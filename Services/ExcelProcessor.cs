using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using JapaneseConverter;
using Text_to_Image.Models;
using Text_to_Image.Data;
using Text_to_Image.Data.Models;

namespace Text_to_Image.Services
{
    public class ExcelProcessor
    {
        private static readonly string FolderPath = @"S:\Anki";

        public static void CleanWorksheet(string filePath)
        {
            try
            {
                Console.WriteLine("Cleaning worksheet data...");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    if (worksheet.Dimension != null)
                    {
                        int lastRow = worksheet.Dimension.End.Row;
                        int lastColumn = worksheet.Dimension.End.Column;

                        worksheet.Cells[1, 1, lastRow, lastColumn].Clear();
                        worksheet.Cells[1, 1, lastRow, lastColumn].Style.Font.Bold = false;
                        worksheet.Cells[1, 1, lastRow, lastColumn].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.None;
                        worksheet.Cells[1, 1, lastRow, lastColumn].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.None);

                        if (worksheet.MergedCells.Count > 0)
                        {
                            var mergedAddresses = new List<string>();
                            foreach (string mergedCell in worksheet.MergedCells)
                            {
                                mergedAddresses.Add(mergedCell);
                            }

                            foreach (string address in mergedAddresses)
                            {
                                worksheet.Cells[address].Merge = false;
                            }

                            Console.WriteLine($"Removed {mergedAddresses.Count} merged cell regions.");
                        }

                        worksheet.ConditionalFormatting.RemoveAll();
                        Console.WriteLine($"Cleaned {lastRow} rows and {lastColumn} columns completely.");
                    }
                    else
                    {
                        Console.WriteLine("Worksheet is empty, no data to clean.");
                    }

                    package.Save();
                }

                Console.WriteLine("Worksheet cleaning completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error cleaning worksheet: {ex.Message}");
                throw;
            }
        }

        // SỬA LẠI: ProcessExcelFile - Hỗ trợ UPDATE duplicates
        public static async Task ProcessExcelFile(ProcessingOptions options)
        {
            try
            {
                RemoveDuplicateRows(options.SelectedFile, options);

                string dateToUse = string.IsNullOrWhiteSpace(options.CustomDate) ?
                    DateTime.Now.ToString("dd-MM-yyyy") : options.CustomDate;

                Console.WriteLine($"Using date: {dateToUse}");

                // Get existing data for duplicate checking
                List<Vocabulary> existingVocabs = new List<Vocabulary>();
                int existingAudioCount = 0;

                try
                {
                    using var dbService = new DatabaseService();
                    string fileType = DetermineFileType(options.FileName);
                    string audioFileType = DetermineAudioFileType(fileType);

                    var existingSummary = await dbService.GetExistingDataSummary(dateToUse, fileType, audioFileType);
                    existingAudioCount = existingSummary.AudioFileCount;
                    existingVocabs = await dbService.GetExistingVocabulariesForDateAsync(dateToUse, fileType);

                    Console.WriteLine("Database Check Results:");
                    Console.WriteLine($"  Existing vocabularies (SAME DAY): {existingSummary.VocabularyCount}");
                    Console.WriteLine($"  Existing audio files: {existingSummary.AudioFileCount}");
                    Console.WriteLine($"  Next audio number starts from: {existingSummary.NextAudioNumber}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Database check failed: {ex.Message}");
                    Console.WriteLine("Will proceed with audio numbering from 1");
                    existingAudioCount = 0;
                }

                // Process Excel
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(options.SelectedFile)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension != null ? worksheet.Dimension.End.Row : 0;

                    if (rowCount == 0)
                    {
                        Console.WriteLine("Worksheet is empty, no data to process.");
                        return;
                    }

                    Console.WriteLine($"Excel file has {rowCount} rows.");
                    Console.WriteLine("Starting processing with conflict resolution logic...");

                    int currentAudioNumber = existingAudioCount + 1;
                    var processedRows = new List<Vocabulary>();
                    var processingResults = new List<VocabularyProcessingInfo>();

                    // PHASE 1: Collect processing info
                    Console.WriteLine("PHASE 1: Collecting processing information...");
                    for (int row = 1; row <= rowCount; row++)
                    {
                        var currentVocab = CreateVocabularyFromRow(worksheet, row, options);
                        var duplicateResult = CheckRowForDuplicate(worksheet, row, existingVocabs, processedRows, options);

                        var processingInfo = new VocabularyProcessingInfo
                        {
                            Vocabulary = currentVocab,
                            DuplicateResult = duplicateResult,
                            RowNumber = row,
                            HasSoundFormulas = HasSoundFormulas(worksheet, row, options)
                        };

                        processingResults.Add(processingInfo);

                        // Track initial CreateNew entries for session duplicate checking
                        if (duplicateResult.Action == "CreateNew")
                        {
                            processedRows.Add(currentVocab);
                        }
                    }

                    // PHASE 2: Apply conflict resolution logic
                    Console.WriteLine("PHASE 2: Applying conflict resolution logic...");
                    ApplyConflictResolutionLogic(processingResults, existingVocabs, options);

                    // PHASE 2.5: Track additional entries that became CreateNewConflict
                    Console.WriteLine("PHASE 2.5: Updating processed rows tracking...");
                    foreach (var info in processingResults)
                    {
                        if (info.DuplicateResult.Action == "CreateNewConflict" &&
                            !processedRows.Any(p => SimilarVocabulary(p, info.Vocabulary, DetermineFileType(options.FileName))))
                        {
                            processedRows.Add(info.Vocabulary);
                        }
                    }

                    // PHASE 3: Execute actions based on processing results
                    Console.WriteLine("PHASE 3: Executing processing actions...");
                    await ExecuteProcessingActions(worksheet, processingResults, options, dateToUse, currentAudioNumber);

                    // PHASE 4: Compact and save
                    Console.WriteLine("PHASE 4: Compacting worksheet and saving...");
                    try
                    {
                        CompactWorksheet(worksheet);
                        package.Save();
                        Console.WriteLine($"Completed! Processed {rowCount} rows.");
                        ShowProcessingSummary(options, processingResults);

                        Console.Write("Open Processed Excel File? (Y/N): ");
                        string openAnswer = Console.ReadLine().Trim().ToUpper();
                        if (openAnswer == "Y")
                        {
                            OpenExcelFile(options.SelectedFile);
                        }
                    }
                    catch (IOException ex)
                    {
                        throw new IOException("Cannot save file. Please ensure the file is not currently open.", ex);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file: {ex.Message}");
                throw;
            }
        }
        private static void ApplyConflictResolutionLogic(List<VocabularyProcessingInfo> processingResults,
    List<Vocabulary> existingVocabs, ProcessingOptions options)
        {
            string fileType = DetermineFileType(options.FileName);

            // Group rows by their conflict target (which existing vocab they conflict with)
            var conflictGroups = new Dictionary<int, List<VocabularyProcessingInfo>>();

            foreach (var info in processingResults)
            {
                if (info.DuplicateResult.IsDuplicate &&
                    info.DuplicateResult.MatchType == "Partial" &&
                    info.DuplicateResult.DuplicateSource == "Database")
                {
                    int existingVocabId = info.DuplicateResult.ExistingVocab.VocabId;

                    if (!conflictGroups.ContainsKey(existingVocabId))
                        conflictGroups[existingVocabId] = new List<VocabularyProcessingInfo>();

                    conflictGroups[existingVocabId].Add(info);
                }
            }

            // Apply conflict resolution rules
            foreach (var group in conflictGroups.Values)
            {
                if (group.Count > 1)
                {
                    // Sort by row number to ensure consistent ordering
                    group.Sort((a, b) => a.RowNumber.CompareTo(b.RowNumber));

                    for (int i = 0; i < group.Count; i++)
                    {
                        var info = group[i];
                        var conflicts = DetectDataConflicts(info.Vocabulary, info.DuplicateResult.ExistingVocab, fileType);

                        if (i == 0)
                        {
                            // FIRST ROW: RESTORE old audio
                            info.DuplicateResult.Action = "RestoreAndUpdate";
                            info.DuplicateResult.IsFirstConflictRow = true;
                            info.DuplicateResult.ConflictFields = conflicts;
                        }
                        else
                        {
                            // SUBSEQUENT ROWS: CREATE new audio
                            info.DuplicateResult.Action = "CreateNewConflict";
                            info.DuplicateResult.ConflictFields = conflicts;
                        }
                    }
                }
                else if (group.Count == 1)
                {
                    // Single partial match
                    var info = group[0];
                    var conflicts = DetectDataConflicts(info.Vocabulary, info.DuplicateResult.ExistingVocab, fileType);
                    info.DuplicateResult.ConflictFields = conflicts;
                }
            }
        }
        private static List<string> DetectDataConflicts(Vocabulary currentVocab, Vocabulary existingVocab, string fileType)
        {
            var conflicts = new List<string>();

            bool HasConflict(string current, string existing, string fieldName)
            {
                if (!string.IsNullOrWhiteSpace(current) && !string.IsNullOrWhiteSpace(existing))
                {
                    if (!SimilarText(current, existing))
                    {
                        conflicts.Add(fieldName);
                        return true;
                    }
                }
                return false;
            }

            switch (fileType.ToLower())
            {
                case "english":
                    HasConflict(currentVocab.VietnameseText, existingVocab.VietnameseText, "Vietnamese");
                    HasConflict(currentVocab.EnglishText, existingVocab.EnglishText, "English");
                    break;
                case "japanese":
                    HasConflict(currentVocab.EnglishText, existingVocab.EnglishText, "English");
                    HasConflict(currentVocab.JapaneseText, existingVocab.JapaneseText, "Japanese");
                    HasConflict(currentVocab.ReadingText, existingVocab.ReadingText, "Reading");
                    break;
                case "chinese":
                    HasConflict(currentVocab.EnglishText, existingVocab.EnglishText, "English");
                    HasConflict(currentVocab.ChineseText, existingVocab.ChineseText, "Chinese");
                    HasConflict(currentVocab.ReadingText, existingVocab.ReadingText, "Reading");
                    break;
                case "tuvung":
                    HasConflict(currentVocab.VietnameseText, existingVocab.VietnameseText, "Vietnamese");
                    HasConflict(currentVocab.JapaneseText, existingVocab.JapaneseText, "Japanese");
                    HasConflict(currentVocab.ReadingText, existingVocab.ReadingText, "Reading");
                    break;
            }

            return conflicts;
        }
        private static async Task ExecuteProcessingActions(ExcelWorksheet worksheet, List<VocabularyProcessingInfo> processingResults,
    ProcessingOptions options, string dateToUse, int startingAudioNumber)
        {
            int currentAudioNumber = startingAudioNumber;

            foreach (var info in processingResults)
            {
                int row = info.RowNumber;
                var duplicateResult = info.DuplicateResult;

                // Text conversion
                if (duplicateResult.Action != "Skip" &&
                    !string.IsNullOrWhiteSpace(options.ColumnInput) &&
                    options.ColumnInput.Length >= 1 &&
                    !options.FileName.Contains("english"))
                {
                    int inputColumn = options.ColumnInput[0] - 'A' + 1;
                    string inputText = worksheet.Cells[row, inputColumn].Text;

                    if (!string.IsNullOrEmpty(inputText))
                    {
                        string convertedText = TextConverter.ConvertToImgTags(inputText);
                        int outputColumn = 7;
                        worksheet.Cells[row, outputColumn].Value = convertedText;
                    }
                }

                // Sound formulas processing
                if (!string.IsNullOrWhiteSpace(options.SoundColumns))
                {
                    switch (duplicateResult.Action)
                    {
                        case "Skip":
                            ClearEntireRow(worksheet, row);
                            Console.WriteLine($"Row {row}: EXACT DUPLICATE - DELETED");
                            break;

                        case "RestoreAndUpdate":
                            await RestoreOldSoundFormulas(worksheet, row, duplicateResult.ExistingVocab, options, dateToUse);
                            if (duplicateResult.ConflictFields.Any())
                            {
                                string conflictFields = string.Join(", ", duplicateResult.ConflictFields);
                                string orderInfo = duplicateResult.IsFirstConflictRow ? " (FIRST)" : "";
                                Console.WriteLine($"Row {row}: CONFLICT{orderInfo} in [{conflictFields}] - RESTORED old audio");
                            }
                            else
                            {
                                Console.WriteLine($"Row {row}: PARTIAL DUPLICATE - RESTORED old audio");
                            }
                            break;

                        case "Clear":
                            ClearSoundFormulas(worksheet, row, options);
                            Console.WriteLine($"Row {row}: SESSION DUPLICATE - CLEARED");
                            break;

                        case "CreateNew":
                            ProcessSoundFormulas(worksheet, row, options, dateToUse, currentAudioNumber);
                            Console.WriteLine($"Row {row}: NEW ENTRY - Created audio {currentAudioNumber} & {currentAudioNumber + 1}");
                            currentAudioNumber += 2;
                            break;

                        case "CreateNewConflict":
                            ProcessSoundFormulas(worksheet, row, options, dateToUse, currentAudioNumber);
                            string conflictInfo = string.Join(", ", duplicateResult.ConflictFields);
                            Console.WriteLine($"Row {row}: CONFLICT (SUBSEQUENT) in [{conflictInfo}] - Created NEW audio {currentAudioNumber} & {currentAudioNumber + 1}");
                            currentAudioNumber += 2;
                            break;
                    }
                }

                // Kanji processing
                if (duplicateResult.Action != "Skip" &&
                    !string.IsNullOrWhiteSpace(options.KanjiColumn) &&
                    options.KanjiColumn.Length >= 1 &&
                    !options.FileName.Contains("english") &&
                    (options.FileName.Contains("tuvung") ||
                     options.FileName.Contains("japanese") ||
                     options.FileName.Contains("chinese")))
                {
                    ProcessKanjiFormatting(worksheet, row, options);
                }
            }
        }
        private static bool HasSoundFormulas(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            if (string.IsNullOrWhiteSpace(options.SoundColumns)) return false;

            if ((options.FileName.Contains("tuvung") ||
                 options.FileName.Contains("japanese") ||
                 options.FileName.Contains("chinese")) && options.SoundColumns.Length == 2)
            {
                string col1Value = worksheet.Cells[row, 5].Text?.Trim();
                string col2Value = worksheet.Cells[row, 6].Text?.Trim();
                return !string.IsNullOrEmpty(col1Value) && !string.IsNullOrEmpty(col2Value) &&
                       col1Value.Contains("[sound:") && col2Value.Contains("[sound:");
            }
            else if (options.FileName.Contains("english") && options.SoundColumns.Length == 2)
            {
                int soundCol1 = options.SoundColumns[0] - 'A' + 1;
                int soundCol2 = options.SoundColumns[1] - 'A' + 1;
                string col1Value = worksheet.Cells[row, soundCol1].Text?.Trim();
                string col2Value = worksheet.Cells[row, soundCol2].Text?.Trim();
                return !string.IsNullOrEmpty(col1Value) && !string.IsNullOrEmpty(col2Value) &&
                       col1Value.Contains("[sound:") && col2Value.Contains("[sound:");
            }
            else if (options.SoundColumns.Length == 1)
            {
                int soundCol = options.SoundColumns[0] - 'A' + 1;
                string colValue = worksheet.Cells[row, soundCol].Text?.Trim();
                return !string.IsNullOrEmpty(colValue) && colValue.Contains("[sound:");
            }

            return false;
        }
        private static void ClearEntireRow(ExcelWorksheet worksheet, int row)
        {
            if (worksheet.Dimension != null)
            {
                int lastColumn = worksheet.Dimension.End.Column;
                for (int col = 1; col <= lastColumn; col++)
                {
                    worksheet.Cells[row, col].Value = "";
                }
            }
        }
        private static void CompactWorksheet(ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension == null) return;

            int totalRows = worksheet.Dimension.End.Row;
            int totalColumns = worksheet.Dimension.End.Column;
            var rowsToDelete = new List<int>();

            Console.WriteLine($"Checking {totalRows} rows for empty content...");

            // TÌM DÒNG TRỐNG
            for (int row = 1; row <= totalRows; row++)
            {
                bool isEmpty = true;
                for (int col = 1; col <= totalColumns; col++)
                {
                    string cellValue = worksheet.Cells[row, col].Text?.Trim();
                    if (!string.IsNullOrWhiteSpace(cellValue))
                    {
                        isEmpty = false;
                        break;
                    }
                }

                if (isEmpty)
                {
                    rowsToDelete.Add(row);
                }
            }

            // XÓA TỪ CUỐI LÊN ĐẦU
            if (rowsToDelete.Any())
            {
                Console.WriteLine($"Found {rowsToDelete.Count} empty rows, removing...");
                for (int i = rowsToDelete.Count - 1; i >= 0; i--)
                {
                    worksheet.DeleteRow(rowsToDelete[i]);
                    Console.WriteLine($"Deleted empty row {rowsToDelete[i]}");
                }

                Console.WriteLine($"Compacted worksheet: removed {rowsToDelete.Count} empty rows");
            }
            else
            {
                Console.WriteLine("No empty rows found - worksheet already clean");
            }
        }
        // SỬA LẠI: CheckRowForDuplicate → trả về duplicate info thay vì boolean
        private static DuplicateCheckResult CheckRowForDuplicate(ExcelWorksheet worksheet, int row,
            List<Vocabulary> existingVocabs, List<Vocabulary> processedRows, ProcessingOptions options)
        {
            var currentVocab = CreateVocabularyFromRow(worksheet, row, options);
            string fileType = DetermineFileType(options.FileName);

            // CHECK 1: EXACT MATCH (tất cả primary fields giống nhau)
            var exactMatch = FindExactMatchInDatabase(currentVocab, existingVocabs, fileType);

            if (exactMatch != null)
            {
                return new DuplicateCheckResult
                {
                    IsDuplicate = true,
                    ExistingVocab = exactMatch,
                    DuplicateSource = "Database",
                    MatchType = "Exact",
                    Action = "Skip"
                };
            }

            // CHECK 2: PARTIAL MATCH (ít nhất 1 primary field giống nhau)
            var partialMatch = FindPartialMatchInDatabase(currentVocab, existingVocabs, fileType);

            if (partialMatch != null)
            {

                return new DuplicateCheckResult
                {
                    IsDuplicate = true,
                    ExistingVocab = partialMatch,
                    DuplicateSource = "Database",
                    MatchType = "Partial",
                    Action = "RestoreAndUpdate"
                };
            }

            // CHECK 3: Current session duplicates (exact match only)
            var sessionExactMatch = FindExactMatchInCurrentSession(currentVocab, processedRows, fileType);

            if (sessionExactMatch != null)
            {
                return new DuplicateCheckResult
                {
                    IsDuplicate = true,
                    ExistingVocab = sessionExactMatch,
                    DuplicateSource = "CurrentSession",
                    MatchType = "Exact",
                    Action = "Clear"
                };
            }

            // CHECK 4: Current session partial duplicates
            var sessionPartialMatch = FindPartialMatchInCurrentSession(currentVocab, processedRows, fileType);

            if (sessionPartialMatch != null)
            {
                return new DuplicateCheckResult
                {
                    IsDuplicate = true,
                    ExistingVocab = sessionPartialMatch,
                    DuplicateSource = "CurrentSession",
                    MatchType = "Partial",
                    Action = "Clear"
                };
            }

            return new DuplicateCheckResult
            {
                IsDuplicate = false,
                Action = "CreateNew"
            };
        }
        private static Vocabulary FindExactMatchInDatabase(Vocabulary currentVocab, List<Vocabulary> existingVocabs, string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.VietnameseText, currentVocab.VietnameseText) &&
                    SimilarText(existing.EnglishText, currentVocab.EnglishText)
                ),

                "japanese" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.EnglishText, currentVocab.EnglishText) &&
                    SimilarText(existing.JapaneseText, currentVocab.JapaneseText)
                ),

                "chinese" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.EnglishText, currentVocab.EnglishText) &&
                    SimilarText(existing.ChineseText, currentVocab.ChineseText)
                ),

                "tuvung" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.VietnameseText, currentVocab.VietnameseText) &&
                    SimilarText(existing.JapaneseText, currentVocab.JapaneseText)
                ),

                _ => null
            };
        }

        private static Vocabulary FindExactMatchInCurrentSession(Vocabulary currentVocab, List<Vocabulary> processedRows, string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => processedRows.FirstOrDefault(processed =>
                    SimilarText(processed.VietnameseText, currentVocab.VietnameseText) &&
                    SimilarText(processed.EnglishText, currentVocab.EnglishText)
                ),

                "japanese" => processedRows.FirstOrDefault(processed =>
                    SimilarText(processed.EnglishText, currentVocab.EnglishText) &&
                    SimilarText(processed.JapaneseText, currentVocab.JapaneseText)
                ),

                "chinese" => processedRows.FirstOrDefault(processed =>
                    SimilarText(processed.EnglishText, currentVocab.EnglishText) &&
                    SimilarText(processed.ChineseText, currentVocab.ChineseText)
                ),

                "tuvung" => processedRows.FirstOrDefault(processed =>
                    SimilarText(processed.VietnameseText, currentVocab.VietnameseText) &&
                    SimilarText(processed.JapaneseText, currentVocab.JapaneseText)
                ),

                _ => null
            };
        }
        private static Vocabulary FindPartialMatchInDatabase(Vocabulary currentVocab, List<Vocabulary> existingVocabs, string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.VietnameseText, currentVocab.VietnameseText) ||
                    SimilarText(existing.EnglishText, currentVocab.EnglishText)
                ),

                "japanese" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.EnglishText, currentVocab.EnglishText) ||
                    SimilarText(existing.JapaneseText, currentVocab.JapaneseText)
                ),

                "chinese" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.EnglishText, currentVocab.EnglishText) ||
                    SimilarText(existing.ChineseText, currentVocab.ChineseText)
                ),

                "tuvung" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.VietnameseText, currentVocab.VietnameseText) ||
                    SimilarText(existing.JapaneseText, currentVocab.JapaneseText)
                ),

                _ => null
            };
        }
        private static Vocabulary FindPartialMatchInCurrentSession(Vocabulary currentVocab, List<Vocabulary> processedRows, string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => processedRows.FirstOrDefault(processed =>
                    SimilarText(processed.VietnameseText, currentVocab.VietnameseText) ||
                    SimilarText(processed.EnglishText, currentVocab.EnglishText)
                ),

                "japanese" => processedRows.FirstOrDefault(processed =>
                    SimilarText(processed.EnglishText, currentVocab.EnglishText) ||
                    SimilarText(processed.JapaneseText, currentVocab.JapaneseText)
                ),

                "chinese" => processedRows.FirstOrDefault(processed =>
                    SimilarText(processed.EnglishText, currentVocab.EnglishText) ||
                    SimilarText(processed.ChineseText, currentVocab.ChineseText)
                ),

                "tuvung" => processedRows.FirstOrDefault(processed =>
                    SimilarText(processed.VietnameseText, currentVocab.VietnameseText) ||
                    SimilarText(processed.JapaneseText, currentVocab.JapaneseText)
                ),

                _ => null
            };
        }


        // THÊM MỚI: Restore old sound formulas từ database
        private static async Task RestoreOldSoundFormulas(ExcelWorksheet worksheet, int row, Vocabulary existingVocab, ProcessingOptions options, string dateToUse)
        {
            if (string.IsNullOrWhiteSpace(options.SoundColumns)) return;

            try
            {
                // Lấy audio files cũ từ database
                using var dbService = new DatabaseService();
                var existingAudioFiles = await dbService.GetAudioFilesByVocabIdAsync(existingVocab.VocabId);

                if (existingAudioFiles != null && existingAudioFiles.Count >= 2)
                {
                    if ((options.FileName.Contains("tuvung") ||
                         options.FileName.Contains("japanese") ||
                         options.FileName.Contains("chinese")) && options.SoundColumns.Length == 2)
                    {
                        // Sắp xếp theo IsOddFile để đảm bảo đúng thứ tự
                        var oddFile = existingAudioFiles.FirstOrDefault(a => a.IsOddFile);
                        var evenFile = existingAudioFiles.FirstOrDefault(a => !a.IsOddFile);

                        if (oddFile != null)
                            worksheet.Cells[row, 5].Value = $"[sound:{oddFile.FileName}]";
                        if (evenFile != null)
                            worksheet.Cells[row, 6].Value = $"[sound:{evenFile.FileName}]";

                        Console.WriteLine($"Row {row}: Restored formulas - {oddFile?.FileName}, {evenFile?.FileName}");
                    }
                    else if (options.FileName.Contains("english") && options.SoundColumns.Length == 2)
                    {
                        int soundCol1 = options.SoundColumns[0] - 'A' + 1;
                        int soundCol2 = options.SoundColumns[1] - 'A' + 1;

                        var oddFile = existingAudioFiles.FirstOrDefault(a => a.IsOddFile);
                        var evenFile = existingAudioFiles.FirstOrDefault(a => !a.IsOddFile);

                        if (oddFile != null)
                            worksheet.Cells[row, soundCol1].Value = $"[sound:{oddFile.FileName}]";
                        if (evenFile != null)
                            worksheet.Cells[row, soundCol2].Value = $"[sound:{evenFile.FileName}]";

                        Console.WriteLine($"Row {row}: Restored formulas - {oddFile?.FileName}, {evenFile?.FileName}");
                    }
                }
                else
                {
                    // Nếu không tìm thấy audio files cũ, clear formulas
                    ClearSoundFormulas(worksheet, row, options);
                    Console.WriteLine($"Row {row}: No existing audio files found - CLEARED formulas");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not restore old formulas for row {row}: {ex.Message}");
                // Fallback: clear formulas
                ClearSoundFormulas(worksheet, row, options);
            }
        }

        private static void ProcessSoundFormulas(ExcelWorksheet worksheet, int row, ProcessingOptions options,
            string dateToUse, int currentAudioNumber)
        {
            int oddNumber = currentAudioNumber;
            int evenNumber = currentAudioNumber + 1;

            if (options.FileName.Contains("tuvung") && options.SoundColumns.Length == 2)
            {
                worksheet.Cells[row, 5].Value = $"[sound:Vocab-{dateToUse}_{oddNumber:00}.mp3]";
                worksheet.Cells[row, 6].Value = $"[sound:Vocab-{dateToUse}_{evenNumber:00}.mp3]";
            }
            else if (options.FileName.Contains("japanese") && options.SoundColumns.Length == 2)
            {
                worksheet.Cells[row, 5].Value = $"[sound:JP-{dateToUse}_{oddNumber:00}.mp3]";
                worksheet.Cells[row, 6].Value = $"[sound:JP-{dateToUse}_{evenNumber:00}.mp3]";
            }
            else if (options.FileName.Contains("chinese") && options.SoundColumns.Length == 2)
            {
                worksheet.Cells[row, 5].Value = $"[sound:ZH-{dateToUse}_{oddNumber:00}.mp3]";
                worksheet.Cells[row, 6].Value = $"[sound:ZH-{dateToUse}_{evenNumber:00}.mp3]";
            }
            else if (options.FileName.Contains("english") && options.SoundColumns.Length == 2)
            {
                int soundCol1 = options.SoundColumns[0] - 'A' + 1;
                int soundCol2 = options.SoundColumns[1] - 'A' + 1;
                worksheet.Cells[row, soundCol1].Value = $"[sound:EN-{dateToUse}_{oddNumber:00}.mp3]";
                worksheet.Cells[row, soundCol2].Value = $"[sound:EN-{dateToUse}_{evenNumber:00}.mp3]";
            }
            else if (options.SoundColumns.Length == 1)
            {
                int soundCol = options.SoundColumns[0] - 'A' + 1;
                string prefix = DetermineLanguagePrefix(options.FileName);
                if (!string.IsNullOrEmpty(prefix))
                {
                    worksheet.Cells[row, soundCol].Value = $"[sound:{prefix}-{dateToUse}_{oddNumber:00}.mp3]";
                }
            }
        }

        private static Vocabulary CreateVocabularyFromRow(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            var vocab = new Vocabulary();
            string fileType = DetermineFileType(options.FileName);

            switch (fileType.ToLower())
            {
                case "english":
                    vocab.VietnameseText = GetCellValue(worksheet, row, 1);
                    vocab.EnglishText = GetCellValue(worksheet, row, 2);
                    break;
                case "japanese":
                    vocab.EnglishText = GetCellValue(worksheet, row, 1);
                    vocab.ReadingText = GetCellValue(worksheet, row, 2);
                    vocab.JapaneseText = GetCellValue(worksheet, row, 3);
                    break;
                case "chinese":
                    vocab.EnglishText = GetCellValue(worksheet, row, 1);
                    vocab.ReadingText = GetCellValue(worksheet, row, 2);
                    vocab.ChineseText = GetCellValue(worksheet, row, 3);
                    break;
                case "tuvung":
                    vocab.VietnameseText = GetCellValue(worksheet, row, 1);
                    vocab.ReadingText = GetCellValue(worksheet, row, 2);
                    vocab.JapaneseText = GetCellValue(worksheet, row, 3);
                    break;
            }

            return vocab;
        }

        private static string GetPrimaryText(Vocabulary vocab)
        {
            return !string.IsNullOrEmpty(vocab.EnglishText) ? vocab.EnglishText :
                   !string.IsNullOrEmpty(vocab.VietnameseText) ? vocab.VietnameseText :
                   !string.IsNullOrEmpty(vocab.JapaneseText) ? vocab.JapaneseText :
                   vocab.ChineseText ?? "";
        }

        private static string GetCellValue(ExcelWorksheet worksheet, int row, int column)
        {
            var value = worksheet.Cells[row, column].Text?.Trim();
            return string.IsNullOrEmpty(value) ? "" : value;
        }

        private static bool SimilarText(string text1, string text2)
        {
            if (string.IsNullOrWhiteSpace(text1) || string.IsNullOrWhiteSpace(text2))
                return false;
            return text1.Trim().Equals(text2.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private static void ClearSoundFormulas(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            if (string.IsNullOrWhiteSpace(options.SoundColumns)) return;

            if ((options.FileName.Contains("tuvung") ||
                 options.FileName.Contains("japanese") ||
                 options.FileName.Contains("chinese")) && options.SoundColumns.Length == 2)
            {
                worksheet.Cells[row, 5].Value = "";
                worksheet.Cells[row, 6].Value = "";
            }
            else if (options.FileName.Contains("english") && options.SoundColumns.Length == 2)
            {
                int soundCol1 = options.SoundColumns[0] - 'A' + 1;
                int soundCol2 = options.SoundColumns[1] - 'A' + 1;
                worksheet.Cells[row, soundCol1].Value = "";
                worksheet.Cells[row, soundCol2].Value = "";
            }
            else if (options.SoundColumns.Length == 1)
            {
                int soundCol = options.SoundColumns[0] - 'A' + 1;
                worksheet.Cells[row, soundCol].Value = "";
            }
        }

        private static void ProcessKanjiFormatting(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            int sourceCol = options.KanjiColumn[0] - 'A' + 1;
            int outputCol = sourceCol;

            if (!string.IsNullOrWhiteSpace(options.KanjiOutputColumn) && options.KanjiOutputColumn.Length >= 1)
            {
                outputCol = options.KanjiOutputColumn[0] - 'A' + 1;
            }

            string cellValue = worksheet.Cells[row, sourceCol].Text;
            if (!string.IsNullOrEmpty(cellValue))
            {
                string formattedText = KanjiHelper.FormatKanjiWithBrackets(cellValue);
                worksheet.Cells[row, outputCol].Value = formattedText;
            }
        }
        private static int RemoveDuplicateRows(string filePath, ProcessingOptions options)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var package = new ExcelPackage(new FileInfo(filePath));
                var worksheet = package.Workbook.Worksheets[0];

                if (worksheet.Dimension == null) return 0;

                int totalRows = worksheet.Dimension.End.Row;
                var duplicateRows = new List<int>();
                var seenData = new HashSet<string>();

                for (int row = 1; row <= totalRows; row++)
                {
                    string rowKey = CreateRowKey(worksheet, row, options);

                    if (string.IsNullOrWhiteSpace(rowKey)) continue;

                    if (seenData.Contains(rowKey))
                    {
                        duplicateRows.Add(row);
                    }
                    else
                    {
                        seenData.Add(rowKey);
                    }
                }

                // Xóa từ cuối lên đầu
                for (int i = duplicateRows.Count - 1; i >= 0; i--)
                {
                    worksheet.DeleteRow(duplicateRows[i]);
                }

                if (duplicateRows.Count > 0)
                {
                    package.Save();
                    Console.WriteLine($"Removed {duplicateRows.Count} duplicate rows.");
                }

                return duplicateRows.Count;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error removing duplicates: {ex.Message}");
                return 0;
            }
        }
        private static string CreateRowKey(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            string fileName = options.FileName.ToLower();

            if (fileName.Contains("english"))
            {
                string vietnamese = GetCellValue(worksheet, row, 1);
                string english = GetCellValue(worksheet, row, 2);
                if (string.IsNullOrWhiteSpace(vietnamese) && string.IsNullOrWhiteSpace(english)) return "";
                return $"VI:{vietnamese}|EN:{english}";
            }
            else if (fileName.Contains("japanese"))
            {
                string english = GetCellValue(worksheet, row, 1);
                string reading = GetCellValue(worksheet, row, 2);
                string japanese = GetCellValue(worksheet, row, 3);
                if (string.IsNullOrWhiteSpace(english) && string.IsNullOrWhiteSpace(reading) && string.IsNullOrWhiteSpace(japanese)) return "";
                return $"EN:{english}|RD:{reading}|JP:{japanese}";
            }
            else if (fileName.Contains("chinese"))
            {
                string english = GetCellValue(worksheet, row, 1);
                string reading = GetCellValue(worksheet, row, 2);
                string chinese = GetCellValue(worksheet, row, 3);
                if (string.IsNullOrWhiteSpace(english) && string.IsNullOrWhiteSpace(reading) && string.IsNullOrWhiteSpace(chinese)) return "";
                return $"EN:{english}|RD:{reading}|ZH:{chinese}";
            }
            else if (fileName.Contains("tuvung"))
            {
                string vietnamese = GetCellValue(worksheet, row, 1);
                string reading = GetCellValue(worksheet, row, 2);
                string japanese = GetCellValue(worksheet, row, 3);
                if (string.IsNullOrWhiteSpace(vietnamese) && string.IsNullOrWhiteSpace(reading) && string.IsNullOrWhiteSpace(japanese)) return "";
                return $"VI:{vietnamese}|RD:{reading}|JP:{japanese}";
            }

            return "";
        }
        private static bool SimilarVocabulary(Vocabulary vocab1, Vocabulary vocab2, string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => SimilarText(vocab1.VietnameseText, vocab2.VietnameseText) &&
                             SimilarText(vocab1.EnglishText, vocab2.EnglishText),
                "japanese" => SimilarText(vocab1.EnglishText, vocab2.EnglishText) &&
                              SimilarText(vocab1.JapaneseText, vocab2.JapaneseText),
                "chinese" => SimilarText(vocab1.EnglishText, vocab2.EnglishText) &&
                             SimilarText(vocab1.ChineseText, vocab2.ChineseText),
                "tuvung" => SimilarText(vocab1.VietnameseText, vocab2.VietnameseText) &&
                            SimilarText(vocab1.JapaneseText, vocab2.JapaneseText),
                _ => false
            };
        }

        private static void ShowProcessingSummary(ProcessingOptions options, List<VocabularyProcessingInfo> processingResults)
        {
            if (!string.IsNullOrWhiteSpace(options.ColumnInput) && !options.FileName.Contains("english"))
            {
                Console.WriteLine($"- <img>: done | {options.ColumnInput[0]} -> G");
            }

            if (!string.IsNullOrWhiteSpace(options.SoundColumns))
            {
                var actionCounts = processingResults
                    .GroupBy(p => p.DuplicateResult.Action)
                    .ToDictionary(g => g.Key, g => g.Count());

                Console.WriteLine("- [sound]: Processing completed");
                Console.WriteLine("  Results:");

                if (actionCounts.ContainsKey("Skip"))
                    Console.WriteLine($"    * EXACT DUPLICATES (deleted): {actionCounts["Skip"]}");
                if (actionCounts.ContainsKey("RestoreAndUpdate"))
                    Console.WriteLine($"    * RESTORED old audio: {actionCounts["RestoreAndUpdate"]}");
                if (actionCounts.ContainsKey("CreateNew"))
                    Console.WriteLine($"    * NEW ENTRIES: {actionCounts["CreateNew"]}");
                if (actionCounts.ContainsKey("CreateNewConflict"))
                    Console.WriteLine($"    * CONFLICT ROWS (new audio): {actionCounts["CreateNewConflict"]}");
                if (actionCounts.ContainsKey("Clear"))
                    Console.WriteLine($"    * SESSION DUPLICATES (cleared): {actionCounts["Clear"]}");
            }

            if (!string.IsNullOrWhiteSpace(options.KanjiColumn))
                Console.WriteLine($"- kanji: done | {options.KanjiColumn}");
        }

        private static string DetermineFileType(string fileName)
        {
            fileName = fileName.ToLower();
            if (fileName.Contains("english")) return "English";
            if (fileName.Contains("japanese")) return "Japanese";
            if (fileName.Contains("chinese")) return "Chinese";
            if (fileName.Contains("tuvung")) return "TuVung";
            return "Unknown";
        }

        private static string DetermineAudioFileType(string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => "VI-EN",
                "japanese" => "JP-EN",
                "chinese" => "ZH-EN",
                "tuvung" => "TUVUNG",
                _ => "VI-EN"
            };
        }

        private static string DetermineLanguagePrefix(string fileName)
        {
            fileName = fileName.ToLower();
            if (fileName.Contains("english")) return "EN";
            if (fileName.Contains("japanese")) return "JP";
            if (fileName.Contains("chinese")) return "ZH";
            return "";
        }
        // THÊM MỚI: Duplicate check result class
        public class DuplicateCheckResult
        {
            public bool IsDuplicate { get; set; }
            public Vocabulary ExistingVocab { get; set; }
            public string DuplicateSource { get; set; } = ""; // "Database" or "CurrentSession"
            public string MatchType { get; set; } = ""; // "Exact" or "Partial"  
            public string Action { get; set; } = ""; // "Skip", "RestoreAndUpdate", "Clear", "CreateNew", "CreateNewConflict"
            public List<string> ConflictFields { get; set; } = new List<string>(); // THÊM MỚI
            public bool IsFirstConflictRow { get; set; } = false; // THÊM MỚI
        }
        public class VocabularyProcessingInfo
        {
            public Vocabulary Vocabulary { get; set; }
            public DuplicateCheckResult DuplicateResult { get; set; }
            public int RowNumber { get; set; }
            public bool HasSoundFormulas { get; set; }
        }
        // CÁC METHODS KHÁC GIỮ NGUYÊN...
        public static System.Diagnostics.Process OpenExcelFile(string filePath)
        {
            try
            {
                var processInfo = new System.Diagnostics.ProcessStartInfo(filePath)
                {
                    UseShellExecute = true
                };
                return System.Diagnostics.Process.Start(processInfo);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error opening Excel file: {ex.Message}");
                return null;
            }
        }

        public static bool SelectExcelFile(ProcessingOptions options)
        {
            string[] excelFiles = Directory.GetFiles(FolderPath, "*.xlsm");

            Console.WriteLine("Available Excel files:");
            for (int i = 0; i < excelFiles.Length; i++)
            {
                Console.WriteLine($"{i + 1}. {Path.GetFileName(excelFiles[i])}");
            }

            Console.Write("Please enter the file number to process (press Enter to exit). ");
            string fileInput = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(fileInput))
                return false;

            int fileChoice = int.Parse(fileInput) - 1;
            if (fileChoice < 0 || fileChoice >= excelFiles.Length)
            {
                throw new ArgumentException("Invalid file number!");
            }

            options.SelectedFile = excelFiles[fileChoice];
            options.FileName = Path.GetFileName(options.SelectedFile).ToLower();
            return true;
        }

        public static bool OpenAndWaitForExcelFile(ProcessingOptions options)
        {
            CleanWorksheet(options.SelectedFile);

            Console.WriteLine("Opening cleaned Excel file...");
            var excelProcess = OpenExcelFile(options.SelectedFile);

            if (excelProcess != null)
            {
                Console.WriteLine("File opened successfully...");
                Console.WriteLine("Waiting for you to close Excel file...");

                excelProcess.WaitForExit();
                Console.WriteLine("Detected Excel file has been closed.");
                return true;
            }

            return false;
        }

        public static async Task ExecuteTasks(ProcessingOptions options)
        {
            if (options.RenameAudioFiles)
            {
                AudioFileManager.RenameAudioFiles(options.AudioFolderPath, options.SelectedDay,
                    options.SelectedMonth, options.SelectedYear, options.FileName);
            }

            if (!string.IsNullOrWhiteSpace(options.ColumnInput) ||
                !string.IsNullOrWhiteSpace(options.SoundColumns) ||
                !string.IsNullOrWhiteSpace(options.KanjiColumn))
            {
                await ProcessExcelFile(options);
            }
            else
            {
                Console.WriteLine("No tasks selected to perform...");
            }
        }

        public static async Task ExecuteTasksWithDatabase(ProcessingOptions options)
        {
            if (options.RenameAudioFiles)
            {
                AudioFileManager.RenameAudioFiles(options.AudioFolderPath, options.SelectedDay,
                    options.SelectedMonth, options.SelectedYear, options.FileName);
            }

            if (!string.IsNullOrWhiteSpace(options.ColumnInput) ||
                !string.IsNullOrWhiteSpace(options.SoundColumns) ||
                !string.IsNullOrWhiteSpace(options.KanjiColumn))
            {
                await ProcessExcelFile(options);
            }
            else
            {
                Console.WriteLine("No tasks selected to perform...");
                return;
            }

            if (options.SaveToDatabase)
            {
                try
                {
                    Console.WriteLine("==================================================");

                    using var dbService = new DatabaseService();
                    var session = await dbService.SaveExcelDataToDatabaseAsync(options);

                    Console.WriteLine("Database Summary:");
                    Console.WriteLine($"Session ID: {session.SessionId}");
                    Console.WriteLine($"NEW entries: {session.ProcessedRows} vocabulary entries");
                    Console.WriteLine($"UPDATED entries: {session.UpdatedRows} vocabulary entries"); // THÊM MỚI
                    Console.WriteLine($"Audio records: {session.AudioFilesCreated}");
                    Console.WriteLine($"File type: {session.FileType}");
                    Console.WriteLine($"Date: {session.DateUsed}");
                    Console.WriteLine("==================================================");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Database save failed: {ex.Message}");
                    Console.WriteLine("Excel processing completed but data not saved to database.");
                    Console.WriteLine("Please check your SQL Server connection.");
                    Console.WriteLine($"Error details: {ex.InnerException?.Message}");
                }
            }
            else
            {
                Console.WriteLine("Database save skipped (user choice).");
                Console.WriteLine("Excel processing completed successfully without database save.");
            }
        }
    }
}