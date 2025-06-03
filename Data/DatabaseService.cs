using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using Text_to_Image.Data.Models;
using Text_to_Image.Models;

namespace Text_to_Image.Data
{
    public class DatabaseService : IDisposable
    {
        private readonly LanguageLearningContext _context;

        public DatabaseService()
        {
            _context = new LanguageLearningContext();
            _context.Database.EnsureCreated();
        }

        // THÊM DEBUG VÀO SaveExcelDataToDatabaseAsync METHOD

        public async Task<ProcessingSession> SaveExcelDataToDatabaseAsync(ProcessingOptions options)
        {
            try
            {
                Console.WriteLine("=== DEBUG: Starting SaveExcelDataToDatabaseAsync ===");
                Console.WriteLine($"Selected file: {options.SelectedFile}");
                Console.WriteLine($"Custom date: {options.CustomDate}");

                string dateToUse = options.CustomDate ?? DateTime.Now.ToString("dd-MM-yyyy");
                DateTime targetDate = DateTime.ParseExact(dateToUse, "dd-MM-yyyy", null);

                Console.WriteLine($"Date to use: {dateToUse}");
                Console.WriteLine($"Target date: {targetDate}");

                // CREATE SESSION
                var session = new ProcessingSession
                {
                    FileName = Path.GetFileName(options.SelectedFile),
                    FileType = DetermineFileType(options.FileName),
                    ProcessedDate = targetDate,
                    DateUsed = dateToUse,
                    Notes = $"Processed with columns: IMG({GetColumnString(options.ColumnInput)}), Sound({options.SoundColumns}), Kanji({options.KanjiColumn})",
                    ProcessedRows = 0,
                    UpdatedRows = 0
                };

                Console.WriteLine($"Session created: FileType={session.FileType}");

                _context.ProcessingSessions.Add(session);
                await _context.SaveChangesAsync();
                Console.WriteLine($"Session saved with ID: {session.SessionId}");

                // READ EXCEL DATA
                Console.WriteLine("=== DEBUG: Reading Excel data ===");
                var allVocabsFromExcel = await ProcessExcelWithUpsertLogic(options, session.SessionId, targetDate);
                Console.WriteLine($"Read {allVocabsFromExcel.Count} vocabularies from Excel");

                // GET EXISTING DATA
                Console.WriteLine("=== DEBUG: Getting existing data ===");
                var existingVocabs = await GetExistingVocabulariesForDateAsync(dateToUse, session.FileType);
                Console.WriteLine($"Found {existingVocabs.Count} existing vocabularies in database");

                // DEBUG: Print existing vocab details
                foreach (var existing in existingVocabs)
                {
                    Console.WriteLine($"  Existing: VocabId={existing.VocabId}, VI='{existing.VietnameseText}', EN='{existing.EnglishText}'");
                }

                // PROCESS EACH VOCABULARY
                Console.WriteLine("=== DEBUG: Processing vocabularies ===");
                int newCount = 0, updateCount = 0, skipCount = 0;
                var newVocabularies = new List<Vocabulary>();
                var updatedVocabularies = new List<Vocabulary>();

                foreach (var currentVocab in allVocabsFromExcel)
                {
                    Console.WriteLine($"\n--- Processing vocab: VI='{currentVocab.VietnameseText}', EN='{currentVocab.EnglishText}' ---");

                    var duplicateResult = CheckVocabularyForDuplicate(currentVocab, existingVocabs, session.FileType);

                    Console.WriteLine($"Duplicate check result: IsDuplicate={duplicateResult.IsDuplicate}, Action={duplicateResult.Action}");

                    if (duplicateResult.ExistingVocab != null)
                    {
                        Console.WriteLine($"  Matched with existing VocabId={duplicateResult.ExistingVocab.VocabId}");
                    }

                    switch (duplicateResult.Action)
                    {
                        case "Skip":
                            skipCount++;
                            Console.WriteLine($"  SKIPPED: {GetPrimaryText(currentVocab, session.FileType)}");
                            break;

                        case "Update":
                            await UpdateExistingVocabulary(duplicateResult.ExistingVocab, currentVocab, options, dateToUse, targetDate);
                            updatedVocabularies.Add(duplicateResult.ExistingVocab);
                            updateCount++;
                            Console.WriteLine($"  UPDATED: {GetPrimaryText(currentVocab, session.FileType)}");
                            break;

                        case "Insert":
                            newVocabularies.Add(currentVocab);
                            newCount++;
                            Console.WriteLine($"  WILL INSERT: {GetPrimaryText(currentVocab, session.FileType)}");
                            break;
                    }
                }

                Console.WriteLine($"\n=== DEBUG: Summary before save ===");
                Console.WriteLine($"New vocabularies to insert: {newVocabularies.Count}");
                Console.WriteLine($"Existing vocabularies to update: {updatedVocabularies.Count}");
                Console.WriteLine($"Vocabularies to skip: {skipCount}");

                // SAVE NEW VOCABULARIES
                if (newVocabularies.Any())
                {
                    Console.WriteLine("=== DEBUG: Saving new vocabularies ===");
                    try
                    {
                        _context.Vocabularies.AddRange(newVocabularies);
                        await _context.SaveChangesAsync();
                        Console.WriteLine($"✓ Successfully saved {newVocabularies.Count} NEW vocabularies");

                        // Print new VocabIds
                        foreach (var newVocab in newVocabularies)
                        {
                            Console.WriteLine($"  New VocabId: {newVocab.VocabId} - {GetPrimaryText(newVocab, session.FileType)}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"✗ Error saving new vocabularies: {ex.Message}");
                        throw;
                    }
                }
                else
                {
                    Console.WriteLine("No new vocabularies to save");
                }

                // SAVE UPDATED VOCABULARIES
                if (updatedVocabularies.Any())
                {
                    Console.WriteLine("=== DEBUG: Saving updated vocabularies ===");
                    try
                    {
                        await _context.SaveChangesAsync();
                        Console.WriteLine($"✓ Successfully updated {updatedVocabularies.Count} existing vocabularies");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"✗ Error updating vocabularies: {ex.Message}");
                        throw;
                    }
                }
                else
                {
                    Console.WriteLine("No vocabularies to update");
                }

                // CREATE AUDIO FILES
                Console.WriteLine("=== DEBUG: Creating audio file records ===");
                if (options.CreateAudioFiles && !string.IsNullOrWhiteSpace(options.SoundColumns))
                {
                    var allVocabsForAudio = newVocabularies.Concat(updatedVocabularies).ToList();
                    Console.WriteLine($"Creating audio files for {allVocabsForAudio.Count} vocabularies");

                    if (allVocabsForAudio.Any())
                    {
                        var audioFiles = CreateAudioFileRecordsFromVocabularies(allVocabsForAudio, options, targetDate);
                        Console.WriteLine($"Created {audioFiles.Count} audio file records");

                        if (audioFiles.Any())
                        {
                            _context.AudioFiles.AddRange(audioFiles);
                            await _context.SaveChangesAsync();
                            session.AudioFilesCreated = audioFiles.Count;
                            Console.WriteLine($"✓ Saved {audioFiles.Count} audio file records to database");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Audio file creation skipped (not enabled or no sound columns)");
                }

                // UPDATE SESSION
                session.ProcessedRows = newCount;
                session.UpdatedRows = updateCount;
                session.TotalRows = allVocabsFromExcel.Count;
                session.IsCompleted = true;
                await _context.SaveChangesAsync();

                Console.WriteLine("=== DEBUG: Final results ===");
                Console.WriteLine($"NEW entries: {newCount}");
                Console.WriteLine($"UPDATED entries: {updateCount}");
                Console.WriteLine($"SKIPPED entries: {skipCount}");
                Console.WriteLine("=== DEBUG: SaveExcelDataToDatabaseAsync completed ===");

                return session;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"=== DEBUG: ERROR in SaveExcelDataToDatabaseAsync ===");
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
                }
                throw;
            }
        }

        private async Task<List<Vocabulary>> ProcessExcelWithUpsertLogic(ProcessingOptions options, int sessionId, DateTime targetDate)
        {
            var vocabularies = new List<Vocabulary>();
            string dateToUse = options.CustomDate ?? DateTime.Now.ToString("dd-MM-yyyy");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(options.SelectedFile));
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension?.End.Row ?? 0;

            if (rowCount == 0) return vocabularies;

            string fileType = DetermineFileType(options.FileName);

            for (int row = 1; row <= rowCount; row++)
            {
                // ĐỌC TẤT CẢ ROWS, KHÔNG CHỈ ROWS CÓ SOUND FORMULAS
                var vocab = CreateVocabularyFromExcelRow(worksheet, row, fileType, options.SelectedFile, targetDate);

                // CHỈ THÊM VÀO LIST NẾU CÓ DATA
                if (HasValidData(vocab, fileType))
                {
                    vocabularies.Add(vocab);
                }
            }

            return vocabularies;
        }
        private DuplicateCheckResult CheckVocabularyForDuplicate(Vocabulary currentVocab, List<Vocabulary> existingVocabs, string fileType)
        {
            // CHECK 1: EXACT MATCH (tất cả primary fields giống nhau)
            var exactMatch = FindExactMatchInDatabase(currentVocab, existingVocabs, fileType);
            if (exactMatch != null)
            {
                return new DuplicateCheckResult
                {
                    IsDuplicate = true,
                    ExistingVocab = exactMatch,
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
                    Action = "Update"
                };
            }

            // CHECK 3: NEW ENTRY
            return new DuplicateCheckResult
            {
                IsDuplicate = false,
                Action = "Insert"
            };
        }

        // THÊM MỚI: Update existing vocabulary
        private async Task UpdateExistingVocabulary(Vocabulary existingVocab, Vocabulary newVocab,
            ProcessingOptions options, string dateToUse, DateTime targetDate)
        {
            // UPDATE VOCABULARY FIELDS
            if (!string.IsNullOrEmpty(newVocab.VietnameseText))
                existingVocab.VietnameseText = newVocab.VietnameseText;
            if (!string.IsNullOrEmpty(newVocab.EnglishText))
                existingVocab.EnglishText = newVocab.EnglishText;
            if (!string.IsNullOrEmpty(newVocab.JapaneseText))
                existingVocab.JapaneseText = newVocab.JapaneseText;
            if (!string.IsNullOrEmpty(newVocab.ChineseText))
                existingVocab.ChineseText = newVocab.ChineseText;
            if (!string.IsNullOrEmpty(newVocab.ReadingText))
                existingVocab.ReadingText = newVocab.ReadingText;
            if (!string.IsNullOrEmpty(newVocab.KanjiText))
                existingVocab.KanjiText = newVocab.KanjiText;
            if (!string.IsNullOrEmpty(newVocab.ImageTags))
                existingVocab.ImageTags = newVocab.ImageTags;

            // UPDATE DATES
            existingVocab.CreatedDate = targetDate; // Update creation date

            // MARK AS MODIFIED
            _context.Entry(existingVocab).State = EntityState.Modified;

            // TẠO AUDIO FILES MỚI CHO UPDATED VOCABULARY (nếu cần)
            if (options.CreateAudioFiles && !string.IsNullOrWhiteSpace(options.SoundColumns))
            {
                // Xóa audio files cũ
                var oldAudioFiles = await _context.AudioFiles
                    .Where(a => a.VocabId == existingVocab.VocabId)
                    .ToListAsync();
                _context.AudioFiles.RemoveRange(oldAudioFiles);

                // Tạo audio files mới
                var newAudioFiles = CreateAudioFileRecordsFromVocabularies(new List<Vocabulary> { existingVocab }, options, targetDate);
                _context.AudioFiles.AddRange(newAudioFiles);
            }
        }

        // THÊM MỚI: Find exact match trong database
        private Vocabulary FindExactMatchInDatabase(Vocabulary currentVocab, List<Vocabulary> existingVocabs, string fileType)
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

        // THÊM MỚI: Find partial match trong database
        private Vocabulary FindPartialMatchInDatabase(Vocabulary currentVocab, List<Vocabulary> existingVocabs, string fileType)
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

        // THÊM MỚI: Helper methods
        private bool SimilarText(string text1, string text2)
        {
            if (string.IsNullOrWhiteSpace(text1) || string.IsNullOrWhiteSpace(text2))
                return false;
            return text1.Trim().Equals(text2.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private string GetPrimaryText(Vocabulary vocab, string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => $"VI:'{vocab.VietnameseText}' / EN:'{vocab.EnglishText}'",
                "japanese" => $"EN:'{vocab.EnglishText}' / JP:'{vocab.JapaneseText}'",
                "chinese" => $"EN:'{vocab.EnglishText}' / ZH:'{vocab.ChineseText}'",
                "tuvung" => $"VI:'{vocab.VietnameseText}' / JP:'{vocab.JapaneseText}'",
                _ => vocab.EnglishText ?? vocab.VietnameseText ?? ""
            };
        }

        private bool HasValidData(Vocabulary vocab, string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => !string.IsNullOrWhiteSpace(vocab.VietnameseText) || !string.IsNullOrWhiteSpace(vocab.EnglishText),
                "japanese" => !string.IsNullOrWhiteSpace(vocab.EnglishText) || !string.IsNullOrWhiteSpace(vocab.JapaneseText),
                "chinese" => !string.IsNullOrWhiteSpace(vocab.EnglishText) || !string.IsNullOrWhiteSpace(vocab.ChineseText),
                "tuvung" => !string.IsNullOrWhiteSpace(vocab.VietnameseText) || !string.IsNullOrWhiteSpace(vocab.JapaneseText),
                _ => true
            };
        }

        private bool CheckRowHasSoundFormulas(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            if (string.IsNullOrWhiteSpace(options.SoundColumns)) return false;

            string fileType = DetermineFileType(options.FileName).ToLower();

            if ((fileType == "japanese" || fileType == "chinese" || fileType == "tuvung") && options.SoundColumns.Length == 2)
            {
                string col1Value = worksheet.Cells[row, 5].Text?.Trim();
                string col2Value = worksheet.Cells[row, 6].Text?.Trim();
                return !string.IsNullOrEmpty(col1Value) && !string.IsNullOrEmpty(col2Value) &&
                       col1Value.Contains("[sound:") && col2Value.Contains("[sound:");
            }
            else if (fileType == "english" && options.SoundColumns.Length >= 2)
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

        private Vocabulary CreateVocabularyFromExcelRow(ExcelWorksheet worksheet, int row, string fileType, string sourceFile, DateTime targetDate)
        {
            var vocab = new Vocabulary
            {
                SourceFile = Path.GetFileName(sourceFile),
                Category = fileType,
                CreatedDate = targetDate,
                ImageTags = ""
            };

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
                    vocab.KanjiText = GetCellValue(worksheet, row, 2);
                    vocab.ImageTags = GetCellValue(worksheet, row, 7) ?? "";
                    break;
                case "chinese":
                    vocab.EnglishText = GetCellValue(worksheet, row, 1);
                    vocab.ReadingText = GetCellValue(worksheet, row, 2);
                    vocab.ChineseText = GetCellValue(worksheet, row, 3);
                    vocab.KanjiText = GetCellValue(worksheet, row, 2);
                    vocab.ImageTags = GetCellValue(worksheet, row, 7) ?? "";
                    break;
                case "tuvung":
                    vocab.VietnameseText = GetCellValue(worksheet, row, 1);
                    vocab.ReadingText = GetCellValue(worksheet, row, 2);
                    vocab.JapaneseText = GetCellValue(worksheet, row, 3);
                    vocab.KanjiText = GetCellValue(worksheet, row, 2);
                    vocab.ImageTags = GetCellValue(worksheet, row, 7) ?? "";
                    break;
            }

            return vocab;
        }

        private List<AudioFile> CreateAudioFileRecordsFromVocabularies(List<Vocabulary> vocabularies, ProcessingOptions options, DateTime targetDate)
        {
            var audioFiles = new List<AudioFile>();
            string dateToUse = options.CustomDate ?? DateTime.Now.ToString("dd-MM-yyyy");

            int existingAudioCount = GetExistingAudioCountForDate(dateToUse, options.AudioFileType);
            int currentAudioNumber = existingAudioCount + 1;

            foreach (var vocab in vocabularies)
            {
                var vocabAudioFiles = CreateAudioFilesForVocabulary(vocab, options, dateToUse, currentAudioNumber, currentAudioNumber + 1, targetDate);
                audioFiles.AddRange(vocabAudioFiles);
                currentAudioNumber += 2;
            }

            return audioFiles;
        }

        // SỬA LẠI DatabaseService.CreateAudioFilesForVocabulary() THEO LOGIC CŨ ĐÚNG
        private List<AudioFile> CreateAudioFilesForVocabulary(Vocabulary vocab, ProcessingOptions options, string dateToUse, int oddNumber, int evenNumber, DateTime targetDate)
        {
            var audioFiles = new List<AudioFile>();

            switch (options.AudioFileType)
            {
                case "VI-EN":
                    // File lẻ: Vietnamese voice + Vietnamese text
                    if (!string.IsNullOrWhiteSpace(vocab.VietnameseText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "VI", $"EN-{dateToUse}_{oddNumber:00}.mp3", "vi-VN-HoaiMyNeural", 1.0m, true, targetDate));
                    // File chẵn: English voice + English text
                    if (!string.IsNullOrWhiteSpace(vocab.EnglishText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "EN", $"EN-{dateToUse}_{evenNumber:00}.mp3", "en-US-JennyNeural", 0.75m, false, targetDate));
                    break;

                case "JP-EN":
                    // File lẻ: English voice + English text
                    if (!string.IsNullOrWhiteSpace(vocab.EnglishText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "EN", $"JP-{dateToUse}_{oddNumber:00}.mp3", "en-US-JennyNeural", 0.75m, true, targetDate));
                    // File chẵn: Japanese voice + Japanese text
                    if (!string.IsNullOrWhiteSpace(vocab.JapaneseText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "JP", $"JP-{dateToUse}_{evenNumber:00}.mp3", "ja-JP-NanamiNeural", 0.7m, false, targetDate));
                    break;

                case "ZH-EN":
                    // File lẻ: English voice + English text
                    if (!string.IsNullOrWhiteSpace(vocab.EnglishText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "EN", $"ZH-{dateToUse}_{oddNumber:00}.mp3", "en-US-JennyNeural", 0.75m, true, targetDate));
                    // File chẵn: Chinese voice + Chinese text
                    if (!string.IsNullOrWhiteSpace(vocab.ChineseText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "ZH", $"ZH-{dateToUse}_{evenNumber:00}.mp3", "zh-CN-XiaoxiaoNeural", 0.7m, false, targetDate));
                    break;

                case "TUVUNG":
                    // File lẻ: English voice + English text (nhưng TuVung không có English text, nên đọc Vietnamese text)
                    if (!string.IsNullOrWhiteSpace(vocab.VietnameseText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "EN", $"Vocab-{dateToUse}_{oddNumber:00}.mp3", "en-US-JennyNeural", 0.75m, true, targetDate));
                    // File chẵn: Vietnamese voice + Vietnamese text
                    if (!string.IsNullOrWhiteSpace(vocab.VietnameseText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "VI", $"Vocab-{dateToUse}_{evenNumber:00}.mp3", "vi-VN-HoaiMyNeural", 1.0m, false, targetDate));
                    break;
            }

            return audioFiles;
        }

        private AudioFile CreateAudioFileRecord(int vocabId, string language, string fileName, string voiceName, decimal speechRate, bool isOddFile, DateTime targetDate)
        {
            return new AudioFile
            {
                VocabId = vocabId,
                Language = language,
                FileName = fileName,
                FilePath = "",
                VoiceName = voiceName,
                SpeechRate = speechRate,
                IsOddFile = isOddFile,
                CreatedDate = targetDate,
                IsGenerated = false
            };
        }

        private string GetCellValue(ExcelWorksheet worksheet, int row, int column)
        {
            var value = worksheet.Cells[row, column].Text?.Trim();
            return string.IsNullOrEmpty(value) ? "" : value;
        }

        private string DetermineFileType(string fileName)
        {
            fileName = fileName.ToLower();
            if (fileName.Contains("english")) return "English";
            if (fileName.Contains("japanese")) return "Japanese";
            if (fileName.Contains("chinese")) return "Chinese";
            if (fileName.Contains("tuvung")) return "TuVung";
            return "Unknown";
        }

        private string GetColumnString(string columnInput)
        {
            return string.IsNullOrWhiteSpace(columnInput) ? "None" : columnInput;
        }

        // FIXED: Method tính audio count chính xác
        private int GetExistingAudioCountForDate(string dateString, string audioFileType)
        {
            try
            {
                string fileNamePattern = GetAudioFileNamePattern(dateString, audioFileType);
                if (string.IsNullOrEmpty(fileNamePattern)) return 0;

                var existingAudioRecords = _context.AudioFiles
                    .Where(a => a.FileName.StartsWith(fileNamePattern))
                    .Select(a => a.FileName)
                    .ToList();

                if (!existingAudioRecords.Any()) return 0;

                int maxNumber = 0;
                foreach (var fileName in existingAudioRecords)
                {
                    var parts = fileName.Split('_');
                    if (parts.Length >= 2)
                    {
                        var numberPart = parts[1].Split('.')[0];
                        if (int.TryParse(numberPart, out int number))
                        {
                            maxNumber = Math.Max(maxNumber, number);
                        }
                    }
                }

                Console.WriteLine($"Audio count for {dateString}: Found {existingAudioRecords.Count} files, max number: {maxNumber}");
                return maxNumber;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not check existing audio records: {ex.Message}");
                return 0;
            }
        }

        private string GetAudioFileNamePattern(string dateString, string audioFileType)
        {
            return audioFileType switch
            {
                "VI-EN" => $"EN-{dateString}_",
                "JP-EN" => $"JP-{dateString}_",
                "ZH-EN" => $"ZH-{dateString}_",
                "TUVUNG" => $"Vocab-{dateString}_",
                _ => ""
            };
        }

        public async Task<List<Vocabulary>> GetExistingVocabulariesForDateAsync(string dateString, string fileType)
        {
            try
            {
                if (!DateTime.TryParseExact(dateString, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime targetDate))
                {
                    return new List<Vocabulary>();
                }

                var existingVocabs = await _context.Vocabularies
                    .Where(v => v.CreatedDate.Date == targetDate.Date && v.Category == fileType)
                    .ToListAsync();

                Console.WriteLine($"Found {existingVocabs.Count} existing vocabularies for {dateString} ({fileType})");
                return existingVocabs;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not get existing vocabularies: {ex.Message}");
                return new List<Vocabulary>();
            }
        }

        public async Task<ExistingDataSummary> GetExistingDataSummary(string dateString, string fileType, string audioFileType)
        {
            var summary = new ExistingDataSummary();

            try
            {
                if (!DateTime.TryParseExact(dateString, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime targetDate))
                {
                    return summary;
                }

                summary.VocabularyCount = await _context.Vocabularies
                    .Where(v => v.CreatedDate.Date == targetDate.Date && v.Category == fileType)
                    .CountAsync();

                summary.AudioFileCount = GetExistingAudioCountForDate(dateString, audioFileType);
                summary.NextAudioNumber = summary.AudioFileCount + 1;

                return summary;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting existing data summary: {ex.Message}");
                return summary;
            }
        }
        public async Task<List<AudioFile>> GetAudioFilesByVocabIdAsync(int vocabId)
        {
            try
            {
                return await _context.AudioFiles
                    .Where(a => a.VocabId == vocabId)
                    .OrderBy(a => a.IsOddFile ? 0 : 1) // Odd file trước, Even file sau
                    .ToListAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting audio files for VocabId {vocabId}: {ex.Message}");
                return new List<AudioFile>();
            }
        }

        public void Dispose()
        {
            _context?.Dispose();
        }

        public class ExistingDataSummary
        {
            public int VocabularyCount { get; set; } = 0;
            public int AudioFileCount { get; set; } = 0;
            public int NextAudioNumber { get; set; } = 1;
        }

        public class ProcessingResult
        {
            public List<Vocabulary> NewVocabularies { get; set; } = new List<Vocabulary>();
            public int SkippedCount { get; set; } = 0;
            public int TotalProcessed { get; set; } = 0;
        }
        public class DuplicateCheckResult
        {
            public bool IsDuplicate { get; set; }
            public Vocabulary ExistingVocab { get; set; }
            public string Action { get; set; } = ""; // "Skip", "Update", "Insert"
        }
    }
}