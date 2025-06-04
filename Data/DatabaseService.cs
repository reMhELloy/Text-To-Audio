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

        public async Task<ProcessingSession> SaveExcelDataToDatabaseAsync(ProcessingOptions options)
        {
            try
            {
                string dateToUse = options.CustomDate ?? DateTime.Now.ToString("dd-MM-yyyy");
                DateTime targetDate = DateTime.ParseExact(dateToUse, "dd-MM-yyyy", null);
                var fileType = DetermineFileType(options.FileName);

                // CREATE SESSION
                var session = new ProcessingSession
                {
                    FileName = Path.GetFileName(options.SelectedFile),
                    FileType = fileType,
                    ProcessedDate = targetDate,
                    DateUsed = dateToUse,
                    Notes = $"Processed with columns: IMG({options.ColumnInput ?? "None"}), Sound({options.SoundColumns}), Kanji({options.KanjiColumn})"
                };

                _context.ProcessingSessions.Add(session);
                await _context.SaveChangesAsync();

                // Load data concurrently
                var excelTask = ProcessExcelWithUpsertLogic(options, session.SessionId, targetDate);
                var existingTask = GetExistingVocabulariesForDateAsync(dateToUse, fileType);

                var allVocabsFromExcel = await excelTask;
                var existingVocabs = await existingTask;

                int skipCount = 0;
                var newVocabularies = new List<Vocabulary>();
                var updatedVocabularies = new List<Vocabulary>();

                foreach (var currentVocab in allVocabsFromExcel)
                {
                    var duplicateResult = CheckVocabularyForDuplicate(currentVocab, existingVocabs, session.FileType);

                    switch (duplicateResult.Action)
                    {
                        case "Skip":
                            skipCount++;
                            break;

                        case "Update":
                            await UpdateExistingVocabulary(duplicateResult.ExistingVocab, currentVocab, options, dateToUse, targetDate);
                            updatedVocabularies.Add(duplicateResult.ExistingVocab);
                            break;

                        case "Insert":
                            newVocabularies.Add(currentVocab);
                            break;
                    }
                }

                // Batch save all vocabulary changes
                if (newVocabularies.Any())
                {
                    _context.Vocabularies.AddRange(newVocabularies);
                }

                // Single save for all vocabulary operations
                if (newVocabularies.Any() || updatedVocabularies.Any())
                {
                    await _context.SaveChangesAsync();
                }

                // ✅ FIX: CENTRALIZED AUDIO PROCESSING (chỉ 1 lần, xử lý đúng)
                if (options.CreateAudioFiles && !string.IsNullOrWhiteSpace(options.SoundColumns))
                {
                    Console.WriteLine($"=== CENTRALIZED AUDIO PROCESSING ===");
                    Console.WriteLine($"New vocabularies: {newVocabularies.Count}");
                    Console.WriteLine($"Updated vocabularies: {updatedVocabularies.Count}");

                    int audioFilesCreated = 0;

                    // STEP 1: Handle UPDATED vocabularies (update existing audio files)
                    foreach (var updatedVocab in updatedVocabularies)
                    {
                        await UpdateExistingAudioFiles(updatedVocab, options, targetDate, dateToUse);
                        audioFilesCreated += 2;
                    }

                    // STEP 2: Handle NEW vocabularies (create new audio files)
                    if (newVocabularies.Any())
                    {
                        var audioFiles = CreateAudioFileRecordsFromVocabularies(newVocabularies, options, targetDate);
                        if (audioFiles.Any())
                        {
                            _context.AudioFiles.AddRange(audioFiles);
                            audioFilesCreated += audioFiles.Count;
                        }
                    }

                    session.AudioFilesCreated = audioFilesCreated;
                    Console.WriteLine($"Total audio files processed: {audioFilesCreated}");
                }

                // Final session update and save
                session.ProcessedRows = newVocabularies.Count;
                session.UpdatedRows = updatedVocabularies.Count;
                session.TotalRows = allVocabsFromExcel.Count;
                session.IsCompleted = true;
                await _context.SaveChangesAsync();

                return session;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Database save error: {ex.Message}");
                throw;
            }
        }
        private async Task UpdateExistingAudioFiles(Vocabulary existingVocab, ProcessingOptions options, DateTime targetDate, string dateToUse)
        {
            try
            {
                Console.WriteLine($"Updating audio files for VocabId {existingVocab.VocabId}");

                // GET EXISTING AUDIO FILES
                var existingAudioFiles = await _context.AudioFiles
                    .Where(a => a.VocabId == existingVocab.VocabId)
                    .OrderBy(a => a.IsOddFile ? 0 : 1)
                    .ToListAsync();

                Console.WriteLine($"Found {existingAudioFiles.Count} existing audio files");

                if (existingAudioFiles.Count >= 2)
                {
                    // ✅ UPDATE EXISTING AUDIO FILES (keep filename, update metadata)
                    var oddFile = existingAudioFiles.FirstOrDefault(a => a.IsOddFile);
                    var evenFile = existingAudioFiles.FirstOrDefault(a => !a.IsOddFile);

                    if (oddFile != null)
                    {
                        oddFile.CreatedDate = targetDate;
                        oddFile.IsGenerated = false; // Mark for regeneration
                        _context.Entry(oddFile).State = EntityState.Modified;
                        Console.WriteLine($"✅ Updated audio file: {oddFile.FileName}");
                    }

                    if (evenFile != null)
                    {
                        evenFile.CreatedDate = targetDate;
                        evenFile.IsGenerated = false;
                        _context.Entry(evenFile).State = EntityState.Modified;
                        Console.WriteLine($"✅ Updated audio file: {evenFile.FileName}");
                    }

                    // ✅ REMOVE DUPLICATE AUDIO FILES (if more than 2)
                    if (existingAudioFiles.Count > 2)
                    {
                        var extraFiles = existingAudioFiles.Skip(2).ToList();
                        _context.AudioFiles.RemoveRange(extraFiles);
                        Console.WriteLine($"🗑️ Removed {extraFiles.Count} duplicate audio files");
                    }
                }
                else
                {
                    Console.WriteLine($"⚠️ VocabId {existingVocab.VocabId} has insufficient audio files ({existingAudioFiles.Count}), will create new ones");

                    // Delete existing incomplete set
                    if (existingAudioFiles.Any())
                    {
                        _context.AudioFiles.RemoveRange(existingAudioFiles);
                        Console.WriteLine($"🗑️ Removed {existingAudioFiles.Count} incomplete audio files");
                    }

                    // Create new audio files with proper numbering
                    int nextAudioNumber = GetNextAvailableAudioNumber(dateToUse, options.AudioFileType);
                    var newAudioFiles = CreateAudioFilesForVocabulary(existingVocab, options, dateToUse, nextAudioNumber, nextAudioNumber + 1, targetDate);
                    _context.AudioFiles.AddRange(newAudioFiles);
                    Console.WriteLine($"✅ Created new audio files: {nextAudioNumber:00}, {nextAudioNumber + 1:00}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error updating audio files for VocabId {existingVocab.VocabId}: {ex.Message}");
                throw;
            }
        }
        private int GetNextAvailableAudioNumber(string dateString, string audioFileType)
        {
            try
            {
                string fileNamePattern = GetAudioFileNamePattern(dateString, audioFileType);
                if (string.IsNullOrEmpty(fileNamePattern)) return 1;

                var existingNumbers = _context.AudioFiles
                    .Where(a => a.FileName.StartsWith(fileNamePattern))
                    .Select(a => a.FileName)
                    .ToList()
                    .Select(fileName => {
                        var parts = fileName.Split('_');
                        if (parts.Length >= 2)
                        {
                            var numberPart = parts[1].Split('.')[0];
                            if (int.TryParse(numberPart, out int number))
                                return number;
                        }
                        return 0;
                    })
                    .Where(n => n > 0)
                    .OrderBy(n => n)
                    .ToList();

                Console.WriteLine($"Existing audio numbers for {dateString}: [{string.Join(", ", existingNumbers)}]");

                // FIND FIRST AVAILABLE ODD NUMBER
                for (int i = 1; i <= existingNumbers.Count + 2; i += 2)
                {
                    if (!existingNumbers.Contains(i) && !existingNumbers.Contains(i + 1))
                    {
                        Console.WriteLine($"Next available audio number pair: {i}, {i + 1}");
                        return i;
                    }
                }

                // NO GAPS, USE NEXT ODD NUMBER
                int nextNumber = existingNumbers.Any() ? existingNumbers.Max() + 1 : 1;
                if (nextNumber % 2 == 0) nextNumber++; // Ensure odd number
                Console.WriteLine($"No gaps, next audio number: {nextNumber}");
                return nextNumber;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting next audio number: {ex.Message}");
                return 1;
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
            Console.WriteLine($"Updated vocabulary data for VocabId {existingVocab.VocabId}");
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