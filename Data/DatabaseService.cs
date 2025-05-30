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
                Console.WriteLine("Saving Excel data to SQL Server...");

                string dateToUse = options.CustomDate ?? DateTime.Now.ToString("dd-MM-yyyy");
                DateTime targetDate = DateTime.ParseExact(dateToUse, "dd-MM-yyyy", null);

                var session = new ProcessingSession
                {
                    FileName = Path.GetFileName(options.SelectedFile),
                    FileType = DetermineFileType(options.FileName),
                    ProcessedDate = targetDate,
                    DateUsed = dateToUse,
                    Notes = $"Processed with columns: IMG({GetColumnString(options.ColumnInput)}), Sound({options.SoundColumns}), Kanji({options.KanjiColumn})"
                };

                _context.ProcessingSessions.Add(session);
                await _context.SaveChangesAsync();

                var result = await ProcessExcelWithSoundFormulasOnly(options, session.SessionId, targetDate);

                if (result.NewVocabularies.Any())
                {
                    _context.Vocabularies.AddRange(result.NewVocabularies);
                    await _context.SaveChangesAsync();

                    Console.WriteLine($"Saved {result.NewVocabularies.Count} NEW vocabulary entries to database.");
                    Console.WriteLine($"Skipped {result.SkippedCount} duplicate entries.");

                    if (options.CreateAudioFiles && !string.IsNullOrWhiteSpace(options.SoundColumns))
                    {
                        var audioFiles = CreateAudioFileRecordsFromVocabularies(result.NewVocabularies, options, targetDate);
                        if (audioFiles.Any())
                        {
                            _context.AudioFiles.AddRange(audioFiles);
                            await _context.SaveChangesAsync();
                            session.AudioFilesCreated = audioFiles.Count;
                            Console.WriteLine($"Created {audioFiles.Count} audio file records.");
                        }
                    }
                }

                session.ProcessedRows = result.NewVocabularies.Count;
                session.TotalRows = result.TotalProcessed;
                session.IsCompleted = true;
                await _context.SaveChangesAsync();

                Console.WriteLine("Excel data successfully saved to SQL Server!");
                return session;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving to database: {ex.Message}");
                throw;
            }
        }

        private async Task<ProcessingResult> ProcessExcelWithSoundFormulasOnly(ProcessingOptions options, int sessionId, DateTime targetDate)
        {
            var result = new ProcessingResult();
            string dateToUse = options.CustomDate ?? DateTime.Now.ToString("dd-MM-yyyy");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(options.SelectedFile));
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension?.End.Row ?? 0;

            if (rowCount == 0) return result;

            string fileType = DetermineFileType(options.FileName);

            for (int row = 1; row <= rowCount; row++)
            {
                bool hasSoundFormulas = CheckRowHasSoundFormulas(worksheet, row, options);

                if (!hasSoundFormulas)
                {
                    result.SkippedCount++;
                    result.TotalProcessed++;
                    continue;
                }

                var vocab = CreateVocabularyFromExcelRow(worksheet, row, fileType, options.SelectedFile, targetDate);
                result.NewVocabularies.Add(vocab);
                result.TotalProcessed++;
            }

            return result;
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
    }
}