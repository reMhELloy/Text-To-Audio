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

            // Ensure database is created
            _context.Database.EnsureCreated();
        }

        public async Task<ProcessingSession> SaveExcelDataToDatabaseAsync(ProcessingOptions options)
        {
            try
            {
                Console.WriteLine("💾 Saving Excel data to SQL Server...");

                // Tạo processing session
                var session = new ProcessingSession
                {
                    FileName = Path.GetFileName(options.SelectedFile),
                    FileType = DetermineFileType(options.FileName),
                    ProcessedDate = DateTime.UtcNow,
                    DateUsed = options.CustomDate ?? DateTime.Now.ToString("dd-MM-yyyy"),
                    Notes = $"Processed with columns: IMG({GetColumnString(options.ColumnInput)}), Sound({options.SoundColumns}), Kanji({options.KanjiColumn})"
                };

                _context.ProcessingSessions.Add(session);
                await _context.SaveChangesAsync();

                // Đọc và xử lý dữ liệu từ Excel với duplicate checking
                var result = await ProcessExcelWithDuplicateChecking(options, session.SessionId);

                if (result.NewVocabularies.Any())
                {
                    _context.Vocabularies.AddRange(result.NewVocabularies);
                    await _context.SaveChangesAsync();

                    Console.WriteLine($"✅ Saved {result.NewVocabularies.Count} NEW vocabulary entries to database.");
                    Console.WriteLine($"⚠️ Skipped {result.SkippedCount} duplicate entries.");

                    // Tạo audio file records và update Excel
                    if (options.CreateAudioFiles && !string.IsNullOrWhiteSpace(options.SoundColumns))
                    {
                        var audioFiles = CreateAudioFileRecordsWithExcelUpdate(result.VocabularyMappings, options);
                        if (audioFiles.Any())
                        {
                            _context.AudioFiles.AddRange(audioFiles);
                            await _context.SaveChangesAsync();

                            session.AudioFilesCreated = audioFiles.Count;
                            Console.WriteLine($"✅ Created {audioFiles.Count} audio file records.");
                        }
                    }
                }

                // Cập nhật session
                session.ProcessedRows = result.NewVocabularies.Count;
                session.TotalRows = result.TotalProcessed;
                session.IsCompleted = true;

                await _context.SaveChangesAsync();

                Console.WriteLine("🎉 Excel data successfully saved to SQL Server!");
                return session;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error saving to database: {ex.Message}");
                throw;
            }
        }

        // 🔍 METHOD MỚI: Xử lý Excel với duplicate checking - CHỈ TRONG CÙNG NGÀY
        private async Task<ProcessingResult> ProcessExcelWithDuplicateChecking(ProcessingOptions options, int sessionId)
        {
            var result = new ProcessingResult();
            string dateToUse = options.CustomDate ?? DateTime.Now.ToString("dd-MM-yyyy");

            // Lấy existing vocabulary cho CÙNG NGÀY và CÙNG FILE TYPE
            var existingVocabsToday = await GetExistingVocabulariesForDate(dateToUse, DetermineFileType(options.FileName));

            Console.WriteLine($"📅 Processing date: {dateToUse}");
            Console.WriteLine($"📂 File type: {DetermineFileType(options.FileName)}");
            Console.WriteLine($"🔍 Existing vocabularies TODAY: {existingVocabsToday.Count}");
            Console.WriteLine($"ℹ️  Note: Only checking duplicates within the SAME DAY");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(options.SelectedFile));
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension?.End.Row ?? 0;

            if (rowCount == 0) return result;

            string fileType = DetermineFileType(options.FileName);

            for (int row = 1; row <= rowCount; row++)
            {
                var vocab = CreateVocabularyFromExcelRow(worksheet, row, fileType, options.SelectedFile);

                // Kiểm tra duplicate - CHỈ với data của cùng ngày
                bool isDuplicate = CheckIfDuplicate(vocab, existingVocabsToday);

                if (isDuplicate)
                {
                    Console.WriteLine($"⚠️ Row {row}: Duplicate found in TODAY's data, skipping...");
                    result.VocabularyMappings.Add(new VocabularyMapping
                    {
                        ExcelRow = row,
                        Vocabulary = null,
                        IsSkipped = true
                    });
                    result.SkippedCount++;
                }
                else
                {
                    result.NewVocabularies.Add(vocab);
                    result.VocabularyMappings.Add(new VocabularyMapping
                    {
                        ExcelRow = row,
                        Vocabulary = vocab,
                        IsSkipped = false
                    });
                    Console.WriteLine($"✅ Row {row}: New vocabulary (not duplicate in today's data)");
                }

                result.TotalProcessed++;
            }

            Console.WriteLine($"\n📊 Summary:");
            Console.WriteLine($"   📝 Total rows processed: {result.TotalProcessed}");
            Console.WriteLine($"   ✅ New vocabularies: {result.NewVocabularies.Count}");
            Console.WriteLine($"   ⚠️ Duplicates (same day): {result.SkippedCount}");

            return result;
        }

        // 🔍 METHOD MỚI: Tạo audio records và update Excel sound formulas
        private List<AudioFile> CreateAudioFileRecordsWithExcelUpdate(List<VocabularyMapping> mappings, ProcessingOptions options)
        {
            var audioFiles = new List<AudioFile>();
            string dateToUse = options.CustomDate ?? DateTime.Now.ToString("dd-MM-yyyy");

            // Lấy existing audio count
            int existingAudioCount = GetExistingAudioCountForDate(dateToUse, options.AudioFileType);
            Console.WriteLine($"🔍 Found {existingAudioCount} existing audio files for date {dateToUse}");

            int currentAudioNumber = existingAudioCount + 1;

            // Mở Excel để update sound formulas
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(options.SelectedFile));
            var worksheet = package.Workbook.Worksheets[0];

            foreach (var mapping in mappings)
            {
                if (mapping.IsSkipped)
                {
                    // Nếu skip thì clear sound formulas trong Excel
                    ClearSoundFormulasInExcel(worksheet, mapping.ExcelRow, options);
                    Console.WriteLine($"🔄 Row {mapping.ExcelRow}: Cleared sound formulas (duplicate)");
                    continue;
                }

                // Tạo audio files cho vocabulary mới
                var vocab = mapping.Vocabulary;
                int oddNumber = currentAudioNumber;
                int evenNumber = currentAudioNumber + 1;

                // Update sound formulas trong Excel
                UpdateSoundFormulasInExcel(worksheet, mapping.ExcelRow, options, dateToUse, oddNumber, evenNumber);

                // Tạo audio file records
                var newAudioFiles = CreateAudioFilesForVocabulary(vocab, options, dateToUse, oddNumber, evenNumber);
                audioFiles.AddRange(newAudioFiles);

                currentAudioNumber += 2; // Mỗi vocabulary tạo 2 files (odd + even)
                Console.WriteLine($"🔄 Row {mapping.ExcelRow}: Updated sound formulas to {oddNumber:00}, {evenNumber:00}");
            }

            // Lưu Excel file với sound formulas đã update
            try
            {
                package.Save();
                Console.WriteLine("✅ Excel sound formulas updated successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Warning: Could not update Excel file: {ex.Message}");
            }

            return audioFiles;
        }

        // 🔍 METHOD MỚI: Update sound formulas trong Excel
        private void UpdateSoundFormulasInExcel(ExcelWorksheet worksheet, int row, ProcessingOptions options,
            string dateToUse, int oddNumber, int evenNumber)
        {
            if (string.IsNullOrWhiteSpace(options.SoundColumns) || options.SoundColumns.Length < 2) return;

            string fileType = DetermineFileType(options.FileName);

            switch (fileType.ToLower())
            {
                case "japanese":
                    // E (odd-EN), F (even-JP)
                    worksheet.Cells[row, 5].Value = $"[sound:EN-{dateToUse}_{oddNumber:00}.mp3]";
                    worksheet.Cells[row, 6].Value = $"[sound:JP-{dateToUse}_{evenNumber:00}.mp3]";
                    break;

                case "chinese":
                    // E (odd-EN), F (even-ZH)
                    worksheet.Cells[row, 5].Value = $"[sound:EN-{dateToUse}_{oddNumber:00}.mp3]";
                    worksheet.Cells[row, 6].Value = $"[sound:ZH-{dateToUse}_{evenNumber:00}.mp3]";
                    break;

                case "tuvung":
                    // E (odd-EN), F (even-VI)
                    worksheet.Cells[row, 5].Value = $"[sound:EN-{dateToUse}_{oddNumber:00}.mp3]";
                    worksheet.Cells[row, 6].Value = $"[sound:VI-{dateToUse}_{evenNumber:00}.mp3]";
                    break;

                case "english":
                    // D, E columns
                    if (options.SoundColumns.Length >= 2)
                    {
                        int col1 = options.SoundColumns[0] - 'A' + 1;
                        int col2 = options.SoundColumns[1] - 'A' + 1;
                        worksheet.Cells[row, col1].Value = $"[sound:EN-{dateToUse}_{oddNumber:00}.mp3]";
                        worksheet.Cells[row, col2].Value = $"[sound:EN-{dateToUse}_{evenNumber:00}.mp3]";
                    }
                    break;
            }
        }

        // 🔍 METHOD MỚI: Clear sound formulas trong Excel cho duplicate rows
        private void ClearSoundFormulasInExcel(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            if (string.IsNullOrWhiteSpace(options.SoundColumns)) return;

            string fileType = DetermineFileType(options.FileName);

            switch (fileType.ToLower())
            {
                case "japanese":
                case "chinese":
                case "tuvung":
                    worksheet.Cells[row, 5].Value = ""; // Column E
                    worksheet.Cells[row, 6].Value = ""; // Column F
                    break;

                case "english":
                    if (options.SoundColumns.Length >= 2)
                    {
                        int col1 = options.SoundColumns[0] - 'A' + 1;
                        int col2 = options.SoundColumns[1] - 'A' + 1;
                        worksheet.Cells[row, col1].Value = "";
                        worksheet.Cells[row, col2].Value = "";
                    }
                    break;
            }
        }

        // 🔍 METHOD MỚI: Tạo audio files cho 1 vocabulary
        private List<AudioFile> CreateAudioFilesForVocabulary(Vocabulary vocab, ProcessingOptions options,
            string dateToUse, int oddNumber, int evenNumber)
        {
            var audioFiles = new List<AudioFile>();

            switch (options.AudioFileType)
            {
                case "VI-EN":
                    if (!string.IsNullOrWhiteSpace(vocab.VietnameseText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "VI",
                            $"EN-{dateToUse}_{oddNumber:00}.mp3", "vi-VN-HoaiMyNeural", 1.0m, true));
                    if (!string.IsNullOrWhiteSpace(vocab.EnglishText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "EN",
                            $"EN-{dateToUse}_{evenNumber:00}.mp3", "en-US-JennyNeural", 0.75m, false));
                    break;

                case "JP-EN":
                    if (!string.IsNullOrWhiteSpace(vocab.EnglishText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "EN",
                            $"JP-{dateToUse}_{oddNumber:00}.mp3", "en-US-JennyNeural", 0.75m, true));
                    if (!string.IsNullOrWhiteSpace(vocab.JapaneseText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "JP",
                            $"JP-{dateToUse}_{evenNumber:00}.mp3", "ja-JP-NanamiNeural", 0.7m, false));
                    break;

                case "ZH-EN":
                    if (!string.IsNullOrWhiteSpace(vocab.EnglishText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "EN",
                            $"ZH-{dateToUse}_{oddNumber:00}.mp3", "en-US-JennyNeural", 0.75m, true));
                    if (!string.IsNullOrWhiteSpace(vocab.ChineseText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "ZH",
                            $"ZH-{dateToUse}_{evenNumber:00}.mp3", "zh-CN-XiaoxiaoNeural", 0.7m, false));
                    break;

                case "TUVUNG":
                    if (!string.IsNullOrWhiteSpace(vocab.VietnameseText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "VI",
                            $"Vocab-{dateToUse}_{oddNumber:00}.mp3", "vi-VN-HoaiMyNeural", 1.0m, true));
                    if (!string.IsNullOrWhiteSpace(vocab.JapaneseText))
                        audioFiles.Add(CreateAudioFileRecord(vocab.VocabId, "JP",
                            $"Vocab-{dateToUse}_{evenNumber:00}.mp3", "ja-JP-NanamiNeural", 0.7m, false));
                    break;
            }

            return audioFiles;
        }

        // 🔍 METHOD MỚI: Lấy existing vocabularies cho CÙNG NGÀY cụ thể
        private async Task<List<Vocabulary>> GetExistingVocabulariesForDate(string dateString, string fileType)
        {
            if (!DateTime.TryParseExact(dateString, "dd-MM-yyyy", null,
                System.Globalization.DateTimeStyles.None, out DateTime targetDate))
            {
                Console.WriteLine($"⚠️ Invalid date format: {dateString}, no duplicate checking");
                return new List<Vocabulary>();
            }

            var existingVocabs = await _context.Vocabularies
                .Where(v => v.CreatedDate.Date == targetDate.Date && v.Category == fileType)
                .ToListAsync();

            Console.WriteLine($"🔍 Checking duplicates for: {targetDate:dd-MM-yyyy} ({fileType})");
            Console.WriteLine($"🔍 Found {existingVocabs.Count} existing vocabularies for TODAY ONLY");

            return existingVocabs;
        }

        // 🔍 METHOD MỚI: Kiểm tra duplicate - CHỈ TRONG CÙNG NGÀY
        private bool CheckIfDuplicate(Vocabulary newVocab, List<Vocabulary> existingVocabsToday)
        {
            // CHỈ kiểm tra với vocabularies của cùng ngày
            bool isDuplicate = existingVocabsToday.Any(existing =>
                (SimilarText(existing.EnglishText, newVocab.EnglishText) ||
                 SimilarText(existing.JapaneseText, newVocab.JapaneseText) ||
                 SimilarText(existing.ChineseText, newVocab.ChineseText) ||
                 SimilarText(existing.VietnameseText, newVocab.VietnameseText))
            );

            if (isDuplicate)
            {
                Console.WriteLine($"   🔍 Duplicate found in TODAY's data");
            }
            else
            {
                Console.WriteLine($"   ✅ New vocabulary (not duplicate in today's data)");
            }

            return isDuplicate;
        }

        // 🔍 METHOD MỚI: So sánh text tương tự
        private bool SimilarText(string text1, string text2)
        {
            if (string.IsNullOrWhiteSpace(text1) || string.IsNullOrWhiteSpace(text2))
                return false;

            return text1.Trim().Equals(text2.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        // 🔍 METHOD MỚI: Tạo vocabulary từ Excel row
        private Vocabulary CreateVocabularyFromExcelRow(ExcelWorksheet worksheet, int row, string fileType, string sourceFile)
        {
            var vocab = new Vocabulary
            {
                SourceFile = Path.GetFileName(sourceFile),
                Category = fileType,
                CreatedDate = DateTime.UtcNow,
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
                    vocab.ImageTags = GetCellValue(worksheet, row, 7);
                    break;
                case "chinese":
                    vocab.EnglishText = GetCellValue(worksheet, row, 1);
                    vocab.ReadingText = GetCellValue(worksheet, row, 2);
                    vocab.ChineseText = GetCellValue(worksheet, row, 3);
                    vocab.KanjiText = GetCellValue(worksheet, row, 2);
                    vocab.ImageTags = GetCellValue(worksheet, row, 7);
                    break;
                case "tuvung":
                    vocab.VietnameseText = GetCellValue(worksheet, row, 1);
                    vocab.ReadingText = GetCellValue(worksheet, row, 2);
                    vocab.JapaneseText = GetCellValue(worksheet, row, 3);
                    vocab.KanjiText = GetCellValue(worksheet, row, 2);
                    vocab.ImageTags = GetCellValue(worksheet, row, 7);
                    break;
            }

            return vocab;
        }

        private async Task<List<Vocabulary>> ReadExcelAndCreateVocabularies(ProcessingOptions options, int sessionId)
        {
            var vocabularies = new List<Vocabulary>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(options.SelectedFile));
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension?.End.Row ?? 0;

            if (rowCount == 0) return vocabularies;

            string fileType = DetermineFileType(options.FileName);

            for (int row = 1; row <= rowCount; row++)
            {
                var vocab = new Vocabulary
                {
                    SourceFile = Path.GetFileName(options.SelectedFile),
                    Category = fileType,
                    CreatedDate = DateTime.UtcNow,
                    ImageTags = "" // Set default empty string instead of null
                };

                // Đọc data theo loại file
                switch (fileType.ToLower())
                {
                    case "english":
                        vocab.VietnameseText = GetCellValue(worksheet, row, 1); // Column A
                        vocab.EnglishText = GetCellValue(worksheet, row, 2);    // Column B
                        break;

                    case "japanese":
                        vocab.EnglishText = GetCellValue(worksheet, row, 1);    // Column A
                        vocab.ReadingText = GetCellValue(worksheet, row, 2);    // Column B
                        vocab.JapaneseText = GetCellValue(worksheet, row, 3);   // Column C
                        vocab.KanjiText = GetCellValue(worksheet, row, 2);      // Column B (processed)
                        vocab.ImageTags = GetCellValue(worksheet, row, 7) ?? ""; // Column G - ensure not null
                        break;

                    case "chinese":
                        vocab.EnglishText = GetCellValue(worksheet, row, 1);    // Column A
                        vocab.ReadingText = GetCellValue(worksheet, row, 2);    // Column B
                        vocab.ChineseText = GetCellValue(worksheet, row, 3);    // Column C
                        vocab.KanjiText = GetCellValue(worksheet, row, 2);      // Column B (processed)
                        vocab.ImageTags = GetCellValue(worksheet, row, 7) ?? ""; // Column G - ensure not null
                        break;

                    case "tuvung":
                        vocab.VietnameseText = GetCellValue(worksheet, row, 1); // Column A
                        vocab.ReadingText = GetCellValue(worksheet, row, 2);    // Column B (processed Kanji)
                        vocab.JapaneseText = GetCellValue(worksheet, row, 3);   // Column C
                        vocab.KanjiText = GetCellValue(worksheet, row, 2);      // Column B (processed)
                        vocab.ImageTags = GetCellValue(worksheet, row, 7) ?? ""; // Column G - ensure not null
                        break;
                }

                // Chỉ thêm nếu có ít nhất 1 field có data
                if (!string.IsNullOrWhiteSpace(vocab.EnglishText) ||
                    !string.IsNullOrWhiteSpace(vocab.VietnameseText) ||
                    !string.IsNullOrWhiteSpace(vocab.JapaneseText) ||
                    !string.IsNullOrWhiteSpace(vocab.ChineseText))
                {
                    vocabularies.Add(vocab);
                }
            }

            return vocabularies;
        }

        private List<AudioFile> CreateAudioFileRecords(List<Vocabulary> vocabularies, ProcessingOptions options)
        {
            var audioFiles = new List<AudioFile>();
            string dateToUse = options.CustomDate ?? DateTime.Now.ToString("dd-MM-yyyy");

            Console.WriteLine($"🔍 Debug AudioFile creation:");
            Console.WriteLine($"   AudioFileType: {options.AudioFileType}");
            Console.WriteLine($"   CreateAudioFiles: {options.CreateAudioFiles}");
            Console.WriteLine($"   SoundColumns: {options.SoundColumns}");
            Console.WriteLine($"   Vocabularies count: {vocabularies.Count}");

            foreach (var vocab in vocabularies.Select((v, i) => new { Vocab = v, Index = i + 1 }))
            {
                Console.WriteLine($"🔍 Processing vocab {vocab.Index} (ID: {vocab.Vocab.VocabId}):");

                switch (options.AudioFileType)
                {
                    case "VI-EN":
                        Console.WriteLine($"   VI-EN mode");
                        // Vietnamese (lẻ) + English (chẵn)
                        if (!string.IsNullOrWhiteSpace(vocab.Vocab.VietnameseText))
                        {
                            var viAudio = CreateAudioFileRecord(vocab.Vocab.VocabId, "VI",
                                $"EN-{dateToUse}_{(vocab.Index * 2 - 1):00}.mp3",
                                "vi-VN-HoaiMyNeural", 1.0m, true);
                            audioFiles.Add(viAudio);
                            Console.WriteLine($"   ✅ Added VI audio: {viAudio.FileName}");
                        }
                        if (!string.IsNullOrWhiteSpace(vocab.Vocab.EnglishText))
                        {
                            var enAudio = CreateAudioFileRecord(vocab.Vocab.VocabId, "EN",
                                $"EN-{dateToUse}_{(vocab.Index * 2):00}.mp3",
                                "en-US-JennyNeural", 0.75m, false);
                            audioFiles.Add(enAudio);
                            Console.WriteLine($"   ✅ Added EN audio: {enAudio.FileName}");
                        }
                        break;

                    case "JP-EN":
                        Console.WriteLine($"   JP-EN mode");
                        // English (lẻ) + Japanese (chẵn)
                        if (!string.IsNullOrWhiteSpace(vocab.Vocab.EnglishText))
                        {
                            var enAudio = CreateAudioFileRecord(vocab.Vocab.VocabId, "EN",
                                $"JP-{dateToUse}_{(vocab.Index * 2 - 1):00}.mp3",
                                "en-US-JennyNeural", 0.75m, true);
                            audioFiles.Add(enAudio);
                            Console.WriteLine($"   ✅ Added EN audio: {enAudio.FileName}");
                        }
                        if (!string.IsNullOrWhiteSpace(vocab.Vocab.JapaneseText))
                        {
                            var jpAudio = CreateAudioFileRecord(vocab.Vocab.VocabId, "JP",
                                $"JP-{dateToUse}_{(vocab.Index * 2):00}.mp3",
                                "ja-JP-NanamiNeural", 0.7m, false);
                            audioFiles.Add(jpAudio);
                            Console.WriteLine($"   ✅ Added JP audio: {jpAudio.FileName}");
                        }
                        break;

                    case "ZH-EN":
                        Console.WriteLine($"   ZH-EN mode");
                        // English (lẻ) + Chinese (chẵn)
                        if (!string.IsNullOrWhiteSpace(vocab.Vocab.EnglishText))
                        {
                            var enAudio = CreateAudioFileRecord(vocab.Vocab.VocabId, "EN",
                                $"ZH-{dateToUse}_{(vocab.Index * 2 - 1):00}.mp3",
                                "en-US-JennyNeural", 0.75m, true);
                            audioFiles.Add(enAudio);
                            Console.WriteLine($"   ✅ Added EN audio: {enAudio.FileName}");
                        }
                        if (!string.IsNullOrWhiteSpace(vocab.Vocab.ChineseText))
                        {
                            var zhAudio = CreateAudioFileRecord(vocab.Vocab.VocabId, "ZH",
                                $"ZH-{dateToUse}_{(vocab.Index * 2):00}.mp3",
                                "zh-CN-XiaoxiaoNeural", 0.7m, false);
                            audioFiles.Add(zhAudio);
                            Console.WriteLine($"   ✅ Added ZH audio: {zhAudio.FileName}");
                        }
                        break;

                    case "TUVUNG":
                        Console.WriteLine($"   TUVUNG mode");
                        // Vietnamese (lẻ) + Japanese (chẵn)
                        if (!string.IsNullOrWhiteSpace(vocab.Vocab.VietnameseText))
                        {
                            var viAudio = CreateAudioFileRecord(vocab.Vocab.VocabId, "VI",
                                $"Vocab-{dateToUse}_{(vocab.Index * 2 - 1):00}.mp3",
                                "vi-VN-HoaiMyNeural", 1.0m, true);
                            audioFiles.Add(viAudio);
                            Console.WriteLine($"   ✅ Added VI audio: {viAudio.FileName}");
                        }
                        if (!string.IsNullOrWhiteSpace(vocab.Vocab.JapaneseText))
                        {
                            var jpAudio = CreateAudioFileRecord(vocab.Vocab.VocabId, "JP",
                                $"Vocab-{dateToUse}_{(vocab.Index * 2):00}.mp3",
                                "ja-JP-NanamiNeural", 0.7m, false);
                            audioFiles.Add(jpAudio);
                            Console.WriteLine($"   ✅ Added JP audio: {jpAudio.FileName}");
                        }
                        break;

                    default:
                        Console.WriteLine($"   ❌ Unknown AudioFileType: {options.AudioFileType}");
                        break;
                }
            }

            Console.WriteLine($"🎯 Total audio files created: {audioFiles.Count}");

            return audioFiles;
        }

        private AudioFile CreateAudioFileRecord(int vocabId, string language, string fileName,
            string voiceName, decimal speechRate, bool isOddFile)
        {
            return new AudioFile
            {
                VocabId = vocabId,
                Language = language,
                FileName = fileName,
                FilePath = "", // Sẽ update sau khi tạo file thực tế
                VoiceName = voiceName,
                SpeechRate = speechRate,
                IsOddFile = isOddFile,
                CreatedDate = DateTime.UtcNow,
                IsGenerated = false
            };
        }

        private string GetCellValue(ExcelWorksheet worksheet, int row, int column)
        {
            var value = worksheet.Cells[row, column].Text?.Trim();
            return string.IsNullOrEmpty(value) ? "" : value; // Return empty string instead of null
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

        // Query methods for future use
        public async Task<List<Vocabulary>> GetVocabulariesByTypeAsync(string fileType, int take = 10)
        {
            return await _context.Vocabularies
                .Where(v => v.Category == fileType && v.IsActive)
                .OrderByDescending(v => v.CreatedDate)
                .Take(take)
                .ToListAsync();
        }

        public async Task<List<ProcessingSession>> GetRecentSessionsAsync(int take = 5)
        {
            return await _context.ProcessingSessions
                .OrderByDescending(s => s.ProcessedDate)
                .Take(take)
                .ToListAsync();
        }

        // 🔍 METHOD MỚI: Lấy summary data cho smart processing
        public async Task<ExistingDataSummary> GetExistingDataSummary(string dateString, string fileType, string audioFileType)
        {
            var summary = new ExistingDataSummary();

            try
            {
                if (!DateTime.TryParseExact(dateString, "dd-MM-yyyy", null,
                    System.Globalization.DateTimeStyles.None, out DateTime targetDate))
                {
                    return summary; // Return empty summary
                }

                // Count existing vocabularies for this date and file type
                summary.VocabularyCount = await _context.Vocabularies
                    .Where(v => v.CreatedDate.Date == targetDate.Date && v.Category == fileType)
                    .CountAsync();

                // Count existing audio files for this date and audio file type
                summary.AudioFileCount = await GetExistingAudioCountForDateAsync(dateString, audioFileType);

                // Calculate next audio number
                summary.NextAudioNumber = summary.AudioFileCount + 1;

                return summary;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Error getting existing data summary: {ex.Message}");
                return summary;
            }
        }
        // Method sync để lấy existing audio count
        private int GetExistingAudioCountForDate(string dateString, string audioFileType)
        {
            try
            {
                // Tạo pattern dựa trên AudioFileType
                string fileNamePattern = "";
                switch (audioFileType)
                {
                    case "VI-EN":
                        fileNamePattern = $"EN-{dateString}_";
                        break;
                    case "JP-EN":
                        fileNamePattern = $"JP-{dateString}_";
                        break;
                    case "ZH-EN":
                        fileNamePattern = $"ZH-{dateString}_";
                        break;
                    case "TUVUNG":
                        fileNamePattern = $"Vocab-{dateString}_";
                        break;
                    default:
                        return 0;
                }

                // Đếm số audio files có pattern này trong database (sync)
                var count = _context.AudioFiles
                    .Where(a => a.FileName.StartsWith(fileNamePattern))
                    .Count(); // Sử dụng Count() thay vì CountAsync()

                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Warning: Could not check existing audio files: {ex.Message}");
                return 0;
            }
        }

        // 🔍 METHOD MỚI: Async version của GetExistingAudioCountForDate
        private async Task<int> GetExistingAudioCountForDateAsync(string dateString, string audioFileType)
        {
            try
            {
                // Tạo pattern dựa trên AudioFileType
                string fileNamePattern = "";
                switch (audioFileType)
                {
                    case "VI-EN":
                        fileNamePattern = $"EN-{dateString}_";
                        break;
                    case "JP-EN":
                        fileNamePattern = $"JP-{dateString}_";
                        break;
                    case "ZH-EN":
                        fileNamePattern = $"ZH-{dateString}_";
                        break;
                    case "TUVUNG":
                        fileNamePattern = $"Vocab-{dateString}_";
                        break;
                    default:
                        return 0;
                }

                // Đếm số audio files có pattern này trong database
                var count = await _context.AudioFiles
                    .Where(a => a.FileName.StartsWith(fileNamePattern))
                    .CountAsync();

                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Warning: Could not check existing audio files: {ex.Message}");
                return 0; // Nếu lỗi thì bắt đầu từ 1
            }
        }
        public void Dispose()
        {
            _context?.Dispose();
        }
    }

    // Helper classes cho existing data summary
    public class ExistingDataSummary
    {
        public int VocabularyCount { get; set; } = 0;
        public int AudioFileCount { get; set; } = 0;
        public int NextAudioNumber { get; set; } = 1;
    }
    public class ProcessingResult
    {
        public List<Vocabulary> NewVocabularies { get; set; } = new List<Vocabulary>();
        public List<VocabularyMapping> VocabularyMappings { get; set; } = new List<VocabularyMapping>();
        public int SkippedCount { get; set; } = 0;
        public int TotalProcessed { get; set; } = 0;
    }

    public class VocabularyMapping
    {
        public int ExcelRow { get; set; }
        public Vocabulary? Vocabulary { get; set; }
        public bool IsSkipped { get; set; }
    }
}
