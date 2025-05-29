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

                // Đọc và lưu dữ liệu từ Excel
                var vocabularies = await ReadExcelAndCreateVocabularies(options, session.SessionId);

                if (vocabularies.Any())
                {
                    _context.Vocabularies.AddRange(vocabularies);
                    await _context.SaveChangesAsync(); // Save vocabularies FIRST để có VocabId

                    Console.WriteLine($"✅ Saved {vocabularies.Count} vocabulary entries to database.");

                    // DEBUG: Check audio creation conditions
                    Console.WriteLine($"🔍 Audio creation check:");
                    Console.WriteLine($"   options.CreateAudioFiles: {options.CreateAudioFiles}");
                    Console.WriteLine($"   options.SoundColumns: '{options.SoundColumns}'");
                    Console.WriteLine($"   SoundColumns null/empty: {string.IsNullOrWhiteSpace(options.SoundColumns)}");

                    // Tạo audio file records nếu có - SAU KHI vocabularies đã có ID
                    if (options.CreateAudioFiles && !string.IsNullOrWhiteSpace(options.SoundColumns))
                    {
                        Console.WriteLine("✅ Conditions met - creating audio files...");
                        var audioFiles = CreateAudioFileRecords(vocabularies, options); // Remove await
                        if (audioFiles.Any())
                        {
                            _context.AudioFiles.AddRange(audioFiles);
                            await _context.SaveChangesAsync();

                            session.AudioFilesCreated = audioFiles.Count;
                            Console.WriteLine($"✅ Created {audioFiles.Count} audio file records.");
                        }
                        else
                        {
                            Console.WriteLine("❌ No audio files generated from CreateAudioFileRecords");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"❌ Audio creation skipped:");
                        Console.WriteLine($"   CreateAudioFiles: {options.CreateAudioFiles}");
                        Console.WriteLine($"   SoundColumns: '{options.SoundColumns}'");
                    }
                }

                // Cập nhật session
                session.ProcessedRows = vocabularies.Count;
                session.TotalRows = vocabularies.Count;
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

        public void Dispose()
        {
            _context?.Dispose();
        }
    }
}