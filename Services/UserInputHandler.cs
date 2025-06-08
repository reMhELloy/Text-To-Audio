using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Text_to_Image.Models;
using Text_to_Image.Services;

namespace Text_to_Image.Services
{
    public static class UserInputHandler
    {
        // Method 1: CHỈ hỏi Excel processing inputs
        public static void GetExcelProcessingInputs(ProcessingOptions options)
        {
            // Kiểm tra loại file dựa trên tên file thực tế
            string fileName = options.FileName.ToLower();
            bool isEnglishFile = fileName.Contains("english");
            bool isTuVungFile = fileName.Contains("tuvung");
            bool isJapaneseFile = fileName.Contains("japanese");
            bool isChineseFile = fileName.Contains("chinese");

            // Câu hỏi 1: Chọn cột chuyển đổi text (BỎ QUA ENGLISH)
            if (isTuVungFile)
            {
                Console.Write("\nConvert to <img> (TuVung). Default: C -> G. (Enter to skip). ");
                string input = Console.ReadLine()?.ToUpper();
                options.ColumnInput = string.IsNullOrWhiteSpace(input) ? "C" : input;
            }
            else if (isJapaneseFile)
            {
                Console.Write("\nConvert to <img> (Japanese). Default: C -> G. (Enter to skip). ");
                string input = Console.ReadLine()?.ToUpper();
                options.ColumnInput = string.IsNullOrWhiteSpace(input) ? "C" : input;
            }
            else if (isChineseFile)
            {
                Console.Write("\nConvert to <img> (Chinese). Default: C -> G. (Enter to skip). ");
                string input = Console.ReadLine()?.ToUpper();
                options.ColumnInput = string.IsNullOrWhiteSpace(input) ? "C" : input;
            }
            else if (!isEnglishFile)
            {
                Console.Write("\nConvert to <img>. Default: C -> G. (Enter to skip). ");
                string input = Console.ReadLine()?.ToUpper();
                options.ColumnInput = string.IsNullOrWhiteSpace(input) ? "C" : input;
            }
            // English file: NO convert to <img> feature

            // Câu hỏi 2: Tạo sound công thức
            if (isTuVungFile)
            {
                Console.Write("\nCreate [sound] formulas (TuVung). Default: EF - E(odd-VI) & F(even-JP). (Enter to skip). ");
                string soundInput = Console.ReadLine()?.ToUpper();
                options.SoundColumns = string.IsNullOrWhiteSpace(soundInput) ? "EF" : soundInput;
            }
            else if (isJapaneseFile)
            {
                Console.Write("\nCreate [sound] formulas (Japanese). Default: EF - E(odd-EN) & F(even-JP). (Enter to skip). ");
                string soundInput = Console.ReadLine()?.ToUpper();
                options.SoundColumns = string.IsNullOrWhiteSpace(soundInput) ? "EF" : soundInput;
            }
            else if (isChineseFile)
            {
                Console.Write("\nCreate [sound] formulas (Chinese). Default: EF - E(odd-EN) & F(even-ZH). (Enter to skip). ");
                string soundInput = Console.ReadLine()?.ToUpper();
                options.SoundColumns = string.IsNullOrWhiteSpace(soundInput) ? "EF" : soundInput;
            }
            else if (isEnglishFile)
            {
                Console.Write("\nConvert to [sound] (English). Default: DE for English. (Enter to skip). ");
                options.SoundColumns = Console.ReadLine()?.ToUpper();
            }
            else
            {
                Console.Write("\nCreate [sound] formulas. Default: EF - E(odd-EN) & F(even-native). (Enter to skip). ");
                string soundInput = Console.ReadLine()?.ToUpper();
                options.SoundColumns = string.IsNullOrWhiteSpace(soundInput) ? "EF" : soundInput;
            }

            // Câu hỏi 3: Chọn cột Kanji (bỏ qua nếu là English)
            if (!isEnglishFile)
            {
                if (isTuVungFile)
                {
                    Console.Write("\nSelect source Kanji column (TuVung). Default: C. (Enter to skip). ");
                    string kanjiInput = Console.ReadLine()?.ToUpper();

                    // CHỈ SET NẾU USER NHẬP GÌ ĐÓ
                    if (!string.IsNullOrWhiteSpace(kanjiInput))
                    {
                        options.KanjiColumn = kanjiInput;
                        options.KanjiOutputColumn = "B"; // LUÔN DÙNG B
                    }
                }
                else if (isJapaneseFile)
                {
                    Console.Write("\nSelect source Kanji column (Japanese). Default: C. (Enter to skip). ");
                    string kanjiInput = Console.ReadLine()?.ToUpper();

                    if (!string.IsNullOrWhiteSpace(kanjiInput))
                    {
                        options.KanjiColumn = kanjiInput;
                        options.KanjiOutputColumn = "B"; // LUÔN DÙNG B
                    }
                }
                else if (isChineseFile)
                {
                    Console.Write("\nSelect source Kanji column (Chinese). Default: C. (Enter to skip). ");
                    string kanjiInput = Console.ReadLine()?.ToUpper();

                    if (!string.IsNullOrWhiteSpace(kanjiInput))
                    {
                        options.KanjiColumn = kanjiInput;
                        options.KanjiOutputColumn = "B"; // LUÔN DÙNG B
                    }
                }
                else
                {
                    Console.Write("\nSelect source Kanji column. Default: C. (Enter to skip). ");
                    string kanjiInput = Console.ReadLine()?.ToUpper();

                    if (!string.IsNullOrWhiteSpace(kanjiInput))
                    {
                        options.KanjiColumn = kanjiInput;
                        options.KanjiOutputColumn = "B"; // LUÔN DÙNG B
                    }
                }
            }
        }

        // Method 2: CHỈ hỏi Audio Creation inputs - SAU KHI Excel hoàn thành
        public static void GetAudioCreationInputs(ProcessingOptions options)
        {
            Console.Write("\nDo you want to create audio files using Azure Speech? (Y/N): ");
            string createAudioAnswer = Console.ReadLine()?.Trim().ToUpper();
            options.CreateAudioFiles = (createAudioAnswer == "Y");

            if (options.CreateAudioFiles)
            {
                // Tự động xác định loại file và cột mặc định cho audio
                SetDefaultAudioConfiguration(options);

                // Chỉ cần chọn thư mục lưu audio (mặc định mở dialog)
                options.AudioOutputFolder = AudioFolderManager.SelectAudioOutputFolder(options);

                if (string.IsNullOrEmpty(options.AudioOutputFolder))
                {
                    Console.WriteLine("No audio output folder selected. Skipping audio creation.");
                    options.CreateAudioFiles = false;
                }
            }
        }

        // Method 3: CHỈ hỏi Audio Rename inputs - SAU KHI Audio Creation hoàn thành
        public static void GetAudioRenameInputs(ProcessingOptions options)
        {
            Console.Write("\nDo you want to rename audio files? (Y/N): ");
            string renameAnswer = Console.ReadLine()?.Trim().ToUpper();
            options.RenameAudioFiles = (renameAnswer == "Y");

            // Nếu chọn đổi tên file âm thanh, cho phép chọn folder
            if (options.RenameAudioFiles)
            {
                options.AudioFolderPath = AudioFolderManager.SelectAudioFolder();
                AudioFolderManager.DisplayFolderInfo(options.AudioFolderPath);
            }
        }
        // Method 4: HỎI VÀ THỰC HIỆN DATABASE SAVE NGAY - SAU KHI Excel và Audio hoàn thành
        public static async Task HandleDatabaseSave(ProcessingOptions options)
        {
            Console.Write("\nDo you want to save data to SQL Server database? (Y/N): ");
            string saveDatabaseAnswer = Console.ReadLine()?.Trim().ToUpper();

            if (saveDatabaseAnswer == "Y")
            {
                try
                {
                    Console.WriteLine("✅ Saving data to SQL Server...");
                    Console.WriteLine(new string('=', 50));

                    using var dbService = new Text_to_Image.Data.DatabaseService();
                    var session = await dbService.SaveExcelDataToDatabaseAsync(options);

                    Console.WriteLine($"\n📊 Database Summary:");
                    Console.WriteLine($"✅ Session ID: {session.SessionId}");
                    Console.WriteLine($"✅ Processed: {session.ProcessedRows} vocabulary entries");
                    Console.WriteLine($"✅ Audio records: {session.AudioFilesCreated}");
                    Console.WriteLine($"✅ File type: {session.FileType}");
                    Console.WriteLine($"✅ Date: {session.DateUsed}");
                    Console.WriteLine(new string('=', 50));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ Database save failed: {ex.Message}");
                    Console.WriteLine("Excel processing completed but data not saved to database.");
                    if (ex.InnerException != null)
                    {
                        Console.WriteLine($"Error details: {ex.InnerException.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("ℹ️ Database save skipped (Excel only mode).");
            }
        }

        private static void SetDefaultAudioConfiguration(ProcessingOptions options)
        {
            string fileName = options.FileName.ToLower();

            if (fileName.Contains("tuvung"))
            {
                options.VietnameseColumn = "A"; // TuVung: VI ở cột A 
                options.JapaneseColumn = "C";   // JP ở cột C 
                options.AudioFileType = "TUVUNG";
            }
            else if (fileName.Contains("english"))
            {
                options.VietnameseColumn = "A"; // English file: VI ở cột A
                options.EnglishColumn = "B";    // EN ở cột B
                options.AudioFileType = "VI-EN";
            }
            else if (fileName.Contains("japanese"))
            {
                options.EnglishColumn = "A";    // Japanese file: EN ở cột A
                options.JapaneseColumn = "C";   // JP ở cột C
                options.AudioFileType = "JP-EN";
            }
            else if (fileName.Contains("chinese"))
            {
                options.EnglishColumn = "A";    // Chinese file: EN ở cột A
                options.ChineseColumn = "C";    // ZH ở cột C
                options.AudioFileType = "ZH-EN";
            }
            else
            {
                // Default to Vietnamese-English
                options.VietnameseColumn = "A";
                options.EnglishColumn = "B";
                options.AudioFileType = "VI-EN";
            }
        }
    }
}