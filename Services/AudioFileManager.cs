using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Text_to_Image.Models;
using Text_to_Image.Config;

namespace Text_to_Image.Services
{
    // Model cho thông tin file audio
    public class AudioFileInfo
    {
        public int RowIndex { get; set; }
        public string SourceColumn { get; set; }
        public string FileName { get; set; }
        public string VoiceName { get; set; }
        public string Language { get; set; }
        public bool IsOddFile { get; set; }
        public string SourceText { get; set; }
        public string OutputPath { get; set; }
    }

    public static class AudioFileManager
    {
        // ========================================
        // AUDIO FILE NAMER SECTION
        // ========================================

        public static List<AudioFileInfo> GenerateAudioFileNames(ProcessingOptions options, int totalRows)
        {
            var audioFiles = new List<AudioFileInfo>();
            string dateToUse = string.IsNullOrWhiteSpace(options.CustomDate)
                ? DateTime.Now.ToString("dd-MM-yyyy")
                : options.CustomDate;

            switch (options.AudioFileType)
            {
                case "VI-EN":
                    audioFiles.AddRange(GenerateViEnAudioFiles(options, totalRows, dateToUse));
                    break;
                case "JP-EN":
                    audioFiles.AddRange(GenerateJpEnAudioFiles(options, totalRows, dateToUse));
                    break;
                case "ZH-EN":
                    audioFiles.AddRange(GenerateZhEnAudioFiles(options, totalRows, dateToUse));
                    break;
                case "TUVUNG":
                    audioFiles.AddRange(GenerateTuVungAudioFiles(options, totalRows, dateToUse));
                    break;
                case "EN-ONLY":
                case "JP-ONLY":
                case "ZH-ONLY":
                    audioFiles.AddRange(GenerateSingleLanguageAudioFiles(options, totalRows, dateToUse));
                    break;
            }

            return audioFiles;
        }

        private static List<AudioFileInfo> GenerateViEnAudioFiles(ProcessingOptions options, int totalRows, string dateToUse)
        {
            var files = new List<AudioFileInfo>();

            for (int row = 1; row <= totalRows; row++)
            {
                // Vietnamese → File số LẺ (1, 3, 5, 7...)
                files.Add(new AudioFileInfo
                {
                    RowIndex = row,
                    SourceColumn = options.VietnameseColumn,
                    FileName = $"EN-{dateToUse}_{(row * 2 - 1):00}.mp3",
                    VoiceName = AppConfig.VOICE_MAPPING["vietnamese"],
                    Language = "vietnamese",
                    IsOddFile = true
                });

                // English → File số CHẴN (2, 4, 6, 8...)
                files.Add(new AudioFileInfo
                {
                    RowIndex = row,
                    SourceColumn = options.EnglishColumn,
                    FileName = $"EN-{dateToUse}_{(row * 2):00}.mp3",
                    VoiceName = AppConfig.VOICE_MAPPING["english"],
                    Language = "english",
                    IsOddFile = false
                });
            }

            return files;
        }

        private static List<AudioFileInfo> GenerateJpEnAudioFiles(ProcessingOptions options, int totalRows, string dateToUse)
        {
            var files = new List<AudioFileInfo>();

            for (int row = 1; row <= totalRows; row++)
            {
                // English → File số LẺ (1, 3, 5, 7...)
                files.Add(new AudioFileInfo
                {
                    RowIndex = row,
                    SourceColumn = options.EnglishColumn,
                    FileName = $"JP-{dateToUse}_{(row * 2 - 1):00}.mp3",
                    VoiceName = AppConfig.VOICE_MAPPING["english"],
                    Language = "english",
                    IsOddFile = true
                });

                // Japanese → File số CHẴN (2, 4, 6, 8...)
                files.Add(new AudioFileInfo
                {
                    RowIndex = row,
                    SourceColumn = options.JapaneseColumn,
                    FileName = $"JP-{dateToUse}_{(row * 2):00}.mp3",
                    VoiceName = AppConfig.VOICE_MAPPING["japanese"],
                    Language = "japanese",
                    IsOddFile = false
                });
            }

            return files;
        }

        private static List<AudioFileInfo> GenerateZhEnAudioFiles(ProcessingOptions options, int totalRows, string dateToUse)
        {
            var files = new List<AudioFileInfo>();

            for (int row = 1; row <= totalRows; row++)
            {
                // English → File số LẺ (1, 3, 5, 7...)
                files.Add(new AudioFileInfo
                {
                    RowIndex = row,
                    SourceColumn = options.EnglishColumn,
                    FileName = $"ZH-{dateToUse}_{(row * 2 - 1):00}.mp3",
                    VoiceName = AppConfig.VOICE_MAPPING["english"],
                    Language = "english",
                    IsOddFile = true
                });

                // Chinese → File số CHẴN (2, 4, 6, 8...)
                files.Add(new AudioFileInfo
                {
                    RowIndex = row,
                    SourceColumn = options.ChineseColumn,
                    FileName = $"ZH-{dateToUse}_{(row * 2):00}.mp3",
                    VoiceName = AppConfig.VOICE_MAPPING["chinese"],
                    Language = "chinese",
                    IsOddFile = false
                });
            }

            return files;
        }

        private static List<AudioFileInfo> GenerateTuVungAudioFiles(ProcessingOptions options, int totalRows, string dateToUse)
        {
            var files = new List<AudioFileInfo>();

            for (int row = 1; row <= totalRows; row++)
            {
                // Vietnamese → File số LẺ (1, 3, 5, 7...)
                files.Add(new AudioFileInfo
                {
                    RowIndex = row,
                    SourceColumn = options.VietnameseColumn,
                    FileName = $"Vocab-{dateToUse}_{(row * 2 - 1):00}.mp3",
                    VoiceName = AppConfig.VOICE_MAPPING["vietnamese"],
                    Language = "vietnamese",
                    IsOddFile = true
                });

                // Japanese → File số CHẴN (2, 4, 6, 8...)
                files.Add(new AudioFileInfo
                {
                    RowIndex = row,
                    SourceColumn = options.JapaneseColumn,
                    FileName = $"Vocab-{dateToUse}_{(row * 2):00}.mp3",
                    VoiceName = AppConfig.VOICE_MAPPING["japanese"],
                    Language = "japanese",
                    IsOddFile = false
                });
            }

            return files;
        }

        private static List<AudioFileInfo> GenerateSingleLanguageAudioFiles(ProcessingOptions options, int totalRows, string dateToUse)
        {
            var files = new List<AudioFileInfo>();
            string column = "";
            string prefix = "";
            string language = "";

            switch (options.AudioFileType)
            {
                case "EN-ONLY":
                    column = options.EnglishColumn;
                    prefix = "EN";
                    language = "english";
                    break;
                case "JP-ONLY":
                    column = options.JapaneseColumn;
                    prefix = "JP";
                    language = "japanese";
                    break;
                case "ZH-ONLY":
                    column = options.ChineseColumn;
                    prefix = "ZH";
                    language = "chinese";
                    break;
            }

            for (int row = 1; row <= totalRows; row++)
            {
                files.Add(new AudioFileInfo
                {
                    RowIndex = row,
                    SourceColumn = column,
                    FileName = $"{prefix}-{dateToUse}_{row:00}.mp3",
                    VoiceName = GetVoiceName(language),
                    Language = language,
                    IsOddFile = true
                });
            }

            return files;
        }

        public static string GetVoiceName(string language)
        {
            if (AppConfig.VOICE_MAPPING.ContainsKey(language.ToLower()))
            {
                return AppConfig.VOICE_MAPPING[language.ToLower()];
            }

            return AppConfig.VOICE_MAPPING["english"]; // Default voice
        }

        // ========================================
        // AUDIO FILE RENAMER SECTION
        // ========================================

        public static void RenameAudioFiles(string folderPath, string day, string month, string year, string fileName)
        {
            try
            {
                // Tìm tất cả file .mp3 trong thư mục
                string[] audioFiles = Directory.GetFiles(folderPath, "*.mp3");

                // Xác định tiền tố dựa trên loại file
                string prefix = DeterminePrefix(fileName);
                if (string.IsNullOrEmpty(prefix))
                {
                    Console.WriteLine("Unable to determine file type for audio renaming.");
                    return;
                }

                // Tạo pattern để tìm file phù hợp
                var matchedFiles = FindMatchingAudioFiles(audioFiles, day, month, year);

                if (matchedFiles.Count == 0)
                {
                    Console.WriteLine($"No audio files found matching date pattern: {day}-{month}-{year}");
                    return;
                }

                ProcessFileRenaming(matchedFiles, prefix, day, month, year);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during audio file renaming: {ex.Message}");
            }
        }

        private static string DeterminePrefix(string fileName)
        {
            fileName = fileName.ToLower();

            if (fileName.Contains("japan") || fileName.Contains("jp"))
                return "JP";
            else if (fileName.Contains("english") || fileName.Contains("en"))
                return "EN";
            else if (fileName.Contains("chinese") || fileName.Contains("china") || fileName.Contains("zh"))
                return "ZH";
            else if (fileName.Contains("tuvung") || fileName.Contains("vocab"))
                return "VC";
            else
                return PromptUserForPrefix();
        }

        private static string PromptUserForPrefix()
        {
            Console.WriteLine("Unable to determine file type automatically.");
            Console.WriteLine("Please select the prefix for audio files:");
            Console.WriteLine("1. JP (Japanese)");
            Console.WriteLine("2. EN (English)");
            Console.WriteLine("3. ZH (Chinese)");
            Console.WriteLine("4. VC (Vocabulary)");
            Console.Write("Enter your choice (1-4): ");

            string choice = Console.ReadLine()?.Trim();
            return choice switch
            {
                "1" => "JP",
                "2" => "EN",
                "3" => "ZH",
                "4" => "VC",
                _ => null
            };
        }

        private static List<string> FindMatchingAudioFiles(string[] audioFiles, string day, string month, string year)
        {
            // Hỗ trợ cả format: 19-2-2025-0001.mp3 và 19-02-2025-0001.mp3
            string pattern1 = $@"^{day}-{month}-{year}-(\d{{4}})\.mp3$";
            string pattern2 = $@"^{day.PadLeft(2, '0')}-{month.PadLeft(2, '0')}-{year}-(\d{{4}})\.mp3$";

            Regex regex1 = new Regex(pattern1, RegexOptions.IgnoreCase);
            Regex regex2 = new Regex(pattern2, RegexOptions.IgnoreCase);

            List<string> matchedFiles = new List<string>();

            foreach (string audioFile in audioFiles)
            {
                string fileName = Path.GetFileName(audioFile);

                if (regex1.IsMatch(fileName) || regex2.IsMatch(fileName))
                {
                    matchedFiles.Add(audioFile);
                }
            }

            return matchedFiles.OrderBy(f => f).ToList();
        }

        private static void ProcessFileRenaming(List<string> matchedFiles, string prefix, string day, string month, string year)
        {
            Console.WriteLine($"Found {matchedFiles.Count} audio file(s) to rename:");

            for (int i = 0; i < matchedFiles.Count; i++)
            {
                string oldFilePath = matchedFiles[i];
                string oldFileName = Path.GetFileName(oldFilePath);

                // Tạo tên file mới với format: JP-19-02-2025_01.mp3
                string newFileName = $"{prefix}-{day.PadLeft(2, '0')}-{month.PadLeft(2, '0')}-{year}_{(i + 1).ToString().PadLeft(2, '0')}.mp3";
                string newFilePath = Path.Combine(Path.GetDirectoryName(oldFilePath), newFileName);

                if (TryRenameFile(oldFilePath, newFilePath, oldFileName, newFileName))
                {
                    Console.WriteLine($"Renamed: {oldFileName} -> {newFileName}");
                }
            }

            Console.WriteLine($"Audio file renaming completed. Processed {matchedFiles.Count} file(s).");
        }

        private static bool TryRenameFile(string oldPath, string newPath, string oldName, string newName)
        {
            try
            {
                if (File.Exists(newPath))
                {
                    Console.WriteLine($"Warning: File {newName} already exists. Skipping {oldName}");
                    return false;
                }

                File.Move(oldPath, newPath);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error renaming {oldName}: {ex.Message}");
                return false;
            }
        }
    }
}