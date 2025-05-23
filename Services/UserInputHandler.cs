using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Text_to_Image.Models;

namespace Text_to_Image.Services
{
    public static class UserInputHandler
    {
        public static void GetProcessingInputs(ProcessingOptions options)
        {
            // Kiểm tra loại file dựa trên tên file thực tế
            string fileName = options.FileName.ToLower();
            bool isEnglishFile = fileName.Contains("english");
            bool isTuVungFile = fileName.Contains("tuvung");
            bool isJapaneseFile = fileName.Contains("japanese");
            bool isChineseFile = fileName.Contains("chinese");

            // Câu hỏi 2: Chọn cột chuyển đổi text (bỏ qua nếu là English)
            if (!isEnglishFile)
            {
                if (isTuVungFile)
                {
                    Console.Write("\nConvert to <img> (TuVung). Default: GJ for TuVung. (Enter to skip). ");
                    options.ColumnInput = Console.ReadLine()?.ToUpper();
                }
                else if (isJapaneseFile)
                {
                    Console.Write("\nConvert to <img> (Japanese). Default: CF for Japanese. (Enter to skip). ");
                    options.ColumnInput = Console.ReadLine()?.ToUpper();
                }
                else if (isChineseFile)
                {
                    Console.Write("\nConvert to <img> (Chinese). Default: CF for Chinese. (Enter to skip). ");
                    options.ColumnInput = Console.ReadLine()?.ToUpper();
                }
                else
                {
                    Console.Write("\nConvert to <img> (JP-ZH). Default: CF for JP-ZH. (Enter to skip). ");
                    options.ColumnInput = Console.ReadLine()?.ToUpper();
                }
            }

            // Câu hỏi 3: Chọn cột âm thanh (hiển thị cho tất cả các file)
            if (isTuVungFile)
            {
                Console.Write("\nConvert to [sound] (TuVung). Default: DH. (Enter to skip). ");
                options.SoundColumns = Console.ReadLine()?.ToUpper();
            }
            else if (isEnglishFile)
            {
                Console.Write("\nConvert to [sound] (English). Default: DE for English. (Enter to skip). ");
                options.SoundColumns = Console.ReadLine()?.ToUpper();
            }
            else if (isJapaneseFile)
            {
                Console.Write("\nConvert to [sound] (Japanese). Default: E for Japanese. (Enter to skip). ");
                options.SoundColumns = Console.ReadLine()?.ToUpper();
            }
            else if (isChineseFile)
            {
                Console.Write("\nConvert to [sound] (Chinese). Default: E for Chinese. (Enter to skip). ");
                options.SoundColumns = Console.ReadLine()?.ToUpper();
            }
            else
            {
                Console.Write("\nConvert to [sound] (EN-JP-ZH). Default:DE for EN | Default: E for JP-ZH. (Enter to skip). ");
                options.SoundColumns = Console.ReadLine()?.ToUpper();
            }

            // Câu hỏi 4: Chọn cột Kanji (bỏ qua nếu là English)
            if (!isEnglishFile)
            {
                if (isTuVungFile)
                {
                    Console.Write("\nSelect source Kanji column (TuVung). Default: F for TuVung. (Enter to skip). ");
                    options.KanjiColumn = Console.ReadLine()?.ToUpper();
                    if (string.IsNullOrWhiteSpace(options.KanjiColumn))
                        options.KanjiColumn = "F"; // Default for TuVung

                    Console.Write("Select column to save processed Kanji results (TuVung). Default: B. (Enter to skip). ");
                    options.KanjiOutputColumn = Console.ReadLine()?.ToUpper();
                    if (string.IsNullOrWhiteSpace(options.KanjiOutputColumn))
                        options.KanjiOutputColumn = "B"; // Default for TuVung
                }
                else if (isJapaneseFile)
                {
                    Console.Write("\nSelect source Kanji column (Japanese). Default: C for Japanese. (Enter to skip). ");
                    options.KanjiColumn = Console.ReadLine()?.ToUpper();
                    if (string.IsNullOrWhiteSpace(options.KanjiColumn))
                        options.KanjiColumn = "C"; // Default for Japanese

                    Console.Write("Select column to save processed Kanji results (Japanese). Default: B. (Enter to skip). ");
                    options.KanjiOutputColumn = Console.ReadLine()?.ToUpper();
                    if (string.IsNullOrWhiteSpace(options.KanjiOutputColumn))
                        options.KanjiOutputColumn = "B"; // Default for Japanese
                }
                else if (isChineseFile)
                {
                    Console.Write("\nSelect source Kanji column (Chinese). Default: C for Chinese. (Enter to skip). ");
                    options.KanjiColumn = Console.ReadLine()?.ToUpper();
                    if (string.IsNullOrWhiteSpace(options.KanjiColumn))
                        options.KanjiColumn = "C"; // Default for Chinese

                    Console.Write("Select column to save processed Kanji results (Chinese). Default: B. (Enter to skip). ");
                    options.KanjiOutputColumn = Console.ReadLine()?.ToUpper();
                    if (string.IsNullOrWhiteSpace(options.KanjiOutputColumn))
                        options.KanjiOutputColumn = "B"; // Default for Chinese
                }
                else
                {
                    Console.Write("\nSelect source Kanji column (JP-ZH). Default: C for JP-ZH. (Enter to skip). ");
                    options.KanjiColumn = Console.ReadLine()?.ToUpper();
                    if (string.IsNullOrWhiteSpace(options.KanjiColumn))
                        options.KanjiColumn = "C"; // Default for JP-ZH

                    Console.Write("Select column to save processed Kanji results (JP-ZH). Default: B. (Enter to skip). ");
                    options.KanjiOutputColumn = Console.ReadLine()?.ToUpper();
                    if (string.IsNullOrWhiteSpace(options.KanjiOutputColumn))
                        options.KanjiOutputColumn = "B"; // Default for JP-ZH
                }
            }

            // Câu hỏi 5: Đổi tên file âm thanh (vẫn hiển thị cho tất cả các file)
            Console.Write("\nDo you want to rename audio files? (Y/N): ");
            string renameAnswer = Console.ReadLine()?.Trim().ToUpper();
            options.RenameAudioFiles = (renameAnswer == "Y");

            // Nếu chọn đổi tên file âm thanh, cho phép chọn folder
            if (options.RenameAudioFiles)
            {
                options.AudioFolderPath = FolderSelector.SelectAudioFolder();
                FolderSelector.DisplayFolderInfo(options.AudioFolderPath);
            }
        }
    }
}
