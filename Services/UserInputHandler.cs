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
            // Câu hỏi 2: Chọn cột chuyển đổi text
            if (options.FileName.Contains("tuvung"))
            {
                Console.Write("\nConvert to <img> (Vocab). Default: GJ for Vocab. (Enter to skip). ");
                options.ColumnInput = Console.ReadLine()?.ToUpper();
            }
            else
            {
                Console.Write("\nConvert to <img> (JP-ZH). Default: CF for JP-ZH. (Enter to skip). ");
                options.ColumnInput = Console.ReadLine()?.ToUpper();
            }

            // Câu hỏi 3: Chọn cột âm thanh
            if (options.FileName.Contains("tuvung"))
            {
                Console.Write("\nConvert to [sound] (Vocab). Default: DH. (Enter to skip). ");
                options.SoundColumns = Console.ReadLine()?.ToUpper();
            }
            else
            {
                Console.Write("\nConvert to [sound] (EN-JP-ZH). Default:DE for EN | Default: E for JP-ZH. (Enter to skip). ");
                options.SoundColumns = Console.ReadLine()?.ToUpper();
            }

            // Câu hỏi 4: Chọn cột Kanji
            if (options.FileName.Contains("tuvung"))
            {
                Console.Write("\nSelect source Kanji column (Vocab). Default: F for Vocab. (Enter to skip). ");
                options.KanjiColumn = Console.ReadLine()?.ToUpper();
                if (string.IsNullOrWhiteSpace(options.KanjiColumn))
                    options.KanjiColumn = "F"; // Default for Vocab
                Console.Write("Select column to save processed Kanji results (example: A, B, D, E...): ");
                string kanjiOutput;
                do
                {
                    kanjiOutput = Console.ReadLine()?.ToUpper();
                    if (string.IsNullOrWhiteSpace(kanjiOutput))
                    {
                        Console.Write("Please enter target column (cannot be empty): ");
                    }
                } while (string.IsNullOrWhiteSpace(kanjiOutput));
                options.KanjiOutputColumn = kanjiOutput;
            }
            else
            {
                Console.Write("\nSelect source Kanji column (JP-ZH). Default: C for JP-ZH. (Enter to skip). ");
                options.KanjiColumn = Console.ReadLine()?.ToUpper();
                if (string.IsNullOrWhiteSpace(options.KanjiColumn))
                    options.KanjiColumn = "C"; // Default for JP-ZH
                Console.Write("Select column to save processed Kanji results (example: A, B, D, E...): ");
                string kanjiOutput;
                do
                {
                    kanjiOutput = Console.ReadLine()?.ToUpper();
                    if (string.IsNullOrWhiteSpace(kanjiOutput))
                    {
                        Console.Write("Please enter target column (cannot be empty): ");
                    }
                } while (string.IsNullOrWhiteSpace(kanjiOutput));
                options.KanjiOutputColumn = kanjiOutput;
            }

            // Câu hỏi 5: Đổi tên file âm thanh
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
