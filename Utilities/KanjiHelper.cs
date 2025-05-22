using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JapaneseConverter
{
    public static class KanjiHelper
    {
        /// <summary>
        /// Kiểm tra xem một ký tự có phải là kanji trong tiếng Nhật hay không
        /// </summary>
        /// <param name="c">Ký tự cần kiểm tra</param>
        /// <returns>true nếu là kanji, false nếu không phải</returns>
        public static bool IsKanji(char c)
        {
            int value = Convert.ToInt32(c);

            // Block 1: CJK Unified Ideographs - Các kanji thông dụng nhất
            // Bao gồm 常用漢字 (Jōyō kanji) và 人名用漢字 (Jinmeiyō kanji)
            if (value >= 0x4E00 && value <= 0x9FFF)
                return true;

            // Block 2: CJK Unified Ideographs Extension A 
            // Chứa các kanji cổ và ít dùng hơn
            if (value >= 0x3400 && value <= 0x4DBF)
                return true;

            // Block 3: CJK Compatibility Ideographs
            // Chứa các biến thể của kanji được sử dụng trong tiếng Nhật
            if (value >= 0xF900 && value <= 0xFAFF)
                return true;

            // Block 4: CJK Compatibility Ideographs Supplement
            // Bổ sung thêm các biến thể đặc biệt
            if (value >= 0x2F800 && value <= 0x2FA1F)
                return true;

            // Kiểm tra CJK Radicals Supplement
            // Chứa các bộ thủ kanji (radical)
            if (value >= 0x2E80 && value <= 0x2EFF)
                return true;

            // Kiểm tra Kangxi Radicals
            // Bộ thủ Kangxi được sử dụng trong từ điển
            if (value >= 0x2F00 && value <= 0x2FDF)
                return true;
            if (value == 0x30FC)
                return true;
            return false;
        }
        public static string FormatKanjiWithBrackets(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            var result = new StringBuilder();
            bool isFirstChar = true;
            bool wasLastCharKanji = false;

            for (int i = 0; i < text.Length; i++)
            {
                char currentChar = text[i];
                bool isCurrentCharKanji = IsKanji(currentChar);

                if (isCurrentCharKanji)
                {
                    // Add space before kanji if it's not the first character
                    if (!isFirstChar && !wasLastCharKanji)
                    {
                        result.Append(' ');
                    }

                    // Add the kanji with brackets
                    result.Append(currentChar);
                    result.Append('[').Append(currentChar).Append(']');

                    // Add space after kanji if it's not the last character
                    if (i < text.Length - 1)
                    {
                        result.Append(' ');
                    }
                }
                else
                {
                    result.Append(currentChar);
                }

                isFirstChar = false;
                wasLastCharKanji = isCurrentCharKanji;
            }

            return result.ToString();
        }
    }

}
