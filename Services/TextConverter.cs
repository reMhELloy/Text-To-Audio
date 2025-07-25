using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JapaneseConverter
{
    public static class TextConverter
    {
        public static string ConvertToImgTags(string japaneseText)
        {
            var result = new StringBuilder();

            for (int i = 0; i < japaneseText.Length; i++)
            {
                // Check for two-character combinations first (existing hiragana/katakana combo logic)
                if (i < japaneseText.Length - 1)
                {
                    string twoChars = japaneseText.Substring(i, 2);

                    if (JapanMapping.HiraganaComboMap.TryGetValue(twoChars, out var hiraganaCombo))
                    {
                        string firstChar = hiraganaCombo.Item1;
                        string secondChar = hiraganaCombo.Item2;

                        if (firstChar.StartsWith("dakuon_"))
                        {
                            result.Append($"<img src=\"hiragana_{firstChar}.png\">");
                        }
                        else
                        {
                            result.Append($"<img src=\"hiragana_{firstChar}.gif\">");
                        }
                        result.Append($"<img src=\"hiragana_{secondChar}.gif\">");
                        i++; // Skip next character as it's part of the combo
                        continue;
                    }

                    if (JapanMapping.KatakanaComboMap.TryGetValue(twoChars, out var katakanaCombo))
                    {
                        string firstChar = katakanaCombo.Item1;
                        string secondChar = katakanaCombo.Item2;

                        if (firstChar.StartsWith("dakuon_"))
                        {
                            result.Append($"<img src=\"katakana_{firstChar}.png\">");
                        }
                        else
                        {
                            result.Append($"<img src=\"katakana_{firstChar}.gif\">");
                        }
                        result.Append($"<img src=\"katakana_{secondChar}.gif\">");
                        i++; // Skip next character as it's part of the combo
                        continue;
                    }
                }

                // Process single character
                char c = japaneseText[i];

                // Check if the character is a kanji (falls within the CJK Unified Ideographs range)
                if (KanjiHelper.IsKanji(c))
                {
                    // Convert kanji to image tag with SVG
                    result.Append($"<img src=\"{c}.svg\"><img src=\"{c}.gif\">");
                }
                // Existing conversion logic for hiragana and katakana
                else if (JapanMapping.HiraganaDakuonMap.ContainsKey(c))
                {
                    result.Append($"<img src=\"hiragana_dakuon_{JapanMapping.HiraganaDakuonMap[c]}.png\">");
                }
                else if (JapanMapping.KatakanaDakuonMap.ContainsKey(c))
                {
                    result.Append($"<img src=\"katakana_dakuon_{JapanMapping.KatakanaDakuonMap[c]}.png\">");
                }
                else if (JapanMapping.KatakanaMap.ContainsKey(c))
                {
                    result.Append($"<img src=\"katakana_{JapanMapping.KatakanaMap[c]}.gif\">");
                }
                else if (JapanMapping.HiraganaMap.ContainsKey(c))
                {
                    result.Append($"<img src=\"hiragana_{JapanMapping.HiraganaMap[c]}.gif\">");
                }
                else
                {
                    result.Append(c);
                }
            }

            return result.ToString();
        }
    }
}
