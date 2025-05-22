using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JapaneseConverter
{
    public static class JapanMapping
    {
        public static readonly Dictionary<char, string> HiraganaDakuonMap = new Dictionary<char, string>
        {
            {'が', "ga"}, {'ぎ', "gi"}, {'ぐ', "gu"}, {'げ', "ge"}, {'ご', "go"},
            {'ざ', "za"}, {'じ', "ji"}, {'ず', "zu"}, {'ぜ', "ze"}, {'ぞ', "zo"},
            {'だ', "da"}, {'ぢ', "ji2"}, {'づ', "zu2"}, {'で', "de"}, {'ど', "do"},
            {'ば', "ba"}, {'び', "bi"}, {'ぶ', "bu"}, {'べ', "be"}, {'ぼ', "bo"},
            {'ぱ', "pa"}, {'ぴ', "pi"}, {'ぷ', "pu"}, {'ぺ', "pe"}, {'ぽ', "po"}
        };

        public static readonly Dictionary<char, string> KatakanaDakuonMap = new Dictionary<char, string>
        {
            {'ガ', "ga"}, {'ギ', "gi"}, {'グ', "gu"}, {'ゲ', "ge"}, {'ゴ', "go"},
            {'ザ', "za"}, {'ジ', "ji"}, {'ズ', "zu"}, {'ゼ', "ze"}, {'ゾ', "zo"},
            {'ダ', "da"}, {'ヂ', "ji2"}, {'ヅ', "zu2"}, {'デ', "de"}, {'ド', "do"},
            {'バ', "ba"}, {'ビ', "bi"}, {'ブ', "bu"}, {'ベ', "be"}, {'ボ', "bo"},
            {'パ', "pa"}, {'ピ', "pi"}, {'プ', "pu"}, {'ペ', "pe"}, {'ポ', "po"}
        };

        public static readonly Dictionary<char, string> KatakanaMap = new Dictionary<char, string>
        {
            {'ア', "a"}, {'イ', "i"}, {'ウ', "u"}, {'エ', "e"}, {'オ', "o"},
            {'カ', "ka"}, {'キ', "ki"}, {'ク', "ku"}, {'ケ', "ke"}, {'コ', "ko"},
            {'サ', "sa"}, {'シ', "shi"}, {'ス', "su"}, {'セ', "se"}, {'ソ', "so"},
            {'タ', "ta"}, {'チ', "chi"}, {'ツ', "tsu"}, {'テ', "te"}, {'ト', "to"},
            {'ナ', "na"}, {'ニ', "ni"}, {'ヌ', "nu"}, {'ネ', "ne"}, {'ノ', "no"},
            {'ハ', "ha"}, {'ヒ', "hi"}, {'フ', "fu"}, {'ヘ', "he"}, {'ホ', "ho"},
            {'マ', "ma"}, {'ミ', "mi"}, {'ム', "mu"}, {'メ', "me"}, {'モ', "mo"},
            {'ヤ', "ya"}, {'ユ', "yu"}, {'ヨ', "yo"},
            {'ラ', "ra"}, {'リ', "ri"}, {'ル', "ru"}, {'レ', "re"}, {'ロ', "ro"},
            {'ワ', "wa"}, {'ヲ', "wo"}, {'ン', "n"}, {'ッ', "tsu"},
            // Thêm các ký tự nhỏ Katakana
            {'ィ', "i"},
            {'ェ', "e"},
            {'ォ', "o"},
            {'ャ', "ya"},
            {'ュ', "yu"},
            {'ョ', "yo"}
        };

        public static readonly Dictionary<char, string> HiraganaMap = new Dictionary<char, string>
        {
            {'あ', "a"}, {'い', "i"}, {'う', "u"}, {'え', "e"}, {'お', "o"},
            {'か', "ka"}, {'き', "ki"}, {'く', "ku"}, {'け', "ke"}, {'こ', "ko"},
            {'さ', "sa"}, {'し', "shi"}, {'す', "su"}, {'せ', "se"}, {'そ', "so"},
            {'た', "ta"}, {'ち', "chi"}, {'つ', "tsu"}, {'て', "te"}, {'と', "to"},
            {'な', "na"}, {'に', "ni"}, {'ぬ', "nu"}, {'ね', "ne"}, {'の', "no"},
            {'は', "ha"}, {'ひ', "hi"}, {'ふ', "fu"}, {'へ', "he"}, {'ほ', "ho"},
            {'ま', "ma"}, {'み', "mi"}, {'む', "mu"}, {'め', "me"}, {'も', "mo"},
            {'や', "ya"}, {'ゆ', "yu"}, {'よ', "yo"},
            {'ら', "ra"}, {'り', "ri"}, {'る', "ru"}, {'れ', "re"}, {'ろ', "ro"},
            {'わ', "wa"}, {'を', "wo"}, {'ん', "n"}, {'っ', "tsu"},
            // Thêm các ký tự nhỏ Hiragana
            {'ぃ', "i"},
            {'ぇ', "e"},
            {'ぉ', "o"},
            {'ゃ', "ya"},
            {'ゅ', "yu"},
            {'ょ', "yo"}
        };
        public static readonly Dictionary<string, (string, string)> HiraganaComboMap = new Dictionary<string, (string, string)>
        {
            {"きゃ", ("ki", "ya")}, {"きゅ", ("ki", "yu")}, {"きょ", ("ki", "yo")},
            {"ぎゃ", ("dakuon_gi", "ya")}, {"ぎゅ", ("dakuon_gi", "yu")}, {"ぎょ", ("dakuon_gi", "yo")},
            {"しゃ", ("shi", "ya")}, {"しゅ", ("shi", "yu")}, {"しょ", ("shi", "yo")},
            {"じゃ", ("dakuon_ji", "ya")}, {"じゅ", ("dakuon_ji", "yu")}, {"じょ", ("dakuon_ji", "yo")},
            {"ちゃ", ("chi", "ya")}, {"ちゅ", ("chi", "yu")}, {"ちょ", ("chi", "yo")},
            {"にゃ", ("ni", "ya")}, {"にゅ", ("ni", "yu")}, {"にょ", ("ni", "yo")},
            {"ひゃ", ("hi", "ya")}, {"ひゅ", ("hi", "yu")}, {"ひょ", ("hi", "yo")},
            {"びゃ", ("dakuon_bi", "ya")}, {"びゅ", ("dakuon_bi", "yu")}, {"びょ", ("dakuon_bi", "yo")},
            {"ぴゃ", ("dakuon_pi", "ya")}, {"ぴゅ", ("dakuon_pi", "yu")}, {"ぴょ", ("dakuon_pi", "yo")},
            {"みゃ", ("mi", "ya")}, {"みゅ", ("mi", "yu")}, {"みょ", ("mi", "yo")},
            {"りゃ", ("ri", "ya")}, {"りゅ", ("ri", "yu")}, {"りょ", ("ri", "yo")}
        };

        public static readonly Dictionary<string, (string, string)> KatakanaComboMap = new Dictionary<string, (string, string)>
        {
            {"キャ", ("ki", "ya")}, {"キュ", ("ki", "yu")}, {"キョ", ("ki", "yo")},
            {"ギャ", ("dakuon_gi", "ya")}, {"ギュ", ("dakuon_gi", "yu")}, {"ギョ", ("dakuon_gi", "yo")},
            {"シャ", ("shi", "ya")}, {"シュ", ("shi", "yu")}, {"ショ", ("shi", "yo")},
            {"ジャ", ("dakuon_ji", "ya")}, {"ジュ", ("dakuon_ji", "yu")}, {"ジョ", ("dakuon_ji", "yo")},
            {"チャ", ("chi", "ya")}, {"チュ", ("chi", "yu")}, {"チョ", ("chi", "yo")},
            {"ニャ", ("ni", "ya")}, {"ニュ", ("ni", "yu")}, {"ニョ", ("ni", "yo")},
            {"ヒャ", ("hi", "ya")}, {"ヒュ", ("hi", "yu")}, {"ヒョ", ("hi", "yo")},
            {"ビャ", ("dakuon_bi", "ya")}, {"ビュ", ("dakuon_bi", "yu")}, {"ビョ", ("dakuon_bi", "yo")},
            {"ピャ", ("dakuon_pi", "ya")}, {"ピュ", ("dakuon_pi", "yu")}, {"ピョ", ("dakuon_pi", "yo")},
            {"ミャ", ("mi", "ya")}, {"ミュ", ("mi", "yu")}, {"ミョ", ("mi", "yo")},
            {"リャ", ("ri", "ya")}, {"リュ", ("ri", "yu")}, {"リョ", ("ri", "yo")}
        };
    }
}
