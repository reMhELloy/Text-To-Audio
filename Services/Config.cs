using System;
using System.IO;
using System.Collections.Generic;

namespace Text_to_Image.Config
{
    // Config.cs - Lớp chứa tất cả các cấu hình
    public static class AppConfig
    {
        // Azure Speech Service Configuration
        public const string SPEECH_KEY = "E27koPDpJmKuZNgAWeua52IrX9dmirppcPmyJUE8JKRDmG2I1D6iJQQJ99BBACYeBjFXJ3w3AAAYACOGE5i9";
        public const string SPEECH_REGION = "eastus";

        // Default Paths
        public static readonly string DEFAULT_EXCEL_FOLDER = @"S:\Anki";
        public static readonly string DEFAULT_AUDIO_FOLDER = @"S:\Anki\Audio";
        public static readonly string BREAK_AUDIO_FOLDER = @"S:\Anki\BreakNew\TestBreakAudio";

        // Audio Settings
        public const float SILENCE_DURATION_SECONDS = 0.5f;
        public const string AUDIO_FORMAT = "mp3"; // Luôn luôn là mp3

        // Voice Settings cho từng ngôn ngữ
        public static readonly Dictionary<string, string> VOICE_MAPPING = new Dictionary<string, string>
        {
            ["vietnamese"] = "vi-VN-NamMinhNeural",
            ["english"] = "en-US-JennyNeural",
            ["japanese"] = "ja-JP-NanamiNeural",
            ["chinese"] = "zh-CN-XiaoxiaoNeural",
            ["korean"] = "ko-KR-SunHiNeural"
        };

        // File Prefixes
        public static readonly Dictionary<string, string> FILE_PREFIXES = new Dictionary<string, string>
        {
            ["tuvung"] = "Vocab",
            ["english"] = "EN",
            ["japanese"] = "JP",
            ["chinese"] = "ZH",
            ["korean"] = "KR"
        };
    }
}