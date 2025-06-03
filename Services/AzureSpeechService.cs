using Microsoft.CognitiveServices.Speech;
using OfficeOpenXml;
using Text_to_Image.Models;
using Text_to_Image.Data.Models;
using Text_to_Image.Services;
using Text_to_Image.Data;

namespace Text_to_Image.Services
{
    public class AzureSpeechService
    {
        private readonly string _speechKey;
        private readonly string _speechRegion;
        private readonly SpeechConfig _speechConfig;

        public AzureSpeechService(string speechKey, string speechRegion)
        {
            _speechKey = speechKey;
            _speechRegion = speechRegion;
            _speechConfig = SpeechConfig.FromSubscription(_speechKey, _speechRegion);

            // Set output format to MP3
            _speechConfig.SetSpeechSynthesisOutputFormat(SpeechSynthesisOutputFormat.Audio16Khz32KBitRateMonoMp3);
        }

        private int GetColumnIndex(string columnLetter)
        {
            if (string.IsNullOrEmpty(columnLetter))
                return 1;

            int result = 0;
            for (int i = 0; i < columnLetter.Length; i++)
            {
                result = result * 26 + (columnLetter[i] - 'A' + 1);
            }
            return result;
        }

        // Tạo file âm thanh từ text với giọng nói tùy chọn - CÓ CHẤT LƯỢNG CHO NGƯỜI MỚI HỌC
        public async Task<bool> CreateAudioFile(string text, string outputPath, string voiceName = "en-US-JennyNeural")
        {
            try
            {
                if (string.IsNullOrWhiteSpace(text))
                {
                    return false;
                }

                _speechConfig.SpeechSynthesisVoiceName = voiceName;

                // Tạo thư mục nếu chưa tồn tại
                string directory = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Tạo SSML với cấu hình phù hợp cho người mới học
                string ssmlText = CreateSSMLForLearners(text, voiceName);

                // Sử dụng synthesizer mà không cần AudioConfig để lấy raw audio data
                using var synthesizer = new SpeechSynthesizer(_speechConfig, null);

                var result = await synthesizer.SpeakSsmlAsync(ssmlText);

                if (result.Reason == ResultReason.SynthesizingAudioCompleted)
                {
                    // Lưu audio data trực tiếp thành file MP3
                    await File.WriteAllBytesAsync(outputPath, result.AudioData);
                    Console.WriteLine($"✓ {Path.GetFileName(outputPath)}");
                    return true;
                }
                else if (result.Reason == ResultReason.Canceled)
                {
                    var cancellation = SpeechSynthesisCancellationDetails.FromResult(result);
                    Console.WriteLine($"✗ {Path.GetFileName(outputPath)}: {cancellation.Reason}");
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ {Path.GetFileName(outputPath)}: {ex.Message}");
                return false;
            }
        }

        // Tạo SSML với cấu hình tối ưu cho người mới học
        private string CreateSSMLForLearners(string text, string voiceName)
        {
            // Xác định tốc độ đọc và style dựa trên giọng nói
            string rate = "1.0"; // Tốc độ bình thường cho tiếng Việt
            string style = "";

            // Cấu hình tùy chỉnh cho từng ngôn ngữ
            if (voiceName.Contains("en-US") || voiceName.Contains("en-GB"))
            {
                rate = "0.75"; // Tiếng Anh đọc chậm hơn nữa
                style = @"style=""calm"""; // Giọng điềm tĩnh cho tiếng Anh
            }
            else if (voiceName.Contains("ja-JP"))
            {
                rate = "0.7"; // Tiếng Nhật đọc rất chậm
                style = @"style=""calm"""; // Giọng điềm tĩnh
            }
            else if (voiceName.Contains("zh-CN") || voiceName.Contains("zh-TW"))
            {
                rate = "0.7"; // Tiếng Trung đọc rất chậm
                style = @"style=""calm"""; // Giọng điềm tĩnh
            }
            else if (voiceName.Contains("vi-VN"))
            {
                rate = "1.0"; // Tiếng Việt đọc bình thường - KHÔNG CHẬM
                style = @"style=""calm"""; // Giọng điềm tĩnh
            }

            // Escape XML characters trong text
            string escapedText = System.Security.SecurityElement.Escape(text);

            // Tạo SSML với silent đầu/cuối và tốc độ phù hợp
            string ssml = $@"<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='en-US'>
                <voice name='{voiceName}'>
                <break time='200ms'/>
                <prosody rate='{rate}' {style}>
                {escapedText}
                </prosody>
                <break time='300ms'/>
                </voice>
                </speak>";

            return ssml;
        }

        // MAIN METHOD: Process Excel for Audio với logic duplicate mới
        public async Task ProcessExcelForAudio(ProcessingOptions options)
        {
            try
            {
                Console.WriteLine("Creating audio files...");

                string dateToUse = string.IsNullOrWhiteSpace(options.CustomDate) ?
                    DateTime.Now.ToString("dd-MM-yyyy") : options.CustomDate;

                // Lấy existing vocabularies từ database để check duplicates
                List<Vocabulary> existingVocabs = new List<Vocabulary>();
                try
                {
                    using var dbService = new DatabaseService();
                    string fileType = DetermineFileType(options.FileName);
                    existingVocabs = await dbService.GetExistingVocabulariesForDateAsync(dateToUse, fileType);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not load existing vocabularies: {ex.Message}");
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var package = new ExcelPackage(new FileInfo(options.SelectedFile));
                var worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension?.End.Row ?? 0;
                if (rowCount == 0)
                {
                    Console.WriteLine("No data in Excel file to process.");
                    return;
                }

                var audioTasks = new List<Task>();
                var processedInSession = new List<Vocabulary>();

                for (int row = 1; row <= rowCount; row++)
                {
                    // Tạo vocabulary từ row hiện tại
                    var currentVocab = CreateVocabularyFromRow(worksheet, row, options);

                    // Check duplicate - Database hoặc Current Session
                    var duplicateResult = CheckRowForDuplicateAudio(currentVocab, existingVocabs, processedInSession, options);

                    if (duplicateResult.IsDuplicate && duplicateResult.DuplicateSource == "Database")
                    {
                        // DATABASE DUPLICATE: Tạo audio với tên file CŨ nhưng nội dung MỚI
                        await CreateAudioForDatabaseDuplicate(worksheet, row, duplicateResult.ExistingVocab, options, audioTasks);
                    }
                    else if (duplicateResult.IsDuplicate && duplicateResult.DuplicateSource == "CurrentSession")
                    {
                        // CURRENT SESSION DUPLICATE: Skip tạo audio
                        Console.WriteLine($"Row {row}: Skipping audio creation (duplicate in current session)");
                    }
                    else
                    {
                        // NEW ENTRY: Tạo audio với tên file MỚI
                        await CreateAudioForNewEntry(worksheet, row, options, audioTasks);
                        processedInSession.Add(currentVocab);
                    }
                }

                // Execute tất cả audio tasks
                if (audioTasks.Count > 0)
                {
                    Console.WriteLine($"Creating {audioTasks.Count} audio files with optimized settings for learners...");

                    // Tạo âm thanh song song (giới hạn 3 files cùng lúc để tránh quá tải API)
                    var semaphore = new SemaphoreSlim(3);
                    var tasks = audioTasks.Select(async task =>
                    {
                        await semaphore.WaitAsync();
                        try
                        {
                            await task;
                        }
                        finally
                        {
                            semaphore.Release();
                        }
                    });

                    await Task.WhenAll(tasks);
                    semaphore.Dispose();

                    Console.WriteLine($"Completed! Created {audioTasks.Count} audio files with learner-friendly settings.");
                }
                else
                {
                    Console.WriteLine("No audio files to create.");
                }

                Console.WriteLine("All audio files have been created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating audio files: {ex.Message}");
                throw;
            }
        }

        // THÊM MỚI: Tạo audio cho database duplicate với tên file cũ
        private async Task CreateAudioForDatabaseDuplicate(ExcelWorksheet worksheet, int row, Vocabulary existingVocab, ProcessingOptions options, List<Task> audioTasks)
        {
            try
            {
                // Lấy audio files cũ từ database
                using var dbService = new DatabaseService();
                var existingAudioFiles = await dbService.GetAudioFilesByVocabIdAsync(existingVocab.VocabId);

                if (existingAudioFiles == null || existingAudioFiles.Count == 0)
                {
                    Console.WriteLine($"Row {row}: No existing audio files found - skipping audio creation");
                    return;
                }

                Console.WriteLine($"Row {row}: Creating audio with OLD filenames but NEW content");

                // Tạo audio cho từng file cũ với nội dung mới
                foreach (var audioFile in existingAudioFiles)
                {
                    string newText = GetTextForAudio(worksheet, row, audioFile.Language, options);
                    if (!string.IsNullOrEmpty(newText))
                    {
                        string outputPath = Path.Combine(options.AudioOutputFolder, audioFile.FileName);
                        var task = CreateAudioFile(newText, outputPath, audioFile.VoiceName);
                        audioTasks.Add(task);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Row {row}: Error creating audio for duplicate - {ex.Message}");
            }
        }

        // THÊM MỚI: Tạo audio cho new entry với tên file mới
        private async Task CreateAudioForNewEntry(ExcelWorksheet worksheet, int row, ProcessingOptions options, List<Task> audioTasks)
        {
            // Check nếu row có sound formulas
            if (!CheckRowHasAudioFormulas(worksheet, row, options))
            {
                Console.WriteLine($"Row {row}: Skipping audio creation (no sound formulas)");
                return;
            }

            // Lấy sound formulas từ Excel
            var soundFormulas = ExtractAudioFilesFromRow(worksheet, row, options);

            foreach (var audioInfo in soundFormulas)
            {
                string text = GetTextForAudio(worksheet, row, ExtractLanguageFromFileName(audioInfo.FileName), options);
                if (!string.IsNullOrEmpty(text))
                {
                    string outputPath = Path.Combine(options.AudioOutputFolder, audioInfo.FileName);
                    var task = CreateAudioFile(text, outputPath, audioInfo.VoiceName);
                    audioTasks.Add(task);
                }
            }
        }

        // THÊM MỚI: Lấy text để tạo audio dựa trên language
        private string GetTextForAudio(ExcelWorksheet worksheet, int row, string language, ProcessingOptions options)
        {
            string sourceColumn = GetSourceColumnForLanguage(language, options);
            int colIndex = GetColumnIndex(sourceColumn);
            return worksheet.Cells[row, colIndex].Text?.Trim() ?? "";
        }

        // THÊM MỚI: Tạo vocabulary từ Excel row
        private Vocabulary CreateVocabularyFromRow(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            var vocab = new Vocabulary();
            string fileType = DetermineFileType(options.FileName);

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
                    break;
                case "chinese":
                    vocab.EnglishText = GetCellValue(worksheet, row, 1);
                    vocab.ReadingText = GetCellValue(worksheet, row, 2);
                    vocab.ChineseText = GetCellValue(worksheet, row, 3);
                    break;
                case "tuvung":
                    vocab.VietnameseText = GetCellValue(worksheet, row, 1);
                    vocab.ReadingText = GetCellValue(worksheet, row, 2);
                    vocab.JapaneseText = GetCellValue(worksheet, row, 3);
                    break;
            }

            return vocab;
        }

        // THÊM MỚI: Check duplicate cho audio processing
        private DuplicateCheckResult CheckRowForDuplicateAudio(Vocabulary currentVocab, List<Vocabulary> existingVocabs, List<Vocabulary> processedInSession, ProcessingOptions options)
        {
            // CHECK 1: Database same day
            var duplicateInDatabase = existingVocabs.FirstOrDefault(existing =>
                SimilarText(existing.EnglishText, currentVocab.EnglishText) ||
                SimilarText(existing.VietnameseText, currentVocab.VietnameseText) ||
                SimilarText(existing.JapaneseText, currentVocab.JapaneseText) ||
                SimilarText(existing.ChineseText, currentVocab.ChineseText)
            );

            if (duplicateInDatabase != null)
            {
                return new DuplicateCheckResult
                {
                    IsDuplicate = true,
                    ExistingVocab = duplicateInDatabase,
                    DuplicateSource = "Database"
                };
            }

            // CHECK 2: Current session
            var duplicateInCurrentSession = processedInSession.FirstOrDefault(processed =>
                SimilarText(processed.EnglishText, currentVocab.EnglishText) ||
                SimilarText(processed.VietnameseText, currentVocab.VietnameseText) ||
                SimilarText(processed.JapaneseText, currentVocab.JapaneseText) ||
                SimilarText(processed.ChineseText, currentVocab.ChineseText)
            );

            if (duplicateInCurrentSession != null)
            {
                return new DuplicateCheckResult
                {
                    IsDuplicate = true,
                    ExistingVocab = duplicateInCurrentSession,
                    DuplicateSource = "CurrentSession"
                };
            }

            return new DuplicateCheckResult { IsDuplicate = false };
        }

        // ORIGINAL METHOD: Check row có audio formulas không
        private bool CheckRowHasAudioFormulas(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            if (string.IsNullOrWhiteSpace(options.SoundColumns)) return false;

            if ((options.FileName.Contains("tuvung") ||
                 options.FileName.Contains("japanese") ||
                 options.FileName.Contains("chinese")) && options.SoundColumns.Length == 2)
            {
                string col1Value = worksheet.Cells[row, 5].Text?.Trim();
                string col2Value = worksheet.Cells[row, 6].Text?.Trim();

                return !string.IsNullOrEmpty(col1Value) && !string.IsNullOrEmpty(col2Value) &&
                       col1Value.Contains("[sound:") && col2Value.Contains("[sound:");
            }
            else if (options.FileName.Contains("english") && options.SoundColumns.Length == 2)
            {
                int soundCol1 = options.SoundColumns[0] - 'A' + 1;
                int soundCol2 = options.SoundColumns[1] - 'A' + 1;

                string col1Value = worksheet.Cells[row, soundCol1].Text?.Trim();
                string col2Value = worksheet.Cells[row, soundCol2].Text?.Trim();

                return !string.IsNullOrEmpty(col1Value) && !string.IsNullOrEmpty(col2Value) &&
                       col1Value.Contains("[sound:") && col2Value.Contains("[sound:");
            }
            else if (options.SoundColumns.Length == 1)
            {
                int soundCol = options.SoundColumns[0] - 'A' + 1;
                string colValue = worksheet.Cells[row, soundCol].Text?.Trim();

                return !string.IsNullOrEmpty(colValue) && colValue.Contains("[sound:");
            }

            return false;
        }

        // ORIGINAL METHOD: Extract audio files từ row
        private List<AudioFileInfo> ExtractAudioFilesFromRow(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            var audioFiles = new List<AudioFileInfo>();

            if ((options.FileName.Contains("tuvung") ||
                 options.FileName.Contains("japanese") ||
                 options.FileName.Contains("chinese")) && options.SoundColumns.Length == 2)
            {
                string col1Formula = worksheet.Cells[row, 5].Text?.Trim(); // Column E
                string col2Formula = worksheet.Cells[row, 6].Text?.Trim(); // Column F

                if (!string.IsNullOrEmpty(col1Formula) && col1Formula.Contains("[sound:"))
                {
                    string fileName = ExtractFileNameFromFormula(col1Formula);
                    string lang = ExtractLanguageFromFileName(fileName);
                    string sourceCol = GetSourceColumnForLanguage(lang, options);

                    audioFiles.Add(new AudioFileInfo
                    {
                        RowIndex = row,
                        SourceColumn = sourceCol,
                        FileName = fileName,
                        VoiceName = GetVoiceNameForLanguage(lang)
                    });
                }

                if (!string.IsNullOrEmpty(col2Formula) && col2Formula.Contains("[sound:"))
                {
                    string fileName = ExtractFileNameFromFormula(col2Formula);
                    string lang = ExtractLanguageFromFileName(fileName);
                    string sourceCol = GetSourceColumnForLanguage(lang, options);

                    audioFiles.Add(new AudioFileInfo
                    {
                        RowIndex = row,
                        SourceColumn = sourceCol,
                        FileName = fileName,
                        VoiceName = GetVoiceNameForLanguage(lang)
                    });
                }
            }
            else if (options.FileName.Contains("english") && options.SoundColumns.Length == 2)
            {
                int soundCol1 = options.SoundColumns[0] - 'A' + 1;
                int soundCol2 = options.SoundColumns[1] - 'A' + 1;

                string col1Formula = worksheet.Cells[row, soundCol1].Text?.Trim();
                string col2Formula = worksheet.Cells[row, soundCol2].Text?.Trim();

                if (!string.IsNullOrEmpty(col1Formula) && col1Formula.Contains("[sound:"))
                {
                    string fileName = ExtractFileNameFromFormula(col1Formula);
                    string lang = ExtractLanguageFromFileName(fileName);
                    string sourceCol = GetSourceColumnForLanguage(lang, options);

                    audioFiles.Add(new AudioFileInfo
                    {
                        RowIndex = row,
                        SourceColumn = sourceCol,
                        FileName = fileName,
                        VoiceName = GetVoiceNameForLanguage(lang)
                    });
                }

                if (!string.IsNullOrEmpty(col2Formula) && col2Formula.Contains("[sound:"))
                {
                    string fileName = ExtractFileNameFromFormula(col2Formula);
                    string lang = ExtractLanguageFromFileName(fileName);
                    string sourceCol = GetSourceColumnForLanguage(lang, options);

                    audioFiles.Add(new AudioFileInfo
                    {
                        RowIndex = row,
                        SourceColumn = sourceCol,
                        FileName = fileName,
                        VoiceName = GetVoiceNameForLanguage(lang)
                    });
                }
            }

            return audioFiles;
        }

        // FIXED: Extract language ĐÚNG cho TẤT CẢ trường hợp
        private string ExtractLanguageFromFileName(string fileName)
        {
            // Extract language from filename based on odd/even number
            if (fileName.Contains("_"))
            {
                var parts = fileName.Split('_');
                if (parts.Length >= 2)
                {
                    var numberPart = parts[1].Split('.')[0];
                    if (int.TryParse(numberPart, out int number))
                    {
                        bool isOdd = (number % 2 == 1);

                        // Xác định language dựa vào file type + odd/even
                        if (fileName.StartsWith("JP-"))
                        {
                            // JAPANESE: Lẻ = EN voice, Chẵn = JP voice
                            return isOdd ? "EN" : "JP";
                        }
                        else if (fileName.StartsWith("ZH-"))
                        {
                            // CHINESE: Lẻ = EN voice, Chẵn = ZH voice  
                            return isOdd ? "EN" : "ZH";
                        }
                        else if (fileName.StartsWith("Vocab-"))
                        {
                            // TUVUNG: Lẻ = EN voice, Chẵn = VI voice
                            return isOdd ? "EN" : "VI";
                        }
                        else if (fileName.StartsWith("EN-"))
                        {
                            // ENGLISH: Lẻ = VI voice, Chẵn = EN voice (NGƯỢC với các loại khác)
                            return isOdd ? "VI" : "EN";
                        }
                    }
                }
            }

            return "EN"; // Default
        }

        // FIXED: GetSourceColumnForLanguage cho TẤT CẢ trường hợp
        private string GetSourceColumnForLanguage(string language, ProcessingOptions options)
        {
            if (options.FileName.Contains("tuvung"))
            {
                // TUVUNG: Cả EN và VI voice đều đọc Vietnamese text từ Column A
                return "A";
            }
            else if (options.FileName.Contains("japanese"))
            {
                // JAPANESE LOGIC CŨ: EN voice đọc English (A), JP voice đọc Japanese (C)
                return language switch
                {
                    "EN" => "A", // English voice đọc English text từ Column A
                    "JP" => "C", // Japanese voice đọc Japanese text từ Column C  
                    _ => "A"
                };
            }
            else if (options.FileName.Contains("chinese"))
            {
                // CHINESE LOGIC CŨ: EN voice đọc English (A), ZH voice đọc Chinese (C)
                return language switch
                {
                    "EN" => "A", // English voice đọc English text từ Column A
                    "ZH" => "C", // Chinese voice đọc Chinese text từ Column C
                    _ => "A"
                };
            }
            else if (options.FileName.Contains("english"))
            {
                // ENGLISH LOGIC CŨ: VI voice đọc Vietnamese (A), EN voice đọc English (B)
                return language switch
                {
                    "VI" => "A", // Vietnamese voice đọc Vietnamese text từ Column A
                    "EN" => "B", // English voice đọc English text từ Column B
                    _ => "A"
                };
            }

            return "A"; // Default
        }

        private string GetVoiceNameForLanguage(string language)
        {
            return language switch
            {
                "EN" => "en-US-JennyNeural",
                "VI" => "vi-VN-HoaiMyNeural",
                "JP" => "ja-JP-NanamiNeural",
                "ZH" => "zh-CN-XiaoxiaoNeural",
                _ => "en-US-JennyNeural"
            };
        }

        private string ExtractFileNameFromFormula(string formula)
        {
            int start = formula.IndexOf("[sound:") + 7;
            int end = formula.IndexOf("]", start);
            if (start > 6 && end > start)
            {
                return formula.Substring(start, end - start);
            }
            return "";
        }

        // THÊM MỚI: Helper methods
        private string GetCellValue(ExcelWorksheet worksheet, int row, int column)
        {
            var value = worksheet.Cells[row, column].Text?.Trim();
            return string.IsNullOrEmpty(value) ? "" : value;
        }

        private bool SimilarText(string text1, string text2)
        {
            if (string.IsNullOrWhiteSpace(text1) || string.IsNullOrWhiteSpace(text2))
                return false;
            return text1.Trim().Equals(text2.Trim(), StringComparison.OrdinalIgnoreCase);
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

        // Method để test SSML output (có thể dùng để debug)
        public string GetSSMLPreview(string text, string voiceName)
        {
            return CreateSSMLForLearners(text, voiceName);
        }

        // Helper classes
        public class DuplicateCheckResult
        {
            public bool IsDuplicate { get; set; }
            public Vocabulary ExistingVocab { get; set; }
            public string DuplicateSource { get; set; } // "Database" or "CurrentSession"
        }
    }
}