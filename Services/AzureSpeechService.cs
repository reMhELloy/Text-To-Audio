using Microsoft.CognitiveServices.Speech;
using OfficeOpenXml;
using Text_to_Image.Models;

namespace Text_to_Image.Services
{
    public class AzureSpeechService
    {
        private readonly string _speechKey;
        private readonly string _speechRegion;
        private readonly SpeechConfig _speechConfig;
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
        public AzureSpeechService(string speechKey, string speechRegion)
        {
            _speechKey = speechKey;
            _speechRegion = speechRegion;
            _speechConfig = SpeechConfig.FromSubscription(_speechKey, _speechRegion);

            // Set output format to MP3
            _speechConfig.SetSpeechSynthesisOutputFormat(SpeechSynthesisOutputFormat.Audio16Khz32KBitRateMonoMp3);
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

        // Thay thế hoàn toàn method ProcessExcelForAudio trong AzureSpeechService
        // Thay thế hoàn toàn method ProcessExcelForAudio trong AzureSpeechService

        public async Task ProcessExcelForAudio(ProcessingOptions options)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var package = new ExcelPackage(new FileInfo(options.SelectedFile));
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension?.End.Row ?? 0;

                if (rowCount == 0)
                {
                    Console.WriteLine("No data found in Excel file.");
                    return;
                }

                // Tạo danh sách các audio files cần tạo - CHỈ CHO ROWS CÓ SOUND FORMULAS
                var audioFiles = GenerateAudioFileNamesFromFormulas(worksheet, options, rowCount);

                if (audioFiles.Count == 0)
                {
                    Console.WriteLine("No audio files to create based on current configuration.");
                    return;
                }

                Console.WriteLine($"Creating {audioFiles.Count} audio files with optimized settings for learners...");

                // Đọc text từ Excel và set đường dẫn output
                foreach (var audioFile in audioFiles)
                {
                    int colIndex = GetColumnIndex(audioFile.SourceColumn);
                    audioFile.SourceText = worksheet.Cells[audioFile.RowIndex, colIndex].Text?.Trim();
                    audioFile.OutputPath = Path.Combine(options.AudioOutputFolder, audioFile.FileName);
                }

                // Tạo âm thanh song song (giới hạn 3 files cùng lúc để tránh quá tải API)
                var semaphore = new SemaphoreSlim(3);
                var tasks = new List<Task>();

                foreach (var audioFile in audioFiles)
                {
                    tasks.Add(ProcessAudioFileTask(audioFile, semaphore));
                }

                await Task.WhenAll(tasks);
                semaphore.Dispose();

                int successCount = audioFiles.Count(af => !string.IsNullOrEmpty(af.SourceText));
                Console.WriteLine($"Completed! Created {successCount} audio files with learner-friendly settings.");

                await Task.Delay(500);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing Excel for audio: {ex.Message}");
            }
        }

        // THÊM CÁC METHODS MỚI VÀO AzureSpeechService:

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

        private List<AudioFileInfo> GenerateAudioFileNamesFromFormulas(ExcelWorksheet worksheet, ProcessingOptions options, int rowCount)
        {
            var audioFiles = new List<AudioFileInfo>();

            for (int row = 1; row <= rowCount; row++)
            {
                if (!CheckRowHasAudioFormulas(worksheet, row, options))
                {
                    Console.WriteLine($"Row {row}: Skipping audio creation (no sound formulas - duplicate entry)");
                    continue;
                }

                var rowAudioFiles = ExtractAudioFilesFromRow(worksheet, row, options);
                audioFiles.AddRange(rowAudioFiles);
            }

            return audioFiles;
        }

        private List<AudioFileInfo> ExtractAudioFilesFromRow(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            var audioFiles = new List<AudioFileInfo>();

            if ((options.FileName.Contains("tuvung") ||
                 options.FileName.Contains("japanese") ||
                 options.FileName.Contains("chinese")) && options.SoundColumns.Length == 2)
            {
                string col1Formula = worksheet.Cells[row, 5].Text?.Trim(); // Column E
                string col2Formula = worksheet.Cells[row, 6].Text?.Trim(); // Column F

                // CHỈ TÃO AUDIO CHO SOUND FORMULAS CÓ SẴN
                if (!string.IsNullOrEmpty(col1Formula) && col1Formula.Contains("[sound:"))
                {
                    string fileName = ExtractFileNameFromFormula(col1Formula);
                    // Xác định language và voice dựa trên filename prefix
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
                    // Xác định language và voice dựa trên filename prefix
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
                    audioFiles.Add(new AudioFileInfo
                    {
                        RowIndex = row,
                        SourceColumn = "B", // English column
                        FileName = fileName,
                        VoiceName = GetVoiceNameForLanguage("EN")
                    });
                }

                if (!string.IsNullOrEmpty(col2Formula) && col2Formula.Contains("[sound:"))
                {
                    string fileName = ExtractFileNameFromFormula(col2Formula);
                    audioFiles.Add(new AudioFileInfo
                    {
                        RowIndex = row,
                        SourceColumn = "B", // English column
                        FileName = fileName,
                        VoiceName = GetVoiceNameForLanguage("EN")
                    });
                }
            }

            return audioFiles;
        }

        // THÊM HELPER METHODS MỚI:
        private string ExtractLanguageFromFileName(string fileName)
        {
            // Extract language from filename: "JP-29-05-2025_02.mp3" -> "JP"
            if (fileName.Contains("-"))
            {
                return fileName.Split('-')[0];
            }
            return "EN"; // Default
        }

        private string GetSourceColumnForLanguage(string language, ProcessingOptions options)
        {
            if (options.FileName.Contains("tuvung"))
            {
                return language switch
                {
                    "EN" => "B", // English text
                    "VI" => "A", // Vietnamese text
                    _ => "A"
                };
            }
            else if (options.FileName.Contains("japanese"))
            {
                // JAPANESE: Cả 2 files JP đều từ các column khác nhau
                return language switch
                {
                    "JP" => "A", // Nếu là JP file đầu tiên (odd) → đọc English text
                    _ => "C"     // Nếu là JP file thứ hai (even) → đọc Japanese text
                };
            }
            else if (options.FileName.Contains("chinese"))
            {
                // CHINESE: Cả 2 files ZH đều từ các column khác nhau
                return language switch
                {
                    "ZH" => "A", // Nếu là ZH file đầu tiên (odd) → đọc English text
                    _ => "C"     // Nếu là ZH file thứ hai (even) → đọc Chinese text
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

        private async Task ProcessAudioFileTask(AudioFileInfo audioFile, SemaphoreSlim semaphore)
        {
            await semaphore.WaitAsync();
            try
            {
                if (!string.IsNullOrEmpty(audioFile.SourceText))
                {
                    await CreateAudioFile(audioFile.SourceText, audioFile.OutputPath, audioFile.VoiceName);
                }
            }
            finally
            {
                semaphore.Release();
            }
        }

        // Method để test SSML output (có thể dùng để debug)
        public string GetSSMLPreview(string text, string voiceName)
        {
            return CreateSSMLForLearners(text, voiceName);
        }
    }

}