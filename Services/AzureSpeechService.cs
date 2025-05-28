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

        // Xử lý tạo âm thanh từ Excel
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

                // Tạo danh sách các audio files cần tạo
                var audioFiles = AudioFileManager.GenerateAudioFileNames(options, rowCount);

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

                // Đợi TẤT CẢ tasks hoàn thành trước khi tiếp tục
                await Task.WhenAll(tasks);

                // Đảm bảo semaphore đã được release hết
                semaphore.Dispose();

                int successCount = audioFiles.Count(af => !string.IsNullOrEmpty(af.SourceText));
                Console.WriteLine($"Completed! Created {successCount} audio files with learner-friendly settings.");

                // Thêm delay nhỏ để đảm bảo tất cả file operations hoàn thành
                await Task.Delay(500);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing Excel for audio: {ex.Message}");
            }
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