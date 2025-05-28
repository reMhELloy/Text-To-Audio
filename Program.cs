using System;
using System.Text;
using System.Collections.Generic;
using OfficeOpenXml;
using System.IO;
using Text_to_Image.Services;
using Text_to_Image.Models;
using System.Threading.Tasks;

public class Program
{
    public static async Task Main()
    {
        bool continueRunning = true;
        while (continueRunning)
        {
            try
            {
                var options = new ProcessingOptions();

                // Step 1: File Selection
                if (!ExcelProcessor.SelectExcelFile(options))
                {
                    Console.WriteLine("The program has ended because no file was selected.");
                    return;
                }

                // Step 2: Date Input
                Console.Write($"\nEnter date (format: day or day-month, current: {DateTime.Now:dd-MM-yyyy}): ");
                string dateInput = Console.ReadLine()?.Trim();
                var dateOptions = DateProcessor.ProcessDateInput(dateInput);

                // Copy date information to main options
                options.CustomDate = dateOptions.CustomDate;
                options.SelectedDay = dateOptions.SelectedDay;
                options.SelectedMonth = dateOptions.SelectedMonth;
                options.SelectedYear = dateOptions.SelectedYear;

                // Step 3: Excel Processing - Open and wait for user to close
                if (!ExcelProcessor.OpenAndWaitForExcelFile(options))
                {
                    Console.WriteLine("Unable to open Excel file...");
                    continue;
                }

                // Step 4: Get User Inputs (ALL inputs including audio)
                UserInputHandler.GetProcessingInputs(options);

                // Step 5: Execute Excel Tasks FIRST - Đợi hoàn thành
                Console.WriteLine("Processing Excel tasks...");
                ExcelProcessor.ExecuteTasks(options);
                Console.WriteLine("Excel processing completed.\n");

                // Step 6: Execute Audio Tasks (if selected) - Đợi hoàn thành
                if (options.CreateAudioFiles)
                {
                    Console.WriteLine("Step 1: Creating audio files...");
                    await ExecuteAudioTasks(options);
                    Console.WriteLine("Step 1: Audio creation completed.\n");
                }

                // Step 7: Execute Audio Rename (if selected) - Đợi hoàn thành
                if (options.RenameAudioFiles)
                {
                    Console.WriteLine("Step 2: Renaming audio files...");
                    AudioFileManager.RenameAudioFiles(options.AudioFolderPath,
                        options.SelectedDay.ToString(), options.SelectedMonth.ToString(),
                        options.SelectedYear.ToString(), options.FileName);
                    Console.WriteLine("Step 2: Audio renaming completed.\n");
                }

                // Step 7: Continue?
                continueRunning = ExcelProcessor.AskToContinue();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError: {ex.Message}");
                Console.Write("\nDo you want to try again? (Y/N): ");
                string answer = Console.ReadLine()?.Trim().ToUpper();
                continueRunning = (answer == "Y");
            }
        }
    }

    // Method để xử lý audio tasks - Đợi hoàn thành
    private static async Task ExecuteAudioTasks(ProcessingOptions options)
    {
        try
        {
            // Sử dụng Azure Speech Service từ config
            var speechService = new AzureSpeechService(
                Text_to_Image.Config.AppConfig.SPEECH_KEY,
                Text_to_Image.Config.AppConfig.SPEECH_REGION
            );

            // Đợi audio processing hoàn toàn hoàn thành
            await speechService.ProcessExcelForAudio(options);

            // Đảm bảo tất cả files đã được tạo xong
            Console.WriteLine("All audio files have been created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating audio files: {ex.Message}");
            throw; // Re-throw để dừng workflow nếu có lỗi
        }
    }
}