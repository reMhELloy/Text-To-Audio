using System;
using OfficeOpenXml;
using System.IO;
using Text_to_Image.Services;
using Text_to_Image.Models;
using System.Threading.Tasks;
using Text_to_Image.Data;

public class Program
{
    // THÊM BIẾN ĐỂ CHỌN SPEECH SERVICE - THÊM GOOGLE TRANSLATE TTS
    private static bool useGoogleTTS = true; // false = Azure, true = Google
    private static bool useGoogleTranslateTTS = false; // THÊM MỚI - Google Translate TTS (Unofficial)
    private static readonly string GOOGLE_CREDENTIALS_PATH = @"S:\Anki\google-credentials.json";

    public static async Task Main()
    {
        bool continueRunning = true;
        while (continueRunning)
        {
            try
            {
                Console.WriteLine("============================================================");
                Console.WriteLine("LANGUAGE LEARNING TOOL");
                Console.WriteLine($"Speech Service: {GetCurrentServiceName()}");
                Console.WriteLine("============================================================");
                Console.WriteLine("1. Process Excel File (Normal Mode)");
                Console.WriteLine("2. Export from Database to Excel");
                Console.WriteLine("3. Switch Speech Service (Azure ↔ Google)");
                Console.WriteLine("0. Exit Program");
                Console.Write("Choose option (0-3): ");

                string mainChoice = Console.ReadLine()?.Trim();

                switch (mainChoice)
                {
                    case "1":
                        await ProcessExcelMode();
                        break;
                    case "2":
                        {
                            using var exportService = new DatabaseExportService();
                            await exportService.ShowExportMenu();
                        }
                        break;
                    case "3":
                        SwitchSpeechService();
                        break;
                    case "0":
                        continueRunning = false;
                        Console.WriteLine("Thanks for using Language Learning Tool!");
                        break;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.Write("Do you want to continue? (Y/N): ");
                string answer = Console.ReadLine()?.Trim().ToUpper();
                continueRunning = (answer == "Y");
            }
        }
    }

    // THÊM METHOD MỚI ĐỂ LẤY TÊN SERVICE HIỆN TẠI
    private static string GetCurrentServiceName()
    {
        if (useGoogleTranslateTTS) return "🟢 Google Translate TTS (Unofficial)";
        if (useGoogleTTS) return "🟢 Google Text-to-Speech";
        return "🔵 Azure Speech Service";
    }

    // THÊM METHOD MỚI ĐỂ SWITCH SPEECH SERVICE - THÊM OPTION CHO GOOGLE TRANSLATE TTS
    private static void SwitchSpeechService()
    {
        Console.WriteLine();
        Console.WriteLine("=== SPEECH SERVICE SELECTION ===");
        Console.WriteLine($"Current: {GetCurrentServiceName()}");
        Console.WriteLine();
        Console.WriteLine("1. 🔵 Azure Speech Service");
        Console.WriteLine("2. 🟢 Google Text-to-Speech");
        Console.WriteLine("3. 🟢 Google Translate TTS (FREE, Unofficial)"); // THÊM MỚI
        Console.WriteLine("0. Cancel");
        Console.Write("Choose (0-3): ");

        string choice = Console.ReadLine()?.Trim();

        switch (choice)
        {
            case "1":
                useGoogleTTS = false;
                useGoogleTranslateTTS = false;
                Console.WriteLine("✅ Switched to Azure Speech Service");
                break;
            case "2":
                if (File.Exists(GOOGLE_CREDENTIALS_PATH))
                {
                    useGoogleTTS = true;
                    useGoogleTranslateTTS = false;
                    Console.WriteLine("✅ Switched to Google Text-to-Speech");
                }
                else
                {
                    Console.WriteLine("❌ Google credentials not found!");
                    Console.WriteLine($"Please place credentials file at: {GOOGLE_CREDENTIALS_PATH}");
                }
                break;
            case "3": // THÊM MỚI - GOOGLE TRANSLATE TTS
                useGoogleTTS = false;
                useGoogleTranslateTTS = true;
                Console.WriteLine("✅ Switched to Google Translate TTS (Unofficial)");
                Console.WriteLine("⚠️ Warning: This uses unofficial API and may be blocked by Google");
                Console.WriteLine("⚠️ Warning: Use at your own risk - for production, use Azure or Google Cloud TTS");
                break;
            case "0":
                Console.WriteLine("Cancelled.");
                break;
            default:
                Console.WriteLine("Invalid choice.");
                break;
        }
        Console.WriteLine();
    }

    private static async Task ProcessExcelMode()
    {
        var options = new ProcessingOptions();

        // Step 1: File Selection
        if (!ExcelProcessor.SelectExcelFile(options))
        {
            Console.WriteLine("No file selected. Returning to main menu.");
            return;
        }

        // Step 2: Date Input
        Console.Write($"Enter date (format: day or day-month, current: {DateTime.Now:dd-MM-yyyy}): ");
        string dateInput = Console.ReadLine()?.Trim();
        var dateOptions = DateProcessor.ProcessDateInput(dateInput);

        options.CustomDate = dateOptions.CustomDate;
        options.SelectedDay = dateOptions.SelectedDay;
        options.SelectedMonth = dateOptions.SelectedMonth;
        options.SelectedYear = dateOptions.SelectedYear;

        // Step 3: Excel Processing - Open and wait for user to close
        if (!ExcelProcessor.OpenAndWaitForExcelFile(options))
        {
            Console.WriteLine("Unable to open Excel file...");
            return;
        }

        // Step 4: Get Excel Processing Inputs
        UserInputHandler.GetExcelProcessingInputs(options);

        // Step 5: Execute Excel Tasks with Database Integration
        Console.WriteLine("Processing Excel tasks with database integration...");
        await ExcelProcessor.ProcessExcelFile(options);
        Console.WriteLine("Excel processing completed.");

        // Step 6: Ask Audio Creation Question
        UserInputHandler.GetAudioCreationInputs(options);

        // Step 7: Execute Audio Creation
        if (options.CreateAudioFiles)
        {
            Console.WriteLine($"Creating audio files with {GetCurrentServiceName()}...");
            await ExecuteAudioTasks(options);
            Console.WriteLine("Audio creation completed.");
        }

        // Step 8: Ask and Handle Database Save
        await UserInputHandler.HandleDatabaseSave(options);

        // Step 9: Ask Audio Rename Question
        UserInputHandler.GetAudioRenameInputs(options);

        // Step 10: Execute Audio Rename
        if (options.RenameAudioFiles)
        {
            Console.WriteLine("Renaming audio files...");
            AudioFileManager.RenameAudioFiles(options.AudioFolderPath,
                options.SelectedDay.ToString(), options.SelectedMonth.ToString(),
                options.SelectedYear.ToString(), options.FileName);
            Console.WriteLine("Audio renaming completed.");
        }

        Console.WriteLine("Excel processing workflow completed!");
    }

    // CẬP NHẬT METHOD NÀY ĐỂ CHỌN SPEECH SERVICE - THÊM GOOGLE TRANSLATE TTS
    private static async Task ExecuteAudioTasks(ProcessingOptions options)
    {
        try
        {
            if (useGoogleTranslateTTS) // THÊM MỚI - GOOGLE TRANSLATE TTS
            {
                // SỬ DỤNG GOOGLE TRANSLATE TTS (UNOFFICIAL)
                var googleTranslateService = new GoogleTranslateTTSService();
                await googleTranslateService.ProcessExcelForAudio(options);
                Console.WriteLine("All audio files have been created successfully with Google Translate TTS.");
            }
            else if (useGoogleTTS)
            {
                // SỬ DỤNG GOOGLE TTS
                var googleService = new GoogleSpeechService(GOOGLE_CREDENTIALS_PATH);
                await googleService.ProcessExcelForAudio(options);
                Console.WriteLine("All audio files have been created successfully with Google TTS.");
            }
            else
            {
                // SỬ DỤNG AZURE TTS (CODE CŨ)
                var azureService = new AzureSpeechService(
                    Text_to_Image.Config.AppConfig.SPEECH_KEY,
                    Text_to_Image.Config.AppConfig.SPEECH_REGION
                );
                await azureService.ProcessExcelForAudio(options);
                Console.WriteLine("All audio files have been created successfully with Azure TTS.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating audio files: {ex.Message}");

            // THÊM FALLBACK LOGIC CHO GOOGLE TRANSLATE TTS
            if (useGoogleTranslateTTS)
            {
                Console.WriteLine();
                Console.Write("Google Translate TTS failed. Try with Google Cloud TTS instead? (Y/N): ");
                string fallback = Console.ReadLine()?.Trim().ToUpper();

                if (fallback == "Y" && File.Exists(GOOGLE_CREDENTIALS_PATH))
                {
                    Console.WriteLine("Switching to Google Cloud TTS as fallback...");
                    var googleService = new GoogleSpeechService(GOOGLE_CREDENTIALS_PATH);
                    await googleService.ProcessExcelForAudio(options);
                    Console.WriteLine("All audio files have been created successfully with Google Cloud TTS (fallback).");
                    return;
                }
                else
                {
                    Console.Write("Try with Azure TTS instead? (Y/N): ");
                    string azureFallback = Console.ReadLine()?.Trim().ToUpper();

                    if (azureFallback == "Y")
                    {
                        Console.WriteLine("Switching to Azure TTS as fallback...");
                        var azureService = new AzureSpeechService(
                            Text_to_Image.Config.AppConfig.SPEECH_KEY,
                            Text_to_Image.Config.AppConfig.SPEECH_REGION
                        );
                        await azureService.ProcessExcelForAudio(options);
                        Console.WriteLine("All audio files have been created successfully with Azure TTS (fallback).");
                        return;
                    }
                }
            }
            // NẾU GOOGLE FAIL, ĐỀ XUẤT CHUYỂN VỀ AZURE
            else if (useGoogleTTS)
            {
                Console.WriteLine();
                Console.Write("Google TTS failed. Try with Azure instead? (Y/N): ");
                string fallback = Console.ReadLine()?.Trim().ToUpper();

                if (fallback == "Y")
                {
                    Console.WriteLine("Switching to Azure TTS as fallback...");
                    var azureService = new AzureSpeechService(
                        Text_to_Image.Config.AppConfig.SPEECH_KEY,
                        Text_to_Image.Config.AppConfig.SPEECH_REGION
                    );
                    await azureService.ProcessExcelForAudio(options);
                    Console.WriteLine("All audio files have been created successfully with Azure TTS (fallback).");
                    return;
                }
                else // THÊM FALLBACK VỀ GOOGLE TRANSLATE TTS
                {
                    Console.Write("Try with Google Translate TTS instead? (Y/N): ");
                    string translateFallback = Console.ReadLine()?.Trim().ToUpper();

                    if (translateFallback == "Y")
                    {
                        Console.WriteLine("Switching to Google Translate TTS as fallback...");
                        var googleTranslateService = new GoogleTranslateTTSService();
                        await googleTranslateService.ProcessExcelForAudio(options);
                        Console.WriteLine("All audio files have been created successfully with Google Translate TTS (fallback).");
                        return;
                    }
                }
            }
            else // AZURE FAIL - THÊM FALLBACK VỀ GOOGLE TRANSLATE TTS
            {
                Console.WriteLine();
                Console.Write("Azure TTS failed. Try with Google Cloud TTS instead? (Y/N): ");
                string fallback = Console.ReadLine()?.Trim().ToUpper();

                if (fallback == "Y" && File.Exists(GOOGLE_CREDENTIALS_PATH))
                {
                    Console.WriteLine("Switching to Google Cloud TTS as fallback...");
                    var googleService = new GoogleSpeechService(GOOGLE_CREDENTIALS_PATH);
                    await googleService.ProcessExcelForAudio(options);
                    Console.WriteLine("All audio files have been created successfully with Google Cloud TTS (fallback).");
                    return;
                }
                else
                {
                    Console.Write("Try with Google Translate TTS instead? (Y/N): ");
                    string translateFallback = Console.ReadLine()?.Trim().ToUpper();

                    if (translateFallback == "Y")
                    {
                        Console.WriteLine("Switching to Google Translate TTS as fallback...");
                        var googleTranslateService = new GoogleTranslateTTSService();
                        await googleTranslateService.ProcessExcelForAudio(options);
                        Console.WriteLine("All audio files have been created successfully with Google Translate TTS (fallback).");
                        return;
                    }
                }
            }

            throw;
        }
    }
}