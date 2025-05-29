using System;
using OfficeOpenXml;
using System.IO;
using Text_to_Image.Services;
using Text_to_Image.Models;
using System.Threading.Tasks;
using Text_to_Image.Data;

public class Program
{
    public static async Task Main()
    {
        bool continueRunning = true;
        while (continueRunning)
        {
            try
            {
                Console.WriteLine("============================================================");
                Console.WriteLine("LANGUAGE LEARNING TOOL");
                Console.WriteLine("============================================================");
                Console.WriteLine("1. Process Excel File (Normal Mode)");
                Console.WriteLine("2. Export from Database to Excel");
                Console.WriteLine("0. Exit Program");
                Console.Write("Choose option (0-2): ");

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
            Console.WriteLine("Creating audio files...");
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

    private static async Task ExecuteAudioTasks(ProcessingOptions options)
    {
        try
        {
            var speechService = new AzureSpeechService(
                Text_to_Image.Config.AppConfig.SPEECH_KEY,
                Text_to_Image.Config.AppConfig.SPEECH_REGION
            );

            await speechService.ProcessExcelForAudio(options);
            Console.WriteLine("All audio files have been created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating audio files: {ex.Message}");
            throw;
        }
    }
}