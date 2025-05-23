using System;
using System.Text;
using System.Collections.Generic;
using OfficeOpenXml;
using System.IO;
using Text_to_Image.Services;
using Text_to_Image.Models;

public class Program
{
    public static void Main()
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

                // Step 3: Excel Processing - ĐỔI TÊN METHOD
                if (!ExcelProcessor.OpenAndWaitForExcelFile(options))
                {
                    Console.WriteLine("Unable to open Excel file...");
                    continue;
                }

                // Step 4: Get User Inputs
                UserInputHandler.GetProcessingInputs(options);

                // Step 5: Execute Tasks
                ExcelProcessor.ExecuteTasks(options);

                // Step 6: Continue?
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



}
