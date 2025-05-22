using System;
using System.Text;
using System.Collections.Generic;
using OfficeOpenXml;
using System.IO;
using Text_to_Image.Services;

public class Program
{
    public static void Main()
    {
        string folderPath = @"S:\Anki";
        bool continueRunning = true;

        while (continueRunning)
        {
            try
            {
                string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsm");

                Console.WriteLine("Available Excel files:");
                for (int i = 0; i < excelFiles.Length; i++)
                {
                    Console.WriteLine($"{i + 1}. {Path.GetFileName(excelFiles[i])}");
                }

                // Câu hỏi 1: Chọn file
                Console.Write("\nPlease enter the file number to process (press Enter to exit). ");
                string fileInput = Console.ReadLine();

                // Nếu người dùng không nhập gì, kết thúc chương trình
                if (string.IsNullOrWhiteSpace(fileInput))
                {
                    Console.WriteLine("The program has ended because no file was selected.");
                    return;
                }

                int fileChoice = int.Parse(fileInput) - 1;
                if (fileChoice < 0 || fileChoice >= excelFiles.Length)
                {
                    throw new ArgumentException("Invalid file number!");
                }

                string selectedFile = excelFiles[fileChoice];
                string fileName = Path.GetFileName(selectedFile).ToLower();

                Console.Write($"\nEnter date (format: day or day-month, current: {DateTime.Now:dd-MM-yyyy}): ");
                string dateInput = Console.ReadLine().Trim();

                string customDate = "";
                if (!string.IsNullOrWhiteSpace(dateInput))
                {
                    try
                    {
                        DateTime currentDate = DateTime.Now;

                        if (dateInput.Contains("-"))
                        {
                            // Format: day-month (e.g., "22-5")
                            string[] parts = dateInput.Split('-');
                            if (parts.Length == 2)
                            {
                                int day = int.Parse(parts[0]);
                                int month = int.Parse(parts[1]);

                                if (day >= 1 && day <= 31 && month >= 1 && month <= 12)
                                {
                                    DateTime customDateTime = new DateTime(currentDate.Year, month, day);
                                    customDate = customDateTime.ToString("dd-MM-yyyy");
                                    Console.WriteLine($"Date selected: {customDate}");
                                }
                                else
                                {
                                    throw new ArgumentException("Day must be 1-31 and month must be 1-12!");
                                }
                            }
                            else
                            {
                                throw new ArgumentException("Invalid format! Use day-month (e.g., 22-5)");
                            }
                        }
                        else
                        {
                            // Format: day only (e.g., "22")
                            int day = int.Parse(dateInput);
                            if (day >= 1 && day <= 31)
                            {
                                DateTime customDateTime = new DateTime(currentDate.Year, currentDate.Month, day);
                                customDate = customDateTime.ToString("dd-MM-yyyy");
                                Console.WriteLine($"Date selected: {customDate}");
                            }
                            else
                            {
                                throw new ArgumentException("Date must be between 1-31!");
                            }
                        }
                    }
                    catch (FormatException)
                    {
                        throw new ArgumentException("Please enter valid numbers! (day or day-month)");
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        throw new ArgumentException("Invalid date for the specified month!");
                    }
                }
                else
                {
                    // If no input, use current date
                    customDate = DateTime.Now.ToString("dd-MM-yyyy");
                    Console.WriteLine($"Using current date: {customDate}");
                }



                // Xử lý dữ liệu sơ bộ - làm sạch worksheet
                ExcelProcessor.CleanWorksheet(selectedFile);

                // Mở file Excel cho người dùng xem
                Console.WriteLine("\nOpening cleaned Excel file...");
                var excelProcess = ExcelProcessor.OpenExcelFile(selectedFile);

                if (excelProcess != null)
                {
                    Console.WriteLine("File opened successfully...");
                    Console.WriteLine("Waiting for you to close Excel file...");

                    // Chờ cho đến khi người dùng đóng file Excel
                    excelProcess.WaitForExit();
                    Console.WriteLine("Detected Excel file has been closed.");

                    // Sau khi người dùng đóng file, tiếp tục với câu hỏi 2
                    // Câu hỏi 2: Chọn cột chuyển đổi text
                    string columnInput = "";
                    if (fileName.Contains("tuvung"))
                    {
                        Console.Write("\nConvert to <img> (Vocab). Default: GJ for Vocab. (Enter to skip). ");
                        columnInput = Console.ReadLine().ToUpper();
                    }
                    else
                    {
                        Console.Write("\nConvert to <img> (JP-ZH). Default: CF for JP-ZH. (Enter to skip). ");
                        columnInput = Console.ReadLine().ToUpper();
                    }

                    // Câu hỏi 3: Chọn cột âm thanh
                    string soundColumns = "";
                    if (fileName.Contains("tuvung"))
                    {
                        Console.Write("\nConvert to [sound] (Vocab). Default: DH. (Enter to skip). ");
                        soundColumns = Console.ReadLine().ToUpper();
                    }
                    else
                    {
                        Console.Write("\nConvert to [sound] (EN-JP-ZH). Default:DE for EN | Default: E for JP-ZH. (Enter to skip). ");
                        soundColumns = Console.ReadLine().ToUpper();
                    }

                    string kanjiColumn = "";
                    if (fileName.Contains("tuvung"))
                    {
                        Console.Write("\nConvert to Kanji (Vocab). Default: F for Vocab. (Enter to skip). ");
                        kanjiColumn = Console.ReadLine().ToUpper();
                    }
                    else
                    {
                        Console.Write("\nConvert to Kanji (JP-ZH). Default: B for JP-ZH. (Enter to skip). ");
                        kanjiColumn = Console.ReadLine().ToUpper();
                    }

                    // Kiểm tra xem có ít nhất một tác vụ cần thực hiện
                    if (string.IsNullOrWhiteSpace(columnInput) && string.IsNullOrWhiteSpace(soundColumns) && string.IsNullOrWhiteSpace(kanjiColumn))
                    {
                        Console.WriteLine("No tasks selected to perform...");
                    }
                    else
                    {
                        // Xử lý dữ liệu sau khi đã có thông tin từ người dùng
                        ExcelProcessor.ProcessExcelFile(selectedFile, fileName, columnInput, soundColumns, kanjiColumn, customDate);
                    }
                }
                else
                {
                    Console.WriteLine("Unable to open Excel file...");
                }

                Console.Write("\nDo it again? (Y/N): ");
                string answer = Console.ReadLine().Trim().ToUpper();

                if (answer != "Y")
                {
                    Console.WriteLine("The program has ended. Thank!");
                    continueRunning = false;
                }
                else
                {
                    Console.WriteLine("\n========================================");
                    Console.WriteLine("Starting new process...");
                    Console.WriteLine("========================================\n");
                    // continueRunning vẫn là true, quay lại vòng lặp while để chọn file mới
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError: {ex.Message}");
                Console.Write("\nDo you want to try again? (Y/N): ");
                string answer = Console.ReadLine().Trim().ToUpper();
                continueRunning = (answer == "Y");
            }
        }
    }


}
