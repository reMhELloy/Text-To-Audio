using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using JapaneseConverter;
using Text_to_Image.Models;
using Text_to_Image.Data;

namespace Text_to_Image.Services
{
    public class ExcelProcessor
    {
        private static readonly string FolderPath = @"S:\Anki";

        public static void CleanWorksheet(string filePath)
        {
            try
            {
                Console.WriteLine("Cleaning worksheet data...");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    if (worksheet.Dimension != null)
                    {
                        int lastRow = worksheet.Dimension.End.Row;
                        int lastColumn = worksheet.Dimension.End.Column;

                        worksheet.Cells[1, 1, lastRow, lastColumn].Clear();
                        worksheet.Cells[1, 1, lastRow, lastColumn].Style.Font.Bold = false;
                        worksheet.Cells[1, 1, lastRow, lastColumn].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.None;
                        worksheet.Cells[1, 1, lastRow, lastColumn].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.None);

                        if (worksheet.MergedCells.Count > 0)
                        {
                            var mergedAddresses = new List<string>();
                            foreach (string mergedCell in worksheet.MergedCells)
                            {
                                mergedAddresses.Add(mergedCell);
                            }

                            foreach (string address in mergedAddresses)
                            {
                                worksheet.Cells[address].Merge = false;
                            }

                            Console.WriteLine($"Removed {mergedAddresses.Count} merged cell regions.");
                        }

                        worksheet.ConditionalFormatting.RemoveAll();
                        Console.WriteLine($"Cleaned {lastRow} rows and {lastColumn} columns completely.");
                    }
                    else
                    {
                        Console.WriteLine("Worksheet is empty, no data to clean.");
                    }

                    package.Save();
                }

                Console.WriteLine("Worksheet cleaning completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error cleaning worksheet: {ex.Message}");
                throw;
            }
        }

        public static async Task ProcessExcelFile(ProcessingOptions options)
        {
            try
            {
                Console.WriteLine("Processing Excel file with database integration...");

                string dateToUse = string.IsNullOrWhiteSpace(options.CustomDate) ?
                    DateTime.Now.ToString("dd-MM-yyyy") : options.CustomDate;

                Console.WriteLine($"Using date: {dateToUse}");

                // Check database for existing audio files
                int existingAudioCount = 0;
                try
                {
                    using var dbService = new DatabaseService();
                    string fileType = DetermineFileType(options.FileName);
                    string audioFileType = DetermineAudioFileType(fileType);

                    var existingSummary = await dbService.GetExistingDataSummary(dateToUse, fileType, audioFileType);
                    existingAudioCount = existingSummary.AudioFileCount;

                    Console.WriteLine("Database Check Results:");
                    Console.WriteLine($"  Existing vocabularies: {existingSummary.VocabularyCount}");
                    Console.WriteLine($"  Existing audio files: {existingSummary.AudioFileCount}");
                    Console.WriteLine($"  Next audio number starts from: {existingSummary.NextAudioNumber}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Database check failed: {ex.Message}");
                    Console.WriteLine("Will proceed with audio numbering from 1");
                    existingAudioCount = 0;
                }

                // Process Excel with correct audio numbering
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(options.SelectedFile)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension != null ? worksheet.Dimension.End.Row : 0;

                    if (rowCount == 0)
                    {
                        Console.WriteLine("Worksheet is empty, no data to process.");
                        return;
                    }

                    Console.WriteLine($"Excel file has {rowCount} rows.");
                    Console.WriteLine("Starting conversion...");

                    for (int row = 1; row <= rowCount; row++)
                    {
                        // Text conversion (skip for English files)
                        if (!string.IsNullOrWhiteSpace(options.ColumnInput) &&
                            options.ColumnInput.Length >= 1 &&
                            !options.FileName.Contains("english"))
                        {
                            int inputColumn = options.ColumnInput[0] - 'A' + 1;
                            string inputText = worksheet.Cells[row, inputColumn].Text;

                            if (!string.IsNullOrEmpty(inputText))
                            {
                                string convertedText = TextConverter.ConvertToImgTags(inputText);
                                int outputColumn = 7; // Column G
                                worksheet.Cells[row, outputColumn].Value = convertedText;
                            }
                        }

                        // Sound formulas with correct numbering based on database
                        if (!string.IsNullOrWhiteSpace(options.SoundColumns))
                        {
                            ProcessSoundFormulas(worksheet, row, options, dateToUse, existingAudioCount);
                        }

                        // Kanji processing (skip for English files)
                        if (!string.IsNullOrWhiteSpace(options.KanjiColumn) &&
                            options.KanjiColumn.Length >= 1 &&
                            !options.FileName.Contains("english") &&
                            (options.FileName.Contains("tuvung") ||
                             options.FileName.Contains("japanese") ||
                             options.FileName.Contains("chinese")))
                        {
                            ProcessKanjiFormatting(worksheet, row, options);
                        }
                    }

                    try
                    {
                        package.Save();
                        Console.WriteLine($"Completed! Processed {rowCount} rows.");
                        ShowProcessingSummary(options);

                        Console.Write("Open Processed Excel File? (Y/N): ");
                        string openAnswer = Console.ReadLine().Trim().ToUpper();
                        if (openAnswer == "Y")
                        {
                            OpenExcelFile(options.SelectedFile);
                        }
                    }
                    catch (IOException ex)
                    {
                        throw new IOException("Cannot save file. Please ensure the file is not currently open.", ex);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file: {ex.Message}");
                throw;
            }
        }

        private static void ProcessSoundFormulas(ExcelWorksheet worksheet, int row, ProcessingOptions options,
            string dateToUse, int existingAudioCount)
        {
            // Calculate correct audio numbers based on existing files in database
            int currentAudioOdd = existingAudioCount + (row * 2 - 1);
            int currentAudioEven = existingAudioCount + (row * 2);

            if (options.FileName.Contains("tuvung") && options.SoundColumns.Length == 2)
            {
                worksheet.Cells[row, 5].Value = $"[sound:EN-{dateToUse}_{currentAudioOdd:00}.mp3]";
                worksheet.Cells[row, 6].Value = $"[sound:VI-{dateToUse}_{currentAudioEven:00}.mp3]";
            }
            else if (options.FileName.Contains("japanese") && options.SoundColumns.Length == 2)
            {
                worksheet.Cells[row, 5].Value = $"[sound:EN-{dateToUse}_{currentAudioOdd:00}.mp3]";
                worksheet.Cells[row, 6].Value = $"[sound:JP-{dateToUse}_{currentAudioEven:00}.mp3]";
            }
            else if (options.FileName.Contains("chinese") && options.SoundColumns.Length == 2)
            {
                worksheet.Cells[row, 5].Value = $"[sound:EN-{dateToUse}_{currentAudioOdd:00}.mp3]";
                worksheet.Cells[row, 6].Value = $"[sound:ZH-{dateToUse}_{currentAudioEven:00}.mp3]";
            }
            else if (options.FileName.Contains("english") && options.SoundColumns.Length == 2)
            {
                int soundCol1 = options.SoundColumns[0] - 'A' + 1;
                int soundCol2 = options.SoundColumns[1] - 'A' + 1;

                worksheet.Cells[row, soundCol1].Value = $"[sound:EN-{dateToUse}_{currentAudioOdd:00}.mp3]";
                worksheet.Cells[row, soundCol2].Value = $"[sound:EN-{dateToUse}_{currentAudioEven:00}.mp3]";
            }
            else if (options.SoundColumns.Length == 1)
            {
                int soundCol = options.SoundColumns[0] - 'A' + 1;
                string prefix = DetermineLanguagePrefix(options.FileName);

                if (!string.IsNullOrEmpty(prefix))
                {
                    int audioNumber = existingAudioCount + row;
                    worksheet.Cells[row, soundCol].Value = $"[sound:{prefix}-{dateToUse}_{audioNumber:00}.mp3]";
                }
            }
        }

        private static void ProcessKanjiFormatting(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            int sourceCol = options.KanjiColumn[0] - 'A' + 1;
            int outputCol = sourceCol;

            if (!string.IsNullOrWhiteSpace(options.KanjiOutputColumn) && options.KanjiOutputColumn.Length >= 1)
            {
                outputCol = options.KanjiOutputColumn[0] - 'A' + 1;
            }

            string cellValue = worksheet.Cells[row, sourceCol].Text;
            if (!string.IsNullOrEmpty(cellValue))
            {
                string formattedText = KanjiHelper.FormatKanjiWithBrackets(cellValue);
                worksheet.Cells[row, outputCol].Value = formattedText;
            }
        }

        private static void ShowProcessingSummary(ProcessingOptions options)
        {
            if (!string.IsNullOrWhiteSpace(options.ColumnInput) && !options.FileName.Contains("english"))
            {
                Console.WriteLine($"- <img>: done | {options.ColumnInput[0]} -> G");
            }

            if (!string.IsNullOrWhiteSpace(options.SoundColumns))
            {
                if ((options.FileName.Contains("tuvung") ||
                     options.FileName.Contains("japanese") ||
                     options.FileName.Contains("chinese")) &&
                    options.SoundColumns.Length == 2)
                {
                    Console.WriteLine("- [sound]: done | E (odd-EN) & F (even-native)");
                }
                else
                {
                    Console.WriteLine($"- [sound]: done | {options.SoundColumns}");
                }
            }

            if (!string.IsNullOrWhiteSpace(options.KanjiColumn))
                Console.WriteLine($"- kanji: done | {options.KanjiColumn}");
        }

        private static string DetermineFileType(string fileName)
        {
            fileName = fileName.ToLower();
            if (fileName.Contains("english")) return "English";
            if (fileName.Contains("japanese")) return "Japanese";
            if (fileName.Contains("chinese")) return "Chinese";
            if (fileName.Contains("tuvung")) return "TuVung";
            return "Unknown";
        }

        private static string DetermineAudioFileType(string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => "VI-EN",
                "japanese" => "JP-EN",
                "chinese" => "ZH-EN",
                "tuvung" => "TUVUNG",
                _ => "VI-EN"
            };
        }

        private static string DetermineLanguagePrefix(string fileName)
        {
            fileName = fileName.ToLower();
            if (fileName.Contains("english")) return "EN";
            if (fileName.Contains("japanese")) return "JP";
            if (fileName.Contains("chinese")) return "ZH";
            return "";
        }

        public static System.Diagnostics.Process OpenExcelFile(string filePath)
        {
            try
            {
                var processInfo = new System.Diagnostics.ProcessStartInfo(filePath)
                {
                    UseShellExecute = true
                };
                return System.Diagnostics.Process.Start(processInfo);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error opening Excel file: {ex.Message}");
                return null;
            }
        }

        public static bool SelectExcelFile(ProcessingOptions options)
        {
            string[] excelFiles = Directory.GetFiles(FolderPath, "*.xlsm");

            Console.WriteLine("Available Excel files:");
            for (int i = 0; i < excelFiles.Length; i++)
            {
                Console.WriteLine($"{i + 1}. {Path.GetFileName(excelFiles[i])}");
            }

            Console.Write("Please enter the file number to process (press Enter to exit). ");
            string fileInput = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(fileInput))
                return false;

            int fileChoice = int.Parse(fileInput) - 1;
            if (fileChoice < 0 || fileChoice >= excelFiles.Length)
            {
                throw new ArgumentException("Invalid file number!");
            }

            options.SelectedFile = excelFiles[fileChoice];
            options.FileName = Path.GetFileName(options.SelectedFile).ToLower();
            return true;
        }

        public static bool OpenAndWaitForExcelFile(ProcessingOptions options)
        {
            CleanWorksheet(options.SelectedFile);

            Console.WriteLine("Opening cleaned Excel file...");
            var excelProcess = OpenExcelFile(options.SelectedFile);

            if (excelProcess != null)
            {
                Console.WriteLine("File opened successfully...");
                Console.WriteLine("Waiting for you to close Excel file...");

                excelProcess.WaitForExit();
                Console.WriteLine("Detected Excel file has been closed.");
                return true;
            }

            return false;
        }

        public static async Task ExecuteTasks(ProcessingOptions options)
        {
            if (options.RenameAudioFiles)
            {
                AudioFileManager.RenameAudioFiles(options.AudioFolderPath, options.SelectedDay,
                    options.SelectedMonth, options.SelectedYear, options.FileName);
            }

            if (!string.IsNullOrWhiteSpace(options.ColumnInput) ||
                !string.IsNullOrWhiteSpace(options.SoundColumns) ||
                !string.IsNullOrWhiteSpace(options.KanjiColumn))
            {
                await ProcessExcelFile(options);
            }
            else
            {
                Console.WriteLine("No tasks selected to perform...");
            }
        }

        public static async Task ExecuteTasksWithDatabase(ProcessingOptions options)
        {
            if (options.RenameAudioFiles)
            {
                AudioFileManager.RenameAudioFiles(options.AudioFolderPath, options.SelectedDay,
                    options.SelectedMonth, options.SelectedYear, options.FileName);
            }

            if (!string.IsNullOrWhiteSpace(options.ColumnInput) ||
                !string.IsNullOrWhiteSpace(options.SoundColumns) ||
                !string.IsNullOrWhiteSpace(options.KanjiColumn))
            {
                await ProcessExcelFile(options);
            }
            else
            {
                Console.WriteLine("No tasks selected to perform...");
                return;
            }

            if (options.SaveToDatabase)
            {
                try
                {
                    Console.WriteLine("==================================================");

                    using var dbService = new DatabaseService();
                    var session = await dbService.SaveExcelDataToDatabaseAsync(options);

                    Console.WriteLine("Database Summary:");
                    Console.WriteLine($"Session ID: {session.SessionId}");
                    Console.WriteLine($"Processed: {session.ProcessedRows} vocabulary entries");
                    Console.WriteLine($"Audio records: {session.AudioFilesCreated}");
                    Console.WriteLine($"File type: {session.FileType}");
                    Console.WriteLine($"Date: {session.DateUsed}");
                    Console.WriteLine("==================================================");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Database save failed: {ex.Message}");
                    Console.WriteLine("Excel processing completed but data not saved to database.");
                    Console.WriteLine("Please check your SQL Server connection.");
                    Console.WriteLine($"Error details: {ex.InnerException?.Message}");
                }
            }
            else
            {
                Console.WriteLine("Database save skipped (user choice).");
                Console.WriteLine("Excel processing completed successfully without database save.");
            }
        }
    }
}