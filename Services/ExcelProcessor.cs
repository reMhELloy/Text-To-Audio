using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

                    // Xóa hoàn toàn mọi dữ liệu trong worksheet
                    if (worksheet.Dimension != null)
                    {
                        // Lấy kích thước của worksheet
                        int lastRow = worksheet.Dimension.End.Row;
                        int lastColumn = worksheet.Dimension.End.Column;

                        // Xóa toàn bộ nội dung, bao gồm cả tiêu đề
                        worksheet.Cells[1, 1, lastRow, lastColumn].Clear();

                        // Xóa cả định dạng, công thức, và comment
                        worksheet.Cells[1, 1, lastRow, lastColumn].Style.Font.Bold = false;
                        worksheet.Cells[1, 1, lastRow, lastColumn].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.None;
                        worksheet.Cells[1, 1, lastRow, lastColumn].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.None);

                        // Xóa các merged cells
                        if (worksheet.MergedCells.Count > 0)
                        {
                            // Cần tạo một danh sách các địa chỉ để xóa sau
                            // Vì không thể xóa trực tiếp trong khi đang duyệt qua collection
                            var mergedAddresses = new List<string>();
                            foreach (string mergedCell in worksheet.MergedCells)
                            {
                                mergedAddresses.Add(mergedCell);
                            }

                            // Xóa từng merged cell
                            foreach (string address in mergedAddresses)
                            {
                                worksheet.Cells[address].Merge = false;
                            }

                            Console.WriteLine($"Removed {mergedAddresses.Count} merged cell regions.");
                        }

                        // Xóa conditional formatting nếu có
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
                throw; // Ném lại ngoại lệ để xử lý ở mức cao hơn
            }
        }

        public static void ProcessExcelFile(ProcessingOptions options)
        {
            try
            {
                Console.WriteLine("Processing Excel file data...");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(options.SelectedFile)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension != null ? worksheet.Dimension.End.Row : 0;
                    string dateToUse = string.IsNullOrWhiteSpace(options.CustomDate) ? DateTime.Now.ToString("dd-MM-yyyy") : options.CustomDate;
                    Console.WriteLine($"Using date: {dateToUse}");

                    if (rowCount == 0)
                    {
                        Console.WriteLine("Worksheet is empty, no data to process.");
                        return;
                    }

                    Console.WriteLine($"\nExcel file has {rowCount} rows.");
                    Console.WriteLine("Starting conversion...");

                    // Xử lý chuyển đổi text và tạo công thức âm thanh
                    for (int row = 1; row <= rowCount; row++)
                    {
                        // Xử lý chuyển đổi text nếu được yêu cầu (BỎ QUA ENGLISH)
                        if (!string.IsNullOrWhiteSpace(options.ColumnInput) &&
                            options.ColumnInput.Length >= 1 &&
                            !options.FileName.Contains("english"))
                        {
                            int inputColumn = options.ColumnInput[0] - 'A' + 1;

                            string inputText = worksheet.Cells[row, inputColumn].Text;
                            if (!string.IsNullOrEmpty(inputText))
                            {
                                string convertedText = TextConverter.ConvertToImgTags(inputText);

                                // Lưu vào cột G cố định cho JP/ZH/TuVung
                                int outputColumn = 7; // Cột G
                                worksheet.Cells[row, outputColumn].Value = convertedText;
                            }
                        }

                        // Xử lý công thức âm thanh nếu được yêu cầu
                        if (!string.IsNullOrWhiteSpace(options.SoundColumns))
                        {
                            if (options.FileName.Contains("tuvung") && options.SoundColumns.Length == 2)
                            {
                                // TuVung: Logic chẵn lẻ - Lẻ = EN (E), Chẵn = VI (F)
                                int soundCol1 = 5; // Cột E - EN (lẻ)
                                int soundCol2 = 6; // Cột F - VI (chẵn)

                                worksheet.Cells[row, soundCol1].Value =
                                    $"[sound:EN-{dateToUse}_{(row * 2 - 1):00}.mp3]";
                                worksheet.Cells[row, soundCol2].Value =
                                    $"[sound:VI-{dateToUse}_{(row * 2):00}.mp3]";
                            }
                            else if (options.FileName.Contains("japanese") && options.SoundColumns.Length == 2)
                            {
                                // Japanese: Logic chẵn lẻ - Lẻ = EN (E), Chẵn = JP (F)
                                int soundCol1 = 5; // Cột E - EN (lẻ)
                                int soundCol2 = 6; // Cột F - JP (chẵn)

                                worksheet.Cells[row, soundCol1].Value =
                                    $"[sound:EN-{dateToUse}_{(row * 2 - 1):00}.mp3]";
                                worksheet.Cells[row, soundCol2].Value =
                                    $"[sound:JP-{dateToUse}_{(row * 2):00}.mp3]";
                            }
                            else if (options.FileName.Contains("chinese") && options.SoundColumns.Length == 2)
                            {
                                // Chinese: Logic chẵn lẻ - Lẻ = EN (E), Chẵn = ZH (F)
                                int soundCol1 = 5; // Cột E - EN (lẻ)
                                int soundCol2 = 6; // Cột F - ZH (chẵn)

                                worksheet.Cells[row, soundCol1].Value =
                                    $"[sound:EN-{dateToUse}_{(row * 2 - 1):00}.mp3]";
                                worksheet.Cells[row, soundCol2].Value =
                                    $"[sound:ZH-{dateToUse}_{(row * 2):00}.mp3]";
                            }
                            else if (options.FileName.Contains("english") && options.SoundColumns.Length == 2)
                            {
                                // English: Logic gốc - Lẻ/chẵn
                                int soundCol1 = options.SoundColumns[0] - 'A' + 1;
                                int soundCol2 = options.SoundColumns[1] - 'A' + 1;

                                worksheet.Cells[row, soundCol1].Value =
                                    $"[sound:EN-{dateToUse}_{(row * 2 - 1):00}.mp3]";
                                worksheet.Cells[row, soundCol2].Value =
                                    $"[sound:EN-{dateToUse}_{(row * 2):00}.mp3]";
                            }
                            else if (options.SoundColumns.Length == 1)
                            {
                                // Logic 1 cột (giữ nguyên)
                                int soundCol = options.SoundColumns[0] - 'A' + 1;
                                string prefix = options.FileName.Contains("english") ? "EN" :
                                              options.FileName.Contains("japanese") ? "JP" :
                                              options.FileName.Contains("chinese") ? "ZH" : "";

                                if (!string.IsNullOrEmpty(prefix))
                                {
                                    worksheet.Cells[row, soundCol].Value =
                                        $"[sound:{prefix}-{dateToUse}_{row:00}.mp3]";
                                }
                            }
                        }

                        // Xử lý định dạng Kanji nếu được yêu cầu (BỎ QUA ENGLISH)
                        if (!string.IsNullOrWhiteSpace(options.KanjiColumn) &&
                            options.KanjiColumn.Length >= 1 &&
                            !options.FileName.Contains("english") &&
                            (options.FileName.Contains("tuvung") ||
                             options.FileName.Contains("japanese") ||
                             options.FileName.Contains("JP") ||
                             options.FileName.Contains("chinese") ||
                             options.FileName.Contains("ZH")))
                        {
                            int sourceCol = options.KanjiColumn[0] - 'A' + 1;

                            // Xác định cột đích - cần thêm tham số kanjiOutputColumn vào method
                            int outputCol = sourceCol; // Mặc định lưu vào cột nguồn
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
                    }

                    try
                    {
                        package.Save();
                        Console.WriteLine($"\nCompleted! Processed {rowCount} rows.");

                        // Hiển thị thông tin về các tác vụ đã thực hiện
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
                                Console.WriteLine($"- [sound]: done | E (odd-EN) & F (even-native)");
                            }
                            else
                            {
                                Console.WriteLine($"- [sound]: done | {options.SoundColumns}");
                            }
                        }
                        if (!string.IsNullOrWhiteSpace(options.KanjiColumn))
                            Console.WriteLine($"- kanji: done | {options.KanjiColumn}");

                        // Hỏi người dùng có muốn mở file đã xử lý không
                        Console.Write("\nOpen Processed Excel File? (Y/N): ");
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
            }
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

            Console.Write("\nPlease enter the file number to process (press Enter to exit). ");
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
            // Xử lý dữ liệu sơ bộ - làm sạch worksheet
            ExcelProcessor.CleanWorksheet(options.SelectedFile);

            // Mở file Excel cho người dùng xem
            Console.WriteLine("\nOpening cleaned Excel file...");
            var excelProcess = ExcelProcessor.OpenExcelFile(options.SelectedFile);

            if (excelProcess != null)
            {
                Console.WriteLine("File opened successfully...");
                Console.WriteLine("Waiting for you to close Excel file...");

                // Chờ cho đến khi người dùng đóng file Excel
                excelProcess.WaitForExit();
                Console.WriteLine("Detected Excel file has been closed.");
                return true;
            }

            return false;
        }

        public static void ExecuteTasks(ProcessingOptions options)
        {
            // Audio file renaming
            if (options.RenameAudioFiles)
            {
                AudioFileManager.RenameAudioFiles(options.AudioFolderPath, options.SelectedDay,
                    options.SelectedMonth, options.SelectedYear, options.FileName);
            }

            // Excel processing
            if (!string.IsNullOrWhiteSpace(options.ColumnInput) ||
                !string.IsNullOrWhiteSpace(options.SoundColumns) ||
                !string.IsNullOrWhiteSpace(options.KanjiColumn))
            {
                ExcelProcessor.ProcessExcelFile(options);
            }
            else
            {
                Console.WriteLine("No tasks selected to perform...");
            }
        }

        public static bool AskToContinue()
        {
            Console.Write("\nDo it again? (Y/N): ");
            string answer = Console.ReadLine()?.Trim().ToUpper();

            if (answer != "Y")
            {
                Console.WriteLine("The program has ended. Thank!");
                return false;
            }
            else
            {
                Console.WriteLine("\n========================================");
                Console.WriteLine("Starting new process...");
                Console.WriteLine("========================================\n");
                return true;
            }
        }
        public static async Task ExecuteTasksWithDatabase(ProcessingOptions options)
        {
            // Audio file renaming
            if (options.RenameAudioFiles)
            {
                AudioFileManager.RenameAudioFiles(options.AudioFolderPath, options.SelectedDay,
                    options.SelectedMonth, options.SelectedYear, options.FileName);
            }

            // Excel processing
            if (!string.IsNullOrWhiteSpace(options.ColumnInput) ||
                !string.IsNullOrWhiteSpace(options.SoundColumns) ||
                !string.IsNullOrWhiteSpace(options.KanjiColumn))
            {
                ExcelProcessor.ProcessExcelFile(options);
            }
            else
            {
                Console.WriteLine("No tasks selected to perform...");
                return;
            }

            // Save to database ONLY if user chose to
            if (options.SaveToDatabase)
            {
                try
                {
                    Console.WriteLine("\n" + new string('=', 50));

                    using var dbService = new DatabaseService();
                    var session = await dbService.SaveExcelDataToDatabaseAsync(options);

                    Console.WriteLine($"\n📊 Database Summary:");
                    Console.WriteLine($"✅ Session ID: {session.SessionId}");
                    Console.WriteLine($"✅ Processed: {session.ProcessedRows} vocabulary entries");
                    Console.WriteLine($"✅ Audio records: {session.AudioFilesCreated}");
                    Console.WriteLine($"✅ File type: {session.FileType}");
                    Console.WriteLine($"✅ Date: {session.DateUsed}");
                    Console.WriteLine(new string('=', 50));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"\n❌ Database save failed: {ex.Message}");
                    Console.WriteLine("Excel processing completed but data not saved to database.");
                    Console.WriteLine("Please check your SQL Server connection.");

                    // Log chi tiết lỗi để debug
                    Console.WriteLine($"Error details: {ex.InnerException?.Message}");
                }
            }
            else
            {
                Console.WriteLine("\n💾 Database save skipped (user choice).");
                Console.WriteLine("Excel processing completed successfully without database save.");
            }
        }
    }
}