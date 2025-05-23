using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using JapaneseConverter;
using Text_to_Image.Models;
using static Microsoft.IO.RecyclableMemoryStreamManager;

namespace Text_to_Image.Services
{
    public class ExcelProcessor
    {
        private static readonly string FolderPath = @"S:\Anki";

        public static void CleanWorksheet(string filePath)
        {
            try
            {
                Console.WriteLine("Đang làm sạch dữ liệu trong worksheet...");

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

                            Console.WriteLine($"Đã xóa {mergedAddresses.Count} vùng cells đã merge.");
                        }

                        // Xóa conditional formatting nếu có
                        worksheet.ConditionalFormatting.RemoveAll();

                        Console.WriteLine($"Đã làm sạch hoàn toàn {lastRow} hàng và {lastColumn} cột dữ liệu.");
                    }
                    else
                    {
                        Console.WriteLine("Worksheet trống, không có dữ liệu để làm sạch.");
                    }

                    package.Save();
                }

                Console.WriteLine("Đã hoàn tất việc làm sạch worksheet.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi làm sạch worksheet: {ex.Message}");
                throw; // Ném lại ngoại lệ để xử lý ở mức cao hơn
            }
        }
        public static void ProcessExcelFile(string filePath, string fileName, string columnInput, string soundColumns, string kanjiColumn, string kanjiOutputColumn, string customDate = "")
        {
            try
            {
                Console.WriteLine("Đang xử lý dữ liệu trong file Excel...");

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension != null ? worksheet.Dimension.End.Row : 0;
                    string dateToUse = string.IsNullOrWhiteSpace(customDate) ? DateTime.Now.ToString("dd-MM-yyyy") : customDate;
                    Console.WriteLine($"Sử dụng ngày: {dateToUse}");


                    if (rowCount == 0)
                    {
                        Console.WriteLine("Worksheet trống, không có dữ liệu để xử lý.");
                        return;
                    }

                    Console.WriteLine($"\nFile Excel có {rowCount} dòng.");
                    Console.WriteLine("Bắt đầu thực hiện chuyển đổi...");

                    // Xử lý chuyển đổi text và tạo công thức âm thanh
                    for (int row = 1; row <= rowCount; row++)
                    {
                        // Xử lý chuyển đổi text nếu được yêu cầu
                        if (!string.IsNullOrWhiteSpace(columnInput) && columnInput.Length == 2)
                        {
                            int inputColumn = columnInput[0] - 'A' + 1;
                            int outputColumn = columnInput[1] - 'A' + 1;

                            string inputText = worksheet.Cells[row, inputColumn].Text;
                            if (!string.IsNullOrEmpty(inputText))
                            {
                                string convertedText = TextConverter.ConvertToImgTags(inputText);
                                worksheet.Cells[row, outputColumn].Value = convertedText;
                            }
                        }

                        // Xử lý công thức âm thanh nếu được yêu cầu
                        if (!string.IsNullOrWhiteSpace(soundColumns))
                        {
                            if (fileName.Contains("tuvung") && soundColumns.Length == 2)
                            {
                                int soundCol1 = soundColumns[0] - 'A' + 1;
                                int soundCol2 = soundColumns[1] - 'A' + 1;

                                worksheet.Cells[row, soundCol1].Value =
                                    $"[sound:Vocab-{dateToUse}_{row:00}.mp3]";

                                worksheet.Cells[row, soundCol2].Value =
                                    $"[sound:Vocab-{dateToUse}_{(row + rowCount):00}.mp3]";
                            }
                            else if (soundColumns.Length == 1)
                            {
                                int soundCol = soundColumns[0] - 'A' + 1;
                                string prefix = fileName.Contains("english") ? "EN" :
                                              fileName.Contains("japanese") ? "JP" :
                                              fileName.Contains("chinese") ? "ZH" : "";

                                if (!string.IsNullOrEmpty(prefix))
                                {
                                    worksheet.Cells[row, soundCol].Value =
                                        $"[sound:{prefix}-{dateToUse}_{row:00}.mp3]";
                                }
                            }
                            else if (soundColumns.Length == 2)
                            {
                                // Xử lý 2 cột - áp dụng logic lẻ/chẵn cho EN/JP/ZH
                                int soundCol1 = soundColumns[0] - 'A' + 1; // Cột cho file số lẻ
                                int soundCol2 = soundColumns[1] - 'A' + 1; // Cột cho file số chẵn
                                string prefix = fileName.Contains("english") ? "EN" :
                                              fileName.Contains("japanese") ? "JP" :
                                              fileName.Contains("chinese") ? "ZH" : "";

                                if (!string.IsNullOrEmpty(prefix))
                                {
                                    // Công thức cho file âm thanh số lẻ (1, 3, 5, 7, ...)
                                    worksheet.Cells[row, soundCol1].Value =
                                        $"[sound:{prefix}-{dateToUse}_{(row * 2 - 1):0000}.mp3]";

                                    // Công thức cho file âm thanh số chẵn (2, 4, 6, 8, ...)
                                    worksheet.Cells[row, soundCol2].Value =
                                        $"[sound:{prefix}-{dateToUse}_{(row * 2):0000}.mp3]";
                                }
                            }

                        }

                        // Xử lý định dạng Kanji nếu được yêu cầu
                        if (!string.IsNullOrWhiteSpace(kanjiColumn) && kanjiColumn.Length >= 1)
                        {
                            int sourceCol = kanjiColumn[0] - 'A' + 1;

                            // Xác định cột đích - cần thêm tham số kanjiOutputColumn vào method
                            int outputCol = sourceCol; // Mặc định lưu vào cột nguồn
                            if (!string.IsNullOrWhiteSpace(kanjiOutputColumn) && kanjiOutputColumn.Length >= 1)
                            {
                                outputCol = kanjiOutputColumn[0] - 'A' + 1;
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
                        Console.WriteLine($"\nHoàn tất! Đã xử lý {rowCount} dòng.");

                        // Hiển thị thông tin về các tác vụ đã thực hiện
                        if (!string.IsNullOrWhiteSpace(columnInput))
                            Console.WriteLine($"- <img>: done | {columnInput[0]} -> {columnInput[1]}");
                        if (!string.IsNullOrWhiteSpace(soundColumns))
                            Console.WriteLine($"- [sound]: done | {soundColumns}");
                        if (!string.IsNullOrWhiteSpace(kanjiColumn))
                            Console.WriteLine($"- kanji: done | {kanjiColumn}");

                        // Hỏi người dùng có muốn mở file đã xử lý không
                        Console.Write("\nOpen Processed Excel File? (Y/N): ");
                        string openAnswer = Console.ReadLine().Trim().ToUpper();
                        if (openAnswer == "Y")
                        {
                            OpenExcelFile(filePath);
                        }
                    }
                    catch (IOException ex)
                    {
                        throw new IOException("Không thể lưu file. Vui lòng đảm bảo file không đang được mở.", ex);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi xử lý file: {ex.Message}");
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
                Console.WriteLine($"Lỗi khi mở file Excel: {ex.Message}");
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
                AudioFileRenamer.RenameAudioFiles(options.AudioFolderPath, options.SelectedDay,
                    options.SelectedMonth, options.SelectedYear, options.FileName);
            }

            // Excel processing
            if (!string.IsNullOrWhiteSpace(options.ColumnInput) ||
                !string.IsNullOrWhiteSpace(options.SoundColumns) ||
                !string.IsNullOrWhiteSpace(options.KanjiColumn))
            {
                ExcelProcessor.ProcessExcelFile(options.SelectedFile, options.FileName,
                    options.ColumnInput, options.SoundColumns, options.KanjiColumn, options.KanjiOutputColumn, options.CustomDate);
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


    }

}
