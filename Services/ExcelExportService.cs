using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Text_to_Image.Services
{
    public class ExcelExportService
    {
        public static void ExportToExcel(List<(string Original, string ImgTags)> data)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string savePath = @"C:\Users\remove\Desktop";
                // Thêm giờ phút giây vào tên file
                string fileName = $"Text_To_Image_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                string fullPath = Path.Combine(savePath, fileName);

                // Xử lý file đang được sử dụng
                int maxRetries = 3;
                int currentRetry = 0; 
                bool fileDeleted = false;

                while (currentRetry < maxRetries && !fileDeleted)
                {
                    try
                    {
                        if (File.Exists(fullPath))
                        {
                            using (FileStream stream = File.Open(fullPath, FileMode.Open, FileAccess.Read, FileShare.None))
                            {
                                stream.Close();
                            }
                            File.Delete(fullPath);
                            fileDeleted = true;
                        }
                        else
                        {
                            fileDeleted = true;
                        }
                    }
                    catch (IOException)
                    {
                        currentRetry++;
                        if (currentRetry < maxRetries)
                        {
                            Console.WriteLine($"File đang được sử dụng, đang thử lại lần {currentRetry + 1}...");
                            Thread.Sleep(1000);
                        }
                        else
                        {
                            throw new IOException("Không thể xóa file vì nó đang được sử dụng. Vui lòng đóng file Excel và thử lại.");
                        }
                    }
                }

                using (var package = new ExcelPackage(new FileInfo(fullPath)))
                {
                    // Kiểm tra và xóa worksheet cũ nếu tồn tại
                    var existingWorksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Conversion Results");
                    if (existingWorksheet != null)
                    {
                        package.Workbook.Worksheets.Delete(existingWorksheet);
                    }

                    // Tạo worksheet mới
                    var worksheet = package.Workbook.Worksheets.Add("Conversion Results");

                    // Set headers
                    worksheet.Cells[1, 1].Value = "Văn bản gốc";
                    worksheet.Cells[1, 2].Value = "Chuỗi thẻ img";

                    // Format headers
                    using (var range = worksheet.Cells[1, 1, 1, 2])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    }

                    // Add data
                    for (int i = 0; i < data.Count; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = data[i].Original;
                        worksheet.Cells[i + 2, 2].Value = data[i].ImgTags;
                    }

                    // Auto-fit columns
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                    // Save file với retry
                    bool saved = false;
                    currentRetry = 0;

                    while (currentRetry < maxRetries && !saved)
                    {
                        try
                        {
                            package.Save();
                            saved = true;
                            Console.WriteLine("Excel file đã được tạo thành công!");
                            Console.WriteLine($"Vị trí file: {fullPath}");
                        }
                        catch (IOException)
                        {
                            currentRetry++;
                            if (currentRetry < maxRetries)
                            {
                                Console.WriteLine($"Không thể lưu file, đang thử lại lần {currentRetry + 1}...");
                                Thread.Sleep(1000);
                            }
                            else
                            {
                                throw;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Có lỗi khi tạo file Excel: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }
    }
}
