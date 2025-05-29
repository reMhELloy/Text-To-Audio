using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using Text_to_Image.Data;
using Text_to_Image.Data.Models;

namespace Text_to_Image.Services
{
    public class DatabaseExportService : IDisposable
    {
        private readonly LanguageLearningContext _context;
        private static readonly string ExportPath = @"S:\Anki";

        public DatabaseExportService()
        {
            _context = new LanguageLearningContext();
        }

        public async Task ShowExportMenu()
        {
            Console.WriteLine("\n" + new string('=', 60));
            Console.WriteLine("📤 DATABASE EXPORT TO EXCEL");
            Console.WriteLine(new string('=', 60));

            while (true)
            {
                Console.WriteLine("\nChoose export option:");
                Console.WriteLine("1. Export by Date");
                Console.WriteLine("2. Export by File Type");
                Console.WriteLine("3. Export Recent Sessions");
                Console.WriteLine("4. View Database Statistics");
                Console.WriteLine("0. Back to Main Menu");

                Console.Write("\nEnter your choice (0-4): ");
                string choice = Console.ReadLine()?.Trim();

                switch (choice)
                {
                    case "1":
                        await ExportByDate();
                        break;
                    case "2":
                        await ExportByFileType();
                        break;
                    case "3":
                        await ExportRecentSessions();
                        break;
                    case "4":
                        await ShowDatabaseStatistics();
                        break;
                    case "0":
                        return;
                    default:
                        Console.WriteLine("❌ Invalid choice. Please try again.");
                        break;
                }
            }
        }

        private async Task ExportByDate()
        {
            Console.WriteLine("\n📅 EXPORT BY DATE");
            Console.WriteLine("Enter date to export (format: dd-MM-yyyy or press Enter for today):");

            string dateInput = Console.ReadLine()?.Trim();
            DateTime targetDate;

            if (string.IsNullOrWhiteSpace(dateInput))
            {
                targetDate = DateTime.Now.Date;
            }
            else
            {
                if (!DateTime.TryParseExact(dateInput, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out targetDate))
                {
                    Console.WriteLine("❌ Invalid date format. Please use dd-MM-yyyy");
                    return;
                }
            }

            // Get sessions for that date
            var sessions = await _context.ProcessingSessions
                .Where(s => s.ProcessedDate.Date == targetDate.Date)
                .OrderByDescending(s => s.ProcessedDate)
                .ToListAsync();

            if (!sessions.Any())
            {
                Console.WriteLine($"❌ No data found for date: {targetDate:dd-MM-yyyy}");
                return;
            }

            Console.WriteLine($"\n📊 Found {sessions.Count} session(s) for {targetDate:dd-MM-yyyy}:");
            for (int i = 0; i < sessions.Count; i++)
            {
                var session = sessions[i];
                Console.WriteLine($"{i + 1}. {session.FileType} - {session.ProcessedRows} words - {session.ProcessedDate:HH:mm:ss}");
            }

            Console.Write("\nSelect session to export (or 0 to export all): ");
            string sessionChoice = Console.ReadLine()?.Trim();

            if (sessionChoice == "0")
            {
                foreach (var session in sessions)
                {
                    await ExportSessionToExcel(session);
                }
            }
            else if (int.TryParse(sessionChoice, out int index) && index > 0 && index <= sessions.Count)
            {
                await ExportSessionToExcel(sessions[index - 1]);
            }
            else
            {
                Console.WriteLine("❌ Invalid selection.");
            }
        }

        private async Task ExportByFileType()
        {
            Console.WriteLine("\n📁 EXPORT BY FILE TYPE");

            var fileTypes = await _context.Vocabularies
                .Where(v => v.Category != null)
                .GroupBy(v => v.Category)
                .Select(g => new { Type = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Count)
                .ToListAsync();

            if (!fileTypes.Any())
            {
                Console.WriteLine("❌ No data found in database.");
                return;
            }

            Console.WriteLine("\nAvailable file types:");
            for (int i = 0; i < fileTypes.Count; i++)
            {
                Console.WriteLine($"{i + 1}. {fileTypes[i].Type} ({fileTypes[i].Count} words)");
            }

            Console.Write("\nSelect file type to export: ");
            string choice = Console.ReadLine()?.Trim();

            if (int.TryParse(choice, out int index) && index > 0 && index <= fileTypes.Count)
            {
                string selectedType = fileTypes[index - 1].Type;
                await ExportByFileTypeToExcel(selectedType);
            }
            else
            {
                Console.WriteLine("❌ Invalid selection.");
            }
        }

        private async Task ExportRecentSessions()
        {
            Console.WriteLine("\n🕒 RECENT SESSIONS");

            var recentSessions = await _context.ProcessingSessions
                .OrderByDescending(s => s.ProcessedDate)
                .Take(10)
                .ToListAsync();

            if (!recentSessions.Any())
            {
                Console.WriteLine("❌ No sessions found in database.");
                return;
            }

            Console.WriteLine("\nRecent sessions:");
            for (int i = 0; i < recentSessions.Count; i++)
            {
                var session = recentSessions[i];
                Console.WriteLine($"{i + 1}. {session.ProcessedDate:dd-MM-yyyy HH:mm} - {session.FileType} - {session.ProcessedRows} words");
            }

            Console.Write("\nSelect session to export: ");
            string choice = Console.ReadLine()?.Trim();

            if (int.TryParse(choice, out int index) && index > 0 && index <= recentSessions.Count)
            {
                await ExportSessionToExcel(recentSessions[index - 1]);
            }
            else
            {
                Console.WriteLine("❌ Invalid selection.");
            }
        }

        private async Task ShowDatabaseStatistics()
        {
            Console.WriteLine("\n📊 DATABASE STATISTICS");
            Console.WriteLine(new string('-', 40));

            var totalVocabs = await _context.Vocabularies.CountAsync();
            var totalSessions = await _context.ProcessingSessions.CountAsync();
            var totalAudioFiles = await _context.AudioFiles.CountAsync();

            Console.WriteLine($"📚 Total Vocabularies: {totalVocabs}");
            Console.WriteLine($"📋 Total Sessions: {totalSessions}");
            Console.WriteLine($"🎵 Total Audio Records: {totalAudioFiles}");

            var byCategory = await _context.Vocabularies
                .GroupBy(v => v.Category)
                .Select(g => new { Category = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Count)
                .ToListAsync();

            Console.WriteLine("\n📁 By Category:");
            foreach (var item in byCategory)
            {
                Console.WriteLine($"   {item.Category}: {item.Count} words");
            }

            var recentDates = await _context.Vocabularies
                .GroupBy(v => v.CreatedDate.Date)
                .Select(g => new { Date = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Date)
                .Take(5)
                .ToListAsync();

            Console.WriteLine("\n📅 Recent Activity:");
            foreach (var item in recentDates)
            {
                Console.WriteLine($"   {item.Date:dd-MM-yyyy}: {item.Count} words");
            }
        }

        private async Task ExportSessionToExcel(ProcessingSession session)
        {
            try
            {
                Console.WriteLine($"\n📤 Exporting session {session.SessionId} ({session.FileType})...");

                // Get vocabularies for this session
                var vocabularies = await _context.Vocabularies
                    .Where(v => v.SourceFile == session.FileName &&
                               v.CreatedDate.Date == session.ProcessedDate.Date)
                    .OrderBy(v => v.VocabId)
                    .ToListAsync();

                if (!vocabularies.Any())
                {
                    Console.WriteLine("❌ No vocabulary data found for this session.");
                    return;
                }

                string fileName = $"{session.FileType}_Export_{session.ProcessedDate:dd-MM-yyyy_HHmm}.xlsm";
                string filePath = Path.Combine(ExportPath, fileName);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("Data");

                // Export based on file type
                switch (session.FileType?.ToLower())
                {
                    case "english":
                        ExportEnglishFormat(worksheet, vocabularies);
                        break;
                    case "japanese":
                        ExportJapaneseFormat(worksheet, vocabularies);
                        break;
                    case "chinese":
                        ExportChineseFormat(worksheet, vocabularies);
                        break;
                    case "tuvung":
                        ExportTuVungFormat(worksheet, vocabularies);
                        break;
                    default:
                        ExportGenericFormat(worksheet, vocabularies);
                        break;
                }

                // Save file
                await package.SaveAsAsync(new FileInfo(filePath));
                Console.WriteLine($"✅ Exported to: {fileName}");
                Console.WriteLine($"📁 Location: {filePath}");

                // Ask to open file
                Console.Write("Open exported Excel file? (Y/N): ");
                if (Console.ReadLine()?.Trim().ToUpper() == "Y")
                {
                    ExcelProcessor.OpenExcelFile(filePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Export failed: {ex.Message}");
            }
        }

        private async Task ExportByFileTypeToExcel(string fileType)
        {
            try
            {
                Console.WriteLine($"\n📤 Exporting all {fileType} data...");

                var vocabularies = await _context.Vocabularies
                    .Where(v => v.Category == fileType)
                    .OrderByDescending(v => v.CreatedDate)
                    .ToListAsync();

                string fileName = $"{fileType}_AllData_Export_{DateTime.Now:dd-MM-yyyy_HHmm}.xlsm";
                string filePath = Path.Combine(ExportPath, fileName);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var package = new ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("Data");

                switch (fileType.ToLower())
                {
                    case "english":
                        ExportEnglishFormat(worksheet, vocabularies);
                        break;
                    case "japanese":
                        ExportJapaneseFormat(worksheet, vocabularies);
                        break;
                    case "chinese":
                        ExportChineseFormat(worksheet, vocabularies);
                        break;
                    case "tuvung":
                        ExportTuVungFormat(worksheet, vocabularies);
                        break;
                    default:
                        ExportGenericFormat(worksheet, vocabularies);
                        break;
                }

                await package.SaveAsAsync(new FileInfo(filePath));
                Console.WriteLine($"✅ Exported {vocabularies.Count} words to: {fileName}");

                Console.Write("Open exported Excel file? (Y/N): ");
                if (Console.ReadLine()?.Trim().ToUpper() == "Y")
                {
                    ExcelProcessor.OpenExcelFile(filePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Export failed: {ex.Message}");
            }
        }

        private void ExportEnglishFormat(ExcelWorksheet worksheet, List<Vocabulary> vocabularies)
        {
            // English format: A=Vietnamese, B=English
            for (int i = 0; i < vocabularies.Count; i++)
            {
                var vocab = vocabularies[i];
                int row = i + 1;

                worksheet.Cells[row, 1].Value = vocab.VietnameseText ?? "";
                worksheet.Cells[row, 2].Value = vocab.EnglishText ?? "";
            }
        }

        private void ExportJapaneseFormat(ExcelWorksheet worksheet, List<Vocabulary> vocabularies)
        {
            // Japanese format: A=English, B=Reading, C=Japanese, G=ImageTags
            for (int i = 0; i < vocabularies.Count; i++)
            {
                var vocab = vocabularies[i];
                int row = i + 1;

                worksheet.Cells[row, 1].Value = vocab.EnglishText ?? "";
                worksheet.Cells[row, 2].Value = vocab.ReadingText ?? "";
                worksheet.Cells[row, 3].Value = vocab.JapaneseText ?? "";
                worksheet.Cells[row, 7].Value = vocab.ImageTags ?? "";
            }
        }

        private void ExportChineseFormat(ExcelWorksheet worksheet, List<Vocabulary> vocabularies)
        {
            // Chinese format: A=English, B=Reading, C=Chinese, G=ImageTags
            for (int i = 0; i < vocabularies.Count; i++)
            {
                var vocab = vocabularies[i];
                int row = i + 1;

                worksheet.Cells[row, 1].Value = vocab.EnglishText ?? "";
                worksheet.Cells[row, 2].Value = vocab.ReadingText ?? "";
                worksheet.Cells[row, 3].Value = vocab.ChineseText ?? "";
                worksheet.Cells[row, 7].Value = vocab.ImageTags ?? "";
            }
        }

        private void ExportTuVungFormat(ExcelWorksheet worksheet, List<Vocabulary> vocabularies)
        {
            // TuVung format: A=Vietnamese, B=Reading, C=Japanese, G=ImageTags
            for (int i = 0; i < vocabularies.Count; i++)
            {
                var vocab = vocabularies[i];
                int row = i + 1;

                worksheet.Cells[row, 1].Value = vocab.VietnameseText ?? "";
                worksheet.Cells[row, 2].Value = vocab.ReadingText ?? "";
                worksheet.Cells[row, 3].Value = vocab.JapaneseText ?? "";
                worksheet.Cells[row, 7].Value = vocab.ImageTags ?? "";
            }
        }

        private void ExportGenericFormat(ExcelWorksheet worksheet, List<Vocabulary> vocabularies)
        {
            // Generic format: All columns
            for (int i = 0; i < vocabularies.Count; i++)
            {
                var vocab = vocabularies[i];
                int row = i + 1;

                worksheet.Cells[row, 1].Value = vocab.EnglishText ?? "";
                worksheet.Cells[row, 2].Value = vocab.VietnameseText ?? "";
                worksheet.Cells[row, 3].Value = vocab.JapaneseText ?? "";
                worksheet.Cells[row, 4].Value = vocab.ChineseText ?? "";
                worksheet.Cells[row, 5].Value = vocab.KanjiText ?? "";
                worksheet.Cells[row, 6].Value = vocab.ReadingText ?? "";
                worksheet.Cells[row, 7].Value = vocab.ImageTags ?? "";
            }
        }

        public void Dispose()
        {
            _context?.Dispose();
        }
    }
}