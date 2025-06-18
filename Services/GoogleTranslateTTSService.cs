using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.IO;
using Text_to_Image.Models;
using System.Collections.Generic;
using System.Threading;
using System.Linq;
using Text_to_Image.Services;
using Text_to_Image.Data.Models;
using Text_to_Image.Data;
using OfficeOpenXml;

namespace Text_to_Image.Services
{
    public class GoogleTranslateTTSService : ISpeechService
    {
        private readonly HttpClient _httpClient;
        private readonly SemaphoreSlim _semaphore;

        public GoogleTranslateTTSService()
        {
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Add("User-Agent",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
            _semaphore = new SemaphoreSlim(3); // 3 request cùng lúc như GoogleSpeechService
        }

        public async Task<bool> CreateAudioFile(string text, string outputPath, string voiceName = "en")
        {
            int maxRetries = 3;
            int retryDelay = 1000;

            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(text))
                        return false;

                    // Tạo thư mục nếu chưa tồn tại
                    string directory = Path.GetDirectoryName(outputPath);
                    if (!Directory.Exists(directory))
                        Directory.CreateDirectory(directory);

                    // Xác định language code từ voice name
                    string languageCode = GetLanguageCode(voiceName);

                    // Tối ưu text cho người mới học (thêm pause)
                    string optimizedText = OptimizeTextForLearners(text, languageCode);

                    // Split text thành chunks nhỏ (Google Translate có giới hạn ~100 ký tự)
                    var textChunks = SplitTextIntoChunks(optimizedText, 100);
                    var audioBytes = new List<byte>();

                    foreach (var chunk in textChunks)
                    {
                        var chunkAudio = await GetGoogleTranslateTTS(chunk.Trim(), languageCode);
                        if (chunkAudio != null && chunkAudio.Length > 0)
                        {
                            audioBytes.AddRange(chunkAudio);
                        }

                        // Delay để tránh rate limiting
                        await Task.Delay(200);
                    }

                    if (audioBytes.Count > 0)
                    {
                        await File.WriteAllBytesAsync(outputPath, audioBytes.ToArray());
                        Console.WriteLine($"✓ [GOOGLE-TRANSLATE] {Path.GetFileName(outputPath)} (Language: {languageCode})");
                        return true;
                    }

                    if (attempt < maxRetries - 1)
                    {
                        Console.WriteLine($"⏳ [GOOGLE-TRANSLATE] {Path.GetFileName(outputPath)}: Retrying... (Attempt {attempt + 1}/{maxRetries})");
                        await Task.Delay(retryDelay);
                        retryDelay *= 2; // Exponential backoff
                        continue;
                    }

                    Console.WriteLine($"✗ [GOOGLE-TRANSLATE] {Path.GetFileName(outputPath)}: No audio data received");
                    return false;
                }
                catch (Exception ex)
                {
                    if (attempt < maxRetries - 1)
                    {
                        Console.WriteLine($"⏳ [GOOGLE-TRANSLATE] {Path.GetFileName(outputPath)}: Error, retrying... {ex.Message}");
                        await Task.Delay(retryDelay);
                        retryDelay *= 2;
                        continue;
                    }

                    Console.WriteLine($"✗ [GOOGLE-TRANSLATE] {Path.GetFileName(outputPath)}: {ex.Message}");
                    return false;
                }
            }

            return false;
        }

        // MAIN METHOD: Process Excel for Audio với logic duplicate từ GoogleSpeechService
        public async Task ProcessExcelForAudio(ProcessingOptions options)
        {
            try
            {
                Console.WriteLine("Creating audio files with GOOGLE TRANSLATE Text-to-Speech...");

                string dateToUse = string.IsNullOrWhiteSpace(options.CustomDate) ?
                    DateTime.Now.ToString("dd-MM-yyyy") : options.CustomDate;

                // Lấy existing vocabularies từ database để check duplicates
                List<Vocabulary> existingVocabs = new List<Vocabulary>();
                try
                {
                    using var dbService = new DatabaseService();
                    string fileType = DetermineFileType(options.FileName);
                    existingVocabs = await dbService.GetExistingVocabulariesForDateAsync(dateToUse, fileType);
                    Console.WriteLine($"Found {existingVocabs.Count} existing vocabularies for {dateToUse} ({fileType})");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not load existing vocabularies: {ex.Message}");
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using var package = new ExcelPackage(new FileInfo(options.SelectedFile));
                var worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension?.End.Row ?? 0;
                if (rowCount == 0)
                {
                    Console.WriteLine("No data in Excel file to process.");
                    return;
                }

                var audioInfoList = new List<AudioCreationInfo>();
                var processedInSession = new List<Vocabulary>();
                var processingResults = new List<AudioProcessingInfo>();

                // PHASE 1: Collect processing info
                Console.WriteLine("PHASE 1: Collecting audio processing information...");
                for (int row = 1; row <= rowCount; row++)
                {
                    var currentVocab = CreateVocabularyFromRow(worksheet, row, options);
                    var duplicateResult = CheckRowForDuplicateAudio(currentVocab, existingVocabs, processedInSession, options);

                    var processingInfo = new AudioProcessingInfo
                    {
                        Vocabulary = currentVocab,
                        DuplicateResult = duplicateResult,
                        RowNumber = row,
                        HasAudioFormulas = CheckRowHasAudioFormulas(worksheet, row, options)
                    };

                    processingResults.Add(processingInfo);

                    // Track CreateNew entries for session duplicate checking
                    if (duplicateResult.Action == "CreateNew")
                    {
                        processedInSession.Add(currentVocab);
                    }
                }

                // PHASE 2: Apply conflict resolution logic
                Console.WriteLine("PHASE 2: Applying audio conflict resolution logic...");
                ApplyAudioConflictResolutionLogic(processingResults, existingVocabs, options);

                // PHASE 2.5: Track additional entries that became CreateNewConflict
                Console.WriteLine("PHASE 2.5: Updating processed rows tracking for CreateNewConflict entries...");
                foreach (var info in processingResults)
                {
                    if (info.DuplicateResult.Action == "CreateNewConflict" &&
                        !processedInSession.Any(p => SimilarVocabulary(p, info.Vocabulary, DetermineFileType(options.FileName))))
                    {
                        processedInSession.Add(info.Vocabulary);
                        Console.WriteLine($"  Added CreateNewConflict row {info.RowNumber} to processed tracking");
                    }
                }

                // PHASE 3: Execute audio creation based on processing results
                Console.WriteLine("PHASE 3: Executing audio creation actions...");
                await CollectAudioCreationInformation(worksheet, processingResults, options, audioInfoList);

                // Execute tất cả audio tasks với semaphore
                //if (audioTasks.Count > 0)
                //{
                //    Console.WriteLine($"Creating {audioTasks.Count} audio files with Google Translate TTS optimized for learners...");

                //    // Tạo âm thanh song song (giới hạn 3 files cùng lúc)
                //    var tasks = audioTasks.Select(async task =>
                //    {
                //        await _semaphore.WaitAsync();
                //        try
                //        {
                //            await task;
                //        }
                //        finally
                //        {
                //            _semaphore.Release();
                //        }
                //    });

                //    await Task.WhenAll(tasks);

                //    Console.WriteLine($"Completed! Created {audioTasks.Count} audio files with Google Translate TTS.");
                //}

                // Execute tất cả audio tasks TUẦN TỰ (thay vì song song)
                if (audioInfoList.Count > 0)
                {
                    // Sắp xếp theo filename để đảm bảo thứ tự
                    var sortedAudioList = audioInfoList
                        .OrderBy(info => ExtractNumberFromFileName(info.FileName))
                        .ThenBy(info => info.FileName)
                        .ToList();

                    Console.WriteLine("=== AUDIO FILES ORDER ===");
                    for (int i = 0; i < sortedAudioList.Count; i++)
                    {
                        Console.WriteLine($"{i + 1:D2}. {sortedAudioList[i].FileName}");
                    }
                    Console.WriteLine("==========================");

                    // Tạo từng file theo thứ tự đã sắp xếp
                    for (int i = 0; i < sortedAudioList.Count; i++)
                    {
                        var audioInfo = sortedAudioList[i];
                        Console.WriteLine($"[{i + 1:D2}/{sortedAudioList.Count:D2}] {audioInfo.FileName}");

                        bool success = await CreateAudioFile(audioInfo.Text, audioInfo.OutputPath, audioInfo.LanguageCode);
                        if (!success)
                        {
                            Console.WriteLine($"✗ FAILED: {audioInfo.FileName}");
                        }
                        await Task.Delay(100);
                    }

                    Console.WriteLine($"Completed {sortedAudioList.Count} audio files.");
                }
                else
                {
                    Console.WriteLine("No audio files to create.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating audio files: {ex.Message}");
                throw;
            }
        }

        // Tối ưu text cho người mới học
        private string OptimizeTextForLearners(string text, string languageCode)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;

            // Thêm dấu chấm và pause cho text dài
            string optimizedText = text.Trim();

            // Thêm pause giữa các từ cho ngôn ngữ khó
            switch (languageCode)
            {
                case "ja":  // Japanese
                case "zh":  // Chinese
                    // Thêm dấu phẩy để tạo pause tự nhiên
                    optimizedText = System.Text.RegularExpressions.Regex.Replace(optimizedText, @"(\S+)", "$1,");
                    optimizedText = optimizedText.TrimEnd(',');
                    break;

                case "en":  // English
                    // Thêm dấu chấm cuối nếu chưa có
                    if (!optimizedText.EndsWith(".") && !optimizedText.EndsWith("!") && !optimizedText.EndsWith("?"))
                    {
                        optimizedText += ".";
                    }
                    break;
            }

            return optimizedText;
        }

        private async Task<byte[]> GetGoogleTranslateTTS(string text, string languageCode)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(text)) return null;

                // URL của Google Translate TTS (unofficial API)
                var encodedText = Uri.EscapeDataString(text);
                var url = $"https://translate.google.com/translate_tts?ie=UTF-8&total=1&idx=0&client=tw-ob&tl={languageCode}&q={encodedText}&textlen={text.Length}";

                var response = await _httpClient.GetAsync(url);

                if (response.IsSuccessStatusCode)
                {
                    var audioData = await response.Content.ReadAsByteArrayAsync();
                    if (audioData.Length > 0)
                        return audioData;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private string GetLanguageCode(string voiceName)
        {
            return voiceName switch
            {
                var name when name.Contains("EN") || name.Contains("en-US") || name.Equals("en") => "en",
                var name when name.Contains("VI") || name.Contains("vi-VN") || name.Equals("vi") => "vi",
                var name when name.Contains("JP") || name.Contains("ja-JP") || name.Equals("ja") => "ja",
                var name when name.Contains("ZH") || name.Contains("cmn-CN") || name.Equals("zh") => "zh",
                _ => "en"
            };
        }

        private List<string> SplitTextIntoChunks(string text, int maxLength)
        {
            var chunks = new List<string>();

            if (text.Length <= maxLength)
            {
                chunks.Add(text);
                return chunks;
            }

            // Split by sentences first
            var sentences = text.Split(new char[] { '.', '!', '?', ';' }, StringSplitOptions.RemoveEmptyEntries);
            var currentChunk = "";

            foreach (var sentence in sentences)
            {
                var trimmedSentence = sentence.Trim();
                if (string.IsNullOrEmpty(trimmedSentence)) continue;

                if ((currentChunk + " " + trimmedSentence).Length <= maxLength)
                {
                    currentChunk += (string.IsNullOrEmpty(currentChunk) ? "" : " ") + trimmedSentence;
                }
                else
                {
                    if (!string.IsNullOrEmpty(currentChunk))
                        chunks.Add(currentChunk.Trim());

                    // If single sentence is too long, split by words
                    if (trimmedSentence.Length > maxLength)
                    {
                        var words = trimmedSentence.Split(' ');
                        var wordChunk = "";

                        foreach (var word in words)
                        {
                            if ((wordChunk + " " + word).Length <= maxLength)
                            {
                                wordChunk += (string.IsNullOrEmpty(wordChunk) ? "" : " ") + word;
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(wordChunk))
                                    chunks.Add(wordChunk.Trim());
                                wordChunk = word;
                            }
                        }

                        if (!string.IsNullOrEmpty(wordChunk))
                            chunks.Add(wordChunk.Trim());

                        currentChunk = "";
                    }
                    else
                    {
                        currentChunk = trimmedSentence;
                    }
                }
            }

            if (!string.IsNullOrEmpty(currentChunk))
                chunks.Add(currentChunk.Trim());

            return chunks.Where(c => !string.IsNullOrWhiteSpace(c)).ToList();
        }

        #region Duplicate Logic Methods (từ GoogleSpeechService)

        private void ApplyAudioConflictResolutionLogic(List<AudioProcessingInfo> processingResults, List<Vocabulary> existingVocabs, ProcessingOptions options)
        {
            string fileType = DetermineFileType(options.FileName);
            var partialMatchGroups = new Dictionary<string, List<AudioProcessingInfo>>();

            // Group partial matches by existing vocab
            foreach (var info in processingResults)
            {
                if (info.DuplicateResult.IsDuplicate && info.DuplicateResult.MatchType == "Partial")
                {
                    string groupKey = $"{info.DuplicateResult.ExistingVocab.VocabId}";

                    if (!partialMatchGroups.ContainsKey(groupKey))
                        partialMatchGroups[groupKey] = new List<AudioProcessingInfo>();

                    partialMatchGroups[groupKey].Add(info);
                }
            }

            // Apply conflict resolution
            foreach (var group in partialMatchGroups.Values)
            {
                if (group.Count > 1)
                {
                    // Multiple rows match same existing vocab
                    for (int i = 0; i < group.Count; i++)
                    {
                        var info = group[i];
                        var conflictFields = GetConflictFields(info.Vocabulary, info.DuplicateResult.ExistingVocab, fileType);

                        if (i == 0)
                        {
                            // First row: RESTORE old audio
                            info.DuplicateResult.Action = "RestoreAndUpdate";
                            Console.WriteLine($"Row {info.RowNumber}: CONFLICT (FIRST) in [{string.Join(", ", conflictFields)}] - RESTORED old audio");
                        }
                        else
                        {
                            // Subsequent rows: CREATE new audio
                            info.DuplicateResult.Action = "CreateNewConflict";
                            Console.WriteLine($"Row {info.RowNumber}: CONFLICT (SUBSEQUENT) in [{string.Join(", ", conflictFields)}] - Created NEW audio");
                        }
                    }
                }
            }
        }

        private DuplicateCheckResult CheckRowForDuplicateAudio(Vocabulary currentVocab, List<Vocabulary> existingVocabs, List<Vocabulary> processedInSession, ProcessingOptions options)
        {
            string fileType = DetermineFileType(options.FileName);

            // CHECK 1: EXACT MATCH trong database
            var exactMatch = FindExactMatchInDatabase(currentVocab, existingVocabs, fileType);
            if (exactMatch != null)
            {
                return new DuplicateCheckResult
                {
                    IsDuplicate = true,
                    ExistingVocab = exactMatch,
                    DuplicateSource = "Database",
                    MatchType = "Exact",
                    Action = "Skip"
                };
            }

            // CHECK 2: PARTIAL MATCH trong database  
            var partialMatch = FindPartialMatchInDatabase(currentVocab, existingVocabs, fileType);
            if (partialMatch != null)
            {
                return new DuplicateCheckResult
                {
                    IsDuplicate = true,
                    ExistingVocab = partialMatch,
                    DuplicateSource = "Database",
                    MatchType = "Partial",
                    Action = "RestoreAndUpdate"
                };
            }

            // CHECK 3: Current session exact duplicates
            var sessionExactMatch = FindExactMatchInCurrentSession(currentVocab, processedInSession, fileType);
            if (sessionExactMatch != null)
            {
                return new DuplicateCheckResult
                {
                    IsDuplicate = true,
                    ExistingVocab = sessionExactMatch,
                    DuplicateSource = "CurrentSession",
                    MatchType = "Exact",
                    Action = "Skip"
                };
            }

            // CHECK 4: Current session partial duplicates
            var sessionPartialMatch = FindPartialMatchInCurrentSession(currentVocab, processedInSession, fileType);
            if (sessionPartialMatch != null)
            {
                return new DuplicateCheckResult
                {
                    IsDuplicate = true,
                    ExistingVocab = sessionPartialMatch,
                    DuplicateSource = "CurrentSession",
                    MatchType = "Partial",
                    Action = "Skip"
                };
            }

            return new DuplicateCheckResult { IsDuplicate = false, Action = "CreateNew" };
        }
        private async Task CollectAudioCreationInformation(ExcelWorksheet worksheet, List<AudioProcessingInfo> processingResults, ProcessingOptions options, List<AudioCreationInfo> audioInfoList)
        {
            foreach (var info in processingResults)
            {
                if (!info.HasAudioFormulas) continue;

                switch (info.DuplicateResult.Action)
                {
                    case "Skip":
                        break;
                    case "RestoreAndUpdate":
                        await CollectAudioInfoForDatabaseDuplicate(worksheet, info.RowNumber, info.DuplicateResult.ExistingVocab, options, audioInfoList);
                        break;
                    case "CreateNew":
                    case "CreateNewConflict":
                        await CollectAudioInfoForNewEntry(worksheet, info.RowNumber, options, audioInfoList);
                        break;
                }
            }
        }

        private async Task CollectAudioInfoForDatabaseDuplicate(ExcelWorksheet worksheet, int row, Vocabulary existingVocab, ProcessingOptions options, List<AudioCreationInfo> audioInfoList)
        {
            try
            {
                using var dbService = new DatabaseService();
                var existingAudioFiles = await dbService.GetAudioFilesByVocabIdAsync(existingVocab.VocabId);
                if (existingAudioFiles == null || existingAudioFiles.Count == 0) return;

                foreach (var audioFile in existingAudioFiles)
                {
                    string newText = GetTextForAudio(worksheet, row, audioFile.Language, options);
                    if (!string.IsNullOrEmpty(newText))
                    {
                        audioInfoList.Add(new AudioCreationInfo
                        {
                            FileName = audioFile.FileName,
                            Text = newText,
                            OutputPath = Path.Combine(options.AudioOutputFolder, audioFile.FileName),
                            LanguageCode = GetLanguageCode(audioFile.Language),
                            RowNumber = row
                        });
                    }
                }
            }
            catch { }
        }

        private async Task CollectAudioInfoForNewEntry(ExcelWorksheet worksheet, int row, ProcessingOptions options, List<AudioCreationInfo> audioInfoList)
        {
            if (!CheckRowHasAudioFormulas(worksheet, row, options)) return;

            var soundFormulas = ExtractAudioFilesFromRow(worksheet, row, options);

            foreach (var audioInfo in soundFormulas)
            {
                string text = GetTextForAudio(worksheet, row, ExtractLanguageFromFileName(audioInfo.FileName), options);
                if (!string.IsNullOrEmpty(text))
                {
                    audioInfoList.Add(new AudioCreationInfo
                    {
                        FileName = audioInfo.FileName,
                        Text = text,
                        OutputPath = Path.Combine(options.AudioOutputFolder, audioInfo.FileName),
                        LanguageCode = GetLanguageCode(ExtractLanguageFromFileName(audioInfo.FileName)),
                        RowNumber = row
                    });
                }
            }
        }

        private int ExtractNumberFromFileName(string fileName)
        {
            try
            {
                // Ví dụ: JP-18-06-2025_03.mp3 -> extract 03
                var parts = fileName.Split('_');
                if (parts.Length >= 2)
                {
                    var numberPart = parts[1].Split('.')[0];
                    if (int.TryParse(numberPart, out int number))
                    {
                        return number;
                    }
                }
            }
            catch { }
            return 0;
        }


        #endregion

        #region Helper Methods (từ GoogleSpeechService)

        private Vocabulary CreateVocabularyFromRow(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            var vocab = new Vocabulary();
            string fileType = DetermineFileType(options.FileName);

            switch (fileType.ToLower())
            {
                case "english":
                    vocab.VietnameseText = GetCellValue(worksheet, row, 1);
                    vocab.EnglishText = GetCellValue(worksheet, row, 2);
                    break;
                case "japanese":
                    vocab.EnglishText = GetCellValue(worksheet, row, 1);
                    vocab.ReadingText = GetCellValue(worksheet, row, 2);
                    vocab.JapaneseText = GetCellValue(worksheet, row, 3);
                    break;
                case "chinese":
                    vocab.EnglishText = GetCellValue(worksheet, row, 1);
                    vocab.ReadingText = GetCellValue(worksheet, row, 2);
                    vocab.ChineseText = GetCellValue(worksheet, row, 3);
                    break;
                case "tuvung":
                    vocab.VietnameseText = GetCellValue(worksheet, row, 1);
                    vocab.ReadingText = GetCellValue(worksheet, row, 2);
                    vocab.JapaneseText = GetCellValue(worksheet, row, 3);
                    break;
            }

            return vocab;
        }

        private bool CheckRowHasAudioFormulas(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            if (string.IsNullOrWhiteSpace(options.SoundColumns)) return false;

            if ((options.FileName.Contains("tuvung") ||
                 options.FileName.Contains("japanese") ||
                 options.FileName.Contains("chinese")) && options.SoundColumns.Length == 2)
            {
                string col1Value = worksheet.Cells[row, 5].Text?.Trim();
                string col2Value = worksheet.Cells[row, 6].Text?.Trim();

                return !string.IsNullOrEmpty(col1Value) && !string.IsNullOrEmpty(col2Value) &&
                       col1Value.Contains("[sound:") && col2Value.Contains("[sound:");
            }
            else if (options.FileName.Contains("english") && options.SoundColumns.Length == 2)
            {
                int soundCol1 = options.SoundColumns[0] - 'A' + 1;
                int soundCol2 = options.SoundColumns[1] - 'A' + 1;

                string col1Value = worksheet.Cells[row, soundCol1].Text?.Trim();
                string col2Value = worksheet.Cells[row, soundCol2].Text?.Trim();

                return !string.IsNullOrEmpty(col1Value) && !string.IsNullOrEmpty(col2Value) &&
                       col1Value.Contains("[sound:") && col2Value.Contains("[sound:");
            }

            return false;
        }

        private List<AudioFileInfo> ExtractAudioFilesFromRow(ExcelWorksheet worksheet, int row, ProcessingOptions options)
        {
            var audioFiles = new List<AudioFileInfo>();

            if ((options.FileName.Contains("tuvung") ||
                 options.FileName.Contains("japanese") ||
                 options.FileName.Contains("chinese")) && options.SoundColumns.Length == 2)
            {
                string col1Formula = worksheet.Cells[row, 5].Text?.Trim(); // Column E
                string col2Formula = worksheet.Cells[row, 6].Text?.Trim(); // Column F

                if (!string.IsNullOrEmpty(col1Formula) && col1Formula.Contains("[sound:"))
                {
                    string fileName = ExtractFileNameFromFormula(col1Formula);
                    string lang = ExtractLanguageFromFileName(fileName);
                    string sourceCol = GetSourceColumnForLanguage(lang, options);

                    audioFiles.Add(new AudioFileInfo
                    {
                        RowIndex = row,
                        SourceColumn = sourceCol,
                        FileName = fileName,
                        VoiceName = GetLanguageCode(lang)
                    });
                }

                if (!string.IsNullOrEmpty(col2Formula) && col2Formula.Contains("[sound:"))
                {
                    string fileName = ExtractFileNameFromFormula(col2Formula);
                    string lang = ExtractLanguageFromFileName(fileName);
                    string sourceCol = GetSourceColumnForLanguage(lang, options);

                    audioFiles.Add(new AudioFileInfo
                    {
                        RowIndex = row,
                        SourceColumn = sourceCol,
                        FileName = fileName,
                        VoiceName = GetLanguageCode(lang)
                    });
                }
            }
            else if (options.FileName.Contains("english") && options.SoundColumns.Length == 2)
            {
                // XỬ LÝ CHO ENGLISH FILES
                int soundCol1 = options.SoundColumns[0] - 'A' + 1;
                int soundCol2 = options.SoundColumns[1] - 'A' + 1;

                string col1Formula = worksheet.Cells[row, soundCol1].Text?.Trim();
                string col2Formula = worksheet.Cells[row, soundCol2].Text?.Trim();

                if (!string.IsNullOrEmpty(col1Formula) && col1Formula.Contains("[sound:"))
                {
                    string fileName = ExtractFileNameFromFormula(col1Formula);
                    string lang = ExtractLanguageFromFileName(fileName);
                    string sourceCol = GetSourceColumnForLanguage(lang, options);

                    audioFiles.Add(new AudioFileInfo
                    {
                        RowIndex = row,
                        SourceColumn = sourceCol,
                        FileName = fileName,
                        VoiceName = GetLanguageCode(lang)
                    });
                }

                if (!string.IsNullOrEmpty(col2Formula) && col2Formula.Contains("[sound:"))
                {
                    string fileName = ExtractFileNameFromFormula(col2Formula);
                    string lang = ExtractLanguageFromFileName(fileName);
                    string sourceCol = GetSourceColumnForLanguage(lang, options);

                    audioFiles.Add(new AudioFileInfo
                    {
                        RowIndex = row,
                        SourceColumn = sourceCol,
                        FileName = fileName,
                        VoiceName = GetLanguageCode(lang)
                    });
                }
            }

            return audioFiles;
        }

        private string ExtractLanguageFromFileName(string fileName)
        {
            // Extract language from filename based on odd/even number
            if (fileName.Contains("_"))
            {
                var parts = fileName.Split('_');
                if (parts.Length >= 2)
                {
                    var numberPart = parts[1].Split('.')[0];

                    if (int.TryParse(numberPart, out int number))
                    {
                        bool isOdd = (number % 2 == 1);

                        // Xác định language dựa vào file type + odd/even
                        if (fileName.StartsWith("JP-"))
                        {
                            return isOdd ? "EN" : "JP";
                        }
                        else if (fileName.StartsWith("ZH-"))
                        {
                            return isOdd ? "EN" : "ZH";
                        }
                        else if (fileName.StartsWith("Vocab-"))
                        {
                            return isOdd ? "EN" : "VI";
                        }
                        else if (fileName.StartsWith("EN-"))
                        {
                            return isOdd ? "VI" : "EN";
                        }
                    }
                }
            }

            return "EN"; // Default
        }

        private string GetSourceColumnForLanguage(string language, ProcessingOptions options)
        {
            if (options.FileName.Contains("tuvung"))
            {
                return "A"; // Cả EN và VI voice đều đọc Vietnamese text từ Column A
            }
            else if (options.FileName.Contains("japanese"))
            {
                return language switch
                {
                    "EN" => "A", // English voice đọc English text từ Column A
                    "JP" => "C", // Japanese voice đọc Japanese text từ Column C  
                    _ => "A"
                };
            }
            else if (options.FileName.Contains("chinese"))
            {
                return language switch
                {
                    "EN" => "A", // English voice đọc English text từ Column A
                    "ZH" => "C", // Chinese voice đọc Chinese text từ Column C
                    _ => "A"
                };
            }
            else if (options.FileName.Contains("english"))
            {
                return language switch
                {
                    "VI" => "A", // Vietnamese voice đọc Vietnamese text từ Column A
                    "EN" => "B", // English voice đọc English text từ Column B
                    _ => "A"
                };
            }

            return "A"; // Default
        }

        private string GetTextForAudio(ExcelWorksheet worksheet, int row, string language, ProcessingOptions options)
        {
            string sourceColumn = GetSourceColumnForLanguage(language, options);
            int colIndex = GetColumnIndex(sourceColumn);
            return worksheet.Cells[row, colIndex].Text?.Trim() ?? "";
        }

        private string ExtractFileNameFromFormula(string formula)
        {
            int start = formula.IndexOf("[sound:") + 7;
            int end = formula.IndexOf("]", start);
            if (start > 6 && end > start)
            {
                return formula.Substring(start, end - start);
            }
            return "";
        }

        private int GetColumnIndex(string columnLetter)
        {
            if (string.IsNullOrEmpty(columnLetter))
                return 1;

            int result = 0;
            for (int i = 0; i < columnLetter.Length; i++)
            {
                result = result * 26 + (columnLetter[i] - 'A' + 1);
            }
            return result;
        }

        private string GetCellValue(ExcelWorksheet worksheet, int row, int column)
        {
            var value = worksheet.Cells[row, column].Text?.Trim();
            return string.IsNullOrEmpty(value) ? "" : value;
        }

        private bool SimilarText(string text1, string text2)
        {
            if (string.IsNullOrWhiteSpace(text1) || string.IsNullOrWhiteSpace(text2))
                return false;
            return text1.Trim().Equals(text2.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private string DetermineFileType(string fileName)
        {
            fileName = fileName.ToLower();
            if (fileName.Contains("english")) return "English";
            if (fileName.Contains("japanese")) return "Japanese";
            if (fileName.Contains("chinese")) return "Chinese";
            if (fileName.Contains("tuvung")) return "TuVung";
            return "Unknown";
        }

        private bool SimilarVocabulary(Vocabulary vocab1, Vocabulary vocab2, string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => SimilarText(vocab1.VietnameseText, vocab2.VietnameseText) &&
                             SimilarText(vocab1.EnglishText, vocab2.EnglishText),
                "japanese" => SimilarText(vocab1.EnglishText, vocab2.EnglishText) &&
                              SimilarText(vocab1.JapaneseText, vocab2.JapaneseText),
                "chinese" => SimilarText(vocab1.EnglishText, vocab2.EnglishText) &&
                             SimilarText(vocab1.ChineseText, vocab2.ChineseText),
                "tuvung" => SimilarText(vocab1.VietnameseText, vocab2.VietnameseText) &&
                            SimilarText(vocab1.JapaneseText, vocab2.JapaneseText),
                _ => false
            };
        }

        private List<string> GetConflictFields(Vocabulary current, Vocabulary existing, string fileType)
        {
            var conflicts = new List<string>();

            switch (fileType.ToLower())
            {
                case "english":
                    if (!SimilarText(current.VietnameseText, existing.VietnameseText))
                        conflicts.Add("Vietnamese");
                    if (!SimilarText(current.EnglishText, existing.EnglishText))
                        conflicts.Add("English");
                    break;

                case "japanese":
                    if (!SimilarText(current.EnglishText, existing.EnglishText))
                        conflicts.Add("English");
                    if (!SimilarText(current.ReadingText, existing.ReadingText))
                        conflicts.Add("Reading");
                    if (!SimilarText(current.JapaneseText, existing.JapaneseText))
                        conflicts.Add("Japanese");
                    break;

                case "chinese":
                    if (!SimilarText(current.EnglishText, existing.EnglishText))
                        conflicts.Add("English");
                    if (!SimilarText(current.ReadingText, existing.ReadingText))
                        conflicts.Add("Reading");
                    if (!SimilarText(current.ChineseText, existing.ChineseText))
                        conflicts.Add("Chinese");
                    break;

                case "tuvung":
                    if (!SimilarText(current.VietnameseText, existing.VietnameseText))
                        conflicts.Add("Vietnamese");
                    if (!SimilarText(current.ReadingText, existing.ReadingText))
                        conflicts.Add("Reading");
                    if (!SimilarText(current.JapaneseText, existing.JapaneseText))
                        conflicts.Add("Japanese");
                    break;
            }

            return conflicts;
        }

        private Vocabulary FindExactMatchInDatabase(Vocabulary currentVocab, List<Vocabulary> existingVocabs, string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.VietnameseText, currentVocab.VietnameseText) &&
                    SimilarText(existing.EnglishText, currentVocab.EnglishText)
                ),

                "japanese" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.EnglishText, currentVocab.EnglishText) &&
                    SimilarText(existing.JapaneseText, currentVocab.JapaneseText)
                ),

                "chinese" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.EnglishText, currentVocab.EnglishText) &&
                    SimilarText(existing.ChineseText, currentVocab.ChineseText)
                ),

                "tuvung" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.VietnameseText, currentVocab.VietnameseText) &&
                    SimilarText(existing.JapaneseText, currentVocab.JapaneseText)
                ),

                _ => null
            };
        }

        private Vocabulary FindPartialMatchInDatabase(Vocabulary currentVocab, List<Vocabulary> existingVocabs, string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.VietnameseText, currentVocab.VietnameseText) ||
                    SimilarText(existing.EnglishText, currentVocab.EnglishText)
                ),

                "japanese" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.EnglishText, currentVocab.EnglishText) ||
                    SimilarText(existing.JapaneseText, currentVocab.JapaneseText)
                ),

                "chinese" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.EnglishText, currentVocab.EnglishText) ||
                    SimilarText(existing.ChineseText, currentVocab.ChineseText)
                ),

                "tuvung" => existingVocabs.FirstOrDefault(existing =>
                    SimilarText(existing.VietnameseText, currentVocab.VietnameseText) ||
                    SimilarText(existing.JapaneseText, currentVocab.JapaneseText)
                ),

                _ => null
            };
        }

        private Vocabulary FindExactMatchInCurrentSession(Vocabulary currentVocab, List<Vocabulary> processedInSession, string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => processedInSession.FirstOrDefault(processed =>
                    SimilarText(processed.VietnameseText, currentVocab.VietnameseText) &&
                    SimilarText(processed.EnglishText, currentVocab.EnglishText)
                ),

                "japanese" => processedInSession.FirstOrDefault(processed =>
                    SimilarText(processed.EnglishText, currentVocab.EnglishText) &&
                    SimilarText(processed.JapaneseText, currentVocab.JapaneseText)
                ),

                "chinese" => processedInSession.FirstOrDefault(processed =>
                    SimilarText(processed.EnglishText, currentVocab.EnglishText) &&
                    SimilarText(processed.ChineseText, currentVocab.ChineseText)
                ),

                "tuvung" => processedInSession.FirstOrDefault(processed =>
                    SimilarText(processed.VietnameseText, currentVocab.VietnameseText) &&
                    SimilarText(processed.JapaneseText, currentVocab.JapaneseText)
                ),

                _ => null
            };
        }

        private Vocabulary FindPartialMatchInCurrentSession(Vocabulary currentVocab, List<Vocabulary> processedInSession, string fileType)
        {
            return fileType.ToLower() switch
            {
                "english" => processedInSession.FirstOrDefault(processed =>
                    SimilarText(processed.VietnameseText, currentVocab.VietnameseText) &&
                    SimilarText(processed.EnglishText, currentVocab.EnglishText)
                ),

                "japanese" => processedInSession.FirstOrDefault(processed =>
                    SimilarText(processed.EnglishText, currentVocab.EnglishText) &&
                    SimilarText(processed.JapaneseText, currentVocab.JapaneseText)
                ),

                "chinese" => processedInSession.FirstOrDefault(processed =>
                    SimilarText(processed.EnglishText, currentVocab.EnglishText) &&
                    SimilarText(processed.ChineseText, currentVocab.ChineseText)
                ),

                "tuvung" => processedInSession.FirstOrDefault(processed =>
                    SimilarText(processed.VietnameseText, currentVocab.VietnameseText) &&
                    SimilarText(processed.JapaneseText, currentVocab.JapaneseText)
                ),

                _ => null
            };
        }

        #endregion

        public string GetSSMLPreview(string text, string voiceName)
        {
            var languageCode = GetLanguageCode(voiceName);
            var optimizedText = OptimizeTextForLearners(text, languageCode);
            return $"Google Translate TTS: {optimizedText} (Language: {languageCode})";
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
            _semaphore?.Dispose();
        }

        // Helper classes
        public class DuplicateCheckResult
        {
            public bool IsDuplicate { get; set; }
            public Vocabulary ExistingVocab { get; set; }
            public string DuplicateSource { get; set; } // "Database" or "CurrentSession"
            public string MatchType { get; set; } = ""; // "Exact" or "Partial"  
            public string Action { get; set; } = ""; // "Skip", "RestoreAndUpdate", "Clear", "CreateNew"
        }

        public class AudioProcessingInfo
        {
            public Vocabulary Vocabulary { get; set; }
            public DuplicateCheckResult DuplicateResult { get; set; }
            public int RowNumber { get; set; }
            public bool HasAudioFormulas { get; set; }
        }

        public class AudioFileInfo
        {
            public int RowIndex { get; set; }
            public string SourceColumn { get; set; }
            public string FileName { get; set; }
            public string VoiceName { get; set; }
        }
        public class AudioCreationInfo
        {
            public string FileName { get; set; }
            public string Text { get; set; }
            public string OutputPath { get; set; }
            public string LanguageCode { get; set; }
            public int RowNumber { get; set; }
        }
    }
}