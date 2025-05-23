using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Text_to_Image.Services
{
    public static class FolderSelector
    {
        public static string SelectAudioFolder(string defaultPath = @"C:\Audio")
        {
            Console.WriteLine("\n========== AUDIO FOLDER SELECTION ==========");
            Console.WriteLine("Note: This is for selecting the AUDIO files folder");
            Console.WriteLine("(Excel files are stored separately in S:\\Anki)");
            Console.WriteLine("\nChoose how you want to select the audio folder:");
            Console.WriteLine("1. Use default audio folder");
            Console.WriteLine("2. Open Windows Explorer (GUI) - Recommended");
            Console.WriteLine("3. Browse folders in console");
            Console.WriteLine("4. Enter custom folder path manually");
            Console.Write("Enter your choice (1-4): ");

            string choice = Console.ReadLine()?.Trim();

            switch (choice)
            {
                case "1":
                    return UseDefaultFolder(defaultPath);

                case "2":
                    return OpenWindowsExplorer(defaultPath);

                case "3":
                    return BrowseAndSelectFolder();

                case "4":
                    return EnterCustomPath();

                default:
                    Console.WriteLine("Invalid choice. Opening Windows Explorer.");
                    return OpenWindowsExplorer(defaultPath);
            }
        }

        private static string OpenWindowsExplorer(string defaultPath)
        {
            try
            {
                Console.WriteLine("\n--- WINDOWS EXPLORER FOLDER SELECTION ---");
                Console.WriteLine("📁 How this works:");
                Console.WriteLine("   1. Windows Explorer will open");
                Console.WriteLine("   2. Navigate to your audio folder containing .mp3 files");
                Console.WriteLine("   3. Click on the address bar and copy the full path");
                Console.WriteLine("   4. Come back to this console window");
                Console.WriteLine("   5. Paste the folder path when prompted");

                Console.WriteLine("\n💡 Tips:");
                Console.WriteLine("   - Look for files like: 19-2-2025-0001.mp3");
                Console.WriteLine("   - You can also drag & drop the folder into this console");

                Console.WriteLine("\nPress Enter to open Windows Explorer...");
                Console.ReadLine();

                // Determine which folder to open
                string explorerPath;
                if (Directory.Exists(defaultPath))
                {
                    explorerPath = defaultPath;
                    Console.WriteLine($"Opening Explorer at: {defaultPath}");
                }
                else
                {
                    // Try common audio locations
                    string[] commonAudioPaths = {
                        Environment.GetFolderPath(Environment.SpecialFolder.MyMusic),
                        @"C:\Audio",
                        @"D:\Audio",
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"),
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                        @"C:\"
                    };

                    explorerPath = @"C:\";
                    foreach (string path in commonAudioPaths)
                    {
                        if (Directory.Exists(path))
                        {
                            explorerPath = path;
                            break;
                        }
                    }
                    Console.WriteLine($"Opening Explorer at: {explorerPath}");
                }

                // Open Windows Explorer
                Process.Start("explorer.exe", $"\"{explorerPath}\"");

                // Wait a moment for Explorer to open
                System.Threading.Thread.Sleep(1500);

                Console.WriteLine("\n🔍 Explorer is now open. Please navigate to your audio folder.");
                Console.WriteLine("📋 Copy the folder path from the address bar (Ctrl+L then Ctrl+C)");

                return GetFolderPathFromUser();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error opening Windows Explorer: {ex.Message}");
                Console.WriteLine("Falling back to manual input method...");
                return EnterCustomPath();
            }
        }

        private static string GetFolderPathFromUser()
        {
            int attempts = 0;
            const int maxAttempts = 3;

            while (attempts < maxAttempts)
            {
                Console.Write("\n📁 Paste your audio folder path here: ");
                string userPath = Console.ReadLine()?.Trim();

                if (string.IsNullOrWhiteSpace(userPath))
                {
                    Console.WriteLine("❌ No path entered.");
                    attempts++;
                    continue;
                }

                // Clean up the path (remove quotes, trim whitespace)
                userPath = userPath.Trim('"').Trim('\'').Trim();

                // Handle drag & drop paths that might have extra characters
                if (userPath.StartsWith("&"))
                {
                    userPath = userPath.Substring(1).Trim();
                }

                if (Directory.Exists(userPath))
                {
                    // Check if folder contains any .mp3 files
                    string[] mp3Files = Directory.GetFiles(userPath, "*.mp3");

                    Console.WriteLine($"✅ Valid folder selected: {userPath}");
                    Console.WriteLine($"📊 Found {mp3Files.Length} .mp3 files in this folder");

                    if (mp3Files.Length == 0)
                    {
                        Console.WriteLine("⚠️  Warning: No .mp3 files found in this folder.");
                        Console.Write("Continue anyway? (Y/N): ");
                        string continueChoice = Console.ReadLine()?.Trim().ToUpper();

                        if (continueChoice != "Y")
                        {
                            attempts++;
                            Console.WriteLine("Please select a different folder.");
                            continue;
                        }
                    }

                    return userPath;
                }
                else
                {
                    Console.WriteLine($"❌ Invalid folder path: {userPath}");
                    Console.WriteLine("Please check the path and try again.");
                    attempts++;

                    if (attempts < maxAttempts)
                    {
                        Console.WriteLine($"💡 Tips for copying path:");
                        Console.WriteLine("   - Click on folder address bar (or press Ctrl+L)");
                        Console.WriteLine("   - The full path will be highlighted");
                        Console.WriteLine("   - Press Ctrl+C to copy, then Ctrl+V to paste here");
                    }
                }
            }

            Console.WriteLine($"\n❌ Maximum attempts ({maxAttempts}) reached.");
            Console.Write("Do you want to try a different method? (Y/N): ");
            string retry = Console.ReadLine()?.Trim().ToUpper();

            if (retry == "Y")
            {
                return SelectAudioFolder();
            }
            else
            {
                Console.WriteLine("Skipping audio folder selection.");
                return null;
            }
        }

        private static string UseDefaultFolder(string defaultPath)
        {
            // Thay đổi default path cho audio folder
            string audioDefaultPath = @"C:\Audio"; // hoặc bất kỳ path nào phù hợp

            Console.WriteLine("Note: Default audio folder is separate from Excel folder.");
            Console.WriteLine($"Excel files are in: S:\\Anki");
            Console.WriteLine($"Default audio folder: {audioDefaultPath}");

            if (Directory.Exists(audioDefaultPath))
            {
                Console.WriteLine($"✅ Using default audio folder: {audioDefaultPath}");
                return audioDefaultPath;
            }
            else
            {
                Console.WriteLine($"❌ Default audio folder doesn't exist: {audioDefaultPath}");
                Console.WriteLine("Please select another option.");
                return OpenWindowsExplorer(defaultPath);
            }
        }

        private static string BrowseAndSelectFolder()
        {
            try
            {
                // Hiển thị các drives có sẵn
                DriveInfo[] drives = DriveInfo.GetDrives()
                    .Where(d => d.IsReady && d.DriveType == DriveType.Fixed)
                    .ToArray();

                Console.WriteLine("\nAvailable drives:");
                for (int i = 0; i < drives.Length; i++)
                {
                    Console.WriteLine($"{i + 1}. {drives[i].Name} ({drives[i].VolumeLabel})");
                }

                Console.Write("Select drive (enter number): ");
                string driveChoice = Console.ReadLine()?.Trim();

                if (int.TryParse(driveChoice, out int driveIndex) &&
                    driveIndex > 0 && driveIndex <= drives.Length)
                {
                    string selectedDrive = drives[driveIndex - 1].Name;
                    return BrowseFolders(selectedDrive);
                }
                else
                {
                    Console.WriteLine("Invalid drive selection.");
                    return EnterCustomPath();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error browsing folders: {ex.Message}");
                return EnterCustomPath();
            }
        }

        private static string BrowseFolders(string currentPath)
        {
            try
            {
                while (true)
                {
                    Console.WriteLine($"\nCurrent path: {currentPath}");

                    // Lấy danh sách thư mục
                    string[] directories = Directory.GetDirectories(currentPath)
                        .Take(20) // Giới hạn 20 thư mục để không quá dài
                        .ToArray();

                    if (directories.Length == 0)
                    {
                        Console.WriteLine("No subdirectories found.");
                    }
                    else
                    {
                        Console.WriteLine("\nAvailable folders:");
                        for (int i = 0; i < directories.Length; i++)
                        {
                            Console.WriteLine($"{i + 1}. {Path.GetFileName(directories[i])}");
                        }
                    }

                    Console.WriteLine("\nOptions:");
                    Console.WriteLine("0. Select current folder");
                    Console.WriteLine("99. Go back to parent folder");
                    Console.WriteLine("88. Enter path manually");
                    Console.Write("Enter your choice: ");

                    string choice = Console.ReadLine()?.Trim();

                    if (choice == "0")
                    {
                        return currentPath;
                    }
                    else if (choice == "99")
                    {
                        DirectoryInfo parent = Directory.GetParent(currentPath);
                        if (parent != null)
                        {
                            currentPath = parent.FullName;
                        }
                        else
                        {
                            Console.WriteLine("Already at root directory.");
                        }
                    }
                    else if (choice == "88")
                    {
                        return EnterCustomPath();
                    }
                    else if (int.TryParse(choice, out int folderIndex) &&
                             folderIndex > 0 && folderIndex <= directories.Length)
                    {
                        currentPath = directories[folderIndex - 1];
                    }
                    else
                    {
                        Console.WriteLine("Invalid choice. Please try again.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error accessing folder: {ex.Message}");
                return EnterCustomPath();
            }
        }

        private static string EnterCustomPath()
        {
            while (true)
            {
                Console.Write("\nEnter the full path to your audio folder: ");
                string customPath = Console.ReadLine()?.Trim();

                if (string.IsNullOrWhiteSpace(customPath))
                {
                    Console.WriteLine("Path cannot be empty. Please try again.");
                    continue;
                }

                // Xử lý path với quotes
                customPath = customPath.Trim('"');

                if (Directory.Exists(customPath))
                {
                    Console.WriteLine($"Folder selected: {customPath}");
                    return customPath;
                }
                else
                {
                    Console.WriteLine("Folder doesn't exist. Please check the path and try again.");
                    Console.Write("Do you want to try again? (Y/N): ");
                    string retry = Console.ReadLine()?.Trim().ToUpper();

                    if (retry != "Y")
                    {
                        // Trả về null nếu user không muốn thử lại
                        Console.WriteLine("Skipping audio folder selection.");
                        return null;
                    }
                }
            }
        }

        public static void DisplayFolderInfo(string folderPath)
        {
            try
            {
                if (Directory.Exists(folderPath))
                {
                    string[] mp3Files = Directory.GetFiles(folderPath, "*.mp3");
                    Console.WriteLine($"\nFolder information:");
                    Console.WriteLine($"Path: {folderPath}");
                    Console.WriteLine($"Total .mp3 files: {mp3Files.Length}");

                    if (mp3Files.Length > 0)
                    {
                        Console.WriteLine("Sample files:");
                        for (int i = 0; i < Math.Min(3, mp3Files.Length); i++)
                        {
                            Console.WriteLine($"  - {Path.GetFileName(mp3Files[i])}");
                        }
                        if (mp3Files.Length > 3)
                        {
                            Console.WriteLine($"  ... and {mp3Files.Length - 3} more files");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading folder information: {ex.Message}");
            }
        }
    }
}
