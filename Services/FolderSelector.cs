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
            Console.WriteLine("Choose how you want to select the audio folder:");
            Console.WriteLine("1. Use default audio folder");
            Console.WriteLine("2. Open folder selection dialog");
            Console.WriteLine("3. Browse folders in console");
            Console.WriteLine("4. Enter custom folder path manually");
            Console.Write("Enter your choice (1-4): ");

            string choice = Console.ReadLine()?.Trim();

            switch (choice)
            {
                case "1":
                    return UseDefaultFolder(defaultPath);

                case "2":
                    return OpenFolderSelectionDialog(defaultPath);

                case "3":
                    return BrowseAndSelectFolder();

                case "4":
                    return EnterCustomPath();

                default:
                    Console.WriteLine("Opening folder selection dialog.");
                    return OpenFolderSelectionDialog(defaultPath);
            }
        }

        private static string OpenFolderSelectionDialog(string defaultPath)
        {
            try
            {
                Console.WriteLine("Opening folder selection dialog...");

                // Create a PowerShell script to show folder browser dialog
                string script = @"
Add-Type -AssemblyName System.Windows.Forms
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = 'Select Audio Folder'
$folderBrowser.ShowNewFolderButton = $true";

                // Set initial directory if it exists
                if (Directory.Exists(defaultPath))
                {
                    script += $"\n$folderBrowser.SelectedPath = '{defaultPath}'";
                }
                else
                {
                    // Try to find a reasonable starting location
                    string[] tryPaths = {
                        @"C:\Audio",
                        @"D:\Audio",
                        Environment.GetFolderPath(Environment.SpecialFolder.MyMusic),
                        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                        @"C:\"
                    };

                    foreach (string path in tryPaths)
                    {
                        if (Directory.Exists(path))
                        {
                            script += $"\n$folderBrowser.SelectedPath = '{path}'";
                            break;
                        }
                    }
                }

                script += @"
$result = $folderBrowser.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    Write-Output $folderBrowser.SelectedPath
} else {
    Write-Output 'CANCELLED'
}";

                ProcessStartInfo psi = new ProcessStartInfo()
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -Command \"{script}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (Process process = Process.Start(psi))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();

                    output = output?.Trim();

                    if (!string.IsNullOrEmpty(error))
                    {
                        Console.WriteLine($"Error: {error}");
                        return FallbackToManualInput();
                    }

                    if (output == "CANCELLED" || string.IsNullOrWhiteSpace(output))
                    {
                        Console.WriteLine("Folder selection was cancelled.");
                        return null;
                    }

                    if (Directory.Exists(output))
                    {
                        Console.WriteLine($"Selected folder: {output}");
                        return output;
                    }
                    else
                    {
                        Console.WriteLine("Selected path is not valid.");
                        return FallbackToManualInput();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error opening folder dialog: {ex.Message}");
                return FallbackToManualInput();
            }
        }

        private static string FallbackToManualInput()
        {
            Console.WriteLine("Falling back to manual input.");
            return EnterCustomPath();
        }

        private static string UseDefaultFolder(string defaultPath)
        {
            string audioDefaultPath = @"C:\Audio";

            if (Directory.Exists(audioDefaultPath))
            {
                Console.WriteLine($"Using default audio folder: {audioDefaultPath}");
                return audioDefaultPath;
            }
            else
            {
                Console.WriteLine($"Default audio folder doesn't exist: {audioDefaultPath}");
                return OpenFolderSelectionDialog(defaultPath);
            }
        }

        private static string BrowseAndSelectFolder()
        {
            try
            {
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

                    string[] directories = Directory.GetDirectories(currentPath)
                        .Take(20)
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
