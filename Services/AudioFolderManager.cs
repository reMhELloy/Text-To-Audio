using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Text_to_Image.Models;

namespace Text_to_Image.Services
{
    // AudioFolderManager.cs - Combined AudioFolderManager + FolderSelector
    public static class AudioFolderManager
    {
        // ========================================
        // MAIN FOLDER SELECTION METHODS
        // ========================================

        public static string SelectAudioOutputFolder(ProcessingOptions options)
        {
            string selectedFolder = SelectAudioFolder();
            if (string.IsNullOrEmpty(selectedFolder))
            {
                selectedFolder = UseDefaultAudioFolder();
            }
            return selectedFolder;
        }

        public static string SelectAudioFolder(string defaultPath = @"C:\Audio")
        {
            string selectedFolder = OpenFolderSelectionDialog(defaultPath);
            if (string.IsNullOrEmpty(selectedFolder))
            {
                return UseDefaultFolder();
            }
            return selectedFolder;
        }

        // ========================================
        // FOLDER SELECTION DIALOG METHODS
        // ========================================

        private static string OpenFolderSelectionDialog(string defaultPath)
        {
            try
            {
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
                    process.WaitForExit();
                    output = output?.Trim();

                    if (output == "CANCELLED" || string.IsNullOrWhiteSpace(output))
                    {
                        return null;
                    }

                    if (Directory.Exists(output))
                    {
                        return output;
                    }
                }
            }
            catch (Exception ex)
            {
                // Silent fail
            }

            return null;
        }

        // ========================================
        // FALLBACK FOLDER METHODS
        // ========================================

        private static string UseDefaultAudioFolder()
        {
            string defaultFolder = @"S:\Anki\Audio";
            try
            {
                if (!Directory.Exists(defaultFolder))
                {
                    Directory.CreateDirectory(defaultFolder);
                }
                return defaultFolder;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private static string UseDefaultFolder()
        {
            string audioDefaultPath = @"C:\Audio";
            try
            {
                if (!Directory.Exists(audioDefaultPath))
                {
                    Directory.CreateDirectory(audioDefaultPath);
                }
                return audioDefaultPath;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        // ========================================
        // FOLDER INFO DISPLAY METHODS
        // ========================================

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