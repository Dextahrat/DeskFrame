using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using IWshRuntimeLibrary;
using File = System.IO.File;

namespace DeskFrame.Util
{
    /// <summary>
    /// Dosya ve klasör i?lemleri için yard?mc? s?n?f.
    /// Drag & Drop kopyalama/ta??ma i?lemlerini yönetir.
    /// </summary>
    public class FileOperationsHelper
    {
        /// <summary>
        /// Bir dosya veya klasörü hedef konuma kopyalar.
        /// </summary>
        /// <param name="sourcePath">Kaynak dosya veya klasör yolu</param>
        /// <param name="destinationPath">Hedef konum</param>
        /// <param name="overwrite">Varolan dosyalar?n üzerine yaz?ls?n m??</param>
        public static void CopyItem(string sourcePath, string destinationPath, bool overwrite = true)
        {
            if (Directory.Exists(sourcePath))
            {
                CopyDirectory(sourcePath, destinationPath, true, overwrite);
            }
            else if (File.Exists(sourcePath))
            {
                File.Copy(sourcePath, destinationPath, overwrite);
            }
            else
            {
                throw new FileNotFoundException($"Source not found: {sourcePath}");
            }
        }

        /// <summary>
        /// Bir dosya veya klasörü hedef konuma ta??r.
        /// </summary>
        /// <param name="sourcePath">Kaynak dosya veya klasör yolu</param>
        /// <param name="destinationPath">Hedef konum</param>
        public static void MoveItem(string sourcePath, string destinationPath)
        {
            if (Directory.Exists(sourcePath))
            {
                Directory.Move(sourcePath, destinationPath);
            }
            else if (File.Exists(sourcePath))
            {
                File.Move(sourcePath, destinationPath);
            }
            else
            {
                throw new FileNotFoundException($"Source not found: {sourcePath}");
            }
        }

        /// <summary>
        /// Klasörü tüm alt klasörleriyle birlikte kopyalar.
        /// </summary>
        /// <param name="sourceDirName">Kaynak klasör yolu</param>
        /// <param name="destDirName">Hedef klasör yolu</param>
        /// <param name="copySubDirs">Alt klasörler de kopyalans?n m??</param>
        /// <param name="overwrite">Varolan dosyalar?n üzerine yaz?ls?n m??</param>
        public static void CopyDirectory(string sourceDirName, string destDirName, bool copySubDirs, bool overwrite = true)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: " + sourceDirName);
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            
            // Hedef klasörü olu?tur
            Directory.CreateDirectory(destDirName);

            // Dosyalar? kopyala
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string tempPath = Path.Combine(destDirName, file.Name);
                file.CopyTo(tempPath, overwrite);
            }

            // Alt klasörleri kopyala
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string tempPath = Path.Combine(destDirName, subdir.Name);
                    CopyDirectory(subdir.FullName, tempPath, copySubDirs, overwrite);
                }
            }
        }

        /// <summary>
        /// Bir dosya veya klasör için k?sayol (.lnk) olu?turur.
        /// .url dosyalar? için do?rudan kopyalama yapar.
        /// </summary>
        /// <param name="targetPath">K?sayolun hedef alaca?? dosya/klasör yolu</param>
        /// <param name="shortcutFolder">K?sayolun olu?turulaca?? klasör (null ise hedefin bulundu?u klasör)</param>
        /// <param name="overwrite">Varolan k?sayolun üzerine yaz?ls?n m??</param>
        /// <returns>Olu?turulan k?sayolun tam yolu</returns>
        public static string CreateShortcut(string targetPath, string shortcutFolder = null, bool overwrite = true)
        {
            // .url dosyalar? için özel i?lem
            if (Path.GetExtension(targetPath).Equals(".url", StringComparison.OrdinalIgnoreCase))
            {
                string folder = !string.IsNullOrEmpty(shortcutFolder) 
                    ? shortcutFolder 
                    : Path.GetDirectoryName(targetPath);
                    
                string destinationPath = Path.Combine(folder, Path.GetFileName(targetPath));
                File.Copy(targetPath, destinationPath, overwrite);
                return destinationPath;
            }

            // .lnk k?sayolu olu?tur
            string shortcutDirectory = !string.IsNullOrEmpty(shortcutFolder) 
                ? shortcutFolder 
                : Path.GetDirectoryName(targetPath);
                
            string shortcutPath = Path.Combine(shortcutDirectory, 
                Path.GetFileNameWithoutExtension(targetPath) + ".lnk");

            // Varolan k?sayolu sil
            if (overwrite && File.Exists(shortcutPath))
            {
                File.Delete(shortcutPath);
            }

            WshShell shell = new WshShell();
            IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);

            shortcut.TargetPath = targetPath;
            shortcut.WorkingDirectory = Path.GetDirectoryName(targetPath);
            
            try
            {
                shortcut.Description = Path.GetFileName(targetPath);
            }
            catch
            {
                // Aç?klama eklenemezse sessizce devam et
            }
            
            shortcut.Save();
            
            return shortcutPath;
        }

        /// <summary>
        /// Drag & Drop i?leminde dosyalar? i?ler.
        /// Copy veya Move operasyonunu belirler ve i?lemi gerçekle?tirir.
        /// </summary>
        /// <param name="droppedFiles">B?rak?lan dosya yollar?</param>
        /// <param name="targetFolder">Hedef klasör</param>
        /// <param name="isCopyOperation">Kopyalama m? yoksa ta??ma m??</param>
        /// <param name="dropIntoSubfolder">Alt klasöre mi b?rak?ld?? (null ise ana hedef)</param>
        /// <param name="createShortcutsInstead">K?sayol olu?turulsun mu?</param>
        /// <returns>??lenen dosya say?s?</returns>
        public static int ProcessDroppedFiles(
            string[] droppedFiles, 
            string targetFolder, 
            bool isCopyOperation,
            string dropIntoSubfolder = null,
            bool createShortcutsInstead = false)
        {
            int processedCount = 0;

            foreach (var sourcePath in droppedFiles)
            {
                try
                {
                    // Hedef yolu belirle
                    string destinationFolder = !string.IsNullOrEmpty(dropIntoSubfolder)
                        ? dropIntoSubfolder
                        : targetFolder;

                    string fileName = Path.GetFileName(sourcePath);
                    string destinationPath = Path.Combine(destinationFolder, fileName);

                    // Ayn? konuma i?lem yap?l?yorsa atla
                    if (Path.GetDirectoryName(sourcePath) == destinationFolder && !createShortcutsInstead)
                    {
                        Debug.WriteLine($"Skipping same location: {sourcePath}");
                        continue;
                    }

                    // ??lem tipine göre i?le
                    if (createShortcutsInstead)
                    {
                        CreateShortcut(sourcePath, destinationFolder);
                        Debug.WriteLine($"Created shortcut: {sourcePath} -> {destinationPath}");
                    }
                    else if (isCopyOperation)
                    {
                        CopyItem(sourcePath, destinationPath);
                        Debug.WriteLine($"Copied: {sourcePath} -> {destinationPath}");
                    }
                    else
                    {
                        MoveItem(sourcePath, destinationPath);
                        Debug.WriteLine($"Moved: {sourcePath} -> {destinationPath}");
                    }

                    processedCount++;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error processing {sourcePath}: {ex.Message}");
                    // Hata durumunda da devam et
                }
            }

            return processedCount;
        }

        /// <summary>
        /// Bir yolun geçerli olup olmad???n? kontrol eder.
        /// </summary>
        public static bool IsValidPath(string path)
        {
            return !string.IsNullOrEmpty(path) && 
                   (File.Exists(path) || Directory.Exists(path));
        }

        /// <summary>
        /// ?ki yolun ayn? olup olmad???n? kontrol eder (büyük/küçük harf duyars?z).
        /// </summary>
        public static bool IsSamePath(string path1, string path2)
        {
            if (string.IsNullOrEmpty(path1) || string.IsNullOrEmpty(path2))
                return false;

            return Path.GetFullPath(path1).Equals(
                Path.GetFullPath(path2), 
                StringComparison.OrdinalIgnoreCase);
        }
    }
}
