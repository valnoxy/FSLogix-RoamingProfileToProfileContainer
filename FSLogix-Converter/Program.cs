using System.Diagnostics;
using System.Management;
using System.Text.RegularExpressions;

namespace FSLogix_Converter
{
    internal class Program
    {
        public class User
        {
            public string? Username { get; set; }
            public string? Sid { get; set; }
            public string? ProfilePath { get; set; }
        }
        public const bool HideBuiltInAccounts = true;
        public const bool HideUnknownSiDs = true;
        public static bool CreateDynamicDisk = false;
        public static bool IsUNC = false;
        public static string? TempPath;
        public static string? FullUserName;
        public static string? TargetDir;
        public static string? DiskSize;
        public static User? TargetUser;

        private static void Main(string[] args)
        {
            switch (args.Length)
            {
                default:
                    ShowHelp();
                    break;
                case 3:
                case 4:
                    FullUserName = args[0];
                    TargetDir = args[1];
                    DiskSize = args[2];
                    foreach (var arg in args)
                    {
                        if (arg.Equals("-dynamic", StringComparison.CurrentCultureIgnoreCase))
                            CreateDynamicDisk = true;
                    }
                    RunMigration();
                    break;
            }
        }

        private static void RunMigration()
        {
            DebugConsole.WriteLine("Fetching users ...");
            var data = GetUsersFromHost(Environment.MachineName);
            foreach (var user in data!.OfType<User>())
            {
                DebugConsole.WriteLine($"Username: {user.Username}");
                DebugConsole.WriteLine($"SID: {user.Sid}");
                DebugConsole.WriteLine($"ProfilePath: {user.ProfilePath}");
                DebugConsole.WriteLine($"----------------------------------------");

                if (string.Equals(user.Username!, FullUserName, StringComparison.CurrentCultureIgnoreCase))
                    TargetUser = user;
            }

            if (TargetUser == null)
            {
                DebugConsole.WriteLine($"User {FullUserName} not found!");
                return;
            }

            DebugConsole.WriteLine($"Found target user on system. Proceed with SID \"{TargetUser.Sid}\"... ");

            // Check if target is UNC path
            if (IsUncPath(TargetDir))
            {
                IsUNC = true;
                TempPath = Path.Combine(Path.GetTempPath(), "FSLogix-Converter", $"Profile_{TargetUser.Sid}.vhdx");
            }

            var regText = $"Windows Registry Editor Version 5.00\r\n[HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList\\{TargetUser.Sid}]\r\n\"ProfileImagePath\"=\"{TargetUser.ProfilePath}\"\r\n\"FSL_OriginalProfileImagePath\"=\"{TargetUser.ProfilePath}\"\r\n\"Flags\"=dword:00000000\r\n\"State\"=dword:00000000\r\n\"ProfileLoadTimeLow\"=dword:00000000\r\n\"ProfileLoadTimeHigh\"=dword:00000000\r\n\"RefCount\"=dword:00000000\r\n\"RunLogonScriptSync\"=dword:00000000\r\n";
            DebugConsole.WriteLine("Created registry file");

            var userSplit = FullUserName!.Split("\\");
            var userName = userSplit[1];
            var targetUserPath = Path.Combine(TargetDir!, $"{TargetUser.Sid}_{userName}", $"Profile_{TargetUser.Sid}.vhdx");
            DebugConsole.WriteLine("New Profile path will be: " + targetUserPath);
            Directory.CreateDirectory(Path.GetDirectoryName(targetUserPath)!);
            if (IsUNC) Directory.CreateDirectory(Path.GetDirectoryName(TempPath)!);
            if (File.Exists(targetUserPath))
            {
                DebugConsole.WriteLine("Failed to create virtual disk: File already exists!", ConsoleColor.Red);
                return;
            }
            var deploymentLetter = GetFreeLetter();
            if (Directory.Exists($"{deploymentLetter}:\\"))
            {
                DebugConsole.WriteLine("Error: No free drive letter available on this system.", ConsoleColor.Red);
                return;
            }

            // Create disk
            DebugConsole.WriteLine("Building virtual disk ...");
            var partDest = new Process();
            partDest.StartInfo.FileName = "diskpart.exe";
            partDest.StartInfo.UseShellExecute = false;
            partDest.StartInfo.CreateNoWindow = true;
            partDest.StartInfo.RedirectStandardInput = true;
            partDest.StartInfo.RedirectStandardOutput = true;
            partDest.Start();

            if (IsUNC)
            {
                partDest.StandardInput.WriteLine(CreateDynamicDisk
                    ? $"create vdisk file=\"{TempPath}\" maximum={DiskSize} type=expandable"
                    : $"create vdisk file=\"{TempPath}\" maximum={DiskSize} type=fixed");
                partDest.StandardInput.WriteLine($"select vdisk file=\"{TempPath}\"");
            }
            else
            {
                partDest.StandardInput.WriteLine(CreateDynamicDisk
                    ? $"create vdisk file=\"{targetUserPath}\" maximum={DiskSize} type=expandable"
                    : $"create vdisk file=\"{targetUserPath}\" maximum={DiskSize} type=fixed");
                partDest.StandardInput.WriteLine($"select vdisk file=\"{targetUserPath}\"");
            }
            partDest.StandardInput.WriteLine("attach vdisk");
            partDest.StandardInput.WriteLine("create partition primary");
            partDest.StandardInput.WriteLine($"format fs=ntfs quick label=\"Profile-{userName}\"");
            partDest.StandardInput.WriteLine($"assign letter={deploymentLetter}");
            partDest.StandardInput.WriteLine("exit");
            partDest.WaitForExit();
            DebugConsole.WriteLine("Disk successfully created.");

            // Assign permissions
            DebugConsole.WriteLine("Assigning access permission ...");
            Directory.CreateDirectory($"{deploymentLetter}:\\Profile");

            var status = StartProcess("icacls", $@"{deploymentLetter}:\Profile /inheritance:r");
            if (status != 0)
            {
                DebugConsole.WriteLine("Failed to run icacls: Exited with code " + status);
                return;
            }

            status = StartProcess("icacls", $@"{deploymentLetter}:\Profile /grant SYSTEM:(OI)(CI)F"); // System
            if (status != 0)
            {
                DebugConsole.WriteLine("Failed to run icacls: Exited with code " + status);
                return;
            }

            status = StartProcess("icacls", $@"{deploymentLetter}:\Profile /grant *S-1-5-32-544:(OI)(CI)F"); // Administrators
            if (status != 0)
            {
                DebugConsole.WriteLine("Failed to run icacls: Exited with code " + status);
                return;
            }

            status = StartProcess("icacls", $@"{deploymentLetter}:\Profile /grant {FullUserName}:(OI)(CI)F"); // User itself
            if (status != 0)
            {
                DebugConsole.WriteLine("Failed to run icacls: Exited with code " + status);
                return;
            }

            // Migration - Cloning to vhd
            DebugConsole.WriteLine("Cloning to disk ...");
            var sourceInfo = new DirectoryInfo(TargetUser.ProfilePath!);
            var targetInfo = new DirectoryInfo($"{deploymentLetter}:\\Profile");
            var totalItems = GetTotalItemCount(sourceInfo);
            var copiedItems = 0;
            CopyAll(sourceInfo, targetInfo, ref copiedItems, totalItems);
            Console.WriteLine();

            // set registry file
            Directory.CreateDirectory($"{deploymentLetter}:\\Profile\\AppData\\Local\\FSLogix");
            if (!File.Exists($"{deploymentLetter}:\\Profile\\AppData\\Local\\FSLogix\\ProfileData.reg"))
            {
                File.WriteAllText($"{deploymentLetter}:\\Profile\\AppData\\Local\\FSLogix\\ProfileData.reg", regText);
                DebugConsole.WriteLine("Registry file written to disk");
            }

            // Set ownership
            status = StartProcess("icacls", $@"{deploymentLetter}:\Profile /setowner SYSTEM");
            if (status != 0)
            {
                DebugConsole.WriteLine("Failed to run icacls: Exited with code " + status);
                return;
            }

            // Detach vhd file
            DebugConsole.WriteLine("Detaching virtual disk ...");
            partDest.Start();
            partDest.StandardInput.WriteLine(IsUNC
                ? $"select vdisk file=\"{TempPath}\""
                : $"select vdisk file=\"{targetUserPath}\"");
            partDest.StandardInput.WriteLine("detach vdisk");
            partDest.StandardInput.WriteLine("exit");

            if (IsUNC)
            {
                DebugConsole.WriteLine("Copying profile to share ...");
                try
                {
                    var userPath = Path.Combine(TargetDir!, $"{TargetUser.Sid}_{userName}");
                    Directory.CreateDirectory(userPath);
                    File.Move(TempPath, Path.Combine(userPath, $"Profile_{TargetUser.Sid}.vhdx"));
                }
                catch (Exception ex)
                {
                    DebugConsole.WriteLine("Failed to move profile to share: " + ex.Message, ConsoleColor.Red);
                }
            }

            DebugConsole.WriteLine("Profile successfully migrated.");
        }

        private static void ShowHelp()
        {
            Console.WriteLine("FSLogix Converter [Version 1.0]");
            Console.WriteLine("\nSyntax: FSLogix-Converter.exe [Domain\\Username] [Path\\To\\Share] [Disk Size in MB] (-dynamic)\n");
            Console.WriteLine("   [Domain\\Username]   -   Define the user you want to migrate");
            Console.WriteLine("   [Path\\To\\Share]     -   File path to the destination of the virtual disk");
            Console.WriteLine("   [Disk Size in MB]   -   Size of the virtual disk in MB");
            Console.WriteLine("   -dynamic            -   Create a dynamic virtual disk");
            Console.WriteLine("\nExample: FSLogix-Converter.exe Contoso\\John.Doe \\\\DC01\\FSLogix 30720 -dynamic\n");
            Environment.Exit(1);
        }

        #region Helpers
        public static bool IsUncPath(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }

            if (Uri.TryCreate(path, UriKind.Absolute, out var uri))
            {
                return uri.IsUnc;
            }

            // Additional manual check for UNC paths that might not be caught by Uri.TryCreate
            return path.StartsWith(@"\\") && path.Trim('\\').Contains(@"\");
        }

        public static int GetTotalItemCount(DirectoryInfo directory)
        {
            var count = 0;
            try
            {
                count += directory.GetFiles().Length;
            }
            catch (Exception ex)
            {
                //DebugConsole.WriteLine($"Error counting files in directory {directory.FullName}: {ex.Message}", ConsoleColor.Yellow);
            }

            foreach (var subDir in directory.GetDirectories())
            {
                try
                {
                    count += GetTotalItemCount(subDir);
                }
                catch (Exception ex)
                {
                    //DebugConsole.WriteLine($"Error counting items in subdirectory {subDir.FullName}: {ex.Message}", ConsoleColor.Yellow);
                }
            }

            return count;
        }

        public static void CopyAll(DirectoryInfo source, DirectoryInfo target, ref int copiedItems, int totalItems)
        {
            if (!source.Exists)
            {
                throw new DirectoryNotFoundException($"Source directory does not exist or could not be found: {source.FullName}");
            }
            if (!target.Exists)
            {
                target.Create();
            }
            foreach (var file in source.GetFiles())
            {
                try
                {
                    var targetFilePath = Path.Combine(target.FullName, file.Name);
                    file.CopyTo(targetFilePath, true);
                    copiedItems++;
                    DisplayProgress(copiedItems, totalItems);
                }
                catch (Exception ex)
                {
                    //DebugConsole.WriteLine($"Error copying file {file.FullName}: {ex.Message}", ConsoleColor.Yellow);
                }
            }

            // Recursively copy all subdirectories
            foreach (var sourceSubDir in source.GetDirectories())
            {
                try
                {
                    var targetSubDir = target.CreateSubdirectory(sourceSubDir.Name);
                    CopyAll(sourceSubDir, targetSubDir, ref copiedItems, totalItems);
                }
                catch (Exception ex)
                {
                    //DebugConsole.WriteLine($"Error copying directory {sourceSubDir.FullName}: {ex.Message}", ConsoleColor.Yellow);
                }
            }

            // Purge files and directories in the target directory that are not in the source directory
            Purge(target, source);
        }

        public static void Purge(DirectoryInfo target, DirectoryInfo source)
        {
            // Remove files in the target directory that do not exist in the source directory
            foreach (var targetFile in target.GetFiles())
            {
                var sourceFilePath = Path.Combine(source.FullName, targetFile.Name);
                if (!File.Exists(sourceFilePath))
                {
                    try
                    {
                        targetFile.Delete();
                    }
                    catch (Exception ex)
                    {
                        //DebugConsole.WriteLine($"Error deleting file {targetFile.FullName}: {ex.Message}", ConsoleColor.Yellow);
                    }
                }
            }

            // Recursively remove directories in the target directory that do not exist in the source directory
            foreach (var targetSubDir in target.GetDirectories())
            {
                var correspondingSourceSubDir = new DirectoryInfo(Path.Combine(source.FullName, targetSubDir.Name));
                if (!correspondingSourceSubDir.Exists)
                {
                    try
                    {
                        targetSubDir.Delete(true);
                    }
                    catch (Exception ex)
                    {
                        DebugConsole.WriteLine($"Error deleting directory {targetSubDir.FullName}: {ex.Message}", ConsoleColor.Yellow);
                    }
                }
                else
                {
                    Purge(targetSubDir, correspondingSourceSubDir);
                }
            }
        }

        public static void DisplayProgress(int copiedItems, int totalItems)
        {
            var progressBarWidth = 50;
            var progress = (int)((copiedItems / (float)totalItems) * progressBarWidth);

            Console.CursorLeft = 0;
            Console.Write("[");
            Console.Write(new string('#', progress));
            Console.Write(new string(' ', progressBarWidth - progress));
            Console.Write($"] {copiedItems}/{totalItems} files copied");
        }

        private static string? GetFreeLetter()
        {
            var availableDriveLetters = new List<char>() { 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z' };
            var drives = DriveInfo.GetDrives();
            foreach (var t in drives)
            {
                availableDriveLetters.Remove(t.Name.ToLower()[0]);
            }
            var freeDisks = availableDriveLetters.ToArray();
            return $"{freeDisks[0]}";
        }

        private static int StartProcess(string fileName, string arguments)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using var process = Process.Start(startInfo);
            process!.WaitForExit();
            return process.ExitCode;
        }

        /// <summary>
        /// Attempts to resolve an identity (Store ID / Security ID) to a DOMAIN\USERNAME string. If failed, returns the given identity string.
        /// </summary>
        /// <param name="id">A Store ID or SID (Security Identifier) to resolve.</param>
        /// <param name="host">If this parameter is specified, then this method will treat the ID parameter as an SID and resolve it on the remote host.</param>
        /// <returns>DOMAIN\USERNAME.</returns>
        private static string? GetUserByIdentity(string id, string? host = null)
        {
            if (host != null)
            {
                try
                {
                    var mo = WMI.GetInstance(host, "Win32_SID.SID=\"" + id + "\"");
                    var ntDomain = mo.GetPropertyValue("ReferencedDomainName").ToString();
                    var ntAccountName = mo.GetPropertyValue("AccountName").ToString();
                    var ntAccount = ntDomain + "\\" + ntAccountName;
                    if (ntDomain == "" || ntAccountName == "") throw new Exception();
                    return ntAccount;
                }
                catch (Exception)
                {
                    return id;
                }
            }

            return null;
        }
        
        private static User GetUserFromHost(string host, ManagementBaseObject userObject)
        {
            var sid = userObject.GetPropertyValue("SID").ToString();
            var username = GetUserByIdentity(sid!, host);

            if (HideBuiltInAccounts && (Regex.IsMatch(sid!, @"^S-1-5-[0-9]+$")))
            {
                return null!;
            }
            if (HideUnknownSiDs && sid == username)
            {
                return null!;
            }

            var profilePath = userObject.GetPropertyValue("LocalPath").ToString();
            return new User
            {
                Sid = sid!,
                ProfilePath = profilePath!,
                Username = username
            };
        }


        private static List<User>? GetUsersFromHost(string host)
        {
            var users = new List<User>();
            var manObjCol = WMI.Query("SELECT SID, LocalPath FROM Win32_UserProfile", host);
            foreach (var man in manObjCol)
            {
                var user = GetUserFromHost(host, man);
                users.Add(user);
            }
            return users;
        }
        #endregion
    }
}