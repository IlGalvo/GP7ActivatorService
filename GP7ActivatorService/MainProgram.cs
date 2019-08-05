using Microsoft.Win32;
using System;
using System.IO;
using System.Windows.Forms;

namespace GP7ActivatorService
{
    public class MainProgram
    {
        #region GLOBAL_VARIABLES
        private static readonly int ArgumentNumber = 1;
        private static readonly string ArgumentValue = (@"-uninstall");

        private static readonly string GP7DefaultName = (@"Guitar Pro 7");
        private static readonly string UninstallRegistryKeyPath = (@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall");

        private static readonly string StartupRegistryKeyPath = (@"SOFTWARE\Microsoft\\Windows\CurrentVersion\Run");
        private static readonly string StartInTrayCommand = (AppDomain.CurrentDomain.BaseDirectory + AppDomain.CurrentDomain.FriendlyName);

        private static readonly string ReminderDateFilePath = (Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            Path.Combine(Application.ProductName, "PostponeDate.txt")));
        private static readonly int MinimumPostponeDays = 25;

        private static readonly string GP7RegistryKeyPath = (@"Software\Arobas Music");
        private static readonly string GP7RegistryName = (@"guitarpro7");
        private static readonly string GP7AppDataDirectory = (@"C:\Users\Andre\AppData\Roaming\Arobas Music");
        #endregion

        #region MAIN
        public static void Main(string[] args)
        {
            bool isInstalled = IsGP7Installed();

            if ((args.Length == ArgumentNumber) && (args[0] == ArgumentValue))
            {
                isInstalled = false;
            }

            try
            {
                bool isEnabled = EnableAutomaticStartup(isInstalled);

                if (ManagePostponeDateFile(isEnabled))
                {
                    ResetGP7TrialPeriodLicenze();
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString(), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region CHECK
        private static bool IsGP7Installed()
        {
            return ((IsGP7InstalledRegistry(RegistryView.Registry32)) || (IsGP7InstalledRegistry(RegistryView.Registry64)));
        }

        private static bool IsGP7InstalledRegistry(RegistryView registryView)
        {
            bool isInstalled = false;

            using (RegistryKey baseRegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, registryView))
            using (RegistryKey subRegistryKey = baseRegistryKey.OpenSubKey(UninstallRegistryKeyPath))
            {
                foreach (string subKeyName in subRegistryKey.GetSubKeyNames())
                {
                    using (RegistryKey subKey = subRegistryKey.OpenSubKey(subKeyName))
                    {
                        if (IsProgramVisible(subKey))
                        {
                            if (subKey.GetValue("DisplayName").ToString() == GP7DefaultName)
                            {
                                isInstalled = true;
                                break;
                            }
                        }
                    }
                }
            }

            return isInstalled;
        }

        private static bool IsProgramVisible(RegistryKey subKey)
        {
            string displayName = ((string)subKey.GetValue("DisplayName"));
            string releaseType = ((string)subKey.GetValue("ReleaseType"));
            int? systemComponent = ((int?)subKey.GetValue("SystemComponent"));
            string parentDisplayName = ((string)subKey.GetValue("ParentDisplayName"));

            return ((!string.IsNullOrEmpty(displayName)) &&
                (string.IsNullOrEmpty(releaseType)) &&
                (string.IsNullOrEmpty(parentDisplayName)) &&
                (systemComponent == null));
        }
        #endregion

        #region STARTUP
        private static bool EnableAutomaticStartup(bool isInstalled)
        {
            bool isEnabled = true;

            using (RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(StartupRegistryKeyPath, true))
            {
                string startupPath = ((string)registryKey.GetValue(Application.ProductName));
                bool isInTray = ((!string.IsNullOrEmpty(startupPath)) && (startupPath == StartInTrayCommand));

                if (isInstalled && !isInTray)
                {
                    registryKey.SetValue(Application.ProductName, StartInTrayCommand);
                }
                else if (!isInstalled && isInTray)
                {
                    registryKey.DeleteValue(Application.ProductName);

                    isEnabled = false;
                }
            }

            return isEnabled;
        }
        #endregion

        #region FILE
        private static bool ManagePostponeDateFile(bool isEnabled)
        {
            bool isManaged = true;

            if (isEnabled)
            {
                string directoryNamePath = Path.GetDirectoryName(ReminderDateFilePath);

                if (!Directory.Exists(directoryNamePath))
                {
                    Directory.CreateDirectory(directoryNamePath);
                }

                if (!File.Exists(ReminderDateFilePath))
                {
                    File.Create(ReminderDateFilePath).Dispose();
                }

                if (DateTime.TryParse(File.ReadAllText(ReminderDateFilePath), out DateTime dateTime))
                {
                    isManaged = ((DateTime.Now.Day - dateTime.Day) >= MinimumPostponeDays);
                }

                if (isManaged)
                {
                    File.WriteAllText(ReminderDateFilePath, DateTime.Now.ToString("dd/MM/yyyy"));
                }
            }
            else
            {
                File.Delete(ReminderDateFilePath);

                isManaged = false;
            }

            return isManaged;
        }
        #endregion

        #region RESET
        private static void ResetGP7TrialPeriodLicenze()
        {
            DeleteGP7LicenceRegistry(RegistryView.Registry32);
            DeleteGP7LicenceRegistry(RegistryView.Registry64);

            if (Directory.Exists(GP7AppDataDirectory))
            {
                Directory.Delete(GP7AppDataDirectory, true);
            }
        }

        private static void DeleteGP7LicenceRegistry(RegistryView registryView)
        {
            using (RegistryKey baseRegistryKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, registryView))
            using (RegistryKey subRegistryKeyRoot = baseRegistryKey.OpenSubKey(GP7RegistryKeyPath, true))
            {
                if (subRegistryKeyRoot != null)
                {
                    using (RegistryKey subRegistryKey = subRegistryKeyRoot.OpenSubKey(GP7RegistryName))
                    {
                        if (subRegistryKey != null)
                        {
                            subRegistryKeyRoot.DeleteSubKeyTree(GP7RegistryName);
                        }
                    }
                }
            }
        }
        #endregion
    }
}