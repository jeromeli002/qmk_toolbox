using QMK_Toolbox.Properties;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace QMK_Toolbox.Helpers
{
    public class DriverInstaller
    {
        private const string DriversListFilename = "drivers.txt";
        private const string InstallerFilename = "qmk_driver_installer.exe";

        public static bool DisplayPrompt()
        {
            var driverPromptResult = MessageBox.Show("确定要安装驱动吗?", "驱动安装", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (driverPromptResult == DialogResult.No || !InstallDrivers())
            {
                return false;
            }
            
            Settings.Default.driversInstalled = true;
            Settings.Default.Save();
            return true;
        }

        private static bool InstallDrivers()
        {
            var driversPath = Path.Combine(Application.LocalUserAppDataPath, DriversListFilename);
            var installerPath = Path.Combine(Application.LocalUserAppDataPath, InstallerFilename);

            if (!File.Exists(driversPath))
            {
                EmbeddedResourceHelper.ExtractResources(DriversListFilename);
            }

            if (!File.Exists(installerPath))
            {
                EmbeddedResourceHelper.ExtractResources(InstallerFilename);
            }

            var process = new Process
            {
                StartInfo = new ProcessStartInfo(installerPath, $"--all --force \"{driversPath}\"")
                {
                    Verb = "runas"
                }
            };

            try
            {
                process.Start();
                return true;
            }
            catch (Win32Exception)
            {
                var tryAgainResult = MessageBox.Show("此操作需要管理员权限，请用管理员权限重新打开操作！！", "错误", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                if (tryAgainResult == DialogResult.Retry)
                {
                    return InstallDrivers();
                }
            }

            return false;
        }
    }
}
