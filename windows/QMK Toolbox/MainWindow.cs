using QMK_Toolbox.Helpers;
using QMK_Toolbox.Hid;
using QMK_Toolbox.KeyTester;
using QMK_Toolbox.Properties;
using QMK_Toolbox.Usb;
using QMK_Toolbox.Usb.Bootloader;
using Syroot.Windows.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Windows.Forms;

namespace QMK_Toolbox
{
    public partial class MainWindow : Form
    {
        private readonly WindowState windowState = new();

        private readonly string _filePassedIn = string.Empty;

        #region Window Events
        public MainWindow()
        {
            InitializeComponent();
        }

        public MainWindow(string path) : this()
        {
            if (path != string.Empty)
            {
                var extension = Path.GetExtension(path)?.ToLower();
                if (extension == ".hex" || extension == ".bin")
                {
                    _filePassedIn = path;
                }
                else
                {
                    MessageBox.Show("不支持此的文件类型", "文件类型错误", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
            windowStateBindingSource.DataSource = windowState;
            windowState.PropertyChanged += AutoFlashEnabledChanged;
            windowState.PropertyChanged += ShowAllDevicesEnabledChanged;
            windowState.ShowAllDevices = Settings.Default.showAllDevices;

            if (Settings.Default.hexFileCollection != null)
            {
                filepathBox.Items.AddRange(Settings.Default.hexFileCollection.ToArray());
            }

            mcuBox.SelectedValue = Settings.Default.targetSetting;

            EmbeddedResourceHelper.ExtractResources(EmbeddedResourceHelper.Resources);

            logTextBox.LogRed($" - JLKB-工具箱 {Application.ProductVersion} (http://jlkb.jlkb.top)");
            logTextBox.LogRed($"*** 注意：刷入固件会清空键盘配置，请提前备份存档");
            logTextBox.LogError("*** 键盘刷入更新固件工具:");
            logTextBox.LogHid("*** 没事不要瞎刷，刷错固件可能会导致键盘变砖。");
            logTextBox.LogCommand("  一、固件写入教程");
            logTextBox.LogError("*** 1、（首次使用需先安装驱动，右键选择管理员运行）工具→安装驱动，安装完成后重启软件。");
            logTextBox.LogError("*** 2、键盘设置复位键（vial-JL里面设置，可到http://jlkb.jlkb.top下载）或按背面复位按钮进入DFU模式");
            logTextBox.LogError("*** 3、地址栏选择对应的固件（文件类型不要错了）");
            logTextBox.LogError("*** 4、键盘进入DFU模式后，如果勾选了自动写入，会自动进行写入操作，如果没有勾选，则需要手动点写入");
            logTextBox.LogError("*** 5、重新插拔或按背面复位键重启键盘");
            logTextBox.LogError("*** 6、更新完成开始使用新固件吧…………");
            logTextBox.LogBootloader("*** ***注意：刷入固件会清空键盘配置，请提前备份存档***");
            logTextBox.LogError("*** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** ");
            
            usbListener.usbDeviceConnected += UsbDeviceConnected;
            usbListener.usbDeviceDisconnected += UsbDeviceDisconnected;
            usbListener.bootloaderDeviceConnected += BootloaderDeviceConnected;
            usbListener.bootloaderDeviceDisconnected += BootloaderDeviceDisconnected;
            usbListener.outputReceived += BootloaderCommandOutputReceived;
            usbListener.Start();

            if (_filePassedIn != string.Empty)
            {
                SetFilePath(_filePassedIn);
            }

            EnableUI();
        }

        private void MainWindow_Shown(object sender, EventArgs e)
        {
            if (Settings.Default.firstStart)
            {
                Settings.Default.Upgrade();
            }

            if (Settings.Default.firstStart)
            {
                DriverInstaller.DisplayPrompt();
                Settings.Default.firstStart = false;
                Settings.Default.Save();
            }
        }

        private void MainWindow_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length == 1)
                {
                    var extension = Path.GetExtension(files.First())?.ToLower();
                    if (extension == ".hex" || extension == ".bin")
                    {
                        e.Effect = DragDropEffects.Copy;
                    }
                }
            }
        }

        private void MainWindow_DragDrop(object sender, DragEventArgs e)
        {
            SetFilePath(((string[])e.Data.GetData(DataFormats.FileDrop, false)).First());
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == NativeMethods.WmShowme)
            {
                ShowMe();
                if (File.Exists(Path.Combine(Path.GetTempPath(), "qmk_toolbox_file.txt")))
                {
                    using (var sr = new StreamReader(Path.Combine(Path.GetTempPath(), "qmk_toolbox_file.txt")))
                    {
                        SetFilePath(sr.ReadLine());
                    }
                    File.Delete(Path.Combine(Path.GetTempPath(), "qmk_toolbox_file.txt"));
                }
            }

            base.WndProc(ref m);
        }

        private void ShowMe()
        {
            if (WindowState == FormWindowState.Minimized)
            {
                WindowState = FormWindowState.Normal;
            }
            Activate();
        }

        private void MainWindow_FormClosing(object sender, FormClosingEventArgs e)
        {
            var arraylist = new ArrayList(filepathBox.Items);
            Settings.Default.hexFileCollection = arraylist;
            Settings.Default.targetSetting = (string)mcuBox.SelectedValue;
            Settings.Default.Save();

            usbListener.Dispose();
        }

        private void MainWindow_FormClosed(object sender, FormClosedEventArgs e)
        {
            Settings.Default.Save();
        }
        #endregion Window Events

        #region USB Devices & Bootloaders
        private readonly UsbListener usbListener = new();

        private void BootloaderDeviceConnected(BootloaderDevice device)
        {
            Invoke(new Action(() =>
            {
                logTextBox.LogBootloader($"{device.Name} 键盘已连接 ({device.Driver}): {device}");

                if (windowState.AutoFlashEnabled)
                {
                    FlashAllAsync();
                }
                else
                {
                    EnableUI();
                }
            }));
        }

        private void BootloaderDeviceDisconnected(BootloaderDevice device)
        {
            Invoke(new Action(() =>
            {
                logTextBox.LogBootloader($"{device.Name} 键盘断开连接 ({device.Driver}): {device}");

                if (!windowState.AutoFlashEnabled)
                {
                    EnableUI();
                }
            }));
        }

        private void BootloaderCommandOutputReceived(BootloaderDevice device, string data, MessageType type)
        {
            Invoke(new Action(() =>
            {
                logTextBox.Log(data, type);
            }));
        }

        private void UsbDeviceConnected(UsbDevice device)
        {
            Invoke(new Action(() =>
            {
                if (windowState.ShowAllDevices)
                {
                    logTextBox.LogUsb($"USB 键盘已连接 ({device.Driver}): {device}");
                }
            }));
        }

        private void UsbDeviceDisconnected(UsbDevice device)
        {
            Invoke(new Action(() =>
            {
                if (windowState.ShowAllDevices)
                {
                    logTextBox.LogUsb($"键盘已断开连接 ({device.Driver}): {device}");
                }
            }));
        }
        #endregion

        #region UI Interaction
        private void AutoFlashEnabledChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "AutoFlashEnabled")
            {
                if (windowState.AutoFlashEnabled)
                {
                    logTextBox.LogInfo("自动写入开启");
                    DisableUI();
                }
                else
                {
                    logTextBox.LogInfo("自动写入关闭");
                    EnableUI();
                }
            }
        }

        private void ShowAllDevicesEnabledChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "ShowAllDevices")
            {
                Settings.Default.showAllDevices = windowState.ShowAllDevices;
            }
        }

        private async void FlashAllAsync()
        {
            string selectedMcu = (string)mcuBox.SelectedValue;
            string filePath = filepathBox.Text;

            if (filePath.Length == 0)
            {
                logTextBox.LogError("请选择对应固件");
                return;
            }

            if (!File.Exists(filePath))
            {
                logTextBox.LogError("固件不存在!");
                return;
            }

            if (!windowState.AutoFlashEnabled)
            {
                Invoke(new Action(DisableUI));
            }

            foreach (BootloaderDevice b in FindBootloaders())
            {
                logTextBox.LogBootloader("正在尝试刷新，请不要删除设备");
                await b.Flash(selectedMcu, filePath);
                logTextBox.LogBootloader("写入完成，重新插拔或按背面复位按钮重启键盘");
                
            }

            if (!windowState.AutoFlashEnabled)
            {
                Invoke(new Action(EnableUI));
            }
        }

        private async void ResetAllAsync()
        {
            string selectedMcu = (string)mcuBox.SelectedValue;

            if (!windowState.AutoFlashEnabled)
            {
                Invoke(new Action(DisableUI));
            }

            foreach (BootloaderDevice b in FindBootloaders())
            {
                if (b.IsResettable)
                {
                    await b.Reset(selectedMcu);
                }
            }

            if (!windowState.AutoFlashEnabled)
            {
                Invoke(new Action(EnableUI));
            }
        }

        private async void ClearEepromAllAsync()
        {
            string selectedMcu = (string)mcuBox.SelectedValue;

            if (!windowState.AutoFlashEnabled)
            {
                Invoke(new Action(DisableUI));
            }

            foreach (BootloaderDevice b in FindBootloaders())
            {
                if (b.IsEepromFlashable)
                {
                    logTextBox.LogBootloader("正在尝试清除空键盘EEPROM，请不要移除设备");
                    await b.FlashEeprom(selectedMcu, "reset.eep");
                    logTextBox.LogBootloader("EEPROM 清空完成");
                }
            }

            if (!windowState.AutoFlashEnabled)
            {
                Invoke(new Action(EnableUI));
            }
        }

        private async void SetHandednessAllAsync(bool left)
        {
            string selectedMcu = (string)mcuBox.SelectedValue;
            string file = left ? "reset_left.eep" : "reset_right.eep";

            if (!windowState.AutoFlashEnabled)
            {
                Invoke(new Action(DisableUI));
            }

            foreach (BootloaderDevice b in FindBootloaders())
            {
                if (b.IsEepromFlashable)
                {
                    logTextBox.LogBootloader("正在尝试设置惯用手，请不要移除设备");
                    await b.FlashEeprom(selectedMcu, file);
                    logTextBox.LogBootloader("EEPROM 写入完成");
                }
            }

            if (!windowState.AutoFlashEnabled)
            {
                Invoke(new Action(EnableUI));
            }
        }

        private void FlashButton_Click(object sender, EventArgs e)
        {
            FlashAllAsync();
        }

        private void ResetButton_Click(object sender, EventArgs e)
        {
            ResetAllAsync();
        }

        private void ClearEepromButton_Click(object sender, EventArgs e)
        {
            ClearEepromAllAsync();
        }

        private void SetHandednessButton_Click(object sender, EventArgs e)
        {
            SetHandednessAllAsync(sender == eepromLeftToolStripMenuItem);
        }

        private List<BootloaderDevice> FindBootloaders()
        {
            return usbListener.Devices.Where(d => d is BootloaderDevice).Select(b => b as BootloaderDevice).ToList();
        }

        private void OpenFileButton_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                SetFilePath(openFileDialog.FileName);
            }
        }

        private void FilepathBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SetFilePath(filepathBox.Text);
                e.Handled = true;
            }
        }

        private void SetFilePath(string filepath)
        {
            if (!string.IsNullOrEmpty(filepath))
            {
                if (filepath.StartsWith("qmk:"))
                {
                    string unwrappedUrl = filepath[(filepath.StartsWith("qmk://") ? 6 : 4)..];
                    DownloadFile(unwrappedUrl);
                }
                else
                {
                    LoadLocalFile(filepath);
                }
            }
        }

        private void LoadLocalFile(string path)
        {
            if (!filepathBox.Items.Contains(path))
            {
                filepathBox.Items.Add(path);
            }
            filepathBox.SelectedItem = path;
        }

        private async void DownloadFile(string url)
        {
            logTextBox.LogInfo($"下载文件: {url}");

            try
            {
                string destFile = Path.Combine(KnownFolders.Downloads.Path, url[(url.LastIndexOf("/") + 1)..]);

                var client = new HttpClient();
                client.DefaultRequestHeaders.UserAgent.ParseAdd("QMK Toolbox");

                var response = await client.GetAsync(url);
                using (var fs = new FileStream(destFile, FileMode.CreateNew))
                {
                    await response.Content.CopyToAsync(fs);
                    logTextBox.LogInfo($"文件保存至: {destFile}");
                }

                LoadLocalFile(destFile);
            }
            catch (Exception e)
            {
                logTextBox.LogError($"无法下载文件: {e.Message}");
            }
        }

        private void DisableUI()
        {
            windowState.CanFlash = false;
            windowState.CanReset = false;
            windowState.CanClearEeprom = false;
        }

        private void EnableUI()
        {
            List<BootloaderDevice> bootloaders = FindBootloaders();
            windowState.CanFlash = bootloaders.Any();
            windowState.CanReset = bootloaders.Any(b => b.IsResettable);
            windowState.CanClearEeprom = bootloaders.Any(b => b.IsEepromFlashable);
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void AboutMenuItem_Click(object sender, EventArgs e)
        {
            new AboutBox().ShowDialog();
        }

        private void InstallDriversMenuItem_Click(object sender, EventArgs e)
        {
            DriverInstaller.DisplayPrompt();
        }

        private void KeyTesterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            KeyTesterWindow.GetInstance().Show(this);
            KeyTesterWindow.GetInstance().Focus();
        }

        private void HidConsoleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            HidConsoleWindow.GetInstance().Show(this);
            HidConsoleWindow.GetInstance().Focus();
        }
        #endregion

        #region Log Box
        private void LogContextMenuStrip_Opening(object sender, CancelEventArgs e)
        {
            copyToolStripMenuItem.Enabled = logTextBox.SelectedText.Length > 0;
            selectAllToolStripMenuItem.Enabled = logTextBox.Text.Length > 0;
            clearToolStripMenuItem.Enabled = logTextBox.Text.Length > 0;
        }

        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            logTextBox.Copy();
        }

        private void SelectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            logTextBox.SelectAll();
        }
        private void ClearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            logTextBox.Clear();
        }
        #endregion
    }
}
