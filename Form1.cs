using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace U盘文件复制
{
    public partial class Form1 : Form
    {
        private ManagementEventWatcher _usbWatcher;
        private string _logPath;
        private const string AllFilesPattern = "*.*";
        private readonly StringBuilder _logBuffer = new StringBuilder();
        private const int MaxLogLines = 200;
        private int _currentLogLines = 0;
        private NotifyIcon notifyIcon;
        private ContextMenuStrip trayMenu;
        private ToolStripMenuItem showMenuItem;
        private ToolStripMenuItem exitMenuItem;

        // 新增并发控制成员
        private readonly SemaphoreSlim _copyLock = new SemaphoreSlim(1, 1);
        private CancellationTokenSource _cts = new CancellationTokenSource();

        public Form1()
        {
            InitializeComponent();
            InitializeUsbWatcher();
            SetupControls();
            SetupCheckboxEvents();
            InitializeTrayIcon();  // 新增的初始化方法
        }

        private void InitializeTrayIcon()
        {
            trayMenu = new ContextMenuStrip();

            showMenuItem = new ToolStripMenuItem("显示主界面");
            exitMenuItem = new ToolStripMenuItem("退出程序");

            // 添加菜单项
            trayMenu.Items.Add(showMenuItem);
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add(exitMenuItem);

            // 配置托盘图标
            notifyIcon = new NotifyIcon();
            notifyIcon.Text = "U盘文件复制工具";
            notifyIcon.Icon = this.Icon;
            notifyIcon.ContextMenuStrip = trayMenu;
            notifyIcon.Visible = false;

            // 事件绑定
            notifyIcon.DoubleClick += (s, e) => ShowMainWindow();
            showMenuItem.Click += (s, e) => ShowMainWindow();
            exitMenuItem.Click += (s, e) => ExitApplication();
        }


        private void SetupCheckboxEvents()
        {
            
        }

        private void checkBox8_CheckedChanged(object sender, EventArgs e)
        {
            // 确保textBox1的启用状态与checkBox8同步
            textBox1.Enabled = !checkBox8.Checked;
            // 强制同步所有复选框的状态
            ToggleAllCheckboxes(checkBox8.Checked);
        }

        private void SyncCheckBoxStates(object sender, EventArgs e)
        {
            var checkBoxes = new[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6, checkBox7, checkBox9 };
            checkBox8.Checked = checkBoxes.All(cb => cb.Checked);
        }

        private void SetupControls()
        {
            checkBox8.CheckedChanged += (s, e) =>
            {
                textBox1.Enabled = !checkBox8.Checked;
                ToggleAllCheckboxes(checkBox8.Checked);
            };
            textBox1.Enabled = !checkBox8.Checked;
        }

        private void ToggleAllCheckboxes(bool state)
        {
            var checkBoxes = new[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6, checkBox7, checkBox9 };
            foreach (var checkBox in checkBoxes)
            {
                checkBox.Checked = state;
            }
        }

        private void InitializeUsbWatcher()
        {
            try
            {
                var query = new WqlEventQuery("SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 2");
                _usbWatcher = new ManagementEventWatcher(query);
                _usbWatcher.EventArrived += async (sender, e) => await HandleUsbEvent();
                _usbWatcher.Start();
            }
            catch (Exception ex)
            {
                LogMessage($"初始化失败: {ex.Message}");
            }
        }

        private async Task HandleUsbEvent()
        {
            if (!await _copyLock.WaitAsync(0)) return;

            try
            {
                await SafeInvokeAsync(() => this.Hide());
                _cts = new CancellationTokenSource();
                await Task.Run(() => CopyFiles(checkBox1.Checked), _cts.Token);
            }
            catch (OperationCanceledException)
            {
                LogMessage("操作已取消");
            }
            finally
            {
                // 移除显示窗口的代码
                _copyLock.Release();
            }
        }

        private Task SafeInvokeAsync(Action action)
        {
            if (InvokeRequired)
                return Task.Factory.FromAsync(BeginInvoke(action), _ => { });
            else
            {
                action();
                return Task.CompletedTask;
            }
        }

        private void CopyFiles(bool silent = false)
        {
            try
            {
                _cts.Token.ThrowIfCancellationRequested();
                _logPath = Path.Combine(textBox2.Text, "CopyLog.txt");
                LogMessage($"==== 开始复制 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ====");

                if (!ValidateTargetDirectory()) return;

                var extensions = GetSelectedExtensions().ToList();
                var searchOptions = extensions.Contains(AllFilesPattern)
                    ? new[] { AllFilesPattern }
                    : extensions.Distinct().ToArray();

                foreach (var drive in GetRemovableDrives())
                {
                    _cts.Token.ThrowIfCancellationRequested();
                    LogMessage($"发现U盘：{drive.Name}");
                    CopyDirectory(drive.RootDirectory, textBox2.Text, searchOptions);
                }

                LogMessage($"==== 复制完成 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ====\n");
            }
            catch (OperationCanceledException)
            {
                LogMessage("用户取消操作");
            }
            catch (Exception ex)
            {
                LogMessage($"全局错误：{ex.Message}");
            }
        }

        private bool ValidateTargetDirectory()
        {
            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                LogMessage("错误：未选择目标目录");
                return false;
            }

            try
            {
                Directory.CreateDirectory(textBox2.Text);
                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"目录创建失败：{ex.Message}");
                return false;
            }
        }

        private IEnumerable<DriveInfo> GetRemovableDrives()
        {
            foreach (var drive in DriveInfo.GetDrives())
            {
                bool isValidDrive = false;
                try
                {
                    isValidDrive = drive.DriveType == DriveType.Removable && drive.IsReady;
                }
                catch (IOException ex)
                {
                    LogMessage($"驱动器访问失败：{drive.Name} - {ex.Message}");
                }

                if (isValidDrive)
                {
                    yield return drive;
                }
            }
        }

        private void CopyDirectory(DirectoryInfo source, string targetParentDir, string[] searchPatterns)
        {
            try
            {
                var destDir = CreateDestinationDirectory(source, targetParentDir);
                CopyFilesWithPatterns(source, destDir, searchPatterns);
                ProcessSubdirectories(source, targetParentDir, searchPatterns);
            }
            catch (Exception ex)
            {
                LogMessage($"目录处理失败: {source.FullName} | 错误：{ex.Message}");
            }
        }

        private string CreateDestinationDirectory(DirectoryInfo source, string targetParentDir)
        {
            var relativePath = source.FullName.Substring(Path.GetPathRoot(source.FullName).Length);
            var destDir = Path.Combine(targetParentDir, relativePath);
            Directory.CreateDirectory(destDir);
            return destDir;
        }

        private void CopyFilesWithPatterns(DirectoryInfo source, string destDir, string[] searchPatterns)
        {
            foreach (var pattern in searchPatterns)
            {
                try
                {
                    foreach (var file in source.EnumerateFiles(pattern))
                    {
                        _cts.Token.ThrowIfCancellationRequested();
                        CopySingleFile(file, destDir);
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    LogMessage($"文件访问被拒绝：{pattern} - {ex.Message}");
                }
            }
        }

        private void CopySingleFile(FileInfo file, string destDir)
        {
            try
            {
                var destPath = Path.Combine(destDir, file.Name);
                using (var sourceStream = File.OpenRead(file.FullName))
                using (var destStream = File.Create(destPath))
                {
                    sourceStream.CopyTo(destStream);
                }
                LogMessage($"成功复制：{file.FullName} -> {destPath}");
            }
            catch (Exception ex)
            {
                LogMessage($"复制失败：{file.FullName} | 错误：{ex.Message}");
            }
        }

        private void ProcessSubdirectories(DirectoryInfo source, string targetParentDir, string[] searchPatterns)
        {
            foreach (var dir in source.EnumerateDirectories())
            {
                _cts.Token.ThrowIfCancellationRequested();
                CopyDirectory(dir, targetParentDir, searchPatterns);
            }
        }

        private IEnumerable<string> GetSelectedExtensions()
        {
            if (checkBox8.Checked) return new[] { AllFilesPattern };

            var extensions = new List<string>();
            AddOfficeExtensions(extensions);
            AddMediaExtensions(extensions);
            AddCompressedExtensions(extensions);
            AddCustomExtensions(extensions);
            return extensions;
        }

        // 定义文件类型扩展名常量
        private static readonly string[] PowerPointExtensions = { "*.ppt", "*.pptx" };
        private static readonly string[] WordExtensions = { "*.doc", "*.docx", "*.txt" };
        private static readonly string[] ExcelExtensions = { "*.xlsx", "*.xls" };
        private static readonly string[] PdfExtensions = { "*.pdf" };
        private static readonly string[] ImageExtensions = { "*.jpg", "*.jpeg", "*.png", "*.gif", "*.bmp", "*.webp" };
        private static readonly string[] VideoExtensions = { "*.mp4", "*.avi", "*.mov", "*.mkv", "*.wmv", "*.flv" };
        private static readonly string[] CompressedExtensions = { "*.zip", "*.rar", "*.7z", "*.tar.gz", "*.gz", "*.bz2", "*.xz", "*.zst", "*.001", ".iso", ".wim", "cab" };

        private void AddOfficeExtensions(List<string> extensions)
        {
            // 使用更具语义化的复选框命名会更理想（如 cbPowerPoint）
            if (checkBox1.Checked) extensions.AddRange(PowerPointExtensions);
            if (checkBox2.Checked) extensions.AddRange(WordExtensions);
            if (checkBox3.Checked) extensions.AddRange(ExcelExtensions);
            if (checkBox4.Checked) extensions.AddRange(PdfExtensions);
        }

        private void AddMediaExtensions(List<string> extensions)
        {
            if (checkBox5.Checked) extensions.AddRange(ImageExtensions);
            if (checkBox6.Checked) extensions.AddRange(VideoExtensions);
        }

        private void AddCompressedExtensions(List<string> extensions)
        {
            if (checkBox9.Checked) extensions.AddRange(CompressedExtensions);
        }

        private void AddCustomExtensions(List<string> extensions)
        {
            if (!checkBox7.Checked || string.IsNullOrWhiteSpace(textBox1.Text)) return;

            extensions.AddRange(
                textBox1.Text.Split(',')
                    .Select(ext => ext.Trim().TrimStart('.'))
                    .Where(ext => !string.IsNullOrWhiteSpace(ext))
                    .Select(ext => $"*.{ext}")
            );
        }

        private void LogMessage(string message)
        {
            var logEntry = $"[{DateTime.Now:HH:mm:ss}] {message}";

            UpdateLogBuffer(logEntry);
            UpdateLogDisplay(logEntry);
            WriteToLogFile(logEntry);
        }

        private void UpdateLogBuffer(string entry)
        {
            lock (_logBuffer)
            {
                _logBuffer.AppendLine(entry);
                _currentLogLines++;

                if (_currentLogLines > MaxLogLines * 2)
                {
                    var lines = _logBuffer.ToString().Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    _logBuffer.Clear();

                    // 使用兼容的拼接方式
                    var keepLines = lines.Skip(lines.Length - MaxLogLines);
                    foreach (var line in keepLines)
                    {
                        _logBuffer.AppendLine(line);
                    }

                    // 修正行数计数器
                    _currentLogLines = keepLines.Count();

                    // 移除最后一个空行
                    if (_logBuffer.Length > 0 && _logBuffer[_logBuffer.Length - 1] == '\n')
                    {
                        _logBuffer.Length--;
                    }
                }
            }
        }

        private void UpdateLogDisplay(string entry)
        {
            if (richTextBox1.InvokeRequired)
            {
                richTextBox1.BeginInvoke((Action)(() => UpdateLogDisplay(entry)));
                return;
            }

            richTextBox1.SuspendLayout();
            try
            {
                if (richTextBox1.Lines.Length >= MaxLogLines)
                {
                    richTextBox1.Select(0, richTextBox1.GetFirstCharIndexFromLine(1));
                    richTextBox1.SelectedText = "";
                }

                richTextBox1.AppendText(entry + Environment.NewLine);
                richTextBox1.ScrollToCaret();
            }
            finally
            {
                richTextBox1.ResumeLayout();
            }
        }

        private void WriteToLogFile(string entry)
        {
            try
            {
                if (!string.IsNullOrEmpty(_logPath))
                {
                    File.AppendAllText(_logPath, entry + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {
                richTextBox1.BeginInvoke((Action)(() =>
                {
                    richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] 日志文件写入失败: {ex.Message}{Environment.NewLine}");
                }));
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;  // 取消关闭
                this.Hide();
                notifyIcon.Visible = true;
                this.ShowInTaskbar = false;
                return;
            }

            // 正常关闭时的清理逻辑
            _cts.Cancel();
            _usbWatcher?.Stop();
            _usbWatcher?.Dispose();
            notifyIcon?.Dispose();  // 释放托盘资源
            base.OnFormClosing(e);
        }


        private void button2_Click_1(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    textBox2.Text = fbd.SelectedPath;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            notifyIcon.Visible = true;
            this.ShowInTaskbar = false;
        }

        private void ShowMainWindow()
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            notifyIcon.Visible = false;
        }

        // 新增的退出方法
        private void ExitApplication()
        {
            notifyIcon.Visible = false;
            Application.Exit();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            notifyIcon.Visible = true;
            this.ShowInTaskbar = false;
        }

    }
}