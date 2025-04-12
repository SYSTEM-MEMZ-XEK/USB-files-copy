using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Text.RegularExpressions;
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
        private const int MaxLogLines = 900;
        private int _currentLogLines = 0;
        private NotifyIcon notifyIcon;
        private ContextMenuStrip trayMenu;
        private ToolStripMenuItem showMenuItem;
        private ToolStripMenuItem exitMenuItem;
        private int _totalFiles = 0;
        private int _successCount = 0;
        private int _failureCount = 0;
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
            var checkBoxes = new[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6, checkBox7, checkBox9, checkBox10 };
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
            var checkBoxes = new[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6, checkBox7, checkBox9, checkBox10 };
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
                LogMessage($"初始化失败: {ex.Message}",true);
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
                LogMessage("操作已取消", true);
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
                LogMessage($"==== 开始复制 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ====", true);
                // 重置计数器
                _totalFiles = 0;
                _successCount = 0;
                _failureCount = 0;
                UpdateCountDisplay();

                if (!ValidateTargetDirectory()) return;

                var extensions = GetSelectedExtensions().ToList();
                var searchOptions = extensions.Contains(AllFilesPattern)
                    ? new[] { AllFilesPattern }
                    : extensions.Distinct().ToArray();

                foreach (var drive in GetRemovableDrives())
                {
                    _cts.Token.ThrowIfCancellationRequested();
                    LogMessage($"发现U盘：{drive.Name}", true);
                    if (ContainsStopFile(drive))
                    {
                        LogMessage($"检测到阻止复制文件，跳过该U盘：{drive.Name}", true);
                        continue;
                    }

                    // 新增：创建U盘专属目录
                    string driveFolder = CreateDriveFolder(drive);
                    CopyDirectory(drive.RootDirectory, driveFolder, searchOptions);
                }

            }
            catch (OperationCanceledException)
            {
                LogMessage("用户取消操作",true);
            }
            catch (Exception ex)
            {
                LogMessage($"全局错误：{ex.Message}", true);
            }
            finally
            {
                // 记录最终统计到日志
                LogMessage($"一共复制{_totalFiles}文件，成功{_successCount}个，失败{_failureCount}个",true);
                LogMessage($"==== 复制完成 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ====\n", true);
            }

         }

        private bool ContainsStopFile(DriveInfo drive)
        {
            try
            {
                string stopFilePath = Path.Combine(drive.RootDirectory.FullName, "stop.copy");
                return File.Exists(stopFilePath);
            }
            catch (Exception ex)
            {
                LogMessage($"检查阻止文件时出错：{ex.Message}", true);
                return true; // 如果无法检查，保守处理为跳过该U盘
            }
        }

        private string CreateDriveFolder(DriveInfo drive)
        {
            try
            {
                string folderName = SanitizeFolderName(drive.VolumeLabel);
                string basePath = textBox2.Text;

                // 确保基础路径有效
                if (!Path.IsPathRooted(basePath))
                {
                    LogMessage($"无效的基础路径：{basePath}", true);
                    throw new ArgumentException("目标路径必须是绝对路径");
                }

                string fullPath = Path.Combine(basePath, folderName);

                // 处理路径长度限制
                if (fullPath.Length > 240)
                {
                    folderName = folderName.Substring(0, 240 - basePath.Length);
                    fullPath = Path.Combine(basePath, folderName);
                }

                // 创建带有序号的目录
                int counter = 1;
                string originalPath = fullPath;
                while (Directory.Exists(fullPath) || File.Exists(fullPath))
                {
                    string suffix = $"_{counter++}";
                    fullPath = originalPath.Length + suffix.Length > 260
                        ? originalPath.Substring(0, 260 - suffix.Length) + suffix
                        : originalPath + suffix;
                }

                Directory.CreateDirectory(fullPath);
                LogMessage($"成功创建目录：{fullPath}",false);
                return fullPath;
            }
            catch (Exception ex)
            {
                LogMessage($"目录创建失败细节：{ex.GetType().Name} - {ex.Message}", true);
                throw new ApplicationException($"无法为U盘 {drive.Name} 创建目录", ex);
            }
        }

        // 新增方法：清理非法字符
        private string SanitizeFolderName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "USB_DRIVE";

            // 移除非法字符
            var invalidChars = Path.GetInvalidFileNameChars();
            var cleanName = new string(name
                .Where(c => !invalidChars.Contains(c))
                .ToArray())
                .Trim();

            // 处理保留设备名称（如CON、PRN等）
            var reservedNames = new[] { "CON", "PRN", "AUX", "NUL",
                              "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
                              "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };
            if (reservedNames.Contains(cleanName.ToUpper())) cleanName += "_DATA";

            // 移除首尾点和空格
            cleanName = cleanName.Trim('.', ' ');

            // 替换连续空格为单个下划线
            cleanName = Regex.Replace(cleanName, @"\s+", "_");

            // 处理空名称情况
            if (string.IsNullOrWhiteSpace(cleanName)) return "USB_DRIVE";

            // 限制长度并添加后缀
            return cleanName.Length > 50
                ? $"{cleanName.Substring(0, 45)}_TRUNC"
                : cleanName;
        }

        private bool ValidateTargetDirectory()
        {
            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                LogMessage("错误：未选择目标目录", true);
                return false;
            }

            try
            {
                // 新增路径格式验证
                if (!Path.IsPathRooted(textBox2.Text))
                {
                    LogMessage("错误：目标路径必须是绝对路径",true);
                    return false;
                }

                var fullPath = Path.GetFullPath(textBox2.Text);
                if (fullPath.StartsWith(@"\\?\")) // 处理长路径格式
                {
                    LogMessage("警告：长路径格式可能需要系统支持", true);
                }

                Directory.CreateDirectory(fullPath);
                textBox2.Text = fullPath; // 标准化路径格式
                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"目录验证失败：{ex.Message}",true);
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
                    LogMessage($"驱动器访问失败：{drive.Name} - {ex.Message}", true);
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
                // 合并两条路径生成语句
                var destDir = CreateDestinationDirectory(source, targetParentDir);
                CopyFilesWithPatterns(source, destDir, searchPatterns);
                ProcessSubdirectories(source, targetParentDir, searchPatterns);
            }
            catch (Exception ex)
            {
                LogMessage($"目录处理失败: {source.FullName} | 错误：{ex.Message}",true);
            }
        }

        // 保持原有的CreateDestinationDirectory方法不变
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
                    LogMessage($"文件访问被拒绝：{pattern} - {ex.Message}",true);
                }
            }
        }

        private void CopySingleFile(FileInfo file, string destDir)
        {
            Interlocked.Increment(ref _totalFiles);
            try
            {
                var destPath = Path.Combine(destDir, file.Name);
                using (var sourceStream = File.OpenRead(file.FullName))
                using (var destStream = File.Create(destPath))
                {
                    sourceStream.CopyTo(destStream);
                }
                LogMessage($"成功复制：{file.FullName} -> {destPath}", false);
                Interlocked.Increment(ref _successCount);
            }
            catch (Exception ex)
            {
                LogMessage($"复制失败：{file.FullName} | 错误：{ex.Message}", true);
                Interlocked.Increment(ref _failureCount);
            }
            finally
            {
                UpdateCountDisplay();
            }
        }

        private void UpdateCountDisplay()
        {
            // 线程安全读取计数器
            int total = Interlocked.CompareExchange(ref _totalFiles, 0, 0);
            int success = Interlocked.CompareExchange(ref _successCount, 0, 0);
            int failure = Interlocked.CompareExchange(ref _failureCount, 0, 0);

            // 安全更新UI
            if (label4.InvokeRequired)
            {
                label4.BeginInvoke((Action)(() =>
                {
                    label4.Text = $"一共复制{total}文件，成功{success}个，失败{failure}个";
                }));
            }
            else
            {
                label4.Text = $"一共复制{total}文件，成功{success}个，失败{failure}个";
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
            AddAudioExtension(extensions);
            return extensions;
        }

        // 定义文件类型扩展名常量
        private static readonly string[] PowerPointExtensions = { "*.ppt", "*.pptx" };
        private static readonly string[] WordExtensions = { "*.doc", "*.docx", "*.txt" };
        private static readonly string[] ExcelExtensions = { "*.xlsx", "*.xls" };
        private static readonly string[] PdfExtensions = { "*.pdf" };
        private static readonly string[] ImageExtensions = { "*.jpg", "*.jpeg", "*.png", "*.gif", "*.bmp", "*.webp" };
        private static readonly string[] VideoExtensions = { "*.mp4", "*.avi", "*.mov", "*.mkv", "*.wmv", "*.flv" };
        private static readonly string[] AudioExtensions = { "*.mp3", "*.wma", "*.wav", "*.ape", "*.ogg","*.flac","*.aac" };
        private static readonly string[] CompressedExtensions = { "*.zip", "*.rar", "*.7z", "*.tar", "*.gz", "*.bz2", "*.xz", "*.zst", "*.001", ".iso", ".wim", "cab" };

        private void AddOfficeExtensions(List<string> extensions)
        {
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
        private void AddAudioExtension(List<string> extensions)
        {
            if (checkBox10.Checked) extensions.AddRange(AudioExtensions);
        }

        private void LogMessage(string message, bool isError = false)
        {
            try
            {
                var logEntry = $"[{DateTime.Now:HH:mm:ss}] {message}";
                var stackTrace = new StackTrace(2, true); // 跳过前两层调用
                var frame = stackTrace.GetFrame(0);
                var lineInfo = $"{Path.GetFileName(frame.GetFileName())}:{frame.GetFileLineNumber()}";

                // 如果是错误，更新UI日志显示
                if (isError)
                {
                    UpdateLogBuffer(logEntry);
                    UpdateLogDisplay(logEntry);
                }

                // 所有日志都写入文件
                WriteToLogFile(logEntry);
            }
            catch (Exception ex)
            {
                var errorMessage = $"[{DateTime.Now:HH:mm:ss}] 日志记录失败: {ex.Message}";
                WriteToLogFile(errorMessage);
                UpdateLogDisplay(errorMessage);
            }
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
                    Task.Run(async () => {
                        await Task.Delay(50);
                        UpdateLogDisplay(_logBuffer.ToString());
                        _logBuffer.Clear();
                    });

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

            richTextBox1.AppendText(entry + Environment.NewLine);
            richTextBox1.ScrollToCaret();
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
                var errorMessage = $"[{DateTime.Now:HH:mm:ss}] 日志文件写入失败: {ex.Message}";
                // 直接更新UI显示错误
                SafeInvokeAsync(() => richTextBox1.AppendText(errorMessage + Environment.NewLine));

                // 尝试写入备用日志
                try
                {
                    string backupLog = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "CopyLog.txt");
                    File.AppendAllText(backupLog, errorMessage + Environment.NewLine);
                }
                catch { }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // 先取消所有后台操作
            _cts?.Cancel();

            // 等待正在进行的日志写入完成
            Task.Run(async () => {
                await Task.Delay(500); // 根据实际情况调整等待时间
                _usbWatcher?.Stop();
                _usbWatcher?.Dispose();
            }).Wait(1000); // 最多等待1秒

            // 确保RichTextBox安全释放
            if (!richTextBox1.IsDisposed)
            {
                richTextBox1.Dispose();
            }

            base.OnFormClosing(e);
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

        private void button2_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    textBox2.Text = fbd.SelectedPath;
                }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            notifyIcon.Visible = true;
            this.ShowInTaskbar = false;
        }
    }
}