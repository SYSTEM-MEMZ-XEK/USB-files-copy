// 在Program.cs中添加以下代码（需确保存在该文件）
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace U盘文件复制
{
    static class Program
    {
        // 添加Win32 API声明
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        private const int SW_RESTORE = 9;

        private static Mutex _mutex;

        [STAThread]
        static void Main()
        {
            // 创建唯一互斥体（使用GUID确保唯一性）
            bool createdNew;
            _mutex = new Mutex(true, "{B7E9A614-5D4B-4A3B-9A5E-3C9D8D7F6C5A}", out createdNew);

            if (!createdNew)
            {
                // 已存在运行实例
                ActivateExistingInstance();
                return;
            }

            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }
            finally
            {
                _mutex?.ReleaseMutex();
            }
        }

        private static void ActivateExistingInstance()
        {
            Process current = Process.GetCurrentProcess();
            foreach (Process process in Process.GetProcessesByName(current.ProcessName))
            {
                if (process.Id != current.Id)
                {
                    IntPtr handle = process.MainWindowHandle;
                    if (handle != IntPtr.Zero)
                    {
                        // 恢复窗口（如果最小化）
                        ShowWindowAsync(handle, SW_RESTORE);
                        // 置前窗口
                        SetForegroundWindow(handle);
                    }
                    break;
                }
            }

            MessageBox.Show("程序已经在运行中", "提示",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);

            Environment.Exit(0);
        }
    }
}