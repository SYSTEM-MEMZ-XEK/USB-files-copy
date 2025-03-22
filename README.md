# U盘文件复制工具

![.NET Framework](https://img.shields.io/badge/.NET%20Framework-4.7.2-blue)
![License](https://img.shields.io/badge/License-MIT-green)

一款基于C#开发的智能U盘文件复制工具，可自动监控USB设备接入并复制指定类型文件到目标目录，支持后台运行和文件类型过滤。

## 功能特性

- 🔌 **自动检测**：实时监控U盘插入事件
- 🗂 **智能过滤**：支持多种常见文件类型组合选择：
  - Office文档（PPT/Word/Excel/PDF）
  - 图片（JPG/PNG/GIF等）
  - 视频（MP4/AVI/MOV等）
  - 压缩文件（ZIP/RAR/7z等）
  - 自定义扩展名（支持通配符）
- 📁 **目录管理**：自动创建与U盘相同的目录结构
- 📝 **日志记录**：完整操作日志记录，支持日志轮转（保留最近200条）
- 🛠 **后台运行**：支持最小化到系统托盘
- ⚡ **并发控制**：使用信号量保证单任务执行
- 🚫 **操作取消**：支持通过CancellationToken取消任务

## 快速开始

### 环境要求
- .NET Framework 4.7.2
- Windows 7及以上系统

### 安装步骤
1. 克隆仓库：
```bash
git clone https://github.com/yourusername/usb-file-copier.git
