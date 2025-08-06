# Excel页边距批量修改工具

[![Python](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-1.1.0-orange.svg)](CHANGELOG.md)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)](https://github.com/your-username/excel-margin-tool)

一个功能完整的Excel页边距批量修改工具，支持GUI界面操作。

## 功能特点

- 🎯 **批量处理**: 一次性处理多个Excel文件
- 📁 **多格式支持**: 支持.xlsx、.xlsm、.xls格式
- ⚙️ **灵活设置**: 可自定义所有页边距参数
- 🎨 **预设方案**: 提供默认、窄边距、宽边距、最小边距等预设
- 💾 **配置管理**: 支持保存和加载配置
- 📊 **进度显示**: 实时显示处理进度
- 📝 **详细日志**: 完整的处理日志记录
- 🔍 **页边距检查**: 查看现有文件的页边距设置

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 方法1: 使用启动脚本（推荐）

1. **Windows用户**：
   - 双击 `start.bat` 或 `start.py` 文件
   - 这些脚本会自动处理编码问题

2. **手动启动**：
   ```bash
   python start.py
   ```

### 方法2: 直接运行

```bash
python excel_margin_app.py
```

## 常见问题

### 乱码问题解决

如果遇到中文乱码问题：

1. **使用启动脚本**：推荐使用 `start.py` 或 `start.bat`
2. **手动设置编码**：
   ```bash
   chcp 65001
   python excel_margin_app.py
   ```
3. **使用Python启动脚本**：
   ```bash
   python start.py
   ```

## 页边距参数说明

- **左边距**: 页面左边距
- **右边距**: 页面右边距  
- **上边距**: 页面上边距
- **下边距**: 页面下边距
- **页眉边距**: 页眉到页面顶部的距离
- **页脚边距**: 页脚到页面底部的距离

**注意**: 所有页边距参数的单位为**厘米**

## 预设方案

- **默认设置**: 标准页边距（左/右: 1.8cm, 上/下: 1.9cm, 页眉/页脚: 0.8cm）
- **窄边距**: 紧凑布局（左/右/上/下: 1.5cm, 页眉/页脚: 0.6cm）
- **宽边距**: 宽松布局（左/右/上/下: 2.0cm, 页眉/页脚: 1.0cm）
- **最小边距**: 最小边距（左/右/上/下: 0.5cm, 页眉/页脚: 0.2cm）

## 注意事项

1. **文件格式支持**:
   - `.xlsx` 和 `.xlsm`: 完全支持页边距修改
   - `.xls`: 功能有限，需要安装xlrd和xlwt库

2. **文件安全**:
   - 建议先备份重要文件
   - 可选择输出文件夹避免覆盖原文件

3. **处理限制**:
   - 大文件处理可能需要较长时间
   - 可以随时点击"停止处理"中断操作

## 配置管理

- **保存配置**: 将当前设置保存为JSON文件
- **加载配置**: 从JSON文件加载之前保存的设置
- **自动加载**: 程序启动时自动加载默认配置文件

## 快速开始

### 从GitHub克隆

```bash
git clone https://github.com/your-username/excel-margin-tool.git
cd excel-margin-tool
pip install -r requirements.txt
python excel_margin_app.py
```

### 从PyPI安装（即将推出）

```bash
pip install excel-margin-tool
excel-margin-tool
```

## 系统要求

- Python 3.8+
- Windows/macOS/Linux
- 足够的磁盘空间存储处理后的文件

## 故障排除

1. **找不到Excel文件**: 检查文件类型选择是否正确
2. **处理失败**: 检查文件是否被其他程序占用
3. **权限错误**: 确保对目标文件夹有写入权限
4. **内存不足**: 减少同时处理的文件数量

## 贡献

我们欢迎所有形式的贡献！请查看 [CONTRIBUTING.md](CONTRIBUTING.md) 了解如何参与项目开发。

## 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 更新日志

详细的更新历史请查看 [CHANGELOG.md](CHANGELOG.md)。

### 最新版本 v1.1.0
- 添加.xls文件支持
- 增加进度条显示
- 添加配置保存/加载功能
- 新增页边距检查功能
- 改进错误处理和日志记录
- 添加预设页边距方案

### v1.0.0
- 初始版本发布
- 支持.xlsx和.xlsm文件
- 基本的页边距修改功能 