# Excel拆分工具使用说明

## 简介

Excel拆分工具是一个用于批量处理Excel文件的应用程序，可以根据指定条件对Excel文件进行拆分，支持多工作表(sheet)并拆功能。本工具可以帮助您轻松处理大量数据，提高工作效率。

## 安装指南

### 1. 安装Python

1. 访问 [Python官网](https://www.python.org/downloads/) 下载最新版Python
2. 安装时勾选"Add Python to PATH"选项
3. 完成安装后，打开命令行窗口，输入`python --version`验证安装成功

### 2. 依赖库安装

本程序会在启动时**自动检测并安装**所需依赖库。主要依赖包括：
- PyQt5
- pandas
- openpyxl

#### 自动安装流程
1. 启动程序后，系统会自动检测缺少的依赖库
2. 弹出依赖安装对话框，点击"安装依赖项"按钮
3. 等待安装完成后，程序将自动启动

#### 手动安装方法（如自动安装失败）
如果自动安装失败，可通过命令行手动安装：
```
pip install PyQt5
pip install pandas
pip install openpyxl
```

或一次性安装所有依赖：
```
pip install PyQt5 pandas openpyxl
```

### 3. 程序运行

双击`excel_splitter_v1.1.0.py`文件或在命令行中运行：
```
python excel_splitter_v1.1.0.py
```

## 使用说明

### 基本流程

1. 选择Excel文件
2. 创建条件组
3. 添加筛选条件
4. 开始批量处理

### 详细步骤

#### 1. 选择Excel文件
1. 点击主界面上的"选择Excel文件"按钮
2. 在弹出的文件选择对话框中选择要处理的Excel文件
3. 程序会自动读取文件中的所有工作表

#### 2. 创建和管理条件组
1. 点击"添加条件组"创建新的条件组
2. 在右侧面板中可编辑条件组名称
3. 使用"删除条件组"按钮可移除不需要的条件组
4. 条件组列表会显示所有创建的条件组

#### 3. 添加筛选条件
1. 选择一个条件组后，点击"添加筛选条件"按钮
2. 在弹出的对话框中依次选择：
   - 工作表
   - 列名
   - 筛选值（可多选）
3. 点击"添加"确认添加该条件
4. 已添加的条件会显示在右侧的条件表格中
5. 可以为同一个条件组添加多个筛选条件

#### 4. 条件组导入导出
1. 点击"导出条件组"将已创建的条件组保存为JSON文件
2. 点击"导入条件组"可从JSON文件中读取预设的条件组

#### 5. 开始批量处理
1. 设置完所有条件组后，点击"开始批量处理"按钮
2. 确认后程序将根据每个条件组创建独立的Excel文件
3. 处理过程中会显示进度条和状态信息
4. 处理完成后会弹出结果提示框

### 高级功能

#### 公式处理
- 程序会自动处理Excel文件中的公式引用，确保在拆分后的文件中公式仍然正确工作
- 对于复杂格式的Excel文件，会使用备用处理方法，但可能会丢失部分格式信息

## 常见问题

### Q: 启动程序时报错，提示缺少依赖库
A: 尝试手动安装依赖库，使用命令：`pip install PyQt5 pandas openpyxl`

### Q: 处理大文件时程序运行很慢
A: 大文件处理需要更多时间和内存，请耐心等待或考虑将文件拆分为多个小文件处理

### Q: 拆分后的文件中公式显示错误或#REF!错误
A: 部分复杂公式可能无法正确调整，特别是包含复杂跨表引用的公式，请检查并手动修正

### Q: 处理后的文件丢失了部分格式
A: 对于某些非标准格式的Excel文件，程序会使用备用方法处理，可能导致部分格式丢失

## 技术支持

如有问题或建议，请提交问题报告或联系开发者。
