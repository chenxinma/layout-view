# layout-view

一个用于分析Excel文件工作表布局和数据密度的Rust库。

## 功能

- 计算Excel文件中每个工作表的数据密度
- 识别数据的起始和结束行列位置
- 跳过隐藏和非常隐藏的工作表
- 输出JSON格式的分析结果
- 基于密度和数据类型混合度对工作表进行分类（行列表 vs 表单）
- 分析工作表中每列的数据类型分布（数值型 vs 文本型）
- 提供C FFI接口，支持从Python等其他语言调用

## 安装

确保您已安装 Rust（2024版本）和 Cargo：

```bash
# 克隆项目
git clone <repository-url>
cd layout-view

# 或者直接使用 Cargo 创建新项目
cargo new layout-view --bin
```

## 依赖项

- `calamine = "0.32.0"` - 用于读取 Excel 文件的库
- `serde = { version = "1.0", features = ["derive"] }` - 用于序列化
- `serde_json = "1.0"` - 用于JSON处理
- `regex = "1.0"` - 用于正则表达式匹配（数据类型识别）
- `lazy_static = "1.4"` - 用于静态正则表达式初始化
- `libc = "0.2"` - 用于C FFI接口

## 构建和运行

### 构建项目
```bash
# 构建项目
cargo build

# 构建发布版本
cargo build --release
```

### 运行测试
```bash
# 运行测试
cargo test
```

### 运行程序
```bash
# 运行程序分析Excel文件
cargo run -- <xlsx_file_path>
```

### 其他常用命令
```bash
# 检查代码
cargo check

# 格式化代码
cargo fmt

# 检查代码风格
cargo clippy
```

## 使用方法

```bash
cargo run -- <xlsx_file_path>
```

程序将输出JSON格式的结果，包含以下信息：
- `sheet_name`: 工作表名称
- `first_row`: 第一个包含数据的行
- `first_col`: 第一个包含数据的列
- `end_row`: 最后一个包含数据的行
- `end_col`: 最后一个包含数据的列
- `total_cells`: 指定范围内的总单元格数
- `data_cells`: 包含实际数据（非空白）的单元格数
- `density`: 数据密度（data_cells / total_cells）
- `visible`: 工作表的可见性状态（"Visible", "Hidden", "VeryHidden"）
- `first_row_first_col_content`: 第一行第一列单元格的内容
- `last_row_first_col_content`: 最后一行第一列单元格的内容
- `data_type_mix`: 数据类型混合程度（使用香农熵计算）
- `column_data_types`: 每列的数据类型信息，包括数值型和文本型的分布
- `sheet_type`: 工作表类型分类（Data/表单/Form/未知/Unknown）
- `classification_reason`: 分类原因说明

## 示例输出

```json
[
  {
    "sheet_name": "Data Sheet",
    "first_row": 0,
    "first_col": 0,
    "end_row": 3,
    "end_col": 2,
    "total_cells": 12,
    "data_cells": 9,
    "density": 0.75,
    "visible": "Visible",
    "first_row_first_col_content": "ID",
    "last_row_first_col_content": "4",
    "data_type_mix": 0.8,
    "column_data_types": [
      {
        "column_index": 0,
        "numeric_count": 4,
        "text_count": 0,
        "total_count": 4,
        "numeric_type_ratio": 1.0
      },
      {
        "column_index": 1,
        "numeric_count": 4,
        "text_count": 0,
        "total_count": 4,
        "numeric_type_ratio": 1.0
      },
      {
        "column_index": 2,
        "numeric_count": 0,
        "text_count": 4,
        "total_count": 4,
        "numeric_type_ratio": 0.0
      }
    ],
    "sheet_type": "Data",
    "classification_reason": "density: 0.750, data_type_mix: 0.800"
  }
]
```

## 算法说明

程序通过以下方式分析Excel工作表：
1. 读取Excel文件并遍历每个工作表
2. 检查工作表的可见性，只处理可见的工作表（跳过隐藏和非常隐藏的工作表）
3. 确定有效数据范围（first_row, first_col, end_row, end_col），排除起始的连续空白行列
4. 统计指定范围内的单元格，将非空白（非空字符串或全空格）单元格计为数据单元格
5. 计算密度为数据单元格数量除以总单元格数量
6. 分析每列的数据类型分布（数值型 vs 文本型），支持整数、小数、千分位数、百分数等数值格式
7. 使用香农熵计算数据类型混合程度，值越高表示数据类型越多样化
8. 基于密度和数据类型混合度对工作表进行分类：
   - 高密度（>0.46）或低密度但高数据类型混合度的工作表分类为 "Data"（行列表）
   - 低密度且低数据类型混合度的工作表分类为 "Form"（表单）

## 项目结构

```
layout-view/
├── Cargo.toml          # 项目配置和依赖
├── Cargo.lock          # 锁定依赖版本
├── src/
│   ├── lib.rs          # 主要库源代码
│   └── main.rs         # 命令行程序入口
├── README.md           # 用户说明文档
└── IFLOW.md            # 项目文档
```

## C FFI 接口

项目提供了 C 语言接口，允许从 Python 等其他语言调用：

- `classify_excel_sheets_c(xlsx_path)`: 分析 Excel 文件并返回 JSON 字符串
- `free_c_string(ptr)`: 释放由 Rust 分配的字符串内存

Python 调用示例：
```python
import ctypes

# 加载动态库
lib = ctypes.CDLL('./target/release/liblayout_view.so')  # Linux
# lib = ctypes.CDLL('./target/release/liblayout_view.dylib')  # macOS
# lib = ctypes.CDLL('./target/release/layout_view.dll')  # Windows

# 定义函数参数和返回类型
lib.classify_excel_sheets_c.argtypes = [ctypes.c_char_p]
lib.classify_excel_sheets_c.restype = ctypes.c_char_p

lib.free_c_string.argtypes = [ctypes.c_char_p]
lib.free_c_string.restype = None

# 调用函数
xlsx_path = b"example.xlsx"
result = lib.classify_excel_sheets_c(xlsx_path)
json_result = ctypes.c_char_p(result).value.decode('utf-8')
print(json_result)

# 释放内存
lib.free_c_string(result)
```

## 开发约定

- 使用 2024 版本的 Rust 语言标准
- 代码应遵循 Rust 的最佳实践和命名约定
- 测试代码应放在 `#[cfg(test)]` 模块中
- 使用 `cargo fmt` 保持代码格式一致
- 使用 `cargo clippy` 保持代码质量