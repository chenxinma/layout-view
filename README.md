# layout-view

一个用于分析Excel文件工作表布局和数据密度的Rust库。

## 功能

- 计算Excel文件中每个工作表的数据密度
- 识别数据的起始和结束行列位置
- 跳过隐藏和非常隐藏的工作表
- 输出JSON格式的分析结果

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
    "visible": "Visible"
  }
]
```

## 算法说明

程序通过以下方式计算数据密度：
1. 读取Excel文件并遍历每个工作表
2. 检查工作表的可见性，只处理可见的工作表（跳过隐藏和非常隐藏的工作表）
3. 确定有效数据范围（first_row, first_col, end_row, end_col），排除起始的连续空白行列
4. 统计指定范围内的单元格，将非空白（非空字符串或全空格）单元格计为数据单元格
5. 计算密度为数据单元格数量除以总单元格数量

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

## 开发约定

- 使用 2024 版本的 Rust 语言标准
- 代码应遵循 Rust 的最佳实践和命名约定
- 测试代码应放在 `#[cfg(test)]` 模块中
- 使用 `cargo fmt` 保持代码格式一致
- 使用 `cargo clippy` 保持代码质量