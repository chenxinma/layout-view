# layout-view 项目文档

## 项目概述

layout-view 是一个使用 Rust 编写的库项目，基于 2024 版本的 Rust。该项目分析Excel文件，计算每个工作表的数据密度，以帮助判别Sheet类型是行列数据表还是一份填报表单。

项目目前依赖于 `calamine` 库（版本 0.32.0）以及 `serde` 和 `serde_json` 库。

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

## 依赖项

- `calamine = "0.32.0"` - 用于读取 Excel 文件的库
- `serde = { version = "1.0", features = ["derive"] }` - 用于序列化
- `serde_json = "1.0"` - 用于JSON处理

## 功能说明

程序接收一个XLSX文件路径作为输入，输出JSON格式的数据密度信息，包括：
- 每个工作表的名称
- 数据范围（first_row, first_col, end_row, end_col）
- 总单元格数和数据单元格数
- 数据密度（数据单元格数/总单元格数）

## 算法说明

1. 读取Excel文件并遍历每个工作表
2. 确定有效数据范围，排除起始的连续空白行列
3. 统计指定范围内的单元格，将非空白（非空字符串或全空格）单元格计为数据单元格
4. 计算密度为数据单元格数量除以总单元格数量

## 开发约定

- 使用 2024 版本的 Rust 语言标准
- 代码应遵循 Rust 的最佳实践和命名约定
- 测试代码应放在 `#[cfg(test)]` 模块中
- 使用 `cargo fmt` 保持代码格式一致
- 使用 `cargo clippy` 保持代码质量


