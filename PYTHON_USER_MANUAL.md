# layout-view Python 用户手册

## 项目概述

`layout-view` 是一个功能强大的 Rust 库，能够分析 Excel 文件的工作表布局和数据密度。该项目提供 C 语言接口（FFI），允许 Python 开发者从 Python 代码中调用其功能。

### 主要功能
- 计算 Excel 文件中每个工作表的数据密度
- 识别数据的起始和结束行列位置
- 跳过隐藏和非常隐藏的工作表
- 输出 JSON 格式的分析结果
- 基于密度和数据类型混合度对工作表进行分类（行列表 vs 表单）
- 分析工作表中每列的数据类型分布（数值型 vs 文本型）

## 安装和构建

### 前提条件
- Rust（2024 版本）
- Cargo
- Python 3.x

### 构建 Rust 库
首先，您需要构建 Rust 库来生成动态链接库：

```bash
# 克隆项目（如果尚未克隆）
git clone <repository-url>
cd layout-view

# 构建发布版本的动态库
cargo build --release
```

构建成功后，您将在 `target/release/` 目录中找到动态库文件：
- **Linux**: `liblayout_view.so`
- **macOS**: `liblayout_view.dylib`
- **Windows**: `layout_view.dll`

## Python 使用方法

### 基本使用

```python
import ctypes
import json

# 加载动态库
lib = ctypes.CDLL('./target/release/liblayout_view.so')  # Linux
# lib = ctypes.CDLL('./target/release/liblayout_view.dylib')  # macOS
# lib = ctypes.CDLL('./target/release/layout_view.dll')  # Windows

# 定义函数参数和返回类型
lib.classify_excel_sheets_c.argtypes = [ctypes.c_char_p]
lib.classify_excel_sheets_c.restype = ctypes.POINTER(ctypes.c_char)  # 返回字符指针

lib.free_c_string.argtypes = [ctypes.POINTER(ctypes.c_char)]
lib.free_c_string.restype = None

def classify_excel_sheets(xlsx_path):
    """
    分类 Excel 工作表
    
    Args:
        xlsx_path (str): Excel 文件路径
    
    Returns:
        list: 分类后的工作表列表或 None（如果出错）
    """
    # 将 Python 字符串转换为 C 字符串
    c_path = ctypes.c_char_p(xlsx_path.encode('utf-8'))
    
    # 调用 Rust 函数
    result_ptr = lib.classify_excel_sheets_c(c_path)
    
    if not result_ptr:
        return None
    
    try:
        # 将结果转换为 Python 字符串
        result_bytes = ctypes.cast(result_ptr, ctypes.c_char_p).value
        if result_bytes is None:
            return None
        
        result_str = result_bytes.decode('utf-8')
        
        # 解析 JSON 结果
        parsed_result = json.loads(result_str)
        
        # 释放 Rust 分配的字符串内存
        lib.free_c_string(result_ptr)
        
        return parsed_result
    except json.JSONDecodeError:
        # 即使 JSON 解析失败，仍需释放字符串
        lib.free_c_string(result_ptr)
        return None
    except Exception:
        # 出现其他错误时，释放字符串并重新抛出异常
        lib.free_c_string(result_ptr)
        raise

# 使用示例
xlsx_file_path = "example.xlsx"
results = classify_excel_sheets(xlsx_file_path)

if results:
    print("工作表分类结果:")
    print(json.dumps(results, indent=2, ensure_ascii=False))
else:
    print("分类失败")
```

### 使用项目提供的 Python 封装

项目提供了一个封装好的 Python 文件 `use_layout_view.py`，可以直接使用：

```python
from use_layout_view import classify_excel_sheets
import json

# 使用封装好的函数
xlsx_file_path = "example.xlsx"
results = classify_excel_sheets(xlsx_file_path)

if results:
    print("工作表分类结果:")
    print(json.dumps(results, indent=2, ensure_ascii=False))
else:
    print("分类失败")
```

## 输出格式说明

函数返回一个包含 `ClassifiedSheet` 对象的列表，每个对象包含以下信息：

### ClassifiedSheet 对象
- `original`: 包含原始分析数据的对象
- `sheet_type`: 工作表类型（Data、Form、Unknown）
- `classification_reason`: 分类原因说明

### Original 对象字段
- `sheet_name`: 工作表名称
- `first_row`: 第一个包含数据的行索引
- `first_col`: 第一个包含数据的列索引
- `end_row`: 最后一个包含数据的行索引
- `end_col`: 最后一个包含数据的列索引
- `total_cells`: 指定范围内的总单元格数
- `data_cells`: 包含实际数据的单元格数
- `density`: 数据密度（data_cells / total_cells）
- `visible`: 工作表的可见性状态（"Visible", "Hidden", "VeryHidden"）
- `first_row_first_col_content`: 第一行第一列单元格的内容
- `last_row_first_col_content`: 最后一行第一列单元格的内容
- `data_type_mix`: 数据类型混合程度（使用香农熵计算）
- `column_data_types`: 每列的数据类型信息

### ColumnDataTypeInfo 对象字段
- `column_index`: 列索引
- `numeric_count`: 数值型数据计数
- `text_count`: 文本型数据计数
- `total_count`: 该列总数据计数
- `numeric_type_ratio`: 数值型数据占比

## 工作表分类算法

### 分类原理

`layout-view` 使用两个关键指标对工作表进行分类：

1. **数据密度（Density）**: 数据单元格占总单元格的比例
2. **数据类型混合度（Data Type Mix）**: 使用香农熵计算的数据类型多样性

### 分类规则

- **高密度（>0.46）** → 分类为 "Data"（行列表）
- **低密度（≤0.46）但高数据类型混合度（>0.35）** → 分类为 "Data"（复杂数据表）
- **低密度且低数据类型混合度** → 分类为 "Form"（表单）

### 数据类型识别

算法支持识别以下数值型格式：
- 整数：`123`, `-45`
- 小数：`3.14`, `-2.5`
- 千分位数：`1,234.56`, `-1,234.56`
- 百分数：`50%`, `-25.5%`

### 香农熵计算

数据类型混合度使用香农熵公式计算：
```
H = -∑(p_i * log(p_i))
```

其中 `p_i` 是第 i 种数据类型的概率。结果被标准化以确保最大值为 1。

## 实际应用示例

### 示例 1：分析 Excel 文件并处理结果

```python
from use_layout_view import classify_excel_sheets
import json

def analyze_excel_file(file_path):
    """分析 Excel 文件并显示详细信息"""
    results = classify_excel_sheets(file_path)
    
    if not results:
        print("无法分析文件")
        return
    
    print(f"分析文件: {file_path}")
    print(f"共找到 {len(results)} 个工作表")
    print("-" * 50)
    
    for i, sheet in enumerate(results):
        original = sheet['original']
        print(f"工作表 {i+1}: {original['sheet_name']}")
        print(f"  类型: {sheet['sheet_type']}")
        print(f"  分类原因: {sheet['classification_reason']}")
        print(f"  数据范围: R{original['first_row']}C{original['first_col']} "
              f"to R{original['end_row']}C{original['end_col']}")
        print(f"  数据密度: {original['density']:.3f}")
        print(f"  数据类型混合度: {original['data_type_mix']:.3f}")
        print(f"  总单元格数: {original['total_cells']}")
        print(f"  数据单元格数: {original['data_cells']}")
        print()
        
        # 显示列的数据类型分布
        if original['column_data_types']:
            print("  列数据类型分布:")
            for col_info in original['column_data_types']:
                print(f"    列 {col_info['column_index']}: "
                      f"数值型 {col_info['numeric_count']}, "
                      f"文本型 {col_info['text_count']}, "
                      f"数值占比 {col_info['numeric_type_ratio']:.2f}")
        print("-" * 50)

# 使用示例
analyze_excel_file("example.xlsx")
```

### 示例 2：根据分类结果进行不同处理

```python
from use_layout_view import classify_excel_sheets

def process_excel_by_sheet_type(file_path):
    """根据工作表类型进行不同处理"""
    results = classify_excel_sheets(file_path)
    
    if not results:
        print("无法分析文件")
        return
    
    data_sheets = []
    form_sheets = []
    
    for sheet in results:
        sheet_type = sheet['sheet_type']
        original = sheet['original']
        
        if sheet_type == 'Data':
            data_sheets.append((original['sheet_name'], original))
        elif sheet_type == 'Form':
            form_sheets.append((original['sheet_name'], original))
    
    print(f"数据表 (Data): {len(data_sheets)} 个")
    for name, data in data_sheets:
        print(f"  - {name}: 密度 {data['density']:.3f}, "
              f"混合度 {data['data_type_mix']:.3f}")
    
    print(f"\n表单 (Form): {len(form_sheets)} 个")
    for name, data in form_sheets:
        print(f"  - {name}: 密度 {data['density']:.3f}, "
              f"混合度 {data['data_type_mix']:.3f}")

# 使用示例
process_excel_by_sheet_type("example.xlsx")
```

## 错误处理和最佳实践

### 错误处理

在使用库时，建议进行适当的错误处理：

```python
def safe_classify_excel_sheets(xlsx_path):
    """安全的 Excel 分类函数，包含错误处理"""
    try:
        if not os.path.exists(xlsx_path):
            raise FileNotFoundError(f"文件不存在: {xlsx_path}")
        
        results = classify_excel_sheets(xlsx_path)
        
        if results is None:
            raise RuntimeError("Excel 工作表分类失败")
        
        return results
    except Exception as e:
        print(f"处理 Excel 文件时出错: {e}")
        return None
```

### 最佳实践

1. **内存管理**: 总是确保正确释放由 Rust 分配的内存
2. **文件验证**: 在处理前验证 Excel 文件是否存在且格式正确
3. **结果验证**: 检查返回结果是否为 None 或空列表
4. **异常处理**: 为可能的 JSON 解析错误和系统错误添加异常处理

## 性能考量

- 该库会分析 Excel 文件的前 100 行数据以确定数据密度
- 对于大型 Excel 文件，这提供了良好的性能与准确性的平衡
- 分类算法经过优化，可以快速处理多个工作表

## 常见问题

### Q: 如何更新 Rust 库？
A: 当 Rust 库更新后，需要重新运行 `cargo build --release` 来生成新的动态库。

### Q: 为什么在某些系统上找不到动态库？
A: 确保动态库文件存在于正确的路径，并且文件名与系统架构匹配（Linux: `.so`, macOS: `.dylib`, Windows: `.dll`）。

### Q: 如何处理中文路径？
A: 确保 Python 字符串正确编码为 UTF-8，如：`xlsx_path.encode('utf-8')`。

## 故障排除

如果遇到问题，可以尝试以下步骤：

1. 确认 Rust 库已正确构建：`cargo build --release`
2. 检查动态库文件是否存在于 `target/release/` 目录
3. 验证 ctypes 函数签名是否正确设置
4. 检查 Excel 文件是否可访问且格式正确