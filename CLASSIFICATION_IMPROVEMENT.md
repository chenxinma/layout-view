# 分类算法改进总结

## 问题
表单型Excel文件（test1_form.xlsx）被错误识别为 Data，特征：
- density: 0.598（较高）
- data_type_mix: 0.298（较低）
- 4列，23行，宽高比 5.75

## 解决方案
引入**行间类型一致性**和**宽高比**两个新特征：

### 1. 行间类型一致性（row_type_consistency）
计算每行的数值型数据占比，然后计算这些占比的标准差，转换为0-1的一致性分数：
- **值越高**：行间数据类型模式越相似 → 可能是数据表
- **值越低**：行间差异大 → 可能是表单

### 2. 宽高比（aspect_ratio）
计算行数/列数：
- **值高（>4）且列数少**：垂直排列的表单（键值对结构）
- **值低（<1）且列数多**：扁平的数据表

## 新分类逻辑

```rust
if density > 0.70 {
    // 极高密度 → 数据表
    Data
} else if aspect_ratio > 4.0 && col_count <= 4 && density > 0.35 {
    // 高瘦结构 + 少列 + 中等密度 → 表单
    Form
} else if density > 0.46 && col_count > 4 && row_consistency > 0.50 {
    // 高密度 + 多列 + 行一致性中等 → 数据表
    Data
} else if density > 0.40 && row_consistency > 0.80 {
    // 密度中等 + 高度行一致 → 数据表
    Data
} else if density > 0.46 && data_type_mix > 0.35 {
    // 原逻辑：高密度 + 高混合度 → 数据表
    Data
} else if density <= 0.46 && data_type_mix > 0.35 {
    // 原逻辑：低密度 + 高混合度 → 数据表
    Data
} else {
    // 其他 → 表单
    Form
}
```

## 测试结果

### test1_form.xlsx（表单）
- 之前：`"sheet_type": "Data"` ❌
- 之后：`"sheet_type": "Form"` ✅
- 特征：density: 0.598, aspect_ratio: 5.75, col_count: 4

### test2_data.xlsx（数据表）
- 之前：`"sheet_type": "Data"` ✅
- 之后：`"sheet_type": "Data"` ✅
- 特征：density: 0.603, aspect_ratio: 0.39, col_count: 18

## 数据结构变化

### SheetDataDensity 新增字段
- `row_type_consistency: f64` - 行间类型一致性
- `aspect_ratio: f64` - 宽高比

### ClassifiedSheet 新增字段
- `row_type_consistency: f64` - 行间类型一致性
- `aspect_ratio: f64` - 宽高比

### classification_reason 更新
现在包含4个指标：
```
"classification_reason": "density: 0.598, data_type_mix: 0.298, row_consistency: 0.727, aspect_ratio: 5.8"
```
