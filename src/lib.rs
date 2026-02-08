use calamine::{open_workbook, Reader, Xlsx};
use lazy_static::lazy_static;
use regex::Regex;
use serde::{Deserialize, Serialize};
use std::ffi::{CStr, CString};
use std::os::raw::c_char;
// use libc;

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct SheetDataDensity {
    pub sheet_name: String,
    pub first_row: u32,
    pub first_col: u32,
    pub end_row: u32,
    pub end_col: u32,
    pub total_cells: u32,
    pub data_cells: u32,
    pub density: f64,
    pub visible: String,                             // 添加可见性字段
    pub first_row_first_col_content: Option<String>, // 采样第一行第一列cell内容
    pub last_row_first_col_content: Option<String>,  // 采样最后一行第一列cell内容
    pub data_type_mix: f64,                          // 数据类型混合程度
    pub column_data_types: Vec<ColumnDataTypeInfo>,  // 每列的数据类型信息
    pub row_type_consistency: f64,                   // 行间类型一致性（0-1，越高越一致）
    pub aspect_ratio: f64,                           // 宽高比（行数/列数）
}

#[derive(Serialize, Deserialize, Debug, Clone)]
pub struct ColumnDataTypeInfo {
    pub column_index: u32,
    pub numeric_count: u32,
    pub text_count: u32,
    pub total_count: u32,
    pub numeric_type_ratio: f64, // 数值型数据占比
}

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
pub enum SheetType {
    Data,    // 行列表
    Form,    // 表单
    Unknown, // 无法确定
}

#[derive(Serialize, Deserialize, Debug)]
pub struct ClassifiedSheet {
    pub sheet_name: String,
    pub first_row: u32,
    pub first_col: u32,
    pub end_row: u32,
    pub end_col: u32,
    pub total_cells: u32,
    pub data_cells: u32,
    pub density: f64,
    pub visible: String,                             // 添加可见性字段
    pub first_row_first_col_content: Option<String>, // 采样第一行第一列cell内容
    pub last_row_first_col_content: Option<String>,  // 采样最后一行第一列cell内容
    pub data_type_mix: f64,                          // 数据类型混合程度
    pub column_data_types: Vec<ColumnDataTypeInfo>,  // 每列的数据类型信息
    pub row_type_consistency: f64,                   // 行间类型一致性（0-1，越高越一致）
    pub aspect_ratio: f64,                           // 宽高比（行数/列数）
    pub sheet_type: SheetType,
    pub classification_reason: String, // 分类原因说明
}

pub fn calculate_sheet_density(
    xlsx_path: &str,
) -> Result<Vec<SheetDataDensity>, Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook(xlsx_path)?;

    let mut results = Vec::new();

    // 获取所有工作表的元数据（包含可见性信息）
    let sheet_metadata: std::collections::HashMap<String, calamine::SheetVisible> = workbook
        .sheets_metadata()
        .iter()
        .map(|sheet| (sheet.name.clone(), sheet.visible))
        .collect();

    for (sheet_name, range) in workbook.worksheets() {
        // 检查工作表是否可见
        let visible_status = sheet_metadata
            .get(&sheet_name)
            .copied()
            .unwrap_or(calamine::SheetVisible::Visible);

        // 只处理可见的工作表
        if visible_status != calamine::SheetVisible::Visible {
            continue; // 跳过隐藏和非常隐藏的工作表
        }

        // 获取数据范围
        let (start_row, start_col, end_row, end_col) = get_effective_range(&range);

        if start_row > end_row || start_col > end_col {
            // 空工作表
            results.push(SheetDataDensity {
                sheet_name: sheet_name.clone(),
                first_row: 0,
                first_col: 0,
                end_row: 0,
                end_col: 0,
                total_cells: 0,
                data_cells: 0,
                density: 0.0,
                visible: format!("{:?}", visible_status), // 记录可见性状态
                first_row_first_col_content: None,
                last_row_first_col_content: None,
                data_type_mix: 0.0,
                column_data_types: Vec::new(),
                row_type_consistency: 0.0,
                aspect_ratio: 0.0,
            });
            continue;
        }

        // 限制分析前100行
        let sample_end_row = std::cmp::min(end_row, start_row + 99); // 最多100行 (0-99)

        // 计算范围内总单元格数和数据单元格数
        let total_cells = (sample_end_row - start_row + 1) * (end_col - start_col + 1);
        let mut data_cells = 0;

        for row in start_row..=sample_end_row {
            for col in start_col..=end_col {
                if let Some(cell) = range.get_value((row, col)) {
                    // 检查是否为非空数据（非空白或全空格）
                    if !is_empty_cell(cell) {
                        data_cells += 1;
                    }
                }
            }
        }

        let density = if total_cells > 0 {
            data_cells as f64 / total_cells as f64
        } else {
            0.0
        };

        // 获取第一行第一列的cell内容
        let first_row_first_col_content = range
            .get_value((start_row, start_col))
            .map(|cell| cell.to_string());

        // 获取最后一行第一列的cell内容
        let last_row_first_col_content = range
            .get_value((sample_end_row, start_col))
            .map(|cell| cell.to_string());

        // 计算每列的数据类型分布
        let column_data_types =
            calculate_column_data_types(&range, start_row, sample_end_row, start_col, end_col);

        // 计算数据类型混合程度
        let data_type_mix = calculate_data_type_mix(&column_data_types);

        // 计算行间类型一致性
        let row_type_consistency =
            calculate_row_type_consistency(&range, start_row, sample_end_row, start_col, end_col);

        // 计算宽高比
        let row_count = (sample_end_row - start_row + 1) as f64;
        let col_count = (end_col - start_col + 1) as f64;
        let aspect_ratio = if col_count > 0.0 {
            row_count / col_count
        } else {
            0.0
        };

        results.push(SheetDataDensity {
            sheet_name: sheet_name.clone(),
            first_row: start_row,
            first_col: start_col,
            end_row: end_row, // 保留原始end_row
            end_col: end_col, // 保留原始end_col
            total_cells,
            data_cells,
            density,
            visible: format!("{:?}", visible_status), // 记录可见性状态
            first_row_first_col_content,
            last_row_first_col_content,
            data_type_mix,
            column_data_types,
            row_type_consistency,
            aspect_ratio,
        });
    }

    Ok(results)
}

fn get_effective_range(range: &calamine::Range<calamine::Data>) -> (u32, u32, u32, u32) {
    let Some((start_row, start_col)) = range.start() else {
        return (0, 0, 0, 0);
    };
    let Some((end_row, end_col)) = range.end() else {
        return (0, 0, 0, 0);
    };

    (start_row, start_col, end_row, end_col)
}

fn is_empty_cell(cell: &calamine::Data) -> bool {
    match cell {
        calamine::Data::Empty | calamine::Data::Error(_) => true,
        calamine::Data::String(s) => s.trim().is_empty(),
        _ => false,
    }
}

// 静态正则表达式，用于匹配数值型数据
lazy_static! {
    static ref PERCENTAGE_RE: Regex = Regex::new(r#"^-?\d*\.?\d+%$"#).unwrap();
    static ref THOUSANDS_RE: Regex = Regex::new(r#"^-?\d{1,3}(,\d{3})*(\.\d+)?$"#).unwrap();
    static ref NUMBER_RE: Regex = Regex::new(r#"^-?\d+\.?\d*$"#).unwrap();
}

/// 检查字符串是否为数值型数据
/// 支持：整数、小数、含千分位数、百分数
fn is_numeric_string(s: &str) -> bool {
    let s_trimmed = s.trim();
    if s_trimmed.is_empty() {
        return false;
    }

    // 检查是否为百分数 (例如: 50%, -25.5%)
    if PERCENTAGE_RE.is_match(s_trimmed) {
        return true;
    }

    // 检查是否为含千分位的数字 (例如: 1,234.56, -1,234.56)
    if THOUSANDS_RE.is_match(s_trimmed) {
        return true;
    }

    // 检查是否为普通小数或整数 (例如: 123.45, -67.89, 123, -45)
    if NUMBER_RE.is_match(s_trimmed) {
        return true;
    }

    false
}

/// 检查单元格是否包含数值型数据
fn is_numeric_cell(cell: &calamine::Data) -> bool {
    match cell {
        calamine::Data::Int(_) | calamine::Data::Float(_) => true,
        calamine::Data::String(s) => is_numeric_string(s),
        _ => false,
    }
}

/// 计算每列的数据类型分布
fn calculate_column_data_types(
    range: &calamine::Range<calamine::Data>,
    start_row: u32,
    end_row: u32,
    start_col: u32,
    end_col: u32,
) -> Vec<ColumnDataTypeInfo> {
    let mut column_info = Vec::new();

    for col in start_col..=end_col {
        let mut numeric_count = 0;
        let mut text_count = 0;
        let mut total_count = 0;

        for row in start_row..=end_row {
            if let Some(cell) = range.get_value((row, col)) {
                if !is_empty_cell(cell) {
                    total_count += 1;
                    if is_numeric_cell(cell) {
                        numeric_count += 1;
                    } else {
                        // 将非数值型数据视为文本型数据
                        text_count += 1;
                    }
                }
            }
        }

        let numeric_type_ratio = if total_count > 0 {
            numeric_count as f64 / total_count as f64
        } else {
            0.0
        };

        column_info.push(ColumnDataTypeInfo {
            column_index: col,
            numeric_count,
            text_count,
            total_count,
            numeric_type_ratio,
        });
    }

    column_info
}

/// 使用多样性指数（香农熵）计算数据类型混合程度
fn calculate_data_type_mix(column_data_types: &[ColumnDataTypeInfo]) -> f64 {
    if column_data_types.is_empty() {
        return 0.0;
    }

    let mut total_mix = 0.0;
    let mut valid_columns = 0;

    for col_info in column_data_types {
        if col_info.total_count > 0 {
            // 计算香农熵
            let p_numeric = col_info.numeric_count as f64 / col_info.total_count as f64;
            let p_text = col_info.text_count as f64 / col_info.total_count as f64;

            let mut entropy = 0.0;
            if p_numeric > 0.0 {
                entropy -= p_numeric * p_numeric.ln();
            }
            if p_text > 0.0 {
                entropy -= p_text * p_text.ln();
            }

            // 标准化熵值 (除以 ln(2) 以确保最大值为 1)
            let normalized_entropy = if entropy > 0.0 {
                entropy / 2.0f64.ln()
            } else {
                0.0
            };

            total_mix += normalized_entropy;
            valid_columns += 1;
        }
    }

    if valid_columns > 0 {
        total_mix / valid_columns as f64
    } else {
        0.0
    }
}

/// 计算行间类型一致性
/// 通过计算每行的数值型占比，然后统计这些占比的标准差
/// 标准差越小，表示行间类型模式越一致（数据表特征）
fn calculate_row_type_consistency(
    range: &calamine::Range<calamine::Data>,
    start_row: u32,
    end_row: u32,
    start_col: u32,
    end_col: u32,
) -> f64 {
    if start_row > end_row {
        return 0.0;
    }

    let mut row_numeric_ratios = Vec::new();

    for row in start_row..=end_row {
        let mut numeric_count = 0;
        let mut total_count = 0;

        for col in start_col..=end_col {
            if let Some(cell) = range.get_value((row, col)) {
                if !is_empty_cell(cell) {
                    total_count += 1;
                    if is_numeric_cell(cell) {
                        numeric_count += 1;
                    }
                }
            }
        }

        if total_count > 0 {
            let ratio = numeric_count as f64 / total_count as f64;
            row_numeric_ratios.push(ratio);
        }
    }

    if row_numeric_ratios.len() < 2 {
        return 0.0;
    }

    // 计算标准差
    let mean = row_numeric_ratios.iter().sum::<f64>() / row_numeric_ratios.len() as f64;
    let variance = row_numeric_ratios
        .iter()
        .map(|&x| (x - mean).powi(2))
        .sum::<f64>()
        / row_numeric_ratios.len() as f64;
    let std_dev = variance.sqrt();

    // 将标准差转换为一致性分数（0-1）
    // 标准差越小，一致性越高
    // 使用 sigmoid 函数映射：consistency = 1 / (1 + std_dev * 2)
    let consistency = 1.0 / (1.0 + std_dev * 3.0);
    consistency
}

/// 根据密度、数据类型混合度、行间类型一致性和宽高比对工作表进行分类
pub fn classify_sheet(sheet_data: &SheetDataDensity) -> ClassifiedSheet {
    // 忽略density=0的sheet
    if sheet_data.density == 0.0 {
        return ClassifiedSheet {
            sheet_name: sheet_data.sheet_name.clone(),
            first_row: sheet_data.first_row,
            first_col: sheet_data.first_col,
            end_row: sheet_data.end_row,
            end_col: sheet_data.end_col,
            total_cells: sheet_data.total_cells,
            data_cells: sheet_data.data_cells,
            density: sheet_data.density,
            visible: sheet_data.visible.clone(),
            first_row_first_col_content: sheet_data.first_row_first_col_content.clone(),
            last_row_first_col_content: sheet_data.last_row_first_col_content.clone(),
            data_type_mix: sheet_data.data_type_mix,
            column_data_types: sheet_data.column_data_types.clone(),
            row_type_consistency: sheet_data.row_type_consistency,
            aspect_ratio: sheet_data.aspect_ratio,
            sheet_type: SheetType::Unknown,
            classification_reason: "Density is zero".to_string(),
        };
    }

    // 综合分类逻辑
    // 1. 极高密度 -> 数据表
    // 2. 高密度 + 高行一致性 + 列数较多 -> 数据表
    // 3. 宽高比 > 5 + 列数 <= 4 -> 表单（垂直排列的键值对表单）
    // 4. 宽表特征（多列+扁平）-> 数据表
    // 5. 低密度 + 高混合度 + 列数较多 -> 数据表
    // 6. 其他情况根据混合度判断
    let col_count = sheet_data.end_col - sheet_data.first_col + 1;

    let sheet_type = if sheet_data.density > 0.70 {
        // 极高密度几乎肯定是数据表
        SheetType::Data
    } else if sheet_data.aspect_ratio > 4.0 && col_count <= 4 && sheet_data.density > 0.35 {
        // 高瘦结构 + 少列 + 中等密度 -> 表单（垂直键值对表单）
        SheetType::Form
    } else if sheet_data.density > 0.46 && col_count > 4 && sheet_data.row_type_consistency > 0.50 {
        // 高密度、多列、行一致性中等以上 -> 数据表
        SheetType::Data
    } else if sheet_data.density > 0.40 && sheet_data.row_type_consistency > 0.80 {
        // 行结构高度一致且密度中等以上 -> 数据表（需要更高的一致性阈值）
        SheetType::Data
    } else if col_count > 10 && sheet_data.aspect_ratio < 0.5 && sheet_data.density > 0.35 {
        // 宽表特征：列数多(>10) + 宽高比低(<0.5) + 密度中等以上 -> 数据表
        SheetType::Data
    } else if sheet_data.density > 0.46 && sheet_data.data_type_mix > 0.35 {
        // 原逻辑：高密度且数据类型混合度高
        SheetType::Data
    } else if sheet_data.density > 0.35 && sheet_data.data_type_mix > 0.35 && col_count > 5 {
        // 密度中等 + 数据类型混合度高 + 列数较多 -> 数据表
        SheetType::Data
    } else {
        // 其他情况视为表单
        SheetType::Form
    };

    let reason = format!(
        "density: {:.3}, data_type_mix: {:.3}, row_consistency: {:.3}, aspect_ratio: {:.1}",
        sheet_data.density,
        sheet_data.data_type_mix,
        sheet_data.row_type_consistency,
        sheet_data.aspect_ratio
    );

    ClassifiedSheet {
        sheet_name: sheet_data.sheet_name.clone(),
        first_row: sheet_data.first_row,
        first_col: sheet_data.first_col,
        end_row: sheet_data.end_row,
        end_col: sheet_data.end_col,
        total_cells: sheet_data.total_cells,
        data_cells: sheet_data.data_cells,
        density: sheet_data.density,
        visible: sheet_data.visible.clone(),
        first_row_first_col_content: sheet_data.first_row_first_col_content.clone(),
        last_row_first_col_content: sheet_data.last_row_first_col_content.clone(),
        data_type_mix: sheet_data.data_type_mix,
        column_data_types: sheet_data.column_data_types.clone(),
        row_type_consistency: sheet_data.row_type_consistency,
        aspect_ratio: sheet_data.aspect_ratio,
        sheet_type,
        classification_reason: reason,
    }
}

/// 对整个Excel文件的所有工作表进行分类（忽略density=0的sheet）
pub fn classify_excel_sheets(
    xlsx_path: &str,
) -> Result<Vec<ClassifiedSheet>, Box<dyn std::error::Error>> {
    let sheets = calculate_sheet_density(xlsx_path)?;

    // 对每个sheet进行分类，忽略density=0的sheet
    let classified_sheets: Vec<ClassifiedSheet> = sheets
        .into_iter()
        .map(|sheet| classify_sheet(&sheet))
        .filter(|classified| classified.density != 0.0) // 过滤掉密度为0的sheet
        .collect();

    Ok(classified_sheets)
}

// C FFI functions for use as a dynamic library from Python

/// C function to classify Excel sheets and return results as JSON string
/// The caller is responsible for freeing the returned string using free_c_string
#[no_mangle]
pub unsafe extern "C" fn classify_excel_sheets_c(xlsx_path: *const c_char) -> *mut c_char {
    if xlsx_path.is_null() {
        return std::ptr::null_mut();
    }

    let path_c_str = unsafe {
        match CStr::from_ptr(xlsx_path).to_str() {
            Ok(s) => s,
            Err(_) => return std::ptr::null_mut(),
        }
    };

    match classify_excel_sheets(path_c_str) {
        Ok(results) => match serde_json::to_string(&results) {
            Ok(json_string) => match CString::new(json_string) {
                Ok(c_string) => c_string.into_raw(),
                Err(_) => std::ptr::null_mut(),
            },
            Err(_) => std::ptr::null_mut(),
        },
        Err(_) => std::ptr::null_mut(),
    }
}

/// C function to free strings allocated by Rust
#[no_mangle]
pub unsafe extern "C" fn free_c_string(ptr: *mut c_char) {
    if !ptr.is_null() {
        let _ = unsafe { CString::from_raw(ptr) };
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_is_empty_cell() {
        use calamine::Data;

        assert!(is_empty_cell(&Data::Empty));
        assert!(is_empty_cell(&Data::String("".to_string())));
        assert!(is_empty_cell(&Data::String("   ".to_string())));
        assert!(!is_empty_cell(&Data::String("data".to_string())));
        assert!(!is_empty_cell(&Data::Int(42)));
    }

    #[test]
    fn test_sheet_data_density_struct() {
        // 测试SheetDataDensity结构体是否包含所有字段
        let sheet_data = SheetDataDensity {
            sheet_name: "Test Sheet".to_string(),
            first_row: 0,
            first_col: 0,
            end_row: 10,
            end_col: 10,
            total_cells: 121,
            data_cells: 50,
            density: 0.5,
            visible: "Visible".to_string(),
            first_row_first_col_content: Some("Test".to_string()),
            last_row_first_col_content: Some("End".to_string()),
            data_type_mix: 0.5,
            column_data_types: vec![ColumnDataTypeInfo {
                column_index: 0,
                numeric_count: 5,
                text_count: 5,
                total_count: 10,
                numeric_type_ratio: 0.5,
            }],
            row_type_consistency: 0.7,
            aspect_ratio: 1.0,
        };

        assert_eq!(sheet_data.sheet_name, "Test Sheet");
        assert_eq!(sheet_data.visible, "Visible");
        assert_eq!(
            sheet_data.first_row_first_col_content,
            Some("Test".to_string())
        );
        assert_eq!(
            sheet_data.last_row_first_col_content,
            Some("End".to_string())
        );
        assert_eq!(sheet_data.data_type_mix, 0.5);
        assert_eq!(sheet_data.column_data_types.len(), 1);
        assert_eq!(sheet_data.column_data_types[0].column_index, 0);
    }

    #[test]
    fn test_is_numeric_string() {
        // 测试正则表达式对数值型数据的识别
        assert!(is_numeric_string("123"));
        assert!(is_numeric_string("-45"));
        assert!(is_numeric_string("3.14"));
        assert!(is_numeric_string("-2.5"));
        assert!(is_numeric_string("1,234.56"));
        assert!(is_numeric_string("-1,234.56"));
        assert!(is_numeric_string("50%"));
        assert!(is_numeric_string("-25.5%"));

        // 测试非数值型数据
        assert!(!is_numeric_string("text"));
        assert!(!is_numeric_string(""));
        assert!(!is_numeric_string("   "));
        assert!(!is_numeric_string("12.34.56"));
        assert!(!is_numeric_string("abc123"));
    }

    #[test]
    fn test_calculate_data_type_mix() {
        // 创建测试数据：包含混合类型的列（数值和文本各占一半）
        let column_data_types = vec![
            ColumnDataTypeInfo {
                column_index: 0,
                numeric_count: 3,
                text_count: 3,
                total_count: 6,
                numeric_type_ratio: 0.5,
            },
            ColumnDataTypeInfo {
                column_index: 1,
                numeric_count: 2,
                text_count: 2,
                total_count: 4,
                numeric_type_ratio: 0.5,
            },
        ];

        // 混合的数据应该有较高的熵值
        let mix = calculate_data_type_mix(&column_data_types);
        // 由于两列都包含混合数据，混合度应该较高
        assert!(mix > 0.6, "混合程度应该较高，当前值为: {}", mix);

        // 创建测试数据：完全一致的列（全数值）
        let column_data_types = vec![
            ColumnDataTypeInfo {
                column_index: 0,
                numeric_count: 5,
                text_count: 0,
                total_count: 5,
                numeric_type_ratio: 1.0,
            },
            ColumnDataTypeInfo {
                column_index: 1,
                numeric_count: 5,
                text_count: 0,
                total_count: 5,
                numeric_type_ratio: 1.0,
            },
        ];

        // 完全一致的数据应该有低混合程度
        let mix = calculate_data_type_mix(&column_data_types);
        assert!(mix < 0.1, "混合程度应该较低，当前值为: {}", mix);
    }
}
