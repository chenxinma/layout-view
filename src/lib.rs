use calamine::{Reader, Xlsx, open_workbook};
use serde::{Deserialize, Serialize};

#[derive(Serialize, Deserialize, Debug)]
pub struct SheetDataDensity {
    pub sheet_name: String,
    pub first_row: u32,
    pub first_col: u32,
    pub end_row: u32,
    pub end_col: u32,
    pub total_cells: u32,
    pub data_cells: u32,
    pub density: f64,
    pub visible: String,  // 添加可见性字段
    pub first_row_first_col_content: Option<String>,  // 采样第一行第一列cell内容
    pub last_row_first_col_content: Option<String>,   // 采样最后一行第一列cell内容
}

pub fn calculate_sheet_density(
    xlsx_path: &str,
) -> Result<Vec<SheetDataDensity>, Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook(xlsx_path)?;

    let mut results = Vec::new();

    // 获取所有工作表的元数据（包含可见性信息）
    let sheet_metadata: std::collections::HashMap<String, calamine::SheetVisible> = 
        workbook.sheets_metadata()
            .iter()
            .map(|sheet| (sheet.name.clone(), sheet.visible))
            .collect();
    
    for (sheet_name, range) in workbook.worksheets() {
        // 检查工作表是否可见
        let visible_status = sheet_metadata.get(&sheet_name)
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
            });
            continue;
        }

        // 计算范围内总单元格数和数据单元格数
        let total_cells = ((end_row - start_row + 1) as u32) * ((end_col - start_col + 1) as u32);
        let mut data_cells = 0;

        for row in start_row..=end_row {
            for col in start_col..=end_col {
                if let Some(cell) = range.get_value((row as u32, col as u32)) {
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
        let first_row_first_col_content = range.get_value((start_row as u32, start_col as u32))
            .map(|cell| cell.to_string());

        // 获取最后一行第一列的cell内容
        let last_row_first_col_content = range.get_value((end_row as u32, start_col as u32))
            .map(|cell| cell.to_string());

        results.push(SheetDataDensity {
            sheet_name: sheet_name.clone(),
            first_row: start_row as u32,
            first_col: start_col as u32,
            end_row: end_row as u32,
            end_col: end_col as u32,
            total_cells,
            data_cells,
            density,
            visible: format!("{:?}", visible_status), // 记录可见性状态
            first_row_first_col_content,
            last_row_first_col_content,
        });
    }

    Ok(results)
}

fn get_effective_range(range: &calamine::Range<calamine::Data>) -> (u32, u32, u32, u32) {
    let Some((start_row, start_col)) = range.start() else { return (0,0,0,0) };
    let Some((end_row, end_col)) = range.end() else { return (0,0,0,0) };

    return (start_row, start_col, end_row, end_col)
}

fn is_empty_cell(cell: &calamine::Data) -> bool {
    match cell {
        calamine::Data::Empty | calamine::Data::Error(_) => true,
        calamine::Data::String(s) => s.trim().is_empty(),
        _ => false,
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
        };
        
        assert_eq!(sheet_data.sheet_name, "Test Sheet");
        assert_eq!(sheet_data.visible, "Visible");
        assert_eq!(sheet_data.first_row_first_col_content, Some("Test".to_string()));
        assert_eq!(sheet_data.last_row_first_col_content, Some("End".to_string()));
    }
}
