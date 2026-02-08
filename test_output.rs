use layout_view::{classify_excel_sheets, SheetType};

fn main() {
    // 示例：展示修改后的 ClassifiedSheet 结构如何生成扁平化的 JSON
    println!("ClassifiedSheet 结构现在直接包含所有 SheetDataDensity 的字段，");
    println!("这样生成的 JSON 将具有更扁平的结构：");
    println!(
        r#"
{
  "sheet_name": "...",
  "first_row": 0,
  "first_col": 0,
  "end_row": 10,
  "end_col": 5,
  "total_cells": 55,
  "data_cells": 40,
  "density": 0.73,
  "visible": "Visible",
  "first_row_first_col_content": "Title",
  "last_row_first_col_content": "End",
  "data_type_mix": 0.45,
  "column_data_types": [...],
  "sheet_type": "Data",
  "classification_reason": "density: 0.730, data_type_mix: 0.450"
}
    "#
    );
}
