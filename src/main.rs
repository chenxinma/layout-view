use layout_view::calculate_sheet_density;
use std::env;
use std::process;

fn main() {
    let args: Vec<String> = env::args().collect();

    if args.len() != 2 {
        eprintln!("Usage: {} <xlsx_file_path>", args[0]);
        process::exit(1);
    }

    let xlsx_path = &args[1];

    match calculate_sheet_density(xlsx_path) {
        Ok(results) => match serde_json::to_string(&results) {
            Ok(json_output) => println!("{}", json_output),
            Err(e) => {
                eprintln!("Error serializing results to JSON: {}", e);
                process::exit(1);
            }
        },
        Err(e) => {
            eprintln!("Error processing XLSX file: {}", e);
            process::exit(1);
        }
    }
}
