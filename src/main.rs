use app::merge_excel_files;
use std::{env, path::PathBuf};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();

    let current_excel_path = args.get(1)
        .map(PathBuf::from)
        .ok_or("Please provide the current excel file path")?;

    let new_excel_path = args.get(2)
        .map(PathBuf::from)
        .ok_or("Please provide the new excel file path")?;


    let output_excel_path = args.get(3).map_or_else(
        || {
            let mut path = env::current_dir().map_err(|_| "Failed to get current directory")?;
            path.push("output.xlsx");
            Ok::<PathBuf, &str>(path)
        },
        |path| Ok(PathBuf::from(path)),
    )?;

    merge_excel_files(current_excel_path, new_excel_path, output_excel_path)
        .map(|_| println!("Successfully merged the excel files"))
        .unwrap_or_else(|e| eprintln!("Error: {}", e));

    Ok(())
}
