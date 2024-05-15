use calamine::{open_workbook, Reader, Xlsx};
use std::env;
use std::io::{self, Write};
use xlsxwriter::{Workbook, Format, FormatAlignment};

fn main() {

    // transform this into a crate
    // need to create an app file for example that receives the params
    // if it is the cli it will receive them through main else it will be an exported library that calls the app fn directly

    let args: Vec<String> = env::args().collect();
    
    // the first arg should be the current excel catalogue and the second one the new one
    let current_excel_path = match args.first() {
        Some(arg) => arg,
        None => {
            println!("Please provide the current excel file path");
            match io::stdout().flush() {
                Ok(_) => return,
                Err(e) => panic!("Error: {}", e),
            }
        }
    };
    
    let new_excel_path = match args.last() {
        Some(arg) => arg,
        None => {
            println!("Please provide the new excel file path");
            match io::stdout().flush() {
                Ok(_) => return,
                Err(e) => panic!("Error: {}", e),
            }
        }
    };
    
    // Read both of the paths provided in the args
    let mut current_excel_data: Xlsx<_> = match open_workbook(current_excel_path) {
        Ok(excel) => excel,
        Err(e) => panic!("Couldn't open the excel in the provided path {}, Error: {}", current_excel_path, e),
    };
    
    let mut new_excel_data: Xlsx<_> = match open_workbook(new_excel_path) {
        Ok(excel) => excel,
        Err(e) => panic!("Couldn't open the excel in the provided path {}, Error: {}", new_excel_path, e),
    };
    
    // TODO: Need to delete the new excel if anything fails
    // also need to check what happens if we already have a current ouput excel from last execution
    let workbook = Workbook::new("output.xlsx");
    // loop through each sheet
    // loop through each row+col combination and compare with the value in the second input arg excel read
    for sheet in current_excel_data.sheet_names().to_owned() {
        if let Some(Ok(r)) = current_excel_data.worksheet_range(&sheet) {
            // need to iterate with indexes
            for (n, row) in r.rows().enumerate() {
                for (m, row_col_value) in row.iter().enumerate() {
                    // Compare value of col with the new excel col value
                    // fix both unwraps below 
                    // first is for a possible unavailable sheet the second one is for a possible unavailable row
                    let new_excel_row_col_value =
                    new_excel_data.worksheet_range(&sheet).unwrap().unwrap()[n][m].clone();
                    println!("Searching through sheet: {}, Row: {}, Col: {}", sheet, n, m);
                    
                    let mut worksheet = workbook.add_worksheet(None).unwrap();

                    if *row_col_value == new_excel_row_col_value {
                        // write to a new ouptut excel file


                        println!("Row: {}, Col: {} is the same in both excel files", n, m);
                    } else {
                        println!("Current excel has value: {}, New excel has value: {}", row_col_value, new_excel_row_col_value);
                        // ask user to stay with value current or new
                        println!("Do you want to keep the current value (Y/N)?");
                        
                        let mut input = String::new();

                        // fix .expect()
                        io::stdin().read_line(&mut input).expect("Failed to read line");
                        let input = input.trim().to_uppercase(); // Remove the newline
                        
                        if input == "Y" {
                            // Keep the current value
                            // ...
                        } else if input == "N" {
                            // Use the new value
                            // ...
                        } else {
                            println!("Invalid input, please enter 'Y' or 'N'");
                        }
                    }
                }
            }
            
        }
    }
}



for (n, row) in r.rows().enumerate() {
    for (m, row_col_value) in row.iter().enumerate() {
        // Write the value to the new file
        worksheet.write_string(n as u16, m as u16, row_col_value, None).unwrap();
    }
}

workbook.close().unwrap();