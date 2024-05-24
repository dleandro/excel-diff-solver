use calamine::{open_workbook, Reader, Xlsx};
use std::env;
use std::io::{self};
use xlsxwriter::Workbook;

fn main() {
    // transform this into a crate
    // need to create an app file for example that receives the params
    // if it is the cli it will receive them through main else it will be an exported library that calls the app fn directly

    let args: Vec<String> = env::args().collect();

    let current_excel_path = args
        .first()
        .expect("Please provide the current excel file path");

    let new_excel_path = args.last().expect("Please provide the new excel file path");

    let mut current_excel_data: Xlsx<_> =
        open_workbook(current_excel_path).expect("Couldn't open the excel in the provided path");

    let mut new_excel_data: Xlsx<_> =
        open_workbook(new_excel_path).expect("Couldn't open the excel in the provided path");

    // TODO: Need to delete the new excel if anything fails
    // also need to check what happens if we already have a current ouput excel from last execution
    // we should prob close the workbook and any other open file
    let workbook = Workbook::new("output.xlsx").expect("Couldn't create the output excel file");
    // loop through each sheet
    // loop through each row+col combination and compare with the value in the second input arg excel read
    for sheet in current_excel_data.sheet_names().to_owned() {
        let mut worksheet = workbook
            .add_worksheet(Some(&sheet))
            .expect("Couldn't add a new worksheet");
        if let Some(Ok(r)) = current_excel_data.worksheet_range(&sheet) {
            for (n, row) in r.rows().enumerate() {
                for (m, row_col_value) in row.iter().enumerate() {
                    // Compare value of col with the new excel col value
                    // fix both unwraps below
                    // first is for a possible unavailable sheet the second one is for a possible unavailable row
                    let new_excel_row_col_value =
                        new_excel_data.worksheet_range(&sheet).unwrap().unwrap()[n][m].clone();
                    println!("Searching through sheet: {}, Row: {}, Col: {}", sheet, n, m);

                    if *row_col_value == new_excel_row_col_value {
                        println!("Row: {}, Col: {} is the same in both excel files writing to the output file...", n, m);
                        worksheet
                            .write_string(
                                n.try_into().unwrap(),
                                m.try_into().unwrap(),
                                &row_col_value.to_string(),
                                None,
                            )
                            .unwrap();
                    } else {
                        println!(
                            "Current excel has value: {}, New excel has value: {}",
                            row_col_value, new_excel_row_col_value
                        );
                        // ask user to stay with value current or new
                        println!("Do you want to keep the current value (Y/N)?");
                        let mut input = String::new();

                        io::stdin()
                            .read_line(&mut input)
                            .expect("Failed to read line");
                        let input = input.trim().to_uppercase();

                        if input == "Y" {
                            println!("Keeping the current value: {}", row_col_value);
                            worksheet
                                .write_string(
                                    n.try_into().unwrap(),
                                    m.try_into().unwrap(),
                                    &row_col_value.to_string(),
                                    None,
                                )
                                .unwrap();
                        } else if input == "N" {
                            println!("Keeping the new value: {}", new_excel_row_col_value);
                            worksheet
                                .write_string(
                                    n.try_into().unwrap(),
                                    m.try_into().unwrap(),
                                    &new_excel_row_col_value.to_string(),
                                    None,
                                )
                                .unwrap();
                        } else {
                            println!("Invalid input, please enter 'Y' or 'N'");
                        }
                    }
                }
            }
        }
    }
}
