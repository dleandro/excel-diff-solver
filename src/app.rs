use calamine::{open_workbook, Reader, Xlsx};
use xlsxwriter::Workbook;
use std::fs::File;
use std::io::{self, BufReader};
use std::path::PathBuf;
use thiserror::Error;

#[derive(Error, Debug)]
pub enum ExcelError {
    #[error("Failed to open workbook: {0}")]
    OpenWorkbookError(String),
    #[error("Failed to create output workbook: {0}")]
    CreateWorkbookError(String),
    #[error("Failed to get worksheet: {0}")]
    GetWorksheetError(String),
    #[error("Failed to write to worksheet: {0}")]
    WriteWorksheetError(String),
    #[error("Failed to delete output file: {0}")]
    DeleteFileError(String),
    #[error("Failed to close workbook: {0}")]
    CloseWorkbookError(String),
    #[error("Failed to copy file: {0}")]
    CopyFileError(String),
}

// need to delete the clones

pub fn merge_excel_files(
    current_excel_path: PathBuf,
    new_excel_path: PathBuf,
    output_excel_path: PathBuf,
) -> Result<(), ExcelError> {
    let mut new_excel_data = open_workbook(&new_excel_path).map_err(|_| {
        ExcelError::OpenWorkbookError(new_excel_path.to_string_lossy().into_owned())
    })?;

    let mut current_excel_data = open_workbook(&current_excel_path).map_err(|_| {
        ExcelError::OpenWorkbookError(current_excel_path.to_string_lossy().into_owned())
    })?;

    let mut new_workbook = Workbook::new(&output_excel_path.to_string_lossy().into_owned()).map_err(|_| {
        ExcelError::CreateWorkbookError(output_excel_path.to_string_lossy().into_owned())
    })?;

    // call method to copy the entire content of the current excel to the output workbook
    copy_current_excel_to_output(&mut current_excel_data, &mut new_workbook)?;
    
    // Perform the comparison and write differences
    // new workbook needs to be connected to the current_excel_data in some way
    merge_new_excel_data(&mut current_excel_data, &mut new_excel_data, &mut new_workbook)?;

    // Ensure all data is written and the file is properly closed
    // do we need this and what does happen if we early exit in the compare call above

    // add back codeguard to close stuff or we could avoid the '?' above and just return the error
    drop(current_excel_data);
    drop(new_excel_data);
    drop(new_workbook);

    Ok(())
}

fn copy_current_excel_to_output(
    current_excel_data: &mut Xlsx<BufReader<File>>,
    workbook: &mut Workbook,
) -> Result<(), ExcelError> {
    for sheet in current_excel_data.sheet_names().to_owned() {
        if let Some(Ok(range)) = current_excel_data.worksheet_range(&sheet) {
            let mut worksheet = workbook.add_worksheet(Some(&sheet)).map_err(|_| ExcelError::WriteWorksheetError(sheet.clone()))?;

            for (n, row) in range.rows().enumerate() {
                for (m, value) in row.iter().enumerate() {
                    worksheet.write_string(n as u32, m as u16, value.to_string().as_str(), None)
                        .map_err(|_| ExcelError::WriteWorksheetError(sheet.clone()))?;
                }
            }
        } else {
            return Err(ExcelError::GetWorksheetError(sheet.clone()));
        }
    }

    Ok(())
}

fn merge_new_excel_data(
    current_excel_data: &mut Xlsx<BufReader<File>>,
    new_excel_data: &mut Xlsx<BufReader<File>>,
    workbook: &mut Workbook,
) -> Result<(), ExcelError> {
    for sheet in new_excel_data.sheet_names().to_owned() {
        let mut worksheet = workbook.get_worksheet(&sheet);

            if worksheet.is_ok() {
                if let Some(Ok(current_range)) = current_excel_data.worksheet_range(&sheet) {
                    if let Some(Ok(new_range)) = new_excel_data.worksheet_range(&sheet) {
                        for (n, row) in new_range.rows().enumerate() {
                            for (m, new_value) in row.iter().enumerate() {
                                let current_value = if n < current_range.height() && m < current_range.width() {
                                    Some(&current_range[n][m])
                                } else {
                                    None
                                };

                                if current_value.is_none() {
                                    println!("Row: {}, Col: {} is out of bounds in the current excel file. Using new value: {}", n, m, new_value);
                                    if let Some(ref mut ws) = worksheet.as_mut().map_err(|_| ExcelError::WriteWorksheetError(sheet.clone()))? {
                                        ws.write_string(n as u32, m as u16, new_value.to_string().as_str(), None)
                                            .map_err(|_| ExcelError::WriteWorksheetError(sheet.clone()))?;
                                    }
                                    continue;
                                }

                                if let Some(current_value) = current_value {
                                    if current_value != new_value {
                                        println!("Row: {}, Col: {} is different in both excel files", n, m);
                                        println!("Current value: {}", current_value);
                                        println!("New value: {}", new_value);
                                        println!("Do you want to keep the current value? 'Y' to keep the current value and 'N' to change to the new value");

                                        loop {
                                            let mut input = String::new();
                                            io::stdin()
                                                .read_line(&mut input)
                                                .map_err(|_| ExcelError::WriteWorksheetError(sheet.clone()))?;
                                            let input = input.trim().to_ascii_uppercase();

                                            match input.chars().next() {
                                                Some('Y') => {
                                                    println!("Keeping the current value: {}", current_value);
                                                    break;
                                                }
                                                Some('N') => {
                                                    println!("Adding the new value: {}", new_value);
                                                    if let Some(ref mut ws) = worksheet.as_mut().map_err(|_| ExcelError::WriteWorksheetError(sheet.clone()))? {
                                                        ws.write_string(n as u32, m as u16, new_value.to_string().as_str(), None)
                                                            .map_err(|_| ExcelError::WriteWorksheetError(sheet.clone()))?;
                                                    }
                                                    break;
                                                }
                                                _ => {
                                                    println!("Invalid input, please enter 'Y' or 'N'");
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            } else {
                worksheet = Ok(Some(workbook.add_worksheet(Some(&sheet)).map_err(|_| ExcelError::WriteWorksheetError(sheet.clone()))?));
                if let Some(Ok(new_range)) = new_excel_data.worksheet_range(&sheet) {
                    for (n, row) in new_range.rows().enumerate() {
                        for (m, new_value) in row.iter().enumerate() {
                            if let Some(ref mut ws) = worksheet.as_mut().map_err(|_| ExcelError::WriteWorksheetError(sheet.clone()))? {
                                ws.write_string(n as u32, m as u16, new_value.to_string().as_str(), None)
                                    .map_err(|_| ExcelError::WriteWorksheetError(sheet.clone()))?;
                            }
                        }
                    }
                }
            }
    }

    Ok(())
}