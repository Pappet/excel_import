use calamine::{open_workbook, Ods, Reader, Xlsx};
use csv::Writer;

use std::collections::HashSet;
use std::env;
use std::error::Error;
use std::path::Path;

fn main() {
    let args: Vec<String> = env::args().collect();
    if args.len() != 4 {
        eprintln!("Verwendung: excel_import <Pfad zur Excel-Datei> <Sheet-Name> <Spaltenname>");
        return;
    }

    let path = &args[1];
    let path_as_path = Path::new(path);
    let output_file_path = "output.csv";
    let sheet_name = &args[2];
    let column_name = &args[3];

    let unique_values_result = match path_as_path.extension().and_then(std::ffi::OsStr::to_str) {
        Some("ods") => read_unique_values_ods(path, sheet_name, column_name),
        Some("xlsx") | _ => read_unique_values_xlsx(path, sheet_name, column_name),
    };

    match unique_values_result {
        Ok(unique_values) => {
            // Speichere das Ergebnis in der Ausgabedatei
            if let Err(e) = write_to_csv_file(unique_values, column_name, output_file_path) {
                eprintln!("Fehler beim Schreiben der CSV-Datei: {}", e);
            } else {
                println!(
                    "Einzigartige Werte in Spalte {} wurden in '{}' gespeichert.",
                    column_name, output_file_path
                );
            }
        }
        Err(e) => eprintln!("Fehler: {}", e),
    }
}

fn get_column_number(
    sheet: &calamine::Range<calamine::DataType>,
    column_name: &str,
) -> Option<usize> {
    let header_row = sheet.rows().nth(0)?;
    header_row
        .iter()
        .position(|cell| cell.to_string() == column_name)
}

fn read_unique_values_xlsx<P: AsRef<std::path::Path>>(
    path: P,
    sheet_name: &str,
    column_name: &str,
) -> Result<HashSet<String>, Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    read_unique_values_from_workbook(&mut workbook, sheet_name, column_name)
}

fn read_unique_values_ods<P: AsRef<std::path::Path>>(
    path: P,
    sheet_name: &str,
    column_name: &str,
) -> Result<HashSet<String>, Box<dyn std::error::Error>> {
    let mut workbook: Ods<_> = open_workbook(path)?;
    read_unique_values_from_workbook(&mut workbook, sheet_name, column_name)
}

fn read_unique_values_from_workbook<R: Reader>(
    workbook: &mut R,
    sheet_name: &str,
    column_name: &str,
) -> Result<HashSet<String>, Box<dyn std::error::Error>> {
    let mut unique_values = HashSet::new();

    if let Some(Ok(worksheet)) = workbook.worksheet_range(sheet_name) {
        if let Some(column) = get_column_number(&worksheet, column_name) {
            for row in worksheet.rows().skip(1) {
                if let Some(cell) = row.get(column) {
                    let value = cell.to_string();
                    if !value.is_empty() {
                        unique_values.insert(value);
                    }
                }
            }
        } else {
            return Err(format!("Spalte '{}' nicht gefunden.", column_name).into());
        }
    } else {
        return Err(format!("Arbeitsblatt '{}' nicht gefunden.", sheet_name).into());
    }

    Ok(unique_values)
}

fn write_to_csv_file(
    unique_values: HashSet<String>,
    _column_name: &str,
    output: &str,
) -> Result<(), Box<dyn Error>> {
    let mut wtr = Writer::from_path(output)?;
    for value in sort_hashset(unique_values) {
        wtr.write_record(&[&value])?;
    }
    wtr.flush()?;
    Ok(())
}

fn sort_hashset(hashset: HashSet<String>) -> Vec<String> {
    let mut vec: Vec<String> = hashset.into_iter().collect();
    vec.sort();
    vec
}
