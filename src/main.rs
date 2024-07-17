use calamine::{open_workbook, Data, Reader, Xlsx};
use rust_xlsxwriter::{Workbook, Worksheet};

const COLUMN_FILTER: [&str; 7] = [
    "Art. nr.",
    "Merk",
    "Omschrijving",
    "Inhoud",
    "BTW",
    "AVP",
    "Aantal",
];

fn write_cell(
    sheet: &mut Worksheet,
    row: usize,
    col: usize,
    cell: &calamine::Data,
) -> Result<(), rust_xlsxwriter::XlsxError> {
    let row = row as u32;
    let col = col as u16;

    match cell {
        calamine::Data::Empty => sheet,
        calamine::Data::String(s) => sheet.write_string(row, col, s)?,
        calamine::Data::Float(f) => sheet.write_number(row, col, *f)?,
        calamine::Data::Int(i) => sheet.write_number(row, col, *i as f64)?,
        calamine::Data::Bool(b) => sheet.write_boolean(row, col, *b)?;
    };

    Ok(())
}

fn main() {
    let mut wb: Xlsx<_> = open_workbook("test.xlsx").expect("cannot open file");

    println!("Found the following sheets: {:#?}", wb.sheet_names());

    if let Ok(range) = wb.worksheet_range("Worksheet") {
        println!("Found the following headers: {:#?}", range.headers());
        let col_to_keep = range
            .headers()
            .unwrap()
            .iter()
            .enumerate()
            .filter(|(_, name)| COLUMN_FILTER.contains(&name.as_str()))
            .map(|(i, _)| i)
            .collect::<Vec<_>>();
        println!("Determined the column indices to keep: {col_to_keep:?}");

        // Create a fresh target workbook
        let mut target = Workbook::new();
        let worksheet = target.add_worksheet();

        for (row_index, row) in range.rows().enumerate() {
            row.iter()
                .enumerate()
                .filter_map(|(i, cell)| col_to_keep.contains(&i).then_some(cell))
                .enumerate()
                .map(|(i, cell)| {
                    write_cell(&mut worksheet, row_index, i, cell);
                });
        }
    }
}
