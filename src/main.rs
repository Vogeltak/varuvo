/// Tool that enriches Varuvo order exports to be useful when Ellen is
/// doing administrative work after receiving the delivery.
///
/// Things it does:
/// 1. Only retain useful columns
/// 2. Insert actual VAT percentages (high vs low bracket)
/// 3. Add subtotal for every row
/// 4. Add VAT for every row
/// 5. Add total (subtotal + VAT) for every row
/// 6. Sum subtotal, vat, and total for all rows
///
use calamine::{open_workbook, Data, Reader, Xlsx};
use rust_xlsxwriter::{Format, Formula, Workbook, Worksheet};

const COLUMN_FILTER: [&str; 7] = [
    "Art. nr.",
    "Merk",
    "Omschrijving",
    "Inhoud",
    "BTW",
    "AVP",
    "Aantal",
];

const COLUMN_MAPPING: [&str; 26] = [
    "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S",
    "T", "U", "V", "W", "X", "Y", "Z",
];

fn col_of(header: &str) -> usize {
    COLUMN_FILTER
        .iter()
        .enumerate()
        .find(|(_, &c)| c == header)
        .unwrap()
        .0
}

fn write_cell(
    sheet: &mut Worksheet,
    row: usize,
    col: u16,
    cell: &calamine::Data,
) -> Result<(), rust_xlsxwriter::XlsxError> {
    let row = row as u32;

    match cell {
        Data::Empty => sheet,
        Data::String(s) => sheet.write_string(row, col, s)?,
        Data::Float(f) => sheet.write_number(row, col, *f)?,
        Data::Int(i) => sheet.write_number(row, col, *i as f64)?,
        Data::Bool(b) => sheet.write_boolean(row, col, *b)?,
        _ => sheet.write_string(row, col, &format!("{:?}", cell))?,
    };

    Ok(())
}

struct NextFreeCol {
    inner: u16,
}

impl NextFreeCol {
    fn new(col: usize) -> Self {
        Self { inner: col as u16 }
    }

    fn get(&mut self) -> u16 {
        let res = self.inner;
        self.inner += 1;
        res
    }
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
        // Force recalculation of formulas in LibreOffice.
        // See https://bugs.documentfoundation.org/show_bug.cgi?id=144819
        worksheet.set_formula_result_default("");

        // First write the headers to the new worksheet.
        if let Some(headers) = range.rows().next() {
            headers
                .iter()
                .enumerate()
                // Select the columns that we want to keep.
                .filter_map(|(i, cell)| col_to_keep.contains(&i).then_some(cell))
                // Give them new column indices.
                .enumerate()
                .for_each(|(i, cell)| write_cell(worksheet, 0, i as u16, cell).unwrap())
        }

        // Add the headers for the columns we're adding to this sheet.
        let mut next_free_header = NextFreeCol::new(COLUMN_FILTER.len());
        write_cell(
            worksheet,
            0,
            next_free_header.get(),
            &Data::String("Subtotaal".to_string()),
        )
        .unwrap();
        write_cell(
            worksheet,
            0,
            next_free_header.get(),
            &Data::String("BTW totaal".to_string()),
        )
        .unwrap();
        write_cell(
            worksheet,
            0,
            next_free_header.get(),
            &Data::String("Totaal".to_string()),
        )
        .unwrap();

        // Create our currency format for the cells we're adding.
        let currency_format = Format::new().set_num_format("$#,##0.00");

        // Process rows containing the actual items.
        for (row_index, row) in range.rows().enumerate().skip(1) {
            let new_row = row
                .iter()
                .enumerate()
                // Select the columns that we want to keep.
                .filter_map(|(i, cell)| col_to_keep.contains(&i).then_some(cell))
                // Give them new column indices.
                .enumerate()
                .collect::<Vec<_>>();

            new_row.iter().for_each(|(i, cell)| {
                // Replace VAT strings with their actual percentages
                let cell = match cell {
                    Data::String(s) => match s.as_str() {
                        "BTW Laag" => &Data::Float(0.09),
                        "BTW Hoog" => &Data::Float(0.21),
                        _ => cell,
                    },
                    _ => cell,
                };
                write_cell(worksheet, row_index, *i as u16, cell).unwrap();
            });

            // Account for 1-based row count in Excel.
            let formula_row = row_index + 1;

            // Start the next free column count right after our retained columns.
            let mut next_free_col = NextFreeCol::new(new_row.len());

            // Calculate subtotaal by multiplying the # of items (aantal) with
            // the sales price (AVP). Add it as a new column to the row.
            let col_subtotaal = next_free_col.get();
            let col_aantal = col_of("Aantal");
            let col_price = col_of("AVP");
            let subtotaal = Formula::new(format!(
                "={}{formula_row}*{}{formula_row}",
                COLUMN_MAPPING[col_price], COLUMN_MAPPING[col_aantal]
            ));
            worksheet
                .write_formula_with_format(
                    row_index as u32,
                    col_subtotaal,
                    subtotaal,
                    &currency_format,
                )
                .unwrap();

            // Calculate the VAT over the subtotal.
            let col_btw_totaal = next_free_col.get();
            let btw_totaal = Formula::new(format!(
                "={}{formula_row}*{}{formula_row}",
                COLUMN_MAPPING[col_subtotaal as usize],
                COLUMN_MAPPING[col_of("BTW")]
            ));
            worksheet
                .write_formula_with_format(
                    row_index as u32,
                    col_btw_totaal,
                    btw_totaal,
                    &currency_format,
                )
                .unwrap();

            // Calculate the total amount.
            let col_totaal = next_free_col.get();
            let totaal = Formula::new(format!(
                "={}{formula_row} + {}{formula_row}",
                COLUMN_MAPPING[col_subtotaal as usize], COLUMN_MAPPING[col_btw_totaal as usize]
            ));
            worksheet
                .write_formula_with_format(row_index as u32, col_totaal, totaal, &currency_format)
                .unwrap();
        }

        let res_format = currency_format.set_bold();

        // Okay, crimes crimes, hardcoding column indices here.

        // Calculate sums for all our totals.
        let res_row = range.rows().skip(1).count();
        // Assumptions: Subtotaal @ 7, BTW totaal @ 8, Totaal @ 9
        for col in [7, 8, 9] {
            let col_idx = COLUMN_MAPPING[col];
            let formula = Formula::new(format!("=SUM({col_idx}2:{col_idx}{res_row})"));
            worksheet
                .write_formula_with_format(res_row as u32, col as u16, formula, &res_format)
                .unwrap();
        }

        worksheet.autofit();
        target.save("out.xlsx").unwrap();
    }
}
