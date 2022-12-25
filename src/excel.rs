use crate::Thing;
use chrono::prelude::*;
use std::collections::HashMap;
use std::fs::{self, remove_file};
use uuid::Uuid;
use xlsxwriter::{DateTime as XLSDateTime, Format, Workbook, Worksheet};

const FONT_SIZE: f64 = 12.0;

pub fn create_xlsx(values: Vec<Thing>) -> Vec<u8> {
    let uuid = Uuid::new_v4().to_string();
    let workbook = Workbook::new(&uuid);
    let mut sheet = workbook.add_worksheet(None).expect("can add sheet");

    let mut width_map: HashMap<u16, usize> = HashMap::new();

    create_headers(&mut sheet, &mut width_map);

    let fmt = workbook
        .add_format()
        .set_text_wrap()
        .set_font_size(FONT_SIZE);

    let date_fmt = workbook
        .add_format()
        .set_num_format("dd/mm/yyyy hh:mm:ss AM/PM")
        .set_font_size(FONT_SIZE);

    for (i, v) in values.iter().enumerate() {
        add_row(i as u32, &v, &mut sheet, &date_fmt, &mut width_map);
    }

    width_map.iter().for_each(|(k, v)| {
        let _ = sheet.set_column(*k as u16, *k as u16, *v as f64 * 1.2, Some(&fmt));
    });

    workbook.close().expect("workbook can be closed");

    let result = fs::read(&uuid).expect("can read file");
    remove_file(&uuid).expect("can delete file");
    result
}

fn add_row(
    row: u32,
    thing: &Thing,
    sheet: &mut Worksheet,
    date_fmt: &Format,
    width_map: &mut HashMap<u16, usize>,
) {
    add_string_column(row, 0, &thing.title, sheet, width_map);
    add_number_column(row, 1, thing.number1, sheet, width_map);
    add_number_column(row, 2, thing.number2, sheet, width_map);
    add_number_column(row, 3, thing.number1 + thing.number2, sheet, width_map);

    let _ = sheet.set_row(row, FONT_SIZE, None);
}

fn add_string_column(
    row: u32,
    column: u16,
    data: &str,
    sheet: &mut Worksheet,
    mut width_map: &mut HashMap<u16, usize>,
) {
    let _ = sheet.write_string(row + 1, column, data, None);
    set_new_max_width(column, data.len(), &mut width_map);
}

fn add_number_column(
    row: u32,
    column: u16,
    data: f64,
    sheet: &mut Worksheet,
    mut width_map: &mut HashMap<u16, usize>,
) {
    let _ = sheet.write_number(row + 1, column, data, None);
    //set_new_max_width(column, data.len(), &mut width_map);
}

fn add_date_column(
    row: u32,
    column: u16,
    date: &DateTime<Utc>,
    sheet: &mut Worksheet,
    mut width_map: &mut HashMap<u16, usize>,
    date_fmt: &Format,
) {
    let d = XLSDateTime::new(
        date.year() as i16,
        date.month() as i8,
        date.day() as i8,
        date.hour() as i8,
        date.minute() as i8,
        date.second() as f64,
    );

    let _ = sheet.write_datetime(row + 1, column, &d, Some(date_fmt));
    set_new_max_width(column, 26, &mut width_map);
}

fn set_new_max_width(col: u16, new: usize, width_map: &mut HashMap<u16, usize>) {
    match width_map.get(&col) {
        Some(max) => {
            if new > *max {
                width_map.insert(col, new);
            }
        }
        None => {
            width_map.insert(col, new);
        }
    };
}

fn create_headers(sheet: &mut Worksheet, mut width_map: &mut HashMap<u16, usize>) {
    let _ = sheet.write_string(0, 0, "Title", None);
    let _ = sheet.write_string(0, 1, "Number1", None);
    let _ = sheet.write_string(0, 2, "Number2", None);
    let _ = sheet.write_string(0, 3, "Sum", None);

    set_new_max_width(0, "Title".len(), &mut width_map);
    set_new_max_width(1, "Number1".len(), &mut width_map);
    set_new_max_width(2, "Number2".len(), &mut width_map);
    set_new_max_width(3, "Sum".len(), &mut width_map);
}