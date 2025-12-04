//! fast_xlsx - High-performance CSV to XLSX converter
//!
//! Usage: fast_xlsx input.csv output.xlsx [--sheet-name "Sheet1"]

use clap::Parser;
use csv::ReaderBuilder;
use rust_xlsxwriter::{Workbook, Worksheet, XlsxError};
use std::error::Error;
use std::fs::File;
use std::io::BufReader;
use std::time::Instant;

#[derive(Parser, Debug)]
#[command(name = "fast_xlsx")]
#[command(about = "Fast CSV to XLSX converter", long_about = None)]
struct Args {
    /// Input CSV file path
    input: String,

    /// Output XLSX file path
    output: String,

    /// Sheet name (default: "Sheet1")
    #[arg(short, long, default_value = "Sheet1")]
    sheet_name: String,

    /// Show progress information
    #[arg(short, long)]
    verbose: bool,
}

fn write_value(worksheet: &mut Worksheet, row: u32, col: u16, value: &str) -> Result<(), XlsxError> {
    // Try to parse as different types for proper Excel formatting
    if value.is_empty() {
        // Empty cell - write empty string
        worksheet.write_string(row, col, "")?;
    } else if let Ok(int_val) = value.parse::<i64>() {
        worksheet.write_number(row, col, int_val as f64)?;
    } else if let Ok(float_val) = value.parse::<f64>() {
        if float_val.is_nan() || float_val.is_infinite() {
            // Handle NaN/Inf as empty or error string
            worksheet.write_string(row, col, "")?;
        } else {
            worksheet.write_number(row, col, float_val)?;
        }
    } else if value.eq_ignore_ascii_case("true") {
        worksheet.write_boolean(row, col, true)?;
    } else if value.eq_ignore_ascii_case("false") {
        worksheet.write_boolean(row, col, false)?;
    } else {
        worksheet.write_string(row, col, value)?;
    }
    Ok(())
}

fn convert_csv_to_xlsx(args: &Args) -> Result<(u32, u16), Box<dyn Error>> {
    let start = Instant::now();

    // Open CSV file
    let file = File::open(&args.input)?;
    let reader = BufReader::with_capacity(1024 * 1024, file); // 1MB buffer
    let mut csv_reader = ReaderBuilder::new()
        .has_headers(false) // We'll handle headers manually
        .flexible(true) // Allow variable record lengths
        .from_reader(reader);

    // Create workbook
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name(&args.sheet_name)?;

    let mut row_count: u32 = 0;
    let mut col_count: u16 = 0;

    // Process all records
    for result in csv_reader.records() {
        let record = result?;
        let num_cols = record.len() as u16;
        if num_cols > col_count {
            col_count = num_cols;
        }

        for (col_idx, value) in record.iter().enumerate() {
            write_value(worksheet, row_count, col_idx as u16, value)?;
        }

        row_count += 1;

        // Progress indicator for verbose mode
        if args.verbose && row_count % 100_000 == 0 {
            eprintln!("  Processed {} rows...", row_count);
        }
    }

    // Save workbook
    workbook.save(&args.output)?;

    if args.verbose {
        let duration = start.elapsed();
        eprintln!(
            "Converted {} rows x {} cols in {:.2}s ({:.0} rows/sec)",
            row_count,
            col_count,
            duration.as_secs_f64(),
            row_count as f64 / duration.as_secs_f64()
        );
    }

    Ok((row_count, col_count))
}

fn main() {
    let args = Args::parse();

    if args.verbose {
        eprintln!("fast_xlsx - CSV to XLSX converter");
        eprintln!("Input:  {}", args.input);
        eprintln!("Output: {}", args.output);
    }

    match convert_csv_to_xlsx(&args) {
        Ok((rows, cols)) => {
            println!("OK {} {}", rows, cols);
        }
        Err(e) => {
            eprintln!("Error: {}", e);
            std::process::exit(1);
        }
    }
}
