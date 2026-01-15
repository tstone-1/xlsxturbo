//! fast_xlsx CLI - High-performance CSV to XLSX converter
//!
//! Usage: fast_xlsx input.csv output.xlsx [--sheet-name "Sheet1"]

use clap::Parser;
use std::time::Instant;
use xlsxturbo::DateOrder;

#[derive(Parser, Debug)]
#[command(name = "fast_xlsx")]
#[command(version)]
#[command(about = "Fast CSV to XLSX converter with automatic type detection")]
#[command(
    long_about = "Converts CSV files to Excel XLSX format with smart type inference:\n\
    - Numbers (integers, floats) -> Excel numbers\n\
    - Booleans (true/false) -> Excel booleans\n\
    - Dates (YYYY-MM-DD) -> Excel dates\n\
    - Datetimes (ISO 8601) -> Excel datetimes\n\
    - NaN/Inf -> Empty cells"
)]
struct Args {
    /// Input CSV file path
    input: String,

    /// Output XLSX file path
    output: String,

    /// Sheet name (default: "Sheet1")
    #[arg(short, long, default_value = "Sheet1")]
    sheet_name: String,

    /// Date order for ambiguous dates like 01-02-2024
    /// auto: ISO first, then European (DMY), then US (MDY)
    /// mdy/us: US format (01-02-2024 = January 2)
    /// dmy/eu: European format (01-02-2024 = February 1)
    #[arg(short, long, default_value = "auto")]
    date_order: String,

    /// Show progress information
    #[arg(short, long)]
    verbose: bool,
}

fn main() {
    let args = Args::parse();

    let date_order = DateOrder::parse(&args.date_order).unwrap_or_else(|| {
        eprintln!(
            "Invalid date_order '{}'. Valid values: auto, mdy, us, dmy, eu",
            args.date_order
        );
        std::process::exit(1);
    });

    if args.verbose {
        eprintln!("fast_xlsx - CSV to XLSX converter");
        eprintln!("Input:  {}", args.input);
        eprintln!("Output: {}", args.output);
        eprintln!("Sheet:  {}", args.sheet_name);
        eprintln!("Dates:  {:?}", date_order);
    }

    let start = Instant::now();

    match xlsxturbo::convert_csv_to_xlsx(&args.input, &args.output, &args.sheet_name, date_order) {
        Ok((rows, cols)) => {
            if args.verbose {
                let duration = start.elapsed();
                let secs = duration.as_secs_f64();
                let rows_per_sec = if secs > 0.0 {
                    format!("{:.0}", rows as f64 / secs)
                } else {
                    "instant".to_string()
                };
                eprintln!(
                    "Converted {} rows x {} cols in {:.2}s ({} rows/sec)",
                    rows, cols, secs, rows_per_sec
                );
            }
            println!("OK {} {}", rows, cols);
        }
        Err(e) => {
            eprintln!("Error: {}", e);
            std::process::exit(1);
        }
    }
}
