//! fast_xlsx CLI - High-performance CSV to XLSX converter
//!
//! Usage: fast_xlsx input.csv output.xlsx [--sheet-name "Sheet1"]

use clap::Parser;
use std::time::Instant;

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

    /// Show progress information
    #[arg(short, long)]
    verbose: bool,
}

fn main() {
    let args = Args::parse();

    if args.verbose {
        eprintln!("fast_xlsx - CSV to XLSX converter");
        eprintln!("Input:  {}", args.input);
        eprintln!("Output: {}", args.output);
        eprintln!("Sheet:  {}", args.sheet_name);
    }

    let start = Instant::now();

    match xlsxturbo::convert_csv_to_xlsx(&args.input, &args.output, &args.sheet_name) {
        Ok((rows, cols)) => {
            if args.verbose {
                let duration = start.elapsed();
                eprintln!(
                    "Converted {} rows x {} cols in {:.2}s ({:.0} rows/sec)",
                    rows,
                    cols,
                    duration.as_secs_f64(),
                    rows as f64 / duration.as_secs_f64()
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
