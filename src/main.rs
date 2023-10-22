use calamine::{open_workbook, DataType, Reader, Xlsx};
use chrono::{Datelike, NaiveDateTime, Timelike};
use clap::Parser;
use csv::Writer;

#[derive(Parser)]
#[command(author, version, about, long_about = None)]
struct Cli {
    /// Input file path
    #[arg(long, short)]
    input: String,

    /// Output file path
    #[arg(long, short)]
    output: String,

    /// Sheet name. Defaults to first sheet in workbook.
    #[arg(long, short)]
    sheet: Option<String>,

    /// render boolean as 1/0 instead of true/false
    #[arg(long)]
    numeric_bool: bool,

    /// format for rendering datetime values
    #[arg(long)]
    datetime_format: Option<String>,

    /// format for rendering time values (if different than datetime)
    #[arg(long)]
    time_format: Option<String>,

    /// format for rendering date values (if different than datetime)
    #[arg(long)]
    date_format: Option<String>,

    /// format for rendering duration
    #[arg(long)]
    duration_hms: bool,

    /// include cells with errors
    #[arg(long)]
    include_errors: bool,
}

/// xlsx stores times with a date of 1899-12-31, so we can use that to detect if a cell is just a time
fn is_time(d: NaiveDateTime) -> bool {
    d.year() == 1899 && d.month() == 12 && d.day() == 31
}

/// technically this finds midnight too
fn is_date(d: NaiveDateTime) -> bool {
    d.hour() == 0 && d.minute() == 0 && d.second() == 0
}

fn parse_cell(cell: &DataType, cli: &Cli) -> String {
    let Cli {
        numeric_bool,
        datetime_format,
        date_format,
        time_format,
        duration_hms: duration_as_hms,
        include_errors,
        ..
    } = cli;
    match cell {
        DataType::Int(x) => x.to_string(),
        DataType::Float(x) => x.to_string(),
        DataType::String(x) => x.clone(),
        DataType::Bool(x) => match (numeric_bool, x) {
            (false, true) => "true".to_string(),
            (false, false) => "false".to_string(),
            (true, true) => "1".to_string(),
            (true, false) => "0".to_string(),
        },
        DataType::DateTime(x) => {
            // if no formatters defined, just stringify immediately
            if datetime_format.is_none() && date_format.is_none() && time_format.is_none() {
                return x.to_string();
            }

            // otherwise, parse and format
            let d = cell.as_datetime().unwrap();
            let fmt = if is_time(d) {
                time_format.as_deref()
            } else if is_date(d) {
                date_format.as_deref()
            } else {
                datetime_format.as_deref()
            };
            match fmt {
                Some(fmt) => d.format(fmt).to_string(),
                None => x.to_string(),
            }
        }
        DataType::Duration(x) => match duration_as_hms {
            false => x.to_string(),
            true => {
                let d = cell.as_duration().unwrap();
                let hours = d.num_hours();
                let minutes = d.num_minutes() - (hours * 60);
                let seconds = d.num_seconds() - (minutes * 60) - (hours * 60 * 60);
                format!("{:02}:{:02}:{:02}", hours, minutes, seconds)
            }
        },
        DataType::DateTimeIso(x) => x.to_string(),
        DataType::DurationIso(x) => x.to_string(),
        DataType::Error(x) => match include_errors {
            true => x.to_string(),
            false => String::default(),
        },
        DataType::Empty => String::default(),
    }
}

fn xlsx_to_csv(cli: &Cli) -> Result<(), Box<dyn std::error::Error>> {
    let Cli {
        input,
        output,
        sheet,
        ..
    } = cli;

    let mut excel: Xlsx<_> = open_workbook(input).map_err(|e| {
        // prepend input filename
        format!("{}: {}", input, e)
    })?;
    let mut csv_file = Writer::from_path(output)?;

    let sheet_name = match sheet {
        Some(name) => name.to_string(),
        None => excel.sheet_names().get(0).unwrap().clone(),
    };

    if let Some(Ok(range)) = excel.worksheet_range(&sheet_name) {
        for row in range.rows() {
            let values = row
                .iter()
                // could alternatively just call cell.to_string() here, but datetimes are funky
                .map(|cell| parse_cell(cell, cli))
                .collect::<Vec<_>>();
            csv_file.write_record(&values)?;
        }
    } else {
        return Err(format!("Couldn't open sheet: '{}'", sheet_name).into());
    }

    csv_file.flush()?;
    Ok(())
}

fn main() {
    let cli = Cli::parse();

    if let Err(e) = xlsx_to_csv(&cli) {
        eprintln!("Error: {}", e);
        std::process::exit(1)
    } else {
        println!("Done");
    }
}
