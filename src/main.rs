use calamine::{open_workbook, DataType, Reader, Xlsx};
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
    #[arg(long, default_value = "%Y-%m-%dT%h:%M:%SZ")]
    datetime_format: String,

    /// format for rendering time values (if different than datetime)
    #[arg(long)]
    time_format: Option<String>,

    /// format for rendering date values (if different than datetime)
    #[arg(long)]
    date_format: Option<String>,

    /// include cells with errors
    #[arg(long)]
    include_errors: bool,
}

fn parse_cell(cell: &DataType, cli: &Cli) -> String {
    let Cli {
        numeric_bool,
        datetime_format,
        date_format,
        time_format,
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
            let days = *x as i64;
            let date =
                chrono::NaiveDate::from_ymd_opt(1900, 1, 1).unwrap() + chrono::Duration::days(days);
            // get the fractional part of the number, which represents the time of day
            let frac = *x - *x as i32 as f64;
            let time =
                chrono::NaiveTime::from_num_seconds_from_midnight_opt((frac * 86400.0) as u32, 0)
                    .unwrap();
            let datetime = chrono::NaiveDateTime::new(date, time);

            // TODO: we should expose cli options to:
            // 1. specify the date, time, and datetime formats. This would also allow the user to decide whether to infer times/dates instead of just datetimes, since they're all stored the same in xlsx
            if x.floor() == 0.0 {
                // datetime.format("%H:%M:%S").to_string()
                datetime
                    .format(time_format.as_ref().unwrap_or(datetime_format))
                    .to_string()
            } else if x % 1.0 == 0.0 {
                // datetime.format("%Y-%m-%d").to_string()
                datetime
                    .format(date_format.as_ref().unwrap_or(datetime_format))
                    .to_string()
            } else {
                // datetime.format("%Y-%m-%d %H:%M:%S").to_string()
                datetime.format(datetime_format).to_string()
            }
        }
        DataType::Duration(x) => x.to_string(),
        DataType::DateTimeIso(x) => x.clone(),
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
    let mut excel: Xlsx<_> = open_workbook(input)?;
    let mut csv_file = Writer::from_path(output)?;

    let sheet_name = match sheet {
        Some(name) => name.to_string(),
        None => excel.sheet_names().get(0).unwrap().clone(),
    };

    if let Some(Ok(range)) = excel.worksheet_range(&sheet_name) {
        for row in range.rows() {
            let values = row
                .iter()
                // TODO: this currently returns an empty string if get_string fails. Should we error instead?
                // .map(|cell| cell.get_string().unwrap_or_default())
                // .map(|cell| cell.as_string().unwrap_or_default())
                // .map(|cell| cell.to_string())
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

    // if let Err(e) = xlsx_to_csv(&cli.input, &cli.output, cli.sheet.as_deref()) {
    if let Err(e) = xlsx_to_csv(&cli) {
        eprintln!("Error: {}", e);
        std::process::exit(1)
    } else {
        println!("Done");
    }
}
