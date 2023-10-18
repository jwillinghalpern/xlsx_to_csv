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
}

fn parse_cell(cell: &DataType) -> String {
    match cell {
        calamine::DataType::Int(x) => x.to_string(),
        calamine::DataType::Float(x) => x.to_string(),
        calamine::DataType::String(x) => x.clone(),
        calamine::DataType::Bool(x) => match x {
            true => "1".to_string(),
            false => "0".to_string(),
        },
        calamine::DataType::DateTime(x) => {
            // TODO: should we determine if the date value is 0 and treat this as just Time instead of DateTime?
            // TODO: if there is no fractional value, should we treat this as just Date instead of DateTime?
            // f64 represents days since Jan 1, 1900
            // convert it to a human readable date
            let days = *x as i32;
            let date =
                chrono::NaiveDate::from_ymd(1900, 1, 1) + chrono::Duration::days(days.into());
            // get the fractional value of the day
            let frac = *x - *x as i32 as f64;
            // convert fractional val to time
            let time =
                chrono::NaiveTime::from_num_seconds_from_midnight((frac * 86400.0) as u32, 0);
            let datetime = chrono::NaiveDateTime::new(date, time);
            datetime.format("%Y-%m-%d %H:%M:%S").to_string()
        }
        calamine::DataType::Duration(x) => {
            todo!()
        }
        calamine::DataType::DateTimeIso(x) => {
            todo!()
        }
        calamine::DataType::DurationIso(x) => {
            todo!()
        }
        calamine::DataType::Error(x) => {
            todo!()
        }
        calamine::DataType::Empty => "".to_string(),
    }
}

fn xlsx_to_csv(
    input: &str,
    output: &str,
    sheet: Option<&str>,
) -> Result<(), Box<dyn std::error::Error>> {
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
                .map(parse_cell)
                // TODO: this currently returns an empty string if get_string fails. Should we error instead?
                // .map(|cell| cell.get_string().unwrap_or_default())
                // .map(|cell| cell.as_string().unwrap_or_default())
                // .map(|cell| cell.to_string())
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

    if let Err(e) = xlsx_to_csv(&cli.input, &cli.output, cli.sheet.as_deref()) {
        eprintln!("Error: {}", e);
        std::process::exit(1)
    } else {
        println!("Done");
    }
}
