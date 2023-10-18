use calamine::{open_workbook, Reader, Xlsx};
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
            let values: Vec<&str> = row
                .iter()
                // TODO: this currently returns an empty string if get_string fails. Should we error instead?
                .map(|cell| cell.get_string().unwrap_or_default())
                .collect();
            csv_file.write_record(&values)?;
        }
    } else {
        // return Err(anyhow!("Couldn't open sheet: '{}'", sheet_name));
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
