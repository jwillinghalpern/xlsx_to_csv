use assert_cmd::Command;
use predicates::prelude::*;
use rand::{distributions::Alphanumeric, Rng};
use std::fs;
use tempfile::NamedTempFile;

type TestResult = Result<(), Box<dyn std::error::Error>>;

const PRG: &str = "xlsx_to_csv";

const TYPES_FILE: &str = "tests/inputs/types.xlsx";

const CUSTOM_SHEET_FILE: &str = "tests/inputs/custom-sheet.xlsx";
const SHEET2: &str = "CustomSheet2";

const DONE: &str = "Done\n";

// --------------------------------------------------
fn gen_bad_file() -> String {
    loop {
        let filename: String = rand::thread_rng()
            .sample_iter(&Alphanumeric)
            .take(7)
            .map(char::from)
            .collect();

        if fs::metadata(&filename).is_err() {
            return filename;
        }
    }
}

// --------------------------------------------------
fn run_outfile(args: &[&str], expected_file: &str) -> TestResult {
    let expected = fs::read_to_string(expected_file)?;
    let tmp_file = NamedTempFile::new()?;
    let tmp_path = &tmp_file.path().to_str().unwrap();

    Command::cargo_bin(PRG)?
        .args(args)
        .args(&["-o", tmp_path])
        .assert()
        .success()
        .stdout(DONE);

    let contents = fs::read_to_string(&tmp_path)?;
    assert_eq!(&expected, &contents);

    Ok(())
}

// --------------------------------------------------
#[test]
fn dies_bad_file() -> TestResult {
    let bad = gen_bad_file();
    let expected = format!("{}: .* [(]os error 2[)]", bad);
    Command::cargo_bin(PRG)?
        .args(&["-i", &bad, "-o", "not-important.csv"])
        .assert()
        .failure()
        .stderr(predicate::str::is_match(expected)?);
    Ok(())
}

// --------------------------------------------------
#[test]
fn dies_no_args() -> TestResult {
    Command::cargo_bin(PRG)?
        .assert()
        .failure()
        .stderr(predicate::str::contains("Usage"));
    Ok(())
}

// --------------------------------------------------
#[test]
fn default_first_sheet() -> TestResult {
    run_outfile(
        &["-i", CUSTOM_SHEET_FILE],
        "tests/expected/default-first-sheet.csv",
    )
}

// --------------------------------------------------
#[test]
fn custom_sheet_name() -> TestResult {
    run_outfile(
        &["-i", CUSTOM_SHEET_FILE, "-s", SHEET2],
        "tests/expected/custom-sheet-name.csv",
    )
}

// --------------------------------------------------
#[test]
fn no_options() -> TestResult {
    run_outfile(&["-i", TYPES_FILE], "tests/expected/no-options.csv")
}

// --------------------------------------------------
#[test]
fn date_format() -> TestResult {
    run_outfile(
        &["-i", TYPES_FILE, "--date-format", "%Y-%m-%d"],
        "tests/expected/date-format.csv",
    )
}

// --------------------------------------------------
#[test]
fn date_format2() -> TestResult {
    run_outfile(
        &["-i", TYPES_FILE, "--date-format", "%m/%d/%y"],
        "tests/expected/date-format2.csv",
    )
}

// --------------------------------------------------
#[test]
fn time_format() -> TestResult {
    run_outfile(
        &["-i", TYPES_FILE, "--time-format", "%H:%M:%S"],
        "tests/expected/time-format.csv",
    )
}
// --------------------------------------------------
#[test]
fn time_format2() -> TestResult {
    run_outfile(
        &["-i", TYPES_FILE, "--time-format", "%r"],
        "tests/expected/time-format2.csv",
    )
}

// --------------------------------------------------
#[test]
fn datetime_format() -> TestResult {
    run_outfile(
        &["-i", TYPES_FILE, "--datetime-format", "%Y-%m-%d %H:%M:%S"],
        "tests/expected/datetime-format.csv",
    )
}

// --------------------------------------------------
#[test]
fn datetime_format2() -> TestResult {
    run_outfile(
        &["-i", TYPES_FILE, "--datetime-format", "%m/%d/%Y %r"],
        "tests/expected/datetime-format2.csv",
    )
}

// --------------------------------------------------
#[test]
fn include_errors() -> TestResult {
    run_outfile(
        &["-i", TYPES_FILE, "--include-errors"],
        "tests/expected/include-errors.csv",
    )
}

// --------------------------------------------------
#[test]
fn duration_hms() -> TestResult {
    run_outfile(
        &["-i", TYPES_FILE, "--duration-hms"],
        "tests/expected/duration-hms.csv",
    )
}

// --------------------------------------------------
#[test]
fn numeric_bool() -> TestResult {
    run_outfile(
        &["-i", TYPES_FILE, "--numeric-bool"],
        "tests/expected/numeric-bool.csv",
    )
}

// --------------------------------------------------
#[test]
fn all() -> TestResult {
    run_outfile(
        &[
            "-i",
            TYPES_FILE,
            "--date-format",
            "%Y-%m-%d",
            "--time-format",
            "%H:%M:%S",
            "--datetime-format",
            "%Y-%m-%d %H:%M:%S",
            "--include-errors",
            "--duration-hms",
            "--numeric-bool",
        ],
        "tests/expected/all-options.csv",
    )
}
