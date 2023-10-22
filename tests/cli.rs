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

macro_rules! expected_file {
    ($str:expr) => {
        concat!("tests/expected/", $str)
    };
}

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
fn run(args: &[&str], expected_file: &str) -> TestResult {
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
fn dies_missing_input() -> TestResult {
    Command::cargo_bin(PRG)?
        .args(&["-o", "not-important.csv"])
        .assert()
        .failure()
        .stderr(predicate::str::contains("Usage"));
    Ok(())
}

// --------------------------------------------------
#[test]
fn dies_missing_output() -> TestResult {
    Command::cargo_bin(PRG)?
        .args(&["-i", TYPES_FILE])
        .assert()
        .failure()
        .stderr(predicate::str::contains("Usage"));
    Ok(())
}

// --------------------------------------------------
#[test]
fn default_first_sheet() -> TestResult {
    run(
        &["-i", CUSTOM_SHEET_FILE],
        expected_file!("default-first-sheet.csv"),
    )
}

// --------------------------------------------------
#[test]
fn custom_sheet_name() -> TestResult {
    run(
        &["-i", CUSTOM_SHEET_FILE, "-s", SHEET2],
        expected_file!("custom-sheet-name.csv"),
    )
}

// --------------------------------------------------
#[test]
fn no_options() -> TestResult {
    run(&["-i", TYPES_FILE], expected_file!("no-options.csv"))
}

// --------------------------------------------------
#[test]
fn date_format() -> TestResult {
    run(
        &["-i", TYPES_FILE, "--date-format", "%Y-%m-%d"],
        expected_file!("date-format.csv"),
    )
}

// --------------------------------------------------
#[test]
fn date_format2() -> TestResult {
    run(
        &["-i", TYPES_FILE, "--date-format", "%m/%d/%y"],
        expected_file!("date-format2.csv"),
    )
}

// --------------------------------------------------
#[test]
fn time_format() -> TestResult {
    run(
        &["-i", TYPES_FILE, "--time-format", "%H:%M:%S"],
        expected_file!("time-format.csv"),
    )
}
// --------------------------------------------------
#[test]
fn time_format2() -> TestResult {
    run(
        &["-i", TYPES_FILE, "--time-format", "%r"],
        expected_file!("time-format2.csv"),
    )
}

// --------------------------------------------------
#[test]
fn datetime_format() -> TestResult {
    run(
        &["-i", TYPES_FILE, "--datetime-format", "%Y-%m-%d %H:%M:%S"],
        expected_file!("datetime-format.csv"),
    )
}

// --------------------------------------------------
#[test]
fn datetime_format2() -> TestResult {
    run(
        &["-i", TYPES_FILE, "--datetime-format", "%m/%d/%Y %r"],
        expected_file!("datetime-format2.csv"),
    )
}

// --------------------------------------------------
#[test]
fn include_errors() -> TestResult {
    run(
        &["-i", TYPES_FILE, "--include-errors"],
        expected_file!("include-errors.csv"),
    )
}

// --------------------------------------------------
#[test]
fn duration_hms() -> TestResult {
    run(
        &["-i", TYPES_FILE, "--duration-hms"],
        expected_file!("duration-hms.csv"),
    )
}

// --------------------------------------------------
#[test]
fn numeric_bool() -> TestResult {
    run(
        &["-i", TYPES_FILE, "--numeric-bool"],
        expected_file!("numeric-bool.csv"),
    )
}

// --------------------------------------------------
#[test]
fn all() -> TestResult {
    run(
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
        expected_file!("all-options.csv"),
    )
}
