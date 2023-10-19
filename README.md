# xlsx_to_csv

CLI tool to convert xlsx files to csv

## Usage

Show help

```bash
xlsx_to_csv -h
```

Convert first sheet from in.xlsx to out.csv.

```bash
xlsx_to_csv -i in.xlsx -o out.csv
```

Convert specific sheet from in.xlsx to out.csv.

```bash
xlsx_to_csv -i in.xlsx -o out.csv --sheet "MySheetName"
```

Customize the datetime, date, and time formats. If not specified, time and date will use the datetime format and everything will be treated as a datetime. BEWARE xlsx treats datetimes without dates as jan 1 1900.

```bash
xlsx_to_csv -i in.xlsx -o out.csv --date-format="%m/%d/%Y" --time-format="%r" --datetime-format="%m/%d/%Y %r"
```

print booleans as 0 or 1 instead of true or false

```bash
xlsx_to_csv -i in.xlsx -o out.csv --numeric-bool
```

include cell errors in output

```bash
xlsx_to_csv -i in.xlsx -o out.csv --include-errors
```

## Build

### Windows

```bash
cargo build --target x86_64-pc-windows-gnu -r
```

### Mac (assuming running on mac)

```bash
cargo build -r
```

### Linux (untested)

```bash
cross build --target x86_64-unknown-linux-gnu -r
```
