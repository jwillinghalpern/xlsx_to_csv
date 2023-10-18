# xlsx_to_csv

CLI tool to convert xlsx files to csv

## Usage

### Convert in.xlsx to out.csv

```bash
xlsx_to_csv -i in.xlsx -o out.csv
```

### Show help

```bash
xlsx_to_csv -h
```

## Build

### Windows

```bash
cargo build --target x86_64-pc-windows-gnu -r
```

### Mac

```bash
cargo build -r
```
