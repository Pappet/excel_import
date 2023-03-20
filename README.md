# excel_import
This Program takes an ods file and extracts a column into a csv file.

Usage:
cargo build --release
sudo ln -s [/path/to/your/binary/in/the/release/folder] /usr/local/bin/excel_import
excel_import "/path/to/your/odsfile" "COLUMN_NAME"
