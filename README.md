# Excel Files Merger

This Python script allows you to merge multiple Excel files from a selected folder into a single Excel file. The script uses a graphical user interface (GUI) to let you select the folder containing the Excel files and the location to save the combined file.

## Features

- Select a folder containing Excel files.
- Combine all Excel files in the selected folder into one.
- Save the combined Excel file to a specified location.

## Requirements

- Python 3.x
- pandas
- openpyxl
- tkinter

## Installation

1. Clone the repository:

   ```sh
   git clone https://github.com/zackha/excel-files-merger.git
   cd excel-files-merger
   ```

2. Install the required Python packages:
   ```sh
   pip install pandas openpyxl
   ```

## Usage

1. Run the script:

   ```sh
   python merge_excel_files.py
   ```

2. A window will appear:
   - Click "Select Folder" to choose the folder containing your Excel files.
   - Click "Select Save Location" to choose where to save the combined Excel file.
   - Click "Combine and Save" to merge the files and save the combined file.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
