# Data Processing Script

This Python script processes log files to generate descriptive plots and Excel reports. It supports handling individual files or aggregating data from multiple files into a single report. Users can opt to save plots as separate image files or embed them directly in the Excel output.

## Features

- **Single or Multiple File Processing**: Process an individual log file or an entire directory of log files.
- **Excel Output**: Generates an Excel file with data extracted from log files. Can combine data from multiple files into a single Excel file.
- **Plot Generation**: Generates scatter plots of data points. Plots can be saved as images or embedded in the Excel output.
- **Customizable Output**: Allows users to specify whether to combine data and whether to save plots as separate images.

## Prerequisites

Before you run the script, make sure you have Python installed on your system. The script requires the following Python packages:
- `pandas`
- `matplotlib`
- `xlsxwriter`

You can install these packages using pip:

    pip install pandas matplotlib xlsxwriter

### Virtual Environment

If you want to use a virtual environment, you can initialize one using this command:

    python -m venv myenv

To activate an existing virtual environment, use this command:

    myenv\Scripts\activate.bat

## Usage

### Command Line Arguments

The script accepts several command-line options to specify input files, output directory, and operational modes:

- `-i`, `--ifile`: Specify the input file or directory.
- `-o`, `--odir`: Specify the output directory where the Excel and plot files will be saved.
- `-s`, `--savepng`: Save plots as PNG files. If not specified, plots are embedded in the Excel file.
- `-c`, `--combine`: Combine data from multiple log files into a single Excel file and plot.

### Running the Script

Here's how you can run the script from the command line:

    python process-data.py -i path_to_log_file_or_directory -o output_directory [-s] [-c]

### Examples

**Process a single file without combining and embed the plot in the Excel file:**

    python process-data.py -i logs/logfile1.log -o output

**Process multiple files, combine the data, and save plots as separate PNG files:**

    python process-data.py -i logs -o output -s -c

## Output

- Excel files and PNG files (if requested) are saved in the specified output directory.
- Excel files include data tables and, optionally, embedded plots.

## Contributing

Feel free to fork this repository and submit pull requests or propose changes by opening issues.
