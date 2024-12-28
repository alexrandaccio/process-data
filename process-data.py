import csv, sys, getopt, os
import pandas as pd
import matplotlib.pyplot as plt

def process_file(input_file, output_dir):
    # Reading and processing data
    with open(input_file, newline='') as csvfile:
        reader = csv.reader(csvfile, delimiter='\t')
        next(reader)  # Skip the date and time header
        next(reader)  # Skip the column headers
        data = list(reader)
    df = pd.DataFrame(data, columns=['Reading', 'Load [ozFin]', 'Time [sec.]'])
    df['Reading'] = pd.to_numeric(df['Reading'])
    df['Load [ozFin]'] = pd.to_numeric(df['Load [ozFin]'])
    df['Time [sec.]'] = pd.to_numeric(df['Time [sec.]'])
    df['ABS(Load)'] = df['Load [ozFin]'].abs()

    # Output file paths
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_xlsx = os.path.join(output_dir, f"{base_name}.xlsx")
    output_plot = os.path.join(output_dir, f"{base_name}_plot.png")

    # Save to Excel
    df.to_excel(output_xlsx, index=False)

    # Plotting
    plt.figure(figsize=(10, 6))
    plt.scatter(df['Time [sec.]'], df['ABS(Load)'], color='blue', label='Load vs Time')
    plt.title(f'Scatter Plot of Load vs Time for {base_name}')
    plt.xlabel('Time [sec.]')
    plt.ylabel('Absolute Load [ozFin]')
    plt.xlim(0, 80)  # Set x-axis limits
    plt.ylim(0, 50)  # Set y-axis limits
    plt.grid(True)
    plt.legend()
    plt.savefig(output_plot)
    plt.close()

def main(argv):
    input_path = None
    output_dir = 'output_files'
    try:
        opts, args = getopt.getopt(argv,"hi:o:",["ifile=", "odir="])
        for opt, arg in opts:
            if opt == '-h':
                print('Usage: process-data.py -i <inputfile or inputdir> -o <outputdir>')
                sys.exit()
            elif opt in ("-i", "--ifile"):
                input_path = arg
            elif opt in ("-o", "--odir"):
                output_dir = arg

        if input_path is None:
            raise ValueError("Input file or directory must be specified with -i option.")

        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)

        # Check if input path is a directory or a file
        if os.path.isdir(input_path):
            # Process each file in the directory
            for filename in os.listdir(input_path):
                if filename.endswith('.log'):  # Adjusted to process .log files
                    file_path = os.path.join(input_path, filename)
                    process_file(file_path, output_dir)
        elif os.path.isfile(input_path):
            # Process a single file
            process_file(input_path, output_dir)
        else:
            print(f"Error: {input_path} is neither a valid file nor a directory")
            sys.exit(1)

    except getopt.GetoptError as err:
        print(str(err))
        print('Usage: process-data.py -i <inputfile or inputdir> -o <outputdir>')
        sys.exit(2)
    except ValueError as ve:
        print(str(ve))
        sys.exit(2)

if __name__ == "__main__":
    main(sys.argv[1:])