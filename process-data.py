import csv, os, sys, getopt
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

def process_file(input_file, output_dir, save_plot_as_png):
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

    # Create a writer with xlsxwriter engine
    writer = pd.ExcelWriter(output_xlsx, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Data', index=False)

    # Create the plot
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.scatter(df['Time [sec.]'], df['ABS(Load)'], color='blue', label='Load vs Time')
    ax.set_title(f'Scatter Plot of Load vs Time for {base_name}')
    ax.set_xlabel('Time [sec.]')
    ax.set_ylabel('Absolute Load [ozFin]')
    ax.set_xlim(0, 80)  # Set x-axis limits
    ax.set_ylim(0, 50)  # Set y-axis limits
    ax.grid(True)
    ax.legend()

    if save_plot_as_png:
        # Save the plot as a PNG file
        plot_filename = os.path.join(output_dir, f"{base_name}_plot.png")
        fig.savefig(plot_filename)
        plt.close()
    else:
        # Save plot to buffer and insert into the Excel file
        imgdata = BytesIO()
        fig.savefig(imgdata, format='png')
        imgdata.seek(0)
        worksheet = writer.sheets['Data']
        worksheet.insert_image('H1', 'plot.png', {'image_data': imgdata})
        plt.close()

    # Close the Pandas Excel writer and output the Excel file
    writer.close()

def main(argv):
    input_path = None
    output_dir = 'output_files'
    save_plot_as_png = False  # Default is False
    try:
        opts, args = getopt.getopt(argv,"hi:o:s",["ifile=", "odir=", "savepng"])
        for opt, arg in opts:
            if opt == '-h':
                print('Usage: process-data.py -i <inputfile or inputdir> -o <outputdir> [-s]')
                sys.exit()
            elif opt in ("-i", "--ifile"):
                input_path = arg
            elif opt in ("-o", "--odir"):
                output_dir = arg
            elif opt in ("-s", "--savepng"):
                save_plot_as_png = True

        if input_path is None:
            raise ValueError("Input file or directory must be specified with -i option.")

        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)

        # Check if input path is a directory or a file
        if os.path.isdir(input_path):
            # Process each file in the directory
            for filename in os.listdir(input_path):
                if filename.endswith('.log'):
                    file_path = os.path.join(input_path, filename)
                    process_file(file_path, output_dir, save_plot_as_png)
        elif os.path.isfile(input_path):
            process_file(input_path, output_dir, save_plot_as_png)
        else:
            print(f"Error: {input_path} is neither a valid file nor a directory")
            sys.exit(1)

    except getopt.GetoptError as err:
        print(str(err))
        print('Usage: process-data.py -i <inputfile or inputdir> -o <outputdir> [-s]')
        sys.exit(2)
    except ValueError as ve:
        print(str(ve))
        sys.exit(2)

if __name__ == "__main__":
    main(sys.argv[1:])