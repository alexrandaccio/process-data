import csv, os, sys, getopt
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

def process_files(file_list, output_dir, save_plot_as_png, combine, input_path):
    combined_df = pd.DataFrame()
    if combine:
        fig, ax = plt.subplots(figsize=(10, 6))
    else:
        fig, ax = None, None

    for input_file in file_list:
        with open(input_file, newline='') as csvfile:
            reader = csv.reader(csvfile, delimiter='\t')
            datetime_str = next(reader)[0]
            next(reader)  # Skip the second header line
            data = list(reader)
        df = pd.DataFrame(data, columns=['Reading', 'Load [ozFin]', 'Time [sec.]'])
        df['Trial'] = os.path.basename(input_file)
        df['Datetime'] = datetime_str
        df['Reading'] = pd.to_numeric(df['Reading'])
        df['Load [ozFin]'] = pd.to_numeric(df['Load [ozFin]'])
        df['Time [sec.]'] = pd.to_numeric(df['Time [sec.]'])
        df['ABS(Load)'] = df['Load [ozFin]'].abs()

        if combine:
            combined_df = pd.concat([combined_df, df])
            ax.scatter(df['Time [sec.]'], df['ABS(Load)'], label=os.path.basename(input_file))
        else:
            save_data_and_plot(df, input_file, output_dir, save_plot_as_png)

    if combine:
        directory_name = os.path.basename(os.path.normpath(input_path))
        output_xlsx = os.path.join(output_dir, f"{directory_name}.xlsx")
        writer = pd.ExcelWriter(output_xlsx, engine='xlsxwriter')
        combined_df.to_excel(writer, sheet_name='Data', index=False)
        ax.legend()
        ax.set_title('Combined Scatter Plot of Load vs Time')
        ax.set_xlabel('Time [sec.]')
        ax.set_ylabel('Absolute Load [ozFin]')
        ax.grid(True)
        finalize_plot(ax, writer, save_plot_as_png, f"{directory_name}_plot.png", output_dir)

def save_data_and_plot(df, input_file, output_dir, save_plot_as_png):
    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_xlsx = os.path.join(output_dir, f"{base_name}.xlsx")
    writer = pd.ExcelWriter(output_xlsx, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Data', index=False)

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.scatter(df['Time [sec.]'], df['ABS(Load)'], color='blue', label='Load vs Time')
    ax.set_title(f'Scatter Plot of Load vs Time for {base_name}')
    ax.set_xlabel('Time [sec.]')
    ax.set_ylabel('Absolute Load [ozFin]')
    ax.grid(True)
    ax.legend()
    finalize_plot(ax, writer, save_plot_as_png, f"{base_name}_plot.png", output_dir)

def finalize_plot(ax, writer, save_plot_as_png, plot_filename, output_dir):
    if save_plot_as_png:
        plot_full_path = os.path.join(output_dir, plot_filename)
        plt.savefig(plot_full_path)
        plt.close()
    else:
        imgdata = BytesIO()
        plt.savefig(imgdata, format='png')
        imgdata.seek(0)
        worksheet = writer.sheets['Data']
        worksheet.insert_image('H1', 'plot.png', {'image_data': imgdata})
        plt.close()
    writer.close()

def main(argv):
    input_path = None
    output_dir = 'output_files'
    save_plot_as_png = False
    combine = False
    try:
        opts, args = getopt.getopt(argv,"hi:o:sc",["ifile=", "odir=", "savepng", "combine"])
        for opt, arg in opts:
            if opt == '-h':
                print('Usage: process-data.py -i <inputfile or inputdir> -o <outputdir> [-s] [-c]')
                sys.exit()
            elif opt in ("-i", "--ifile"):
                input_path = arg
            elif opt in ("-o", "--odir"):
                output_dir = arg
            elif opt in ("-s", "--savepng"):
                save_plot_as_png = True
            elif opt in ("-c", "--combine"):
                combine = True

        if input_path is None:
            raise ValueError("Input file or directory must be specified with -i option.")

        os.makedirs(output_dir, exist_ok=True)
        file_list = [os.path.join(input_path, f) for f in os.listdir(input_path)] if os.path.isdir(input_path) else [input_path]

        process_files(file_list, output_dir, save_plot_as_png, combine, input_path)

    except getopt.GetoptError as err:
        print(str(err))
        print('Usage: process-data.py -i <inputfile or inputdir> -o <outputdir> [-s] [-c]')
        sys.exit(2)
    except ValueError as ve:
        print(str(ve))
        sys.exit(2)

if __name__ == "__main__":
    main(sys.argv[1:])