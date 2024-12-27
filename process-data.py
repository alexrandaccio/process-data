import csv, sys, getopt
import pandas as pd
import matplotlib.pyplot as plt

def main(argv):
    outputfile = 'output.xlsx'
    try:
        opts, args = getopt.getopt(argv,"ho:",["ofile="])
        input_file = args[0]
        with open(input_file, newline='') as csvfile:
            reader = csv.reader(csvfile, delimiter='\t')
            next(reader)  # Skip the date and time header
            next(reader)  # Skip the column headers
            
            # Reading and processing data
            data = list(reader)
            df = pd.DataFrame(data, columns=['Reading', 'Load [ozFin]', 'Time [sec.]'])
            df['Reading'] = pd.to_numeric(df['Reading'])
            df['Load [ozFin]'] = pd.to_numeric(df['Load [ozFin]'])
            df['Time [sec.]'] = pd.to_numeric(df['Time [sec.]'])
            df['ABS(Load)'] = df['Load [ozFin]'].abs()

    except getopt.GetoptError:
        print('Usage: process-data.py -o <outputfile> <input_file>')
        sys.exit(2)
    except IndexError:
        print('Input file required.')
        sys.exit(2)
    except FileNotFoundError:
        print('File not found.')
        sys.exit(2)

    for opt, arg in opts:
        if opt == '-h':
            print('Usage: process-data.py -o <outputfile> <input_file>')
            sys.exit()
        elif opt in ("-o", "--ofile"):
            outputfile = arg

    # Save to Excel
    df.to_excel(outputfile, index=False)

    # Plotting
    plt.figure(figsize=(10, 6))
    plt.scatter(df['Time [sec.]'], df['ABS(Load)'], color='blue', label='Load vs Time')
    plt.title('Load vs Time')
    plt.xlabel('Time (s)')
    plt.ylabel('Load (in-oz)')
    plt.xlim(0, 80)  # Set x-axis limits
    plt.ylim(0, 50)  # Set y-axis limits
    plt.grid(True)
    plt.legend()
    plt.savefig(outputfile.replace('.xlsx', '_plot.png'))
    plt.show()

if __name__ == "__main__":
    main(sys.argv[1:])
