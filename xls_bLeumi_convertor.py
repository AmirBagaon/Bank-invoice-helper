import glob, os, sys
import pandas as pd
from collections import defaultdict


def main(argv):
    source_path, dest_path = parseArgs(argv)
    print(f"Source path: {source_path}")
    print(f"Destination path: {dest_path}")
    lst_for_df = analyzeFiles(source_path)
    export(lst_for_df, dest_path)

def parseArgs(argv):
    """
    Parse and validate the arguments
    """
    dest_path = os.path.join(os.getcwd(), "output.xlsx")
    source_path = os.getcwd()
    if argv:
        help_words = {"help", "-help", "--help", "-h", "--h" }
        if argv[0] in help_words:
            info = """
This is a software to sum and merge couple of Bank Leumi monthly excels files.

Usage:
Put the desired xlsx files in one directory. Put the script in the same directory and run it.
The output will be in the same directory with the name 'output.xlsx'.

Optional:
Pass first argument for source directory.
Pass second argument for dest directory.

Example commands:
py xls_bLeumi_convertor.py 
py xls_bLeumi_convertor.py <source_directory>
py xls_bLeumi_convertor.py <source_directory> <dest_directory> 
            """
            print(info)
            exit(0)

        source_path = argv[0]
        if not os.path.isdir(source_path):
            print(f"Path \"{source_path}\" is not valid directory. Aborting")
            exit(-1)
        if len(argv) > 1:
            if not (str(argv[1]).endswith(('.xls', '.xlsx', 'csv'))):
                print("Output path must end with .xlsx/.xls/.csv extension. Aborting")
                exit(-1)

            dest_path = argv[1]
    return source_path, dest_path

def analyzeFiles(path):
    """
    Gathers all the files in the source directory,
    sums the total amount of each purchase/product/'name',
    and merge them into unique instance per 'name', following by its sum.
    """
    d = defaultdict(list)
    for filename in glob.glob(os.path.join(path, '*.xls')):
        with open(filename, 'r', encoding='utf-16') as f: # open in readonly mode
            contents = f.readlines()
        for line in contents:
            sorted_dict = line.strip().split("\t")
            if len(sorted_dict) > 3:
                name = sorted_dict[1]
                date = sorted_dict[0]
                amount = sorted_dict[3]
                if "שם בית העסק" in name:
                    continue
                if not name:
                    if 'סה"כ' in sorted_dict[0]:
                        name = sorted_dict[0]
                        d[name].append([amount,""])
                        continue
                
                d[name].append([amount,date])

    #Calculate the total price for each 'name'
    for item in d:
        total_price = 0
        for purchase in d[item]:
            sorted_dict = purchase[0].strip().replace(',','').strip("₪")
            total_price += float(purchase[0].strip().replace(',','').strip("₪"))
        
        #Insert total price to the start of the list
        d[item].insert(0, total_price)

    #Sort the dict by the total price and get a list of tuples
    sorted_dict = sorted(d.items(), key=lambda k_v: k_v[1][0], reverse=True)

    #Create a list for the data frame, with tuples of (name,total,details) 
    lst_for_df = []
    for tup in sorted_dict:
        name = tup[0]
        total = tup[1][0]
        details = tup[1][1:]
        lst_for_df.append((name,total,details))

    return lst_for_df

def export(lst_for_df, output_path):
    """
    Exports the analyzed list to an Excel file
    """
    df = pd.DataFrame(lst_for_df, columns=["שם","סך הכל","פירוט"])
    with pd.ExcelWriter(output_path) as writer:
        df.to_excel(writer)




if __name__ == "__main__":
    main(sys.argv[1:])