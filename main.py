import argparse
from typing import Any
import pandas as pd
import re
from dataclasses import dataclass

@dataclass(init=True)
class phage:
    year: int
    name: str
    lat: float
    long: float
    host: str
    azureus: Any = None
    coelicolor: Any = None
    distatochomrogenes: Any = None
    griseus: Any = None
    mirabilis: Any = None
    scabiei: Any = None

    """def __init__(self, year, name, coords, host, azureus, coelicolor, distatochomrogenes, griseus, mirabilis, scabiei):
        self.year = year
        self.name = name
        self.coords = coords
        self.host = host
        self.azureus = azureus
        self.coelicolor = coelicolor
        self.distatochomrogenes = distatochomrogenes
        self.griseus = griseus
        self.mirabilis = mirabilis
        self.scabiei = scabiei"""

def phage_from_df_row(year, row):
    print(row["Phage Name"])
    lat = 0
    long = 0
    host = "Streptomyces"
    phage(year, row["Phage Name"], lat, long, host)



def main(args = []):
    for file in args.inputFiles:
        year = re.search(r"([0-9]{4})", file.name)
        if (year == None): 
            print (f'Year could not be found in file name {file.name}')
        else:
            year = year.group(0)
        sheet_df = pd.read_excel(file.name, engine='openpyxl', header=1)
        sheet_df = sheet_df[sheet_df['#'] != "ex."]
        sheet_df = sheet_df[sheet_df['Phage Name'].notna()]
        sheet_df.drop(sheet_df.tail(1).index,inplace=True)
        sheet_df.reset_index(inplace=True)

        for index, row in sheet_df.iterrows():
            phage_from_df_row(year, row)
        print()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
                    prog = 'phage-processor',
                    description = 'Processes UMBC Phage Excel Records')
    parser.add_argument('inputFiles', type=argparse.FileType('r'), nargs='+')
    parser.add_argument("-o", "--output", help="Directs the output to a name of your choice")
    args = parser.parse_args()
    main(args)