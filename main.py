import argparse
from typing import Any
import pandas as pd
import re
from dataclasses import dataclass, field
from geopy import distance as ds

def get_item(li, index, default=None):
    try:
        return li[index]
    except:
        return default

def convert(tude):
    tude = tude.strip()
    print(tude, re.match("\d+\.\d+", tude))
    multiplier = 1 if tude.find("N")!=-1 or tude.find("E")!=-1 else -1
    return multiplier * float(get_item(re.match("\d+\.\d+", tude), 0, 0))

@dataclass(init=True)
class phage:
    year: int
    name: str
    coords: str
    distance: str = field(init=False)
    location: str | None
    host: str
    azureus: Any = None
    coelicolor: Any = None
    distatochomrogenes: Any = None
    griseus: Any = None
    mirabilis: Any = None
    scabiei: Any = None

    def __post_init__(self):
        umbc_md = (39.2434636, -76.7139453)
        try:
            coord_location = tuple(map(convert, self.coords.split(",")))
            self.distance = (f'{ds.distance(umbc_md, coord_location).miles:.2f} miles')
        except:
            self.distance = None
        


def phage_from_df_row(year, row):
    location = get_item(row.filter(regex=re.compile("location", re.IGNORECASE)), 0)
    gps_coords = get_item(row.filter(regex=re.compile("GPS Coordinate", re.IGNORECASE)), 0)
    host = "Streptomyces" #To-do
    azureus = get_item(row.filter(regex=re.compile("azureus", re.IGNORECASE)), 0)
    coelicolor = get_item(row.filter(regex=re.compile("coelicolor", re.IGNORECASE)), 0)
    distatochomrogenes = get_item(row.filter(regex=re.compile("diastatochromogenes", re.IGNORECASE)), 0)
    griseus = get_item(row.filter(regex=re.compile("griseus", re.IGNORECASE)), 0)
    mirabilis = get_item(row.filter(regex=re.compile("mirabilis", re.IGNORECASE)), 0)
    scabiei = get_item(row.filter(regex=re.compile("scabiei", re.IGNORECASE)), 0)
    return phage(year, row["Phage Name"], gps_coords, location, host, azureus, coelicolor, distatochomrogenes, griseus, mirabilis, scabiei)



def main(args = []):
    phages = []
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
            phages.append(phage_from_df_row(year, row))
    phage_df = pd.DataFrame(phages)
    writer = pd.ExcelWriter(args.output, engine='xlsxwriter') 
    phage_df.to_excel(writer, sheet_name='phages', index=False, na_rep='None')

    for column in phage_df:
        column_length = max(phage_df[column].astype(str).map(len).max(), len(column))
        col_idx = phage_df.columns.get_loc(column)
        writer.sheets['phages'].set_column(col_idx, col_idx, column_length)

    writer.save()

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
                    prog = 'phage-processor',
                    description = 'Processes UMBC Phage Excel Records')
    parser.add_argument('inputFiles', type=argparse.FileType('r'), nargs='+')
    parser.add_argument("-o", "--output", default="master_sheet.xlsx", help="Directs the output to a name of your choice")
    args = parser.parse_args()
    main(args)