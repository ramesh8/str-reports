import pandas as pd
import sys

file="04.04.2023"
if len(sys.argv) > 1:
    file = sys.argv[1]

filename = f"files/{file}.xlsx"

# get meta information 
xl = pd.ExcelFile(filename)
sheets = xl.sheet_names
# dfs = [xl.parse(sheet) for sheet in sheets]

config = {
    "skip_sheets":[0,6,7],
    "meta_rows":{
        1:[range(5),range(-3,0)],
        2:[range(5),range(-3,0)],
        3:[range(5),range(-3,0)],
        4:[range(5),range(-3,0)],
        5:[range(5),range(-3,0)],        
        },
    "remove_empty_rows" : True,
    "remove_empty_cols" : True,
    "split_dfs":{1:True, 2:True, 3:True, 4:True, 5:True},

}
dfs = []
meta_rows = [list for x in sheets]
for index, sheet in enumerate(sheets):
    if index in config["skip_sheets"]:
        # print(sheet , "skipping")
        continue
    df = xl.parse(sheet,header=None)
    meta_rows[index]=[]
    
    if index in config["meta_rows"].keys():
        for rows in config["meta_rows"][index]:
            meta_rows[index].append(df.iloc[rows])
            df.drop(df.index[rows], inplace=True)
    
    if config["remove_empty_rows"]!=None:
        df.dropna(axis='columns',how='all', inplace=True)
    if config["remove_empty_cols"]!=None:
        df.dropna(axis='rows',how='all', inplace=True)

    df.reset_index(drop=True, inplace=True)
    df.columns = range(df.columns.size)

    if index in config["split_dfs"].keys():
        df["nancnt"] = df.isnull().sum(axis=1)
        maxcnt = max(df["nancnt"])
        maxrows = df.index[df["nancnt"]>=maxcnt-1].to_list()        
        df.drop(columns=["nancnt"], inplace=True)
        for i in range(len(maxrows)):
            if i == 0:
                tdf = df.iloc[:maxrows[1],:]
            elif i==len(maxrows)-1:
                tdf = df.iloc[maxrows[i]:,:]
            else:
                tdf = df.iloc[maxrows[i]:maxrows[i+1],:]
            # print(tdf)
            dfs.append({"sheet":sheet, "index":index,"df":tdf})        
    else:
        dfs.append({"sheet":sheet, "index":index,"df":df})

for index, dfo in enumerate(dfs):
    print("saving sheet ", {dfo["sheet"]})
    dfo["df"].to_csv(f"output/{file}-{dfo['sheet']}-{index}.csv", index=None, header=None)

# print(meta_rows)