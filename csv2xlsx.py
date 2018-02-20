# -*- coding: utf-8 -*-
"""
Created on Mon Feb 19 22:57:18 2018

@author: Tyler
"""
import argparse
import pandas as pd 
import pandas.io.formats.excel

def tab_xyz(df): #TODO Implement when needed
    df = pd.DataFrame(columns=['X','Y','Z','miniD'])
    return df
 
def tab_types(df):
    cols = ['ID','Name','CellNumber']
    df = df.tail(1).transpose() #take last row and transpose
    df = df.reset_index()
    df = df.reindex(df.index.rename(cols[0])) #add new index, rename
    df.rename(columns={'index': cols[1], len(df):cols[2]}, inplace=True) #rename columns
    df[cols[2]] = df[cols[2]].astype(int) #format cellnumbers as numbers
    return df

def tab_groups(df):
    return tab_types(df)

def tab_internalconns(df):
    start = '['
    close = ']'
    df = df[:-1]#remove last line
    df.insert(0,'', df.columns.values) #put first row as first column too
    for idx,row in enumerate(df.values): #for each row
        for i in range(1,len(row)): #in each column
            vars = row[i].replace(start, "").replace(close, "").split()
            str = "{},{},5,0,0,-10,2,{},0,1,2,4"
            row[i] = str.format(float(vars[0]),(idx+1)*100, int(vars[1]))
    return df

def main(file = 'input.csv', out = 'output.xlsx'):  
    pandas.io.formats.excel.header_style = None #Don't style headers in excel output
    writer = pd.ExcelWriter(out)
    df = pd.read_csv(file)
    
    tab_xyz(df.copy()).to_excel(writer,'XYZ')
    tab_types(df.copy()).to_excel(writer,'Types')
    tab_groups(df.copy()).to_excel(writer,'Groups')
    tab_internalconns(df.copy()).to_excel(writer,'InternalConns', index=False)
    
    writer.save()
    
    
parser = argparse.ArgumentParser(description='Convert CI model csv to xlsx.')
parser.add_argument('input', metavar='input',
                   help='Input csv file')
parser.add_argument('output', metavar='output',
                   help='Output xlsx file')
args = parser.parse_args()

main(file = args.input, out = args.output)