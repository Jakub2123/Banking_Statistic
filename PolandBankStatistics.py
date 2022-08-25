import requests
import pandas as pd
import urllib.request
import numpy as np

from io import BytesIO
from zipfile import ZipFile
from urllib.request import  urlopen

Url= 'https://www.nbp.pl/en/statystyka/oproc/mir_new/stopy_proc_en_srdw.xlsx'
urllib.request.urlretrieve(Url,'Stopy_proc.xlsx')

ZipUrl = 'https://www.nbp.pl/en/statystyka/czasowe_dwn/nalez_zobow_mif_en.zip'

with urlopen(ZipUrl) as zip:
    with ZipFile(BytesIO(zip.read())) as zfile:
        zfile.extractall()
    
    
print( "Downloading complete")

filename = r'Poland.xlsx'
previousSheet='Previous'
currentSheet='Latest'
startCol= 0
startRow= 15

dfSource =pd.read_excel(filename,sheet_name=currentSheet,index_col=0,nrows=12)

print('load source parameters complete')

for i, column in enumerate (dfSource.columns,1):
    
    print('Processing: ',column)
    
    f=dfSource[column]
    
#read file
    df=pd.read_excel(f['Website link'],sheet_name=f['Source tab'],index_col=0,header=f['Identyfier row']-1)

 #Saving Data
    data=pd.read_excel(filename,sheet_name=currentSheet)
    with pd.ExcelWriter(filename) as writer:
        data.to_excel(writer,sheet_name=previousSheet,index=None)
        data.to_excel(writer,sheet_name=currentSheet,index=None)
        df[f['Source identyfier']].to_excel(writer,sheet_name=currentSheet,header=False,index=False,startcol=i,startrow=f['Start row'])


print('DONE!')
