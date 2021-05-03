import numpy as np
import xlrd
import openpyxl
import pandas as pd
import matplotlib as plt
import seaborn as sns
%matplotlib notebook



dataset1=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='Detail_67_1_1')
dataset2=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='Detail_67_1_1_1')
dataset3=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='Detail_67_1_1_2')
dataset4=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='Detail_67_1_1_3')
dataset5=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='Detail_67_1_1_4')
dataset6=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='Detail_67_1_1_5')
dataset7=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='Detail_67_1_1_6')

dataset.info()
dataset1.info()

#dataset.head()
d=pd.concat([dataset1, dataset2, dataset3, dataset4, dataset5, dataset6, dataset7])
d

d.to_csv (r'C:\Users\HP\Desktop\detail.csv', index = False, header=True)


vol1=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='DetailVol_67_1_1')
vol2=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='DetailVol_67_1_1_1')
vol3=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='DetailVol_67_1_1_2')
vol4=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='DetailVol_67_1_1_3')
vol5=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='DetailVol_67_1_1_4')
vol6=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='DetailVol_67_1_1_5')
vol7=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='DetailVol_67_1_1_6')

detailVol=pd.concat([vol1, vol2, vol3, vol4, vol5, vol6, vol7])
detailVol
detailVol.to_csv (r'C:\Users\HP\Desktop\detailVol.csv', index = False, header=True)

temp1=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='DetailTemp_67_1_1')
temp2=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='DetailTemp_67_1_1_1')
temp3=pd.read_excel(r'C:\Users\HP\Desktop\data.xlsx',sheet_name='DetailTemp_67_1_1_2')
temp4=pd.read_excel(r'C:\Users\HP\Desktop\data_1.xlsx',sheet_name='DetailTemp_67_1_1_3')
temp5=pd.read_excel(r'C:\Users\HP\Desktop\data_1.xlsx',sheet_name='DetailTemp_67_1_1_4')
temp6=pd.read_excel(r'C:\Users\HP\Desktop\data_1.xlsx',sheet_name='DetailTemp_67_1_1_5')
temp7=pd.read_excel(r'C:\Users\HP\Desktop\data_1.xlsx',sheet_name='DetailTemp_67_1_1_6')

detailTemp=pd.concat([temp1, temp2, temp3, temp4, temp5, temp6, temp7])

detailTemp.to_csv (r'C:\Users\HP\Desktop\detailTemp.csv', index = False, header=True)
detailTemp


df=pd.read_csv('C:\\Users\\HP\\Desktop\\detailTemp.csv',parse_dates=['Realtime'])
df
#type(df.Relative Time(h:min:s.ms))
detailTempDownsampled=df.resample('1Min',on='Realtime').mean()

#detailTempDownsampled.to_csv (r'C:\Users\HP\Desktop\detailTempDownsampled.csv', header=True)
detailTempDownsampled

df1=pd.read_csv('C:\\Users\\HP\\Desktop\\detailVol.csv',index_col=['Record ID'],parse_dates=['Realtime'])
df1

detailVolDownsampled=df1.resample('1Min',on='Realtime').mean()
#df_x.resample('5Min').agg({'price': 'mean', 'vol': 'sum'})
detailVolDownsampled.to_csv (r'C:\Users\HP\Desktop\detailVolDownsampled.csv', header=True)
detailVolDownsampled

df2=pd.read_csv('C:\\Users\\HP\\Desktop\\detail.csv',index_col=['Record Index'],parse_dates=['Absolute Time'])
df2

detailDownsampled=df2.resample('1Min',on='Absolute Time').mean()
#df_x.resample('5Min').agg({'price': 'mean', 'vol': 'sum'})
detailDownsampled.to_csv (r'C:\Users\HP\Desktop\detailDownsampled.csv', header=True)
detailDownsampled
