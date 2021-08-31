import pandas as pd 
import os
import numpy as np
import openpyxl
import warnings


#Shipment Order
with warnings.catch_warnings(record=True):
    warnings.simplefilter("always")
    SO = pd.read_excel('SO.xlsx', engine= 'openpyxl')
    SO ['Order No'] = SO ['OrderNO']
    SO ['salesorder_no'] = SO ['Invoice Ref Number'] 
    SO = SO[['SO Status','Carrier Name','WaveNO','Last Edit Time','Order No','salesorder_no']]
    SO = SO.drop_duplicates(subset=['salesorder_no'])

#Order Composition
with warnings.catch_warnings(record=True):
    warnings.simplefilter("always")
    OC = pd.read_excel('Order Composition.xlsx', engine='openpyxl')
    OC ['salesorder_no'] = OC ['Invoice No']
    OC ['Delivery No'] = OC ['Delivery ConfirmNO']
    OC = OC[['salesorder_no','Delivery No']]

#Merged SO + OC 
Compiled_SO_OC = SO.merge(OC, on = ['salesorder_no'], how='left')

#Siap Kirim
SK = pd.read_csv('Siap Kirim.csv')
SK = SK[['salesorder_no','transaction_date','shipper','source_name']]
SK['salesorder_no'] = SK['salesorder_no'].str.replace('LZ-','')
SK['salesorder_no'] = SK['salesorder_no'].str.replace('SP-','')
SK['salesorder_no'] = SK['salesorder_no'].str.replace('TP-','')
SK['salesorder_no'] = SK['salesorder_no'].str.replace('-24908','')


#Merged Siap Kirim 
Compiled_DF1 = SK.merge(Compiled_SO_OC, on =['salesorder_no'], how='left')
Compiled_DF1.to_excel('Merged Siap Kirim.xlsx', index=False)


#Siap Dikirim
SD = pd.read_csv('Siap Dikirim.csv')
SD = SD[['salesorder_no','transaction_date','shipper','source_name','status']]
SD['salesorder_no'] = SD['salesorder_no'].str.replace('LZ-','')
SD['salesorder_no'] = SD['salesorder_no'].str.replace('SP-','')
SD['salesorder_no'] = SD['salesorder_no'].str.replace('TP-','')
SD['salesorder_no'] = SD['salesorder_no'].str.replace('-24908','')
SD = SD.drop_duplicates(subset=['salesorder_no'])
wta = ['To Confirm Receive','Shipped','ORDER SHIPPING','ORDER CONFLICTED','Retry Ship']
SD = SD[SD['status'].apply(lambda txt: not any([wta in txt for wta in wta]))]

#Merged Siap Dikirim
Compiled_DF2 = SD.merge(Compiled_SO_OC, on=['salesorder_no'], how = 'left')
Compiled_DF2.to_excel('Merged Siap Dikirim.xlsx', index=False)




