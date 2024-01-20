
import pandas as pd
import numpy as np
# time
import time
from time import strftime
from datetime import datetime, timedelta
import os


def barcode_29(userfoldertimepath):
    df1 = pd.read_excel(userfoldertimepath+"/BarcodeFromAPL_29.xlsx")
    df2 = df1.copy()
    df2.columns = ['Barcode', '1', 'cus1', 'cus2',
                   'size', 'Qty', 'SAPcusPO', '2', '3', '4']
    df2['SAPcusArticle'] = df2['cus1'] + '-' + df2['cus2']
    df2['SAPsize'] = 'US_' + df2['size'].astype(str)
    df2['CusCTN'] = df2['Barcode']
    df2['CTNbarcode'] = df2['Barcode']
    df3 = df2[['CusCTN', 'SAPcusPO', 'SAPcusArticle',
               'SAPsize', 'Qty', 'CTNbarcode']].copy()
    df3 = df3.sort_values(
        by=['CusCTN', 'SAPcusArticle', 'SAPsize', 'CTNbarcode'], ascending=True)

    tm = (datetime.now()+timedelta(hours=7)).strftime("%Y%m%d_%H%M%S")
    outputfilename = "BarcodeFromAPL_29_"+tm+".xlsx"
    write = pd.ExcelWriter(userfoldertimepath+"/"+outputfilename)
    df3.to_excel(write, sheet_name="BarcodeFromAPL_29",
                 index=False, freeze_panes=(1, 0))
    write.close()

    return outputfilename


def NWGW_29(userfoldertimepath, tolerance):
    P2list = []
    for file in os.scandir(userfoldertimepath):
        if "P2_29" in file.name:
            P2list.append(file.name)
    dfM18 = pd.read_excel(userfoldertimepath+"/M18_29.xlsx")
    dfP2 = pd.read_excel(userfoldertimepath+"/"+P2list[0])
    dfnew = pd.DataFrame(columns=dfM18.columns)
    for j in range(len(dfM18)):
        for i in range(dfM18['Item Start'][j], dfM18['Item End'][j] + 1):
            dfM18_1 = dfM18[j: j + 1].copy()
            dfM18_1['ctn No.'] = i
            dfnew = pd.concat([dfnew, dfM18_1], axis=0,
                              ignore_index=True, sort=False)
            dfnew['ctn No.'] = dfnew['ctn No.'].astype(int)
    dfP2_1 = pd.DataFrame(columns=dfP2.columns)
    for file in P2list:
        dfP2_2 = pd.read_excel(userfoldertimepath+"/"+file)
        dfP2_1 = pd.concat([dfP2_1, dfP2_2], axis=0,
                           ignore_index=True, sort=False)
    dfP2 = dfP2_1.sort_values(dfP2.columns[[1, 2]].tolist()).drop_duplicates()
    dfP2['SO No.'] = dfP2['Packing Plan'].str.slice(2, 10).astype(int)
    dfbarcode = pd.merge(dfnew, dfP2,
                         left_on=['Sales Document', 'Purchase Order Number',
                                  'Packing Plan No.', 'ctn No.'],
                         right_on=['SO No.', 'SOLD-TO PO',
                                   'Packing Plan', 'Carton No'],
                         how='left', sort=False)
    dfbarcode['Max'] = dfbarcode['Carton G.W'] * (1+int(tolerance)*0.01)
    dfbarcode['min'] = dfbarcode['Carton G.W'] * (1-int(tolerance)*0.01)
    dfbc = dfbarcode[['Sales Document',
                      'Carton Barcode', 'Carton G.W', 'Max', 'min']]
    dfbcnew = dfbc.sort_values(dfbc.columns[[0, 1]].tolist())
    tm = (datetime.now()+timedelta(hours=7)).strftime("%Y%m%d_%H%M%S")
    outputfilename = "NWGW_29_"+tm+".xlsx"
    write = pd.ExcelWriter(userfoldertimepath+"/"+outputfilename)
    dfbcnew.to_excel(write, sheet_name="NWGW_29",
                     index=False, freeze_panes=(1, 0))
    write.close()

    return outputfilename
