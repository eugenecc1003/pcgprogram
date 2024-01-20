import pandas as pd
import numpy as np
# time
import time
from time import strftime
from datetime import datetime, timedelta
import os


def zo13_vlookup(userfoldertimepath):

    dfzo13 = pd.read_excel(userfoldertimepath +
                           "/zo13rawdata.xlsx", dtype=str).fillna('')
    dfcuinit = pd.read_excel(userfoldertimepath +
                             "/cudata.xlsx", dtype=str).fillna('')
    dfcu = dfcuinit[dfcuinit.columns[[6, 10, 7, 9]]].copy()
    dfcu = dfcuinit[dfcuinit.columns[[6, 10, 7, 9, 2]]].copy()
    dfcu[dfcu.columns[0]] = dfcu[dfcu.columns[0]].str.strip()
    dfcu[dfcu.columns[0]] = dfcu[dfcu.columns[0]].apply(
        lambda x: x.split('-V')[0] if "-V" in x else x)
    dfcu[dfcu.columns[0]] = dfcu[dfcu.columns[0]].apply(
        lambda x: x[:11] if len(x) == 14 else x)
    dfcu[dfcu.columns[0]] = dfcu[dfcu.columns[0]].apply(
        lambda x: x[:15] if len(x) == 18 else x)
    dfcu[dfcu.columns[4]] = dfcu[dfcu.columns[4]].str.strip()
    dfcu = dfcu.drop_duplicates(
        keep='first', inplace=False).reset_index(drop=True).copy()
    df1 = dfzo13.copy()
    df1['GA'] = df1[df1.columns[5]]
    cdtGA1 = df1[df1.columns[5]].str.len() == 14
    cdtGA2 = df1[df1.columns[5]].str.len() == 18
    df1.loc[cdtGA1, 'GA'] = df1[df1.columns[5]].str[:11]
    df1.loc[cdtGA2, 'GA'] = df1[df1.columns[5]].str[:15]
    dfmerge = pd.merge(df1, dfcu, left_on='GA',
                       right_on=dfcu.columns[4], how='left', sort=False)

    dfexport = dfmerge[dfmerge.columns[:14]].copy()
    dfexport = dfexport[dfexport.columns[[
        0, 1, 2, 3, 4, 5, 6, 7, 8, 11, 13, 12]]]
    dfexport.columns = ['採購單號碼', '採購單日期', '請求交貨日期', '到貨港', '出貨類型', '物料號碼',
                        '訂單數量', 'SEGMENT', '客戶號碼', 'Hs Code', '內文-Description of Goods', '內文-Measurement']
    dfexport['內文-Specification'] = ''
    dfexport['內文-Remark'] = ''
    tm = (datetime.now()+timedelta(hours=7)).strftime("%Y%m%d_%H%M%S")
    outputfilename = "ZO13new_"+tm+".xlsx"
    write = pd.ExcelWriter(userfoldertimepath+"/"+outputfilename)
    dfexport.to_excel(write, sheet_name='ZO13data',
                      index=False, freeze_panes=(1, 0))
    write.close()
    return outputfilename


def ZO13_transfer(userfoldertimepath, org, channel, fty):
    df1 = pd.read_excel(userfoldertimepath +
                        "/zo13transfer.xlsx", dtype=str).fillna('')
    df2 = df1.copy()
    df2['PURCH_NO_C'] = df2[df2.columns[0]]
    df2['PURCH_DATE'] = df2[df2.columns[1]]
    df2['ZZARV_PORT'] = df2[df2.columns[3]]
    df2['SALES_ORG'] = org
    df2['DISTR_CHAN'] = channel
    df2['DIVISION'] = '13'
    df2['REQ_DATE_H'] = df2[df2.columns[2]]
    df2['DOC_DATE'] = strftime("%Y%m%d", time.localtime())
    df2['PRICE_DATE'] = strftime("%Y%m%d", time.localtime())
    df2['DOC_TYPE'] = 'ZO13'
    df2['SHIP_TYPE'] = df2[df2.columns[4]]
    df2['PLANT'] = fty
    df2['SGT_RCAT_ITM'] = df2[df2.columns[7]]
    df2['MATNR'] = df2[df2.columns[5]]
    POlist = list(set(df2['PURCH_NO_C'].tolist()))
    for i in range(len(POlist)):
        check1 = df2['PURCH_NO_C'] == POlist[i]
        MTRlist = list(set(df2[check1]['MATNR'].tolist()))
        for j in range(len(MTRlist)):
            check2 = df2['MATNR'] == MTRlist[j]
            df2.loc[check1 & check2, 'ZSO_ITEM'] = str(
                MTRlist.index(MTRlist[j]) + 1).zfill(6)
    df2['qty'] = df2[df2.columns[6]].astype(float)
    df2['PO_QUAN'] = df2.groupby(['PURCH_NO_C', 'SGT_RCAT_ITM', 'MATNR'])[
        'qty'].transform('sum')
    df2['ZZSTAWN'] = df2[df2.columns[9]]
    df2['PARTN_ROLE'] = 'WE'
    df2['PARTN_NUM_CUST'] = df2[df2.columns[8]]
    df2['LANGU'] = 'E'
    df2['TEXT_ID'] = 'Z070'
    df2['1'] = df2[df2.columns[10]]
    df2['2'] = df2[df2.columns[11]]
    df2['3'] = df2[df2.columns[12]]
    df2['4'] = df2[df2.columns[13]]
    sheetAcol = ['PURCH_NO_C', 'PURCH_DATE', 'DOC_TYPE', 'SALES_ORG', 'DISTR_CHAN',
                 'DIVISION', 'REQ_DATE_H', 'DOC_DATE', 'PRICE_DATE', 'SHIP_TYPE', 'ZZARV_PORT', 'PLANT']
    sheetBcol = ['PURCH_NO_C', 'ZSO_ITEM', 'MATNR',
                 'SGT_RCAT_ITM', 'PO_QUAN', 'ZZSTAWN']
    sheetCcol = ['PURCH_NO_C', 'PARTN_ROLE', 'PARTN_NUM_CUST']
    sheetDcol = ['PURCH_NO_C', 'ZSO_ITEM',
                 'PURCH_DATE', 'LANGU', 'TEXT_ID', 'TEXT_LINE']
    sheetA = df2[sheetAcol].copy()
    sheetA.drop_duplicates(inplace=True, keep='first')
    sheetB = df2[sheetBcol].copy()
    sheetB.drop_duplicates(inplace=True, keep='first')
    sheetC = df2[sheetCcol].copy()
    sheetC.drop_duplicates(inplace=True, keep='first')
    sheetD = pd.DataFrame(columns=sheetDcol)
    contentcollist = ['1', '2', '3', '4']
    contentlist = ['[Description of Goods]:',
                   '[Measurement]:', '[Specification]:', '[Remark]:']
    for i in range(4):
        sheetD1 = df2[['PURCH_NO_C', 'ZSO_ITEM', 'PURCH_DATE',
                       'LANGU', 'TEXT_ID', contentcollist[i]]].copy()
        sheetD1[contentcollist[i]] = contentlist[i] + \
            sheetD1[contentcollist[i]]
        sheetD1.columns = sheetDcol
        sheetD = pd.concat([sheetD, sheetD1], ignore_index=True).fillna('')
    sheetD.drop_duplicates(subset=sheetDcol, inplace=True, keep='first')
    sheetD = sheetD.sort_values(sheetDcol, ascending=True)
    sheetD = sheetD.reset_index(drop=True)
    sheetAblank = pd.DataFrame(
        np.full((3, len(sheetAcol)), ''), columns=sheetAcol)
    sheetBblank = pd.DataFrame(
        np.full((3, len(sheetBcol)), ''), columns=sheetBcol)
    sheetCblank = pd.DataFrame(
        np.full((3, len(sheetCcol)), ''), columns=sheetCcol)
    sheetDblank = pd.DataFrame(
        np.full((3, len(sheetDcol)), ''), columns=sheetDcol)
    sheetAnew = pd.concat([sheetAblank, sheetA])
    sheetBnew = pd.concat([sheetBblank, sheetB])
    sheetCnew = pd.concat([sheetCblank, sheetC])
    sheetDnew = pd.concat([sheetDblank, sheetD])
    tm = (datetime.now()+timedelta(hours=7)).strftime("%Y%m%d_%H%M%S")
    outputfilename = "ZO13to6sheet_"+tm+".xlsx"
    write = pd.ExcelWriter(userfoldertimepath+"/"+outputfilename)
    sheetAnew.to_excel(write, sheet_name='ZSDTSD001A',
                       index=False, freeze_panes=(1, 0))
    sheetBnew.to_excel(write, sheet_name='ZSDTSD001B',
                       index=False, freeze_panes=(1, 0))
    sheetCnew.to_excel(write, sheet_name='ZSDTSD001C',
                       index=False, freeze_panes=(1, 0))
    sheetDnew.to_excel(write, sheet_name='ZSDTSD001D',
                       index=False, freeze_panes=(1, 0))
    write.close()
    return outputfilename
