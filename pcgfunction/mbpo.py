import pandas as pd
import numpy as np
# time
import time
from time import strftime
from datetime import datetime, timedelta
import os


def MBPO_27(userfoldertimepath):
    filenamelist = os.listdir(userfoldertimepath)
    dfgtn = pd.read_excel(os.path.join(
        userfoldertimepath, filenamelist[0]), dtype={30: str})

    # A表

    # 客戶採購單號碼
    dfgtn['PURCH_NO_C'] = dfgtn[dfgtn.columns[1]].astype(str)  # PO No.
    # GANO
    # 客戶採購單的號碼項目
    dfgtn['ZSO_GANO'] = dfgtn[dfgtn.columns[6]] + \
        "-" + dfgtn[dfgtn.columns[18]].str[11:15]
    dfgtn['PO_ITM_NO'] = ""
    POlist = list(set(dfgtn['PURCH_NO_C'].tolist()))
    for PO in POlist:
        check1 = dfgtn['PURCH_NO_C'] == PO
        GANOlist = list(set(dfgtn[check1]['ZSO_GANO'].tolist()))
        for GANO in GANOlist:
            check2 = dfgtn['ZSO_GANO'] == GANO
            dfgtn.loc[check1 & check2, 'PO_ITM_NO'] = str(
                GANOlist.index(GANO) + 1).zfill(6)
    # 採購單日期
    dfgtn['PURCH_DATE'] = dfgtn[dfgtn.columns[2]].dt.strftime("%Y%m%d")
    # 收貨人採購單號碼
    dfgtn['PURCH_NO_S'] = dfgtn[dfgtn.columns[3]]
    # 終端客戶採購單號碼
    dfgtn['ZZZCBSTKD_E'] = dfgtn[dfgtn.columns[21]].str[3:]
    # 銷售組織
    dfgtn['SALES_ORG'] = 4020
    # 配銷通路
    dfgtn['DISTR_CHAN'] = 27
    # 部門
    dfgtn['DIVISION'] = 10
    # 品牌請求交貨日期
    dfgtn['REQ_DATE_H'] = dfgtn[dfgtn.columns[16]].dt.strftime("%Y%m%d")
    # 文件日期
    dfgtn['DOC_DATE'] = strftime("%Y%m%d", time.localtime())
    # 銷售文件類型
    dfgtn['DOC_TYPE'] = 'ZO10'
    # 定價及匯率日期
    dfgtn['PRICE_DATE'] = dfgtn['PURCH_DATE']
    # 出貨類型
    dfgtn['SHIP_TYPE'] = ""
    cdtocean = dfgtn[dfgtn.columns[4]].str.lower() == "ocean"
    cdtair = dfgtn[dfgtn.columns[4]].str.lower() == "air"
    cdtexpress = dfgtn[dfgtn.columns[4]].str.lower() == "express"
    dfgtn.loc[cdtocean, 'SHIP_TYPE'] = "21"
    dfgtn.loc[cdtair, 'SHIP_TYPE'] = "11"
    dfgtn.loc[cdtexpress, 'SHIP_TYPE'] = "51"
    # 工廠
    dfgtn['PLANT'] = '612A'
    # 需求類型
    dfgtn['SGT_RCAT'] = 'FLT'
    # 生產年月
    dfgtn['ZZPRD_PERIOD'] = dfgtn[dfgtn.columns[16]].dt.strftime("%Y%m")
    # 包裝方式
    dfgtn['ZZMIXPACK'] = '10'
    # Brand Factory Name
    dfgtn['ZZART_PLANTNAME'] = 'PLV'
    # Variant Value Conversion Type
    dfgtn['FSH_SC_VCTYP'] = 'US'
    # 數量
    dfgtn['ZDOUBLE'] = '6'
    # 品牌工廠代號
    dfgtn['ZZART_PLANT'] = '61374'

    # B表

    # 客戶採購單日期
    dfgtn['PURCH_DATE_ITM'] = dfgtn['PURCH_DATE']
    dfgtn['WRF_CHARSTC2'] = "US_" + dfgtn[dfgtn.columns[10]].astype(str)
    # 物料號碼
    dfgtn['MATNR'] = ''
    # 客戶料號
    dfgtn['KDMAT'] = dfgtn[dfgtn.columns[6]]
    # 訂單數量
    dfgtn['PO_QUAN'] = dfgtn[dfgtn.columns[11]]
    # 訂單數量單位
    dfgtn['PO_UNIT'] = 'PAA'
    # 條件價格
    dfgtn['KSCHL_EDI2_ITM'] = dfgtn[dfgtn.columns[12]]
    # SD 文件幣別
    dfgtn['WAERK_EDI2_ITM'] = dfgtn[dfgtn.columns[13]]
    # 報價和銷售訂單的拒收原因
    dfgtn['REASON_REJ'] = ''
    # Material Type
    dfgtn['ZZMTART'] = 'FG'
    # 外貿的商品代碼/進口代號
    dfgtn['ZZSTAWN'] = dfgtn[dfgtn.columns[30]]

    # C表

    dfgtn['PARTN_ROLE'] = 'WE'
    dfgtn['PARTN_SRT_CUST'] = dfgtn[dfgtn.columns[9]].astype(str)

    # export

    sheetAcol = ['PURCH_NO_C', 'PO_ITM_NO', 'PURCH_DATE', 'PURCH_NO_S', 'ZZZCBSTKD_E', 'SALES_ORG', 'DISTR_CHAN', 'DIVISION', 'REQ_DATE_H', 'DOC_DATE',
                 'DOC_TYPE', 'PRICE_DATE', 'SHIP_TYPE', 'PLANT', 'SGT_RCAT', 'ZZPRD_PERIOD', 'ZZMIXPACK', 'ZZART_PLANTNAME', 'FSH_SC_VCTYP', 'ZDOUBLE', 'ZZART_PLANT']
    sheetBcol = ['PURCH_NO_C', 'PO_ITM_NO', 'ZSO_GANO', 'PURCH_DATE_ITM', 'WRF_CHARSTC2', 'MATNR', 'KDMAT',
                 'PO_QUAN', 'PO_UNIT', 'KSCHL_EDI2_ITM', 'WAERK_EDI2_ITM', 'REASON_REJ', 'ZZMTART', 'ZZSTAWN']
    sheetCcol = ['PURCH_NO_C', 'PO_ITM_NO', 'PURCH_DATE', 'PARTN_ROLE']
    sheetA = dfgtn[sheetAcol].copy()
    sheetA.drop_duplicates(inplace=True, keep='first')
    sheetB = dfgtn[sheetBcol].copy()
    sheetB.drop_duplicates(inplace=True, keep='first')
    sheetC = dfgtn[sheetCcol].copy()
    sheetC.drop_duplicates(inplace=True, keep='first')
    sheetAblank = pd.DataFrame(
        np.full((3, len(sheetAcol)), ''), columns=sheetAcol)
    sheetBblank = pd.DataFrame(
        np.full((3, len(sheetBcol)), ''), columns=sheetBcol)
    sheetCblank = pd.DataFrame(
        np.full((3, len(sheetCcol)), ''), columns=sheetCcol)
    sheetAnew = pd.concat([sheetAblank, sheetA])
    sheetBnew = pd.concat([sheetBblank, sheetB])
    sheetCnew = pd.concat([sheetCblank, sheetC])
    sheetblank = pd.DataFrame()

    tm = (datetime.now()+timedelta(hours=7)).strftime("%Y%m%d_%H%M%S")
    outputfilename = "GTNto6sheet_Merrell_"+tm+".xlsx"
    write = pd.ExcelWriter(userfoldertimepath+"/"+outputfilename)

    sheetAnew.to_excel(write, sheet_name='ZSDTSD001A',
                       index=False, freeze_panes=(1, 0))
    sheetBnew.to_excel(write, sheet_name='ZSDTSD001B',
                       index=False, freeze_panes=(1, 0))
    sheetCnew.to_excel(write, sheet_name='ZSDTSD001C',
                       index=False, freeze_panes=(1, 0))
    sheetblank.to_excel(write, sheet_name='ZSDTSD001D',
                        index=False, freeze_panes=(1, 0))
    sheetblank.to_excel(write, sheet_name='ZSDTSD001E29',
                        index=False, freeze_panes=(1, 0))
    sheetblank.to_excel(write, sheet_name='ZSDTSD001F',
                        index=False, freeze_panes=(1, 0))
    # write._save()
    write.close()

    return outputfilename


def MBPO_29(userfoldertimepath, option):
    dfxpc = pd.read_excel(userfoldertimepath +
                          '/APL29.xlsx', dtype={'Size': str})
    dfpar = pd.read_excel(userfoldertimepath+'/PF29.xlsx')
    dfcus = pd.read_excel(userfoldertimepath+'/CC29.xlsx')
    size = [['01', 'US_1.0'], ['015', 'US_1.5'], ['02', 'US_2.0'], ['025', 'US_2.5'], ['03', 'US_3.0'], ['035', 'US_3.5'], ['04', 'US_4.0'], ['045', 'US_4.5'], ['05', 'US_5.0'], ['055', 'US_5.5'], ['06', 'US_6.0'], ['065', 'US_6.5'], ['07', 'US_7.0'], ['075', 'US_7.5'], ['08', 'US_8.0'], ['085', 'US_8.5'], ['09', 'US_9.0'], ['095', 'US_9.5'], ['10', 'US_10.0'], [
        '105', 'US_10.5'], ['11', 'US_11.0'], ['115', 'US_11.5'], ['12', 'US_12.0'], ['125', 'US_12.5'], ['13', 'US_13.0'], ['14', 'US_14.0'], ['15', 'US_15.0'], ['16', 'US_16.0'], ['17', 'US_17.0'], ['18', 'US_18.0'], ['105K', 'US_10.5K'], ['11K', 'US_11.0K'], ['115K', 'US_11.5K'], ['12K', 'US_12.0K'], ['125K', 'US_12.5K'], ['13K', 'US_13.0K'], ['135K', 'US_13.5K']]
    dfsize = pd.DataFrame(size, columns=['size1', 'size2'])
    dfxpc1 = dfxpc.copy()
    dfpar1 = dfpar.copy()
    dfpar1.columns = ['brand', 'partner', 'ship-to', 'consigneeZB',
                      'notifyZD', 'notifyZE', 'notifyZF', 'forwarderSP', 'File Nm']
    dfcus1 = dfcus.copy()
    dfcus1.columns = ['brand', 'customercode', 'market', 'customername',
                      'Attribute3', 'Attribute4', 'Attribute5', 'shiptoData', 'File Nm']
    dfxpc2 = pd.merge(dfxpc1, dfsize,
                      left_on=['Size'],
                      right_on=['size1'],
                      how='left', sort=False)
    dfxpc3 = dfxpc2.copy()
    df1 = pd.merge(dfxpc3, dfcus1,
                   left_on=['Market', 'Customer'],
                   right_on=['market', 'customername'],
                   how='left', sort=False)
    df2 = pd.merge(df1, dfpar1,
                   left_on=['shiptoData'],
                   right_on=['ship-to'],
                   how='left', sort=False)
    df3 = df2.copy()
    # A
    # 客戶採購單號碼
    df3['PURCH_NO_C'] = df3[df3.columns[7]].astype(str)  # PO No.
    # 客戶採購單的號碼項目
    purlist = list(set(df3['PURCH_NO_C'].tolist()))
    df3['PO_ITM_NO'] = ''
    # for i in range(len(df3)): df3['PO_ITM_NO'][i] = str(purlist.index(str(df3['PURCH_NO_C'][i])) + 1).zfill(6)
    # 採購單日期
    df3['PURCH_DATE'] = strftime("%Y%m%d", time.localtime())
    # 收貨人採購單號碼
    df3['PURCH_NO_S'] = df3[df3.columns[8]]  # PO Reference No.
    # 終端客戶採購單號碼
    df3['ZZZCBSTKD_E'] = df3[df3.columns[9]]  # Customer Order No.
    # 銷售組織
    df3['SALES_ORG'] = 4020
    # 配銷通路
    df3['DISTR_CHAN'] = 29
    # 部門
    df3['DIVISION'] = 10
    # 品牌請求交貨日期
    df3['REQ_DATE_H'] = pd.to_datetime(
        df3[df3.columns[22]]).dt.strftime('%Y%m%d')  # Orig Req XFD
    # 文件日期
    df3['DOC_DATE'] = strftime("%Y%m%d", time.localtime())
    # 銷售文件類型
    df3['DOC_TYPE'] = 'ZO10'
    # 定價及匯率日期
    df3['PRICE_DATE'] = pd.to_datetime(
        df3[df3.columns[24]]).dt.strftime('%Y%m%d')  # PO Release Date
    # 出貨類型
    df3['SHIP_TYPE'] = '21'
    # 工廠
    df3['PLANT'] = '608A'
    # 需求類型
    df3['SGT_RCAT'] = 'FLT'
    # 客戶需求交貨日
    df3['ZZEND_CUST_REQ'] = pd.to_datetime(df3[df3.columns[34]]).dt.strftime(
        '%Y%m%d')  # Current Customer Target XFD
    # 包裝方式
    df3['ZZMIXPACK'] = '10'
    # VAS備註
    df3['ZZ13VASTAG'] = '10'
    # Brand Factory Name
    df3['ZZART_PLANTNAME'] = 'PYV'
    # Service Identifier
    df3['ZZSERV_IDNO'] = ''
    # 客戶採購單HS Code
    df3['ZZCUSHSCDE'] = ''
    # Variant Value Conversion Type
    df3['FSH_SC_VCTYP'] = 'US'

    # B
    # 客戶採購單日期
    df3['PURCH_DATE_ITM'] = df3['PURCH_DATE']
    # 尺碼
    df3['WRF_CHARSTC2'] = df3[df3.columns[36]]  # size2
    # for i in range(len(df3)):
    #     # 有半碼
    #     if str(df3['Size'][i])[-1:] == '5' and len(str(df3['Size'][i])) > 1 :
    #         df3['WRF_CHARSTC2'][i] = 'US_' + str(df3['Size'][i])[:-1] + '.5'
    #     else:
    #         df3['WRF_CHARSTC2'][i] = 'US_' + str(df3['Size'][i]) + '.0'
    # 物料號碼
    df3['MATNR'] = ''
    # 客戶料號
    df3['KDMAT'] = df3[df3.columns[13]] + '-' + \
        df3[df3.columns[16]]  # Style/Part No.  Color/Width
    # 訂單數量
    df3['PO_QUAN'] = df3[df3.columns[18]]  # 'Quantity'
    # 訂單數量單位
    df3['PO_UNIT'] = 'PAA'
    # 報價和銷售訂單的拒收原因
    df3['REASON_REJ'] = ''
    # 預告訂單數量更新方式
    df3['ZFCST_UPD_KIND'] = ''
    # 終端客戶採購單號碼
    df3['ZZZCBSTKD_E_ITM'] = ''
    # 收貨人採購單號碼
    df3['ZPURCH_NO_S_ITM'] = ''
    # 出貨地址代碼
    df3['ZDELV_ADDR_CODE'] = ''
    # 計劃決策群組
    df3['STRGR'] = ''
    # 客戶需求的需求類型
    df3['BEDKU'] = ''
    # 型體代號+顏色代號
    df3['ZZMDMARK'] = df3['KDMAT']
    # Material Type
    df3['ZZMTART'] = 'FG'
    # 楦頭肥度
    df3['ZWIDTH_ITM'] = df3[df3.columns[16]]  # Color/Width
    df3['ZSO_GANO'] = df3['KDMAT'] + '-' + df3[df3.columns[15]] + \
        '-Line' + df3[df3.columns[12]].astype(str)  # 'Ship Mode' 'Line No.'
    df3['ZSO_ITEM'] = ""  # df3.index + 1

    # 1SO1GA_1item or 1SO1GA_Mitem處理
    POlist = list(set(df3['PURCH_NO_C'].tolist()))
    for i in range(len(POlist)):
        check1 = df3['PURCH_NO_C'] == POlist[i]
        GANOlist = list(set(df3[check1]['ZSO_GANO'].tolist()))
        for j in range(len(GANOlist)):
            check2 = df3['ZSO_GANO'] == GANOlist[j]
            sizelist = list(set(df3[check1 & check2]['WRF_CHARSTC2'].tolist()))
            for k in range(len(sizelist)):
                check3 = df3['WRF_CHARSTC2'] == sizelist[k]
                df3.loc[check1 & check2 & check3, 'ZSO_ITEM'] = str(
                    sizelist.index(sizelist[k]) + 1).zfill(6)

    if option == '1SO1GA1item':
        for i in range(len(POlist)):
            check1 = df3['PURCH_NO_C'] == POlist[i]
            GANOlist = list(set(df3[check1]['ZSO_GANO'].tolist()))
            for j in range(len(GANOlist)):
                check2 = df3['ZSO_GANO'] == GANOlist[j]
                df3.loc[check1 & check2, 'PO_ITM_NO'] = str(
                    GANOlist.index(GANOlist[j]) + 1).zfill(6)
    else:
        for i in range(len(POlist)):
            check1 = df3['PURCH_NO_C'] == POlist[i]
            KDMATlist = list(set(df3[check1]['KDMAT'].tolist()))
            for j in range(len(KDMATlist)):
                check2 = df3['KDMAT'] == KDMATlist[j]
                df3.loc[check1 & check2, 'PO_ITM_NO'] = str(
                    KDMATlist.index(KDMATlist[j]) + 1).zfill(6)

    # C
    df4 = df3.copy()
    sheetCcolumns = ['PURCH_NO_C', 'PO_ITM_NO', 'ZSO_ITEM', 'PURCH_DATE', 'PARTN_ROLE', 'PARTN_NUM_CUST',
                     'PARTN_SRT_CUST', 'PARTN_NUM_VEND', 'PARTN_SRT_VEND', 'PARTN_SRT_NAME', 'PARTN_EXT_CUST', 'PARTN_EXT_VEND']
    sheetCcombine = pd.DataFrame(columns=sheetCcolumns)
    dfCWE = df4[['PURCH_NO_C', 'PO_ITM_NO',
                 'ZSO_ITEM', 'PURCH_DATE', 'ship-to']].copy()
    dfCWE['PARTN_SRT_CUST'] = dfCWE['ship-to']
    dfCWE['PARTN_ROLE'] = 'WE'
    dfmergeWE = pd.concat([sheetCcombine, dfCWE], ignore_index=True)
    dfCZB = df4[['PURCH_NO_C', 'PO_ITM_NO', 'ZSO_ITEM',
                 'PURCH_DATE', 'consigneeZB']].copy()
    dfCZB['PARTN_SRT_CUST'] = dfCZB['consigneeZB']
    dfCZB['PARTN_ROLE'] = 'ZB'
    dfmergeWEZB = pd.concat([dfmergeWE, dfCZB], ignore_index=True)
    dfCZD = df4[['PURCH_NO_C', 'PO_ITM_NO',
                 'ZSO_ITEM', 'PURCH_DATE', 'notifyZD']].copy()
    dfCZD['PARTN_SRT_CUST'] = dfCZD['notifyZD']
    dfCZD['PARTN_ROLE'] = 'ZD'
    dfmergeWEZBZD = pd.concat([dfmergeWEZB, dfCZD], ignore_index=True)
    dfCZE = df4[['PURCH_NO_C', 'PO_ITM_NO',
                 'ZSO_ITEM', 'PURCH_DATE', 'notifyZE']].copy()
    dfCZE['PARTN_SRT_CUST'] = dfCZE['notifyZE']
    dfCZE['PARTN_ROLE'] = 'ZE'
    dfmergeWEZBZDZE = pd.concat([dfmergeWEZBZD, dfCZE], ignore_index=True)
    dfCZF = df4[['PURCH_NO_C', 'PO_ITM_NO',
                 'ZSO_ITEM', 'PURCH_DATE', 'notifyZF']].copy()
    dfCZF['PARTN_SRT_CUST'] = dfCZF['notifyZF']
    dfCZF['PARTN_ROLE'] = 'ZF'
    dfmergeWEZBZDZEZF = pd.concat([dfmergeWEZBZDZE, dfCZF], ignore_index=True)
    dfCSP = df4[['PURCH_NO_C', 'PO_ITM_NO', 'ZSO_ITEM',
                 'PURCH_DATE', 'forwarderSP']].copy()
    dfCSP['PARTN_NUM_VEND'] = dfCSP['forwarderSP']
    dfCSP['PARTN_ROLE'] = 'SP'
    dfmergeWEZBZDZEZFSP = pd.concat(
        [dfmergeWEZBZDZEZF, dfCSP], ignore_index=True)

    cdt1 = dfmergeWEZBZDZEZFSP['PARTN_SRT_CUST'].astype(str) != 'nan'
    cdt2 = dfmergeWEZBZDZEZFSP['PARTN_NUM_VEND'].astype(str) != 'nan'
    dfpartner = dfmergeWEZBZDZEZFSP[cdt1 | cdt2].copy()
    dfpartnererror = dfmergeWEZBZDZEZFSP[(dfmergeWEZBZDZEZFSP['PARTN_ROLE'].astype(str) == 'WE') &
                                         (dfmergeWEZBZDZEZFSP['PARTN_SRT_CUST'].astype(str) == 'nan')].copy()

    sheetAcol = ['PURCH_NO_C', 'PO_ITM_NO', 'PURCH_DATE', 'PURCH_NO_S', 'ZZZCBSTKD_E', 'SALES_ORG', 'DISTR_CHAN', 'DIVISION', 'REQ_DATE_H', 'DOC_DATE', 'DOC_TYPE',
                 'PRICE_DATE', 'SHIP_TYPE', 'PLANT', 'SGT_RCAT', 'ZZEND_CUST_REQ', 'ZZMIXPACK', 'ZZ13VASTAG', 'ZZART_PLANTNAME', 'ZZSERV_IDNO', 'ZZCUSHSCDE', 'FSH_SC_VCTYP']
    sheetBcol = ['PURCH_NO_C', 'PO_ITM_NO', 'ZSO_GANO', 'ZSO_ITEM', 'PURCH_DATE_ITM', 'WRF_CHARSTC2', 'MATNR', 'KDMAT', 'PO_QUAN', 'PO_UNIT', 'SHIP_TYPE',
                 'REASON_REJ', 'ZFCST_UPD_KIND', 'ZZZCBSTKD_E_ITM', 'ZPURCH_NO_S_ITM', 'STRGR', 'BEDKU', 'ZZMDMARK', 'ZZMTART', 'ZWIDTH_ITM', 'ZDELV_ADDR_CODE']
    sheetCcol = ['PURCH_NO_C', 'PO_ITM_NO', 'ZSO_ITEM', 'PURCH_DATE', 'PARTN_ROLE', 'PARTN_NUM_CUST',
                 'PARTN_SRT_CUST', 'PARTN_NUM_VEND', 'PARTN_SRT_VEND', 'PARTN_SRT_NAME', 'PARTN_EXT_CUST', 'PARTN_EXT_VEND']

    sheetA = df4[sheetAcol].copy()
    sheetA.drop_duplicates(inplace=True, keep='first')
    sheetB = df4[sheetBcol].copy()
    sheetB.drop_duplicates(inplace=True, keep='first')
    sheetC = dfpartner[sheetCcol].copy()
    sheetC.drop_duplicates(inplace=True, keep='first')
    sheetAblank = pd.DataFrame(
        np.full((3, len(sheetAcol)), ''), columns=sheetAcol)
    sheetBblank = pd.DataFrame(
        np.full((3, len(sheetBcol)), ''), columns=sheetBcol)
    sheetCblank = pd.DataFrame(
        np.full((3, len(sheetCcol)), ''), columns=sheetCcol)
    sheetAnew = pd.concat([sheetAblank, sheetA])
    sheetBnew = pd.concat([sheetBblank, sheetB])
    sheetCnew = pd.concat([sheetCblank, sheetC])
    sheetblank = pd.DataFrame()

    tm = (datetime.now()+timedelta(hours=7)).strftime("%Y%m%d_%H%M%S")

    if option == '1SO1GAMitem':
        outputfilename = "XPCto6sheet1SO1GA_Mitem_"+tm+".xlsx"
    else:
        outputfilename = "XPCto6sheet1SO1GA_1item_"+tm+".xlsx"
    write = pd.ExcelWriter(userfoldertimepath + "/" + outputfilename)

    sheetAnew.to_excel(write, sheet_name='ZSDTSD001A',
                       index=False, freeze_panes=(1, 0))
    sheetBnew.to_excel(write, sheet_name='ZSDTSD001B',
                       index=False, freeze_panes=(1, 0))
    sheetCnew.to_excel(write, sheet_name='ZSDTSD001C',
                       index=False, freeze_panes=(1, 0))
    sheetblank.to_excel(write, sheet_name='ZSDTSD001D',
                        index=False, freeze_panes=(1, 0))
    sheetblank.to_excel(write, sheet_name='ZSDTSD001E29',
                        index=False, freeze_panes=(1, 0))
    sheetblank.to_excel(write, sheet_name='ZSDTSD001F',
                        index=False, freeze_panes=(1, 0))
    if len(dfpartnererror) != 0:
        dfpartnererror.to_excel(
            write, sheet_name='withoutshipto', index=False, freeze_panes=(1, 0))
    write.close()

    return outputfilename


def MBPO_38(userfoldertimepath):
    dfgtn = pd.read_excel(userfoldertimepath + '/GTN38.xlsx', dtype=str)
    dfpar = pd.read_excel(userfoldertimepath+'/PF38.xlsx', dtype=str)
    dfcus = pd.read_excel(userfoldertimepath+'/CC38.xlsx', dtype=str)
    dfgtn1 = dfgtn.copy()
    dfpar1 = dfpar.copy()
    dfpar1.columns = ['brand', 'partner', 'ship-to', 'consigneeZB',
                      'notifyZD', 'notifyZE', 'notifyZF', 'forwarderSP', 'File Nm']
    dfcus1 = dfcus.copy()
    dfcus1.columns = ['brand', 'customercode', 'search term', 'Attribute2',
                      'Attribute3', 'Attribute4', 'Attribute5', 'customer code', 'File Nm']
    ship = [['air', '11'], ['ocean', '21'], ['truck', '31'], [
        'courier', '51'], ['tb-boat truck', '21'], ['fast vessel', '21']]
    dfship = pd.DataFrame(ship, columns=['shipment', 'shipmentcode'])
    dfgtn1['Shipment Method'] = dfgtn1['Shipment Method'].str.lower()
    df = pd.merge(dfgtn1, dfship, left_on=['Shipment Method'], right_on=[
                  'shipment'], how='left', sort=False)
    df1 = pd.merge(df, dfcus1, left_on=['Destination Code (Plant)'], right_on=[
                   'search term'], how='left', sort=False)
    df2 = pd.merge(df1, dfpar1, left_on=['customer code'], right_on=[
                   'ship-to'], how='left', sort=False)
    df3 = df2.copy()

    # A
    # 客戶採購單號碼
    df3['PURCH_NO_C'] = df3['PR Number']  # PO No.
    # 客戶採購單的號碼項目
    df3['PO_ITM_NO'] = ''
    # 採購單日期
    df3['PURCH_DATE'] = pd.to_datetime(df3['Issue Date']).dt.strftime('%Y%m%d')
    # 收貨人採購單號碼
    df3['PURCH_NO_S'] = df3['收貨人採購單號']
    # 取代的客戶採購單號碼
    df3['ZRPCPURNO'] = df3['PO #']
    # 取代的客戶採購單項次
    df3['ZRPCPRITM'] = ''
    # 銷售組織
    df3['SALES_ORG'] = 4020
    # 配銷通路
    df3['DISTR_CHAN'] = 38
    # 部門
    df3['DIVISION'] = 10
    # 品牌請求交貨日期
    df3['REQ_DATE_H'] = pd.to_datetime(
        df3['Revised Production End Date']).dt.strftime('%Y%m%d')
    # 文件日期
    df3['DOC_DATE'] = strftime("%Y%m%d", time.localtime())
    # 銷售文件類型
    df3['DOC_TYPE'] = 'ZO10'
    # 定價及匯率日期
    df3['PRICE_DATE'] = pd.to_datetime(df3['Issue Date']).dt.strftime('%Y%m%d')
    # 出貨類型
    df3['SHIP_TYPE'] = df3['shipmentcode']
    # 工廠
    df3['PLANT'] = '612A'
    # 需求類型
    df3['SGT_RCAT'] = 'FLT'
    # 客戶需求交貨日
    df3['ZZEND_CUST_REQ'] = pd.to_datetime(
        df3['Vendor Last Revised CRD']).dt.strftime('%Y%m%d')
    # 生產年月
    df3['ZZPRD_PERIOD'] = pd.to_datetime(
        df3['Revised Production End Date']).dt.strftime('%Y%m')
    # 包裝方式
    df3['ZZMIXPACK'] = '10'
    # Brand Factory Name
    df3['ZZART_PLANTNAME'] = ''
    # 客戶採購單HS Code
    df3['ZZCUSHSCDE'] = ''
    # Variant Value Conversion Type
    df3['FSH_SC_VCTYP'] = 'US'
    # 數量
    df3['ZDOUBLE'] = '6'

    # B
    df3['ZSO_GANO'] = df3['Buyer Item #'] + \
        df3['Size'].str[-1:] + df3['shipmentcode']
    POlist = list(set(df3['PURCH_NO_C'].tolist()))
    for i in range(len(POlist)):
        check1 = df3['PURCH_NO_C'] == POlist[i]
        GANOlist = list(set(df3[check1]['ZSO_GANO'].tolist()))
        for j in range(len(GANOlist)):
            check2 = df3['ZSO_GANO'] == GANOlist[j]
            df3.loc[check1 & check2, 'PO_ITM_NO'] = str(
                GANOlist.index(GANOlist[j]) + 1).zfill(6)
            df3.loc[check1 & check2, 'ZRPCPRITM'] = str(
                GANOlist.index(GANOlist[j]) + 1).zfill(6)
    # 客戶採購單日期
    df3['PURCH_DATE_ITM'] = df3['PURCH_DATE']
    # 尺碼
    df3['WRF_CHARSTC2'] = df3['Size'].str[:-
                                          1].astype(int).map(lambda x: "US_" + str(float(x/10)).rstrip("0").rstrip("."))
    # 物料號碼
    df3['MATNR'] = ''
    # 客戶料號
    df3['KDMAT'] = df3['Buyer Item #'] + "-" + df3['Size'].str[-1:]
    # 訂單數量
    # df3['PO_QUAN'] = df3['Quantity']
    df3['Quantity'] = df3['Quantity'].astype(int)
    df3['PO_QUAN'] = df3.groupby(['PURCH_NO_C', 'PO_ITM_NO', 'ZSO_GANO', 'WRF_CHARSTC2'])[
        'Quantity'].transform('sum')
    # 訂單數量單位
    df3['PO_UNIT'] = 'PAA'
    # 報價和銷售訂單的拒收原因
    df3['REASON_REJ'] = ''
    # Material Type
    df3['ZZMTART'] = 'FG'
    # 外貿的商品代碼/進口代號
    df3['ZZSTAWN'] = ''

    # C
    df4 = df3.copy()
    sheetCcolumns = ['PURCH_NO_C', 'PO_ITM_NO', 'PURCH_DATE', 'PARTN_ROLE', 'PARTN_NUM_CUST', 'PARTN_SRT_CUST',
                     'PARTN_NUM_VEND', 'PARTN_SRT_VEND', 'PARTN_SRT_NAME', 'PARTN_EXT_CUST', 'PARTN_EXT_VEND']
    sheetCcombine = pd.DataFrame(columns=sheetCcolumns)
    dfCWE = df4[['PURCH_NO_C', 'PO_ITM_NO', 'PURCH_DATE', 'ship-to']].copy()
    dfCWE['PARTN_NUM_CUST'] = dfCWE['ship-to']
    dfCWE['PARTN_ROLE'] = 'WE'
    dfmergeWE = pd.concat([sheetCcombine, dfCWE], ignore_index=True)
    dfCZB = df4[['PURCH_NO_C', 'PO_ITM_NO',
                 'PURCH_DATE', 'consigneeZB']].copy()
    dfCZB['PARTN_NUM_CUST'] = dfCZB['consigneeZB']
    dfCZB['PARTN_ROLE'] = 'ZB'
    dfmergeWEZB = pd.concat([dfmergeWE, dfCZB], ignore_index=True)
    dfCZD = df4[['PURCH_NO_C', 'PO_ITM_NO', 'PURCH_DATE', 'notifyZD']].copy()
    dfCZD['PARTN_NUM_CUST'] = dfCZD['notifyZD']
    dfCZD['PARTN_ROLE'] = 'ZD'
    dfmergeWEZBZD = pd.concat([dfmergeWEZB, dfCZD], ignore_index=True)
    dfCZE = df4[['PURCH_NO_C', 'PO_ITM_NO', 'PURCH_DATE', 'notifyZE']].copy()
    dfCZE['PARTN_NUM_CUST'] = dfCZE['notifyZE']
    dfCZE['PARTN_ROLE'] = 'ZE'
    dfmergeWEZBZDZE = pd.concat([dfmergeWEZBZD, dfCZE], ignore_index=True)
    dfCZF = df4[['PURCH_NO_C', 'PO_ITM_NO', 'PURCH_DATE', 'notifyZF']].copy()
    dfCZF['PARTN_NUM_CUST'] = dfCZF['notifyZF']
    dfCZF['PARTN_ROLE'] = 'ZF'
    dfmergeWEZBZDZEZF = pd.concat([dfmergeWEZBZDZE, dfCZF], ignore_index=True)
    dfCSP = df4[['PURCH_NO_C', 'PO_ITM_NO',
                 'PURCH_DATE', 'forwarderSP']].copy()
    dfCSP['PARTN_NUM_VEND'] = dfCSP['forwarderSP']
    dfCSP['PARTN_ROLE'] = 'SP'
    dfmergeWEZBZDZEZFSP = pd.concat(
        [dfmergeWEZBZDZEZF, dfCSP], ignore_index=True)
    cdt1 = dfmergeWEZBZDZEZFSP['PARTN_NUM_CUST'].astype(str) != 'nan'
    cdt2 = dfmergeWEZBZDZEZFSP['PARTN_NUM_VEND'].astype(str) != 'nan'
    dfpartner = dfmergeWEZBZDZEZFSP[cdt1 | cdt2].copy()
    dfpartnererror = dfmergeWEZBZDZEZFSP[(dfmergeWEZBZDZEZFSP['PARTN_ROLE'].astype(str) == 'WE') &
                                         (dfmergeWEZBZDZEZFSP['PARTN_NUM_CUST'].astype(str) == 'nan')].copy()

    # F
    df4['ZCUS_CONTRACT'] = ''

    # column check
    sheetAcol = ['PURCH_NO_C', 'PO_ITM_NO', 'PURCH_DATE', 'PURCH_NO_S', 'ZRPCPURNO', 'ZRPCPRITM', 'SALES_ORG', 'DISTR_CHAN', 'DIVISION', 'REQ_DATE_H', 'DOC_DATE',
                 'DOC_TYPE', 'PRICE_DATE', 'SHIP_TYPE', 'PLANT', 'SGT_RCAT', 'ZZEND_CUST_REQ', 'ZZPRD_PERIOD', 'ZZMIXPACK', 'ZZART_PLANTNAME', 'ZZCUSHSCDE', 'FSH_SC_VCTYP', 'ZDOUBLE']
    sheetBcol = ['PURCH_NO_C', 'PO_ITM_NO', 'ZSO_GANO', 'PURCH_DATE_ITM', 'WRF_CHARSTC2',
                 'MATNR', 'KDMAT', 'PO_QUAN', 'PO_UNIT', 'REASON_REJ', 'ZZMTART', 'ZZSTAWN']
    sheetCcol = ['PURCH_NO_C', 'PO_ITM_NO', 'PURCH_DATE', 'PARTN_ROLE', 'PARTN_NUM_CUST', 'PARTN_SRT_CUST',
                 'PARTN_NUM_VEND', 'PARTN_SRT_VEND', 'PARTN_SRT_NAME', 'PARTN_EXT_CUST', 'PARTN_EXT_VEND']
    sheetFcol = ['PURCH_NO_C', 'PO_ITM_NO', 'ZCUS_CONTRACT']
    sheetA = df4[sheetAcol].copy()
    sheetA.drop_duplicates(inplace=True, keep='first')
    sheetB = df4[sheetBcol].copy()
    sheetB.drop_duplicates(inplace=True, keep='first')
    sheetC = dfpartner[sheetCcol].copy()
    sheetC.drop_duplicates(inplace=True, keep='first')
    sheetF = df4[sheetFcol].copy()
    sheetF.drop_duplicates(inplace=True, keep='first')
    sheetAblank = pd.DataFrame(
        np.full((3, len(sheetAcol)), ''), columns=sheetAcol)
    sheetBblank = pd.DataFrame(
        np.full((3, len(sheetBcol)), ''), columns=sheetBcol)
    sheetCblank = pd.DataFrame(
        np.full((3, len(sheetCcol)), ''), columns=sheetCcol)
    sheetFblank = pd.DataFrame(
        np.full((3, len(sheetFcol)), ''), columns=sheetFcol)
    sheetAnew = pd.concat([sheetAblank, sheetA])
    sheetBnew = pd.concat([sheetBblank, sheetB])
    sheetCnew = pd.concat([sheetCblank, sheetC])
    sheetFnew = pd.concat([sheetFblank, sheetF])
    sheetblank = pd.DataFrame()

    # output
    tm = (datetime.now()+timedelta(hours=7)).strftime("%Y%m%d_%H%M%S")
    outputfilename = "GTNto6sheet_TBL_"+tm+".xlsx"
    write = pd.ExcelWriter(userfoldertimepath+"/"+outputfilename)

    sheetAnew.to_excel(write, sheet_name='ZSDTSD001A',
                       index=False, freeze_panes=(1, 0))
    sheetBnew.to_excel(write, sheet_name='ZSDTSD001B',
                       index=False, freeze_panes=(1, 0))
    sheetCnew.to_excel(write, sheet_name='ZSDTSD001C',
                       index=False, freeze_panes=(1, 0))
    sheetblank.to_excel(write, sheet_name='ZSDTSD001D',
                        index=False, freeze_panes=(1, 0))
    sheetblank.to_excel(write, sheet_name='ZSDTSD001E29',
                        index=False, freeze_panes=(1, 0))
    sheetFnew.to_excel(write, sheet_name='ZSDTSD001F',
                       index=False, freeze_panes=(1, 0))

    if len(dfpartnererror) != 0:
        dfpartnererror.to_excel(
            write, sheet_name='withoutshipto', index=False, freeze_panes=(1, 0))

    write.close()

    return outputfilename
