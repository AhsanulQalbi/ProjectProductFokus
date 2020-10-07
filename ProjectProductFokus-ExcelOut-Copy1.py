#!/usr/bin/env python
# coding: utf-8

# In[35]:


import pandas as pd
import numpy as np
import pandasql as ps
import xlsxwriter
import os
import pathlib
import shutil

path = pathlib.Path().absolute()

dataSet1 = pd.read_excel('cleanData.xlsx')
dataSet2 = pd.read_excel('descData.xlsx')

dataFrame = pd.DataFrame(dataSet1)
dataFrameDesc = pd.DataFrame(dataSet2)

dataFrame = dataFrame.replace(np.nan, '', regex=True)
dataFrameDesc = dataFrameDesc.replace(np.nan, '', regex=True)

query = "select DISTINCT AM from  dataFrame;"
dataUserTemp = np.array(ps.sqldf(query)).flatten().tolist()

query = "select NAMA from  dataFrameDesc"
usernames = np.array(ps.sqldf(query)).flatten().tolist()


print("Berikut merupakan List Nama User : ")
print(usernames)

index = 0
while (index< len(usernames)) : 
    
    file_name = usernames[index] + '.xlsx'
    if (usernames[index] not in dataUserTemp) :
        print("data {} Tidak Ditemukan\n\n".format(usernames[index]))
        index = index + 1
        continue
    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet()
    
    print("data {}".format(usernames[index]))
    
    dataFrameAmFilter = dataFrame.copy()
    query = "select * from  dataFrameAmFilter where AM = '{}';".format(usernames[index])
    dataFrameAmFilter = ps.sqldf(query)
    
    dataProductFokus = dataFrameAmFilter.copy()
    query = "select * from  dataFrameAmFilter where NAMAPROD like 'D-VIT%' "
    query = query + "or NAMAPROD like 'CALNIC%'"
    query = query + "or NAMAPROD like 'VOXIN%'"
    query = query + "or NAMAPROD like 'CEFAROX%'"
    query = query + "or NAMAPROD like 'GASTROLAN%'"
    query = query + "or NAMAPROD like 'VOMISTOP%'"
    query = query + "or NAMAPROD like 'ESTIN%'"
    query = query + "or NAMAPROD like 'ZEMINDO%'"
    query = query + "or NAMAPROD like 'BUSMIN%'"
    query = query + "or NAMAPROD like 'HI-SQUA%'"
    dataProductFokus = ps.sqldf(query)

    query = "select APT_PANEL1 from  dataFrameDesc where NAMA = \"{}\"".format(usernames[index])
    dispenList = np.array(ps.sqldf(query)).flatten().tolist()
    query = "select APT_PANEL2 from  dataFrameDesc where NAMA = \"{}\"".format(usernames[index])
    dispenList.extend(np.array(ps.sqldf(query)).flatten().tolist())
    query = "select APT_PANEL3 from  dataFrameDesc where NAMA = \"{}\"".format(usernames[index])
    dispenList.extend(np.array(ps.sqldf(query)).flatten().tolist())
    query = "select APOTIK_TENDER from  dataFrameDesc where NAMA = \"{}\"".format(usernames[index])
    dispenList.extend(np.array(ps.sqldf(query)).flatten().tolist())

    
    listApotekDispen = pd.DataFrame(dispenList, columns = ['dispen'])
    query = "select * from  dataProductFokus where \"NAMALANG\" not in (select dispen from listApotekDispen)"
    filterDataProductFokus = ps.sqldf(query)
   
    query = "select AT from  dataFrameDesc where NAMA = \"{}\"".format(usernames[index])
    AT = (np.array(ps.sqldf(query)).flatten().tolist())
   
    query = "select DISKON from  dataFrameDesc where NAMA = \"{}\"".format(usernames[index])
    diskon = np.array(ps.sqldf(query)).flatten().tolist()
    if(len(diskon[0]) == 0) : 
        diskonTotal = 0
    else : 
        diskonTotal = int(float(diskon[0]))
    
    query = "select sum(THNA) from  filterDataProductFokus"
    sumSales = np.array(ps.sqldf(query)).flatten().tolist()
    
    totalPresentaseDiskon = diskonTotal/sumSales[0]
    
    if AT[0] >= 1 and totalPresentaseDiskon <= 0.35: 
        query = "select ({} - 5000000) * 7.5/100 as \"Incentive Product Fokus\" from filterDataProductFokus ".format(sumSales[0])
        IPF = np.array(ps.sqldf(query)).flatten().tolist()
        TotalIncentive = IPF[0]
    else :
        TotalIncentive = 0
    
    title_format = workbook.add_format({'bold': True, 'bg_color': '#FFCAB8', 'font_color' : 'black', 'underline' : True, 'font_size':18})
    header_1_format = workbook.add_format({'bold': True, 'bg_color': '#6F6E6D', 'font_color' : 'white','border': 1})
    row_1_format = workbook.add_format({'bold': True, 'bg_color': 'white', 'font_color' : 'black','border': 1})
    header_2_format = workbook.add_format({'bold': True, 'bg_color': '#AFCED6', 'font_color' : 'black','border': 1})
    header_3_format = workbook.add_format({'bold': True, 'bg_color': '#FFF700', 'font_color' : 'black','border': 1})
    row_2_format = workbook.add_format({'bold': False, 'bg_color': 'C7C7C7', 'font_color' : 'black'})
    row_2_format_bdr_top = workbook.add_format({'bold': False, 'bg_color': 'C7C7C7', 'font_color' : 'black', 'top' : 2})
    row_2_format_bdr_bottom = workbook.add_format({'bold': False, 'bg_color': 'C7C7C7', 'font_color' : 'black', 'bottom' : 2})
    row_2_format_bdr_right = workbook.add_format({'bold': False, 'bg_color': 'C7C7C7', 'font_color' : 'black', 'right' : 2})
    row_2_format_bdr_left = workbook.add_format({'bold': False, 'bg_color': 'C7C7C7', 'font_color' : 'black', 'left' : 2})
    
    row_2_format_bdr_top_right = workbook.add_format({'bold': False, 'bg_color': 'C7C7C7', 'font_color' : 'black', 'top' : 2, 'right' : 2})
    row_2_format_bdr_top_left = workbook.add_format({'bold': False, 'bg_color': 'C7C7C7', 'font_color' : 'black', 'top' : 2, 'left' : 2})
    
    row_2_format_bdr_bottom_right = workbook.add_format({'bold': False, 'bg_color': 'C7C7C7', 'font_color' : 'black', 'bottom' : 2, 'right' : 2})
    row_2_format_bdr_bottom_left = workbook.add_format({'bold': False, 'bg_color': 'C7C7C7', 'font_color' : 'black', 'bottom' : 2, 'left' : 2})
    
    
    row_3_format = workbook.add_format({'bold': False, 'bg_color': 'white', 'font_color' : 'black'})
    dots_cell_format = workbook.add_format({'bold': False, 'bg_color': 'white', 'font_color' : 'black','pattern' : 4})
    
    row_1_format_bdrless = workbook.add_format({'bold': True, 'bg_color': 'white', 'font_color' : 'black'})
    row_3_format_bdr = workbook.add_format({'bold': False, 'bg_color': 'white', 'font_color' : 'black', 'border' : 1})
    
    worksheet.merge_range('C1:I1','INCENTIVE PRODUCT FOKUS',title_format)
    
    worksheet.write(2, 3, "NAMA",header_1_format)
    worksheet.write(2, 4, "TEAM",header_1_format)
    worksheet.write(2, 5, "BULAN",header_1_format)
    
    worksheet.write(3, 3, usernames[index], row_1_format)
    worksheet.write(3, 4, dataFrameAmFilter.TEAM[0], row_1_format)
    bulan = dataFrameAmFilter.Bulan[0][5:7] +  '/' + dataFrameAmFilter.Bulan[0][:4]
    worksheet.write(3, 5, bulan, row_1_format)
    
    string = str(dataFrameAmFilter.AT[0]) + "%"
    worksheet.write(5, 3, 'A/T', header_2_format)
    worksheet.write(5,4 , string, row_1_format)
    
    totalPresentaseDiskon = float("{:.2f}".format(totalPresentaseDiskon*100))
    string = str(totalPresentaseDiskon) + "%"
    worksheet.write(6, 3, 'BIAYA', header_2_format)
    worksheet.write(6,4 , string, row_1_format)
    
    worksheet.write(8, 2, '', row_2_format_bdr_top_left)
    worksheet.write(8, 3, '', row_2_format_bdr_top)
    worksheet.write(8, 4, '', row_2_format_bdr_top)
    worksheet.write(8, 5, '', row_2_format_bdr_top)
    worksheet.write(8, 6, '', row_2_format_bdr_top)
    worksheet.write(8, 7, '' , row_2_format_bdr_top)
    worksheet.write(8, 8, '', row_2_format_bdr_top_right)
    worksheet.write(9, 2, '', row_2_format_bdr_left)
    worksheet.write(9, 8, '', row_2_format_bdr_right)
    worksheet.merge_range('D10:H10','SALES NON PANEL / TENDER',header_3_format)
    
    worksheet.write(10, 2, '', row_2_format_bdr_left)
    worksheet.write(10, 3, 'NAMA APOTIK', header_2_format)
    worksheet.write(10, 4, 'PRODUK FOKUS', header_2_format)
    worksheet.write(10, 5, 'HNA', header_2_format)
    worksheet.write(10, 6, 'UNIT', header_2_format)
    worksheet.write(10, 7, 'SALES', header_2_format)
    worksheet.write(10, 8, '', row_2_format_bdr_right)
    
    row = 11
    col = 3
    index2 = 0
    flag_color = 0
    listLangNonPanel = set(filterDataProductFokus.NAMALANG)
    listLangNonPanel = list(listLangNonPanel)
    while(index2<len((listLangNonPanel))):
        if(flag_color == 0) : 
            worksheet.write(row, col, listLangNonPanel[index2], row_2_format)
        else :    
            worksheet.write(row, col, listLangNonPanel[index2], row_3_format)
        index3 = 0
        query = "SELECT * FROM filterDataProductFokus WHERE NAMALANG = \"{}\"".format(listLangNonPanel[index2])
        listProduct = ps.sqldf(query)
        while(index3 < len(listProduct)) : 
            worksheet.write(row, 2, '', row_2_format_bdr_left)
            worksheet.write(row, 8, '', row_2_format_bdr_right)
            col = col + 1
            if(flag_color == 0) :
                worksheet.write(row, col, listProduct.NAMAPROD[index3], row_2_format)
            else :    
                worksheet.write(row, col, listProduct.NAMAPROD[index3], row_3_format)
            
            col = col + 1
            if(flag_color == 0) :
                worksheet.write(row, col, listProduct.HNA[index3], row_2_format)
            else : 
                worksheet.write(row, col, listProduct.HNA[index3], row_3_format)
                
            col = col + 1
            if(flag_color == 0) :
                worksheet.write(row, col, listProduct.BANYAK[index3], row_2_format)
            else :
                worksheet.write(row, col, listProduct.BANYAK[index3], row_3_format)
            
            col = col + 1
            if(flag_color == 0) :
                worksheet.write(row, col, listProduct.THNA[index3], row_2_format)
                flag_color = 1
            else :
                worksheet.write(row, col, listProduct.THNA[index3], row_3_format)
                flag_color = 0
                
            col = 3
            index3 = index3 + 1
            row = row + 1
            if(flag_color == 0) :
                worksheet.write(row, col, ' ', row_2_format)
            else :
                worksheet.write(row, col, ' ', row_3_format)
        index2 = index2 + 1
    
    worksheet.write(row, 2, '', row_2_format_bdr_left)
    worksheet.write(row, 3, '', dots_cell_format)
    worksheet.write(row, 4, '', dots_cell_format)
    worksheet.write(row, 5, '', dots_cell_format)
    worksheet.write(row, 6, '', dots_cell_format)
    worksheet.write(row, 7, '' , dots_cell_format)
    worksheet.write(row, 8, '', row_2_format_bdr_right)
    
    row = row + 1
    
    worksheet.write(row, 2, '', row_2_format_bdr_left)
    worksheet.write(row, 3, 'TOTAL', header_2_format)
    worksheet.write(row, 4, '', header_2_format)
    worksheet.write(row, 5, '', header_2_format)
    worksheet.write(row, 6, '', header_2_format)
    worksheet.write(row, 7, sumSales[0] , header_2_format)
    worksheet.write(row, 8, '', row_2_format_bdr_right)
    
    row = row + 1
    worksheet.write(row, 2, '', row_2_format_bdr_bottom_left)
    worksheet.write(row, 3, '', row_2_format_bdr_bottom)
    worksheet.write(row, 4, '', row_2_format_bdr_bottom)
    worksheet.write(row, 5, '', row_2_format_bdr_bottom)
    worksheet.write(row, 6, '', row_2_format_bdr_bottom)
    worksheet.write(row, 7, '' , row_2_format_bdr_bottom)
    worksheet.write(row, 8, '', row_2_format_bdr_bottom_right)
    
    row = row + 2
    worksheet.write(row, 2, '', row_2_format_bdr_top_left)
    worksheet.write(row, 3, '', row_2_format_bdr_top)
    worksheet.write(row, 4, '', row_2_format_bdr_top)
    worksheet.write(row, 5, '', row_2_format_bdr_top)
    worksheet.write(row, 6, '', row_2_format_bdr_top)
    worksheet.write(row, 7, '' , row_2_format_bdr_top)
    worksheet.write(row, 8, '', row_2_format_bdr_top_right)
    
    row = row + 1
    worksheet.write(row, 2, '', row_2_format_bdr_left)
    worksheet.write(row, 8, '', row_2_format_bdr_right)
    worksheet.merge_range(row,3,row,7,'SALES PANEL / TENDER',header_1_format)

    row = row + 1
    worksheet.write(row, 2, '', row_2_format_bdr_left)
    worksheet.write(row, 3, 'NAMA APOTIK', header_2_format)
    worksheet.write(row, 4, 'PRODUK FOKUS', header_2_format)
    worksheet.write(row, 5, 'HNA', header_2_format)
    worksheet.write(row, 6, 'UNIT', header_2_format)
    worksheet.write(row, 7, 'SALES', header_2_format)
    worksheet.write(row, 8, '', row_2_format_bdr_right)
    
    
    row = row + 1
    col = 3
    index2 = 0
    query = "select * from  dataProductFokus where \"NAMALANG\" in (select dispen from listApotekDispen) and  PRSND <= 35"
    filterDataTenderProductFokus = ps.sqldf(query)
    
    flag_color = 0
    while(index2<len((dispenList))):
        if(flag_color == 0) :
            worksheet.write(row, col, dispenList[index2], row_2_format)
        else :
            worksheet.write(row, col, dispenList[index2], row_3_format)
            
        index3 = 0
        query = "SELECT * FROM filterDataTenderProductFokus WHERE NAMALANG = \"{}\"".format(dispenList[index2])
        listProduct = ps.sqldf(query)
        while(index3 < len(listProduct)) :
            worksheet.write(row, 2, '', row_2_format_bdr_left)
            worksheet.write(row, 8, '', row_2_format_bdr_right)
            col = col + 1
            if(flag_color == 0) :
                worksheet.write(row, col, listProduct.NAMAPROD[index3], row_2_format)
            else:
                worksheet.write(row, col, listProduct.NAMAPROD[index3], row_3_format)
                
            col = col + 1
            if(flag_color == 0) :
                worksheet.write(row, col, listProduct.HNA[index3], row_2_format)
            else :
                worksheet.write(row, col, listProduct.HNA[index3], row_3_format)
                
            col = col + 1
            if(flag_color == 0) :
                worksheet.write(row, col, listProduct.BANYAK[index3], row_2_format)
            else :
                worksheet.write(row, col, listProduct.BANYAK[index3], row_3_format)
                
            
            col = col + 1
            if(flag_color == 0) :
                worksheet.write(row, col, listProduct.THNA[index3], row_2_format)
                flag_color = 1
            else :
                worksheet.write(row, col, listProduct.THNA[index3], row_3_format)
                flag_color = 0
            col = 3
            index3 = index3 + 1
            row = row + 1

        index2 = index2 + 1
    
    worksheet.write(row, 2, '', row_2_format_bdr_left)
    worksheet.write(row, 3, '', dots_cell_format)
    worksheet.write(row, 4, '', dots_cell_format)
    worksheet.write(row, 5, '', dots_cell_format)
    worksheet.write(row, 6, '', dots_cell_format)
    worksheet.write(row, 7, '' , dots_cell_format)
    worksheet.write(row, 8, '', row_2_format_bdr_right)
    
    row = row + 1
    
    query = "select sum(THNA) from  filterDataTenderProductFokus"
    sumSalesTender = np.array(ps.sqldf(query)).flatten().tolist()
        
    worksheet.write(row, 2, '', row_2_format_bdr_left)
    worksheet.write(row, 3, 'TOTAL', header_2_format)
    worksheet.write(row, 4, '', header_2_format)
    worksheet.write(row, 5, '', header_2_format)
    worksheet.write(row, 6, '', header_2_format)
    worksheet.write(row, 7, sumSalesTender[0] , header_2_format)
    worksheet.write(row, 8, '', row_2_format_bdr_right)
    
    row = row + 1
    
    worksheet.write(row, 2, '', row_2_format_bdr_bottom_left)
    worksheet.write(row, 3, '', row_2_format_bdr_bottom)
    worksheet.write(row, 4, '', row_2_format_bdr_bottom)
    worksheet.write(row, 5, '', row_2_format_bdr_bottom)
    worksheet.write(row, 6, '', row_2_format_bdr_bottom)
    worksheet.write(row, 7, '' , row_2_format_bdr_bottom)
    worksheet.write(row, 8, '', row_2_format_bdr_bottom_right)
    
    row = row + 2
    worksheet.merge_range(row,5,row,6,'NETT SALES',header_2_format)
    worksheet.write(row, 7, sumSales[0]-5000000, row_1_format)
    
    row = row + 1
    worksheet.merge_range(row,5,row,6,'INCENTIVE PRODUK FOKUS',header_3_format)
    worksheet.write(row, 7, TotalIncentive, row_1_format)
    
    row = row + 3
    worksheet.write(row, 2, 'NOTE:', row_1_format_bdrless)
    
    last_header_format = workbook.add_format({'bold': True, 'bg_color': '#FF9069', 'font_color' : 'black', 'border' : 1})
    
    row = row + 1
    worksheet.write(row, 2, 'NO', last_header_format)
    worksheet.write(row, 3, 'KOLOM', last_header_format)
    worksheet.merge_range(row,4,row,8,'KETERANGAN',last_header_format)
    
    row = row + 1
    worksheet.write(row, 2, '1', row_3_format_bdr)
    worksheet.write(row, 3, 'A/T', header_2_format)
    worksheet.merge_range(row,4,row,8,'PERSENTASE SALES DIBAGI TARGET',row_3_format_bdr)

    row = row + 1
    worksheet.write(row, 2, '2', row_3_format_bdr)
    worksheet.write(row, 3, 'NON PANEL/TENDER', header_2_format)
    worksheet.merge_range(row,4,row,8,'PERSENTASE DISKON PRODUK FOKUS NON PANEL/TENDER',row_3_format_bdr)

    row = row + 1
    worksheet.write(row, 2, '3', row_3_format_bdr)
    worksheet.write(row, 3, 'NAMA APOTIK', header_2_format) 
    worksheet.merge_range(row,4,row,8,'NAMA APOTIK YANG DICOVER TP',row_3_format_bdr)
    
    row = row + 1
    worksheet.write(row, 2, '4', row_3_format_bdr)
    worksheet.write(row, 3, 'PRODUK FOKUS', header_2_format) 
    worksheet.merge_range(row,4,row,8,'NAMA PRODUK FOKUS',row_3_format_bdr)
    
    row = row + 1
    worksheet.write(row, 2, '5', row_3_format_bdr)
    worksheet.write(row, 3, 'UNIT', header_2_format) 
    worksheet.merge_range(row,4,row,8,'TOTAL UNIT PRODUK DALAM 1 BULAN',row_3_format_bdr)
    
    row = row + 1
    worksheet.write(row, 2, '6', row_3_format_bdr)
    worksheet.write(row, 3, 'SALES', header_2_format) 
    worksheet.merge_range(row,4,row,8,'UNIT x HNA',row_3_format_bdr)
    
    row = row + 1
    worksheet.write(row, 2, '7', row_3_format_bdr)
    worksheet.write(row, 3, 'VALUE PRODUK FOKUS', header_2_format) 
    worksheet.merge_range(row,4,row,8,'SALES NON PANEL/TENDER - 5 JUTA',row_3_format_bdr)

    row = row + 1
    worksheet.write(row, 2, '8', row_3_format_bdr)
    worksheet.write(row, 3, 'INCENTIVE PRODUK FOKUS', header_3_format) 
    worksheet.merge_range(row,4,row,8,'VALUE PRODUK FOKUS x 7,5 %',row_3_format_bdr)
    
    workbook.close()
    source = str(path) + "\\" + file_name
    destination = str(path) + "\\" + "results" + "\\" + file_name
    shutil.move(source, destination)
    print("File data {} telah terbuat\n\n".format(file_name))
  
    index = index + 1

print("Program has successfully executed")


# In[22]:


print(30/100)


# In[21]:




