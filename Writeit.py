#Xcel write module
#for 0.0.37


def Export(name,result,cellno,sheetname,filenames,saveloc):
    import xlwt


#Different styles
    style0 = xlwt.easyxf('font: name Times New Roman, color-index Black, bold on',
	num_format_str='#,##0')#can also do #,##0.00 if we need decimal places
#Workbook object defined
    wbk = xlwt.Workbook()

#A sheet created
    sheet = wbk.add_sheet('mysheet')

#Create name of file with filename provided by user
    
    name_final=saveloc + "\\" + name + ".xls"

#This count helps us decide how many cycles to run while writing
    loop_count=len(result)

#Writing format in xlwt is row,col,value
    
    sheet.write(0,0,'S.No',style0)
    sheet.write(0,1,'Results',style0)
    sheet.write(0,2,'Cell No.',style0)
    sheet.write(0,3,'Sheet Name',style0)
    sheet.write(0,4,'File Name',style0)
    s_no = 1
    for i in range(loop_count):
        sheet.write(i+2,0,s_no )
        sheet.write(i+2,1,result[i])
        sheet.write(i+2,2,cellno[i])
        sheet.write(i+2,3,sheetname[i])
        sheet.write(i+2,4,filenames[i])
        s_no+=1


    sheet.col(1).width=10000 #3333=1"  #This sets the width of column 1
    sheet.col(3).width=5000 #3333=1"  #This sets the width of column 3
    sheet.col(4).width=30000 #3333=1"  #This sets the width of column 4


#^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


    sheet.write(loop_count+3,1,"Search result = ",style0)
    print loop_count
    sheet.write(loop_count+3,2,loop_count,style0) #loop_count+3 because we want a diff of 3 rows b/w result and last entry
    
    wbk.save(name_final)
    


