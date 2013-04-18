#FU3_5

def xl(adrs,patn,cb_case,reg_case,EX_case,exclude):
    import xlrd
    import re
    
#-^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^-
#:::::::::::::::::::::::::::These are our lists :::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::
    
    final_list=[] #This list carries the results to be returned
    cell_name =[] #This list carries cell names respectively to be returned
    sheet_name=[] #This list carries sheet names respectively to be returned
    file_names=[] #This list carries file names respectively to be returned
    compileobj = []

    
    for patn in patn:

        
        if cb_case==False and EX_case ==False:
            x=re.compile(patn) #Returns 'None' if result negetive
            compileobj.append(x)
        elif cb_case==True and EX_case == True:
            patn_new = "^"+ patn + "$"
            x=re.compile(patn_new,flags=re.IGNORECASE) #For case insensitivity
            compileobj.append(x)
        elif cb_case==False and EX_case==True:
            patn_new = "^" + patn + "$" 
            x=re.compile(patn_new)
            compileobj.append(x)
        elif cb_case==True and EX_case==False:
            x=re.compile(patn,flags=re.IGNORECASE)
            compileobj.append(x)
    
    
    Total_files=len(adrs) #Total no.of files to be processed count
    for m in range(Total_files):
        book = xlrd.open_workbook(adrs[m])
        no_sheets = book.nsheets
        sh_name=book.sheet_names()
        
        
        for cnt in range(no_sheets):
            sh = book.sheet_by_index(cnt)

            for i in range(sh.nrows):
                for j in range(sh.ncols):
                    cell_value=repr(sh.cell_value(rowx=i, colx=j)) #Return a string containing a printable representation of an object.
                    #Repr return a wierd result while converting a unicode object.It add u and ' '. So we need to remove them

#-^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^--^-
#:::::::::::::::::::::::::::This section is to test for conditions to search:::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

                    try:
                        cell_value_new=cell_value.lstrip("u").lstrip("'").rstrip("'")#Removing unicode 'u' from the left and right ends "'"
                    except:
                        pass

                    for obj in compileobj:
                        x1=obj.search(cell_value_new)
                    

##                    if exclude==True:
##                        if not x1:
##                            cell_name.append(xlrd.cellname(rowx=i,colx=j))
##                            final_list.append(sh.cell_value(rowx=i, colx=j))
##                            sheet_name.append(sh_name[cnt])
##                            file_names.append(adrs[m])
##                        else:pass
##                    else:
                        if x1:
                            cell_name.append(xlrd.cellname(rowx=i,colx=j))
                            final_list.append(sh.cell_value(rowx=i, colx=j))
                            sheet_name.append(sh_name[cnt])
                            file_names.append(adrs[m])
                            break
                        else:pass

    return(final_list,cell_name,sheet_name,file_names)
                        

 
     
