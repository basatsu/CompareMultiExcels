#Compare 2 old and new version excels in two folders and check diff

#in case same file name and same sheets
#_diff.txt shows where diff is. result outputs to  _diff.xlsx

#in case same file name and not same sheets
#output to _SheetDiff.txt

import xlrd 
from pathlib import Path
import glob,os,re,openpyxl
from itertools import zip_longest

oldbook = input('Old folder pass:')
newbook = input('New folder pass:')
output = input('Output diff.xlsx to which folder pass:')
outtext = input('Output diff.txt to which folder pass:')
outsheet = input('(not same sheets)diff.xlsx to which folder pass:')

for old in glob.glob(os.path.join(oldbook, '*.xlsx')): 
    for new in glob.glob(os.path.join(newbook, '*.xlsx')):
        rb1 = xlrd.open_workbook(old) 
        rb2 = xlrd.open_workbook(new) 
        oldfile=os.path.basename(old)
        newfile=os.path.basename(new)
        
        if oldfile==newfile and rb1.nsheets == rb2.nsheets:
            wb = openpyxl.Workbook()
            outfile=os.path.splitext(os.path.basename(old))[0]
        
            txt=outtext+"\\"+outfile+"_diff.txt"
            with open(txt, mode='a',encoding='utf-8', newline="\n") as f:
                f.write(oldfile+"と"+newfile+"\n")
                f.close() 
            for shnum in range(max(rb1.nsheets, rb2.nsheets)):  
                sheet1 = rb1.sheet_by_index(shnum) 
                sheet2 = rb2.sheet_by_index(shnum)          
            
                with open(txt, mode='a',encoding='utf-8', newline="\n") as f:
                    f.write(rb1.sheet_by_index(shnum).name+"\n")
                    f.close()   

                newsheet=wb.create_sheet(title=rb1.sheet_by_index(shnum).name,index=shnum)
                #new file has less rows than old file
                if sheet1.nrows > sheet2.nrows:
                    print("new file "+newfile+","+rb2.sheet_by_index(shnum).name+": not match rows")

                else:
                    for rownum in range(max(sheet1.nrows, sheet2.nrows)): 
                        if rownum < sheet1.nrows: 

                            row_rb1 = sheet1.row_values(rownum) 
                            row_rb2 = sheet2.row_values(rownum)
 
                            for colnum, (c1, c2) in enumerate(zip_longest(row_rb1, row_rb2)): 
                                if c1 != c2: 
                                    newsheet.cell(row=rownum+1, column=colnum+1).value=("{}→{}").format(c1,c2)
                                    with open(txt, mode='a',encoding='utf-8', newline="\n") as f:
                                        f.write("Row{},Col{}:{}and{}".format(rownum+1, colnum+1, c1, c2)+"\n")
                                        f.close() 
                        else: 
                            print("oldfile "+oldfile+",Sheetname:"+rb2.sheet_by_index(shnum).name+",less rows than new file "+newfile)
                            print("Row {} missing".format(rownum+1)) 
            wb.remove(wb["Sheet"])
            wb.save(output+"\\"+outfile+"_diff.xlsx")

        elif oldfile==newfile and rb1.nsheets != rb2.nsheets:
            outfile2=os.path.splitext(os.path.basename(old))[0]
            txtsheet=outsheet+"\\"+outfile2+"_SheetDiff.txt"
            with open(txtsheet, mode='a',encoding='utf-8', newline="\n") as t:
                t.write(oldfile+"and"+newfile+":Not match number of sheets\n")
                t.close()
        else:
            pass
print("Successful compare diff")


