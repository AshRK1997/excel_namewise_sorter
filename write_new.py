import pandas as pd
import xlsxwriter
import os.path

def col():
    global col_find
    col_find = input("Enter the coloumn name to be searched in : ")
    
    try:
       col_val = dfs.columns.get_loc(col_find)
       
    except:
        print("Enter the correct Coloumn headers")
        col()
def file():
    global file_name
    file_name = input("Enter the path and name of the file with extension : ")
    #print(file_name)
    if(os.path.exists(file_name)):
        q=0
    else:
        print("file does not exists")
        file()
    
        
def shet():
    global sheet
    sheet = input("enter the sheet number : ")
    try:
        global dfs
        dfs = pd.read_excel(file_name, sheet_name=sheet)
        dfs = dfs.fillna(" - ")
        
    except:
        print("The specified sheet or file does not exist")
        shet()
def print_search():
    global i,j
    global found
    i,j=0,0
    found=False
    for index, row in dfs.iterrows():
        
        if name_search==row[col_find]:
            found = True
            for cell in row:
                worksheet.write(i, j, cell)
                #print(cell,sep=" ",end=" ")
                j+=1
            i+=1
            j=0
            #print("\n")
    if found==False:
        print("the specified name's not found")
        
def mainer():
    global col_len,name_search,workbook,worksheet
    file()
    shet()
    col()
    while(1):
        name_search = input("Enter name to be searched or if you want to exit search enter y : ")
        if name_search=="y":
            break
        else:
            workbook = xlsxwriter.Workbook('demo.xlsx')
            worksheet = workbook.add_worksheet(sheet) 
            worksheet.set_column('A:Z', 200000)
            col_len = len(dfs.columns)
            print_search()
            try:
                workbook.close()
            except:
                print("please delete demo.xlsx and restart the process")
                mainer()

mainer()
        



        

    
        
        
    





'''import xlwt

DATA = (("The Essential Calvin and Hobbes", 1988,),
        ("The Authoritative Calvin and Hobbes", 1990,),
        ("The Indispensable Calvin and Hobbes", 1992,),
        ("Attack of the Deranged Mutant Killer Monster Snow Goons", 1992,),
        ("The Days Are Just Packed", 1993,),
        ("Homicidal Psycho Jungle Cat", 1994,),
        ("There's Treasure Everywhere", 1996,),
        ("It's a Magical World", 1996,),)

wb = xlwt.Workbook()
ws = wb.add_sheet("My Sheet")
for i, row in enumerate(DATA):
    for j, col in enumerate(row):
        ws.write(i, j, col)
ws.col(0).width = 256 * max([len(row[0]) for row in DATA])
wb.save("myworkbook.xls")
'''
