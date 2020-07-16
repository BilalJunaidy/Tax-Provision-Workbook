# The following code is going to be embedded into a vba script which is going to call this function when the user
# wants to initiate the tax provision workbook.
# Before the process begins, the user has to upload all relavent documents within the following directory
# C:\Users\E067090\OneDrive - RSM\TTC\DB to excel and back feature\Solution integration\docs
# I am using this test directory for now, but when this is deployed it will be some directory within the M drive.
# At a minimum the user needs to provide the Trial Balance and the Completed Provision Mapping Tool.
# Care should be taken not to change the structure of the Provision Mapping Tool by the tax professional.

# This program will take in the user provided documents noted above, and execute the following:

# A. Create a sqlite Database within the directory noted above (When deployed the directory will change)
# B. Create two tables in the database.
#    One for the Accounting to Tax Net Income calculation
#    The other for the Accounting vs Tax temp differences and subsequent DTA/DTL calculation

# C. The program will then read into memory the Mapping provision tool

# D. From the Mapping tool, the program will create the calculation of Accounting vs Tax Net Income, and calculate
#    the net income for tax purposes. 

# E. The program will then write this calculation within the opened excel file where the VBA script was run from.

# F. The program will then post these items into the database table created for this purpose.

# G. The program will then create the calculation of Timing differences and eventual DTA/DTL.

# H. The program will then write this calculation within the opened excel file where the VBA script was run from.

# I. The program will then post these items into the database table created for this purpose.

#    The feature that will let the user send in manual adjustments is going to be run through a seperate python file

###BEGINNING OF PROGRAM###


#Import all dependencies
import openpyxl
import sqlite3
from sqlite3 import Error
import win32com.client as win32
from datetime import datetime
import os


#The following function will create a connection with a sqlite database file.
#When initially used, this will actually create a sqlite database file
def create_connection(path):
    connection = None

    try:
        connection = sqlite3.connect(path)
        print("Connection to SQlite DB succcesful")
    except Error as e:
        print(f"{e} error occurred while trying to connect to DB")

    return connection

#The following calls the create_connection function
connection = create_connection(r'C:\Users\E067090\OneDrive - RSM\TTC\DB to excel and back feature\Solution integration\Tax_Provision.db')

#The following will create a table in the sqlite database 
def execute_query(connection, query):
#The following create a cursor object upon which the query is going to be executed
    cursor = connection.cursor()

    try:
        cursor.execute(query)
#Making sure to commit the execution above to help ensure that the query has actually updated the DB
        connection.commit()
        print("Query executed successfully")
    except Error as e:
        print(f"{e} occurred when executing query")


create_accounting_to_tax_NetIncome_table = """
    CREATE TABLE IF NOT EXISTS Tax_NetIncome (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    FSA TEXT NOT NULL,
    AMOUNT INTEGER,
    DATE_LAST_UPDATED TEXT,
    LAST_UPDATED_BY TEXT
    );
    """

create_accounting_vs_tax_value_differences = """
    CREATE TABLE IF NOT EXISTS Accounting_vs_Tax_diff (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    FSA TEXT NOT NULL,
    ACC_VAL INTEGER,
    TAX VAL INTEGER,
    TAX_vs_ACC INTEGER,
    DATE_LAST_UPDATED TEXT,
    LAST_UPDATED_BY TEXT    
    );
    """

execute_query(connection, create_accounting_to_tax_NetIncome_table)
execute_query(connection, create_accounting_vs_tax_value_differences)

def read_Provision_Mapping_Tool(connection):
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    workbook = excel.Workbooks.Open(r'C:\Users\E067090\OneDrive - RSM\TTC\DB to excel and back feature\Solution integration\docs\Provision_Mapping_Tool.xlsx')

    # Accessing sheet 1 here to get information from workbook to be able to compute net income for tax purposes

    sheet1 = workbook.Sheets(1)

    Acc_NI = int(sheet1.Range("D3").Value)
    Tax_Net_Income = Acc_NI

    #Defining Table1 here. This is a list with the values that are going to be entered into the first Table defined above
    table_1 = []

    table_1.append(sheet1.Range("A3").Value)
    table_1.append(Acc_NI)

    #Defining range for inclusion, accessing value and appending to table_1

    #Defining xlconstants
    xlRight = win32.constants.xlToRight
    xlDown = win32.constants.xlDown

    #get last row and last columns in our range
    inclusion_LastCol = sheet1.Range("F3").End(xlRight).Column  
    inclusion_LastRow = sheet1.Range("F3").End(xlDown).Row

    #Define the first and last cell in our range
    inclusion_First_cell = sheet1.Cells(3, 6)
    inclusion_Last_cell = sheet1.Cells(inclusion_LastRow, inclusion_LastCol)

    raw_data = sheet1.Range(inclusion_First_cell, inclusion_Last_cell).Value

    for i in range(len(raw_data)):
        #Appending FSA
        table_1.append(raw_data[i][0])

        #Appending Value

        #THE FOLLOWING HAS TO BE UNCOMMENTED - IMPORTANT
        Tax_Net_Income += int(raw_data[i][3])
        table_1.append(raw_data[i][3]) 

    #Defining range for deductions, accessing value and appending to table_1
    deduction_LastCol = sheet1.Range("K3").End(xlRight).Column  
    deduction_LastRow = sheet1.Range("K3").End(xlDown).Row

    #Define the first and last cell in our range
    deduction_First_cell = sheet1.Cells(3, 11)
    deduction_Last_cell = sheet1.Cells(deduction_LastRow, deduction_LastCol)

    raw_data = sheet1.Range(deduction_First_cell, deduction_Last_cell).Value

    for i in range(len(raw_data)):
        #Appending FSA
        table_1.append(raw_data[i][0])

        #Appending Value
        #THE FOLLOWING HAS TO BE UNCOMMENTED - IMPORTANT
        Tax_Net_Income -= int(raw_data[i][3])
        table_1.append(raw_data[i][3])

    #Appending to table_1 the calculated tax net income
    table_1.append("Tax Net Income")
    table_1.append(f"{Tax_Net_Income}")


    #Inserting table_1 into the Tax_NetIncome database table
    len_table = int(len(table_1))
    curr_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    curr_user = homedir = os.environ['HOME'].split("\\")[-1]   
    cursor = connection.cursor()
    
    for i in range(0, len_table, 2):
        
        FSA = table_1[i]
        AMOUNT = table_1[i+1]
        #THE FOLLOWING HAS TO BE UNCOMMENTED - IMPORTANT
        #Tax_Net_Income += AMOUNT
        cursor.execute("INSERT INTO Tax_NetIncome (FSA, AMOUNT, DATE_LAST_UPDATED, LAST_UPDATED_BY) VALUES(?, ?, ?, ?)",
                   (FSA, int(AMOUNT), curr_time, curr_user))
        connection.commit()

    #The following is updating the Tax_Provision_Workbook for the table_1 data (i.e. reconciliation between accounting and tax net income)

    for wb in excel.Workbooks:
        if "Tax_Provision_Workbook" in wb.Name:
            
            Main_Sheet = wb.ActiveSheet        

            #Get # of rows and columns
            Row_Len = len(table_1)/2
            
            Col_Len = 2
            

            #Defining the first and last cells in our range
            First_Cell = Main_Sheet.Cells(2,2)
            Last_Cell = Main_Sheet.Cells(1 + Row_Len, 3)
                                          

            #Defining the range
            main_sheet_Range = Main_Sheet.Range(First_Cell, Last_Cell)

            #Populate the excel sheet starting with cells (2,2)

            table = []

            for i in range(0, len(table_1), 2):
                table_new = []
                table_new.append(table_1[i])
                table_new.append(table_1[i+1])
                table.append(table_new)
                
            main_sheet_Range.Value = table
        

    
    # Accessing sheet 2 here to get information from workbook to be able to compute DTA/DTL

    sheet2 = workbook.Sheets(2)

    #Defining Table1 here. This is a list with the values that are going to be entered into the first Table defined above
    table_2 = []

    table_2.append("Tax Net Income")
    table_2.append(f"{Tax_Net_Income}")

    #Defining range for , accessing value and appending to table_1

    #Defining xlconstants
    xlRight = win32.constants.xlToRight
    xlDown = win32.constants.xlDown

    #get last row and last columns in our range
    LastCol = sheet2.Range("A3").End(xlRight).Column  
    LastRow = sheet2.Range("A3").End(xlDown).Row

    #Define the first and last cell in our range
    First_cell = sheet2.Cells(3, 1)
    Last_cell = sheet2.Cells(LastRow, LastCol)

    raw_data = sheet2.Range(First_cell, Last_cell).Value

    for i in range(len(raw_data) - 1):
        #Appending FSA
        table_2.append(raw_data[i][0])

        #Appending Value
        table_2.append(raw_data[i][2]) 

    #Inserting table_2 into the Accounting_vs_Tax_diff database table
    len_table_2_half = int(len(table_2)/3) - 1   
    cursor = connection.cursor()
    
    for i in range(0, len_table_1_half, 3):
        
        FSA = table_2[i]
        ACC_VAL = table_2[i+1]
        TAX_VAL = table_2[i+2]
        TAX_vs_ACC = TAX_VAL - ACC_VAL
        Tax_Net_Income += TAX_vs_ACC        
        cursor.execute("INSERT INTO Accounting_vs_Tax_diff (FSA, ACC_VAL, TAX_VAL, TAX_vs_ACC, DATE_LAST_UPDATED, LAST_UPDATED_BY) VALUES(?, ?, ?, ?, ?, ?)",
                   (FSA, (ACC_VAL), (TAX_VAL), (TAX_vs_ACC), DATE_LAST_UPDATED, LAST_UPDATED_BY))
        connection.commit()

    #The following is updating the Tax_Provision_Workbook for the table_1 data (i.e. reconciliation between accounting and tax net income)

    for wb in excel.Workbooks:
        if "Tax_Provision_Workbook" in wb.Name:
        
            Main_Sheet = wb.ActiveSheet        

            #Get # of rows and columns
            Row_Len = len(table_1[0])
            
            Col_Len = len(table_1)
            

            #Defining the first and last cells in our range
            First_Cell = Main_Sheet.Cells(2,8)
            Last_Cell = Main_Sheet.Cells(2+Col_Len, 7+Row_Len) 

            #Defining the range
            main_sheet_Range = Main_Sheet.Range(First_Cell, Last_Cell)

            #Populate the excel sheet starting with cells (2,8)
            table = []

            for i in range(0, len(table_1), 2):
                table_new = []
                table_new.append(table_1[i])
                table_new.append(table_1[i+1])
                table.append(table_new)
                
            main_sheet_Range.Value = table

read_Provision_Mapping_Tool(connection)

###END OF PROGRAM###



    
    
    
    
    






