# The following code is going to both the tax provision workbook and the database depending on updated values
# provided by the user in the excel document under the manual adjustment columns.

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
        #print(f"{e} error occurred while trying to connect to DB")

    return connection

#The following calls the create_connection function
connection = create_connection(r'C:\Users\E067090\OneDrive - RSM\TTC\DB to excel and back feature\Solution integration\Tax_Provision.db')


###The following will create a table in the sqlite database 
##def execute_query(connection, query):
###The following create a cursor object upon which the query is going to be executed
##    cursor = connection.cursor()
##
##    try:
##        cursor.execute(query)
###Making sure to commit the execution above to help ensure that the query has actually updated the DB
##        connection.commit()
##        #print("Query executed successfully")
##    except Error as e:
##        #print(f"{e} occurred when executing query")
##
##
##update_tax_NetIncome_table = """
##    UPDATE TABLE Tax_NetIncome {
##    
##    };
##    """
def Update_Provision_Mapping_Tool(connection):
    excel = win32.gencache.EnsureDispatch('Excel.Application')

    #The following helps access the main sheet of the tax provision workbook, so that we can access updates sent by user
    for wb in excel.Workbooks:
        if "Tax_Provision_Workbook" in wb.Name:
            Main_Sheet = wb.ActiveSheet

            #Defining xlconstants
            xlRight = win32.constants.xlToRight
            xlDown = win32.constants.xlDown

            #get last row and last columns in our range
            LastCol = Main_Sheet.Range("E2").End(xlRight).Column  
            LastRow = Main_Sheet.Range("E2").End(xlDown).Row

            #Define the first and last cell in our range
            First_cell = Main_Sheet.Cells(2, 5)
            Last_cell = Main_Sheet.Cells(LastRow, LastCol)

            #The raw_data represents the updated numbers entered by the users
            raw_data = Main_Sheet.Range(First_cell, Last_cell).Value
            print(raw_data) 
            

            curr_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            curr_user = os.environ['USERNAME']
            cursor = connection.cursor()
            table = []

            for i in range(len(raw_data)):
                AMOUNT = int(raw_data[i][0])
                ID = i+1
                table_new = []
                table_new.append(AMOUNT)
                table.append(table_new)
                
                print(AMOUNT)

                query = """UPDATE Tax_NetIncome SET AMOUNT = ?, DATE_LAST_UPDATED = ?, LAST_UPDATED_BY = ? WHERE id = ?"""
                data = ((AMOUNT), curr_time, curr_user, ID)
                cursor.execute(query, data)
                print("cursor executed")
                connection.commit()
                print("Connection commited")

            print(table)

            #The following code is to update the Tax provision workbook in real time
            #Get # of rows and columns
            Row_Len = len(table)
            
            Col_Len = 1
            

            #Defining the first and last cells in our range
            First_Cell = Main_Sheet.Cells(2,3)
            Last_Cell = Main_Sheet.Cells(1 + Row_Len, 3)
                                          

            #Defining the range
            main_sheet_Range = Main_Sheet.Range(First_Cell, Last_Cell)

            #Populate the excel sheet starting with cells (2,3)
            main_sheet_Range.Value = table
                

            

            
            

            






Update_Provision_Mapping_Tool(connection)
