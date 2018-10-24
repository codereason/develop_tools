import xlrd
import pymysql
from config import *
# Open the workbook and define the worksheet
book = xlrd.open_workbook(TABLE_PATH)
sheet = book.sheet_by_index(0)

# Establish a MySQL connection
conn = pymysql.connect(host=HOST, user=USER, passwd=PASSWORD,
                       db=DATABASE)

# Get the cursor, which is used to traverse the database, line by line
cursor = conn.cursor()

# Create the INSERT INTO sql query
table_name = TABLE_NAME
create_table_query = " Create table " + table_name + " (A VARCHAR(255) ,B VARCHAR(255) ,C VARCHAR(255) ,D VARCHAR(255) ,E VARCHAR(255) ,F VARCHAR(255) ,G VARCHAR(255) ,H VARCHAR(255) ,I VARCHAR(255) ,J VARCHAR(255),K VARCHAR(255),L VARCHAR(255));"

import_data_query = "INSERT INTO " + table_name + " (A, B, C, D, E, F, G, H, I, J, K, L) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"

# Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
cursor.execute(create_table_query)
for r in range(1, sheet.nrows):
    A = sheet.cell(r, 0).value
    B = sheet.cell(r, 1).value
    C = sheet.cell(r, 2).value
    D = sheet.cell(r, 3).value
    E = sheet.cell(r, 4).value
    F = sheet.cell(r, 5).value
    G = sheet.cell(r, 6).value
    H = sheet.cell(r, 7).value
    I = sheet.cell(r, 8).value
    J = sheet.cell(r, 9).value
    K = sheet.cell(r, 10).value
    L = sheet.cell(r, 11).value

    # Assign values from each row
    values = (A, B, C, D, E, F, G, H, I, J, K, L)

    # Execute sql Query
    cursor.execute(import_data_query, values)

# Close the cursor
cursor.close()

# Commit the transaction
conn.commit()

# Close the database connection
conn.close()
print("exporting from " + TABLE_PATH + " into database is done... table : "+ TABLE_NAME)
