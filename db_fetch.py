import pymysql
import json
from openpyxl import load_workbook

def add_data_xl(filename="testfile.xlsx", cell_add="B5", data="test_data"):
    template = load_workbook(filename)
    main_sheet = template.active
    main_sheet[cell_add] = data
    template.save(filename)

def SelSqlFun():
    """" This function will fetch the information from table for certain critiria() """
    with open(r"mysql_querry", "r") as file:
        data = json.load(file)
    db = pymysql.connect(host=data["DB_HOST"], user=data["DB_USER"], password=data["DB_PASSWD"], database=data["DB_NAME"], port=data["DB_PORT"])
    cursor = db.cursor()
    sql = f'SELECT {data["DB_FIELD"]} FROM {data["TABLE_NAME"]} WHERE app_code IN ({data["APP_CODE"]});'
    print(sql)
    try:
       # Execute the SQL command
       cursor.execute(sql)
       # Fetch all the rows in a list of lists.
       results = cursor.fetchall()

       next_row_po = data["ROW_NO"]
       next_row_pd = data["ROW_NO"]
       next_row_no = data["ROW_NO"]
       next_row_nd = data["ROW_NO"]

       for row in results:
           if ( row[6] == 'PROD' or row[6] == 'DR'):
               if( row[3] == "Windows" or row[3] == "Linux"):
                   add_data_xl(filename=data["PROD_OS_XL"], cell_add=f"B{next_row_po}", data=row[5])
                   add_data_xl(filename=data["PROD_OS_XL"], cell_add=f"F{next_row_po}", data=row[4])
                   next_row_po += 1

               elif( row[3] == "Oracle" or row[3] == "SQLServer"):
                   add_data_xl(filename=data["PROD_DB_XL"], cell_add=f"B{next_row_pd}", data=row[5])
                   add_data_xl(filename=data["PROD_DB_XL"], cell_add=f"G{next_row_pd}", data=row[4])
                   next_row_pd += 1

           elif ( row[6] == 'DEV' or row[6] == 'QA'):
               if( row[3] == "Windows" or row[3] == "Linux"):
                   add_data_xl(filename=data["NONPROD_OS_XL"], cell_add=f"B{next_row_no}", data=row[5])
                   add_data_xl(filename=data["NONPROD_OS_XL"], cell_add=f"D{next_row_no}", data=row[4])
                   next_row_no += 1

               elif( row[3] == "Oracle" or row[3] == "SQLServer"):
                   add_data_xl(filename=data["NONPROD_DB_XL"], cell_add=f"B{next_row_nd}", data=row[5])
                   add_data_xl(filename=data["NONPROD_DB_XL"], cell_add=f"c{next_row_nd}", data=row[4])
                   next_row_nd += 1

    except Exception as e:
       print (f"Error: unable to fetch data with error {e}")

    # disconnect from server
    db.close()
