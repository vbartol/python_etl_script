import psycopg2
import xlrd

#connect to the db
conn = psycopg2.connect(
        host="localhost",
        database="postgres",
        user="postgres",
        password="sammir",
        port="5432"
    )

cursor = conn.cursor()

#create table
try:
    cursor.execute("CREATE TABLE last_Week_Report (Store_NO varchar(20) PRIMARY KEY, Store Varchar(45), TY_Units Varchar(30), LY_Units Varchar(30), TW_Sales Varchar(30), LW_Sales Varchar(30), LW_Var Varchar(30), LY_Sales Varchar(30), LY_Var  Varchar(30), YTD_Sales Varchar(30), LYTD_Sales Varchar(30), LYTD_Var  Varchar(30));")
except:
    print("Table already exist")


query= """INSERT INTO last_Week_Report(Store_NO, Store, TY_Units, LY_Units, TW_Sales, LW_Sales, LW_Var, LY_Sales, LY_Var, YTD_Sales, LYTD_Sales, LYTD_Var) VALUES(%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""

#extract
def extract():
    workbook = xlrd.open_workbook("C:/Users/38595/Downloads/SpaceNK_20181222.xlsx")
    sheet = workbook.sheet_by_index(0)
    for r in range(6, sheet.nrows-1):
        Store_NO = sheet.cell(r, 2).value
        Store = sheet.cell(r,3).value
        if(sheet.cell(r,5).value<0):
            TY_Units = transform(int(round(sheet.cell(r,5).value)))
        else:
            TY_Units = int(round(sheet.cell(r,5).value))
        if(sheet.cell(r,6).value<0):
            LY_Units=transform(int(round(sheet.cell(r,6).value)))
        else:
            LY_Units = int(round(sheet.cell(r,6).value))
        if(sheet.cell(r,7).value<0):
            TW_Sales=transform(int(round(sheet.cell(r,7).value)))
        else:
            TW_Sales = int(round(sheet.cell(r,7).value))
        if(sheet.cell(r,8).value<0):
            LW_Sales=transform(int(round(sheet.cell(r,8).value)))
        else:
            LW_Sales = int(round(sheet.cell(r,8).value))
        LW_Var = transform(sheet.cell(r,9).value)
        if(sheet.cell(r,10).value<0):
            LY_Sales=transform(round(int(sheet.cell(r,10).value)))
        else:
            LY_Sales = int(round(sheet.cell(r,10).value))
        LY_Var = transform(sheet.cell(r,11).value)
        if(sheet.cell(r,12).value<0):
            YTD_Sales=transform(int(round(sheet.cell(r,12).value)))
        else:
            YTD_Sales = int(round(sheet.cell(r,12).value))
        if(sheet.cell(r,13).value<=0):
            LYTD_Sales=transform(round(int(sheet.cell(r,13).value)))
        else:
            LYTD_Sales = int(round(sheet.cell(r,13).value))
        LYTD_Var = transform(sheet.cell(r,14).value)

        values = (Store_NO, Store, TY_Units, LY_Units, TW_Sales, LW_Sales, LW_Var , LY_Sales, LY_Var, YTD_Sales, LYTD_Sales, LYTD_Var)
        load(values)

    # Close the cursor
    cursor.close()

    # Commit the transaction
    conn.commit()

    #close the connection
    conn.close()

def transform(value):
    if(type(value) == float):
        return percentage(value)
    if(type(value) == int):
        return number(value)
    else:
        return value


def number(value):
    if(value<0):
        value *= -1
        value = str(value)
        value = '(' + value + ')'
        return value


def percentage(value):
        if(value<0):
            value *= -1 * 100;
            value = int(value)
            value = str(value)
            value = ('('+ value + ')%')
            return value
        else:
            value*=100
            value = int(value)
            value= str(value)
            value = (value + '%')
            return value


#load
def load(values):
    cursor.execute(query,values)


def main():
    extract()

if __name__ == "__main__":
        main()
