import sys
import pandas as panda
from datetime import datetime
import re
import os

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

# Get path of sales data CSV file from the command line
def get_sales_csv():
    # Check whether command line parameter provided
    if(len(sys.argv) == 1):
        print('Please Provide Command Line argument to file')
    else:
        # Check whether provide parameter is valid path of file
        try:
            panda.read_csv(sys.argv[1])
            return os.path.abspath(sys.argv[1])
        except:
            print('Not a valid File')
            quit()
    return

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    sales_folder = os.path.dirname(sales_csv)
    # Determine the name and path of the directory to hold the order data files
    date = datetime.now().strftime('%Y-%m-%d')
    filename = f'Orders_{date}'
    # Create the order directory if it does not already exist
    orders_folder = os.path.join(sales_folder, filename)
    if not os.path.isdir(orders_folder):
        os.makedirs(orders_folder)
    return orders_folder

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    sales_dataframe = panda.read_csv(sales_csv)
    # Insert a new "TOTAL PRICE" column into the DataFrame
    total_price = sales_dataframe['ITEM QUANTITY'] * sales_dataframe['ITEM PRICE']
    sales_dataframe.insert(7,'TOTAL PRICE',total_price)
    # Remove columns from the DataFrame that are not needed
    sales_dataframe.drop(columns=['ADDRESS','CITY','STATE','POSTAL CODE','COUNTRY'],inplace=True)
    # Group the rows in the DataFrame by order ID
    for order_id,order_data in sales_dataframe.groupby('ORDER ID'):
        # For each order ID:
        # Remove the "ORDER ID" column
        order_data.drop(columns=['ORDER ID'],inplace=True)
        # Sort the items by item number
        order_data.sort_values(by='ITEM NUMBER')
        # Append a "GRAND TOTAL" row
        grand_total = order_data['TOTAL PRICE'].sum()
        grand_total_frame = panda.DataFrame({'ITEM PRICE':['GRAND TOTAL:'],'TOTAL PRICE':[grand_total]})
        order_data = panda.concat([order_data,grand_total_frame])
        # Determine the file name and full path of the Excel sheet
        customer_name=order_data['CUSTOMER NAME'].values[0]
        customer_name=re.sub('\W','',customer_name)
        order_file_name=f"Order{order_id}_{customer_name}.xlsx"
        order_file_path=os.path.join(orders_dir,order_file_name)
        # Export the data to an Excel sheet
        sheet_name=f'Order {order_id}'
        # Format the Excel sheet (
        # TODO: Format the Excel sheet
        writer = panda.ExcelWriter(order_file_path, engine='xlsxwriter')
        order_data.to_excel(writer,index=False,sheet_name=sheet_name)
        workbook= writer.book
        worksheet=writer.sheets[sheet_name]
        # Define format for the money columns
        format1=workbook.add_format({'num_format':'$#,##0.00'})
        worksheet.set_column('F:G', 13, format1)
        # Format each colunm
        worksheet.set_column('A:A',11)
        worksheet.set_column('B:B',13)
        worksheet.set_column('C:C',15)
        worksheet.set_column('D:D',15)
        worksheet.set_column('E:E',15)
        worksheet.set_column('H:H',10)
        worksheet.set_column('I:I',30)
        # close the sheet
        writer.close()        


if __name__ == '__main__':
    main()