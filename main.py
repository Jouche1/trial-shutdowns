from read import read_files
from database import query_snowflake
from config import *
from datetime import datetime
import os
import pandas as pd

def main():
    # Read all data from Excel file(s) --  order_id_list is created from "return combined_data['ORDER_ID'].to_list()" in read.py
    order_id_list = read_files(SOURCE_DIRECTORY)
    if not order_id_list:
        print("No order IDs found, exiting.")
        return

    # Snowflake query
    # result_df = query_snowflake(order_id_list)
    # if result_df.empty:
    #     print("No data returned from Snowflake, exiting.")
    #     return
    

    # 7/12 at 9:10pm - trying to figure out how to remove extra serials/orders from output.. should be 20 but get extras (some 4S)
    # Remove any extra serials (extra order numbers) that come from the query to keep the list the same 
    # filtered_df = result_df[result_df['SERIAL'].isin(order_id_list)]
    # merged_df = pd.merge(result_df, all_data, on='ORDER_ID', how='inner')

    # Export results to Excel
    # output_path = os.path.join(TARGET_DIRECTORY, 'test-Sheet 13.xlsx' )
    # sorted_df = result_df.sort_values(by='DOCUMENT_NUMBER')
    # sorted_df.to_excel(output_path, index=False)
    # print(f"Export to Excel successful: {output_path}")

    # sorted_df = result_df.sort_values(by='DOCUMENT_NUMBER')
    # print(sorted_df)

if __name__ == "__main__":
    main()
