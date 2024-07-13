import os
import glob
import pandas as pd

def read_files(directory):
    try:
        order_data = {}
        for file in glob.glob(os.path.join(directory, '*.xlsx')):
            df = pd.read_excel(file)
            order_ids = df['ORDER_ID'].drop_duplicates().to_list()
            order_data[file] = order_ids
        
        return order_data
    except Exception as e:
        print(f"Error reading files: {e}")
        return {}
    
