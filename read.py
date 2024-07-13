import os
import glob
import pandas as pd

def read_files(directory):
    try:
        all_data = []
        for file in glob.glob(os.path.join(directory, '*.xlsx')):
            df = pd.read_excel(file)
            all_data.append(df)
        combined_data = pd.concat(all_data)
        
        combined_data = combined_data.drop_duplicates(subset=['ORDER_ID'])
        
        return combined_data['ORDER_ID'].to_list()
    except Exception as e:
        print(f"Error reading files: {e}")
        return []
    
