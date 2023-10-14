import pandas as pd
import numpy as np

def excel_to_numpy(excel_file):
    # Read the Excel file into a Pandas DataFrame
    df = pd.read_excel(excel_file)

    # Convert the DataFrame to a NumPy array
    np_array = df.to_numpy()

    # Return the NumPy array
    return np_array

 

