import os
import pandas as pd
from processor import ExcelRAGProcessor

# Initialize the processor
processor = ExcelRAGProcessor()

def preprocess_file(file_path):
    """Preprocess the file to remove rows before the first occurrence of 3 sequential rows with both null and non-null values."""
    if file_path.endswith(".csv"):
        df = pd.read_csv(file_path, header=None)
    else:
        df = pd.read_excel(file_path, header=None, engine="openpyxl")
    
    # Detect the first occurrence of 3 sequential rows with both null and non-null values
    for i in range(len(df) - 2):
        row1 = df.iloc[i]
        row2 = df.iloc[i + 1]
        row3 = df.iloc[i + 2]
        
        if (
            row1.isnull().any() and row1.notnull().any() and
            row2.isnull().any() and row2.notnull().any() and
            row3.isnull().any() and row3.notnull().any()
        ):
            # Remove all rows before these 3 sequential rows
            df = df.iloc[i:]
            break
    
    # Save the preprocessed file to a temporary file
    temp_file_path = f"temp_{os.path.basename(file_path)}"
    if file_path.endswith(".csv"):
        df.to_csv(temp_file_path, index=False, header=False)
    else:
        df.to_excel(temp_file_path, index=False, header=False, engine="openpyxl")
    
    return temp_file_path

# Load and preprocess all Excel files from the current directory
uploaded_files = []
for file_name in os.listdir("."):
    if file_name.endswith(".xlsx") or file_name.endswith(".csv"):
        preprocessed_file = preprocess_file(file_name)
        uploaded_files.append(open(preprocessed_file, "rb"))

# Process the files
results = processor.process_files(uploaded_files)
print("Processed Results:", results)