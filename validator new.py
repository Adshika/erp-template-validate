import pandas as pd
import os

# Ensure xlsxwriter is installed
try:
    import xlsxwriter
except ImportError:
    print("xlsxwriter module not found. Installing...")
    os.system("pip install xlsxwriter")
    import xlsxwriter

# Define the provided column names
column_names = [
    'CustomerTypeDescription', 'MemberTypeDescription', 'AdmissionNo', 'MemberSurName', 'MemberName',
    'FatherName', 'SpouseName', 'MemberNameRegional', 'FatherNameinRegional', 'SpouseNameinRegional', 'DOB', 
    'Age', 'AdmissionDate', 'GenderDescription', 'MaritalStatusDesc', 'CommunityDescription', 'CasteDescription', 
    'FarmerTypeDescription', 'Address1', 'Address2', 'VillageDescription', 'LedgerFolioNo', 'ContactNo', 
    'ShareBalance', 'ThriftBalance', 'DividentBalance', 'AdhaarCardNo', 'DCCBSBACNO', 'PacsIDPKey', 'BranchId'
]

# Input directory path
input_dir = "C:\\Users\\adity\\Desktop\\input"

# Output directory path
output_dir = "C:\\Users\\adity\\Desktop\\output"

# Create the output directory if it doesn't exist
os.makedirs(output_dir, exist_ok=True)

# Function to clean up names using proper and trim
def clean_name(name):
    if pd.isna(name):
        return name
    return ' '.join([word.capitalize() for word in name.split()])

# Iterate over each file in the input directory
for filename in os.listdir(input_dir):
    if filename.endswith(".xlsx"):  # Process only Excel files
        input_excel_path = os.path.join(input_dir, filename)
        
        try:
            print(f"Processing file: {input_excel_path}")
            
            # Read the existing Excel file
            try:
                df = pd.read_excel(input_excel_path)
            except Exception as e:
                print(f"Failed to read {input_excel_path}: {e}")
                continue
            
            # Debugging: Print the shape of the dataframe
            print(f"Original DataFrame shape: {df.shape}")
            
            # Initialize a DataFrame with the required columns
            df_filtered = pd.DataFrame(columns=column_names)
            
            # Populate the DataFrame with the columns from the input file
            for col in column_names:
                if col in df.columns:
                    df_filtered[col] = df[col]

            # Apply proper and trim functions to MemberName and FatherName columns
            if 'MemberName' in df_filtered.columns:
                df_filtered['MemberName'] = df_filtered['MemberName'].apply(clean_name)
            if 'FatherName' in df_filtered.columns:
                df_filtered['FatherName'] = df_filtered['FatherName'].apply(clean_name)
            
            # Fill missing 'Age' column with 33
            df_filtered['Age'] = 33
            
            # Ensure the specified columns have 'Not Available' if missing
            mandatory_columns = ['MaritalStatusDesc', 'CommunityDescription', 'CasteDescription', 'FarmerTypeDescription']
            for col in mandatory_columns:
                if col not in df_filtered.columns or df_filtered[col].isnull().all():
                    df_filtered[col] = 'Not Available'
            
            # Debugging: Print the first few rows to check the changes
            print(df_filtered.head())

            # Construct the output file path
            output_filename = os.path.splitext(filename)[0] + "_Template.xlsx"
            output_excel_path = os.path.join(output_dir, output_filename)
            
            # Debugging: Print the output file path
            print(f"Output file path: {output_excel_path}")
            
            # Create a new Excel writer object
            writer = pd.ExcelWriter(output_excel_path, engine='xlsxwriter')
            
            # Write the filtered dataframe to the new Excel file
            df_filtered.to_excel(writer, index=False)
            
            # Save and close the Excel file
            writer.close()
            
            print(f"New Excel file created successfully: {output_excel_path}")
        except Exception as e:
            print(f"An error occurred while processing {filename}: {e}")

print("Processing completed.")
