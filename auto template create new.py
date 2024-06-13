import os
import pandas as pd

# Define the headings for the ERP template
erp_headings = [
    "CustomerTypeDescription",
    "MemberTypeDescription",
    "AdmissionNo",
    "MemberSurName",
    "MemberName",
    "FatherName",
    "SpouseName",
    "MemberNameRegional",
    "FatherNameinRegional",
    "SpouseNameinRegional",
    "DOB",
    "Age",
    "AdmissionDate",
    "GenderDescription",
    "MaritalStatusDesc",
    "CommunityDescription",
    "CasteDescription",
    "FarmerTypeDescription",
    "Address1",
    "Address2",
    "VillageDescription",
    "LedgerFolioNo",
    "ContactNo",
    "ShareBalance",
    "ThriftBalance",
    "DividentBalance",
    "AdhaarCardNo",
    "DCCBSBACNO",
    "PacsIDPKey",
    "BranchId"
]

# Input and output paths
input_dir = r'C:\Users\adity\Documents\SocietyData'
output_dir = r'C:\Users\adity\Documents\Automation data conversion'

# Ensure the output directory exists
os.makedirs(output_dir, exist_ok=True)

# Process each Excel file in the input directory
for filename in os.listdir(input_dir):
    if filename.endswith('.xlsx'):
        input_path = os.path.join(input_dir, filename)
        society_name = os.path.splitext(filename)[0]  # Extract society name
        df = pd.read_excel(input_path)

        # Dictionary to store data for each PACS
        pacs_data = {}

        for _, row in df.iterrows():
            pacs_name = row['BpacsName']
            if pacs_name not in pacs_data:
                pacs_data[pacs_name] = []
            pacs_data[pacs_name].append(row)

        # Process data for each PACS
        for pacs_name, data_rows in pacs_data.items():
            # Replace invalid characters in PACS name with underscores
            invalid_chars = r'[\]:*?/\\'  # Use raw string to avoid SyntaxWarning
            pacs_name_cleaned = ''.join('_' if c in invalid_chars else c for c in pacs_name)
            output_path = os.path.join(output_dir, f'{pacs_name_cleaned}.xlsx')
            
            erp_data_list = []

            for row in data_rows:
                # Map society data to ERP template
                share_rupees = str(row["ShareRupees"])  # Convert to string to handle non-int values
                share_balance = str(int(share_rupees) - 21) if share_rupees.endswith("21") else share_rupees

                erp_data = {
                    "CustomerTypeDescription": "Member",
                    "MemberTypeDescription": "A Type",
                    "AdmissionNo": "",
                    "MemberSurName": "Mrs" if row["Gender"] == "F" else "Mr",
                    "MemberName": row["ApplicantName"],
                    "FatherName": row["FatherName"],
                    "SpouseName": "",
                    "MemberNameRegional": "",
                    "FatherNameinRegional": "",
                    "SpouseNameinRegional": "",
                    "DOB": row["DOB"],
                    "Age": "33",  # Assuming a default value, can be computed or derived
                    "AdmissionDate": row["RegistrationDate"],
                    "GenderDescription": "Male" if row["Gender"] == "M" else "Female",
                    "MaritalStatusDesc": "Not Available",
                    "CommunityDescription": "Not Available",
                    "CasteDescription": "Not Available",
                    "FarmerTypeDescription": "Not Available",
                    "Address1": "",
                    "Address2": "",
                    "VillageDescription": row["GramPanchyatName"],
                    "LedgerFolioNo": "",
                    "ContactNo": row["MobileNo"],
                    "ShareBalance": share_balance,
                    "ThriftBalance": "0",
                    "DividentBalance": "0",
                    "AdhaarCardNo": row["AadharNo"],
                    "DCCBSBACNO": "",
                    "PacsIDPKey": "",
                    "BranchId": ""
                }
                erp_data_list.append(erp_data)

            # Convert list of dictionaries to DataFrame and save to Excel file
            erp_df = pd.DataFrame(erp_data_list, columns=erp_headings)
            erp_df.to_excel(output_path, index=False)
            print(f"ERP data written to: {output_path}")
