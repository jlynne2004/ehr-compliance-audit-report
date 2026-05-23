# Messy_EHR_Sample_Export.py
# This script is designed to export a sample of patient records to test the EHR_Compliance_Audit_Report script.

# MESSY DATA PARAMETERS #
# Invalid/Missing NPI Numbers
# Discharge Dates Before Admission Dates
# Missing ICD-10 Codes
# Duplicate Provider Names with Typos with Same NPI
# Future Encounter Date Errors

import pandas as pd
import random
from datetime import datetime, timedelta

# Set a random seed for reproducible "messy" data
random.seed(42)

# Base template pools for generating records
providers = [
    {"id": "P001", "name": "Dr. Sarah Jenkins", "npi": "1234567890"},
    {"id": "P002", "name": "Dr Robert Chen", "npi": "987654321"},  # Malformed: 9 digits
    {"id": "P003", "name": "Dr. Maria Rodriguez", "npi": None},     # Missing NPI
    {"id": "P004", "name": "Dr. Sarah Jenkins ", "npi": "1234567890"} # Duplicate name typo
]

icd10_codes = ["U07.1", "J11.1", "B34.2", "J06.9", None] # Includes missing code

# Initialize our data list
mock_data = []
base_date = datetime(2026, 5, 1)

# Generate 30 mock patient tracking records
for i in range(1, 31):
    patient_id = f"PT{1000 + i}"
    
    # Select a provider and clinical code at random
    provider = random.choice(providers)
    icd_code = random.choice(icd10_codes)
    
    # Generate logical baseline dates
    days_to_add = random.randint(1, 15)
    admission = base_date + timedelta(days=days_to_add)
    length_of_stay = random.randint(2, 7)
    discharge = admission + timedelta(days=length_of_stay)
    
    # Inject an intentional Date Logic Error into Record #12
    if i == 12:
        discharge = admission - timedelta(days=3) # Discharge before admission
        
    mock_data.append({
        "patient_id": patient_id,
        "admission_date": admission.strftime("%Y-%m-%d"),
        "discharge_date": discharge.strftime("%Y-%m-%d"),
        "provider_id": provider["id"],
        "provider_name": provider["name"],
        "provider_npi": provider["npi"],
        "icd_10_code": icd_code
    })

# Convert to DataFrame and save to the matching filename
df_mock = pd.DataFrame(mock_data)
df_mock.to_csv("messy_ehr_export.csv", index=False)

print("Successfully generated 'messy_ehr_export.csv' with intentional compliance defects.")
