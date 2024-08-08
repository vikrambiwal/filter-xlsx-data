import pandas as pd
import numpy as np

file_path = './data.xlsm'

df = pd.read_excel(file_path, engine='openpyxl', sheet_name=1)

current_entry = {"group": "", "emails": [], "timezone": ""}
all_list = []

for index, row in df.iterrows():
    location_group = row.get("location_group")
    
    if pd.isna(location_group):
        current_entry["emails"] + [row.get("PAM"), row.get("SAM")]
    else:
        
        if current_entry["group"]: 
            all_list.append(current_entry)
        current_entry = {
            "group": row.get("location_group"),
            "emails": [],
            "timezone": ""
        }
        if pd.isna(row.get("PAM")) is False:
            current_entry["emails"] = current_entry["emails"] + [row.get("PAM")]
            
        if pd.isna(row.get("SAM")) is False:
            current_entry["emails"] = current_entry["emails"] + [row.get("SAM")]
            
        if pd.isna(row.get("Time Zone")) is False:
            current_entry["timezone"] = row.get("Time Zone")

if current_entry["group"]:
    all_list.append(current_entry)


print(all_list)

for entry in all_list:
    entry["emails"] = ", ".join(entry["emails"])
df = pd.DataFrame(all_list)
df.to_excel("output.xlsx")
