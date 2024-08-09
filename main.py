
import pandas as pd
import itertools
import json


file_path = 'CDL.xlsx' 
xls = pd.ExcelFile(file_path)

# Loading all sheets into DataFrames 
signature_df = pd.read_excel(xls, 'signature ', dtype=str)
general_df = pd.read_excel(xls, 'general', dtype=str)
Take_2 = pd.read_excel(xls, '2', dtype=str)
Take_3 = pd.read_excel(xls, '3', dtype=str)
Take_4 = pd.read_excel(xls, '4', dtype=str)
Take_5 = pd.read_excel(xls, '5', dtype=str)
Take_6 = pd.read_excel(xls, '6', dtype=str)
Take_7 = pd.read_excel(xls, '7', dtype=str)

# Generating all possible combinations from the signature sheet
signature_combinations = list(itertools.product(*[
    signature_df[col].dropna().tolist() for col in signature_df.columns
]))

# Retrieving the corresponding data from a specific DataFrame
def get_data(df, key_col, key_value):
    row = df[df[key_col] == key_value]
    if not row.empty:
        return {k: str(v) for k, v in row.iloc[0].to_dict().items()}
    return {}

# Generating all possible mutations
mutations = []
for combination in signature_combinations:
    mutation = {
        "signature": ''.join(combination),
        "Organization": general_df.at[0, 'Organization'],
        "Standard": general_df.at[0, 'Standard'],
        " Description Take_2": get_data(Take_2, 'Take_2', combination[1]).get("Description Take_2", ""),
        " Temperature Take_2": get_data(Take_2, 'Take_2', combination[1]).get("Temperature Take_2", ""),
        " Rating": get_data(Take_2, 'Take_2', combination[1]).get("Rating", ""),
        "Size": combination[2],
        "Identification": get_data(Take_4, 'Take_4', combination[3]).get("Identification", ""),
        "Coverage": str(float(get_data(Take_4, 'Take_4', combination[3]).get("Coverage", "0")) * 100) + '%',
        "Number": get_data(Take_5, 'Take_5', combination[4]).get("Number", ""),
        "Color": get_data(Take_5, 'Take_5', combination[4]).get("Color", ""),
        "Material": get_data(Take_6, 'Take_6', combination[5]).get("Material", ""),
        "Temperature Take_6": get_data(Take_6, 'Take_6', combination[5]).get("Temperature Take_6", ""),
        "Temperature Take_7": get_data(Take_7, 'Take_7', combination[6]).get("Temperature Take_7", ""),
    }

    mutations.append(mutation)

# Converting the mutations list to JSON without escaping special characters
mutations_json = json.dumps(mutations, indent=4, ensure_ascii=False)

# Saving the JSON to a file
with open('virus_mutations.json', 'w', encoding='utf-8') as json_file:
    json_file.write(mutations_json)

print("All possible mutations have been generated and saved to 'virus_mutations.json'.")
