import openpyxl
from io import BytesIO
import pandas as pd
import json

def read_excel_to_bytes(file_path):
    # Load the workbook from the file path
    workbook = openpyxl.load_workbook(file_path)
    
    # Create a BytesIO object to hold the byte content
    byte_stream = BytesIO()
    
    # Save the workbook content into the BytesIO object
    workbook.save(byte_stream)
    
    # Get the byte content
    byte_content = byte_stream.getvalue()
    
    return byte_content

def process(content):
    df = pd.read_excel(content)

    with open('settings.json', 'r') as file:
        data = json.load(file)

    def mapDuplicates(element, newElement):
        if element["name"] == newElement["name"]:
            return {"name": newElement["name"], "quantity": element["quantity"] + newElement["quantity"], "ordering": newElement["ordering"], "title": v["title"]}
        return element

    results = []
    i = 0
    for row in df["Номенклатура"]:
        for k, v in data.items():
            for name in v["names"]:
                if name["old"] == row:
                        
                    newValue = {"name": name["new"], "quantity": df["Количество"][i].item(), "ordering": v["ordering"], "title": v["title"]}
                    filtered = list(filter(lambda x: x["name"] == newValue["name"], results))
                    if len(filtered) > 0:
                        results= list(map(lambda x: mapDuplicates(x, newValue), results))
                    else:
                        results.append(newValue)
        i += 1

    results.sort(key=lambda x: x["ordering"]) 
    names = []
    for result in results:
        if result["title"] not in names:
            names.append(result["title"])
        names.append(result["name"])
    print(names)

    new_df = pd.DataFrame({'Номенклатура': names})
    new_df.to_excel("new.xlsx", sheet_name='Sheet1', index=False)
    return read_excel_to_bytes("new.xlsx")

