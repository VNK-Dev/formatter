import openpyxl
from openpyxl.styles import Font
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

def mapDuplicates(existingProduct, product):
        if existingProduct.name == product.name:
            existingProduct.quantity += product.quantity
        return existingProduct

class Table:

    def __init__(self, file):
        self.dataframe = pd.read_excel(file)

    def transform(self, mapping):
        products = self._parse_products(mapping)
        products.sort(key=lambda x: x.ordering) 
        nomenclatureColumn, quantityColumn = self._columns_from_products(products)
        self.dataframe = pd.DataFrame({'Номенклатура': nomenclatureColumn, 'Количество': quantityColumn})

    def _parse_products(self, mapping):
        products = []
        for _, row in self.dataframe.iterrows():
            product = mapping.create_product(row["Номенклатура"], row["Количество"])
            if product is None:
                continue
            duplicates = list(filter(lambda existingProduct: existingProduct.name == product.name, products))
            if len(duplicates) > 0:
                products = list(map(lambda existingProduct: mapDuplicates(existingProduct, product), products))
                continue
            products.append(product)
        return products

    def _columns_from_products(self, products):
        nomenclatureColumn = []
        quantityColumn = []
        for product in products:
            if product.category not in nomenclatureColumn:
                nomenclatureColumn.append(product.category)
                quantityColumn.append('')
            nomenclatureColumn.append(product.name)
            quantityColumn.append(product.quantity)
        return nomenclatureColumn, quantityColumn

    def to_excel(self, fileName):
        with pd.ExcelWriter(fileName, engine='openpyxl') as writer:
            # Convert the DataFrame to an XlsxWriter Excel object
            self.dataframe.to_excel(writer, index=False, sheet_name='Sheet1')
            
            # Get the workbook and the writer's worksheet
            worksheet = writer.sheets['Sheet1']

            # Set the font size for the header
            header_font = Font(size=16)  # Set the header font size
            for cell in worksheet[1]:  # The first row contains the headers
                cell.font = header_font
            
            # Set the font size for the rest of the cells
            cell_font = Font(size=14)  # Set the cell font size
            for row in worksheet.iter_rows(min_row=2):
                for cell in row:
                    cell.font = cell_font
            
            # Set the column widths
            for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter  # Get the column name
                    
                    # Iterate over the rows in the column to find the max length
                    for cell in col:
                        try:
                            # Consider a scaling factor for font size
                            scaling_factor = 1.1 if cell.font.size == 12 else 1.2
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value)) * scaling_factor
                        except:
                            pass
                    adjusted_width = (max_length + 2)  # Add some padding
                    worksheet.column_dimensions[column].width = adjusted_width


class Mapping():
   
    def __init__(self, filePath):
        with open(filePath, 'r') as file:
            self.data = json.load(file)

    def create_product(self, nomenclature, quantity):
        for v in self.data.values():
            for name in v["names"]:
                if name["old"] == nomenclature:
                    return Product(name["new"], v["category"], v["ordering"], quantity)
        return None

        
class Product():

    def __init__(self, name, category, ordering, quantity):
        self.name = name
        self.category = category
        self.ordering = ordering
        self.quantity = quantity


def process(content):
    table = Table(content)
    table.transform(Mapping('settings.json'))
    table.to_excel("result.xlsx")
    return read_excel_to_bytes("result.xlsx")
