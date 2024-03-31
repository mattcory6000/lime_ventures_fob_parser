import pandas as pd

def preprocess_and_parse_excel_sheet(file_path):
    xl = pd.ExcelFile(file_path)
    df = xl.parse(xl.sheet_names[0])  # Assuming data is in the first sheet

    processed_records = []
    category = ""
    producer_name = ""

    for index, row in df.iterrows():
        # Check if the row is a category header in ALLCAPS
        if row['Supplier / Product Name'].isupper():
            category = row['Supplier / Product Name']
            continue

        # Check if the row is a new producer name (not in ALLCAPS and no other data in the row)
        if not row['Supplier / Product Name'].isupper() and pd.isna(row['Product ID']):
            # Update producer name, strip any notes starting with two hyphens
            producer_name = row['Supplier / Product Name'].split(' --')[0]
            continue

        # Process specific product details
        if producer_name:
            # Extract and clean the product name
            product_name_components = row['Supplier / Product Name'].split()
            container = None
            if "CANS" in product_name_components:
                container = "CANS"
                product_name_components.remove("CANS")  # Remove "CANS"
            elif "PET" in product_name_components:
                container = "Keg"  # Indicating this is a keg product
                product_name_components.remove("PET")  # Remove "PET"
            product_name = ' '.join(product_name_components).replace(producer_name, "").strip()

            formatted_price = f"${row['Price']:.2f}" if pd.notnull(row['Price']) else None

            # Parse volume amount and unit
            if pd.notnull(row['Package']):
                package = row['Package'].lower()  # Convert to lowercase to standardize the checks
                # Initial default values
                container, volume_amount, volume_unit, pack_size = None, None, None, '1'

                if 'x' in package:
                    pack_size, volume_info = package.split('x', 1)
                    volume_amount = ''.join(filter(str.isdigit, volume_info))
                    volume_unit = ''.join(filter(str.isalpha, volume_info)).replace("cs", "").replace("cans",
                                                                                                      "").replace(
                        "cider", "").replace("wine", "")
                    pack_size = pack_size.strip() if pack_size else '1'  # Default pack_size to '1' if not defined
                elif 'pet' in package:
                    volume_amount = ''.join(filter(str.isdigit, package))
                    volume_unit = 'L'
                    container = 'PET'
                elif 'bbl' in package:
                    container = 'Keg'
                    if '1/2 bbl' in package:
                        volume_amount = '15.5'
                        volume_unit = 'gallons'
                    elif '1/6 bbl' in package:
                        volume_amount = '5.2'
                        volume_unit = 'gallons'
                    else:
                        # Handle generic bbl without specific size
                        volume_amount = ''.join(filter(str.isdigit, package))
                        volume_unit = 'gallons'

                # Normalize volume unit and pack size to proper capitalization
                volume_unit = volume_unit.capitalize() if volume_unit else volume_unit
                pack_size = pack_size.strip() if pack_size else '1'  # Ensure pack_size is a string and has no leading/trailing spaces

            # Compile the processed record with the new "Container" field
            processed_record = {
                "VPN": row['Product ID'] if pd.notnull(row['Product ID']) else None,
                "Category": category,
                "Producer Name": producer_name,
                "Product Name": product_name,
                "Container": container,
                "FOB": formatted_price,
                "Volume Amount": volume_amount,
                "Volume Unit": volume_unit,
                "Pack Size": pack_size,
                # Map additional columns E-L
                "Style": row["Style"] if pd.notnull(row["Style"]) else None,
                "ABV": row["ABV"] if pd.notnull(row["ABV"]) else None,
                "Country": row["Country"] if pd.notnull(row["Country"]) else None,
                "Coupler": row["Coupler"] if pd.notnull(row["Coupler"]) else None,
                "UPC/EAN": row["UPC/EAN"] if pd.notnull(row["UPC/EAN"]) else None,
                "COLA": row["COLA"] if pd.notnull(row["COLA"]) else None,
                "Cases per Layer": row["Cases per Layer"] if pd.notnull(row["Cases per Layer"]) else None,
                "Cases per Pallet": row["Cases per Pallet"] if pd.notnull(row["Cases per Pallet"]) else None,
            }
            processed_records.append(processed_record)

    return processed_records

# Example usage
file_path = 'Lime Ventures Price List 2.7.2024 East Coast_for distributors.xlsx'
processed_records = preprocess_and_parse_excel_sheet(file_path)

# Convert the list of dictionaries to a DataFrame
df = pd.DataFrame(processed_records)

# Write the DataFrame to a CSV file
df.to_csv('lime_ventures_fob_parsed.csv', index=False)
