import argparse
import pandas as pd
import json

def create_jsonld_with_conditions(schemas, item_map, unit_map, context_toplevel, context_connector):
    jsonld = {
        "@context": {},
        "Battery": {
            "@type": "battery:Battery"
        }
    }

    # Build the @context part
    for _, row in context_toplevel.iterrows():
        jsonld["@context"][row['Item']] = row['Key']

    connectors = set()
    for _, row in context_connector.iterrows():
        jsonld["@context"][row['Item']] = row['Key']
        connectors.add(row['Item'])  # Track connectors to avoid redefining types

    # Helper function to add nested structures with type annotations
    def add_to_structure(path, value, unit):
        current_level = jsonld["Battery"]
        # Iterate through the path to create or navigate the structure
        for idx, part in enumerate(path[1:]):
            is_last = idx == len(path) - 2  # Check if current part is the last in the path

            if part not in current_level:
                if part in connectors:
                    current_level[part] = {}
                else:
                    if part in item_map:
                        current_level[part] = {"@type": item_map[part]['Key']}
                    else:
                        raise ValueError(f"Connector or item '{part}' is not defined in any relevant sheet.")

            if not is_last:
                current_level = current_level[part]
            else:
                # Handle the unit and value structure for the last item
                final_type = item_map.get(part, {}).get('Key', '')
                if unit != 'No Unit':
                    if pd.isna(unit):
                        raise ValueError(f"The value '{value}' is filled in the wrong row, please check the schemas")
                    unit_info = unit_map[unit]
                    current_level[part] = {
                        "@type": final_type,
                        "hasNumberValue": {
                            "@type": "emmo:hasNumberValue",
                            "value": value,
                            "unit": {
                                "label": unit_info['Label'],
                                "symbol": unit_info['Symbol'],
                                "@type": unit_info['Key']
                            }
                        }
                    }
                else:
                    current_level[part] = {
                        "@type": final_type,
                        "value": value
                    }

    # Process each schema entry to construct the JSON-LD output
    for _, row in schemas.iterrows():
        if pd.isna(row['Value']) or row['Ontology link'] == 'NotOntologize':
            continue
        if pd.isna(row['Unit']):
            raise ValueError(f"The value '{row['Value']}' is filled in the wrong row, please check the schemas")

        ontology_path = row['Ontology link'].split('-')
        add_to_structure(ontology_path, row['Value'], row['Unit'])

    return jsonld

def convert_excel_to_jsonld(excel_file):
    excel_data = pd.ExcelFile(excel_file)
    
    schemas = pd.read_excel(excel_data, 'Schemas')
    item_map = pd.read_excel(excel_data, 'Ontology - Item').set_index('Item').to_dict(orient='index')
    unit_map = pd.read_excel(excel_data, 'Ontology - Unit').set_index('Item').to_dict(orient='index')
    context_toplevel = pd.read_excel(excel_data, '@context-TopLevel')
    context_connector = pd.read_excel(excel_data, '@context-Connector')

    jsonld_output = create_jsonld_with_conditions(schemas, item_map, unit_map, context_toplevel, context_connector)
    
    return jsonld_output

def main():
    parser = argparse.ArgumentParser(description='Convert an Excel file to JSON-LD format.')
    parser.add_argument('--path_to_excel_file', required=True, help='Path to the Excel file to convert.')

    args = parser.parse_args()

    jsonld = convert_excel_to_jsonld(args.path_to_excel_file)
    jsonld_file_path = args.path_to_excel_file.replace('.xlsx', '.json')

    with open(jsonld_file_path, 'w') as f:
        json.dump(jsonld, f, indent=4)
    
    print(f"JSON-LD file has been saved to {jsonld_file_path}")

if __name__ == "__main__":
    main()
