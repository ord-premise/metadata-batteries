import pandas as pd
import json

file_path = # Give the path to the input Excel metadata file.
jsonld_file_path = # Give the path to the output jsonLD file. 

def create_jsonld_with_conditions(schemas, item_id_map, connector_id_map):
    jsonld = {
        "@context": {
            "@vocab": "http://emmo.info/electrochemistry#",
            "xsd": "http://www.w3.org/2001/XMLSchema#"
        },
        "@graph": []
    }

    top_level_items = {}

    for _, row in schemas.iterrows():
        if pd.isna(row['Value']) or row['Ontology link'] == 'NotOntologize':
            continue  # Skip rows with empty value or 'NotOntologize' in Ontology link

        ontology_links = row['Ontology link'].split('-')
        if len(ontology_links) < 5:
            continue  # Skip rows with insufficient parts in Ontology link

        value = row['Value']
        unit = row['Unit']

        main_item = ontology_links[0]
        relation = ontology_links[1]
        sub_item = ontology_links[2]
        property_relation = ontology_links[3]
        property_item = ontology_links[4]

        if main_item not in top_level_items:
            top_level_items[main_item] = {
                "@type": main_item,
                relation: []
            }
            jsonld["@graph"].append(top_level_items[main_item])

        sub_item_structure = {
            "@type": sub_item,
            property_relation: [{
                "@type": property_item,
                "hasValue": value,
                "hasUnits": unit
            }]
        }

        jsonld["@context"].update({
            main_item: {"@id": item_id_map.get(main_item, ""), "@type": "@id"},
            sub_item: {"@id": item_id_map.get(sub_item, ""), "@type": "@id"},
            property_item: {"@id": item_id_map.get(property_item, ""), "@type": "@id"}
        })

        top_level_items[main_item][relation].append(sub_item_structure)

    return jsonld

# Load the provided Excel file
excel_data = pd.ExcelFile(file_path)

# Read the necessary sheets
schemas = pd.read_excel(excel_data, sheet_name='Schemas')
item_id_map = pd.read_excel(excel_data, sheet_name='Ontology - item').set_index('Item')['ID'].to_dict()
connector_id_map = pd.read_excel(excel_data, sheet_name='Ontology - Connector').set_index('Item')['ID'].to_dict()

# Create the updated JSON-LD structure
jsonld_output_updated = create_jsonld_with_conditions(schemas, item_id_map, connector_id_map)

# Save the updated JSON-LD output to a file
with open(jsonld_file_path, 'w') as file:
    json.dump(jsonld_output_updated, file, indent=4)
