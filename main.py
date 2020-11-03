from parse import Parser
from pprint import pprint

file_path = "docs/Demo_Assessment_Model_08.18.20.xlsx"
sheet_name = "KPI Dashboard"

kpi_parser = Parser(file_path, sheet_name)

schema = kpi_parser.get_category_schema()


def output_to_json(data):
    import json
    import os
    loc = "refactor.json"
    with open(loc, "w") as output_file:
        json.dump(data, output_file, indent=4)
        output_file.close()
    print(f"File saved to: {os.path.abspath(loc)}")


pprint(schema)

output_to_json(schema)
