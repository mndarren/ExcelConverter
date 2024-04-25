#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
current_module:: excel2json_xml.py
created_by:: Darren Xie
created_on:: 04/04/2021

Convert Excel data to XML data
"""
import json
import re
import warnings
from typing import Dict

from dicttoxml import dicttoxml
from pandas import read_excel
from speedup_work_lib.simple_log import SimpleLog

CONFIG_FILE = "config.json"
SHEET_NAME = "Test Result"
INPUT_FILE = "input"
OUTPUT_JSON = "output_json"
OUTPUT_XML = "output_xml"
HEADER_ROWS = "header_rows"
COL_LIST = ["Step", "Task", "ExpectedResult", "PassFail", "Note", "AppComments", "AutoSimFunction", "Control",
            "Timestamp"]


class Excel2JsonXml(SimpleLog):
    """Parse Excel data to XML data"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def get_args(self):
        """
        Get arguments, including input file and output filename.
        The default path is the current path.
        """
        try:
            with open(CONFIG_FILE, 'r') as fh:
                json_data = json.load(fh)
            self.print_log(json_data)
            return json_data
        except IOError:
            raise IOError(f"Cannot open file {CONFIG_FILE}")
        except Exception:
            raise Exception(f"Cannot open file {CONFIG_FILE}")

    def main(self):
        """
        Convert data to JSON and XML
        """
        self.start_log()
        config_dict = self.get_args()
        # ignore the openpyxl UserWarning
        with (warnings.catch_warnings()):
            warnings.filterwarnings("ignore", category=UserWarning, module=re.escape('openpyxl.styles.stylesheet'))
            excel_data_header = read_excel(config_dict[INPUT_FILE], sheet_name=SHEET_NAME, header=None,
                                           nrows=config_dict[HEADER_ROWS], names=['obj', 'value'])
            header_dict = self.get_json_header(excel_data_header)
            excel_data_df = read_excel(config_dict[INPUT_FILE], sheet_name=SHEET_NAME,
                                       skiprows=lambda x: 0 <= x <= config_dict[HEADER_ROWS]
                                       ).dropna(how='all').dropna(how='all', axis=1)

        json_data_str = excel_data_df.to_json(orient='records')

        print_dict = {'header': header_dict, 'data': json.loads(json_data_str)}
        json_str = json.dumps(print_dict)
        with open(config_dict[OUTPUT_JSON], 'w') as outfile:
            outfile.write(json_str)
        self.print_log("Generated JSON file")

        # xml_str = excel_data_df.to_xml(attr_cols=COL_LIST)
        my_item_func = lambda x: 'test_case'
        xml_str = dicttoxml(print_dict, custom_root='TestResult', attr_type=False, item_func=my_item_func).decode()
        with open(config_dict[OUTPUT_XML], 'w') as outfile:
            outfile.write(xml_str)
        self.print_log("Generated XML file")
        self.stop_log()

    def get_json_header(self, excel_data_header) -> Dict:
        """
        Convert the dataframe to dict
        """
        header_dict: [Dict] = {}
        for index, row in excel_data_header.iterrows():
            header_dict[row['obj']] = row['value']

        return header_dict


if __name__ == "__main__":
    Excel2JsonXml().main()
