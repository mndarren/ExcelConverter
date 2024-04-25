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
import tkinter as tk
import warnings
from datetime import datetime
from inspect import getframeinfo, stack
from typing import Dict

from dicttoxml import dicttoxml
from pandas import read_excel

TIME_FORMAT = '%m/%d/%Y %H:%M:%S'

CONFIG_FILE = "config.json"
SHEET_NAME = "Test Result"
INPUT_FILE = "input"
OUTPUT_JSON = "output_json"
OUTPUT_XML = "output_xml"
HEADER_ROWS = "header_rows"

WINDOW_SIZE = "800x400"
WINDOW_TITLE = "Excel Converter"
BUTTON_TEXT = "Run"
LABEL_TEXT = "Show log data here"
LABEL_WRAP_LENGTH = 700

log_text = ""
start_time = datetime.now()


def start_log():
    """print out start time"""
    global start_time
    start_time = datetime.now()
    return print_log(f"Starting program: {getframeinfo(stack()[1][0]).filename}") + break_section()


def print_log(msg=''):
    """
    Print out the log message
    :param msg: Log message
    """
    caller = getframeinfo(stack()[1][0])
    return f"[{_curr_time()}][{caller.function} - {caller.lineno}]: {msg}\n"


def stop_log():
    """print out duration time"""
    duration = datetime.now() - start_time
    return break_section() + print_log(f"Done to spend time: {str(duration)}")


def _curr_time():
    """
    return: current time
    """
    return datetime.now().strftime(TIME_FORMAT)


def break_section():
    """
    :return: break section lines
    """
    return f"{'-' * 120}\n"


def update_label():
    global label
    label["text"] = log_text


def get_args():
    """
    Get arguments, including input file and output filename.
    The default path is the current path.
    """
    try:
        global log_text
        with open(CONFIG_FILE, 'r') as fh:
            json_data = json.load(fh)
        log_text = log_text + f"{print_log(json_data)}"
        update_label()
        return json_data
    except IOError:
        raise IOError(f"IOError: Cannot open file {CONFIG_FILE}")
    except Exception:
        raise Exception(f"Exception: Cannot open file {CONFIG_FILE}")


def main(input_path=None):
    """
    Convert data to JSON and XML
    """
    global log_text
    log_text = start_log()
    update_label()
    config_dict = get_args()
    if input_path is None:
        input_path = config_dict[INPUT_FILE]
    # ignore the openpyxl UserWarning
    with (warnings.catch_warnings()):
        warnings.filterwarnings("ignore", category=UserWarning, module=re.escape('openpyxl.styles.stylesheet'))
        excel_data_header = read_excel(input_path, sheet_name=SHEET_NAME, header=None,
                                       nrows=config_dict[HEADER_ROWS], names=['obj', 'value'])
        header_dict = get_json_header(excel_data_header)
        excel_data_df = read_excel(input_path, sheet_name=SHEET_NAME,
                                   skiprows=lambda x: 0 <= x <= config_dict[HEADER_ROWS]
                                   ).dropna(how='all').dropna(how='all', axis=1)

    json_data_str = excel_data_df.to_json(orient='records')

    print_dict = {'header': header_dict, 'data': json.loads(json_data_str)}
    json_str = json.dumps(print_dict)
    with open(config_dict[OUTPUT_JSON], 'w') as outfile:
        outfile.write(json_str)
    log_text = log_text + print_log("Generated JSON file")
    update_label()

    # xml_str = excel_data_df.to_xml(attr_cols=COL_LIST)
    my_item_func = lambda x: 'test_case'
    xml_str = dicttoxml(print_dict, custom_root='TestResult', attr_type=False, item_func=my_item_func).decode()
    with open(config_dict[OUTPUT_XML], 'w') as outfile:
        outfile.write(xml_str)
    log_text = log_text + print_log("Generated XML file")
    update_label()
    log_text = log_text + stop_log()
    update_label()


def get_json_header(excel_data_header) -> Dict:
    """
    Convert the dataframe to dict
    """
    header_dict: [Dict] = {}
    for index, row in excel_data_header.iterrows():
        header_dict[row['obj']] = row['value']

    return header_dict


window = tk.Tk()
window.geometry(WINDOW_SIZE)
window.title(WINDOW_TITLE)
frame = tk.Frame(window)
frame.pack()

# widgets
button = tk.Button(frame, text=BUTTON_TEXT, command=lambda: main())
button.pack()

label = tk.Label(frame, text=LABEL_TEXT, justify="left", anchor="w", wraplength=LABEL_WRAP_LENGTH)
label.pack()

window.mainloop()
