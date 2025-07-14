import openai
from api.models import SheetData, CellData, SheetCacheData
from masker import get_regions
from attribs import *
from prompter import header_prompt, table_detection_prompt
from utils.conversion import address_to_coord
from utils.plot import plot_regions, plot_ranges, plot_tables
import matplotlib.pyplot as plt
import numpy as np
import json
import traceback

def get_values_in_range(cells: list, rng: str):
    """
    Extract values from a range of cells
    """
    def format_value(c: CellData):
        v = c.value
        if v is None:
            return ''
        else: 
            return str(v)

    st, end = rng.split(':')
    (r1, c1), (r2, c2) = address_to_coord(st), address_to_coord(end)
    data = []
    for r in range(r1, r2+1):
        for c in range(c1, c2+1):
            if r < len(cells) and c < len(cells[0]):
                data.append(format_value(cells[r][c]))
            else:
                data.append('')
    
    return data

def parse_ranges_response(s):
    """
    Parse the header ranges response from GPT
    """
    s = [r.strip() for r in s.strip('[]').split(',')]
    return s

def parse_table_detection_response(s):
    """
    Parse the table detection response from GPT
    """
    s = [r.strip('() ').split(',') for r in s.strip().split(';')]
    tables = []
    for r in s:
        if len(r) >= 3:
            table = {
                'data': r[0].strip(),
                'row_hdr': r[1].strip(),
                'col_hdr': r[2].strip()
            }
            tables.append(table)
    return tables

def detect_tables(sheet: SheetData):
    """
    Detect tables in an Excel sheet using OpenAI API
    """
    client = openai.OpenAI()
    model = 'gpt-4.1'
    
    try:
        font_rgs = get_regions(sheet, key=font_key)
        format_rgs = get_regions(sheet, key=format_key)
        formula_rgs = get_regions(sheet, key=formula_key, cmp=formula_cmp)
        color_rgs = get_regions(sheet, key=color_key)

        fig, (ax1, ax2, ax3, ax4) = plt.subplots(1, 4, figsize=(5*4, 10))
        shape = sheet.shape
        plot_regions(font_rgs, shape, 'Text', ax1)
        plot_regions(format_rgs, shape, 'Format', ax2)
        plot_regions(formula_rgs, shape, 'Formula', ax3)
        plot_regions(color_rgs, shape, 'Color', ax4)
        plt.show()
        
        header_prompt_text = header_prompt([font_rgs, format_rgs, formula_rgs, color_rgs])
        
        header_response = client.responses.create(
        model=model,
        input=[
            {
                "role": "user",
                "content": header_prompt_text
            }
        ]
        )
        
        info_ranges = parse_ranges_response(header_response.output_text)
        print(info_ranges)
        fig, (ax1) = plt.subplots(1, 1, figsize=(4, 4))
        plot_ranges(info_ranges, shape, 'Requested Header/Info Ranges', ax1)
        plt.show()
        
        header_data = {rng: '[' + ', '.join(get_values_in_range(sheet.data, rng)) + ']' for rng in info_ranges}
        header_data_input = '\n'.join([f'{h[0].replace("$", "")} {h[1]}' for h in header_data.items()])

        table_detection_prompt_text = table_detection_prompt(header_data_input)
        
        table_response = client.responses.create(
            model=model,
            previous_response_id=header_response.id,
            input=[
                {
                    "role": "user",
                    "content": table_detection_prompt_text
                }
            ]
        )

        tables = parse_table_detection_response(table_response.output_text)
        print(tables)

        fig, (ax1) = plt.subplots(1, 1, figsize=(4, 4))
        plot_tables(tables, shape, 'Tables', ax1)
        plt.show()

        return info_ranges, tables
        
    except Exception as e:
        print(f"Error detecting tables: {e}")
        traceback.print_exc()
        return []