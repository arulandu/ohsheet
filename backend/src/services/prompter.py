from api.models import RegionData

header_prompt_template = """You are a table detection expert. You are given information describing the structure of an Excel spreadsheet. Your task is to output Excel ranges corresponding to the headers. 

The input contains 2 structural attributes: "Format", "Color". For each *attribute*, you will be given a newline separated list of regions.
A region represents a contiguous region of spreadsheet cells that have nearly the same attribute value.
Each region is specified as a semicolon-separated tuple of the value of the attribute as a backtick-quoted string along with a comma-separated list of Excel ranges whose union is the region e.g. (`value` ; ['A1:B3', 'C1:D3']).
Ranges are specified like 'A1:B3' which selections the first to second row and first to third column, inclusive. 

Format regions are grouped by a unique key encoding the number format and font of the cell. Color regions are grouped by background cell color. 

Your task is to output a comma separated list, surrounded by [], of non-overlapping Excel ranges e.g. 'A1:B3' corresponding to the header regions of the data. An overestimate is preferred to an underestimate. You should select any regions that may contain information for describing the data, but you should omit regions that contain the data itself.
DO NOT ADD OTHER WORDS OR EXPLANATION.

INPUT:

*Format*
{}

*Color*
{}
"""

def encode_attribute(rgs):
    def encode_region(rg):
        val = rg["val"].replace("Table", "").replace("table", "")
        return f'(`{val}` ; [{", ".join(rg["ranges"]).replace("$", "")}])'

    s = '\n'.join([encode_region(rg) for rg in rgs])
    return s

def header_prompt(regions: RegionData):
    encodings = [encode_attribute(regions.format), encode_attribute(regions.color)]
    return header_prompt_template.format(*encodings)

table_detection_template = """
You are now given the data corresponding to the header/informative regions that you outputted in your previous response.
The ranges are newline separated. For each range, you are given the range, a space, and a comma separated list of backtick-quoted cell values in row-major order where e.g. A1:A3 [`value1`, , `value2`] implies that the second cell is empty. 

Using this information and the previous attribute region information, your task is to detect every table present in the spreadsheet.
You should output a semicolon-separated list of tables. For each table, you should give a comma-separated, parenthesized tuple of three ranges: the data range, the row header range, and the column header range. If either of these headers are blank, you should leave it blank in the tuple: e.g. (A2:D7, A2:D2, ) specifies a blank column header. 

DO NOT ADD OTHER WORDS OR EXPLANATION.

INPUT:
{}
"""

def table_detection_prompt(header_data_input):
    return table_detection_template.format(header_data_input)

query_prompt_template = """
You are a spreadsheet genius. Your input consists of a set of new-line separated list of tables, a query, and a cell address. Your job is to answer the query regarding the cell address using the information provided to you in the tables.

Each table consists of three semicolon-separated segments: the data segment, the row header segment, and the column header segment. 
Each segment is given by the range followed by a space and a comma-separated list of backtick-quoted cell values in row-major order e.g. A1:A3 [`value1`, `value2`, `value3`]. 

You should contemplate what the cell is, what quantities in the spreadsheet it is related to, and any other information according to the query. If you are unsure with high confidence, you should state that you do not know with certainty. ONLY OUTPUT THE ANSWER AS PLAIN TEXT.

INPUT:

TABLES:
{}

QUERY:
{}

CELL ADDRESS:
{}
"""