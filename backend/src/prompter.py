header_prompt_template = """You are a table detection expert. You are given information describing the structure of an Excel spreadsheet. Your task is to output Excel ranges corresponding to the headers. 

The input contains 5 structural attributes: "Text", "Value Format", "Formula", "Color". For each *attribute*, you will be given a newline separated list of regions.
A region represents a contiguous region of spreadsheet cells that have nearly the same attribute value.
Each region is specified as a semicolon-separated tuple of the value of the attribute as a backtick-quoted string along with a comma-separated list of Excel ranges whose union is the region e.g. (`value` ; ['A1:B3', 'C1:D3']).
Ranges are specified like 'A1:B3' which selections the first to second row and first to third column, inclusive. 

Text regions account for differences in font as well. Value Format is the number format applied to the cell value. A Formula region contains cells whose formula only differs from the region value by cell reference changes, common when applying a formula to a range of cells. Color is the background cell color. 

Your task is to output a comma separated list, surrounded by [], of non-overlapping Excel ranges e.g. 'A1:B3' corresponding to the header regions of the data. An overestimate is preferred to an underestimate. You should select any regions that may contain information for describing the data, but you should omit regions that contain the data itself.
DO NOT ADD OTHER WORDS OR EXPLANATION.

INPUT:

*Text*
{}

*Value Format*
{}

*Formula*
{}

*Color*
{}
"""

def encode_attribute(rgs):
    def encode_region(rg):
        return f'(`{rg["val"]}` ; [{", ".join(rg["ranges"])}])'

    s = '\n'.join([encode_region(rg) for rg in rgs])
    return s

def header_prompt(rgss):
    encodings = [encode_attribute(rg) for rg in rgss]
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
