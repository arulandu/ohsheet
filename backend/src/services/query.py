import openai
from api.models import SheetCacheData
from services.prompter import query_prompt_template
from utils.parser import ExcelBook
from utils.path import extract_save_dir


def format_values_for_prompt(values):
    """
    Format raw Excel values into the prompt format
    """
    if values is None:
        return '` `'

    if isinstance(values, list):
        if len(values) == 0:
            return '` `'
        if isinstance(values[0], list):
            # 2D array - flatten it
            flat_values = []
            for row in values:
                for cell in row:
                    flat_values.append(cell)
            return ', '.join([f'`{str(cell)}`' if cell is not None else '` `' for cell in flat_values])
        else:
            # 1D array
            return ', '.join([f'`{str(cell)}`' if cell is not None else '` `' for cell in values])
    else:
        # Single value
        return f'`{str(values)}`' if values is not None else '` `'

def get_formatted_table_data(excel_book: ExcelBook, sheet_index: int, tables: list):
    """
    Get formatted table data for the prompt
    """
    table_data = []
    
    for table in tables:
        table_info = []
        
        # Get data range
        if table.data:
            data_values = excel_book.get_range_values(sheet_index, table.data)
            data_str = format_values_for_prompt(data_values)
            table_info.append(f"{table.data} [{data_str}]")
        else:
            table_info.append("")
        
        # Get row header range
        if table.row_hdr:
            row_values = excel_book.get_range_values(sheet_index, table.row_hdr)
            row_str = format_values_for_prompt(row_values)
            table_info.append(f"{table.row_hdr} [{row_str}]")
        else:
            table_info.append("")
        
        # Get column header range
        if table.col_hdr:
            col_values = excel_book.get_range_values(sheet_index, table.col_hdr)
            col_str = format_values_for_prompt(col_values)
            table_info.append(f"{table.col_hdr} [{col_str}]")
        else:
            table_info.append("")
        
        table_data.append('; '.join(table_info))
    
    return '\n'.join(table_data)

def query(prompt, cell_address, sheet_cache: SheetCacheData, file_path: str, sheet_index: int, save=False):
    """
    Process a query about a specific cell using the detected tables
    """
    try:
        print("query", prompt)
        excel_book = ExcelBook(file_path)
        # Get formatted table data
        tables_data = get_formatted_table_data(excel_book, sheet_index, sheet_cache.tables)
        
        if not tables_data:
            return "Error: Could not read table data from Excel file"
        
        formatted_prompt = query_prompt_template.format(
            tables_data,
            prompt,
            cell_address
        )

        save_dir = extract_save_dir(file_path)
        if save:
            with open(f'{save_dir}/{sheet_cache.id}_query.txt', "w") as f:
                f.write(formatted_prompt)
            
        client = openai.OpenAI()
        model = 'gpt-4.1'
        
        response = client.responses.create(
            model=model,
            input=[
                {
                    "role": "user",
                    "content": formatted_prompt
                }
            ]
        )
        return response.output_text
        
    except Exception as e:
        return f"Error processing query: {str(e)}"