# Challenge
Understand and query relationships between potentially multiple Excel files. 

You can just embed the whole spreadsheet. Issue is spreadsheets are large, ~500k rows (200K tokes). infini-attention etc only gets you to 500K

Solution: 
- parse and aggregate into tables with metadata.
    - a table is a group of columns, each col has a label, a start / end cell, a orientation direction
- turn formulas between cells into relationships between columns / variables. 
- visualize variable dependencies

Step 1: Spreadsheet compression
Data types only, compress like cells into ranges, get headers, dump it into LLM and ask for the header range and data range. There can be both horizontal and veritcal headers. 

There is no formula usage in spreadsheet LLM. Formula string parser to see what is column dependent or not?

## Gameplan
1. Get e2e deliverable to work
    - Table detection
        * GPT to detect headers first
        * Send header data and ask for a list of tables, each table has a data range + row label range + column label range 
        * Validate dimensions matching. Store the headers and the tables. 
    - Queries
        * Send header data, table ranges. Given question, return some formula. 
        * Won't handle multi-sheet. So, you do table detection for each sheet first.  
2.  Performance optimization.
    - Handroll formatting support to pyxlsb. COM call issue, want to actually support xlsb and not xlsx. Appscript. Don't want ppl to have to install LibreOffice. 
    - Write the fastest xlsb parser imaginable in Rust or C++
3. Extras
    - Value to format parser. Should be easy. Like Spreadsheet LLM