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
    [x] Table detection
        * GPT to detect headers first
        * Send header data and ask for a list of tables, each table has a data range + row label range + column label range 
        * Validate dimensions matching. Store the headers and the tables. 
2. Plugin
    - [x] On load, make a request to the python server running locally, sending all the data. Python should cache the tables into a .sai file. Plugin uses this api to query against headers etc.
    - Functionality: sheet select: show related sheets. on cell change: show connected sheet cells in UI, use backend to pull info abt the connection using the header. 
    - Chat functionality for asking questions in the spreadsheet. No chat history. Question hits python server, uses the .sai cache. 
2.  Performance optimization.
    - Handroll formatting support to pyxlsb. COM call issue, want to actually support xlsb and not xlsx. Appscript. Don't want ppl to have to install LibreOffice. 
    - Write the fastest xlsb parser imaginable in Rust or C++
3. Extras
    - Value to format parser. Should be easy. Like Spreadsheet LLM

Main thing:
- Figure out a way to improve table detection. It kinda sucks right now, unfort

### Architecture Constraints
Plugins run sandboxed and can't file save. Use GPT to save to index file. 

COM
- 10.4s/728 -> 142ms/cell
- You can only batch COM value, formulas, and numberformat. You can't get rendered text.

Excel JS API
- "46731.00""ms" for 419898 -> 0.11 ms/cell (format / fill data)
- Can batch everything except color / font. Can batch get rendered text. 

