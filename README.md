# Oh Sheet!
A simple way to grok spreadsheet relationships! There is a bit of LLM use, but most of the magic is smart caching and aggregation for table detection.

## Problem
LLMs and large spreadsheets, think ~500k rows, don't mix. You can't just embed the whole spreadsheet, but wouldn't it be nice to answer queries about arbitrary data across arbitrary files? This is the end goal, and this project makes some steps toward it. 

Context windows and more importantly needle-in-a-haystack accuracy limits require us to create efficient abstractions of spreadsheet data. Oh Sheet! does exactly this. While there are a few performance bottlenecks, funnily enough, this is due to latency in COM calls / Excel's query API and not due to tool calls or reasoning:

COM
- 10.4s/728 -> 142ms/cell
- You can only batch COM value, formulas, and numberformat. You can't get rendered text.

Excel JS API
- "46731.00""ms" for 419898 -> 0.11 ms/cell (format / fill data)
- Can batch everything except color / font. Can batch get rendered text. 

With a custom .xlsb parser in the future, Oh Sheet! could do something more meaningful. 

## Solution
Oh Sheet! ships with a plugin to avoid the 100x COM latency bottleneck. The plugin parses the sheet with Office.js and sends the data to a local FastAPI server which handles the rest. Features:
1. Table detection
    * DFS with clever keys and comparators to aggregate spreadsheet into regions. Segment regions into rectangular ranges, which are sent to GPT 4.1 Spreadsheet-LLM style "ish". No actual data is sent here. 
    * GPT 4.1 first detects the headers / descriptive ranges. Data is then sent for these ranges, and GPT 4.1 detects the table ranges with row/col headers.
2. Formula clustering
    * A key and comparator inspired by git diffs!
3. Formula dependency querying
    * Given a cell, we simply look up the connected formula component for itself and each precedent/dependent. 
4. Querying [TODO]
    * Send each table ranges in a sheet along with data for just the headers + queried cell address + question. Return a valid Excel formula.

Most importantly, Oh Sheet! caches most of the pre-computation in an `.osht` file. The vision is that once the files are indexed (e.g. regions + table ranges), queries in the future can be much faster. 

## Instructions
This is guaranteed to run on MacOS, though Windows should be easier. If you are on Linux, sorry. There is no production build. Running the development build requires a node and python installation. Unfortunately, due to Microsoft Same-Origin issues, you need a SSL cert. for your localhost. 
- Run `./scripts/setup.sh` to configure everything, including this.
- Run `./scripts/start-server.sh` to start the backend dev server.
- Run `./scripts/start-plugin.sh` to sideload the Excel plugin.
