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

<p float="left">
   <img width="32%"  alt="image" src="https://github.com/user-attachments/assets/55a6bf37-e5b0-419b-90bb-eb233fb1f79c" />
<img width="32%" alt="image" src="https://github.com/user-attachments/assets/f8470193-05f2-4f13-9390-a78b1cbc5646" />
<img width="32%"  alt="image" src="https://github.com/user-attachments/assets/c03a5d12-439c-4e37-856f-126b23e0b677" />
</p>
<p float="left">
<img width="48%"alt="image" src="https://github.com/user-attachments/assets/d30ce6e5-3e6d-48d1-9159-acf93dff3363" />
<img width="48%" alt="image" src="https://github.com/user-attachments/assets/1e79c2f1-e69b-4beb-91a5-f0970daf05dd" />
</p>

If you look at the formula regions above, it correctly unions cells that have isomorphic dependency structures.


## Instructions
This is guaranteed to run on MacOS, though Windows should be easier. If you are on Linux, sorry. There is no production build. Running the development build requires a node and python installation. Unfortunately, due to Microsoft Same-Origin issues, you need a SSL cert. for your localhost. Before starting, place a `.env` file in `/backend` with `OPENAI_API_KEY="..."`. Don't worry, token usage is very small. 
- Run `./scripts/setup.sh` to configure everything, including this.
- Run `./scripts/start-server.sh` to start the backend dev server.
- Run `./scripts/start-plugin.sh` to sideload the Excel plugin.
<p float="left">
  <img src="https://github.com/user-attachments/assets/2572edfb-68c9-4e84-923d-64f2206f6a00" width="45%" />
  <img src="https://github.com/user-attachments/assets/3967a1a6-5550-4491-9fbb-d709fbf1d3ce" width="45%" />
</p>

- Click the `Show Task Pane` button in the top left of Excel. Go to any sheet and click analyze. 
- Check invalidate to force reset cache for the sheet. Check debug to see some pretty backend plots.
- Set the highlight to `Table` and move your cursor around cells to see the active table. Set the highlight to `Formula` to see linked formula blocks. 

