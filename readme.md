# json-2-csv-2-xlsx

CouchDB JSON to Excel

This project converts two JSON files (scraped from CouchDB) to csv and then creates Excels worksheets for each file into
a single Excel .xslx. This project is specifically coded for my own needs: filenames are hard-coded; however, it should
not be too hard to modify to suit others' needs.

# Dependencies:
 - https://github.com/mrodrig/json-2-csv - to convert JSON to csv
 - https://github.com/SheetJS/js-xlsx - to create the Excel worksheets

# Thanks!

Many thanks to http://sheetjs.com/ for the useful examples and code!

# Installation

    npm install

Copy your JSON into the two files, A.json and B.json.

To process the files, enter

    node index.js

The output file is named report.xlsx.
