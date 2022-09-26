# Split csv

Split CSV/XLSX file into multiple files

## Install
``` bash
$ npm install
```
## Usage
``` bash
$ node index.js [options]
```
<details> <summary> Options </summary>
  
``` bash
  -path (string): Path to file
  -name (string): File name (if path is not specified)
  -rows (number): Number of rows per file, default: 1000
  -empty (number): Number of empty rows to add below header, default: 0
  -filter (number): Column index that the row will be removed if data is empty at that column, default: no
  -sheet (number): Ordinal number of the sheet in file, default: 0
  --no-header (flag)
```
</details>

Example:

```bash
$ node index.js -path ./example.xlsx -rows 2000 -filter 0
```
