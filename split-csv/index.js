/**
 * Script to split csv from file
 */

// Import section
const XLSX = require('xlsx');
const _ = require('lodash');

// Program
const readArguments = () => {
    const argv = require('minimist')(process.argv.slice(2));
    const argPath = argv.path;
    const argFileName = argv.name;
    const rowsPerSheet = argv.rows | 1000;
    const path = argPath || `./${argFileName}`;
    const emptyHeaderRows = argv.empty || 0;
    const range = argv.range || 1;
    const filter = argv.filter;
    const sheetNumber = argv.sheet || 0;
    const isNoHeader = argv._.includes('no-header');

    return {
        path,
        rowsPerSheet,
        emptyHeaderRows,
        range,
        sheetNumber,
        isNoHeader,
        filter,
    };
};

const getHeader = (path) => {
    const workbook = XLSX.readFile(path, { sheetRows: 1 });
    let config = { defval: null, header: 1 };
    const jsonData = XLSX.utils.sheet_to_json(
        workbook.Sheets[workbook.SheetNames[0]],
        config
    );
    return _.filter(jsonData[0], (v) => v !== null);
};
const XLSXToJSON = (arguments) => {
    const workbook = XLSX.readFile(arguments.path);
    let header = undefined;
    let config = { defval: null };
    if (!arguments.isNoHeader) {
        header = getHeader(arguments.path);
        config.header = header;
        config.range = arguments.range;
    }
    const jsonData = XLSX.utils.sheet_to_json(
        workbook.Sheets[workbook.SheetNames[arguments.sheetNumber]],
        config
    );
    let data = jsonData;
    if (arguments.filterColumn) {
        data = _.filter(jsonData, (v) => {
            const key = Object.keys(jsonData)[arguments.filterColumn];
            return !_.isEmpty(v[key]);
        });
    }

    return {
        header,
        data,
    };
};

const sheetToArray = (data, rowsPerSheet) => {
    const sheetArray = [];
    while (data.length > 0) {
        const chunk = data.splice(0, rowsPerSheet);
        sheetArray.push(chunk);
    }
    return sheetArray;
};

const main = () => {
    const arguments = readArguments();
    const jsonData = XLSXToJSON(arguments);
    const sheetArray = sheetToArray(jsonData.data, arguments.rowsPerSheet);
    let index = 0;
    for (let sheet of sheetArray) {
        // Add dummy data row for header
        for (let i = 0; i < arguments.emptyHeaderRows; ++i) {
            sheet.unshift({});
        }
        const worksheet = XLSX.utils.json_to_sheet(sheet);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet');
        if (!jsonData.header) {
            XLSX.utils.sheet_add_aoa(worksheet, [jsonData.header], {
                origin: 'A1',
            });
        }
        const fileName = arguments.path
            .replace(/^.*[\\\/]/, '')
            .replace(/\.[^/.]+$/, '');
        XLSX.writeFile(workbook, `${fileName}_splitted_${index}.xlsx`);
        index++;
    }
};

main();
