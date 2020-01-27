const MVLoaderBase = require('mvloader/src/mvloaderbase');
const XLSX = require('xlsx');

class MVLExcelController extends MVLoaderBase {

    constructor (...config) {
        let locaDefaults = {
            sheetName: 'sheet',
        };
        super(locaDefaults, ...config);
    }

    async init() {
        super.init();
        this.Excel = XLSX;
    }

    writeFile (filename, data = []) {
        let workbook = XLSX.utils.book_new();
        if (Array.isArray(data)) {

        }

        let ws_name = this.config.sheetName;
        let workSheet = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(workbook, workSheet, ws_name);
        return XLSX.writeFile(workbook, filename);
    }

}

module.exports = MVLExcelController;