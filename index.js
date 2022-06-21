const excel = require('exceljs');

const workbook = new excel.Workbook();

//const filePath = 'grm_MOS3_2.xlsx';
const filePath = 'ExcelTest1.xlsx';

const handleSheet = (sheet) => {
    
    console.log(sheet.name);

    for(let rowIndex = 1; rowIndex <= sheet.rowCount; rowIndex++) {
        let row = sheet.getRow(rowIndex);
        let cell = row.getCell(1);
        let cellValue = cell.value;
        let str = cellValue?.toString();
        console.log(str);
    }

};

const onWorkbookOpenSuccess = () => {
    console.log('Открыт успешно.');

    workbook.worksheets.forEach(handleSheet);
};

const onWorkbookOpenError = reason => console.log(reason);

workbook.xlsx.readFile(filePath).then(onWorkbookOpenSuccess, onWorkbookOpenError);