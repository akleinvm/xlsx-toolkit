import JSZip, { file } from 'jszip';
import ExcelStyles from './ExcelStyles';
import ExcelSharedStrings from './ExcelSharedStrings';
import ExcelWorkbook from './ExcelWorkbook';
import ExcelWorksheet from './ExcelWorksheet';

export default class ExcelDocument {
    private sharedStrings!: ExcelSharedStrings | null | undefined;
    private styles!: ExcelStyles;
    private workbook!: ExcelWorkbook;
    private worksheets!: Map<number, ExcelWorksheet>;
    private zipFiles!: JSZip | null;

    public async loadXLSX (arrayBuffer: ArrayBuffer) {
        console.log('Extracting xlsx document...');
        const zip = new JSZip();
        await zip.loadAsync(arrayBuffer);

        console.log('Retrieving files...');
        this.zipFiles = zip.folder("");
        if(!this.zipFiles) throw new Error('no file found');
        this.zipFiles.remove('xl/calcChain.xml');

        console.log('Parsing workbook.xml');
        this.workbook = new ExcelWorkbook();
        const fileWorkbook = this.zipFiles.file('xl/workbook.xml');
        if(!fileWorkbook) throw new Error('workbook not found');
        const xmlWorkbook = await fileWorkbook.async('binarystring');
        if(!xmlWorkbook) throw new Error('workbook is null or undefined');
        this.workbook.fromXML(xmlWorkbook);

        console.log('Parsing sharedStrings.xml...');
        this.sharedStrings = new ExcelSharedStrings();
        const fileSharedStrings = this.zipFiles.file('xl/sharedStrings.xml');
        if(!fileSharedStrings) {
            console.log('sharedStrings not found')
        }
        else {
            const xmlSharedStrings = await fileSharedStrings.async('binarystring');
            if(!xmlSharedStrings) throw new Error('sharedStrings is null or undefined');
            this.sharedStrings.fromXML(xmlSharedStrings)
        }

        console.log('Parsing styles.xml...');
        this.styles = new ExcelStyles();
        const fileStyles = this.zipFiles.file("xl/styles.xml");
        if(!fileStyles) throw new Error('styles not found');
        const xmlStyles = await fileStyles.async('binarystring');
        if(!xmlStyles) throw new Error('styles is null or undefined');
        this.styles.fromXML(xmlStyles);

        this.worksheets = new Map();
    }

    public async getWorksheet (sheetNo: number): Promise<ExcelWorksheet> {
        if(!this.zipFiles) throw new Error('no file found');

        console.log(`Parsing sheet${sheetNo}.xml...`);

        const fileWorksheet = this.zipFiles.file(`xl/worksheets/sheet${sheetNo}.xml`);
        if(!fileWorksheet) throw new Error('worksheet not found');
        const xmlWorksheet = await fileWorksheet.async('binarystring');
        if(!xmlWorksheet) throw new Error('worksheet is null or undefined');
        const worksheet = new ExcelWorksheet(this.sharedStrings);
        worksheet.fromXML(xmlWorksheet);
        
        const worksheetRowNo = worksheet.rowsMap.size;
        const worksheetCellNo = worksheet.cellsMap.size;
        console.log(`${worksheetRowNo} row(s) and ${worksheetCellNo} cell(s) detected.`);
        
        this.worksheets.set(sheetNo, worksheet);
        return worksheet;
    }

    public async saveXLSX (): Promise<ArrayBuffer> {
        if(!this.zipFiles) throw new Error('no file found');

        if(this.sharedStrings) {
            console.log('Updating sharedString.xml...');
            this.zipFiles.file("xl/sharedStrings.xml", this.sharedStrings.toString());
        }  

        console.log('Updating styles.xml...');
        this.zipFiles.file("xl/styles.xml", this.styles.toString());

        console.log('Updating worksheets.xml...');
        for (const [sheetNo, worksheet] of this.worksheets) {
            this.zipFiles?.file(`xl/worksheets/sheet${sheetNo}.xml`, worksheet.toString());
        }

        const arrayBuffer = await this.zipFiles.generateAsync({type:'arraybuffer', compression: "DEFLATE", compressionOptions: {level: 9}});

        return arrayBuffer;
    }
/*
    public getSharedString () {
        return JSON.stringify(this.sharedStrings.sharedStringArray);
    }*/
}