import JSZip, { file } from 'jszip';
import ExcelStyles from './ExcelStyles';
import ExcelSharedStrings from './ExcelSharedStrings';
import ExcelWorksheet from './ExcelWorksheet';
import ExcelColumnConverter from './ExcelColumnConverter';


export default class ExcelDocument {
    private files = new Map<string, string>;
    private sharedStrings = new ExcelSharedStrings();
    private styles = new ExcelStyles();

    public async fromArrayBuffer (arrayBuffer: ArrayBuffer) {
        const zip = new JSZip();
        await zip.loadAsync(arrayBuffer);

        const zipFolder = zip.folder("");
        if(zipFolder) {
            for(const [relativePath, file] of Object.entries(zipFolder.files)) {
                const content = await file.async('binarystring');
                this.files.set(relativePath, content);
            }
        }

        console.log('Parsing sharedStrings.xml...');
        const xmlSharedStrings = this.files.get('xl/sharedStrings.xml') ?? '';
        if(xmlSharedStrings === '') throw new Error('sharedStrings is null or undefined');
        this.sharedStrings.fromXML(xmlSharedStrings);

        console.log('Parsing styles.xml...');
        const xmlStyles = this.files.get("xl/styles.xml") ?? '';
        if(xmlStyles === '') throw new Error('styles is null or undefined');
        this.styles.fromXML(xmlStyles);
    }

    public getRange (sheetNo: number, startCell: string, endCell: string): string[][] {
        console.log(`Parsing sheet${sheetNo}.xml...`);
        const xmlWorksheet = this.files.get(`xl/worksheets/sheet${sheetNo}.xml`) ?? "";
        if(xmlWorksheet === '') throw new Error('worksheet is null or undefined');
        const worksheet = new ExcelWorksheet();
        worksheet.fromXML(xmlWorksheet);

        return worksheet.getRange(startCell, endCell);

    }
    
}