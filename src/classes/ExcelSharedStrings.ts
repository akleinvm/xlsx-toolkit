import jsDOM from './SharedParserSerializer';

export default class ExcelSharedStrings {
    private static xmlDocument: XMLDocument;
    private static sstElement: Element;
    private static namespace: string;
    private static stringsMap: Map<string, number>;
    private static numbersMap: Map<number, string>;

    public fromXML(xmlString: string) {
        ExcelSharedStrings.xmlDocument = jsDOM.parser.parseFromString(xmlString, 'text/xml');
        ExcelSharedStrings.sstElement = ExcelSharedStrings.xmlDocument.getElementsByTagName('sst')[0];
        ExcelSharedStrings.namespace = ExcelSharedStrings.sstElement.getAttribute('xmlns') ?? '';
        ExcelSharedStrings.stringsMap = new Map();
        ExcelSharedStrings.numbersMap = new Map();
        
        const siElements = ExcelSharedStrings.sstElement.getElementsByTagName('si');
        for(let i = 0; i < siElements.length; i++) {
            const tElement = siElements[i].getElementsByTagName('t')[0];
            if(tElement && tElement.textContent) {
                ExcelSharedStrings.stringsMap.set(tElement.textContent, i);
                ExcelSharedStrings.numbersMap.set(i, tElement.textContent);
            }
        }
    }

    public toString(): string {
        ExcelSharedStrings.sstElement.setAttribute('uniqueCount', ExcelSharedStrings.stringsMap.size.toString());

        return jsDOM.serializer.serializeToString(ExcelSharedStrings.xmlDocument);
    }

    public static getStringIndex(string: string): number {
        let stringIndex = ExcelSharedStrings.stringsMap.get(string);

        if(!stringIndex) {
            stringIndex = ExcelSharedStrings.stringsMap.size;
            ExcelSharedStrings.stringsMap.set(string, ExcelSharedStrings.stringsMap.size);
            ExcelSharedStrings.numbersMap.set(ExcelSharedStrings.stringsMap.size, string);
            const siElement = this.xmlDocument.createElementNS(this.namespace, 'si',);
            const tElement = this.xmlDocument.createElementNS(this.namespace, 't');
            tElement.textContent = string;
            siElement.appendChild(tElement);
            this.sstElement.appendChild(siElement);
        }
        return stringIndex
    }

    public static getIndexString (index: number): string {
        return ExcelSharedStrings.numbersMap.get(index) ?? '';
    }
}