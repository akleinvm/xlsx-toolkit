export default class ExcelSharedStrings {
    private static xmlDocument: XMLDocument;
    private static sstElement: Element;
    private static namespace: string;
    private static stringsMap: Map<string, number>;

    public fromXML(xmlString: string) {
        ExcelSharedStrings.xmlDocument = new DOMParser().parseFromString(xmlString, 'text/xml');
        ExcelSharedStrings.sstElement = ExcelSharedStrings.xmlDocument.getElementsByTagName('sst')[0];
        ExcelSharedStrings.namespace = ExcelSharedStrings.sstElement.getAttribute('xmlns') ?? '';
        ExcelSharedStrings.stringsMap = new Map();
        
        const siElements = ExcelSharedStrings.sstElement.getElementsByTagName('si');
        for(let i = 0; i < siElements.length; i++) {
            const tElement = siElements[i].getElementsByTagName('t')[0];
            if(tElement && tElement.textContent) {
                ExcelSharedStrings.stringsMap.set(tElement.textContent, i)
            }
        }
    }

    public toString(): string {
        ExcelSharedStrings.sstElement.setAttribute('uniqueCount', ExcelSharedStrings.stringsMap.size.toString());

        const xmlSerializer = new XMLSerializer();
        return xmlSerializer.serializeToString(ExcelSharedStrings.xmlDocument);
    }

    public static getStringIndex(string: string): number {
        let stringIndex = ExcelSharedStrings.stringsMap.get(string);

        if(!stringIndex) {
            stringIndex = ExcelSharedStrings.stringsMap.size;
            ExcelSharedStrings.stringsMap.set(string, ExcelSharedStrings.stringsMap.size);
            const siElement = this.xmlDocument.createElementNS(this.namespace, 'si',);
            const tElement = this.xmlDocument.createElementNS(this.namespace, 't');
            tElement.textContent = string;
            siElement.appendChild(tElement);
            this.sstElement.appendChild(siElement);
        }
        return stringIndex
    }
}