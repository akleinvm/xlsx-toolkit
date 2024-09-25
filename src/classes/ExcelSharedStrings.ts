export class ExcelSharedStrings {
    private xmlDocument!: XMLDocument;
    private sstElement!: Element;
    private namespace!: string;
    private stringsMap!: Map<string, number>;

    public fromXML(xmlString: string) {
        this.xmlDocument = new DOMParser().parseFromString(xmlString, 'text/xml');
        this.sstElement = this.xmlDocument.getElementsByTagName('sst')[0];
        this.namespace = this.sstElement.getAttribute('xmlns') ?? '';
        this.stringsMap = new Map();
        
        const siElements = this.sstElement.getElementsByTagName('si');
        for(let i = 0; i < siElements.length; i++) {
            const tElement = siElements[i].getElementsByTagName('t')[0];
            if(tElement && tElement.textContent) {
                this.stringsMap.set(tElement.textContent, i)
            }
        }
    }

    public toString(): string {
        this.sstElement.setAttribute('uniqueCount', this.stringsMap.size.toString());

        const xmlSerializer = new XMLSerializer();
        return xmlSerializer.serializeToString(this.xmlDocument);
    }

    public getStringIndex(string: string): number {
        let stringIndex = this.stringsMap.get(string);

        if(!stringIndex) {
            stringIndex = this.stringsMap.size;
            this.stringsMap.set(string, this.stringsMap.size);
            const siElement = this.xmlDocument.createElementNS(this.namespace, 'si',);
            const tElement = this.xmlDocument.createElementNS(this.namespace, 't');
            tElement.textContent = string;
            siElement.appendChild(tElement);
            this.sstElement.appendChild(siElement);
        }
        return stringIndex
    }
}