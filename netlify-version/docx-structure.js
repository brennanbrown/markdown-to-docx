/**
 * DOCX Structure Generator
 * Contains all the XML templates needed to create valid DOCX files
 */

class DocxStructure {
    static getContentTypesXml() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>`;
    }

    static getRelsXml() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
    }

    static getDocumentRelsXml() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>`;
    }

    static getStylesXml() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:style w:type="paragraph" w:styleId="Normal">
        <w:name w:val="Normal"/>
        <w:qFormat/>
        <w:pPr>
            <w:spacing w:after="120"/>
        </w:pPr>
        <w:rPr>
            <w:sz w:val="22"/>
        </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading1">
        <w:name w:val="heading 1"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr>
            <w:spacing w:before="240" w:after="120"/>
        </w:pPr>
        <w:rPr>
            <w:b/>
            <w:sz w:val="32"/>
            <w:color w:val="2F5597"/>
        </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading2">
        <w:name w:val="heading 2"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr>
            <w:spacing w:before="200" w:after="120"/>
        </w:pPr>
        <w:rPr>
            <w:b/>
            <w:sz w:val="26"/>
            <w:color w:val="2F5597"/>
        </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading3">
        <w:name w:val="heading 3"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr>
            <w:spacing w:before="160" w:after="120"/>
        </w:pPr>
        <w:rPr>
            <w:b/>
            <w:sz w:val="24"/>
            <w:color w:val="2F5597"/>
        </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading4">
        <w:name w:val="heading 4"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr>
            <w:spacing w:before="140" w:after="120"/>
        </w:pPr>
        <w:rPr>
            <w:b/>
            <w:sz w:val="22"/>
            <w:color w:val="2F5597"/>
        </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading5">
        <w:name w:val="heading 5"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr>
            <w:spacing w:before="120" w:after="120"/>
        </w:pPr>
        <w:rPr>
            <w:b/>
            <w:sz w:val="22"/>
            <w:color w:val="2F5597"/>
        </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Heading6">
        <w:name w:val="heading 6"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr>
            <w:spacing w:before="120" w:after="120"/>
        </w:pPr>
        <w:rPr>
            <w:b/>
            <w:sz w:val="22"/>
            <w:color w:val="2F5597"/>
        </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Quote">
        <w:name w:val="Quote"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr>
            <w:spacing w:before="120" w:after="120"/>
            <w:ind w:left="720"/>
        </w:pPr>
        <w:rPr>
            <w:i/>
            <w:color w:val="666666"/>
        </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="Code">
        <w:name w:val="Code"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr>
            <w:spacing w:before="120" w:after="120"/>
            <w:ind w:left="360"/>
        </w:pPr>
        <w:rPr>
            <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>
            <w:sz w:val="18"/>
            <w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>
        </w:rPr>
    </w:style>
</w:styles>`;
    }

    static getNumberingXml() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:abstractNum w:abstractNumId="0">
        <w:lvl w:ilvl="0">
            <w:start w:val="1"/>
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1."/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="720" w:hanging="360"/>
            </w:pPr>
        </w:lvl>
    </w:abstractNum>
    <w:abstractNum w:abstractNumId="1">
        <w:lvl w:ilvl="0">
            <w:start w:val="1"/>
            <w:numFmt w:val="bullet"/>
            <w:lvlText w:val="•"/>
            <w:lvlJc w:val="left"/>
            <w:pPr>
                <w:ind w:left="720" w:hanging="360"/>
            </w:pPr>
            <w:rPr>
                <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/>
            </w:rPr>
        </w:lvl>
    </w:abstractNum>
    <w:num w:numId="1">
        <w:abstractNumId w:val="0"/>
    </w:num>
    <w:num w:numId="2">
        <w:abstractNumId w:val="1"/>
    </w:num>
</w:numbering>`;
    }

    static getDocumentXml(content) {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        ${content}
    </w:body>
</w:document>`;
    }
}
