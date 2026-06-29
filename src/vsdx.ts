import JSZip from 'jszip';
import { create } from 'xmlbuilder2';
import { GraphData } from './parser.js';

export class VsdxGenerator {
    private zip: JSZip;
    private pageHeight: number = 11; // inches
    private pageWidth: number = 8.5; // inches
    private dpi: number = 96;

    constructor() {
        this.zip = new JSZip();
    }

    public async generate(graph: GraphData): Promise<Buffer> {
        // Adjust page size if graph is too big
        const graphWidthIn = graph.width / this.dpi;
        const graphHeightIn = graph.height / this.dpi;
        
        if (graphWidthIn > this.pageWidth) this.pageWidth = graphWidthIn + 1;
        if (graphHeightIn > this.pageHeight) this.pageHeight = graphHeightIn + 1;

        this.addContentTypes();
        this.addRels();
        this.addDocProps();
        this.addDocumentXml();
        this.addThemeXml();
        this.addWindowsXml();
        this.addPagesXml(graph);
        this.addPageXml(graph);

        return await this.zip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 },
        });
    }

    private addContentTypes() {
        const xml = create({ encoding: 'UTF-8', standalone: true })
            .ele('Types', { xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types' })
                .ele('Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }).up()
                .ele('Default', { Extension: 'xml', ContentType: 'application/xml' }).up()
                .ele('Override', { PartName: '/visio/document.xml', ContentType: 'application/vnd.ms-visio.drawing.main+xml' }).up()
                .ele('Override', { PartName: '/visio/windows.xml', ContentType: 'application/vnd.ms-visio.windows+xml' }).up()
                .ele('Override', { PartName: '/visio/pages/pages.xml', ContentType: 'application/vnd.ms-visio.pages+xml' }).up()
                .ele('Override', { PartName: '/visio/pages/page1.xml', ContentType: 'application/vnd.ms-visio.page+xml' }).up()
                .ele('Override', { PartName: '/visio/theme/theme1.xml', ContentType: 'application/vnd.openxmlformats-officedocument.theme+xml' }).up()
                .ele('Override', { PartName: '/docProps/core.xml', ContentType: 'application/vnd.openxmlformats-package.core-properties+xml' }).up()
                .ele('Override', { PartName: '/docProps/app.xml', ContentType: 'application/vnd.openxmlformats-officedocument.extended-properties+xml' }).up()
                .ele('Override', { PartName: '/docProps/custom.xml', ContentType: 'application/vnd.openxmlformats-officedocument.custom-properties+xml' }).up()
            .up();
        this.zip.file('[Content_Types].xml', xml.end({ prettyPrint: true }));
    }

    private addRels() {
        const xml = create({ encoding: 'UTF-8', standalone: true })
            .ele('Relationships', { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' })
                .ele('Relationship', { Id: 'rId1', Type: 'http://schemas.microsoft.com/visio/2010/relationships/document', Target: 'visio/document.xml' }).up()
                .ele('Relationship', { Id: 'rId2', Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties', Target: 'docProps/app.xml' }).up()
                .ele('Relationship', { Id: 'rId3', Type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', Target: 'docProps/core.xml' }).up()
                .ele('Relationship', { Id: 'rId4', Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties', Target: 'docProps/custom.xml' }).up()
            .up();
        this.zip.folder('_rels')?.file('.rels', xml.end({ prettyPrint: true }));
    }

    private addDocProps() {
        const app = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Template></Template><Application>Microsoft Visio</Application><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Pages</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Page-1</vt:lpstr></vt:vector></TitlesOfParts><Manager></Manager><Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinkBase></HyperlinkBase><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>`;

        const core = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title></dc:title><dc:subject></dc:subject><dc:creator>mermaid2visio</dc:creator><cp:keywords></cp:keywords><dc:description></dc:description><cp:lastModifiedBy></cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:modified><cp:category></cp:category><dc:language>en-US</dc:language></cp:coreProperties>`;

        const custom = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="_VPID_ALTERNATENAMES"><vt:lpwstr></vt:lpwstr></property><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="IsMetric"><vt:bool>false</vt:bool></property><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="BuildNumberCreated"><vt:i4>0</vt:i4></property><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="BuildNumberEdited"><vt:i4>1074806719</vt:i4></property></Properties>`;

        this.zip.folder('docProps')?.file('app.xml', app);
        this.zip.folder('docProps')?.file('core.xml', core);
        this.zip.folder('docProps')?.file('custom.xml', custom);
    }

    private addThemeXml() {
        // Minimal Office theme — Visio expects a theme part referenced from document.xml.rels
        const theme = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office Theme"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4672C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office Theme"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office Theme"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="sq" cmpd="sng"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="9525" cap="sq" cmpd="sng"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="9525" cap="sq" cmpd="sng"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;

        this.zip.folder('visio')?.folder('theme')?.file('theme1.xml', theme);
    }

    private addDocumentXml() {
        const xml = create({ encoding: 'UTF-8', standalone: true })
            .ele('VisioDocument', {
                xmlns: 'http://schemas.microsoft.com/office/visio/2012/main',
                'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            })
                .ele('DocumentSettings', { TopPage: '0', DefaultTextStyle: '0', DefaultLineStyle: '0', DefaultFillStyle: '0', DefaultGuideStyle: '0' })
                    .ele('GlueSettings').txt('9').up()
                    .ele('SnapSettings').txt('65847').up()
                    .ele('SnapExtensions').txt('34').up()
                    .ele('SnapAngles').up()
                    .ele('DynamicGridEnabled').txt('1').up()
                    .ele('ProtectStyles').txt('0').up()
                    .ele('ProtectShapes').txt('0').up()
                    .ele('ProtectMasters').txt('0').up()
                    .ele('ProtectBkgnds').txt('0').up()
                .up()
                .ele('Colors')
                    .ele('ColorEntry', { IX: '0', RGB: '#000000' }).up()
                    .ele('ColorEntry', { IX: '1', RGB: '#FFFFFF' }).up()
                .up()
                .ele('FaceNames')
                    .ele('FaceName', { NameU: 'Calibri', UnicodeRanges: '-469750017 -1040178053 9 0', CharSets: '536871423 0', Panose: '2 15 5 2 2 2 4 3 2 4', Flags: '357' }).up()
                .up()
                .ele('StyleSheets')
                    .ele('StyleSheet', { ID: '0', NameU: 'No Style', IsCustomNameU: '1', Name: 'No Style', IsCustomName: '1' })
                        .ele('Cell', { N: 'EnableLineProps', V: '1' }).up()
                        .ele('Cell', { N: 'EnableFillProps', V: '1' }).up()
                        .ele('Cell', { N: 'EnableTextProps', V: '1' }).up()
                        .ele('Cell', { N: 'LineWeight', V: '0.01041666666666667' }).up()
                        .ele('Cell', { N: 'LineColor', V: '0' }).up()
                        .ele('Cell', { N: 'LinePattern', V: '1' }).up()
                        .ele('Cell', { N: 'Rounding', V: '0' }).up()
                        .ele('Cell', { N: 'EndArrowSize', V: '2' }).up()
                        .ele('Cell', { N: 'BeginArrow', V: '0' }).up()
                        .ele('Cell', { N: 'EndArrow', V: '0' }).up()
                        .ele('Cell', { N: 'LineCap', V: '0' }).up()
                        .ele('Cell', { N: 'BeginArrowSize', V: '2' }).up()
                        .ele('Cell', { N: 'LineColorTrans', V: '0' }).up()
                        .ele('Cell', { N: 'FillForegnd', V: '1' }).up()
                        .ele('Cell', { N: 'FillBkgnd', V: '0' }).up()
                        .ele('Cell', { N: 'FillPattern', V: '1' }).up()
                        .ele('Cell', { N: 'ShdwForegnd', V: '0' }).up()
                        .ele('Cell', { N: 'ShdwPattern', V: '0' }).up()
                        .ele('Cell', { N: 'FillForegndTrans', V: '0' }).up()
                        .ele('Cell', { N: 'FillBkgndTrans', V: '0' }).up()
                        .ele('Cell', { N: 'ShdwForegndTrans', V: '0' }).up()
                        .ele('Cell', { N: 'ShapeShdwType', V: '0' }).up()
                        .ele('Cell', { N: 'ShapeShdwOffsetX', V: '0' }).up()
                        .ele('Cell', { N: 'ShapeShdwOffsetY', V: '0' }).up()
                        .ele('Cell', { N: 'ShapeShdwScaleFactor', V: '1' }).up()
                        .ele('Cell', { N: 'ShapeShdwBlur', V: '0' }).up()
                        .ele('Cell', { N: 'ShapeShdwShow', V: '0' }).up()
                        .ele('Cell', { N: 'LeftMargin', V: '0' }).up()
                        .ele('Cell', { N: 'RightMargin', V: '0' }).up()
                        .ele('Cell', { N: 'TopMargin', V: '0' }).up()
                        .ele('Cell', { N: 'BottomMargin', V: '0' }).up()
                        .ele('Cell', { N: 'VerticalAlign', V: '1' }).up()
                        .ele('Cell', { N: 'TextBkgnd', V: '0' }).up()
                        .ele('Cell', { N: 'DefaultTabStop', V: '0.5' }).up()
                        .ele('Cell', { N: 'TextDirection', V: '0' }).up()
                        .ele('Cell', { N: 'TextBkgndTrans', V: '0' }).up()
                        .ele('Cell', { N: 'LineGradientEnabled', V: '0' }).up()
                        .ele('Cell', { N: 'FillGradientEnabled', V: '0' }).up()
                        .ele('Cell', { N: 'QuickStyleLineColor', V: '100' }).up()
                        .ele('Cell', { N: 'QuickStyleFillColor', V: '100' }).up()
                        .ele('Cell', { N: 'QuickStyleShadowColor', V: '100' }).up()
                        .ele('Cell', { N: 'QuickStyleFontColor', V: '100' }).up()
                        .ele('Cell', { N: 'QuickStyleLineMatrix', V: '100' }).up()
                        .ele('Cell', { N: 'QuickStyleFillMatrix', V: '100' }).up()
                        .ele('Cell', { N: 'QuickStyleEffectsMatrix', V: '100' }).up()
                        .ele('Cell', { N: 'QuickStyleFontMatrix', V: '100' }).up()
                        .ele('Cell', { N: 'QuickStyleType', V: '0' }).up()
                        .ele('Cell', { N: 'QuickStyleVariation', V: '0' }).up()
                        .ele('Section', { N: 'Character' })
                            .ele('Row', { IX: '0' })
                                .ele('Cell', { N: 'Font', V: 'Calibri' }).up()
                                .ele('Cell', { N: 'Color', V: '0' }).up()
                                .ele('Cell', { N: 'Style', V: '0' }).up()
                                .ele('Cell', { N: 'Size', V: '0.1666666666666667' }).up()
                                .ele('Cell', { N: 'LangID', V: 'en-US' }).up()
                            .up()
                        .up()
                        .ele('Section', { N: 'Paragraph' })
                            .ele('Row', { IX: '0' })
                                .ele('Cell', { N: 'IndFirst', V: '0' }).up()
                                .ele('Cell', { N: 'IndLeft', V: '0' }).up()
                                .ele('Cell', { N: 'IndRight', V: '0' }).up()
                                .ele('Cell', { N: 'SpLine', V: '-1.2' }).up()
                                .ele('Cell', { N: 'SpBefore', V: '0' }).up()
                                .ele('Cell', { N: 'SpAfter', V: '0' }).up()
                                .ele('Cell', { N: 'HorzAlign', V: '1' }).up()
                                .ele('Cell', { N: 'Bullet', V: '0' }).up()
                            .up()
                        .up()
                    .up()
                .up()
                .ele('DocumentSheet', { NameU: 'TheDoc', IsCustomNameU: '1', Name: 'TheDoc', IsCustomName: '1', LineStyle: '0', FillStyle: '0', TextStyle: '0' })
                    .ele('Cell', { N: 'DocLangID', V: 'en-US' }).up()
                .up()
            .up();

        this.zip.folder('visio')?.file('document.xml', xml.end({ prettyPrint: true }));

        // Document Rels — points at the pages collection, windows state part, and theme
        const rels = create({ encoding: 'UTF-8', standalone: true })
            .ele('Relationships', { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' })
                .ele('Relationship', { Id: 'rId1', Type: 'http://schemas.microsoft.com/visio/2010/relationships/pages', Target: 'pages/pages.xml' }).up()
                .ele('Relationship', { Id: 'rId2', Type: 'http://schemas.microsoft.com/visio/2010/relationships/windows', Target: 'windows.xml' }).up()
                .ele('Relationship', { Id: 'rId3', Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme', Target: 'theme/theme1.xml' }).up()
            .up();
        this.zip.folder('visio')?.folder('_rels')?.file('document.xml.rels', rels.end({ prettyPrint: true }));
    }

    private addWindowsXml() {
        const xml = create({ encoding: 'UTF-8', standalone: true })
            .ele('Windows', {
                xmlns: 'http://schemas.microsoft.com/office/visio/2012/main',
                ClientWidth: '1280',
                ClientHeight: '720',
            })
                .ele('Window', {
                    ID: '0',
                    WindowType: 'Drawing',
                    WindowState: '1',
                    WindowLeft: '0',
                    WindowTop: '0',
                    WindowWidth: '1280',
                    WindowHeight: '720',
                    ContainerType: 'Page',
                    Page: '0',
                    ViewScale: '-1',
                    ViewCenterX: (this.pageWidth / 2).toString(),
                    ViewCenterY: (this.pageHeight / 2).toString(),
                })
                    .ele('ShowRulers').txt('1').up()
                    .ele('ShowGrid').txt('1').up()
                    .ele('ShowPageBreaks').txt('0').up()
                    .ele('ShowGuides').txt('1').up()
                    .ele('ShowConnectionPoints').txt('1').up()
                    .ele('GlueSettings').txt('9').up()
                    .ele('SnapSettings').txt('65847').up()
                    .ele('SnapExtensions').txt('34').up()
                    .ele('SnapAngles').up()
                    .ele('DynamicGridEnabled').txt('1').up()
                    .ele('TabSplitterPos').txt('0.5').up()
                .up()
            .up();

        this.zip.folder('visio')?.file('windows.xml', xml.end({ prettyPrint: true }));
    }

    private addPagesXml(graph: GraphData) {
        const xml = create({ encoding: 'UTF-8', standalone: true })
            .ele('Pages', {
                xmlns: 'http://schemas.microsoft.com/office/visio/2012/main',
                'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            })
                .ele('Page', { ID: '0', NameU: 'Page-1', Name: 'Page-1', ViewScale: '-1', ViewCenterX: (this.pageWidth / 2).toString(), ViewCenterY: (this.pageHeight / 2).toString() })
                    .ele('PageSheet', { LineStyle: '0', FillStyle: '0', TextStyle: '0' })
                        .ele('Cell', { N: 'PageWidth', V: this.pageWidth.toString() }).up()
                        .ele('Cell', { N: 'PageHeight', V: this.pageHeight.toString() }).up()
                        .ele('Cell', { N: 'ShdwOffsetX', V: '0.125' }).up()
                        .ele('Cell', { N: 'ShdwOffsetY', V: '-0.125' }).up()
                        .ele('Cell', { N: 'PageScale', V: '1', U: 'IN_F' }).up()
                        .ele('Cell', { N: 'DrawingScale', V: '1', U: 'IN_F' }).up()
                        .ele('Cell', { N: 'DrawingSizeType', V: '0' }).up()
                        .ele('Cell', { N: 'DrawingScaleType', V: '0' }).up()
                        .ele('Cell', { N: 'InhibitSnap', V: '0' }).up()
                        .ele('Cell', { N: 'PageLockReplace', V: '0', U: 'BOOL' }).up()
                        .ele('Cell', { N: 'PageLockDuplicate', V: '0', U: 'BOOL' }).up()
                        .ele('Cell', { N: 'UIVisibility', V: '0' }).up()
                        .ele('Cell', { N: 'ShdwType', V: '0' }).up()
                        .ele('Cell', { N: 'ShdwObliqueAngle', V: '0' }).up()
                        .ele('Cell', { N: 'ShdwScaleFactor', V: '1' }).up()
                    .up()
                    .ele('Rel', { 'r:id': 'rId1' }).up()
                .up()
            .up();

        this.zip.folder('visio')?.folder('pages')?.file('pages.xml', xml.end({ prettyPrint: true }));

        // pages.xml.rels → page1.xml
        const rels = create({ encoding: 'UTF-8', standalone: true })
            .ele('Relationships', { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' })
                .ele('Relationship', { Id: 'rId1', Type: 'http://schemas.microsoft.com/visio/2010/relationships/page', Target: 'page1.xml' }).up()
            .up();
        this.zip.folder('visio')?.folder('pages')?.folder('_rels')?.file('pages.xml.rels', rels.end({ prettyPrint: true }));

        // page1.xml.rels — empty Relationships container (LibreOffice/Visio expect this part to exist)
        const pageRels = create({ encoding: 'UTF-8', standalone: true })
            .ele('Relationships', { xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships' })
            .up();
        this.zip.folder('visio')?.folder('pages')?.folder('_rels')?.file('page1.xml.rels', pageRels.end({ prettyPrint: true }));
    }

    private addPageXml(graph: GraphData) {
        const root = create({ encoding: 'UTF-8', standalone: true })
            .ele('PageContents', { xmlns: 'http://schemas.microsoft.com/office/visio/2012/main' })
            .ele('Shapes');

        let shapeId = 1;
        const nodeIdToShapeId = new Map<string, number>();
        const connects: Array<{ fromSheet: number, fromCell: string, toSheet: number, toCell: string }> = [];

        // Helper for LinePattern
        const getLinePattern = (dash: string | undefined) => {
             if (!dash || dash === 'none' || dash === '0') return null;
             return '2'; 
        };

        // Helper for Fonts
        const getFontStyle = (weight?: string, style?: string) => {
            let s = 0;
            if (weight === 'bold' || weight === '700' || weight === '800') s |= 1; // Bold
            if (style === 'italic') s |= 2; // Italic
            return s;
        };

        const getFontSize = (sizeStr?: string) => {
            if (!sizeStr) return null;
            const px = parseFloat(sizeStr);
            if (isNaN(px)) return null;
            // VSDX (OPC) Size cell value is in INTERNAL units = inches, regardless of any U attribute.
            // 1 px = 1/96 inch at 96 DPI. (Equivalent to px*0.75pt then /72pt-per-inch.)
            return (px / this.dpi).toFixed(4);
        };

        const getHorzAlign = (align?: string) => {
            if (align === 'left') return '0';
            if (align === 'right') return '2';
            return '1'; // Center default
        };

        // Geometry/Connection cells require a NUMERIC V= with the formula in F=
        // expr is a Visio formula like 'Width*0.5' or 'Height*0.15' or 'Width-Height*0.5'
        // Evaluate it numerically given the shape's actual Width/Height in inches.
        const evalFormula = (expr: string | number, W: number, H: number): { V: string, F?: string } => {
            if (typeof expr === 'number') return { V: expr.toString() };
            const trimmed = String(expr).trim();
            // Plain number?
            if (/^-?\d+(\.\d+)?$/.test(trimmed)) return { V: trimmed };
            // Build numeric evaluation
            const replaced = trimmed.replace(/Width/g, `(${W})`).replace(/Height/g, `(${H})`);
            // Safe eval: only digits, ops, parens, spaces, decimal, dot
            if (!/^[-+*/().0-9\s]+$/.test(replaced)) {
                return { V: '0', F: trimmed };
            }
            let val = 0;
            try {
                // eslint-disable-next-line no-new-func
                val = Function(`"use strict"; return (${replaced});`)() as number;
                if (!isFinite(val)) val = 0;
            } catch {
                val = 0;
            }
            return { V: val.toString(), F: trimmed };
        };

        const addCellExpr = (parent: any, N: string, expr: string | number, W: number, H: number) => {
            const { V, F } = evalFormula(expr, W, H);
            const attrs: any = { N, V };
            if (F) attrs.F = F;
            parent.ele('Cell', attrs).up();
        };

        // Visio rejects CSS `rgb(...)` and named colors. Coerce to #RRGGBB.
        const toHexColor = (c?: string): string | undefined => {
            if (!c) return undefined;
            const s = c.trim();
            if (!s || s === 'none' || s === 'transparent') return undefined;
            if (s.startsWith('#')) {
                if (s.length === 4) {
                    return '#' + s[1] + s[1] + s[2] + s[2] + s[3] + s[3];
                }
                return s.toUpperCase();
            }
            const m = s.match(/^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
            if (m) {
                const r = Math.min(255, parseInt(m[1], 10));
                const g = Math.min(255, parseInt(m[2], 10));
                const b = Math.min(255, parseInt(m[3], 10));
                return '#' + [r, g, b].map(v => v.toString(16).padStart(2, '0').toUpperCase()).join('');
            }
            return undefined; // drop named colors we can't translate
        };

        // 1. Add Clusters (Subgraphs) - Draw first (background)
        if (graph.clusters) {
            for (const cluster of graph.clusters) {
                const w = cluster.width / this.dpi;
                const h = cluster.height / this.dpi;
                const x = cluster.x / this.dpi; 
                const y = cluster.y / this.dpi;
                
                const pinX = x + w/2;
                const pinY = this.pageHeight - (y + h/2);

                const shape = root.ele('Shape', { ID: shapeId.toString(), Type: 'Group', LineStyle: '0', FillStyle: '0', TextStyle: '0' });
                
                shape.ele('Cell', { N: 'PinX', V: pinX.toString() }).up();
                shape.ele('Cell', { N: 'PinY', V: pinY.toString() }).up();
                shape.ele('Cell', { N: 'Width', V: w.toString() }).up();
                shape.ele('Cell', { N: 'Height', V: h.toString() }).up();
                shape.ele('Cell', { N: 'LocPinX', V: (w / 2).toString(), F: 'Width*0.5' }).up();
                shape.ele('Cell', { N: 'LocPinY', V: (h / 2).toString(), F: 'Height*0.5' }).up();

                // Style — Visio defaults FillPattern/LinePattern to 0 (invisible) without explicit Cells.
                const cFill = toHexColor(cluster.style?.fill);
                const cStroke = toHexColor(cluster.style?.stroke);
                shape.ele('Cell', { N: 'FillForegnd', V: cFill || '#FFFFFF' }).up();
                shape.ele('Cell', { N: 'FillPattern', V: cFill ? '1' : '0' }).up();
                shape.ele('Cell', { N: 'LineColor', V: cStroke || '#000000' }).up();
                if (cluster.style?.strokeWidth) {
                    const px = parseFloat(cluster.style.strokeWidth) || 1;
                    shape.ele('Cell', { N: 'LineWeight', V: (px * 0.01).toString() }).up();
                }
                const cLp = getLinePattern(cluster.style?.strokeDasharray);
                shape.ele('Cell', { N: 'LinePattern', V: cLp || '1' }).up();
                // Top-align cluster title; give it left margin so it doesn't sit flush against the edge
                shape.ele('Cell', { N: 'VerticalAlign', V: '0' }).up();
                shape.ele('Cell', { N: 'LeftMargin', V: '0.08' }).up();
                shape.ele('Cell', { N: 'TopMargin', V: '0.04' }).up();

                // Geometry - Rectangle
                const geom = shape.ele('Section', { N: 'Geometry', IX: '0' });
                geom.ele('Cell', { N: 'NoFill', V: '0' }).up();
                geom.ele('Cell', { N: 'NoLine', V: '0' }).up();
                geom.ele('Cell', { N: 'NoShow', V: '0' }).up();
                geom.ele('Cell', { N: 'NoSnap', V: '0' }).up();
                {
                    const r1 = geom.ele('Row', { T: 'MoveTo', IX: '1' });
                    addCellExpr(r1, 'X', 0, w, h); addCellExpr(r1, 'Y', 0, w, h); r1.up();
                    const r2 = geom.ele('Row', { T: 'LineTo', IX: '2' });
                    addCellExpr(r2, 'X', 'Width', w, h); addCellExpr(r2, 'Y', 0, w, h); r2.up();
                    const r3 = geom.ele('Row', { T: 'LineTo', IX: '3' });
                    addCellExpr(r3, 'X', 'Width', w, h); addCellExpr(r3, 'Y', 'Height', w, h); r3.up();
                    const r4 = geom.ele('Row', { T: 'LineTo', IX: '4' });
                    addCellExpr(r4, 'X', 0, w, h); addCellExpr(r4, 'Y', 'Height', w, h); r4.up();
                    const r5 = geom.ele('Row', { T: 'LineTo', IX: '5' });
                    addCellExpr(r5, 'X', 0, w, h); addCellExpr(r5, 'Y', 0, w, h); r5.up();
                }
                geom.up();

                // Bold the cluster title via Character section
                if (cluster.text) {
                    const cs = shape.ele('Section', { N: 'Character', IX: '0' });
                    const cr = cs.ele('Row', { IX: '0' });
                    cr.ele('Cell', { N: 'Style', V: '1' }).up(); // Bold
                    const tColor = toHexColor(cluster.style?.color);
                    if (tColor) cr.ele('Cell', { N: 'Color', V: tColor }).up();
                    cr.up().up();
                }

                // Paragraph: left-align cluster title
                const pg = shape.ele('Section', { N: 'Paragraph', IX: '0' });
                pg.ele('Row', { IX: '0' })
                    .ele('Cell', { N: 'HorzAlign', V: '0' }).up()
                .up().up();

                // Text last
                if (cluster.text) {
                    shape.ele('Text').txt(cluster.text).up();
                }

                shapeId++;
            }
        }

        // 2. Add Nodes
        for (const node of graph.nodes) {
            // Register ID
            nodeIdToShapeId.set(node.id, shapeId);

            // Convert to inches
            const w = node.width / this.dpi;
            const h = node.height / this.dpi;
            const x = node.x / this.dpi; 
            const y = this.pageHeight - (node.y / this.dpi);

            const shape = root.ele('Shape', { ID: shapeId.toString(), Type: 'Shape', LineStyle: '0', FillStyle: '0', TextStyle: '0' });
            
            // Transform
            shape.ele('Cell', { N: 'PinX', V: x.toString() }).up();
            shape.ele('Cell', { N: 'PinY', V: y.toString() }).up();
            shape.ele('Cell', { N: 'Width', V: w.toString() }).up();
            shape.ele('Cell', { N: 'Height', V: h.toString() }).up();
            shape.ele('Cell', { N: 'LocPinX', V: (w / 2).toString(), F: 'Width*0.5' }).up();
            shape.ele('Cell', { N: 'LocPinY', V: (h / 2).toString(), F: 'Height*0.5' }).up();

            // Rounding
            if (node.rounding && node.rounding > 0) {
                // Convert px to inches
                const rIn = node.rounding / this.dpi;
                shape.ele('Cell', { N: 'Rounding', V: rIn.toString() }).up();
            }

            // Styles — Visio defaults FillPattern/LinePattern to 0 (invisible) without explicit Cells.
            const nFill = toHexColor(node.style?.fill);
            const nStroke = toHexColor(node.style?.stroke);
            shape.ele('Cell', { N: 'FillForegnd', V: nFill || '#FFFFFF' }).up();
            shape.ele('Cell', { N: 'FillPattern', V: nFill ? '1' : '0' }).up();
            shape.ele('Cell', { N: 'LineColor', V: nStroke || '#000000' }).up();
            if (node.style?.strokeWidth) {
                const px = parseFloat(node.style.strokeWidth) || 1;
                shape.ele('Cell', { N: 'LineWeight', V: (px * 0.01).toString() }).up();
            }
            const nLp = getLinePattern(node.style?.strokeDasharray);
            shape.ele('Cell', { N: 'LinePattern', V: nLp || '1' }).up();

            // Tight text margins so multi-line labels fit inside the shape
            shape.ele('Cell', { N: 'LeftMargin', V: '0.02' }).up();
            shape.ele('Cell', { N: 'RightMargin', V: '0.02' }).up();
            shape.ele('Cell', { N: 'TopMargin', V: '0.02' }).up();
            shape.ele('Cell', { N: 'BottomMargin', V: '0.02' }).up();

            // Paragraph (Alignment)
            if (node.style && node.style.textAlign) {
                shape.ele('Section', { N: 'Paragraph', IX: '0' })
                    .ele('Row', { IX: '0' })
                        .ele('Cell', { N: 'HorzAlign', V: getHorzAlign(node.style.textAlign) }).up()
                    .up().up();
            }

            // Hyperlink
            if (node.url) {
                shape.ele('Section', { N: 'Hyperlink', IX: '0' })
                    .ele('Row', { IX: '0' })
                        .ele('Cell', { N: 'Address', V: node.url }).up()
                    .up().up();
            }

            // Geometry
            const geom = shape.ele('Section', { N: 'Geometry', IX: '0' });
            geom.ele('Cell', { N: 'NoFill', V: '0' }).up();
            geom.ele('Cell', { N: 'NoLine', V: '0' }).up();
            geom.ele('Cell', { N: 'NoShow', V: '0' }).up();
            geom.ele('Cell', { N: 'NoSnap', V: '0' }).up();

            const addRow = (sec: any, T: string, ix: number, cells: Array<[string, string | number]>) => {
                const row = sec.ele('Row', { T, IX: ix.toString() });
                for (const [N, expr] of cells) addCellExpr(row, N, expr, w, h);
                row.up();
            };
            if (node.type === 'circle' || node.type === 'ellipse') {
                addRow(geom, 'Ellipse', 1, [
                    ['X', 'Width*0.5'], ['Y', 'Height*0.5'],
                    ['A', 'Width'], ['B', 'Height*0.5'],
                    ['C', 'Width*0.5'], ['D', 'Height'],
                ]);
            } else if (node.type === 'diamond') {
                addRow(geom, 'MoveTo', 1, [['X', 'Width*0.5'], ['Y', 0]]);
                addRow(geom, 'LineTo', 2, [['X', 'Width'], ['Y', 'Height*0.5']]);
                addRow(geom, 'LineTo', 3, [['X', 'Width*0.5'], ['Y', 'Height']]);
                addRow(geom, 'LineTo', 4, [['X', 0], ['Y', 'Height*0.5']]);
                addRow(geom, 'LineTo', 5, [['X', 'Width*0.5'], ['Y', 0]]);
            } else if (node.type === 'stadium') {
                addRow(geom, 'MoveTo', 1, [['X', 'Height*0.5'], ['Y', 0]]);
                addRow(geom, 'LineTo', 2, [['X', 'Width - Height*0.5'], ['Y', 0]]);
                addRow(geom, 'ArcTo', 3, [['X', 'Width - Height*0.5'], ['Y', 'Height'], ['A', 'Width']]);
                addRow(geom, 'LineTo', 4, [['X', 'Height*0.5'], ['Y', 'Height']]);
                addRow(geom, 'ArcTo', 5, [['X', 'Height*0.5'], ['Y', 0], ['A', 0]]);
            } else if (node.type === 'parallelogram') {
                addRow(geom, 'MoveTo', 1, [['X', 'Width*0.2'], ['Y', 0]]);
                addRow(geom, 'LineTo', 2, [['X', 'Width'], ['Y', 0]]);
                addRow(geom, 'LineTo', 3, [['X', 'Width*0.8'], ['Y', 'Height']]);
                addRow(geom, 'LineTo', 4, [['X', 0], ['Y', 'Height']]);
                addRow(geom, 'LineTo', 5, [['X', 'Width*0.2'], ['Y', 0]]);
            } else if (node.type === 'cylinder') {
                // Cylinder: single closed outline drawn as one filled path so the inside is
                // clean (no inner half-oval visible). The top is a half-arc (visible cap),
                // the sides are straight, and the bottom is a half-arc (visible curve).
                // Path: top-left → arc-over to top-right (the top cap) → down right side →
                // arc-under to bottom-left → up left side → close.
                addRow(geom, 'MoveTo', 1, [['X', 0], ['Y', 'Height*0.15']]);
                // Top half-arc: bows UPWARD to Y=0. For a horizontal chord and left→right
                // motion, "up" is right-of-motion → A < 0.
                addRow(geom, 'ArcTo', 2, [['X', 'Width'], ['Y', 'Height*0.15'], ['A', '-Height*0.15']]);
                addRow(geom, 'LineTo', 3, [['X', 'Width'], ['Y', 'Height*0.85']]);
                // Bottom half-arc: bows DOWNWARD to Y=Height. For right→left motion,
                // "down" is right-of-motion → A < 0 again. Both arcs use A = -Height*0.15.
                addRow(geom, 'ArcTo', 4, [['X', 0], ['Y', 'Height*0.85'], ['A', '-Height*0.15']]);
                addRow(geom, 'LineTo', 5, [['X', 0], ['Y', 'Height*0.15']]);

                // Second geometry: front curve of the top cap (the visible "lip" inside the
                // top of the cylinder). No fill, just the outline so the body fill shows
                // through underneath. This is the line that suggests 3D depth.
                const geom2 = shape.ele('Section', { N: 'Geometry', IX: '1' });
                geom2.ele('Cell', { N: 'NoFill', V: '1' }).up();
                geom2.ele('Cell', { N: 'NoLine', V: '0' }).up();
                geom2.ele('Cell', { N: 'NoShow', V: '0' }).up();
                addRow(geom2, 'MoveTo', 1, [['X', 0], ['Y', 'Height*0.15']]);
                // Same chord as the top, but bow DOWNWARD (inside the cylinder): A positive.
                addRow(geom2, 'ArcTo', 2, [['X', 'Width'], ['Y', 'Height*0.15'], ['A', 'Height*0.15']]);
            } else if (node.type === 'subroutine') {
                addRow(geom, 'MoveTo', 1, [['X', 0], ['Y', 0]]);
                addRow(geom, 'LineTo', 2, [['X', 'Width'], ['Y', 0]]);
                addRow(geom, 'LineTo', 3, [['X', 'Width'], ['Y', 'Height']]);
                addRow(geom, 'LineTo', 4, [['X', 0], ['Y', 'Height']]);
                addRow(geom, 'LineTo', 5, [['X', 0], ['Y', 0]]);

                const subGeom1 = shape.ele('Section', { N: 'Geometry', IX: '1' });
                subGeom1.ele('Cell', { N: 'NoFill', V: '1' }).up();
                addRow(subGeom1, 'MoveTo', 1, [['X', '0.1*Width'], ['Y', 0]]);
                addRow(subGeom1, 'LineTo', 2, [['X', '0.1*Width'], ['Y', 'Height']]);

                const subGeom2 = shape.ele('Section', { N: 'Geometry', IX: '2' });
                subGeom2.ele('Cell', { N: 'NoFill', V: '1' }).up();
                addRow(subGeom2, 'MoveTo', 1, [['X', '0.9*Width'], ['Y', 0]]);
                addRow(subGeom2, 'LineTo', 2, [['X', '0.9*Width'], ['Y', 'Height']]);
            } else {
                // Default: Rectangle
                addRow(geom, 'MoveTo', 1, [['X', 0], ['Y', 0]]);
                addRow(geom, 'LineTo', 2, [['X', 'Width'], ['Y', 0]]);
                addRow(geom, 'LineTo', 3, [['X', 'Width'], ['Y', 'Height']]);
                addRow(geom, 'LineTo', 4, [['X', 0], ['Y', 'Height']]);
                addRow(geom, 'LineTo', 5, [['X', 0], ['Y', 0]]);
            }
            geom.up();

            // Character Props (must come before Text per Visio schema: Cells, Sections, then Text)
            if (node.text && node.style) {
                const charSec = shape.ele('Section', { N: 'Character', IX: '0' });
                const charRow = charSec.ele('Row', { IX: '0' });

                const col = toHexColor(node.style.color);
                if (col) charRow.ele('Cell', { N: 'Color', V: col }).up();

                const fs = getFontSize(node.style.fontSize);
                if (fs) charRow.ele('Cell', { N: 'Size', V: fs }).up(); // Unit=Points

                const st = getFontStyle(node.style.fontWeight, node.style.fontStyle);
                if (st > 0) charRow.ele('Cell', { N: 'Style', V: st.toString() }).up();

                charRow.up().up(); // Row, Section
            }

            // Connection Points
            const conn = shape.ele('Section', { N: 'Connection', IX: '0' });
            const connRow = (ix: number, xExpr: string, yExpr: string) => {
                const r = conn.ele('Row', { IX: ix.toString() });
                addCellExpr(r, 'X', xExpr, w, h);
                addCellExpr(r, 'Y', yExpr, w, h);
                r.up();
            };
            connRow(0, 'Width*0.5', 'Height*0.5');
            connRow(1, 'Width*0.5', 'Height');
            connRow(2, 'Width*0.5', '0');
            connRow(3, '0', 'Height*0.5');
            connRow(4, 'Width', 'Height*0.5');
            conn.up();

            // Text (must be last child of Shape)
            if (node.text) {
                shape.ele('Text').txt(node.text).up();
            }

            shapeId++;
        }

        // 3. Add Connectors (Edges)
        if (graph.edges) {
            for (const edge of graph.edges) {
                if (!edge.d) continue;

                // Create a shape for the connector
                // In Visio, a connector is just a 1D shape (a line). 
                // However, since we are drawing a complex path, we treat it as a 2D shape with no fill, just stroke.
                // We will position the shape at (0,0) of the page and draw absolute coordinates.
                // This is the easiest way to preserve the exact path.
                
                const shape = root.ele('Shape', { ID: shapeId.toString(), Type: 'Shape', LineStyle: '0', FillStyle: '0', TextStyle: '0' });
                
                // Pin at (0,0) of the page, so local coordinates match page coordinates (mostly)
                // But wait, Visio local coords are relative to PinX/PinY.
                // If PinX=0, PinY=0 (bottom-left), then local (x,y) are offsets from there.
                shape.ele('Cell', { N: 'PinX', V: '0' }).up();
                shape.ele('Cell', { N: 'PinY', V: '0' }).up();
                shape.ele('Cell', { N: 'Width', V: this.pageWidth.toString() }).up();
                shape.ele('Cell', { N: 'Height', V: this.pageHeight.toString() }).up();
                shape.ele('Cell', { N: 'LocPinX', V: '0' }).up();
                shape.ele('Cell', { N: 'LocPinY', V: '0' }).up();

                // Styling
                shape.ele('Cell', { N: 'FillPattern', V: '0' }).up(); // No fill

                // Add Arrowhead (Default to EndArrow=13 for standard arrow)
                // TODO: Check edge style for "arrowhead: none" if needed.
                shape.ele('Cell', { N: 'EndArrow', V: '13' }).up();

                let edgeStrokeHex = toHexColor(edge.style?.stroke) || '#000000';
                let edgeWeight = '0.01';
                if (edge.style?.strokeWidth) {
                    const px = parseFloat(edge.style.strokeWidth) || 1;
                    edgeWeight = (px * 0.01).toString();
                }
                shape.ele('Cell', { N: 'LineColor', V: edgeStrokeHex }).up();
                shape.ele('Cell', { N: 'LineWeight', V: edgeWeight }).up();

                // ObjType must come before any Section per Visio schema (Cells precede Sections)
                shape.ele('Cell', { N: 'ObjType', V: '2' }).up();

                // 1D endpoints
                shape.ele('Cell', { N: 'BeginX', V: '0' }).up();
                shape.ele('Cell', { N: 'BeginY', V: '0' }).up();
                shape.ele('Cell', { N: 'EndX', V: '0' }).up();
                shape.ele('Cell', { N: 'EndY', V: '0' }).up();

                const geom = shape.ele('Section', { N: 'Geometry', IX: '0' });
                geom.ele('Cell', { N: 'NoFill', V: '1' }).up();
                geom.ele('Cell', { N: 'NoLine', V: '0' }).up();
                geom.ele('Cell', { N: 'NoShow', V: '0' }).up();

                this.parsePathToVisio(edge.d, geom);
                geom.up();

                // Routing Style (1 = Right Angle)
                shape.ele('Section', { N: 'ShapeLayout', IX: '0' })
                    .ele('Row', { IX: '0' })
                        .ele('Cell', { N: 'ConLineRouteExt', V: '1' }).up()
                    .up().up();

                // Logic for connection
                if (edge.startId && edge.endId) {
                    const startShapeId = nodeIdToShapeId.get(edge.startId);
                    const endShapeId = nodeIdToShapeId.get(edge.endId);

                    if (startShapeId && endShapeId) {
                        // Queue the connection
                        connects.push({ fromSheet: shapeId, fromCell: 'BeginX', toSheet: startShapeId, toCell: 'PinX' });
                        connects.push({ fromSheet: shapeId, fromCell: 'EndX', toSheet: endShapeId, toCell: 'PinX' });
                    }
                }

                shapeId++;
            }
        }

        // 4. Add Edge Labels
        if (graph.labels) {
            for (const label of graph.labels) {
                const w = (label.width || 10) / this.dpi;
                const h = (label.height || 10) / this.dpi;
                // Mermaid edge labels are placed via transform=translate(cx, cy), so (x, y) is
                // the CENTER of the label bbox. Same convention nodes use, so the pin maps 1:1.
                const pinX = label.x / this.dpi;
                const pinY = this.pageHeight - (label.y / this.dpi);

                const shape = root.ele('Shape', { ID: shapeId.toString(), Type: 'Shape', LineStyle: '0', FillStyle: '0', TextStyle: '0' });
                
                shape.ele('Cell', { N: 'PinX', V: pinX.toString() }).up();
                shape.ele('Cell', { N: 'PinY', V: pinY.toString() }).up();
                shape.ele('Cell', { N: 'Width', V: w.toString() }).up();
                shape.ele('Cell', { N: 'Height', V: h.toString() }).up();
                shape.ele('Cell', { N: 'LocPinX', V: (w / 2).toString(), F: 'Width*0.5' }).up();
                shape.ele('Cell', { N: 'LocPinY', V: (h / 2).toString(), F: 'Height*0.5' }).up();

                // Invisible box
                shape.ele('Cell', { N: 'FillPattern', V: '0' }).up();
                shape.ele('Cell', { N: 'LinePattern', V: '0' }).up();

                // Background fill if specified
                const labelFill = toHexColor(label.style?.fill);
                if (labelFill) {
                    shape.ele('Cell', { N: 'FillPattern', V: '1' }).up();
                    shape.ele('Cell', { N: 'FillForegnd', V: labelFill }).up();
                }

                // Char Props (must come before Text)
                const charSec = shape.ele('Section', { N: 'Character', IX: '0' });
                const charRow = charSec.ele('Row', { IX: '0' });

                if (label.style) {
                    const labelColor = toHexColor(label.style.color);
                    if (labelColor) charRow.ele('Cell', { N: 'Color', V: labelColor }).up();

                    const fs = getFontSize(label.style.fontSize);
                    if (fs) charRow.ele('Cell', { N: 'Size', V: fs }).up();

                    const st = getFontStyle(label.style.fontWeight, label.style.fontStyle);
                    if (st > 0) charRow.ele('Cell', { N: 'Style', V: st.toString() }).up();
                }
                charRow.up().up();

                // Text last
                shape.ele('Text').txt(label.text).up();

                shapeId++;
            }
        }
        
        root.up(); // Exit Shapes (Wait, I need to check xmlbuilder2 nesting. I think I need to go up from Shapes before Connects? NO. Connects is sibling of Shapes in PageContents)
        // Shapes was created with root.ele('Shapes').
        // So root is PageContents.
        // We are inside Shapes right now? No, loop finished.
        // root points to PageContents -> Shapes.
        root.up(); // Now root points to PageContents.

        // 5. Add Connections
        if (connects.length > 0) {
            const connXml = root.ele('Connects');
            for (const c of connects) {
                connXml.ele('Connect', { FromSheet: c.fromSheet.toString(), FromCell: c.fromCell, ToSheet: c.toSheet.toString(), ToCell: c.toCell }).up();
            }
            connXml.up();
        }

        // root is still PageContents
        this.zip.folder('visio')?.folder('pages')?.file('page1.xml', root.end({ prettyPrint: true }));
    }

    private parsePathToVisio(d: string, geomXml: any) {
        // Simple regex-based SVG path parser
        // Supports M (Move), L (Line), C (Cubic Bezier)
        // Coordinates in d are separate by space or comma
        
        // Normalize: Put spaces around commands
        const normalized = d.replace(/([a-zA-Z])/g, ' $1 ').trim();
        const tokens = normalized.split(/[\s,]+|(?=[a-zA-Z])/).filter(t => t.trim().length > 0);

        let currentX = 0;
        let currentY = 0;
        let rowIx = 1;

        const toVisioY = (y: number) => this.pageHeight - (y / this.dpi);
        const toVisioX = (x: number) => x / this.dpi;

        for (let i = 0; i < tokens.length; i++) {
            const cmd = tokens[i];
            
            if (cmd === 'M') {
                const x = parseFloat(tokens[++i]);
                const y = parseFloat(tokens[++i]);
                currentX = x;
                currentY = y;
                
                geomXml.ele('Row', { T: 'MoveTo', IX: rowIx.toString() })
                    .ele('Cell', { N: 'X', V: toVisioX(x).toString() }).up()
                    .ele('Cell', { N: 'Y', V: toVisioY(y).toString() }).up().up();
                rowIx++;
            } else if (cmd === 'L') {
                const x = parseFloat(tokens[++i]);
                const y = parseFloat(tokens[++i]);
                currentX = x;
                currentY = y;

                geomXml.ele('Row', { T: 'LineTo', IX: rowIx.toString() })
                    .ele('Cell', { N: 'X', V: toVisioX(x).toString() }).up()
                    .ele('Cell', { N: 'Y', V: toVisioY(y).toString() }).up().up();
                rowIx++;
            } else if (cmd === 'C') {
                // Cubic Bezier: C x1 y1, x2 y2, x y
                const x1 = parseFloat(tokens[++i]);
                const y1 = parseFloat(tokens[++i]);
                const x2 = parseFloat(tokens[++i]);
                const y2 = parseFloat(tokens[++i]);
                const x = parseFloat(tokens[++i]);
                const y = parseFloat(tokens[++i]);

                // Flatten the curve into lines for simplicity and robustness
                // 5 segments
                const steps = 10;
                for (let s = 1; s <= steps; s++) {
                    const t = s / steps;
                    const mt = 1 - t;
                    // Bezier formula
                    const bx = mt*mt*mt*currentX + 3*mt*mt*t*x1 + 3*mt*t*t*x2 + t*t*t*x;
                    const by = mt*mt*mt*currentY + 3*mt*mt*t*y1 + 3*mt*t*t*y2 + t*t*t*y;
                    
                    geomXml.ele('Row', { T: 'LineTo', IX: rowIx.toString() })
                        .ele('Cell', { N: 'X', V: toVisioX(bx).toString() }).up()
                        .ele('Cell', { N: 'Y', V: toVisioY(by).toString() }).up().up();
                    rowIx++;
                }

                currentX = x;
                currentY = y;
            }
        }
    }
}
