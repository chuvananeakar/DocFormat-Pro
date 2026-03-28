import JSZip from 'jszip';
import { ProcessResult, DocxOptions, HeaderType } from '../types';

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const TWIPS_PER_CM = 567;
const TWIPS_PER_PT = 20;

const DOC_TYPE_KEYWORDS = [
  "NGHỊ QUYẾT", "QUYẾT ĐỊNH", "CHỈ THỊ", "KẾT LUẬN", "QUY CHẾ", "QUY ĐỊNH", 
  "HƯỚNG DẪN", "BÁO CÁO", "KẾ HOẠCH", "CHƯƠNG TRÌNH", "THÔNG BÁO", "THÔNG TRI", 
  "CÔNG VĂN", "TỜ TRÌNH", "BIÊN BẢN", "PHƯƠNG ÁN", "ĐỀ ÁN", "DỰ ÁN", 
  "HỢP ĐỒNG", "BẢN THỎA THUẬN", "GIẤY ỦY QUYỀN", "GIẤY MỜI", "GIẤY GIỚI THIỆU", "GIẤY NGHỈ PHÉP"
];

const DEFAULT_OPTIONS: DocxOptions = {
  headerType: HeaderType.NONE,
  removeNumbering: false,
  margins: { top: 2, bottom: 2, left: 3, right: 1.5 },
  font: { family: "Times New Roman", sizeNormal: 14, sizeTable: 13 },
  paragraph: { lineSpacing: 1.15, after: 6, indent: 1.27 },
  table: { rowHeight: 0.8 }
};

const ACRONYMS_LIST = [
  "UBND", "THCS", "THPT", "BGDĐT", "SGDĐT", "PGDĐT", "ĐTN", "CĐ", "ĐCS", "VN", 
  "GDĐT", "CNTT", "KHTN", "KHXH", "GDCD", "TDTT", "BCH", "CSCS", "CMHS", "ĐĐ", 
  "BĐD", "STT", "GV", "HS", "SKKN",
  "NQ", "QĐ", "CT", "KL", "QC", "QYĐ", "HD", "BC", "KH", "CTR", "TB", "TTR", "CV", "BB",
  "PA", "ĐA", "DA", "HĐ", "BTT", "GUQ", "GM", "GGT", "GNP"
];

const smartNormalizeText = (text: string): string => {
    let t = text.trim();
    if (!t) return "";
    const lowerRegex = /[a-zàáâãèéêìíòóôõùúăđĩũơưạảấầẩẫậắằẳẵặẹẻẽềềểễệỉịọỏốồổỗộớờởỡợụủứừửữựỳýỷỹỵ]/;
    const specialRegex = /[0-9\-\/]/;
    
    t = t.split(/\s+/).map(word => {
        if (lowerRegex.test(word) || specialRegex.test(word)) {
            return word;
        }
        return word.replace(/[A-ZÀÁÂÃÈÉÊÌÍÒÓÔÕÙÚĂĐĨŨƠƯẠẢẤẦẨẪẬẮẰẲẴẶẸẺẼỀỀỂỄỆỈỊỌỎỐỒỔỖỘỚỜỞỠỢỤỦỨỪỬỮỰỲÝỶỸỴ]+/g, (match) => {
            if (ACRONYMS_LIST.includes(match)) return match;
            return match.toLowerCase();
        });
    }).join(' ');

    if (t.length > 0) {
        t = t.charAt(0).toUpperCase() + t.slice(1);
    }
    return t;
};

const normalizeTableHeader = (text: string): string => {
    let t = text.trim();
    if (!t) return "";
    t = t.toLowerCase();
    t = t.charAt(0).toUpperCase() + t.slice(1);
    ACRONYMS_LIST.forEach(acro => {
        const regex = new RegExp(`\\b${acro.toLowerCase()}\\b`, 'gi');
        t = t.replace(regex, acro);
    });
    return t;
};

const normalizeSummary = (text: string): string => {
    let summary = text.trim();
    if (!summary) return "";
    summary = summary.replace(/^[:-]\s*/, '').trim();
    summary = smartNormalizeText(summary);
    
    summary = summary.replace(/(^|\s)(?:-|–)?\s*tháng\s+(\d{1,2})(?:\/|-)(\d{4})/gi, (match, prefix, m, y) => {
        return `${prefix}tháng ${m.padStart(2, '0')} năm ${y}`;
    });
    
    const currentYear = new Date().getFullYear();
    summary = summary.replace(/(^|\s)(?:-|–)?\s*tháng\s+(\d{1,2})\b(?!\s*năm|\/|-)/gi, (match, prefix, m) => {
        return `${prefix}tháng ${m.padStart(2, '0')} năm ${currentYear}`;
    });
    
    return summary.trim();
};

const isParagraphBold = (p: Element): boolean => {
    const runs = Array.from(p.getElementsByTagNameNS(W_NS, "r"));
    for (const r of runs) {
        const t = r.getElementsByTagNameNS(W_NS, "t")[0];
        if (t && t.textContent && t.textContent.trim().length > 0) {
            const rPr = r.getElementsByTagNameNS(W_NS, "rPr")[0];
            if (rPr) {
                const b = rPr.getElementsByTagNameNS(W_NS, "b")[0];
                if (b) {
                    const val = b.getAttributeNS(W_NS, "val");
                    if (val !== "false" && val !== "0") return true;
                }
            }
        }
    }
    return false;
};

const enforceSchema = (doc: Document) => {
    const orders: Record<string, string[]> = {
        "w:pPr": [
            "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl", 
            "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens", 
            "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", 
            "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", 
            "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment", 
            "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange"
        ],
        "w:rPr": [
            "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike", 
            "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", 
            "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", 
            "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", 
            "w:eastAsianLayout", "w:specVanish", "w:oMath", "w:rPrChange"
        ],
        "w:tblPr": [
            "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual", "w:tblStyleRowBandSize", 
            "w:tblStyleColBandSize", "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders", 
            "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription", "w:tblPrChange"
        ],
        "w:trPr": [
            "w:cnfStyle", "w:divId", "w:gridBefore", "w:gridAfter", "w:wBefore", "w:wAfter", "w:cantSplit", 
            "w:trHeight", "w:tblHeader", "w:tblCellSpacing", "w:jc", "w:hidden", "w:trPrChange"
        ],
        "w:tcPr": [
            "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders", "w:shd", "w:noWrap", 
            "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark", "w:headers", "w:cellIns", 
            "w:cellDel", "w:tcPrChange"
        ],
        "w:r": [
            "w:rPr", "w:t", "w:br", "w:tab", "w:drawing", "w:pict", "w:object", "w:sym", "w:fldChar", 
            "w:instrText", "w:delText", "w:insText"
        ]
    };

    Object.keys(orders).forEach(tagName => {
        const localName = tagName.split(":")[1];
        const elements = Array.from(doc.getElementsByTagNameNS(W_NS, localName));
        elements.forEach(el => {
            const order = orders[tagName];
            const children = Array.from(el.childNodes);
            children.sort((a, b) => {
                const nameA = a.nodeName.includes(":") ? a.nodeName : `w:${a.nodeName}`;
                const nameB = b.nodeName.includes(":") ? b.nodeName : `w:${b.nodeName}`;
                const indexA = order.indexOf(nameA);
                const indexB = order.indexOf(nameB);

                if (indexA !== -1 && indexB !== -1) return indexA - indexB;
                if (indexA === -1 && indexB !== -1) return -1;
                if (indexA !== -1 && indexB === -1) return 1;
                return 0;
            });
            children.forEach(c => el.appendChild(c));
        });
    });
};

export const processDocx = async (file: File, options: DocxOptions = DEFAULT_OPTIONS): Promise<ProcessResult> => {
  const logs: string[] = [];
  try {
    logs.push(`Loading file: ${file.name}`);
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);
    const docXmlPath = "word/document.xml";
    const docXmlContent = await zip.file(docXmlPath)?.async("string");

    if (!docXmlContent) throw new Error("Invalid DOCX: missing word/document.xml");

    const parser = new DOMParser();
    const doc = parser.parseFromString(docXmlContent, "application/xml");
    const body = doc.getElementsByTagNameNS(W_NS, "body")[0];
    
    const createElement = (tagName: string) => doc.createElementNS(W_NS, tagName);
    const getOrCreate = (parent: Element, tagName: string): Element => {
      const localName = tagName.includes(":") ? tagName.split(":")[1] : tagName;
      let child = parent.getElementsByTagNameNS(W_NS, localName)[0];
      if (!child) {
        child = doc.createElementNS(W_NS, tagName);
        parent.appendChild(child);
      }
      return child;
    };

    const isTableParagraph = (p: Element): boolean => {
      let parent = p.parentNode;
      while(parent) {
        if (parent.nodeName === 'w:tbl' || parent.nodeName === 'tbl') return true;
        parent = parent.parentNode;
      }
      return false;
    };

    const paragraphsForCleaning = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
    for (const p of paragraphsForCleaning) {
        const textNodes = Array.from(p.getElementsByTagNameNS(W_NS, "t"));
        if (textNodes.length > 0) {
            const firstNode = textNodes[0];
            if (firstNode.textContent) firstNode.textContent = firstNode.textContent.trimStart();
            const lastNode = textNodes[textNodes.length - 1];
            if (lastNode.textContent) lastNode.textContent = lastNode.textContent.trimEnd();
        }
        const fullText = textNodes.map(n => n.textContent || "").join("");
        const hasContent = p.getElementsByTagNameNS(W_NS, "drawing").length > 0 || 
                           p.getElementsByTagNameNS(W_NS, "pict").length > 0 || 
                           p.getElementsByTagNameNS(W_NS, "object").length > 0 || 
                           p.getElementsByTagNameNS(W_NS, "br").length > 0;
        if (!hasContent && fullText.length === 0) p.parentNode?.removeChild(p);
    }

    if (options.removeNumbering) {
        const allParagraphs = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
        for (const p of allParagraphs) {
            const pPr = p.getElementsByTagNameNS(W_NS, "pPr")[0];
            if (pPr) {
                const numPr = pPr.getElementsByTagNameNS(W_NS, "numPr")[0];
                if (numPr) pPr.removeChild(numPr);
                const pStyle = getOrCreate(pPr, "w:pStyle");
                pStyle.setAttributeNS(W_NS, "w:val", "Normal");
            }
            const firstRun = p.getElementsByTagNameNS(W_NS, "r")[0];
            if (firstRun) {
                const firstText = firstRun.getElementsByTagNameNS(W_NS, "t")[0];
                if (firstText && firstText.textContent) {
                    const bulletRegex = /^[\s]*([•\-\–\—\*]|(\d+\.))[\s]+/;
                    if (bulletRegex.test(firstText.textContent)) {
                        firstText.textContent = firstText.textContent.replace(bulletRegex, "").trimStart();
                    }
                }
            }
        }
    }

    const sectPrsDoc = Array.from(doc.getElementsByTagNameNS(W_NS, "sectPr"));
    for (const sectPr of sectPrsDoc) {
        const pgSz = getOrCreate(sectPr, "w:pgSz");
        const wStr = pgSz.getAttributeNS(W_NS, "w");
        const hStr = pgSz.getAttributeNS(W_NS, "h");
        const orient = pgSz.getAttributeNS(W_NS, "orient");

        let w = parseInt(wStr || "0");
        let h = parseInt(hStr || "0");
        
        let isLandscape = orient === "landscape" || w > h;

        if (isLandscape) {
            pgSz.setAttributeNS(W_NS, "w", String(Math.round(29.7 * TWIPS_PER_CM)));
            pgSz.setAttributeNS(W_NS, "w:h", String(Math.round(21 * TWIPS_PER_CM)));
            pgSz.setAttributeNS(W_NS, "w:orient", "landscape");
        } else {
            pgSz.setAttributeNS(W_NS, "w", String(Math.round(21 * TWIPS_PER_CM)));
            pgSz.setAttributeNS(W_NS, "w:h", String(Math.round(29.7 * TWIPS_PER_CM)));
            pgSz.setAttributeNS(W_NS, "w:orient", "portrait");
        }

        const pgMar = getOrCreate(sectPr, "w:pgMar");
        pgMar.setAttributeNS(W_NS, "w:top", String(Math.round(options.margins.top * TWIPS_PER_CM)));
        pgMar.setAttributeNS(W_NS, "w:bottom", String(Math.round(options.margins.bottom * TWIPS_PER_CM)));
        pgMar.setAttributeNS(W_NS, "w:left", String(Math.round(options.margins.left * TWIPS_PER_CM)));
        pgMar.setAttributeNS(W_NS, "w:right", String(Math.round(options.margins.right * TWIPS_PER_CM)));
    }

    const rebuildParagraph = (p: Element, text: string, isBold: boolean, fontSize: string, isTitle: boolean) => {
        Array.from(p.childNodes).forEach(child => {
            if (child.nodeName !== "w:pPr") p.removeChild(child);
        });
        const r = doc.createElementNS(W_NS, "w:r");
        const rPr = doc.createElementNS(W_NS, "w:rPr");
        r.appendChild(rPr); 
        const pPr = getOrCreate(p, "w:pPr");
        const jc = getOrCreate(pPr, "w:jc");
        jc.setAttributeNS(W_NS, "w:val", "center");
        const ind = getOrCreate(pPr, "w:ind");
        ind.setAttributeNS(W_NS, "w:left", "0");
        ind.setAttributeNS(W_NS, "w:right", "0");
        ind.setAttributeNS(W_NS, "w:firstLine", "0");
        ind.removeAttributeNS(W_NS, "hanging");
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", isTitle ? "480" : "0");
        spacing.setAttributeNS(W_NS, "w:after", "0"); 
        spacing.setAttributeNS(W_NS, "w:line", "240");
        spacing.setAttributeNS(W_NS, "w:lineRule", "auto");
        const rFonts = getOrCreate(rPr, "w:rFonts");
        rFonts.setAttributeNS(W_NS, "w:ascii", options.font.family);
        rFonts.setAttributeNS(W_NS, "w:hAnsi", options.font.family);
        rFonts.setAttributeNS(W_NS, "w:cs", options.font.family);
        rFonts.setAttributeNS(W_NS, "w:eastAsia", options.font.family);
        const b = getOrCreate(rPr, "w:b");
        b.setAttributeNS(W_NS, "w:val", isBold ? "true" : "false"); 
        const iEl = getOrCreate(rPr, "w:i");
        iEl.setAttributeNS(W_NS, "w:val", "false");
        const sz = getOrCreate(rPr, "w:sz");
        sz.setAttributeNS(W_NS, "w:val", fontSize);
        const szCs = getOrCreate(rPr, "w:szCs");
        szCs.setAttributeNS(W_NS, "w:val", fontSize);
        const t = doc.createElementNS(W_NS, "w:t");
        t.textContent = text;
        r.appendChild(t);
        p.appendChild(r);
    };

    const createTitleUnderlineFrag = (protectedElements: Set<Element>, lineTables: Set<Element>): DocumentFragment => {
        const frag = doc.createDocumentFragment();
        const tbl = doc.createElementNS(W_NS, "w:tbl");
        lineTables.add(tbl); 
        const tblPr = doc.createElementNS(W_NS, "w:tblPr");
        tbl.appendChild(tblPr);
        const jcTbl = doc.createElementNS(W_NS, "w:jc");
        jcTbl.setAttributeNS(W_NS, "w:val", "center");
        tblPr.appendChild(jcTbl);
        const tblW = doc.createElementNS(W_NS, "w:tblW");
        tblW.setAttributeNS(W_NS, "w:w", "1500");
        tblW.setAttributeNS(W_NS, "w:type", "dxa");
        tblPr.appendChild(tblW);
        const tblLayout = doc.createElementNS(W_NS, "w:tblLayout");
        tblLayout.setAttributeNS(W_NS, "w:type", "fixed");
        tblPr.appendChild(tblLayout);
        const tblGrid = doc.createElementNS(W_NS, "w:tblGrid");
        const gridCol = doc.createElementNS(W_NS, "w:gridCol");
        gridCol.setAttributeNS(W_NS, "w:w", "1500");
        tblGrid.appendChild(gridCol);
        tbl.appendChild(tblGrid);
        const tr = doc.createElementNS(W_NS, "w:tr");
        tbl.appendChild(tr);
        const trPr = doc.createElementNS(W_NS, "w:trPr");
        const trHeight = doc.createElementNS(W_NS, "w:trHeight");
        trHeight.setAttributeNS(W_NS, "w:val", String(Math.round(0.1 * TWIPS_PER_CM)));
        trHeight.setAttributeNS(W_NS, "w:hRule", "exact");
        trPr.appendChild(trHeight);
        tr.appendChild(trPr);
        const tc = doc.createElementNS(W_NS, "w:tc");
        tr.appendChild(tc);
        const tcPr = doc.createElementNS(W_NS, "w:tcPr");
        tc.appendChild(tcPr);
        const tcW = doc.createElementNS(W_NS, "w:tcW");
        tcW.setAttributeNS(W_NS, "w:w", "1500");
        tcW.setAttributeNS(W_NS, "w:type", "dxa");
        tcPr.appendChild(tcW);
        
        const tcMar = doc.createElementNS(W_NS, "w:tcMar");
        ["top", "bottom", "left", "right"].forEach(side => {
            const mar = doc.createElementNS(W_NS, `w:${side}`);
            mar.setAttributeNS(W_NS, "w:w", "0");
            mar.setAttributeNS(W_NS, "w:type", "dxa");
            tcMar.appendChild(mar);
        });
        tcPr.appendChild(tcMar);

        const tcBorders = doc.createElementNS(W_NS, "w:tcBorders");
        const top = doc.createElementNS(W_NS, "w:top");
        top.setAttributeNS(W_NS, "w:val", "single");
        top.setAttributeNS(W_NS, "w:sz", "6"); 
        top.setAttributeNS(W_NS, "w:space", "0");
        top.setAttributeNS(W_NS, "w:color", "000000");
        tcBorders.appendChild(top);
        tcPr.appendChild(tcBorders);
        const p = doc.createElementNS(W_NS, "w:p");
        const pPr = doc.createElementNS(W_NS, "w:pPr");
        p.appendChild(pPr);
        const spacing = doc.createElementNS(W_NS, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "0");
        spacing.setAttributeNS(W_NS, "w:line", "24"); 
        spacing.setAttributeNS(W_NS, "w:lineRule", "exact");
        pPr.appendChild(spacing);
        tc.appendChild(p);
        protectedElements.add(p);
        frag.appendChild(tbl);
        const safeP = doc.createElementNS(W_NS, "w:p");
        const safePPr = doc.createElementNS(W_NS, "w:pPr");
        safeP.appendChild(safePPr);
        const safeSpacing = doc.createElementNS(W_NS, "w:spacing");
        safeSpacing.setAttributeNS(W_NS, "w:before", "0");
        safeSpacing.setAttributeNS(W_NS, "w:after", "120"); 
        safeSpacing.setAttributeNS(W_NS, "w:line", "2"); 
        safeSpacing.setAttributeNS(W_NS, "w:lineRule", "exact");
        safePPr.appendChild(safeSpacing);
        protectedElements.add(safeP);
        frag.appendChild(safeP);
        return frag;
    };

    const createPartyDashLine = (protectedElements: Set<Element>): Element => {
        const p = doc.createElementNS(W_NS, "w:p");
        const pPr = getOrCreate(p, "w:pPr");
        const jc = getOrCreate(pPr, "w:jc");
        jc.setAttributeNS(W_NS, "w:val", "center");
        const ind = getOrCreate(pPr, "w:ind");
        ind.setAttributeNS(W_NS, "w:left", "0");
        ind.setAttributeNS(W_NS, "w:right", "0");
        ind.setAttributeNS(W_NS, "w:firstLine", "0");
        ind.removeAttributeNS(W_NS, "hanging");
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "120"); 
        spacing.setAttributeNS(W_NS, "w:line", "240"); 
        spacing.setAttributeNS(W_NS, "w:lineRule", "auto");
        const r = doc.createElementNS(W_NS, "w:r");
        const rPr = getOrCreate(r, "w:rPr");
        const rFonts = getOrCreate(rPr, "w:rFonts");
        rFonts.setAttributeNS(W_NS, "w:ascii", options.font.family);
        rFonts.setAttributeNS(W_NS, "w:hAnsi", options.font.family);
        rFonts.setAttributeNS(W_NS, "w:cs", options.font.family);
        rFonts.setAttributeNS(W_NS, "w:eastAsia", options.font.family);
        const b = getOrCreate(rPr, "w:b");
        b.setAttributeNS(W_NS, "w:val", "false"); 
        const sz = getOrCreate(rPr, "w:sz");
        sz.setAttributeNS(W_NS, "w:val", String(options.font.sizeNormal * 2));
        const szCs = getOrCreate(rPr, "w:szCs");
        szCs.setAttributeNS(W_NS, "w:val", String(options.font.sizeNormal * 2));
        const t = doc.createElementNS(W_NS, "w:t");
        t.textContent = "-----";
        r.appendChild(t);
        p.appendChild(r);
        protectedElements.add(p);
        return p;
    };

    const paragraphs = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
    let detectedDocType = ""; 
    const docTypeElements = new Set<Element>();
    const abstractElements = new Set<Element>();
    const protectedElements = new Set<Element>();
    const lineTables = new Set<Element>();

    const limit = Math.min(paragraphs.length, 20); 
    
    for (let i = 0; i < limit; i++) {
        const p = paragraphs[i];
        if (isTableParagraph(p)) continue;
        const text = p.textContent?.trim() || "";
        if (!text) continue;

        const cleanText = text.toUpperCase().replace(/^[^A-ZÀ-Ỹ]+/, '');
        const matchedKeyword = DOC_TYPE_KEYWORDS.find(k => cleanText.startsWith(k));

        if (matchedKeyword) {
            docTypeElements.add(p);
            detectedDocType = matchedKeyword; 
            
            const summaryParagraphs: Element[] = [];
            const originalUpper = text.toUpperCase();
            const keywordIndex = originalUpper.indexOf(matchedKeyword);
            const remainingText = text.slice(keywordIndex + matchedKeyword.length).trim();
            
            rebuildParagraph(p, matchedKeyword, true, "28", true); 

            if (remainingText.length > 3) {
                const newP = doc.createElementNS(W_NS, "w:p");
                if (p.nextSibling) p.parentNode?.insertBefore(newP, p.nextSibling);
                else p.parentNode?.appendChild(newP);
                summaryParagraphs.push(newP);
                rebuildParagraph(newP, normalizeSummary(remainingText), true, String(options.font.sizeNormal * 2), false); 
            }

            let linesCaptured = 0;
            let currentIndex = i + 1; 
            let hasFoundBoldSummary = false; 

            while (currentIndex < paragraphs.length && linesCaptured < 8) { 
                const tempP = paragraphs[currentIndex];
                if (isTableParagraph(tempP)) break;
                const tempText = tempP.textContent?.trim() || "";

                if (tempText.length > 0) {
                    const upperText = tempText.toUpperCase();
                    
                    if (tempText.startsWith("Căn cứ") || tempText.startsWith("Xét") || tempText.startsWith("Theo") || tempText.startsWith("Kính gửi") || tempText.startsWith("Hôm nay") || tempText.startsWith("Thời gian:") || tempText.startsWith("Đồng kính gửi")) break;
                    
                    if (/^([IVXLCDM]+|[0-9]+)[\.\)]\s/.test(tempText)) break;
                    if (tempText.length > 250) break; 

                    const isBold = isParagraphBold(tempP);
                    
                    if (hasFoundBoldSummary && !isBold) {
                        break;
                    }

                    if (isBold) {
                        hasFoundBoldSummary = true;
                    }

                    if (!hasFoundBoldSummary && linesCaptured >= 3) break;

                    summaryParagraphs.push(tempP);
                    rebuildParagraph(tempP, normalizeSummary(tempText), true, String(options.font.sizeNormal * 2), false); 
                    linesCaptured++;
                } else {
                    abstractElements.add(tempP);
                }
                currentIndex++;
            }

            summaryParagraphs.forEach(sp => abstractElements.add(sp));

            const targetNode = summaryParagraphs.length > 0 ? summaryParagraphs[summaryParagraphs.length - 1] : p;
            
            if (options.headerType === HeaderType.PARTY) {
                const dashP = createPartyDashLine(protectedElements);
                if (targetNode.nextSibling) targetNode.parentNode?.insertBefore(dashP, targetNode.nextSibling);
                else targetNode.parentNode?.appendChild(dashP);
            } else {
                const underlineFrag = createTitleUnderlineFrag(protectedElements, lineTables);
                if (targetNode.nextSibling) targetNode.parentNode?.insertBefore(underlineFrag, targetNode.nextSibling);
                else targetNode.parentNode?.appendChild(underlineFrag);
            }
            break; 
        }
    }

    const finalParagraphs = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
    for (const p of finalParagraphs) {
      if (docTypeElements.has(p) || abstractElements.has(p) || protectedElements.has(p)) continue; 
      
      const isTable = isTableParagraph(p);
      if (isTable) continue; 
      
      const pPr = getOrCreate(p, "w:pPr");
      const pText = p.textContent?.trim() || "";
      const upperText = pText.toUpperCase();

      let isDecisionSpecialLine = false;
      if (detectedDocType === "QUYẾT ĐỊNH" && pText.length > 0) {
          if (upperText === "QUYẾT ĐỊNH:" || upperText === "QUYẾT ĐỊNH") {
              isDecisionSpecialLine = true;
          } 
          else if (pText === upperText && pText.length < 150 && /[A-ZÀ-Ỹ]/.test(upperText)) {
              isDecisionSpecialLine = true;
          }
      }

      if (isDecisionSpecialLine) {
        const jc = getOrCreate(pPr, "w:jc");
        jc.setAttributeNS(W_NS, "w:val", "center");
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "120");
        spacing.setAttributeNS(W_NS, "w:after", "120");
        spacing.setAttributeNS(W_NS, "w:line", "240"); 
        spacing.setAttributeNS(W_NS, "w:lineRule", "auto");
        const ind = getOrCreate(pPr, "w:ind");
        ind.setAttributeNS(W_NS, "w:left", "0");
        ind.setAttributeNS(W_NS, "w:right", "0");
        ind.setAttributeNS(W_NS, "w:firstLine", "0");
        ind.removeAttributeNS(W_NS, "hanging");

        const targetSize = options.font.sizeNormal * 2;
        const runs = Array.from(p.getElementsByTagNameNS(W_NS, "r"));
        for (const r of runs) {
            const rPr = getOrCreate(r, "w:rPr");
            const b = getOrCreate(rPr, "w:b");
            b.setAttributeNS(W_NS, "w:val", "true");
            const sz = getOrCreate(rPr, "w:sz");
            sz.setAttributeNS(W_NS, "w:val", String(targetSize));
            const szCs = getOrCreate(rPr, "w:szCs");
            szCs.setAttributeNS(W_NS, "w:val", String(targetSize));
        }
        continue; 
      }

      const jc = getOrCreate(pPr, "w:jc");
      jc.setAttributeNS(W_NS, "w:val", "both");
      const spacing = getOrCreate(pPr, "w:spacing");
      spacing.setAttributeNS(W_NS, "w:before", "0");
      spacing.setAttributeNS(W_NS, "w:after", String(Math.round(options.paragraph.after * TWIPS_PER_PT)));
      spacing.setAttributeNS(W_NS, "w:line", String(Math.round(options.paragraph.lineSpacing * 240))); 
      spacing.setAttributeNS(W_NS, "w:lineRule", "auto");
      const ind = getOrCreate(pPr, "w:ind");
      ind.setAttributeNS(W_NS, "w:left", "0");
      ind.setAttributeNS(W_NS, "w:right", "0");
      ind.setAttributeNS(W_NS, "w:firstLine", String(Math.round(options.paragraph.indent * TWIPS_PER_CM)));
      
      const targetSize = options.font.sizeNormal * 2;
      const runs = Array.from(p.getElementsByTagNameNS(W_NS, "r"));
      for (const r of runs) {
          const rPr = getOrCreate(r, "w:rPr");
          const sz = getOrCreate(rPr, "w:sz");
          sz.setAttributeNS(W_NS, "w:val", String(targetSize));
          const szCs = getOrCreate(rPr, "w:szCs");
          szCs.setAttributeNS(W_NS, "w:val", String(targetSize));
      }
    }

    const tables = Array.from(doc.getElementsByTagNameNS(W_NS, "tbl"));
    for (const tbl of tables) {
        if (lineTables.has(tbl)) continue;
        
        let sttColIndex = -1;

        const rows = Array.from(tbl.getElementsByTagNameNS(W_NS, "tr"));
        for (let rIndex = 0; rIndex < rows.length; rIndex++) {
            const tr = rows[rIndex];
            const trPr = getOrCreate(tr, "w:trPr");

            let isHeaderRow = (rIndex === 0);
            if (!isHeaderRow && trPr.getElementsByTagNameNS(W_NS, "tblHeader").length > 0) {
                isHeaderRow = true;
            }

            let trHeight = getOrCreate(trPr, "w:trHeight");
            trHeight.setAttributeNS(W_NS, "w:val", String(Math.round(options.table.rowHeight * TWIPS_PER_CM))); 
            trHeight.setAttributeNS(W_NS, "w:hRule", "atLeast"); 

            const cells = Array.from(tr.getElementsByTagNameNS(W_NS, "tc"));
            let logicalColIndex = 0;

            for (const tc of cells) {
                const tcPr = getOrCreate(tc, "w:tcPr");
                const vAlign = getOrCreate(tcPr, "w:vAlign");
                vAlign.setAttributeNS(W_NS, "w:val", "center"); 
                
                const gridSpanEl = tcPr.getElementsByTagNameNS(W_NS, "gridSpan")[0];
                const gridSpan = gridSpanEl ? parseInt(gridSpanEl.getAttributeNS(W_NS, "val") || "1") : 1;

                const tcParagraphs = Array.from(tc.getElementsByTagNameNS(W_NS, "p"));
                
                if (isHeaderRow) {
                    const cellFullText = tc.textContent || "";
                    const cleanText = cellFullText.toUpperCase().replace(/\./g, '').trim();
                    if (cleanText === "STT" || cleanText === "SỐ TT") {
                        sttColIndex = logicalColIndex;
                    }
                }

                for (const p of tcParagraphs) {
                    const pPr = getOrCreate(p, "w:pPr");
                    const jc = getOrCreate(pPr, "w:jc");
                    
                    if (isHeaderRow || logicalColIndex === sttColIndex) {
                        jc.setAttributeNS(W_NS, "w:val", "center");
                    } else {
                        jc.setAttributeNS(W_NS, "w:val", "left");
                    }
                    
                    const ind = getOrCreate(pPr, "w:ind");
                    ind.setAttributeNS(W_NS, "w:left", "0");
                    ind.setAttributeNS(W_NS, "w:right", "0");
                    ind.setAttributeNS(W_NS, "w:firstLine", "0");
                    ind.removeAttributeNS(W_NS, "hanging");

                    const runs = Array.from(p.getElementsByTagNameNS(W_NS, "r"));
                    
                    if (isHeaderRow) {
                        const fullText = p.textContent || "";
                        const hasLetters = /[A-ZÀ-Ỹa-zà-ỹ]/.test(fullText);
                        
                        if (hasLetters) {
                            const cleanT = normalizeTableHeader(fullText);
                            
                            Array.from(p.childNodes).forEach(child => {
                                if (child.nodeName !== "w:pPr") p.removeChild(child);
                            });
                            
                            const r = doc.createElementNS(W_NS, "w:r");
                            const rPr = doc.createElementNS(W_NS, "w:rPr");
                            r.appendChild(rPr);
                            
                            const b = doc.createElementNS(W_NS, "w:b");
                            b.setAttributeNS(W_NS, "w:val", "true"); 
                            rPr.appendChild(b);
                            
                            const sz = doc.createElementNS(W_NS, "w:sz");
                            sz.setAttributeNS(W_NS, "w:val", "26"); 
                            rPr.appendChild(sz);
                            const szCs = doc.createElementNS(W_NS, "w:szCs");
                            szCs.setAttributeNS(W_NS, "w:val", "26");
                            rPr.appendChild(szCs);
                            
                            const t = doc.createElementNS(W_NS, "w:t");
                            t.textContent = cleanT;
                            r.appendChild(t);
                            p.appendChild(r);
                        } else {
                            for (const r of runs) {
                                const rPr = getOrCreate(r, "w:rPr");
                                const b = getOrCreate(rPr, "w:b");
                                b.setAttributeNS(W_NS, "w:val", "true");
                                const sz = getOrCreate(rPr, "w:sz");
                                sz.setAttributeNS(W_NS, "w:val", "26");
                                const szCs = getOrCreate(rPr, "w:szCs");
                                szCs.setAttributeNS(W_NS, "w:val", "26");
                            }
                        }
                    } else {
                        for (const r of runs) {
                            const rPr = getOrCreate(r, "w:rPr");
                            const sz = getOrCreate(rPr, "w:sz");
                            sz.setAttributeNS(W_NS, "w:val", "26"); 
                            const szCs = getOrCreate(rPr, "w:szCs");
                            szCs.setAttributeNS(W_NS, "w:val", "26");
                        }
                    }
                }
                logicalColIndex += gridSpan;
            }
        }
    }
    
    if (options.headerType !== HeaderType.NONE && body) {
        const headerTable = createHeaderTemplate(doc, options);
        if (body.firstChild) body.insertBefore(headerTable, body.firstChild);
        else body.appendChild(headerTable);

        const sectPrs = body.getElementsByTagNameNS(W_NS, "sectPr");
        const lastSectPr = sectPrs.length > 0 ? sectPrs[sectPrs.length - 1] : null;
        
        const blankP = doc.createElementNS(W_NS, "w:p");
        const signatureBlock = createSignatureBlock(doc, options as any, detectedDocType);
        
        if (lastSectPr && lastSectPr.parentNode === body) {
            body.insertBefore(blankP, lastSectPr);
            body.insertBefore(signatureBlock, lastSectPr);
        } else {
            body.appendChild(blankP);
            body.appendChild(signatureBlock);
        }
    }

    const allRunsInDoc = Array.from(doc.getElementsByTagNameNS(W_NS, "r"));
    for (const r of allRunsInDoc) { getOrCreate(r, "w:rPr"); }
    const allPPrsInDoc = Array.from(doc.getElementsByTagNameNS(W_NS, "pPr"));
    for (const pPr of allPPrsInDoc) { getOrCreate(pPr, "w:rPr"); }

    const allRPrs = Array.from(doc.getElementsByTagNameNS(W_NS, "rPr"));
    for (const rPr of allRPrs) {
        const rFonts = getOrCreate(rPr, "w:rFonts");
        rFonts.setAttributeNS(W_NS, "w:ascii", options.font.family);
        rFonts.setAttributeNS(W_NS, "w:hAnsi", options.font.family);
        rFonts.setAttributeNS(W_NS, "w:cs", options.font.family);
        rFonts.setAttributeNS(W_NS, "w:eastAsia", options.font.family);
        ["asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"].forEach(theme => {
            rFonts.removeAttributeNS(W_NS, theme);
            rFonts.removeAttribute(theme); 
        });
    }

    const serializer = new XMLSerializer();
    const modifyXmlFonts = async (filePath: string) => {
        const xmlContent = await zip.file(filePath)?.async("string");
        if (xmlContent) {
            const extDoc = parser.parseFromString(xmlContent, "application/xml");
            const getOrCreateExt = (parent: Element, tagName: string): Element => {
                const localName = tagName.includes(":") ? tagName.split(":")[1] : tagName;
                let child = parent.getElementsByTagNameNS(W_NS, localName)[0];
                if (!child) {
                    child = extDoc.createElementNS(W_NS, tagName);
                    parent.appendChild(child);
                }
                return child;
            };
            
            const rPrsExt = Array.from(extDoc.getElementsByTagNameNS(W_NS, "rPr"));
            for (const rPr of rPrsExt) {
                const rFonts = getOrCreateExt(rPr, "w:rFonts");
                rFonts.setAttributeNS(W_NS, "w:ascii", options.font.family);
                rFonts.setAttributeNS(W_NS, "w:hAnsi", options.font.family);
                rFonts.setAttributeNS(W_NS, "w:cs", options.font.family);
                rFonts.setAttributeNS(W_NS, "w:eastAsia", options.font.family);
                ["asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"].forEach(theme => {
                    rFonts.removeAttributeNS(W_NS, theme);
                    rFonts.removeAttribute(theme);
                });
            }

            const rPrDefaults = Array.from(extDoc.getElementsByTagNameNS(W_NS, "rPrDefault"));
            for (const rPrDef of rPrDefaults) {
                const rPr = getOrCreateExt(rPrDef, "w:rPr");
                const rFonts = getOrCreateExt(rPr, "w:rFonts");
                rFonts.setAttributeNS(W_NS, "w:ascii", options.font.family);
                rFonts.setAttributeNS(W_NS, "w:hAnsi", options.font.family);
                rFonts.setAttributeNS(W_NS, "w:cs", options.font.family);
                rFonts.setAttributeNS(W_NS, "w:eastAsia", options.font.family);
                ["asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"].forEach(theme => {
                    rFonts.removeAttributeNS(W_NS, theme);
                    rFonts.removeAttribute(theme);
                });
            }
            
            enforceSchema(extDoc); 
            zip.file(filePath, serializer.serializeToString(extDoc));
        }
    };

    await modifyXmlFonts("word/styles.xml");
    await modifyXmlFonts("word/numbering.xml");

    enforceSchema(doc);

    const fontSize = options.font.sizeTable * 2;
    const fontFamily = options.font.family;
    const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r>
                <w:rPr><w:rFonts w:ascii="${fontFamily}" w:hAnsi="${fontFamily}" w:cs="${fontFamily}" w:eastAsia="${fontFamily}"/><w:sz w:val="${fontSize}"/><w:szCs w:val="${fontSize}"/></w:rPr>
                <w:fldChar w:fldCharType="begin"/>
            </w:r>
            <w:r>
                <w:rPr><w:rFonts w:ascii="${fontFamily}" w:hAnsi="${fontFamily}" w:cs="${fontFamily}" w:eastAsia="${fontFamily}"/><w:sz w:val="${fontSize}"/><w:szCs w:val="${fontSize}"/></w:rPr>
                <w:instrText xml:space="preserve"> PAGE </w:instrText>
            </w:r>
            <w:r>
                <w:rPr><w:rFonts w:ascii="${fontFamily}" w:hAnsi="${fontFamily}" w:cs="${fontFamily}" w:eastAsia="${fontFamily}"/><w:sz w:val="${fontSize}"/><w:szCs w:val="${fontSize}"/></w:rPr>
                <w:fldChar w:fldCharType="separate"/>
            </w:r>
            <w:r>
                <w:rPr><w:rFonts w:ascii="${fontFamily}" w:hAnsi="${fontFamily}" w:cs="${fontFamily}" w:eastAsia="${fontFamily}"/><w:sz w:val="${fontSize}"/><w:szCs w:val="${fontSize}"/><w:noProof/></w:rPr>
                <w:t></w:t>
            </w:r>
            <w:r>
                <w:rPr><w:rFonts w:ascii="${fontFamily}" w:hAnsi="${fontFamily}" w:cs="${fontFamily}" w:eastAsia="${fontFamily}"/><w:sz w:val="${fontSize}"/><w:szCs w:val="${fontSize}"/></w:rPr>
                <w:fldChar w:fldCharType="end"/>
            </w:r>
        </w:p>
    </w:hdr>`;
    zip.file("word/header_custom.xml", headerXml);

    let contentTypesXml = await zip.file("[Content_Types].xml")?.async("string");
    if (contentTypesXml && !contentTypesXml.includes('PartName="/word/header_custom.xml"')) {
        contentTypesXml = contentTypesXml.replace(
            '</Types>',
            '<Override PartName="/word/header_custom.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'
        );
        zip.file("[Content_Types].xml", contentTypesXml);
    }

    let relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");
    if (relsXml && !relsXml.includes('Target="header_custom.xml"')) {
        relsXml = relsXml.replace(
            '</Relationships>',
            '<Relationship Id="rIdCustomHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header_custom.xml"/></Relationships>'
        );
        zip.file("word/_rels/document.xml.rels", relsXml);
    }

    const sectPrs = Array.from(doc.getElementsByTagNameNS(W_NS, "sectPr"));
    const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    for (const sPr of sectPrs) {
        getOrCreate(sPr, "w:titlePg");
        const headerRefs = Array.from(sPr.getElementsByTagNameNS(W_NS, "headerReference"));
        for (const hr of headerRefs) {
            if (hr.getAttributeNS(W_NS, "type") === "default") sPr.removeChild(hr);
        }
        const newHdrRef = doc.createElementNS(W_NS, "w:headerReference");
        newHdrRef.setAttributeNS(W_NS, "w:type", "default");
        newHdrRef.setAttributeNS(R_NS, "r:id", "rIdCustomHdr");
        if (sPr.firstChild) sPr.insertBefore(newHdrRef, sPr.firstChild);
        else sPr.appendChild(newHdrRef);
    }

    const newDocXml = serializer.serializeToString(doc);
    zip.file(docXmlPath, newDocXml);
    const generatedBlob = await zip.generateAsync({ type: "blob" });

    return { success: true, blob: generatedBlob, fileName: `formatted_${file.name}`, logs };
  } catch (error) {
    return { success: false, error: error instanceof Error ? error.message : "Unknown error", logs };
  }
};

const createHeaderTemplate = (doc: Document, options: DocxOptions): Element => {
    const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const createElement = (tagName: string) => doc.createElementNS(W_NS, tagName);
    const getOrCreate = (parent: Element, tagName: string): Element => {
      const localName = tagName.includes(":") ? tagName.split(":")[1] : tagName;
      let child = parent.getElementsByTagNameNS(W_NS, localName)[0];
      if (!child) {
        child = doc.createElementNS(W_NS, tagName);
        parent.appendChild(child);
      }
      return child;
    };

    const createStyledP = (text: string, isBold: boolean, isItalic: boolean, customSize?: number): Element => {
        const p = createElement("w:p");
        const pPr = getOrCreate(p, "w:pPr");
        const jc = getOrCreate(pPr, "w:jc");
        jc.setAttributeNS(W_NS, "w:val", "center");
        const ind = getOrCreate(pPr, "w:ind");
        ind.setAttributeNS(W_NS, "w:left", "0");
        ind.setAttributeNS(W_NS, "w:right", "0");
        ind.setAttributeNS(W_NS, "w:firstLine", "0");
        ind.removeAttributeNS(W_NS, "hanging");
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "0");
        spacing.setAttributeNS(W_NS, "w:line", "240"); 
        spacing.setAttributeNS(W_NS, "w:lineRule", "auto");
        const r = createElement("w:r");
        p.appendChild(r);
        const rPr = getOrCreate(r, "w:rPr");
        const sizeToUse = customSize ? customSize * 2 : options.font.sizeTable * 2;
        const sz = getOrCreate(rPr, "w:sz");
        sz.setAttributeNS(W_NS, "w:val", String(sizeToUse));
        const szCs = getOrCreate(rPr, "w:szCs");
        szCs.setAttributeNS(W_NS, "w:val", String(sizeToUse));
        if (isBold) rPr.appendChild(createElement("w:b"));
        if (isItalic) rPr.appendChild(createElement("w:i"));
        const t = createElement("w:t");
        t.textContent = text;
        r.appendChild(t);
        return p;
    };

    const createMottoP = (text: string, isBold: boolean, customSize?: number): Element => {
        const p = createElement("w:p");
        const pPr = getOrCreate(p, "w:pPr");
        const jc = getOrCreate(pPr, "w:jc");
        jc.setAttributeNS(W_NS, "w:val", "center");
        const ind = getOrCreate(pPr, "w:ind");
        ind.setAttributeNS(W_NS, "w:left", "0");
        ind.setAttributeNS(W_NS, "w:right", "0");
        ind.setAttributeNS(W_NS, "w:firstLine", "0");
        ind.removeAttributeNS(W_NS, "hanging");
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "0");
        spacing.setAttributeNS(W_NS, "w:line", "240"); 
        spacing.setAttributeNS(W_NS, "w:lineRule", "auto");
        const r = createElement("w:r");
        p.appendChild(r);
        const rPr = getOrCreate(r, "w:rPr");
        const sizeToUse = customSize ? customSize * 2 : options.font.sizeTable * 2;
        const sz = getOrCreate(rPr, "w:sz");
        sz.setAttributeNS(W_NS, "w:val", String(sizeToUse));
        const szCs = getOrCreate(rPr, "w:szCs");
        szCs.setAttributeNS(W_NS, "w:val", String(sizeToUse));
        if (isBold) rPr.appendChild(createElement("w:b"));
        
        const t = createElement("w:t");
        t.textContent = text;
        r.appendChild(t);
        return p;
    };

    const appendSafeTable = (tc: Element, tbl: Element) => {
        tc.appendChild(tbl);
        const p = createElement("w:p");
        const pPr = getOrCreate(p, "w:pPr");
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "0");
        spacing.setAttributeNS(W_NS, "w:line", "2"); 
        spacing.setAttributeNS(W_NS, "w:lineRule", "exact");
        tc.appendChild(p);
    };

    const createShortLineTable = (): Element => {
        const tbl = createElement("w:tbl");
        const tblPr = getOrCreate(tbl, "w:tblPr");
        const jcTbl = getOrCreate(tblPr, "w:jc");
        jcTbl.setAttributeNS(W_NS, "w:val", "center");
        const tblW = getOrCreate(tblPr, "w:tblW");
        tblW.setAttributeNS(W_NS, "w:w", "1000");
        tblW.setAttributeNS(W_NS, "w:type", "dxa");
        const tblLayout = getOrCreate(tblPr, "w:tblLayout");
        tblLayout.setAttributeNS(W_NS, "w:type", "fixed");
        const tblGrid = getOrCreate(tbl, "w:tblGrid");
        const gridCol = createElement("w:gridCol");
        gridCol.setAttributeNS(W_NS, "w:w", "1000");
        tblGrid.appendChild(gridCol);
        const tr = createElement("w:tr");
        tbl.appendChild(tr);
        const tc = createElement("w:tc");
        tr.appendChild(tc);
        const tcPr = getOrCreate(tc, "w:tcPr");
        const tcW = getOrCreate(tcPr, "w:tcW");
        tcW.setAttributeNS(W_NS, "w:w", "1000");
        tcW.setAttributeNS(W_NS, "w:type", "dxa");
        
        const tcMar = getOrCreate(tcPr, "w:tcMar");
        ["top", "bottom", "left", "right"].forEach(side => {
            const mar = getOrCreate(tcMar, `w:${side}`);
            mar.setAttributeNS(W_NS, "w:w", "0");
            mar.setAttributeNS(W_NS, "w:type", "dxa");
        });

        const tcBorders = getOrCreate(tcPr, "w:tcBorders");
        const top = getOrCreate(tcBorders, "w:top"); 
        top.setAttributeNS(W_NS, "w:val", "single");
        top.setAttributeNS(W_NS, "w:sz", "4"); 
        top.setAttributeNS(W_NS, "w:space", "0");
        top.setAttributeNS(W_NS, "w:color", "000000");
        const p = createElement("w:p");
        tc.appendChild(p);
        const pPr = getOrCreate(p, "w:pPr");
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "0");
        spacing.setAttributeNS(W_NS, "w:line", "24"); 
        spacing.setAttributeNS(W_NS, "w:lineRule", "exact");
        return tbl;
    };

    const createMottoLineTable = (widthTwips: string): Element => {
        const tbl = createElement("w:tbl");
        const tblPr = getOrCreate(tbl, "w:tblPr");
        const jcTbl = getOrCreate(tblPr, "w:jc");
        jcTbl.setAttributeNS(W_NS, "w:val", "center");
        const tblW = getOrCreate(tblPr, "w:tblW");
        tblW.setAttributeNS(W_NS, "w:w", widthTwips);
        tblW.setAttributeNS(W_NS, "w:type", "dxa");
        const tblLayout = getOrCreate(tblPr, "w:tblLayout");
        tblLayout.setAttributeNS(W_NS, "w:type", "fixed");
        const tblGrid = getOrCreate(tbl, "w:tblGrid");
        const gridCol = createElement("w:gridCol");
        gridCol.setAttributeNS(W_NS, "w:w", widthTwips);
        tblGrid.appendChild(gridCol);
        const tr = createElement("w:tr");
        tbl.appendChild(tr);
        const tc = createElement("w:tc");
        tr.appendChild(tc);
        const tcPr = getOrCreate(tc, "w:tcPr");
        const tcW = getOrCreate(tcPr, "w:tcW");
        tcW.setAttributeNS(W_NS, "w:w", widthTwips);
        tcW.setAttributeNS(W_NS, "w:type", "dxa");

        const tcMar = getOrCreate(tcPr, "w:tcMar");
        ["top", "bottom", "left", "right"].forEach(side => {
            const mar = getOrCreate(tcMar, `w:${side}`);
            mar.setAttributeNS(W_NS, "w:w", "0");
            mar.setAttributeNS(W_NS, "w:type", "dxa");
        });

        const tcBorders = getOrCreate(tcPr, "w:tcBorders");
        const top = getOrCreate(tcBorders, "w:top"); 
        top.setAttributeNS(W_NS, "w:val", "single");
        top.setAttributeNS(W_NS, "w:sz", "6"); 
        top.setAttributeNS(W_NS, "w:space", "0");
        top.setAttributeNS(W_NS, "w:color", "000000");
        const p = createElement("w:p");
        tc.appendChild(p);
        const pPr = getOrCreate(p, "w:pPr");
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "0");
        spacing.setAttributeNS(W_NS, "w:line", "24"); 
        spacing.setAttributeNS(W_NS, "w:lineRule", "exact");
        return tbl;
    };

    const tbl = createElement("w:tbl");
    const tblPr = getOrCreate(tbl, "w:tblPr");
    const tblBorders = getOrCreate(tblPr, "w:tblBorders");
    ["top", "left", "bottom", "right", "insideH", "insideV"].forEach(side => {
        const border = getOrCreate(tblBorders, `w:${side}`);
        border.setAttributeNS(W_NS, "w:val", "none");
    });
    
    const tblLayout = getOrCreate(tblPr, "w:tblLayout");
    tblLayout.setAttributeNS(W_NS, "w:type", "fixed");

    const tblW = getOrCreate(tblPr, "w:tblW");
    tblW.setAttributeNS(W_NS, "w:w", "9350"); 
    tblW.setAttributeNS(W_NS, "w:type", "dxa");

    const tblGrid = getOrCreate(tbl, "w:tblGrid");
    const col1 = createElement("w:gridCol");
    col1.setAttributeNS(W_NS, "w:w", "3800"); 
    tblGrid.appendChild(col1);
    const col2 = createElement("w:gridCol");
    col2.setAttributeNS(W_NS, "w:w", "5550"); 
    tblGrid.appendChild(col2);

    const tr = createElement("w:tr");
    tbl.appendChild(tr);

    const tc1 = createElement("w:tc");
    tr.appendChild(tc1);
    const tc1Pr = getOrCreate(tc1, "w:tcPr");
    const tc1W = getOrCreate(tc1Pr, "w:tcW");
    tc1W.setAttributeNS(W_NS, "w:w", "3800");
    tc1W.setAttributeNS(W_NS, "w:type", "dxa");
    
    // ĐÃ SỬA LỖI TYPO tc1Mar Ở ĐÂY
    const tc1Mar = getOrCreate(tc1Pr, "w:tcMar");
    ["top", "bottom", "left", "right"].forEach(side => {
        const mar = getOrCreate(tc1Mar, `w:${side}`);
        mar.setAttributeNS(W_NS, "w:w", "0");
        mar.setAttributeNS(W_NS, "w:type", "dxa");
    });
    
    const tc2 = createElement("w:tc");
    tr.appendChild(tc2);
    const tc2Pr = getOrCreate(tc2, "w:tcPr");
    const tc2W = getOrCreate(tc2Pr, "w:tcW");
    tc2W.setAttributeNS(W_NS, "w:w", "5550");
    tc2W.setAttributeNS(W_NS, "w:type", "dxa");

    const tc2Mar = getOrCreate(tc2Pr, "w:tcMar");
    ["top", "bottom", "left", "right"].forEach(side => {
        const mar = getOrCreate(tc2Mar, `w:${side}`);
        mar.setAttributeNS(W_NS, "w:w", "0");
        mar.setAttributeNS(W_NS, "w:type", "dxa");
    });

    const docDate = options.documentDate ? new Date(options.documentDate) : new Date();
    const day = String(docDate.getDate()).padStart(2, '0');
    const month = String(docDate.getMonth() + 1).padStart(2, '0');
    const year = docDate.getFullYear();
    const currentDateStr = `Ea Kar, ngày ${day} tháng ${month} năm ${year}`;

    switch (options.headerType) {
        case HeaderType.PARTY:
            let partySymbolStr = "Số: ... - .../CB";
            if (options.docSymbol || options.docSuffix) {
                const symbol = options.docSymbol || "...";
                const suffix = options.docSuffix || "CB";
                partySymbolStr = `Số: ... - ${symbol}/${suffix}`;
            }

            tc1.appendChild(createStyledP("ĐẢNG BỘ XÃ EA KAR", false, false));
            tc1.appendChild(createStyledP("CHI BỘ TRƯỜNG THCS CHU VĂN AN", true, false));
            tc1.appendChild(createStyledP("*", false, false)); 
            tc1.appendChild(createStyledP(partySymbolStr, false, false));

            tc2.appendChild(createMottoP("ĐẢNG CỘNG SẢN VIỆT NAM", true, 13)); 
            appendSafeTable(tc2, createMottoLineTable("3400")); 
            tc2.appendChild(createStyledP("", false, false));
            tc2.appendChild(createStyledP(currentDateStr, false, true, 14));
            break;

        case HeaderType.DEPARTMENT:
            const deptName = options.departmentName || "TỔ TOÁN - TIN";
            tc1.appendChild(createStyledP("TRƯỜNG THCS CHU VĂN AN", false, false));
            tc1.appendChild(createStyledP(deptName, true, false));
            appendSafeTable(tc1, createShortLineTable()); 
            tc1.appendChild(createStyledP("", false, false));  
            tc1.appendChild(createStyledP("Số: ... /...", false, false)); 

            tc2.appendChild(createStyledP("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", true, false, 13));
            tc2.appendChild(createMottoP("Độc lập - Tự do - Hạnh phúc", true, 13)); 
            appendSafeTable(tc2, createMottoLineTable("3200")); 
            tc2.appendChild(createStyledP("", false, false)); 
            tc2.appendChild(createStyledP(currentDateStr, false, true, 14));
            break;

        case HeaderType.SCHOOL:
        default:
            let schoolSymbolStr = "Số: ... /...";
            if (options.docSymbol || options.docSuffix) {
                const symbol = options.docSymbol || "...";
                const suffix = options.docSuffix || "...";
                schoolSymbolStr = `Số: ... /${symbol}-${suffix}`;
            }

            tc1.appendChild(createStyledP("UBND XÃ EA KAR", false, false));
            tc1.appendChild(createStyledP("TRƯỜNG THCS CHU VĂN AN", true, false));
            appendSafeTable(tc1, createShortLineTable());
            tc1.appendChild(createStyledP("", false, false));
            tc1.appendChild(createStyledP(schoolSymbolStr, false, false));

            tc2.appendChild(createStyledP("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", true, false, 13));
            tc2.appendChild(createMottoP("Độc lập - Tự do - Hạnh phúc", true, 13)); 
            appendSafeTable(tc2, createMottoLineTable("3200")); 
            tc2.appendChild(createStyledP("", false, false)); 
            tc2.appendChild(createStyledP(currentDateStr, false, true, 14));
            break;
    }

    return tbl;
};

const createSignatureBlock = (doc: Document, options: any, docType: string): Element => {
    const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const createElement = (tagName: string) => doc.createElementNS(W_NS, tagName);
    const getOrCreate = (parent: Element, tagName: string): Element => {
      const localName = tagName.includes(":") ? tagName.split(":")[1] : tagName;
      let child = parent.getElementsByTagNameNS(W_NS, localName)[0];
      if (!child) {
        child = doc.createElementNS(W_NS, tagName);
        parent.appendChild(child);
      }
      return child;
    };

    const createTightP = (text: string, isBold: boolean, isItalic: boolean, isUnderline: boolean, align: string, customSize?: number) => {
        const p = createElement("w:p");
        const pPr = getOrCreate(p, "w:pPr");
        const jc = getOrCreate(pPr, "w:jc");
        jc.setAttributeNS(W_NS, "w:val", align);
        const ind = getOrCreate(pPr, "w:ind");
        ind.setAttributeNS(W_NS, "w:left", "0");
        ind.setAttributeNS(W_NS, "w:right", "0");
        ind.setAttributeNS(W_NS, "w:firstLine", "0");
        ind.removeAttributeNS(W_NS, "hanging");
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "0");
        spacing.setAttributeNS(W_NS, "w:line", "240"); 
        spacing.setAttributeNS(W_NS, "w:lineRule", "auto"); 
        const r = createElement("w:r");
        p.appendChild(r);
        const rPr = getOrCreate(r, "w:rPr");
        const sizeToUse = customSize ? customSize * 2 : options.font?.sizeTable * 2 || 26;
        const sz = getOrCreate(rPr, "w:sz");
        sz.setAttributeNS(W_NS, "w:val", String(sizeToUse)); 
        const szCs = getOrCreate(rPr, "w:szCs");
        szCs.setAttributeNS(W_NS, "w:val", String(sizeToUse));
        if (isBold) {
            const b = getOrCreate(rPr, "w:b");
            b.setAttributeNS(W_NS, "w:val", "true");
        }
        if (isItalic) {
            const i = getOrCreate(rPr, "w:i");
            i.setAttributeNS(W_NS, "w:val", "true");
        }
        if (isUnderline) {
            const u = getOrCreate(rPr, "w:u");
            u.setAttributeNS(W_NS, "w:val", "single");
        }
        const t = createElement("w:t");
        t.textContent = text;
        r.appendChild(t);
        return p;
    };

    const tbl = createElement("w:tbl");
    const tblPr = getOrCreate(tbl, "w:tblPr");
    const tblBorders = getOrCreate(tblPr, "w:tblBorders");
    ["top", "left", "bottom", "right", "insideH", "insideV"].forEach(side => {
        const border = getOrCreate(tblBorders, `w:${side}`);
        border.setAttributeNS(W_NS, "w:val", "none");
    });
    
    const tblLayout = getOrCreate(tblPr, "w:tblLayout");
    tblLayout.setAttributeNS(W_NS, "w:type", "fixed");

    const tblW = getOrCreate(tblPr, "w:tblW");
    tblW.setAttributeNS(W_NS, "w:w", "9350");
    tblW.setAttributeNS(W_NS, "w:type", "dxa"); 

    const isMinutes = (docType && docType.toUpperCase().includes("BIÊN BẢN")) || options.isMinutes === true;
    
    const w1 = isMinutes ? "4675" : "3800";
    const w2 = isMinutes ? "4675" : "5550";

    const tblGrid = getOrCreate(tbl, "w:tblGrid");
    const col1 = createElement("w:gridCol");
    col1.setAttributeNS(W_NS, "w:w", w1);
    tblGrid.appendChild(col1);
    const col2 = createElement("w:gridCol");
    col2.setAttributeNS(W_NS, "w:w", w2);
    tblGrid.appendChild(col2);

    const tr = createElement("w:tr");
    tbl.appendChild(tr);

    const tc1 = createElement("w:tc");
    tr.appendChild(tc1);
    const tc1Pr = getOrCreate(tc1, "w:tcPr");
    const tc1W = getOrCreate(tc1Pr, "w:tcW");
    tc1W.setAttributeNS(W_NS, "w:w", w1);
    tc1W.setAttributeNS(W_NS, "w:type", "dxa");
    
    const tc1Mar = getOrCreate(tc1Pr, "w:tcMar");
    ["top", "bottom", "left", "right"].forEach(side => {
        const mar = getOrCreate(tc1Mar, `w:${side}`);
        mar.setAttributeNS(W_NS, "w:w", "0");
        mar.setAttributeNS(W_NS, "w:type", "dxa");
    });

    const tc2 = createElement("w:tc");
    tr.appendChild(tc2);
    const tc2Pr = getOrCreate(tc2, "w:tcPr");
    const tc2W = getOrCreate(tc2Pr, "w:tcW");
    tc2W.setAttributeNS(W_NS, "w:w", w2);
    tc2W.setAttributeNS(W_NS, "w:type", "dxa");
    
    const tc2Mar = getOrCreate(tc2Pr, "w:tcMar");
    ["top", "bottom", "left", "right"].forEach(side => {
        const mar = getOrCreate(tc2Mar, `w:${side}`);
        mar.setAttributeNS(W_NS, "w:w", "0");
        mar.setAttributeNS(W_NS, "w:type", "dxa");
    });

    const signerTitle = options.signerTitle ? options.signerTitle.trim().toUpperCase() : "";
    const signerName = options.signerName ? options.signerName.trim() : "";
    const presiderName = options.presiderName ? options.presiderName.trim() : "";
    const secretaryName = options.secretaryName ? options.secretaryName.trim() : "";

    if (isMinutes) {
        tc1.appendChild(createTightP("THƯ KÝ", true, false, false, "center", 14));
        tc1.appendChild(createTightP("", false, false, false, "center", 14));
        tc1.appendChild(createTightP("", false, false, false, "center", 14));
        tc1.appendChild(createTightP("", false, false, false, "center", 14));
        if (secretaryName) tc1.appendChild(createTightP(secretaryName, true, false, false, "center", 14));

        tc2.appendChild(createTightP("CHỦ TỌA", true, false, false, "center", 14));
        tc2.appendChild(createTightP("", false, false, false, "center", 14));
        tc2.appendChild(createTightP("", false, false, false, "center", 14));
        tc2.appendChild(createTightP("", false, false, false, "center", 14));
        if (presiderName) tc2.appendChild(createTightP(presiderName, true, false, false, "center", 14));
    } else {
        switch (options.headerType) {
            case HeaderType.PARTY:
                tc1.appendChild(createTightP("Nơi nhận:", false, false, true, "left", 14));
                tc1.appendChild(createTightP("- Đảng ủy xã Ea Kar (b/c),", false, false, false, "left", 12));
                tc1.appendChild(createTightP("- Chi ủy và Lãnh đạo trường,", false, false, false, "left", 12));
                tc1.appendChild(createTightP("- BT Chi Đoàn, TPT Đội,", false, false, false, "left", 12));
                tc1.appendChild(createTightP("- Đảng viên (t/h),", false, false, false, "left", 12));
                tc1.appendChild(createTightP("- Lưu HSCB.", false, false, false, "left", 12));

                tc2.appendChild(createTightP("T/M CHI BỘ", true, false, false, "center", 14));
                tc2.appendChild(createTightP(signerTitle || "BÍ THƯ", true, false, false, "center", 14));
                tc2.appendChild(createTightP("", false, false, false, "center", 14));
                tc2.appendChild(createTightP("", false, false, false, "center", 14));
                tc2.appendChild(createTightP("", false, false, false, "center", 14));
                if (signerName) tc2.appendChild(createTightP(signerName, true, false, false, "center", 14));
                break;
            case HeaderType.DEPARTMENT:
                tc1.appendChild(createTightP("Nơi nhận:", true, true, false, "left", 12));
                tc1.appendChild(createTightP("- Lãnh đạo trường (b/c);", false, false, false, "left", 11));
                tc1.appendChild(createTightP("- Thành viên Tổ (t/h);", false, false, false, "left", 11));
                tc1.appendChild(createTightP("- Lưu HSTCM.", false, false, false, "left", 11));

                tc2.appendChild(createTightP(signerTitle || "TỔ TRƯỞNG", true, false, false, "center", 14));
                tc2.appendChild(createTightP("", false, false, false, "center", 14));
                tc2.appendChild(createTightP("", false, false, false, "center", 14));
                tc2.appendChild(createTightP("", false, false, false, "center", 14));
                if (signerName) tc2.appendChild(createTightP(signerName, true, false, false, "center", 14));
                break;
            case HeaderType.SCHOOL:
            default:
                tc1.appendChild(createTightP("Nơi nhận:", true, true, false, "left", 12));
                tc1.appendChild(createTightP("- Sở Giáo dục và Đào tạo Đắk Lắk (b/c);", false, false, false, "left", 11));
                tc1.appendChild(createTightP("- Phòng Văn hóa – Xã hội Ea Kar (b/c);", false, false, false, "left", 11));
                tc1.appendChild(createTightP("- Cấp ủy chi bộ (b/c);", false, false, false, "left", 11));
                tc1.appendChild(createTightP("- Các tổ chuyên môn, Văn phòng(t/h);", false, false, false, "left", 11));
                tc1.appendChild(createTightP("- Giáo viên, nhân viên (t/h);", false, false, false, "left", 11));
                tc1.appendChild(createTightP("- Lưu VT, EDOC.", false, false, false, "left", 11));

                if (signerTitle === "PHÓ HIỆU TRƯỞNG") {
                    tc2.appendChild(createTightP("KT. HIỆU TRƯỞNG", true, false, false, "center", 14));
                    tc2.appendChild(createTightP("PHÓ HIỆU TRƯỞNG", true, false, false, "center", 14));
                } else {
                    tc2.appendChild(createTightP(signerTitle || "HIỆU TRƯỞNG", true, false, false, "center", 14));
                }
                tc2.appendChild(createTightP("", false, false, false, "center", 14));
                tc2.appendChild(createTightP("", false, false, false, "center", 14));
                tc2.appendChild(createTightP("", false, false, false, "center", 14));
                if (signerName) tc2.appendChild(createTightP(signerName, true, false, false, "center", 14));
                break;
        }
    }
    return tbl;
}