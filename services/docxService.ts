import JSZip from 'jszip';
import { ProcessResult, DocxOptions, HeaderType } from '../types';

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const TWIPS_PER_CM = 567;
const TWIPS_PER_PT = 20;

const DOC_TYPE_KEYWORDS = [
  "NGHỊ QUYẾT", "QUYẾT ĐỊNH", "THÔNG BÁO", "BÁO CÁO", "TỜ TRÌNH", 
  "KẾ HOẠCH", "CHƯƠNG TRÌNH", "CÔNG VĂN", "GIẤY MỜI", "BIÊN BẢN"
];

const DEFAULT_OPTIONS: DocxOptions = {
  headerType: HeaderType.NONE,
  removeNumbering: false,
  margins: { top: 2, bottom: 2, left: 3, right: 1.5 },
  font: { family: "Times New Roman", sizeNormal: 14, sizeTable: 13 },
  paragraph: { lineSpacing: 1.15, after: 6, indent: 1.27 },
  table: { rowHeight: 0.8 }
};

const normalizeSummary = (text: string): string => {
    let summary = text.trim();
    if (!summary) return "";
    summary = summary.replace(/^[:-]\s*/, '').trim();
    summary = summary.replace(/(?:-|–)?\s*tháng\s+(\d{1,2})(?:\/|-)(\d{4})/gi, (match, m, y) => {
        return `tháng ${m.padStart(2, '0')} năm ${y}`;
    });
    const currentYear = new Date().getFullYear();
    summary = summary.replace(/(?:-|–)?\s*tháng\s+(\d{1,2})(?!\s*năm|\/|-)/gi, (match, m) => {
        return `tháng ${m.padStart(2, '0')} năm ${currentYear}`;
    });
    if (summary.length > 0) {
        summary = summary.charAt(0).toUpperCase() + summary.slice(1).toLowerCase();
    }
    return summary;
};

// "CẢNH SÁT KIỂM DUYỆT" - Khóa cứng trật tự XML chuẩn MS Word
const enforceSchema = (doc: Document) => {
    const orders: Record<string, string[]> = {
        "w:pPr": ["w:pStyle", "w:spacing", "w:ind", "w:jc", "w:rPr"],
        "w:rPr": ["w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:color", "w:sz", "w:szCs", "w:u"],
        "w:tblPr": ["w:tblW", "w:jc", "w:tblBorders", "w:tblLayout"],
        "w:tcPr": ["w:tcW", "w:tcBorders", "w:tcMar", "w:vAlign"],
        "w:r": ["w:rPr", "w:t"] 
    };
    Object.keys(orders).forEach(tagName => {
        const localName = tagName.split(":")[1];
        const elements = Array.from(doc.getElementsByTagNameNS(W_NS, localName));
        elements.forEach(el => {
            const order = orders[tagName];
            const children = Array.from(el.childNodes);
            children.sort((a, b) => {
                const nameA = a.nodeName;
                const nameB = b.nodeName;
                const indexA = order.indexOf(nameA);
                const indexB = order.indexOf(nameB);
                if (indexA === -1 && indexB === -1) return 0;
                if (indexA === -1) return 1;
                if (indexB === -1) return -1;
                return indexA - indexB;
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

    // --- BƯỚC 1: DỌN DẸP RÁC ---
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

    const body = doc.getElementsByTagNameNS(W_NS, "body")[0];
    if (body) {
      const sectPr = getOrCreate(body, "w:sectPr");
      const pgSz = getOrCreate(sectPr, "w:pgSz");
      pgSz.setAttributeNS(W_NS, "w:w", String(Math.round(21 * TWIPS_PER_CM)));
      pgSz.setAttributeNS(W_NS, "w:h", String(Math.round(29.7 * TWIPS_PER_CM)));
      pgSz.setAttributeNS(W_NS, "w:orient", "portrait");
      const pgMar = getOrCreate(sectPr, "w:pgMar");
      pgMar.setAttributeNS(W_NS, "w:top", String(Math.round(options.margins.top * TWIPS_PER_CM)));
      pgMar.setAttributeNS(W_NS, "w:bottom", String(Math.round(options.margins.bottom * TWIPS_PER_CM)));
      pgMar.setAttributeNS(W_NS, "w:left", String(Math.round(options.margins.left * TWIPS_PER_CM)));
      pgMar.setAttributeNS(W_NS, "w:right", String(Math.round(options.margins.right * TWIPS_PER_CM)));
    }

    // --- BƯỚC 2: KHÓA MỤC TIÊU TIÊU ĐỀ & TRÍCH YẾU ---
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

    // ĐƯỜNG KẺ MỎNG (BẢO VỆ HEIGHT 0.1CM)
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

    // 5 DẤU GẠCH ĐẢNG
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
            
            let summaryP: Element | null = null;
            const originalUpper = text.toUpperCase();
            const keywordIndex = originalUpper.indexOf(matchedKeyword);
            const remainingText = text.slice(keywordIndex + matchedKeyword.length).trim();
            
            rebuildParagraph(p, matchedKeyword, true, "28", true); 

            if (remainingText.length > 3) {
                const newP = doc.createElementNS(W_NS, "w:p");
                if (p.nextSibling) p.parentNode?.insertBefore(newP, p.nextSibling);
                else p.parentNode?.appendChild(newP);
                summaryP = newP;
                rebuildParagraph(summaryP, normalizeSummary(remainingText), true, String(options.font.sizeNormal * 2), false); 
                abstractElements.add(newP);
            } else {
                for (let step = 1; step <= 3; step++) {
                    if (i + step < paragraphs.length) {
                        const tempP = paragraphs[i + step];
                        if (isTableParagraph(tempP)) break;
                        const tempText = tempP.textContent?.trim() || "";
                        if (tempText.length > 0) {
                            summaryP = tempP;
                            rebuildParagraph(summaryP, normalizeSummary(tempText), true, String(options.font.sizeNormal * 2), false); 
                            abstractElements.add(summaryP);
                            for(let k = 1; k < step; k++){
                                abstractElements.add(paragraphs[i+k]);
                            }
                            break;
                        }
                    }
                }
            }

            const targetNode = summaryP || p;
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

    // --- BƯỚC 3: CĂN CHỈNH NỘI DUNG VĂN BẢN BÊN DƯỚI VÀ NHẬN DIỆN "QUYẾT ĐỊNH" ---
    const finalParagraphs = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
    for (const p of finalParagraphs) {
      if (docTypeElements.has(p) || abstractElements.has(p) || protectedElements.has(p)) continue; 
      
      const isTable = isTableParagraph(p);
      const pPr = getOrCreate(p, "w:pPr");
      const pText = p.textContent?.trim() || "";
      const upperText = pText.toUpperCase();

      let isDecisionSpecialLine = false;
      if (detectedDocType === "QUYẾT ĐỊNH" && !isTable && pText.length > 0) {
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

      if (!isTable) {
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
      } else {
        const jc = getOrCreate(pPr, "w:jc");
        jc.setAttributeNS(W_NS, "w:val", "center");
        const ind = getOrCreate(pPr, "w:ind");
        ind.setAttributeNS(W_NS, "w:left", "0");
        ind.setAttributeNS(W_NS, "w:right", "0");
        ind.setAttributeNS(W_NS, "w:firstLine", "0");
        ind.removeAttributeNS(W_NS, "hanging");
      }
      
      const targetSize = isTable ? (options.font.sizeTable * 2) : (options.font.sizeNormal * 2);
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

        const rows = Array.from(tbl.getElementsByTagNameNS(W_NS, "tr"));
        for (const tr of rows) {
            const trPr = getOrCreate(tr, "w:trPr");
            let trHeight = trPr.getElementsByTagNameNS(W_NS, "trHeight")[0];
            if (!trHeight) {
                trHeight = doc.createElementNS(W_NS, "w:trHeight");
                trHeight.setAttributeNS(W_NS, "w:val", String(Math.round(options.table.rowHeight * TWIPS_PER_CM)));
                trHeight.setAttributeNS(W_NS, "w:hRule", "atLeast");
                trPr.appendChild(trHeight);
            }

            const cells = Array.from(tr.getElementsByTagNameNS(W_NS, "tc"));
            for (const tc of cells) {
                const tcPr = getOrCreate(tc, "w:tcPr");
                const vAlign = getOrCreate(tcPr, "w:vAlign");
                vAlign.setAttributeNS(W_NS, "w:val", "center");
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

    // --- BƯỚC CẬP NHẬT MỚI: THANH TRỪNG THEME ĐỂ ÉP FONT TIMES NEW ROMAN ---
    const allRPrs = Array.from(doc.getElementsByTagNameNS(W_NS, "rPr"));
    for (const rPr of allRPrs) {
        const rFonts = getOrCreate(rPr, "w:rFonts");
        rFonts.setAttributeNS(W_NS, "w:ascii", options.font.family);
        rFonts.setAttributeNS(W_NS, "w:hAnsi", options.font.family);
        rFonts.setAttributeNS(W_NS, "w:cs", options.font.family);
        rFonts.setAttributeNS(W_NS, "w:eastAsia", options.font.family);
        
        // Diệt cỏ tận gốc các theme mặc định của Word gây lỗi hiển thị font
        ["asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"].forEach(theme => {
            rFonts.removeAttributeNS(W_NS, theme);
            rFonts.removeAttribute(`w:${theme}`); // Gỡ bỏ an toàn trên mọi parser
        });
    }

    enforceSchema(doc);

    // --- STEP 7: AUTO PAGE NUMBERING ---
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
                <w:t>2</w:t>
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

    const serializer = new XMLSerializer();
    const newDocXml = serializer.serializeToString(doc);
    zip.file(docXmlPath, newDocXml);
    const generatedBlob = await zip.generateAsync({ type: "blob" });

    return { success: true, blob: generatedBlob, fileName: `formatted_${file.name}`, logs };
  } catch (error) {
    return { success: false, error: error instanceof Error ? error.message : "Unknown error", logs };
  }
};

// ====================================================
// ============= TEMPLATE CREATORS ====================
// ====================================================

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
        
        const u = getOrCreate(rPr, "w:u");
        u.setAttributeNS(W_NS, "w:val", "single"); 

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
    col1.setAttributeNS(W_NS, "w:w", "4000"); 
    tblGrid.appendChild(col1);
    const col2 = createElement("w:gridCol");
    col2.setAttributeNS(W_NS, "w:w", "5350"); 
    tblGrid.appendChild(col2);

    const tr = createElement("w:tr");
    tbl.appendChild(tr);

    const tc1 = createElement("w:tc");
    tr.appendChild(tc1);
    const tc1Pr = getOrCreate(tc1, "w:tcPr");
    const tc1W = getOrCreate(tc1Pr, "w:tcW");
    tc1W.setAttributeNS(W_NS, "w:w", "4000");
    tc1W.setAttributeNS(W_NS, "w:type", "dxa");
    
    const tc2 = createElement("w:tc");
    tr.appendChild(tc2);
    const tc2Pr = getOrCreate(tc2, "w:tcPr");
    const tc2W = getOrCreate(tc2Pr, "w:tcW");
    tc2W.setAttributeNS(W_NS, "w:w", "5350");
    tc2W.setAttributeNS(W_NS, "w:type", "dxa");

    const docDate = options.documentDate ? new Date(options.documentDate) : new Date();
    const day = String(docDate.getDate()).padStart(2, '0');
    const month = String(docDate.getMonth() + 1).padStart(2, '0');
    const year = docDate.getFullYear();
    const currentDateStr = `Ea Kar, ngày ${day} tháng ${month} năm ${year}`;

    switch (options.headerType) {
        case HeaderType.PARTY:
            tc1.appendChild(createStyledP("ĐẢNG BỘ XÃ EA KAR", false, false));
            tc1.appendChild(createStyledP("CHI BỘ TRƯỜNG THCS CHU VĂN AN", true, false));
            tc1.appendChild(createStyledP("*", false, false)); 
            tc1.appendChild(createStyledP("Số: ... - .../CB", false, false));

            tc2.appendChild(createMottoP("ĐẢNG CỘNG SẢN VIỆT NAM", true, 13)); 
            tc2.appendChild(createStyledP("", false, false));
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
            tc2.appendChild(createStyledP("", false, false)); 
            tc2.appendChild(createStyledP(currentDateStr, false, true, 14));
            break;

        case HeaderType.SCHOOL:
        default:
            tc1.appendChild(createStyledP("UBND XÃ EA KAR", false, false));
            tc1.appendChild(createStyledP("TRƯỜNG THCS CHU VĂN AN", true, false));
            appendSafeTable(tc1, createShortLineTable());
            tc1.appendChild(createStyledP("", false, false));
            tc1.appendChild(createStyledP("Số: ... /...", false, false));

            tc2.appendChild(createStyledP("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", true, false, 13));
            tc2.appendChild(createMottoP("Độc lập - Tự do - Hạnh phúc", true, 13)); 
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

    const isMinutes = docType && docType.toUpperCase().includes("BIÊN BẢN");
    const w1 = isMinutes ? "4675" : "4000";
    const w2 = isMinutes ? "4675" : "5350";

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

    const tc2 = createElement("w:tc");
    tr.appendChild(tc2);
    const tc2Pr = getOrCreate(tc2, "w:tcPr");
    const tc2W = getOrCreate(tc2Pr, "w:tcW");
    tc2W.setAttributeNS(W_NS, "w:w", w2);
    tc2W.setAttributeNS(W_NS, "w:type", "dxa");

    const signerTitle = options.signerTitle ? options.signerTitle.trim().toUpperCase() : "";
    const signerName = options.signerName ? options.signerName.trim() : "";

    if (isMinutes) {
        tc1.appendChild(createTightP("THƯ KÝ", true, false, false, "center", 14));
        tc2.appendChild(createTightP("CHỦ TRÌ", true, false, false, "center", 14));
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