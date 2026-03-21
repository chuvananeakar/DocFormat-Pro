import JSZip from 'jszip';
import { ProcessResult, DocxOptions, HeaderType } from '../types';

// Constants for XML Namespaces and Units
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const TWIPS_PER_CM = 567;
const TWIPS_PER_PT = 20;

const HEADER_SETUP = {
  lineSpacing: 240, // Single spacing (1.0 lines)
  spacingAfter: 6 * TWIPS_PER_PT,
};

const DOC_TYPE_KEYWORDS = [
  "NGHỊ QUYẾT", 
  "QUYẾT ĐỊNH", 
  "THÔNG BÁO", 
  "BÁO CÁO", 
  "TỜ TRÌNH", 
  "KẾ HOẠCH", 
  "CHƯƠNG TRÌNH", 
  "CÔNG VĂN", 
  "GIẤY MỜI", 
  "BIÊN BẢN"
];

// Fallback defaults if options are missing (though App should provide them)
const DEFAULT_OPTIONS: DocxOptions = {
  headerType: HeaderType.NONE,
  removeNumbering: false,
  margins: { top: 2, bottom: 2, left: 3, right: 1.5 },
  font: { family: "Times New Roman", sizeNormal: 14, sizeTable: 13 },
  paragraph: { lineSpacing: 1.15, after: 6, indent: 1.27 },
  table: { rowHeight: 0.8 }
};

export const processDocx = async (file: File, options: DocxOptions = DEFAULT_OPTIONS): Promise<ProcessResult> => {
  const logs: string[] = [];
  try {
    logs.push(`Loading file: ${file.name}`);
    logs.push(`Applying settings: Margins(${options.margins.top}, ${options.margins.bottom}, ${options.margins.left}, ${options.margins.right}), Font(${options.font.sizeNormal}pt)`);
    
    const arrayBuffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(arrayBuffer);

    const docXmlPath = "word/document.xml";
    const docXmlContent = await zip.file(docXmlPath)?.async("string");

    if (!docXmlContent) {
      throw new Error("Invalid DOCX: missing word/document.xml");
    }

    logs.push("Parsing XML structure...");
    const parser = new DOMParser();
    const doc = parser.parseFromString(docXmlContent, "application/xml");

    // Helper to creating elements with namespace
    const createElement = (tagName: string) => doc.createElementNS(W_NS, tagName);

    // Helper to get or create a child element
    const getOrCreate = (parent: Element, tagName: string): Element => {
      let child = parent.getElementsByTagNameNS(W_NS, tagName)[0];
      if (!child) {
        child = createElement(tagName);
        parent.appendChild(child);
      }
      return child;
    };

    // --- DEDICATED FUNCTION: Deep Formatting for Fonts ---
    // Forces the font family on every single run in the document
    const forceFontApplication = (document: Document, fontFamily: string) => {
        const allRuns = Array.from(document.getElementsByTagNameNS(W_NS, "r"));
        let count = 0;
        
        for (const r of allRuns) {
            const rPr = getOrCreate(r, "rPr");
            const rFonts = getOrCreate(rPr, "rFonts");
            
            // Forcefully set all font types to the target font
            // This overrides any existing direct formatting
            rFonts.setAttributeNS(W_NS, "w:ascii", fontFamily);
            rFonts.setAttributeNS(W_NS, "w:hAnsi", fontFamily);
            rFonts.setAttributeNS(W_NS, "w:cs", fontFamily);
            
            // CRITICAL: Set EastAsia font for Vietnamese/Unicode support
            rFonts.setAttributeNS(W_NS, "w:eastAsia", fontFamily);
            
            count++;
        }
        logs.push(`Deep Formatting: Enforced '${fontFamily}' on ${count} text runs.`);
    };

    // Helper to check if a paragraph is inside a table
    const isTableParagraph = (p: Element): boolean => {
      let parent = p.parentNode;
      while(parent) {
        if (parent.nodeName === 'w:tbl' || parent.nodeName === 'tbl') {
          return true;
        }
        parent = parent.parentNode;
      }
      return false;
    };

    // --- STEP 1: Data Cleaning ---
    logs.push("Cleaning data (removing empty paragraphs & trimming whitespace)...");
    
    const paragraphsForCleaning = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
    let removedCount = 0;

    for (const p of paragraphsForCleaning) {
        const textNodes = Array.from(p.getElementsByTagNameNS(W_NS, "t"));

        // A. Trim Whitespace
        if (textNodes.length > 0) {
            const firstNode = textNodes[0];
            if (firstNode.textContent) firstNode.textContent = firstNode.textContent.trimStart();
            const lastNode = textNodes[textNodes.length - 1];
            if (lastNode.textContent) lastNode.textContent = lastNode.textContent.trimEnd();
        }

        // B. Remove Empty Paragraphs
        const fullText = textNodes.map(n => n.textContent || "").join("");
        const drawings = p.getElementsByTagNameNS(W_NS, "drawing");
        const picts = p.getElementsByTagNameNS(W_NS, "pict");
        const objects = p.getElementsByTagNameNS(W_NS, "object");
        const breaks = p.getElementsByTagNameNS(W_NS, "br"); 
        
        const hasContent = drawings.length > 0 || picts.length > 0 || objects.length > 0 || breaks.length > 0;

        if (!hasContent && fullText.length === 0) {
            const parent = p.parentNode as Element;
            if (parent && (parent.localName === "body" || parent.nodeName.indexOf("body") !== -1)) {
                parent.removeChild(p);
                removedCount++;
            }
        }
    }
    if (removedCount > 0) logs.push(`Removed ${removedCount} empty paragraph(s).`);

    // --- STEP 1.5: Remove Bullets & Numbering (If Enabled) ---
    if (options.removeNumbering) {
        logs.push("Removing automatic numbering and bullets...");
        const allParagraphs = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
        let numberingRemovedCount = 0;

        for (const p of allParagraphs) {
            const pPr = p.getElementsByTagNameNS(W_NS, "pPr")[0];
            if (pPr) {
                // Remove <w:numPr> to kill auto-numbering
                const numPr = pPr.getElementsByTagNameNS(W_NS, "numPr")[0];
                if (numPr) {
                    pPr.removeChild(numPr);
                    numberingRemovedCount++;
                }

                // Force Style to Normal
                const pStyle = getOrCreate(pPr, "pStyle");
                pStyle.setAttributeNS(W_NS, "w:val", "Normal");
            }

            // Safety Net: Strip Manual Bullets/Numbers from text content
            // Regex matches starts of lines like "1. ", "- ", "* ", "• "
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
        if (numberingRemovedCount > 0) logs.push(`Removed automatic numbering from ${numberingRemovedCount} paragraphs.`);
    }

    // --- STEP 2: Page Setup ---
    logs.push("Applying Page Setup...");
    const body = doc.getElementsByTagNameNS(W_NS, "body")[0];
    if (body) {
      const sectPr = getOrCreate(body, "sectPr");
      const pgSz = getOrCreate(sectPr, "pgSz");
      // Fixed A4 Size for now
      pgSz.setAttributeNS(W_NS, "w:w", String(Math.round(21 * TWIPS_PER_CM)));
      pgSz.setAttributeNS(W_NS, "w:h", String(Math.round(29.7 * TWIPS_PER_CM)));
      pgSz.setAttributeNS(W_NS, "w:orient", "portrait");
      
      const pgMar = getOrCreate(sectPr, "pgMar");
      pgMar.setAttributeNS(W_NS, "w:top", String(Math.round(options.margins.top * TWIPS_PER_CM)));
      pgMar.setAttributeNS(W_NS, "w:bottom", String(Math.round(options.margins.bottom * TWIPS_PER_CM)));
      pgMar.setAttributeNS(W_NS, "w:left", String(Math.round(options.margins.left * TWIPS_PER_CM)));
      pgMar.setAttributeNS(W_NS, "w:right", String(Math.round(options.margins.right * TWIPS_PER_CM)));
    }

    // --- STEP 3: Identification & Formatting ---
    logs.push("Formatting Paragraphs...");
    const paragraphs = Array.from(doc.getElementsByTagNameNS(W_NS, "p"));
    
    // Identify Special Sections (Limit to first 20 paragraphs)
    const docTypeIndices = new Set<number>();
    const abstractIndices = new Set<number>();

    logs.push("Scanning for Document Type and Abstract (Decree 30/2020/NĐ-CP standards)...");
    
    // Scan only first 20 paragraphs for efficiency and accuracy
    const limit = Math.min(paragraphs.length, 20);
    
    for (let i = 0; i < limit; i++) {
        const p = paragraphs[i];
        if (isTableParagraph(p)) continue;

        const text = p.textContent?.trim() || "";
        if (!text) continue;
        
        const upperText = text.toUpperCase();
        // Check if paragraph contains one of the keywords (case-insensitive check)
        const match = DOC_TYPE_KEYWORDS.some(k => upperText.includes(k));

        if (match) {
            docTypeIndices.add(i);
            
            // Transform text to UPPERCASE immediately
            const textNodes = Array.from(p.getElementsByTagNameNS(W_NS, "t"));
            for (const tNode of textNodes) {
                if (tNode.textContent) {
                    tNode.textContent = tNode.textContent.toUpperCase();
                }
            }

            // The next paragraph is the Abstract
            if (i + 1 < paragraphs.length) {
                const nextP = paragraphs[i + 1];
                if (!isTableParagraph(nextP)) {
                    abstractIndices.add(i + 1);
                    logs.push(`Found Document Type at #${i+1} and Abstract at #${i + 2}`);
                }
            }
            // Stop after finding the first valid Document Type header to avoid false positives
            break; 
        }
    }

    // Apply Formatting Loop
    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const isTable = isTableParagraph(p);
      const pPr = getOrCreate(p, "pPr");

      const isDocType = docTypeIndices.has(i);
      const isAbstract = abstractIndices.has(i);

      if (isDocType) {
        // Document Type Styling: Center, Bold, Upper(done), Spacing Before 12pt, After 0, Single Line
        const jc = getOrCreate(pPr, "jc");
        jc.setAttributeNS(W_NS, "w:val", "center");

        const spacing = getOrCreate(pPr, "spacing");
        spacing.setAttributeNS(W_NS, "w:before", "240"); // 12pt
        spacing.setAttributeNS(W_NS, "w:after", "0");
        spacing.setAttributeNS(W_NS, "w:line", "240"); // Single
        spacing.setAttributeNS(W_NS, "w:lineRule", "auto");

        // Remove Indentation
        const ind = getOrCreate(pPr, "ind");
        ind.setAttributeNS(W_NS, "w:left", "0");
        ind.setAttributeNS(W_NS, "w:right", "0");
        ind.setAttributeNS(W_NS, "w:firstLine", "0");
        ind.removeAttributeNS(W_NS, "hanging");

        // Force Bold on all runs
        const runs = Array.from(p.getElementsByTagNameNS(W_NS, "r"));
        for (const r of runs) {
             const rPr = getOrCreate(r, "rPr");
             const b = getOrCreate(rPr, "b");
             b.setAttributeNS(W_NS, "w:val", "true"); // Force Bold
        }

      } else if (isAbstract) {
        // Abstract Styling: Center, Bold, Spacing Before 0, After 12pt, Single Line
        const jc = getOrCreate(pPr, "jc");
        jc.setAttributeNS(W_NS, "w:val", "center");

        const spacing = getOrCreate(pPr, "spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "240"); // 12pt
        spacing.setAttributeNS(W_NS, "w:line", "240"); // Single
        spacing.setAttributeNS(W_NS, "w:lineRule", "auto");

        // Remove Indentation
        const ind = getOrCreate(pPr, "ind");
        ind.setAttributeNS(W_NS, "w:left", "0");
        ind.setAttributeNS(W_NS, "w:right", "0");
        ind.setAttributeNS(W_NS, "w:firstLine", "0");
        ind.removeAttributeNS(W_NS, "hanging");

        // Force Bold on all runs
        const runs = Array.from(p.getElementsByTagNameNS(W_NS, "r"));
        for (const r of runs) {
             const rPr = getOrCreate(r, "rPr");
             const b = getOrCreate(rPr, "b");
             b.setAttributeNS(W_NS, "w:val", "true"); // Force Bold
        }

      } else if (!isTable) {
        // Normal Body
        const jc = getOrCreate(pPr, "jc");
        jc.setAttributeNS(W_NS, "w:val", "both");

        const spacing = getOrCreate(pPr, "spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", String(Math.round(options.paragraph.after * TWIPS_PER_PT)));
        // Line spacing calculation: e.g. 1.15 * 240 = 276
        spacing.setAttributeNS(W_NS, "w:line", String(Math.round(options.paragraph.lineSpacing * 240))); 
        spacing.setAttributeNS(W_NS, "w:lineRule", "auto");

        const ind = getOrCreate(pPr, "ind");
        ind.setAttributeNS(W_NS, "w:left", "0");
        ind.setAttributeNS(W_NS, "w:right", "0");
        ind.setAttributeNS(W_NS, "w:firstLine", String(Math.round(options.paragraph.indent * TWIPS_PER_CM)));
      } else {
        // Table Paragraph
        const jc = getOrCreate(pPr, "jc");
        jc.setAttributeNS(W_NS, "w:val", "center");

        const ind = getOrCreate(pPr, "ind");
        ind.setAttributeNS(W_NS, "w:left", "0");
        ind.setAttributeNS(W_NS, "w:right", "0");
        ind.setAttributeNS(W_NS, "w:firstLine", "0");
        ind.removeAttributeNS(W_NS, "hanging");
      }

      // Font Normalization - Size Only
      // Note: Font Family is now handled by forceFontApplication in Step 6
      const targetSize = isTable ? (options.font.sizeTable * 2) : (options.font.sizeNormal * 2);
      
      const runs = Array.from(p.getElementsByTagNameNS(W_NS, "r"));
      for (const r of runs) {
          const rPr = getOrCreate(r, "rPr");
          
          const sz = getOrCreate(rPr, "sz");
          sz.setAttributeNS(W_NS, "w:val", String(targetSize));
          const szCs = getOrCreate(rPr, "szCs");
          szCs.setAttributeNS(W_NS, "w:val", String(targetSize));
      }
    }

    // --- STEP 4: Table Row Properties ---
    const tables = Array.from(doc.getElementsByTagNameNS(W_NS, "tbl"));
    for (const tbl of tables) {
        const rows = Array.from(tbl.getElementsByTagNameNS(W_NS, "tr"));
        for (const tr of rows) {
            const trPr = getOrCreate(tr, "trPr");
            const trHeight = getOrCreate(trPr, "trHeight");
            trHeight.setAttributeNS(W_NS, "w:val", String(Math.round(options.table.rowHeight * TWIPS_PER_CM)));
            trHeight.setAttributeNS(W_NS, "w:hRule", "atLeast");

            const cells = Array.from(tr.getElementsByTagNameNS(W_NS, "tc"));
            for (const tc of cells) {
                const tcPr = getOrCreate(tc, "tcPr");
                const vAlign = getOrCreate(tcPr, "vAlign");
                vAlign.setAttributeNS(W_NS, "w:val", "center");
            }
        }
    }
    
    // --- STEP 5: Insert Header Template (If Selected) ---
    if (options.headerType !== HeaderType.NONE && body) {
        logs.push(`Inserting Header Template: ${options.headerType}...`);
        const headerTable = createHeaderTemplate(doc, options);
        if (body.firstChild) {
            body.insertBefore(headerTable, body.firstChild);
        } else {
            body.appendChild(headerTable);
        }
    }

    // --- STEP 6: Deep Font Formatting (Force Times New Roman) ---
    // Executed last to ensure everything, including inserted templates, is standardized.
    logs.push("Step 6: Performing Deep Font Formatting (Force Times New Roman)...");
    forceFontApplication(doc, options.font.family);

    logs.push("Rebuilding DOCX file...");
    const serializer = new XMLSerializer();
    const newDocXml = serializer.serializeToString(doc);
    
    zip.file(docXmlPath, newDocXml);

    const generatedBlob = await zip.generateAsync({ type: "blob" });
    logs.push("Done!");

    return {
      success: true,
      blob: generatedBlob,
      fileName: `formatted_${file.name}`,
      logs
    };

  } catch (error) {
    console.error(error);
    return {
      success: false,
      error: error instanceof Error ? error.message : "Unknown error",
      logs
    };
  }
};

// Helper function to create the Standard Header Table
const createHeaderTemplate = (doc: Document, options: DocxOptions): Element => {
    const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const createElement = (tagName: string) => doc.createElementNS(W_NS, tagName);
    const getOrCreate = (parent: Element, tagName: string): Element => {
      let child = parent.getElementsByTagNameNS(W_NS, tagName)[0];
      if (!child) {
        child = createElement(tagName);
        parent.appendChild(child);
      }
      return child;
    };

    // Helper to create a paragraph with specific text and styling
    const createStyledP = (text: string, isBold: boolean, isItalic: boolean): Element => {
        const p = createElement("w:p");
        const pPr = getOrCreate(p, "w:pPr");
        
        // Center Align
        const jc = getOrCreate(pPr, "w:jc");
        jc.setAttributeNS(W_NS, "w:val", "center");
        
        // Remove spacing
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "0");
        spacing.setAttributeNS(W_NS, "w:line", "240"); // Single spacing
        spacing.setAttributeNS(W_NS, "w:lineRule", "auto");

        const r = createElement("w:r");
        p.appendChild(r);

        const rPr = getOrCreate(r, "w:rPr");
        
        // Note: Font family is now handled globally by forceFontApplication
        
        // Size (Using Table Size ~13pt usually)
        const sz = getOrCreate(rPr, "w:sz");
        sz.setAttributeNS(W_NS, "w:val", String(options.font.sizeTable * 2));
        const szCs = getOrCreate(rPr, "w:szCs");
        szCs.setAttributeNS(W_NS, "w:val", String(options.font.sizeTable * 2));

        if (isBold) {
            const b = createElement("w:b");
            rPr.appendChild(b);
        }
        if (isItalic) {
            const i = createElement("w:i");
            rPr.appendChild(i);
        }

        const t = createElement("w:t");
        t.textContent = text;
        r.appendChild(t);

        return p;
    };

    // Helper to create a shrink-to-fit nested table with a thin 1/4pt bottom border
    const createExactLineTable = (text: string, isBold: boolean, isItalic: boolean): Element => {
        const tbl = createElement("w:tbl");
        const tblPr = getOrCreate(tbl, "w:tblPr");
        
        const jcTbl = getOrCreate(tblPr, "w:jc");
        jcTbl.setAttributeNS(W_NS, "w:val", "center");
        
        // CRITICAL: AutoFit forces the table to shrink to exact text width
        const tblLayout = getOrCreate(tblPr, "w:tblLayout");
        tblLayout.setAttributeNS(W_NS, "w:type", "autofit");

        const tr = createElement("w:tr");
        tbl.appendChild(tr);

        const tc = createElement("w:tc");
        tr.appendChild(tc);
        const tcPr = getOrCreate(tc, "w:tcPr");

        const tcW = getOrCreate(tcPr, "w:tcW");
        tcW.setAttributeNS(W_NS, "w:w", "0");
        tcW.setAttributeNS(W_NS, "w:type", "auto");

        // Thin 1/4pt Bottom border
        const tcBorders = getOrCreate(tcPr, "w:tcBorders");
        const bottom = getOrCreate(tcBorders, "w:bottom");
        bottom.setAttributeNS(W_NS, "w:val", "single");
        bottom.setAttributeNS(W_NS, "w:sz", "2"); 
        bottom.setAttributeNS(W_NS, "w:space", "0");
        bottom.setAttributeNS(W_NS, "w:color", "000000");

        // Remove cell margins so border hugs the text
        const tcMar = getOrCreate(tcPr, "w:tcMar");
        ["top", "bottom", "left", "right"].forEach(side => {
            const mar = getOrCreate(tcMar, `w:${side}`);
            mar.setAttributeNS(W_NS, "w:w", "0");
            mar.setAttributeNS(W_NS, "w:type", "dxa");
        });

        const p = createElement("w:p");
        tc.appendChild(p);
        const pPr = getOrCreate(p, "w:pPr");
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "0");

        const r = createElement("w:r");
        p.appendChild(r);
        const rPr = getOrCreate(r, "w:rPr");
        if (isBold) rPr.appendChild(createElement("w:b"));
        if (isItalic) rPr.appendChild(createElement("w:i"));

        const t = createElement("w:t");
        t.textContent = text;
        r.appendChild(t);

        return tbl;
    };

    const tbl = createElement("w:tbl");
    const tblPr = getOrCreate(tbl, "w:tblPr");
    
    // No Borders
    const tblBorders = getOrCreate(tblPr, "w:tblBorders");
    const sides = ["top", "left", "bottom", "right", "insideH", "insideV"];
    sides.forEach(side => {
        const border = getOrCreate(tblBorders, `w:${side}`);
        border.setAttributeNS(W_NS, "w:val", "none");
    });
    
    // Width 100% (approx)
    const tblW = getOrCreate(tblPr, "w:tblW");
    tblW.setAttributeNS(W_NS, "w:w", "5000"); // Auto
    tblW.setAttributeNS(W_NS, "w:type", "pct");

    // Grid (Columns) - Approx 40% (6cm) and 60% (9cm)
    const tblGrid = getOrCreate(tbl, "w:tblGrid");
    const col1 = createElement("w:gridCol");
    col1.setAttributeNS(W_NS, "w:w", "3600"); // ~6.35cm
    tblGrid.appendChild(col1);
    const col2 = createElement("w:gridCol");
    col2.setAttributeNS(W_NS, "w:w", "5400"); // ~9.5cm
    tblGrid.appendChild(col2);

    const tr = createElement("w:tr");
    tbl.appendChild(tr);

    // --- Left Cell (Agency Info) ---
    const tc1 = createElement("w:tc");
    tr.appendChild(tc1);
    const tc1Pr = getOrCreate(tc1, "w:tcPr");
    const tc1W = getOrCreate(tc1Pr, "w:tcW");
    tc1W.setAttributeNS(W_NS, "w:w", "3600");
    tc1W.setAttributeNS(W_NS, "w:type", "dxa");
    
    // --- Right Cell (Motto) ---
    const tc2 = createElement("w:tc");
    tr.appendChild(tc2);
    const tc2Pr = getOrCreate(tc2, "w:tcPr");
    const tc2W = getOrCreate(tc2Pr, "w:tcW");
    tc2W.setAttributeNS(W_NS, "w:w", "5400");
    tc2W.setAttributeNS(W_NS, "w:type", "dxa");

    const today = new Date();
    const day = String(today.getDate()).padStart(2, '0');
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const year = today.getFullYear();
    const currentDateStr = `Ea Kar, ngày ${day} tháng ${month} năm ${year}`;

    switch (options.headerType) {
        case HeaderType.PARTY:
            tc1.appendChild(createStyledP("ĐẢNG BỘ XÃ EA KAR", false, false));
            tc1.appendChild(createStyledP("CHI BỘ TRƯỜNG THCS CHU VĂN AN", true, false));
            tc1.appendChild(createStyledP("*", false, false));
            tc1.appendChild(createStyledP("Số: ... - .../CB", false, false));

            // Nested table for thin, text-length line
            tc2.appendChild(createExactLineTable("ĐẢNG CỘNG SẢN VIỆT NAM", true, false)); 
            // Add 1 empty single line
            tc2.appendChild(createStyledP("", false, false));
            // Updated date text
            tc2.appendChild(createStyledP(currentDateStr, false, true));
            break;

        case HeaderType.DEPARTMENT:
            const deptName = options.departmentName || "TỔ TOÁN - TIN";
            tc1.appendChild(createStyledP("TRƯỜNG THCS CHU VĂN AN", false, false));
            tc1.appendChild(createStyledP(deptName, true, false));
            tc1.appendChild(createStyledP("-------", false, false)); // Department uses static dashed line
            tc1.appendChild(createStyledP("Số: ... /...", false, false));

            tc2.appendChild(createStyledP("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", true, false));
            // Nested table for thin, text-length line
            tc2.appendChild(createExactLineTable("Độc lập - Tự do - Hạnh phúc", true, false));
            // Add 1 empty single line
            tc2.appendChild(createStyledP("", false, false));
            // Updated date text
            tc2.appendChild(createStyledP(currentDateStr, false, true));
            break;

        case HeaderType.SCHOOL:
        default:
            tc1.appendChild(createStyledP("UBND XÃ EA KAR", false, false));
            tc1.appendChild(createStyledP("TRƯỜNG THCS CHU VĂN AN", true, false));
            tc1.appendChild(createStyledP("-------", false, false));
            tc1.appendChild(createStyledP("Số: ... /...", false, false));

            tc2.appendChild(createStyledP("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", true, false));
            // Nested table for thin, text-length line
            tc2.appendChild(createExactLineTable("Độc lập - Tự do - Hạnh phúc", true, false));
            // Add 1 empty single line
            tc2.appendChild(createStyledP("", false, false));
            // Updated date text
            tc2.appendChild(createStyledP(currentDateStr, false, true));
            break;
    }

    return tbl;
};