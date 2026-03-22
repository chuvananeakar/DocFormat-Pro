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

// Helper to parse and normalize bad document titles
const normalizeTitleText = (rawText: string): { docType: string, summary: string } | null => {
    // Only match if it starts with the exact keywords
    const regex = new RegExp(`^\\s*(${DOC_TYPE_KEYWORDS.join('|')})\\s*(.*)$`, 'i');
    const match = rawText.match(regex);

    if (!match) return null;

    const docType = match[1].toUpperCase();
    let summary = match[2].trim();

    if (!summary || summary.length < 3) return null; // Already separated or too short

    // Clean up leading characters
    summary = summary.replace(/^[:-]\s*/, '').trim();

    // Fix date formats
    summary = summary.replace(/(?:-|–)?\s*tháng\s+(\d{1,2})(?:\/|-)(\d{4})/gi, 'tháng $1 năm $2');
    const currentYear = new Date().getFullYear();
    summary = summary.replace(/(?:-|–)?\s*tháng\s+(\d{1,2})(?!\s*năm|\/|-)/gi, `tháng $1 năm ${currentYear}`);

    // Sentence case
    summary = summary.charAt(0).toUpperCase() + summary.slice(1).toLowerCase();

    return { docType, summary };
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
    let hasNormalizedTitle = false;
    let detectedDocType = "";
    for (let i = 0; i < paragraphs.length; i++) {
      const p = paragraphs[i];
      const pIndex = i + 1;
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

      // --- SAFELY INJECT TITLE DETECTION HERE ---
      const runsForTitle = Array.from(p.getElementsByTagNameNS(W_NS, "r"));
      let pText = "";
      runsForTitle.forEach(r => {
          const tNodes = r.getElementsByTagNameNS(W_NS, "t");
          if (tNodes.length > 0) pText += tNodes[0].textContent || "";
      });

      // CRITICAL SAFETY LOCK: Only check first 10 paragraphs, stop if already found
      if (!hasNormalizedTitle && pIndex <= 10 && pText.length > 0 && pText.length < 200) {
          // To avoid false positives in body text, ensure the text is somewhat "title-like" 
          // (e.g. it is entirely uppercase or starts with uppercase doc type)
          const upperPText = pText.toUpperCase().trim();
          const isLikelyTitle = DOC_TYPE_KEYWORDS.some(k => upperPText.startsWith(k));
          
          if (isLikelyTitle) {
              let combinedText = pText;
              let nextP = p.nextSibling;
              let extraParagraphs: Element[] = [];

              // Look ahead up to 3 paragraphs for split titles (usually short and ALL CAPS)
              for (let i = 0; i < 3; i++) {
                  if (nextP && nextP.nodeName === "w:p") {
                      let nextText = "";
                      const nextRuns = Array.from((nextP as Element).getElementsByTagNameNS(W_NS, "r"));
                      nextRuns.forEach(r => {
                          const tNodes = r.getElementsByTagNameNS(W_NS, "t");
                          if (tNodes.length > 0) nextText += tNodes[0].textContent || "";
                      });

                      const trimmedNext = nextText.trim();
                      // If it's short and completely uppercase, it's a continuation of the title
                      const isAllCaps = trimmedNext.length > 0 && trimmedNext === trimmedNext.toUpperCase();
                      
                      if (trimmedNext.length > 0 && trimmedNext.length < 150 && isAllCaps) {
                          combinedText += " " + trimmedNext;
                          extraParagraphs.push(nextP as Element);
                          nextP = nextP.nextSibling;
                      } else {
                          break; // Stop looking if it's normal body text
                      }
                  } else {
                      break;
                  }
              }

              const normalized = normalizeTitleText(combinedText);
              if (normalized) {
                  hasNormalizedTitle = true; // LOCK ENGAGED
                  detectedDocType = normalized.docType;
                  
                  // Delete the extra paragraphs we just merged
                  extraParagraphs.forEach(ep => ep.parentNode?.removeChild(ep));

                  // Clear original paragraph runs
                  runsForTitle.forEach(r => p.removeChild(r));

                  // 1. DocType Run (e.g., "BIÊN BẢN")
                  const rType = doc.createElementNS(W_NS, "w:r");
                  const rPrType = getOrCreate(rType, "w:rPr");
                  rPrType.appendChild(doc.createElementNS(W_NS, "w:b")); 
                  const szType = getOrCreate(rPrType, "w:sz");
                  szType.setAttributeNS(W_NS, "w:val", "28"); 
                  const szCsType = getOrCreate(rPrType, "w:szCs");
                  szCsType.setAttributeNS(W_NS, "w:val", "28"); 
                  const tType = doc.createElementNS(W_NS, "w:t");
                  tType.textContent = normalized.docType;
                  rType.appendChild(tType);
                  p.appendChild(rType);

                  // 2. Summary Paragraph (e.g., "Triển khai kế hoạch sinh hoạt...")
                  const newP = doc.createElementNS(W_NS, "w:p");
                  const newPPr = getOrCreate(newP, "w:pPr");
                  const jcNew = getOrCreate(newPPr, "w:jc");
                  jcNew.setAttributeNS(W_NS, "w:val", "center"); 
                  
                  const rSum = doc.createElementNS(W_NS, "w:r");
                  const rPrSum = getOrCreate(rSum, "w:rPr");
                  rPrSum.appendChild(doc.createElementNS(W_NS, "w:b")); 
                  const szSum = getOrCreate(rPrSum, "w:sz");
                  szSum.setAttributeNS(W_NS, "w:val", "28"); 
                  const szCsSum = getOrCreate(rPrSum, "w:szCs");
                  szCsSum.setAttributeNS(W_NS, "w:val", "28"); 
                  
                  const tSum = doc.createElementNS(W_NS, "w:t");
                  tSum.textContent = normalized.summary;
                  rSum.appendChild(tSum);
                  newP.appendChild(rSum);

                  if (p.nextSibling) {
                      p.parentNode?.insertBefore(newP, p.nextSibling);
                  } else {
                      p.parentNode?.appendChild(newP);
                  }
              }
          }
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
    
    // --- STEP 5: Insert Header and Signature Templates ---
    if (options.headerType !== HeaderType.NONE && body) {
        logs.push("Inserting Standard Header & Signature Templates...");
        
        // Insert Header at the top
        const headerTable = createHeaderTemplate(doc, options);
        if (body.firstChild) {
            body.insertBefore(headerTable, body.firstChild);
        } else {
            body.appendChild(headerTable);
        }

        // Insert Signature Block at the end (safely before final sectPr)
        const sectPrs = body.getElementsByTagNameNS(W_NS, "sectPr");
        const lastSectPr = sectPrs.length > 0 ? sectPrs[sectPrs.length - 1] : null;
        
        // Add a blank paragraph before the signature block for spacing
        const blankP = doc.createElementNS(W_NS, "w:p");
        
        const signatureBlock = createSignatureBlock(doc, options, detectedDocType);
        
        if (lastSectPr && lastSectPr.parentNode === body) {
            body.insertBefore(blankP, lastSectPr);
            body.insertBefore(signatureBlock, lastSectPr);
        } else {
            body.appendChild(blankP);
            body.appendChild(signatureBlock);
        }
    }

    // --- STEP 6: Deep Font Formatting (Force Times New Roman) ---
    // Executed last to ensure everything, including inserted templates, is standardized.
    logs.push("Step 6: Performing Deep Font Formatting (Force Times New Roman)...");
    forceFontApplication(doc, options.font.family);

    // --- STEP 7: AUTO PAGE NUMBERING (Different First Page, Top Center) ---
    logs.push("Configuring Auto Page Numbering (Top Center, starting from Page 2)...");
    
    // 7.1 Create the header XML file with a Page Field
    const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:p>
            <w:pPr>
                <w:jc w:val="center"/>
            </w:pPr>
            <w:r>
                <w:fldChar w:fldCharType="begin"/>
            </w:r>
            <w:r>
                <w:instrText xml:space="preserve"> PAGE </w:instrText>
            </w:r>
            <w:r>
                <w:fldChar w:fldCharType="separate"/>
            </w:r>
            <w:r>
                <w:rPr><w:noProof/></w:rPr>
                <w:t>2</w:t>
            </w:r>
            <w:r>
                <w:fldChar w:fldCharType="end"/>
            </w:r>
        </w:p>
    </w:hdr>`;
    zip.file("word/header_custom.xml", headerXml);

    // 7.2 Update [Content_Types].xml to register the new header part
    let contentTypesXml = await zip.file("[Content_Types].xml")?.async("string");
    if (contentTypesXml && !contentTypesXml.includes('PartName="/word/header_custom.xml"')) {
        contentTypesXml = contentTypesXml.replace(
            '</Types>',
            '<Override PartName="/word/header_custom.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'
        );
        zip.file("[Content_Types].xml", contentTypesXml);
    }

    // 7.3 Update document.xml.rels to link the header part
    let relsXml = await zip.file("word/_rels/document.xml.rels")?.async("string");
    if (relsXml && !relsXml.includes('Target="header_custom.xml"')) {
        relsXml = relsXml.replace(
            '</Relationships>',
            '<Relationship Id="rIdCustomHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header_custom.xml"/></Relationships>'
        );
        zip.file("word/_rels/document.xml.rels", relsXml);
    }

    // 7.4 Update Document sections (sectPr)
    const sectPrs = Array.from(doc.getElementsByTagNameNS(W_NS, "sectPr"));
    const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    
    for (const sPr of sectPrs) {
        // A. Enable "Different First Page" (w:titlePg) so page 1 ignores the default header
        getOrCreate(sPr, "titlePg");
        
        // B. Clean up existing default headers to avoid conflict
        const headerRefs = Array.from(sPr.getElementsByTagNameNS(W_NS, "headerReference"));
        for (const hr of headerRefs) {
            if (hr.getAttributeNS(W_NS, "type") === "default") {
                sPr.removeChild(hr);
            }
        }
        
        // C. Link the custom header to "default" (which means page 2 onwards)
        const newHdrRef = doc.createElementNS(W_NS, "w:headerReference");
        newHdrRef.setAttributeNS(W_NS, "w:type", "default");
        newHdrRef.setAttributeNS(R_NS, "r:id", "rIdCustomHdr");
        
        // Ensure headerReference is inserted before certain elements in sectPr for strict OOXML validation
        // Simply appending it usually works in modern Word, but putting it at the top of sectPr is safer
        if (sPr.firstChild) {
            sPr.insertBefore(newHdrRef, sPr.firstChild);
        } else {
            sPr.appendChild(newHdrRef);
        }
    }

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

        const tcBorders = getOrCreate(tcPr, "w:tcBorders");
        const bottom = getOrCreate(tcBorders, "w:bottom");
        bottom.setAttributeNS(W_NS, "w:val", "single");
        bottom.setAttributeNS(W_NS, "w:sz", "2"); 
        bottom.setAttributeNS(W_NS, "w:space", "0");
        bottom.setAttributeNS(W_NS, "w:color", "000000");

        const tcMar = getOrCreate(tcPr, "w:tcMar");
        ["top", "bottom", "left", "right"].forEach(side => {
            const mar = getOrCreate(tcMar, `w:${side}`);
            mar.setAttributeNS(W_NS, "w:w", "0");
            mar.setAttributeNS(W_NS, "w:type", "dxa");
        });

        // Create paragraph inside the cell with centered alignment
        const p = createElement("w:p");
        tc.appendChild(p);
        
        const pPr = getOrCreate(p, "w:pPr");
        const jcP = getOrCreate(pPr, "w:jc");
        jcP.setAttributeNS(W_NS, "w:val", "center");
        
        // CRITICAL SPACING FIX FOR Tightness
        // Explicitly set 0pt spacing before/after and a tight 12pt line height
        // We MUST use 'exact' lineRule to strictly override inheriting generic spacing
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "0");
        spacing.setAttributeNS(W_NS, "w:line", "240"); // Explicit 240 twips (12pt) spacing
        spacing.setAttributeNS(W_NS, "w:lineRule", "exact"); // Force exact rule

        // Create the text run with styling
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

    // Helper to create a short solid black line (approx 1/3 width of the text)
    const createShortLineTable = (): Element => {
        const tbl = createElement("w:tbl");
        const tblPr = getOrCreate(tbl, "w:tblPr");
        
        const jcTbl = getOrCreate(tblPr, "w:jc");
        jcTbl.setAttributeNS(W_NS, "w:val", "center");
        
        const tblLayout = getOrCreate(tblPr, "w:tblLayout");
        tblLayout.setAttributeNS(W_NS, "w:type", "fixed");

        const tr = createElement("w:tr");
        tbl.appendChild(tr);

        const tc = createElement("w:tc");
        tr.appendChild(tc);
        const tcPr = getOrCreate(tc, "w:tcPr");

        // Set width exactly to ~1/3 of "TRƯỜNG THCS CHU VĂN AN" (approx 900 twips)
        const tcW = getOrCreate(tcPr, "w:tcW");
        tcW.setAttributeNS(W_NS, "w:w", "900");
        tcW.setAttributeNS(W_NS, "w:type", "dxa");

        const tcBorders = getOrCreate(tcPr, "w:tcBorders");
        // Using top border so it sits extremely close to the text above
        const top = getOrCreate(tcBorders, "w:top"); 
        top.setAttributeNS(W_NS, "w:val", "single");
        top.setAttributeNS(W_NS, "w:sz", "2"); // 1/4 pt (thin line)
        top.setAttributeNS(W_NS, "w:space", "0");
        top.setAttributeNS(W_NS, "w:color", "000000");

        const tcMar = getOrCreate(tcPr, "w:tcMar");
        ["top", "bottom", "left", "right"].forEach(side => {
            const mar = getOrCreate(tcMar, `w:${side}`);
            mar.setAttributeNS(W_NS, "w:w", "0");
            mar.setAttributeNS(W_NS, "w:type", "dxa");
        });

        // Blank tight paragraph inside the cell
        const p = createElement("w:p");
        tc.appendChild(p);
        const pPr = getOrCreate(p, "w:pPr");
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "0");
        spacing.setAttributeNS(W_NS, "w:line", "24"); // Ultra tight height (1.2pt)
        spacing.setAttributeNS(W_NS, "w:lineRule", "exact");

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

    const docDate = options.documentDate ? new Date(options.documentDate) : new Date();
    const day = String(docDate.getDate()).padStart(2, '0');
    const month = String(docDate.getMonth() + 1).padStart(2, '0');
    const year = docDate.getFullYear();
    const currentDateStr = `Ea Kar, ngày ${day} tháng ${month} năm ${year}`;

    switch (options.headerType) {
        case HeaderType.PARTY:
            tc1.appendChild(createStyledP("ĐẢNG BỘ XÃ EA KAR", false, false));
            tc1.appendChild(createStyledP("CHI BỘ TRƯỜNG THCS CHU VĂN AN", true, false));
            tc1.appendChild(createStyledP("*", false, false)); // Mandatory asterisk for Party documents
            tc1.appendChild(createStyledP("", false, false));  // Empty line for vertical alignment
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
            tc1.appendChild(createShortLineTable()); // Use OOXML line
            tc1.appendChild(createStyledP("", false, false));
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
            tc1.appendChild(createShortLineTable());
            tc1.appendChild(createStyledP("", false, false));
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

// Helper to create the Signature Block at the end of the document
const createSignatureBlock = (doc: Document, options: DocxOptions, docType: string): Element => {
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

    // Special tight-spacing P creator for the signature block
    const createTightP = (text: string, isBold: boolean, isItalic: boolean, align: string) => {
        const p = createElement("w:p");
        const pPr = getOrCreate(p, "w:pPr");
        
        const jc = getOrCreate(pPr, "w:jc");
        jc.setAttributeNS(W_NS, "w:val", align);
        
        const spacing = getOrCreate(pPr, "w:spacing");
        spacing.setAttributeNS(W_NS, "w:before", "0");
        spacing.setAttributeNS(W_NS, "w:after", "0");
        spacing.setAttributeNS(W_NS, "w:line", "240"); // Explicit 12 twips spacing
        spacing.setAttributeNS(W_NS, "w:lineRule", "auto"); // Single Spacing is correct here

        const r = createElement("w:r");
        p.appendChild(r);
        const rPr = getOrCreate(r, "w:rPr");
        
        const sz = getOrCreate(rPr, "w:sz");
        sz.setAttributeNS(W_NS, "w:val", String(options.font.sizeTable * 2)); // Use table font size (13pt)
        const szCs = getOrCreate(rPr, "w:szCs");
        szCs.setAttributeNS(W_NS, "w:val", String(options.font.sizeTable * 2));

        if (isBold) rPr.appendChild(createElement("w:b"));
        if (isItalic) rPr.appendChild(createElement("w:i"));

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
    
    // 1. CRITICAL FIX: Set Table Width to 100% (5000 pct)
    const tblW = getOrCreate(tblPr, "w:tblW");
    tblW.setAttributeNS(W_NS, "w:w", "5000");
    tblW.setAttributeNS(W_NS, "w:type", "pct"); // Changed from dxa to pct

    const tr = createElement("w:tr");
    tbl.appendChild(tr);

    // --- Left Cell ---
    const tc1 = createElement("w:tc");
    tr.appendChild(tc1);
    const tc1Pr = getOrCreate(tc1, "w:tcPr");

    // --- Right Cell ---
    const tc2 = createElement("w:tc");
    tr.appendChild(tc2);
    const tc2Pr = getOrCreate(tc2, "w:tcPr");

    // Check if it is a Meeting Minutes document
    const isMinutes = docType === "BIÊN BẢN";

    if (isMinutes) {
        // 50/50 Width Split for Minutes
        const tc1W = getOrCreate(tc1Pr, "w:tcW");
        tc1W.setAttributeNS(W_NS, "w:w", "2500"); // 50%
        tc1W.setAttributeNS(W_NS, "w:type", "pct");
        
        const tc2W = getOrCreate(tc2Pr, "w:tcW");
        tc2W.setAttributeNS(W_NS, "w:w", "2500"); // 50%
        tc2W.setAttributeNS(W_NS, "w:type", "pct");

        // Insert Minutes Signatures
        tc1.appendChild(createTightP("THƯ KÝ", true, false, "center"));
        tc2.appendChild(createTightP("CHỦ TRÌ", true, false, "center"));
    } else {
        // Standard 40/60 Width Split for other documents
        const tc1W = getOrCreate(tc1Pr, "w:tcW");
        tc1W.setAttributeNS(W_NS, "w:w", "2000"); // 40%
        tc1W.setAttributeNS(W_NS, "w:type", "pct");
        
        const tc2W = getOrCreate(tc2Pr, "w:tcW");
        tc2W.setAttributeNS(W_NS, "w:w", "3000"); // 60%
        tc2W.setAttributeNS(W_NS, "w:type", "pct");

        // Insert "Nơi nhận"
        tc1.appendChild(createTightP("Nơi nhận:", true, true, "left"));
        tc1.appendChild(createTightP("- Như trên;", false, false, "left"));
        tc1.appendChild(createTightP("- Lưu: VT...", false, false, "left"));

        // Insert Standard Signatures based on HeaderType
        switch (options.headerType) {
            case 'PARTY':
                tc2.appendChild(createTightP("T/M CHI BỘ", true, false, "center"));
                tc2.appendChild(createTightP("BÍ THƯ", true, false, "center"));
                break;
            case 'DEPARTMENT':
                tc2.appendChild(createTightP("TỔ TRƯỞNG", true, false, "center"));
                break;
            case 'SCHOOL':
            default:
                tc2.appendChild(createTightP("HIỆU TRƯỞNG", true, false, "center"));
                break;
        }
    }

    return tbl;
}
