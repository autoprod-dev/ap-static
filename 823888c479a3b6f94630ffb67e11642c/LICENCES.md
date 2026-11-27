# Licences

My Office is a private, personal tool (not for sale, not published). It is built only from open-source libraries whose licences allow personal use, copying and changes. Full licence texts are in `vendor/licences/`.

| Library | Version | Used for | Licence |
|---|---|---|---|
| mammoth | 1.13.0 | Opening .docx | BSD-2-Clause |
| docx (+ docx/html) | 9.16.1 | Saving .docx | MIT |
| ExcelJS | 4.4.0 | Opening and saving .xlsx (with formatting) | MIT |
| fast-formula-parser | 1.0.19 | Spreadsheet formulas | MIT |
| @formulajs/formulajs | 4.6.1 | Extra spreadsheet functions (MAX, MATCH, PMT and so on) | MIT |
| PptxGenJS | 4.0.1 | Saving .pptx | MIT |
| JSZip | 3.10.2 | Reading .pptx, backup files | MIT (dual MIT or GPL-3.0; used under MIT) |
| pdf.js (pdfjs-dist) | 6.5.438 | Viewing PDFs | Apache-2.0 |

Their bundled dependencies are MIT, ISC, BSD, Apache-2.0, Unlicense or BlueOak licensed.

Not used:

- **HyperFormula**: GPL-3.0 (or a paid commercial key). GPL would be allowed for private use, but the brief preferred MIT/Apache/MPL.
- **Univer**: Apache-2.0 and allowed, but too heavy (see README).
- No Microsoft software, fonts, icons or branding.
