# Preview provenance

The PNG files are rasterized, unmodified PDF exports of the actual **Panel** worksheet from LibreOffice Calc. They contain the workbook's native chart and conditional formatting; no dashboard was drawn separately.

The rendering application was LibreOffice Portable **26.2.4.2** (build `0229ac93fcf0d7cbc6376066c6f35021cef002dc`), downloaded through the [LibreOffice portable distribution page](https://www.libreoffice.org/download-other/#libreoffice-portable) and PortableApps.com. The downloaded Standard package matched its published SHA-256:

```text
4bde93374aef4409243505b20d16561a4628ac7591457dd01fc6e1ccf571ba65
```

For each sample month, a temporary copy retained all original formulas, removed their cached values and set **Panel!B4** to the selected month. Headless LibreOffice opened and resaved that copy as XLSX. Python read the newly computed values and checked them against independent raw-input totals, including all 122 panel comparisons over the two months. The delivered workbook was never resaved through LibreOffice.

The recalculated copies were exported using the `calc_pdf_Export` filter with `PageRange` set to `1`, preserving the worksheet's landscape A4 print area. Each PDF has one page. PyMuPDF **2.0** rendered each page at twice the PDF's coordinate resolution, producing **1684 × 1191** PNGs. The two images were visually inspected for readable totals, tables, charts and the absence of evaluation watermarks.

The portable application, Python environments and temporary copies are excluded from the package. Neither LibreOffice nor PyMuPDF is needed by `generate.py`, `verify.py` or the CI workflow. This preview check does not cover mouse/keyboard interactions in Excel.
