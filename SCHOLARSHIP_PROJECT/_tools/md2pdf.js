// usage: node md2pdf.js in.html out.pdf [landscape]
const { chromium } = require('playwright');
(async () => {
  const [inp, out, orient] = process.argv.slice(2);
  const b = await chromium.launch();
  const p = await b.newPage();
  await p.goto('file://' + inp, { waitUntil: 'load' });
  if (orient === 'landscape') await p.addStyleTag({ content: '@page { size: A4 landscape; }' });
  await p.pdf({ path: out, format: 'A4', landscape: orient === 'landscape', printBackground: true, preferCSSPageSize: true,
    displayHeaderFooter: true, headerTemplate: '<span></span>',
    footerTemplate: '<div style="font-size:7pt;color:#777;width:100%;text-align:center;font-family:Arial">Page <span class="pageNumber"></span> of <span class="totalPages"></span></div>',
    margin: { top: '14mm', bottom: '14mm', left: '12mm', right: '12mm' } });
  await b.close();
})();
