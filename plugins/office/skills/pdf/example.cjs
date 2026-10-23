const { writeFile } = require('node:fs/promises');
const { resolve } = require('node:path');
module.exports = async ({ PDFLib }) => {
  const {PDFDocument,StandardFonts,rgb} = PDFLib;
  const document = await PDFDocument.create();
  const font = await document.embedFont(StandardFonts.Helvetica);
  const page = document.addPage([595.28,841.89]);
  page.drawText('Rapport de synthèse', {x:48,y:780,size:24,font,color:rgb(0,0.35,0.52)});
  page.drawText('Résultats vérifiés et prochaines étapes.', {x:48,y:730,size:12,font});
  const output = resolve('rapport.pdf');
  const bytes = await document.save();
  if ((await PDFDocument.load(bytes)).getPageCount() !== 1) throw Error('PDF invalide');
  await writeFile(output,bytes,{flag:'wx'}); console.log(output);
};
