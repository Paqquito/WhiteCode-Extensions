const { resolve } = require('node:path');
const { writeFile } = require('node:fs/promises');
module.exports = async ({ PptxGenJS }) => {
  const presentation = new PptxGenJS();
  presentation.layout = 'LAYOUT_WIDE';
  presentation.author = 'WhiteCode';
  presentation.subject = 'Revue de projet';
  const slide = presentation.addSlide();
  slide.addText('Revue de projet', {x:0.6,y:0.5,w:12,h:0.7,fontSize:30,bold:true,color:'005A85'});
  slide.addText('Résultats et prochaines étapes', {x:0.6,y:1.5,w:12,h:1,fontSize:20});
  const output = resolve('presentation.pptx');
  await writeFile(output, await presentation.write({outputType:'nodebuffer'}), {flag:'wx'});
  console.log(output);
};
