const { writeFile } = require('node:fs/promises');
const { resolve } = require('node:path');
module.exports = async ({ docx }) => {
  const {Document,Packer,Paragraph,HeadingLevel,Table,TableRow,TableCell} = docx;
  const document = new Document({sections:[{children:[
    new Paragraph({text:'Compte rendu',heading:HeadingLevel.TITLE}),
    new Paragraph({text:'Décisions',heading:HeadingLevel.HEADING_1}),
    new Paragraph('La revue du projet est terminée.'),
    new Table({rows:[new TableRow({children:['Action','Responsable'].map(text => new TableCell({children:[new Paragraph(text)]}))})]})
  ]}]});
  const output = resolve('compte-rendu.docx');
  await writeFile(output, await Packer.toBuffer(document), {flag:'wx'});
  console.log(output);
};
