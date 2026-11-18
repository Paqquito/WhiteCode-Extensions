import { expect, test } from 'bun:test'
import { execFile } from 'node:child_process'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import JSZip from 'jszip'
import ExcelJS from 'exceljs'
import { PDFDocument, PDFArray, PDFRawStream, decodePDFRawStream } from 'pdf-lib'

const run = promisify(execFile)
const runtime = join(import.meta.dir, '..', 'office.cjs')

test('le convertisseur livré écrit des fichiers Office et préserve les originaux', async () => {
  const cwd = await mkdtemp(join(tmpdir(), 'office conversion é '))
  const convert = (input: string, output: string) => run(process.execPath, [runtime, '--convert', input, output], { cwd })
  try {
    await writeFile(join(cwd, 'README.md'), '# Recette\n\n## Décisions\n\n- **Livré** à temps\n\n| Équipe | Lot |\n|---|---|\n| Études | 42 |\n\n```js\nconst A = 1\n```\n')
    expect((await convert('README.md', 'note.docx')).stdout.trim()).toBe(join(cwd, 'note.docx'))
    const original = await readFile(join(cwd, 'note.docx'))
    const archive = await JSZip.loadAsync(original)
    const body = await archive.file('word/document.xml')!.async('string')
    for (const text of ['Heading1', 'Heading2', 'Décisions', 'Livré', '<w:tbl>', 'Études', 'const A = 1']) expect(body).toContain(text)
    await expect(convert('README.md', 'note.docx')).rejects.toThrow('EEXIST')
    expect(await readFile(join(cwd, 'note.docx'))).toEqual(original)
    await writeFile(join(cwd, 'table.json'), JSON.stringify({ sheets: [{ name: 'Budget', rows: [['Libellé','Valeur'],['Études',42],['Confirmé',true]] }] }))
    await convert('table.json', 'budget.xlsx')
    const book = new ExcelJS.Workbook()
    await book.xlsx.readFile(join(cwd, 'budget.xlsx'))
    expect(book.getWorksheet('Budget')!.getCell('B2').value).toBe(42)
    expect(book.getWorksheet('Budget')!.getCell('B3').value).toBe(true)
    await writeFile(join(cwd, 'slides.json'), JSON.stringify({ slides: [{ title: 'Décision', bullets: ['Projet validé','Envoi demain'], notes: 'Réunion' }, { title: 'Fin', bullets: [] }] }))
    await convert('slides.json', 'synthèse.pptx')
    const deck = await JSZip.loadAsync(await readFile(join(cwd, 'synthèse.pptx')))
    expect(await deck.file('ppt/slides/slide1.xml')!.async('string')).toContain('Projet validé')
    expect(await deck.file('ppt/slides/slide2.xml')!.async('string')).toContain('Fin')
    expect(await deck.file('ppt/notesSlides/notesSlide1.xml')!.async('string')).toContain('Réunion')
    await expect(convert('table.json', 'invalid.pptx')).rejects.toThrow()
    await expect(readFile(join(cwd, 'invalid.pptx'))).rejects.toThrow('ENOENT')
  } finally {
    await rm(cwd, { recursive: true, force: true })
  }
}, 30000)

test('un lot garde les décisions ouvertes, les réserves et les données de la même source dans les quatre fichiers', async () => {
  const cwd = await mkdtemp(join(tmpdir(), 'office source '))
  const source = '# Essais\n\n## Suivi\n\n- [ ] Valider le lot A.\n- Ne pas ouvrir le lot B avant accord.\n\n| Lot | Essais |\n| --- | --- |\n| 007 | 17 |\n'
  const outputs = ['note.docx', 'lots.xlsx', 'revue.pptx', 'fiche.pdf']
  try {
    await writeFile(join(cwd, 'README.md'), source)
    const batch = (...paths: string[]) => run(process.execPath, [runtime, '--convert', 'README.md', ...paths], { cwd })
    await batch(...outputs)
    const xml = async (name: string, member: string) => (await JSZip.loadAsync(await readFile(join(cwd, name)))).file(member)!.async('string')
    const book = new ExcelJS.Workbook()
    await book.xlsx.readFile(join(cwd, 'lots.xlsx'))
    expect(book.worksheets[0]!.getRow(2).values).toEqual([undefined, '007', 17])
    const pdf = await PDFDocument.load(await readFile(join(cwd, 'fiche.pdf')))
    expect(pdf.getPageCount()).toBe(1)
    const contents = pdf.getPage(0).node.Contents() as PDFArray
    const operators = contents.asArray().map(ref => Buffer.from(decodePDFRawStream(pdf.context.lookup(ref, PDFRawStream)).decode()).toString('latin1')).join('')
    const pdfText = Array.from(operators.matchAll(/<([0-9a-f]+)> Tj/gi), match => Buffer.from(match[1]!, 'hex').toString('latin1')).join(' ')
    for (const text of [await xml('note.docx', 'word/document.xml'), await xml('revue.pptx', 'ppt/slides/slide2.xml'), JSON.stringify(book.getWorksheet('Source')!.getSheetValues()), pdfText]) {
      expect(text.split('[ ]')).toHaveLength(2)
      for (const fact of ['[ ]', 'Valider le lot A.', 'Ne pas ouvrir le lot B avant accord.']) expect(text).toContain(fact)
    }
    await expect(batch('nouveau.pdf', 'lots.xlsx')).rejects.toThrow('EEXIST')
    await expect(readFile(join(cwd, 'nouveau.pdf'))).rejects.toThrow('ENOENT')
    expect(await readFile(join(cwd, 'README.md'), 'utf8')).toBe(source)
  } finally { await rm(cwd, { recursive: true, force: true }) }
}, 30000)
