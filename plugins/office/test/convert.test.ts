import { expect, test } from 'bun:test'
import { execFile } from 'node:child_process'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import JSZip from 'jszip'
import ExcelJS from 'exceljs'

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
