import { readFile, open, unlink } from 'node:fs/promises'
import { extname, resolve } from 'node:path'
import ExcelJS from 'exceljs'
import PptxGenJS from 'pptxgenjs'
import { z } from 'zod'
import { renderSource } from './sourceDocuments'
import { markdownWord } from './markdownWord'

const cell = z.union([z.string(), z.number(), z.boolean(), z.null()])
const workbookInput = z.object({ sheets: z.array(z.object({
  name: z.string().min(1), rows: z.array(z.array(cell).min(1)).min(1),
}).strict()).min(1) }).strict()
const slideInput = z.object({ slides: z.array(z.object({
  title: z.string().min(1), bullets: z.array(z.string().min(1)), notes: z.string().optional(),
}).strict()).min(1) }).strict()

async function workbook(source: string): Promise<Uint8Array> {
  const data = workbookInput.parse(JSON.parse(source))
  const book = new ExcelJS.Workbook()
  for (const { name, rows } of data.sheets) {
    const sheet = book.addWorksheet(name)
    sheet.addRows(rows)
    sheet.views = [{ state: 'frozen', ySplit: 1 }]
    sheet.getRow(1).font = { bold: true }
    sheet.columns.forEach(column => {
      let width = 10
      column.eachCell?.(c => { width = Math.max(width, Math.min(60, c.text.length + 2)) })
      column.width = width
    })
  }
  return new Uint8Array(await book.xlsx.writeBuffer())
}

async function slides(source: string): Promise<Uint8Array> {
  const data = slideInput.parse(JSON.parse(source))
  const deck = new PptxGenJS()
  deck.layout = 'LAYOUT_WIDE'
  for (const { title, bullets, notes } of data.slides) {
    const slide = deck.addSlide()
    slide.addText(title, { x: 0.6, y: 0.4, w: 12, h: 1.1, fontSize: 30, bold: true, fit: 'shrink' })
    if (bullets.length) slide.addText(bullets.map((text, i) => ({ text, options: { bullet: true, breakLine: i < bullets.length - 1 } })),
      { x: 0.8, y: 1.8, w: 11.7, h: 5.0, fontSize: 20, paraSpaceAfter: 12, fit: 'shrink', valign: 'top' })
    if (notes) slide.addNotes(notes)
  }
  return new Uint8Array(await deck.write({ outputType: 'uint8array' }) as Uint8Array)
}

/** Render every requested format from the same immutable source snapshot. */
export async function convert(input: string, outputs: string[]): Promise<string[]> {
  const source = (await readFile(input, 'utf8')).replace(/^\\uFEFF/, '')
  const sourceType = extname(input).toLowerCase()
  const targets = outputs.map(output => resolve(output))
  const created: string[] = []
  try {
    for (const target of targets) {
      const ext = extname(target).toLowerCase()
      let data: Uint8Array
      if (['.md', '.markdown', '.txt'].includes(sourceType) && ['.docx', '.xlsx', '.pptx', '.pdf'].includes(ext)) {
        data = ext === '.docx' ? await markdownWord(source) : await renderSource(source, ext)
      } else if (sourceType === '.json' && ext === '.xlsx') data = await workbook(source)
      else if (sourceType === '.json' && ext === '.pptx') data = await slides(source)
      else throw Error('Conversions disponibles : Markdown ou texte vers DOCX, XLSX, PPTX ou PDF, JSON vers XLSX ou PPTX.')
      const handle = await open(target, 'wx')
      created.push(target)
      try { await handle.writeFile(data) } finally { await handle.close() }
    }
    return targets
  } catch (error) {
    for (const target of created) await unlink(target).catch(failure => { if (failure.code !== 'ENOENT') throw failure })
    throw error
  }
}
