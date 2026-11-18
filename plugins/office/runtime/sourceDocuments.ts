import { marked, type Token, type Tokens } from 'marked'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import ExcelJS from 'exceljs'
import PptxGenJS from 'pptxgenjs'

type Block = { kind: 'heading' | 'text'; text: string } | { kind: 'table'; rows: string[][] }

function inline(tokens: Token[]): string {
  return tokens.map(token => {
    if (token.type === 'checkbox') return ''
    if (token.type === 'br') return '\n'
    if (token.type === 'image') return token.text || token.href
    if ('tokens' in token && token.tokens) return inline(token.tokens)
    return 'text' in token ? String(token.text) : token.raw
  }).join('')
}

function blocks(tokens: Token[]): Block[] {
  const result: Block[] = []
  for (const token of tokens) {
    if (['space', 'def', 'hr', 'checkbox'].includes(token.type)) continue
    if (token.type === 'table') {
      result.push({ kind: 'table', rows: [token.header, ...token.rows].map(row => row.map((c: Tokens.TableCell) => inline(c.tokens))) })
    } else if (token.type === 'list') {
      token.items.forEach((item: Tokens.ListItem, index: number) => {
        const content = blocks(item.tokens)
        const first = content[0]
        if (first && first.kind === 'text') {
          const marker = item.task ? (item.checked ? '[x]' : '[ ]') : token.ordered ? String(Number(token.start || 1) + index) + '.' : '•'
          first.text = marker + ' ' + first.text
        }
        result.push(...content)
      })
    } else if (token.type === 'blockquote') result.push(...blocks(token.tokens))
    else result.push({ kind: token.type === 'heading' ? 'heading' : 'text', text: inline([token]) })
  }
  return result
}

function lines(block: Block): string[] {
  return block.kind === 'table' ? block.rows.map(row => row.join('  |  ')) : block.text.split(/\r?\n/)
}

function cell(text: string): string | number {
  // Preserve identifiers with leading zeroes, percentages, dates and precise decimal spellings.
  const value = Number(text)
  return /^-?(0|[1-9]\d*)$/.test(text) && Number.isSafeInteger(value) && String(value) === text ? value : text
}

async function spreadsheet(content: Block[]) {
  const book = new ExcelJS.Workbook()
  const tables = content.filter(b => b.kind === 'table')
  tables.forEach((table, i) => {
    const sheet = book.addWorksheet('Tableau ' + (i + 1))
    sheet.addRows([table.rows[0], ...table.rows.slice(1).map(row => row.map(cell))])
    sheet.getRow(1).font = { bold: true }
    sheet.views = [{ state: 'frozen', ySplit: 1 }]
    sheet.columns.forEach(column => {
      let width = 10
      column.eachCell?.(c => { width = Math.min(60, Math.max(width, c.text.length + 2)) })
      column.width = width
    })
  })
  const notes = content.filter(b => b.kind !== 'table')
  if (notes.length) {
    const sheet = book.addWorksheet('Source')
    sheet.getColumn(1).width = 90
    for (const block of notes) {
      const row = sheet.addRow([block.text])
      row.getCell(1).alignment = { wrapText: true }
      if (block.kind === 'heading') row.font = { bold: true }
    }
  }
  return new Uint8Array(await book.xlsx.writeBuffer())
}

async function presentation(content: Block[]) {
  const deck = new PptxGenJS()
  deck.layout = 'LAYOUT_WIDE'
  let heading = '', body: string[] = []
  function flush() {
    if (!heading && !body.length) return
    const slide = deck.addSlide()
    if (heading) slide.addText(heading, { x: 0.5, y: 0.3, w: 12.2, h: 0.9, fontSize: 29, bold: true, fit: 'shrink' })
    if (body.length) slide.addText(body.join('\n'), { x: 0.6, y: heading ? 1.4 : 0.5, w: 12, h: 5.6, fontSize: 20, fit: 'shrink', valign: 'top' })
    body = []
  }
  for (const block of content) {
    if (block.kind === 'heading') { flush(); heading = block.text }
    else for (const line of lines(block)) {
      if (body.length >= 12) flush()
      body.push(line)
    }
  }
  flush()
  return new Uint8Array(await deck.write({ outputType: 'uint8array' }) as Uint8Array)
}

async function pdf(content: Block[]) {
  const doc = await PDFDocument.create()
  const regular = await doc.embedFont(StandardFonts.Helvetica)
  const bold = await doc.embedFont(StandardFonts.HelveticaBold)
  const size = [595, 842] as [number, number]
  let page = doc.addPage(size), y = 790
  for (const block of content) {
    const font = block.kind === 'heading' ? bold : regular
    const fontSize = block.kind === 'heading' ? 16 : 11
    function draw(line: string) {
      if (y < 52) { page = doc.addPage(size); y = 790 }
      page.drawText(line, { x: 45, y, size: fontSize, font })
      y -= fontSize * 1.45
    }
    for (const line of lines(block)) {
      let current = ''
      for (const word of line.replace(/\t/g, '    ').split(/\s+/u)) {
        if (current && font.widthOfTextAtSize(current + ' ' + word, fontSize) > 500) { draw(current); current = '' }
        if (current) current += ' '
        for (const char of word) {
          if (current && font.widthOfTextAtSize(current + char, fontSize) > 500) { draw(current); current = '' }
          current += char
        }
      }
      draw(current)
    }
    y -= 6
  }
  return doc.save()
}

export async function renderSource(source: string, extension: string) {
  const content = blocks(marked.lexer(source))
  if (!content.length) throw Error('Le document source est vide.')
  if (extension === '.xlsx') return spreadsheet(content)
  if (extension === '.pptx') return presentation(content)
  return pdf(content)
}
