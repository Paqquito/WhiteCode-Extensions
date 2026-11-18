import { marked, type Token, type Tokens } from 'marked'
import { Document, ExternalHyperlink, HeadingLevel, Packer, Paragraph, Table, TableCell, TableRow, TextRun, WidthType, type IRunOptions } from 'docx'

type Inline = TextRun | ExternalHyperlink
const headings = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6]

function inline(tokens: Token[], style: IRunOptions = {}): Inline[] {
  return tokens.flatMap<Inline>(token => {
    switch (token.type) {
      case 'strong': return inline(token.tokens, { ...style, bold: true })
      case 'em': return inline(token.tokens, { ...style, italics: true })
      case 'del': return inline(token.tokens, { ...style, strike: true })
      case 'br': return [new TextRun({ ...style, break: 1 })]
      case 'codespan': return [new TextRun({ ...style, font: 'Consolas', text: token.text })]
      case 'link': {
        if (/^(https?:|mailto:)/i.test(token.href)) return [new ExternalHyperlink({ link: token.href, children: [new TextRun({ ...style, style: 'Hyperlink', text: plain(token.tokens) })] })]
        return inline(token.tokens, style)
      }
      case 'image': return [new TextRun({ ...style, italics: true, text: token.text || token.href })]
      default: return [new TextRun({ ...style, text: String('text' in token ? token.text : token.raw).replace(/\r?\n/g, ' ') })]
    }
  })
}
function plain(tokens: Token[]): string {
  return tokens.map(token => 'tokens' in token && token.tokens ? plain(token.tokens) : String('text' in token ? token.text : token.raw)).join('')
}
function tableRow(cells: Tokens.TableCell[], header: boolean): TableRow {
  return new TableRow({ tableHeader: header, children: cells.map(cell => new TableCell({
    children: [new Paragraph({ children: inline(cell.tokens, { bold: header }) })],
  })) })
}
function blocks(tokens: Token[], depth = 0): Array<Paragraph | Table> {
  return tokens.flatMap<Paragraph | Table>(token => {
    switch (token.type) {
      case 'space': case 'def': return []
      case 'heading': return [new Paragraph({ heading: headings[token.depth - 1], children: inline(token.tokens) })]
      case 'blockquote': return blocks(token.tokens, depth)
      case 'code': return String(token.text).split('\n').map(line => new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ font: 'Consolas', text: line })] }))
      case 'table': return [new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [tableRow(token.header, true), ...token.rows.map((row: Tokens.TableCell[]) => tableRow(row, false))] })]
      case 'list': return token.items.flatMap((item: Tokens.ListItem, index: number) => {
        let first = true
        return item.tokens.flatMap(part => {
          if (part.type !== 'text' && part.type !== 'paragraph') return blocks([part], depth + 1)
          const children = inline(part.tokens ?? [part])
          const prefix = first && token.ordered ? Number(token.start || 1) + index + '. ' : ''
          const bullet = first && !token.ordered ? { level: Math.min(depth, 8) } : undefined
          first = false
          return [new Paragraph({ bullet, children: [...(prefix ? [new TextRun(prefix)] : []), ...children] })]
        })
      })
      case 'hr': return [new Paragraph({ thematicBreak: true })]
      default: return [new Paragraph({ children: inline('tokens' in token && token.tokens ? token.tokens : [token]) })]
    }
  })
}
export async function markdownWord(source: string): Promise<Buffer> {
  const children = blocks(marked.lexer(source.replace(/^\uFEFF/, '')))
  if (!children.length) throw Error('Le document source est vide.')
  return Packer.toBuffer(new Document({ sections: [{ children }] }))
}
