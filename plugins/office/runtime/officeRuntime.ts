/** Bundled libraries for agent-authored scripts, executed through the ordinary shell permission flow. */
import ExcelJS from 'exceljs'
import * as docx from 'docx'
import PptxGenJS from 'pptxgenjs'
import * as PDFLib from 'pdf-lib'
import { createRequire } from 'node:module'
import { resolve, extname } from 'node:path'

import { convert } from './convert'

async function main() {
  if (process.argv[2] === '--convert') {
    if (process.argv.length !== 5) throw Error('Usage : office.cjs --convert <source> <sortie>')
    console.log(await convert(process.argv[3]!, process.argv[4]!))
    return
  }
  const script = process.argv[2]
  if (!script || extname(script) !== '.cjs') throw new Error('Usage : office.cjs <script.cjs>. Le script doit exporter une fonction async ({ ExcelJS, docx, PptxGenJS, PDFLib }).')
  const run = createRequire(resolve(script))(resolve(script))
  if (typeof run !== 'function') throw new Error('Le script doit exporter une fonction avec module.exports.')
  await run({ ExcelJS, docx, PptxGenJS, PDFLib })
}
main().catch(error => { console.error(error instanceof Error ? error.message : String(error)); process.exitCode = 1 })
