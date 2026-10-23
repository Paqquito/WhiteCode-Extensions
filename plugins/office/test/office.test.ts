/**
 * Le plugin Bureautique tel qu'il s'installe : copié ailleurs, sans
 * node_modules, son `office.cjs` construit doit suffire à produire les quatre
 * documents de ses exemples.
 *
 * Le desktop exécute ce runtime avec Electron en mode Node
 * (`ELECTRON_RUN_AS_NODE`) ; ce test l'exécute avec Bun, son équivalent ici.
 * Le passage par Electron est vérifié par le test d'intégration du desktop.
 */

import { expect, test } from 'bun:test'
import { cp, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { execFile } from 'node:child_process'
import { promisify } from 'node:util'
import JSZip from 'jszip'
import ExcelJS from 'exceljs'
import { PDFDocument } from 'pdf-lib'

const PLUGIN = join(import.meta.dir, '..')
const run = promisify(execFile)

test('le plugin déplacé crée quatre documents valides sans dépendance installée', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'whitecode office é '))
  const plugin = join(directory, 'plugin')
  await cp(PLUGIN, plugin, { recursive: true, filter: source => !source.includes(`${join(PLUGIN, 'runtime')}`) && !source.includes(`${join(PLUGIN, 'test')}`) })
  for (const skill of ['excel', 'word', 'powerpoint', 'pdf']) {
    const { stdout } = await run(process.execPath, [join(plugin, 'office.cjs'), join(plugin, 'skills', skill, 'example.cjs')], {
      cwd: directory, windowsHide: true, timeout: 20000, env: { ...process.env, NODE_PATH: '' },
    })
    expect(stdout).toContain(directory)
  }
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(join(directory, 'budget.xlsx'))
  expect(workbook.getWorksheet('Budget')!.getCell('B4').value).toEqual({ formula: 'SUM(B2:B3)', result: 2000 })
  expect(workbook.getWorksheet('Budget')!.getRow(1).font.bold).toBe(true)
  const word = await JSZip.loadAsync(await readFile(join(directory, 'compte-rendu.docx')))
  const body = await word.file('word/document.xml')!.async('string')
  expect(body).toContain('Décisions')
  expect(body).toContain('<w:tbl>')
  const deck = await JSZip.loadAsync(await readFile(join(directory, 'presentation.pptx')))
  expect(await deck.file('ppt/slides/slide1.xml')!.async('string')).toContain('Revue de projet')
  expect((await PDFDocument.load(await readFile(join(directory, 'rapport.pdf')))).getPageCount()).toBe(1)
  const bad = join(directory, 'failure.cjs')
  await writeFile(bad, 'module.exports = async () => { throw Error("document failed") }')
  await expect(run(process.execPath, [join(plugin, 'office.cjs'), bad], { windowsHide: true })).rejects.toThrow('document failed')
}, 90000)

test('le manifeste, le catalogue et les licences sont cohérents', async () => {
  const manifest = JSON.parse(await readFile(join(PLUGIN, '.claude-plugin', 'plugin.json'), 'utf8'))
  const catalogue = JSON.parse(await readFile(join(PLUGIN, '..', '..', '.claude-plugin', 'marketplace.json'), 'utf8'))
  const entry = catalogue.plugins.find((item: { name: string }) => item.name === manifest.name)
  expect(entry.version).toBe(manifest.version)
  const notices = await readFile(join(PLUGIN, 'THIRD_PARTY_NOTICES.txt'), 'utf8')
  for (const name of ['exceljs', 'docx', 'pptxgenjs', 'pdf-lib']) expect(notices).toContain(`## ${name} `)
})
