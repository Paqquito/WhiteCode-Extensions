/**
 * Construit `plugins/office/office.cjs` et ses mentions de licence.
 *
 * Le fichier construit est versionné : un plugin s'installe par `git clone`,
 * sans aucune construction sur le poste de l'utilisateur. Relancez ce script
 * après toute modification du runtime ou des bibliothèques, puis committez le
 * résultat avec la nouvelle version du plugin.
 *
 *   bun install --frozen-lockfile
 *   bun run build:office
 */

import { readdir, readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'

const ROOT = join(import.meta.dir, '..')
const PLUGIN = join(ROOT, 'plugins', 'office')
/** Les bibliothèques réellement embarquées dans `office.cjs`. */
const BUNDLED = ['exceljs', 'docx', 'pptxgenjs', 'pdf-lib', 'marked', 'zod']

const result = await Bun.build({
  entrypoints: [join(PLUGIN, 'runtime', 'officeRuntime.ts')],
  outdir: PLUGIN,
  naming: 'office.cjs',
  target: 'node',
  format: 'cjs',
  minify: true,
})
if (!result.success) {
  console.error('Échec de la construction du runtime bureautique :')
  for (const log of result.logs) console.error(`  ${log.message}`)
  process.exit(1)
}

// Mentions de licence : les paquets embarqués et toutes leurs dépendances,
// lus dans l'arbre installé plutôt que dans le fichier de verrouillage.
const notices = new Map<string, string>()
async function collect(directory: string): Promise<void> {
  let metadata: { name: string; version: string; dependencies?: Record<string, string> }
  try { metadata = JSON.parse(await readFile(join(directory, 'package.json'), 'utf8')) } catch { return }
  const key = `${metadata.name} ${metadata.version}`
  if (notices.has(key)) return
  notices.set(key, '')
  const texts: string[] = []
  for (const file of await readdir(directory, { withFileTypes: true })) {
    if (file.isFile() && /^(licen[sc]e|copying|notice)(\.|$)/i.test(file.name)) texts.push(await readFile(join(directory, file.name), 'utf8'))
  }
  notices.set(key, texts.join('\n'))
  for (const dependency of Object.keys(metadata.dependencies ?? {})) {
    const nested = join(directory, 'node_modules', dependency)
    await collect(await exists(nested) ? nested : join(ROOT, 'node_modules', dependency))
  }
}
async function exists(path: string): Promise<boolean> {
  try { await readdir(path); return true } catch { return false }
}
for (const name of BUNDLED) await collect(join(ROOT, 'node_modules', name))
await writeFile(join(PLUGIN, 'THIRD_PARTY_NOTICES.txt'),
  [...notices].sort(([a], [b]) => a.localeCompare(b)).map(([key, text]) => `\n## ${key}\n\n${text}`).join('\n'))

const size = result.outputs.reduce((total, output) => total + output.size, 0)
console.log(`office.cjs construit (${Math.round(size / 1024)} Ko), ${notices.size} paquets dans THIRD_PARTY_NOTICES.txt.`)
