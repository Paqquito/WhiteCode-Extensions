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

import { writeFile } from 'node:fs/promises'
import { dependencyNotices } from './dependency-notices'
import { join } from 'node:path'

const ROOT = join(import.meta.dir, '..')
const PLUGIN = join(ROOT, 'plugins', 'office')
/** Les bibliothèques réellement embarquées dans `office.cjs`. */
const BUNDLED = ['exceljs', 'docx', 'pptxgenjs', 'pdf-lib', 'marked', 'zod', 'jszip']
const notices = dependencyNotices(ROOT, BUNDLED)

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
await writeFile(join(PLUGIN, 'THIRD_PARTY_NOTICES.txt'), notices)

const size = result.outputs.reduce((total, output) => total + output.size, 0)
console.log(`office.cjs construit (${Math.round(size / 1024)} Ko) avec ses notices de dépendances.`)
