import { existsSync, readFileSync, readdirSync, realpathSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'

const read = (path: string) => readFileSync(path, 'utf8').replace(/\r\n/g, '\n').trim()
const mit = read(join(import.meta.dir, '../LICENSE')).split('\n\n').slice(2).join('\n\n')
const isc = 'Permission to use, copy, modify, and/or distribute this software for any purpose with or without fee is hereby granted, provided that the above copyright notice and this permission notice appear in all copies.\n\nTHE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.'

function licenseText(dir: string, pkg: { name: string; version: string; license?: string; author?: string | { name: string } }): string {
  const texts = readdirSync(dir, { withFileTypes: true })
    .filter(f => f.isFile() && /^(?:licen[cs]es?|copying|notices?|third[-_ ]party(?:[-_ ]notices)?)(?:[.\-_ ]|$)/i.test(f.name))
    .map(f => f.name).sort().map(f => read(join(dir, f))).filter(Boolean)
  if (pkg.name === 'node-pty') texts.push(read(join(dir, 'deps/winpty/LICENSE')))
  if (pkg.name === '@pdf-lib/standard-fonts') {
    const original = read(join(dir, 'README.md')).split("## Original Repo's License")[1]
    if (!original?.trim()) throw Error('Original font metrics attribution missing')
    texts.push(original.trim())
  }
  if (pkg.name === 'pako') {
    const source = read(join(dir, 'lib/zlib/trees.js'))
    // pako also incorporates zlib's C implementation, whose notice lives in the source.
    const notice = source.match(/(?:\/\*[\s\S]*?\*\/|(?:^\/\/[^\n]*\n)+)/gm)?.find(s => /Jean-loup|Gailly|Adler/.test(s))
    if (!notice || !/altered|misrepresented/i.test(notice)) throw Error('Missing zlib notice: ' + pkg.version)
    texts.push(notice)
  }
  if (texts.length) return texts.join('\n\n')
  const key = pkg.name + '@' + pkg.version
  if (key === 'saxes@5.0.1') return read(join(import.meta.dir, 'licenses/saxes-5.0.1.txt'))
  if (key === 'hash.js@1.1.7' || key === 'isarray@1.0.0') {
    const section = read(join(dir, 'README.md')).split(/^#{1,6}\s+license\s*$/im)[1]?.split(/^#{1,6}\s/m)[0]?.trim()
    if (!section || !/permission is hereby granted/i.test(section)) throw Error('Incomplete README license: ' + key)
    return section.replaceAll('&lt;', '<').replaceAll('&gt;', '>')
  }
  const author = typeof pkg.author === 'string' ? pkg.author : pkg.author?.name
  let license = pkg.license
  let source = 'License and author as declared in the upstream package.json.'
  if (key === 'buffers@0.1.1') {
    license = 'MIT/X11'
    source = 'Upstream license declaration by James Halliday: https://github.com/bitpay/node-buffers/commit/1b745ee35d33eb166e15ef1866073a07c6d7de87'
  }
  // These releases identify their license and author but omit a complete license file.
  const declarations = ['binary@0.3.0', 'buffers@0.1.1', 'chainsaw@0.1.0', 'https@1.0.0', 'standardwebhooks@1.1.1', 'standardwebhooks@1.3.0']
  if (declarations.includes(key) && author && (license === 'MIT' || license === 'MIT/X11' || license === 'ISC')) {
    return source + '\n' + license + ' — ' + author + '\n\n' + (license === 'ISC' ? isc : mit)
  }
  throw Error('Missing upstream license text for ' + key)
}

export function dependencyNotices(root: string, packages: string[]): string {
  const entries = new Map<string, string>()
  const visited = new Set<string>()
  const modulesRoot = realpathSync(join(root, 'node_modules'))
  function collect(dir: string) {
    dir = realpathSync(dir)
    if (visited.has(dir)) return
    visited.add(dir)
    const pkg = JSON.parse(read(join(dir, 'package.json')))
    const key = pkg.name + ' ' + pkg.version
    const license = licenseText(dir, pkg)
    if (entries.has(key) && entries.get(key) !== license) throw Error('Conflicting upstream notices: ' + key)
    entries.set(key, license)
    for (const dep of Object.keys({ ...pkg.dependencies, ...pkg.optionalDependencies })) {
      let parent = dir
      let nested: string | undefined
      while (parent !== modulesRoot && dirname(parent) !== parent) {
        const candidate = join(parent, 'node_modules', dep)
        if (existsSync(join(candidate, 'package.json'))) { nested = candidate; break }
        parent = dirname(parent)
      }
      const child = nested ?? join(modulesRoot, dep)
      if (pkg.optionalDependencies?.[dep] && !existsSync(join(child, 'package.json'))) continue
      collect(child)
    }
  }
  for (const name of packages) collect(join(modulesRoot, name))
  const sections = [...entries.keys()].sort().map(k => '## ' + k + '\n\n' + entries.get(k))
  const supplements = join(root, 'licenses')
  if (existsSync(supplements)) {
    for (const f of readdirSync(supplements).sort()) sections.push(read(join(supplements, f)))
  }
  return sections.join('\n\n') + '\n'
}

if (import.meta.main) {
  const [dir, ...extra] = process.argv.slice(2)
  if (!dir) throw Error('Usage: bun dependency-notices.ts <repository> [additional-runtime-package ...]')
  const root = resolve(dir)
  const pkg = JSON.parse(read(join(root, 'package.json')))
  writeFileSync(join(root, 'THIRD_PARTY_NOTICES.txt'), dependencyNotices(root, [...Object.keys(pkg.dependencies ?? {}), ...extra]))
}
