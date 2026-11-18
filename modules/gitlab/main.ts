import { safeStorage } from 'electron'
import { randomUUID } from 'node:crypto'
import { readFile, rename, rm, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import type { ModuleHost } from '@whitecode/desktop-modules'
import { gitlabRequest, instanceUrl, projectPath, type Account } from './api.js'

export function register(host: ModuleHost) {
  const file = join(host.directory, 'account.json')
  async function account(): Promise<(Account & { encrypted: string }) | null> {
    try { return JSON.parse(await readFile(file, 'utf8')) }
    catch (error) { if ((error as NodeJS.ErrnoException).code === 'ENOENT') return null; throw error }
  }
  host.registerIpc('status', async () => {
    const saved = await account()
    return saved && { instance: saved.instance, username: saved.username }
  })
  host.registerIpc('connect', async input => {
    const { instance, token } = (input ?? {}) as { instance?: string; token?: string }
    const origin = instanceUrl(instance)
    if (typeof token !== 'string' || !token || /\s/.test(token) || token.length > 4096) throw Error('Jeton personnel invalide.')
    if (!safeStorage.isEncryptionAvailable() || (process.platform === 'linux' && safeStorage.getSelectedStorageBackend() === 'basic_text')) throw Error('Le coffre du système est indisponible.')
    const user = await gitlabRequest(origin, token, 'user') as { username?: string }
    if (typeof user.username !== 'string' || !user.username) throw Error('Compte GitLab introuvable.')
    const tmp = file + '.' + randomUUID()
    try {
      await writeFile(tmp, JSON.stringify({ instance: origin, username: user.username, encrypted: safeStorage.encryptString(token).toString('base64') }), { mode: 0o600, flag: 'wx' })
      await rename(tmp, file)
    } finally { await rm(tmp, { force: true }) }
    return { instance: origin, username: user.username }
  })
  host.registerIpc('disconnect', async () => { await rm(file, { force: true }) })
  host.registerIpc('requests', async input => {
    const project = projectPath(input)
    const saved = await account()
    if (!saved) throw Error('Connectez votre compte dans Paramètres → GitLab.')
    const token = safeStorage.decryptString(Buffer.from(saved.encrypted, 'base64'))
    return gitlabRequest(saved.instance, token, 'projects/' + project + '/merge_requests?state=opened&per_page=50')
  })
}
