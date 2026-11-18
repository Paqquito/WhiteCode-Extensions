import { confirmedWrite, type WriteForm, type GitLabCall } from './writes.js'
import { dialog, safeStorage } from 'electron'
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
  async function connection() {
    const saved = await account()
    if (!saved) throw Error('Connectez votre compte dans Paramètres → GitLab.')
    const token = safeStorage.decryptString(Buffer.from(saved.encrypted, 'base64'))
    const call: GitLabCall = (path, method, body) => gitlabRequest(saved.instance, token, path, method, body)
    return { saved, call }
  }
  host.registerIpc('requests', async input => {
    const { call } = await connection()
    return call('projects/' + projectPath(input) + '/merge_requests?state=opened&per_page=50')
  })
  host.registerIpc('project', async project => {
    const { saved, call } = await connection()
    const rows = await call('projects/' + projectPath(project) + '/merge_requests?state=opened&per_page=50')
    return { username: saved.username, instance: saved.instance, rows }
  })
  host.registerIpc('detail', async input => {
    const { project, iid } = (input ?? {}) as WriteForm
    if (!Number.isSafeInteger(iid) || iid! <= 0) throw Error('Numéro de demande invalide.')
    const { call } = await connection()
    return call('projects/' + projectPath(project) + '/merge_requests/' + iid)
  })
  host.registerIpc('choices', async input => {
    const { project, iid, pipeline, kind } = (input ?? {}) as WriteForm & { kind: string }
    if (!Number.isSafeInteger(iid) || iid! < 1) throw Error('Demande invalide.')
    const { call } = await connection()
    const prefix = 'projects/' + projectPath(project)
    if (kind === 'jobs' && Number.isSafeInteger(pipeline) && pipeline! > 0) {
      const jobs = await call(prefix + '/pipelines/' + pipeline + '/jobs?per_page=100') as Array<{ id: number; name: string; status: string }>
      return jobs.map(j => ({ id: j.id, label: '#' + j.id + ' · ' + j.name + ' · ' + j.status }))
    }
    if (kind !== 'discussions') throw Error('Sélection GitLab invalide.')
    const discussions = await call(prefix + '/merge_requests/' + iid + '/discussions?per_page=100') as Array<{ id: string; notes: Array<{ resolvable: boolean; resolved: boolean; body: string }> }>
    return discussions.flatMap(d => {
      const note = d.notes.find(n => n.resolvable && !n.resolved)
      return note ? [{ id: d.id, label: note.body.slice(0, 100) }] : []
    })
  })
  let writing = false
  host.registerIpc('write', async value => {
    if (writing) throw Error('Une écriture GitLab est déjà en cours.')
    writing = true
    try {
      const { saved, call } = await connection()
      const input = value as WriteForm
      if (saved.instance !== input?.instance || saved.username !== input.username) throw Error('Compte GitLab modifié. Actualisez le projet.')
      return await confirmedWrite(call, input, async (label, detail) => {
        const { response } = await dialog.showMessageBox({ type: 'question', title: 'GitLab', message: label,
          detail, buttons: ['Annuler', 'Confirmer'], defaultId: 0, cancelId: 0, noLink: true })
        return response === 1
      }, async () => {
        const current = await account()
        if (current?.encrypted !== saved.encrypted || current.instance !== saved.instance || current.username !== saved.username) throw Error('Compte GitLab modifié. Actualisez le projet.')
      })
    } finally { writing = false }
  })
}
