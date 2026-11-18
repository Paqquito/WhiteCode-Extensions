import { expect, test } from 'bun:test'
import { confirmedWrite, prepareWrite, type GitLabCall, type WriteForm } from './writes.js'

const sha = 'a'.repeat(40)
const common = { project: 'group/app', instance: 'https://git.example', username: 'tester', iid: 2, sha, updatedAt: 'one' }
const branch = { commit: { id: sha } }
function fixture(jobStatus = 'manual', pipelineStatus = 'failed') {
  const mr = { iid: 2, state: 'opened', sha, updated_at: 'one', head_pipeline: { id: 7, sha }, diff_refs: { base_sha: sha, start_sha: sha, head_sha: sha } }
  const writes: Array<{ path: string; method: string; body: unknown }> = []
  const call: GitLabCall = async (path, method, body) => {
    if (method && method !== 'GET') { writes.push({ path, method, body }); return {} }
    if (path.includes('/branches/')) return branch
    if (path.endsWith('/merge_requests/2')) return mr
    if (path.endsWith('/discussions/ABC')) return { id: 'ABC', notes: [{ resolvable: true, resolved: false }] }
    if (path.endsWith('/pipelines/7')) return { id: 7, sha, status: pipelineStatus }
    if (path.endsWith('/jobs/4')) return { id: 4, pipeline: { id: 7 }, commit: { id: sha }, status: jobStatus }
    if (path === 'projects/group%2Fapp') return { merge_method: 'ff', squash_option: 'never' }
    throw Error('Chemin inattendu ' + path)
  }
  return { mr, writes, call }
}
test('chaque action est construite dans le projet et confirmée avant sa première écriture', async () => {
  const cases: Array<[Partial<WriteForm> & Pick<WriteForm, 'action'>, string, string]> = [
    [{ action: 'create', source: 'feature/one', target: 'main', title: 'Titre', description: '', draft: true }, 'POST', '/merge_requests'],
    [{ action: 'edit', title: 'Titre', description: 'Révisée', draft: false, reviewers: '1,2' }, 'PUT', '/merge_requests/2'],
    [{ action: 'comment', body: 'Note' }, 'POST', '/notes'], [{ action: 'resolve', discussion: 'ABC' }, 'PUT', '/discussions/ABC'],
    [{ action: 'line', oldPath: 'a.ts', newPath: 'a.ts', newLine: 3, body: 'Note' }, 'POST', '/discussions'],
    [{ action: 'merge' }, 'PUT', '/merge'], [{ action: 'retryPipeline', pipeline: 7 }, 'POST', '/pipelines/7/retry'],
    [{ action: 'cancelPipeline', pipeline: 7 }, 'POST', '/pipelines/7/cancel'],
    [{ action: 'retryJob', pipeline: 7, job: 4 }, 'POST', '/jobs/4/retry'],
    [{ action: 'cancelJob', pipeline: 7, job: 4 }, 'POST', '/jobs/4/cancel'],
    [{ action: 'playJob', pipeline: 7, job: 4 }, 'POST', '/jobs/4/play'],
  ]
  for (const [args, method, suffix] of cases) {
    const { call, writes } = fixture(args.action === 'retryJob' ? 'failed' : 'manual', args.action === 'cancelPipeline' ? 'running' : 'failed')
    const result = await confirmedWrite(call, { ...common, ...args }, async (_, detail) => {
      expect(writes).toHaveLength(0); expect(detail).toContain(sha); return true
    }, async () => {})
    expect(result.performed).toBe(true); expect(writes).toHaveLength(1)
    expect(writes[0]!.method).toBe(method); expect(writes[0]!.path.endsWith(suffix)).toBe(true)
    if (args.action === 'merge') expect(writes[0]!.body).toMatchObject({ sha })
  }
})
test('annulation, changement de compte ou de révision et identifiants invalides bloquent la mutation', async () => {
  const { call, writes, mr } = fixture()
  const form: WriteForm = { ...common, action: 'comment', body: 'Note' }
  expect(await confirmedWrite(call, form, async () => false, async () => {})).toEqual({ performed: false })
  await expect(confirmedWrite(call, form, async () => true, async () => { throw Error('Compte modifié') })).rejects.toThrow()
  await expect(prepareWrite(call, { ...form, iid: -1 })).rejects.toThrow()
  await expect(prepareWrite(call, { ...form, action: 'playJob', pipeline: 8, job: 4 })).rejects.toThrow()
  await expect(confirmedWrite(call, form, async () => { mr.sha = 'b'.repeat(40); return true }, async () => {})).rejects.toThrow()
  expect(writes).toEqual([])
})

test('une branche cible déplacée pendant la confirmation empêche la création', async () => {
  const { call, writes } = fixture()
  let target = sha
  const branches: GitLabCall = (path, method, body) => path.endsWith('/branches/main')
    ? Promise.resolve({ commit: { id: target } }) : call(path, method, body)
  await expect(confirmedWrite(branches, { ...common, action: 'create', source: 'feature', target: 'main',
    title: 'Ajout', draft: true, description: '' }, async () => { target = 'c'.repeat(40); return true }, async () => {})).rejects.toThrow()
  expect(writes).toHaveLength(0)
})
