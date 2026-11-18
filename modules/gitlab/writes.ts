import { projectPath } from './api.js'

export type GitLabCall = (path: string, method?: 'GET' | 'POST' | 'PUT', body?: Record<string, unknown>) => Promise<unknown>
export type Action = 'create' | 'edit' | 'comment' | 'line' | 'resolve' | 'merge' | 'retryPipeline' | 'cancelPipeline' | 'retryJob' | 'cancelJob' | 'playJob'
export interface WriteForm {
  action: Action
  project: string
  instance: string
  username: string
  iid?: number
  sha?: string
  updatedAt?: string
  title?: string
  description?: string
  draft?: boolean
  source?: string
  target?: string
  labels?: string
  reviewers?: string
  body?: string
  oldPath?: string
  newPath?: string
  oldLine?: number
  newLine?: number
  discussion?: string
  pipeline?: number
  job?: number
}
interface Mutation {
  label: string
  path: string
  method: 'PUT' | 'POST'
  body: Record<string, unknown>
  revision: string
  targetRevision?: string
}
function object(value: unknown): Record<string, any> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw Error('Réponse GitLab invalide.')
  return value as Record<string, any>
}
function id(value: unknown): number {
  if (typeof value !== 'number' || !Number.isSafeInteger(value) || value < 1) throw Error('Identifiant GitLab invalide.')
  return value
}
function text(value: unknown, max: number, empty = false): string {
  if (typeof value !== 'string' || (!empty && !value.trim()) || value.length > max || value.includes('\0')) throw Error('Valeur absente ou trop longue.')
  return value
}
function hash(value: unknown): string {
  if (typeof value !== 'string' || !/^(?:[a-f0-9]{40}|[a-f0-9]{64})$/.test(value)) throw Error('Commit GitLab invalide.')
  return value
}
function requireMatch(matches: boolean) {
  if (!matches) throw Error('La demande, le commit ou le pipeline a changé. Actualisez le panneau.')
}
function mrFields(input: WriteForm) {
  const title = text(input.title, 255).replace(/^(?:draft|wip):\s*/i, '').trim()
  if (!title || typeof input.draft !== 'boolean') throw Error('Titre ou état du brouillon invalide.')
  const body: Record<string, unknown> = { title: (input.draft ? 'Draft: ' : '') + title, description: text(input.description, 65000, true) }
  if (input.labels !== undefined) body.labels = text(input.labels, 2000, true)
  if (input.reviewers !== undefined) {
    const values = text(input.reviewers, 512, true).trim()
    const ids = values ? values.split(',').map(s => id(Number(s.trim()))) : []
    body.reviewer_ids = ids
  }
  return body
}
export async function prepareWrite(call: GitLabCall, input: WriteForm): Promise<Mutation> {
  const project = 'projects/' + projectPath(input.project)
  const collection = project + '/merge_requests'
  if (input.action === 'create') {
    const source = text(input.source, 255), target = text(input.target, 255)
    if (source === target) throw Error('Les deux branches doivent être différentes.')
    const readBranch = async (name: string) => object(await call(project + '/repository/branches/' + encodeURIComponent(name)))
    const [branch, destination] = await Promise.all([readBranch(source), readBranch(target)])
    return { label: 'Créer la demande de fusion', path: collection, method: 'POST', revision: hash(branch.commit?.id), targetRevision: hash(destination.commit?.id),
      body: { ...mrFields(input), source_branch: source, target_branch: target } }
  }
  const path = collection + '/' + id(input.iid)
  const mr = object(await call(path))
  const revision = hash(input.sha)
  requireMatch(mr.state === 'opened' && mr.sha === revision && mr.iid === input.iid)
  switch (input.action) {
    case 'edit':
      requireMatch(typeof input.updatedAt === 'string' && mr.updated_at === input.updatedAt)
      return { label: 'Modifier la demande de fusion', path, method: 'PUT', body: mrFields(input), revision }
    case 'comment':
      return { label: 'Publier le commentaire', path: path + '/notes', method: 'POST', body: { body: text(input.body, 65000) }, revision }
    case 'line': {
      const refs = object(mr.diff_refs)
      if (input.newLine === undefined && input.oldLine === undefined) throw Error('Une ligne du diff est requise.')
      const position = { position_type: 'text', base_sha: hash(refs.base_sha), start_sha: hash(refs.start_sha), head_sha: hash(refs.head_sha),
        old_path: text(input.oldPath, 1024), new_path: text(input.newPath, 1024),
        ...(input.newLine === undefined ? {} : { new_line: id(input.newLine) }), ...(input.oldLine === undefined ? {} : { old_line: id(input.oldLine) }) }
      requireMatch(position.head_sha === revision)
      return { label: 'Ouvrir une discussion sur le diff', path: path + '/discussions', method: 'POST', body: { body: text(input.body, 65000), position }, revision }
    }
    case 'resolve': {
      const discussion = text(input.discussion, 128)
      if (!/^[a-zA-Z0-9_-]+$/.test(discussion)) throw Error('Identifiant de discussion invalide.')
      const target = path + '/discussions/' + discussion
      const state = object(await call(target))
      requireMatch(state.id === discussion && Array.isArray(state.notes) && state.notes.some((n: any) => n.resolvable && !n.resolved))
      return { label: 'Résoudre la discussion', path: target, method: 'PUT', body: { resolved: true }, revision }
    }
    case 'merge': {
      if (mr.draft || mr.work_in_progress || mr.has_conflicts) throw Error('Fusion indisponible pour cette demande.')
      const settings = object(await call(project))
      const method = text(settings.merge_method, 40)
      const squash = settings.squash_option === 'always' || (settings.squash_option !== 'never' && Boolean(mr.squash))
      return { label: 'Fusionner (' + method + (squash ? ', squash' : '') + ')', path: path + '/merge', method: 'PUT',
        body: { sha: revision, squash, should_remove_source_branch: false }, revision }
    }
    case 'retryPipeline': case 'cancelPipeline': case 'retryJob': case 'playJob': case 'cancelJob': {
      const pid = id(input.pipeline)
      requireMatch(mr.head_pipeline?.id === pid && mr.head_pipeline.sha === revision)
      const isJob = ['retryJob', 'playJob', 'cancelJob'].includes(input.action)
      const target = isJob ? project + '/jobs/' + id(input.job) : project + '/pipelines/' + pid
      const item = object(await call(target))
      requireMatch(isJob ? item.pipeline?.id === pid && item.id === input.job && item.commit?.id === revision : item.id === pid && item.sha === revision)
      const verb = input.action === 'playJob' ? 'play' : input.action.startsWith('retry') ? 'retry' : 'cancel'
      const states = verb === 'play' ? ['manual'] : verb === 'retry' ? ['failed', 'canceled', ...(isJob ? ['success'] : [])] : ['created', 'waiting_for_resource', 'preparing', 'pending', 'running', 'scheduled', 'manual']
      requireMatch(states.includes(item.status))
      const label = ({ play: 'Lancer', retry: 'Relancer', cancel: 'Annuler' })[verb] + (isJob ? ' le job ' + input.job : ' le pipeline ' + pid)
      return { label, path: target + '/' + verb, method: 'POST', body: {}, revision }
    }
    default: throw Error('Écriture GitLab inconnue.')
  }
}
export async function confirmedWrite(call: GitLabCall, input: WriteForm, confirm: (title: string, detail: string) => Promise<boolean>, checkAccount: () => Promise<void>) {
  const initial = await prepareWrite(call, input)
  const detail = [input.username + ' · ' + input.instance, input.project + (input.iid ? ' !' + input.iid : ''),
    initial.revision + (initial.targetRevision ? ' → ' + initial.targetRevision : ''),
    JSON.stringify(initial.body, null, 2)].join('\n')
  if (!await confirm(initial.label, detail)) return { performed: false }
  await checkAccount()
  const current = await prepareWrite(call, input)
  requireMatch(JSON.stringify(initial) === JSON.stringify(current))
  await checkAccount()
  await call(current.path, current.method, current.body)
  return { performed: true }
}
