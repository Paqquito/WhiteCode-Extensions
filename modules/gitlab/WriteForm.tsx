import * as React from 'react'
import type { ModuleCall } from '@whitecode/desktop-modules'
import type { Account, MergeRequest } from './api.js'
import type { Action, WriteForm } from './writes.js'

export type Detail = MergeRequest & {
  sha: string; updated_at: string; draft?: boolean; labels?: string[]
  reviewers?: Array<{ id: number }>; head_pipeline?: { id: number; status: string; sha: string }
}
const actions: Array<[Action, string]> = [
  ['edit', 'Modifier la demande'], ['comment', 'Commenter'], ['line', 'Commenter une ligne'],
  ['resolve', 'Résoudre une discussion'], ['merge', 'Fusionner'], ['retryPipeline', 'Relancer le pipeline'],
  ['cancelPipeline', 'Annuler le pipeline'], ['playJob', 'Lancer un job manuel'], ['retryJob', 'Relancer un job'], ['cancelJob', 'Annuler un job'],
]
export function GitLabWriteForm({ call, project, account, mr, onSaved, onClose }: {
  call: ModuleCall; project: string; account: Account; mr?: Detail
  onSaved(): Promise<void>; onClose(): void
}) {
  const [action, select] = React.useState<Action>(mr ? 'comment' : 'create')
  const [busy, setBusy] = React.useState(false), [message, inform] = React.useState('')
  const [choices, setChoices] = React.useState<Array<{ id: string | number; label: string }>>([])
  const jobs = ['playJob', 'retryJob', 'cancelJob'].includes(action)
  React.useEffect(() => {
    let active = true
    setChoices([])
    if (mr && (jobs || action === 'resolve')) {
      void call('choices', { project, iid: mr.iid, pipeline: jobs ? mr.head_pipeline?.id : undefined, kind: jobs ? 'jobs' : 'discussions' })
        .then(value => { if (active) setChoices(value as typeof choices) }, error => { if (active) inform(String(error)) })
    }
    return () => { active = false }
  }, [action, mr, project])
  const fields = action === 'edit' || action === 'create'
  async function submit(e: React.FormEvent<HTMLFormElement>) {
    e.preventDefault()
    if (busy) return
    const data = new FormData(e.currentTarget)
    const string = (name: string) => data.has(name) ? String(data.get(name)) : undefined
    const number = (name: string) => string(name) ? Number(string(name)) : undefined
    const input: WriteForm = { action, project, ...account, iid: mr?.iid, sha: mr?.sha, updatedAt: mr?.updated_at,
      source: string('source'), target: string('target'), title: string('title'), description: string('description'),
      draft: data.has('draft'), labels: string('labels'), reviewers: string('reviewers'), body: string('body'),
      oldPath: string('oldPath'), newPath: string('newPath'), oldLine: number('oldLine'), newLine: number('newLine'),
      discussion: string('discussion'), job: number('job'), pipeline: mr?.head_pipeline?.id }
    setBusy(true); inform('')
    try {
      const result = await call('write', input) as { performed: boolean }
      if (result.performed) await onSaved()
      else inform('Action annulée.')
    } catch (error) { inform(String(error)) }
    finally { setBusy(false) }
  }
  return <form onSubmit={submit} className="settings-form"><fieldset disabled={busy}>
    <legend>{mr ? 'Demande !' + mr.iid : 'Créer une demande de fusion'} · {project}</legend>
    {mr && <label>Action<select value={action} onChange={e => { select(e.target.value as Action); inform('') }}>
      {actions.filter(([key]) => mr.head_pipeline || !/Job|Pipeline/.test(key)).map(([key, label]) => <option key={key} value={key}>{label}</option>)}
    </select></label>}
    {action === 'create' && <><label>Branche source<input name="source" required /></label><label>Branche cible<input name="target" required /></label></>}
    {fields && <><label>Titre<input name="title" defaultValue={mr?.title.replace(/^(?:Draft|WIP):\s*/i, '')} maxLength={255} required /></label>
      <label>Description<textarea name="description" defaultValue={mr?.description || ''} maxLength={65000} /></label>
      <label><input name="draft" type="checkbox" defaultChecked={mr ? mr.draft : true} /> Brouillon</label>
      <label>Étiquettes séparées par des virgules<input name="labels" defaultValue={mr?.labels?.join(',')} /></label>
      <label>Identifiants des relecteurs séparés par des virgules<input name="reviewers" defaultValue={mr?.reviewers?.map(r => r.id).join(',')} /></label></>}
    {(action === 'comment' || action === 'line') && <label>Commentaire<textarea name="body" required maxLength={65000} /></label>}
    {action === 'line' && <><label>Chemin avant modification<input name="oldPath" required /></label><label>Chemin après modification<input name="newPath" required /></label>
      <label>Ancien numéro de ligne<input name="oldLine" type="number" min="1" /></label><label>Nouveau numéro de ligne<input name="newLine" type="number" min="1" /></label></>}
    {(action === 'resolve' || jobs) && <label>{jobs ? 'Job' : 'Discussion ouverte'}<select required name={jobs ? 'job' : 'discussion'} key={action} defaultValue="">
      <option value="">Sélectionner…</option>{choices.map(c => <option key={c.id} value={c.id}>{c.label}</option>)}</select></label>}
    {mr?.head_pipeline && /Job|Pipeline/.test(action) && <p>Pipeline #{mr.head_pipeline.id} · {mr.head_pipeline.status}</p>}
    <button type="submit">Vérifier et confirmer…</button> <button type="button" onClick={onClose}>Fermer</button>
  </fieldset>{message && <p role="status">{message}</p>}</form>
}
