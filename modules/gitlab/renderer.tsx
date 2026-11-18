import { GitLabWriteForm, type Detail } from './WriteForm.js'
import * as React from 'react'
import type { ModuleViews, ModuleViewProps, ModulePanelProps } from '@whitecode/desktop-modules'
import type { Account, MergeRequest } from './api.js'

function Settings({ call }: ModuleViewProps) {
  const [account, setAccount] = React.useState<Account | null>(null)
  const [instance, setInstance] = React.useState('')
  const [token, setToken] = React.useState('')
  const [error, setError] = React.useState('')
  const [busy, setBusy] = React.useState(false)
  React.useEffect(() => {
    let active = true
    void call('status').then(value => {
      if (!active) return
      const saved = value as Account | null
      setAccount(saved)
      if (saved) setInstance(saved.instance)
    }, reason => { if (active) setError(String(reason)) })
    return () => { active = false }
  }, [])
  async function connect(event: React.FormEvent) {
    event.preventDefault()
    setBusy(true); setError('')
    try {
      setAccount(await call('connect', { instance, token }) as Account)
      setToken('')
    } catch (reason) { setError(String(reason)) }
    finally { setBusy(false) }
  }
  return <section><h2>GitLab</h2>
    {account && <p>Compte connecté · {account.username} ({account.instance})</p>}
    <form onSubmit={connect} className="settings-form">
      <label>Instance GitLab<input type="url" required placeholder="https://gitlab.com" value={instance} onChange={e => setInstance(e.target.value)} /></label>
      <label>Jeton personnel<input type="password" required autoComplete="off" value={token} onChange={e => setToken(e.target.value)} /></label>
      <p>Le droit <code>read_api</code> permet la consultation. Les écritures demandent le droit <code>api</code> et les permissions du projet.</p>
      <button disabled={busy}>Vérifier et enregistrer</button>
    </form>
    {account && <button disabled={busy} onClick={async () => {
      setBusy(true); setError('')
      try { await call('disconnect'); setAccount(null); setToken('') }
      catch (reason) { setError(String(reason)) }
      finally { setBusy(false) }
    }}>Déconnecter le compte</button>}
    {error && <p role="alert">{error}</p>}
  </section>
}

function Panel({ call, onClose }: ModulePanelProps) {
  const [project, setProject] = React.useState('')
  const [account, setAccount] = React.useState<Account | null>(null)
  const [editor, setEditor] = React.useState<Detail | 'create' | null>(null)
  const [notice, setNotice] = React.useState('')
  async function load() {
    const result = await call('project', project.trim()) as Account & { rows: MergeRequest[] }
    setAccount({ instance: result.instance, username: result.username }); setRows(result.rows)
  }
  const [rows, setRows] = React.useState<MergeRequest[] | null>(null)
  const [error, setError] = React.useState('')
  const [busy, setBusy] = React.useState(false)
  return <section className="panel"><header className="panel-header"><strong>Demandes de fusion GitLab</strong>
    <button onClick={onClose} aria-label="Fermer GitLab">Fermer</button></header>
    <div className="panel-body"><form onSubmit={async event => {
      event.preventDefault(); setBusy(true); setError(''); setRows(null)
      try { setEditor(null); setNotice(''); await load() }
      catch (reason) { setError(String(reason)) }
      finally { setBusy(false) }
    }}><label>Projet<input required placeholder="groupe/projet" value={project} onChange={e => { setProject(e.target.value); setRows(null); setEditor(null); setNotice('') }} disabled={busy} /></label>
    <button disabled={busy}>Consulter les demandes ouvertes</button></form>
    {error && <p role="alert">{error}</p>}
    {rows && account && <p>{account.username} · {account.instance} <button disabled={busy || !!editor} onClick={() => { setNotice(''); setEditor('create') }}>Créer une demande</button></p>}
    {notice && <p role="status">{notice}</p>}
    {editor && account && <GitLabWriteForm key={editor === 'create' ? 'create' : editor.iid} call={call} project={project.trim()} account={account}
      mr={editor === 'create' ? undefined : editor} onClose={() => setEditor(null)} onSaved={async () => {
        setEditor(null); setNotice('Action effectuée dans GitLab.'); setBusy(true)
        try { await load() } catch (reason) { setError('Actualisation impossible : ' + reason) } finally { setBusy(false) }
      }} />}
    {rows?.length === 0 && <p>Aucune demande ouverte.</p>}
    {rows?.map(row => <details key={row.iid}><summary>!{row.iid} · {row.title}</summary>
      <p>{row.source_branch} → {row.target_branch}</p><p>{row.description}</p><button disabled={busy || !!editor} onClick={async () => {
        setBusy(true); setError(''); setNotice('')
        try { setEditor(await call('detail', { project, iid: row.iid }) as Detail) }
        catch (reason) { setError(String(reason)) } finally { setBusy(false) }
      }}>Actions sur la demande</button></details>)}
    </div>
  </section>
}

export const views: ModuleViews = { label: 'GitLab', Panel, Settings }
