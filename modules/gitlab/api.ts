export interface Account {
  instance: string
  username: string
}
export interface MergeRequest {
  iid: number
  title: string
  state: string
  source_branch: string
  target_branch: string
  description?: string
}
export function instanceUrl(input: unknown): string {
  if (typeof input !== 'string') throw Error('Adresse HTTPS de l’instance attendue.')
  const url = new URL(input)
  if (url.protocol !== 'https:' || url.username || url.password || url.search || url.hash) throw Error('Adresse HTTPS sans identifiants attendue.')
  return url.href.replace(/\/+$/, '')
}
export function projectPath(input: unknown): string {
  if (typeof input !== 'string' || !/^[\w.-]+(?:\/[\w.-]+)+$/.test(input) || input.split('/').some(p => p === '..' || p === '.')) throw Error('Projet attendu sous la forme groupe/projet.')
  return encodeURIComponent(input)
}
export async function gitlabRequest(instance: string, token: string, path: string): Promise<unknown> {
  const res = await fetch(instanceUrl(instance) + '/api/v4/' + path, {
    headers: { 'PRIVATE-TOKEN': token }, redirect: 'error', signal: AbortSignal.timeout(15000),
  })
  if (!res.ok) throw Error('GitLab a répondu HTTP ' + res.status + '. Vérifiez le projet et les droits de votre jeton.')
  return res.json()
}
