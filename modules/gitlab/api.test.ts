import { expect, test } from 'bun:test'
import { instanceUrl, projectPath } from './api.js'

test('les requêtes restent sur l’instance HTTPS et dans le projet sélectionnés', () => {
  expect(instanceUrl('https://git.example/team/')).toBe('https://git.example/team')
  for (const value of ['https://user:pass@git.example', 'http://git.example', 'https://git.example?q=1']) expect(() => instanceUrl(value)).toThrow()
  expect(projectPath('team/subgroup/desktop')).toBe('team%2Fsubgroup%2Fdesktop')
  for (const value of ['group/../app', 'x?y', '/outside', 'group/%2fteam']) expect(() => projectPath(value)).toThrow()
})
