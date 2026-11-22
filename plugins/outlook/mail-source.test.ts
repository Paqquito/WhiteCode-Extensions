import { expect, test } from 'bun:test'
import { spawnSync } from 'node:child_process'
import { join } from 'node:path'

test.skipIf(process.platform !== 'win32')('folder browsing keeps online stores visible and rejects an unmounted selection', () => {
  const script = `
$ErrorActionPreference='Stop'
. $env:OUTLOOK_SOURCE
$folder=[pscustomobject]@{StoreID='00112233';EntryID='aabbccdd';Name='Pilote';FolderPath='Pilote';DefaultItemType=0}
$online=[pscustomobject]@{DisplayName='Pilote';StoreID=$folder.StoreID}
$online|Add-Member ScriptMethod GetRootFolder { $folder }
$offline=[pscustomobject]@{DisplayName='Archives'}
$offline|Add-Member ScriptMethod GetRootFolder { throw 'Unavailable store' }
$stores=[pscustomobject]@{Count=2}
$stores|Add-Member ScriptMethod Item { param($i) if($i -eq 1){return $offline};return $online }
$session=[pscustomobject]@{Stores=$stores}
$roots=@(Get-OutlookRoots $session)
$rejected=$false
try { Get-OutlookFolder $session 'ffffffff' $folder.EntryID inbox @() } catch { $rejected=$true }
@{roots=$roots;rejected=$rejected}|ConvertTo-Json -Depth 4 -Compress
`
  const result = spawnSync('powershell.exe', ['-NoProfile','-NonInteractive','-Command',script], {
    encoding:'utf8', windowsHide:true, timeout:10000,
    env:{...process.env,OUTLOOK_SOURCE:join(import.meta.dir,'skills/outlook/mail-source.ps1')},
  })
  expect(result.status).toBe(0)
  expect(JSON.parse(result.stdout)).toMatchObject({rejected:true,roots:[{unavailable:true,mail:false},{id:'aabbccdd',mail:true}]})
})
