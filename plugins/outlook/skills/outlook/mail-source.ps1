# Shared by interactive Outlook reads and consumers of the read-only mail stream.
function Get-OutlookFolderInfo($folder) {
    @{store=[string]$folder.StoreID; id=[string]$folder.EntryID; name=[string]$folder.Name; path=[string]$folder.FolderPath; mail=($folder.DefaultItemType -eq 0)}
}

function Get-OutlookRoots($session) {
    $stores = $session.Stores
    for ($i=1; $i -le $stores.Count; $i++) {
        $label = "Banque Outlook $i"
        try {
            $store = $stores.Item($i)
            $label = $store.DisplayName
            Get-OutlookFolderInfo ($store.GetRootFolder())
        } catch {
            @{store='';id="offline-$i";name="$label (indisponible)";path='';mail=$false;unavailable=$true}
        }
    }
}

function Get-OutlookFolder($session, [string]$storeId, [string]$folderId, [string]$default, [string[]]$children) {
    if (!$storeId -and !$folderId) {
        $kind = if ($default -eq 'sent') { 5 } else { 6 }
        $source = $session.GetDefaultFolder($kind)
        foreach ($part in $children) { $source = $source.Folders.Item($part) }
        return ,$source
    }
    if ($children.Count -or $storeId -notmatch '^[a-fA-F0-9]{8,8192}$' -or $folderId -notmatch '^[a-fA-F0-9]{8,8192}$') {
        throw 'Supply both StoreId and FolderId for the selected folder.'
    }
    # GetFolderFromID alone can reopen a detached PST.
    $stores = $session.Stores
    for ($i=1; $i -le $stores.Count; $i++) {
        try { $key = $stores.Item($i).StoreID } catch { continue }
        if ($key -ne $storeId) { continue }
        $folder = $session.GetFolderFromID($folderId,$storeId)
        if ($folder -and $folder.StoreID -eq $storeId -and $folder.EntryID -eq $folderId) { return ,$folder }
        break
    }
    throw 'Dossier ou banque Outlook indisponible. Reconnectez la source et resynchronisez.'
}

function Get-OutlookMailDateField($folder) {
    try { $sent = $folder.Store.GetDefaultFolder(5) } catch { $sent = $null }
    if ($sent -and ($folder.EntryID -eq $sent.EntryID -or $folder.FolderPath.StartsWith(($sent.FolderPath+'\'),[StringComparison]::OrdinalIgnoreCase))) { return 'SentOn' }
    return 'ReceivedTime'
}

function Get-OutlookPeriodItems($folder, [string]$field, [datetime]$from, [datetime]$until) {
    if ($until -le $from) { throw 'UntilDate must be later than FromDate (exclusive).' }
    # Jet date literals use the Outlook user's regional short date and time.
    $filter = "[$field] >= '" + $from.ToString('g') + "' AND [$field] < '" + $until.ToString('g') + "'"
    return ,$folder.Items.Restrict($filter)
}

function Read-OutlookMailStream($folder, $items, [string]$field, [datetime]$from, [datetime]$until) {
    $expected = $items.Count
    $ids = @{}
    for ($i=1; $i -le $expected; $i++) {
        $mail = $items.Item($i)
        if (!$mail -or !$mail.EntryID -or $ids.ContainsKey($mail.EntryID)) { throw 'Folder changed during export. Retry.' }
        $ids[$mail.EntryID] = $true
        if ($mail.Class -ne 43) { continue }
        $date = [datetime]$mail.$field
        if ($date -lt $from -or $date -ge $until -or $mail.Parent.EntryID -ne $folder.EntryID) { throw 'Folder changed during export. Retry.' }
        [pscustomobject]@{id=[string]$mail.EntryID; store=[string]$folder.StoreID; folder=[string]$folder.FolderPath;
            subject=[string]$mail.Subject; sender=[string]$mail.SenderName; at=$date; body=[string]$mail.Body}
    }
    if ($expected -ne $items.Count) { throw 'Folder changed during export. Retry.' }
}
