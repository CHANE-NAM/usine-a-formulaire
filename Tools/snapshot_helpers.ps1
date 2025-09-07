$repo = "G:\Mon Drive\APPLI TEST Personnalité Drive\Projet USINE à FORMULAIRE GoogleForm"
. "$repo\Tools\snapshot_helpers.ps1"

# Re-génère manifest/brief/diff pour le dernier snapshot :
$exp  = Join-Path $repo 'export-onglets-csv'
$last = Get-ChildItem $exp -Dir | ? Name -like 'SNAPSHOT_*' | Sort LastWriteTime -desc | Select -First 1

$man  = Write-Manifest -SnapshotDir $last.FullName -RepoRoot $repo
Write-BriefMd -SnapshotDir $last.FullName -Manifest $man

$prev = Get-ChildItem $exp -Dir |
        ? { $_.FullName -ne $last.FullName -and (Test-Path (Join-Path $_.FullName 'manifest.json')) } |
        Sort LastWriteTime -desc | Select -First 1
if ($prev) {
  Write-DiffMd -PrevManifestPath (Join-Path $prev.FullName 'manifest.json') `
               -CurrManifestPath (Join-Path $last.FullName 'manifest.json') `
               -OutPath          (Join-Path $last.FullName 'diff.md') | Out-Null
}

