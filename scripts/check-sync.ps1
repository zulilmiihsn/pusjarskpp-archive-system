$gsFiles = Get-ChildItem -Path "D:\Projects\Portal Arsip" -Filter "*.gs"
foreach ($f in $gsFiles) {
    $jsPath = $f.FullName -replace '\.gs$','.js'
    if (Test-Path $jsPath) {
        $r = Compare-Object (Get-Content $f.FullName) (Get-Content $jsPath)
        if ($r) {
            Write-Host "OUT OF SYNC: $($f.Name)"
        } else {
            Write-Host "OK: $($f.Name)"
        }
    } else {
        Write-Host "MISSING JS: $($f.Name)"
    }
}
