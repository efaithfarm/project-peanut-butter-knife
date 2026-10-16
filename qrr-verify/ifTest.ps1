$hostFiles = @()

$output = Get-ChildItem -path C:\Windows\System32 -recurse
forEach ($item in $output) {
    if ($item.FullName -match "akhbahean;sifhjfg;aisnvihv") {
        $hostFiles += $item.Name
    }
}
if ($hostFiles.count -gt 0) {
    Write-Host "Host file found! See? $hostFiles"
} else {
    Write-Host "Sorry, no matches!"
}