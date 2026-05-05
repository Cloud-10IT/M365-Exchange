$publicScripts = Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Public') -Filter '*.ps1' -ErrorAction SilentlyContinue
$privateScripts = Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Private') -Filter '*.ps1' -ErrorAction SilentlyContinue

foreach ($script in @($privateScripts) + @($publicScripts)) {
    . $script.FullName
}

$functionsToExport = $publicScripts.BaseName
Export-ModuleMember -Function $functionsToExport