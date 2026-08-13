$Private = @(Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Private') -Filter '*.ps1' -ErrorAction SilentlyContinue)
$Public  = @(Get-ChildItem -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Public')  -Filter '*.ps1' -ErrorAction SilentlyContinue)

foreach ($File in @($Private + $Public)) {
    try {
        . $File.FullName
    }
    catch {
        throw "Failed to import function from '$($File.FullName)': $_"
    }
}

Export-ModuleMember -Function $Public.BaseName
