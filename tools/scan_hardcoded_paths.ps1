Param(
    [string]$Root = $PSScriptRoot + "\..",
    [string[]]$IncludeExtensions = @(".vbs", ".ps1", ".md", ".txt", ".csv"),
    [string]$OutFile = ""
)

$rootPath = (Resolve-Path $Root).Path
$pattern = '(?i)[A-Z]:\\[^\r\n"'']+'

$results = Get-ChildItem -Path $rootPath -Recurse -File |
    Where-Object { $IncludeExtensions -contains $_.Extension } |
    ForEach-Object {
        $file = $_.FullName
        $lineNumber = 0
        Get-Content -LiteralPath $file | ForEach-Object {
            $lineNumber++
            $line = $_
            $pathMatches = [regex]::Matches($line, $pattern)
            foreach ($m in $pathMatches) {
                [PSCustomObject]@{
                    File = $file
                    Line = $lineNumber
                    Path = $m.Value
                }
            }
        }
    }

if ($OutFile -ne "") {
    $results | Export-Csv -NoTypeInformation -Path $OutFile
} else {
    $results | Format-Table -AutoSize
}
