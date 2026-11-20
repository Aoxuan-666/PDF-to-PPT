param(
    [Parameter(Mandatory = $true)]
    [string]$PptxPath,

    [Parameter(Mandatory = $true)]
    [string]$OutPdfPath
)

$ErrorActionPreference = 'Stop'
$PptxFull = [System.IO.Path]::GetFullPath($PptxPath)
$OutPdfFull = [System.IO.Path]::GetFullPath($OutPdfPath)

$outDir = [System.IO.Path]::GetDirectoryName($OutPdfFull)
if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}
if (Test-Path $OutPdfFull) {
    Remove-Item $OutPdfFull -Force
}

$app = $null
$pres = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    # ReadOnly = msoTrue (-1), Untitled = msoFalse (0), WithWindow = msoFalse (0)
    $pres = $app.Presentations.Open($PptxFull, -1, 0, 0)
    # ppSaveAsPDF = 32
    $pres.SaveAs($OutPdfFull, 32)
    Write-Output "Exported PDF: $OutPdfFull"
} finally {
    if ($pres -ne $null) {
        try { $pres.Close() } catch {}
    }
    if ($app -ne $null) {
        try { $app.Quit() } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
