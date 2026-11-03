# samples/_raw/*.xlsx 를 실제 Excel 로 열어 전체 재계산 후 samples/ 에 저장한다.
# 사용: powershell -ExecutionPolicy Bypass -File scripts/excel_recalc.ps1
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$raw = Join-Path $root "samples\_raw"
$out = Join-Path $root "samples"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    Get-ChildItem $raw -Filter *.xlsx | ForEach-Object {
        $wb = $excel.Workbooks.Open($_.FullName)
        $excel.CalculateFull()
        $dest = Join-Path $out $_.Name
        if (Test-Path $dest) { Remove-Item $dest -Force }
        $wb.SaveAs($dest, 51)  # 51 = xlOpenXMLWorkbook (.xlsx)
        $wb.Close($false)
        Write-Output "recalculated: $dest"
    }
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
