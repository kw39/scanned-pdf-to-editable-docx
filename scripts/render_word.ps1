<# Read-only Word rendering through a separate empty COM instance. Windows + desktop Word required. #>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Docx,
    [Parameter(Mandatory=$true)][string]$Pdf,
    [string]$MapJson
)
$ErrorActionPreference = 'Stop'
$wordApp = $null
$taskDoc = $null
$activeApp = $null
$taskCopy = $null
$owned = $false
$activeInstanceDetected = $false
function Release-Com($obj) {
    if ($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
    }
}
function Com-Identity($obj) {
    $pointer = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($obj)
    try { return $pointer.ToInt64() } finally { [void][Runtime.InteropServices.Marshal]::Release($pointer) }
}
function Require-NewOutput([string]$path, [string]$extension) {
    $resolved = [IO.Path]::GetFullPath($path)
    if ([IO.Path]::GetExtension($resolved) -ine $extension) { throw "Output must end with $extension" }
    if (Test-Path -LiteralPath $resolved) { throw "Refusing to overwrite output: $resolved" }
    $parent = [IO.Path]::GetDirectoryName($resolved)
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) { throw "Create the output directory first: $parent" }
    return $resolved
}
try {
    $sourcePath = (Resolve-Path -LiteralPath $Docx).Path
    if ([IO.Path]::GetExtension($sourcePath) -ine '.docx') { throw 'Input must be a DOCX file' }
    $pdfPath = Require-NewOutput $Pdf '.pdf'
    $mapPath = if ($MapJson) { Require-NewOutput $MapJson '.json' } else { $null }
    if ($null -eq [Type]::GetTypeFromProgID('Word.Application')) { throw 'Desktop Microsoft Word is unavailable' }
    try { $activeApp = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') } catch { $activeApp = $null }
    $activeInstanceDetected = $null -ne $activeApp
    $wordApp = New-Object -ComObject Word.Application
    if ($null -ne $activeApp -and (Com-Identity $wordApp) -eq (Com-Identity $activeApp)) {
        throw 'COM activation returned the existing Word instance; refusing to control it'
    }
    if ($wordApp.Documents.Count -ne 0) { throw 'Word instance already contains documents; refusing to control it' }
    $owned = $true
    Release-Com $activeApp
    $activeApp = $null
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0
    $wordApp.AutomationSecurity = 3
    $taskCopy = Join-Path ([IO.Path]::GetTempPath()) ('scan-docx-qa-' + [Guid]::NewGuid().ToString('N') + '.docx')
    Copy-Item -LiteralPath $sourcePath -Destination $taskCopy
    # FileName, ConfirmConversions=false, ReadOnly=true, AddToRecentFiles=false.
    # OpenAndRepair is left at its normal false default.
    $taskDoc = $wordApp.Documents.Open($taskCopy, $false, $true, $false)
    $taskDoc.Repaginate()
    $taskDoc.ExportAsFixedFormat($pdfPath, 17)
    $paragraphs = @()
    if ($mapPath) {
        foreach ($paragraph in $taskDoc.Paragraphs) {
            $range = $paragraph.Range
            $startRange = $range.Duplicate
            $endRange = $range.Duplicate
            $startRange.Collapse(1)
            $endRange.Collapse(0)
            try {
                # Page numbers can differ when the scan's printed numbering restarts or skips.
                $paragraphs += [ordered]@{
                    text = $range.Text.TrimEnd([char]13, [char]7)
                    physical_page = $startRange.Information(3)
                    physical_end_page = $endRange.Information(3)
                    adjusted_page = $startRange.Information(1)
                    adjusted_end_page = $endRange.Information(1)
                    top_pt = $startRange.Information(6)
                    alignment = [int]$paragraph.Alignment
                    rendered_lines = $range.ComputeStatistics(1)
                    in_table = [bool]$range.Information(12)
                }
            } finally { Release-Com $startRange; Release-Com $endRange; Release-Com $range; Release-Com $paragraph }
        }
        $map = [ordered]@{
            source_docx = $sourcePath
            renderer = 'Microsoft Word'
            existing_instance_detected = $activeInstanceDetected
            separate_instance_guard = $true
            version = [string]$wordApp.Version
            build = [string]$wordApp.Build
            physical_page_count = $taskDoc.ComputeStatistics(2)
            scope = 'Main-story paragraphs. Headers, footers, drawing labels and visual geometry require rendered-page checks.'
            paragraphs = $paragraphs
        }
        [IO.File]::WriteAllText($mapPath, ($map | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($false)))
    }
    [ordered]@{renderer='Microsoft Word'; version=[string]$wordApp.Version; pdf=$pdfPath; map=$mapPath} | ConvertTo-Json -Compress
} finally {
    $cleanupErrors = @()
    if ($null -ne $taskDoc) {
        try {
            $closeSaveChanges = 0
            $taskDoc.Close([ref]$closeSaveChanges)
        } catch { $cleanupErrors += $_.Exception.Message }
        finally { Release-Com $taskDoc; $taskDoc = $null }
    }
    try {
        if ($owned -and $null -ne $wordApp -and $wordApp.Documents.Count -eq 0) {
            $quitSaveChanges = 0
            $wordApp.Quit([ref]$quitSaveChanges)
        }
    } catch { $cleanupErrors += $_.Exception.Message }
    Release-Com $activeApp
    Release-Com $wordApp
    try {
        if ($taskCopy -and (Test-Path -LiteralPath $taskCopy)) { Remove-Item -LiteralPath $taskCopy }
    } catch { $cleanupErrors += $_.Exception.Message }
    if ($cleanupErrors.Count) { throw ('Task Word cleanup failed: ' + ($cleanupErrors -join '; ')) }
}
