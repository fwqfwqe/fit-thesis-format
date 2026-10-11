# Windows + an already installed Microsoft Word desktop application.
# The default mode opens the source read-only, never calls Fields.Update, and
# never saves the source. Do not run this script on an untrusted macro document.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPdf,
    [switch]$UpdateFields,
    [switch]$AuthorizedCopy,
    [string]$SavedPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Get-AbsoluteOutputPath([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { throw 'An output path is empty.' }
    if (-not [System.IO.Path]::IsPathRooted($Value)) {
        $Value = Join-Path -Path (Get-Location).ProviderPath -ChildPath $Value
    }
    return [System.IO.Path]::GetFullPath($Value)
}

function Release-OwnedComObject($Value) {
    if ($null -ne $Value -and [System.Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

if ($env:OS -ne 'Windows_NT') { throw 'word_render.ps1 requires Windows and an installed Microsoft Word.' }
$taskInput = (Resolve-Path -LiteralPath $InputPath -ErrorAction Stop).ProviderPath
if (-not (Test-Path -LiteralPath $taskInput -PathType Leaf)) { throw 'InputPath must be an existing file.' }
if ([System.IO.Path]::GetExtension($taskInput) -ine '.docx') { throw 'Only .docx input is accepted.' }
$taskPdf = Get-AbsoluteOutputPath $OutputPdf
if ([System.IO.Path]::GetExtension($taskPdf) -ine '.pdf') { throw 'OutputPdf must end in .pdf.' }
if (Test-Path -LiteralPath $taskPdf) { throw "Refusing to overwrite: $taskPdf" }
$taskSaved = $null
if ($UpdateFields) {
    if (-not $AuthorizedCopy -or [string]::IsNullOrWhiteSpace($SavedPath)) {
        throw 'Updating fields requires -UpdateFields -AuthorizedCopy -SavedPath together. The caller must already have authorization for that copy.'
    }
    $taskSaved = Get-AbsoluteOutputPath $SavedPath
    if ([System.IO.Path]::GetExtension($taskSaved) -ine '.docx') { throw 'SavedPath must end in .docx.' }
    if ($taskSaved -ieq $taskInput) { throw 'SavedPath must differ from InputPath.' }
    if (Test-Path -LiteralPath $taskSaved) { throw "Refusing to overwrite: $taskSaved" }
} elseif ($AuthorizedCopy -or -not [string]::IsNullOrWhiteSpace($SavedPath)) {
    throw '-AuthorizedCopy and -SavedPath are only valid together with -UpdateFields.'
}
if ($taskInput -ieq $taskPdf -or ($null -ne $taskSaved -and $taskSaved -ieq $taskPdf)) {
    throw 'Input, PDF, and saved DOCX paths must be distinct.'
}

$taskBefore = (Get-FileHash -LiteralPath $taskInput -Algorithm SHA256).Hash
[void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($taskPdf))
if ($null -ne $taskSaved) {
    [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($taskSaved))
}
$taskApp = $null
$taskDocuments = $null
$taskDoc = $null
$taskOptions = $null
$taskOriginalOptions = @{}
$taskOriginalSecurity = $null
$taskOriginalAlerts = $null
$taskReport = $null
try {
    # Always create a dedicated instance; never attach to an active user session.
    $taskApp = New-Object -ComObject Word.Application
    $taskOriginalAlerts = $taskApp.DisplayAlerts
    $taskOriginalSecurity = $taskApp.AutomationSecurity
    $taskApp.Visible = $false
    $taskApp.DisplayAlerts = 0
    $taskApp.AutomationSecurity = 3
    $taskOptions = $taskApp.Options
    foreach ($taskName in @('UpdateFieldsAtPrint', 'UpdateLinksAtPrint', 'UpdateLinksAtOpen')) {
        $taskOriginalOptions[$taskName] = $taskOptions.$taskName
        $taskOptions.$taskName = $false
    }
    $taskDocuments = $taskApp.Documents
    # ConfirmConversions=false, ReadOnly=true, AddToRecentFiles=false.
    $taskDoc = $taskDocuments.Open($taskInput, $false, $true, $false)
    if (-not $taskDoc.ReadOnly) { throw 'Word did not open the source read-only.' }
    if ($UpdateFields) {
        # Two passes settle common TOC/SEQ/REF dependencies and page references.
        for ($taskPass = 0; $taskPass -lt 2; $taskPass++) {
            $taskStories = $null
            try {
                $taskStories = $taskDoc.StoryRanges
                foreach ($taskFirstStory in $taskStories) {
                    $taskStory = $taskFirstStory
                    while ($null -ne $taskStory) {
                        $taskNextStory = $null
                        $taskFields = $null
                        try {
                            $taskFields = $taskStory.Fields
                            [void]$taskFields.Update()
                            $taskNextStory = $taskStory.NextStoryRange
                        } finally {
                            Release-OwnedComObject $taskFields
                            Release-OwnedComObject $taskStory
                        }
                        $taskStory = $taskNextStory
                    }
                }
            } finally {
                Release-OwnedComObject $taskStories
            }
            $taskTocs = $null
            try {
                $taskTocs = $taskDoc.TablesOfContents
                for ($taskTocIndex = 1; $taskTocIndex -le $taskTocs.Count; $taskTocIndex++) {
                    $taskToc = $taskTocs.Item($taskTocIndex)
                    try { $taskToc.Update() } finally { Release-OwnedComObject $taskToc }
                }
            } finally {
                Release-OwnedComObject $taskTocs
            }
            $taskDoc.Repaginate()
        }
        # ReadOnly protects the input; SaveAs writes only the authorized new copy.
        if (Test-Path -LiteralPath $taskSaved) { throw "Refusing to overwrite: $taskSaved" }
        $taskDoc.SaveAs2($taskSaved, 16)
    }
    if (Test-Path -LiteralPath $taskPdf) { throw "Refusing to overwrite: $taskPdf" }
    $taskDoc.ExportAsFixedFormat($taskPdf, 17, $false)
    if (-not (Test-Path -LiteralPath $taskPdf -PathType Leaf)) { throw 'Word did not create the PDF.' }
    $taskReport = [ordered]@{
        schema_version = 2
        engine = 'Microsoft Word'
        engine_version = [string]$taskApp.Version
        input = $taskInput
        pdf = $taskPdf
        pdf_sha256 = (Get-FileHash -LiteralPath $taskPdf -Algorithm SHA256).Hash.ToLowerInvariant()
        saved_docx = $taskSaved
        saved_docx_sha256 = $null
        explicitly_updated_fields = [bool]$UpdateFields
        input_opened_read_only = $true
        pages = $taskDoc.ComputeStatistics(2)
        source_sha256_before = $taskBefore
    }
} finally {
    # Close only objects created above. No process-name termination is used.
    if ($null -ne $taskDoc) {
        try { $taskDoc.Close(0) } catch { Write-Warning "Could not close the owned document: $_" }
        Release-OwnedComObject $taskDoc
    }
    if ($null -ne $taskOptions) {
        foreach ($taskName in $taskOriginalOptions.Keys) {
            try { $taskOptions.$taskName = $taskOriginalOptions[$taskName] } catch { Write-Warning "Could not restore Word option $taskName" }
        }
        Release-OwnedComObject $taskOptions
    }
    Release-OwnedComObject $taskDocuments
    if ($null -ne $taskApp) {
        try {
            if ($null -ne $taskOriginalSecurity) { $taskApp.AutomationSecurity = $taskOriginalSecurity }
            if ($null -ne $taskOriginalAlerts) { $taskApp.DisplayAlerts = $taskOriginalAlerts }
        } catch { Write-Warning 'Could not restore an owned Word application setting.' }
        try { $taskApp.Quit(0) } catch { Write-Warning "Could not quit the owned Word instance: $_" }
        Release-OwnedComObject $taskApp
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $taskAfter = (Get-FileHash -LiteralPath $taskInput -Algorithm SHA256).Hash
    if ($taskBefore -cne $taskAfter) { throw 'Source SHA256 changed. Stop and investigate; do not deliver this run.' }
}
$taskReport['source_sha256_after'] = $taskAfter
$taskReport['source_unchanged'] = $true
if ($null -ne $taskSaved) {
    $taskReport['saved_docx_sha256'] = (Get-FileHash -LiteralPath $taskSaved -Algorithm SHA256).Hash.ToLowerInvariant()
}
$taskReport | ConvertTo-Json -Depth 4
