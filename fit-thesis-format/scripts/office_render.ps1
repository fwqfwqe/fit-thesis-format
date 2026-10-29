# Native, read-only DOCX renderer. Only explicitly selected layout fields may
# be refreshed, and only into a NEW DOCX. Never attaches to a user document.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidateSet('word','wps')][string]$Engine,
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutputPdf,
    [switch]$UpdateFields,
    [string]$SavedPath,
    [string]$StatePath
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$allowedFields=@('TOC','SEQ','REF','PAGEREF','STYLEREF','PAGE','NUMPAGES','SECTIONPAGES')
function Absolute-Output([string]$value,[string]$extension) {
    if([string]::IsNullOrWhiteSpace($value)){throw 'Empty output path.'}
    $path=[IO.Path]::GetFullPath($value)
    if([IO.Path]::GetExtension($path) -ine $extension){throw "Output must end in $extension"}
    if(Test-Path -LiteralPath $path){throw "Refusing overwrite: $path"}
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
    return $path
}
function Release-Com($obj){
    if($null -ne $obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)){
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj)
    }
}
function Sha([string]$path){return (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()}
function Set-RenderPhase([string]$phase){
    if(-not [string]::IsNullOrWhiteSpace($StatePath)){
        $taskState['phase']=$phase
        $taskState['updated_at_utc']=[DateTime]::UtcNow.ToString('o')
        [IO.File]::WriteAllText($StatePath,($taskState | ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
    }
}
function Page-Info($range){
    $logical=$null;$physical=$null;$style=$null
    try{$logical=[int]$range.Information(1)}catch{}
    try{$physical=[int]$range.Information(3)}catch{}
    try{$style=[int]$range.Sections.Item(1).Footers.Item(1).PageNumbers.NumberStyle}catch{}
    return @{actual_page=$logical;physical_page=$physical;number_style=$style}
}
function Field-Entries($doc,[bool]$refresh){
    $entries=[Collections.Generic.List[object]]::new()
    foreach($firstStory in $doc.StoryRanges){
        $story=$firstStory
        while($null -ne $story){
            $next=$null
            try{
                # Reverse iteration permits TOC expansion without invalidating later indices.
                for($index=$story.Fields.Count;$index -ge 1;$index--){
                    $field=$story.Fields.Item($index)
                    try{
                        $code=[string]$field.Code.Text
                        $kind=if($code -match '^\s*([A-Za-z]+)\b'){$Matches[1].ToUpperInvariant()}else{'UNKNOWN'}
                        $updated=$false
                        if($refresh){
                            if($kind -in $allowedFields){[void]$field.Update();$updated=$true}
                            # TOC.Update can replace its COM field/range object. A
                            # fresh pass below inspects the resulting collection.
                            $entries.Add([ordered]@{kind=$kind;code=$code;updated=$updated})
                            continue
                        }
                        $result=[string]$field.Result.Text
                        # WPS can repaginate the full document on Information().
                        # Valid field results need no per-field page lookup; only
                        # errors need a location. TOC targets are mapped below.
                        $physical=$null
                        if($result -match '(?i)Error!|错误[!！]|未找到引用源|未定义书签|Reference source not found|Bookmark not defined'){
                            $info=Page-Info $field.Result;$physical=$info.physical_page
                        }
                        $entries.Add([ordered]@{kind=$kind;code=$code;result=$result;updated=$updated;physical_page=$physical})
                    }finally{Release-Com $field}
                }
                $next=$story.NextStoryRange
            }finally{Release-Com $story}
            $story=$next
        }
    }
    return $entries.ToArray()
}
if($env:OS -ne 'Windows_NT'){throw 'Native Word/WPS automation requires Windows.'}
$taskInput=(Resolve-Path -LiteralPath $InputPath).ProviderPath
if([IO.Path]::GetExtension($taskInput) -ine '.docx'){throw 'Only DOCX is accepted.'}
$taskPdf=Absolute-Output $OutputPdf '.pdf'
$taskSaved=$null
if($UpdateFields){$taskSaved=Absolute-Output $SavedPath '.docx'}
elseif(-not [string]::IsNullOrWhiteSpace($SavedPath)){throw 'SavedPath requires UpdateFields.'}
if($taskSaved -ieq $taskInput){throw 'Input may not be overwritten.'}
$taskBefore=Sha $taskInput
$taskState=[ordered]@{schema_version=10;engine=$Engine;input=$taskInput;source_sha256=$taskBefore;phase='starting';closed=$false;owned_processes=@();preexisting_pids=@();com_hwnd=$null}
if(-not [string]::IsNullOrWhiteSpace($StatePath)){
    $StatePath=Absolute-Output $StatePath '.json'
    Set-RenderPhase 'starting'
}
$taskApp=$null;$taskDoc=$null;$taskDocuments=$null;$taskOptions=$null
$taskOptionsBefore=@{};$taskAppBefore=@{};$taskMayQuit=$false;$taskReport=$null
$taskFieldSnapshot=@();$taskFailure=$null;$taskWarnings=[Collections.Generic.List[string]]::new()
$taskProcessName=if($Engine -eq 'word'){'WINWORD'}else{'wps'}
$taskBeforeProcesses=@(Get-Process -Name $taskProcessName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$taskState['preexisting_pids']=$taskBeforeProcesses
try{
    $taskProgId=if($Engine -eq 'word'){'Word.Application'}else{'KWPS.Application'}
    $taskApp=New-Object -ComObject $taskProgId
    $taskDocuments=$taskApp.Documents
    # Some WPS releases return a shared server. Never use or quit one with user documents.
    if($taskDocuments.Count -ne 0){throw 'COM returned a nonempty application instance; refusing to touch user documents.'}
    $taskAppPath=[string]$taskApp.Path
    $taskExecutable=Join-Path $taskAppPath ($taskProcessName.ToLowerInvariant()+'.exe')
    if(-not(Test-Path -LiteralPath $taskExecutable -PathType Leaf)){throw 'Cannot verify the actual office executable.'}
    $taskProduct=[Diagnostics.FileVersionInfo]::GetVersionInfo($taskExecutable)
    if($Engine -eq 'wps' -and ($taskProduct.ProductName -notmatch 'WPS|Kingsoft|金山') -and $taskAppPath -notmatch 'Kingsoft|WPS'){throw 'Requested WPS but actual executable is not WPS.'}
    if($Engine -eq 'word' -and $taskProduct.ProductName -notmatch 'Microsoft'){throw 'Requested Word but actual executable is not Microsoft Word.'}
    $taskNewProcesses=@(Get-Process -Name $taskProcessName -ErrorAction SilentlyContinue | Where-Object {$_.Id -notin $taskBeforeProcesses})
    $taskMayQuit=$taskNewProcesses.Count -gt 0
    if(-not $taskMayQuit){throw 'COM did not create a dedicated new office process; close or isolate the existing automation instance before retrying.'}
    $taskState['owned_processes']=@($taskNewProcesses | ForEach-Object {@{id=$_.Id;start_utc=$_.StartTime.ToUniversalTime().ToString('o');start_filetime_utc=$_.StartTime.ToUniversalTime().ToFileTimeUtc();path=$_.Path}})
    $taskState['engine_executable']=$taskExecutable
    try{$taskState['com_hwnd']=$taskApp.Hwnd}catch{}
    Set-RenderPhase 'owned_instance_created'
    foreach($prop in @('Visible','DisplayAlerts','AutomationSecurity')){
        $taskAppBefore[$prop]=$taskApp.$prop
    }
    $taskApp.Visible=$false;$taskApp.DisplayAlerts=0;$taskApp.AutomationSecurity=3
    $taskOptions=$taskApp.Options
    foreach($prop in @('UpdateFieldsAtPrint','UpdateLinksAtPrint','UpdateLinksAtOpen')){
        try{$taskOptionsBefore[$prop]=$taskOptions.$prop;$taskOptions.$prop=$false}catch{
            throw "Required safe-open option unavailable: $prop"
        }
    }
    Set-RenderPhase 'opening_document'
    $taskDoc=$taskDocuments.Open($taskInput,$false,$true,$false)
    if(-not $taskDoc.ReadOnly){throw 'Office did not open input read-only.'}
    if($UpdateFields){
        Set-RenderPhase 'updating_layout_fields'
        for($pass=0;$pass -lt 2;$pass++){
            [void](Field-Entries $taskDoc $true)
            $taskDoc.Repaginate()
        }
        if(Test-Path -LiteralPath $taskSaved){throw 'Saved output appeared during operation.'}
        Set-RenderPhase 'saving_native_intermediate'
        try{$taskDoc.SaveAs2($taskSaved,16)}catch{
            if(Test-Path -LiteralPath $taskSaved){throw}
            $taskDoc.SaveAs($taskSaved,16)
        }
    }
    Set-RenderPhase 'repaginating'
    $taskDoc.Repaginate()
    Set-RenderPhase 'reading_field_results'
    # WPS NextStoryRange can cycle across identical linked section stories.
    # Read-only acceptance uses the actual PDF text plus native TOC targets;
    # avoid that unbounded COM walk on the full manuscript.
    $taskFieldSnapshot=if($Engine -eq 'wps' -and -not $UpdateFields){@()}else{@(Field-Entries $taskDoc $false)}
    $taskBookmarks=[ordered]@{}
    $taskShowHidden=$taskDoc.Bookmarks.ShowHidden
    Set-RenderPhase 'mapping_bookmark_pages'
    try{
        $taskDoc.Bookmarks.ShowHidden=$true
        foreach($bookmark in $taskDoc.Bookmarks){
            try{$taskBookmarks[[string]$bookmark.Name]=Page-Info $bookmark.Range}finally{Release-Com $bookmark}
        }
    }finally{$taskDoc.Bookmarks.ShowHidden=$taskShowHidden}
    Set-RenderPhase 'mapping_table_pages'
    $taskTableSpans=[Collections.Generic.List[object]]::new()
    for($tableIndex=1;$tableIndex -le $taskDoc.Tables.Count;$tableIndex++){
        $taskTable=$null;$taskTableRange=$null;$taskTableStart=$null;$taskTableEnd=$null
        try{
            $taskTable=$taskDoc.Tables.Item($tableIndex)
            $taskTableRange=$taskTable.Range
            $taskTableStart=$taskTableRange.Duplicate
            $taskTableEnd=$taskTableRange.Duplicate
            [void]$taskTableStart.SetRange($taskTableRange.Start,$taskTableRange.Start)
            $endPosition=[Math]::Max($taskTableRange.Start,$taskTableRange.End-1)
            [void]$taskTableEnd.SetRange($endPosition,$endPosition)
            $firstPhysical=[int]$taskTableStart.Information(3)
            $lastPhysical=[int]$taskTableEnd.Information(3)
            if($firstPhysical -lt 1 -or $lastPhysical -lt $firstPhysical){throw "Invalid physical page span for table $tableIndex"}
            $span=[ordered]@{table_index=$tableIndex;start_page=$firstPhysical;end_page=$lastPhysical;row_pages_available=$false;row_start_pages=@()}
            if($Engine -eq 'word' -and $lastPhysical -gt $firstPhysical){
                try{
                    $rows=[Collections.Generic.List[object]]::new()
                    for($ri=1;$ri -le $taskTable.Rows.Count;$ri++){
                        $row=$null;$range=$null
                        try{
                            $row=$taskTable.Rows.Item($ri);$range=$row.Range.Duplicate
                            [void]$range.SetRange($range.Start,$range.Start)
                            $rows.Add(@{row_index=$ri;physical_page=[int]$range.Information(3);heading_format=[int]$row.HeadingFormat})
                        }finally{Release-Com $range;Release-Com $row}
                    }
                    $span['row_start_pages']=$rows.ToArray();$span['row_pages_available']=$true
                }catch{$span['row_pages_error']=[string]$_}
            }
            $taskTableSpans.Add($span)
        }finally{
            Release-Com $taskTableStart;Release-Com $taskTableEnd;Release-Com $taskTableRange;Release-Com $taskTable
        }
    }
    $taskMathCount=$null
    try{$taskMathCount=[int]$taskDoc.OMaths.Count}catch{$taskWarnings.Add('Native equation collection unavailable; editability needs visual/application review.')}
    $taskFieldErrors=@($taskFieldSnapshot | Where-Object {$_.result -match '(?i)Error!|错误[!！]|未找到引用源|未定义书签|Reference source not found|Bookmark not defined'})
    if(Test-Path -LiteralPath $taskPdf){throw 'PDF output appeared during operation.'}
    Set-RenderPhase 'exporting_pdf'
    try{$taskDoc.ExportAsFixedFormat($taskPdf,17,$false)}catch{
        # WPS API versions without ExportAsFixedFormat expose the same fixed-format
        # export via SaveAs PDF. This changes only the owned in-memory document.
        if(Test-Path -LiteralPath $taskPdf){throw}
        if($Engine -ne 'wps'){throw}
        $taskDoc.SaveAs($taskPdf,17)
    }
    if(-not(Test-Path -LiteralPath $taskPdf -PathType Leaf)){throw 'Office created no PDF.'}
    $taskReport=[ordered]@{
        schema_version=10;engine_key=$Engine
        engine=if($Engine -eq 'word'){'Microsoft Word'}else{'WPS Writer'}
        engine_reported_name=[string]$taskApp.Name;engine_version=[string]$taskApp.Version
        engine_executable=$taskExecutable;engine_product=[string]$taskProduct.ProductName
        engine_file_version=[string]$taskProduct.FileVersion;owned_process_ids=@($taskNewProcesses | ForEach-Object {$_.Id})
        input=$taskInput;pdf=$taskPdf;pdf_sha256=(Sha $taskPdf)
        saved_docx=$taskSaved;saved_docx_sha256=$null
        explicitly_updated_fields=[bool]$UpdateFields;allowed_field_updates=$allowedFields
        input_opened_read_only=$true;pages=[int]$taskDoc.ComputeStatistics(2)
        source_sha256_before=$taskBefore;bookmark_pages=$taskBookmarks
        field_results=$taskFieldSnapshot;field_errors=$taskFieldErrors
        table_spans=$taskTableSpans.ToArray();table_spans_complete=$true
        native_math_count=$taskMathCount;native_math_collection_available=($null -ne $taskMathCount)
        warnings=$taskWarnings.ToArray()
    }
}catch{$taskFailure=$_}
finally{
    if($null -ne $taskDoc){try{$taskDoc.Close(0)}catch{$taskWarnings.Add('Owned document could not be closed.')};Release-Com $taskDoc}
    if($null -ne $taskOptions){
        foreach($prop in $taskOptionsBefore.Keys){try{$taskOptions.$prop=$taskOptionsBefore[$prop]}catch{$taskWarnings.Add("Could not restore option $prop")}}
        Release-Com $taskOptions
    }
    Release-Com $taskDocuments
    if($null -ne $taskApp){
        foreach($prop in $taskAppBefore.Keys){try{$taskApp.$prop=$taskAppBefore[$prop]}catch{$taskWarnings.Add("Could not restore setting $prop")}}
        if($taskMayQuit){try{$taskApp.Quit(0)}catch{$taskWarnings.Add('Owned application could not be closed.')}}
        Release-Com $taskApp
    }
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    $taskAfter=Sha $taskInput
    if($taskBefore -cne $taskAfter){throw 'Input SHA256 changed; this run is invalid.'}
    $taskState['closed']=($taskWarnings.Count -eq 0)
    $taskState['failed']=($null -ne $taskFailure)
    Set-RenderPhase 'finished'
}
if($null -ne $taskFailure){throw $taskFailure}
$taskReport.source_sha256_after=$taskAfter
$taskReport.source_unchanged=$true
$taskReport.warnings=$taskWarnings.ToArray()
$taskReport['state_file']=$StatePath
if($taskSaved){$taskReport.saved_docx_sha256=Sha $taskSaved}
$taskReport | ConvertTo-Json -Depth 12
