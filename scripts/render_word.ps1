param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference='Stop'
$source=(Resolve-Path -LiteralPath $InputPath).Path
$output=[System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $output -Force | Out-Null
$rendered=Join-Path $output ([System.IO.Path]::GetFileName($source))
if($rendered -eq $source){throw 'Use a separate output directory; reference originals are not modified.'}
Copy-Item -LiteralPath $source -Destination $rendered -Force
$zip=[System.IO.Compression.ZipFile]::OpenRead($rendered)
try {
    $reader=[System.IO.StreamReader]::new($zip.GetEntry('word/document.xml').Open())
    $documentXml=[xml]$reader.ReadToEnd()
    $reader.Dispose()
    $namespace=[System.Xml.XmlNamespaceManager]::new($documentXml.NameTable)
    $namespace.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $sectionFormats=@($documentXml.SelectNodes('//w:sectPr',$namespace) | ForEach-Object {
        $number=$_.SelectSingleNode('w:pgNumType',$namespace)
        if($number){$number.GetAttribute('fmt','http://schemas.openxmlformats.org/wordprocessingml/2006/main')}else{''}
    })
} finally {$zip.Dispose()}
$pdf=Join-Path $output ([System.IO.Path]::GetFileNameWithoutExtension($source)+'.pdf')
$wordApp=$null
$doc=$null
try {
    $wordApp=New-Object -ComObject Word.Application
    $wordApp.Visible=$false
    $wordApp.DisplayAlerts=0
    $doc=$wordApp.Documents.Open($rendered,$false,$false)
    $doc.Fields.Update() | Out-Null
    foreach($toc in $doc.TablesOfContents){$toc.Update()}
    $doc.Repaginate()
    # Odd-page section breaks may insert unnumbered blank backs.
    # Restart each Roman front section at the next content page number.
    $nextRoman=1
    $sectionIndex=0
    $layout=@()
    foreach($section in $doc.Sections){
        $startRange=$section.Range.Duplicate
        $startRange.Collapse(1)
        $endRange=$section.Range.Duplicate
        $endRange.End=$endRange.End-2
        $endRange.Collapse(0)
        $contentPages=$endRange.Information(3)-$startRange.Information(3)+1
        $role='body'
        if($sectionFormats[$sectionIndex] -eq 'upperRoman'){$role='front'}
        elseif($sectionIndex -eq 0 -and $sectionFormats.Count -gt 1 -and $sectionFormats[1] -eq 'upperRoman'){$role='cover'}
        $layout+=@{role=$role;pages=$contentPages}
        if($sectionFormats[$sectionIndex] -eq 'upperRoman'){
            foreach($footer in $section.Footers){
                $footer.PageNumbers.RestartNumberingAtSection=$true
                $footer.PageNumbers.StartingNumber=$nextRoman
            }
            $nextRoman+=$contentPages
        }
        $sectionIndex++
    }
    $doc.Repaginate()
    $doc.Fields.Update() | Out-Null
    foreach($toc in $doc.TablesOfContents){$toc.UpdatePageNumbers()}
    foreach($section in $doc.Sections){
        foreach($header in $section.Headers){$header.Range.Fields.Update() | Out-Null}
        foreach($footer in $section.Footers){$footer.Range.Fields.Update() | Out-Null}
    }
    $doc.Save()
    if($documentXml.SelectNodes('//w:sectPr/w:type[@w:val="oddPage"]',$namespace).Count -gt 0){
        $layoutPath=Join-Path $output 'section-layout.json'
        [System.IO.File]::WriteAllText($layoutPath,($layout | ConvertTo-Json -Depth 3),[System.Text.UTF8Encoding]::new($false))
        $doc.Close(0)
        $doc=$null
        $pythonPath=Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
        if(-not (Test-Path -LiteralPath $pythonPath)){throw 'Bundled Python path not found; provide an environment-specific renderer.'}
        & $pythonPath -X utf8 (Join-Path $PSScriptRoot 'prepare_duplex.py') $rendered $layoutPath
        if($LASTEXITCODE -ne 0){throw 'Duplex page preparation failed'}
        $doc=$wordApp.Documents.Open($rendered,$false,$false)
        $doc.Fields.Update() | Out-Null
        foreach($toc in $doc.TablesOfContents){$toc.Update()}
        $doc.Repaginate()
        foreach($toc in $doc.TablesOfContents){$toc.UpdatePageNumbers()}
        foreach($section in $doc.Sections){
            foreach($header in $section.Headers){$header.Range.Fields.Update() | Out-Null}
            foreach($footer in $section.Footers){$footer.Range.Fields.Update() | Out-Null}
        }
        $doc.Save()
    }
    $doc.ExportAsFixedFormat($pdf,17)
    Write-Output "Pages: $($doc.ComputeStatistics(2))"
    Write-Output $rendered
    Write-Output $pdf
} finally {
    if($doc){$doc.Close(0)}
    if($wordApp){$wordApp.Quit()}
}
