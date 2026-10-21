$ErrorActionPreference='Stop'
$pptApp=New-Object -ComObject PowerPoint.Application
try{
 $deck=$pptApp.Presentations.Open2007((Join-Path $PSScriptRoot 'candidate_normalized.pptx'),$false,$false,$false,$true)
 if($deck.Slides.Count -ne 18){throw 'Unexpected slide count'}
 $deck.SaveAs((Join-Path $PSScriptRoot 'candidate_native.pptx'),24)
 $deck.Close()
 Write-Output 'PowerPoint native normalization saved: 18 slides'
}finally{$pptApp.Quit();[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptApp)}
