$ErrorActionPreference='Stop'
function Add-Text($slide,$text,$x,$y,$w,$h,$size,$color=0x3D2613,$center=$true){
 $s=$slide.Shapes.AddTextbox(1,$x*72,$y*72,$w*72,$h*72)
 $s.TextFrame.MarginLeft=0;$s.TextFrame.MarginRight=0;$s.TextFrame.MarginTop=0;$s.TextFrame.MarginBottom=0
 $s.TextFrame.VerticalAnchor=3
 $s.TextFrame.TextRange.Text=$text
 $s.TextFrame.TextRange.Font.Name='Yu Gothic';$s.TextFrame.TextRange.Font.NameFarEast='Yu Gothic'
 $s.TextFrame.TextRange.Font.Size=$size;$s.TextFrame.TextRange.Font.Bold=-1;$s.TextFrame.TextRange.Font.Color.RGB=$color
 if($center){$s.TextFrame.TextRange.ParagraphFormat.Alignment=2}
}
function Add-Box($slide,$x,$y,$w,$h,$color=0xF6F3ED){
 $s=$slide.Shapes.AddShape(1,$x*72,$y*72,$w*72,$h*72)
 $s.Fill.ForeColor.RGB=$color;$s.Line.Visible=0
}
function Add-Line($slide,$x1,$y1,$x2,$y2,$arrow=$false){
 $s=$slide.Shapes.AddLine($x1*72,$y1*72,$x2*72,$y2*72)
 $s.Line.ForeColor.RGB=0x746152;$s.Line.Weight=1.5
 if($arrow){$s.Line.EndArrowheadStyle=2}
}
$pptApp=New-Object -ComObject PowerPoint.Application
try{
 $deck=$pptApp.Presentations.Open((Join-Path $PSScriptRoot 'candidate_native.pptx'),$false,$false,$false)
 $slide=$deck.Slides.Item(3)
 for($i=$slide.Shapes.Count;$i -ge 1;$i--){$s=$slide.Shapes.Item($i);if($s.Top -ge 1.8*72 -and $s.Top -lt 3.8*72){$s.Delete()}}
 Add-Box $slide 0.7 2.1 3.2 1.4
 Add-Text $slide 'サイバー攻撃' 0.85 2.4 2.9 0.65 27
 Add-Box $slide 5.0 1.85 3.3 0.8
 Add-Text $slide '動画サービス停止' 5.1 1.93 3.1 0.64 23
 Add-Box $slide 5.0 2.95 3.3 0.8
 Add-Text $slide '出版事業に影響' 5.1 3.03 3.1 0.64 23
 Add-Box $slide 9.5 2.1 3.1 1.4 0xEEEcf8
 Add-Text $slide "営業利益への`r影響" 9.62 2.25 2.86 1.08 24 0x3729AE
 Add-Line $slide 3.9 2.8 4.4 2.8
 Add-Line $slide 4.4 2.25 4.4 3.35
 Add-Line $slide 4.4 2.25 4.95 2.25 $true
 Add-Line $slide 4.4 3.35 4.95 3.35 $true
 Add-Line $slide 8.3 2.25 8.85 2.25
 Add-Line $slide 8.3 3.35 8.85 3.35
 Add-Line $slide 8.85 2.25 8.85 3.35
 Add-Line $slide 8.85 2.8 9.45 2.8 $true
 foreach($s in $deck.Slides.Item(7).Shapes){
  if($s.HasTextFrame -eq -1 -and $s.TextFrame.HasText -eq -1){
   switch($s.TextFrame.TextRange.Text){
    '監視・権限分離' {$s.TextFrame.TextRange.Text="監視`r権限分離"}
    'バックアップ・復旧' {$s.TextFrame.TextRange.Text="バックアップ`r復旧演習"}
   }
  }
 }
 $slide=$deck.Slides.Item(8)
 for($i=$slide.Shapes.Count;$i -ge 1;$i--){$s=$slide.Shapes.Item($i);if($s.HasTextFrame -eq -1 -and $s.TextFrame.HasText -eq -1){if($s.TextFrame.TextRange.Text -like '8.2.35*8.5.12'){$s.Delete()}}}
 $versions=@('8.2.35','8.3.36','8.4.27','8.5.12')
 for($i=0;$i -lt 4;$i++){Add-Text $slide $versions[$i] (0.7+3.15*$i) 5.3 2.4 0.75 32 0x686B00 $false}
 $slide=$deck.Slides.Item(10)
 for($i=$slide.Shapes.Count;$i -ge 1;$i--){$s=$slide.Shapes.Item($i);if($s.Type -eq 9 -and $s.Top -gt 2*72 -and $s.Line.EndArrowheadStyle -ne 1){$s.Delete()}}
 Add-Line $slide 4.2 4.775 4.55 4.775
 Add-Line $slide 4.55 4.775 4.55 3.775
 Add-Line $slide 4.55 3.775 4.9 3.775 $true
 Add-Line $slide 8.4 3.775 8.75 3.775
 Add-Line $slide 8.75 3.775 8.75 2.775
 Add-Line $slide 8.75 2.775 9.1 2.775 $true
 # Suppress obsolete transition prompts carried over from the earlier slide order.
 foreach($page in $deck.Slides){foreach($s in $page.NotesPage.Shapes){if($s.HasTextFrame -eq -1 -and $s.TextFrame.HasText -eq -1){$text=$s.TextFrame.TextRange.Text;if($text.Contains('次への接続：')){$s.TextFrame.TextRange.Text=[regex]::Replace($text,'次への接続：[^\r\n]*(\r\n|\r|\n)?','')}}}}
 $deck.SaveAs((Join-Path $PSScriptRoot 'candidate_polished.pptx'),24)
 $deck.Close()
 Write-Output 'Revised causal branches, readable labels and version separation; retained all 18 slides.'
}finally{$pptApp.Quit();[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pptApp)}
