$file = "p:\impress\Software Engineering 1 - 01 - OOP Basics.pptx"
$outFile = ".\Software Engineering 1 - 01 - OOP Basics.html"

$xPos = 0
$yPos = 0
$width = 1000
$height = 800
$colCount = 5
$col = 0


'<!doctype html>' | out-file $outFile
'<html>' | out-file $outFile -Append
'<link href="content/impress.css" rel="stylesheet" />' | out-file $outFile -Append
'<body class="impress-not-supported">' | out-file $outFile -Append
'<div id="impress">' | out-file $outFile -Append

Add-type -AssemblyName office
$app = New-Object -ComObject powerpoint.application

$presentation = $app.Presentations.open($file)

foreach($slide in $presentation.Slides) {
    ("-> " + $slide.Name) | out-host
    '<div class="step slide" data-x="' + $xPos + '" data-y="' + $yPos + '">' | out-file $outFile -Append
    foreach($shape in $slide.Shapes) {
        ('<p>' + $shape.TextFrame.TextRange.Text + '</p>') | out-file $outFile -Append
    }
    '</div>' | out-file $outFile -Append
    $xPos += $width
    $col++
    
    if($col -ge $colCount) {
        $xPos = 0
        $col = 0
        $yPos += $height
    }
}

'<div id="overview" class="step" data-x="3000" data-y="1500" data-scale="10"/>' | out-file $outFile -Append
'</div>' | out-file $outFile -Append
'<script src="content/impress.js"></script>' | out-file $outFile -Append
'<script>impress().init();</script>' | out-file $outFile -Append
'</body>' | out-file $outFile -Append
'</html>' | out-file $outFile -Append

$app.quit()
$app = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()

"finished...." | out-host

& .\$outFile