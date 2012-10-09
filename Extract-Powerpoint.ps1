<# 
.SYNOPSIS 
    extract powerpoint
.DESCRIPTION 
    Extracts power point slides to impress js using powershell
.EXAMPLE
	.\Extract-Powerpoint.ps1 'MyPresentation.pptx'
	ls *.pptx | % { .\Extract-Powerpoint.ps1 $_ }
.NOTES 
    Author     : Arthur Zaczek, arthur@dasz.at
	License    : GNU General Public License (GPL)
.LINK 
    http://dasz.at
#> 

param(
	[Parameter(HelpMessage="Path to the Powerpoint to extract")]
	[string]$file,
	
	[Parameter(HelpMessage="Simple mode. Extract only shapes with no formating")]
	[switch]$Simple,

	[Parameter(HelpMessage="Extract position of shapes")]
	[switch]$Position,

	[Parameter(HelpMessage="Extract font style")]
	[switch]$FontStyle,
	
	[Parameter(HelpMessage="Open the result when finished")]
	[switch]$Open
)

if(!$file) {
	get-help .\Extract-Powerpoint.ps1 -Full
	exit 1
}

if(!$Simple) {
# full style
	$Position = $true
	$FontStyle = $true
}

# ---------------- Init variables ------------------------------
$file = resolve-path $file
$outFile = [System.IO.Path]::GetFileNameWithoutExtension($file) + '.html'

$xPos = 0
$yPos = 0
$zPos = 0

$width = 1024
$height = 768
$gap = 100
$zGap = -50

$colCount = 5
$col = 0

# ---------------- helper functions ------------------------------
function nameToClass($name) {
	return $name.Replace(" ", "-")
}

function getStyleAttribute([string[]]$styles) {
	if(!$styles) { return '' }
	if($styles.length -eq 0) { return '' }
	
	return ' style="' + [string]::join(';', $styles) + '"'
}

# ---------------- render functions ------------------------------
function renderHeader() {
	'<!doctype html>' | out-file $outFile
	'<html>' | out-file $outFile -Append
	'<link href="content/impress.css" rel="stylesheet" />' | out-file $outFile -Append
	'<link href="content/ppt.css" rel="stylesheet" />' | out-file $outFile -Append
	'<body class="impress-not-supported">' | out-file $outFile -Append
	'<div id="impress">' | out-file $outFile -Append
}

function renderFooter() {
	'<div id="overview" class="step" data-x="3000" data-y="1500" data-scale="10"/>' | out-file $outFile -Append
	'</div>' | out-file $outFile -Append
	'<script src="content/impress.js"></script>' | out-file $outFile -Append
	'<script>impress().init();</script>' | out-file $outFile -Append
	'</body>' | out-file $outFile -Append
	'</html>' | out-file $outFile -Append
}

function renderTextShape($shape) {
	$styles = @()
	if($Position) {
		$styles += ('position: absolute')
        $styles += ('top:' + $shape.Top + 'px')
        $styles += ('left:' + $shape.Left + 'px')
	}
	'    <div' + (getStyleAttribute $styles) + ' class="' + (nameToClass $shape.Name) + '">' | out-file $outFile -Append
	foreach($p in $shape.TextFrame2.TextRange.Paragraphs()) {
		$styles = @()
		if($FontStyle) {
			$styles += 'font-size: ' + $p.Font.Size + 'pt'
		}
		if($p.Text -and $p.Text.Trim()) {
			if($p.ParagraphFormat.Bullet.Visible) {
				$margin = ($p.ParagraphFormat.IndentLevel - 1) * 20
                if($margin -gt 0) {
                    $styles += ('margin-left:' + $margin + 'px')
                }
				('        <li' + (getStyleAttribute $styles) + '>' + $p.Text + '</li>') | out-file $outFile -Append
			} else {
				('        <p' + (getStyleAttribute $styles) + '>' + $p.Text.TrimEnd() + '</p>') | out-file $outFile -Append
			} 
		}
		else {
			'        <p' + (getStyleAttribute $styles) + '>&nbsp;</p>' | out-file $outFile -Append
		}
	}
	'    </div>' | out-file $outFile -Append
}

function renderSlide($slide) {
    '<!-- ' + $slide.Name + ' -->' | out-file $outFile -Append
    '<div class="step slide ' + (nameToClass $slide.CustomLayout.Name) + '" data-x="' + $xPos + '" data-y="' + $yPos + '" data-z="' + $zPos + '">' | out-file $outFile -Append
    foreach($shape in $slide.Shapes) {
        if($shape.HasTextFrame) {
			renderTextShape $shape
        }
    }
    '</div>' | out-file $outFile -Append
}

# ---------------- Main ------------------------------
'Extracting "' + $file + '"' | out-host
'to         "' + $outFile + '"' | out-host
renderHeader

# init powerpoint
Add-type -AssemblyName office
$app = New-Object -ComObject powerpoint.application
$presentation = $app.Presentations.open($file)

foreach($slide in $presentation.Slides) {
    ("-> " + $slide.Name) | out-host

	renderSlide $slide
	
    $xPos += $width + $gap
    $zPos += $zGap
    $col++
    
    if($col -ge $colCount) {
        $xPos = 0
        $col = 0
        $yPos += $height + $gap
    }
}

renderFooter

# Quit powerpoint
$app.quit()
$app = $null
[gc]::collect()
[gc]::WaitForPendingFinalizers()

"finished...." | out-host

if($Open) {
	& .\$outFile
}
