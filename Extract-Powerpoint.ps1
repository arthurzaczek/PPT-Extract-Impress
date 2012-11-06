<# 
.SYNOPSIS 
    Extracts power point slides to impress js
.DESCRIPTION 
    Supports font-size, syntax highlighing, positioning and css classes for theming
.EXAMPLE
	.\Extract-Powerpoint.ps1 'MyPresentation.pptx'
.EXAMPLE
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

	[Parameter(HelpMessage="Extract each text frame without any bullet as source code")]
	[switch]$SourceCode,
	
	[Parameter(HelpMessage="Does not render the overview slide")]
	[switch]$NoOverview,
	
	[Parameter(HelpMessage="Column layout")]
	[switch]$LayoutColumns,
	[Parameter(HelpMessage="Horizontal layout (default)")]
	[switch]$LayoutHorizontal,
	[Parameter(HelpMessage="Vertical layout (default)")]
	[switch]$LayoutVertical,
	[Parameter(HelpMessage="Enables rotation during layout")]
	[switch]$LayoutRotation,
	
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

if($LayoutColumns) {
	$posMode = "columns"
} elseif($LayoutHorizontal) {
	$posMode = "horizontal"
} elseif($LayoutVertical) {
	$posMode = "vertical"
} else {
	$posMode = "horizontal"
}

$xPos = 0
$yPos = 0
$zPos = 0

$width = 1024
$height = 768
$gap = 100
$zGap = -50

$colCount = 5
$col = 0

$rot = 0
$rotDelta = 30

# ---------------- helper functions ------------------------------
function nameToClass($name) {
	if(!$name) { return "" }
	return ($name -replace "\d","").Trim().Replace(" ", "-")
}

function getStyleAttribute([string[]]$styles) {
	if(!$styles) { return '' }
	if($styles.length -eq 0) { return '' }
	
	return ' style="' + [string]::join(';', $styles) + '"'
}

function hasBullets($paragraphs) {
    foreach($p in $paragraphs) {
	   if($p.ParagraphFormat.Bullet.Visible) {
            return $true; 
       }
    }
    return $false;
}

function isSingleParagraph($paragraphs) {
    return $paragraphs.Count -le 1;
}

function out-result {
    $input | out-file $outFile -Append -Encoding "UTF8"
}

# ---------------- render functions ------------------------------
function renderHeader() {
	'<!doctype html>' | out-file $outFile -Encoding "UTF8"
	'<html>' | out-result
	'<link href="content/impress.css" rel="stylesheet" />' | out-result
	'<link href="content/ppt.css" rel="stylesheet" />' | out-result
    if($SourceCode) {
        '<link href="content/syntaxhighlighter/styles/shCore.css" rel="stylesheet" type="text/css" />' | out-result
        '<link href="content/syntaxhighlighter/styles/shThemeDefault.css" rel="stylesheet" type="text/css" />' | out-result
    }
	'<link href="content/custom.css" rel="stylesheet" />' | out-result
	'<body class="impress-not-supported">' | out-result
	'<div id="impress">' | out-result
}

function renderFooter() {
    if($SourceCode) {
        '<script src="content/syntaxhighlighter/scripts/shCore.js" type="text/javascript"></script>' | out-result
        '<script src="content/syntaxhighlighter/scripts/shAutoloader.js" type="text/javascript"></script>' | out-result
         
        '<script type="text/javascript">' | out-result
        'SyntaxHighlighter.autoloader(' | out-result
        "  'js jscript javascript  content/syntaxhighlighter/scripts/shBrushJScript.js'," | out-result
        "  'java                   content/syntaxhighlighter/scripts/shBrushJava.js'," | out-result
        "  'cpp                    content/syntaxhighlighter/scripts/shBrushCpp.js'," | out-result
        "  'csharp                 content/syntaxhighlighter/scripts/shBrushCSharp.js'" | out-result
        ');' | out-result
         
        'SyntaxHighlighter.all();' | out-result
        '</script>' | out-result
    }
	if(!$NoOverview) {
		'<div id="overview" class="step" data-x="3000" data-y="1500" data-scale="10"/>' | out-result
	}
	'</div>' | out-result
	'<script src="content/impress.js"></script>' | out-result
	'<script>impress().init();</script>' | out-result
	'</body>' | out-result
	'</html>' | out-result
}

function renderParagraphs($paragraphs) {
	foreach($p in $paragraphs) {
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
				('        <li' + (getStyleAttribute $styles) + '>' + $p.Text + '</li>') | out-result
			} else {
				('        <p' + (getStyleAttribute $styles) + '>' + $p.Text.TrimEnd() + '</p>') | out-result
			} 
		}
		else {
			'        <p' + (getStyleAttribute $styles) + '>&nbsp;</p>' | out-result
		}
	}
}

function renderSourceCode($shape) {
    '        <pre class="brush: csharp">' + $shape.TextFrame2.TextRange.Text.Replace("<", "&lt;")  + '</pre>' | out-result
}

function renderTextShape($shape) {
	$styles = @()
	if($Position) {
		$styles += ('position: absolute')
        $styles += ('top:' + $shape.Top + 'px')
        $styles += ('left:' + $shape.Left + 'px')
	}
	'    <div' + (getStyleAttribute $styles) + ' class="' + (nameToClass $shape.Name) + '">' | out-result
	$paragraphs = $shape.TextFrame2.TextRange.Paragraphs()
    if($SourceCode -and !(hasBullets $paragraphs) -and !(isSingleParagraph $paragraphs)) {
        renderSourceCode $shape
    } else {
        renderParagraphs $paragraphs
    }
	'    </div>' | out-result
}

function renderSlide($slide) {
    '<!-- ' + $slide.Name + ' -->' | out-result
    '<div class="step slide ' + (nameToClass $slide.CustomLayout.Name) + '" data-x="' + $xPos + '" data-y="' + $yPos + '" data-z="' + $zPos + '" data-rotate="' + $rot + '">' | out-result
    foreach($shape in $slide.Shapes) {
        if($shape.HasTextFrame) {
			renderTextShape $shape
        }
    }
    '</div>' | out-result
}

function updatePositions() {
	switch($posMode) {
		"columns" {
			$script:xPos += $width + $gap
			$script:col++
			
			if($col -ge $colCount) {
				$script:xPos = 0
				$script:col = 0
				$script:yPos += $height + $gap
			}
			$script:zPos += $zGap
		}
		"horizontal" {
			$script:xPos += $width + $gap
			$script:zPos += $zGap
		}
		"vertical" {
			$script:yPos += $height + $gap
			$script:zPos += $zGap
		}
		default { "unknown position mode " + $posMode | out-host }
	}
	
	if($LayoutRotation) {
		$script:rot += $rotDelta
	}
}

# ---------------- Main ------------------------------
'Extracting "' + $file + '"' | out-host
'to         "' + $outFile + '"' | out-host
renderHeader

# init powerpoint
Add-type -AssemblyName office
$app = New-Object -ComObject powerpoint.application
$app.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$presentation = $app.Presentations.open($file)

foreach($slide in $presentation.Slides) {
    ("-> " + $slide.Name) | out-host
	renderSlide $slide
	updatePositions
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
