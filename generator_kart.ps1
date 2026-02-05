# ===========================
#  GENERATOR KART A7 NA A4
#  Tekst + Obrazki → PDF
#  Podgląd = tymczasowy HTML
# ===========================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

[System.Windows.Forms.Application]::EnableVisualStyles()


# --- Stałe układu strony / kart ---
$global:COLS = 4
$global:ROWS = 2
$global:SLOTS_PER_PAGE = $global:COLS * $global:ROWS

# --- Globalne dane obrazków ---
$global:ImagePaths = New-Object System.Collections.Generic.List[string]

$global:ProjectDirty = $false
$global:IsProgrammaticUpdate = $false

# --- Stałe do obliczania wymiarów pola tekstu na karcie ---
$global:FontName           = "Times New Roman"   # domyślna czcionka kart
$global:CardInnerWidthMm   = 64                  # przybliżona szerokość pola na wyraz (mm)
$global:CardTextHeightMm   = 44                  # przybliżona wysokość pola na wyraz (mm) w połowie karty

$global:PreviewBrowser = $null
$global:PreviewTimer   = $null

# --- Globalne ustawienia wyglądu (tekst + szata) ---

# Tekst
$global:TextColorHex      = "#000000"
$global:TextShadowEnabled = $false

# Tło karty
$global:CardBackgroundMode          = "none"     # none / color / gradient
$global:CardBackgroundColorHex      = "#FFFFFF"
$global:CardBackgroundGradColor1Hex = "#FFFFFF"
$global:CardBackgroundGradColor2Hex = "#E0ECFF"
$global:CardBackgroundGradientType  = "linear"   # linear / radial
$global:CardBackgroundGradientDirection = "diag" # vertical / horizontal / diag

# Obramowanie karty
$global:CardBorderEnabled  = $true
$global:CardBorderColorHex = "#3366CC"
$global:CardBorderWidthPx  = 2
$global:CardBorderStyle    = "solid"            # solid / dashed / dotted

# Środkowa linia
$global:CardLineEnabled  = $true
$global:CardLineColorHex = "#3366CC"
$global:CardLineWidthPx  = 2
$global:CardLineStyle    = "solid"             # solid / dashed / dotted

# Cień karty
$global:CardShadowEnabled = $false

# Linie cięcia
$global:ShowCutLines = $false

# Zaokrąglenie
$global:CardRoundedEnabled = $true

# --- Motywy jako zestawy ustawień ---

$global:CardStyles = @(

    # 1. KLASYCZNY
    [pscustomobject]@{
        Id   = 'classic'
        Name = 'Klasyczny'
        Settings = @{
            TextColorHex        = "#000000"
            TextShadowEnabled   = $false

            CardBackgroundMode          = "none"
            CardBackgroundColorHex      = "#FFFFFF"
            CardBackgroundGradColor1Hex = "#FFFFFF"
            CardBackgroundGradColor2Hex = "#E0ECFF"
            CardBackgroundGradientType  = "linear"
            CardBackgroundGradientDirection = "vertical"

            CardBorderEnabled   = $true
            CardBorderColorHex  = "#3366CC"
            CardBorderWidthPx   = 2
            CardBorderStyle     = "solid"

            CardLineEnabled     = $true
            CardLineColorHex    = "#3366CC"
            CardLineWidthPx     = 2
            CardLineStyle       = "solid"

            CardShadowEnabled   = $false
            CardRoundedEnabled  = $false
            ShowCutLines        = $false
        }
    },

    # 2. MIĘKKI GRADIENT
    [pscustomobject]@{
        Id   = 'soft'
        Name = 'Miękki'
        Settings = @{
            TextColorHex        = "#003366"
            TextShadowEnabled   = $false

            CardBackgroundMode          = "gradient"
            CardBackgroundColorHex      = "#FFFFFF"
            CardBackgroundGradColor1Hex = "#FFFFFF"
            CardBackgroundGradColor2Hex = "#CCE0FF"
            CardBackgroundGradientType  = "linear"
            CardBackgroundGradientDirection = "vertical"

            CardBorderEnabled   = $true
            CardBorderColorHex  = "#6699CC"
            CardBorderWidthPx   = 2
            CardBorderStyle     = "solid"

            CardLineEnabled     = $true
            CardLineColorHex    = "#6699CC"
            CardLineWidthPx     = 2
            CardLineStyle       = "solid"

            CardShadowEnabled   = $true
            CardRoundedEnabled  = $true
            ShowCutLines        = $false
        }
    },

    # 3. RAMKA / OBWÓDKA
    [pscustomobject]@{
        Id   = 'outline'
        Name = 'Obwódka'
        Settings = @{
            TextColorHex        = "#000000"
            TextShadowEnabled   = $false

            CardBackgroundMode          = "color"
            CardBackgroundColorHex      = "#FFFFFF"
            CardBackgroundGradColor1Hex = "#FFFFFF"
            CardBackgroundGradColor2Hex = "#FFFFFF"
            CardBackgroundGradientType  = "linear"
            CardBackgroundGradientDirection = "horizontal"

            CardBorderEnabled   = $true
            CardBorderColorHex  = "#000000"
            CardBorderWidthPx   = 2
            CardBorderStyle     = "dashed"

            CardLineEnabled     = $true
            CardLineColorHex    = "#000000"
            CardLineWidthPx     = 2
            CardLineStyle       = "dotted"

            CardShadowEnabled   = $false
            CardRoundedEnabled  = $false
            ShowCutLines        = $true
        }
    },

    # 4. PASTELOWY
    [pscustomobject]@{
        Id   = 'pastel'
        Name = 'Pastelowy'
        Settings = @{
            TextColorHex        = "#334455"
            TextShadowEnabled   = $false

            CardBackgroundMode          = "gradient"
            CardBackgroundColorHex      = "#FFFFFF"
            CardBackgroundGradColor1Hex = "#FFF5FF"
            CardBackgroundGradColor2Hex = "#E0F7FF"
            CardBackgroundGradientType  = "radial"
            CardBackgroundGradientDirection = "diag"

            CardBorderEnabled   = $true
            CardBorderColorHex  = "#FF9999"
            CardBorderWidthPx   = 2
            CardBorderStyle     = "solid"

            CardLineEnabled     = $true
            CardLineColorHex    = "#FF9999"
            CardLineWidthPx     = 2
            CardLineStyle       = "solid"

            CardShadowEnabled   = $true
            CardRoundedEnabled  = $true
            ShowCutLines        = $false
        }
    },

    # 5. RETRO 80s
    [pscustomobject]@{
        Id   = 'retro'
        Name = 'Retro 80s'
        Settings = @{
            TextColorHex        = "#FFFFFF"
            TextShadowEnabled   = $true

            CardBackgroundMode          = "gradient"
            CardBackgroundColorHex      = "#222233"
            CardBackgroundGradColor1Hex = "#222233"
            CardBackgroundGradColor2Hex = "#FF66CC"
            CardBackgroundGradientType  = "linear"
            CardBackgroundGradientDirection = "diag"

            CardBorderEnabled   = $true
            CardBorderColorHex  = "#00FFFF"
            CardBorderWidthPx   = 3
            CardBorderStyle     = "solid"

            CardLineEnabled     = $true
            CardLineColorHex    = "#FFFFFF"
            CardLineWidthPx     = 2
            CardLineStyle       = "dashed"

            CardShadowEnabled   = $true
            CardRoundedEnabled  = $true
            ShowCutLines        = $false
        }
    }
	
	    # 6. STYL UŻYTKOWNIKA (specjalny)
    [pscustomobject]@{
        Id   = 'user'
        Name = 'Użytkownika'
        Settings = @{}   # specjalnie puste – niczego nie narzuca
    }
)

# ================= POMOCNICZE FUNKCJE =================
function Set-UserStyleSelected {
    if ($global:UserStyleIndex -eq $null) { return }

    if ($cmbStyle.SelectedIndex -ne $global:UserStyleIndex) {
        $global:StyleChangingInternally = $true
        $cmbStyle.SelectedIndex = $global:UserStyleIndex
        $global:StyleChangingInternally = $false
    }
}

function Set-ProjectDirty {
    if (-not $global:IsProgrammaticUpdate) {
        $global:ProjectDirty = $true
    }
}

function Clear-ProjectDirty {
    $global:ProjectDirty = $false
}



function HtmlEncode([string]$s) {
    if ($null -eq $s) { return "" }
    $r = $s -replace '&','&amp;'
    $r = $r -replace '<','&lt;'
    $r = $r -replace '>','&gt;'
    $r = $r -replace '"','&quot;'
    $r = $r -replace "'","&#39;"
    return $r
}

function Get-ImageMimeType([string]$path) {
    $ext = [System.IO.Path]::GetExtension($path).ToLower()
    switch ($ext) {
        '.png'  { 'image/png'  }
        '.jpg'  { 'image/jpeg' }
        '.jpeg' { 'image/jpeg' }
        '.jfif' { 'image/jpeg' }
        '.bmp'  { 'image/bmp'  }
        '.gif'  { 'image/gif'  }
        '.tif'  { 'image/tiff' }
        '.tiff' { 'image/tiff' }
        '.ico'  { 'image/x-icon' }
        default { 'image/png'  }
    }
}

function Get-UniquePath([string]$path) {
    if (-not (Test-Path $path)) { return $path }
    $dir  = [System.IO.Path]::GetDirectoryName($path)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($path)
    $ext  = [System.IO.Path]::GetExtension($path)
    $i = 1
    while ($true) {
        $candidate = Join-Path $dir ("{0}_{1}{2}" -f $name, $i, $ext)
        if (-not (Test-Path $candidate)) { return $candidate }
        $i++
    }
}

function Compute-OptimalFontPt {
    param($Items)

    if (-not $Items -or $Items.Count -eq 0) { return 28 }

    $allLines = New-Object 'System.Collections.Generic.List[System.Collections.Generic.List[string]]'
    $maxLines = 1

    foreach ($it in $Items) {
        $linesHtml = Convert-RunsToLinesHtml $it.Runs

        $linesPlain = New-Object 'System.Collections.Generic.List[string]'
        foreach ($ln in $linesHtml) {
            $plain = ($ln -replace '<[^>]+>', '')
            if (-not [string]::IsNullOrWhiteSpace($plain)) {
                $linesPlain.Add($plain) | Out-Null
            }
        }

        if ($linesPlain.Count -gt 0) {
            $allLines.Add($linesPlain) | Out-Null
            if ($linesPlain.Count -gt $maxLines) { $maxLines = $linesPlain.Count }
        }
    }

    if ($allLines.Count -eq 0) { return 28 }

    $bmp = New-Object System.Drawing.Bitmap 2000,800
    $gfx = [System.Drawing.Graphics]::FromImage($bmp)
    $gfx.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias
    $fontName = $global:FontName
    $dpiX = $gfx.DpiX
    $dpiY = $gfx.DpiY

    $cardWidthPx  = [math]::Round($global:CardInnerWidthMm * $dpiX / 25.4)
    $cardHeightPx = [math]::Round($global:CardTextHeightMm * $dpiY / 25.4)

    $min  = 8
    $max  = 60
    $best = 14

    $widthSafety  = 0.95
    $heightSafety = 0.90

    while ($min -le $max) {
        $mid  = [int](($min + $max) / 2)
        $font = New-Object System.Drawing.Font($fontName, $mid)

        $fitsW = $true

        foreach ($group in $allLines) {
            for ($i = 0; $i -lt $group.Count; $i++) {
                $ln = $group[$i]
                if ([string]::IsNullOrEmpty($ln)) { continue }

                $testText = if ($i -eq $group.Count - 1) { $ln + " ?" } else { $ln }

                $sz = $gfx.MeasureString($testText, $font)
                if ($sz.Width -gt $cardWidthPx * $widthSafety) {
                    $fitsW = $false
                    break
                }
            }
            if (-not $fitsW) { break }
        }

        $szHeight = $gfx.MeasureString("Ag", $font)
        $lineH    = $szHeight.Height
        $totalH   = $lineH * $maxLines * 1.05
        $fitsH    = ($totalH -le $cardHeightPx * $heightSafety)

        $font.Dispose()

        if ($fitsW -and $fitsH) {
            $best = $mid
            $min  = $mid + 1
        } else {
            $max = $mid - 1
        }
    }

    $gfx.Dispose()
    $bmp.Dispose()

    return $best
}

# ================= CSS / HTML =================

function Save-Project {
    $project = [pscustomobject]@{
        Version = 1
        Mode    = if ($global:TabsControl.SelectedIndex -eq 0) { "text" } else { "image" }

        TextRtf   = if ($global:RichText) { $global:RichText.Rtf } else { $null }
        ImageList = if ($global:ImagePaths) { $global:ImagePaths } else { @() }

        Settings  = [pscustomobject]@{
            TextColorHex        = $global:TextColorHex
            TextShadowEnabled   = $global:TextShadowEnabled

            CardBackgroundMode          = $global:CardBackgroundMode
            CardBackgroundColorHex      = $global:CardBackgroundColorHex
            CardBackgroundGradColor1Hex = $global:CardBackgroundGradColor1Hex
            CardBackgroundGradColor2Hex = $global:CardBackgroundGradColor2Hex
            CardBackgroundGradientType  = $global:CardBackgroundGradientType
            CardBackgroundGradientDirection = $global:CardBackgroundGradientDirection

            CardBorderEnabled   = $global:CardBorderEnabled
            CardBorderColorHex  = $global:CardBorderColorHex
            CardBorderWidthPx   = $global:CardBorderWidthPx
            CardBorderStyle     = $global:CardBorderStyle

            CardLineEnabled     = $global:CardLineEnabled
            CardLineColorHex    = $global:CardLineColorHex
            CardLineWidthPx     = $global:CardLineWidthPx
            CardLineStyle       = $global:CardLineStyle

            CardShadowEnabled   = $global:CardShadowEnabled
            CardRoundedEnabled  = $global:CardRoundedEnabled
            ShowCutLines        = $global:ShowCutLines
            FontName            = $global:FontName
            FontSizePt          = $global:FontSizePt
        }
    }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "Projekt kart|*.jamam;*.json|Wszystkie pliki|*.*"
    $sfd.FileName = "projekt_kart.jamam"

    if ($sfd.ShowDialog() -eq 'OK') {
        $json = $project | ConvertTo-Json -Depth 10
        [IO.File]::WriteAllText($sfd.FileName, $json, [Text.Encoding]::UTF8)
		Clear-ProjectDirty
    }
	
}

function Load-Project {

    if ($global:ProjectDirty) {
        $res = [System.Windows.Forms.MessageBox]::Show(
            "Masz niezapisane zmiany. Kontynuacja spowoduje ich utratę.`n`nCzy chcesz kontynuować?",
            "Niezapisany projekt",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    }

    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Projekt kart|*.jamam;*.json|Wszystkie pliki|*.*"

    if ($ofd.ShowDialog() -eq 'OK') {

        $json = [IO.File]::ReadAllText($ofd.FileName, [Text.Encoding]::UTF8)
        $project = $json | ConvertFrom-Json

        # >> TU WYŁĄCZAMY „brudzenie” <<
        $global:IsProgrammaticUpdate = $true

        try {
            # Tryb
            if ($project.Mode -eq "text") {
                $global:TabsControl.SelectedIndex = 0
            } else {
                $global:TabsControl.SelectedIndex = 1
            }

            # Tekst
            if ($project.TextRtf -and $global:RichText) {
                $global:RichText.Rtf = $project.TextRtf
            }

            # Obrazki
            $global:ImagePaths.Clear()
            $listView.Items.Clear()
            $imgList.Images.Clear()

            foreach ($path in $project.ImageList) {
                if ([string]::IsNullOrWhiteSpace($path)) { continue }
                if (-not (Test-Path -Path $path)) { continue }

                try {
                    $img = [System.Drawing.Image]::FromFile($path)
                    $thumbBmp = New-Object System.Drawing.Bitmap($imgList.ImageSize.Width, $imgList.ImageSize.Height)
                    $g = [System.Drawing.Graphics]::FromImage($thumbBmp)
                    $g.Clear([System.Drawing.Color]::White)
                    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                    $scale = [Math]::Min($imgList.ImageSize.Width / $img.Width, $imgList.ImageSize.Height / $img.Height)
                    $dw = [int]([Math]::Round($img.Width * $scale))
                    $dh = [int]([Math]::Round($img.Height * $scale))
                    $dx = [int](($imgList.ImageSize.Width - $dw)/2)
                    $dy = [int](($imgList.ImageSize.Height - $dh)/2)
                    $g.DrawImage($img,$dx,$dy,$dw,$dh)
                    $g.Dispose()
                    $img.Dispose()

                    $key = [Guid]::NewGuid().ToString()
                    $imgList.Images.Add($key, $thumbBmp) | Out-Null
                    $item = New-Object System.Windows.Forms.ListViewItem([System.IO.Path]::GetFileName($path), $imgList.Images.IndexOfKey($key))
                    $listView.Items.Add($item) | Out-Null

                    $global:ImagePaths.Add($path) | Out-Null
                }
                catch {}
            }

            # Ustawienia
            $s = $project.Settings

            $global:TextColorHex        = $s.TextColorHex
            $global:TextShadowEnabled   = $s.TextShadowEnabled

            $global:CardBackgroundMode          = $s.CardBackgroundMode
            $global:CardBackgroundColorHex      = $s.CardBackgroundColorHex
            $global:CardBackgroundGradColor1Hex = $s.CardBackgroundGradColor1Hex
            $global:CardBackgroundGradColor2Hex = $s.CardBackgroundGradColor2Hex
            $global:CardBackgroundGradientType  = $s.CardBackgroundGradientType
            $global:CardBackgroundGradientDirection = $s.CardBackgroundGradientDirection

            $global:CardBorderEnabled   = $s.CardBorderEnabled
            $global:CardBorderColorHex  = $s.CardBorderColorHex
            $global:CardBorderWidthPx   = $s.CardBorderWidthPx
            $global:CardBorderStyle     = $s.CardBorderStyle

            $global:CardLineEnabled     = $s.CardLineEnabled
            $global:CardLineColorHex    = $s.CardLineColorHex
            $global:CardLineWidthPx     = $s.CardLineWidthPx
            $global:CardLineStyle       = $s.CardLineStyle

            $global:CardShadowEnabled   = $s.CardShadowEnabled
            $global:CardRoundedEnabled  = $s.CardRoundedEnabled
            $global:ShowCutLines        = $s.ShowCutLines
            $global:FontName            = $s.FontName
            $global:FontSizePt          = $s.FontSizePt
        }
        finally {
            $global:IsProgrammaticUpdate = $false
        }

        Clear-ProjectDirty
        Schedule-PreviewRefresh
    }
}

function Build-Css {
    param(
        [int]  $FontPt,
        [bool] $ShowCutLines
    )

    # --- Tekst ---
    $textHex  = $global:TextColorHex

    $textShadowCss =
        if ($global:TextShadowEnabled) {
            "text-shadow: 0 0.4mm 0.8mm rgba(0,0,0,0.35);"
        } else {
            "text-shadow: none;"
        }

    # --- Obramowanie karty ---
    $borderColor = $global:CardBorderColorHex
    $borderPx    = 0
    try { $borderPx = [int]$global:CardBorderWidthPx } catch { $borderPx = 0 }
    if ($borderPx -lt 0) { $borderPx = 0 }

    if (-not $global:CardBorderEnabled -or $borderPx -eq 0) {
        $borderCss = "border: none;"
    } else {
        $borderCss = "border: ${borderPx}px $($global:CardBorderStyle) $borderColor;"
    }

    # --- Linia środkowa ---
    $lineColor = $global:CardLineColorHex
    $linePx    = 0
    try { $linePx = [int]$global:CardLineWidthPx } catch { $linePx = 0 }
    if ($linePx -lt 0) { $linePx = 0 }

    if (-not $global:CardLineEnabled -or $linePx -eq 0) {
        $lineBaseCss = "border-top: none;"
    } else {
        $lineBaseCss = "border-top: ${linePx}px $($global:CardLineStyle) $lineColor;"
    }

    # --- Tło karty ---
    $bgCss = "background: #ffffff;"

    switch ($global:CardBackgroundMode) {
        'color' {
            $bgCss = "background: $($global:CardBackgroundColorHex);"
        }
		'gradient' {
			if ($global:CardBackgroundGradientType -eq 'radial') {
				$grad = "radial-gradient(circle at center, $($global:CardBackgroundGradColor1Hex), $($global:CardBackgroundGradColor2Hex))"
			} else {
				# kierunek dla gradientu liniowego
				$angle =
					switch ($global:CardBackgroundGradientDirection) {
						'vertical'   { '180deg' }  # góra-dół
						'horizontal' { '90deg'  }  # lewo-prawo
						default      { '135deg' }  # ukośny 45°
					}

				$grad = "linear-gradient($angle, $($global:CardBackgroundGradColor1Hex), $($global:CardBackgroundGradColor2Hex))"
			}
			$bgCss = "background: $grad;"
		}
    }

    # --- Cień karty ---
    $shadowCss =
        if ($global:CardShadowEnabled) {
            "box-shadow: 0 2mm 4mm rgba(0,0,0,0.18);"
        } else {
            "box-shadow: none;"
        }

    $radiusCss =
        if ($global:CardRoundedEnabled) {
            "border-radius: 4mm; overflow: hidden;"
        } else {
            "border-radius: 0; overflow: visible;"
        }

    # --- CSS właściwy ---
    $css = @"
@page {
    size: A4 landscape;
    margin: 0;
}

html, body {
    margin: 0;
    padding: 0;
    width: 100%;
    height: 100%;
}

body {
    font-family: '__FONTNAME__', serif;
    background: #ffffff;
}

/* Strona A4 */
.page {
    box-sizing: border-box;
    width: 297mm;
    height: 210mm;
    margin: 0 auto;
    background: white;
    display: block;
    position: relative;
    page-break-after: always;
}

/* Podgląd ekranowy */
@media screen {
    body {
        background: #d0d0d0;
    }

    .page {
		margin: 6mm auto 8mm auto;
        box-shadow: 0 0 4mm rgba(0,0,0,0.25);
    }
}

.page-table {
    border-collapse: collapse;
    width: 100%;
    height: 100%;
    table-layout: fixed;
}

.page-table tr {
    height: 50%;
}

.page-table td {
    width: 25%;
    padding: 0;
    margin: 0;
    vertical-align: top;
}

.card {
    box-sizing: border-box;
    padding: 4mm;
    margin: 1.5mm;
    position: relative;
    width: calc(100% - 3mm);
    height: calc(100% - 3mm);
}

.card-inner {
    position: relative;
    width: 100%;
    height: 100%;
    box-sizing: border-box;
    $borderCss
    $bgCss
    $shadowCss
    $radiusCss
}

/* środkowa linia */
.card-line {
    position: absolute;
    left: 10%;
    right: 10%;
    top: 50%;
    $lineBaseCss
    z-index: 10;
    pointer-events: none;
}

.half {
    position: absolute;
    left: 0;
    right: 0;
    box-sizing: border-box;
    padding: 3mm;
}

.half.top {
    top: 0;
    bottom: 50%;
}

.half.bottom {
    top: 50%;
    bottom: 0;
}

.half-content {
    display: table;
    width: 100%;
    height: 100%;
}

.half-content-inner {
    display: table-cell;
    vertical-align: middle;
    text-align: center;
}

/* KARTY TEKSTOWE */

.text-card .header,
.text-card .word {
    font-size: __FONTPT__pt;
    line-height: 1.1;
    word-wrap: break-word;
    color: $textHex;
    $textShadowCss
}

.text-card .header {
    margin-bottom: 2mm;
}

/* KARTY Z OBRAZAMI */

.image-card .header {
    font-size: 28 pt;
    margin-bottom: 2mm;
    color: $textHex;
    text-align: center;
    display: block;
    width: 100%;
    $textShadowCss
}


.image-card .half-content {
    width: 100%;
    height: 100%;
    display: block;
}

.image-card .half-content-inner {
    width: 100%;
    height: 100%;
    text-align: center;
    display: block;
}

.image-card .img-box {
    display: block;
    margin: 0 auto;
    width: 100%;
    height: calc(100% - 5mm);
    box-sizing: border-box;
    padding: 1mm;
    overflow: hidden;
    border-radius: 12px;
}

.image-card img {
    display: inline-block;
    border-radius: 12px;
    max-width: 100%;
    max-height: 100%;
    width: auto;
    height: auto;
}


.bold      { font-weight: bold; }
.italic    { font-style: italic; }
.underline { text-decoration: underline; }
"@

    if ($ShowCutLines) {
        $css += @"

.cutlines {
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    pointer-events: none;
    z-index: 50;
}

.cutline-h {
    position: absolute;
    top: 50%;
    left: 0;
    width: 100%;
    border-top: 1px dashed #999;
}

.cutline-v {
    position: absolute;
    top: 0;
    height: 100%;
    border-left: 1px dashed #999;
}

.cutline-v.x1 { left: 25%; }
cutline-v.x2 { left: 50%; }
.cutline-v.x3 { left: 75%; }
"@
    }

    $css = $css -replace '__FONTNAME__', $global:FontName
    $css = $css -replace '__FONTPT__',   $FontPt.ToString()

    return $css
}


function Build-FullHtml {
    param(
        [string]$Css,
        [string]$Body
    )

    return @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <title>Karty</title>
    <style>
$Css
    </style>
</head>
<body>
$Body
</body>
</html>
"@
}

# ================= KONWERSJA RUNÓW → HTML =================

function Convert-RunPartToSpanHtml {
    param(
        [string]$text,
        [System.Drawing.FontStyle]$style
    )

    if ([string]::IsNullOrEmpty($text)) { return "" }

    $enc = HtmlEncode $text

    $classes = @()
    if ($style -band [System.Drawing.FontStyle]::Bold)      { $classes += 'bold' }
    if ($style -band [System.Drawing.FontStyle]::Italic)    { $classes += 'italic' }
    if ($style -band [System.Drawing.FontStyle]::Underline) { $classes += 'underline' }

    if ($classes.Count -gt 0) {
        return ('<span class="{0}">{1}</span>' -f ([string]::Join(' ',$classes), $enc))
    } else {
        return $enc
    }
}

function Convert-RunsToLinesHtml {
    param($runs)

    $lines = New-Object System.Collections.Generic.List[string]
    if (-not $runs -or $runs.Count -eq 0) { return $lines }

    $currentSb = New-Object System.Text.StringBuilder
    $needNewLineNext = $false

    foreach ($r in $runs) {
        $text  = $r.Text
        $style = $r.Style
        if ([string]::IsNullOrEmpty($text)) { continue }

        $buf = New-Object System.Text.StringBuilder

        for ($i = 0; $i -lt $text.Length; $i++) {
            $ch = $text[$i]

            if ($ch -eq ' ') {
                if ($buf.Length -gt 0) {
                    $partHtml = Convert-RunPartToSpanHtml -text $buf.ToString() -style $style
                    [void]$currentSb.Append($partHtml)
                    $buf.Clear() | Out-Null
                }
                $needNewLineNext = $true
            }
            elseif ($ch -eq '-') {
                $buf.Append($ch) | Out-Null
                $partHtml = Convert-RunPartToSpanHtml -text $buf.ToString() -style $style
                [void]$currentSb.Append($partHtml)
                $buf.Clear() | Out-Null
                $needNewLineNext = $true
            }
            else {
                if ($needNewLineNext) {
                    if ($currentSb.Length -gt 0) {
                        $lines.Add($currentSb.ToString()) | Out-Null
                        $currentSb = New-Object System.Text.StringBuilder
                    }
                    $needNewLineNext = $false
                }
                $buf.Append($ch) | Out-Null
            }
        }

        if ($buf.Length -gt 0) {
            $partHtml = Convert-RunPartToSpanHtml -text $buf.ToString() -style $style
            [void]$currentSb.Append($partHtml)
            $buf.Clear() | Out-Null
        }
    }

    if ($currentSb.Length -gt 0) {
        $lines.Add($currentSb.ToString()) | Out-Null
    }

    return $lines
}

# ================= BUDOWANIE HTML STRON =================

function Build-TextPageBody {
    param(
        $Items,
        [int]$PageIndex
    )

    $array = @($Items)
    if ($null -eq $array) { return "" }
    $n = $array.Count
    if ($n -le 0) { return "" }

    $start = $PageIndex * $global:SLOTS_PER_PAGE
    if ($start -ge $n) { return "" }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<div class="page">')

	if ($global:ShowCutLines) {
		[void]$sb.AppendLine('  <div class="cutlines">')
		[void]$sb.AppendLine('    <div class="cutline-h"></div>')
		[void]$sb.AppendLine('    <div class="cutline-v x1"></div>')
		[void]$sb.AppendLine('    <div class="cutline-v x2"></div>')
		[void]$sb.AppendLine('    <div class="cutline-v x3"></div>')
		[void]$sb.AppendLine('  </div>')
	}

    [void]$sb.AppendLine('  <table class="page-table">')

    for ($row = 0; $row -lt $global:ROWS; $row++) {
        [void]$sb.AppendLine('<tr>')
        for ($col = 0; $col -lt $global:COLS; $col++) {
            $idx = $start + $row * $global:COLS + $col
            if ($idx -ge $n) {
                [void]$sb.AppendLine('<td></td>')
                continue
            }

            $i = $idx
            $j = ($idx + 1) % $n

            $topLines    = Convert-RunsToLinesHtml $array[$i].Runs
            $bottomLines = Convert-RunsToLinesHtml $array[$j].Runs

            $topArray    = @($topLines)
            $bottomArray = @($bottomLines)

            if ($bottomArray.Count -gt 0) {
                $last = $bottomArray.Count - 1
                $bottomArray[$last] = $bottomArray[$last] + "<span>&nbsp;?</span>"
            }

            $topHtml    = ($topArray    -join '<br/>')
            $bottomHtml = ($bottomArray -join '<br/>')

$card = @"
<td>
  <div class="card text-card">
    <div class="card-inner">
      <div class="card-line"></div>

      <div class="half top">
        <div class="half-content">
          <div class="half-content-inner">
            <div class="header">Ja mam</div>
            <div class="word">$topHtml</div>
          </div>
        </div>
      </div>

      <div class="half bottom">
        <div class="half-content">
          <div class="half-content-inner">
            <div class="header">Kto ma</div>
            <div class="word">$bottomHtml</div>
          </div>
        </div>
      </div>

    </div>
  </div>
</td>
"@
            [void]$sb.AppendLine($card)
        }
        [void]$sb.AppendLine('</tr>')
    }

    [void]$sb.AppendLine('</table>')
    [void]$sb.AppendLine('</div>')
    return $sb.ToString()
}

function Build-TextFullBody {
    param($Items)

    $array = @($Items)
    if ($null -eq $array -or $array.Count -eq 0) { return "" }
    $pages = [math]::Ceiling($array.Count / $global:SLOTS_PER_PAGE)
    $sb = New-Object System.Text.StringBuilder
    for ($p = 0; $p -lt $pages; $p++) {
        [void]$sb.AppendLine((Build-TextPageBody -Items $array -PageIndex $p))
    }
    return $sb.ToString()
}

function Build-ImagePageBody {
    param(
        [string[]]$ImagePaths,
        [int]     $PageIndex
    )

    if (-not $ImagePaths) { return "" }
    $n = $ImagePaths.Count
    if ($n -le 0) { return "" }

    $start = $PageIndex * $global:SLOTS_PER_PAGE
    if ($start -ge $n) { return "" }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<div class="page">')

	if ($global:ShowCutLines) {
		[void]$sb.AppendLine('  <div class="cutlines">')
		[void]$sb.AppendLine('    <div class="cutline-h"></div>')
		[void]$sb.AppendLine('    <div class="cutline-v x1"></div>')
		[void]$sb.AppendLine('    <div class="cutline-v x2"></div>')
		[void]$sb.AppendLine('    <div class="cutline-v x3"></div>')
		[void]$sb.AppendLine('  </div>')
	}

    [void]$sb.AppendLine('  <table class="page-table">')

    for ($row = 0; $row -lt $global:ROWS; $row++) {
        [void]$sb.AppendLine('<tr>')
        for ($col = 0; $col -lt $global:COLS; $col++) {
            $idx = $start + $row * $global:COLS + $col
            if ($idx -ge $n) {
                [void]$sb.AppendLine('<td></td>')
                continue
            }

            $i = $idx
            $j = ($idx + 1) % $n

            $pathTop    = $ImagePaths[$i]
            $pathBottom = $ImagePaths[$j]

            $mimeTop    = Get-ImageMimeType $pathTop
            $mimeBottom = Get-ImageMimeType $pathBottom

            $bytesTop    = [System.IO.File]::ReadAllBytes($pathTop)
            $bytesBottom = [System.IO.File]::ReadAllBytes($pathBottom)

            $b64Top    = [Convert]::ToBase64String($bytesTop)
            $b64Bottom = [Convert]::ToBase64String($bytesBottom)

            $srcTop    = "data:$mimeTop;base64,$b64Top"
            $srcBottom = "data:$mimeBottom;base64,$b64Bottom"

$card = @"
<td>
  <div class="card image-card">
    <div class="card-inner">
      <div class="card-line"></div>

      <div class="half top">
        <div class="half-content">
          <div class="half-content-inner">
            <div class="header">Ja mam</div>
            <div class="img-box">
              <img src="$srcTop" />
            </div>
          </div>
        </div>
      </div>

      <div class="half bottom">
        <div class="half-content">
          <div class="half-content-inner">
            <div class="header">Kto ma?</div>
            <div class="img-box">
              <img src="$srcBottom" />
            </div>
          </div>
        </div>
      </div>

    </div>
  </div>
</td>
"@
            [void]$sb.AppendLine($card)
        }
        [void]$sb.AppendLine('</tr>')
    }

    [void]$sb.AppendLine('</table>')
    [void]$sb.AppendLine('</div>')
    return $sb.ToString()
}

function Build-ImageFullBody {
    param([string[]]$ImagePaths)

    if (-not $ImagePaths -or $ImagePaths.Count -eq 0) { return "" }
    $pages = [math]::Ceiling($ImagePaths.Count / $global:SLOTS_PER_PAGE)
    $sb = New-Object System.Text.StringBuilder
    for ($p = 0; $p -lt $pages; $p++) {
        [void]$sb.AppendLine((Build-ImagePageBody -ImagePaths $ImagePaths -PageIndex $p))
    }
    return $sb.ToString()
}

# ================= EDGE / PDF =================

function Convert-HtmlToPdf {
    param(
        [string]$FullHtml,
        [string]$PdfPath
    )

    $htmlTemp = Join-Path $env:TEMP ("karty_{0}.html" -f (Get-Random))
    [System.IO.File]::WriteAllText($htmlTemp, $FullHtml, [System.Text.Encoding]::UTF8)

    $edgeCandidates = @(
        "$env:ProgramFiles (x86)\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
    )
    $edgePath = $edgeCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $edgePath) {
        [System.Windows.Forms.MessageBox]::Show(
            "Nie znaleziono Microsoft Edge (Chromium).",
            "Błąd"
        ) | Out-Null
        return
    }

    $uri = "file:///" + ($htmlTemp -replace '\\','/')

    $args = @(
        "--headless",
        "--disable-gpu",
        "--print-to-pdf=$PdfPath",
        $uri
    )

    $proc = Start-Process -FilePath $edgePath -ArgumentList $args -Wait -PassThru

    if ($proc.ExitCode -ne 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Edge zwrócił kod $($proc.ExitCode). PDF mógł nie zostać wygenerowany.",
            "Błąd generowania PDF"
        ) | Out-Null
    }
	$proc.Close()
	$proc.Dispose()

}

# ================= TEKST – RTF → RUNY =================

function Get-TextItems {
    param([System.Windows.Forms.RichTextBox]$rtb)

    $items = New-Object System.Collections.Generic.List[object]

    $origStart = $rtb.SelectionStart
    $origLen   = $rtb.SelectionLength

    for ($line = 0; $line -lt $rtb.Lines.Count; $line++) {
        $full = $rtb.Lines[$line]
        if ($null -eq $full) { continue }
        $trim = $full.Trim()
        if ([string]::IsNullOrWhiteSpace($trim)) { continue }

        $first = 0
        while ($first -lt $full.Length -and [char]::IsWhiteSpace($full[$first])) { $first++ }
        $last = $full.Length - 1
        while ($last -ge 0 -and [char]::IsWhiteSpace($full[$last])) { $last-- }
        if ($last -lt $first) { continue }

        $absStart = $rtb.GetFirstCharIndexFromLine($line) + $first
        $len      = $last - $first + 1

        $runs = New-Object System.Collections.Generic.List[object]

        if ($len -gt 0) {
            $rtb.Select($absStart, 1)
            $baseFont  = $rtb.SelectionFont
            if (-not $baseFont) { $baseFont = $rtb.Font }
            $curStyle = $baseFont.Style
            $sb = New-Object System.Text.StringBuilder

            for ($i=0; $i -lt $len; $i++) {
                $rtb.Select($absStart + $i, 1)
                $fi = $rtb.SelectionFont
                if (-not $fi) { $fi = $baseFont }
                $st = $fi.Style
                if ($i -eq 0 -or $st -eq $curStyle) {
                    [void]$sb.Append($full[$first + $i])
                } else {
                    $runs.Add([pscustomobject]@{
                        Text  = $sb.ToString()
                        Style = $curStyle
                    }) | Out-Null
                    $sb.Clear() | Out-Null
                    [void]$sb.Append($full[$first + $i])
                    $curStyle = $st
                }
            }
            if ($sb.Length -gt 0) {
                $runs.Add([pscustomobject]@{
                    Text  = $sb.ToString()
                    Style = $curStyle
                }) | Out-Null
            }
        }

        $items.Add([pscustomobject]@{
            Plain = $trim
            Runs  = $runs
        }) | Out-Null
    }

    $rtb.Select($origStart, $origLen)

    return $items
}

function Get-TextCountFast([System.Windows.Forms.RichTextBox]$rtb) {
    $count = 0
    foreach ($ln in $rtb.Lines) {
        if (-not [string]::IsNullOrWhiteSpace( ($ln -as [string]).Trim() )) {
            $count++
        }
    }
    return $count
}

# ================= PODGLĄD =================

function Update-PreviewFromState {
    if (-not $global:PreviewBrowser) { return }

    $emptyHtml = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<style>
body {
    font-family: Segoe UI, sans-serif;
    background: #f5f5f5;
    color: #777;
    text-align: center;
    padding-top: 40px;
}
</style>
</head>
<body>
Brak danych do podglądu.
</body>
</html>
"@

    $fontSize = 28
    $body     = ""
    $showCut  = $global:ShowCutLines

    if ($global:TabsControl.SelectedIndex -eq 0) {
        $items = Get-TextItems -rtb $global:RichText
        if ($items.Count -lt 2) {
            $global:PreviewBrowser.DocumentText = $emptyHtml
            return
        }

        $fontSize = Compute-OptimalFontPt $items
        $body     = Build-TextFullBody -Items $items
    }
    else {
        if ($global:ImagePaths.Count -lt 2) {
            $global:PreviewBrowser.DocumentText = $emptyHtml
            return
        }

        $fontSize = 40
        $body     = Build-ImageFullBody -ImagePaths $global:ImagePaths.ToArray()
    }

    $css  = Build-Css -FontPt $fontSize -ShowCutLines:$showCut
    $html = Build-FullHtml -Css $css -Body $body
    $global:PreviewBrowser.DocumentText = $html
}

if (-not $global:PreviewTimer) {
    $global:PreviewTimer = New-Object System.Windows.Forms.Timer
    $global:PreviewTimer.Interval = 400
    $global:PreviewTimer.Add_Tick({
        $global:PreviewTimer.Stop()
        Update-PreviewFromState
    })
}

function Schedule-PreviewRefresh {
    if (-not $global:PreviewBrowser) { return }
    $global:PreviewTimer.Stop()
    $global:PreviewTimer.Start()
}

# ================= UI =================

# Paleta kolorów
$colorBgForm      = [System.Drawing.Color]::FromArgb(245,247,250)
$colorPanel       = [System.Drawing.Color]::FromArgb(252,252,254)
$colorAccent      = [System.Drawing.Color]::FromArgb(52,120,212)
$colorAccentDark  = [System.Drawing.Color]::FromArgb(36,94,176)
$colorTextMain    = [System.Drawing.Color]::FromArgb(30,30,35)
$colorTextMuted   = [System.Drawing.Color]::FromArgb(110,115,125)

function Set-FlatButtonStyle {
    param(
        [System.Windows.Forms.Button]$btn,
        [System.Drawing.Color]$bg,
        [System.Drawing.Color]$fg
    )

    $btn.FlatStyle = 'Flat'
    $btn.FlatAppearance.BorderSize = 0
    $btn.FlatAppearance.MouseOverBackColor = $colorAccentDark
    $btn.FlatAppearance.MouseDownBackColor = $colorAccentDark

    $btn.BackColor = $bg
    $btn.ForeColor = $fg
    $btn.Margin    = '4,4,0,4'
    $btn.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Karty Ja mam - kto ma? w formacie A7"
$form.StartPosition = "CenterScreen"


$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

# Ustawienia okna jako 80% ekranu
$winW = [int]($screen.Width  * 0.80)
$winH = [int]($screen.Height * 0.80)

$form.Size = New-Object System.Drawing.Size($winW, $winH)

# sensowne minimum: 60% ekranu
$form.MinimumSize = New-Object System.Drawing.Size(
    [int]($screen.Width  * 0.60),
    [int]($screen.Height * 0.60)
)

$form.BackColor = $colorBgForm
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Padding = '8,8,8,8'

# --- Górny pasek ---
$topBar = New-Object System.Windows.Forms.TableLayoutPanel
$topBar.ColumnCount = 2
$topBar.RowCount    = 1
$topBar.Dock        = 'Top'
$topBar.Height      = 50
$topBar.Margin      = '0,0,0,0'
$topBar.Padding     = '8,8,8,4'
$topBar.BackColor   = $colorPanel
$topBar.CellBorderStyle = 'None'
$topBar.ColumnStyles.Add(
    (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,70))
) | Out-Null
$topBar.ColumnStyles.Add(
    (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,30))
) | Out-Null

# lewa część: opcje globalne
$panelOptions = New-Object System.Windows.Forms.FlowLayoutPanel
$panelOptions.Dock = 'Fill'
$panelOptions.FlowDirection = 'LeftToRight'
$panelOptions.WrapContents  = $true
$panelOptions.Margin  = '0,0,0,0'
$panelOptions.Padding = '0,0,0,0'
$panelOptions.BackColor = [System.Drawing.Color]::Transparent
$topBar.Controls.Add($panelOptions,0,0) | Out-Null

# linie cięcia
$chkCutlines = New-Object System.Windows.Forms.CheckBox
$chkCutlines.Text = "Linie cięcia"
$chkCutlines.AutoSize = $true
$chkCutlines.Margin   = '4,9,4,4'
$chkCutlines.ForeColor = $colorTextMuted
$chkCutlines.Checked = $global:ShowCutLines
$panelOptions.Controls.Add($chkCutlines) | Out-Null

$chkCutlines.Add_CheckedChanged({
    $global:ShowCutLines = $chkCutlines.Checked
    Set-ProjectDirty
	Schedule-PreviewRefresh
})


# wybór motywu
$lblStyleTop = New-Object System.Windows.Forms.Label
$lblStyleTop.Text = "Motyw:"
$lblStyleTop.AutoSize = $true
$lblStyleTop.Margin   = '20,9,4,4'
$lblStyleTop.ForeColor = $colorTextMuted
$panelOptions.Controls.Add($lblStyleTop) | Out-Null

$cmbStyle = New-Object System.Windows.Forms.ComboBox
$cmbStyle.DropDownStyle = 'DropDownList'
$cmbStyle.Width        = 140
$cmbStyle.DropDownWidth = 220
$cmbStyle.Margin  = '0,6,0,4'
$panelOptions.Controls.Add($cmbStyle) | Out-Null

$global:UserStyleIndex = $null
$global:StyleChangingInternally = $false

foreach ($style in $global:CardStyles) {
    [void]$cmbStyle.Items.Add($style.Name)
}
if ($cmbStyle.Items.Count -gt 0) {
    $cmbStyle.SelectedIndex = 0
}

# znajdź indeks stylu "user"
for ($i = 0; $i -lt $global:CardStyles.Count; $i++) {
    if ($global:CardStyles[$i].Id -eq 'user') {
        $global:UserStyleIndex = $i
        break
    }
}

# prawa część topBar – przyciski
$panelButtons = New-Object System.Windows.Forms.FlowLayoutPanel
$panelButtons.Dock = 'Fill'
$panelButtons.FlowDirection = 'RightToLeft'
$panelButtons.WrapContents  = $false
$panelButtons.Margin  = '0,0,0,0'
$panelButtons.Padding = '0,0,0,0'
$panelButtons.BackColor = [System.Drawing.Color]::Transparent
$topBar.Controls.Add($panelButtons,1,0) | Out-Null

$btnSavePdf = New-Object System.Windows.Forms.Button
$btnSavePdf.Text   = "Eksportuj PDF"
$btnSavePdf.Width  = 110
$btnSavePdf.Height = 32
$btnSavePdf.Enabled = $false
Set-FlatButtonStyle $btnSavePdf $colorAccent ([System.Drawing.Color]::White)
$panelButtons.Controls.Add($btnSavePdf) | Out-Null

# --- ZAPISZ PROJEKT ---
$btnSaveProject = New-Object System.Windows.Forms.Button
$btnSaveProject.Text = "Zapisz projekt"
$btnSaveProject.Width = 110
$btnSaveProject.Height = 32
Set-FlatButtonStyle $btnSaveProject $colorAccent ([System.Drawing.Color]::White)
$panelButtons.Controls.Add($btnSaveProject)

# --- OTWÓRZ PROJEKT ---
$btnLoadProject = New-Object System.Windows.Forms.Button
$btnLoadProject.Text = "Otwórz projekt"
$btnLoadProject.Width = 110
$btnLoadProject.Height = 32
Set-FlatButtonStyle $btnSaveProject $colorAccent ([System.Drawing.Color]::White)
$panelButtons.Controls.Add($btnLoadProject)


# --- Panel główny: lewa (zakładki), prawa (podgląd + ustawienia) ---
$panelMain = New-Object System.Windows.Forms.Panel
$panelMain.Dock    = 'Fill'
$panelMain.Margin  = '0,4,0,0'
$panelMain.Padding = '0,0,0,0'
$panelMain.BackColor = $colorBgForm

$form.Controls.Add($panelMain) | Out-Null
$form.Controls.Add($topBar)    | Out-Null

$splitMain = New-Object System.Windows.Forms.SplitContainer
$splitMain.Dock         = 'Fill'
$splitMain.Orientation  = 'Vertical'
$splitMain.BorderStyle  = 'None'
$splitMain.Panel1MinSize = 320
$splitMain.Panel2Collapsed = $false
$panelMain.Controls.Add($splitMain) | Out-Null

# --- Zakładki po lewej ---
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock   = 'Fill'
$tabs.Margin = '0,0,0,0'
$tabs.Padding = New-Object System.Drawing.Point(16, 4)
$tabs.Font   = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$tabs.Appearance = 'Normal'
$tabs.SizeMode   = 'Fixed'
$tabs.ItemSize   = New-Object System.Drawing.Size(120, 26)
$splitMain.Panel1.Controls.Add($tabs) | Out-Null

$global:TabsControl = $tabs

# --- Prawa strona: podgląd + panel ustawień ---
# --- Prawa strona: podgląd + panel ustawień (skalowane splitterem) ---
$splitRight = New-Object System.Windows.Forms.SplitContainer
$splitRight.Dock        = 'Fill'
$splitRight.Orientation = 'Horizontal'  # góra/dół
$splitRight.BorderStyle = 'None'
$splitRight.Panel1MinSize = 100
$splitRight.Panel2MinSize = 150
$splitRight.SplitterWidth = 6
$splitRight.IsSplitterFixed = $false
$splitRight.BackColor   = $colorBgForm

$splitMain.Panel2.Controls.Add($splitRight) | Out-Null


# ========= GÓRA: PODGLĄD =========
$previewBrowser = New-Object System.Windows.Forms.WebBrowser
$previewBrowser.Dock = 'Fill'
$previewBrowser.ScriptErrorsSuppressed = $true
$previewBrowser.AllowWebBrowserDrop   = $false
$previewBrowser.IsWebBrowserContextMenuEnabled = $false
$previewBrowser.WebBrowserShortcutsEnabled    = $false
$splitRight.Panel1.Controls.Add($previewBrowser) | Out-Null
$global:PreviewBrowser = $previewBrowser

# ================= PANEL USTAWIEŃ (EDYTOR WYGLĄDU) =================

$settingsPanel = New-Object System.Windows.Forms.Panel
$settingsPanel.Dock = 'Fill'
$settingsPanel.BackColor = [System.Drawing.Color]::FromArgb(246,248,252)
$settingsPanel.AutoScroll = $true
$settingsPanel.AutoScrollMinSize = New-Object System.Drawing.Size(0, 300)

# wewnętrzny layout: nagłówek + tabcontrol
$settingsRoot = New-Object System.Windows.Forms.TableLayoutPanel
$settingsRoot.Dock = 'Fill'
$settingsRoot.ColumnCount = 1
$settingsRoot.RowCount    = 2
$settingsRoot.RowStyles.Add(
    (New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::AutoSize))
) | Out-Null
$settingsRoot.RowStyles.Add(
    (New-Object System.Windows.Forms.RowStyle(
        [System.Windows.Forms.SizeType]::Percent,100))
) | Out-Null
$settingsRoot.Padding = '4,4,4,4'
$settingsPanel.Controls.Add($settingsRoot)

# podpis panelu
$lblAppearanceTitle = New-Object System.Windows.Forms.Label
$lblAppearanceTitle.Text = "Edytor wyglądu"
$lblAppearanceTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$lblAppearanceTitle.AutoSize = $true
$lblAppearanceTitle.Margin   = '2,0,2,4'
$settingsRoot.Controls.Add($lblAppearanceTitle,0,0) | Out-Null

# główny TabControl: TEKST / SZATA
$tabsAppearance = New-Object System.Windows.Forms.TabControl
$tabsAppearance.Dock   = 'Fill'
$tabsAppearance.Padding = New-Object System.Drawing.Point(12,4)
$tabsAppearance.ItemSize = New-Object System.Drawing.Size(90,24)
$tabsAppearance.SizeMode = 'Fixed'
$settingsRoot.Controls.Add($tabsAppearance,0,1) | Out-Null

# ------------------ ZAKŁADKA: TEKST ------------------

$tabTextSettings = New-Object System.Windows.Forms.TabPage
$tabTextSettings.Text     = "Tekst"
$tabTextSettings.BackColor = [System.Drawing.Color]::FromArgb(245,245,249)
$tabTextSettings.Padding  = '6,6,6,6'
$tabsAppearance.TabPages.Add($tabTextSettings) | Out-Null

$tlpTextSettings = New-Object System.Windows.Forms.TableLayoutPanel
$tlpTextSettings.Dock = 'Top'
$tlpTextSettings.AutoSize = $true
$tlpTextSettings.AutoSizeMode = 'GrowAndShrink'
$tlpTextSettings.ColumnCount = 1
$tlpTextSettings.RowCount    = 6
$tlpTextSettings.Padding = '2,2,2,2'
$tlpTextSettings.BackColor = [System.Drawing.Color]::FromArgb(245,245,249)
$tabTextSettings.Controls.Add($tlpTextSettings)

# nagłówek "Tekst"
$lblTextSettings = New-Object System.Windows.Forms.Label
$lblTextSettings.Text = "Tekst na kartach"
$lblTextSettings.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblTextSettings.AutoSize = $true
$lblTextSettings.Margin = '0,0,0,6'
$tlpTextSettings.Controls.Add($lblTextSettings,0,0) | Out-Null

# kolor tekstu
$btnTextColor = New-Object System.Windows.Forms.Button
$btnTextColor.Text = "Kolor tekstu"
$btnTextColor.Width = 140
$btnTextColor.Height = 26
$btnTextColor.Margin = '0,0,0,6'
$btnTextColor.BackColor = [System.Drawing.ColorTranslator]::FromHtml($global:TextColorHex)
$tlpTextSettings.Controls.Add($btnTextColor,0,1) | Out-Null

$btnTextColor.Add_Click({
    $cd = New-Object System.Windows.Forms.ColorDialog
    $cd.Color = [System.Drawing.ColorTranslator]::FromHtml($global:TextColorHex)
    if ($cd.ShowDialog() -eq 'OK') {
        $global:TextColorHex = ("#{0:X2}{1:X2}{2:X2}" -f $cd.Color.R,$cd.Color.G,$cd.Color.B)
        $btnTextColor.BackColor = $cd.Color
        Set-UserStyleSelected
        Set-ProjectDirty
	Schedule-PreviewRefresh
    }
})

# czcionka
$lblFontSettings = New-Object System.Windows.Forms.Label
$lblFontSettings.Text = "Czcionka:"
$lblFontSettings.AutoSize = $true
$lblFontSettings.Margin = '0,4,0,2'
$tlpTextSettings.Controls.Add($lblFontSettings,0,2) | Out-Null

$cmbFontSettings = New-Object System.Windows.Forms.ComboBox
$cmbFontSettings.DropDownStyle = 'DropDownList'
$cmbFontSettings.Width  = 180
$cmbFontSettings.Margin = '0,0,0,6'
$tlpTextSettings.Controls.Add($cmbFontSettings,0,3) | Out-Null

$installedFontsSettings = New-Object System.Drawing.Text.InstalledFontCollection
$familiesSettings = $installedFontsSettings.Families | Sort-Object Name
foreach ($fam in $familiesSettings) { [void]$cmbFontSettings.Items.Add($fam.Name) }

if ($cmbFontSettings.Items.Contains($global:FontName)) {
    $cmbFontSettings.SelectedItem = $global:FontName
} elseif ($cmbFontSettings.Items.Count -gt 0) {
    $cmbFontSettings.SelectedIndex = 0
    $global:FontName = $cmbFontSettings.SelectedItem
}

# cień tekstu
$chkTextShadow = New-Object System.Windows.Forms.CheckBox
$chkTextShadow.Text = "Cień tekstu"
$chkTextShadow.Checked = $global:TextShadowEnabled
$chkTextShadow.AutoSize = $true
$chkTextShadow.Margin = '0,2,0,0'
$tlpTextSettings.Controls.Add($chkTextShadow,0,4) | Out-Null

$chkTextShadow.Add_CheckedChanged({
    $global:TextShadowEnabled = $chkTextShadow.Checked
    Set-UserStyleSelected
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

# event zmiany czcionki → wpływa na edytor i PDF
$cmbFontSettings.Add_SelectedIndexChanged({
    $name = $cmbFontSettings.SelectedItem -as [string]
    if ([string]::IsNullOrWhiteSpace($name)) { return }
    try {
        $currentSize = $global:RichText.Font.Size
        $newFont = New-Object System.Drawing.Font($name, $currentSize)
        $global:RichText.Font = $newFont
        $global:FontName = $name
    } catch {}
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

# udostępniamy dalej (inne fragmenty skryptu tego używają)
$global:CmbFontSettings = $cmbFontSettings
$global:BtnTextColor    = $btnTextColor
$global:ChkTextShadow   = $chkTextShadow

# ------------------ ZAKŁADKA: SZATA ------------------

$tabSzataSettings = New-Object System.Windows.Forms.TabPage
$tabSzataSettings.Text     = "Szata"
$tabSzataSettings.BackColor = [System.Drawing.Color]::FromArgb(245,245,249)
$tabSzataSettings.Padding  = '4,4,4,4'
$tabsAppearance.TabPages.Add($tabSzataSettings) | Out-Null

# wewnętrzny TabControl: Tło / Obramowanie / Linia / Dodatki
$tabsSzata = New-Object System.Windows.Forms.TabControl
$tabsSzata.Dock   = 'Fill'
$tabsSzata.Padding = New-Object System.Drawing.Point(10,3)
$tabsSzata.ItemSize = New-Object System.Drawing.Size(90,22)
$tabsSzata.SizeMode = 'Fixed'
$tabSzataSettings.Controls.Add($tabsSzata)

# ===== PODZAKŁADKA: TŁO =====

$pageBg = New-Object System.Windows.Forms.TabPage
$pageBg.Text = "Tło"
$pageBg.BackColor = [System.Drawing.Color]::FromArgb(245,245,249)
$pageBg.Padding = '6,6,6,6'
$tabsSzata.TabPages.Add($pageBg) | Out-Null

$tlpBg = New-Object System.Windows.Forms.TableLayoutPanel
$tlpBg.Dock = 'Top'
$tlpBg.AutoSize = $true
$tlpBg.AutoSizeMode = 'GrowAndShrink'
$tlpBg.ColumnCount = 1
$tlpBg.Padding = '2,2,2,2'
$pageBg.Controls.Add($tlpBg)

$chkBgEnabled = New-Object System.Windows.Forms.CheckBox
$chkBgEnabled.Text = "Aktywne tło szaty"
$chkBgEnabled.Checked = ($global:CardBackgroundMode -ne "none")
$chkBgEnabled.AutoSize = $true
$chkBgEnabled.Margin = '0,0,0,4'
$tlpBg.Controls.Add($chkBgEnabled,0,0) | Out-Null

# tryb: kolor / gradient
$panelBgMode = New-Object System.Windows.Forms.FlowLayoutPanel
$panelBgMode.FlowDirection = 'LeftToRight'
$panelBgMode.AutoSize = $true
$panelBgMode.WrapContents = $false
$panelBgMode.Margin = '0,0,0,4'
$tlpBg.Controls.Add($panelBgMode,0,1) | Out-Null

$rbBgColor = New-Object System.Windows.Forms.RadioButton
$rbBgColor.Text = "Kolor"
$rbBgColor.AutoSize = $true
$rbBgColor.Margin = '0,0,10,0'
$panelBgMode.Controls.Add($rbBgColor)

$rbBgGradient = New-Object System.Windows.Forms.RadioButton
$rbBgGradient.Text = "Gradient"
$rbBgGradient.AutoSize = $true
$panelBgMode.Controls.Add($rbBgGradient)

# kolor tła
$btnBgColor = New-Object System.Windows.Forms.Button
$btnBgColor.Text = "Kolor tła"
$btnBgColor.Width = 120
$btnBgColor.Height = 24
$btnBgColor.Margin = '0,0,0,4'
$btnBgColor.BackColor = [System.Drawing.ColorTranslator]::FromHtml($global:CardBackgroundColorHex)
$tlpBg.Controls.Add($btnBgColor,0,2) | Out-Null

# gradient: typ + 2 kolory
$lblBgGrad = New-Object System.Windows.Forms.Label
$lblBgGrad.Text = "Gradient:"
$lblBgGrad.AutoSize = $true
$lblBgGrad.Margin = '0,6,0,2'
$tlpBg.Controls.Add($lblBgGrad,0,3) | Out-Null

$panelBgGrad = New-Object System.Windows.Forms.FlowLayoutPanel
$panelBgGrad.FlowDirection = 'LeftToRight'
$panelBgGrad.AutoSize = $true
$panelBgGrad.WrapContents = $false
$panelBgGrad.Margin = '0,0,0,0'
$tlpBg.Controls.Add($panelBgGrad,0,4) | Out-Null

$cmbBgType = New-Object System.Windows.Forms.ComboBox
$cmbBgType.DropDownStyle = 'DropDownList'
$cmbBgType.Width = 160
$cmbBgType.Margin = "0,0,0,4"
$cmbBgType.Items.AddRange(@("Pionowy","Poziomy","Ukośny 45°","Radialny"))
$cmbBgType.SelectedIndex = 0
$panelBgGrad.Controls.Add($cmbBgType)

$btnBgGrad1 = New-Object System.Windows.Forms.Button
$btnBgGrad1.Text = "Kolor 1"
$btnBgGrad1.Width = 80
$btnBgGrad1.Height = 24
$btnBgGrad1.Margin = '0,0,4,0'
$btnBgGrad1.BackColor = [System.Drawing.ColorTranslator]::FromHtml($global:CardBackgroundGradColor1Hex)
$panelBgGrad.Controls.Add($btnBgGrad1)

$btnBgGrad2 = New-Object System.Windows.Forms.Button
$btnBgGrad2.Text = "Kolor 2"
$btnBgGrad2.Width = 80
$btnBgGrad2.Height = 24
$btnBgGrad2.Margin = '0,0,0,0'
$btnBgGrad2.BackColor = [System.Drawing.ColorTranslator]::FromHtml($global:CardBackgroundGradColor2Hex)
$panelBgGrad.Controls.Add($btnBgGrad2)

# inicjalizacja stanu radiobuttonów i comboboxa
switch ($global:CardBackgroundMode) {
    'gradient' { $rbBgGradient.Checked = $true }
    'color'    { $rbBgColor.Checked    = $true }
    default    { $rbBgColor.Checked    = $true }
}

if ($global:CardBackgroundGradientType -eq 'radial') {
    $cmbBgType.SelectedItem = "Radialny"
} else {
    switch ($global:CardBackgroundGradientDirection) {
        'vertical'   { $cmbBgType.SelectedItem = "Pionowy" }
        'horizontal' { $cmbBgType.SelectedItem = "Poziomy" }
        default      { $cmbBgType.SelectedItem = "Ukośny 45°" }
    }
}

function Update-BgControlsState {
    $enabled = $chkBgEnabled.Checked

    # włączanie/wyłączanie przełączników kolor/gradient
    $rbBgColor.Enabled    = $enabled
    $rbBgGradient.Enabled = $enabled

    if (-not $enabled) {
        # wszystko schowane gdy tło jest wyłączone
        $btnBgColor.Visible = $false
        $lblBgGrad.Visible  = $false
        $panelBgGrad.Visible = $false
        return
    }

    if ($rbBgColor.Checked) {
        # tryb KOLOR: pokazujemy przycisk koloru, chowamy gradient
        $btnBgColor.Visible = $true
        $btnBgColor.Enabled = $true

        $lblBgGrad.Visible   = $false
        $panelBgGrad.Visible = $false
    } else {
        # tryb GRADIENT: chowamy kolor, pokazujemy panel gradientu
        $btnBgColor.Visible = $false

        $lblBgGrad.Visible   = $true
        $panelBgGrad.Visible = $true

        $cmbBgType.Enabled = $true
        $btnBgGrad1.Enabled = $true
        $btnBgGrad2.Enabled = $true
    }
}


$chkBgEnabled.Add_CheckedChanged({
    if (-not $chkBgEnabled.Checked) {
        $global:CardBackgroundMode = "none"
    } else {
        if ($rbBgGradient.Checked) { $global:CardBackgroundMode = "gradient" }
        else { $global:CardBackgroundMode = "color" }
    }
    Update-BgControlsState
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

$rbBgColor.Add_CheckedChanged({
    if ($rbBgColor.Checked -and $chkBgEnabled.Checked) {
        $global:CardBackgroundMode = "color"
    }
    Update-BgControlsState
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

$rbBgGradient.Add_CheckedChanged({
    if ($rbBgGradient.Checked -and $chkBgEnabled.Checked) {
        $global:CardBackgroundMode = "gradient"
    }
    Update-BgControlsState
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

$btnBgColor.Add_Click({
    $cd = New-Object System.Windows.Forms.ColorDialog
    $cd.Color = [System.Drawing.ColorTranslator]::FromHtml($global:CardBackgroundColorHex)
    if ($cd.ShowDialog() -eq 'OK') {
        $global:CardBackgroundColorHex = ("#{0:X2}{1:X2}{2:X2}" -f $cd.Color.R,$cd.Color.G,$cd.Color.B)
        $btnBgColor.BackColor = $cd.Color
        if ($global:CardBackgroundMode -eq "none") { $global:CardBackgroundMode = "color" }
        Set-UserStyleSelected
		Set-ProjectDirty
	Schedule-PreviewRefresh
    }
})

$cmbBgType.Add_SelectedIndexChanged({
    switch ($cmbBgType.SelectedItem) {
        "Radialny" {
            $global:CardBackgroundGradientType      = "radial"
            $global:CardBackgroundGradientDirection = "diag"   # na wszelki wypadek
        }
        "Pionowy" {
            $global:CardBackgroundGradientType      = "linear"
            $global:CardBackgroundGradientDirection = "vertical"
        }
        "Poziomy" {
            $global:CardBackgroundGradientType      = "linear"
            $global:CardBackgroundGradientDirection = "horizontal"
        }
        "Ukośny 45°" {
            $global:CardBackgroundGradientType      = "linear"
            $global:CardBackgroundGradientDirection = "diag"
        }
    }
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

$btnBgGrad1.Add_Click({
    $cd = New-Object System.Windows.Forms.ColorDialog
    $cd.Color = [System.Drawing.ColorTranslator]::FromHtml($global:CardBackgroundGradColor1Hex)
    if ($cd.ShowDialog() -eq 'OK') {
        $global:CardBackgroundGradColor1Hex = ("#{0:X2}{1:X2}{2:X2}" -f $cd.Color.R,$cd.Color.G,$cd.Color.B)
        $btnBgGrad1.BackColor = $cd.Color
        if ($global:CardBackgroundMode -eq "none") { $global:CardBackgroundMode = "gradient" }
        Set-ProjectDirty
		Schedule-PreviewRefresh
    }
})

$btnBgGrad2.Add_Click({
    $cd = New-Object System.Windows.Forms.ColorDialog
    $cd.Color = [System.Drawing.ColorTranslator]::FromHtml($global:CardBackgroundGradColor2Hex)
    if ($cd.ShowDialog() -eq 'OK') {
        $global:CardBackgroundGradColor2Hex = ("#{0:X2}{1:X2}{2:X2}" -f $cd.Color.R,$cd.Color.G,$cd.Color.B)
        $btnBgGrad2.BackColor = $cd.Color
        if ($global:CardBackgroundMode -eq "none") { $global:CardBackgroundMode = "gradient" }
        Set-ProjectDirty
		Schedule-PreviewRefresh
    }
})

Update-BgControlsState

# ===== PODZAKŁADKA: OBRAMOWANIE =====

$pageBorder = New-Object System.Windows.Forms.TabPage
$pageBorder.Text = "Obramowanie"
$pageBorder.BackColor = [System.Drawing.Color]::FromArgb(245,245,249)
$pageBorder.Padding = '6,6,6,6'
$tabsSzata.TabPages.Add($pageBorder) | Out-Null

$tlpBorder = New-Object System.Windows.Forms.TableLayoutPanel
$tlpBorder.Dock = 'Top'
$tlpBorder.AutoSize = $true
$tlpBorder.AutoSizeMode = 'GrowAndShrink'
$tlpBorder.ColumnCount = 1
$tlpBorder.Padding = '2,2,2,2'
$pageBorder.Controls.Add($tlpBorder)

$chkBorderEnabled = New-Object System.Windows.Forms.CheckBox
$chkBorderEnabled.Text = "Obramowanie aktywne"
$chkBorderEnabled.Checked = $global:CardBorderEnabled
$chkBorderEnabled.AutoSize = $true
$chkBorderEnabled.Margin = '0,0,0,4'
$tlpBorder.Controls.Add($chkBorderEnabled,0,0) | Out-Null

$btnBorderColor = New-Object System.Windows.Forms.Button
$btnBorderColor.Text = "Kolor obramowania"
$btnBorderColor.Width = 160
$btnBorderColor.Height = 24
$btnBorderColor.Margin = '0,0,0,4'
$btnBorderColor.BackColor = [System.Drawing.ColorTranslator]::FromHtml($global:CardBorderColorHex)
$tlpBorder.Controls.Add($btnBorderColor,0,1) | Out-Null

$panelBorderLine = New-Object System.Windows.Forms.FlowLayoutPanel
$panelBorderLine.FlowDirection = 'LeftToRight'
$panelBorderLine.AutoSize = $true
$panelBorderLine.WrapContents = $false
$panelBorderLine.Margin = '0,0,0,0'
$tlpBorder.Controls.Add($panelBorderLine,0,2) | Out-Null

$lblBorderWidth = New-Object System.Windows.Forms.Label
$lblBorderWidth.Text = "Grubość:"
$lblBorderWidth.AutoSize = $true
$lblBorderWidth.Margin = '0,4,4,0'
$panelBorderLine.Controls.Add($lblBorderWidth)

$numBorderWidth = New-Object System.Windows.Forms.NumericUpDown
$numBorderWidth.Minimum = 0
$numBorderWidth.Maximum = 10
$numBorderWidth.Value   = [decimal]$global:CardBorderWidthPx
$numBorderWidth.Width   = 60
$numBorderWidth.Margin  = '0,0,8,0'
$panelBorderLine.Controls.Add($numBorderWidth)

$cmbBorderStyle = New-Object System.Windows.Forms.ComboBox
$cmbBorderStyle.DropDownStyle = 'DropDownList'
$cmbBorderStyle.Width = 120
$cmbBorderStyle.Items.AddRange(@(
    "ciągła",      # solid
    "kreskowana",  # dashed
    "kropkowana",  # dotted
    "podwójna",    # double
    "żłobkowana",  # groove
    "wypukła",     # ridge
    "wcięta",      # inset
    "wystająca"    # outset
))
$cmbBorderStyle.Margin = '0,0,0,0'
$panelBorderLine.Controls.Add($cmbBorderStyle)

switch ($global:CardBorderStyle) {
    'solid'  { $cmbBorderStyle.SelectedItem = "ciągła" }
    'dashed' { $cmbBorderStyle.SelectedItem = "kreskowana" }
    'dotted' { $cmbBorderStyle.SelectedItem = "kropkowana" }
    'double' { $cmbBorderStyle.SelectedItem = "podwójna" }
    'groove' { $cmbBorderStyle.SelectedItem = "żłobkowana" }
    'ridge'  { $cmbBorderStyle.SelectedItem = "wypukła" }
    'inset'  { $cmbBorderStyle.SelectedItem = "wcięta" }
    'outset' { $cmbBorderStyle.SelectedItem = "wystająca" }
    default  { $cmbBorderStyle.SelectedItem = "ciągła" }
}

function Update-BorderControlsEnabled {
    $enabled = $chkBorderEnabled.Checked
    $btnBorderColor.Enabled = $enabled
    $numBorderWidth.Enabled = $enabled
    $cmbBorderStyle.Enabled = $enabled
}

$chkBorderEnabled.Add_CheckedChanged({
    $global:CardBorderEnabled = $chkBorderEnabled.Checked
    Update-BorderControlsEnabled
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

$btnBorderColor.Add_Click({
    $cd = New-Object System.Windows.Forms.ColorDialog
    $cd.Color = [System.Drawing.ColorTranslator]::FromHtml($global:CardBorderColorHex)
    if ($cd.ShowDialog() -eq 'OK') {
        $global:CardBorderColorHex = ("#{0:X2}{1:X2}{2:X2}" -f $cd.Color.R,$cd.Color.G,$cd.Color.B)
        $btnBorderColor.BackColor = $cd.Color
        Set-ProjectDirty
		Schedule-PreviewRefresh
    }
})

$numBorderWidth.Add_ValueChanged({
    $global:CardBorderWidthPx = [int]$numBorderWidth.Value
    Set-UserStyleSelected
	Set-ProjectDirty
	Schedule-PreviewRefresh
})

$cmbBorderStyle.Add_SelectedIndexChanged({
	switch ($cmbBorderStyle.SelectedItem) {
		"ciągła"      { $global:CardBorderStyle = "solid" }
		"kreskowana"  { $global:CardBorderStyle = "dashed" }
		"kropkowana"  { $global:CardBorderStyle = "dotted" }
		"podwójna"    { $global:CardBorderStyle = "double" }
		"żłobkowana"  { $global:CardBorderStyle = "groove" }
		"wypukła"     { $global:CardBorderStyle = "ridge" }
		"wcięta"      { $global:CardBorderStyle = "inset" }
		"wystająca"   { $global:CardBorderStyle = "outset" }
	}
	Set-ProjectDirty
	Schedule-PreviewRefresh
})

Update-BorderControlsEnabled

# ===== PODZAKŁADKA: LINIA ŚRODKOWA =====

$pageLine = New-Object System.Windows.Forms.TabPage
$pageLine.Text = "Środkowa linia"
$pageLine.BackColor = [System.Drawing.Color]::FromArgb(245,245,249)
$pageLine.Padding = '6,6,6,6'
$tabsSzata.TabPages.Add($pageLine) | Out-Null

$tlpLine = New-Object System.Windows.Forms.TableLayoutPanel
$tlpLine.Dock = 'Top'
$tlpLine.AutoSize = $true
$tlpLine.AutoSizeMode = 'GrowAndShrink'
$tlpLine.ColumnCount = 1
$tlpLine.Padding = '2,2,2,2'
$pageLine.Controls.Add($tlpLine)

$chkLineEnabled = New-Object System.Windows.Forms.CheckBox
$chkLineEnabled.Text = "Środkowa linia aktywna"
$chkLineEnabled.Checked = $global:CardLineEnabled
$chkLineEnabled.AutoSize = $true
$chkLineEnabled.Margin = '0,0,0,4'
$tlpLine.Controls.Add($chkLineEnabled,0,0) | Out-Null

$btnLineColor = New-Object System.Windows.Forms.Button
$btnLineColor.Text = "Kolor linii"
$btnLineColor.Width = 120
$btnLineColor.Height = 24
$btnLineColor.Margin = '0,0,0,4'
$btnLineColor.BackColor = [System.Drawing.ColorTranslator]::FromHtml($global:CardLineColorHex)
$tlpLine.Controls.Add($btnLineColor,0,1) | Out-Null

$panelLine = New-Object System.Windows.Forms.FlowLayoutPanel
$panelLine.FlowDirection = 'LeftToRight'
$panelLine.AutoSize = $true
$panelLine.WrapContents = $false
$panelLine.Margin = '0,0,0,0'
$tlpLine.Controls.Add($panelLine,0,2) | Out-Null

$lblLineWidth = New-Object System.Windows.Forms.Label
$lblLineWidth.Text = "Grubość:"
$lblLineWidth.AutoSize = $true
$lblLineWidth.Margin = '0,4,4,0'
$panelLine.Controls.Add($lblLineWidth)

$numLineWidth = New-Object System.Windows.Forms.NumericUpDown
$numLineWidth.Minimum = 0
$numLineWidth.Maximum = 10
$numLineWidth.Value   = [decimal]$global:CardLineWidthPx
$numLineWidth.Width   = 60
$numLineWidth.Margin  = '0,0,8,0'
$panelLine.Controls.Add($numLineWidth)

$cmbLineStyle = New-Object System.Windows.Forms.ComboBox
$cmbLineStyle.DropDownStyle = 'DropDownList'
$cmbLineStyle.Width = 120
$cmbLineStyle.Items.AddRange(@("ciągła","kreskowana","kropkowana"))
$cmbLineStyle.Margin = '0,0,0,0'
$panelLine.Controls.Add($cmbLineStyle)

switch ($global:CardLineStyle) {
    'dashed' { $cmbLineStyle.SelectedItem = "kreskowana" }
    'dotted' { $cmbLineStyle.SelectedItem = "kropkowana" }
    default  { $cmbLineStyle.SelectedItem = "ciągła" }
}

function Update-LineControlsEnabled {
    $enabled = $chkLineEnabled.Checked
    $btnLineColor.Enabled = $enabled
    $numLineWidth.Enabled = $enabled
    $cmbLineStyle.Enabled = $enabled
}

$chkLineEnabled.Add_CheckedChanged({
    $global:CardLineEnabled = $chkLineEnabled.Checked
    Update-LineControlsEnabled
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

$btnLineColor.Add_Click({
    $cd = New-Object System.Windows.Forms.ColorDialog
    $cd.Color = [System.Drawing.ColorTranslator]::FromHtml($global:CardLineColorHex)
    if ($cd.ShowDialog() -eq 'OK') {
        $global:CardLineColorHex = ("#{0:X2}{1:X2}{2:X2}" -f $cd.Color.R,$cd.Color.G,$cd.Color.B)
        $btnLineColor.BackColor = $cd.Color
        Set-ProjectDirty
	Schedule-PreviewRefresh
    }
})

$numLineWidth.Add_ValueChanged({
    $global:CardLineWidthPx = [int]$numLineWidth.Value
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

$cmbLineStyle.Add_SelectedIndexChanged({
    switch ($cmbLineStyle.SelectedItem) {
        "kreskowana" { $global:CardLineStyle = "dashed" }
        "kropkowana" { $global:CardLineStyle = "dotted" }
        default      { $global:CardLineStyle = "solid" }
    }
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

Update-LineControlsEnabled

# ===== PODZAKŁADKA: DODATKI (CIEŃ + ROGI) =====

$pageExtras = New-Object System.Windows.Forms.TabPage
$pageExtras.Text = "Dodatki"
$pageExtras.BackColor = [System.Drawing.Color]::FromArgb(245,245,249)
$pageExtras.Padding = '6,6,6,6'
$tabsSzata.TabPages.Add($pageExtras) | Out-Null

$tlpExtras = New-Object System.Windows.Forms.TableLayoutPanel
$tlpExtras.Dock = 'Top'
$tlpExtras.AutoSize = $true
$tlpExtras.AutoSizeMode = 'GrowAndShrink'
$tlpExtras.ColumnCount = 1
$tlpExtras.Padding = '2,2,2,2'
$pageExtras.Controls.Add($tlpExtras)

$chkCardShadow = New-Object System.Windows.Forms.CheckBox
$chkCardShadow.Text = "Cień szaty"
$chkCardShadow.Checked = $global:CardShadowEnabled
$chkCardShadow.AutoSize = $true
$chkCardShadow.Margin = '0,0,0,4'
$tlpExtras.Controls.Add($chkCardShadow,0,0) | Out-Null

$chkCardRounded = New-Object System.Windows.Forms.CheckBox
$chkCardRounded.Text = "Zaokrąglone rogi szaty"
$chkCardRounded.Checked = $global:CardRoundedEnabled
$chkCardRounded.AutoSize = $true
$chkCardRounded.Margin = '0,2,0,0'
$tlpExtras.Controls.Add($chkCardRounded,0,1) | Out-Null

$chkCardShadow.Add_CheckedChanged({
    $global:CardShadowEnabled = $chkCardShadow.Checked
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

$chkCardRounded.Add_CheckedChanged({
    $global:CardRoundedEnabled = $chkCardRounded.Checked
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

# wreszcie dodaj panel ustawień do prawej kolumny
$splitRight.Panel2.Controls.Add($settingsPanel) | Out-Null

# ================= ZAKŁADKA TEKST =================

$tabText = New-Object System.Windows.Forms.TabPage
$tabText.Text    = "Tekst"
$tabText.Padding = '8,8,8,8'
$tabText.BackColor = $colorPanel
$tabs.TabPages.Add($tabText) | Out-Null

$tlpText = New-Object System.Windows.Forms.TableLayoutPanel
$tlpText.Dock = 'Fill'
$tlpText.ColumnCount = 1
$tlpText.RowCount    = 3
$tlpText.RowStyles.Add((New-Object System.Windows.Forms.RowStyle(
    [System.Windows.Forms.SizeType]::AutoSize)))    | Out-Null
$tlpText.RowStyles.Add((New-Object System.Windows.Forms.RowStyle(
    [System.Windows.Forms.SizeType]::Percent,100))) | Out-Null
$tlpText.RowStyles.Add((New-Object System.Windows.Forms.RowStyle(
    [System.Windows.Forms.SizeType]::Absolute,32))) | Out-Null
$tlpText.BackColor = $colorPanel
$tabText.Controls.Add($tlpText) | Out-Null

$lblTextInfo = New-Object System.Windows.Forms.Label
$lblTextInfo.Text = "Wpisz / wklej wyrazy (po jednym w wierszu). Formatowanie B / I / U będzie odwzorowane."
$lblTextInfo.AutoSize = $true
$lblTextInfo.Margin   = '2,0,2,6'
$lblTextInfo.ForeColor = $colorTextMuted
$tlpText.Controls.Add($lblTextInfo,0,0) | Out-Null

$tlpEditorText = New-Object System.Windows.Forms.TableLayoutPanel
$tlpEditorText.Dock = 'Fill'
$tlpEditorText.ColumnCount = 1
$tlpEditorText.RowCount    = 2
$tlpEditorText.RowStyles.Add((New-Object System.Windows.Forms.RowStyle(
    [System.Windows.Forms.SizeType]::AutoSize)))    | Out-Null
$tlpEditorText.RowStyles.Add((New-Object System.Windows.Forms.RowStyle(
    [System.Windows.Forms.SizeType]::Percent,100))) | Out-Null
$tlpEditorText.BackColor = $colorPanel
$tlpText.Controls.Add($tlpEditorText,0,1) | Out-Null

$toolStripText = New-Object System.Windows.Forms.ToolStrip
$toolStripText.RenderMode       = 'Professional'
$toolStripText.GripStyle        = 'Hidden'
$toolStripText.BackColor        = $colorPanel
$toolStripText.Dock             = 'Fill'
$toolStripText.ShowItemToolTips = $true
$toolStripText.ImageScalingSize = New-Object System.Drawing.Size(1,1)
$tlpEditorText.Controls.Add($toolStripText,0,0) | Out-Null

$btnSize   = New-Object System.Drawing.Size(34, 30)
$btnMargin = '3,3,3,3'

$btnBoldItem = New-Object System.Windows.Forms.ToolStripButton
$btnBoldItem.Text         = 'B'
$btnBoldItem.DisplayStyle = 'Text'
$btnBoldItem.Font         = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$btnBoldItem.CheckOnClick = $true
$btnBoldItem.ToolTipText  = 'Pogrubienie'
$btnBoldItem.AutoSize     = $false
$btnBoldItem.Size         = $btnSize
$btnBoldItem.Margin       = $btnMargin
$btnBoldItem.ForeColor    = [System.Drawing.Color]::FromArgb(30,50,110)
$btnBoldItem.BackColor    = [System.Drawing.Color]::FromArgb(220,232,252)
[void]$toolStripText.Items.Add($btnBoldItem)

$btnItalicItem = New-Object System.Windows.Forms.ToolStripButton
$btnItalicItem.Text         = 'I'
$btnItalicItem.DisplayStyle = 'Text'
$btnItalicItem.Font         = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Italic)
$btnItalicItem.CheckOnClick = $true
$btnItalicItem.ToolTipText  = 'Kursywa'
$btnItalicItem.AutoSize     = $false
$btnItalicItem.Size         = $btnSize
$btnItalicItem.Margin       = $btnMargin
$btnItalicItem.ForeColor    = [System.Drawing.Color]::FromArgb(20,90,60)
$btnItalicItem.BackColor    = [System.Drawing.Color]::FromArgb(218,242,232)
[void]$toolStripText.Items.Add($btnItalicItem)

$btnUnderlineItem = New-Object System.Windows.Forms.ToolStripButton
$btnUnderlineItem.Text         = 'U'
$btnUnderlineItem.DisplayStyle = 'Text'
$btnUnderlineItem.Font         = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Underline)
$btnUnderlineItem.CheckOnClick = $true
$btnUnderlineItem.ToolTipText  = 'Podkreślenie'
$btnUnderlineItem.AutoSize     = $false
$btnUnderlineItem.Size         = $btnSize
$btnUnderlineItem.Margin       = $btnMargin
$btnUnderlineItem.ForeColor    = [System.Drawing.Color]::FromArgb(120,80,20)
$btnUnderlineItem.BackColor    = [System.Drawing.Color]::FromArgb(252,237,214)
[void]$toolStripText.Items.Add($btnUnderlineItem)

$sepTextRight = New-Object System.Windows.Forms.ToolStripSeparator
$sepTextRight.Alignment = [System.Windows.Forms.ToolStripItemAlignment]::Right
[void]$toolStripText.Items.Add($sepTextRight)

$btnClearText = New-Object System.Windows.Forms.ToolStripButton
$btnClearText.Text         = 'X'
$btnClearText.DisplayStyle = 'Text'
$btnClearText.Alignment    = [System.Windows.Forms.ToolStripItemAlignment]::Right
$btnClearText.ToolTipText  = 'Wyczyść tekst'
$btnClearText.AutoSize     = $false
$btnClearText.Size         = $btnSize
$btnClearText.Margin       = $btnMargin
$btnClearText.Font         = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$btnClearText.ForeColor    = [System.Drawing.Color]::FromArgb(150,30,30)
$btnClearText.BackColor    = [System.Drawing.Color]::FromArgb(252,226,226)
[void]$toolStripText.Items.Add($btnClearText)

$rtbWords = New-Object System.Windows.Forms.RichTextBox
$rtbWords.Dock       = 'Fill'
$rtbWords.Multiline  = $true
$rtbWords.AcceptsTab = $true
$rtbWords.DetectUrls = $false
$rtbWords.Font       = New-Object System.Drawing.Font($global:FontName, 14)
$rtbWords.Margin     = '0,4,0,0'
$rtbWords.BorderStyle = 'FixedSingle'
$rtbWords.BackColor   = [System.Drawing.Color]::White
$tlpEditorText.Controls.Add($rtbWords,0,1) | Out-Null

$global:RichText = $rtbWords

$bottomTextPanel = New-Object System.Windows.Forms.TableLayoutPanel
$bottomTextPanel.ColumnCount = 1
$bottomTextPanel.Dock        = 'Fill'
$bottomTextPanel.Margin      = '0,6,0,0'
$bottomTextPanel.Padding     = '2,0,2,0'
$bottomTextPanel.BackColor   = $colorPanel
$tlpText.Controls.Add($bottomTextPanel,0,2) | Out-Null

$lblCountText = New-Object System.Windows.Forms.Label
$lblCountText.Text    = "Wyrazów: 0"
$lblCountText.AutoSize = $true
$lblCountText.Margin   = '2,4,2,0'
$lblCountText.ForeColor = $colorTextMuted
$bottomTextPanel.Controls.Add($lblCountText,0,0) | Out-Null

$btnClearText.Add_Click({
    $rtbWords.Clear()
})

# ================= ZAKŁADKA OBRAZKI =================

$tabImg = New-Object System.Windows.Forms.TabPage
$tabImg.Text    = 'Obrazki'
$tabImg.Padding = '8,8,8,8'
$tabImg.BackColor = $colorPanel
$tabs.TabPages.Add($tabImg) | Out-Null

$tlpImg = New-Object System.Windows.Forms.TableLayoutPanel
$tlpImg.Dock = 'Fill'
$tlpImg.ColumnCount = 1
$tlpImg.RowCount    = 3
$tlpImg.RowStyles.Add((New-Object System.Windows.Forms.RowStyle(
    [System.Windows.Forms.SizeType]::AutoSize)))    | Out-Null
$tlpImg.RowStyles.Add((New-Object System.Windows.Forms.RowStyle(
    [System.Windows.Forms.SizeType]::Percent,100))) | Out-Null
$tlpImg.RowStyles.Add((New-Object System.Windows.Forms.RowStyle(
    [System.Windows.Forms.SizeType]::Absolute,32))) | Out-Null
$tlpImg.Margin  = '0,0,0,0'
$tlpImg.Padding = '0,0,0,0'
$tlpImg.BackColor = $colorPanel
$tabImg.Controls.Add($tlpImg) | Out-Null

$lblImgInfo = New-Object System.Windows.Forms.Label
$lblImgInfo.Text = 'Dodaj obrazy (PNG / JPG / JPEG / JFIF / BMP / GIF / TIF / TIFF / ICO).'
$lblImgInfo.AutoSize = $true
$lblImgInfo.Margin   = '2,0,2,6'
$lblImgInfo.ForeColor = $colorTextMuted
$tlpImg.Controls.Add($lblImgInfo,0,0) | Out-Null

$tlpImgEditor = New-Object System.Windows.Forms.TableLayoutPanel
$tlpImgEditor.Dock = 'Fill'
$tlpImgEditor.ColumnCount = 1
$tlpImgEditor.RowCount    = 2
$tlpImgEditor.RowStyles.Add((New-Object System.Windows.Forms.RowStyle(
    [System.Windows.Forms.SizeType]::AutoSize)))    | Out-Null
$tlpImgEditor.RowStyles.Add((New-Object System.Windows.Forms.RowStyle(
    [System.Windows.Forms.SizeType]::Percent,100))) | Out-Null
$tlpImgEditor.Margin  = '0,0,0,0'
$tlpImgEditor.Padding = '0,0,0,0'
$tlpImgEditor.BackColor = $colorPanel
$tlpImg.Controls.Add($tlpImgEditor,0,1) | Out-Null

$toolStripImg = New-Object System.Windows.Forms.ToolStrip
$toolStripImg.RenderMode       = 'Professional'
$toolStripImg.GripStyle        = 'Hidden'
$toolStripImg.BackColor        = $colorPanel
$toolStripImg.Dock             = 'Fill'
$toolStripImg.ShowItemToolTips = $true
$toolStripImg.ImageScalingSize = New-Object System.Drawing.Size(1,1)
$tlpImgEditor.Controls.Add($toolStripImg,0,0) | Out-Null

$btnSizeImg   = New-Object System.Drawing.Size(34, 30)
$btnMarginImg = '3,3,3,3'

$btnAddImg = New-Object System.Windows.Forms.ToolStripButton
$btnAddImg.Text         = '+'
$btnAddImg.DisplayStyle = 'Text'
$btnAddImg.ToolTipText  = 'Dodaj obraz(y)'
$btnAddImg.AutoSize     = $false
$btnAddImg.Size         = $btnSizeImg
$btnAddImg.Margin       = $btnMarginImg
$btnAddImg.Font         = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$btnAddImg.ForeColor    = [System.Drawing.Color]::FromArgb(20,120,40)
$btnAddImg.BackColor    = [System.Drawing.Color]::FromArgb(216,242,225)
[void]$toolStripImg.Items.Add($btnAddImg)

$btnRemoveImg = New-Object System.Windows.Forms.ToolStripButton
$btnRemoveImg.Text         = '–'
$btnRemoveImg.DisplayStyle = 'Text'
$btnRemoveImg.ToolTipText  = 'Usuń zaznaczone obrazki'
$btnRemoveImg.AutoSize     = $false
$btnRemoveImg.Size         = $btnSizeImg
$btnRemoveImg.Margin       = $btnMarginImg
$btnRemoveImg.Font         = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$btnRemoveImg.ForeColor    = [System.Drawing.Color]::FromArgb(150,95,10)
$btnRemoveImg.BackColor    = [System.Drawing.Color]::FromArgb(252,239,214)
[void]$toolStripImg.Items.Add($btnRemoveImg)

# przycisk wyboru widoku (ikony / lista)
$ddView = New-Object System.Windows.Forms.ToolStripDropDownButton
$ddView.Text = "Widok"
$ddView.ToolTipText = "Widok listy obrazków"
$ddView.AutoSize = $true
$ddView.Margin = '6,3,3,3'
[void]$toolStripImg.Items.Add($ddView)

$miIcons = New-Object System.Windows.Forms.ToolStripMenuItem "Ikony"
$miList  = New-Object System.Windows.Forms.ToolStripMenuItem "Lista"

$ddView.DropDownItems.AddRange(@($miIcons,$miList))

$miIcons.Checked = $true   # start: widok ikon

$sepImgRight = New-Object System.Windows.Forms.ToolStripSeparator
$sepImgRight.Alignment = [System.Windows.Forms.ToolStripItemAlignment]::Right
[void]$toolStripImg.Items.Add($sepImgRight)

$btnClearImg = New-Object System.Windows.Forms.ToolStripButton
$btnClearImg.Text         = 'X'
$btnClearImg.DisplayStyle = 'Text'
$btnClearImg.Alignment    = [System.Windows.Forms.ToolStripItemAlignment]::Right
$btnClearImg.ToolTipText  = 'Wyczyść listę obrazów'
$btnClearImg.AutoSize     = $false
$btnClearImg.Size         = $btnSizeImg
$btnClearImg.Margin       = $btnMarginImg
$btnClearImg.Font         = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)
$btnClearImg.ForeColor    = [System.Drawing.Color]::FromArgb(160,30,30)
$btnClearImg.BackColor    = [System.Drawing.Color]::FromArgb(252,226,226)
[void]$toolStripImg.Items.Add($btnClearImg)

# duże ikony
$imgList = New-Object System.Windows.Forms.ImageList
$imgList.ColorDepth = 'Depth32Bit'
$imgList.ImageSize  = New-Object System.Drawing.Size(96,96)

# małe ikony do widoku "Lista"
$imgListSmall = New-Object System.Windows.Forms.ImageList
$imgListSmall.ColorDepth = 'Depth32Bit'
$imgListSmall.ImageSize  = New-Object System.Drawing.Size(32,32)

$listView = New-Object System.Windows.Forms.ListView
$listView.View           = 'LargeIcon'
$listView.LargeImageList = $imgList
$listView.SmallImageList = $imgListSmall

$listView.MultiSelect    = $true
$listView.HideSelection  = $false
$listView.Dock           = 'Fill'
$listView.Margin         = '0,4,0,0'
$listView.BorderStyle    = 'FixedSingle'
$listView.BackColor      = [System.Drawing.Color]::White

# KLUCZOWE USTAWIENIA DLA KRESKI I UKŁADU
$listView.AllowDrop      = $true
$listView.Sorting        = [System.Windows.Forms.SortOrder]::None

$listView.AutoArrange    = $true          # musi być TRUE dla InsertionMark
$listView.Alignment      = [System.Windows.Forms.ListViewAlignment]::Top

$listView.InsertionMark.Color = [System.Drawing.Color]::Black

$tlpImgEditor.Controls.Add($listView,0,1) | Out-Null


# indeks przeciąganego elementu
$global:DragImgIndex = -1

$listView.Add_ItemDrag({
    param($sender, $e)

    if ($listView.SelectedIndices.Count -eq 0) { return }

    $global:DragImgIndex = $listView.SelectedIndices[0]
    $listView.DoDragDrop($e.Item, [System.Windows.Forms.DragDropEffects]::Move) | Out-Null
})

$listView.Add_DragEnter({
    param($sender, $e)

    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
    elseif ($e.Data.GetDataPresent([System.Windows.Forms.ListViewItem])) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
    }
    else {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
})

$listView.Add_DragOver({
    param($sender, $e)

    $pt = $listView.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))

    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
    elseif ($e.Data.GetDataPresent([System.Windows.Forms.ListViewItem])) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
    }
    else {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::None
        $listView.InsertionMark.Index = -1
        return
    }

    # Najbliższy element pod kursorem
    $targetIndex = $listView.InsertionMark.NearestIndex($pt)

    if ($targetIndex -lt 0) {
        $listView.InsertionMark.Index = -1
        return
    }

    $itemBounds = $listView.Items[$targetIndex].GetBounds(
        [System.Windows.Forms.ItemBoundsPortion]::Entire
    )

    if ($pt.X -gt $itemBounds.Left + ($itemBounds.Width / 2)) {
        $listView.InsertionMark.AppearsAfterItem = $true
    } else {
        $listView.InsertionMark.AppearsAfterItem = $false
    }

    $listView.InsertionMark.Index = $targetIndex
})

$listView.Add_DragLeave({
    param($sender, $e)
    $listView.InsertionMark.Index = -1
})

$listView.Add_DragDrop({
    param($sender, $e)

    # indeks z insertion mark
    $insertIndex = $listView.InsertionMark.Index
    if ($insertIndex -ge 0 -and $listView.InsertionMark.AppearsAfterItem) {
        $insertIndex++
    }
    $listView.InsertionMark.Index = -1

    # 1) przeciąganie z Explorera (FileDrop)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {

        $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)

        if ($insertIndex -lt 0 -or $insertIndex -gt $listView.Items.Count) {
            $insertIndex = $listView.Items.Count
        }

        foreach ($path in $files) {
            if (-not (Test-Path $path)) { continue }
            $ext = [System.IO.Path]::GetExtension($path).ToLower()
            if ($ext -notin @('.png','.jpg','.jpeg','.jfif','.bmp','.gif','.tif','.tiff','.ico')) { continue }

            try {
                $img = [System.Drawing.Image]::FromFile($path)
                $key = [Guid]::NewGuid().ToString()

                # duża miniatura
                $thumbLarge = New-Object System.Drawing.Bitmap($imgList.ImageSize.Width, $imgList.ImageSize.Height)
                $gL = [System.Drawing.Graphics]::FromImage($thumbLarge)
                $gL.Clear([System.Drawing.Color]::White)
                $gL.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $scaleL = [Math]::Min($imgList.ImageSize.Width / $img.Width, $imgList.ImageSize.Height / $img.Height)
                $dwL = [int]([Math]::Round($img.Width * $scaleL))
                $dhL = [int]([Math]::Round($img.Height * $scaleL))
                $dxL = [int](($imgList.ImageSize.Width - $dwL)/2)
                $dyL = [int](($imgList.ImageSize.Height - $dhL)/2)
                $gL.DrawImage($img,$dxL,$dyL,$dwL,$dhL)
                $gL.Dispose()
                $imgList.Images.Add($key, $thumbLarge) | Out-Null
                $thumbLarge.Dispose()

                # mała miniatura
                $thumbSmall = New-Object System.Drawing.Bitmap($imgListSmall.ImageSize.Width, $imgListSmall.ImageSize.Height)
                $gS = [System.Drawing.Graphics]::FromImage($thumbSmall)
                $gS.Clear([System.Drawing.Color]::White)
                $gS.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $scaleS = [Math]::Min($imgListSmall.ImageSize.Width / $img.Width, $imgListSmall.ImageSize.Height / $img.Height)
                $dwS = [int]([Math]::Round($img.Width * $scaleS))
                $dhS = [int]([Math]::Round($img.Height * $scaleS))
                $dxS = [int](($imgListSmall.ImageSize.Width - $dwS)/2)
                $dyS = [int](($imgListSmall.ImageSize.Height - $dhS)/2)
                $gS.DrawImage($img,$dxS,$dyS,$dwS,$dhS)
                $gS.Dispose()
                $imgListSmall.Images.Add($key, $thumbSmall) | Out-Null
                $thumbSmall.Dispose()

                $img.Dispose()

                $iconIndex = $imgList.Images.IndexOfKey($key)
                $item = New-Object System.Windows.Forms.ListViewItem(
                    [System.IO.Path]::GetFileName($path),
                    $iconIndex
                )

                $listView.Items.Insert($insertIndex, $item)
                $global:ImagePaths.Insert($insertIndex, $path)
                $insertIndex++
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Nie udało się wczytać: $path") | Out-Null
            }
        }

        Update-ImgCountLabel -lbl $lblCountImg -btn $btnSavePdf
        Update-ImageOrderLabels
		$listView.ArrangeIcons()
        Set-ProjectDirty
        Schedule-PreviewRefresh
        return
    }

    # 2) przeciąganie w obrębie listy
    if ($global:DragImgIndex -lt 0) { return }

    $fromIndex = $global:DragImgIndex
    $global:DragImgIndex = -1

    if ($insertIndex -lt 0 -or $insertIndex -gt $listView.Items.Count) {
        $insertIndex = $listView.Items.Count
    }

    if ($insertIndex -gt $fromIndex) {
        $insertIndex--
    }

    if ($insertIndex -eq $fromIndex) { return }

    if ($fromIndex -lt 0 -or $fromIndex -ge $global:ImagePaths.Count) { return }
    if ($insertIndex -lt 0 -or $insertIndex -gt  $global:ImagePaths.Count) { return }

    $path = $global:ImagePaths[$fromIndex]
    $global:ImagePaths.RemoveAt($fromIndex)
    $global:ImagePaths.Insert($insertIndex, $path)

    $item = $listView.Items[$fromIndex]
    $listView.Items.RemoveAt($fromIndex)
    $listView.Items.Insert($insertIndex, $item)

    $listView.SelectedIndices.Clear()
    $item.Selected = $true
    $item.Focused  = $true
    $item.EnsureVisible()

	$listView.ArrangeIcons()
    Update-ImageOrderLabels
    Set-ProjectDirty
    Schedule-PreviewRefresh
})

$bottomImgPanel = New-Object System.Windows.Forms.TableLayoutPanel
$bottomImgPanel.ColumnCount = 1
$bottomImgPanel.Dock        = 'Fill'
$bottomImgPanel.Margin      = '0,6,0,0'
$bottomImgPanel.Padding     = '2,0,2,0'
$bottomImgPanel.BackColor   = $colorPanel
$tlpImg.Controls.Add($bottomImgPanel,0,2) | Out-Null

$lblCountImg = New-Object System.Windows.Forms.Label
$lblCountImg.Text    = 'Obrazków: 0'
$lblCountImg.AutoSize = $true
$lblCountImg.Margin   = '2,4,2,0'
$lblCountImg.ForeColor = $colorTextMuted
$bottomImgPanel.Controls.Add($lblCountImg,0,0) | Out-Null

# ================= LOGIKA UI =================

function Update-ImageOrderLabels {
    for ($i = 0; $i -lt $listView.Items.Count; $i++) {
        $path = $global:ImagePaths[$i]
        $name = [System.IO.Path]::GetFileName($path)
        $num  = $i + 1
        $listView.Items[$i].Text = ("{0}. {1}" -f $num, $name)
    }
}

function Set-ImgView([string]$mode) {
    $miIcons.Checked = $false
    $miList.Checked  = $false

    switch ($mode) {
        'Icons' {
            $listView.View = 'LargeIcon'
            $miIcons.Checked = $true
        }
        'List' {
            $listView.View = 'List'   # użyje SmallImageList
            $miList.Checked  = $true
        }
    }
}

$miIcons.Add_Click({ Set-ImgView 'Icons' })
$miList.Add_Click({  Set-ImgView 'List'  })

function Update-WordCountLabel {
    param(
        [System.Windows.Forms.RichTextBox]$rtb,
        [System.Windows.Forms.Label]$lbl,
        [System.Windows.Forms.Button]$btn
    )
    $c = Get-TextCountFast $rtb
    $lbl.Text = "Wyrazów: $c"
    if ($global:TabsControl.SelectedIndex -eq 0) {
        $btn.Enabled = ($c -ge 2)
    }
}

function Update-ImgCountLabel {
    param(
        [System.Windows.Forms.Label]$lbl,
        [System.Windows.Forms.Button]$btn
    )
    $lbl.Text = "Obrazków: " + $global:ImagePaths.Count
    if ($global:TabsControl.SelectedIndex -eq 1) {
        $btn.Enabled = ($global:ImagePaths.Count -ge 2)
    }
}

$rtbWords.Add_TextChanged({
    Update-WordCountLabel -rtb $rtbWords -lbl $lblCountText -btn $btnSavePdf
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

Update-WordCountLabel -rtb $rtbWords -lbl $lblCountText -btn $btnSavePdf
Update-ImgCountLabel -lbl $lblCountImg -btn $btnSavePdf

$tabs.Add_SelectedIndexChanged({
    if ($global:TabsControl.SelectedIndex -eq 0) {
        Update-WordCountLabel -rtb $rtbWords -lbl $lblCountText -btn $btnSavePdf
    } else {
        Update-ImgCountLabel -lbl $lblCountImg -btn $btnSavePdf
    }
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

# formatowanie B/I/U

function Toggle-SelectionStyle {
    param(
        [System.Windows.Forms.RichTextBox]$rtb,
        [System.Drawing.FontStyle]$flag
    )

    $selStart = $rtb.SelectionStart
    $selLen   = $rtb.SelectionLength

    if ($selLen -eq 0) {
        $base = $rtb.SelectionFont
        if (-not $base) { $base = $rtb.Font }
        $newStyle = $base.Style -bxor $flag
        $rtb.SelectionFont = New-Object System.Drawing.Font($base, $newStyle)
    } else {
        $rtb.Select($selStart,1)
        $base = $rtb.SelectionFont
        if (-not $base) { $base = $rtb.Font }
        $targetStyle = $base.Style -bxor $flag
        $rtb.Select($selStart,$selLen)
        $rtb.SelectionFont = New-Object System.Drawing.Font($base, $targetStyle)
        $rtb.Select($selStart,$selLen)
    }
}



$btnBoldItem.Add_Click({      Toggle-SelectionStyle -rtb $rtbWords -flag ([System.Drawing.FontStyle]::Bold) })
$btnItalicItem.Add_Click({    Toggle-SelectionStyle -rtb $rtbWords -flag ([System.Drawing.FontStyle]::Italic) })
$btnUnderlineItem.Add_Click({ Toggle-SelectionStyle -rtb $rtbWords -flag ([System.Drawing.FontStyle]::Underline) })

$rtbWords.Add_SelectionChanged({
    $f = $rtbWords.SelectionFont
    if (-not $f) { $f = $rtbWords.Font }
    $btnBoldItem.Checked      = [bool]($f.Style -band [System.Drawing.FontStyle]::Bold)
    $btnItalicItem.Checked    = [bool]($f.Style -band [System.Drawing.FontStyle]::Italic)
    $btnUnderlineItem.Checked = [bool]($f.Style -band [System.Drawing.FontStyle]::Underline)
})

# Obrazki: dodawanie / usuwanie

$btnAddImg.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = 'Obrazy|*.png;*.jpg;*.jpeg;*.jfif;*.bmp;*.gif;*.tif;*.tiff;*.ico'
    $ofd.Multiselect = $true
    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    foreach ($path in $ofd.FileNames) {
        try {
            $img = [System.Drawing.Image]::FromFile($path)
            $key = [Guid]::NewGuid().ToString()

            # duża miniatura (ikony)
            $thumbLarge = New-Object System.Drawing.Bitmap($imgList.ImageSize.Width, $imgList.ImageSize.Height)
            $gL = [System.Drawing.Graphics]::FromImage($thumbLarge)
            $gL.Clear([System.Drawing.Color]::White)
            $gL.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $scaleL = [Math]::Min($imgList.ImageSize.Width / $img.Width, $imgList.ImageSize.Height / $img.Height)
            $dwL = [int]([Math]::Round($img.Width * $scaleL))
            $dhL = [int]([Math]::Round($img.Height * $scaleL))
            $dxL = [int](($imgList.ImageSize.Width - $dwL)/2)
            $dyL = [int](($imgList.ImageSize.Height - $dhL)/2)
            $gL.DrawImage($img,$dxL,$dyL,$dwL,$dhL)
            $gL.Dispose()
            $imgList.Images.Add($key, $thumbLarge) | Out-Null
            $thumbLarge.Dispose()

            # mała miniatura (widok "Lista")
            $thumbSmall = New-Object System.Drawing.Bitmap($imgListSmall.ImageSize.Width, $imgListSmall.ImageSize.Height)
            $gS = [System.Drawing.Graphics]::FromImage($thumbSmall)
            $gS.Clear([System.Drawing.Color]::White)
            $gS.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $scaleS = [Math]::Min($imgListSmall.ImageSize.Width / $img.Width, $imgListSmall.ImageSize.Height / $img.Height)
            $dwS = [int]([Math]::Round($img.Width * $scaleS))
            $dhS = [int]([Math]::Round($img.Height * $scaleS))
            $dxS = [int](($imgListSmall.ImageSize.Width - $dwS)/2)
            $dyS = [int](($imgListSmall.ImageSize.Height - $dhS)/2)
            $gS.DrawImage($img,$dxS,$dyS,$dwS,$dhS)
            $gS.Dispose()
            $imgListSmall.Images.Add($key, $thumbSmall) | Out-Null
            $thumbSmall.Dispose()

            $img.Dispose()

            $iconIndex = $imgList.Images.IndexOfKey($key)
            $item = New-Object System.Windows.Forms.ListViewItem(
                [System.IO.Path]::GetFileName($path),
                $iconIndex
            )
            $listView.Items.Add($item) | Out-Null
            $global:ImagePaths.Add($path) | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Nie udało się wczytać: $path") | Out-Null
        }
    }

    Update-ImgCountLabel -lbl $lblCountImg -btn $btnSavePdf
	Update-ImageOrderLabels
    Set-ProjectDirty
    Schedule-PreviewRefresh
})

$btnRemoveImg.Add_Click({
    if ($listView.SelectedIndices.Count -le 0) { return }
    for ($k = $listView.SelectedIndices.Count - 1; $k -ge 0; $k--) {
        $idx = $listView.SelectedIndices[$k]
        $listView.Items.RemoveAt($idx)
        $global:ImagePaths.RemoveAt($idx)
    }
    Update-ImgCountLabel -lbl $lblCountImg -btn $btnSavePdf
	Update-ImageOrderLabels
	$listView.ArrangeIcons()
    Set-ProjectDirty
	Schedule-PreviewRefresh
})

$btnClearImg.Add_Click({
    $listView.Items.Clear()
    $imgList.Images.Clear()
    $imgListSmall.Images.Clear()
    $global:ImagePaths.Clear()
    Update-ImgCountLabel -lbl $lblCountImg -btn $btnSavePdf
    Set-ProjectDirty
    Schedule-PreviewRefresh
})

# Motywy – zastosowanie + odświeżenie panelu

$cmbStyle.Add_SelectedIndexChanged({

    if ($global:StyleChangingInternally) { return }

    $idx = $cmbStyle.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $global:CardStyles.Count) { return }

    $style = $global:CardStyles[$idx]

    # 1) Jeśli to NIE jest styl użytkownika – wczytaj preset
    if ($style.Id -ne 'user') {

        $global:TextColorHex        = $style.Settings.TextColorHex
        $global:TextShadowEnabled   = $style.Settings.TextShadowEnabled

        $global:CardBackgroundMode          = $style.Settings.CardBackgroundMode
        $global:CardBackgroundColorHex      = $style.Settings.CardBackgroundColorHex
        $global:CardBackgroundGradColor1Hex = $style.Settings.CardBackgroundGradColor1Hex
        $global:CardBackgroundGradColor2Hex = $style.Settings.CardBackgroundGradColor2Hex
        $global:CardBackgroundGradientType  = $style.Settings.CardBackgroundGradientType
        $global:CardBackgroundGradientDirection = $style.Settings.CardBackgroundGradientDirection

        $global:CardBorderEnabled   = $style.Settings.CardBorderEnabled
        $global:CardBorderColorHex  = $style.Settings.CardBorderColorHex
        $global:CardBorderWidthPx   = $style.Settings.CardBorderWidthPx
        $global:CardBorderStyle     = $style.Settings.CardBorderStyle

        $global:CardLineEnabled     = $style.Settings.CardLineEnabled
        $global:CardLineColorHex    = $style.Settings.CardLineColorHex
        $global:CardLineWidthPx     = $style.Settings.CardLineWidthPx
        $global:CardLineStyle       = $style.Settings.CardLineStyle

        $global:CardShadowEnabled   = $style.Settings.CardShadowEnabled
        $global:CardRoundedEnabled  = $style.Settings.CardRoundedEnabled
        $global:ShowCutLines        = $style.Settings.ShowCutLines
    }

    # 2) W każdym przypadku – zsynchronizuj UI z aktualnymi globalnymi

    # TEKST
    if ($global:BtnTextColor) {
        $global:BtnTextColor.BackColor =
            [System.Drawing.ColorTranslator]::FromHtml($global:TextColorHex)
    }
    if ($global:ChkTextShadow) {
        $global:ChkTextShadow.Checked = $global:TextShadowEnabled
    }

    # TŁO
    if ($chkBgEnabled) {
        $chkBgEnabled.Checked = ($global:CardBackgroundMode -ne "none")

        switch ($global:CardBackgroundMode) {
            'gradient' { $rbBgGradient.Checked = $true }
            'color'    { $rbBgColor.Checked    = $true }
            default    { $rbBgColor.Checked    = $true }
        }

        $btnBgColor.BackColor  =
            [System.Drawing.ColorTranslator]::FromHtml($global:CardBackgroundColorHex)
        $btnBgGrad1.BackColor  =
            [System.Drawing.ColorTranslator]::FromHtml($global:CardBackgroundGradColor1Hex)
        $btnBgGrad2.BackColor  =
            [System.Drawing.ColorTranslator]::FromHtml($global:CardBackgroundGradColor2Hex)

        if ($global:CardBackgroundGradientType -eq 'radial') {
            $cmbBgType.SelectedItem = "Radialny"
        } else {
            switch ($global:CardBackgroundGradientDirection) {
                'vertical'   { $cmbBgType.SelectedItem = "Pionowy" }
                'horizontal' { $cmbBgType.SelectedItem = "Poziomy" }
                default      { $cmbBgType.SelectedItem = "Ukośny 45°" }
            }
        }

        Update-BgControlsState
    }

    # OBRAMOWANIE
    if ($chkBorderEnabled) {
        $chkBorderEnabled.Checked = $global:CardBorderEnabled
        $btnBorderColor.BackColor =
            [System.Drawing.ColorTranslator]::FromHtml($global:CardBorderColorHex)
        $numBorderWidth.Value = [decimal]$global:CardBorderWidthPx

        switch ($global:CardBorderStyle) {
            'dashed' { $cmbBorderStyle.SelectedItem = "kreskowana" }
            'dotted' { $cmbBorderStyle.SelectedItem = "kropkowana" }
            default  { $cmbBorderStyle.SelectedItem = "ciągła" }
        }

        Update-BorderControlsEnabled
    }

    # LINIA ŚRODKOWA
    if ($chkLineEnabled) {
        $chkLineEnabled.Checked = $global:CardLineEnabled
        $btnLineColor.BackColor =
            [System.Drawing.ColorTranslator]::FromHtml($global:CardLineColorHex)
        $numLineWidth.Value = [decimal]$global:CardLineWidthPx

        switch ($global:CardLineStyle) {
            'dashed' { $cmbLineStyle.SelectedItem = "kreskowana" }
            'dotted' { $cmbLineStyle.SelectedItem = "kropkowana" }
            default  { $cmbLineStyle.SelectedItem = "ciągła" }
        }

        Update-LineControlsEnabled
    }

    # DODATKI
    if ($chkCardShadow) {
        $chkCardShadow.Checked = $global:CardShadowEnabled
    }
    if ($chkCardRounded) {
        $chkCardRounded.Checked = $global:CardRoundedEnabled
    }
    if ($chkCutlines) {
        $chkCutlines.Checked = $global:ShowCutLines
    }

    Set-ProjectDirty
	Schedule-PreviewRefresh
})

# Zapis PDF

$btnSavePdf.Add_Click({
    if ($global:TabsControl.SelectedIndex -eq 0) {
        $items = Get-TextItems -rtb $rtbWords
        if ($items.Count -lt 2) {
            [System.Windows.Forms.MessageBox]::Show("Podaj przynajmniej 2 wyrazy.","Zapis") | Out-Null
            return
        }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "PDF|*.pdf"
        $sfd.FileName = "karty_txt.pdf"
        if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $outPath = Get-UniquePath $sfd.FileName

        $fontSize = Compute-OptimalFontPt $items
        $body     = Build-TextFullBody -Items $items
        $css      = Build-Css -FontPt $fontSize -ShowCutLines:$global:ShowCutLines
        $fullHtml = Build-FullHtml -Css $css -Body $body

        Convert-HtmlToPdf -FullHtml $fullHtml -PdfPath $outPath
        [System.Windows.Forms.MessageBox]::Show("Zapisano PDF:`n$outPath","Gotowe") | Out-Null
    }
    else {
        if ($global:ImagePaths.Count -lt 2) {
            [System.Windows.Forms.MessageBox]::Show("Dodaj przynajmniej 2 obrazki.","Zapis") | Out-Null
            return
        }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "PDF|*.pdf"
        $sfd.FileName = "karty_img.pdf"
        if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $outPath = Get-UniquePath $sfd.FileName

        $fontSize = 40
        $body     = Build-ImageFullBody -ImagePaths $global:ImagePaths.ToArray()
        $css      = Build-Css -FontPt $fontSize -ShowCutLines:$global:ShowCutLines
        $fullHtml = Build-FullHtml -Css $css -Body $body

        Convert-HtmlToPdf -FullHtml $fullHtml -PdfPath $outPath
        [System.Windows.Forms.MessageBox]::Show("Zapisano PDF:`n$outPath","Gotowe") | Out-Null
    }
})

$btnSaveProject.Add_Click({ Save-Project })
$btnLoadProject.Add_Click({ Load-Project })


# Ustaw początkowy podział splittera
$form.Add_Shown({
    $w = $splitMain.Width
    if ($w -le 0) { return }

    # lewa / prawa
    $left = [int]($w * 0.5)
    if ($left -lt $splitMain.Panel1MinSize) {
        $left = $splitMain.Panel1MinSize
    }
    $splitMain.SplitterDistance = $left

    # góra (podgląd) / dół (ustawienia)
    if ($splitRight -ne $null -and $splitRight.Height -gt 0) {
        $splitRight.SplitterDistance = [int]($splitRight.Height * 0.45)  # ~45% na podgląd
    }

    Clear-ProjectDirty
    Schedule-PreviewRefresh
})

$form.Add_FormClosing({
    param($sender, $e)

    if ($global:ProjectDirty) {
        $res = [System.Windows.Forms.MessageBox]::Show(
            "Masz niezapisane zmiany. Zamknąć mimo to?",
            "Niezapisany projekt",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($res -ne [System.Windows.Forms.DialogResult]::Yes) {
            $e.Cancel = $true
        }
    }
})

# Start
[void]$form.ShowDialog()
