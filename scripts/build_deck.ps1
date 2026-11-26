<#
.SYNOPSIS
    Builds a native PowerPoint (.pptx) presentation from a compact JSON specification
    using Microsoft PowerPoint COM automation, with optional 1920x1080 PNG export for visual QA.
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$SpecPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$ExportPngDir = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Office.Interop.PowerPoint")
[void][System.Reflection.Assembly]::LoadWithPartialName("office")

function Convert-HexToBgr {
    param([string]$Hex)
    if ([string]::IsNullOrWhiteSpace($Hex)) { return 0 }
    $clean = $Hex.Trim().TrimStart('#')
    if ($clean.Length -ne 6) {
        throw "Invalid hex color: '$Hex'. Expected #RRGGBB."
    }
    $r = [Convert]::ToInt32($clean.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($clean.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($clean.Substring(4, 2), 16)
    return [int]($r + ($g * 256) + ($b * 65536))
}

function Get-Prop {
    param($Obj, [string]$Name, $Default = $null)
    if ($null -eq $Obj) { return $Default }
    $prop = $Obj.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value) { return $Default }
    return $prop.Value
}

function Get-AlignInt {
    param([string]$Align)
    switch (($Align + "").ToLowerInvariant()) {
        "center"  { return 2 }
        "right"   { return 3 }
        "justify" { return 4 }
        default   { return 1 }
    }
}

function Get-AutoShapeInt {
    param([string]$ShapeType)
    switch (($ShapeType + "").ToLowerInvariant()) {
        "rounded_rect" { return 5 }
        "oval"         { return 9 }
        "triangle"     { return 7 }
        "right_arrow"  { return 33 }
        "chevron"      { return 52 }
        default        { return 1 }
    }
}

function Get-TransitionInt {
    param([string]$TransitionType)
    switch (($TransitionType + "").ToLowerInvariant()) {
        "morph"        { return 3954 }
        "fade"         { return 3849 }
        "fade_basic"   { return 1793 }
        "push_left"    { return 3853 }
        "push_right"   { return 3854 }
        "push_up"      { return 3855 }
        "push_down"    { return 3852 }
        "cover_left"   { return 1281 }
        "cover_down"   { return 1284 }
        "uncover_left" { return 2049 }
        "uncover_down" { return 2052 }
        "zoom"         { return 3345 }
        "zoom_center"  { return 3349 }
        default        { return 0 }
    }
}

function Get-AnimEffectInt {
    param([string]$EffectName)
    switch (($EffectName + "").ToLowerInvariant()) {
        "appear"            { return 1 }
        "fly_in"            { return 2 }
        "fade"              { return 10 }
        "wipe"              { return 22 }
        "zoom"              { return 23 }
        "grow_turn"         { return 31 }
        "faded_zoom"        { return 48 }
        "change_fill_color" { return 54 }
        "change_font_color" { return 56 }
        "grow_shrink"       { return 59 }
        "spin"              { return 61 }
        "pulse"             { return 75 }
        "flash_bulb"        { return 75 }
        "grow_with_color"   { return 77 }
        "teeter"            { return 80 }
        default             { return 10 }
    }
}

function Get-AnimTriggerInt {
    param([string]$TriggerName)
    switch (($TriggerName + "").ToLowerInvariant()) {
        "with_previous"  { return 2 }
        "after_previous" { return 3 }
        default          { return 1 }
    }
}

function Get-MergeCmdInt {
    param([string]$Mode)
    switch (($Mode + "").ToLowerInvariant()) {
        "union"     { return 1 }
        "combine"   { return 2 }
        "intersect" { return 3 }
        "subtract"  { return 4 }
        "fragment"  { return 5 }
        default     { throw "Unknown MergeShapes mode: '$Mode'" }
    }
}

function Add-CoverImage {
    param(
        $Slide,
        [string]$ImagePath,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$Name = ""
    )
    if (-not (Test-Path -LiteralPath $ImagePath)) {
        throw "Image file not found: '$ImagePath'"
    }
    $absImg = (Resolve-Path -LiteralPath $ImagePath).Path
    $pic = $Slide.Shapes.AddPicture($absImg, 0, -1, $Left, $Top, -1, -1)
    $origW = [double]$pic.Width
    $origH = [double]$pic.Height

    $scaleX = $Width / $origW
    $scaleY = $Height / $origH
    $scale = [Math]::Max($scaleX, $scaleY)

    $pic.LockAspectRatio = -1
    $pic.Width = [single]($origW * $scale)
    $newW = [double]$pic.Width
    $newH = [double]$pic.Height

    if ($newW -gt $Width) {
        $excessPts = $newW - $Width
        $cropOrig = [single](($excessPts / 2.0) / $scale)
        $pic.PictureFormat.CropLeft = $cropOrig
        $pic.PictureFormat.CropRight = $cropOrig
    }
    if ($newH -gt $Height) {
        $excessPts = $newH - $Height
        $cropOrig = [single](($excessPts / 2.0) / $scale)
        $pic.PictureFormat.CropTop = $cropOrig
        $pic.PictureFormat.CropBottom = $cropOrig
    }

    $pic.Left = [single]$Left
    $pic.Top = [single]$Top
    if (-not [string]::IsNullOrWhiteSpace($Name)) {
        $pic.Name = $Name
    }
    return $pic
}

function Apply-TextFormatting {
    param($Shape, $El)
    $text = [string](Get-Prop $El "text" "")
    $font = [string](Get-Prop $El "font" "Segoe UI")
    $size = [single](Get-Prop $El "size" 18)
    $colorHex = [string](Get-Prop $El "color" "#111111")
    $bold = [bool](Get-Prop $El "bold" $false)
    $italic = [bool](Get-Prop $El "italic" $false)
    $align = [string](Get-Prop $El "align" "left")
    $lineSpacing = [single](Get-Prop $El "line_spacing" 1.2)
    $spaceAfter = [single](Get-Prop $El "space_after" 6)
    $valign = [string](Get-Prop $El "valign" "top")

    $tf = $Shape.TextFrame
    $tf.WordWrap = -1
    $tf.AutoSize = 0
    $tf.MarginLeft = [single](Get-Prop $El "margin_left" 4)
    $tf.MarginRight = [single](Get-Prop $El "margin_right" 4)
    $tf.MarginTop = [single](Get-Prop $El "margin_top" 4)
    $tf.MarginBottom = [single](Get-Prop $El "margin_bottom" 4)
    switch ($valign.ToLowerInvariant()) {
        "middle" { $tf.VerticalAnchor = 3 }
        "bottom" { $tf.VerticalAnchor = 4 }
        default  { $tf.VerticalAnchor = 1 }
    }

    $tr = $tf.TextRange
    $tr.Text = $text
    $tr.Font.Name = $font
    $tr.Font.Size = $size
    $tr.Font.Color.RGB = Convert-HexToBgr $colorHex
    $tr.Font.Bold = if ($bold) { -1 } else { 0 }
    $tr.Font.Italic = if ($italic) { -1 } else { 0 }

    $alignInt = Get-AlignInt $align
    for ($p = 1; $p -le $tr.Paragraphs().Count; $p++) {
        $para = $tr.Paragraphs($p)
        $para.ParagraphFormat.Alignment = $alignInt
        $para.ParagraphFormat.SpaceWithin = $lineSpacing
        $para.ParagraphFormat.SpaceAfter = $spaceAfter
    }
}

$resolvedSpec = (Resolve-Path -LiteralPath $SpecPath).Path
$spec = Get-Content -LiteralPath $resolvedSpec -Raw -Encoding UTF8 | ConvertFrom-Json

$outDir = Split-Path -Parent $OutputPath
if (-not [string]::IsNullOrWhiteSpace($outDir) -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}
$absOutput = [System.IO.Path]::GetFullPath($OutputPath)

$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Add(0) # msoFalse (headless)
    $pres.PageSetup.SlideWidth = 960  # 13.333 in (modern 16:9 widescreen)
    $pres.PageSetup.SlideHeight = 540 # 7.5 in

    $slides = @(Get-Prop $spec "slides" @())
    for ($sIdx = 0; $sIdx -lt $slides.Count; $sIdx++) {
        $sSpec = $slides[$sIdx]
        $slide = $pres.Slides.Add($sIdx + 1, 12) # ppLayoutBlank
        $shapeMap = @{}

        # 1. Slide Background (Solid color + optional full-bleed image + optional Sanzo Wada overlay)
        $bgSpec = Get-Prop $sSpec "background" $null
        $bgColor = Get-Prop $bgSpec "color" "#FFFFFF"
        $slide.FollowMasterBackground = 0
        $slide.Background.Fill.Visible = -1
        $slide.Background.Fill.Solid()
        $slide.Background.Fill.ForeColor.RGB = Convert-HexToBgr $bgColor

        $bgImage = Get-Prop $bgSpec "image" ""
        if (-not [string]::IsNullOrWhiteSpace($bgImage)) {
            $bgPic = Add-CoverImage -Slide $slide -ImagePath $bgImage -Left 0 -Top 0 -Width 960 -Height 540 -Name "_bg_image"
            $shapeMap["_bg_image"] = $bgPic
        }

        $overlayHex = Get-Prop $bgSpec "overlay_color" ""
        if (-not [string]::IsNullOrWhiteSpace($overlayHex)) {
            $ov = $slide.Shapes.AddShape(1, 0, 0, 960, 540)
            $ov.Name = "_bg_overlay"
            $ov.Line.Visible = 0
            $ov.Fill.Visible = -1
            $ov.Fill.Solid()
            $ov.Fill.ForeColor.RGB = Convert-HexToBgr $overlayHex
            $ov.Fill.Transparency = [single](Get-Prop $bgSpec "overlay_transparency" 0.35)
            $shapeMap["_bg_overlay"] = $ov
        }

        # 2. Slide Elements
        $elements = @(Get-Prop $sSpec "elements" @())
        foreach ($el in $elements) {
            $kind = ([string](Get-Prop $el "type" "textbox")).ToLowerInvariant()
            $name = [string](Get-Prop $el "name" "")
            $left = [single](Get-Prop $el "left" 0)
            $top = [single](Get-Prop $el "top" 0)
            $width = [single](Get-Prop $el "width" 100)
            $height = [single](Get-Prop $el "height" 50)
            $created = $null

            if ($kind -eq "textbox") {
                $created = $slide.Shapes.AddTextbox(1, $left, $top, $width, $height)
                $bgEl = Get-Prop $el "bg_color" ""
                if (-not [string]::IsNullOrWhiteSpace($bgEl)) {
                    $created.Fill.Visible = -1
                    $created.Fill.Solid()
                    $created.Fill.ForeColor.RGB = Convert-HexToBgr $bgEl
                    $created.Fill.Transparency = [single](Get-Prop $el "bg_transparency" 0.0)
                } else {
                    $created.Fill.Visible = 0
                }
                $lineEl = Get-Prop $el "line_color" ""
                if (-not [string]::IsNullOrWhiteSpace($lineEl)) {
                    $created.Line.Visible = -1
                    $created.Line.ForeColor.RGB = Convert-HexToBgr $lineEl
                    $created.Line.Weight = [single](Get-Prop $el "line_weight" 1)
                } else {
                    $created.Line.Visible = 0
                }
                Apply-TextFormatting -Shape $created -El $el
            }
            elseif ($kind -eq "shape") {
                $shapeKind = [string](Get-Prop $el "shape_type" "rect")
                if ($shapeKind.ToLowerInvariant() -eq "line") {
                    $x2 = [single](Get-Prop $el "x2" ($left + $width))
                    $y2 = [single](Get-Prop $el "y2" $top)
                    $created = $slide.Shapes.AddLine($left, $top, $x2, $y2)
                    $created.Line.Visible = -1
                    $created.Line.ForeColor.RGB = Convert-HexToBgr ([string](Get-Prop $el "line_color" "#000000"))
                    $created.Line.Weight = [single](Get-Prop $el "line_weight" 1.5)
                } else {
                    $autoInt = Get-AutoShapeInt $shapeKind
                    $created = $slide.Shapes.AddShape($autoInt, $left, $top, $width, $height)
                    $fillEl = Get-Prop $el "fill_color" ""
                    if (-not [string]::IsNullOrWhiteSpace($fillEl)) {
                        $created.Fill.Visible = -1
                        $created.Fill.Solid()
                        $created.Fill.ForeColor.RGB = Convert-HexToBgr $fillEl
                        $created.Fill.Transparency = [single](Get-Prop $el "fill_transparency" 0.0)
                    } else {
                        $created.Fill.Visible = 0
                    }
                    $lineEl = Get-Prop $el "line_color" ""
                    if (-not [string]::IsNullOrWhiteSpace($lineEl)) {
                        $created.Line.Visible = -1
                        $created.Line.ForeColor.RGB = Convert-HexToBgr $lineEl
                        $created.Line.Weight = [single](Get-Prop $el "line_weight" 1)
                    } else {
                        $created.Line.Visible = 0
                    }
                    $txt = Get-Prop $el "text" ""
                    if (-not [string]::IsNullOrEmpty($txt)) {
                        Apply-TextFormatting -Shape $created -El $el
                    }
                }
            }
            elseif ($kind -eq "image") {
                $imgPath = [string](Get-Prop $el "path" "")
                $fit = ([string](Get-Prop $el "fit" "cover")).ToLowerInvariant()
                if ($fit -eq "stretch") {
                    $absImg = (Resolve-Path -LiteralPath $imgPath).Path
                    $created = $slide.Shapes.AddPicture($absImg, 0, -1, $left, $top, $width, $height)
                } else {
                    $created = Add-CoverImage -Slide $slide -ImagePath $imgPath -Left $left -Top $top -Width $width -Height $height
                }
            }

            if ($null -ne $created) {
                $rot = Get-Prop $el "rotation" $null
                if ($null -ne $rot) {
                    $created.Rotation = [single]$rot
                }
                if (-not [string]::IsNullOrWhiteSpace($name)) {
                    $created.Name = $name
                    $shapeMap[$name] = $created
                }
            }
        }

        # 3. Native Merge Shapes
        $merges = @(Get-Prop $sSpec "merge_shapes" @())
        foreach ($m in $merges) {
            $baseName = [string](Get-Prop $m "base" "")
            $others = @(Get-Prop $m "others" @())
            $mode = [string](Get-Prop $m "mode" "union")
            $resultName = [string](Get-Prop $m "result_name" $baseName)

            $namesList = New-Object System.Collections.Generic.List[string]
            $namesList.Add($baseName)
            foreach ($o in $others) { $namesList.Add([string]$o) }

            $existingBefore = @{}
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $sh = $slide.Shapes.Item($i)
                if ($namesList -notcontains $sh.Name) {
                    $existingBefore[$sh.Id] = $true
                }
            }

            $cmdInt = Get-MergeCmdInt $mode
            $baseShape = $shapeMap[$baseName]
            $range = $slide.Shapes.Range([string[]]$namesList.ToArray())
            $range.MergeShapes([Microsoft.Office.Core.MsoMergeCmd]$cmdInt, $baseShape)

            # Find newly merged shape(s) and update map
            $mergedShape = $null
            for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                $sh = $slide.Shapes.Item($i)
                if (-not $existingBefore.ContainsKey($sh.Id)) {
                    $mergedShape = $sh
                    break
                }
            }
            if ($null -ne $mergedShape -and -not [string]::IsNullOrWhiteSpace($resultName)) {
                $mergedShape.Name = $resultName
                $shapeMap[$resultName] = $mergedShape
            }
        }

        # 4. Slide Transition (with safe fallback if Office build predates Morph)
        $trans = Get-Prop $sSpec "transition" $null
        if ($null -ne $trans) {
            $tType = [string](Get-Prop $trans "type" "none")
            $tDur = [single](Get-Prop $trans "duration" 0.8)
            $effInt = Get-TransitionInt $tType
            try {
                $slide.SlideShowTransition.EntryEffect = $effInt
                $slide.SlideShowTransition.Duration = $tDur
            } catch {
                Write-Warning "Slide $($sIdx + 1): transition '$tType' ($effInt) unsupported in Office $($ppt.Version); falling back to smooth fade (3849)."
                $slide.SlideShowTransition.EntryEffect = 3849
                $slide.SlideShowTransition.Duration = $tDur
            }
        }

        # 5. Dynamic Animations
        $anims = @(Get-Prop $sSpec "animations" @())
        foreach ($anim in $anims) {
            $targetName = [string](Get-Prop $anim "target" "")
            if ([string]::IsNullOrWhiteSpace($targetName)) { continue }

            $targetShape = $null
            if ($shapeMap.ContainsKey($targetName)) {
                $targetShape = $shapeMap[$targetName]
            } else {
                for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
                    if ($slide.Shapes.Item($i).Name -eq $targetName) {
                        $targetShape = $slide.Shapes.Item($i)
                        break
                    }
                }
            }
            if ($null -eq $targetShape) {
                Write-Warning "Slide $($sIdx + 1): animation target '$targetName' not found."
                continue
            }

            $effInt = Get-AnimEffectInt ([string](Get-Prop $anim "effect" "fade"))
            $trigInt = Get-AnimTriggerInt ([string](Get-Prop $anim "trigger" "on_click"))
            $dur = [single](Get-Prop $anim "duration" 0.6)
            $delay = [single](Get-Prop $anim "delay" 0.0)

            $seq = $slide.TimeLine.MainSequence
            $eff = $seq.AddEffect($targetShape, $effInt, 0, $trigInt)
            $eff.Timing.Duration = $dur
            $eff.Timing.TriggerDelayTime = $delay
        }
    }

    $pres.SaveAs($absOutput, 24) # ppSaveAsOpenXMLPresentation
    Write-Output "SAVED_PPTX=$absOutput"

    if (-not [string]::IsNullOrWhiteSpace($ExportPngDir)) {
        if (-not (Test-Path -LiteralPath $ExportPngDir)) {
            New-Item -ItemType Directory -Path $ExportPngDir -Force | Out-Null
        }
        $absPngDir = (Resolve-Path -LiteralPath $ExportPngDir).Path
        $pres.Export($absPngDir, "PNG", 1920, 1080)
        Write-Output "EXPORTED_PNG_DIR=$absPngDir"
    }
}
finally {
    if ($null -ne $pres) {
        try { $pres.Close() } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null
    }
    if ($null -ne $ppt) {
        try { $ppt.Quit() } catch {}
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
