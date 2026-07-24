Add-Type -AssemblyName System.Drawing
$srcPath = "C:\Users\Angel A Rodriguez\.gemini\antigravity-ide\brain\43c582c3-4305-4c55-88d9-c70d7d3020a2\meet_greet_section_1784906368746.png"
$src = [System.Drawing.Bitmap]::FromFile($srcPath)

# Exact photo coordinates (x: 149, y: 118, w: 420, h: 230)
$x = 149
$y = 118
$w = 420
$h = 230

$crop = New-Object System.Drawing.Bitmap $w, $h
$graphics = [System.Drawing.Graphics]::FromImage($crop)
$rectSrc = New-Object System.Drawing.Rectangle $x, $y, $w, $h
$rectDst = New-Object System.Drawing.Rectangle 0, 0, $w, $h

$graphics.DrawImage($src, $rectDst, $rectSrc, [System.Drawing.GraphicsUnit]::Pixel)

$crop.Save("c:\Users\Angel A Rodriguez\.gemini\antigravity-ide\scratch\questionnaire-new-generation-2026\public\meet_greet_photo.png", [System.Drawing.Imaging.ImageFormat]::Png)
$crop.Save("c:\Users\Angel A Rodriguez\.gemini\antigravity-ide\scratch\questionnaire-new-generation-2026\meet_greet_photo.png", [System.Drawing.Imaging.ImageFormat]::Png)

$graphics.Dispose()
$crop.Dispose()
$src.Dispose()
Write-Host "Successfully cropped exact Tradicion.org photo without headers!"
