Add-Type -AssemblyName System.Drawing

$width = 1024
$height = 1024
$bitmap = New-Object System.Drawing.Bitmap $width, $height
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

# Background (rounded rect approximation using clear + rects, or just leave it since the blue box is centered)
$graphics.Clear([System.Drawing.Color]::White)

# Blue Base Rectangle
$blueBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(40, 90, 185))
$rect = New-Object System.Drawing.Rectangle(200, 200, 624, 624)
$graphics.FillRectangle($blueBrush, $rect)

# Light Blue Screen
$lightBlueBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(217, 226, 255))
$screenRect = New-Object System.Drawing.Rectangle(296, 312, 432, 120)
$graphics.FillRectangle($lightBlueBrush, $screenRect)

# White Buttons
$whiteBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
$graphics.FillRectangle($whiteBrush, 296, 512, 96, 96)
$graphics.FillRectangle($whiteBrush, 464, 512, 96, 96)
$graphics.FillRectangle($whiteBrush, 632, 512, 96, 96)
$graphics.FillRectangle($whiteBrush, 296, 656, 96, 96)

# Equal Button
$equalBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(138, 178, 255))
$graphics.FillRectangle($equalBrush, 464, 656, 264, 96)

$path = "d:\HermesJang\CMCC\tools\Benefitcostanalysis\app-icon.png"
$bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)

$graphics.Dispose()
$bitmap.Dispose()
