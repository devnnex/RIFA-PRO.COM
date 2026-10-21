# Draw the code-native ticket mark at each required PWA size.
Add-Type -AssemblyName System.Drawing
$projectRoot = Split-Path -Parent $PSScriptRoot
$outputs = @(@{Name='icon-192.png';Size=192},@{Name='icon-512.png';Size=512},@{Name='icon-maskable-512.png';Size=512},@{Name='apple-touch-icon.png';Size=180})
foreach ($spec in $outputs) {
  $size = $spec.Size
  $bitmap = New-Object System.Drawing.Bitmap($size,$size)
  $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
  $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
  $graphics.ScaleTransform(($size/512.0),($size/512.0))
  $navy = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#17283d'))
  $cream = New-Object System.Drawing.SolidBrush([System.Drawing.ColorTranslator]::FromHtml('#f5f2ea'))
  $gold = New-Object System.Drawing.Pen([System.Drawing.ColorTranslator]::FromHtml('#a88a4d'),5)
  $gold.DashStyle = [System.Drawing.Drawing2D.DashStyle]::Dash
  $graphics.Clear($navy.Color)
  $graphics.FillRectangle($cream,105,145,302,222)
  $graphics.FillEllipse($navy,72,223,66,66)
  $graphics.FillEllipse($navy,374,223,66,66)
  $graphics.DrawLine($gold,338,160,338,352)
  $font = New-Object System.Drawing.Font('Georgia',140,[System.Drawing.FontStyle]::Bold,[System.Drawing.GraphicsUnit]::Pixel)
  $graphics.DrawString('R',$font,$navy,[System.Drawing.PointF]::new(157,174))
  $starFont = New-Object System.Drawing.Font('Georgia',35,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Pixel)
  $starBrush = New-Object System.Drawing.SolidBrush($gold.Color)
  $graphics.DrawString([string][char]0x2605,$starFont,$starBrush,[System.Drawing.PointF]::new(349,231))
  $bitmap.Save((Join-Path $projectRoot ('images/'+$spec.Name)),[System.Drawing.Imaging.ImageFormat]::Png)
  $starBrush.Dispose(); $starFont.Dispose(); $font.Dispose(); $gold.Dispose(); $navy.Dispose(); $cream.Dispose(); $graphics.Dispose(); $bitmap.Dispose()
}
Write-Output 'PWA icons generated: 192, 512, maskable 512 and Apple 180.'
