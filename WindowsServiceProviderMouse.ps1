Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$rand = New-Object System.Random

while ($true) {
	$screenWidth = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width
	$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height

	$x = $rand.Next(0, $screenWidth)
	$y = $rand.Next(0, $screenHeight)

	[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)

	Start-Sleep -Milliseconds 3000

}
