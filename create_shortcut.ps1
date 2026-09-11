$ws = New-Object -ComObject WScript.Shell
$desktop = [System.Environment]::GetFolderPath('Desktop')
$sc = $ws.CreateShortcut("$desktop\FormsMobile App.lnk")
$sc.TargetPath = "d:\appkhaosat\CHAY_APP.bat"
$sc.WorkingDirectory = "d:\appkhaosat"
$sc.IconLocation = "shell32.dll,14"
$sc.Save()
Write-Host "Done shortcut creation"
