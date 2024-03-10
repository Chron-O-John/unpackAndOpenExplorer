param (
    [string]$zipFilePath
)

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Program {
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@

# Check if the provided file exists
if (-not (Test-Path -Path $zipFilePath)) {
    Write-Host "File '$zipFilePath' not found."
    exit
}

#SetDestinationPath
$destinationPath = (Join-Path -Path (Split-Path -Path $zipFilePath) -ChildPath (Split-Path -Path $zipFilePath -Leaf).Replace('.zip', ''))

# Unzip the file
$arguments = "x `"$zipFilePath`" -o`"$destinationPath\`""

# Replaced with 7zip, as Expand-Archive is too slow
# Expand-Archive -Path $zipFilePath -DestinationPath $destinationPath
Start-Process -NoNewWindow -Wait -FilePath "C:\Program Files\7-Zip\7z.exe" -ArgumentList $arguments

pause
# Delete the zip file
Remove-Item -Path $zipFilePath -Force

# Open an Explorer window in foreground of the extracted folder
$explorer = Start-Process -FilePath "explorer.exe" -ArgumentList "/root\,$destinationPath" -PassThru
Start-Sleep -Milliseconds 500
$foldername = Split-Path $destinationPath -Leaf

    (New-Object -ComObject 'Shell.Application').Windows() | ForEach-Object { 
        if ($_.locationName -contains "$foldername") {
             $pwnd = $_.HWND
        }
    }
	
[Program]::SetForegroundWindow($pwnd)