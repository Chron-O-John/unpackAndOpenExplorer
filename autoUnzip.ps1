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
Expand-Archive -Path $zipFilePath -DestinationPath $destinationPath

# Delete the zip file
Remove-Item -Path $zipFilePath -Force

# Open an Explorer window in foreground of the extracted folder
$explorer = Start-Process -FilePath "explorer.exe" -ArgumentList "/root\,$destinationPath" -PassThru
Start-Sleep -Milliseconds 500
$foldername = Split-Path $destinationPath -Leaf
#echo $explorer
    (New-Object -ComObject 'Shell.Application').Windows() | ForEach-Object { 
        if ($_.locationName -contains "$foldername") {
             $pwnd = $_.HWND
        }
    }


#$hwnd = (Get-Process -Id $explorer.Id).MainWindowHandle
[Program]::SetForegroundWindow($pwnd)