' ============================================================
'  baslat.vbs  -  Tek tikla baslatici (VBScript orkestrasyonu)
'
'  1) PowerShell HTTP koprusunu (server.ps1) ayri pencerede baslatir
'  2) Sunucunun ayaga kalkmasini bekler
'  3) Modern tarayiciyi "uygulama modu" ile acar (kenarliksiz, app gibi)
'
'  Boylece: VBScript baslatir -> PowerShell koprer -> Chromium gosterir.
' ============================================================
Option Explicit
Dim shell, fso, here, ps1, url, port
Set shell = CreateObject("WScript.Shell")
Set fso   = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
ps1  = fso.BuildPath(here, "server.ps1")
port = 8765
url  = "http://localhost:" & port & "/"

' 1) PowerShell sunucusu (gorunur pencere: loglari ve hatalari gormek icin)
shell.CurrentDirectory = here
shell.Run "powershell -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """", 1, False

' 2) Kisa bekleme
WScript.Sleep 1200

' 3) Modern tarayiciyi uygulama modunda ac (once Edge, sonra Chrome, olmazsa varsayilan)
If Not OpenApp("msedge.exe", url) Then
    If Not OpenApp("chrome.exe", url) Then
        shell.Run url, 1, False   ' varsayilan tarayicida normal sekme
    End If
End If

Function OpenApp(exeName, address)
    On Error Resume Next
    Dim exePath
    exePath = shell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & exeName & "\")
    If Err.Number <> 0 Or Len(exePath) = 0 Then
        OpenApp = False
        Exit Function
    End If
    shell.Run """" & exePath & """ --app=" & address & " --new-window", 1, False
    OpenApp = (Err.Number = 0)
End Function
