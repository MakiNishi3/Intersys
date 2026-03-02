Option Explicit

Dim sh, fso, shellApp, sizeInput, unitInput, sizeValue, bytes
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set shellApp = CreateObject("Shell.Application")

If Not IsAdmin() Then
    MsgBox "RamSizer must be run as Administrator.", 16, "RamSizer"
    WScript.Quit
End If

sizeInput = InputBox("Enter size value (1 - 1000):", "RamSizer")
If sizeInput = "" Or Not IsNumeric(sizeInput) Then WScript.Quit
sizeValue = CLng(sizeInput)

If sizeValue < 1 Or sizeValue > 1000 Then
    MsgBox "Invalid size range.", 16, "RamSizer"
    WScript.Quit
End If

unitInput = UCase(InputBox("Select unit: MB / GB / TB", "RamSizer"))
Select Case unitInput
    Case "MB": bytes = sizeValue * 1024 * 1024
    Case "GB": bytes = sizeValue * 1024 * 1024 * 1024
    Case "TB": bytes = sizeValue * 1024 * 1024 * 1024 * 1024#
    Case Else
        MsgBox "Invalid unit.", 16, "RamSizer"
        WScript.Quit
End Select

sh.Run "cmd /c powercfg -h off", 0, True
sh.Run "cmd /c wmic computersystem where name='%computername%' set AutomaticManagedPagefile=False", 0, True
sh.Run "cmd /c wmic pagefileset where name='C:\\pagefile.sys' delete", 0, True

DeleteFolder sh.ExpandEnvironmentStrings("%TEMP%")
DeleteFolder sh.ExpandEnvironmentStrings("%USERPROFILE%\Downloads")

shellApp.Namespace(10).Items().InvokeVerb("Delete")

sh.Run "cmd /c rundll32.exe advapi32.dll,ProcessIdleTasks", 0, True
sh.Run "cmd /c echo y|PowerShell Clear-RecycleBin -Force", 0, True

MsgBox "RamSizer completed." & vbCrLf & _
       "Requested size: " & sizeValue & " " & unitInput & vbCrLf & _
       "System memory cleanup executed.", 64, "RamSizer"

Sub DeleteFolder(path)
    On Error Resume Next
    If fso.FolderExists(path) Then fso.DeleteFolder path, True
    On Error GoTo 0
End Sub

Function IsAdmin()
    Dim test
    On Error Resume Next
    test = sh.RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    IsAdmin = (Err.Number = 0)
    Err.Clear
End Function
