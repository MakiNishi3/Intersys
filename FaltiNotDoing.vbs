Option Explicit

Dim shell, choice
Set shell = CreateObject("WScript.Shell")

choice = MsgBox( _
"FaltiNotDoing" & vbCrLf & vbCrLf & _
"Choose simulation mode:" & vbCrLf & vbCrLf & _
"YES  = Crash simulation (BSOD style)" & vbCrLf & _
"NO   = Hang simulation (HFTD style)" & vbCrLf & _
"CANCEL = Exit", _
vbYesNoCancel + vbQuestion, _
"FaltiNotDoing")

If choice = vbYes Then
    SimulatedCrash
ElseIf choice = vbNo Then
    SimulatedHang
Else
    WScript.Quit
End If

Sub SimulatedCrash()
    Dim html
    html = "<html>" & _
    "<head>" & _
    "<title>FaltiNotDoing</title>" & _
    "<hta:application border='none' caption='no' showintaskbar='no' />" & _
    "<style>" & _
    "html,body{margin:0;width:100%;height:100%;background:rgb(0,0,0);" & _
    "color:#00aaff;font-family:Consolas;font-size:20px;}" & _
    ".center{padding:60px;}" & _
    "</style>" & _
    "</head>" & _
    "<body>" & _
    "<div class='center'>" & _
    "<h1>:(</h1>" & _
    "<p>Your PC ran into a problem and needs to restart.</p>" & _
    "<p><b>FaltiNotDoing – Crash Example</b></p>" & _
    "<p>If you crash on my computer you will lose:</p>" & _
    "<ul>" & _
    "<li>BSOD</li>" & _
    "<li>Colors Screen of Death</li>" & _
    "</ul>" & _
    "<p>Crash on your computer.</p>" & _
    "<p>Pick a color: RGB(0,0,0)</p>" & _
    "<br>" & _
    "<p>Stop Code: FALTI_NOT_DOING</p>" & _
    "<p>100% complete</p>" & _
    "<br>" & _
    "<p style='font-size:14px;'>This is a simulation. No system crash occurred.</p>" & _
    "</div>" & _
    "</body></html>"

    ShowHTA html
End Sub

Sub SimulatedHang()
    Dim html
    html = "<html>" & _
    "<head>" & _
    "<title>FaltiNotDoing</title>" & _
    "<hta:application border='none' caption='no' showintaskbar='no' />" & _
    "<style>" & _
    "html,body{margin:0;width:100%;height:100%;background:rgb(0,0,0);" & _
    "color:#ff0000;font-family:Consolas;font-size:20px;}" & _
    ".center{padding:60px;}" & _
    ".blink{animation:blink 1s infinite;}" & _
    "@keyframes blink{0%{opacity:1;}50%{opacity:0;}100%{opacity:1;}}" & _
    "</style>" & _
    "</head>" & _
    "<body>" & _
    "<div class='center'>" & _
    "<h1 class='blink'>SYSTEM NOT RESPONDING</h1>" & _
    "<p><b>FaltiNotDoing – Hang Example</b></p>" & _
    "<p>If you hang on my computer you will lose:</p>" & _
    "<ul>" & _
    "<li>HFTD</li>" & _
    "<li>Hang Fail Threat Display</li>" & _
    "</ul>" & _
    "<p>Hang on your computer.</p>" & _
    "<p>Pick a color: RGB(0,0,0)</p>" & _
    "<br>" & _
    "<p>Status: THREAD DEADLOCK DETECTED</p>" & _
    "<p>Input disabled...</p>" & _
    "<br>" & _
    "<p style='font-size:14px;'>Simulation only. Press Ctrl+Alt+Del to exit.</p>" & _
    "</div>" & _
    "</body></html>"

    ShowHTA html
End Sub

Sub ShowHTA(content)
    Dim fso, temp, file
    Set fso = CreateObject("Scripting.FileSystemObject")
    temp = fso.GetSpecialFolder(2)
    file = temp & "\falti_sim.hta"

    With fso.CreateTextFile(file, True)
        .Write content
        .Close
    End With

    shell.Run "mshta.exe """ & file & """", 1, False
End Sub
