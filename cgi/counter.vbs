' Simple CGI counter for IIS 5.1
Option Explicit

Const ForReading = 1
Const ForWriting = 2
Const COUNTER_RELATIVE_PATH = "..\data\visit-count.txt"

Dim fso, scriptFolder, counterPath, currentCount
Set fso = CreateObject("Scripting.FileSystemObject")
scriptFolder = fso.GetParentFolderName(WScript.ScriptFullName)
counterPath = fso.BuildPath(scriptFolder, COUNTER_RELATIVE_PATH)

currentCount = ReadCount(counterPath)
currentCount = currentCount + 1
Call WriteCount(counterPath, currentCount)

EmitResponse currentCount

Sub EmitResponse(count)
    Dim payload
    payload = "{""count"":" & CStr(count) & "}"
    WScript.StdOut.Write "Content-Type: application/json" & vbCrLf
    WScript.StdOut.Write "Cache-Control: no-cache" & vbCrLf & vbCrLf
    WScript.StdOut.Write payload
End Sub

Function ReadCount(path)
    On Error Resume Next
    Dim value, stream, tmp
    value = 0
    If fso.FileExists(path) Then
        Set stream = fso.OpenTextFile(path, ForReading, False)
        If Not stream.AtEndOfStream Then
            tmp = Trim(stream.ReadLine)
            If IsNumeric(tmp) Then
                value = CLng(tmp)
            End If
        End If
        stream.Close
    End If
    If Err.Number <> 0 Then
        value = 0
        Err.Clear
    End If
    ReadCount = value
End Function

Sub WriteCount(path, count)
    On Error Resume Next
    Dim stream
    ' Overwrite with the latest number so restarts stay consistent.
    Set stream = fso.OpenTextFile(path, ForWriting, True)
    stream.Write count & vbCrLf
    stream.Close
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub
