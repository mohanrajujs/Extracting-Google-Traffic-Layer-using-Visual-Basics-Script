Option Explicit

'Private Declare Sub google_traffic Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
ByVal szURL As String, ByVal szFileName As String, _
ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Dim Ret As Long
Dim TimeToRun

Sub google_traffic()

Dim ws As Worksheet
Dim LastRow As Long, i As Long
Dim strPath As String
Dim destinationDir As String
Dim strDirectory As String
Dim objFolder As Variant
Dim FolderName As Variant
Dim fso As Variant
    
destinationDir = "E:\Data_Google_Traffic\"
Set fso = CreateObject("Scripting.FileSystemObject")
'Const OverwriteExisting = True
strDirectory = destinationDir & Replace(Date, "/", "_ ") & " " & Format(Time, "hh_mm_ss")
Set objFolder = fso.CreateFolder(strDirectory)

'Not needed, timer takes care of this issue
'If Not fso.FolderExists(strDirectory) Then
'   Set objFolder = fso.CreateFolder(strDirectory)
'End If
'Set FolderName = fso

Set ws = Sheets("Google_Trafffic")
LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
For i = 2 To LastRow
    strPath = objFolder & "\" & ws.Range("A" & i).Value & ".jpg"
    Ret = URLDownloadToFile(0, ws.Range("B" & i).Value, strPath, 0, 0)
    If Ret = 0 Then
       ws.Range("C" & i).Value = "File successfully downloaded"
    Else
       ws.Range("C" & i).Value = "Unable to download the file"
    End If
Application.Wait Now + TimeValue("00:00:10"):
Next i
Call Main
End Sub
 
Sub Main()

TimeToRun = Now + TimeValue("00:01:00")

    Application.OnTime TimeToRun, "google_Traffic"

End Sub


Sub auto_open()

    Call Main

End Sub

Sub auto_close()

    On Error Resume Next
    Application.OnTime TimeToRun, "google_Traffic", , False
    
End Sub
