Attribute VB_Name = "LogWork"
Option Explicit

Const gLogFileName = "MailLog.txt"
Const gInitFalse = 0
Const gInitTrue = 15
Const gLogFormStringsLimit = 200
Const gLogFormEnabled = False
Dim gNewLogString, gLogFileFolder, gInitResult, gLogFormStrings

'TimeStamp for LOG
Private Function uGetTimeStamp()
    uGetTimeStamp = "[" & Format(Now(), "MM.DD hh:mm:ss") & "] "
End Function

'Simple LOGGER
Public Function uDebugPrint(inText)
    Debug.Print uGetTimeStamp & inText
End Function

'Advanced LOGGER Class 1
Public Function uADebugPrint(inTag, inText)
    Debug.Print uGetTimeStamp & inTag & ": " & inText
End Function

'Advanced LOGGER Class 1 with StringReturn
Public Function uBDebugPrint(inTag, inTextVariable, inText)
    Debug.Print uGetTimeStamp & inTag & ": " & inText
    inTextVariable = inText
End Function

'Advanced LOGGER Class 2
Public Function uCDebugPrint(inTag, inType, inText, Optional inSilentMode = False)
Dim tType, tLogString
    If inSilentMode Then: Exit Function
    Select Case inType
        Case 1: tType = "WARN"
        Case 2: tType = "CRIT"
        Case Else: tType = "INFO"
    End Select
    tLogString = uGetTimeStamp & tType & "." & inTag & ": " & inText
    Debug.Print tLogString
End Function

'Advanced LOGGER Class 2 with StringReturn
Public Function uC2DebugPrint(inTag, inType, inTextVariable, inText, Optional inSilentMode = False)
Dim tType, tLogString
    If inSilentMode Then: Exit Function
    Select Case inType
        Case 1: tType = "WARN"
        Case 2: tType = "CRIT"
        Case Else: tType = "INFO"
    End Select
    tLogString = uGetTimeStamp & tType & "." & inTag & ": " & inText
    inTextVariable = inText
End Function

Public Sub sLogInit(inFolder)
    gInitResult = gInitFalse
    gNewLogString = vbNullString
    gLogFileFolder = inFolder
    If uFileExists(inFolder) Then: gInitResult = gInitTrue
End Sub

Public Sub sLogIt(inText)
    If gNewLogString <> vbNullString Then
        gNewLogString = Now() & vbTab & inText & vbCrLf & gNewLogString
    Else
        gNewLogString = Now() & vbTab & inText
    End If
End Sub

Public Sub sLogWrite()
Dim tOldLogText
Dim FSObj As Object
Dim tTextFile
    If gInitResult <> gInitTrue Then: Exit Sub
    Set FSObj = CreateObject("Scripting.FileSystemObject")
    sLogIt "[КОНЕЦ] Завершение работы."
    If FSObj.FileExists(gLogFileFolder & gLogFileName) Then
        Set tTextFile = FSObj.OpenTextFile(gLogFileFolder & gLogFileName, 1)
        tOldLogText = tTextFile.ReadAll
        tTextFile.Close
    Else
        tOldLogText = vbNullString
    End If
    Set tTextFile = FSObj.OpenTextFile(gLogFileFolder & gLogFileName, 2, True)
    If tOldLogText <> vbNullString Then
        tTextFile.WriteLine gNewLogString
        tTextFile.Write tOldLogText
    Else
        tTextFile.Write gNewLogString
    End If
    tTextFile.Close
End Sub
