Attribute VB_Name = "UMODOv1"
'UTILITY MODULE OUTLOOK v1
'07.12.2015
Option Explicit

'clipboard
Public Const cHandle = &H42
Public Const cTextMode = 1 'CF_TEXT
Public Const cMaxSize = 4096

'common
Public uD2S(255) As String
Public ClipBoard_Error As String

'keyboard work
Public Enum enGetAsyncKeyState
    vbEsc = &H1B
    vbShift = &H10
    vbControl = &H11
    vbAlt = &H12
    vbLButton = &H1
    vbRButton = &H2
    vbEnter = &HD
    vbSpace = &H20
End Enum

'api
#If VBA7 Then
    Private Declare PtrSafe Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
    Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As enGetAsyncKeyState) As Integer
    Declare PtrSafe Function ShowWindow Lib "user32" (ByVal lHwnd As Long, ByVal lCmdShow As Long) As Boolean
    Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#Else
    Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
    Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Declare Function CloseClipboard Lib "User32" () As Long
    Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
    Declare Function EmptyClipboard Lib "User32" () As Long
    Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any,  ByVal lpString2 As Any) As Long
    Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Declare Function GetClipboardData Lib "User32" (ByVal wFormat As Long) As Long
    Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As enGetAsyncKeyState) As Integer
    Declare Function ShowWindow Lib "User32" (ByVal lHwnd As Long, ByVal lCmdShow As Long) As Boolean
    Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Declare Function GetTickCount Lib "kernel32" () As Long
    Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
#End If

'safe UBOUND for any value (even not arrays)
Public Function uSafeUBound(inArray)
    uSafeUBound = -2
    On Error Resume Next
        uSafeUBound = UBound(inArray)
        If IsArray(inArray) And Err.Number <> 0 Then: uSafeUBound = -1
    On Error GoTo 0
End Function

'Get file MD5 hash
'parameter full path with name of file returned in the function as an MD5 hash
'Set a reference to mscorlib 4.0 64-bit
'Make sure that Net Framework 3.5 (includes .Net 2 and .Net 3) is installed and enabled and not only the Net Framework 4.8 Advanced Services
Public Function uFileToMD5(inFilePath, outHashString, Optional inCodingB64 = False, Optional inFileSizeLimit = 0) As Boolean
    Dim tMD5CryptoService, tFSO, tFile
    Dim tBytes() As Byte
    Dim tFileSizeLimit, tDefaultSizeLimit, tMaxSizeLimit
    
    'defaults
    uFileToMD5 = False
    outHashString = vbNullString
    Debug.Print "#1"
    
    'file size routines \\ 1mb = 1 048 576
    tDefaultSizeLimit = 20971520
    tMaxSizeLimit = 209715200
    If IsNumeric(inFileSizeLimit) Then
        tFileSizeLimit = Fix(inFileSizeLimit)
        If tFileSizeLimit <= 0 Then
            tFileSizeLimit = tDefaultSizeLimit
        Else
            tFileSizeLimit = tFileSizeLimit * 1048576
            If tFileSizeLimit > tMaxSizeLimit Then: tFileSizeLimit = tMaxSizeLimit
        End If
    Else
        tFileSizeLimit = tDefaultSizeLimit
    End If
    
    'turn off errcontrol
    On Error GoTo SafeFinisher
    
    Debug.Print "#2"
    
    'is file exists?
    Set tFSO = CreateObject("Scripting.FileSystemObject")
    'If tFSO.FileExists(inFilePath) Then: Exit Function
    Set tFile = tFSO.GetFile(inFilePath)
    
    Debug.Print "#3"
    
    'filesize check with error.raising
    'Debug.Print tFile.Size & " \\ " & tFileSizeLimit
    If tFile.Size = 0 Or tFile.Size > tFileSizeLimit Then
        Err.Raise 20032, "uFileToMD5", "Current file size not allowed!", "123", "123"
    End If
    
    Debug.Print "#4"

    'Convert the string to a byte array and hash it
    Set tMD5CryptoService = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    tBytes = fGetFileBytes(tFile)
    Debug.Print "#4A:" & UBound(tBytes)
    tBytes = tMD5CryptoService.ComputeHash_2(tBytes)
    
    Debug.Print "#5"
    
    If inCodingB64 Then
       outHashString = fGetBaseConvertedString(tBytes, 1) 'base-64
    Else
       outHashString = fGetBaseConvertedString(tBytes, 0) 'hex
    End If
    
    'no error - return success
    uFileToMD5 = True
        
'clearing after and error finisher
SafeFinisher:
    Debug.Print "ERR: " & Err.Description & " // " & Err.Source & " // " & Err.Number
    Err.Clear
    Set tFSO = Nothing
    Set tFile = Nothing
    Set tMD5CryptoService = Nothing
    'On Error GoTo 0
End Function

Public Function fTestMD5File()
    Dim tPath, tOutString, tHashResult
    tPath = "C:\Users\haustov\Desktop\��������\������������PBELKA26-TP.xlsx"
    tHashResult = uFileToMD5(tPath, tOutString, , 2)
    
    Debug.Print "R=" & tHashResult & "; MD5=" & tOutString
End Function

'makes byte array from file
'Set a reference to mscorlib 4.0 64-bit
'Make sure that Net Framework 3.5 (includes .Net 2 and .Net 3) is installed and enabled and not only the Net Framework 4.8 Advanced Services
Private Function fGetFileBytes(inFile) As Byte()
    Dim tFileNumber, tByteValue() As Byte
    
    'With CreateObject("Adodb.Stream")
    '    .Type = 1 ' adTypeBinary
    '    .Open
    '    .LoadFromFile inFile.Path
    '    .Position = 0
    '    tByteValue = .Read
    '    .Close
    'End With
    
    tFileNumber = FreeFile
    
    ''// Does file exist?
    If LenB(Dir(inFile.Path)) Then
        
        Open inFile.Path For Binary Access Read As tFileNumber
        'a zero length file content will give error 9 here
        
        ReDim tByteValue(0 To LOF(tFileNumber) - 1&) As Byte
        Get tFileNumber, , tByteValue
        Close tFileNumber
    'Else
        'Err.Raise 53 'File not found
    End If
    
    'fin
    fGetFileBytes = tByteValue
    Erase tByteValue
End Function

'used to produce a base-64/hex output
'Set a reference to mscorlib 4.0 64-bit
'Make sure that Net Framework 3.5 (includes .Net 2 and .Net 3) is installed and enabled and not only the Net Framework 4.8 Advanced Services
Function fGetBaseConvertedString(inValue, Optional inMode = 0)
    Dim tDocument, tModeText
      
    Set tDocument = CreateObject("MSXML2.DOMDocument")
    
    Select Case inMode
        Case 1: tModeText = "bin.base64"
        Case Else: tModeText = "bin.Hex"
    End Select
    
    With tDocument
        .LoadXML "<root />"
        .DocumentElement.DataType = tModeText
        .DocumentElement.nodeTypedValue = inValue
    End With
    
    fGetBaseConvertedString = Replace(tDocument.DocumentElement.Text, vbLf, "")
    
    Set tDocument = Nothing
End Function

'sleep in seconds
Public Sub uSleep(inTime)
    If inTime <= 0 Then: Exit Sub
    Sleep inTime * 1000
End Sub

Public Sub uAddToList(inList, inItem, Optional inSeparator = ";")
    If inList = vbNullString Then
        inList = inItem
    Else
        inList = inList & inSeparator & inItem
    End If
End Sub

Public Sub uAddToListUnique(outList, inItem, Optional inSeparator = ";")
Dim tElements, tElement
    'unique check
    tElements = Split(outList, inSeparator)
    For Each tElement In tElements
        If tElement = inItem Then: Exit Sub      'not unique > exit
    Next
    'add unique item to list
    If outList = vbNullString Then
        outList = inItem
    Else
        outList = outList & inSeparator & inItem
    End If
End Sub

Public Function uItemInList(inList, inItem, Optional inSeparator = ";")
Dim tElements, tElement, tIndex
    'unique check
    tElements = Split(inList, inSeparator)
    tIndex = -1
    uItemInList = tIndex
    For Each tElement In tElements
        tIndex = tIndex + 1
        If tElement = inItem Then
            uItemInList = tIndex
            Exit Function 'found item > exit
        End If
    Next
End Function

'======================================================================[FUNC][SECTION][>]
'001. ����� ���������� "�� ���������"
' - inNewPath - ����� ����� ����
Public Sub uChangeDir(inNewPath As String)
Dim fCurDir As String
    fCurDir = CurDir
    SetCurrentDirectory inNewPath
End Sub

'002. ����� � ��������� ����������� ������� (��������� ������ uD2S)
Public Sub uD2SInit()
Dim tTotalSize, tCounterSize As Variant
Dim tCounter()
Dim i, j As Variant
    If uD2S(1) = "A" Then: Exit Sub
    tTotalSize = UBound(uD2S)
    tCounterSize = 0
    ReDim tCounter(tCounterSize)
    tCounter(0) = 65
    'n = 65
    For i = 1 To tTotalSize
        uD2S(i) = vbNullString
        For j = tCounterSize To 0 Step -1
            uD2S(i) = uD2S(i) & Chr(tCounter(j))
        Next j
        '=INC
        tCounter(0) = tCounter(0) + 1
        For j = 0 To tCounterSize
            If tCounter(j) = 91 Then
                tCounter(j) = 65
                If j < tCounterSize Then
                    tCounter(j + 1) = tCounter(j + 1) + 1
                Else
                    tCounterSize = tCounterSize + 1
                    ReDim Preserve tCounter(tCounterSize)
                    tCounter(tCounterSize) = 65
                    Exit For
                End If
            End If
        Next j
    Next i
End Sub

'007. ��������� ��� ����� �� ���� � ���� (���� ������ - ��� ����� �� ������)
' - inPath  - ���� � �����
Public Function uGetFileName(inPath As Variant) As Variant
    On Error Resume Next
        uGetFileName = Right(inPath, Len(inPath) - InStrRev(inPath, "\"))
        If Err.Number > 0 Then
            uGetFileName = vbNullString
        End If
    On Error GoTo 0
End Function

'009.1. �������� �� ������������� �����\���������� �� ��������� ����
' - inFileName      - ���� � ����� ��� ����������
Public Function uFileExists(inFileName As Variant) As Boolean
Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    uFileExists = objFSO.FileExists(inFileName) Or objFSO.FolderExists(inFileName)
End Function

'009.2. �������� �����
' - inFileName      - ���� � ����� ��� ����������
Public Function uDeleteFile(inFileName As Variant) As Boolean
Dim objFSO As Object
    uDeleteFile = True
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(inFileName) Then
        On Error Resume Next
            objFSO.DeleteFile inFileName, True
            If objFSO.FileExists(inFileName) Or Err.Number > 0 Then
                uDeleteFile = False
            End If
        On Error GoTo 0
    End If
End Function

'009.3. �������� ������ ��������
' - inPathName      - ���� ����� ����������
Public Function uFolderCreate(inPathName As Variant) As Boolean
Dim objFSO As Object
    If uFileExists(inPathName) Then
        uFolderCreate = True
        Exit Function
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
        objFSO.CreateFolder (inPathName)
        If Err.Number > 0 Then
            uFolderCreate = False
        Else
            uFolderCreate = True
        End If
    On Error GoTo 0
End Function

'009.4. ����������� �����
' - inSourceFilePath            - ���� � �����
' - inDestinationFilePath       - ���� � �����
Public Function uMoveFile(inSourceFilePath, inDestinationFilePath) As Boolean
Dim objFSO As Object
    uMoveFile = False
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(inSourceFilePath) Then
        On Error Resume Next
            objFSO.CopyFile inSourceFilePath, inDestinationFilePath
            If Err.Number > 0 Then: Exit Function
            objFSO.DeleteFile inSourceFilePath, True
            If Err.Number > 0 Then: Exit Function
        On Error GoTo 0
    End If
    uMoveFile = True
End Function

Public Function uCopyFile(inSourceFilePath, inDestinationFilePath) As Boolean
Dim objFSO As Object
    uCopyFile = False
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(inSourceFilePath) Then
        On Error Resume Next
            'STEP 1 - Delete DESTINATION file if exists
            If objFSO.FileExists(inDestinationFilePath) Then
                objFSO.DeleteFile inDestinationFilePath, True
                If Err.Number > 0 Then: Exit Function
            End If
            'STEP 2 - Try to COPY SOURCE to DESTINATION
            objFSO.CopyFile inSourceFilePath, inDestinationFilePath
            If Err.Number > 0 Then: Exit Function
        On Error GoTo 0
    End If
    uCopyFile = True
End Function

'012. ��������� ���������� ���� � ������
' - inMonth         - ����� (����� ��� �������, ��� � �������)
' - inYear          - ���
Public Function uDaysPerMonth(inMonth, inYear As Variant) As Variant
    uDaysPerMonth = 0
    Select Case LCase(inMonth)
        Case "������", 1:       uDaysPerMonth = 31
        Case "�������", 2:
            If (inYear Mod 4) = 0 Then
                                uDaysPerMonth = 29
            Else
                                uDaysPerMonth = 28
            End If
        Case "����", 3:         uDaysPerMonth = 31
        Case "������", 4:       uDaysPerMonth = 30
        Case "���", 5:          uDaysPerMonth = 31
        Case "����", 6:         uDaysPerMonth = 30
        Case "����", 7:         uDaysPerMonth = 31
        Case "������", 8:       uDaysPerMonth = 31
        Case "��������", 9:     uDaysPerMonth = 30
        Case "�������", 10:     uDaysPerMonth = 31
        Case "������", 11:      uDaysPerMonth = 30
        Case "�������", 12:     uDaysPerMonth = 31
    End Select
    If inYear <= 0 Then: uDaysPerMonth = 0
End Function

'013. ������������ �������� ������ � ��� �����
' - inMonth         - ����� � ��������� ����
Public Function uMonthC2D(inMonth As Variant) As Variant
    uMonthC2D = 0
    Select Case Trim(LCase(inMonth))
        Case "������":      uMonthC2D = 1
        Case "�������":     uMonthC2D = 2
        Case "����":        uMonthC2D = 3
        Case "������":      uMonthC2D = 4
        Case "���":         uMonthC2D = 5
        Case "����":        uMonthC2D = 6
        Case "����":        uMonthC2D = 7
        Case "������":      uMonthC2D = 8
        Case "��������":    uMonthC2D = 9
        Case "�������":     uMonthC2D = 10
        Case "������":      uMonthC2D = 11
        Case "�������":     uMonthC2D = 12
    End Select
End Function

'014. ������������ ����� ������ � ��� ��������
' - inMonth         - ����� � �������� ����
Public Function uMonthD2C(inMonth As Variant) As Variant
    uMonthD2C = vbNullString
    Select Case inMonth
        Case 1:     uMonthD2C = "������"
        Case 2:     uMonthD2C = "�������"
        Case 3:     uMonthD2C = "����"
        Case 4:     uMonthD2C = "������"
        Case 5:     uMonthD2C = "���"
        Case 6:     uMonthD2C = "����"
        Case 7:     uMonthD2C = "����"
        Case 8:     uMonthD2C = "������"
        Case 9:     uMonthD2C = "��������"
        Case 10:    uMonthD2C = "�������"
        Case 11:    uMonthD2C = "������"
        Case 12:    uMonthD2C = "�������"
    End Select
End Function

'016. ������� ��������� ��������� �������� � ����� ������ (��� ������� ���������� ����)
'inText - ����� ������������ � ����� ������
Function ClipBoard_Write(inText As String) As Boolean
Dim tGlobalMemory As Long, tLockedGlobalMemory As Long
Dim tClipMemory As Long, tBuf As Long
    ClipBoard_Save = True
    ClipBoard_Error = vbNullString
'1 // ��������� ������ �� ����� ��������� ������
    tGlobalMemory = GlobalAlloc(cHandle, Len(inText) + 1)
'2 // ���������� ����������� ������� ������ � ��������� ��� ������
    tLockedGlobalMemory = GlobalLock(tGlobalMemory)
'3 // �������� ����� � ���������� ���� ������
    tLockedGlobalMemory = lstrcpy(tLockedGlobalMemory, inText)
'4 // ������������ ������
    If GlobalUnlock(tGlobalMemory) = 0 Then
'5 // ������� ����� ������ ��� ����������� ������
        If OpenClipboard(0&) <> 0 Then
'6 // ������� ����� ������
            tBuf = EmptyClipboard()
'7 // ����� ����� � ����� ������
            tClipMemory = SetClipboardData(cTextMode, tGlobalMemory)
'8 // ������� ����� ������
            If CloseClipboard() = 0 Then
                ClipBoard_Error = "Clipboard: �� ������� ������� ����� ������."
                ClipBoard_Save = False
            End If
        Else
            ClipBoard_Error = "Clipboard: �� ������� ������� ����� ������."
            ClipBoard_Save = False
        End If
    Else
        ClipBoard_Error = "Clipboard: �� ������� �������������� ������."
        ClipBoard_Save = False
    End If
End Function

'017. ������� ��������� ��������� ������ �� ������ ������ (��� ������� ���������� ����, ��� ����� ���������� ������ �� ������ ������)
Function ClipBoard_Read()
Dim tClipMemory As Long
Dim tLockedClipMemory As Long
Dim tText As String
Dim tBuf As Long
    ClipBoard_Read = False
    ClipBoard_Error = vbNullString
'1 // ������� ����� ������ ��� ����������� ������
    If OpenClipboard(0&) <> 0 Then
'2 // �������� ��������� �� ���� ������ � �������� ����������� � ������ ������
        tClipMemory = GetClipboardData(cTextMode)
'3 // ���� ��������� �� ������ �� ����� ��������
        If Not (IsNull(tClipMemory)) Then
'4 // ��������� ������� ������ ��������� ������� ������
            tLockedClipMemory = GlobalLock(tClipMemory)
'5 // ���� ��������� �� ������ �� ����� ��������
            If Not IsNull(tLockedClipMemory) Then
'6 // ������� ��� ��������� ������ �������� ������
                tText = Space$(cMaxSize)
'7 // ��������� ������ � ���� ��������� ���������� �� ������ ������
                tBuf = lstrcpy(tText, tLockedClipMemory)
'8 // ������������ ���� ������ ������ ������
                tBuf = GlobalUnlock(tClipMemory)
'9 // ������� ��������� ���������� � ����������� ����� �� ������� (null termination character (CHR=0))
                tText = Mid(tText, 1, InStr(1, tText, Chr$(0), 0) - 1)
'10 // ��������� ����������� ������ ����� �������
                ClipBoard_Read = tText
            Else
                ClipBoard_Error = "Clipboard: �� ���������� ���������� ����� ������ ��� ������."
                'EXIT
            End If
        Else
            ClipBoard_Error = "Clipboard: �� ������� �������� ��������� �� ���� ������ ������ ������."
            'EXIT
        End If
'11 // ������� ����� ������
        tBuf = CloseClipboard()
    Else
        ClipBoard_Error = "Clipboard: �� ������� ������� ����� ������."
        'EXIT
    End If
End Function

'006. ��������� ����� �� ��������� ���� (���� ������ - ������)
' - inPathName  - ���� � ����������� �����
Public Function uWorkBookOpen(inPathName As Variant) As Boolean
    On Error Resume Next
        Workbooks.Open FileName:=inPathName, UpdateLinks:=False, ReadOnly:=True
        ThisWorkbook.Activate
        Select Case Err.Number
        Case Is = 0
            uWorkBookOpen = True
            Windows(inPathName).Visible = False
        Case Else
            uWorkBookOpen = False
        End Select
    On Error GoTo 0
End Function

'018. ������ ����
Public Function uHideWindow(inWindowName) As Boolean
Dim tHandle, tRes
    uHideWindow = False
    On Error Resume Next
        tHandle = FindWindow(vbNullString, inWindowName)
        If tHandle > 0 Then 'Err.Number
            uHideWindow = ShowWindow(tHandle, 0) '5 - SHOW
        End If
    On Error GoTo 0
    'uHideWindow = True
End Function

Public Function uGetFileExtension(inFileName) As String
Dim tValue
    uGetFileExtension = vbNullString
    tValue = InStrRev(inFileName, ".")
    If tValue > 0 Then
        uGetFileExtension = UCase(Right(inFileName, Len(inFileName) - tValue))
    End If
End Function

'======================================================================[FUNC][SECTION][<]
