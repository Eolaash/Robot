VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CATSDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CATSDownloader v001 23.06.2020
'����� ������������ ������ � ������� ������� ���
'v001 - �����

Option Explicit

Private Const pClassName = "CATSDownloader"
Private Const pClassVersion = "001"
Private Const pClassDateStamp = "23.06.2020"
Private Const pATSDomainAddress = "www.atsenergo.ru"

Private pQuickLog, pPrintLog, pErrorText
Private pIsObjectActive
Private pPartCode, pUserName, pPassword
Private pWinHTTP, pFSO, pRegExp, pShell 'objects
Private pProxyEnabled, pProxyAddress, pProxyPort, pProxyLogin, pProxyPassword, pLoggedIn
Private pATSAuthLink, pATSNAuthLink, pHostLink, pOriginLink, pReportShortLink, pReportPersonlaLink, pNReportFileBaseLink

'----------------------------FUNC PART------
Private Sub fLogReset()
    pErrorText = vbNullString
    pQuickLog = vbNullString
End Sub

Private Function fGetLogTag(inTag)
    fGetLogTag = pClassName & "." & inTag
End Function

Private Function fGetTimeStamp()
    fGetTimeStamp = "[" & Format(Now(), "MM.DD hh:mm:ss") & "] "
End Function

Private Sub fLogAdd(inTag, inType, inText)
Dim tType, tLogString
    Select Case inType
        Case 1: tType = "WARN"
        Case 2: tType = "CRIT"
        Case Else: tType = "INFO"
    End Select
    tLogString = fGetTimeStamp & tType & "." & inTag & ": " & inText
    If inType = 2 Then: pErrorText = tLogString
    If pQuickLog <> vbNullString Then
        pQuickLog = pQuickLog & vbCrLf & tLogString
    Else
        pQuickLog = tLogString
    End If
    If pPrintLog Then: Debug.Print tLogString
End Sub

Private Sub Class_Initialize()
    Dim tLogTag
    On Error Resume Next
        pIsObjectActive = False
        tLogTag = fGetLogTag("INIT")
        fLogReset
        pPrintLog = True
        
        'Objects init
        Set pWinHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
        Set pFSO = CreateObject("Scripting.FileSystemObject")
        Set pRegExp = CreateObject("VBScript.RegExp")
        Set pShell = CreateObject("Shell.Application")
        
        'Default settings
        pRegExp.IgnoreCase = True
        pRegExp.Global = True
        '
        pProxyEnabled = False
        pATSNAuthLink = "https://" & pATSDomainAddress & "/nauth"
        pHostLink = pATSDomainAddress
        pOriginLink = "https://" & pATSDomainAddress
        pATSAuthLink = "https://" & pATSDomainAddress & "/auth"
        pReportShortLink = pATSDomainAddress & "/nreports" 'with S
        pNReportFileBaseLink = "https://" & pATSDomainAddress & "/nreport" 'without S
        pReportPersonlaLink = "https://" & pATSDomainAddress & "/nreports?access=personal"
        pLoggedIn = False
        ClearNReportCredentials
        
        'ERROR Check
        If Err.Number = 0 Then
            pIsObjectActive = True
            fLogAdd tLogTag, 0, "������������� ������ �������! [DOMAIN=" & pATSDomainAddress & "][CLASSINFO: " & pClassName & " v" & pClassVersion & " release " & pClassDateStamp & "]"
        Else
            fLogAdd tLogTag, 2, Err.Source & ": " & Err.Description
        End If
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
Dim tLogTag
    tLogTag = fGetLogTag("KILL")
    pIsObjectActive = False
    
    'Drop proxy
    SetProxy False, vbNullString, 0, vbNullString, vbNullString
    
    'Objects release
    Set pWinHTTP = Nothing
    Set pFSO = Nothing
    Set pRegExp = Nothing
    Set pShell = Nothing
    
    'Clear memory
    ClearNReportCredentials
    
    'Logging
    fLogAdd tLogTag, 0, "������ ���������."
    fLogReset
End Sub

Public Sub ClearNReportCredentials()
    Dim tLogTag
    tLogTag = fGetLogTag("ClearNReportCredentials")
    'fLogAdd tLogTag, 0, "��������� ���������������� NREPORT ��������."
    pPartCode = vbNullString
    pUserName = vbNullString
    pPassword = vbNullString
End Sub

Public Sub SetNReportCredentials(ioPartCode, ioUserName, ioPassword, Optional inForceResetParams = False)
    Dim tLogTag
    
    On Error Resume Next
        tLogTag = fGetLogTag("SetNReportCredentials")
        'fLogAdd tLogTag, 0, "��������� ���������������� NREPORT ������."
        pPartCode = ioPartCode
        pUserName = ioUserName
        pPassword = ioPassword
        
        If inForceResetParams Then
            ioPartCode = vbNullString
            ioUserName = vbNullString
            ioPassword = vbNullString
        End If
        
     On Error GoTo 0
End Sub

Private Function fNReportLoginAttempt()
    Dim tLogTag, tCookie, tPostData, tRedirectLink
    
    On Error Resume Next
    
        ' 00 \\ ���������������
        tLogTag = fGetLogTag("LoginAttempt")
        fLogAdd tLogTag, 0, "������� �����������: " & pATSNAuthLink
        fNReportLoginAttempt = False
        pLoggedIn = False
        
        ' 01 \\ ��������� ��������� ������ �� ��������� �������� ����������� ����� ���
        'pWinHTTP.Open "GET", pATSNAuthLink, False
        'pWinHTTP.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        'pWinHTTP.SetRequestHeader "Accept-Charset", "windows-1251,utf-8;q=0.7,*;q=0.7"
        'pWinHTTP.SetRequestHeader "Host", pHostLink
        'pWinHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        'pWinHTTP.SetRequestHeader "Connection", "keep-alive"
        'pWinHTTP.SetRequestHeader "Origin", pOriginLink
        'pWinHTTP.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36"
        
        pWinHTTP.Open "GET", "https://www.atsenergo.ru/nauth", False
        pWinHTTP.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        pWinHTTP.SetRequestHeader "Accept-Charset", "windows-1251,utf-8;q=0.7,*;q=0.7"
        pWinHTTP.SetRequestHeader "Host", "www.atsenergo.ru"
        pWinHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        pWinHTTP.SetRequestHeader "Connection", "keep-alive"
        pWinHTTP.SetRequestHeader "Origin", "https://www.atsenergo.ru"
        pWinHTTP.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36"
           
        ' 02 \\ ���������� ��������� ������ � ���� ������
        pWinHTTP.Send
                
        If Err.Number <> 0 Then
            fLogAdd tLogTag, 2, "������ pWinHTTP ������ #1 (SEND) > " & Err.Source & ": " & Err.Description
            Exit Function
        End If
        
        pWinHTTP.WaitForResponse
        
        If Err.Number <> 0 Then
            fLogAdd tLogTag, 2, "������ pWinHTTP ������ #1 (WAIT) > " & Err.Source & ": " & Err.Description
            Exit Function
        End If
        
        ' 03 \\ ���������� ������ ��������� ���� �� ������� �������������� ������ (���� ��� ��� - ������ ��������)
        tCookie = pWinHTTP.getResponseHeader("Set-Cookie") ' ����� ��������� Set-Cookie �� ������ �������
        'fLogAdd tLogTag, 0, "COOKIE=" & pWinHTTP.GetAllResponseHeaders()
        'If InStr(tCookie, "JSESSIONID=") > 0 Then
        '    pSessionCookie = tCookie
        '    fLogAdd tLogTag, 0, "������ ����������������!"
        'End If
        
        If InStr(tCookie, "JSESSIONID=") = 0 Then
            fLogAdd tLogTag, 2, "����������� ��������! ������ ��������� �������������� ������!"
            Exit Function
        End If
        
        'fLogAdd tLogTag, 0, pSessionCookie
        
        ' 04 \\ ��������� �������������� ��������� (�������� ����� ����� ���������� ����� - https://docs.microsoft.com/en-us/windows/win32/winhttp/winhttprequestoption)
        pWinHTTP.Option(6) = False 'WinHttpRequestOption_EnableRedirects \\ The default value of this option is VARIANT_TRUE to indicate that requests are automatically redirected.
        
        ' 05 \\ ������ �� ����� ����������� ������ �� �������� ������ ����������� �� �������� ����������� (��������� ������)
        pWinHTTP.Open "POST", pATSAuthLink, False
        tPostData = "partcode=" & pPartCode & "&username=" & pUserName & "&password=" & pPassword
        
        pWinHTTP.SetRequestHeader "Cookie", tCookie '���������� ������ ���������� cookie ��� �������� �������������� ������������� � ���������� ������� ������
        pWinHTTP.SetRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        pWinHTTP.SetRequestHeader "Accept-Charset", "windows-1251,utf-8;q=0.7,*;q=0.7"
        pWinHTTP.SetRequestHeader "Host", pHostLink
        pWinHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        pWinHTTP.SetRequestHeader "Connection", "keep-alive"
        pWinHTTP.SetRequestHeader "Origin", pOriginLink
        pWinHTTP.SetRequestHeader "Referer", pATSAuthLink
        pWinHTTP.SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36"
        
        ' 06 \\ ���������� ��������� ������ � ���� ������
        pWinHTTP.Send tPostData
        
        If Err.Number <> 0 Then
            fLogAdd tLogTag, 2, "������ pWinHTTP ������ #2 (SEND) > " & Err.Source & ": " & Err.Description
            Exit Function
        End If
        
        pWinHTTP.WaitForResponse
        
        If Err.Number <> 0 Then
            fLogAdd tLogTag, 2, "������ pWinHTTP ������ #2 (WAIT) > " & Err.Source & ": " & Err.Description
            Exit Function
        End If
        
        ' 07 \\ ��������� ������, ���� � ��������� ���� ������ ����� (��������� � ��������), ������� ��� ����������� ������
        tRedirectLink = pWinHTTP.getResponseHeader("Location") ' ��� ������� ����������� ���� ������������� �� ��������� ��������
        If Not InStr(tRedirectLink, pReportShortLink) > 0 Then
            fLogAdd tLogTag, 2, "����������� ��������! ������ �� ����� �������� ������ �����������! ��������� ������, ����� � ��� ���������."
            Exit Function
        End If
        
        ' 08 \\ ����������� �������
        fNReportLoginAttempt = True
        pLoggedIn = True
        fLogAdd tLogTag, 0, "����������� �������!"
        
    On Error GoTo 0
End Function

Private Function fNReportByClassGet(inReportClass, inMode, inPartCode, inGTPCode, inRegion, inZoneCode, inYear, inMonth, inDay, outResult)
    
    'inMode: ANY - Check Availability \ 1 - Get PATTERN for RegExp
    fNReportByClassGet = True
    outResult = vbNullString
    
    Select Case inReportClass
        '81. ����������� ����� �� ������� ��������, ������������ ��� ���������� ��������� ������������
        '������ ����� 20200601_BELKAMKO_ZONE1_TG_power_consumer_avans.xls
        Case "TG_power_consumer_avans":
            If inMode = 1 Then
                outResult = fGetDateStamp(inYear, inMonth, inDay) & "_" & inPartCode & "_" & inZoneCode & "_TG_power_consumer_avans.xls"
            End If
        '28. ����� �� ����������� �����
        '������ ����� 20200626_BELKAMKO_PBELKA11_buy_norem.xls
        Case "buy_norem":
            If inMode = 1 Then
                outResult = fGetDateStamp(inYear, inMonth, inDay) & "_" & inPartCode & "_" & inGTPCode & "_buy_norem.xls"
            End If
        Case Else: fNReportByClassGet = False
    End Select
End Function

Private Function fNReportIsLoggedIn()
    Dim tLogTag

    On Error Resume Next
    
        tLogTag = fGetLogTag("fNReportIsLoggedIn")
        fNReportIsLoggedIn = False
                
        pWinHTTP.Open "GET", pReportPersonlaLink, False
        pWinHTTP.Send
        
        If Err.Number <> 0 Then
            fLogAdd tLogTag, 2, "������ pWinHTTP ������ #1 (SEND) > " & Err.Source & ": " & Err.Description
            Exit Function
        End If
        
        pWinHTTP.WaitForResponse
        
        If Err.Number <> 0 Then
            fLogAdd tLogTag, 2, "������ pWinHTTP ������ #1 (WAIT) > " & Err.Source & ": " & Err.Description
            Exit Function
        End If
        
        If pWinHTTP.Status = 200 Then: fNReportIsLoggedIn = True
    
    On Error GoTo 0
End Function

Private Function fNReportDownloadFile(inYear, inMonth, inDay, inReportClass, inPartCode, inGTPCode, inRegion, inSaveFolder, outFilePath, Optional inZipMode = True, Optional inSaveFolderCreate = True, Optional inForcedReload = False)
    Dim tLogTag
    Dim tRegion, tZoneName, tDate, tReportPageLink, tResult, tValueString, tFileLink, tZipFileLink, tPos, tFileName, tFileNameZipped, tSavePath, tZipSavePath, tSaveFolder
    Dim tMatches
    
    On Error Resume Next
        
        ' 00 // ����������
        fNReportDownloadFile = 1 '�������������
        outFilePath = vbNullString
        tLogTag = fGetLogTag("DownloadFile")
        fLogAdd tLogTag, 0, "���������� �����-������ [CLASS=" & inReportClass & "; YEAR=" & inYear & "; MONTH=" & inMonth & "; DAY=" & inDay & "; PARTCODE=" & inPartCode & "; GTPCODE=" & inGTPCode & "; REGION=" & inRegion & "; SAVEPATH=" & inSaveFolder & "; ZIPMODE=" & inZipMode & "]"
        
        ' 01 // �������� �� �������� ������ ������ ������ ������ �������
        If Not fNReportByClassGet(inReportClass, 0, "", "", "", "", "", "", "", tValueString) Then
            fLogAdd tLogTag, 2, "����-����� ������ <" & inReportClass & "> �� ����� ���� �������� ���� ������� ������. ���������� � ������������."
            Exit Function
        End If
        
        ' 02 // �������� ����� ��������
        If Not pFSO.FolderExists(inSaveFolder) Then
            If inSaveFolderCreate Then '���� ��������� ����������� � ������� ��������
                pFSO.CreateFolder inSaveFolder
                If Not pFSO.FolderExists(inSaveFolder) Then
                    fLogAdd tLogTag, 2, "���� ���������� �� ��������� � ������� ��� ������� ��������: " & inSaveFolder
                    Exit Function
                End If
            Else
                fLogAdd tLogTag, 2, "���� ���������� �� ���������: " & inSaveFolder
                Exit Function
            End If
        End If
        
        tSaveFolder = inSaveFolder
        If Right(tSaveFolder, 1) <> "\" Then: tSaveFolder = tSaveFolder & "\"
        
        ' 03 // ��������� �������� ���������� ��� �������� ����� ����� ������ �� ����
        fNReportDownloadFile = 2
        
        ' 03.01 // ������ - tRegion (eur \ sib)
        Select Case inRegion
            Case "eur", 1, "1", "ZONE1":
                tZoneName = "ZONE1"
                tRegion = "&region=eur"
            Case "sib", 2, "2", "ZONE2":
                tZoneName = "ZONE2"
                tRegion = "&region=sib"
            Case Else:
                tRegion = vbNullString
        End Select
        
        ' 03.02 // ���� - tDate
        tDate = "&rdate=" & fGetDateStamp(inYear, inMonth, inDay)
        
        ' 03.03 // ����� ������ ��������� � �������� ��������� ������
        tReportPageLink = "https://" & pATSDomainAddress & "/nreport?rname=" & inReportClass & tDate & tRegion
        fLogAdd tLogTag, 0, "������ �������: " & tReportPageLink
        
        ' 04 // ������� ������� ��������� �� �������� ������
        pWinHTTP.Open "GET", tReportPageLink, False
        
        pWinHTTP.Send
        
        If Err.Number <> 0 Then
            fLogAdd tLogTag, 2, "������ pWinHTTP ������ #1 (SEND) > " & Err.Source & ": " & Err.Description
            Exit Function
        End If
        
        pWinHTTP.WaitForResponse
        
        If Err.Number <> 0 Then
            fLogAdd tLogTag, 2, "������ pWinHTTP ������ #1 (WAIT) > " & Err.Source & ": " & Err.Description
            Exit Function
        End If
        
        ' 05 // ��������� ������ �� ������
        tResult = pWinHTTP.ResponseText
        If tResult = vbNullString Then
            fLogAdd tLogTag, 2, "������ ������ ��������� ����������! ������ �����."
            Exit Function
        End If
        
        ' 06 // ����� ������ �� ���� �� ���������� ���������
        fNReportDownloadFile = 3
        
        ' 06.01 // ������� ������� ����� �����-������ ��� ������ �� ����������� ���������
        If Not fNReportByClassGet(inReportClass, 1, inPartCode, inGTPCode, inRegion, tZoneName, inYear, inMonth, inDay, tValueString) Then
            fLogAdd tLogTag, 2, "�� ������� �������� ������� ����� �����-������ (ERR)."
            Exit Function
        ElseIf tValueString = vbNullString Then
            fLogAdd tLogTag, 2, "�� ������� �������� ������� ����� �����-������ (EMPTY STR)."
            Exit Function
        End If
        pRegExp.Pattern = tValueString
        tFileName = tValueString
        tSavePath = tSaveFolder & tFileName
        outFilePath = tSavePath
        fLogAdd tLogTag, 0, "������� �����-������: " & tValueString
        
        ' 06.02 // ���������� �� ���� ���
        If pFSO.FileExists(tSavePath) Then
            If inForcedReload Then
                pFSO.DeleteFile tSavePath
                If Err.Number <> 0 Then
                    fLogAdd tLogTag, 2, "������ �������� ����� (" & tSavePath & ") > " & Err.Source & ": " & Err.Description
                    fNReportDownloadFile = 4 '�������� ������
                    Exit Function
                End If
            Else
                fNReportDownloadFile = -1 '���� ��� ���������� - ��������� �� ����
                fLogAdd tLogTag, 1, "���� ��� ��������� � �����. �������� �� ���������."
                Exit Function
            End If
        End If
        
        ' 06.03 // ZIP-Mode
        If inZipMode Then
            tPos = InStrRev(tFileName, ".")
            tFileNameZipped = Left(tFileName, tPos - 1) & ".zip"
            tZipSavePath = tSaveFolder & tFileNameZipped
            
            If pFSO.FileExists(tZipSavePath) Then
                If inForcedReload Then
                    pFSO.DeleteFile tZipSavePath
                    If Err.Number <> 0 Then
                        fLogAdd tLogTag, 2, "������ �������� ����� (" & tZipSavePath & ") > " & Err.Source & ": " & Err.Description
                        fNReportDownloadFile = 4 '�������� ������
                        Exit Function
                    End If
                Else
                    If fUnzipByWindows(tZipSavePath, tSaveFolder) Then
                        If pFSO.FileExists(tSavePath) Then
                            pFSO.DeleteFile tZipSavePath '������� ������ �� �����
                            fNReportDownloadFile = -1 '���� �������� �� ��� ������������� ������ - ��������� �� ����
                            fLogAdd tLogTag, 1, "���� ��� ��������� � �����. �������� �� ���������."
                            Exit Function
                        Else
                            '���� � ������ �� ��������� ������� ����� - ������� �����
                            pFSO.DeleteFile tZipSavePath
                            If Err.Number <> 0 Then
                                fLogAdd tLogTag, 2, "������ �������� ����� (" & tZipSavePath & ") > " & Err.Source & ": " & Err.Description
                                fNReportDownloadFile = 4 '�������� ������
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        ' 06.04 // ����� �� ��������
        Set tMatches = pRegExp.Execute(tResult)
        If tMatches.Count <> 1 Then
            fLogAdd tLogTag, 2, "������� ������ ������� <" & tMatches.Count & "> ����������� ������ <1>. ���-�� ����� �� ���.."
            Exit Function
        End If
        
        'HELP - "�������� ������������: " & tMatch.Value & vbCrlf & "����� ������� � ������: " & tMatch.FirstIndex & vbCrlf & "�����: " & tMatch.Length
        tFileLink = fHRefExtract(tResult, tMatches.Item(0).FirstIndex, tMatches.Item(0).Length)
        If tFileLink = vbNullString Then
            fLogAdd tLogTag, 2, "�� ������� ���������� HREF ������ �� ����-�����. [PATTERN=" & tMatches.Item(0).Value & "]"
            Exit Function
        End If
        tFileLink = pNReportFileBaseLink & tFileLink 'add href part
        tZipFileLink = tFileLink & "&zip=1" 'link adjusting
        
        ' 07 // ���� �������� ����� �� ������ (���� ����� �� ����� ����� ������ ���� ����� ���, ���� ��� ����� ������� ������)
        If inZipMode Then
            If fGetFileByStream(tZipFileLink, tZipSavePath) Then
                If fUnzipByWindows(tZipSavePath, tSaveFolder) Then
                    pFSO.DeleteFile tZipSavePath '������� ������ �� �����
                    If pFSO.FileExists(tSavePath) Then
                        fNReportDownloadFile = 0 '���� �������� �� ��� ������������� ������ - ��������� �� ����
                        Exit Function
                    End If
                End If
            End If
        End If
        
        If fGetFileByStream(tFileLink, tSavePath) Then
            If pFSO.FileExists(tSavePath) Then
                fNReportDownloadFile = 0 'OK
                Exit Function
            End If
        Else
            fNReportDownloadFile = 5 '�������
        End If
        
    On Error GoTo 0
End Function

Private Function fGetFileByStream(inFileLink, inSavePath)
    Dim tStream, tLogTag
    
    On Error Resume Next
    
        ' 01 // ����������
        tLogTag = fGetLogTag("fGetFileByStream")
        fLogAdd tLogTag, 0, "������ �� ����-�����: " & inFileLink

        fGetFileByStream = False
        
        ' 02 // �������� �������
        pWinHTTP.Open "GET", inFileLink, False
        pWinHTTP.Send
        
        If Err.Number <> 0 Then
            fLogAdd tLogTag, 2, "������ pWinHTTP ������ #1 (SEND) > " & Err.Source & ": " & Err.Description
            Exit Function
        End If
        
        ' 03 // Stream �������� �����
        If pWinHTTP.Status = 200 Then
            
            Set tStream = CreateObject("ADODB.Stream")
            With tStream
                .Type = 1 'adTypeBinary
                .Open
                .Write pWinHTTP.ResponseBody
                .SaveToFile inSavePath
                .Close
            End With
            Set tStream = Nothing
            
            If Err.Number <> 0 Then
                fLogAdd tLogTag, 2, "������ pWinHTTP ������ #2 (STREAM) > " & Err.Source & ": " & Err.Description
                Exit Function
            End If
                        
        End If
        
        fGetFileByStream = True
    On Error GoTo 0
End Function

Private Function fUnzipByWindows(inZipFilePath, inExctractPath)
    Dim tLogTag, tSourceFile, tDestFolder, tFilesInZip
    
    On Error Resume Next
        tLogTag = fGetLogTag("fUnzipByWindows")
        fUnzipByWindows = False
        
        tSourceFile = pFSO.GetAbsolutePathName(inZipFilePath)
        tDestFolder = pFSO.GetAbsolutePathName(inExctractPath)
           
        Set tFilesInZip = pShell.NameSpace(tSourceFile).Items()
        pShell.NameSpace(tDestFolder).copyHere tFilesInZip, 16
        
        Set tFilesInZip = Nothing
        
        If Err.Number <> 0 Then
            fLogAdd tLogTag, 2, "������ > " & Err.Source & ": " & Err.Description
        Else
            fUnzipByWindows = True
        End If
    On Error GoTo 0
End Function

Private Function fHRefExtract(inText, inStart, inLength)
    Dim tPosHRef, tPosA, tPosB
    
    fHRefExtract = vbNullString
    
    ' 02 // Lock HREF
    tPosHRef = InStrRev(inText, "href=", inStart)
    If tPosHRef < 1 Then: Exit Function
    
    ' 03 // Lock A
    tPosA = InStrRev(inText, ">", inStart)
    If tPosA < 1 Then: Exit Function
    
    ' 04 // Lock B
    tPosB = InStrRev(inText, "<", inStart)
    If tPosB < 1 Then: Exit Function
    
    ' 05 // HREF extraction
    fHRefExtract = Mid(inText, tPosHRef + 6, tPosA - (tPosHRef + 6) - 1)
End Function


Private Function fGetDateStamp(inYear, inMonth, inDay)
    Dim tMonth, tDay
    
    tMonth = Fix(inMonth)
    If tMonth < 10 Then: tMonth = "0" & tMonth
    tDay = Fix(inDay)
    If tDay < 10 Then: tDay = "0" & tDay
    fGetDateStamp = inYear & tMonth & tDay
End Function

Public Function GetNReportPersonalFileByClass(inYear, inMonth, inDay, inReportClass, inGTPCode, inRegion, inSaveFolder, outFilePath, Optional inZipMode = True, Optional inSaveFolderCreate = True, Optional inForcedReload = False)
    Dim tLogTag
    
    ' 01 // ����������
    tLogTag = fGetLogTag("GetNReportPersonalFileByClass")
    GetNReportPersonalFileByClass = False
    outFilePath = vbNullString
    
    ' 02 // ��������
    If pWinHTTP Is Nothing Or pIsObjectActive = False Then
        fLogAdd tLogTag, 2, "����� ��� ������ HTTP �� ���������������!"
        Exit Function
    End If
    
    If pPartCode = vbNullString Or pUserName = vbNullString Or pPassword = vbNullString Then
        fLogAdd tLogTag, 2, "�� ������ ��������� ����� � ������������ ������. ������� ����������� SetNReportCredentials(inPartCode, inUserName, inPasssword)!"
        Exit Function
    End If
    
    ' 03 // ����� ��� ������� ����� � ������� (���� ������ �� ����������)
    If Not pLoggedIn Then
        If Not fNReportLoginAttempt Then
            fLogAdd tLogTag, 2, "������ ����� � ������������ ������!"
            Exit Function
        End If
    End If
    
    ' 04 // ������� �������� ������� �����-������
    If fNReportDownloadFile(inYear, inMonth, inDay, inReportClass, pPartCode, inGTPCode, inRegion, inSaveFolder, outFilePath, inZipMode, inSaveFolderCreate, inForcedReload) > 0 Then
        fLogAdd tLogTag, 2, "�� ������� �������� ����-�����!"
        Exit Function
    End If
    
    ' 05 // ����������
    GetNReportPersonalFileByClass = True
End Function

Public Sub SetProxy(inEnabled, inAddress, inPort, inLogin, inPassword)
    Dim tLogTag
    
    On Error Resume Next
    
    tLogTag = fGetLogTag("SetProxy")
    
        ' preventer
        If (pWinHTTP Is Nothing Or pIsObjectActive = False) And inEnabled Then
            pProxyEnabled = False
            pProxyAddress = vbNullString
            pProxyPort = vbNullString
            pProxyLogin = vbNullString
            pProxyPassword = vbNullString
            fLogAdd tLogTag, 0, "������ ������������ ������, �.�. ������ HTTP �� ���������������!"
        End If
        
        ' settings
        pProxyEnabled = inEnabled
        pProxyAddress = inAddress
        pProxyPort = inPort
        pProxyLogin = inLogin
        pProxyPassword = inPassword
        
        If pProxyEnabled Then
            'PARAMS: HTTPREQUEST_PROXYSETTING_DEFAULT = 0; HTTPREQUEST_PROXYSETTING_PRECONFIG = 0; HTTPREQUEST_PROXYSETTING_DIRECT = 1; HTTPREQUEST_PROXYSETTING_PROXY = 2;
            'EXAMPLE: Use proxy_server for all requests outside of the microsoft.com domain. -> WinHttpReq.SetProxy( HTTPREQUEST_PROXYSETTING_PROXY, "proxy_server:80", "*.microsoft.com");
            pWinHTTP.SetProxy 2, inAddress & ":" & inPort, vbNullString 'no bypass using
            pWinHTTP.SetCredentials pProxyLogin, pProxyPassword, 1 'HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0; HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1
            fLogAdd tLogTag, 1, "��������� ������ ����������� ��� ������: " & inAddress & ":" & inPort
        Else
            pWinHTTP.SetProxy 0, vbNullString, vbNullString
            pProxyAddress = vbNullString
            pProxyPort = vbNullString
            pProxyLogin = vbNullString
            pProxyPassword = vbNullString
            fLogAdd tLogTag, 1, "��������� ������ �������� �� ��������� �� ���������."
        End If
        
     On Error GoTo 0
End Sub

Public Property Get ErrorText()
    ErrorText = pErrorText
End Property

Public Property Get IsActive()
    IsActive = pIsObjectActive
End Property

Public Property Let PrintLog(inMode)
    If VarType(inMode) = vbBoolean Then: pPrintLog = inMode
End Property
