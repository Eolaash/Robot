Attribute VB_Name = "M_F63"
'M_F63 version 1.04 (09-02-2021)
'1.01 - fix delay timer and API object kill; delay in seconds -> cnstSendDelay
'1.02 - gate ignore option added
'1.03 - gate calculation now by MSK timezone
'1.04 - uncomm ignore option added
Option Explicit

Private Const cnstModuleName = "M_F63"
Private Const cnstModuleVersion = "1.04"
Private Const cnstModuleDate = "09-02-2021"
Private Const cnstElementSplitter = ";"
Private Const cnstSendDelay = 35 '������� 30 ������ ��� ����� ��������

Private gLocalInit

Private Function fGetLogTag(inTagText)
    fGetLogTag = cnstModuleName & "." & inTagText
End Function

Private Function fLocalInit(Optional inForceInit As Boolean = False)
Dim tLogTag
    fLocalInit = False
    tLogTag = fGetLogTag("F63INI")
    uCDebugPrint tLogTag, 0, "������������� > ���: " & cnstModuleName & "; ������: " & cnstModuleVersion & "; ��������: " & cnstModuleDate
    If inForceInit Or Not gLocalInit Then
        ' // variables
        gLocalInit = False
        ' // objects
        ' // paths
        'gForecastDataPath = gDataPath & "\BRForecast"
        'If Not (gFSO.FolderExists(gForecastDataPath)) Then
        '    If Not (uFolderCreate(gForecastDataPath)) Then
        '        uADebugPrint tLogTag, "�� ������� ����� ����� ������ <gForecastDataPath> �� ����: " & gForecastDataPath
        '        gDataPath = vbNullString
        '        Exit Function
        '    Else
        '        uADebugPrint tLogTag, "������� ����� ������ <gForecastDataPath> �� ����: " & gForecastDataPath
        '    End If
        'End If
        gLocalInit = True
    End If
    fLocalInit = True
End Function

Public Sub fForm63Main()
    If Not fLocalInit Then: Exit Sub
    fF63ExchangeScan gLocalUTC, gTraderInfo.ID
End Sub

'fForm63Manual - standalone TESTING Entry point
Public Sub fForm63Manual()
    Dim tResult, tFirstDay, tErrorText, tCurrentDateTime
    
    If Not fConfiguratorInit Then: Exit Sub
    If Not fLocalInit Then: Exit Sub
    If Not fXMLSmartUpdate("BASIS,FRAME,CALENDAR,R80020DB,XSD80020V2,DICTIONARY,CALCROUTE,CALCDB,XSD80040V2,F63DB,CREDENTIALS") Then: Exit Sub
    
    fF63ExchangeScan gLocalUTC, gTraderInfo.ID, "PBELKA11", True, True, True, -9, 0
    'fF63ExchangeScan gLocalUTC, gTraderInfo.ID, "PBELKAM7", True, True, False, -2, 0
    'tCurrentDateTime = Now()
    'tResult = fWorkDayShiftAdv(CDate(tCurrentDateTime), -3, 0, tFirstDay, tErrorText)
    'Debug.Print tFirstDay
End Sub

' F-01 // ������� ������� ������� �� �������� ��� BASIS
Private Sub fF63ExchangeScan(inLocalUTC, inTraderID, Optional inGTPCode = vbNullString, Optional inForceIgnoreGate = False, Optional inForceIgnoreTimeLimits = False, Optional inForceUncommIgnore = False, Optional inForceDayShift = 0, Optional inForceDaysToReport = 0)
Dim tXPathString, tNode, tNodes, tLogTag
' 00 // ��������������� � ��������� ��������
    tLogTag = fGetLogTag("F63EXSCN")
    If Not gXMLBasis.Active Then: Exit Sub
    If Not gXMLCalendar.Active Then: Exit Sub
    If Not gF63DB.Active Then: Exit Sub
' 01 // ����� ��� ��������������� � ������ � F63 � BASIS
    If inGTPCode = vbNullString Then
        tXPathString = "//trader[@id='" & inTraderID & "']/gtp/exchange/item[(@id='F63' and @enabled='1')]"
    Else
        tXPathString = "//trader[@id='" & inTraderID & "']/gtp[@id='" & inGTPCode & "']/exchange/item[(@id='F63' and @enabled='1')]"
    End If
    
    Set tNodes = gXMLBasis.XML.SelectNodes(tXPathString)
    
    If inGTPCode <> vbNullString Then: uCDebugPrint tLogTag, 1, "����� ��� ������� - " & tNodes.Length
    
    If tNodes.Length = 0 Then: Exit Sub '��� ��������� ��� ������

' 02 // ��������� ��������� ���
    For Each tNode In tNodes
        fF63ExchangeItem tNode, inLocalUTC, inForceIgnoreGate, inForceIgnoreTimeLimits, Fix(inForceDayShift), Fix(inForceDaysToReport), inForceUncommIgnore
    Next
End Sub

' F-02 // ������� ������� ��������� ��������� � BASIS ���
Private Sub fF63ExchangeItem(inNode, inLocalUTC, inForceIgnoreGate, inForceIgnoreTimeLimits, inForceDayShift, inForceDaysToReport, inForceUncommIgnore) ', inTraderID)
Dim tLogTag, tGTPID, tTimeZoneUTC, tErrorText, tActiveDays, tIndex, tCurrentDay, tFirstDay, tCurrentDateTime, tCalcValue, tSectionList, tIsUpdated, tCalcReady, tCreateNewNode, tXMLChanged
Dim tCurrentReportNode, tTempValue, tNeedToSend, tLogin, tPassword, tGateIgnore, tUncommIgnore, tTraderID
' 00 // ���������������
    tLogTag = fGetLogTag("fF63ExchangeItem")
    
' 01 // ������ ���������� �������� ����
    If Not fReadBasisData(inNode, tTraderID, tGTPID, tTimeZoneUTC, tLogin, tPassword, tGateIgnore, tUncommIgnore, tErrorText) Then
        uCDebugPrint tLogTag, 2, "fReadBasisData > " & tErrorText
        Exit Sub
    End If
    
    If inForceIgnoreGate Then
        uCDebugPrint tLogTag, 1, "������������� ����� ������ ��� ������� �������� ��������! inForceIgnoreGate = True"
        tGateIgnore = True
    End If
    
    If inForceIgnoreTimeLimits Then
        uCDebugPrint tLogTag, 1, "������������� ����� ������ ���������� �������� ��������! inForceIgnoreTimeLimits = True"
    End If
    
    If inForceUncommIgnore Then
        uCDebugPrint tLogTag, 1, "������������� �������������� ���������� ��� ������� ����������� �������� ��������! inForceUncommIgnore = True"
    End If
    
' 02 // ����������� ��������� ����� ��������
    If Not fIsActivePeriod(inNode, inForceIgnoreTimeLimits, tTimeZoneUTC, inLocalUTC, tCurrentDateTime, tErrorText) Then
        If tErrorText <> vbNullString Then: uCDebugPrint tLogTag, 2, tErrorText '����� �����, ���� ��� ������� ����� ���������
        'uCDebugPrint tLogTag, 0, "Out Period"
        Exit Sub
    End If
    uCDebugPrint tLogTag, 0, "������ ������ �� ��� " & tGTPID
    
' 03 // ����������� ���� �� ������� ����� ������ ������ (�� ���������)
    If inForceDayShift = 0 Then
        If Not fWorkDayShiftAdv(CDate(tCurrentDateTime), -1, 0, tFirstDay, tErrorText) Then
            uCDebugPrint tLogTag, 2, tErrorText
            Exit Sub
        End If
        
        tActiveDays = Fix(tCurrentDateTime) - tFirstDay
        
    ElseIf Abs(inForceDayShift) <= 10 Then
        tFirstDay = CDate(Fix(tCurrentDateTime) + inForceDayShift)
        
        tActiveDays = Fix(tCurrentDateTime) - tFirstDay
        If inForceDaysToReport > 0 Then
            If inForceDaysToReport <= tActiveDays Then: tActiveDays = inForceDaysToReport
        End If
        
        uCDebugPrint tLogTag, 1, "������������ ��������� �������� �� [" & Fix(inForceDayShift) & "] ����� ��� ������ �������."
        uCDebugPrint tLogTag, 1, "���� ������ [" & Format(tFirstDay, "YYYY-MM-DD") & "]; ������� ���� [" & Format(Fix(tCurrentDateTime), "YYYY-MM-DD") & "]; ���������� ���� ��� ������ [" & tActiveDays & "]"
    Else
        uCDebugPrint tLogTag, 2, "inForceDayShift (" & inForceDayShift & ") �� ������ ���� ������ �� ��������� ����� �������� [-5:+5]!"
        Exit Sub
    End If
        
    tCurrentDay = tFirstDay
    'uCDebugPrint tLogTag, 0, "tFirstDay=" & Format(tFirstDay, "YYYYMMDD")
    
' 04 // ��������� ������ ������� ��� �������
    If Not fGetSectionList(inNode, tSectionList, cnstElementSplitter, tErrorText) Then
        uCDebugPrint tLogTag, 2, "������ ��������� ������ ������� �� ��� " & tGTPID & ": " & tErrorText
        Exit Sub
    End If
    uCDebugPrint tLogTag, 0, "��������� ������ ������: tActiveDays=" & tActiveDays & "; tSectionList=" & tSectionList & "; tGateIgnore=" & tGateIgnore & "; tUncommIgnore=" & tUncommIgnore
    
' 05 // ��� ������� ��� � ������ ���������� ��������� ������ ������� � �������� ������
    For tIndex = 1 To tActiveDays
' 06 // ������� ������ �������
        If Not fCalculateDayValue(tTraderID, tGTPID, tSectionList, cnstElementSplitter, tCurrentDay, tUncommIgnore, tCalcValue, tIsUpdated, tErrorText) Then
            uCDebugPrint tLogTag, 2, tErrorText
            tCalcReady = False '������ ��������
        Else
            'tCalcValue = Round(tCalcValue, 0) 'v1
            tCalcValue = -tCalcValue 'v2 ����������� ��� ������������� �������� � ��������� ��������� ������������ ������ ������ ������� (����� - ������; ���� - �����)
            uCDebugPrint tLogTag, 0, "CALCVAL=" & tCalcValue & "; ISUPDATED=" & tIsUpdated
            tCalcReady = True '������ ������
        End If
' 07 // �������� �� �� ��������� ������������������ ����� �� ���� ��� �� ���� ����
        If Not fGetCurrentReportF63DB(gF63DB, tTraderID, tGTPID, tCurrentDay, tCurrentReportNode, tErrorText) Then
            uCDebugPrint tLogTag, 2, tErrorText
            Exit Sub
        End If
' 08 // �������� � ������� ������� (���� ����� �� �������� ��� �� ����� ����������, �� ������ ����� report)
        tCreateNewNode = True 'tCalcReady '���� ������ ������ - �� �������������� �� ������ ������� ����� �����, ���� �������� - �������� ������ �� ��������
        tXMLChanged = False
        If Not tCurrentReportNode Is Nothing Then
            tTempValue = tCurrentReportNode.GetAttribute("calcstatus")
            If Not IsNull(tTempValue) Then
                '4 ��������� ��������: 1. X -> X; 2. X -> V; 3. V -> V; 4. V -> X
                '������� 1 (XX) �� ������ ������������ �������� ����� ������, �� ���� ��������� ������� ��� ���������� (���������� ����� ��� 4� ��������)
                If tTempValue = "0" And Not tCalcReady Then
                    tCreateNewNode = False
                '������� 3 (VV) ���������� ������ - ������� ���������� �� ��������� �� ������
                ElseIf tTempValue = "1" And tCalcReady Then
                    tTempValue = tCurrentReportNode.GetAttribute("value")
                    If Not IsNull(tTempValue) Then
                        If IsNumeric(tTempValue) Then
                            tTempValue = CDec(tTempValue) - tCalcValue
                            If tTempValue = 0 Then: tCreateNewNode = False
                        End If
                    End If
                End If
                '������� 2 � 4 ������������ ��� ����������� � �������� ����� ������
            End If
        End If
' 09 // ����� ���� report
        If tCreateNewNode Then
            If Not fCreateNewReportF63DB(gF63DB, tTraderID, tGTPID, tCurrentDay, tCalcValue, tCalcReady, tCurrentReportNode, tErrorText) Then
                fReloadXMLDB gF63DB, False 'RollBack ANY Changes
                uCDebugPrint tLogTag, 2, tErrorText
                Exit Sub
            End If
            tXMLChanged = True
        End If
' 10 // �������� ��������� �������� ������, ��� �������� ������� � ������������� ��������
        If Not tCurrentReportNode Is Nothing Then
            tNeedToSend = False
            tTempValue = tCurrentReportNode.GetAttribute("sent")
            If IsNull(tTempValue) Then
                tNeedToSend = True
            ElseIf tTempValue = vbNullString Then
                tNeedToSend = True
            End If
            If tNeedToSend Then '���� ���������� ���� � �������� - ���������� ��������
                tTempValue = tCurrentReportNode.GetAttribute("calcstatus")
                If Not IsNull(tTempValue) Then
                    If tTempValue = "0" Then: tNeedToSend = False
                End If
            End If
            'Sending?
            If tNeedToSend Then
                uSleep cnstSendDelay 'antispam filter avoid
                If Not fReportSending(tCurrentReportNode, tLogin, tPassword, tGTPID, tCurrentDay, tXMLChanged, tGateIgnore, tErrorText) Then
                    uCDebugPrint tLogTag, 2, tErrorText
                    Exit Sub
                End If
            End If
        End If
        '���������� ������ � ��
        If tXMLChanged Then: fSaveXMLDB gF63DB, False
        '� ���������� ���
        tCurrentDay = tCurrentDay + 1
    Next
    'uCDebugPrint tLogTag, 0, "tWorkDays=" & tWorkDays
    'uCDebugPrint tLogTag, 2, tGTPID & "  " & tTimeZoneUTC

End Sub

Private Function fReportSending(inReportNode, inLogin, inPassword, inGTPID, inDate, inXMLChanged, inGateIgnore, outErrorText)
Dim tEnergyAPI As New CEnergyAPI
Dim tNode, tCalcValue, tSentTry, tLogTag, tIsSendFailed
' 00 // ���������������
    fReportSending = False
    outErrorText = vbNullString
    tLogTag = fGetLogTag("F63REPORTSEND")
' 01 // ���������� ������ �� ���� report
    tCalcValue = inReportNode.GetAttribute("value")
    tSentTry = inReportNode.GetAttribute("senttrycount")
    If IsNumeric(tSentTry) Then
        tSentTry = CDec(tSentTry) + 1
    Else
        tSentTry = 1
    End If
    uCDebugPrint tLogTag, 1, "������� #" & tSentTry & " �������� ������ �� ��� " & inGTPID & " �� ���� " & Format(inDate, "DD.MM.YYYY")
    inXMLChanged = True
    inReportNode.SetAttribute "senttrycount", tSentTry
    inReportNode.SetAttribute "senttry", Format(Now(), "YYYYMMDD hh:mm:ss")
' 02 // �������� ������ ����� API Energy2010
    tIsSendFailed = False
    tEnergyAPI.PrintLog = False
    If tEnergyAPI.IsActive Then
        tEnergyAPI.SetCredentials inLogin, inPassword
        If Not tEnergyAPI.SendRequest(2, inDate, 1, inGTPID & ":" & tCalcValue, inGateIgnore) Then
            uCDebugPrint tLogTag, 2, tEnergyAPI.ErrorText
            tIsSendFailed = True
        Else
            If Not tEnergyAPI.ReadResponse Then
                uCDebugPrint tLogTag, 2, tEnergyAPI.ErrorText
                tIsSendFailed = True
            End If
        End If
    Else
        uCDebugPrint tLogTag, 2, tEnergyAPI.ErrorText
        tIsSendFailed = True
    End If
' 03 // �� ����������� ������ API ������� �������
    If Not tIsSendFailed Then
        inReportNode.SetAttribute "sent", Format(Now(), "YYYYMMDD hh:mm:ss")
        inReportNode.SetAttribute "errortext", vbNullString
        uCDebugPrint tLogTag, 0, "������ ����������!"
    Else
        inReportNode.SetAttribute "errortext", tEnergyAPI.ErrorText
        uCDebugPrint tLogTag, 1, "������ �� ����������!"
    End If
' XX // ����������
    Set tEnergyAPI = Nothing
    fReportSending = True
End Function

Private Function fCreateNewReportF63DB(inXMLDB As TXMLDataBaseFile, inTraderID, inGTPID, inDate, inCalcValue, inCalcStatus, outReportNode, outErrorText)
Dim tXPathString, tRootNode, tDayNode, tNode, tLastReportIndex, tYear, tMonth, tDay
' 00 // ���������������
    fCreateNewReportF63DB = False
    Set outReportNode = Nothing
    outErrorText = vbNullString
' 00 // �������� ������ � ������� XML
    If inXMLDB.XML Is Nothing Then
        outErrorText = "�� <" & inXMLDB.ClassTag & "> �� ������!"
        Exit Function
    End If
    '������ ��� �� ���������� ��������� ������
    Select Case inXMLDB.Version
        Case 1, "1":
            tYear = Format(Year(inDate), "0000")
            tMonth = Format(Month(inDate), "00")
            tDay = Format(Day(inDate), "00")
            tXPathString = "child::trader[@id='" & inTraderID & "']/gtp[@id='" & inGTPID & "']/year[@id='" & tYear & "']/month[@id='" & tMonth & "']/day[@id='" & tDay & "']"
            Set tRootNode = inXMLDB.XML.DocumentElement
            Set tDayNode = tRootNode.SelectSingleNode(tXPathString)
            If tDayNode Is Nothing Then
                'Lv 01 \\ TRADER
                tXPathString = "child::trader[@id='" & inTraderID & "']"
                Set tNode = tRootNode.SelectSingleNode(tXPathString)
                If tNode Is Nothing Then
                    Set tNode = tRootNode.AppendChild(inXMLDB.XML.CreateElement("trader"))
                    tNode.SetAttribute "id", inTraderID
                End If
                Set tRootNode = tNode
                'Lv 02 \\ GTP
                tXPathString = "child::gtp[@id='" & inGTPID & "']"
                Set tNode = tRootNode.SelectSingleNode(tXPathString)
                If tNode Is Nothing Then
                    Set tNode = tRootNode.AppendChild(inXMLDB.XML.CreateElement("gtp"))
                    tNode.SetAttribute "id", inGTPID
                End If
                Set tRootNode = tNode
                'Lv 03 \\ YEAR
                tXPathString = "child::year[@id='" & tYear & "']"
                Set tNode = tRootNode.SelectSingleNode(tXPathString)
                If tNode Is Nothing Then
                    Set tNode = tRootNode.AppendChild(inXMLDB.XML.CreateElement("year"))
                    tNode.SetAttribute "id", tYear
                End If
                Set tRootNode = tNode
                'Lv 04 \\ MONTH
                tXPathString = "child::month[@id='" & tMonth & "']"
                Set tNode = tRootNode.SelectSingleNode(tXPathString)
                If tNode Is Nothing Then
                    Set tNode = tRootNode.AppendChild(inXMLDB.XML.CreateElement("month"))
                    tNode.SetAttribute "id", tMonth
                End If
                Set tRootNode = tNode
                'Lv 05 \\ DAY
                tXPathString = "child::day[@id='" & tDay & "']"
                Set tNode = tRootNode.SelectSingleNode(tXPathString)
                If tNode Is Nothing Then
                    Set tNode = tRootNode.AppendChild(inXMLDB.XML.CreateElement("day"))
                    tNode.SetAttribute "id", tDay
                End If
                Set tDayNode = tNode
            End If
            '������ ������� ���������� ������
            tLastReportIndex = tDayNode.GetAttribute("lastreport")
            If Not IsNull(tLastReportIndex) Then
                If Not IsNumeric(tLastReportIndex) Then
                    outErrorText = "�� ������� ��������� �������� <lastreport>, �� �������� ���������� <" & tLastReportIndex & ">!"
                    Exit Function
                End If
            Else
                tLastReportIndex = 0
            End If
            '�������� ������ ������
            tLastReportIndex = tLastReportIndex + 1 '������ ��������� ���������
            Set tNode = tDayNode.AppendChild(inXMLDB.XML.CreateElement("report"))
            tNode.SetAttribute "id", tLastReportIndex
            tNode.SetAttribute "created", Format(Now(), "YYYYMMDD hh:mm:ss")
            tNode.SetAttribute "sent", vbNullString
            tNode.SetAttribute "value", inCalcValue
            If inCalcStatus Then
                tNode.SetAttribute "calcstatus", "1"
            Else
                tNode.SetAttribute "calcstatus", "0"
            End If
            tNode.SetAttribute "senttry", vbNullString
            tNode.SetAttribute "senttrycount", 0
            tNode.SetAttribute "errortext", vbNullString
            tDayNode.SetAttribute "lastreport", tLastReportIndex
            '�������� ������� ���� � ���������
            Set outReportNode = tNode
        Case Else:
            outErrorText = "������ <" & inXMLDB.Version & "> �� <" & inXMLDB.ClassTag & "> �� ����� �����������! ��������!"
            Exit Function
    End Select
' XX // ����������
    fCreateNewReportF63DB = True
End Function

Private Function fGetCurrentReportF63DB(inXMLDB As TXMLDataBaseFile, inTraderID, inGTPID, inDate, outReportNode, outErrorText)
Dim tXPathString, tRootNode, tDayNode, tLastReportIndex
' 00 // ���������������
    fGetCurrentReportF63DB = False
    Set outReportNode = Nothing
    outErrorText = vbNullString
' 00 // �������� ������ � ������� XML
    If inXMLDB.XML Is Nothing Then
        outErrorText = "�� <" & inXMLDB.ClassTag & "> �� ������!"
        Exit Function
    End If
    '������ ��� �� ���������� ��������� ������
    Select Case inXMLDB.Version
        Case 1, "1":
            tXPathString = "child::trader[@id='" & inTraderID & "']/gtp[@id='" & inGTPID & "']/year[@id='" & Format(Year(inDate), "0000") & "']/month[@id='" & Format(Month(inDate), "00") & "']/day[@id='" & Format(Day(inDate), "00") & "']"
            Set tRootNode = inXMLDB.XML.DocumentElement
            Set tDayNode = tRootNode.SelectSingleNode(tXPathString)
            If Not tDayNode Is Nothing Then
                tLastReportIndex = tDayNode.GetAttribute("lastreport")
                If Not IsNull(tLastReportIndex) Then
                    If IsNumeric(tLastReportIndex) Then
                        tXPathString = "child::report[@id='" & tLastReportIndex & "']"
                        Set outReportNode = tDayNode.SelectSingleNode(tXPathString)
                        If outReportNode Is Nothing Then
                            outErrorText = "���������� �������� � �� <" & inXMLDB.ClassTag & ">! ���� <report> � �������� <" & tLastReportIndex & "> �� �������! XPath <" & tXPathString & ">"
                            Exit Function
                        End If
                    Else
                        outErrorText = "���������� �������� � �� <" & inXMLDB.ClassTag & ">! ������ <" & tLastReportIndex & "> ��� ���� <report> ������ �� ������! XPath <" & tXPathString & ">"
                        Exit Function
                    End If
                End If
            End If
        Case Else:
            outErrorText = "������ <" & inXMLDB.Version & "> �� <" & inXMLDB.ClassTag & "> �� ����� �����������! ��������!"
            Exit Function
    End Select
' XX // ����������
    fGetCurrentReportF63DB = True
End Function

'��������� ������ ������� � ��������(������) ������� �� ����� ���� �������� � ������ ���� ���
Private Function fGetSectionList(inNode, outSectionList, inSplitter, outErrorText)
Dim tXPathString, tNode, tSectionID, tNodes, tSectionListCount, tIndex
Dim tSectionList()
' 00 // ���������������
    fGetSectionList = False
    outErrorText = vbNullString
    outSectionList = vbNullString
' 01 // ������� �� ������������ ���� ���
    tXPathString = "ancestor::gtp"
    Set tNode = inNode.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        outErrorText = "�� ������� �������� ������������ ���� �� XPath <" & tXPathString & ">!"
        Exit Function
    End If
' 02 // ����� �������� ��� �������� �������
    tXPathString = "descendant::section[version[@status='active']]"
    Set tNodes = tNode.SelectNodes(tXPathString)
    If tNodes.Length = 0 Then
        outErrorText = "�� ������� �������� ���� �������� ������ ������� �� XPath <" & tXPathString & ">!"
        Exit Function
    End If
' 03 // ���� ������� � �������� �� ������������ ������ ��� �������
    tSectionListCount = -1
    For Each tNode In tNodes
        tSectionID = tNode.GetAttribute("id")
        If IsNull(tSectionID) Then
            outErrorText = "�� ������� �������� �������� <id> ���� �������!"
            Exit Function
        End If
        '����� SetctionID � ������
        For tIndex = 0 To tSectionListCount
            If tSectionList(tIndex) = tSectionID Then
                outErrorText = "���������� ��������� �������� ������ ������� <" & tSectionID & ">! ��������!"
                Exit Function
            End If
        Next
        '�������� � ������
        tSectionListCount = tSectionListCount + 1
        ReDim Preserve tSectionList(tSectionListCount)
        tSectionList(tSectionListCount) = tSectionID
        '���������� ���������� �������
        If outSectionList = vbNullString Then
            outSectionList = tSectionID
        Else
            outSectionList = outSectionList & cnstElementSplitter & tSectionID
        End If
    Next
' 04 // �������� ��������� ������
    If outSectionList = vbNullString Or tSectionListCount = -1 Then
        outErrorText = "��� ������ �������� ������ ������� �������� ��������, �� ������� ��������� ������!"
        Exit Function
    End If
' XX // ����������
    fGetSectionList = True
End Function

Private Function fGetActiveSectionVersion(inTraderID, inGTPID, inSection)
Dim tXPathString, tNodes, tTempValue, tLogTag
    tLogTag = fGetLogTag("GETACTIVESECTION")
    fGetActiveSectionVersion = 0 'version zero - mean error
    If Not gXMLBasis.Active Then: Exit Function
    
    tXPathString = "//trader[@id='" & inTraderID & "']/gtp[@id='" & inGTPID & "']/section[@id='" & inSection & "']/version[@status='active']"
    Set tNodes = gXMLBasis.XML.SelectNodes(tXPathString)
    If tNodes.Length <> 1 Then '���������� ���������� ��������� �������� �������
        uCDebugPrint tLogTag, 2, "���������� ������ ��������� ������� ������ ���� 1 [������� - " & tNodes.Length & "]; tXPathString=[" & tXPathString & "]"
        Exit Function
    End If
    
    tTempValue = tNodes(0).GetAttribute("id")
    If IsNull(tTempValue) Then '������� ������
        uCDebugPrint tLogTag, 2, "������ ID ������ ��������� �������; tXPathString=[" & tXPathString & "]"
        Exit Function
    End If
    
    uCDebugPrint tLogTag, 0, "������ ��������� �������� ���������� ��� [" & tTempValue & "]; tXPathString=[" & tXPathString & "]"
    fGetActiveSectionVersion = tTempValue
End Function

Private Function fCalculateDayValue(inTraderID, inGTPID, inSectionList, inSplitter, inDate, inUncomIgnore, outValue, outIsUpdated, outErrorText)
Dim tLogTag, tIsUpdated, tSections, tSection, tResultDateStart, tResultDateEnd, tErrorText, tError, tStatusLine, tActiveSectionVersionID
Dim tResult()
' 00 // ���������������
    tLogTag = fGetLogTag("CALCDAYVALUE")
    fCalculateDayValue = False
    outErrorText = vbNullString
    outValue = 0
    outIsUpdated = False
' 01 // ������ ������
    tSections = Split(inSectionList, inSplitter)
    For Each tSection In tSections
        '������� ������ � ������ CALCROUTE
        tActiveSectionVersionID = fGetActiveSectionVersion(inTraderID, inGTPID, tSection)
        'tError = fGetFactCalculation(inTraderID, inGTPID, tSection, tActiveSectionVersionID, "FULL", 0, "T", 0, "d", inDate, inDate, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText)
        tError = fGetFactCalculation(inTraderID, inGTPID, tSection, tActiveSectionVersionID, "FULL", 0, True, "T", 1, "d", inDate, inDate, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText, inUncomIgnore)
        If tError <> 0 Then
            outErrorText = tErrorText
            uCDebugPrint tLogTag, 1, tErrorText
            Exit Function
        End If
        '�������� tStatusLine > �������� 2 �����, � ������� ��������� ��������� ��������� ������� 80020 � 80040, ��� 1 - ������ ���� � �� ���������, 2 - �����-�� ������ ������� ����� ���� �������
        'If Len(tStatusLine) <> 2 Then 'v1
        If Len(tStatusLine) <> 3 Then 'v2
            tErrorText = "#E1# ������ ������������, ��������� ��! [" & tStatusLine & "]"
            outErrorText = tErrorText
            uCDebugPrint tLogTag, 1, tErrorText
            Exit Function
        End If
        'If Left(tStatusLine, 1) <> "1" Then 'v1
        If Left(tStatusLine, 1) <> "0" Then 'v2 problem mark with "1"
            tErrorText = "#E2# ������ ������������, ��������� ��! [" & tStatusLine & "]"
            outErrorText = tErrorText
            uCDebugPrint tLogTag, 1, tErrorText
            Exit Function
        End If
        If Not IsNumeric(tResult(0)) Then 'v2
            tErrorText = "#E3# ������� ��������! RESULT=" & tResult(0)
            outErrorText = tErrorText
            uCDebugPrint tLogTag, 1, tErrorText
            Exit Function
        End If
        outValue = outValue + tResult(0)
        outIsUpdated = outIsUpdated Or tIsUpdated
    Next
' XX // ����������
    fCalculateDayValue = True
End Function

Private Function fIsActivePeriod(inNode, inForceIgnoreTimeLimits, inTimeZoneUTC, inLocalUTC, outCurrentDateTime, outErrorText)
Dim tCurrentTime, tStartTime, tEndTime, tShiftTime
' 00 // ���������������
    fIsActivePeriod = False
    outErrorText = vbNullString
    outCurrentDateTime = 0
    
' 01 // �������� �������� �����
    '�������� �� ��������
    If Not IsNumeric(inTimeZoneUTC) Then
        outErrorText = "�������� �������� ����� ������ ���� �������� ��������� � �������� [-12..+12], � �������� <" & inTimeZoneUTC & ">!"
        Exit Function
    End If
    inTimeZoneUTC = CDec(inTimeZoneUTC)
    '�������� �� ��������� � ���������� ��������
    If Not (Abs(inTimeZoneUTC) <= 12) Then
        outErrorText = "�������� �������� ����� ������ ���� �������� ��������� � �������� [-12..+12], � �������� <" & inTimeZoneUTC & ">!"
        Exit Function
    End If
    
' 02 // ������ ����������� �������
    If Not inForceIgnoreTimeLimits Then
        If Not fExtractTimeFromHourText(inNode.GetAttribute("start"), tStartTime, outErrorText) Then
            outErrorText = "���� F63 �������� @start > " & outErrorText
            Exit Function
        End If
        If Not fExtractTimeFromHourText(inNode.GetAttribute("end"), tEndTime, outErrorText) Then
            outErrorText = "���� F63 �������� @end > " & outErrorText
            Exit Function
        End If
        '������ ������ �������
        If tEndTime < tStartTime Then
            outErrorText = "������ ������, ����� ������ ������� <" & Format(tStartTime, "hhmm") & "> ��������� ����� ��� ��������� <" & Format(tEndTime, "hhmm") & ">!"
            Exit Function
        End If
    End If
    
' 03 // ������� ����� � ��� ���������
    tShiftTime = (-inLocalUTC + inTimeZoneUTC) / 24
    tCurrentTime = Time() + tShiftTime
    outCurrentDateTime = Now() + tShiftTime
    
' 04 // ��������� �������� ������� � ��� ���������� ��������
    'Debug.Print "START=" & Format(tStartTime, "hhmm") & " NOW=" & Format(tCurrentTime, "hhmm") & " END=" & Format(tEndTime, "hhmm")
    If Not inForceIgnoreTimeLimits Then
        If Not (tCurrentTime > tStartTime And tCurrentTime < tEndTime) Then
            Exit Function
        End If
    End If
    
' XX // ����������
    fIsActivePeriod = True
End Function

Private Function fExtractTimeFromHourText(inValue, outValue, outErrorText)
Dim tHours, tMinutes
' 00 // ���������������
    fExtractTimeFromHourText = False
    outErrorText = vbNullString
    outValue = -1
' 01 // ��������
    '������� �� ���������?
    If IsNull(inValue) Then
        outErrorText = "�� ������� ��������� �������� �������� ��� ����������� � ����!"
        Exit Function
    End If
    '����� ��������� � ��������
    If Len(inValue) <> 4 Then
        outErrorText = "������� ��������� ������������� ��������� ��������� <" & inValue & ">, � ������ ���� [����]!"
        Exit Function
    End If
    '�������� �� ��������?
    If Not IsNumeric(inValue) Then
        outErrorText = "������� ��������� ������������� ��������� ��������� <" & inValue & ">, � ������ ���� [����]!"
        Exit Function
    End If
' 02 // ������ ���������
    tHours = CDec(Left(inValue, 2))
    tMinutes = CDec(Right(inValue, 2))
    '�������� ����
    If tHours < 0 Or tHours > 24 Then
        outErrorText = "������� ��������� ������������� ��������� ��������� <" & inValue & ">, � ������ ���� [����] (�� 00-23)!"
        Exit Function
    End If
    '�������� �����
    If tMinutes < 0 Or tMinutes > 59 Then
        outErrorText = "������� ��������� ������������� ��������� ��������� <" & inValue & ">, � ������ ���� [����] (�� 00-59)!"
        Exit Function
    End If
' XX // ����������
    outValue = tHours / 24 + tMinutes / 1440 '�������� � ������ ������� �������
    fExtractTimeFromHourText = True
End Function

Private Function fReadBasisData(inNode, outTraderID, outGTPID, outTimeZoneUTC, outLogin, outPassword, outGateIgnore, outUncommIgnore, outErrorText)
Dim tNode, tXPathString, tTempNode, tValue, tBasisGTPNode
' 00 // ���������������
    fReadBasisData = False
    outErrorText = vbNullString
    outGTPID = vbNullString
    outTimeZoneUTC = -100
' 01 // ������ ������ ��� ������� Energy2010 �� BASIS
    tXPathString = "self::node()"
    If Not fGetAttributeCFG(gXMLBasis, tXPathString, "login", outLogin, tNode, outErrorText, inNode) Then: Exit Function
' 02 // ������ �������� ������������� �����
    If Not fGetAttributeCFG(gXMLBasis, tXPathString, "gateignore", outGateIgnore, tNode, outErrorText, inNode) Then: Exit Function
    If outGateIgnore = "1" Then
        outGateIgnore = True
    Else
        outGateIgnore = False
    End If
' 03 // ������ �������� ������������� �������������� ���������� � �������� �� ������� 80020/80040
    tValue = fGetAttributeCFG(gXMLBasis, tXPathString, "uncommignore", outUncommIgnore, tNode, outErrorText, inNode)
    If outUncommIgnore = "1" Then
        outUncommIgnore = True
    Else
        outUncommIgnore = False
    End If
' 04 // ������ ���� ��� �� BASIS
    tXPathString = "ancestor::gtp"
    If Not fGetAttributeCFG(gXMLBasis, tXPathString, "id", outGTPID, tBasisGTPNode, outErrorText, inNode) Then: Exit Function
' 05 // ������ ���� �������� �� BASIS
    tXPathString = "ancestor::trader"
    If Not fGetAttributeCFG(gXMLBasis, tXPathString, "id", outTraderID, tNode, outErrorText, inNode) Then: Exit Function
' 06 // ������ ������ ��� ������� Energy2010 �� BASIS
    tXPathString = "//trader[@id='" & outTraderID & "']/service[@id='soenergy2010']/item[@login='" & outLogin & "']"
    If Not fGetAttributeCFG(gXMLCredentials, tXPathString, "password", outPassword, tNode, outErrorText) Then: Exit Function
' 07 // ������ ���� ������� ���� ��������� ��� �� BASIS � Dictionary
    ' .01 // ������ ���� ������� �� BASIS
    tXPathString = "child::settings"
    If Not fGetAttributeCFG(gXMLBasis, tXPathString, "subjectid", tValue, tTempNode, outErrorText, tBasisGTPNode) Then: Exit Function
    ' .02 // ������ ���� ������� ����
    tXPathString = "//subjects/subject[@id='" & tValue & "']"
    If Not fGetAttributeCFG(gXMLDictionary, tXPathString, "tradezone", tValue, tTempNode, outErrorText) Then: Exit Function
    ' .03 // ������ ���� �������� ����� ������� ����
    tXPathString = "//tradezones/tradezone[@id='" & tValue & "']"
    If Not fGetAttributeCFG(gXMLDictionary, tXPathString, "timezone", tValue, tTempNode, outErrorText) Then: Exit Function
    ' .04 // ������ �������� �����
    tXPathString = "//timezones/timezone[@id='" & tValue & "']"
    If Not fGetAttributeCFG(gXMLDictionary, tXPathString, "utc", outTimeZoneUTC, tTempNode, outErrorText) Then: Exit Function
' XX // ����������
    fReadBasisData = True
End Function




