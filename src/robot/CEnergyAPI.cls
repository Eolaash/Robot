VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEnergyAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CEnergyAPI v002 13.01.2019
'������ ��� ������� � �������� "�������" ���������� ���������
'v001 - �����
'v002 - ������������� ����� (�����)
'v003 - ���� � �������� ������� �� ��� �������
Option Explicit

Private Const pClassName = "CEnergyAPI"
Private Const pClassVersion = "002"
Private Const pClassDateStamp = "10.01.2019"
Private pWinHTTP
Private pContentType
Private pWebServiceURL
Private pLogin, pPassword
Private pIsObjectActive
Private pErrorText
Private pActionID
Private pSOAPRequest
Private pSOAPResponse
Private pXML
Private pResponse
Private pActiveDateXML, pActiveDate, pDataMode, pDataModeIndex
Private pBlockLimiter, pElementLimiter, pSubLimiter
Private pQuickLog, pPrintLog
Private pMulti

Private Sub fResetRequestParams()
    '����� ������� ������� �������
    pDataMode = vbNullString
    pActiveDateXML = vbNullString
    pActiveDate = 0
    pActionID = -1
    pDataModeIndex = -1
    '���������� ������
    pResponse = vbNullString
    pSOAPResponse = vbNullString
End Sub

Private Sub fLogReset()
    pErrorText = vbNullString
    pQuickLog = vbNullString
End Sub

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
        tLogTag = pClassName & ".INIT"
        fLogReset
        pPrintLog = True
        '�������� ������ ��� ����������
        Set pWinHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
        '� ��������� � ������� ������ ������ � ������ �������� ������ ������������
        pWinHTTP.Option(4) = 13056
        '��������� ��������� (��): ResolveTimeout, ConnectTimeout, SendTimeout, ReceiveTimeout
        'https://docs.microsoft.com/en-us/windows/desktop/winhttp/iwinhttprequest-settimeouts
        pWinHTTP.SetTimeouts 30000, 30000, 30000, 30000
        '��������� ��� �������� ����������
        pContentType = "application/soap+xml;charset=UTF-8"
        '���� ���������� �� ���������
        pWebServiceURL = "https://ws.so-ups.ru:8500/Energy2010/webservices/armservice.asmx"
        '����� � ������ �� ���������
        pLogin = vbNullString
        pPassword = vbNullString
        '����� ���������� ������ �������
        fResetRequestParams
        'XML ������ ��� ������
        Set pXML = CreateObject("Msxml2.DOMDocument.6.0")
        pXML.SetProperty "SelectionNamespaces", "xmlns:x='http://schemas.xmlsoap.org/soap/envelope/' " & "xmlns:m='http://tempuri.org/'" & " xmlns:n='http://www.w3.org/2001/XMLSchema-instance'"
        '�����������
        pBlockLimiter = "^"
        pElementLimiter = ";"
        pSubLimiter = ":"
        '��������� ��� �������� �� ���� � ����
        pMulti = 1000000
        '������ �������������?
        If Err.Number = 0 Then
            pIsObjectActive = True
            fLogAdd tLogTag, 0, "������������� ������ �������! WebServiceURL = [" & pWebServiceURL & "]"
        Else
            fLogAdd tLogTag, 2, Err.Source & ": " & Err.Description
        End If
    On Error GoTo 0
End Sub

Public Sub SetCredentials(inLogin, inPassword)
    pLogin = inLogin
    pPassword = inPassword
End Sub

Private Function fGetChildNode(inRootNode, inChildName)
Dim tChildCount, tIndex, tXML
    Set fGetChildNode = Nothing
    If inRootNode Is Nothing Then: Exit Function
    tChildCount = inRootNode.ChildNodes.Length
    For tIndex = 0 To tChildCount - 1
        If inRootNode.ChildNodes(tIndex).BaseName = inChildName Then
            Set fGetChildNode = inRootNode.ChildNodes(tIndex)
            Exit Function
        End If
    Next
End Function

Private Function fGetResponseClass(outResponseClass, outErrorText)
Dim tNode, tXPathString
    fGetResponseClass = False
    outErrorText = vbNullString
    outResponseClass = vbNullString
    
    'test
    If Not fXMLCheck(pSOAPResponse) Then
        outErrorText = "����� �� �������� XML - ������ ������!"
        Exit Function
    End If
    
    '�������� � pXML ������ ��������� � fXMLCheck
    tXPathString = "/x:Envelope"
    Set tNode = pXML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        outErrorText = "����������� ����� XML ������! �������� ������� ������� �� <*:Envelope>."
        Exit Function
    End If
    
    '������ �� �������
    tXPathString = "//x:Fault/faultstring"
    Set tNode = pXML.SelectSingleNode(tXPathString)
    If Not tNode Is Nothing Then
        outErrorText = "������ ���������� > " & tNode.Text
        Exit Function
    End If
    
    '��������� �������
    tXPathString = "//x:Body"
    Set tNode = pXML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then '��������
        outErrorText = "�� ������� ���� <Body> � ������!"
        Exit Function
    End If
    
    If tNode.ChildNodes.Length <> 1 Then '��������
        outErrorText = "���� <Body> �������� ���������� ���������� �������� ��� (" & tNode.ChildNodes.Length & ")!"
        Exit Function
    End If
    
    '������ ������ SOAPResponse
    outResponseClass = tNode.ChildNodes(0).NodeName
    fGetResponseClass = True
End Function

Private Function fReadResponse_ValidateUser(outErrorText)
Dim tNode, tXPathString
    fReadResponse_ValidateUser = False
        
    tXPathString = "//x:Body/m:ValidateUserResponse/m:ValidateUserResponse/m:IsValid"
    Set tNode = pXML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then '��������
        outErrorText = "����� <ValidateUserResponse> � ������� � ���������!"
        Exit Function
    End If
    
    pResponse = pActionID & pBlockLimiter
    If LCase(tNode.Text) = "true" Then
        pResponse = pResponse & "1"
    Else
        pResponse = pResponse & "0"
    End If
    fReadResponse_ValidateUser = True
End Function

Private Function fReadResponse_SaveData(outErrorText)
Dim tNode, tXPathString
    fReadResponse_SaveData = False
        
    tXPathString = "//x:Body/m:SaveDataResponse"
    Set tNode = pXML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then '��������
        outErrorText = "����� <SaveDataResponse> � ������� � ���������!"
        Exit Function
    End If
    
    pResponse = pActionID & pBlockLimiter
    If tNode.ChildNodes.Length = 0 Then
        pResponse = pResponse & "1"
    Else
        pResponse = pResponse & "0"
    End If
    fReadResponse_SaveData = True
End Function

Private Function fReadResponse_GetData(inDataMode, outErrorText)
Dim tNode, tNodes, tXPathString, tIndex, tTempNode, tTempResponse, tCode, tValue, tLogTag
    fReadResponse_GetData = False
    tLogTag = pClassName & ".READGETDATA"
    tTempResponse = vbNullString
     
    tXPathString = "//x:Body/m:GetDataResponse/m:ArmResponse/m:Data/m:ArmEntity"
    Set tNodes = pXML.SelectNodes(tXPathString)
    If tNodes.Length = 0 Then '��������
        outErrorText = "����� <GetDataResponse> �� �������� (EnlistNodes)!"
        Exit Function
    End If
    tTempResponse = pActionID & pBlockLimiter
    
    tIndex = -1
    'for each node
    For Each tNode In tNodes
        tIndex = tIndex + 1
        If tIndex > 0 Then: tTempResponse = tTempResponse & pElementLimiter
        
        'CODE Extract
        tXPathString = "descendant::m:Code"
        Set tTempNode = tNode.SelectSingleNode(tXPathString)
        If tTempNode Is Nothing Then
            outErrorText = "����� <GetDataResponse> �� �������� (GetCode)!"
            Exit Function
        End If
        tTempResponse = tTempResponse & tTempNode.Text & pSubLimiter
        tCode = tTempNode.Text
        
        'If inDataMode Then
        'End If
        'VALUE Extract
        tXPathString = "descendant::m:Values"
        Set tTempNode = tNode.SelectSingleNode(tXPathString)
        If tTempNode Is Nothing Then
            outErrorText = "����� <GetDataResponse> �� ��� " & tCode & " �� �������� (GetValues)!"
            Exit Function
        End If
        '����� ��������� �� ���� ������
        'Select Case inDataMode
        '    Case 4: tNodeName = "ValueAbsPt"
        '    Case Else: tNodeName = "Value"
        'End Select
        '������
        tValue = vbNullString
        tXPathString = "descendant::m:Value"
        Set tTempNode = tNode.SelectSingleNode(tXPathString)
        If tTempNode Is Nothing Then
            fLogAdd tLogTag, 1, "����� <GetDataResponse> �� ��� " & tCode & " ��� ������ <Value>!"
        Else
            tValue = tTempNode.Text
        End If
        '�����������
        tValue = Replace(tValue, ".", ",")
        If IsNumeric(tValue) Then: tValue = Fix(tValue * pMulti)
        '����������
        tTempResponse = tTempResponse & tValue
    Next
    
    'final
    pResponse = tTempResponse
    fReadResponse_GetData = True
End Function

'������ ������
Public Function ReadResponse()
Dim tActionName, tErrorText, tLogTag
' 00 // ����������
    ReadResponse = False
    fLogReset
    tLogTag = pClassName & ".READRESP"
    pErrorText = vbNullString
    pResponse = vbNullString
    fLogAdd tLogTag, 0, "��������� ������ ������� // pActionID=[" & pActionID & "]; pActiveDate=[" & pActiveDate & "]; pDataMode=[" & pDataMode & "]"
    
' 01 // ������ ������ ������ (� ��� ��������� ��������; �������� � pXML)
    fLogAdd tLogTag, 0, "����������� ������ ������..."
    If Not fGetResponseClass(tActionName, tErrorText) Then
        fLogAdd tLogTag, 2, tErrorText
        Exit Function
    End If
    fLogAdd tLogTag, 0, "��������� ����� ������: " & tActionName
    
' 02 // ��������� ������ ��������� ������
    Select Case tActionName
' 02.A // ����� ValidateUserResponse
        Case "ValidateUserResponse":
            If Not fReadResponse_ValidateUser(tErrorText) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
' 02.B // ����� GetDataResponse
        Case "GetDataResponse":
            If Not fReadResponse_GetData(pDataModeIndex, tErrorText) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
' 02.C // ����� SaveDataResponse
        Case "SaveDataResponse":
            If Not fReadResponse_SaveData(tErrorText) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
' 02.C // ����� ����������
        Case Else:
            fLogAdd tLogTag, 2, "����� [" & tActionName & "] ���������� �����������!"
            Exit Function
    End Select
    
' �� // ����������
    fLogAdd tLogTag, 0, "����� ������ [" & tActionName & "]: " & pResponse
    fLogAdd tLogTag, 0, "��������� ������ �������."
    ReadResponse = True
End Function

Private Function fGetDataModeByIndex(inIndex)
    fGetDataModeByIndex = vbNullString
    pDataModeIndex = -1
    Select Case inIndex
        Case 0, "0":
            fGetDataModeByIndex = "None"
            pDataModeIndex = 0
        Case 1, "1":
            fGetDataModeByIndex = "OperationalData"
            pDataModeIndex = 1
        Case 2, "2":
            fGetDataModeByIndex = "MonthlyPlanning"
            pDataModeIndex = 2
        Case 3, "3":
            fGetDataModeByIndex = "LongPlanning"
            pDataModeIndex = 3
        Case 4, "4":
            fGetDataModeByIndex = "MonthlyData"
            pDataModeIndex = 4
        Case 5, "5":
            fGetDataModeByIndex = "NormativeData"
            pDataModeIndex = 5
        Case Else: fGetDataModeByIndex = "UnknownDataModeIndex(" & inIndex & ")"
    End Select
End Function

Private Function fDateParamCheck(inValue, outValueXML, outValue, outErrorText)
    fDateParamCheck = False
    outValueXML = vbNullString
    outValue = 0
    outErrorText = vbNullString
    If Not IsDate(inValue) Then
        outErrorText = "�������� ���� <" & inValue & "> ������ ���� �����!"
        Exit Function
    End If
    outValue = CDate(Fix(inValue))
    outValueXML = Format(outValue, "YYYY-MM-DDT00:00:00")
    fDateParamCheck = True
End Function

Private Function fDataModeParamCheck(inValue, outValue, outErrorText)
Dim tTempValue
    fDataModeParamCheck = False
    outValue = 0
    outErrorText = vbNullString
    If Not IsNumeric(inValue) Then
        outErrorText = "�������� ������ ������ <" & inValue & "> ������ ���� ������!"
        Exit Function
    End If
    tTempValue = fGetDataModeByIndex(inValue)
    If Not (tTempValue = "OperationalData" Or tTempValue = "MonthlyData") Then 'limiting by 2 modes
        outErrorText = "�������� ������ ������ <" & tTempValue & "> ������ ���� OperationalData(1) ��� MonthlyData(4)!"
        Exit Function
    End If
    outValue = tTempValue
    fDataModeParamCheck = True
End Function

Private Function fDataParamCheck(inValue, outValue, outErrorText)
Dim tElements, tElement, tParams
    fDataParamCheck = False
    outValue = vbNullString
    outErrorText = vbNullString
    If inValue = vbNullString Then
        outErrorText = "�������� ������ <" & inValue & "> ������ �� ������ ������� ������� [���_���1" & pSubLimiter & "�����_����" & pElementLimiter & "���_���2" & pSubLimiter & "�����_����...]!"
        Exit Function
    End If
    tElements = Split(inValue, pElementLimiter)
    For Each tElement In tElements
        If outValue <> vbNullString Then: outValue = outValue & pElementLimiter
        tParams = Split(tElement, pSubLimiter)
        If UBound(tParams) <> 1 Then
            outErrorText = "������� ��������� ������ <" & tElement & "> ������ ���� ������� [���_���" & pSubLimiter & "�����_����]!"
            Exit Function
        End If
        If Len(tParams(0)) <> 8 Then
            outErrorText = "������ ������ ��������� 8 ��������! ������� ��������� ������ <" & tElement & "> ������ ���� ������� [���_���" & pSubLimiter & "�����_����]!"
            Exit Function
        End If
        If Not IsNumeric(tParams(1)) Then
            outErrorText = "�����_���� ������ ���� ������! ������� ��������� ������ <" & tElement & "> ������ ���� ������� [���_���" & pSubLimiter & "�����_����]!"
            Exit Function
        End If
        outValue = outValue & UCase(tParams(0)) & pSubLimiter & Fix(tParams(1))
    Next
    fDataParamCheck = True
End Function

'��������� ��� � pXML ��� �������� ����� �� SOAP GetData
Public Function fExtractDataNode(outDataNode, outErrorText)
Dim tXPathString
    fExtractDataNode = False
    Set outDataNode = Nothing
    outErrorText = vbNullString
    
    tXPathString = "//x:Body/m:GetDataResponse/m:ArmResponse/m:Data"
    Set outDataNode = pXML.SelectSingleNode(tXPathString)
    
    If outDataNode Is Nothing Then
        outErrorText = "�� ������� ������� ���� <Data> �� ������ SOAP GetData!"
        Exit Function
    End If
    
    fExtractDataNode = True
End Function

Private Function fXMLDataTimeConvert(inString, outValue, outErrorText)
Dim tDate, tTime, tElements, tSubElements
    '2019-01-15T10:30:00
    fXMLDataTimeConvert = False
    outValue = 0
    outErrorText = vbNullString
    '�������� #1
    If Len(inString) <> 19 Then
        outErrorText = "������ <" & inString & "> �� �������� dateTime, ����� ������ �� 19!"
        Exit Function
    End If
    tElements = Split(inString, "T")
    If UBound(tElements) <> 1 Then
        outErrorText = "������ <" & inString & "> �� �������� dateTime, ��� ����������� T!"
        Exit Function
    End If
    
    If Len(tElements(0)) <> 10 Then
        outErrorText = "������ <" & inString & "> �� �������� dateTime, ������ [YYYY-MM-DDTHH:mm:SS]!"
        Exit Function
    End If
    
    tSubElements = Split(tElements(0), "-")
    If UBound(tSubElements) <> 2 Then
        outErrorText = "���� ���� <" & inString & "> �� �������� dateTime, ������ [YYYY-MM-DDTHH:mm:SS]!"
        Exit Function
    End If
    
    tDate = CDate(tSubElements(2) & "." & tSubElements(1) & "." & tSubElements(0))
    If Not IsDate(tDate) Then
        outErrorText = "�� ������� ���������� ���� �� <" & inString & "> ������� dateTime [YYYY-MM-DDTHH:mm:SS]!"
        Exit Function
    End If
    
    tSubElements = Split(tElements(1), ":")
    If UBound(tSubElements) <> 2 Then
        outErrorText = "���� ������� <" & inString & "> �� �������� dateTime, ������ [YYYY-MM-DDTHH:mm:SS]!"
        Exit Function
    End If
    
    tTime = CDate(tSubElements(0) & ":" & tSubElements(1) & ":" & tSubElements(2))
    If Not IsDate(tTime) Then
        outErrorText = "�� ������� ���������� ���� �� <" & inString & "> ������� dateTime [YYYY-MM-DDTHH:mm:SS]!"
        Exit Function
    End If
    
    outValue = tDate + tTime
    fXMLDataTimeConvert = True
End Function

'��������� ��� � pXML ��� �������� ����� �� SOAP GetData
Public Function fExtractCloseGate(outCloseGateLocal, outCloseGateMSK, outErrorText)
Dim tXPathString, tNode, tErrorText
    fExtractCloseGate = False
    outCloseGateLocal = 0
    outCloseGateMSK = 0
    outErrorText = vbNullString
    
    'LocalTime
    tXPathString = "//x:Body/m:GetDataResponse/m:ArmResponse/m:CloseGateDateLocal"
    Set tNode = pXML.SelectSingleNode(tXPathString)
    
    If tNode Is Nothing Then
        outErrorText = "�� ������� ������� ���� <CloseGateDateLocal> �� ������ SOAP GetData!"
        Exit Function
    End If
    
    If Not fXMLDataTimeConvert(tNode.Text, outCloseGateLocal, tErrorText) Then
        outErrorText = tErrorText
        Exit Function
    End If
    
    'MSKTime
    tXPathString = "//x:Body/m:GetDataResponse/m:ArmResponse/m:CloseGateDateMsk"
    Set tNode = pXML.SelectSingleNode(tXPathString)
    
    If tNode Is Nothing Then
        outErrorText = "�� ������� ������� ���� <CloseGateDateMsk> �� ������ SOAP GetData!"
        Exit Function
    End If
    
    If Not fXMLDataTimeConvert(tNode.Text, outCloseGateMSK, tErrorText) Then
        outErrorText = tErrorText
        Exit Function
    End If
    
    fExtractCloseGate = True
End Function

Private Function fInjectValuesToDataNode(inDataString, inDataNode, outData, outErrorText)
Dim tXPathString, tArmEntityNode, tXMLText, tNode, tElements, tElement, tParams, tGTPID, tValue
    fInjectValuesToDataNode = False
    
    tXMLText = vbNullString

    If inDataNode Is Nothing Then
        outErrorText = "���� <Data> ��������� �����! �������� ������ ����������!"
        Exit Function
    End If
    
    tElements = Split(inDataString, pElementLimiter)
    For Each tElement In tElements
        tParams = Split(tElement, pSubLimiter)
        '�������� #1: ������������ ������
        If UBound(tParams) <> 1 Then '��������
            outErrorText = "������ ������ ����� �������� ������! �������� [���_���" & pSubLimiter & "�����_����]"
            Exit Function
        End If
        '������� ������ � ���������
        tGTPID = tParams(0)
        tValue = CDec(tParams(1)) / pMulti '������� � ����
        tValue = Replace(CStr(tValue), ",", ".") '������� � �������� ������������� ����� ��� XML
        '����� ���� ��������� ���
        tXPathString = "descendant::m:ArmEntity[m:Code='" & tGTPID & "']"
        Set tArmEntityNode = inDataNode.SelectSingleNode(tXPathString)
        If tArmEntityNode Is Nothing Then
            outErrorText = "��� <" & tGTPID & "> �� ������� � ������ ��������� � ������ ������! �������� �������������� ��� ��� � ���!"
            Exit Function
        End If
        
        '����� ���� ������� ��� ����������
        tXPathString = "descendant::m:Value"
        Set tNode = tArmEntityNode.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            outErrorText = "� ���� <Data> ��� ��� <" & tGTPID & "> �� ������� ���� �������, �������� ����������� ���� <Data>!"
            Exit Function
        End If
        
        '���������� ����������
        If tNode.Attributes.Length = 1 Then
            tNode.RemoveAttribute (tNode.Attributes(0).Name)
        ElseIf tNode.Attributes.Length > 1 Then
            outErrorText = "� ���� <Value> ���� <Data> ��� ��� <" & tGTPID & "> ������� ���������� ����� 1 ����������� <" & tNode.Attributes.Length & ">!"
            Exit Function
        End If
        
        '�������� ������
        tNode.Text = tValue
        
        '���������� ���� � ������ �������� ���������
        tXMLText = tXMLText & tArmEntityNode.XML
    Next
    
    
    outData = tXMLText
    'If outDataNode.Attributes.Length = 1 Then: outDataNode.RemoveAttribute (outDataNode.Attributes(0).Name)
    
    fInjectValuesToDataNode = True
End Function

Private Function fServiceOpenSendingChannel(outErrorText)
    fServiceOpenSendingChannel = False
'01 // ������� �� ������ �����?
    If Not pIsObjectActive Then
        outErrorText = "WinHttp �� ��������!"
        Exit Function
    End If
    
'02 // �������� ���������� �� ������ ����������
    pWinHTTP.Open "POST", pWebServiceURL, False
    If Err.Number <> 0 Then
        outErrorText = Err.Source & ": " & Err.Description
        Exit Function
    End If
            
'03 // �������� ������ � ������ �� ������ ��� ������� � �����������
    pWinHTTP.SetCredentials pLogin, pPassword, 0
    If Err.Number <> 0 Then
        outErrorText = Err.Source & ": " & Err.Description
        Exit Function
    End If
            
'04 // �������� ��������� ��� ����������� ��������
    pWinHTTP.SetRequestHeader "Content-Type", pContentType
    If Err.Number <> 0 Then
        outErrorText = Err.Source & ": " & Err.Description
        Exit Function
    End If
'XX // Over
    fServiceOpenSendingChannel = True
End Function

Private Function fGetCurrentTimeZoneOffset()
    Dim tThisComputer, tWMIService, tWMIItems, tItem
    
    fGetCurrentTimeZoneOffset = vbNullString 'default value
    
    tThisComputer = "." 'local machine
    Set tWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & tThisComputer & "\root\cimv2")

    'extract WMI items by SQL query
    Set tWMIItems = tWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    
    'preventer
    If tWMIItems.Count <> 1 Then: Exit Function

    'extractor
    For Each tItem In tWMIItems
        fGetCurrentTimeZoneOffset = tItem.CurrentTimeZone / 60
        Exit For
    Next
End Function

Private Function fGetLocalUTC(inUTC)
    Dim tLocalUTC

'01 // � ���������� ������ �������� �� ����� � �������� ��� � �������� [-12..+12]
    If IsNumeric(inUTC) Then
        If Abs(Fix(inUTC)) <= 12 Then
            fGetLocalUTC = Fix(inUTC)
            Exit Function
        End If
    End If
    
'02 // � ������ ���� ���������� ����� �� ������
     tLocalUTC = fGetCurrentTimeZoneOffset()
     If IsNumeric(tLocalUTC) Then
        If Abs(Fix(tLocalUTC)) <= 12 Then
            fGetLocalUTC = Fix(tLocalUTC)
            Exit Function
        End If
    End If
    
'03 // ���� ������ �� ���������� �� ���������� UTC+3
    fGetLocalUTC = 3
    
End Function

Public Function SendRequest(inActionID, Optional inDate = vbNullString, Optional inDataMode = vbNullString, Optional inData = vbNullString, Optional inIgnoreGate = False, Optional inLocalUTC = vbNullString)
Dim tDataMode, tDate, tData, tErrorText, tResponseClass, tDataNode, tCurrentDateTime, tDataBlock, tLogTag, tCloseGateLocal, tCloseGateMSK, tLocalUTC, tMSKUTC, tUTCShiftValue
'00 // ����������
    SendRequest = False
    tLogTag = pClassName & ".SENDREQ"
    fLogReset
    fResetRequestParams
    tMSKUTC = 3
    tLocalUTC = fGetLocalUTC(inLocalUTC)
    fLogAdd tLogTag, 0, "����� SendRequest // inActionID=[" & inActionID & "]; inDate=[" & inDate & "]; inDataMode=[" & inDataMode & "]; inData=[" & inData & "]; tLocalUTC=[" & tLocalUTC & "]"
'01 // ��������� ������������� �������
    tDataMode = vbNullString
    tData = vbNullString
    Select Case inActionID
'01.A // ������ �� ������������� ������������ ValidateUser
        Case 0, "0":
            fLogAdd tLogTag, 0, "������ �� ������� ValidateUser(" & inActionID & ")."
            pActionID = 0
            If Not fDateParamCheck(Date, pActiveDateXML, pActiveDate, tErrorText) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
'01.B // ������ �� ������ ������ GetData \\ inParam1 - ������������ ����, inParam2 - ��� ������
        Case 1, "1":
            fLogAdd tLogTag, 0, "������ �� ������� GetData(" & inActionID & ")."
            pActionID = 1
            If Not fDateParamCheck(inDate, pActiveDateXML, pActiveDate, tErrorText) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
            If Not fDataModeParamCheck(inDataMode, tDataMode, tErrorText) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
            pDataMode = tDataMode
'01.C // ������ �� �������� ������ SaveData \\ inParam1 - ������������ ����, inParam2 - ��� ������, inParam3 - ������ � ������� ���1:�����;���2:����� � �����
        Case 2, "2":
            fLogAdd tLogTag, 0, "������ �� ������� SaveData(" & inActionID & ")."
            pActionID = 2
            If Not fDateParamCheck(inDate, pActiveDateXML, pActiveDate, tErrorText) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
            If Not fDataModeParamCheck(inDataMode, tDataMode, tErrorText) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
            pDataMode = tDataMode
            If Not fDataParamCheck(inData, tData, tErrorText) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
'01.D // ����������� ������
        Case Else:
            fLogAdd tLogTag, 2, "������� � �������� <" & inActionID & "> ����������!"
            Exit Function
    End Select
    
    If pDataModeIndex = 4 And Day(pActiveDate) <> 1 Then
        tDate = DateSerial(Year(pActiveDate), Month(pActiveDate), 1)
        If Not fDateParamCheck(tDate, pActiveDateXML, pActiveDate, tErrorText) Then
            fLogAdd tLogTag, 2, tErrorText
            Exit Function
        End If
        fLogAdd tLogTag, 1, "����� ������ <" & pDataMode & "> ������ ����� ������ ���� 01.MM.YYYY. ����������� ������������� ���� [" & inDate & "] >>> [" & pActiveDate & "]!"
    End If
    
'02 // �������� ������� �� ����������
    fLogAdd tLogTag, 0, "������� � ������������ ��������."
    Select Case pActionID
'02.A // ������� ������������������
        Case 0, 1:
            If Not fSendSimpleRequest(tErrorText, pActiveDateXML, pDataMode, tData) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
'02.B // ����������� ������ ��� SaveData
        Case 2:
        ' .01 // ������������ �������; ����� ��� ��� ������� ������ SaveData ���������� �������� ������ � ������ �� ������ GetData
            pActionID = 1 'GetData
            If Not fSendSimpleRequest(tErrorText, pActiveDateXML, pDataMode, tData) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
        ' .02 // ��������� ������
        ' .02.01 // ����������� ������ ������
            fLogAdd tLogTag, 0, "��������� ������ ������� [" & pActionID & "]..."
            If Not fGetResponseClass(tResponseClass, tErrorText) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
        ' .02.02 // �������� �� ����� ������ ��� GetDataResponse
            If Not tResponseClass = "GetDataResponse" Then
                fLogAdd tLogTag, 2, "����������� ����� ��������� ��������� <" & tResponseClass & ">, �������� <GetDataResponse>!"
                Exit Function
            End If
        ' .02.03 // ���������� ������ � �������� ����� ������ ������
            If Not fExtractCloseGate(tCloseGateLocal, tCloseGateMSK, tErrorText) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
        ' .02.04 // ������ �������� ������� � ������� ����� �� ����������� ������ ������
            tUTCShiftValue = (1 / 24) * (tMSKUTC - tLocalUTC) '�������� ������ ��� ����������� ���������� ������� (������� ��� ������� ���� ���) � ����������
            tCurrentDateTime = Now() + tUTCShiftValue
            Debug.Print "������ ������ (�� ���) <" & tCloseGateMSK & ">, � ������ <" & tCurrentDateTime & ">! [������������� ����� = " & inIgnoreGate & "]"
            If tCloseGateMSK < tCurrentDateTime Then
                fLogAdd tLogTag, 2, "������ ������ (�� ���) ��������� <" & tCloseGateMSK & ">, � ������ <" & tCurrentDateTime & ">! [������������� ����� = " & inIgnoreGate & "]"
                If Not inIgnoreGate Then: Exit Function
            End If
        ' .02.05 // ���������� ����� <Data> �� ������ GetDataResponse
            fLogAdd tLogTag, 0, "�������� ������ � ���� <Data> ������� [" & pActionID & "]..."
            If Not fExtractDataNode(tDataNode, tErrorText) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
        ' .02.06 // �������� ����� ������ � ��������������� ����� <ArmEntity> ����� <Data>
            If Not fInjectValuesToDataNode(tData, tDataNode, tDataBlock, tErrorText) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
        ' .03 // ������ ������������, ��������� ������ SaveData
            pActionID = 2 'SaveData
            If Not fSendSimpleRequest(tErrorText, pActiveDateXML, pDataMode, tDataBlock) Then
                fLogAdd tLogTag, 2, tErrorText
                Exit Function
            End If
'02.C // ������� ������� ����������� �������
        Case Else:
            fLogAdd tLogTag, 2, "������� <" & inActionID & "> ����������!"
            Exit Function
    End Select
'XX // ����������
    SendRequest = True
End Function

Private Function fSendSimpleRequest(outErrorText, Optional inDate = vbNullString, Optional inDataMode = vbNullString, Optional inData = vbNullString)
Dim tErrorText, tLogTag
    fSendSimpleRequest = False
    tLogTag = pClassName & ".SIMPLESENDREQ"
    On Error Resume Next
' 01 // �������� ������ ����������
        fLogAdd tLogTag, 0, "�������� ���������� ��� ������� [" & pActionID & "]..."
        If Not fServiceOpenSendingChannel(tErrorText) Then
            outErrorText = tErrorText
            Exit Function
        End If
' 02 // ���������� �������
        fLogAdd tLogTag, 0, "���������� ������� [" & pActionID & "]..."
        If Not fSOAPCreator(tErrorText, inDate, inDataMode, inData) Then
            outErrorText = tErrorText
            Exit Function
        End If
' 03 // �������� �������
        fLogAdd tLogTag, 0, "�������� ������� [" & pActionID & "]..."
        pWinHTTP.Send pSOAPRequest
        If Err.Number <> 0 Then
            outErrorText = Err.Source & ": " & Err.Description
            Exit Function
        End If
        fLogAdd tLogTag, 0, "������ ������� ��������� [" & pActionID & "]."
' 04 // �������� ������
        pWinHTTP.WaitForResponse
' 05 // ���������� ������� � ����������
        fLogAdd tLogTag, 0, "���������� ������ ������� [" & pActionID & "]..."
        pSOAPResponse = pWinHTTP.ResponseText
        If Err.Number <> 0 Then
            outErrorText = Err.Source & ": " & Err.Description
            Exit Function
        End If
        fLogAdd tLogTag, 0, "����� ������� [" & pActionID & "] ��������."
    On Error GoTo 0
    fSendSimpleRequest = True
End Function

Private Sub fAddSOAPString(inText)
    If pSOAPRequest <> vbNullString Then
        pSOAPRequest = pSOAPRequest & vbCrLf & inText
    Else
        pSOAPRequest = inText
    End If
End Sub

Private Function fXMLCheck(inText)
    fXMLCheck = False
    pXML.LoadXML inText
    If pXML.parseError.ErrorCode = 0 Then: fXMLCheck = True
End Function

Private Function fSOAPCreator(outErrorText, Optional inDate = vbNullString, Optional inDataMode = vbNullString, Optional inData = vbNullString)
    fSOAPCreator = False
    outErrorText = vbNullString
    pSOAPRequest = vbNullString
    Select Case pActionID
        '������ "ValidateUser"
        Case 0:
            fAddSOAPString "<?xml version=""1.0"" encoding=""utf-8""?>"
            fAddSOAPString "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
            fAddSOAPString "<soap:Body>"
            fAddSOAPString "<ValidateUser xmlns=""http://tempuri.org/"">"
            fAddSOAPString "<ValidateUserRequest>"
            fAddSOAPString "<Login>" & pLogin & "</Login>"
            fAddSOAPString "<Password>" & pPassword & "</Password>"
            fAddSOAPString "</ValidateUserRequest>"
            fAddSOAPString "</ValidateUser>"
            fAddSOAPString "</soap:Body>"
            fAddSOAPString "</soap:Envelope>"
        '������ "GetData"
        Case 1:
            fAddSOAPString "<?xml version=""1.0"" encoding=""utf-8""?>"
            fAddSOAPString "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
            fAddSOAPString "<soap:Body>"
            fAddSOAPString "<GetData xmlns=""http://tempuri.org/"">"
            fAddSOAPString "<ArmRequest>"
            fAddSOAPString "<Login>" & pLogin & "</Login>"
            fAddSOAPString "<OperationDate>" & inDate & "</OperationDate>"
            fAddSOAPString "<Password>" & pPassword & "</Password>"
            fAddSOAPString "<Variant>" & inDataMode & "</Variant>"
            fAddSOAPString "</ArmRequest>"
            fAddSOAPString "</GetData>"
            fAddSOAPString "</soap:Body>"
            fAddSOAPString "</soap:Envelope>"
        '������ "SaveData"
        Case 2:
            fAddSOAPString "<?xml version=""1.0"" encoding=""utf-8""?>"
            fAddSOAPString "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
            fAddSOAPString "<soap:Body>"
            fAddSOAPString "<SaveData xmlns=""http://tempuri.org/"">"
            fAddSOAPString "<ArmSaveRequest>"
            fAddSOAPString "<Login>" & pLogin & "</Login>"
            fAddSOAPString "<OperationDate>" & inDate & "</OperationDate>"
            fAddSOAPString "<Password>" & pPassword & "</Password>"
            fAddSOAPString "<Variant>" & inDataMode & "</Variant>"
            fAddSOAPString "<Data>"
            fAddSOAPString inData
            fAddSOAPString "</Data>"
            fAddSOAPString "</ArmSaveRequest>"
            fAddSOAPString "</SaveData>"
            fAddSOAPString "</soap:Body>"
            fAddSOAPString "</soap:Envelope>"
    End Select
    '�������� �� XML �������
    If Not fXMLCheck(pSOAPRequest) Then
        outErrorText = "������ ������������ ������� " & pActionID & ". ������� XML ��������. ������� �������������!"
        Exit Function
    End If
    fSOAPCreator = True
End Function

Private Sub Class_Terminate()
Dim tLogTag
    tLogTag = pClassName & ".KILL"
    Set pWinHTTP = Nothing
    Set pXML = Nothing
    pIsObjectActive = False
    fLogAdd tLogTag, 0, "������ ���������."
    fLogReset
End Sub

Public Property Get ErrorText()
    ErrorText = pErrorText
End Property

Public Property Get IsActive()
    IsActive = pIsObjectActive
End Property

Public Property Get ActiveDate()
    ActiveDate = pActiveDateXML
End Property

Public Property Get ResponseSOAPText()
    ResponseSOAPText = pSOAPResponse
End Property

Public Property Get Response()
    Response = pResponse
End Property

Public Property Let PrintLog(inMode)
    If VarType(inMode) = vbBoolean Then: pPrintLog = inMode
End Property
