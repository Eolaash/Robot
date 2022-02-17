VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type TReason
    Reasons() As String
    ReasonsCount As Variant
    Short As String
    IsInternal As Boolean
End Type

Private Type TSource
    FileName As String
    Class As String
    RecievedTimeStamp As String
    ProcessedTimeStamp As String
    Object As String
    Period As String
    Owner As String
End Type

'Private pDate As String
Private pSource As TSource
'Public Object As String
Public Module As String
Public FuncName As String
Public ExchangeNode
Private pReason As TReason
Private pCommand() As String
Private pCommandCount As Variant
Private pDecision As Integer
Private pStatus As String 'text
Private pHasCommands As Boolean
Private pReportList As String
Private pReportListCount As Variant
Private pFilteredList As String
Public ReportToSenderOnly As Boolean
Public SenderAddress As String


Public Sub ClearCommand()
    pCommandCount = -1
    pHasCommands = False
End Sub

Public Sub AddCommand(inCommand)
    If inCommand <> vbnullsting Then
        pHasCommands = True
        pCommandCount = pCommandCount + 1
        ReDim Preserve pCommand(pCommandCount)
        pCommand(pCommandCount) = pSource.Class & ":" & inCommand
    End If
End Sub

Public Function GetCommand(inIndex)
    If Not IsNumeric(inIndex) Then: Exit Function
    If Not (inIndex >= 0 And inIndex <= pCommandCount) Then: Exit Function
    GetCommand = pCommand(inIndex)
End Function

Private Sub Class_Initialize()
    WipeData
End Sub

Public Sub ClearAddressList()
    pFilteredList = vbnullsting
    pReportList = vbNullString
    pReportListCount = -1
End Sub

Private Sub WipeData()
    ReportToSenderOnly = False
    SenderAddress = vbNullString
    pSource.Class = vbNullString
    pSource.FileName = vbNullString
    pSource.Object = vbNullString
    pSource.Period = vbNullString
    pSource.Owner = vbNullString
    Module = vbNullString
    FuncName = vbNullString
    fDecisionAssign -1, True
    ClearCommand
    ClearReason
    ClearAddressList
    Set ExchangeNode = Nothing
End Sub

Private Function fDecisionText(inValue, Optional inLang = 0)
Dim tTempStatus
    tTempStatus = vbNullString
    Select Case inValue
        Case -1:
            If inLang = 0 Then
                tTempStatus = "Неизвестно"
            ElseIf inLang = 1 Then
                tTempStatus = "Unknown"
            End If
        Case 0 To 9:
            If inLang = 0 Then
                tTempStatus = "Принято"
            ElseIf inLang = 1 Then
                tTempStatus = "Accepted"
            End If
        Case 10 To 19:
            If inLang = 0 Then
                tTempStatus = "Отклонено"
            ElseIf inLang = 1 Then
                tTempStatus = "Rejected"
            End If
        Case 20 To 29:
            If inLang = 0 Then
                tTempStatus = "Задержано"
            ElseIf inLang = 1 Then
                tTempStatus = "Suspended"
            End If
        Case 30 To 39:
            If inLang = 0 Then
                tTempStatus = "Передано оператору"
            ElseIf inLang = 1 Then
                tTempStatus = "Manual"
            End If
    End Select
    fDecisionText = tTempStatus
End Function

Private Function fDecisionAssign(inValue, Optional inForceAssign = False)
Dim tTempStatus
    fDecisionAssign = False
    'preventer
    If (pDecision >= inValue) And Not inForceAssign Then: Exit Function
    'selector
    tTempStatus = fDecisionText(inValue)
    'result
    If tTempStatus <> vbNullString Then
        pDecision = inValue
        pStatus = tTempStatus
        fDecisionAssign = True
    End If
End Function

Public Function SetSource(inClass, inFileName, inOwner)
    pSource.Class = inClass
    pSource.FileName = inFileName
    pSource.Owner = inOwner
End Function

Public Function ClearReason()
    With pReason
        .IsInternal = False
        .Short = vbNullString
        .ReasonsCount = -1
    End With
End Function

Public Function AddReason(inText, Optional inShortText = vbNullString)
    With pReason
        .ReasonsCount = .ReasonsCount + 1
        ReDim Preserve .Reasons(.ReasonsCount)
        .Reasons(.ReasonsCount) = inText
        'для коротюльки
        If .ReasonsCount = 0 Then
            If inShortText <> vbNullString Then
                .Short = inShortText
            Else
                .Short = "Какая-то ошибка в функции <" & FuncName & "> модуля <" & Module & ">!"
            End If
        Else
            .Short = "Множественные ошибки функции <" & FuncName & "> модуля <" & Module & ">!"
        End If
    End With
End Function

Public Function GetReason(Optional inMode = 0) As String
Dim tIndex
    Select Case inMode
        Case 0: GetReason = pSource.Object & ": " & pStatus & " (" & pReason.Short & ")"
        Case 1: 'полное описание ошибки
            If pReason.ReasonsCount = 0 Then 'если одна ошибка - отобразим полное её описание
                GetReason = pSource.Object & ": " & pStatus & " (" & pReason.Reasons(0) & ")"
            ElseIf pReason.ReasonsCount < 0 Then 'ошибок нет
                GetReason = pSource.Object & ": " & pStatus & " (Ошибок не определено)"
            Else 'ошибок много
                GetReason = pSource.Object & ": " & pStatus & " (" & pReason.Short & ")"
            End If
        Case 2:
            GetReason = "Источник: " & pSource.FileName & vbCrLf
            GetReason = GetReason & "Получен от: " & pSource.Owner & vbCrLf
            If pSource.RecievedTimeStamp <> vbNullString Then: GetReason = GetReason & "Получен: " & pSource.RecievedTimeStamp & vbCrLf
            If pSource.ProcessedTimeStamp <> vbNullString Then: GetReason = GetReason & "Обработан: " & pSource.ProcessedTimeStamp & vbCrLf
            If pSource.Class <> vbNullString Then: GetReason = GetReason & "Класс: " & pSource.Class & vbCrLf
            If pSource.Object <> vbNullString Then: GetReason = GetReason & "Объект: " & pSource.Object & vbCrLf
            If pSource.Period <> vbNullString Then: GetReason = GetReason & "Период: " & pSource.Period & vbCrLf
            If pReason.IsInternal Then
                GetReason = GetReason & "Статус: " & pStatus & vbCrLf & vbCrLf & "Причина:" & vbCrLf & "Внутренняя ошибка модуля <" & Module & ">!" & vbCrLf & "Обратитесь к разработчику."
            Else
                GetReason = GetReason & "Статус: " & pStatus & vbCrLf & vbCrLf & "Причина:"
                If pReason.ReasonsCount < 0 Then
                    GetReason = GetReason & vbCrLf & "Ошибок не определено"
                Else
                    For tIndex = 0 To pReason.ReasonsCount
                        GetReason = GetReason & vbCrLf & pReason.Reasons(tIndex)
                    Next
                End If
            End If
    End Select
End Function

Public Function GetHeader() As String
    GetHeader = pSource.FileName & ":" & fDecisionText(pDecision, 1)
End Function

Private Function fIsAddressEqual(inAddress, inAddressCheck, inDomainCheck)
Dim tExtractedDomain, tFixDomainCheck, tFixAddressCheck, tFixAddress
    fIsAddressEqual = 0
    tFixAddress = LCase(Trim(inAddress))
    If tFixAddress = vbNullString Then: Exit Function
    tFixAddressCheck = LCase(Trim(inAddressCheck))
    tFixDomainCheck = LCase(Trim(inDomainCheck))
    If tFixAddressCheck <> vbNullString Then
        If tFixAddress = tFixAddressCheck Then: fIsAddressEqual = 1
    ElseIf tFixDomainCheck <> vbNullString Then
        tExtractedDomain = Right(tFixAddress, Len(tFixAddress) - InStrRev(tFixAddress, "@") + 1)
        If Left(tFixDomainCheck, 1) <> "@" Then: tFixDomainCheck = "@" & tFixDomainCheck
        If tExtractedDomain = tFixDomainCheck Then: fIsAddressEqual = 2
    End If
End Function

Private Sub fAddToListUnique(outList, inItem, inAttr, Optional inMainSeparator = ";", Optional inPartSeparator = ":")
Dim tElements, tElement, tParts
    'unique check
    tElements = Split(outList, inMainSeparator)
    For Each tElement In tElements
        tParts = Split(tElement, inPartSeparator)
        If tParts(0) = inItem Then: Exit Sub 'not unique > exit
    Next
    'add unique item to list
    If outList = vbNullString Then
        outList = inItem & inPartSeparator & inAttr
    Else
        outList = outList & inMainSeparator & inItem & inPartSeparator & inAttr
    End If
End Sub

Private Function fGetAttributeAdvanced(inNode, inAttributeName, Optional inDefaulValue = vbNullString)
Dim tValue
    fGetAttributeAdvanced = inDefaulValue
    tValue = inNode.GetAttribute(inAttributeName)
    If IsEmpty(tValue) Or IsNull(tValue) Or tValue = vbNullString Then: Exit Function
    fGetAttributeAdvanced = tValue
End Function

Private Sub fWarnLevelNormalizer(inValue)
    If Not IsNumeric(inValue) Then: inValue = 0
    inValue = CLng(inValue)
    If inValue < 0 Then: inValue = 0
    If inValue > 39 Then: inValue = 39
End Sub

Private Function fGetReportList()
Dim tResultList, tExchangeNode, tSilent, tAddress, tEnabled, tClass, tDomain, tCounter, tItemNode, tElementNode, tWarnLevel
    ClearAddressList
    fGetReportList = pReportListCount
    tResultList = vbNullString
' 01 // Если отчет только отправителю
    If ReportToSenderOnly Then
        fAddToListUnique tResultList, SenderAddress, 0
        pReportListCount = pReportListCount + 1
    Else
        If ExchangeNode Is Nothing Then: Exit Function
' 02 // Поиск в EXCHANGE ноды нужного класса
        For Each tItemNode In ExchangeNode.ChildNodes
            tClass = tItemNode.GetAttribute("id")
            If tClass = pSource.Class Then
' 03 // Перебор записей ноды искомого класса
                For Each tElementNode In tItemNode.ChildNodes
                    tEnabled = tElementNode.GetAttribute("enabled")
                    If tEnabled = "1" Then
                        Select Case LCase(tElementNode.NodeName)
                            'RECIEVE FROM
                            Case "recievefrom":
                                tAddress = tElementNode.GetAttribute("address")
                                tDomain = tElementNode.GetAttribute("domain")
                                If fIsAddressEqual(SenderAddress, tAddress, tDomain) > 0 Then
                                    tSilent = tElementNode.GetAttribute("silent")
                                    tWarnLevel = fGetAttributeAdvanced(tElementNode, "warnlevel", 0)
                                    fWarnLevelNormalizer tWarnLevel
                                    If tSilent = "0" Then
                                        fAddToListUnique tResultList, SenderAddress, tWarnLevel
                                        pReportListCount = pReportListCount + 1
                                    End If
                                End If
                            'REPORT TO
                            Case "reportto":
                                tAddress = tElementNode.GetAttribute("address")
                                tWarnLevel = fGetAttributeAdvanced(tElementNode, "warnlevel", 0)
                                fWarnLevelNormalizer tWarnLevel
                                fAddToListUnique tResultList, tAddress, tWarnLevel
                                pReportListCount = pReportListCount + 1
                        End Select
                    End If
                Next 'class
            End If
        Next 'exchange
    End If
' 04 // Завершение
    pReportList = tResultList
    fGetReportList = pReportListCount
End Function

Public Sub FilterAdressList(Optional inWarnLevelOverride = -1)
Dim tMainSeparator, tPartSeparator, tElements, tElement, tParts, tWarnFilter, tWarnLevel, tAddress
    If fGetReportList >= 0 Then
        tMainSeparator = ";"
        tPartSeparator = ":"
        If inWarnLevelOverride >= 0 Then
            tWarnFilter = inWarnLevelOverride
            fWarnLevelNormalizer tWarFilter
        Else
            tWarnFilter = pDecision
        End If
        tElements = Split(pReportList, tMainSeparator)
        For Each tElement In tElements
            tParts = Split(tElement, tPartSeparator)
            tAddress = tParts(0)
            tWarnLevel = CLng(tParts(1))
            If tWarnFilter >= tWarnLevel Then
                If pFilteredList <> vbNullString Then
                    pFilteredList = pFilteredList & tMainSeparator & tAddress
                Else
                    pFilteredList = tAddress
                End If
            End If
        Next
    End If
End Sub

Public Property Let Decision(inValue)
    fDecisionAssign inValue
End Property

Public Property Get Decision()
    Decision = pDecision
End Property

Public Property Get Status() As String
    Status = pStatus
End Property

Public Property Get CommandCount()
    CommandCount = pCommandCount
End Property

Public Property Get ReasonText() As String
    Reason = pReason.Text
End Property

Public Property Let Object(inText)
    pSource.Object = inText
    'pSource.Object = "1"
End Property

Public Property Get Object()
    Object = pSource.Object
End Property

Public Property Let Period(inText)
    pSource.Period = inText
End Property

Public Property Get Period()
    Period = pSource.Period
End Property

Public Property Let RecievedTimeStamp(inText)
    pSource.RecievedTimeStamp = inText
End Property

Public Property Get RecievedTimeStamp()
    Period = pSource.RecievedTimeStamp
End Property

Public Property Let ProcessedTimeStamp(inText)
    pSource.ProcessedTimeStamp = inText
End Property

Public Property Get ProcessedTimeStamp()
    Period = pSource.ProcessedTimeStamp
End Property

Public Property Let IsInternal(inTrigger)
    pReason.IsInternal = inTrigger
End Property

Public Property Get IsInternal()
    IsInternal = pReason.IsInternal
End Property

Public Property Get HasCommands()
    HasCommands = pHasCommands
End Property

Public Property Get ReportList()
    ReportList = pFilteredList
End Property

Private Sub Class_Terminate()
    Set ExchangeNode = Nothing
End Sub
