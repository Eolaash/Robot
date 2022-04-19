Attribute VB_Name = "Interceptor"
'INTERCEPTOR version 4 (16-08-2018)
Option Explicit

Const cnstModuleName = "INTERCEPTOR"
Const cnstModuleVersion = 4
Const cnstModuleDate = "16-08-2018"

'Const cnstCertSign = "34A59E99C731911FDAD87E3825D689434AB0F3EB" 'ends at 2022-02-05 (haustovya)
Const cnstCertSign = "EB70B04BF0D17307B1E19A1E08D4B8BB687766F1" 'ends at 2023-01-19 (haustovya)
Const cnstCertEncrypt = "DDF9881359B613365A0A256942698E82271C0ABA" 'ends at 2023-02-15

Const cnstDelimiter = ":"
Const cnstMailListDelimiter = ";"
Const cnstMailListParamDelimiter = ":"
Const cnstFastRespond = True
Const clsTagM80020 = "80020"
Const clsTagM30308 = "30308"
Const cnstSendingDealay = 12 'sending delay for item with in a single AIIS area space

Enum EFrameOperations
    ForceZero = 0 '"op_forcezero"
    ForceUncom = 1 '"op_forceuncom"
    ForceCom = 2 ' "op_forcecom"
    IgonreUncom = 3 ' "ignore_uncom"
End Enum

Type TReportAssist
    Reports() As CReport
    ReportsCount As Variant
End Type

Type TM30308DataBlock
    'date
    Date As Date
    Year As String
    Month As String
    Day As String
    'header
    TraderID As String
    GTPID As String
    Version As Variant
    'GTPName As String
    Mode As Variant
    Source As String
    RecievedTime As Variant
    EmailSender As String
    Number As Variant
    'data
    Consume(24) As Variant
    Generate(24) As Variant
    Total(24) As Variant
    'info
    Ready As Boolean
    Comment As String
End Type

Type TReportBuyNoremDataBlock
    'date
    Date As Date
    Year As String
    Month As String
    Day As String
    'header
    TraderID As String
    GTPID As String
    Version As Variant
    FileName As String
    ModificationDate As Variant
    'data
    ValueATS(24) As Variant
    ValueATSCorrection(24) As Variant
    ValueSO(24) As Variant
    ValueAccepted(24) As Variant
    PriceHourAverage(24) As Variant
    ValueGeneration(24) As Variant
    ValueRSV(24) As Variant
    PriceHourRSV(24) As Variant
    'info
    Ready As Boolean
    Comment As String
End Type

Type TSendTimeZone
    Name As String
    Class As Variant
    ID As Integer
    UTC As Variant
    StartDate As Date
    EndDate As Date
    DayLimit As Variant
    Now As Date
    StartDateFormated As String
    EndDateFormated As String
    NowFormated As String
End Type

Type TSenderAreaItem
    'AreaNode As Object
    AreaID As Variant
    SectionID As Variant
    OutNum As Variant
    Active As Boolean
    Class As Variant
    Status As Variant
    Error As Variant
    MailToCount As Variant
    MailToList() As String
    SentCount As Variant
    SentList() As String
    FileName As Variant
End Type

Type TSendTimeZoneList
    TimeZone() As TSendTimeZone
    TimeZoneCount As Variant
End Type

Type TMessageInfo
    SenderEMail As String
    Recieved As Date
End Type

Type TSendAreaItem
    SectionID As Variant
    ID As Variant
    Class As Variant
    TimeZone As Variant
    SendDaysList() As Variant
'    LastDaysList() As Variant
    SendDaysCount As Variant
    ExchangeNode As Object
    MailList As Variant
    SendEnabled As Boolean
End Type

Type TSendGTPItem
    ID As Variant
    Node As Object
    ActiveAreaList() As TSendAreaItem
    ActiveAreaCount As Variant
End Type

Type tPathList
    Root As Variant
    Processed As Variant
    Incoming As Variant
    Done As Variant
End Type

Type TCommonReport
    Date As String
    Owner As String
    RecievedTimeStamp As String
    ProcessedTimeStamp As String
    Source As String
    SourceClass As String
    Object As String
    ReasonInternal As String
    Reason As String
    ReasonShort As String
    Status As String 'text
    Decision As Integer
End Type

Dim gTempDropList As New Collection
Dim gMailCopyList As New Collection
Dim gTempMailFolderName
Dim gCurrentMessage As TMessageInfo
Dim gStackMailFolderName
'=
Dim g30308FolderTag
Dim g80020FolderTag
Dim gTempFolderTag
'=
Dim gLocalInit

Public Sub TestExample()
    'MsgBox "OK"
    'fMainFlow
End Sub

'fGetOpertaionByEnumID - return text command represetation by enum id
Private Function fGetOpertaionByEnumID(inEnumID)
    Select Case inEnumID
        Case 0: fGetOpertaionByEnumID = "op_forcezero"
        Case 1: fGetOpertaionByEnumID = "op_forceuncom"
        Case 2: fGetOpertaionByEnumID = "op_forcecom"
        Case 3: fGetOpertaionByEnumID = "ignore_uncom"
        Case Else: fGetOpertaionByEnumID = vbNullString
    End Select
End Function

Private Function fGetLogTag(inTagText)
    fGetLogTag = cnstModuleName & "." & inTagText
End Function

'INIT \\

'после закрытия outlook переменные сбрасываются; т.е. инициализация при нормальной работе запускается единожды за сессию работы outlook
Private Function fLocalInit(Optional inForceInit As Boolean = False)
Dim tLogTag, tAccount, tMainAccountExists, tPathList, tToAll, tErrorText
Dim tTempFolder As Outlook.Folder
    fLocalInit = False
    tLogTag = "LOCALINI"
    uCDebugPrint tLogTag, 0, "Старт инициализации (принудтельный > " & inForceInit & ")"
    'If Not fConfiguratorInit Then: Exit Function
    If inForceInit Or Not gLocalInit Then
        '=
        gLocalInit = False
        'технические папки для писем
        gTempMailFolderName = "Temp"
        If Not fOutlookFolderCreator(gTempMailFolderName, fGetOutlookRootFolder, tTempFolder) Then
            uCDebugPrint tLogTag, 2, "Не удалось инициализировать папку временного хранения писем <gTempMailFolderName>!"
            Exit Function
        End If
        gStackMailFolderName = "Stack"
        If Not fOutlookFolderCreator(gStackMailFolderName, fGetOutlookRootFolder, tTempFolder) Then
            uCDebugPrint tLogTag, 2, "Не удалось инициализировать папку хранения писем повторной обработки <gStackMailFolderName>!"
            Exit Function
        End If
        Set tTempFolder = Nothing 'clear object
        'местные пути
        g30308FolderTag = "m30308" 'Environ("HOMEPATH") & "\Desktop\Суточные"
        g80020FolderTag = "m80020" '"Z:\temp\80020"
        gTempFolderTag = "temp" 'Environ("TEMP")     'путь до временной локальной папки
        'проверим пути
        If Not fGetStorageListByTag(g30308FolderTag, tPathList, tToAll, tErrorText) Then
            uCDebugPrint tLogTag, 2, tErrorText
            Exit Function
        End If
        If Not fGetStorageListByTag(g80020FolderTag, tPathList, tToAll, tErrorText) Then
            uCDebugPrint tLogTag, 2, tErrorText
            Exit Function
        End If
        If Not fGetStorageListByTag(gTempFolderTag, tPathList, tToAll, tErrorText) Then
            uCDebugPrint tLogTag, 2, tErrorText
            Exit Function
        End If
        'show info
        'uCDebugPrint tLogTag, 0, "gTempFolder > " & gTempFolder
        'uCDebugPrint tLogTag, 0, "gSideDropPath > " & gSideDropPath
        'uCDebugPrint tLogTag, 0, "g30308Path > " & g30308Path
        'check 80020cfg
        If Not (gXML80020CFG.Active) Then
            uCDebugPrint tLogTag, 2, "Не удалось инициировать XML80020CFG!"
            Exit Function
        End If
        'all fine
        gLocalInit = True
    End If
    uCDebugPrint tLogTag, 0, "Инициализация завершена."
    fLocalInit = True
End Function

Public Sub InterceptTest()
Dim tApplication As New Outlook.Application
Dim tExplorer As Outlook.Explorer
Dim tSelection As Outlook.Selection
Dim tMailItem As Outlook.MailItem
Dim tItemIndex, tText, tAcc
    '==
    'MsgBox Application.Session.Accounts.Count
    'For Each tAcc In Application.Session.Accounts
    '    tText = tText & tAcc & vbCrLf
    'Next
    'MsgBox "T1"
    'MsgBox Application.Session.Accounts.Item(1).DisplayName
    'Exit Sub
    uDebugPrint "TST: Start"
    Set tExplorer = tApplication.ActiveExplorer     ' Get the ActiveExplorer.
    Set tSelection = tExplorer.Selection            ' Get the selection.
    uDebugPrint "TST: Selected items - " & tSelection.Count
    'Set oItem = oSel.Item(1)
    For tItemIndex = 1 To tSelection.Count
        Set tMailItem = tSelection.Item(tItemIndex)
        fMailReprocessor tMailItem
    Next
    
    'If (oItem.MessageClass = "IPM.Note") Then
    '    Set oMailItem = oItem
    '    uDebugPrint "TST: Subject - " & oMailItem.Subject
    '    Main oMailItem
    'End If
    
    uDebugPrint "TST: Over"
End Sub

Private Function fDropAttachment(inAttachment, inPath, Optional inClearAfter = False) As String
Dim tDropPath
    On Error Resume Next
        uDebugPrint "DRP: Start"
        fDropAttachment = vbNullString
        tDropPath = inPath & "\" & inAttachment.FileName
        uDeleteFile (tDropPath)
        If uFileExists(tDropPath) Then
            uDebugPrint "DRP: Result > Can't DELETE old file > " & tDropPath
            Exit Function
        End If
        inAttachment.SaveAsFile tDropPath
        If uFileExists(tDropPath) Then
            fDropAttachment = tDropPath
            If inClearAfter Then: gTempDropList.Add tDropPath
        End If
        uDebugPrint "DRP: Result > " & fDropAttachment
    On Error GoTo 0
End Function

Private Sub fTempDropCleaner()
    Do Until gTempDropList.Count = 0
        If uDeleteFile(gTempDropList.Item(1)) Then
            uDebugPrint "CLR: DELETED > " & gTempDropList.Item(1)
            gTempDropList.Remove 1
        End If
    Loop
End Sub

Private Sub fMailCopyListCleaner()
    Set gMailCopyList = New Collection
    'Debug.Print "CLEAR"
    'Do Until gMailCopyList.Count = 0
    '    gMailCopyList.Remove 1
    'Loop
End Sub

Function fIsGTPName(inText) As Boolean
    fIsGTPName = False
    If Len(inText) <> 8 Then: Exit Function
    If UCase(Left(inText, 1)) <> "P" Then: Exit Function
    fIsGTPName = True
End Function

Private Sub fClassificator_XML(inFilePath, outClass, outFileNameStatus, outStructureStatus, outComment)
Dim tXMLDoc, tNode, tValue, tDay, tNumber, tINN, tAIISCode, tElement, tAreaCodeTemp, tBNode, tTempNode, tError
Dim tFileName, tRootNode, tRootNodeName, tAttribute
Dim tLogTag
'00 // Подготовка возвращаемых переменных и базовых значений
    tLogTag = "CLS_XML"
    outClass = vbNullString
    outComment = vbNullString
    outFileNameStatus = -1
    outStructureStatus = -1 'unknown
'01 // Работа с путём файла
    If Not gFSO.FileExists(inFilePath) Then
        outComment = "Нет файла"
        uCDebugPrint tLogTag, 2, "По неизвестным причинам не доступен файл по пути: " & inFilePath
        Exit Sub
    End If
'02 // Загрузка целевого XML для распознания
    Set tXMLDoc = CreateObject("Msxml2.DOMDocument.6.0")
    tXMLDoc.ASync = False
    tXMLDoc.Load inFilePath
'03 // Проверим итоги парсинга
    If tXMLDoc.parseError.ErrorCode = 0 Then 'Parsed?
'04 // Поиск первичных признаков класса
        Set tRootNode = tXMLDoc.DocumentElement
        tRootNodeName = LCase(tRootNode.NodeName)
        Select Case tRootNodeName
            Case "message":
                tValue = UCase(tRootNode.GetAttribute("class"))
                If Not IsNull(tValue) Then
                    If tValue = clsTagM80020 Then
                        outClass = clsTagM80020
                    End If
                End If
        End Select
    Else
        outComment = "Ошибка парсинга"
        uCDebugPrint tLogTag, 2, "Ошибка парсинга " & tXMLDoc.parseError.ErrorCode & " [LINE:" & tXMLDoc.parseError.Line & "/POS:" & tXMLDoc.parseError.LinePos & "]: " & tXMLDoc.parseError.Reason
        Set tXMLDoc = Nothing
        Exit Sub
    End If
'05 // Более глубокий анализ структуры ВЕРОЯТНОГО класса используя СХЕМЫ
    Select Case outClass
        Case clsTagM80020: Set tXMLDoc.Schemas = gXML80020CFG.XSD20V2.XML
    End Select
    'если схема подгружена
    If Not IsNull(tXMLDoc.Schemas) Then
        Set tError = tXMLDoc.Validate()
        If tError.ErrorCode = 0 Then
            outStructureStatus = 0 'normal
        Else
            outStructureStatus = 1 'damaged
            uCDebugPrint tLogTag, 1, "Ошибка по схеме " & tError.ErrorCode & " [LINE:" & tError.Line & "]: " & tError.Reason
        End If
    'XML без схемы по умолчанию считаются правильными (не безопасно, но что поделать)
    Else
        outStructureStatus = 0 'normal
    End If
'06 // Если структура подверждена, то можно анализировать содержимое (если нужно)
    If outStructureStatus = 0 Then
        Select Case outClass
            Case clsTagM80020: outFileNameStatus = fFileNameValidate_XML(outClass, tXMLDoc, inFilePath)
        End Select
    End If
'07 // Выход
    uCDebugPrint tLogTag, 0, "Классификация завершена! > Class=" & outClass & "; NameStatus=" & outFileNameStatus & "; StructureStatus=" & outStructureStatus
    Set tXMLDoc = Nothing
End Sub

Private Function fFileNameValidate_XML(inClass, inXML, inFilePath)
Dim tFileName, tNode, tValue
Dim tNumber, tAreaCodeTemp, tDay, tINN, tAIISCode, tElement
    fFileNameValidate_XML = 0
    tFileName = uGetFileName(inFilePath)
    If (tFileName <> vbNullString) And (InStrRev(tFileName, ".") > 0) Then
        tFileName = Left(tFileName, InStrRev(tFileName, ".") - 1)
    End If
    'CLASS DEFINED SELECTOR
    Select Case inClass
        Case clsTagM80020:
            fFileNameValidate_XML = 1
            tFileName = Split(tFileName, "_")
            If UBound(tFileName) <> 4 Then: Exit Function
            Set tNode = inXML.DocumentElement
            tNumber = tNode.GetAttribute("number")
            If IsNull(tNumber) Then: Exit Function
            tDay = vbNullString
            tINN = vbNullString
            tAIISCode = vbNullString
            Set tNode = inXML.SelectSingleNode("//datetime/day")
            If Not (tNode Is Nothing) Then: tDay = tNode.Text
            Set tNode = inXML.SelectSingleNode("//sender/inn")
            If Not (tNode Is Nothing) Then: tINN = tNode.Text
            Set tNode = inXML.SelectNodes("//area/inn")
            If Not (tNode Is Nothing) Then
                tAreaCodeTemp = vbNullString
                For Each tElement In tNode
                    'P1
                    If (tAIISCode = vbNullString) And (Len(tElement.Text) > 3) Then
                        tAIISCode = Left(tElement.Text, Len(tElement.Text) - 2)
                    End If
                    'P2
                    If tAIISCode <> vbNullString Then
                        tAreaCodeTemp = Left(tElement.Text, Len(tElement.Text) - 2)
                        If tAreaCodeTemp <> tAIISCode Then 'async
                            tAIISCode = vbNullString
                            Exit For
                        End If
                    Else 'not locked
                        tAIISCode = vbNullString
                        Exit For
                    End If
                Next tElement
                If tAIISCode <> vbNullString Then: tAIISCode = tAIISCode & "00"
            End If
            If (clsTagM80020 = tFileName(0)) And (tINN = tFileName(1)) And (tDay = tFileName(2)) And (tNumber = tFileName(3)) And (tAIISCode = tFileName(4)) Then: fFileNameValidate_XML = 0
    End Select
End Function

Private Function fM30308XLSFileNameValidate(inWorkBook, inWorkSheetIndex, outVersion, inFile, outComment)
Dim tFileName, tMode, tDate, tGTPID, tNameSample, tNameSampleFormat, tExtension, tDateText
    fM30308XLSFileNameValidate = -1
    If inWorkBook Is Nothing Then: Exit Function
    With inWorkBook.WorkSheets(inWorkSheetIndex)
        fM30308XLSFileNameValidate = 1
        tFileName = inFile.Name
        tExtension = uGetFileExtension(tFileName)
        If (tFileName <> vbNullString) And (InStrRev(tFileName, ".") > 0) Then
            tFileName = Left(tFileName, InStrRev(tFileName, ".") - 1)
        End If
        tMode = LCase(.Cells(9, 3).Value)
        tDate = .Cells(10, 3).Value
        tGTPID = UCase(.Cells(7, 3).Value)
        tNameSampleFormat = "КодГТП_ГГГГММ"
        tDateText = Format(tDate, "YYYYMM")
        If tMode = "сутки" Then
            tDateText = tDateText & Format(tDate, "DD")
            tNameSampleFormat = tNameSampleFormat & "ДД"
        End If
        tNameSample = tGTPID & "_" & tDateText & "." & tExtension
        tFileName = Split(tFileName, "_")
        If UBound(tFileName) <> 1 Then
            outComment = "Неверный формат имени файла (Ожидаемый формат - " & tNameSampleFormat & ")!"
            Exit Function
        End If
        If tFileName(0) <> tGTPID Then
            outComment = "КодГТП в имени файла <" & tFileName(0) & "> отличается от кода ГТП внутри файла <" & tGTPID & "> (Ожидаемый формат - " & tNameSampleFormat & ")!"
            Exit Function
        End If
        If tFileName(1) <> tDateText Then
            outComment = "Дата в имени файла <" & tFileName(1) & "> отличается от ожидаемой <" & tDateText & "> по данным внутри файла (Ожидаемый формат - " & tNameSampleFormat & ")!"
            Exit Function
        End If
    End With
    outComment = vbNullString
    fM30308XLSFileNameValidate = 0
End Function

Private Function fM30308XLSStructureValidate(inWorkBook, inWorkSheetIndex, outVersion, outDataSet, outComment)
Dim tMarkA, tMarkB, tIndex
    fM30308XLSStructureValidate = -1 'unknown
    outVersion = 0
    outComment = "Не удалось определить стуктуру!"
    If inWorkBook Is Nothing Then: Exit Function
    With inWorkBook.WorkSheets(inWorkSheetIndex)
        'VERSION DETECT
        tMarkA = LCase(.Cells(1, 7).Value)
        tMarkB = LCase(.Cells(2, 7).Value)
        If tMarkA = "заявка" And tMarkB = "ежедневка" Then: outVersion = 1
        If outVersion = 0 Then
            tMarkA = LCase(.Cells(1, 9).Value)
            tMarkB = LCase(.Cells(2, 9).Value)
            If tMarkA = "заявка" And tMarkB = "ежедневка" Then
                tMarkA = LCase(.Cells(3, 9).Value)
                If Len(tMarkA) > 7 Then
                    tMarkA = Right(tMarkA, Len(tMarkA) - 7)
                    If IsNumeric(tMarkA) Then: outVersion = CInt(tMarkA)
                End If
            End If
        End If
        'STRUCTURAL CHECK
        fM30308XLSStructureValidate = 1 'damaged
        'ГТП?
        tMarkA = UCase(.Cells(7, 3).Value)
        If Not fIsGTPName(tMarkA) Then
            outComment = "Ячейка C7 должна содержать код ГТП (ФОРМАТ: PSSSXXX)!"
            Exit Function
        End If
        uAddToList outDataSet, "GTP:" & tMarkA
        'Метки размерности
        tMarkA = LCase(Trim(.Cells(13, 2).Value))
        tMarkB = LCase(Trim(.Cells(39, 2).Value))
        If Not (tMarkA = "время" And tMarkB = "итого:") Then
            outComment = "Ячейка B13 или B39 содержат недопустимые значения (нарушен шаблон)!"
            Exit Function
        End If
        'режим заявки
        tMarkA = LCase(Trim(.Cells(9, 3).Value))
        If Not (tMarkA = "сутки" Or tMarkA = "месяц") Then
            outComment = "Ячейка C9 может принимать значения только <Сутки> или <Месяц>, а содержит <" & tMarkA & ">!"
            Exit Function
        End If
        uAddToList outDataSet, "MODE:" & tMarkA
        'дата?
        tMarkA = .Cells(10, 3).Value
        If Not IsDate(tMarkA) Then
            outComment = "Ячейка C10 может принимать только значения ДАТЫ!"
            Exit Function
        End If
        uAddToList outDataSet, "DATE:" & Fix(CDate(tMarkA))
        'STRUCTURAL BY VERSION CHECK
        Select Case outVersion
            Case 1:
                For tIndex = 0 To 23
                    tMarkA = .Cells(15 + tIndex, 6) 'потребление
                    If Not IsNumeric(tMarkA) Then
                        outComment = "Ячейка C" & 15 + tIndex & " должна содержать цифровое значение (потребление)!"
                        Exit Function
                    End If
                Next
            Case 2 To 5:
                For tIndex = 0 To 23
                    tMarkA = .Cells(15 + tIndex, 8) 'потребление + генерация
                    tMarkB = .Cells(15 + tIndex, 9) 'генерация
                    If Not IsNumeric(tMarkB) Then
                        outComment = "Ячейка D" & 15 + tIndex & " должна содержать цифровое значение (генерация)!"
                        Exit Function
                    End If
                    If Not IsNumeric(tMarkA) Then
                        outComment = "Ячейка E" & 15 + tIndex & " должна содержать цифровое значение (потребление)!"
                        Exit Function
                    End If
                Next
        End Select
    End With
    'NORMAL EXIT
    outComment = vbNullString
    fM30308XLSStructureValidate = 0 'normal
End Function

Private Sub fClassificator_XLS(inFile, outClass, outVersion, outFileNameStatus, outStructureStatus, outDataSet, outComment)
Dim tFileName, tWorkBook, tLogTag, tIndex
Dim tMarks()
Dim tMarksSize
'00 // Подготовка возвращаемых переменных и базовых значений
    tLogTag = "CLS_XLS"
    outVersion = 0
    outClass = vbNullString
    outComment = vbNullString
    outDataSet = vbNullString 'class-specific subdata reads
    outFileNameStatus = -1
    outStructureStatus = -1 'unknown
'01 // Работа с путём файла
    If Not gFSO.FileExists(inFile.Path) Then
        outComment = "Нет файла"
        uCDebugPrint tLogTag, 2, "По неизвестным причинам не доступен файл по пути: " & inFile.Path
        Exit Sub
    End If
'02 // Откроем книгу
    'сделай проверку! gExcel
    Set tWorkBook = gExcel.Workbooks.Open(inFile.Path, False, True)
    If tWorkBook Is Nothing Then
        outComment = "Не удалось открыть книгу"
        uCDebugPrint tLogTag, 2, "По неизвестным не удалось открыть книгу EXCEL: " & inFile.Path
        Exit Sub
    End If
    fExcelControl -1, -1, -1, -1 'disable excel controls
    With tWorkBook
'03 // Определение класса первичное
        tMarksSize = 4
        ReDim Marks(tMarksSize)
    '03.01 // M30308 Check V2..5
        If outClass = vbNullString Then
            Marks(0) = LCase(.WorkSheets(1).Cells(1, 9).Value)
            Marks(1) = LCase(.WorkSheets(1).Cells(2, 9).Value)
            If Marks(0) = "заявка" And Marks(1) = "ежедневка" Then: outClass = "30308"
        End If
    '03.02 // M30308 Check V1
        If outClass = vbNullString Then
            Marks(0) = LCase(.WorkSheets(1).Cells(1, 7).Value)
            Marks(1) = LCase(.WorkSheets(1).Cells(2, 7).Value)
            If Marks(0) = "заявка" And Marks(1) = "ежедневка" Then: outClass = "30308"
        End If
'04 // Проверка структуры класса
        Select Case outClass
            Case "30308": outStructureStatus = fM30308XLSStructureValidate(tWorkBook, 1, outVersion, outDataSet, outComment)
        End Select
'05 // Определение соотвествия имени файла содержимому
        If outStructureStatus = 0 Then
            Select Case outClass
                Case "30308": outFileNameStatus = fM30308XLSFileNameValidate(tWorkBook, 1, outVersion, inFile, outComment)
            End Select
        End If
'04 // Определение класса первичное
        .Close SaveChanges:=False 'silent close book without saving
    End With
    Set tWorkBook = Nothing 'clear object
    fExcelControl 1, 1, 1, 1 'restore excel controls
    uCDebugPrint tLogTag, 0, "Классификация завершена! > Class=" & outClass & "; Version=" & outVersion & "; NameStatus=" & outFileNameStatus & "; StructureStatus=" & outStructureStatus
    ' // over
    'Set tExcel = Nothing
End Sub

Private Sub fXML80020AttachmentReprocess(inAttachmentPath, outCommandString)
Dim tXMLDoc, tResult, tGTPNameList, tGTPName
    Set tXMLDoc = CreateObject("Msxml2.DOMDocument.6.0")
    tXMLDoc.ASync = False
    tXMLDoc.Load inAttachmentPath
    
    tResult = fXML80020Reprocess(gFSO.GetFile(inAttachmentPath), tXMLDoc, tGTPNameList)
    'Debug.Print "AtR1"
    If tResult > 0 Then
        uDebugPrint "XML80020AREP: Areas Extracted > " & tResult
    End If
    'Debug.Print "AtR2"
    outCommandString = tGTPNameList
    
    'If tGTPNameList <> vbNullString Then
    '    tGTPNameList = Split(tGTPNameList, cnstDelimiter)
    '    For Each tGTPName In tGTPNameList
    '        fMailListAdd "80020" & cnstDelimiter & tGTPName
    '    Next
    'End If
    Set tXMLDoc = Nothing
End Sub

Private Sub fAttachmentReprocess_XML(inFilePath, outClass, outCommandString)
Dim tValidStructure, tValidName, tComment
    If gFSO.FileExists(inFilePath) Then
' 01 // Internal check with CLASS resolver
        fClassificator_XML inFilePath, outClass, tValidName, tValidStructure, tComment
' 02 // CLASS-defined attachment processing
        If tValidStructure = 0 Then
            Select Case outClass
                Case clsTagM80020: fXML80020AttachmentReprocess inFilePath, outCommandString
            End Select
            'Debug.Print "AtRXML1"
        End If
    End If
End Sub

' Обработчик почты (запускается ТОЛЬКО после того как сканнер подготовит список писем, т.к. без списка нечего обрабатывать)
Private Sub fMailReprocess(inScanMailAccountName)
Dim tNode, tItemNode, tIsProcessed, tEntryID, tUnknownItem, tXMLChanges, tLogTag, tMailBoxFailture
Dim tMailItem As Outlook.MailItem

    ' 00 // Подготовка
    tXMLChanges = False
    tMailBoxFailture = False
    tLogTag = fGetLogTag("fMailReprocess")
    
    'Debug.Print "F1-" & inScanMailAccountName
    
    ' 01 // Проверим есть ли письма для работы
    If gMailScanDB.Active Then
        Set tNode = gMailScanDB.XML.SelectSingleNode("/message/account[@id='" & inScanMailAccountName & "']")
        If tNode Is Nothing Then: Exit Sub
    Else
        uCDebugPrint tLogTag, 2, gMailScanDB.ClassTag & " не доступна. Отмена обработки писем."
    End If
    
    'Debug.Print "F2"
            
    On Error Resume Next
        For Each tItemNode In tNode.ChildNodes
            tIsProcessed = tItemNode.GetAttribute("processed")
            If tIsProcessed = "0" Then
                tEntryID = tItemNode.GetAttribute("entryid")
                
                'Debug.Print "F3-" & tEntryID
                For Each tUnknownItem In fGetOutlookRootFolder.Items
                    If Not tUnknownItem Is Nothing Then
                        If TypeName(tUnknownItem) = "MailItem" Then
                            If tUnknownItem.EntryID = tEntryID Then
                                Set tMailItem = tUnknownItem
                                'Debug.Print "F4"
                                ' Если ошибок получения письма нет - письмо считается доступным для обработки
                                If Err.Number = 0 Then
                                    ' с 05.10.2020 внесена корректировка - требуется возврат результата работы с письмом (в случае ошибки доступа к ящику - письмо считается необработанным)
                                    If fMailReprocessor(tMailItem, False) Then
                                        'Debug.Print "F5-OK"
                                        tItemNode.SetAttribute "processed", "1"
                                        tXMLChanges = True
                                    End If
                                Else
                                    tMailBoxFailture = True
                                End If
                                
                                Exit For 'в любом случае завершаем скан EntryID
                            End If
                        End If
                    End If
                Next
                
            End If
            
            ' // если был сбой чтения письма - выход
            If tMailBoxFailture Then
                uCDebugPrint tLogTag, 2, "Похоже почтовый ящика <" & inScanMailAccountName & "> не доступен. Отмена дальнейшей обработки писем."
                Exit For
            End If
        Next
    On Error GoTo 0
    
    ' // Внесем корректировки в БД обработки писем
    If tXMLChanges Then
        'fSaveXMLChanges gMailScanDB.XML, gMailScanDB.Path, True
        fSaveXMLDB gMailScanDB, False, , True, , tLogTag & " вносит корректировки!"
    End If
End Sub

Public Function fGetMail()
    Dim tLogTag, tScanMailAccountName, tNewMailCount, tXMLChanges
    
    tLogTag = "fGetMail"
    fGetMail = False
    tScanMailAccountName = fGetMainMailAccountAsString()
    
    ' 01 // local init
    If Not fLocalInit Then: Exit Function
    
    ' 03 // update mailbox items status (trottling if no new items)
    tNewMailCount = fMailScanner(tScanMailAccountName, tXMLChanges)
    'if tNewMailCount = -1 >>> reread scanDB if XMLChanges
    'Debug.Print "X1"
    
    If tXMLChanges Then: fSaveXMLDB gMailScanDB, False, , True, , tLogTag & " вносит корректировки! [fMailScanner]"
    
    'Debug.Print "X2"
    
    If tNewMailCount <= 0 Then: Exit Function
    
    fMailReprocess tScanMailAccountName
End Function

'gMailScanDB
Private Function fMailScanner(inScanMailAccountName, outXMLChanges)
    Dim tNode, tMailNode, tRoot, tLogTag, tIndex, tEntryID, tLock, tValue, tMailAdded, tMailRemoved, tMailToProcess, tMailIndex
    Dim tMailCount
    Dim tUnknownItem As Object
    Dim tErrorText
    Dim tMailFolder As Outlook.MAPIFolder

' 00 // Предопределения
    fMailScanner = -1
    tLogTag = fGetLogTag("fMailScanner")
    tMailAdded = 0
    tMailRemoved = 0
    tMailToProcess = 0
    outXMLChanges = False
    uCDebugPrint tLogTag, 0, "Начало сканирования входящей почты."

' 01 // Предварительная проверка доступности ящика
    Set tMailFolder = fGetMainMailAccount(inScanMailAccountName)
    If tMailFolder Is Nothing Then
        uCDebugPrint tLogTag, 2, "Аварийное завершение. Почтовый ящик <" & inScanMailAccountName & "> оказался недоступен."
        Exit Function
    End If
    
    tMailCount = fGetOutlookRootFolder.Items.Count
    uCDebugPrint tLogTag, 0, "Всего писем в ящике: " & tMailCount
    
' 02 // Поиск текущего основного аккаунта в XML MailScan и создание аккаута если он не найден
    'Debug.Print "1"
    If Not fGetAccountNode(inScanMailAccountName, tNode, tErrorText, outXMLChanges) Then
        uCDebugPrint tLogTag, 2, tErrorText
        uCDebugPrint tLogTag, 2, "Аварийное завершение. Не удалось прочитать блок ACCOUNT файла <" & gMailScanDB.ClassTag & ">."
        Exit Function
    End If
    
    'Debug.Print "2"
' 03 // Все текущие элементы сбрасывают признак синхронизации перед сканирование папки входящей почты
    For Each tMailNode In tNode.ChildNodes
        tMailNode.SetAttribute "sync", "0"
    Next
    
' 05 // Сканирование папки входящих писем; если письмо есть в XML то признак синхронизации "1"; если письма нет в XML - добавим
    
    'Debug.Print "3"
    For tMailIndex = tMailCount To 1 Step -1 'REVERSE MODE - более ранние письма пришли первыми и проверять надо начинать с них (в хронологическом порядке)
        
        ' // субпроверка каждого письма на доступность (при ошибке прекратить работу сканера)
        'Debug.Print "3-1-" & tMailIndex
        Set tUnknownItem = fGetMailItem(tMailFolder, tMailIndex, tErrorText)
        If tUnknownItem Is Nothing Then
            uCDebugPrint tLogTag, 2, "Reading mail item failed! Desc: " & tErrorText
            uCDebugPrint tLogTag, 2, "Аварийное завершение. Почтовый ящик <" & inScanMailAccountName & "> оказался недоступен или же попытка чтения письма закончилась неудачно. Сканер остановлен."
            Exit Function
        End If
            
        ' // обработка письма
        'Debug.Print "3-2-" & tUnknownItem.Subject
        If TypeName(tUnknownItem) = "MailItem" Then
            
            tEntryID = tUnknownItem.EntryID
            
            'поиск письма среди писем XML
            tLock = False
            For Each tMailNode In tNode.ChildNodes
                tValue = tMailNode.GetAttribute("entryid")
                If tValue = tEntryID Then 'если письмо нашлось - восстановим синхронизацию
                    tMailNode.SetAttribute "sync", "1" 'sync complete
                    tLock = True
                    Exit For
                End If
            Next
                
            'новое письмо для XML
            If Not tLock Then
                Set tMailNode = tNode.AppendChild(gMailScanDB.XML.CreateElement("item"))
                tMailNode.SetAttribute "entryid", tEntryID
                tMailNode.SetAttribute "processed", "0"
                tMailNode.SetAttribute "sync", "1" 'sync complete
                tMailAdded = tMailAdded + 1
                outXMLChanges = True
            End If
        End If
        'Debug.Print "3-3"
    Next
    
    'Debug.Print "4"
    
' 04 // Теперь удалим письма в XML для которых не нашлось сихронизации (т.е. письма удалены из папки входящих) ОБХОД с КОНЦА т.к. удаляем по ходу цикла детишек ноды аккаунта
    For tIndex = tNode.ChildNodes.Length - 1 To 0 Step -1
        If tNode.ChildNodes(tIndex).GetAttribute("sync") = "0" Then 'нет синхрованизации - удаляем
            tNode.RemoveChild tNode.ChildNodes(tIndex)
            tMailRemoved = tMailRemoved + 1
            outXMLChanges = True
        ElseIf tNode.ChildNodes(tIndex).GetAttribute("processed") = "0" Then 'подсчет признаков писем "к обработке"
            tMailToProcess = tMailToProcess + 1
        End If
    Next
    'Debug.Print "5"
' 05 // Поиск сообщений в папке сообщений на повторную обработку (STACK)
    'gStackMailFolderName
' 06 // Завершение сканирования; возврат количества писем "к обработке" и сохраниение XML если были изменения (новые письма, удаленные письма)
    fMailScanner = tMailToProcess
    'If tXMLChanges Then: fSaveXMLChanges gMailScanDB.XML, gMailScanDB.Path, True
    uCDebugPrint tLogTag, 0, "Новых писем - " & tMailAdded & "; Удалено писем - " & tMailRemoved & "; Писем к обработке - " & tMailToProcess
    uCDebugPrint tLogTag, 0, "Завершение сканирования входящей почты."
End Function

'safe getter
Private Function fGetMailItem(inFolder As Outlook.MAPIFolder, inIndex, outErrorText) As Object
    outErrorText = vbNullString
    On Error Resume Next
        'Debug.Print "tick"
        Set fGetMailItem = inFolder.Items(inIndex)
        'Debug.Print "tock"
        If Err.Number <> 0 Then
            outErrorText = Err.Description
            Set fGetMailItem = Nothing
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
    On Error GoTo 0
End Function

'gMailScanDB
Private Function fGetAccountNode(inScanMailAccountName, outNode, outErrorText, ioXMLChanges, Optional inAutoCreate = True)
    Dim tRoot, tTempNode
    
    fGetAccountNode = False
    outErrorText = vbNullString
    Set outNode = Nothing
    
    'preventer
    If Not gMailScanDB.Active Then
        outErrorText = "MailDB not active! ClassTag=[" & gMailScanDB.ClassTag & "]"
        Exit Function
    End If
    
    'try to get node
    Set tTempNode = gMailScanDB.XML.SelectSingleNode("/message/account[@id='" & inScanMailAccountName & "']")
    
    'autocreate node if not locked
    If inAutoCreate And tTempNode Is Nothing Then
        Set tRoot = gMailScanDB.XML.DocumentElement
        Set tTempNode = fGetChildNodeByName(tRoot, "account", True)
        If Not tTempNode Is Nothing Then 'preventer
            tTempNode.SetAttribute "id", inScanMailAccountName
        End If
        ioXMLChanges = True
    End If
    
    'last check
    If Not tTempNode Is Nothing Then
        Set outNode = tTempNode
        fGetAccountNode = True
    Else
        outErrorText = "Failed to get node for account [" & inScanMailAccountName & "]! inAutoCreate=[" & inAutoCreate & "]"
    End If
    
    'fin
    Set tRoot = Nothing
End Function


Private Function fGetDataSetItemByTag(inDataSet, inTag)
Dim tDataItems, tDataItem, tDataItemElements, tTag, tValue, tTargetTag
    fGetDataSetItemByTag = vbNullString
    tDataItems = Split(inDataSet, ";")
    tTargetTag = UCase(inTag)
    For Each tDataItem In tDataItems
        tDataItemElements = Split(tDataItem, ":")
        If UBound(tDataItemElements) = 1 Then
            tTag = UCase(tDataItemElements(0))
            tValue = tDataItemElements(1)
            If tTag = tTargetTag Then
                fGetDataSetItemByTag = tValue
                Exit Function
            End If
        End If
    Next
End Function

Private Function fCheckRecieveLegality(inGTPID, inAddress, outReport As CReport)
Dim tXPathString, tValidGTPIDList, tNodes, tNode, tLegalAddress, tLegalDomain, tCurrentGTPID, tLogTag
    fCheckRecieveLegality = False
    tLogTag = "LEGAL30308"
    outReport.FuncName = tLogTag
    uCDebugPrint tLogTag, 0, "Адрес источника <" & inAddress & "> проверяется на легальность.."
    't1
    If Not gXMLBasis.Active Then
        outReport.Decision = 20
        outReport.IsInternal = True
        outReport.AddReason "Конфиг BASIS не доступен!"
        Exit Function
    End If
    't2
    If gXMLBasis.XML Is Nothing Then
        outReport.Decision = 20
        outReport.IsInternal = True
        outReport.AddReason "XML конфига BASIS не доступна!"
        Exit Function
    End If
    
    'get list of available items to recive from inAddress
    tValidGTPIDList = vbNullString
    tXPathString = "//gtp/exchange/item[@id='" & clsTagM30308 & "' and @enabled='1']/recievefrom[@enabled='1']"
    Set tNodes = gXMLBasis.XML.SelectNodes(tXPathString)
    For Each tNode In tNodes
        tLegalAddress = tNode.GetAttribute("address")
        tLegalDomain = tNode.GetAttribute("domain")
        tCurrentGTPID = tNode.ParentNode.ParentNode.ParentNode.GetAttribute("id")
        If tCurrentGTPID <> vbNullString And Not IsNull(tCurrentGTPID) Then
            If fIsAddressEqual(inAddress, tLegalAddress, tLegalDomain) Then: uAddToListUnique tValidGTPIDList, tCurrentGTPID
        End If
    Next
    
    'check is current inGTPID in that list
    If uItemInList(tValidGTPIDList, inGTPID) >= 0 Then
        fCheckRecieveLegality = True
        uCDebugPrint tLogTag, 0, "Адрес источника <" & inAddress & "> допущен к работе."
    Else
        uCDebugPrint tLogTag, 2, "Адрес источника <" & inAddress & "> не допущен к работе."
        outReport.Decision = 10
        If tValidGTPIDList = vbNullString Then
            outReport.AddReason "Адрес <" & inAddress & "> не может быть источником для приёма заявок 30308 потребления!"
        Else
            outReport.AddReason "Неверно заполнено поле ГТП <" & inGTPID & ">, с адреса <" & inAddress & "> могут быть приняты заявки только для следующих ГТП:" & vbCrLf & tValidGTPIDList
        End If
    End If
End Function

Private Function fM30308XLSExtractDataFromFile(inFile, inMessageInfo As TMessageInfo, outDataBlock As TM30308DataBlock, outReport As CReport, inTraderID)
Dim tWorkBook, tListIndex, tValue, tHourIndex, tValueGen, tDataRowShift, tLogTag
    fM30308XLSExtractDataFromFile = False
    tLogTag = "EXTCT30308"
    outReport.FuncName = tLogTag
    outDataBlock.Ready = False
' 01 \\ Открытие книги
    Set tWorkBook = gExcel.Workbooks.Open(inFile.Path, False, True)
    If tWorkBook Is Nothing Then
        outDataBlock.Comment = "Внутренняя ошибка"
        outReport.Decision = 20
        outReport.IsInternal = True
        outReport.AddReason "По неизвестным причинам не удалось открыть книгу EXCEL: " & inFile.Path
        Set tWorkBook = Nothing
        Exit Function
    End If
    fExcelControl -1, -1, -1, -1 'disable excel controls
' 02 \\ Откроем лист #1
    tListIndex = 1
    On Error Resume Next
        With tWorkBook.WorkSheets(tListIndex)
            tValue = .Cells(1, 1).Value 'read check
            If Err.Number = 0 Then
' 03 \\ Чтение данных заголовка
                outDataBlock.Version = .Cells(11, 3).Value 'версия
                If outDataBlock.Version = vbNullString Or Not IsNumeric(outDataBlock.Version) Then
                    outDataBlock.Version = 1
                ElseIf outDataBlock.Version < 1 Then
                    outDataBlock.Version = 1
                End If
                outDataBlock.Date = Fix(CDate(.Cells(10, 3).Value)) 'дата с нормализацией (отсечение часов и т.п.)
                outDataBlock.Day = Format(Day(outDataBlock.Date), "00")
                outDataBlock.Month = Format(Month(outDataBlock.Date), "00")
                outDataBlock.Year = Format(Year(outDataBlock.Date), "0000")
                outDataBlock.Mode = LCase(.Cells(9, 3).Value) 'режим
                'outDataBlock.GTPName = .Cells(8, 3).Value
                outDataBlock.GTPID = UCase(Trim(.Cells(7, 3).Value))
                outDataBlock.TraderID = inTraderID
                outDataBlock.RecievedTime = fGetRecievedTimeStamp(inMessageInfo.Recieved)
                outDataBlock.EmailSender = inMessageInfo.SenderEMail
' 04 \\ Чтение данных потребления и генерации
                outDataBlock.Consume(0) = 0
                outDataBlock.Generate(0) = 0
                outDataBlock.Total(0) = 0
                For tHourIndex = 1 To 24
                    Select Case outDataBlock.Version
                        Case 2 To 5, "2" To "5":
                            tDataRowShift = 14
                            tValue = .Cells(tDataRowShift + tHourIndex, 8).Value
                            tValueGen = .Cells(tDataRowShift + tHourIndex, 9).Value
                        Case Else:
                            tDataRowShift = 14
                            tValue = .Cells(tDataRowShift + tHourIndex, 6).Value
                            tValueGen = 0
                    End Select
                    'check
                    If Not (IsNumeric(tValue) And IsNumeric(tValueGen)) Then
                        outDataBlock.Comment = "Нечисловое значение данных в строке #" & tDataRowShift + tHourIndex
                        outReport.Decision = 10
                        outReport.AddReason outDataBlock.Comment
                        Exit For
                    End If
                    'sum
                    outDataBlock.Consume(0) = outDataBlock.Consume(0) + tValue - tValueGen
                    outDataBlock.Generate(0) = outDataBlock.Generate(0) + tValueGen
                    outDataBlock.Total(0) = outDataBlock.Total(0) + tValue
                    'hour values
                    outDataBlock.Consume(tHourIndex) = tValue - tValueGen
                    outDataBlock.Generate(tHourIndex) = tValueGen
                    outDataBlock.Total(tHourIndex) = tValue
                Next
' 05 \\ Данные прочитаны?
                If tHourIndex = 25 Then
                    If outDataBlock.Mode = "месяц" Then
                        outDataBlock.Mode = 1
                        outDataBlock.Day = "00"
                    Else
                        outDataBlock.Mode = 0
                    End If
                    outDataBlock.Ready = True
                End If
            Else 'ошибка чтения листа
                outDataBlock.Comment = "Внутренняя ошибка"
                outReport.Decision = 20
                outReport.IsInternal = True
                outReport.AddReason "По неизвестным причинам не удалось открыть лист (#" & tListIndex & "): " & inFile.Path
            End If
        End With
    On Error GoTo 0
    tWorkBook.Close SaveChanges:=False 'silent close book without saving
    Set tWorkBook = Nothing 'clear object
    fExcelControl 1, 1, 1, 1 'restore excel controls
    fM30308XLSExtractDataFromFile = outDataBlock.Ready
End Function

Private Function fAddReportAssist(outAReports As TReportAssist)
    With outAReports
        .ReportsCount = .ReportsCount + 1
        ReDim Preserve .Reports(.ReportsCount)
        Set .Reports(.ReportsCount) = New CReport
        Set fAddReportAssist = .Reports(.ReportsCount)
    End With
End Function

Private Function fCurrentTimeOnDutyCorrection(tCurrentTime, tIsCurrentTimeChange)
    Dim tLogTag, tTime, tIndex, tNewCurrentTime
    
    ' 00 // Подготовка
    tIsCurrentTimeChange = False
    tLogTag = "M30308TIMLIM_ONDUTY"
    If Not gXMLCalendar.Active Then
        uCDebugPrint tLogTag, 1, "Календарь недоступен! Пропуск проверки ONDUTY!"
        Exit Function
    End If
    
    ' 01 // Проверка рабочее ли сейчас время и день (именно со стороны сотрудника)
    'uCDebugPrint tLogTag, 1, "TIME=" & tCurrentTime & " TIME2=" & TimeSerial(17, 30, 0) - (tCurrentTime - Fix(tCurrentTime))
    If fIsOnDutyDay(tCurrentTime) Then
        tTime = TimeSerial(17, 30, 0)
        If tTime - (tCurrentTime - Fix(tCurrentTime)) > 0 Then: Exit Function
    End If
    
    ' 02 // Время или день оказались нерабочими (ищем следующий рабочий день)
    tTime = TimeSerial(7, 30, 0)
    For tIndex = 1 To 15
        tNewCurrentTime = Fix(tCurrentTime) + tIndex + tTime
        If fIsOnDutyDay(tNewCurrentTime) Then
            uCDebugPrint tLogTag, 1, "ON_DUTY"
            tIsCurrentTimeChange = True
            tCurrentTime = tNewCurrentTime
            Exit Function
        End If
    Next
End Function

Private Function fM30308TimeLimitsCheck(inGTPID, inTraderID, outSemiMode, inDataBlock As TM30308DataBlock, inPrevDataBlock As TM30308DataBlock, outReport As CReport)
    Dim tSubjectID, tCurrentTime, tRealCurrentTime, tLocalTime, tTimeLimitGEN, tTimeLimitSO, tTimeLimitATS, tOverTimeSO, tOverTimeATS, tLogTag, tComment, tIsGenChange, tIndex, tIsCurrentTimeChange
    Dim tSubjectData As TSubjectInfo
    
    ' 00 // Подготовка
    tLogTag = "M30308TIMLIM"
    fM30308TimeLimitsCheck = False
    outSemiMode = False
    uCDebugPrint tLogTag, 0, "Проверка временных рамок.."
    
    ' 01 // Извлечение SubjectID
    If Not fBasisGetGTPSettings(inGTPID, "subjectid", tSubjectID, tComment) Then
        outReport.Decision = 20
        outReport.IsInternal = True
        outReport.AddReason tComment
        Exit Function
    End If
    
    ' 02 // Извлечение Subject Data
    If Not fDictionaryGetSubjectInfo(tSubjectID, tSubjectData) Then
        outReport.Decision = 20
        outReport.IsInternal = True
        outReport.AddReason tSubjectData.Comment
        Exit Function
    End If
    
    ' 04 // Проверка изменения генерации
    If inPrevDataBlock.Ready Then
        tIsGenChange = False
        For tIndex = 1 To 24
            If inPrevDataBlock.Generate(tIndex) <> inDataBlock.Generate(tIndex) Then
                tIsGenChange = True 'раз есть изменения в генерации значит требуется переподача
                Exit For
            End If
        Next
    Else
        tIsGenChange = True 'если это новая заявка (т.к. нет предыдущей, то требуется подача генерации)
    End If
    
    ' 05 // Расчёт временных лимитов подачи заявок
    tOverTimeSO = False
    tOverTimeATS = False
    tLocalTime = Now()
    tTimeLimitSO = inDataBlock.Date - 1 + TimeSerial(8, 20, 0) '08:20
    tTimeLimitATS = inDataBlock.Date - 1 + TimeSerial(13, 15, 0) '13:15
    tTimeLimitGEN = inDataBlock.Date - 1
    
    tRealCurrentTime = tLocalTime + (tSubjectData.TradeZoneUTC - gLocalUTC) / 24
    
    fCurrentTimeOnDutyCorrection tLocalTime, tIsCurrentTimeChange
    tCurrentTime = tLocalTime + (tSubjectData.TradeZoneUTC - gLocalUTC) / 24
    
    If tIsCurrentTimeChange Then
        tComment = "Произведена автоматическая КОРРЕКЦИЯ ВРЕМЕНИ ближайшей возможной ПОДАЧИ заявки с <" & tRealCurrentTime & " " & fGetUTCForm(tSubjectData.TradeZoneUTC) & "> на <" & tCurrentTime & " " & fGetUTCForm(tSubjectData.TradeZoneUTC) & "> по причине получения в НЕРАБОЧЕЕ ВРЕМЯ оператора!"
        outReport.AddReason tComment & vbCrLf
        uCDebugPrint tLogTag, 1, tComment
    End If
    
    'analyze limits
    If tCurrentTime > tTimeLimitSO Then
        tOverTimeSO = True
        tComment = "Ворота приёма заявки в СО на дату <" & inDataBlock.Date & "> уже закрылись! Закрыто <" & tTimeLimitSO & " " & fGetUTCForm(tSubjectData.TradeZoneUTC) & ">. Сейчас <" & tRealCurrentTime & " " & fGetUTCForm(tSubjectData.TradeZoneUTC) & ">. Внимание! Используемое время это время ценовой зоны, которой принадлежит текущая ГТП."
        outReport.AddReason tComment
        uCDebugPrint tLogTag, 1, tComment
        Select Case tSubjectData.TradeMode
            Case 0:
                'over
                outReport.Decision = 10
                tComment = "Время приёма заявки истекло! Приём заявки невозможен!"
                outReport.AddReason tComment
                uCDebugPrint tLogTag, 2, tComment
                Exit Function
            Case 1:
                If tCurrentTime > tTimeLimitATS Then
                    tOverTimeATS = True
                    tComment = "Ворота приёма заявки в АТС на дату <" & inDataBlock.Date & "> уже закрылись! Закрыто <" & tTimeLimitATS & " " & fGetUTCForm(tSubjectData.TradeZoneUTC) & ">. Сейчас <" & tCurrentTime & " " & fGetUTCForm(tSubjectData.TradeZoneUTC) & ">."
                    outReport.AddReason tComment
                    uCDebugPrint tLogTag, 1, tComment
                    'over
                    outReport.Decision = 10
                    tComment = "Время приёма заявки истекло! Приём заявки невозможен!"
                    outReport.AddReason tComment
                    uCDebugPrint tLogTag, 2, tComment
                    Exit Function
                Else
                    tComment = "Внимание! Заявка получена в промежуточном режиме (приём СО - уже заблокирован, приём АТС - еще активен)."
                    outReport.Decision = 5 'semi mode - reporting
                    outReport.AddReason tComment
                    uCDebugPrint tLogTag, 1, tComment
                    outReport.AddReason "Приём имеет ограничения! Возможно исправление уже принятой СО заявки только в сторону УМЕНЬШЕНИЯ почасовых объемов."
                    outSemiMode = True
                End If
        End Select
        
        'addition
        If tIsCurrentTimeChange Then
            tComment = "Заявка пришла в нерабочее время или выходной день и может быть подана только <" & tCurrentTime & " " & fGetUTCForm(tSubjectData.TradeZoneUTC) & ">. Свяжитесь с оператором."
            outReport.AddReason tComment
            uCDebugPrint tLogTag, 1, tComment
        End If
    End If
    fM30308TimeLimitsCheck = True
    uCDebugPrint tLogTag, 0, "Проверка пройдена.."
End Function

Private Function fGetM30308Node(inXML, inGTPID, inTraderID, inYear, inMonth, inDay, inNodeName)
Dim tXPathString
    Set fGetM30308Node = Nothing
    If inXML Is Nothing Then: Exit Function
    tXPathString = "//trader[@id='" & inTraderID & "']/gtp[@id='" & inGTPID & "']/year[@id='" & inYear & "']/month[@id='" & inMonth & "']/day[@id='" & inDay & "']/" & inNodeName
    Set fGetM30308Node = inXML.SelectSingleNode(tXPathString)
End Function

Private Function fM30308NodeExtract(inNode, inDataBlock As TM30308DataBlock)
Dim tIndex, tNode, tResult, tHour, tTotal, tGenerate, tValue
Dim tCheckList(24)
    fM30308NodeExtract = False
    inDataBlock.Ready = False
    inDataBlock.Comment = "Нет данных"
    If inNode Is Nothing Then: Exit Function 'no record
    If inNode.ChildNodes.Length <> 24 Then
        inDataBlock.Comment = "Данные повреждены! Количество часовых нод (" & inNode.ChildNodes.Length & ") не соотвествует искомому (24)!"
        Exit Function 'corrupted record
    End If
    'precast
    'Source As String
    'RecievedTime As Variant
    'EmailSender As String
    inDataBlock.Day = inNode.ParentNode.GetAttribute("id")
    inDataBlock.Source = inNode.GetAttribute("source")
    inDataBlock.EmailSender = inNode.GetAttribute("email_sender")
    inDataBlock.RecievedTime = inNode.GetAttribute("recieved")
    inDataBlock.Version = inNode.GetAttribute("version")
    inDataBlock.Number = inNode.GetAttribute("number")
    inDataBlock.Month = inNode.ParentNode.ParentNode.GetAttribute("id")
    inDataBlock.Year = inNode.ParentNode.ParentNode.ParentNode.GetAttribute("id")
    inDataBlock.GTPID = inNode.ParentNode.ParentNode.ParentNode.ParentNode.GetAttribute("id")
    inDataBlock.TraderID = inNode.ParentNode.ParentNode.ParentNode.ParentNode.ParentNode.GetAttribute("id") 'wtf
    If inDataBlock.Day = "00" Then
        inDataBlock.Mode = 1
        inDataBlock.Date = DateSerial(inDataBlock.Year, inDataBlock.Month, 1)
    Else
        inDataBlock.Mode = 0
        inDataBlock.Date = DateSerial(inDataBlock.Year, inDataBlock.Month, inDataBlock.Day)
    End If
    For tIndex = 1 To 24
        tCheckList(tIndex) = False
    Next
    'cast
    On Error Resume Next
        For Each tNode In inNode.ChildNodes
            'tCheckList(tIndex) = False
            tHour = tNode.GetAttribute("id")
            tTotal = tNode.GetAttribute("total")
            tGenerate = tNode.GetAttribute("generate")
            If Not (IsNull(tHour) And IsNull(tTotal) And IsNull(tGenerate)) Then
                'normalize local
                tTotal = Replace(tTotal, ".", ",")
                tGenerate = Replace(tGenerate, ".", ",")
                If IsNumeric(tHour) And IsNumeric(tTotal) And IsNumeric(tGenerate) Then
                    tHour = CInt(tHour)
                    tTotal = CDbl(tTotal)
                    tGenerate = CDbl(tGenerate)
                    If Err.Number = 0 Then
                        If (tHour >= 1 And tHour <= 24) And tTotal >= 0 And tGenerate >= 0 Then
                            tCheckList(tHour) = True
                            inDataBlock.Generate(tHour) = tGenerate
                            inDataBlock.Total(tHour) = tTotal
                            inDataBlock.Consume(tHour) = tTotal - tGenerate
                            inDataBlock.Generate(0) = inDataBlock.Generate(0) + inDataBlock.Generate(tHour)
                            inDataBlock.Total(0) = inDataBlock.Total(0) + inDataBlock.Total(tHour)
                            inDataBlock.Consume(0) = inDataBlock.Consume(0) + inDataBlock.Consume(tHour)
                        End If
                    End If
                End If
            End If
            If Err.Number <> 0 Then: Err.Clear
        Next
    On Error GoTo 0
    'postcast
    tResult = True
    For tIndex = 1 To 24
        tResult = tResult And tCheckList(tIndex)
    Next
    'over
    inDataBlock.Ready = tResult
    If inDataBlock.Ready Then
        inDataBlock.Comment = vbNullString
    Else
        inDataBlock.Comment = "Неполные данные! Возможно.. нарушение индексации, отсутствие значений или нецифровые значения!"
    End If
    fM30308NodeExtract = tResult
End Function

Private Function fM30308GetContainerNode(inXML, inTraderCode, inGTPCode, inYear, inMonth, inDay)
    Dim tXPathString, tRootNode, tNode, tContainerNode
    
    tXPathString = "/message/trader[@id='" & inTraderCode & "']/gtp[@id='" & inGTPCode & "']/year[@id='" & inYear & "']/month[@id='" & inMonth & "']/day[@id='" & inDay & "']"
    Set tContainerNode = inXML.SelectSingleNode(tXPathString)
    
    If tContainerNode Is Nothing Then
        'trader
        Set tRootNode = inXML.DocumentElement
        tXPathString = "/message/trader[@id='" & inTraderCode & "']"
        Set tNode = inXML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("trader"))
            tNode.SetAttribute "id", inTraderCode
        End If
        
        'gtp
        Set tRootNode = tNode
        tXPathString = tXPathString & "/gtp[@id='" & inGTPCode & "']"
        Set tNode = inXML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("gtp"))
            tNode.SetAttribute "id", inGTPCode
        End If
        
        'year
        Set tRootNode = tNode
        tXPathString = tXPathString & "/year[@id='" & inYear & "']"
        Set tNode = inXML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("year"))
            tNode.SetAttribute "id", inYear
        End If
        
        'month
        Set tRootNode = tNode
        tXPathString = tXPathString & "/month[@id='" & inMonth & "']"
        Set tNode = inXML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("month"))
            tNode.SetAttribute "id", inMonth
        End If
        
        'day
        Set tRootNode = tNode
        tXPathString = tXPathString & "/day[@id='" & inDay & "']"
        Set tNode = inXML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("day"))
            tNode.SetAttribute "id", inDay
        End If
        
        'container ready
        Set tContainerNode = tNode
    End If
    
    Set fM30308GetContainerNode = tContainerNode
End Function

Private Function fM30308Inject(inXML, tDataBlock As TM30308DataBlock)
Dim tXPathString, tNode, tRootNode, tNodes, tIndex, tOldNode, tContainerNode
    fM30308Inject = False
    If inXML Is Nothing Then: Exit Function
    If Not tDataBlock.Ready Then: Exit Function
    With tDataBlock
    'get container node
        Set tContainerNode = fM30308GetContainerNode(inXML, .TraderID, .GTPID, .Year, .Month, .Day)
        If tContainerNode Is Nothing Then: Exit Function 'logic corrupted
        
    'prepare container - twins check-kill
        'tXPathString = "/message/trader[@id='" & .TraderID & "']/gtp[@id='" & .GTPID & "']/year[@id='" & .Year & "']/month[@id='" & .Month & "']/day[@id='" & .Day & "']/request"
        tXPathString = "child::request"
        Set tNodes = tContainerNode.SelectNodes(tXPathString)
        For Each tNode In tNodes
            Set tOldNode = tNode.ParentNode.RemoveChild(tNode)
        Next
    'inject day-node
        
        Set tNode = tContainerNode.AppendChild(inXML.CreateElement("request"))
        'tNode.setAttribute "id", .Day
        tNode.SetAttribute "number", .Number
        tNode.SetAttribute "version", .Version
        tNode.SetAttribute "source", "email"
        tNode.SetAttribute "recieved", .RecievedTime
        tNode.SetAttribute "email_sender", .EmailSender
        Set tRootNode = tNode
    'hour injection
        For tIndex = 1 To 24
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("hour"))
            tNode.SetAttribute "id", Format(tIndex, "00")
            tNode.SetAttribute "total", .Total(tIndex)
            tNode.SetAttribute "generate", .Generate(tIndex)
        Next
    End With
    fM30308Inject = True
End Function

Private Function fM30308RSVInject(inXML, tDataBlock As TReportBuyNoremDataBlock)
Dim tXPathString, tNode, tRootNode, tNodes, tIndex, tOldNode, tContainerNode
    fM30308RSVInject = False
    If inXML Is Nothing Then: Exit Function
    If Not tDataBlock.Ready Then: Exit Function
    With tDataBlock
    'get container node
        Set tContainerNode = fM30308GetContainerNode(inXML, .TraderID, .GTPID, .Year, .Month, .Day)
        If tContainerNode Is Nothing Then: Exit Function 'logic corrupted
        
    'prepare container - twins check-kill
        'tXPathString = "/message/trader[@id='" & .TraderID & "']/gtp[@id='" & .GTPID & "']/year[@id='" & .Year & "']/month[@id='" & .Month & "']/day[@id='" & .Day & "']/request"
        tXPathString = "child::trade"
        Set tNodes = tContainerNode.SelectNodes(tXPathString)
        For Each tNode In tNodes
            Set tOldNode = tNode.ParentNode.RemoveChild(tNode)
        Next
    'inject day-node
        
        Set tNode = tContainerNode.AppendChild(inXML.CreateElement("trade"))
        tNode.SetAttribute "filename", .FileName
        tNode.SetAttribute "version", .Version
        tNode.SetAttribute "lastmodified", .ModificationDate
        Set tRootNode = tNode
    'hour injection
        For tIndex = 1 To 24
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("hour"))
            tNode.SetAttribute "id", Format(tIndex, "00")
            'main
            tNode.SetAttribute "rsvvalue", .ValueRSV(tIndex)
            tNode.SetAttribute "rsvprice", .PriceHourRSV(tIndex)
            'info
            tNode.SetAttribute "atsvalue", .ValueATS(tIndex)
            tNode.SetAttribute "atscorvalue", .ValueATSCorrection(tIndex)
            tNode.SetAttribute "sovalue", .ValueSO(tIndex)
            tNode.SetAttribute "acceptedvalue", .ValueAccepted(tIndex)
            tNode.SetAttribute "generationvalue", .ValueGeneration(tIndex)
            tNode.SetAttribute "avgprice", .PriceHourAverage(tIndex)
        Next
    End With
    fM30308RSVInject = True
End Function

Private Function fGetRecievedTimeStamp(inDate, Optional inMode = 0, Optional inLocalUTC = 0, Optional inTargetUTC = 0)
Dim tValueDate, tResultString, tHourShift
    tValueDate = inDate
    If Not (IsDate(inDate)) Then: tValueDate = Now()
    Select Case inMode
        Case 1: fGetRecievedTimeStamp = Format(tValueDate, "YYYY-MM-DD Hh:Nn:Ss")
        Case 2:
            tHourShift = inTargetUTC - inLocalUTC
            tResultString = Format(tValueDate + (tHourShift) / 24, "YYYY-MM-DD Hh:Nn:Ss")
            tResultString = tResultString & " UTC"
            If inTargetUTC >= 0 Then: tResultString = tResultString & "+"
            fGetRecievedTimeStamp = tResultString & inTargetUTC
        Case Else: fGetRecievedTimeStamp = Format(tValueDate, "YYYYMMDDHhNnSs")
    End Select
End Function

Private Sub fM30308XLSAttachmentReprocess(inFile, inVersion, outAReports As TReportAssist, inValidStructure, inValidName, inDataSet, inComment)
Dim tNode, tValue, tGTPID, tWorkBook, tFileName, tXPathString, tReportList, tDate, tMode, tComment, tSemiAcceptMode, tRecieved, tKillNode, tOldNode, tNumber, tTraderID, tM30308Node
Dim tTempVar, tDropFolder, tResultCollector, tResultElement, tPathList, tToAll, tErrorText, tDropTriggered
Dim tGraphPicturePath, tRetroComment
Dim tPrevDataBlock As TM30308DataBlock
Dim tDataBlock As TM30308DataBlock
Dim tReport As CReport
Dim tLogTag
    'outCommandString = vbNullString
    tLogTag = fGetLogTag("M30308REP")
    tRecieved = fGetRecievedTimeStamp(gCurrentMessage.Recieved)
    tTraderID = gTraderInfo.ID
    ' \\ REPORT Init
    Set tReport = fAddReportAssist(outAReports) 'new report
    'outAReports.Reports (outAReports.ReportsCount)
    tReport.Module = cnstModuleName
    tReport.SetSource clsTagM30308, inFile.Name, gCurrentMessage.SenderEMail
'01 // Предварительная обработка проверок состояния объекта информации
    tGTPID = UCase(fGetDataSetItemByTag(inDataSet, "GTP"))
    tDate = fGetDataSetItemByTag(inDataSet, "DATE")
    tMode = fGetDataSetItemByTag(inDataSet, "MODE")
    tReport.Object = "Заявка " & tMode & " " & tGTPID
    tReport.RecievedTimeStamp = fGetRecievedTimeStamp(gCurrentMessage.Recieved, 2, gLocalUTC, 3)
    tReport.ProcessedTimeStamp = fGetRecievedTimeStamp(Now(), 2, gLocalUTC, 3)
    tReport.Period = tDate
    tReport.Decision = 0 'accept by default
    tReport.SenderAddress = gCurrentMessage.SenderEMail
    If Not gXML30308DB.Active Then
        tReport.Decision = 20
        tReport.IsInternal = True
        tComment = "База данных для заявок <gXML30308DB> не готова!"
        uCDebugPrint tLogTag, 2, tComment
        tReport.AddReason tComment
        Exit Sub
    End If
'02 // Проверка на легальность источника данных
    'uCDebugPrint tLogTag, 0, "Адрес источника <" & gCurrentMessage.SenderEMail & "> проверяется на легальность.."
    If Not fCheckRecieveLegality(tGTPID, gCurrentMessage.SenderEMail, tReport) Then
        uCDebugPrint tLogTag, 2, tReport.GetReason(1)
        tReport.ReportToSenderOnly = True 'только отправителю
        Exit Sub
    End If
    'uCDebugPrint tLogTag, 0, "Адрес источника <" & gCurrentMessage.SenderEMail & "> допущен к работе."
'03 // Извлечение ноды EXCHANGE
    If tGTPID <> vbNullString Then
        tXPathString = "//trader[@id='" & tTraderID & "']/gtp[@id='" & tGTPID & "']/exchange"
        Set tReport.ExchangeNode = gXMLBasis.XML.SelectSingleNode(tXPathString)
    End If
'04 // Проверка на структурное сооствествие
    If inValidStructure <> 0 Or inValidName <> 0 Then
        tReport.Decision = 10
        tReport.AddReason inComment
        Exit Sub
    End If
'05 // Извлечем данные
    If Not fM30308XLSExtractDataFromFile(inFile, gCurrentMessage, tDataBlock, tReport, tTraderID) Then
        uCDebugPrint tLogTag, 2, tReport.GetReason(1)
        'fFastRespond gCurrentMessage.SenderEMail, tReport.Source & ":Rejected", fReportExpose(tReport, 2)
        Exit Sub
    End If
    
'06 // Извлечение ноды записи заявки из БД на данное время по объекту
    Set tM30308Node = fGetM30308Node(gXML30308DB.XML, tGTPID, tTraderID, tDataBlock.Year, tDataBlock.Month, tDataBlock.Day, "request")
    fM30308NodeExtract tM30308Node, tPrevDataBlock 'Чтение данных из ноды (если ноды нет или данные не прочитаны tPrevDataBlock.Ready будет FALSE)
        
'07 // Проверка на дату и время
    If Not fM30308TimeLimitsCheck(tGTPID, tTraderID, tSemiAcceptMode, tDataBlock, tPrevDataBlock, tReport) Then: Exit Sub

'08 // Вынесение решения о возможности замены (даже если нода не поддаётся чтению, её надо будет заместить)
    If Not tM30308Node Is Nothing Then
        tKillNode = True
        'If fM30308NodeExtract(tM30308Node, tPrevDataBlock) Then
        If tPrevDataBlock.Ready Then
            uCDebugPrint tLogTag, 0, "RECIEVED CHECK >> PREV_MSG = " & tPrevDataBlock.RecievedTime & " NEW_MSG = " & tRecieved
            'If DateDiff("s", tPrevDataBlock.RecievedTime, tRecieved) <= 0 Then: tKillNode = False
            If CDbl(tRecieved) < CDbl(tPrevDataBlock.RecievedTime) Then: tKillNode = False
            'datediff is Date2 - Date1
        End If
        'node eraser
        If tKillNode Then
            Set tOldNode = tM30308Node.ParentNode.RemoveChild(tM30308Node) 'self-erase
        Else
            tReport.Decision = 10
            tComment = "Временная аномалия! Есть более актуальная заявка принятая <" & tPrevDataBlock.RecievedTime & " " & fGetUTCForm(gLocalUTC) & "> с адреса <" & tPrevDataBlock.EmailSender & ">!"
            uCDebugPrint tLogTag, 2, tComment
            tReport.AddReason tComment
            Exit Sub
        End If
    End If
'09 // Внесение заявки в БД (оповещение оператора если это замена)
    'задание номера
    If tPrevDataBlock.Ready Then
        tNumber = tPrevDataBlock.Number
        If Not IsNumeric(tNumber) Then 'fixit
            tNumber = 1
        ElseIf tNumber < 1 Then
            tNumber = 1
        Else
            tNumber = tNumber + 1 'normal line
        End If
    Else
        tNumber = 1
    End If
    tDataBlock.Number = tNumber
    'сброс заявки в БД
    If Not fM30308Inject(gXML30308DB.XML, tDataBlock) Then
        tReport.IsInternal = True
        tReport.Decision = 20
        tComment = "Запись заявки в БД не удалась!"
        uCDebugPrint tLogTag, 2, tComment
        tReport.AddReason tComment
        Exit Sub
    End If
    
'10 // Сохранение БД
    fSaveXMLDB gXML30308DB, False, , , , tLogTag & " добавила заявку!"
    
'11 // Проверка на необходимость вложения в ответ ретроспективы факт\план
    'If fM30308Retrospective(tGraphPicturePath, tRetroComment, tReport.ExchangeNode) Then
        
    'End If
    
    uCDebugPrint tLogTag, 0, "Присвоен номер #" & tNumber

'12 // Проведение экспресс анализа данных на аномалии по временному ряду
    If tNumber = 1 Then
        tReport.AddReason "Заявка принята и зарегистрирована под номером - " & tDataBlock.Number & "."
        tReport.AddReason "Заявленный объём потребления ГТП от внешней электросети (МВт*ч): " & tDataBlock.Total(0)
    Else
        tReport.AddReason "Корректирующая заявка принята и зарегистрирована под номером - " & tDataBlock.Number & "."
        tReport.AddReason "Заявленный объём потребления ГТП от внешней электросети (МВт*ч): " & tDataBlock.Total(0) & " (предыдущий составлял - " & tPrevDataBlock.Total(0) & ")"
        tReport.Decision = 1 'accepted but - correction report
    End If
'13 // Сброс заявки
    'tFileName = inFile.Name
    If Not fGetStorageListByTag(g30308FolderTag, tPathList, tToAll, tErrorText) Then
        uCDebugPrint tLogTag, 2, "Копирование файлов заявок невозможно! Ошибка получения каталогов хранения заявок: " & tErrorText
        Exit Sub
    End If
    
    tResultCollector = Empty
    tDropTriggered = False
    For Each tDropFolder In tPathList
    
        If tResultCollector And Not tToAll Then: Exit For
        
        tResultElement = uCopyFile(inFile.Path, tDropFolder & "\" & inFile.Name)
        
        If tResultElement Then: tDropTriggered = True
        
        If IsEmpty(tResultCollector) Then
            tResultCollector = tResultElement
        Else
            tResultCollector = tResultCollector And tResultElement
        End If
                
        uCDebugPrint tLogTag, 0, "ЗАЯВКА перемещена > " & tDropFolder & "\" & inFile.Name
    Next
        
    If tDropTriggered Then: tReport.AddCommand tGTPID
End Sub

Private Function fM30308Retrospective(outGraphPiturePath, outComment, inExchangeParentNode)
    Dim tExchangeNode, tLogTag, tXPathString, tDepth, tRSVResult
    Dim tNode, tBasisGTPNode, tGTPCode, tTraderCode, tTradeZoneCode, tSubjectID, tErrorText

    ' 00 // Prepare
    fM30308Retrospective = False
    outGraphPiturePath = vbNullString
    outComment = vbNullString
    tLogTag = fGetLogTag("M30308Retro")
    
    ' 01 // Get EXCHANGE 30308 Node
    If inExchangeParentNode Is Nothing Then
        uCDebugPrint tLogTag, 2, "Нода Exchange 30308 не определена на входе!"
        Exit Function
    End If
    
    tXPathString = "child::item[(@id='" & clsTagM30308 & "' and @enabled='1')]"
    Set tExchangeNode = inExchangeParentNode.SelectSingleNode(tXPathString)
    If tExchangeNode Is Nothing Then
        uCDebugPrint tLogTag, 2, "Нода Exchange 30308 не определена! XPath = " & tXPathString
        Exit Function
    End If
    
    ' 02 // Get RETRO DEPTH attr
    If Not fGetAttr(tExchangeNode, "retroreport", tDepth) Then: Exit Function
    If IsNumeric(tDepth) Then
        tDepth = Fix(tDepth)
        If tDepth < 0 Then: tDepth = 0
        If tDepth > 30 Then: tDepth = 30
        If tDepth = 0 Then: Exit Function 'trottle
    Else
        Exit Function
    End If
    
    ' 03 // Чтение кода ГТП из BASIS (tBasisGTPNode - будет хранить ссылку на ноду ГТП)
    tXPathString = "ancestor::gtp"
    If Not fGetAttributeCFG(gXMLBasis, tXPathString, "id", tGTPCode, tBasisGTPNode, tErrorText, inExchangeParentNode) Then: Exit Function
    
    ' 04 // Чтение кода Торговца из BASIS
    tXPathString = "ancestor::trader"
    If Not fGetAttributeCFG(gXMLBasis, tXPathString, "id", tTraderCode, tNode, tErrorText, tBasisGTPNode) Then: Exit Function
        
    ' 05 // Чтение кода ценовой зоны выбранной ГТП из BASIS и Dictionary
    ' .01 // Чтение кода региона из BASIS
    tXPathString = "child::settings"
    If Not fGetAttributeCFG(gXMLBasis, tXPathString, "subjectid", tSubjectID, tNode, tErrorText, tBasisGTPNode) Then: Exit Function
    
    ' .02 // Чтение кода ценовой зоны
    tXPathString = "//subjects/subject[@id='" & tSubjectID & "']"
    If Not fGetAttributeCFG(gXMLDictionary, tXPathString, "tradezone", tTradeZoneCode, tNode, tErrorText) Then: Exit Function
    
    uCDebugPrint tLogTag, 0, "Запрос на ретроспективу для [" & tTraderCode & ":" & tGTPCode & ":" & tSubjectID & ":" & tTradeZoneCode & "] глубиной " & tDepth & " дней."
    
 '   Exit Function
    
    ' 05 // Get RSV by depth
    If fGetRSVTimeLine(tRSVResult, tTraderCode, tGTPCode, tTradeZoneCode, Fix(Now() - tDepth), Fix(Now())) Then
    End If
    
End Function

Private Function fGetRSVTimeLine(outResult, inTraderCode, inGTPCode, inTradeZoneCode, inStartDate, inEndDate)
    Dim tSize, tDayIndex, tHourIndex, tDayCounter, tChildNode, tValue
    Dim tRSVData
    Dim tNode
    
    ' 01 // Предопределения
    fGetRSVTimeLine = False
    
    tSize = (inEndDate - inStartDate + 1) * 24 - 1
    ReDim outResult(tSize)
    
    ' 02 // Подготовка пустышки
    For tHourIndex = 0 To tSize
        outResult(tHourIndex) = -1
    Next
    
    ' 03 // Подготовка данных
    tDayCounter = 0
    For tDayIndex = inStartDate To inEndDate
        If fGetRSVDataNodeByDate(inTraderCode, inGTPCode, inTradeZoneCode, tDayIndex, tNode, True) Then
            For Each tChildNode In tNode.ChildNodes
                If tChildNode.NodeName = "hour" Then
                    If fGetTypedAttributeByName(tChildNode, "id", "INT", tValue) Then
                        tHourIndex = tDayCounter * 24 + tValue - 1
                        If fGetTypedAttributeByName(tChildNode, "rsvvalue", "DEC", tValue) Then
                            outResult(tHourIndex) = tValue
                        End If
                    End If
                End If
            Next
        End If
        tDayCounter = tDayCounter + 1
    Next
    
End Function

Private Function fGetTypedAttributeByName(inNode, inAttributeName, inAttributeType, outValue)
Dim tValue
    
    ' 01 // Предопределения
    fGetTypedAttributeByName = False
    outValue = 0
    
    ' 02 // Существует ли нода
    If inNode Is Nothing Then: Exit Function
    
    ' 03 // Сущеустрвует ли аттрибут
    tValue = inNode.GetAttribute(inAttributeName)
    If IsNull(tValue) Then: Exit Function
    
    ' 04 // Конвертация в требуемый тип
    Select Case UCase(inAttributeType)
        Case "INT":
            If Not IsNumeric(tValue) Then: Exit Function
            outValue = Fix(tValue)
        Case "STR":
            outValue = tValue
        Case "DEC":
            If Not IsNumeric(tValue) Then: Exit Function
            outValue = CDec(tValue)
        Case Else: Exit Function
    End Select
    
    ' 05 // Выход
    fGetTypedAttributeByName = True
End Function

' Общая функция для работы с путями файлов NReport (должна быть единой для всего что работает с этими файлами)
Public Function fGetNReportFilePath(inNReportAlias, inTraderCode, inGTPCode, inTradeZoneCode, inYear, inMonth, inDay, outDirPath, outFileName, outErrorText, Optional inSubCreationEnabled = True, Optional inOverrideHomeDir = vbNullString, Optional inFileNameOnly = False)
    Dim tHomeDir, tSubDir, tContainerDir, tFileName
    Dim tLogTag, tYear, tMonth, tDay
    
    ' 00 \\ Предопределения
    fGetNReportFilePath = False
    outErrorText = vbNullString
    outDirPath = vbNullString
    outFileName = vbNullString
    tLogTag = "fGetNReportFilePath"
       
    ' 01 \\ Форматирование
    tYear = Format(inYear, "0000")
    tMonth = Format(inMonth, "00")
    tDay = Format(Day(inDay), "00")
    
    ' 02 \\ Построение пути для отчета по NReportAlias
    tSubDir = inTraderCode
    Select Case inNReportAlias
        Case "buy_norem":
            tSubDir = tSubDir & "\" & inNReportAlias & "\" & tYear & "\" & tMonth & "\" & tDay
            tFileName = tYear & tMonth & tDay & "_" & inTraderCode & "_" & inGTPCode & "_buy_norem.xls" '20200626_BELKAMKO_PBELKA12_buy_norem.xls
            
        Case Else:
            outErrorText = tLogTag & " > Алиас (" & inNReportAlias & ") отчета не обнаружен в списке допустимых!"
            Exit Function
    End Select
    
    ' 03 \\ Поиск пути и(или) его формирование (если опция имени выбрана - не трогаем создание папок)
    If Not inFileNameOnly Then
        
        ' 03.01 \\ Домашняя директория
        If inOverrideHomeDir = vbNullString Then
            tHomeDir = gDataPath
        Else
            tHomeDir = inOverrideHomeDir
        End If
        
        ' 03.02 \\ Есть ли директория?
        If Not uFileExists(tHomeDir) Then
            outErrorText = tLogTag & " > Не обнаружена домашняя директория: " & tHomeDir
            Exit Function
        End If
    
        ' 03.03 \\ Коррекция
        If Right(tHomeDir, 1) <> "\" Then: tHomeDir = tHomeDir & "\" 'fix it
    
        ' 03.04 \\ Создание субдиректории
        If inSubCreationEnabled Then
            If Not fDirPathAutoBuilder(tHomeDir, tSubDir, tContainerDir, outErrorText) Then
                outErrorText = tLogTag & " > " & outErrorText
                Exit Function
            End If
        Else
            tContainerDir = tHomeDir
        End If
    End If
    
    ' 04 \\ В зависимости от опции формируем результат
    If inFileNameOnly Then
        fGetNReportFilePath = True
        outFileName = tFileName
    Else
        fGetNReportFilePath = True
        outDirPath = tContainerDir
        outFileName = tFileName
    End If
    
End Function

Public Function fDirPathAutoBuilder(inParentDir, inSubPath, outFinalPath, outErrorText)
    Dim tLogTag, tParentDir, tFinalPath
    Dim tSubFolderList, tSubFolder

    ' 00 \\ Предопределения
    tLogTag = "fDirPathAutoBuilder"
    fDirPathAutoBuilder = False
    outFinalPath = vbNullString
    outErrorText = vbNullString
    
    ' 01 \\ Правка
    tParentDir = inParentDir
    If Right(tParentDir, 1) <> "\" Then: tParentDir = tParentDir & "\" 'fix it
    
    ' 02 \\ Предпроверка (вдруг уже создано)
    tFinalPath = tParentDir & inSubPath
    If uFileExists(tFinalPath) Then 'forced exit
        fDirPathAutoBuilder = True
        outFinalPath = tFinalPath
        Exit Function
    End If
    
    ' 03 \\ Проверка наличия родительской дирктории
    If Not uFileExists(inParentDir) Then
        outErrorText = tLogTag & "Не обнаружена родительская директория: " & inParentDir
        Exit Function
    End If
    
    ' 04 \\ Воссоздание субдиректорий путём сплита пути
    tSubFolderList = Split(inSubPath, "\")
    
    tFinalPath = tParentDir
    For Each tSubFolder In tSubFolderList
        tFinalPath = tFinalPath & "\" & tSubFolder
        If Not uFolderCreate(tFinalPath) Then
            outErrorText = tLogTag & "Не удалось создать директорию <" & tSubFolder & ">: " & tFinalPath
            Exit Function
        End If
    Next
    
    ' 05 \\ Успешное завершение
    fDirPathAutoBuilder = True
    outFinalPath = tFinalPath
End Function

' Download NReport File
' RESOURCES USED: CREDENTIALS
Private Function fDownloadNReportFile(inNReportAlias, inTraderCode, inGTPCode, inTradeZoneCode, inYear, inMonth, inDay, outNReportFileName, outDownloadDir, Optional inAutoSubDir = True, Optional inUseZip = True, Optional inCanDownloadFiles = True, Optional inForcedDownload = False, Optional inOverrideHomeDir = vbNullString)
    Dim tLogTag, tErrorText
    Dim tReportFilePath, tReportFileDir, tReportFileName, tFileLocked, tDownloadReportFilePath, tPos
    Dim tFileDownloader As CATSDownloader
    Dim tPartCode, tUserName, tPassword
    Dim tXPathString, tValue, tNode
    
    ' 00 \\ Предопределения
    tLogTag = fGetLogTag("DownloadNReportFile")
    fDownloadNReportFile = False
    outNReportFileName = vbNullString
    outDownloadDir = vbNullString
    uCDebugPrint tLogTag, 0, "Инициация [PARAMS:NReport=" & inNReportAlias & ", Trader=" & inTraderCode & ", GTP=" & inGTPCode & ", TZ=" & inTradeZoneCode & ", Y=" & inYear & ", M=" & inMonth & ", D=" & inDay & "]" & _
    "// SETTINGS: [AutoDir=" & inAutoSubDir & "][UseZip=" & inUseZip & "][CanDownload=" & inCanDownloadFiles & "][ForceDownload=" & inForcedDownload & "][OverrideHomeDir=" & inOverrideHomeDir & "]"
    
    ' 01 \\ Формирование пути к отчету и его имени файла
    If Not fGetNReportFilePath(inNReportAlias, inTraderCode, inGTPCode, inTradeZoneCode, inYear, inMonth, inDay, tReportFileDir, tReportFileName, tErrorText, inAutoSubDir, inOverrideHomeDir) Then
        uCDebugPrint tLogTag, 2, tErrorText
        Exit Function
    End If
    
    ' 02 \\ Проверка пути (есть ли файл с этим именем?)
    tReportFilePath = tReportFileDir & "\" & tReportFileName
    tFileLocked = uFileExists(tReportFilePath)
    
    ' 03 \\ Попытка удалить файл если был применен параметр жесткой перезаписи файла
    If inForcedDownload And tFileLocked Then
        If Not uDeleteFile(tReportFilePath) Then
            uCDebugPrint tLogTag, 2, "Не удалось удалить файл! Путь: " & tReportFilePath
            Exit Function
        End If
        tFileLocked = False
    End If
    
    ' 04 \\ Файл уже существует? Скачивание не требуется
    If tFileLocked Then
        outNReportFileName = tReportFileDir
        outDownloadDir = tReportFileName
        fDownloadNReportFile = True
        uCDebugPrint tLogTag, 0, "Файл получен! Путь: " & tReportFilePath
        Exit Function
    ElseIf Not inCanDownloadFiles Then
        uCDebugPrint tLogTag, 2, "Файл не найден, а скачивание запрещено! Путь: " & tReportFilePath
        Exit Function
    End If
    
    ' 05 \\ Для скачивания нужна авторизация
    If Not gXMLCredentials.Active Then
        uCDebugPrint tLogTag, 2, "Ресурс недоступен - " & gXMLCredentials.ClassTag
        Exit Function
    End If
    
     ' 06 \\ Получение данных аутидентификации из конфига
    tPartCode = inTraderCode 'PARTCODE
        
    tXPathString = "//trader[@id='" & inTraderCode & "']/service[@id='atsenergo']/item[@partcode='" & inTraderCode & "']"
    If Not fGetAttributeCFG(gXMLCredentials, tXPathString, "username", tValue, tNode, tErrorText) Then
        uCDebugPrint tLogTag, 2, tErrorText
        Exit Function
    End If
    tUserName = tValue 'USERNAME
        
    If Not fGetAttributeCFG(gXMLCredentials, vbNullString, "password", tValue, tNode, tErrorText, tNode) Then
        uCDebugPrint tLogTag, 2, tErrorText
        Exit Function
    End If
    tPassword = tValue 'PASSWORD
    
    ' 07 \\ Скачивание файла
    Set tFileDownloader = New CATSDownloader ' new object
        
    If tFileDownloader.IsActive Then
        tFileDownloader.SetNReportCredentials tPartCode, tUserName, tPassword, True
        tFileLocked = tFileDownloader.GetNReportPersonalFileByClass(inYear, inMonth, inDay, inNReportAlias, inGTPCode, inTradeZoneCode, tReportFileDir, tDownloadReportFilePath, , , True)
    End If
        
    Set tFileDownloader = Nothing ' kill obj
    
    ' 08 \\ Скачивание файла
    If tFileLocked Then
        tPos = InStrRev(tDownloadReportFilePath, "\") 'not failsafe?
        outNReportFileName = Right(tDownloadReportFilePath, Len(tDownloadReportFilePath) - tPos)
        outDownloadDir = Left(tDownloadReportFilePath, tPos - 1)
        fDownloadNReportFile = True
        uCDebugPrint tLogTag, 0, "Файл получен! Путь: " & tReportFilePath
        Exit Function
    End If
    
End Function

' NEED EXCEL OBJECT to WORK
' gExcel internal object on use
Private Function fReportBuyNoremDataExtraction(inFilePath, inFileName, inGTPCode, inTraderCode, inYear, inMonth, inDay, outDataBlock As TReportBuyNoremDataBlock)
Dim tLogTag, tWorkBook, tDateValue, tValue, tTempValue, tSuccessReads, tIsReadingOK
Dim tStartRow, tCurrentRow, tRowStep, tHourIndex, tColumnIndex, tFile

    ' 00 \\ Предопределения
    tLogTag = fGetLogTag("ReportBuyNoremDataExtraction")
    fReportBuyNoremDataExtraction = False
    tStartRow = 15
    tRowStep = 4
        
    On Error Resume Next
    
        ' 01 \\ Попытка открыть файл
        Set tWorkBook = gExcel.Workbooks.Open(inFilePath, False, True)
        
        If Err.Number <> 0 Then
            uCDebugPrint tLogTag, 2, "Ошибка (" & Err.Number & "): " & Err.Description
            uCDebugPrint tLogTag, 2, "Не удалось открыть книгу Excel! Путь: " & tReportFilePath
            Exit Function
        End If
        
        ' 01 \\ Попытка чтения содержимого
        fExcelControl -1, -1, -1, -1 'сброс контроля
        
        With tWorkBook.WorkSheets(1)
            
        ' 02 \\ Блок проверок
            
            '02.01 \\ DATE check
            tDateValue = inYear & "-" & inMonth & "-" & inDay
            tValue = .Cells(1, 7).Value
            
            If tDateValue <> tValue Then
                uCDebugPrint tLogTag, 2, "Ошибка проверки! Ожидалась дата <" & tDateValue & ">, а прочитана <" & tValue & ">!"
                uCDebugPrint tLogTag, 2, "Путь к файлу-отчету: " & tReportFilePath
                tWorkBook.Close SaveChanges:=False 'silent close book without saving
                Set tWorkBook = Nothing
                fExcelControl 1, 1, 1, 1 'возврат контроля
                Exit Function
            End If
            
            '02.02 \\ GTP check
            tValue = UCase(.Cells(4, 3).Value)
            If inGTPCode <> tValue Then
                uCDebugPrint tLogTag, 2, "Ошибка проверки! Ожидалась ГТП <" & inGTPCode & ">, а прочитана <" & tValue & ">!"
                uCDebugPrint tLogTag, 2, "Путь к файлу-отчету: " & tReportFilePath
                tWorkBook.Close SaveChanges:=False 'silent close book without saving
                Set tWorkBook = Nothing
                fExcelControl 1, 1, 1, 1 'возврат контроля
                Exit Function
            End If
            
            '02.03 \\ TRADER check (пропустим)
        
        ' 03 \\ Чтение данных
            tSuccessReads = 0
            tCurrentRow = tStartRow
            For tHourIndex = 1 To 24
                
                tIsReadingOK = True
                
                ' Проверка позиции
                tTempValue = Format(tHourIndex - 1, "00")
                If tHourIndex < 24 Then
                    tTempValue = tTempValue & "-" & Format(tHourIndex, "00")
                Else
                    tTempValue = tTempValue & "-00"
                End If
                tValue = .Cells(tCurrentRow, 1).Value
                
                If tValue = tTempValue Then
                'Чтение строки данных
                
                    'C4 / ATS Value
                    tColumnIndex = 4
                    tValue = .Cells(tCurrentRow, tColumnIndex).Value
                    If IsNumeric(tValue) Then
                        outDataBlock.ValueATS(0) = outDataBlock.ValueATS(0) + tValue
                        outDataBlock.ValueATS(tHourIndex) = tValue
                    Else
                        tIsReadingOK = False
                        uCDebugPrint tLogTag, 2, "Ошибка чтения позиции: В ячейке[R:" & tCurrentRow & " C:" & tColumnIndex & "] ожидалось число, а получено <" & tValue & ">!"
                    End If
                    
                    'C5 / ATSCorrection Value
                    tColumnIndex = 5
                    tValue = .Cells(tCurrentRow, tColumnIndex).Value
                    If IsNumeric(tValue) Then
                        outDataBlock.ValueATSCorrection(0) = outDataBlock.ValueATSCorrection(0) + tValue
                        outDataBlock.ValueATSCorrection(tHourIndex) = tValue
                    Else
                        tIsReadingOK = False
                        uCDebugPrint tLogTag, 2, "Ошибка чтения позиции: В ячейке[R:" & tCurrentRow & " C:" & tColumnIndex & "] ожидалось число, а получено <" & tValue & ">!"
                    End If
                    
                    'C7 / SO Value
                    tColumnIndex = 7
                    tValue = .Cells(tCurrentRow, tColumnIndex).Value
                    If IsNumeric(tValue) Then
                        outDataBlock.ValueSO(0) = outDataBlock.ValueSO(0) + tValue
                        outDataBlock.ValueSO(tHourIndex) = tValue
                    Else
                        tIsReadingOK = False
                        uCDebugPrint tLogTag, 2, "Ошибка чтения позиции: В ячейке[R:" & tCurrentRow & " C:" & tColumnIndex & "] ожидалось число, а получено <" & tValue & ">!"
                    End If
                    
                    'C9 / Accepted Value
                    tColumnIndex = 9
                    tValue = .Cells(tCurrentRow, tColumnIndex).Value
                    If IsNumeric(tValue) Then
                        outDataBlock.ValueAccepted(0) = outDataBlock.ValueAccepted(0) + tValue
                        outDataBlock.ValueAccepted(tHourIndex) = tValue
                    Else
                        tIsReadingOK = False
                        uCDebugPrint tLogTag, 2, "Ошибка чтения позиции: В ячейке[R:" & tCurrentRow & " C:" & tColumnIndex & "] ожидалось число, а получено <" & tValue & ">!"
                    End If
                    
                    'C10 / Average Price
                    tColumnIndex = 10
                    tValue = .Cells(tCurrentRow, tColumnIndex).Value
                    If IsNumeric(tValue) Then
                        outDataBlock.PriceHourAverage(0) = outDataBlock.PriceHourAverage(0) + tValue
                        outDataBlock.PriceHourAverage(tHourIndex) = tValue
                    Else
                        tIsReadingOK = False
                        uCDebugPrint tLogTag, 2, "Ошибка чтения позиции: В ячейке[R:" & tCurrentRow & " C:" & tColumnIndex & "] ожидалось число, а получено <" & tValue & ">!"
                    End If
                    
                    'C11 / Generation Value
                    tColumnIndex = 11
                    tValue = .Cells(tCurrentRow, tColumnIndex).Value
                    If IsNumeric(tValue) Then
                        outDataBlock.ValueGeneration(0) = outDataBlock.ValueGeneration(0) + tValue
                        outDataBlock.ValueGeneration(tHourIndex) = tValue
                    Else
                        tIsReadingOK = False
                        uCDebugPrint tLogTag, 2, "Ошибка чтения позиции: В ячейке[R:" & tCurrentRow & " C:" & tColumnIndex & "] ожидалось число, а получено <" & tValue & ">!"
                    End If
                    
                    'C13 / RSV Value
                    tColumnIndex = 13
                    tValue = .Cells(tCurrentRow, tColumnIndex).Value
                    If IsNumeric(tValue) Then
                        outDataBlock.ValueRSV(0) = outDataBlock.ValueRSV(0) + tValue
                        outDataBlock.ValueRSV(tHourIndex) = tValue
                    Else
                        tIsReadingOK = False
                        uCDebugPrint tLogTag, 2, "Ошибка чтения позиции: В ячейке[R:" & tCurrentRow & " C:" & tColumnIndex & "] ожидалось число, а получено <" & tValue & ">!"
                    End If
                    
                    'C15 / RSV Price
                    tColumnIndex = 15
                    tValue = .Cells(tCurrentRow, tColumnIndex).Value
                    If IsNumeric(tValue) Then
                        outDataBlock.PriceHourRSV(0) = outDataBlock.PriceHourRSV(0) + tValue
                        outDataBlock.PriceHourRSV(tHourIndex) = tValue
                    Else
                        tIsReadingOK = False
                        uCDebugPrint tLogTag, 2, "Ошибка чтения позиции: В ячейке[R:" & tCurrentRow & " C:" & tColumnIndex & "] ожидалось число, а получено <" & tValue & ">!"
                    End If
                    
                    'FINALIZE ROW READ
                    If tIsReadingOK Then
                        tSuccessReads = tSuccessReads + 1
                    Else
                        uCDebugPrint tLogTag, 2, "Путь к файлу-отчету: " & tReportFilePath
                        Exit For
                    End If
                    'END
                Else
                    uCDebugPrint tLogTag, 2, "Ошибка чтения позиции: Ячейка[R:" & tCurrentRow & " C:1] ожидалось значение <" & tTempValue & ">, а получено <" & tValue & ">!"
                    uCDebugPrint tLogTag, 2, "Путь к файлу-отчету: " & tReportFilePath
                    Exit For
                End If
                
                tCurrentRow = tCurrentRow + tRowStep
            Next
            
        End With
                
        ' 04 \\ Проверка через итог таблицы правильности сбора информации
        If tSuccessReads = 24 Then
        'Row select
            tCurrentRow = tCurrentRow - tRowStep + 1
        'Column to check
            tColumnIndex = 13
        'Read
            tValue = .Cells(tCurrentRow, tColumnIndex).Value
        'Checks
            If IsNumeric(tValue) Then
                If Round(outDataBlock.ValueRSV(0) - tValue, 3) <> 0 Then
                    tSuccessReads = 0
                    uCDebugPrint tLogTag, 2, "Ошибка чтения итога: В ячейке[R:" & tCurrentRow & " C:" & tColumnIndex & "] ожидалось <" & outDataBlock.ValueRSV(0) & ">, а получено <" & tValue & ">!"
                End If
            Else
                tSuccessReads = 0
                uCDebugPrint tLogTag, 2, "Ошибка чтения позиции: В ячейке[R:" & tCurrentRow & " C:" & tColumnIndex & "] ожидалось число, а получено <" & tValue & ">!"
            End If
        End If
                
        ' 05 \\ Закрытие книги
        tWorkBook.Close SaveChanges:=False 'silent close book without saving
        Set tWorkBook = Nothing
        fExcelControl 1, 1, 1, 1 'возврат контроля
        
        ' 06 \\ Подготовка и проверка итога
        If tSuccessReads <> 24 Then: Exit Function
        With outDataBlock
            .PriceHourRSV(0) = .PriceHourRSV(0) / 24
            .PriceHourAverage(0) = .PriceHourAverage(0) / 24
            .Day = inDay
            .Month = inMonth
            .Year = inYear
            .GTPID = inGTPCode
            .Ready = True
            .TraderID = inTraderCode
            .Version = 1
            .Date = DateSerial(inYear, inMonth, inDay)
            .Comment = vbNullString
            .FileName = inFileName
            
            Set tFile = gFSO.GetFile(inFilePath)
            .ModificationDate = tFile.DateLastModified
            Set tFile = Nothing
        End With
        
        ' 07 \\ Самопроверка
        If Err.Number <> 0 Then
            outDataBlock.Ready = False
            uCDebugPrint tLogTag, 2, "Ошибка (" & Err.Number & "): " & Err.Description
            uCDebugPrint tLogTag, 2, "Не удалось успешно прочитать отчет! Путь: " & tReportFilePath
            Exit Function
        End If
        
    On Error GoTo 0
    
    fReportBuyNoremDataExtraction = True

End Function

' RESOURCES USED: R30308DB*, CREDENTIALS*
' Return a NODE with RSVData in M30308DB
Private Function fM30308InjectRSVData(inXMLDB, inGTPCode, inTraderCode, inTradeZoneCode, inYear, inMonth, inDay, Optional inDownloadFiles = True)
    Dim tLogTag, tErrorText, tValue, tReportFileDir, tReportFileName
    'Dim tPartCode, tUserName, tPassword
    Dim tNReportAlias, tReportFilePath, tFileLocked
    Dim tContainerNode, tXPathString, tNodes, tNode, tTempNode
    Dim tDataBlock As TReportBuyNoremDataBlock
    
    ' 00 \\ Предопределения
    tLogTag = fGetLogTag("M30308InjectRSVData")
    Set fM30308InjectRSVData = Nothing
    
    ' 01 \\ Запрос на расположение или получение требуемого файла
    tNReportAlias = "buy_norem"
    If Not fDownloadNReportFile(tNReportAlias, inTraderCode, inGTPCode, inTradeZoneCode, inYear, inMonth, inDay, tReportFileName, tReportFileDir, , , inDownloadFiles) Then
        uCDebugPrint tLogTag, 2, "Файл не обнаружен!"
        Exit Function
    End If
  
    ' 02 \\ Проверка на наличии файла (по идее нужна только для дебага)
    tReportFilePath = tReportFileDir & "\" & tReportFileName
    tFileLocked = uFileExists(tReportFilePath)
    If Not tFileLocked Then: Exit Function 'logic corrupted
    
    ' 03 \\ Извлечение данных из файла-отчета (подготовка датаблока)
    If Not fReportBuyNoremDataExtraction(tReportFilePath, tReportFileName, inGTPCode, inTraderCode, inYear, inMonth, inDay, tDataBlock) Then: Exit Function
    
    ' 04 \\ Инъекция блока данных в XMLDB
    If Not fM30308RSVInject(gXML30308DB.XML, tDataBlock) Then
        uCDebugPrint tLogTag, 2, "Инъекция не удалась!"
        fReloadXMLDB gXML30308DB, False 'forced reloadDB - to clear any changes \\ rollback db
        Exit Function
    Else
        fSaveXMLDB gXML30308DB, False
    End If
  
    ' 05 \\ Завершение
    Set fM30308InjectRSVData = fGetM30308Node(gXML30308DB.XML, inGTPCode, inTraderCode, inYear, inMonth, inDay, "trade")
End Function


' RESOURCES USED: R30308DB, CREDENTIALS
Private Function fGetRSVDataNodeByDate(inTraderCode, inGTPCode, inTradeZoneCode, inDate, outDataNode, Optional inDownloadFiles = False)
    Dim tLogTag, tDay, tMonth, tYear
    Dim tNode, tXPathString
    
    ' 00 \\ Предопределения
    tLogTag = fGetLogTag("GetRSVDataNodeByDate")
    fGetRSVDataNodeByDate = False
    Set outDataNode = Nothing
    
    If Not IsDate(inDate) Then
        uCDebugPrint tLogTag, 1, "Входящая дата не является датой! inDate=" & inDate
        Exit Function
    End If
    
    tYear = Format(Year(inDate), "0000")
    tMonth = Format(Month(inDate), "00")
    tDay = Format(Day(inDate), "00")
        
    ' 01 \\ Проверки ресурсов
    If Not gXML30308DB.Active Then
        uCDebugPrint tLogTag, 1, "Ресурс недоступен - " & gXML30308DB.ClassTag
        Exit Function
    End If
    
    ' 02 \\ Получение ноды из БД
    Set tNode = fGetM30308Node(gXML30308DB.XML, inGTPCode, inTraderCode, tYear, tMonth, tDay, "trade")
    
    If (tNode Is Nothing) And inDownloadFiles Then
        Set tNode = fM30308InjectRSVData(gXML30308DB.XML, inGTPCode, inTraderCode, inTradeZoneCode, tYear, tMonth, tDay) 'default - TRUE to download files
    End If
    
    ' 03 \\ Возврат ноды или Nothing
    If Not tNode Is Nothing Then
        fGetRSVDataNodeByDate = True
        Set outDataNode = tNode
    End If
End Function

Private Sub fAttachmentReprocess_XLS(inFilePath, outClass, outAReports As TReportAssist)
Dim tValue, tNode, tClass, tGTP, tValidName, tValidStructure, tComment, tFileName, tVersion, tDataSet, tFile
    If gFSO.FileExists(inFilePath) Then
        Set tFile = gFSO.GetFile(inFilePath)
        fClassificator_XLS tFile, outClass, tVersion, tValidName, tValidStructure, tDataSet, tComment
        Select Case outClass
            Case clsTagM30308: fM30308XLSAttachmentReprocess tFile, tVersion, outAReports, tValidStructure, tValidName, tDataSet, tComment
        End Select
    End If
End Sub

'Обработчик ВЛОЖЕНИЙ
Private Function fAttachmentHandler(inAttachment As Attachment)
Dim tFileExtension, tLogTag, tDropPath, tClass, tCommandString, tCommandList, tCommand, tIndex, tSubIndex, tPathList, tToAll, tErrorText, tTempFolder
Dim tAReports As TReportAssist
    
' 00 // Подготовка
    tLogTag = fGetLogTag("AttachmentHandler")
    fAttachmentHandler = False
    tClass = vbNullString
    tCommandString = vbNullString
        
    On Error Resume Next ' On Error GoTo 0
        
'01 //
        tFileExtension = uGetFileExtension(inAttachment.FileName)
        uCDebugPrint tLogTag, 0, "Расширение вложения - " & tFileExtension
        tAReports.ReportsCount = -1 'INIT
        
        ' FATAL EXIT
        If Err.Number <> 0 Then
            uCDebugPrint tLogTag, 2, "Не удалось обработать вложение ошибка (#" & Err.Number & "): " & Err.Description
            Exit Function
        End If
        
        If Not fGetStorageListByTag(gTempFolderTag, tPathList, tToAll, tErrorText) Then
            uCDebugPrint tLogTag, 2, "Не удалось обработать вложение! Ошибка временной папки: " & tErrorText
            Exit Function
        End If
        
        tTempFolder = tPathList(0) 'system folder (so ToAll will not be never used)
'02 //
        Select Case tFileExtension
            'Эксемельки
            Case "XML":
                tDropPath = fDropAttachment(inAttachment, tTempFolder, True)
                fAttachmentReprocess_XML tDropPath, tClass, tCommandString
            'Эксельки
            Case "XLSX", "XLS", "XLSM":
                tDropPath = fDropAttachment(inAttachment, tTempFolder, True)
                fAttachmentReprocess_XLS tDropPath, tClass, tAReports
        End Select

'03 // добавление СООБЩЕНИЯ (ПИСЬМА) в автоматический сортировщик по ТЕКУЩЕМУ ВЛОЖЕНИЮ
        If tClass <> vbNullString And tCommandString <> vbNullString Then
            tCommandList = Split(tCommandString, cnstDelimiter)
            For Each tCommand In tCommandList
                fMailListAdd tClass & cnstDelimiter & tCommand 'дубликаты не пройдут, каждая пара CLASS:COMMAND уникальна
            Next
        End If
    
'04 // CLEAR
    For tIndex = 0 To tAReports.ReportsCount
        
        'final actions
        With tAReports.Reports(tIndex)
            If .HasCommands Then
                For tSubIndex = 0 To .CommandCount
                    fMailListAdd .GetCommand(tSubIndex)
                Next
            End If
            
            'reporter
            .FilterAdressList
            fFastRespond .ReportList, .GetHeader, .GetReason(2)
        End With
        
        'kill
        Set tAReports.Reports(tIndex) = Nothing
    Next
    tAReports.ReportsCount = -1
    
'05 //
    fAttachmentHandler = True
End Function

Private Function fOutlookFolderCreator(inPath, inRootFolder, outFolder, Optional inDelimiter = "\") As Boolean
Dim tSubFolders, tDepth, tSubFolder, tLogTag
'Dim tStr As String
Dim tCurrentFolder As Outlook.Folder
'00 // Предопределения
    tLogTag = "OLFLDCREATOR"
    fOutlookFolderCreator = False               'флаг успеха
    Set outFolder = Nothing                     'папка возврата
    tSubFolders = Split(inPath, inDelimiter)    'получим элементы необходимой структуры подпапок
    tDepth = 0                                  'глубина подпапок
    Set tCurrentFolder = fGetOutlookRootFolder     'корневая папка структуры подпапок
'01 // Проверка корневой папки
    If IsEmpty(tCurrentFolder) Then
        uCDebugPrint tLogTag, 2, "Корневая папка не задана!"
        Exit Function
    End If
'02 // Проход по пути
    On Error Resume Next
        For Each tSubFolder In tSubFolders
            tDepth = tDepth + 1
            tSubFolder = CStr(tSubFolder)
            Set tCurrentFolder = tCurrentFolder.Folders(tSubFolder) 'Получим подпапку
            If Err.Number <> 0 Then
                If Err.Number = -2147221233 Then
                    uCDebugPrint tLogTag, 1, "Подпапка <" & tSubFolder & "> уровень " & tDepth & " не найдена. Создание искомой."
                    Err.Clear
                    Set tCurrentFolder = tCurrentFolder.Folders.Add(tSubFolder) 'создание новой подпапки
                    If Err.Number <> 0 Then
                        uCDebugPrint tLogTag, 2, "Не удалось создать подпапку <" & tSubFolder & "> уровень " & tDepth & " > Ошибка " & Err.Number & " > " & Err.Description
                        Exit Function
                    End If
                Else
                    uCDebugPrint tLogTag, 2, "Не удалось полученить подпапку <" & tSubFolder & "> уровень " & tDepth & " > Ошибка " & Err.Number & " > " & Err.Description
                    Exit Function
                End If
            End If
        Next
    On Error GoTo 0
'03 // Завершение
    Set outFolder = tCurrentFolder
    fOutlookFolderCreator = True
End Function

Private Sub fMailCopyListHandler(inItem As Outlook.MailItem, inFolderToProcess)
Dim tLogTag, tIndex, tIndexLimit, tCommandElements, tClass, tCommand, tSubFolderPath, tErrors, tSuccessCommands, tTotalCommands
Dim tTargetFolder As Outlook.MAPIFolder
Dim tCopiedItem As Outlook.MailItem
Dim tTempFolder As Outlook.MAPIFolder
'00 // Предопределения
    tLogTag = "MAILCOPYHNDL"
'01 // Нужна ли сортировка вообще
    If gMailCopyList.Count <= 0 Then
        uCDebugPrint tLogTag, 0, "Это сообщение не нуждается в автоматической сортировке!"
        Exit Sub
    End If
'02 // Если есть команды проверим.. можно ли применять сортировку к этой папке
    uCDebugPrint tLogTag, 0, "Папка автоматической сортировки <" & fGetOutlookRootFolder & ">."
    If inItem.Parent <> fGetOutlookRootFolder Then
        uCDebugPrint tLogTag, 0, "Папка сообщения <" & inItem.Parent & "> не подлежит сортировке!"
        Exit Sub
    End If
'03 // Получим папку временного хранения сообщений
    If Not fOutlookFolderCreator(gTempMailFolderName, fGetOutlookRootFolder, tTempFolder) Then
        uCDebugPrint tLogTag, 2, "Папка временного хранения сообщений <gTempMailFolderName> не определена!"
        Exit Sub
    End If
    'Set tTempFolder = fGetTempFolder(fGetOutlookRootFolder, gTempMailFolderName)
    On Error Resume Next
        Set inItem = inItem.Move(tTempFolder) 'перенос сообщения во временную папку gTempMailFolderName
        If Err.Number <> 0 Then
            uCDebugPrint tLogTag, 2, "Перенос сообщения во временную папку не удался > " & Err.Description
            Exit Sub
        End If
    On Error GoTo 0
    'проверим еще раз
    If gTempMailFolderName <> inItem.Parent Then
        uCDebugPrint tLogTag, 2, "Неопределнная ошибка! Папка текущего сообщения <" & inItem.Parent & ">, а должна быть <" & gTempMailFolderName & ">"
        Exit Sub
    End If
'04 // Перебор команд на автоматическую сортировку текущего сообщения
    tIndex = 0                                  'индекс для предотвращения бесконечного цикла, опирается на tIndexLimit
    tErrors = 0                                 'сборщик количества ошибок исполнения команд
    tSuccessCommands = 0                        'сборщик количества исполненых команд
    tTotalCommands = gMailCopyList.Count
    tIndexLimit = gMailCopyList.Count * 2
    Do Until (gMailCopyList.Count = 0 Or tIndex >= tIndexLimit)
        tIndex = tIndex + 1 'индекс для защиты от бесонечного цикла
        tCommandElements = Split(gMailCopyList.Item(1), cnstDelimiter)
        If UBound(tCommandElements) = 1 Then
'05 // Разбивка команды на элементы
            tClass = tCommandElements(0)
            tCommand = tCommandElements(1)
            tSubFolderPath = vbNullString
'06 // ФАЗА 1. По классу определим целевую подпапку
            Select Case tClass
                Case "80020": tSubFolderPath = "Макеты\M80020\" & tCommand
                Case "30308": tSubFolderPath = "Заявки контрагентов\" & tCommand
                Case Else:
                    uCDebugPrint tLogTag, 1, "Получена команда <" & gMailCopyList.Item(1) & "> с неожиданным классом <" & tClass & ">!"
                    tErrors = tErrors + 1
            End Select
'07 // ФАЗА 2. Раз целевая папка определена, то проверим существует ли она и создадим её
            If tSubFolderPath <> vbNullString Then
                If fOutlookFolderCreator(tSubFolderPath, fGetOutlookRootFolder, tTargetFolder) Then
'08 // ФАЗА 3. Раз целевая папка определена и существует, то произведем копирование сообщения в неё
                    On Error Resume Next
                        Set tCopiedItem = inItem.Copy
                        tCopiedItem.UnRead = False
                        tCopiedItem.Move tTargetFolder
                        If Err.Number = 0 Then
                            uCDebugPrint tLogTag, 0, "Произведено копирование сообщения в папку <" & tTargetFolder & ">"
                            tSuccessCommands = tSuccessCommands + 1
                        Else
                            uCDebugPrint tLogTag, 2, "Не удалось копирование сообщения > Ошибка " & Err.Number & " > " & Err.Description
                            tErrors = tErrors + 1
                        End If
                    On Error GoTo 0
                Else
                    uCDebugPrint tLogTag, 1, "Команда <" & gMailCopyList.Item(1) & "> не может быть обработана!"
                    tErrors = tErrors + 1
                End If
            End If
        Else
            uCDebugPrint tLogTag, 1, "Получена неверная команда <" & gMailCopyList.Item(1) & ">, ожидалась <CLASS:COMMAND>!"
            tErrors = tErrors + 1
        End If
        gMailCopyList.Remove 1
    Loop
'XX // Завершение сортировки
    If tErrors = 0 Then
        inItem.Delete 'если ошибок нет.. то сообщение оригинал удаляем
        uCDebugPrint tLogTag, 0, "Сортировка закончена; комманд обработано " & tSuccessCommands & " из " & tTotalCommands & "; сообщение удалено из временной папки!"
    Else
        uCDebugPrint tLogTag, 1, "Сортировка закончена; комманд обработано " & tSuccessCommands & " из " & tTotalCommands & "; сообщение ОСТАЛОСЬ во временной папке!"
    End If
End Sub

'MAIN \\ InterceptorMain ... inItem - intecepted e-mail message
'ЭТАП 1 \\ Обработка вложений и распределение писем по категориям и соотвествиям
Public Function fMailReprocessor(inItem As Outlook.MailItem, Optional inForceInit = True) 'inItem As Outlook.MailItem
Dim tNode, tTempPath, tExtension, tDropPath, tClass, tGTPID, tGTP, tVersion, tValidName, tValidStructure, tValue, tDay, tTempDir, tElements, tErrors, tErrorReason, tReadyToCopy, tLogTag
Dim tAttachmentFailIndexList, tAttachmentIndex, tAttachmentReadFail
    
'00 // Опознание типа сообщения
    tLogTag = fGetLogTag("MailReprocessor")
    fMailReprocessor = False

    'Debug.Print "F1-MailReprocessor"
    On Error Resume Next
    
'01 // Чтение параметров
        uCDebugPrint tLogTag, 0, "Старт обработки сообщения (принудительная инициализация - " & inForceInit & ")"
        uCDebugPrint tLogTag, 0, "Класс - " & inItem.MessageClass & "; Тема - " & inItem.Subject
        
        If inItem.MessageClass <> "IPM.Note" Then
            uCDebugPrint tLogTag, 1, "Класс <" & inItem.MessageClass & "> не подлежит обработке, обрабатывается только <IPM.Note>!"
            Exit Function
        End If
        
        'Debug.Print "F2"
        gCurrentMessage.Recieved = inItem.ReceivedTime
        gCurrentMessage.SenderEMail = vbNullString
        
'02 // Определим адрес источника письма
        'Debug.Print "F3"
        If inItem.SenderEmailType = "EX" Then 'если письмо с учетки эксчейнажа
            gCurrentMessage.SenderEMail = inItem.Sender.GetExchangeUser.PrimarySmtpAddress
        Else
            gCurrentMessage.SenderEMail = inItem.SenderEmailAddress
        End If
        
        'Debug.Print "F4"
        uCDebugPrint tLogTag, 0, "Отправитель (" & inItem.SenderEmailType & ") - " & gCurrentMessage.SenderEMail

'03 // Проверка на наличие вложений в письме (если вложений нет - выход)
        uCDebugPrint tLogTag, 0, "Количество вложений - " & inItem.Attachments.Count
        If inItem.Attachments.Count = 0 Then: Exit Function 'пока только работа с вложениями
        
'04 // Отсев на ошибку чтения параметров письма (скорее всего ящик не доступен)
        'Debug.Print "F5"
        If Err.Number <> 0 Then
            uCDebugPrint tLogTag, 2, "Обработка остановлена. Произошла непредвиденная ошибка чтения письма (#" & Err.Number & "): " & Err.Description
            Exit Function
        End If
        
'05 // Инициализация данных
        'Debug.Print "F6"
        If inForceInit Then
            If Not fLocalInit Then: Exit Function        'глобальные переменные
            If Not fXMLSmartUpdate Then: Exit Function 'конфигурация и базы данных
        End If
        
        'подготовка списка для копирования писем
        'Debug.Print "F7"
        fMailCopyListCleaner
        'Debug.Print "F7E"
'06 // Обработка вложений
        For tAttachmentIndex = 1 To inItem.Attachments.Count
            'Debug.Print "F8-A" & tAttachmentIndex
            If Not fAttachmentHandler(inItem.Attachments.Item(tAttachmentIndex)) Then
                uCDebugPrint tLogTag, 0, "Не удалось обработать вложение (индекс - " & tAttachmentIndex & ")! Обработка сообщения завершена с ошибкой!"
                Exit Function
            End If
        Next tAttachmentIndex
        
        'очистка папки от временных файлов
        fTempDropCleaner
    
'07 // Сортировка сообщения по целевым папкам после разбора
        fMailCopyListHandler inItem, fGetOutlookRootFolder
        
    On Error GoTo 0
    
'08 // Завершение
    fMailReprocessor = True
    fMailCopyListCleaner
    uCDebugPrint tLogTag, 0, "Обработка успешно сообщения завершена."
End Function

Private Sub fMailListAdd(inText)
Dim tIndex
    For tIndex = 1 To gMailCopyList.Count
        If gMailCopyList.Item(tIndex) = inText Then: Exit Sub
    Next tIndex
    gMailCopyList.Add inText
End Sub

'ЭТАП 2 \\ Обработка макетов
Public Sub fEngageReprocessor()
    fActivateTimer 15, 1
    uDebugPrint "TIM: Reproccessor engaged!"
End Sub

Private Function fXML80020Reprocess(inFile, inXML, inGTPNameList) As Variant
Dim tNode, tValue, tDate, tAreaIndex, tGTPName, tTraderINN
'00 \\ Подготовка
    fXML80020Reprocess = 0
    inGTPNameList = vbNullString
    uDebugPrint "REP: Обработка файла > " & inFile.Name
'01 \\ Парсинг
    If inXML.parseError.ErrorCode <> 0 Then
        uDebugPrint "REP: Ошибка парсинга XML!"
        Exit Function
    End If
'02 \\ Проверка схемы
    Set inXML.Schemas = gXML80020CFG.XSD20V2.XML
    Set tValue = inXML.Validate()
    If tValue <> 0 Then
        uDebugPrint "REP: Нарушение структуры XML 80020 > " & tValue.Reason
        Exit Function
    End If
'03 \\ Работа с датой
    Set tNode = inXML.SelectSingleNode("//datetime/day")
    tDate = tNode.Text
    If (tNode Is Nothing) Then
        uDebugPrint "REP: Нарушение структуры XML 80020 > Не удалось получить дату макета!"
        Exit Function
    End If
'04 \\ Извлечение кода торговца (оказалось код AREA не уникальный для множества торговцев, а только для одного торговца)
    Set tNode = inXML.SelectSingleNode("//sender/inn")
    tTraderINN = tNode.Text
    If (tNode Is Nothing) Then
        uDebugPrint "REP: Нарушение структуры XML 80020 > Не удалось получить ИНН отправителя макета!"
        Exit Function
    End If
'05 \\ Перебор AREA
    Set tNode = inXML.SelectNodes("//area")
    For tAreaIndex = 0 To tNode.Length - 1
        If fXML80020AreaReprocess(inFile, tNode(tAreaIndex), tTraderINN, tDate, tGTPName) Then
            fXML80020Reprocess = fXML80020Reprocess + 1
        End If
        'Debug.Print "AR1"
        ' \\ подготовка списка имён ГТП, которым принадлежала AREA
        'If inGTPNameList <> vbNullString Then: inGTPNameList = inGTPNameList & cnstDelimiter
        'inGTPNameList = inGTPNameList & tGTPName
        uAddToList inGTPNameList, tGTPName
        'Debug.Print "AR2"
    Next
'00 \\ Подготовка
'00 \\ Подготовка
'00 \\ Подготовка
'00 \\ Подготовка
End Function

Private Function fXML80020AreaConverter(ByRef inAreaNode)
Dim tNode, tAreaID, tNewAreaID, tSenderINN, tConverterNode, tMPLock, tChLock, tIsConvertable, tIndex, tSubIndex, tMPointCodeA, tMPointCodeB, tChIndex, tChSubIndex, tChCodeA, tChCodeB, tValue, tChIndexList, tElements
Dim tIndexList()
'00 \\ Проверка
    fXML80020AreaConverter = False
    If Not gXMLConverter.Active Then: Exit Function
'01 \\ Извлечение кода AREA
    Set tNode = inAreaNode.SelectSingleNode("inn")
    If (tNode Is Nothing) Then
        uDebugPrint "CONV: Нарушение структуры XML 80020 > Не удалось получить Код AREA!"
        Exit Function
    End If
    tAreaID = tNode.Text
'02 \\ Извлечение ИНН отравителя
    Set tNode = inAreaNode.ParentNode.SelectSingleNode("sender/inn")
    If (tNode Is Nothing) Then
        uDebugPrint "CONV: Нарушение структуры XML 80020 > Не удалось получить ИНН отправителя!"
        Exit Function
    End If
    tSenderINN = tNode.Text
'03 \\ Поиск в конфигурации конвертера текущей AREA удовлетворяющей условиям конвертации
    Set tConverterNode = gXMLConverter.XML.SelectSingleNode("//trader[@id='" & gTraderInfo.ID & "']/source[@inn='" & tSenderINN & "']/area[@sourceid='" & tAreaID & "']")
    If (tConverterNode Is Nothing) Then: Exit Function
    tNewAreaID = tConverterNode.GetAttribute("toid")
    uDebugPrint "CONV: AREA " & tAreaID & " требует конвертации в AREA " & tNewAreaID & "."
'04 \\ Определение возможности конвертации
    tIsConvertable = False
    ReDim tIndexList(tConverterNode.ChildNodes.Length - 1)
    'uDebugPrint "CONV: BASE-CHILDS " & inAreaNode.ChildNodes.Length & "; CONV-CHILDS " & tConverterNode.ChildNodes.Length & "."
    If inAreaNode.ChildNodes.Length > tConverterNode.ChildNodes.Length Then
        For tIndex = 0 To tConverterNode.ChildNodes.Length - 1
            tMPLock = False 'precast
            tMPointCodeA = tConverterNode.ChildNodes(tIndex).GetAttribute("sourcecode") 'CONV measurepoint SOURCEcode read
            For tSubIndex = 0 To inAreaNode.ChildNodes.Length - 1
                If inAreaNode.ChildNodes(tSubIndex).NodeName = "measuringpoint" Then 'only <measuringpoint> childs
                    tMPointCodeB = inAreaNode.ChildNodes(tSubIndex).GetAttribute("code") '80020 measurepoint code read
                    If tMPointCodeA = tMPointCodeB Then
                        'channels compare
                        tMPLock = True
                        tChIndexList = vbNullString
                        'tCHLock = False
                        If tConverterNode.ChildNodes(tIndex).ChildNodes.Length <= inAreaNode.ChildNodes(tSubIndex).ChildNodes.Length Then
                            For tChIndex = 0 To tConverterNode.ChildNodes(tIndex).ChildNodes.Length - 1
                                tChCodeA = tConverterNode.ChildNodes(tIndex).ChildNodes(tChIndex).GetAttribute("sourcecode")
                                tChLock = False
                                For tChSubIndex = 0 To inAreaNode.ChildNodes(tSubIndex).ChildNodes.Length - 1
                                     tChCodeB = inAreaNode.ChildNodes(tSubIndex).ChildNodes(tChSubIndex).GetAttribute("code")
                                     If tChCodeA = tChCodeB Then
                                        tChLock = True
                                        tChIndexList = tChIndexList & cnstDelimiter & tChSubIndex
                                        Exit For
                                     End If
                                Next tChSubIndex
                                'res chan
                                If Not tChLock Then
                                    uDebugPrint "CONV: Точка измерения " & tMPointCodeA & " не имеет канала " & tChCodeA & " в исходной AREA."
                                    Exit For
                                End If
                            Next tChIndex
                        Else
                            tChLock = False
                            uDebugPrint "CONV: Точка измерения " & tMPointCodeA & " должна иметь не менее " & tConverterNode.ChildNodes(tIndex).ChildNodes.Length & " каналов в исходной AREA."
                        End If
                        tIndexList(tIndex) = tSubIndex & tChIndexList 'IndexList of LOCKED childs
                        Exit For 'back to convpoint list
                    End If
                End If
            Next tSubIndex
            If Not tMPLock Then
                uDebugPrint "CONV: Точка измерения " & tMPointCodeA & " не определена в исходной AREA."
                Exit For 'can't lock sourcecode MPoint in current AREA
            End If
            If Not tChLock Then
                tMPLock = False
                Exit For
            End If
        Next tIndex
        tIsConvertable = tMPLock
    End If
    uDebugPrint "CONV: Возможность конвертации > " & tIsConvertable
    If Not tIsConvertable Then: Exit Function
'05 \\ Конвертация AREA
    For tIndex = 0 To tConverterNode.ChildNodes.Length - 1
        '05.01 \\ Изъятие сохранненной позции точки измерения исходной AREA готовой к конвертации
        tElements = Split(tIndexList(tIndex), cnstDelimiter)
        tSubIndex = CLng(tElements(0))
        '05.02 \\ Установка нового КОДА ТИ
        tValue = tConverterNode.ChildNodes(tIndex).GetAttribute("tocode")
        inAreaNode.ChildNodes(tSubIndex).SetAttribute "code", tValue
        '05.03 \\ Установка нового ИМЕНИ ТИ
        tValue = tConverterNode.ChildNodes(tIndex).GetAttribute("toname")
        inAreaNode.ChildNodes(tSubIndex).SetAttribute "name", tValue
        '05.04 \\ Установка новых КОДОВ КАНАЛОВ ТИ
        For tChIndex = 1 To tConverterNode.ChildNodes(tIndex).ChildNodes.Length 'обрати внимание с 1 до Length (т.к. нулевой элемент индекс ТИ)
            tChSubIndex = tElements(tChIndex)
            tValue = tConverterNode.ChildNodes(tIndex).ChildNodes(tChIndex - 1).GetAttribute("tocode") 'tChIndex - 1  ... т.к. мы обращаемся к базовой очередности детишек ТИ (т.е. каналов в конвертере)
            inAreaNode.ChildNodes(tSubIndex).ChildNodes(tChSubIndex).SetAttribute "code", tValue
            tValue = tConverterNode.ChildNodes(tIndex).ChildNodes(tChIndex - 1).GetAttribute("todesc")
            inAreaNode.ChildNodes(tSubIndex).ChildNodes(tChSubIndex).SetAttribute "desc", tValue
        Next
        '05.05 \\ Удаление каналов не подлежащих конвертации
        tValue = inAreaNode.ChildNodes(tSubIndex).ChildNodes.Length - 1
        For tChSubIndex = tValue To 0 Step -1
            tChLock = False
            ' \\ this child is locked?
            For tChIndex = 1 To tConverterNode.ChildNodes(tIndex).ChildNodes.Length
                If tChSubIndex = CInt(tElements(tChIndex)) Then
                    tChLock = True
                    Exit For
                End If
            Next tChIndex
            ' \\ kill child if not locked
            If Not tChLock Then: inAreaNode.ChildNodes(tSubIndex).RemoveChild inAreaNode.ChildNodes(tSubIndex).ChildNodes(tChSubIndex)
        Next tChSubIndex
        '05.06 \\ Уберем из индексации данные индексов каналов
        tIndexList(tIndex) = CLng(tSubIndex)
    Next tIndex
    '05.07 \\ Удаление ТИ не подлежащих конвертации (очистка)
    tValue = inAreaNode.ChildNodes.Length - 1
    For tSubIndex = tValue To 0 Step -1
        'uDebugPrint "CONV: [" & tSubIndex & "]" & inAreaNode.ChildNodes(tSubIndex).NodeName
        If inAreaNode.ChildNodes(tSubIndex).NodeName = "measuringpoint" Then 'Фильтр только по ТИ
            tMPLock = False
            ' \\ this child is locked?
            For tIndex = 0 To tConverterNode.ChildNodes.Length - 1
                If tSubIndex = tIndexList(tIndex) Then
                    tMPLock = True
                    Exit For
                End If
            Next tIndex
            ' \\ kill child if not locked
            If Not tMPLock Then: inAreaNode.RemoveChild inAreaNode.ChildNodes(tSubIndex)
            'uDebugPrint "CONV: [" & tSubIndex & "] KILL? " & Not (tMPLock)
        End If
    Next tSubIndex
    '05.08 \\ Установка нового КОДА AREA
    tValue = tConverterNode.GetAttribute("toid")
    Set tNode = inAreaNode.SelectSingleNode("inn")
    tNode.Text = tValue
    '05.09 \\ Установка нового ИМЕНИ AREA
    tValue = tConverterNode.GetAttribute("toname")
    Set tNode = inAreaNode.SelectSingleNode("name")
    tNode.Text = tValue
'00 \\ Подготовка
    uDebugPrint "CONV: Конвертация завершена."
    fXML80020AreaConverter = True
'00 \\ Подготовка
'00 \\ Подготовка
End Function

Private Sub fReportAdd(inReport, inData)
    If inReport <> vbNullString Then: inReport = inReport & "!#"
    inReport = inReport & inData
End Sub

Private Function fUnPackReport(inReport)
    fUnPackReport = Replace(inReport, "!#", vbCrLf)
End Function

Private Function fCheckAreaFrame(inFile, inTraderID, inAreaNode, inXMLAreaDBNode, inFrameNode, inClass, inReport As TCommonReport) As Boolean
Dim tLogTag, tNode, tLogString, tRepString, tValue, tAreaID, tXPathString, tIndex, tSubIndex, tTempValueA, tTempValueB, tLock, tFrameNode, tLinkListCount, tChLinkListCount, tChIndex, tMPNameA, tMPNameB, tSubLock, tChLock
Dim tLinkList()
Dim tChLinkList()
'01 \\ Проверим готовность FRAME cfg
    'inReport.Reason = vbNullString
    'inReport.ReasonInternal = vbNullString
    'inReport.ReasonShort = vbNullString
    fCheckAreaFrame = False
    tLogTag = "FRMCHK"
    Set inFrameNode = Nothing
    uBDebugPrint tLogTag, tLogString, "Начата проверка структуры AREA."
'02 \\ Определим код AREA
    Set tNode = inAreaNode.SelectSingleNode("inn")
    If (tNode Is Nothing) Then
        inReport.Reason = "Нарушение структуры XML " & inClass & " > Не удалось получить Код AREA!"
        'inReport.Reason = "Внутренняя ошибка автоматики"
        inReport.Status = "Отклонен"
        inReport.Decision = 1
        uADebugPrint tLogTag, inReport.ReasonInternal
        'uBDebugPrint tLogTag, tLogString, "Нарушение структуры XML " & inClass & " > Не удалось получить Код AREA!"
        Exit Function
    End If
    tAreaID = tNode.Text
'03 \\ Поиск AREA в FRAME
    tLock = False
    'get framenode
    tXPathString = "//trader[@id='" & inTraderID & "']/gtp/area[@id='" & tAreaID & "']"
    Set tFrameNode = gXMLFrame.XML.SelectSingleNode(tXPathString)
    If (tFrameNode Is Nothing) Then 'no node
        tLock = True
    ElseIf tFrameNode.ChildNodes.Length = 0 Then 'no childs
        tLock = True
    End If
    'empty?
    If tLock Then
        inReport.ReasonInternal = "FRAME не содержит структуры для AREA " & tAreaID & "! Необходимо заполнить."
        inReport.Reason = "Внутренняя ошибка автоматики"
        inReport.Status = "Обработка приостановлена до решения проблемы"
        inReport.Decision = 2
        uADebugPrint tLogTag, inReport.ReasonInternal
        Exit Function
    End If
'04 \\ Поиск дубликатов ТИ в полученной AREA
    tLock = False
    For tIndex = 0 To inAreaNode.ChildNodes.Length - 2 'Last index can't be forward-scanned
        If inAreaNode.ChildNodes(tIndex).NodeName = "measuringpoint" Then 'filter MP-only
            tValue = 0 'example counter
            tTempValueA = CLngLng(inAreaNode.ChildNodes(tIndex).GetAttribute("code"))
            For tSubIndex = tIndex + 1 To inAreaNode.ChildNodes.Length - 1 'forward-scan runs FOR X > FROM X+1 TO LASTINDEX
                If inAreaNode.ChildNodes(tSubIndex).NodeName = "measuringpoint" Then
                    tTempValueB = CLngLng(inAreaNode.ChildNodes(tSubIndex).GetAttribute("code"))
                    If tTempValueA = tTempValueB Then: tValue = tValue + 1
                End If
            Next tSubIndex
            If tValue > 0 Then
                uBDebugPrint tLogTag, tLogString, "ТИ с кодом " & tTempValueA & " входящей AREA имеет дубликаты(" & tValue & ")"
                fReportAdd inReport.Reason, tLogString
                tLock = True
            End If
        End If
    Next tIndex
    'результат поиска дубликатов
    If tLock Then
        uBDebugPrint tLogTag, tLogString, "Были найдены дубликаты ТИ в структуре входящей AREA " & tAreaID & ". Проверка провалена."
        inReport.Status = "Отклонена"
        inReport.Decision = 1
        'fReportAdd inReport.Reason, tLogString
        Exit Function
    End If
'05 \\ Поиск связей ожидаемой (FRAME) структуры с входящей (AREA) структурой и запись в tLinkList
    tLock = False
    ' \\ Подготовка списка связей
    tLinkListCount = inAreaNode.ChildNodes.Length - 1
    ReDim tLinkList(tLinkListCount)
    For tIndex = 0 To tLinkListCount
        If inAreaNode.ChildNodes(tIndex).NodeName = "measuringpoint" Then
            tLinkList(tIndex) = -1 'normal node
        Else
            tLinkList(tIndex) = -100 'ignoring this node
        End If
    Next tIndex
    ' \\ Поиск связей
    For tIndex = 0 To tFrameNode.ChildNodes.Length - 1
        tValue = -1 'no link precast
        tTempValueA = CLngLng(tFrameNode.ChildNodes(tIndex).GetAttribute("code"))
        For tSubIndex = 0 To inAreaNode.ChildNodes.Length - 1
            If tLinkList(tSubIndex) = -1 Then
                tTempValueB = CLngLng(inAreaNode.ChildNodes(tSubIndex).GetAttribute("code"))
                If tTempValueA = tTempValueB Then
                    tValue = tSubIndex 'link locked
                    Exit For
                End If
            End If
        Next tSubIndex
        If tValue = -1 Then
            tLinkListCount = tLinkListCount + 1
            ReDim Preserve tLinkList(tLinkListCount)
            tLinkList(tLinkListCount) = tIndex
            uBDebugPrint tLogTag, tLogString, "Ожидаемая ТИ " & tTempValueA & " не найдена во входящей AREA!"
            'inReport.Decision = 1
            fReportAdd inReport.Reason, tLogString
            tLock = True
        Else
            tLinkList(tValue) = tIndex 'write link to list
        End If
    Next tIndex
    ' \\ Перечисление ТИ AREA не входящих в состав FRAME (если такие есть)
    For tIndex = 0 To tLinkListCount
        If tLinkList(tIndex) = -1 Then
            tTempValueA = CLngLng(inAreaNode.ChildNodes(tIndex).GetAttribute("code"))
            uBDebugPrint tLogTag, tLogString, "Обнаруженная ТИ " & tTempValueA & " не должна присутствовать во входящей AREA!"
            fReportAdd inReport.Reason, tLogString
            tLock = True
        End If
    Next tIndex
    ' \\ Обработка результата
    If tLock Then
        uBDebugPrint tLogTag, tLogString, "Имеются нарушения структуры ТИ AREA " & tAreaID & ". Проверка провалена."
        inReport.Status = "Отклонена"
        inReport.Decision = 1
        Exit Function
    End If
'06 \\ Поиск связей ожидаемой (FRAME) структуры с входящей (AREA) структурой (ПО КАНАЛАМ)
    tChLock = False
    For tIndex = 0 To tLinkListCount
        If tLinkList(tIndex) <> -100 Then
            tMPNameA = CLngLng(tFrameNode.ChildNodes(tLinkList(tIndex)).GetAttribute("code"))
            tMPNameB = CLngLng(inAreaNode.ChildNodes(tIndex).GetAttribute("code"))
            'Поиск дубликатов
            tLock = False
            For tSubIndex = 0 To inAreaNode.ChildNodes(tIndex).ChildNodes.Length - 2 'Last index can't be forward-scanned
                tValue = 0 'example counter
                tTempValueA = inAreaNode.ChildNodes(tIndex).ChildNodes(tSubIndex).GetAttribute("code")
                For tChIndex = tSubIndex + 1 To inAreaNode.ChildNodes(tIndex).ChildNodes.Length - 1 'forward-scan runs FOR X > FROM X+1 TO LASTINDEX
                    tTempValueB = inAreaNode.ChildNodes(tIndex).ChildNodes(tChIndex).GetAttribute("code")
                    If tTempValueA = tTempValueB Then: tValue = tValue + 1
                Next tChIndex
                If tValue > 0 Then
                    uBDebugPrint tLogTag, tLogString, "Канал ТИ " & tMPNameA & "/" & tTempValueA & " входящей AREA имеет дубликаты(" & tValue & ")"
                    fReportAdd inReport.Reason, tLogString
                    tLock = True
                    tChLock = True
                End If
            Next tSubIndex
            'Если дубликатов нет, то проверим соответствия каналов
            If Not tLock Then
                '//1
                tChLinkListCount = inAreaNode.ChildNodes(tIndex).ChildNodes.Length - 1
                ReDim tChLinkList(tChLinkListCount)
                For tSubIndex = 0 To tChLinkListCount
                    tChLinkList(tSubIndex) = -1
                Next tSubIndex
                '//2
                For tChIndex = 0 To tFrameNode.ChildNodes(tLinkList(tIndex)).ChildNodes.Length - 1
                    tValue = -1 'no link precast
                    tTempValueA = tFrameNode.ChildNodes(tLinkList(tIndex)).ChildNodes(tChIndex).GetAttribute("code")
                    For tSubIndex = 0 To inAreaNode.ChildNodes(tIndex).ChildNodes.Length - 1
                        If tChLinkList(tSubIndex) = -1 Then
                            tTempValueB = inAreaNode.ChildNodes(tIndex).ChildNodes(tSubIndex).GetAttribute("code")
                            If tTempValueA = tTempValueB Then
                                tValue = tSubIndex 'link locked
                                Exit For
                            End If
                        End If
                    Next tSubIndex
                    If tValue = -1 Then
                        tChLinkListCount = tChLinkListCount + 1
                        ReDim Preserve tChLinkList(tChLinkListCount)
                        tChLinkList(tChLinkListCount) = tChIndex
                        uBDebugPrint tLogTag, tLogString, "Ожидаемый канал ТИ " & tMPNameA & "/" & tTempValueA & " не найден во входящей AREA!"
                        fReportAdd inReport.Reason, tLogString
                        tLock = True
                        tChLock = True
                    Else
                        tChLinkList(tValue) = tChIndex 'write link to list
                    End If
                Next tChIndex
                '//3
                For tSubIndex = 0 To tChLinkListCount
                    If tChLinkList(tSubIndex) = -1 Then
                        tTempValueB = inAreaNode.ChildNodes(tIndex).ChildNodes(tSubIndex).GetAttribute("code")
                        uBDebugPrint tLogTag, tLogString, "Обнаруженный канал ТИ " & tMPNameB & "/" & tTempValueB & " не должен присутствовать во входящей AREA!"
                        fReportAdd inReport.Reason, tLogString
                        tLock = True
                        tChLock = True
                    End If
                Next tSubIndex
            End If
        End If
    Next tIndex
    ' \\ Обработка результата
    If tChLock Then
        uBDebugPrint tLogTag, tLogString, "Имеются нарушения структуры ТИ AREA " & tAreaID & ". Проверка провалена."
        inReport.Status = "Отклонена"
        inReport.Decision = 1
        Exit Function
    End If
'00 \\ Подготовка
    Set inFrameNode = tFrameNode
    fCheckAreaFrame = True
    uBDebugPrint tLogTag, tLogString, "Проверка структуры успешно пройдена."
'00 \\ Подготовка
'00 \\ Подготовка
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

Private Function fGetReportList(inClass, inRootNode, inAddress)
Dim tLogTag, tResultList, tExchangeNode, tSilent, tAddress, tEnabled, tClass, tDomain, tCounter, tItemNode, tElementNode
    fGetReportList = vbNullString
    tLogTag = "GETREPORTLIST"
    uCDebugPrint tLogTag, 0, "Формируем список адресов для отправки отчётов.."
    If inRootNode Is Nothing Then
        uCDebugPrint tLogTag, 2, "По непонятным причинам корневая нода на входе недействительна!"
        Exit Function
    End If
    tResultList = vbNullString
    tCounter = 0
' 01 // Поиск дочерней ноды EXCHANGE в корневой ноде
    Set tExchangeNode = fGetChildNodeByName(inRootNode, "exchange")
    If tExchangeNode Is Nothing Then
        uCDebugPrint tLogTag, 1, "Отсутствует блок EXCHANGE в корневой ноде " & UCase(inRootNode.NodeName) & " (конфиг BASIS)!"
        Exit Function
    End If
' 02 // Поиск в EXCHANGE ноды нужного класса
    For Each tItemNode In tExchangeNode.ChildNodes
        tClass = tItemNode.GetAttribute("id")
        If tClass = inClass Then
' 03 // Перебор записей ноды искомого класса
            For Each tElementNode In tItemNode.ChildNodes
                tEnabled = tElementNode.GetAttribute("enabled")
                If tEnabled = "1" Then
                    Select Case LCase(tElementNode.NodeName)
                        'RECIEVE FROM
                        Case "recievefrom":
                            tAddress = tElementNode.GetAttribute("address")
                            tDomain = tElementNode.GetAttribute("domain")
                            If fIsAddressEqual(inAddress, tAddress, tDomain) > 0 Then
                                tSilent = tElementNode.GetAttribute("silent")
                                If tSilent = "0" Then
                                    uAddToListUnique tResultList, inAddress
                                    tCounter = tCounter + 1
                                End If
                            End If
                        'REPORT TO
                        Case "reportto":
                            tAddress = tElementNode.GetAttribute("address")
                            uAddToListUnique tResultList, tAddress
                            tCounter = tCounter + 1
                    End Select
                End If
            Next 'class
        End If
    Next 'exchange
' 04 // Завершение
    fGetReportList = tResultList
    uCDebugPrint tLogTag, 0, "Список адресов сформирован! Адресатов: " & tCounter
End Function

'legal checker
Private Function fLegalSourceCheck(inClass, inRootNode, inAddress)
Dim tItemNode, tEnabled, tAddress, tDomain, tElementNode, tExchangeNode, tClass, tLogTag, tResult
    fLegalSourceCheck = False
    tLogTag = "LEGCHK"
    uCDebugPrint tLogTag, 0, "Проверяем легальность источника информации <" & inAddress & ">"
    If inRootNode Is Nothing Then
        uCDebugPrint tLogTag, 2, "По непонятным причинам корневая нода на входе недействительна!"
        Exit Function
    End If
' 01 // Поиск дочерней ноды EXCHANGE в корневой ноде
    Set tExchangeNode = fGetChildNodeByName(inRootNode, "exchange")
    If tExchangeNode Is Nothing Then
         uCDebugPrint tLogTag, 1, "Проверка не удалась, отсутствует блок EXCHANGE в корневой ноде " & UCase(inRootNode.NodeName) & " (конфиг BASIS)!"
         Exit Function
    End If
' 02 // Поиск в EXCHANGE ноды нужного класса
    For Each tItemNode In tExchangeNode.ChildNodes
        tClass = tItemNode.GetAttribute("id")
        If tClass = inClass Then
' 03 // Перебор записей ноды искомого класса
            For Each tElementNode In tItemNode.ChildNodes
                If tElementNode.NodeName = "recievefrom" Then
                    tEnabled = tElementNode.GetAttribute("enabled")
                    'if enabled
                    If tEnabled = "1" Then
                        tAddress = tElementNode.GetAttribute("address")
                        tDomain = tElementNode.GetAttribute("domain")
                        tResult = fIsAddressEqual(inAddress, tAddress, tDomain)
                        If tResult > 0 Then
                            Select Case tResult
                                Case 1: uCDebugPrint tLogTag, 0, "Адрес легален и допущен к работе."
                                Case 2: uCDebugPrint tLogTag, 0, "Адрес легален и допущен к работе. Разрешено принадлежности домену <" & tDomain & ">."
                            End Select
                            fLegalSourceCheck = True
                            Exit Function
                        End If
                    End If
                End If
            Next 'class
            'fin
            uCDebugPrint tLogTag, 1, "Адрес не допущен к работе."
            Exit Function
        End If
    Next 'exchange
' 04 // Завершнение
    uCDebugPrint tLogTag, 1, "Проверка не удалась, отсутствует элемент класса <" & inClass & "> в блоке EXCHANGE в корневой ноде " & UCase(inRootNode.NodeName) & " (конфиг BASIS)!"
End Function

Private Function fLockChannelNodeInAreaNode(inAreaNode, inMPointCode, inChannelCode)
    Dim tMPointNode, tChanngelNode
    
    'Default
    Set fLockChannelNodeInAreaNode = Nothing
    
    'Preventer
    If inAreaNode Is Nothing Then: Exit Function
    
    'Scanner
    For Each tMPointNode In inAreaNode.ChildNodes
        If tMPointNode.GetAttribute("code") = inMPointCode Then
            For Each tChanngelNode In tMPointNode.ChildNodes
                If tChanngelNode.GetAttribute("code") = inChannelCode Then
                    'Quit with result
                    Set fLockChannelNodeInAreaNode = tChanngelNode
                    Set tMPointNode = Nothing
                    Set tChanngelNode = Nothing
                    Exit Function
                End If
            Next
        End If
    Next
    
    'Quit no result
    Set tMPointNode = Nothing
    Set tChanngelNode = Nothing
End Function

Private Function fOperationByFrame(inAreaNode, inFrameNode)
    Dim tFMPointNode, tFCHNode, tValue, tMPointCode, tCHCode, tOPDescription, tSum, tAPeriodNode, tChannelNode
    Dim tLogTag

'00 \\ Предустановки
    tLogTag = fGetLogTag("fOperationByFrame")
    fOperationByFrame = False ' False - значит операции не было; True - была операция ??? а если ошибка?

'01 \\ если FRAME не доступен - выход
    If inFrameNode Is Nothing Then
        uCDebugPrint tLogTag, 1, "Не обнаружено ноды FRAME для работы с текущей AREA!"
        Exit Function
    End If
    
'02 \\ поиск операторов во FRAME
    For Each tFMPointNode In inFrameNode.ChildNodes
        tMPointCode = tFMPointNode.GetAttribute("code")
        
        For Each tFCHNode In tFMPointNode.ChildNodes
            tCHCode = tFCHNode.GetAttribute("code")
            
            ' Операция 1 \\ ПРИНУДИТЕЛЬНОЕ НУЛЕНИЕ
            tValue = tFCHNode.GetAttribute(fGetOpertaionByEnumID(EFrameOperations.ForceZero))
            If tValue = "1" Then
                tOPDescription = "ПРИНУДИТЕЛЬНОЕ НУЛЕНИЕ"
                ' Поиск кода в AREANODE
                Set tChannelNode = fLockChannelNodeInAreaNode(inAreaNode, tMPointCode, tCHCode)
                If Not tChannelNode Is Nothing Then
                    tSum = 0
                    For Each tAPeriodNode In tChannelNode.ChildNodes
                        If IsNumeric(tAPeriodNode.FirstChild.Text) Then: tSum = tSum + CLng(tAPeriodNode.FirstChild.Text)
                        tAPeriodNode.FirstChild.Text = 0
                    Next
                    uCDebugPrint tLogTag, 1, "Канал " & tMPointCode & "[" & tCHCode & "] подвергнут операции: " & tOPDescription & " (" & tSum & " >> 0)"
                End If
            End If
            
            ' Операция 2 \\ ПРИНУДИТЕЛЬНАЯ НЕКОММЕРЦИЯ
            tValue = tFCHNode.GetAttribute(fGetOpertaionByEnumID(EFrameOperations.ForceUncom))
            If tValue = "1" Then
                tOPDescription = "ПРИНУДИТЕЛЬНАЯ НЕКОММЕРЦИЯ"
                ' Поиск кода в AREANODE
                Set tChannelNode = fLockChannelNodeInAreaNode(inAreaNode, tMPointCode, tCHCode)
                If Not tChannelNode Is Nothing Then
                    tSum = 0
                    For Each tAPeriodNode In tChannelNode.ChildNodes
                        If IsNumeric(tAPeriodNode.FirstChild.Text) Then: tSum = tSum + CLng(tAPeriodNode.FirstChild.Text)
                        tAPeriodNode.FirstChild.SetAttribute "status", 1 'Child of PERIOD is VALUE node, so it SETs STATUS = 1 to VALUE first(can't be more) node
                    Next
                    uCDebugPrint tLogTag, 1, "Канал " & tMPointCode & "[" & tCHCode & "] подвергнут операции: " & tOPDescription & " (Объёмы подверженные операции: " & tSum & ")"
                End If
            End If
            
            ' Операция 3 \\ ПРИНУДИТЕЛЬНАЯ КОММЕРЦИЯ
            tValue = tFCHNode.GetAttribute(fGetOpertaionByEnumID(EFrameOperations.ForceCom))
            If tValue = "1" Then
                tOPDescription = "ПРИНУДИТЕЛЬНАЯ КОММЕРЦИЯ"
                ' Поиск кода в AREANODE
                Set tChannelNode = fLockChannelNodeInAreaNode(inAreaNode, tMPointCode, tCHCode)
                If Not tChannelNode Is Nothing Then
                    tSum = 0
                    For Each tAPeriodNode In tChannelNode.ChildNodes
                        If IsNumeric(tAPeriodNode.FirstChild.Text) Then: tSum = tSum + CLng(tAPeriodNode.FirstChild.Text)
                        tAPeriodNode.FirstChild.SetAttribute "status", 0 'Child of PERIOD is VALUE node, so it SETs STATUS = 0 to VALUE first(can't be more) node
                    Next
                    uCDebugPrint tLogTag, 1, "Канал " & tMPointCode & "[" & tCHCode & "] подвергнут операции: " & tOPDescription & " (Объёмы подверженные операции: " & tSum & ")"
                End If
            End If
            
            ' Конец операций
        Next
    Next
End Function

Private Function fGetIndexByStartTime(inStartTime)
Dim tHour, tMin
    fGetIndexByStartTime = -1
    If Len(inStartTime) = 4 Then
        If IsNumeric(inStartTime) Then
            tHour = CInt(Left(inStartTime, 2))
            tMin = CInt(Right(inStartTime, 2))
            fGetIndexByStartTime = tHour * 2
            If tMin = 30 Then: fGetIndexByStartTime = fGetIndexByStartTime + 1
        End If
    End If
End Function

Function fGetStartTimeByIndex(inIndex)
Dim tHour, tMin
    fGetStartTimeByIndex = -1
    If IsNumeric(inIndex) Then
        tHour = inIndex \ 2
        tMin = (tHour * 2 < inIndex)
        If tHour < 10 Then
            fGetStartTimeByIndex = "0" & tHour
        Else
            fGetStartTimeByIndex = tHour
        End If
        If tMin Then
            fGetStartTimeByIndex = fGetStartTimeByIndex & "30"
        Else
            fGetStartTimeByIndex = fGetStartTimeByIndex & "00"
        End If
    End If
End Function

'fAreaAnalyzer - проверка AREA на состав точек и каналов, а так же некоммерцию
Private Function fAreaAnalyzer(inAreaNode, inFrameNode, inClass, outUnComDetected, ioReport As TCommonReport)
    Dim tValue, AValueNode, tStatus, tFMPNode, tFCHNode, tIgnoreUncom, tIndex, tAPeriodNode, tChLogLine, tPreviousValue, tStartIndex, tUnComDetected, tUnComStop, tLogTag, tLogString
    Dim tUnComArray(48)
    Dim tUnComValueArray(48)
    Dim tAMPNodes, tAMPNode, tXPathString, tAMPCode, tAMPName, tACHNodes, tACHNode, tACHCode, tFrameNodeExists, tValueSum, tTotalValueSum
    
'01 \\ Предопределения
    fAreaAnalyzer = True
    tLogTag = fGetLogTag("fAreaAnalyzer")
    tUnComDetected = False
    tUnComStop = False
    outUnComDetected = False
    tIgnoreUncom = 0 'default ignore uncom status
    
    'tOperAlias_IgnoreUncom = "ignore_uncom" 'just for ignoring
    'tOperAlias_ForceUncom = "op_forceuncom" 'for CHANGING data -> will remove uncom status
    
'02 \\ Проверка входных данных
    If TypeName(inAreaNode) <> "IXMLDOMElement" Then
        uCDebugPrint tLogTag, 2, "Входная нода inAreaNode не является нодой!"
        Exit Function
    End If
        
    'is frame node exists?
    If inFrameNode Is Nothing Then
        tFrameNodeExists = False
    Else
        tFrameNodeExists = (TypeName(inFrameNode) = "IXMLDOMElement")
    End If
    
    'read frame OPERATIONS status to whole AREA
    If tFrameNodeExists Then
        If inFrameNode.GetAttribute(fGetOpertaionByEnumID(EFrameOperations.IgonreUncom)) = "1" Or inFrameNode.GetAttribute(fGetOpertaionByEnumID(EFrameOperations.ForceUncom)) = "1" Then
            tIgnoreUncom = 30 'high status (FOR WHOLE AREA)
        End If
    End If
        
'03 \\ Выборка AMP нод (Area Measuring Point = AMP)
    tXPathString = "child::measuringpoint"
    Set tAMPNodes = inAreaNode.SelectNodes(tXPathString)
    
'04 \\ Переборка найденых AMP нод // measuringpoint
    For Each tAMPNode In tAMPNodes
    
        'read header attributes
        tAMPCode = tAMPNode.GetAttribute("code")
        tAMPName = tAMPNode.GetAttribute("name")
        
        'get data from frame nodes // measuringpoint
        If tFrameNodeExists And tIgnoreUncom < 30 Then
            Set tFMPNode = inFrameNode.SelectSingleNode("child::measuringpoint[@code='" & tAMPCode & "']")
            If tFMPNode.GetAttribute(fGetOpertaionByEnumID(EFrameOperations.IgonreUncom)) = "1" Or tFMPNode.GetAttribute(fGetOpertaionByEnumID(EFrameOperations.ForceUncom)) = "1" Then
                tIgnoreUncom = 20
            End If
        End If
        
        'prepare nodes for channel scan
        tXPathString = "child::measuringchannel"
        Set tACHNodes = tAMPNode.SelectNodes(tXPathString)
        
'05 \\ Переборка найденых ACH нод // measuringchannel
        For Each tACHNode In tACHNodes
            
            'read header attributes // measuringchannel
            tACHCode = tACHNode.GetAttribute("code")
            
            'get data from frame nodes // measuringchannel
            If tFrameNodeExists And tIgnoreUncom < 20 Then
                Set tFCHNode = tFMPNode.SelectSingleNode("child::measuringchannel[@code='" & tACHCode & "']")
                If tFCHNode.GetAttribute(fGetOpertaionByEnumID(EFrameOperations.IgonreUncom)) = "1" Or tFCHNode.GetAttribute(fGetOpertaionByEnumID(EFrameOperations.ForceUncom)) = "1" Then
                    tIgnoreUncom = 10
                End If
            End If
            
'06 \\ Переборка данных содержащихся в ACH нодах
            
            '06.01 \\ Подготовка массива маркеров некоммерции (status = 1)
            tUnComArray(48) = False 'fake element for finishing with FALSE anyway \\ 48+1
             tUnComValueArray(48) = 0
            For tIndex = 0 To 47
                tUnComArray(tIndex) = True
                tUnComValueArray(tIndex) = -1
            Next
            
            '06.02 \\ Поиск некоммерции в данных текущего канала tACHNode + чтение самих значений
            For Each tAPeriodNode In tACHNode.ChildNodes
                tIndex = fGetIndexByStartTime(tAPeriodNode.GetAttribute("start"))
                If tIndex > -1 Then
                
                    'status
                    tValue = tAPeriodNode.FirstChild.GetAttribute("status")
                    If (tValue = "0") Or IsNull(tValue) Then: tUnComArray(tIndex) = False
                    
                    'value
                    tValue = tAPeriodNode.FirstChild.Text
                    If Not IsNull(tValue) Then
                        If IsNumeric(tValue) Then
                            tUnComValueArray(tIndex) = Fix(tValue)
                        End If
                    End If
                End If
            Next
            
            '06.03 // Подготовка строки отчета
            tChLogLine = vbNullString
            tTotalValueSum = 0
            tPreviousValue = False 'для составления отрезков, надо знать предыдущее значение
                
            For tIndex = 0 To 48 '48 coz we need finisher (real data 0..47)
                If tUnComArray(tIndex) And (Not tPreviousValue) Then 'starting uncom part
                    tStartIndex = tIndex
                    tValueSum = 0
                ElseIf (Not tUnComArray(tIndex)) And tPreviousValue Then 'ending uncom part
                    If tChLogLine <> vbNullString Then: tChLogLine = tChLogLine & ", " 'splitter for visual adapt
                    tChLogLine = tChLogLine & fGetStartTimeByIndex(tStartIndex) & "-" & fGetStartTimeByIndex(tIndex) & "[" & tValueSum & "]"
                    tTotalValueSum = tTotalValueSum + tValueSum 'for whole uncom status items
                End If
                
                'accumulate value
                If tUnComArray(tIndex) Then: tValueSum = tValueSum + tUnComValueArray(tIndex)
                    
                'save current value for next iteration
                tPreviousValue = tUnComArray(tIndex)
            Next
            
            '06.04 // Отработка отчета (при наличии)
            If tChLogLine <> vbNullString Then
                tUnComDetected = True
            
                'should we ignore uncom?
                If tIgnoreUncom > 0 Then
                    tChLogLine = tChLogLine & " (РАЗРЕШЕНО)"
                Else
                    tUnComStop = True
                End If
                                
                tChLogLine = "НеКОМ > " & tAMPCode & "\" & tACHCode & "(" & tAMPName & "):" & tChLogLine
                uC2DebugPrint tLogTag, 1, tLogString, tChLogLine
                fReportAdd ioReport.Reason, tLogString
            End If
            
            'restore status (exclude higher level)
            If tIgnoreUncom < 20 Then: tIgnoreUncom = 0
        Next ' // ACH end
        
        'restore status (exclude higher level)
        If tIgnoreUncom < 30 Then: tIgnoreUncom = 0
    Next ' // AMP end
    
'XX \\ Завершение анализа
    fAreaAnalyzer = tUnComStop
    outUnComDetected = tUnComDetected
End Function

Private Sub fAreaStatusChange(inXMLNode, inNewStatus, inReport As TCommonReport)
Dim tNode, tIndex, tLockIndex
    ' 01 // Status update
    inXMLNode.SetAttribute "status", inNewStatus
    ' 02 // Report work
    Set tNode = fGetChildNodeByName(inXMLNode, "report", False)
    If Not (tNode Is Nothing) Then: inXMLNode.RemoveChild (tNode)
    If inReport.Reason <> vbNullString Then
        inXMLNode.SetAttribute "report", inReport.Reason
    End If
End Sub

Private Function fReportExpose(inReport As TCommonReport, Optional inMode = 0) As String
    Select Case inMode
        Case 0: fReportExpose = inReport.Object & ": " & inReport.Status & " (" & inReport.ReasonShort & ")"
        Case 1: fReportExpose = inReport.Object & ": " & inReport.Status & " (" & inReport.Reason & ")"
        Case 2:
            fReportExpose = "Источник: " & inReport.Source & vbCrLf
            If inReport.Owner <> vbNullString Then: fReportExpose = fReportExpose & "Получен от: " & inReport.Owner & vbCrLf
            If inReport.RecievedTimeStamp <> vbNullString Then: fReportExpose = fReportExpose & "Получен: " & inReport.RecievedTimeStamp & vbCrLf
            If inReport.ProcessedTimeStamp <> vbNullString Then: fReportExpose = fReportExpose & "Обработан: " & inReport.ProcessedTimeStamp & vbCrLf
            If inReport.SourceClass <> vbNullString Then: fReportExpose = fReportExpose & "Класс источника: " & inReport.SourceClass & vbCrLf
            If inReport.Object <> vbNullString Then: fReportExpose = fReportExpose & "Объект: " & inReport.Object & vbCrLf
            If inReport.Date <> vbNullString Then: fReportExpose = fReportExpose & "Дата: " & inReport.Date & vbCrLf
            fReportExpose = fReportExpose & "Статус: " & inReport.Status & vbCrLf & vbCrLf & "Причина:" & vbCrLf & fUnPackReport(inReport.Reason)
    End Select
End Function

Private Function fXML80020AreaReprocess(inFile, inAreaNode, inTraderINN, inDate, inGTPName) As Boolean
Dim tNow, tWorkDayLimit, tNode, tTempNode, tTraderID, tTraderNode, tAreaID, tVersionNode, tGTPID, tLinkGTPID, tVersion, tStatus, tSectionNode, tMainNode, tNewNumber, tXML80020Node, tFrameNode, tClass, tCheckStatus, tCheckReport, tResultString, tLogTag, tReportList, tValue, tErrorText, tTempDate, tUnComDetected
Dim tReport As TCommonReport
    fXML80020AreaReprocess = False
    tClass = "80020"
    tWorkDayLimit = 3
    inGTPName = vbNullString
    tResultString = vbNullString
    tLogTag = "REP"
    tReportList = vbNullString
'01 \\ Прохождение AREA через CONVERTOR
    fXML80020AreaConverter inAreaNode
'02 \\ Извлечение кода AREA
    Set tNode = inAreaNode.SelectSingleNode("inn")
    If (tNode Is Nothing) Then
        uADebugPrint tLogTag, "Нарушение структуры XML " & tClass & " > Не удалось получить Код AREA!"
        Exit Function
    End If
    tAreaID = tNode.Text
    Set tTempNode = gXMLBasis.XML.SelectSingleNode("//trader[@inn='" & inTraderINN & "']")
    If tTempNode Is Nothing Then
        tTraderID = "UNKNOWN (INN=" & inTraderINN & ")"
    Else
        tTraderID = tTempNode.GetAttribute("id")
    End If
    ' \\ REPORT Init
    tReport.Object = "AREA " & tAreaID & " (" & tTraderID & ")"
    tReport.Source = inFile.Name
    tReport.RecievedTimeStamp = fGetRecievedTimeStamp(gCurrentMessage.Recieved, 2, gLocalUTC, 3)
    tReport.ProcessedTimeStamp = fGetRecievedTimeStamp(Now(), 2, gLocalUTC, 3)
    tReport.Owner = gCurrentMessage.SenderEMail
    'SetSource
    tReport.Date = inDate
    tReport.SourceClass = tClass
    tReport.Status = vbNullString
    tReport.Reason = vbNullString
    tReport.ReasonShort = vbNullString
    tReport.ReasonInternal = vbNullString
    tReport.Decision = 0 '0 - accept; 1 - reject; 2 -  to stack (internal error); 3 - to manual (problem need to be solved by manual)
    'tReport.
'03 \\ Поиск AREA в конфиругации
    Set tVersionNode = gXMLBasis.XML.SelectSingleNode("//trader[@id='" & tTraderID & "']/gtp/section/version[area[@id='" & tAreaID & "' and @type='1']]")
    If (tVersionNode Is Nothing) Then
        tReport.Status = "Отклонена"
        tReport.ReasonShort = "Не найдена в конфиге BASIS"
        tReport.Decision = 1 'reject
        uADebugPrint tLogTag, fReportExpose(tReport)
        Exit Function
    End If
'04 \\ Чтение данных из BASIS
    Set tSectionNode = tVersionNode.ParentNode
    Set tMainNode = tSectionNode.ParentNode
    Set tTraderNode = tMainNode.ParentNode
    'tTraderID = tTraderNode.GetAttribute("id")
    tGTPID = tMainNode.GetAttribute("id")
    inGTPName = tGTPID 'передача имени ГТП вышестоящей процедуре (это для создания папок в аутлуке)
    tLinkGTPID = tSectionNode.GetAttribute("id")
    tVersion = tVersionNode.GetAttribute("id")
    tStatus = tVersionNode.GetAttribute("status")
    tReport.Object = tReport.Object & " (" & tTraderID & " : переток " & tGTPID & "-" & tLinkGTPID & " v" & tVersion & ")"
'05.01 \\ Проверка на легитимность отправителя данной AREA
    'fGetReportList tReportList
    If Not fLegalSourceCheck(tClass, tVersionNode, gCurrentMessage.SenderEMail) Then
        tReport.Reason = "Адрес источника данных не является легальным! В обработке отказано"
        tReport.Decision = 1 'reject
        tReport.Status = "Отклонена"
        tReport.ReasonShort = "Нелегальный источник данных"
        uADebugPrint tLogTag, fReportExpose(tReport, 1)
        Exit Function
    End If
'05.02 \\ Сбор списка адресов для отчёта
    tReportList = fGetReportList(tClass, tVersionNode, gCurrentMessage.SenderEMail)
'06 \\ Проверка статуса AREA в BASIS
    If tStatus = "closed" Then
        tValue = tVersionNode.GetAttribute("closed")
        tReport.Status = "Отклонена"
        tReport.ReasonShort = "Не актуальна"
        tReport.Reason = "AREA закрыта для приёма. Срок актуальности истёк " & tValue & "."
        tReport.Decision = 1 'reject
        uADebugPrint tLogTag, fReportExpose(tReport, 1)
        fFastRespond tReportList, tReport.Source & ":Rejected", fReportExpose(tReport, 2) 'gCurrentMessage.SenderEMail
        Exit Function
    End If
'07 \\ Проверка даты входящего макета
    tNow = CLng(Format(Now(), "YYYYMMDD"))
    'If Not (tNow <= fWorkDayShift(inDate, tWorkDayLimit)) Then 'A > не более 3х рабочих дней
    If Not fWorkDayShiftAdv(inDate, tWorkDayLimit, 1, tTempDate, tErrorText) Then
        tReport.Reason = "Ошибка работы с датой: " & tErrorText
        tReport.Decision = 1 'reject
    Else
        If Not (tNow <= CLng(tTempDate)) Then
            tReport.Reason = "Прошло больше " & tWorkDayLimit & " рабочих дней; или ошибка даты"
            tReport.Decision = 1 'reject
        ElseIf inDate - tNow >= 0 Then 'B > дата не может быть старше, чем сегодня (физически неисполнимо)
            tReport.Reason = "Дата объекта " & inDate & " не может быть старше даты приёма " & tNow
            tReport.Decision = 1 'reject
        End If
    End If
    If tReport.Decision = 1 Then
        tReport.Status = "Отклонена"
        tReport.ReasonShort = "Не актуальна по дате"
        uADebugPrint tLogTag, fReportExpose(tReport, 1)
        fFastRespond tReportList, tReport.Source & ":OutDated", fReportExpose(tReport, 2) 'gCurrentMessage.SenderEMail
        Exit Function
    End If
'08 \\ Проверка БД 80020 на наличие записи для данной AREA и выдача порядкового номера для извлекаемой AREA
    fSplitCheckArea inFile, inAreaNode, inDate, tVersionNode, tAreaID, tNewNumber, tXML80020Node, tReport
    If tNewNumber = 0 Then
        fReloadXML gXML80020DB.XML, gXML80020DB.Path 'Отмена изменений в файле
        uADebugPrint tLogTag, fReportExpose(tReport, 1)
        fFastRespond tReportList, tReport.Source & ":Rejected", fReportExpose(tReport, 2)
        Exit Function
    End If
'09 \\ Получив запись в БД присвоим текущий статус "0"
    fAreaStatusChange tXML80020Node, 0, tReport
'10 \\ Проверка структуры AREA
    If Not (fCheckAreaFrame(inFile, tTraderID, inAreaNode, tXML80020Node, tFrameNode, tClass, tReport)) Then
        fReloadXML gXML80020DB.XML, gXML80020DB.Path 'Отмена изменений в файле
        uADebugPrint tLogTag, fReportExpose(tReport, 1)
        fFastRespond tReportList, tReport.Source & ":Rejected", fReportExpose(tReport, 2)
        Exit Function
    End If
'11 \\ Операции над AREA
    fOperationByFrame inAreaNode, tFrameNode
'12 \\ Поиск некоммерческой информации
    If fAreaAnalyzer(inAreaNode, tFrameNode, tClass, tUnComDetected, tReport) Then
        fAreaStatusChange tXML80020Node, 1, tReport 'CHANGE STATUS to 1
        'fSaveXMLChanges gXML80020DB.XML, gXML80020DB.Path, True 'Сохранение в БД
        fSaveXMLDB gXML80020DB, False, , , , tLogTag & " нашла некоммерцию и завершила анализ!"
        tReport.Status = "Задержана"
        tReport.ReasonShort = "Содержит некоммерческую информацию"
        uADebugPrint tLogTag, fReportExpose(tReport, 0)
        'tResultString = "Ошибка:" & vbCrLf & fUnPackReport(tCheckReport) & vbCrLf & vbCrLf & tResultString
        fFastRespond tReportList, tReport.Source & ":UnComm", fReportExpose(tReport, 2) 'gCurrentMessage.SenderEMail
        Exit Function
    ElseIf tUnComDetected Then
        'uADebugPrint tLogTag, fReportExpose(tReport, 0)
        tReport.Status = "Принята (с условиями)"
        tReport.ReasonShort = "Содержит разрешенную некоммерческую информацию"
        fFastRespond tReportList, tReport.Source & ":UnComm_OK", fReportExpose(tReport, 2) 'gCurrentMessage.SenderEMail
    End If
'13 \\ Выделение AREA в отдельный файл и регистрация изменений в БД 80020
    If Not fSplitDropArea(inFile, inAreaNode, inDate, tVersionNode, tAreaID, tNewNumber, tXML80020Node, tReport) Then
        'uADebugPrint tLogTag, "AREA " & tAreaID & " (переток " & tGTPID & "-" & tLinkGTPID & " v" & tVersion & ") отклонён."
        Exit Function
    End If
    'fSaveXMLChanges gXML80020DB.XML, gXML80020DB.Path, True
    fSaveXMLDB gXML80020DB, False, , , , tLogTag & " завершила анализ!"
'14 \\ Подготовка
'00 \\ Подготовка
    uADebugPrint tLogTag, tReport.Object & " принят."
    'fSaveXMLChanges gXML80020DB.XML, gXML80020DB.Path
    fXML80020AreaReprocess = True
End Function

Private Function fFastRespond(inAddressList, inHeader, inBody, Optional inSign = vbNullString, Optional inPicturePath = vbNullString, Optional inAttachmentPath = vbNullString)
    Dim tAutoSign, tLogTag, tPictureCode, tCIDCode, tBody
    Dim tOutMail As Outlook.MailItem
    Dim tPAccessor As Outlook.PropertyAccessor
    Const PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

    tLogTag = fGetLogTag("FASTRESP")
    If Not (cnstFastRespond) Then: Exit Function 'Preventor #1
    If inAddressList = vbNullString Then: Exit Function 'Preventor #2

    'Debug.Print "FF1"
    'On Error Resume Next
        ' Тело письма
        tBody = inBody
        tBody = Replace(tBody, "<", "&lt;") 'to HTMP
        tBody = Replace(tBody, ">", "&gt;") 'to HTMP
        tBody = Replace(tBody, vbCrLf, "<br>") 'to HTMP
        '<font face="Arial"></font>
        
        ' Подпись (если на входе пусто - подпись по умолчанию)
        If inSign = vbNullString Then
            tAutoSign = "<br><br>// Данное сообщение - ответ автоматической системы проверки"
        Else
            tAutoSign = "<br><br>" & Replace(inSign, vbCrLf, "<br>")
        End If
    
    'Debug.Print "FF2"
        ' Картинка встроенная в тело письма
        If inPicturePath <> vbNullString Then
            If gFSO.FileExists(inPicturePath) Then
                tCIDCode = uGetFileName(inPicturePath)
                tPictureCode = "<br><img src=""cid:" & tCIDCode & """><br>"
                If InStr(tBody, "##PIX##") > 0 Then
                    tBody = Replace(tBody, "##PIX##", tPictureCode)
                Else
                    tBody = tBody & tPictureCode
                End If
            End If
        End If
    
    'Debug.Print "FF3"
        ' Слияние с подписью
        tBody = tBody & tAutoSign
        
        ' Шрифтование
        tBody = "<font face=""Calibri"">" & tBody & "</font>"
        
        ' Создание письма
    'Debug.Print "FF4"
        Set tOutMail = Outlook.Application.CreateItem(olMailItem) 'Outlook.Application.CreateItem(0)
    
    'Debug.Print "FF5"
        ' Заполним поля письма
        With tOutMail
            .SendUsingAccount = fGetMailAccount()
            .To = inAddressList
            .CC = ""
            .BCC = ""
            .Subject = "ROBOT:" & inHeader
            
            ' Картинка
            If inPicturePath <> vbNullString Then
                .Attachments.Add (inPicturePath)
                Set tPAccessor = .Attachments.Item(.Attachments.Count).PropertyAccessor
                tPAccessor.SetProperty PR_ATTACH_CONTENT_ID, tCIDCode
            End If
            
            ' Вложения без встройки в тело
            If inAttachmentPath <> vbNullString Then: .Attachments.Add inAttachmentPath
            
            .HTMLBody = tBody
           
            .Send   'or use .Display
        End With
        
    'Debug.Print "FF6"
        ' Обработка ошибок
        If Err.Number <> 0 Then
            uCDebugPrint tLogTag, 2, "ERROR > " & Err.Description
        End If
    'On Error GoTo 0
    'Debug.Print "FF7"
    ' Уничтожение объектов
    Set tOutMail = Nothing
    Set tPAccessor = Nothing
    'Debug.Print "FF8"
End Function

Private Sub fReloadXML(inXML, inPath)
    inXML.Load inPath
End Sub

Private Sub fSplitCheckArea(inFile, inAreaNode, inDate, inVersionNode, tAreaID, inNewNumber, inXML80020Node, inReport As TCommonReport)
Dim tNode, tYear, tMonth, tDay, tTempNode, tNumber, tValue, tRootNode, tTraderID, tAIISCode, tRootString, tNodeString, tMainXPath, tInNum, tLogTag, tCurrentTime
'00 \\ Подготовка
    inNewNumber = 0
    'inReport.ReasonInternal = vbNullString
    'inReport.Reason = vbNullString
    'inReport.Status = vbNullString
    tLogTag = "SCA"
    tCurrentTime = CDate(Now())
    Set inXML80020Node = Nothing
'01 \\ Извлечение даты
    If Not (IsTimeStamp(inDate, tYear, tMonth, tDay)) Then
        inReport.ReasonInternal = "Не удалось определить дату из [" & inDate & "]."
        inReport.Reason = "Внутренняя ошибка автоматики"
        inReport.Status = "Обработка приостановлена до решения проблемы"
        inReport.Decision = 2
        uADebugPrint tLogTag, inReport.ReasonInternal
        Exit Sub
    End If
'02 \\ Вытяжка номера входящего макета 80020 и TRADER ID из BASIS
    Set tTempNode = inAreaNode.ParentNode
    tNumber = CLng(tTempNode.GetAttribute("number"))
    Set tTempNode = inVersionNode.ParentNode.ParentNode.ParentNode 'Переход на ноду TRADER конфига BASIS по текущей AREA (-3)
    tTraderID = tTempNode.GetAttribute("id")
    Set tTempNode = inVersionNode.ParentNode.ParentNode 'Переход на ноду GTP конфига BASIS по текущей AREA (-2)
    tAIISCode = tTempNode.GetAttribute("aiiscode")
    tMainXPath = "/message/trader[@id='" & tTraderID & "']/year[@id='" & tYear & "']/month[@id='" & tMonth & "']/day[@id='" & tDay & "']/aiis[@id='" & tAIISCode & "']/area[@id='" & tAreaID & "' and @class='80020']"
'03 \\ Поиск AREA в базе данных полученных AREA
    Set tNode = gXML80020DB.XML.SelectNodes(tMainXPath)
    If tNode.Length = 0 Then 'Создание первичной записи (ПРОХОД ПО УРОВНЯМ)
        '03.1 \\ Уровень TRADER
        tRootString = "/message"
        tNodeString = tRootString & "/trader[@id='" & tTraderID & "']"
        Set tRootNode = gXML80020DB.XML.SelectSingleNode(tRootString)
        Set tNode = gXML80020DB.XML.SelectNodes(tNodeString)
        If tNode.Length = 0 Then
            Set tNode = tRootNode.AppendChild(gXML80020DB.XML.CreateElement("trader"))
            tNode.SetAttribute "id", tTraderID 'ID
            Set tTempNode = inVersionNode.ParentNode.ParentNode.ParentNode
            tValue = tTempNode.GetAttribute("name")
            tNode.SetAttribute "name", tValue 'NAME
            tValue = tTempNode.GetAttribute("inn")
            tNode.SetAttribute "inn", tValue 'INN
        ElseIf tNode.Length > 1 Then
            inReport.ReasonInternal = "БД " & gXML80020DB.ClassTag & " > нарушение структуры уровня TRADER. Элементов " & tTraderID & " более 1."
            inReport.Reason = "Внутренняя ошибка автоматики"
            inReport.Status = "Обработка приостановлена до решения проблемы"
            inReport.Decision = 2
            uADebugPrint tLogTag, inReport.ReasonInternal
            Exit Sub
        End If
        '03.2 \\ Уровень YEAR
        tRootString = tNodeString
        tNodeString = tRootString & "/year[@id='" & tYear & "']"
        Set tRootNode = gXML80020DB.XML.SelectSingleNode(tRootString)
        Set tNode = gXML80020DB.XML.SelectNodes(tNodeString)
        If tNode.Length = 0 Then
            Set tNode = tRootNode.AppendChild(gXML80020DB.XML.CreateElement("year"))
            tNode.SetAttribute "id", tYear 'ID
        ElseIf tNode.Length > 1 Then
            inReport.ReasonInternal = "БД " & gXML80020DB.ClassTag & " > нарушение структуры уровня YEAR. Элементов " & tYear & " более 1."
            inReport.Reason = "Внутренняя ошибка автоматики"
            inReport.Status = "Обработка приостановлена до решения проблемы"
            inReport.Decision = 2
            uADebugPrint tLogTag, inReport.ReasonInternal
            Exit Sub
        End If
        '03.3 \\ Уровень MONTH
        tRootString = tNodeString
        tNodeString = tRootString & "/month[@id='" & tMonth & "']"
        Set tRootNode = gXML80020DB.XML.SelectSingleNode(tRootString)
        Set tNode = gXML80020DB.XML.SelectNodes(tNodeString)
        If tNode.Length = 0 Then
            Set tNode = tRootNode.AppendChild(gXML80020DB.XML.CreateElement("month"))
            tNode.SetAttribute "id", tMonth 'ID
        ElseIf tNode.Length > 1 Then
            inReport.ReasonInternal = "БД " & gXML80020DB.ClassTag & " > нарушение структуры уровня MONTH. Элементов " & tMonth & " более 1."
            inReport.Reason = "Внутренняя ошибка автоматики"
            inReport.Status = "Обработка приостановлена до решения проблемы"
            inReport.Decision = 2
            uADebugPrint tLogTag, inReport.ReasonInternal
            Exit Sub
        End If
        '03.4 \\ Уровень DAY
        tRootString = tNodeString
        tNodeString = tRootString & "/day[@id='" & tDay & "']"
        Set tRootNode = gXML80020DB.XML.SelectSingleNode(tRootString)
        Set tNode = gXML80020DB.XML.SelectNodes(tNodeString)
        If tNode.Length = 0 Then
            Set tNode = tRootNode.AppendChild(gXML80020DB.XML.CreateElement("day"))
            tNode.SetAttribute "id", tDay 'ID
        ElseIf tNode.Length > 1 Then
            inReport.ReasonInternal = "БД " & gXML80020DB.ClassTag & " > нарушение структуры уровня DAY. Элементов " & tDay & " более 1."
            inReport.Reason = "Внутренняя ошибка автоматики"
            inReport.Status = "Обработка приостановлена до решения проблемы"
            inReport.Decision = 2
            uADebugPrint tLogTag, inReport.ReasonInternal
            Exit Sub
        End If
        '03.5 \\ Уровень AIIS
        tRootString = tNodeString
        tNodeString = tRootString & "/aiis[@id='" & tAIISCode & "']"
        Set tRootNode = gXML80020DB.XML.SelectSingleNode(tRootString)
        Set tNode = gXML80020DB.XML.SelectNodes(tNodeString)
        If tNode.Length = 0 Then
            Set tNode = tRootNode.AppendChild(gXML80020DB.XML.CreateElement("aiis"))
            tNode.SetAttribute "id", tAIISCode 'ID
            Set tTempNode = inVersionNode.ParentNode.ParentNode
            tValue = tTempNode.GetAttribute("id")
            tNode.SetAttribute "gtpid", tValue 'GTPID
            tNode.SetAttribute "number", 0 'NUMBER
        ElseIf tNode.Length > 1 Then
            inReport.ReasonInternal = "БД " & gXML80020DB.ClassTag & " > нарушение структуры уровня AIIS. Элементов " & tAIISCode & " более 1."
            inReport.Reason = "Внутренняя ошибка автоматики"
            inReport.Status = "Обработка приостановлена до решения проблемы"
            inReport.Decision = 2
            uADebugPrint tLogTag, inReport.ReasonInternal
            Exit Sub
        End If
        '03.6 \\ Уровень AREA
        tRootString = tNodeString
        tNodeString = tRootString & "/area[@id='" & tAreaID & "']"
        Set tRootNode = gXML80020DB.XML.SelectSingleNode(tRootString)
        Set tNode = gXML80020DB.XML.SelectNodes(tNodeString)
        If tNode.Length = 0 Then
            Set tNode = tRootNode.AppendChild(gXML80020DB.XML.CreateElement("area"))
            tNode.SetAttribute "id", tAreaID 'ID
            tNode.SetAttribute "class", 80020 'CLASS
            tNode.SetAttribute "status", 0 'STATUS
            tNode.SetAttribute "recieved", 0 'RECIEVED
            tNode.SetAttribute "innum", 0 'INNUM
            tNode.SetAttribute "outnum", 0 'NAME
            Set tRootNode = tNode
            Set tNode = tRootNode.AppendChild(gXML80020DB.XML.CreateElement("infile"))
            Set tNode = tRootNode.AppendChild(gXML80020DB.XML.CreateElement("outfile"))
        ElseIf tNode.Length > 1 Then
            inReport.ReasonInternal = "БД " & gXML80020DB.ClassTag & " > нарушение структуры уровня AREA. Элементов " & tAreaID & " более 1."
            inReport.Reason = "Внутренняя ошибка автоматики"
            inReport.Status = "Обработка приостановлена до решения проблемы"
            inReport.Decision = 2
            uADebugPrint tLogTag, inReport.ReasonInternal
            Exit Sub
        End If
    ElseIf tNode.Length > 1 Then 'Возможно стоит сделать автопочинку (стирание всех элементов и создание нового - вынести создание на отдельную SUB)
        inReport.ReasonInternal = "БД " & gXML80020DB.ClassTag & " > нарушение структуры. Элементов " & tAreaID & " более 1 на дату " & inDate & "."
        inReport.Reason = "Внутренняя ошибка автоматики"
        inReport.Status = "Обработка приостановлена до решения проблемы"
        inReport.Decision = 2
        uADebugPrint tLogTag, inReport.ReasonInternal
        Exit Sub
    End If
'04 \\ Поиск AREA в базе данных полученных AREA (ПОВТОРНО)
    Set tNode = gXML80020DB.XML.SelectNodes(tMainXPath)
    If tNode.Length <> 1 Then
        inReport.ReasonInternal = "БД " & gXML80020DB.ClassTag & " > Аномальное количество элементов (" & tNode.Length & ") AREA " & tAreaID & " на дату " & inDate & " (должно быть 1)."
        inReport.Reason = "Внутренняя ошибка автоматики"
        inReport.Status = "Обработка приостановлена до решения проблемы"
        inReport.Decision = 2
        uADebugPrint tLogTag, inReport.ReasonInternal
        Exit Sub
    End If
    Set tRootNode = gXML80020DB.XML.SelectSingleNode(tMainXPath)
'05 \\ Проверка номера входяшего макета с номером уже имеющегося входящего по этой AREA (если номер меньше или равен - отказать в приёме)
    tInNum = CLng(tRootNode.GetAttribute("innum"))
    If tNumber <= tInNum Then
        inReport.Reason = "Входящий макет отклонён. Номер макета " & tNumber & " меньше или равен уже принятому " & tInNum & "."
        inReport.Status = "Отклонена"
        inReport.Decision = 1
        uADebugPrint tLogTag, inReport.Reason
        Exit Sub
    End If
    'tInNum = tRootNode.getAttribute("class")
'06 \\ Проверки пройдены, можно выдать номер макету извлекаемой AREA
    Set tTempNode = tRootNode.ParentNode
    tValue = CLng(tTempNode.GetAttribute("number"))
    inNewNumber = tValue + 1
    Set inXML80020Node = tRootNode
'07 \\ Первичная регистрация текущего макета (в случае ошибки будет отменена)
    ' 07.1 // Номер исходящего макета по всей ГТП (АИИС)
    fRemoveChilds inXML80020Node
    Set tNode = inXML80020Node.ParentNode
    tNode.SetAttribute "number", inNewNumber
    ' 07.2 // Имя входяшего файла
    Set tNode = fGetChildNodeByName(inXML80020Node, "infile", True)
    tNode.Text = inFile.Name
    ' 07.4 // Номер входящего файла
    inXML80020Node.SetAttribute "innum", tNumber
    ' 07.5 // Номер исходящего файла
    inXML80020Node.SetAttribute "outnum", inNewNumber
    ' 07.6 // Дата и время получения данных
    inXML80020Node.SetAttribute "recieved", tCurrentTime
    ' 07.7 // Убираем аттрибут отчета (если он был)
    inXML80020Node.RemoveAttribute "report"
End Sub

Private Function fSplitDropArea(inFile, inAreaNode, inDate, inVersionNode, inAreaID, inNewNumber, inXML80020Node, inReport As TCommonReport) As Boolean 'SDA
    Dim tNumber, tYear, tMonth, tDay, tXML80020Blank, tGTPID, tSectionNode, tMainNode, tTraderNode, tDropPath, tFileName, tDropFullPath, tSenderINN, tSenderName, tAIISCode, tTimeZone, tSource80020Node, tDayLightSaving, tRoot, tNode, tComment, tCurrentTime, tSideDropPath, tLogTag
    Dim tPathList, tToAll, tErrorText, tResultDropPath
    
    fSplitDropArea = False
    tLogTag = "SDA"
'01 // Извлечение даты
    If Not (IsTimeStamp(inDate, tYear, tMonth, tDay)) Then
        inReport.ReasonInternal = "Не удалось определить дату из [" & inDate & "]."
        inReport.Reason = "Внутренняя ошибка автоматики"
        inReport.Status = "Обработка приостановлена до решения проблемы"
        inReport.Decision = 2
        uADebugPrint tLogTag, inReport.ReasonInternal
        Exit Function
    End If
'02 // Вытяжка данных для заполнения нового XML
    Set tSectionNode = inVersionNode.ParentNode
    Set tMainNode = tSectionNode.ParentNode
    Set tTraderNode = tMainNode.ParentNode
    tSenderINN = tTraderNode.GetAttribute("inn")
    tSenderName = tTraderNode.GetAttribute("name")
    tGTPID = tMainNode.GetAttribute("id")
    tAIISCode = tMainNode.GetAttribute("aiiscode")
    tTimeZone = tSectionNode.GetAttribute("timezone")
    Set tNode = inAreaNode.ParentNode
    tNumber = CLng(tNode.GetAttribute("number"))
    tCurrentTime = CDate(Now())
    Set tSource80020Node = inAreaNode.SelectNodes("//datetime/daylightsavingtime")
    If tSource80020Node.Length = 1 Then
        tDayLightSaving = tSource80020Node(0).Text
    Else
        tDayLightSaving = 0
    End If
'03 // Создание пути
    'MainFolder
    If Not fBuildM80020DropFolder(gXML80020CFG.Path.Processed, tDropPath, tYear, tMonth, tGTPID, True, inReport.ReasonInternal) Then
        inReport.Reason = "Внутренняя ошибка автоматики #1-" & tLogTag
        inReport.Status = "Обработка приостановлена до решения проблемы"
        inReport.Decision = 2
        uADebugPrint tLogTag, inReport.ReasonInternal
        Exit Function
    End If
    
    If Not fGetStorageListByTag(g80020FolderTag, tPathList, tToAll, tErrorText) Then
        inReport.Reason = "Внутренняя ошибка автоматики #2-" & tLogTag
        inReport.Status = "Обработка приостановлена до решения проблемы"
        inReport.Decision = 2
        inReport.ReasonInternal = tErrorText
        uCDebugPrint tLogTag, 2, inReport.ReasonInternal
        Exit Function
    End If
    
    tResultDropPath = tPathList(0)
    
    'SideFolder
    If Not fBuildM80020DropFolder(tResultDropPath, tSideDropPath, tYear, tMonth, tGTPID, True, inReport.ReasonInternal) Then
        inReport.Reason = "Внутренняя ошибка автоматики #3-" & tLogTag
        inReport.Status = "Обработка приостановлена до решения проблемы"
        inReport.Decision = 2
        uADebugPrint tLogTag, inReport.ReasonInternal
        Exit Function
    End If
    'Папка периода
    'tDropPath = gXML80020CFG.Path.Processed & "\" & tYear & "-" & Format(tMonth, "00")
    'If Not (uFileExists(tDropPath)) Then
    '    If Not (uFolderCreate(tDropPath)) Then
    '        inReport.ReasonInternal = "Не удалось создать папку > " & tDropPath
    '        inReport.Reason = "Внутренняя ошибка автоматики"
    '        inReport.Status = "Обработка приостановлена до решения проблемы"
    '        inReport.Decision = 2
    '        uADebugPrint tLogTag, inReport.ReasonInternal
    '        Exit Function
    '    End If
    'End If
    'gSideDropPath *REMOVE IT* >>>>
    'tSideDropPath = gSideDropPath & "\" & tYear & "-" & Format(tMonth, "00")
    'If Not (uFileExists(tSideDropPath)) Then
    '    If Not (uFolderCreate(tSideDropPath)) Then
    '        inReport.ReasonInternal = "Не удалось создать папку > " & tSideDropPath
    '        inReport.Reason = "Внутренняя ошибка автоматики"
    '        inReport.Status = "Обработка приостановлена до решения проблемы"
    '        inReport.Decision = 2
    '        uADebugPrint tLogTag, inReport.ReasonInternal
    '        Exit Function
    '    End If
    'End If
    '<<<<
    'Папка ГТП
    'tDropPath = tDropPath & "\" & tGTPID
    'If Not (uFileExists(tDropPath)) Then
    '    If Not (uFolderCreate(tDropPath)) Then
    '        inReport.ReasonInternal = "Не удалось создать папку > " & tDropPath
    '        inReport.Reason = "Внутренняя ошибка автоматики"
    '        inReport.Status = "Обработка приостановлена до решения проблемы"
    '        inReport.Decision = 2
    '        uADebugPrint tLogTag, inReport.ReasonInternal
    '        Exit Function
    '    End If
    'End If
    'Папка ГТП REMOVE IT
    'tSideDropPath = tSideDropPath & "\" & tGTPID
    'If Not (uFileExists(tSideDropPath)) Then
    '    If Not (uFolderCreate(tSideDropPath)) Then
    '        inReport.ReasonInternal = "Не удалось создать папку > " & tSideDropPath
    '        inReport.Reason = "Внутренняя ошибка автоматики"
    '        inReport.Status = "Обработка приостановлена до решения проблемы"
    '        inReport.Decision = 2
    '        uADebugPrint tLogTag, inReport.ReasonInternal
    '        Exit Function
    '    End If
    'End If
'04 // Подготовка имён
    tFileName = "80020_" & tSenderINN & "_" & inDate & "_" & inNewNumber & "_" & tAIISCode & ".xml"
    tDropFullPath = tDropPath & "\" & tFileName
    tSideDropPath = tSideDropPath & "\" & tFileName '<<<<<<<< REMOVE IT
'05 // Создание бланка 80020 макета
    tComment = "Сформировано " & tCurrentTime & " Outlook Interceptor Tool (номер оригинального макета - " & tNumber & ")"
    fBlank80020Create tXML80020Blank, tDropFullPath, inNewNumber, inDate, tDayLightSaving, tSenderName, tSenderINN, tComment
'06 // Копирование блока AREA из оригинального макета 80020 в подготовленный бланк
    Set tRoot = tXML80020Blank.DocumentElement
    Set tNode = tRoot.AppendChild(inAreaNode.CloneNode(True))
'07 // Регистрация новых макетов в БД 80020
    ' 07.1 // Имя исходяшего файла
    Set tNode = fGetChildNodeByName(inXML80020Node, "outfile", True)
    tNode.Text = tFileName
'08 // Сохраним новый макет
    fSaveXMLChanges tXML80020Blank, tDropFullPath, inUseRebuild:=True
    '<<<<<<<< REMOVE IT
    On Error Resume Next
        gFSO.CopyFile tDropFullPath, tSideDropPath
    On Error GoTo 0
    '<<<<<<<< REMOVE IT
'00 // Подготовка
    fSplitDropArea = True
'00 // Подготовка
'00 // Подготовка
'00 // Подготовка
'00 // Подготовка
End Function

Private Sub fBlank80020Create(inXML, inDropFullPath, inNumber, inDate, inDayLightSaving, inSenderName, inSenderINN, inCommentString)
Dim tFilePath, tCurrentTime, tRoot, tNode, tRecord, tComment, tIntro
'00 // Подготовка
    Set inXML = Nothing
    If Not uDeleteFile(inDropFullPath) Then
        uDebugPrint "BLK80020: Не удалось удалить файл > " & inDropFullPath
        Exit Sub
    End If
    
'01 // Подготовка XML
    tCurrentTime = Now()
    Set inXML = CreateObject("Msxml2.DOMDocument.6.0")
    inXML.ASync = False
    inXML.Load (tFilePath)
    
'02 // Кореневая нода макета MESSAGE
    Set tRoot = inXML.CreateElement("message")
    inXML.AppendChild tRoot
    tRoot.SetAttribute "class", 80020 'CLASS
    tRoot.SetAttribute "version", 2 'VERSION
    tRoot.SetAttribute "number", inNumber 'NUMBER
    
'03 // Нода времени DATETIME
    Set tNode = tRoot.AppendChild(inXML.CreateElement("datetime"))
    Set tRecord = tNode.AppendChild(inXML.CreateElement("timestamp"))
    tRecord.Text = Format(tCurrentTime, "YYYYMMDDhhmmss")
    Set tRecord = tNode.AppendChild(inXML.CreateElement("daylightsavingtime"))
    tRecord.Text = inDayLightSaving
    Set tRecord = tNode.AppendChild(inXML.CreateElement("day"))
    tRecord.Text = inDate
    
'04 // Нода отправителя SENDER
    Set tNode = tRoot.AppendChild(inXML.CreateElement("sender"))
    Set tRecord = tNode.AppendChild(inXML.CreateElement("name"))
    tRecord.Text = inSenderName
    Set tRecord = tNode.AppendChild(inXML.CreateElement("inn"))
    tRecord.Text = inSenderINN
    
'05 // Комментарий
    Set tComment = inXML.CreateComment(inCommentString)
    inXML.InsertBefore tComment, inXML.ChildNodes(0)
    
'06 // Инструкция обработчикам XML
    Set tIntro = inXML.CreateProcessingInstruction("xml", "version='1.0' encoding='Windows-1251' standalone='yes'")
    inXML.InsertBefore tIntro, inXML.ChildNodes(0)
    
'07 // Первое сохранение шаблона
    'fSaveXMLChanges inXML, tFilePath
End Sub

'Public Sub fXML80020Reprocessor()
'Dim tFile, tIncomingFolder, tIndex, tXMLDoc, tExtractedAreaCount
'    fInit                                    'глобальные переменные
'    If Not fXMLSmartUpdate Then: Exit Sub   'выбор конфигурации базиса
'    fReloadXML gXML80020DB.XML, gXML80020DB.Path 'принудительное чтение БД
'    uDebugPrint "REP: Start"
'    tIndex = 0
'    Set tXMLDoc = CreateObject("Msxml2.DOMDocument.6.0")
'    tXMLDoc.ASync = False
'    Set tIncomingFolder = gFSO.GetFolder(gXML80020CFG.Path.Incoming)
'    tExtractedAreaCount = 0
'===
'    For Each tFile In tIncomingFolder.Files
'        tXMLDoc.Load tFile.Path
'        tExtractedAreaCount = tExtractedAreaCount + fXML80020Reprocess(tFile, tXMLDoc)
'        If Not uMoveFile(tFile.Path, gXML80020CFG.Path.Done & "\" & tFile.Name) Then
 '           uDebugPrint "REP: Не удалось переместить файл в папку отработанных макетов."
'        End If
'        tIndex = tIndex + 1
'    Next
'===
'    Set tXMLDoc = Nothing
'    uDebugPrint "REP: End. Processed files = " & tIndex & "; Exctracted AREAs = " & tExtractedAreaCount
'=== sender delayed activation
'    If tExtractedAreaCount > 0 Then: fEngageXMLASender
'End Sub


'ЭТАП 3 \\ Рассылка макетов
'Public Sub fEngageXMLASender()
'    fActivateTimer 120, 2
'    uDebugPrint "TIM: XMLASender engaged!"
'End Sub

Private Function fGetSendPeriod(inTimeZones As TSendTimeZoneList)
Dim tLogTag, tNow, tNowFormated, tStartDateFormated, tEndDateFormated, tStartDate, tEndDate, tTimeZoneIndex, tHourShiftCorrection, tErrorText 'tLocalUTC
'TODO: Функция не имеет понятия о Зимнем Времени (если введут - надо допилить)
'00 // Подготовка и инициация
    tLogTag = "GSP"
    'tWorkDayLimit = 3
    'tLocalUTC = 4 'Локальный часовой пояс
    fGetSendPeriod = False
'01 // Зададим параметры часовых поясов и классовые отличия
    inTimeZones.TimeZoneCount = 1
    ReDim inTimeZones.TimeZone(inTimeZones.TimeZoneCount)
    With inTimeZones.TimeZone(0)
        .Class = "80020"
        .ID = 1
        .Name = "МСК"
        .UTC = 3
        .DayLimit = 3
    End With
    With inTimeZones.TimeZone(1)
        .Class = "80020"
        .ID = 3
        .Name = "ХАБ"
        .UTC = 10
        .DayLimit = 3
    End With
'02 // Определим периоды рассылки для каждого из часовых
    For tTimeZoneIndex = 0 To inTimeZones.TimeZoneCount
        With inTimeZones.TimeZone(tTimeZoneIndex)
'02.1 // Получим дату начала периода отсутпив на tWorkDayLimit рабочих дней от текущей даты в прошлое
            tHourShiftCorrection = (.UTC - gLocalUTC) / 24 'расчитаем смещение от ЛОКАЛЬНОГО ЧАСОВОГО ПОЯСА
            tNow = Now() + tHourShiftCorrection 'учтем смещение от ЛОКАЛЬНОГО ЧАСОВОГО ПОЯСА
            tNowFormated = CLng(Format(tNow, "YYYYMMDD"))
            Select Case .Class
                Case "80020":
                    'tStartDateFormated = fWorkDayShift(tNowFormated, -.DayLimit)
                    'если возникла проблема с шифтером, то следует прекратить процедуру рассылки
                    'If Len(tStartDateFormated) <> 8 Then
                    If Not fWorkDayShiftAdv(tNowFormated, -.DayLimit, 1, tStartDateFormated, tErrorText) Then
                        uADebugPrint tLogTag, tErrorText
                        uADebugPrint tLogTag, "Не удалось определить начальную дату периода рассылки! Пояс - " & .Name
                        Exit Function
                    End If
                    tStartDateFormated = CLng(tStartDateFormated)
                    tStartDate = DateSerial(Left(tStartDateFormated, 4), Mid(tStartDateFormated, 5, 2), Right(tStartDateFormated, 2))
'02.2 // Получим дату конца периода рассылки - вчерашний день
                    tEndDate = tNow - 1
                    tEndDateFormated = CLng(Format(tEndDate, "YYYYMMDD"))
            End Select
'02.3 // Проверки логики дат
            'проверка на соотвествие полученных дат реальным датам
            If Not (IsDate(tStartDate) And IsDate(tEndDate)) Then
                uADebugPrint tLogTag, "Не удалось определить даты периода рассылки! Пояс - " & .Name
                Exit Function
            End If
            'логическая проверка
            If tEndDateFormated < tStartDateFormated Then
                uADebugPrint tLogTag, "Ошибка логики дата конца периода рассылки (" & tEndDateFormated & ") меньше даты начала периода (" & tStartDateFormated & ")! Пояс - " & .Name
                Exit Function
            End If
'02.4 // Проверки пройдены, внесем полученные данные
            uADebugPrint tLogTag, "Период рассылки (класс " & .Class & "): " & tStartDateFormated & " - " & tEndDateFormated & "; Пояс - " & .Name & "; Текущее время - " & Format(tNow, "YYYYMMDD HH:mm")
            .StartDate = tStartDate
            .StartDateFormated = Fix(tStartDateFormated)
            .EndDate = tEndDate
            .EndDateFormated = Fix(tEndDateFormated)
            .Now = tNow
            .NowFormated = Fix(tNowFormated)
        End With
    Next
'XX // Завершение
    If inTimeZones.TimeZoneCount < 0 Then
        uADebugPrint tLogTag, "Не внесены часовые пояса периодов рассылки!"
        Exit Function
    End If
    fGetSendPeriod = True
End Function

Private Function fGetTimeZoneIndex(inTimeZones As TSendTimeZoneList, inTimeZone, inClass)
Dim tIndex
    fGetTimeZoneIndex = -1
    For tIndex = 0 To inTimeZones.TimeZoneCount
        If inTimeZones.TimeZone(tIndex).Class = inClass And inTimeZones.TimeZone(tIndex).ID = inTimeZone Then
            fGetTimeZoneIndex = tIndex
            Exit Function
        End If
    Next
End Function

Private Function fDateFormatedToDate(inFormatedDate)
    fDateFormatedToDate = 0
    If IsNumeric(inFormatedDate) And Len(inFormatedDate) = 8 Then
        On Error Resume Next
            fDateFormatedToDate = DateSerial(Left(inFormatedDate, 4), Mid(inFormatedDate, 5, 2), Right(inFormatedDate, 2))
            If Err.Number <> 0 Then: fDateFormatedToDate = 0
        On Error GoTo 0
    End If
End Function

Private Function fDateToDateFormated(inDate)
    fDateToDateFormated = vbNullString
    If IsDate(inDate) Then: fDateToDateFormated = Format(Year(inDate), "0000") & Format(Month(inDate), "00") & Format(Day(inDate), "00")
End Function

'Поиск ноды Item в блоке Exchange по классу
Private Function fGetExchangeNodeByClass(inVersionNode, inClass)
Dim tNode, tItemNode, tClass
    Set fGetExchangeNodeByClass = Nothing
    If inVersionNode Is Nothing Then: Exit Function
    If inVersionNode.ChildNodes.Length = 0 Then: Exit Function
    For Each tNode In inVersionNode.ChildNodes
        If LCase(tNode.NodeName) = "exchange" And tNode.ChildNodes.Length > 0 Then
            For Each tItemNode In tNode.ChildNodes
                tClass = LCase(tItemNode.GetAttribute("id"))
                If tClass = inClass Then
                    Set fGetExchangeNodeByClass = tItemNode
                    'uADebugPrint "GENBC", "Locked"
                    Exit Function
                End If
            Next
        End If
    Next
End Function

'Получение списка активных Area на текущий момент и готовых к работе с пересылкой
Private Sub fGetActiveAreaList(inSendGTP As TSendGTPItem, inTimeZones As TSendTimeZoneList)
Dim tSectionNode, tSectionNodes, tVersionNode, tVersionNodes, tCreatedDate, tClosedDate, tAreaNode, tTZIndex, tStartDate, tEndDate, tClass, tActive, tAreaID, tIndexDate, tIndex
Dim tSectionID, tTimeZone, tLogTag, tXPathString, tValue, tAreaNodes, tTZWindowStartDate, tTZWindowEndDate, tVersionID

'00 // Подготовка и инициация
    tLogTag = fGetLogTag("GetActiveAreaList")
    inSendGTP.ActiveAreaCount = -1
    If inSendGTP.Node Is Nothing Then: Exit Sub
    
'01 // Определяем ноды сечений в исходной ноде ГТП
    tXPathString = "child::section"
    Set tSectionNodes = inSendGTP.Node.SelectNodes(tXPathString)
    If tSectionNodes.Length = 0 Then: Exit Sub
    
'01 // Перебор сечений
    For Each tSectionNode In tSectionNodes
    
        ' // Параметры сечения
        tSectionID = UCase(tSectionNode.GetAttribute("id"))
        
        tTimeZone = 1
        If fGetTypedAttributeByName(tSectionNode, "timezone", "INT", tValue) Then
            If tValue = 3 Then: tTimeZone = 3
        End If
        
        ' // Определим набор версий сечения
        tXPathString = "child::version"
        Set tVersionNodes = tSectionNode.SelectNodes(tXPathString)
        
'02 // Перебор версий
        For Each tVersionNode In tVersionNodes
            
            ' // Параметры актуальности версии сечения
            tCreatedDate = -1
            tClosedDate = 0
            
            If fGetTypedAttributeByName(tVersionNode, "created", "INT", tValue) Then: tCreatedDate = tValue
            If fGetTypedAttributeByName(tVersionNode, "closed", "INT", tValue) Then: tClosedDate = tValue
            If fGetTypedAttributeByName(tVersionNode, "id", "INT", tValue) Then: tVersionID = tValue
               
            ' // Определим набор версий сечения
            tXPathString = "child::area"
            Set tAreaNodes = tVersionNode.SelectNodes(tXPathString)
            
' 03 // Перебор AREA
            For Each tAreaNode In tAreaNodes
                            
                ' // Параметры AREA
                tAreaID = UCase(tAreaNode.GetAttribute("id"))
                tClass = UCase(tAreaNode.GetAttribute("type"))
                If tClass = "1" Then
                    tClass = "80020"
                Else
                    tClass = "80040"
                End If
                
                ' // Временные зоны
                tTZIndex = fGetTimeZoneIndex(inTimeZones, tTimeZone, tClass)
                
                If tTZIndex > -1 Then
                    
                    ' // Окно времени по временной зоне
                    tTZWindowStartDate = Fix(inTimeZones.TimeZone(tTZIndex).StartDateFormated)
                    tTZWindowEndDate = Fix(inTimeZones.TimeZone(tTZIndex).EndDateFormated)
                    
                    ' // Определение активных временных рамок для текущей AREA
                    tActive = True
                    If tClosedDate > 0 And tTZWindowStartDate > tClosedDate Then: tActive = False
                    If tTZWindowEndDate < tCreatedDate Then: tActive = False
                    
                    ' // В случае АКТИВНОСТИ необходимы последующий действия
                    If tActive Then
                                             
                        ' // Подстройка даты начала АКТИВНОСТИ
                        If tTZWindowStartDate < tCreatedDate Then
                            tStartDate = tCreatedDate 'fDateFormatedToDate(tCreatedDate)
                        Else
                            tStartDate = tTZWindowStartDate
                        End If
                        
                        ' // Подстройка даты конца АКТИВНОСТИ
                        If tClosedDate > tTZWindowEndDate Or tClosedDate = 0 Then
                            tEndDate = tTZWindowEndDate
                        Else
                            tEndDate = tClosedDate
                        End If
                                        
                        ' // Отсечение времени от даты (корректировка)
                        tStartDate = Fix(fDateFormatedToDate(tStartDate))
                        tEndDate = Fix(fDateFormatedToDate(tEndDate))
                        uCDebugPrint tLogTag, 0, "TZIndex=" & tTZIndex & "; P=" & inSendGTP.ID & "-" & tSectionID & " v" & tVersionID & "; A=" & tAreaID & "; PERIOD=" & tStartDate & "-" & tEndDate
                        
                        ' // Добавление AREA в список активных
                        inSendGTP.ActiveAreaCount = inSendGTP.ActiveAreaCount + 1
                        ReDim Preserve inSendGTP.ActiveAreaList(inSendGTP.ActiveAreaCount)
                        
                        With inSendGTP.ActiveAreaList(inSendGTP.ActiveAreaCount)
                            .Class = tClass
                            .ID = tAreaID
                            .TimeZone = tTZIndex
                            .SectionID = tSectionID
                            .SendDaysCount = DateDiff("d", tStartDate, tEndDate)
                            Set .ExchangeNode = fGetExchangeNodeByClass(tVersionNode, .Class)
                            ReDim .SendDaysList(.SendDaysCount)
                                            
                            tIndex = -1
                            For tIndexDate = tStartDate To tEndDate
                                tIndex = tIndex + 1
                                .SendDaysList(tIndex) = tIndexDate
                            Next
                        End With
                    End If 'ACTIVE
                End If 'TZ
            Next
        Next
    Next
End Sub

' fGetMailString - form internal mail list string with params
' USED EXTERNAL: cnstMailListParamDelimiter, cnstMailListDelimiter, uAddToList
Private Function fGetMailString(inExchangeItemNode)
    Dim tMailToNodes, tXPathString, tNode, tEmailAddress, tEncrypt, tSign, tElement
    
    'default
    fGetMailString = vbNullString
    
    'quick check input var
    If inExchangeItemNode Is Nothing Then: Exit Function
    If TypeName(inExchangeItemNode) <> "IXMLDOMElement" Then: Exit Function
    
    'scan for child nodes
    tXPathString = "child::mailto[@enabled='1']"
    Set tMailToNodes = inExchangeItemNode.SelectNodes(tXPathString)
    
    For Each tNode In tMailToNodes
        
        'read
        fGetAttr tNode, "address", tEmailAddress
        fGetAttr tNode, "encrypt", tEncrypt
        fGetAttr tNode, "sign", tSign
        
        'fix and check
        tEmailAddress = Trim(tEmailAddress)
        If Not (tEncrypt = "1" Or tEncrypt = 1) Then: tEncrypt = "0"
        If Not (tSign = "1" Or tSign = 1) Then: tSign = "0"
        
        If tEmailAddress <> vbNullString Then
            'form mail element
            tElement = LCase(tEmailAddress & cnstMailListParamDelimiter & tEncrypt & cnstMailListParamDelimiter & tSign) ' - is delimiter lv1
        
            'add element to list
            uAddToList fGetMailString, tElement, cnstMailListDelimiter ' - delimiter lv2
        End If
    Next
End Function

'Сравнивает два списка (список отправленных и список к отправке) и убирет из списка к отправке элементы присутсвующие в списке отправленных
Private Function fMailListAdjust(inMailString, inSentString)
Dim tFinalString, tMailElements, tSentElements, tMailAddress, tPosA, tSentAddress, tPosB, tMailElement, tSentElement, tAlreadySent
    fMailListAdjust = vbNullString
    If IsNull(inSentString) Then: inSentString = vbNullString
    inMailString = LCase(inMailString)
    inSentString = LCase(inSentString)
    tMailElements = Split(inMailString, cnstMailListDelimiter)
    tSentElements = Split(inSentString, cnstMailListDelimiter)
    If UBound(tSentElements) < 0 Then
        fMailListAdjust = inMailString
        Exit Function 'nothing to adjust - all items should be sent
    End If
    If UBound(tMailElements) < 0 Then: Exit Function 'nothing to adjust - no items on input
    tFinalString = vbNullString
    For Each tMailElement In tMailElements
        tPosA = InStr(tMailElement, cnstMailListParamDelimiter)
        If tPosA > 0 Then
            tMailAddress = Left(tMailElement, tPosA - 1)
            tAlreadySent = False
            For Each tSentElement In tSentElements
                tPosB = InStr(tSentElement, cnstMailListParamDelimiter)
                If tPosB > 0 Then
                    tSentAddress = Left(tSentElement, tPosB - 1)
                    If tSentAddress = tMailAddress Then
                        tAlreadySent = True
                        Exit For
                    End If
                End If
            Next
            If Not tAlreadySent Then
                uAddToList tFinalString, tMailElement, cnstMailListDelimiter
            End If
        End If
    Next
    fMailListAdjust = tFinalString
End Function

Private Function fDateSplitter(inDate, outDateText, outYear, outMonth, outDay)
    fDateSplitter = False
    outDateText = vbNullString
    outYear = vbNullString
    outMonth = vbNullString
    outDay = vbNullString
    If Not IsDate(inDate) Then: Exit Function
    outYear = Format(Year(inDate), "0000")
    outMonth = Format(Month(inDate), "00")
    outDay = Format(Day(inDate), "00")
    outDateText = outYear & outMonth & outDay
    fDateSplitter = True
End Function

Private Function fMailListIsEqual(inListA, inListB)
Dim tElementsA, tElementsB, tElementA, tElementB, tPosA, tPosB, tAddressA, tAddressB, tIsEqual
    fMailListIsEqual = False
    'autocorrection
    If IsNull(inListA) Then: inListA = vbNullString
    If IsNull(inListB) Then: inListB = vbNullString
    'compare P1 Easy
    If Len(inListA) <> Len(inListB) Then: Exit Function
    'compare P2 Deep
    tElementsA = Split(inListA, cnstMailListDelimiter)
    tElementsB = Split(inListB, cnstMailListDelimiter)
    If UBound(tElementsA) <> UBound(tElementsB) Then: Exit Function
    'for each element of list A
    For Each tElementA In tElementsA
        tIsEqual = False
        tPosA = InStr(tElementA, cnstMailListParamDelimiter)
        If tPosA > 0 Then
            tAddressA = Left(tElementA, tPosA - 1)
            'for each element of list A compare to each element of list B > to find equal address
            For Each tElementB In tElementsB
                tPosB = InStr(tElementB, cnstMailListParamDelimiter)
                If tPosB > 0 Then
                    tAddressB = Left(tElementB, tPosB - 1)
                    If tAddressA = tAddressB Then
                        tIsEqual = True
                        Exit For 'Leave cycle B
                    End If
                End If
            Next
        End If
        'if element A has no pair in list B > lists not equal > exit
        If Not tIsEqual Then: Exit Function
    Next
    'lists are equal
    fMailListIsEqual = True
End Function

' fMailListClearing - delete dublicates (and empty items) from mail-list and return cleared list
' EXTERNALS: cnstMailListDelimiter, cnstMailListParamDelimiter
' UMOD: uAddToList
' Added 2021-08-12
Private Function fMailListClearing(inMailList)
    Dim tMailElements, tMainIndex, tResultList, tMailDetails, tInternalIndex, tTempMailCount, tLock, tCurrentElement
    Dim tTempMailList()
    
    'by default return original
    fMailListClearing = inMailList
    tResultList = vbNullString
    tTempMailCount = -1
    
    'split to main elements (email:param1:param2)
    tMailElements = Split(Trim(inMailList), cnstMailListDelimiter)
    
    'run by elements
    For tMainIndex = 0 To UBound(tMailElements)
    
        tCurrentElement = Trim(tMailElements(tMainIndex))
        
        'safe guarding
        If tCurrentElement <> vbNullString Then
            tMailDetails = Split(tMailElements(tMainIndex), cnstMailListParamDelimiter)
            
            'safe guarding
            If UBound(tMailDetails) = 2 Then
                'look back for same address (dublicate)
                tLock = False
                For tInternalIndex = 0 To tTempMailCount
                    If tMailDetails(0) = tTempMailList(tInternalIndex) Then
                        tLock = True
                        Exit For
                    End If
                Next
                
                'if not locked dublicate - add address to list
                If Not tLock Then
                    uAddToList tResultList, tMailElements(tMainIndex), cnstMailListDelimiter
                    tTempMailCount = tTempMailCount + 1
                    ReDim Preserve tTempMailList(tTempMailCount)
                    tTempMailList(tTempMailCount) = tMailDetails(0)
                End If
            End If
        End If
    Next
    
    'returning
    If tResultList <> vbNullString Then: fMailListClearing = tResultList
End Function

'fAssignSendList - производит распределение списков рассылки по активным дням и активным AREA текущей ГТП
Private Sub fAssignSendList(inSendGTP As TSendGTPItem)
    Dim tXPathString, tAreaIndex, tDateIndex, tCurrentFormattedDate, tYear, tMonth, tDay, tAreaID, tClass, tNode, tMailNode, tOldMailString, tMailString, tSentString, tIsChanged
    Dim tInternalIndex, tLogTag

    ' 01 // Значение триггера об изменениях в XML по умолчанию (т.е. показывает необходимость сохранить изменения)
    tIsChanged = False
    tLogTag = fGetLogTag("AssignSendList")
    
    ' 02 // Подготовка листов рассылки + удаление задвоений при одном коде AREA -- 2021-08-13 fix
    ' исправление объединяет список рассылки одинаковых AREA в единый список (т.к. рассылка по коду AREA не может быть множественной)
    For tAreaIndex = 0 To inSendGTP.ActiveAreaCount
        With inSendGTP.ActiveAreaList(tAreaIndex)
            .SendEnabled = (TypeName(.ExchangeNode) = "IXMLDOMElement" And Not (.ExchangeNode Is Nothing))
            If .SendEnabled Then
                .MailList = fGetMailString(.ExchangeNode) 'critical to structure of node?
            Else
                .MailList = vbNullString
            End If
        End With
        
        'accumulating areas sendlist with same ID to first one in list
        For tInternalIndex = 0 To tAreaIndex - 1
            If inSendGTP.ActiveAreaList(tInternalIndex).Class = inSendGTP.ActiveAreaList(tAreaIndex).Class And inSendGTP.ActiveAreaList(tInternalIndex).ID = inSendGTP.ActiveAreaList(tAreaIndex).ID And inSendGTP.ActiveAreaList(tInternalIndex).SendEnabled And inSendGTP.ActiveAreaList(tAreaIndex).SendEnabled Then
                inSendGTP.ActiveAreaList(tAreaIndex).SendEnabled = False 'disabling current area and accumulate its sendlist to active one
                uAddToList inSendGTP.ActiveAreaList(tInternalIndex).MailList, inSendGTP.ActiveAreaList(tAreaIndex).MailList 'accumulate sendlist
                inSendGTP.ActiveAreaList(tInternalIndex).MailList = fMailListClearing(inSendGTP.ActiveAreaList(tInternalIndex).MailList) 'remove dublicates
                uCDebugPrint tLogTag, 1, "Аккумуляция списка рассылки по AREA " & inSendGTP.ActiveAreaList(tInternalIndex).ID & " (CLASS " & inSendGTP.ActiveAreaList(tInternalIndex).Class & ")"
                Exit For 'it's nothing to look for more
            End If
        Next
    Next
    
    ' 03 // Перебор переданных АКТИВНЫХ Area
    For tAreaIndex = 0 To inSendGTP.ActiveAreaCount
        With inSendGTP.ActiveAreaList(tAreaIndex)
            
            If .SendEnabled Then 'если рассылка разрешена по этой AREA (с учетом вышестоящего правила об аккумуляции списков)
                
                ' 03 // Перебор активных дней для этой AREA
                For tDateIndex = 0 To .SendDaysCount
                
                    ' get request params
                    fDateSplitter .SendDaysList(tDateIndex), tCurrentFormattedDate, tYear, tMonth, tDay
                    tAreaID = .ID
                    tClass = .Class
                    tXPathString = "/message/trader[@id='" & gTraderInfo.ID & "']/year[@id='" & tYear & "']/month[@id='" & tMonth & "']/day[@id='" & tDay & "']/aiis[@gtpid='" & inSendGTP.ID & "']/area[@id='" & tAreaID & "' and @class='" & tClass & "']"
                    
                    ' get requested node
                    Set tNode = gXML80020DB.XML.SelectSingleNode(tXPathString)
                    
                    ' form or update sendlist data by requested node
                    If Not (tNode Is Nothing) Then
                         
                         tMailString = .MailList ' fGetMailString(.ExchangeNode) // 2021-08-13 fix
                         
                         If tMailString <> vbNullString Then
                            'uADebugPrint "ASL", "A=" & tAreaID & " READY TO SEND for date = " & tCurrentDate & " >> MailString=" & tMailString
                            Set tMailNode = fGetChildNodeByName(tNode, "mail", True)
                            tOldMailString = LCase(tMailNode.GetAttribute("mailto"))
                            tSentString = LCase(tMailNode.GetAttribute("sent"))
                            tMailString = fMailListAdjust(tMailString, tSentString)
                            ''Debug.Print "A2: " & GetTickCount
                            If Not fMailListIsEqual(tMailString, tOldMailString) Then
                                tMailNode.SetAttribute "mailto", tMailString
                                tIsChanged = True
                            End If
                            ''Debug.Print "NL: " & tMailString & " // OL:" & tOldMailString
                            ''Debug.Print "A3: " & GetTickCount
                            'fSaveXMLChanges gXML80020DB.XML, gXML80020DB.Path, True
                        End If
                    End If
                Next
                
            End If
        End With
    Next
    'finalyze
    If tIsChanged Then: fSaveXMLDB gXML80020DB, False, , , , tLogTag & " внесла изменения в списки рассылок!"
End Sub

Private Sub fSenderCommandAddElement(ioCommand, inBlock, inElementSplitter, Optional inAddToStart = False)
    If inBlock = vbNullString Then: Exit Sub
    
    If ioCommand <> vbNullString Then
        If inAddToStart Then
            ioCommand = inBlock & inElementSplitter & ioCommand
        Else 'add as tail
            ioCommand = ioCommand & inElementSplitter & inBlock
        End If
    Else
        ioCommand = inBlock
    End If
End Sub

'fCreateSenderCommand - Формирует строку-команду для функции рассыльщика по формату >> TRADER@ГТП:DATE1@AREA1,AREA2:DATE2@AREA1,AREA2
Private Function fCreateSenderCommand(inSendGTP As TSendGTPItem, inTraderID, Optional inBlockSplitter = ":", Optional inSubSplitter = "@", Optional inEnumSplitter = ",", Optional inDebugMode = False)
    Dim tDayList()
    Dim tDayListCount, tAreaIndex, tDayIndex, tDaySubIndex, tLock, tDayElement, tGTPElement, tResultCommand, tActive, tTickStart
    
    If inDebugMode Then: tTickStart = GetTickCount
    fCreateSenderCommand = vbNullString
    tResultCommand = vbNullString
    tDayListCount = -1
    
    'daylist create
    For tAreaIndex = 0 To inSendGTP.ActiveAreaCount
        For tDayIndex = 0 To inSendGTP.ActiveAreaList(tAreaIndex).SendDaysCount
            tLock = -1
            For tDaySubIndex = 0 To tDayListCount
                If tDayList(tDaySubIndex) = inSendGTP.ActiveAreaList(tAreaIndex).SendDaysList(tDayIndex) Then
                    tLock = tDaySubIndex
                    Exit For
                End If
            Next
            If tLock = -1 Then
                tDayListCount = tDayListCount + 1
                ReDim Preserve tDayList(tDayListCount)
                tDayList(tDayListCount) = inSendGTP.ActiveAreaList(tAreaIndex).SendDaysList(tDayIndex)
            End If
        Next
    Next
    'check logic
    If tDayListCount = -1 Then: Exit Function
    
    'header-block
    tGTPElement = inTraderID & inSubSplitter & inSendGTP.ID
    fSenderCommandAddElement tResultCommand, tGTPElement, inBlockSplitter
    
    'create day elements
    tActive = False
    For tDayIndex = 0 To tDayListCount
        'tLock = False
        tDayElement = vbNullString
        For tAreaIndex = 0 To inSendGTP.ActiveAreaCount
            For tDaySubIndex = 0 To inSendGTP.ActiveAreaList(tAreaIndex).SendDaysCount
                If tDayList(tDayIndex) = inSendGTP.ActiveAreaList(tAreaIndex).SendDaysList(tDaySubIndex) Then
                    'tDayElement = tDayElement & "," & inSendGTP.ActiveAreaList(tAreaIndex).ID
                    fSenderCommandAddElement tDayElement, inSendGTP.ActiveAreaList(tAreaIndex).ID, inEnumSplitter
                    'tLock = True
                    Exit For
                End If
            Next
        Next
        
        If tDayElement <> vbNullString Then
            fSenderCommandAddElement tDayElement, fDateToDateFormated(tDayList(tDayIndex)), inSubSplitter, True 'day-block header
            fSenderCommandAddElement tResultCommand, tDayElement, inBlockSplitter
            tActive = True
        End If
    Next
    
    'check logic
    If tActive Then
        fCreateSenderCommand = tResultCommand
    End If
    
    If inDebugMode Then
        tTickStart = GetTickCount - tTickStart
        'Debug.Print "fCreateSenderCommand time:" & tTickStart
    End If
End Function

Private Function fInjectAreaItemChanges(inAreaItem As TSenderAreaItem, inAreaNodes, ioSaveActivated, Optional inMailBlockSplitter = ";", Optional inMailParamSplitter = ":")
Dim tMailNode, tIndex, tTempString, tXPathString
    With inAreaItem
        tXPathString = "parent::*/child::area[@id='" & .AreaID & "']/mail"
        'Set tMailNode = inAreaNodes(0).SelectSingleNode(tXPathString)
        'Set tMailNode = Nothing
        'Debug.Print TypeName(gXMLBasis.XML)
        'Debug.Print TypeName(gXMLBasis.XML.DocumentElement)
        'Debug.Print TypeName(inAreaNodes)
        'Debug.Print TypeName(inAreaNodes(0))
        Set tMailNode = fGetNodeSafe(inAreaNodes, , , tXPathString)
        ''Debug.Print "MAILNODE=" & Not (tMailNode Is Nothing)
        
        'Set tMailNode = fGetChildNodeByName(tAreaNode, "mail")
        
        If Not tMailNode Is Nothing Then
            'mailto
            tTempString = vbNullString
            For tIndex = 0 To .MailToCount
                If .MailToList(tIndex) <> vbNullString Then
                    If tTempString = vbNullString Then
                        tTempString = .MailToList(tIndex)
                    Else
                        tTempString = tTempString & inMailBlockSplitter & .MailToList(tIndex)
                    End If
                End If
            Next
            tMailNode.SetAttribute "mailto", tTempString
            'sent
            tTempString = vbNullString
            For tIndex = 0 To .SentCount
                If tTempString = vbNullString Then
                    tTempString = .SentList(tIndex)
                Else
                    tTempString = tTempString & inMailBlockSplitter & .SentList(tIndex)
                End If
            Next
            tMailNode.SetAttribute "sent", tTempString
            tMailNode.SetAttribute "updated", fGetRecievedTimeStamp(Now())
            'save changes
            'fSaveXMLChanges gXML80020DB.XML, gXML80020DB.Path, True
            ioSaveActivated = True
        End If
    End With
End Function

Private Function fKillDoubleMails(inMailToString, inMailSentString, Optional inMailBlockSplitter = ";", Optional inMailParamSplitter = ":")
    Dim tMailToList, tMailSentList, tMailAddress, tMailParams, tIndex, tSubIndex, tSubMailParams, tMailToCount, tMailSentCount, tIsDouble, tResultMailToList
    
    fKillDoubleMails = vbNullString
    If inMailToString = vbNullString Then: Exit Function
    tResultMailToList = vbNullString
    
    If inMailSentString = vbNullString Then
        tMailSentCount = -1
    Else
        tMailSentList = Split(inMailSentString, inMailBlockSplitter)
        tMailSentCount = UBound(tMailSentList)
    End If
    
    tMailToList = Split(inMailToString, inMailBlockSplitter)
    tMailToCount = UBound(tMailToList)
    For tIndex = 0 To tMailToCount
        tIsDouble = False
        tMailParams = Split(tMailToList(tIndex), inMailParamSplitter)
        
        'self list scan
        For tSubIndex = tIndex + 1 To tMailToCount
            tSubMailParams = Split(tMailToList(tSubIndex), inMailParamSplitter)
            If tMailParams(0) = tSubMailParams(0) Then
                tIsDouble = True
                Exit For
            End If
        Next
        
        'sent list scan
        If Not tIsDouble Then
            For tSubIndex = 0 To tMailSentCount
                tSubMailParams = Split(tMailSentList(tSubIndex), inMailParamSplitter)
                If tMailParams(0) = tSubMailParams(0) Then
                    tIsDouble = True
                    Exit For
                End If
            Next
        End If
        
        'erase?
        If Not tIsDouble Then
            If tResultMailToList = vbNullString Then
                tResultMailToList = tMailToList(tIndex) 'inMailParamSplitter
            Else
                tResultMailToList = tResultMailToList & inMailBlockSplitter & tMailToList(tIndex) 'inMailParamSplitter
            End If
        End If
    Next
    
    'return result
    fKillDoubleMails = tResultMailToList
End Function

Private Sub fExtractAreaItem(inAreaItem As TSenderAreaItem, inAreaNode, Optional inMailBlockSplitter = ";", Optional inMailParamSplitter = ":")
    Dim tMailNode, tMailToList, tSentList, tFileNode, tTempValue
    Dim tMailToString, tMailSentString
    
    With inAreaItem
        '.AreaID = inAreaNode.GetAttribute("id")
        '.Class = inAreaNode.GetAttribute("class")
        tTempValue = fGetAttr(inAreaNode, "id", .AreaID)
        tTempValue = fGetAttr(inAreaNode, "class", .Class)
        .Error = vbNullString
        
        tTempValue = fGetAttr(inAreaNode, "outnum", .OutNum)
        If IsNumeric(.OutNum) Then
            .OutNum = Fix(.OutNum)
        Else
            .OutNum = Empty
            .Error = "OutNum not numeric"
        End If
        
        .Active = False
        .SectionID = vbNullString 'optional
        tTempValue = fGetAttr(inAreaNode, "status", .Status, 0)
        
        .MailToCount = -1
        Erase .MailToList
        .SentCount = -1
        Erase .SentList
        
        Set tMailNode = fGetChildNodeByName(inAreaNode, "mail")
        If Not tMailNode Is Nothing Then
            'tMailToList = tMailNode.GetAttribute("mailto")
            'If IsNull(tMailToList) Then: tMailToList = vbNullString
            'tSentList = tMailNode.GetAttribute("sent")
            'If IsNull(tSentList) Then: tSentList = vbNullString
            
            tTempValue = fGetAttr(tMailNode, "sent", tSentList)
            If tSentList <> vbNullString Then
                tMailSentString = tSentList
                .SentList = Split(tSentList, inMailBlockSplitter)
                .SentCount = UBound(.SentList)
            End If
            
            tTempValue = fGetAttr(tMailNode, "mailto", tMailToList)
            If tMailToList <> vbNullString Then
                tMailToString = tMailToList
                tMailToString = fKillDoubleMails(tMailToString, tMailSentString, inMailBlockSplitter, inMailParamSplitter)
                If tMailToString <> vbNullString Then
                    .MailToList = Split(tMailToString, inMailBlockSplitter)
                    .MailToCount = UBound(.MailToList)
                End If
            End If
        End If
        
        Set tFileNode = fGetChildNodeByName(inAreaNode, "outfile")
        If Not tFileNode Is Nothing Then
            .FileName = tFileNode.Text
        End If
    End With
End Sub

Private Function fAutoSenderGetCommonParams(inMainNode, outAuthEnabled, outAuthLogin, outAuthPassword, outSMTPServer, outSMTPPort, outMailFrom, outSSLMode, outSSLVersion, outConnectionTimeout, outErrorText)
    Dim tXPathString, tNode, tValue, tConnectionNode
    
    'prepare
    fAutoSenderGetCommonParams = False
    outErrorText = vbNullString
    
    'defaults
    outAuthEnabled = False
    outAuthLogin = vbNullString
    outAuthPassword = vbNullString
    outSMTPServer = vbNullString
    outSMTPPort = 0
    outSSLMode = 0
    outSSLVersion = vbNullString
    outMailFrom = vbNullString
    outConnectionTimeout = 0
    
    tXPathString = vbNullString
    
    'AUTH_ENABLED
    If Not fGetAttributeCFG(gXMLCredentials, tXPathString, "authenabled", tValue, tNode, outErrorText, inMainNode) Then: Exit Function
    outAuthEnabled = (tValue = "1" Or tValue = 1)
    
    If outAuthEnabled Then
        'AUTH_LOGIN
        If Not fGetAttributeCFG(gXMLCredentials, tXPathString, "username", tValue, tNode, outErrorText, inMainNode) Then: Exit Function
        outAuthLogin = tValue
        
        'AUTH_PASSWORD
        If Not fGetAttributeCFG(gXMLCredentials, tXPathString, "password", tValue, tNode, outErrorText, inMainNode) Then: Exit Function
        outAuthPassword = tValue
    End If
    
    'MAIL_FROM
    If Not fGetAttributeCFG(gXMLCredentials, tXPathString, "email", tValue, tNode, outErrorText, inMainNode) Then: Exit Function
    outMailFrom = tValue
    
    'SMTP_SERVER
    If Not fGetAttributeCFG(gXMLCredentials, tXPathString, "smtpserver", tValue, tNode, outErrorText, inMainNode) Then: Exit Function
    outSMTPServer = tValue
    
    'SSL_MODE
    If Not fGetAttributeCFG(gXMLCredentials, tXPathString, "sslmode", tValue, tNode, outErrorText, inMainNode) Then: Exit Function
    outSSLMode = tValue
    
    'SMTP_SERVER_PORT
    tXPathString = "child::connections/connection[@id='" & outSSLMode & "']"
    If Not fGetAttributeCFG(gXMLCredentials, tXPathString, "port", tValue, tConnectionNode, outErrorText, inMainNode) Then: Exit Function
    outSMTPPort = tValue
    
    'SSL_VERSION
    tXPathString = vbNullString
    If Not fGetAttributeCFG(gXMLCredentials, tXPathString, "sslversion", tValue, tNode, outErrorText, tConnectionNode) Then: Exit Function
    outSSLVersion = tValue
    
    'SERVER CONNECTION TIMEOUT
    tXPathString = vbNullString
    If Not fGetAttributeCFG(gXMLCredentials, tXPathString, "timeout", tValue, tNode, outErrorText, inMainNode) Then: Exit Function
    outConnectionTimeout = tValue
    
    'fin
    fAutoSenderGetCommonParams = True
End Function

Private Function fAutoSenderGetSenderByRole(inMainNode, inRequiredRole, outSenderNode, outErrorText)
    Dim tXPathString, tSenderNodes, tSenderNode, tValue, tNode, tMainRole, tSenderFilePath
    
    'prepare
    fAutoSenderGetSenderByRole = False
    outErrorText = vbNullString
    Set outSenderNode = Nothing
    
    'select senders by role
    tXPathString = "child::senders/sender[roles/role[@id='" & inRequiredRole & "']]"
    Set tSenderNodes = inMainNode.SelectNodes(tXPathString)
    
    'look deeper
    For Each tSenderNode In tSenderNodes
        If fGetAttributeCFG(gXMLCredentials, vbNullString, "mainrole", tMainRole, tNode, outErrorText, tSenderNode) Then
            tXPathString = "child::filepath"
            Set tNode = tSenderNode.SelectSingleNode(tXPathString)
            If Not tNode Is Nothing Then
                tSenderFilePath = tNode.Text
                If gFSO.FileExists(tSenderFilePath) Then
                    Set outSenderNode = tSenderNode
                    If tMainRole = inRequiredRole Then: Exit For
                End If
            End If
        End If
    Next
    
    'checker
    If outSenderNode Is Nothing Then
        outErrorText = "Не найдено подходящих для роли <" & inRequiredRole & "> доступных интерфейсов!"
        Exit Function
    End If
    
    'fin
    fAutoSenderGetSenderByRole = True
End Function

Private Function fAutoSenderGetKeysForEncrypt(inMainNode, inSignNeed, inEcryptNeed, outSignKey, outEncryptKey, outErrorText)
    Dim tXPathString, tValue, tNode, tTodayTag, tKeyNodes, tKeyNode, tKeyExpireDate, tKeyType
    
    'prepare
    fAutoSenderGetKeysForEncrypt = False
    outErrorText = vbNullString
    outSignKey = vbNullString
    outEncryptKey = vbNullString
    
    'getdate
    tTodayTag = Fix(Format(Now(), "YYYYMMDD"))
    
    'get SIGN KEY
    If inSignNeed Then
        tXPathString = "child::keys/sign"
        Set tKeyNodes = inMainNode.SelectNodes(tXPathString)
        For Each tKeyNode In tKeyNodes
            If tKeyNode.Text <> vbNullString Then
                If fGetAttributeCFG(gXMLCredentials, vbNullString, "type", tKeyType, tNode, outErrorText, tKeyNode) And fGetAttributeCFG(gXMLCredentials, vbNullString, "expire", tKeyExpireDate, tNode, outErrorText, tKeyNode) Then
                    If IsNumeric(tKeyExpireDate) Then
                        tKeyExpireDate = Fix(tKeyExpireDate)
                        If tTodayTag < tKeyExpireDate Then
                            outSignKey = tKeyNode.Text
                            If tKeyType = "main" Then: Exit For
                        End If
                    End If
                End If
            End If
        Next
        
        'check
        If outSignKey = vbNullString Then
            outErrorText = "Не найден подходящий ключ подписи! Возможно срок ключа истёк."
            Exit Function
        End If
    End If
    
    'get ENCRYPT KEY
    If inEcryptNeed Then
        tXPathString = "child::keys/encrypt"
        Set tKeyNodes = inMainNode.SelectNodes(tXPathString)
        For Each tKeyNode In tKeyNodes
            If tKeyNode.Text <> vbNullString Then
                If fGetAttributeCFG(gXMLCredentials, vbNullString, "type", tKeyType, tNode, outErrorText, tKeyNode) And fGetAttributeCFG(gXMLCredentials, vbNullString, "expire", tKeyExpireDate, tNode, outErrorText, tKeyNode) Then
                    If IsNumeric(tKeyExpireDate) Then
                        tKeyExpireDate = Fix(tKeyExpireDate)
                        If tTodayTag < tKeyExpireDate Then
                            outEncryptKey = tKeyNode.Text
                            If tKeyType = "main" Then: Exit For
                        End If
                    End If
                End If
            End If
        Next
        
        'check
        If outEncryptKey = vbNullString Then
            outErrorText = "Не найден подходящий ключ шифрования! Возможно срок ключа истёк."
            Exit Function
        End If
    End If
    
    'fin
    fAutoSenderGetKeysForEncrypt = True
End Function

Private Function fAutoSenderCommandLineBlockReader(inCommandLineRootNode, inChildName, inIsActive, ioHostCommandLine, inSplitter, outErrorText, Optional inSelector = -1)
    Dim tXPathString, tTempNode, tNode, tValue, tIsRequired, tDefaultIndex, tSelectedIndex, tIsGroup, tLogTag, tItemNode
    
    'prepare
    fAutoSenderCommandLineBlockReader = False
    outErrorText = vbNullString
    tLogTag = fGetLogTag("fAutoSenderCommandLineBlockReader")
    tSelectedIndex = inSelector
    
    'lock child node
    tXPathString = "child::" & inChildName
    Set tNode = inCommandLineRootNode.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        If inIsActive Then
            outErrorText = "Ошибка формирования бланка командной строки! Не найдена нода XPath=[ " & tXPathString & " ]"
            Exit Function
        'drop unexistant (warning)
        Else
            fAutoSenderCommandLineBlockReader = True
            Exit Function
        End If
    End If
    
    'get required param
    If Not fGetAttributeCFG(gXMLCredentials, vbNullString, "required", tIsRequired, tTempNode, outErrorText, tNode) Then: Exit Function
    tIsRequired = tIsRequired = 1
    
    'get default group tag
    tIsGroup = False
    If fGetAttributeCFG(gXMLCredentials, vbNullString, "default", tDefaultIndex, tTempNode, outErrorText, tNode) Then
        tIsGroup = True
        tXPathString = tXPathString & "/item[@id='" & tSelectedIndex & "']"
        Set tNode = inCommandLineRootNode.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            uCDebugPrint tLogTag, 1, "Внимание! Не найден искомый [" & tSelectedIndex & "] элемент группы, будет использован индекс по умолчанию [" & tDefaultIndex & "]! XPath [ " & tXPathString & " ]"
            tSelectedIndex = tDefaultIndex
            tXPathString = tXPathString & "/item[@id='" & tSelectedIndex & "']"
            Set tNode = inCommandLineRootNode.SelectSingleNode(tXPathString)
            If tNode Is Nothing Then
                outErrorText = "Элемент [" & tSelectedIndex & "] по умолчанию группы не найден! XPath [ " & tXPathString & " ]"
                Exit Function
            End If
        End If
    End If
    
    'logic check
    If tIsRequired And Not inIsActive Then
        outErrorText = "Нода <" & tXPathString & "> является обязательной! tIsRequired[" & tIsRequired & "] inIsActive[" & inIsActive & "]"
        Exit Function
    End If
    
    'select item
    If inIsActive Then
        tXPathString = "child::used"
    Else
        tXPathString = "child::notused"
    End If
    
    Set tItemNode = tNode.SelectSingleNode(tXPathString)
    If tItemNode Is Nothing Then
        outErrorText = "Не найдена нода действия для [" & inChildName & "][IsGroup=" & tIsGroup & "/Index:" & tSelectedIndex & "]! XPath [ " & tXPathString & " ]"
        Exit Function
    End If
    
    'build command line
    ioHostCommandLine = ioHostCommandLine & inSplitter & tItemNode.Text
    
    'fin
    fAutoSenderCommandLineBlockReader = True
End Function

Private Function fAutoSenderCommandLineConstruct(inSenderNode, inAuthEnabled, inSSLMode, inSignNeeded, inEcryptNeeded, inHasSubject, inAttachmentNeeded, outBlankString, outErrorText)
    Dim tXPathString, tResultString, tNode, tCommandLineRootNode, tCommandSplitter, tCommandBlock, tCommandBlockNode, tIsRequired
    'prepare
    fAutoSenderCommandLineConstruct = False
    outErrorText = vbNullString
    outBlankString = vbNullString
    
    'exec path
    tXPathString = "child::filepath"
    Set tNode = inSenderNode.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        outErrorText = "Ошибка формирования бланка командной строки! Не найдена нода XPath=[ " & tXPathString & " ]"
        Exit Function
    End If
    tResultString = tNode.Text
    
    'get commandline node
    tXPathString = "child::commandline"
    Set tCommandLineRootNode = inSenderNode.SelectSingleNode(tXPathString)
    If tCommandLineRootNode Is Nothing Then
        outErrorText = "Ошибка формирования бланка командной строки! Не найдена нода XPath=[ " & tXPathString & " ]"
        Exit Function
    End If
    
    'SPLITTER
    tXPathString = vbNullString
    If Not fGetAttributeCFG(gXMLCredentials, tXPathString, "splitter", tCommandSplitter, tNode, outErrorText, tCommandLineRootNode) Then: Exit Function
    
    'CONNECTION
    If Not fAutoSenderCommandLineBlockReader(tCommandLineRootNode, "connection", True, tResultString, tCommandSplitter, outErrorText) Then: Exit Function
        
    'MAIL_TO
    If Not fAutoSenderCommandLineBlockReader(tCommandLineRootNode, "mailto", True, tResultString, tCommandSplitter, outErrorText) Then: Exit Function
    
    'MAIL_FROM
    If Not fAutoSenderCommandLineBlockReader(tCommandLineRootNode, "mailfrom", True, tResultString, tCommandSplitter, outErrorText) Then: Exit Function
    
    'AUTH
    If Not fAutoSenderCommandLineBlockReader(tCommandLineRootNode, "auth", inAuthEnabled, tResultString, tCommandSplitter, outErrorText) Then: Exit Function
    
    'SSL_MODE
    If Not fAutoSenderCommandLineBlockReader(tCommandLineRootNode, "sslmode", True, tResultString, tCommandSplitter, outErrorText, inSSLMode) Then: Exit Function
    
    'SIGN
    If Not fAutoSenderCommandLineBlockReader(tCommandLineRootNode, "sign", inSignNeeded, tResultString, tCommandSplitter, outErrorText) Then: Exit Function
    
    'ENCRYPT
    If Not fAutoSenderCommandLineBlockReader(tCommandLineRootNode, "encrypt", inEcryptNeeded, tResultString, tCommandSplitter, outErrorText) Then: Exit Function
    
    'SUBJECT
    If Not fAutoSenderCommandLineBlockReader(tCommandLineRootNode, "subject", inHasSubject, tResultString, tCommandSplitter, outErrorText) Then: Exit Function
    
    'ATTACHMENT
    If Not fAutoSenderCommandLineBlockReader(tCommandLineRootNode, "attachment", inAttachmentNeeded, tResultString, tCommandSplitter, outErrorText) Then: Exit Function
    
    'fin
    outBlankString = tResultString
    fAutoSenderCommandLineConstruct = True
End Function

Public Sub fAutoSenderNew_Test()
    Dim tResult, tResString
    If Not fConfiguratorInit Then: Exit Sub
    If Not fXMLSmartUpdate("BASIS,CREDENTIALS") Then: Exit Sub
    tResult = fAutoSender("BELKAMKO", "haustov@izhenergy.ru", 1, 1, "C:\Users\haustov\GTPCFG\TODO.txt", "Test", tResString)
End Sub

'fAutoSender - Function to send attachments using external sender-intefaces
Private Function fAutoSender(inTraderCode, inAddress, inEncrypt, inSign, inFilePath, inSubject, outErrorText)
    Dim tLogTag, tSimpleMode, tMainNode, tValue, tNode, tXPathString, tMainCFG
    Dim tAuthEnabled, tAuthLogin, tAuthPassword, tSMTPServer, tSMTPPort, tMailFrom, tSSLMode, tSSLVersion, tRoleRequired, tSenderNode, tConnectionTimeout
    Dim tRoleTagSimple, tRoleTagEncrypt, tSignKey, tEncryptKey, tSignNeeded, tEcryptNeeded, tBlankString, tHasSubject, tHasAttachment, tResultValue
    
    'prepare
    fAutoSender = False
    tLogTag = fGetLogTag("fAutoSender")
    outErrorText = vbNullString
    tRoleTagSimple = "simple"
    tRoleTagEncrypt = "encrypt"
    tSignNeeded = inSign = 1
    tEcryptNeeded = inEncrypt = 1
    tHasSubject = inSubject <> vbNullString
    
    'check sources
    If Not gXMLCredentials.Active Then
        uCDebugPrint tLogTag, 2, "Ресурс недоступен - " & tMainCFG.ClassTag
        Exit Function
    End If
    
    'get mainnode
    tXPathString = "//trader[@id='" & inTraderCode & "']/service[(@id='mailbox' and @version='1')]/item[@type='mainsender']"
    Set tMainNode = gXMLCredentials.XML.SelectSingleNode(tXPathString)
    If tMainNode Is Nothing Then
        uCDebugPrint tLogTag, 2, "Не удалось получить основую ноду из <" & tMainCFG.ClassTag & "> для работы! XPath=[" & tXPathString & "]"
        Exit Function
    End If
    
    'get common settings for service
    If Not fAutoSenderGetCommonParams(tMainNode, tAuthEnabled, tAuthLogin, tAuthPassword, tSMTPServer, tSMTPPort, tMailFrom, tSSLMode, tSSLVersion, tConnectionTimeout, outErrorText) Then
        uCDebugPrint tLogTag, 2, outErrorText
        Exit Function
    End If
    
    'Select sender role required
    If tSignNeeded Or tEcryptNeeded Then
        tRoleRequired = tRoleTagEncrypt
    Else
        tRoleRequired = tRoleTagSimple
    End If
    
    'Select sender-node by required role
    If Not fAutoSenderGetSenderByRole(tMainNode, tRoleRequired, tSenderNode, outErrorText) Then
        uCDebugPrint tLogTag, 2, outErrorText
        Exit Function
    End If
    
    'Read keys for encrypt role
    If tRoleRequired = tRoleTagEncrypt Then
        If Not fAutoSenderGetKeysForEncrypt(tMainNode, tSignNeeded, tEcryptNeeded, tSignKey, tEncryptKey, outErrorText) Then
            uCDebugPrint tLogTag, 2, outErrorText
            Exit Function
        End If
    End If
    
    'attachment test
    If inFilePath <> vbNullString Then
        If Not gFSO.FileExists(inFilePath) Then
            uCDebugPrint tLogTag, 2, "Файл-вложение для письма не найден по указанному пути: " & inFilePath
            Exit Function
        End If
        tHasAttachment = True
    Else
        tHasAttachment = False
    End If
    
    'Construct command line blank for sender-interface
    If Not fAutoSenderCommandLineConstruct(tSenderNode, tAuthEnabled, tSSLMode, tSignNeeded, tEcryptNeeded, tHasSubject, tHasAttachment, tBlankString, outErrorText) Then
        uCDebugPrint tLogTag, 2, outErrorText
        Exit Function
    End If
    
    'Replace placeholder with data
    tBlankString = Replace(tBlankString, "##SERVER##", tSMTPServer)
    tBlankString = Replace(tBlankString, "##PORT##", tSMTPPort)
    tBlankString = Replace(tBlankString, "##TIMEOUT##", tConnectionTimeout)
    tBlankString = Replace(tBlankString, "##MAIL-TO##", inAddress)
    tBlankString = Replace(tBlankString, "##MAIL-FROM##", tMailFrom)
    If tAuthEnabled Then
        tBlankString = Replace(tBlankString, "##USER-LOGIN##", tAuthLogin)
        tBlankString = Replace(tBlankString, "##USER-PASSWORD##", tAuthPassword)
    End If
    tBlankString = Replace(tBlankString, "##SSL-VERSION##", tSSLVersion)
    If tSignNeeded Then: tBlankString = Replace(tBlankString, "##SIGN-KEY##", tSignKey)
    If tEcryptNeeded Then: tBlankString = Replace(tBlankString, "##ENCRYPT-KEY##", tEncryptKey)
    If tHasAttachment Then: tBlankString = Replace(tBlankString, "##MAIL-ATTACHMENT-PATH##", inFilePath)
    If tHasSubject Then: tBlankString = Replace(tBlankString, "##MAIL-SUBJECT##", inSubject)
    
    'Use interface to send mail
    tResultValue = gWShell.Run(tBlankString, 0, True)
    If tResultValue <> 0 Then
        uCDebugPrint tLogTag, 2, tBlankString 'строка отправки
        outErrorText = "Ошибка(" & tResultValue & ")! Отправка на адрес <" & inAddress & "> файла <" & inFilePath & "> не удалась!"
        uCDebugPrint tLogTag, 2, outErrorText
        Exit Function
    End If

    'fin
    fAutoSender = True
End Function

Private Function fAutoSenderOld(inTraderID, inAddress, inEncrypt, inSign, inFilePath, inHeader, inResultString) As Boolean
    Dim tLogTag, tCSMFullPath, tGMSFullPath, tEasyMode, tResultValue
    Dim tSMTPServer, tSMTPPort, tMailTo, tTimeOut, tSMTPBlockKey
    Dim tMailFrom, tSSLMode, tAuthEnabled, tAuthLogin, tAuthPassword, tSSLVer, tSSLCheckCert, tSSLCheckCertOnline, tAccountBlockKey
    Dim tSign, tCertSign, tEncrypt, tCertEncrypt, tSignKey, tEncryptKey
    Dim tResultKey, tBodyBlockKey
'00 // Предустановки
    fAutoSender = False
    tEasyMode = False
    tLogTag = "AUTOSND"
    tCSMFullPath = "C:\Users\haustov\Desktop\CSM\CryptoSendMail.exe" 'main
    tGMSFullPath = "C:\Users\haustov\Desktop\CSM\GoogleMailSend.exe" 'easy
'01 // Проверим есть ли файл по указанному пути
    If Not (gFSO.FileExists(inFilePath)) Then
        inResultString = "Отправляемый на адрес <" & inAddress & "> файл <" & inFilePath & "> не обнаружен."
        uADebugPrint tLogTag, inResultString
        Exit Function
    End If
    If Not (gFSO.FileExists(tCSMFullPath)) Then
        inResultString = "Программа отправки CSM не обнаружена по пути <" & tCSMFullPath & ">."
        uADebugPrint tLogTag, inResultString
        Exit Function
    End If
    If Not (gFSO.FileExists(tGMSFullPath)) Then
        inResultString = "Программа отправки GMS не обнаружена по пути <" & tGMSFullPath & ">."
        uADebugPrint tLogTag, inResultString
        Exit Function
    End If
'02 // Настройки почтового сервера
    tSMTPServer = "mail.izhenergy.ru"
    tSMTPPort = "25"
    tMailTo = inAddress
    tTimeOut = 10
'03 // Настройки аккаунта отправки
    tMailFrom = "robot@izhenergy.ru"
    tAuthEnabled = True
    tAuthLogin = "robot"
    tAuthPassword = "Akashi90)"
    tSSLMode = 1 '1 - STARTTLS; 2 - SSL/TLS with port; 0 - disable
    tSSLVer = "auto"
    tSSLCheckCert = "N"
    tSSLCheckCertOnline = "N"
'04 // Настройки сертификатов
    'SIGN
    If inSign = 1 Then
        tSign = "Y"
        tCertSign = cnstCertSign
    Else
        tSign = "N"
        tCertSign = vbNullString
    End If
    'ENCRYPT
    If inEncrypt = 1 Then
        tEncrypt = "Y"
        tCertEncrypt = cnstCertEncrypt
    Else
        tEncrypt = "N"
        tCertEncrypt = vbNullString
    End If
'05 // Решение о применении программы
    If inSign + inEncrypt = 0 Then: tEasyMode = True
'06 // Формирование блоков командной строки сервера
    If tEasyMode Then
        tSMTPBlockKey = "-q -ct " & tTimeOut & " -smtp " & tSMTPServer & " -port " & tSMTPPort & " -t " & tMailTo
    Else
        tSMTPBlockKey = "/smtp_timeout=" & tTimeOut & " /smtp_host=" & tSMTPServer & " /smtp_port=" & tSMTPPort & " /to=" & tMailTo
    End If
    tResultKey = tResultKey & " " & tSMTPBlockKey
'07 // Формирование блоков командной строки аккаунта отправки
    If tEasyMode Then
        tAccountBlockKey = "-f " & tMailFrom
        If tAuthEnabled Then: tAccountBlockKey = tAccountBlockKey & " -auth -user " & tAuthLogin & " -pass " & tAuthPassword
        Select Case tSSLMode
            Case 1: tAccountBlockKey = tAccountBlockKey & " -starttls"
            Case 2: tAccountBlockKey = tAccountBlockKey & " -ssl"
        End Select
    Else
        tAccountBlockKey = "/from=" & tMailFrom
        If tAuthEnabled Then
            tAccountBlockKey = tAccountBlockKey & " /smtp_auth=Y /smtp_user=" & tAuthLogin & " /smtp_password=" & tAuthPassword
        Else
            tAccountBlockKey = tAccountBlockKey & " /smtp_auth=N"
        End If
        Select Case tSSLMode
            Case 1: tAccountBlockKey = tAccountBlockKey & " /ssl_mode=1 /ssl_ver=" & tSSLVer & " /ssl_check_cert=" & tSSLCheckCert & " /ssl_check_cert_online=" & tSSLCheckCertOnline
            Case 2: tAccountBlockKey = tAccountBlockKey & " /ssl_mode=2 /ssl_ver=" & tSSLVer & " /ssl_check_cert=" & tSSLCheckCert & " /ssl_check_cert_online=" & tSSLCheckCertOnline
        End Select
    End If
    tResultKey = tResultKey & " " & tAccountBlockKey
'07 // Формирование блоков серитификатов
    If Not tEasyMode Then
        tSignKey = "/s=" & tSign
        If tCertSign <> vbNullString Then: tSignKey = tSignKey & " /cs=" & tCertSign
        tEncryptKey = "/e=" & tEncrypt
        If tCertEncrypt <> vbNullString Then: tEncryptKey = tEncryptKey & " /es=" & tCertEncrypt
        tResultKey = tResultKey & " " & tSignKey & " " & tEncryptKey
    End If
'08 // Формирование блока содержимого (темы, тела и вложения)
    If tEasyMode Then
        tBodyBlockKey = "-sub """ & inHeader & """ -attach """ & inFilePath & """"
    Else
        tBodyBlockKey = "/subj=""" & inHeader & """ """ & inFilePath & """"
    End If
    tResultKey = tResultKey & " " & tBodyBlockKey
'09 // Формирование финальной строки
    If tEasyMode Then
        tResultKey = """" & tGMSFullPath & """ " & tResultKey
    Else
        tResultKey = """" & tCSMFullPath & """ " & tResultKey
    End If
    tResultValue = gWShell.Run(tResultKey, 0, True)
    If tResultValue <> 0 Then
        uADebugPrint tLogTag, tResultKey 'строка отправки
        inResultString = "Ошибка(" & tResultValue & ")! Отправка на адрес <" & inAddress & "> файла <" & inFilePath & "> не удалась!"
        uADebugPrint tLogTag, inResultString
        Exit Function
    End If
    'uADebugPrint tLogTag, tResultKey
    'uADebugPrint tLogTag, tResultKey
    fAutoSender = True
End Function

'fSendItResourcesCheck - resource check for work of fSendIt
Private Function fSendItResourcesCheck(outErrorText) As Boolean
    
    fSendItResourcesCheck = False
    outErrorText = vbNullString
    
    On Error Resume Next
        
        'init check
        If Not gXML80020DB.Active Then
            outErrorText = "Неожиданная ошибка! " & gXML80020DB.ClassTag & " не активен!"
            On Error GoTo 0
            Exit Function
        End If
        
        'xml exist?
        If gXML80020DB.XML Is Nothing Then
            outErrorText = "Неожиданная ошибка! " & gXML80020DB.ClassTag & " не имеет активной XML!"
            On Error GoTo 0
            Exit Function
        End If
        
        'is it real XML object?
        If TypeName(gXML80020DB.XML) <> "DOMDocument60" Then
            outErrorText = "Неожиданная ошибка! " & gXML80020DB.ClassTag & " неожиданный тип параметра XML:" & TypeName(gXML80020DB.XML) & "!"
            On Error GoTo 0
            Exit Function
        End If
        
        If Err.Number <> 0 Then
            outErrorText = "Ошибка! Источник(" & Err.Source & "): " & Err.Description
            On Error GoTo 0
            Exit Function
        End If
        
    On Error GoTo 0
    
    fSendItResourcesCheck = True
End Function

Private Function fGetSendCommandHeader(inSendCommand, inBlockSplitter, inSubSplitter, outTraderCode, outGTPCode, outDayCommand, outErrorText)
    Dim tBlockElements, tTempElements, tPosIndex
    
    fGetSendCommandHeader = False
    outTraderCode = vbNullString
    outGTPCode = vbNullString
    outDayCommand = vbNullString
    outErrorText = vbNullString
    
    ' 01 /// Empty
    If inSendCommand = vbNullString Then
        outErrorText = "Ошибка синтаксиса команды-посылки! Команда оказалась пуста!"
        Exit Function
    End If
    
    ' 02 /// Block count
    tBlockElements = Split(inSendCommand, inBlockSplitter)
    If UBound(tBlockElements) < 1 Then
        outErrorText = "Ошибка синтаксиса команды-посылки! Элементов ожидалось не менее 2, а получено - " & UBound(tBlockElements) + 1
        Exit Function
    End If
    
     ' 03 /// Read header
    tTempElements = Split(UCase(tBlockElements(0)), inSubSplitter)
    If UBound(tTempElements) <> 1 Then
        outErrorText = "Ошибка команды-посылки! Блок #0 не читаем, ожидалось 2 элемента, а получено - " & UBound(tTempElements) + 1
        Exit Function
    End If
    
    ' 04 // return values
    outTraderCode = tTempElements(0)
    outGTPCode = tTempElements(1)
    
    tPosIndex = InStr(inSendCommand, inBlockSplitter)
    outDayCommand = Right(inSendCommand, Len(inSendCommand) - tPosIndex)
    
    ' XX // return
    fGetSendCommandHeader = True
End Function

Private Function fDayCommandReprocess(inDayCommand, outYear, outMonth, outDay, outDate, outAreaList, outAreaCount, inSubSplitter, inEnumSplitter, outErrorText)
    Dim tElements, tValue
    
    fDayCommandReprocess = False
    outErrorText = vbNullString
    outYear = vbNullString
    outMonth = vbNullString
    outDay = vbNullString
    outDate = 0
    outAreaList = vbNullString
    outAreaCount = -1
    
    ' 01 \\\ Empty
    If inDayCommand = vbNullString Then
        outErrorText = "Ошибка синтаксиса! Команда-дня оказалась пуста!"
        Exit Function
    End If
    
    ' 02 \\\ Struct pass
    tElements = Split(inDayCommand, inSubSplitter)
    If UBound(tElements) <> 1 Then
        outErrorText = "Ошибка синтаксиса! Ожидалось 2 элемента команды дата и список AREA [получено элементов: " & UBound(tElements) + 1 & "]"
        Exit Function
    End If
    
    ' 03 \\\ Date get
    If Not (IsNumeric(tElements(0)) And Len(tElements(0)) = 8) Then
        outErrorText = "Ошибка синтаксиса! Ожидалась дата формата YYYYMMDD [получено: " & tElements(0) & "]"
        Exit Function
    End If
    
    On Error Resume Next
        tValue = DateSerial(Left(tElements(0), 4), Mid(tElements(0), 5, 2), Right(tElements(0), 2))
        If Err.Number <> 0 Then
            outErrorText = "Ошибка синтаксиса! Ожидалась дата формата YYYYMMDD [получено: " & tElements(0) & "] (неправильная дата)"
            On Error GoTo 0
            Exit Function
        Else
            outDay = Right(tElements(0), 2)
            outMonth = Mid(tElements(0), 5, 2)
            outYear = Left(tElements(0), 4)
            outDate = tValue
        End If
    On Error GoTo 0
    
    ' 04 \\\ Get AREA list
    outAreaList = Split(tElements(1), inEnumSplitter)
    outAreaCount = UBound(outAreaList)
        
    fDayCommandReprocess = True
End Function

Private Sub fSortAreaItemsByOutNum(inAreaItems() As TSenderAreaItem, inAreaItemsCount, Optional inSilent = False)
    Dim tSorted, tAreaItemsCount, tIndex, tResultString
    Dim tTempAreaItem As TSenderAreaItem

    'nothing to sort if it's just one element or NONE
    If inAreaItemsCount <= 0 Then: Exit Sub

    tSorted = False
    While Not tSorted
        tSorted = True
        For tIndex = 0 To inAreaItemsCount - 1
            If inAreaItems(tIndex).OutNum > inAreaItems(tIndex + 1).OutNum Then
                tSorted = False
                tTempAreaItem = inAreaItems(tIndex)
                inAreaItems(tIndex) = inAreaItems(tIndex + 1)
                inAreaItems(tIndex + 1) = tTempAreaItem
            End If
        Next
    Wend
    
    'checker? visual
    If Not inSilent Then
        tResultString = "tResultString:"
        For tIndex = 0 To inAreaItemsCount - 1
            tResultString = tResultString & " " & inAreaItems(tIndex).OutNum
            If inAreaItems(tIndex).OutNum > inAreaItems(tIndex + 1).OutNum Then
                tSorted = False
                tResultString = tResultString & " **FAILED**"
            End If
        Next
        tResultString = tResultString & " " & inAreaItems(inAreaItemsCount).OutNum
        uCDebugPrint "fSortAreaItemsByOutNum", 0, tResultString
    End If
End Sub

Private Function fSendLockerCheck()
    fSendLockerCheck = False
    fSendLockerCheck = True
End Function

Private Sub fShowDebugTimer(inTag, ioIndex, ioTick, ioTickSave)
    ioTickSave = GetTickCount - ioTick
    ioTick = GetTickCount
    ioIndex = ioIndex + 1
    Debug.Print inTag & " #" & ioIndex & ": " & ioTickSave
End Sub

Private Function fIsMailAddressEqual(ioRawAddress, inTargetAddress, inMailParamSplitter)
    Dim tElements
    
    fIsMailAddressEqual = False
    
    tElements = Split(ioRawAddress, inMailParamSplitter)
    
    If UBound(tElements) <> 2 Then: Exit Function
    
    ioRawAddress = tElements(0)
    
    'return compare result
    fIsMailAddressEqual = (ioRawAddress = inTargetAddress)
End Function

Private Sub fCheckSendLocks(inAreaItemList() As TSenderAreaItem, inMailAddress, inAreaCurrentIndex, inAreaLoBound, inAreaHiBound, inDate, outLock, inMailParamSplitter)
    Dim tLogTag, tIndex, tMailIndex, tLockType, tLock, tMailListRaw
    
    'pre
    'fCheckSendLocks = False
    tLogTag = "fCheckSendLocks"
    
    'defaults
    outLock = False
    tLock = False
    
    'check all other AIIS Active areas
    For tIndex = inAreaLoBound To inAreaHiBound
        
        If inAreaItemList(tIndex).Active And tIndex <> inAreaCurrentIndex Then
            
            'check MailToList
            If tIndex < inAreaCurrentIndex Then
                
                tLockType = "DOWN"
                For tMailIndex = 0 To inAreaItemList(tIndex).MailToCount
                    tMailListRaw = inAreaItemList(tIndex).MailToList(tMailIndex)
                    
                    If fIsMailAddressEqual(tMailListRaw, inMailAddress, inMailParamSplitter) Then
                        tLock = True
                        uCDebugPrint tLogTag, 1, tLockType & " LOCK! Date=" & inDate & "; AREA=" & inAreaItemList(inAreaCurrentIndex).AreaID & " OUTNUM=" & inAreaItemList(inAreaCurrentIndex).OutNum & "; Locked by [AREA=" & inAreaItemList(tIndex).AreaID & "][OUTNUM=" & inAreaItemList(tIndex).OutNum & "][ADR=" & inMailAddress & "]"
                        Exit For 'L2-For exit
                    End If
                Next
            
            'check MailSentList
            ElseIf tIndex > inAreaCurrentIndex Then
                
                tLockType = "UP"
                For tMailIndex = 0 To inAreaItemList(tIndex).SentCount
                    tMailListRaw = inAreaItemList(tIndex).SentList(tMailIndex)
                    
                    If fIsMailAddressEqual(tMailListRaw, inMailAddress, inMailParamSplitter) Then
                        tLock = True
                        uCDebugPrint tLogTag, 1, tLockType & " LOCK! Date=" & inDate & "; AREA=" & inAreaItemList(inAreaCurrentIndex).AreaID & " OUTNUM=" & inAreaItemList(inAreaCurrentIndex).OutNum & "; Locked by [AREA=" & inAreaItemList(tIndex).AreaID & "][OUTNUM=" & inAreaItemList(tIndex).OutNum & "][ADR=" & inMailAddress & "]"
                        Exit For 'L2-For exit
                    End If
                Next
                
            End If
        End If
                                
        If tLock Then: Exit For
    Next
                            
    'fin
    outLock = tLock
End Sub

'fSendIt - обработчик SendCommand; Принцип его сводится к тому, что ASender формирует команды на рассылку на базе текущей даты и времени
'и указывает для ГТП список дат и активных на эту дату AREA.. а fSendIt функция уже просматривает на эту дату для этой ГТП полученные макеты 80020 в R80020DB
'и уже на базе полученных макетов смотрит какие AREA по SendCommand необходимо послать.. т.е. SendCommand бьёт по площади, а R80020DB смотрит
'по наличию макетов.. какие AREA он может послать сейчас и формиует для них списки рассылки прямо в R80020DB
Private Sub fSendIt(inSendCommand, Optional inBlockSplitter = ":", Optional inSubSplitter = "@", Optional inEnumSplitter = ",", Optional inMailBlockSplitter = ";", Optional inMailParamSplitter = ":", Optional inSendDelay = cnstSendingDealay)
    Dim tLogTag, tMainElements, tTempElements, tTraderID, tGTPID, tDayIndex, tAreaCount, tDate, tYear, tMonth, tDay, tIndex, tSubIndex, tDays, tOutNum, tSorted, tLock, tMailIndex, tMailSubIndex, tMailElement, tTempMailElement, tFilePath, tMailListModified, tSaveActive, tHeader, tResultString
    Dim tXPathString, tNode, tAreaNode, tSectionNode, tSectionID, tErrorText, tDayCommand, tAreaElements, tAreaNodes, tDayCommandElements
    Dim tAreaItems() As TSenderAreaItem
    Dim tAreaItemsCount, tRootNode
    Dim tTempAreaItem As TSenderAreaItem
    Dim tDebugMode, tTick, tTickSave, tTickIndex
        
    tLogTag = "fSendIt"
    tDebugMode = False
    
    If tDebugMode Then
        tTick = GetTickCount
        tTickIndex = 0
    End If
    
'00 // На всякий случай проверим
    If Not fSendItResourcesCheck(tErrorText) Then
        uCDebugPrint tLogTag, 2, tErrorText
        Exit Sub
    End If
    
    'и перезагрузим
    'fReloadXML gXML80020DB.XML, gXML80020DB.Path
    'uCDebugPrint tLogTag, 0, "inSendCommand: " & inSendCommand
    
    If Not fGetSendCommandHeader(inSendCommand, inBlockSplitter, inSubSplitter, tTraderID, tGTPID, tDayCommand, tErrorText) Then
        uCDebugPrint tLogTag, 2, tErrorText
        uCDebugPrint tLogTag, 2, "inSendCommand: " & inSendCommand
        Exit Sub
    End If
    
    tDayCommandElements = Split(tDayCommand, inBlockSplitter)
    tDays = UBound(tDayCommandElements)
    
    'makes rootnode to speedup childnode scan dramatically
    tXPathString = "//trader[@id='" & tTraderID & "']"
    Set tRootNode = gXML80020DB.XML.SelectSingleNode(tXPathString)
    
'03 // Перебор полученных дат
    'tDays = UBound(tMainElements)
    tSaveActive = False 'moved here to decrease saves count
    For tDayIndex = 0 To tDays
        If tDebugMode Then: fShowDebugTimer tLogTag & "-" & tDayIndex, tTickIndex, tTick, tTickSave
        If fDayCommandReprocess(tDayCommandElements(tDayIndex), tYear, tMonth, tDay, tDate, tAreaElements, tAreaCount, inSubSplitter, inEnumSplitter, tErrorText) Then

'04 // Получим текущий статус AREA обозначенный в R80020DB на текущий tDayIndex для текущей ГТП и Продавца
            'get area nodes
            tXPathString = "child::year[@id='" & tYear & "']/month[@id='" & tMonth & "']/day[@id='" & tDay & "']/aiis[@gtpid='" & tGTPID & "']/area"
            Set tAreaNodes = tRootNode.SelectNodes(tXPathString)

            'extract areas from R80020DB for compare
            If tDebugMode Then: fShowDebugTimer tLogTag, tTickIndex, tTick, tTickSave
            tAreaItemsCount = -1
            For Each tAreaNode In tAreaNodes
                tAreaItemsCount = tAreaItemsCount + 1
                ReDim Preserve tAreaItems(tAreaItemsCount)
                fExtractAreaItem tAreaItems(tAreaItemsCount), tAreaNode, inMailBlockSplitter, inMailParamSplitter
                            
                'Error Check
                If tAreaItems(tAreaItemsCount).Error <> vbNullString Then
                    uCDebugPrint tLogTag, 2, " Ошибка при извлечении AREA=" & tAreaItems(tIndex).AreaID & "; GTP=" & tGTPID & "; DATE=" & tDate & " (Источник:" & gXML80020DB.ClassTag & "): " & tAreaItems(tAreaItemsCount).Error
                    Exit Sub
                End If
            Next
            
            'sort areas BY OUTNUM order
            fSortAreaItemsByOutNum tAreaItems, tAreaItemsCount, True
            
            'active resolver
            For tIndex = 0 To tAreaItemsCount
                For tSubIndex = 0 To tAreaCount
                    If tAreaItems(tIndex).AreaID = tAreaElements(tSubIndex) And tAreaItems(tIndex).Status = 0 Then '2021-10-13 Filtering by STATUS
                        tAreaItems(tIndex).Active = True
                        
                        'cosmetic
                        tXPathString = "//version[@status!='closed']/area[@id='" & tAreaItems(tIndex).AreaID & "']/ancestor::section"
                        'Set tSectionNode = gXMLBasis.XML.SelectSingleNode(tXPathString)
                        Set tSectionNode = fGetNodeSafe(gXMLBasis.XML, , , tXPathString)
                        ''Debug.Print ">>" & Not (tSectionNode Is Nothing)
                        fGetAttr tSectionNode, "id", tSectionID
                        tAreaItems(tIndex).SectionID = tSectionID
                        
                        Exit For
                    End If
                Next
            Next
            
            'sening active areas
            For tIndex = 0 To tAreaItemsCount
                
                If tAreaItems(tIndex).Active Then
                    
                    tMailListModified = False
                    If fBuildM80020DropFolder(gXML80020CFG.Path.Processed, tFilePath, tYear, tMonth, tGTPID, False, tErrorText) Then: tFilePath = tFilePath & "\" & tAreaItems(tIndex).FileName
                                
                    'tFilePath = gXML80020CFG.Path.Processed & "\" & tYear & "-" & Format(tMonth, "00") & "\" & tGTPID & "\" & tAreaItems(tIndex).FileName
                    If gFSO.FileExists(tFilePath) Then
                        
                        'for each MailTo item for current AREA
                        For tMailIndex = 0 To tAreaItems(tIndex).MailToCount
                            tMailElement = Split(tAreaItems(tIndex).MailToList(tMailIndex), inMailParamSplitter) 'split to params
                                        
                            'down lock check
                            'tDownLock = False
                            'For tIndexDown = 0 To tIndex - 1
                            '    'tMailSubIndex for mailto
                            '    If tAreaItems(tIndexDown).Active Then
                            '        For tMailSubIndex = 0 To tAreaItems(tIndexDown).MailToCount
                            '            tTempMailElement = Split(tAreaItems(tIndexDown).MailToList(tMailSubIndex), inMailParamSplitter)
                            '            If UBound(tTempMailElement) = 2 Then
                            '                If tTempMailElement(0) = tMailElement(0) Then
                            '                    tDownLock = True
                            '                    uCDebugPrint tLogTag, 1, "DOWN LOCK! AREA=" & tAreaItems(tIndex).AreaID & "; DATE=" & tDate & "; ADR=[" & tMailElement(0) & "]"
                            '                    Exit For
                            '                End If
                            '            End If
                            '        Next
                            '    End If
                            '
                            '    If tDownLock Then: Exit For
                            'Next
                                        
                            'up lock check
                            'tUpLock = False
                            'For tIndexUp = tIndex + 1 To tAreaItemsCount
                            '    If tAreaItems(tIndexUp).Active Then
                            '        For tMailSubIndex = 0 To tAreaItems(tIndexUp).SentCount
                            '            tTempMailElement = Split(tAreaItems(tIndexUp).SentList(tMailSubIndex), inMailParamSplitter)
                            '            If UBound(tTempMailElement) = 2 Then '??? no need?
                            '                If tTempMailElement(0) = tMailElement(0) Then
                            '                    tUpLock = True
                            '                    uCDebugPrint tLogTag, 1, "UP LOCK! AREA=" & tAreaItems(tIndex).AreaID & "; DATE=" & tDate & "; ADR=[" & tMailElement(0) & "]"
                            '                    Exit For
                            '                End If
                            '            End If
                            '        Next
                            '    End If
                            '
                            '    If tUpLock Then: Exit For
                            'Next
                            
                            fCheckSendLocks tAreaItems, tMailElement(0), tIndex, 0, tAreaItemsCount, tDate, tLock, inMailParamSplitter
                            
                            'next step
                            If Not tLock Then
                                tHeader = tAreaItems(tIndex).Class & " P:" & tGTPID & "-" & tAreaItems(tIndex).SectionID & " A:" & tAreaItems(tIndex).AreaID & " N:" & fNZeroAdd(tAreaItems(tIndex).OutNum, 6) & " " & tYear & "-" & tMonth & "-" & tDay
                                If fAutoSender(tTraderID, tMailElement(0), tMailElement(1), tMailElement(2), tFilePath, tHeader, tResultString) Then
                                'If True Then
                                    uCDebugPrint tLogTag, 0, "Отправлено <" & tHeader & "> на <" & tMailElement(0) & ">"
                                    tAreaItems(tIndex).SentCount = tAreaItems(tIndex).SentCount + 1
                                    ReDim Preserve tAreaItems(tIndex).SentList(tAreaItems(tIndex).SentCount)
                                    tAreaItems(tIndex).SentList(tAreaItems(tIndex).SentCount) = tAreaItems(tIndex).MailToList(tMailIndex) 'tAreaItems(tIndex).SentList(tAreaItems(tIndex).SentCount)
                                    tAreaItems(tIndex).MailToList(tMailIndex) = vbNullString
                                    tMailListModified = True
                                    'fGetRecievedTimeStamp
                                End If
                            End If
                        Next 'mailto
                        
                        'adjusting changes
                        If tMailListModified Then
                            fInjectAreaItemChanges tAreaItems(tIndex), tAreaNodes, tSaveActive, inMailBlockSplitter, inMailParamSplitter 'save changes to node
                            'uCDebugPrint tLogTag, 1, "Delay"
                            uSleep inSendDelay 'delay after sending
                        End If
                    Else
                        uCDebugPrint tLogTag, 2, "ERR! NO OUTFILE or UNCOM! AREA=" & tAreaItems(tIndex).AreaID & "; GTP=" & tGTPID & "; DATE=" & tDate
                    End If
                End If
            Next 'area
            
            Erase tAreaItems 'clear afterall
                    
        End If
        If tDebugMode Then: fShowDebugTimer tLogTag & "-" & tDayIndex, tTickIndex, tTick, tTickSave
    Next
    
    'If tDebugMode Then: fShowDebugTimer tLogTag, tTickIndex, tTick, tTickSave
    If tSaveActive Then: fSaveXMLDB gXML80020DB, False, , , , tLogTag & " внесла изменения в рассылки!" 'moved here to decrease saves count
End Sub

'fNZeroAdd
Private Function fNZeroAdd(inValue, inDigiCount)
    Dim tHighStack, tIndex
    fNZeroAdd = inValue
    tHighStack = inDigiCount - Len(inValue)
    If tHighStack > 0 Then
        For tIndex = 1 To tHighStack
            fNZeroAdd = "0" & fNZeroAdd
        Next
    End If
End Function

'fGetGTPNodes - return GTP nodes from BASIS-type XML structure by TraderID (outError - store error status)
Private Function fGetGTPNodes(inXML, inTraderID, outError)
    Dim tXPathString, tNodes
    
    Set fGetGTPNodes = Nothing
    outError = vbNullString
    
    If inXML Is Nothing Then
        outError = "XML object not exists."
        Exit Function
    End If
    
    On Error Resume Next
    
    tXPathString = "//trader[@id='" & inTraderID & "']/gtp"
    Set tNodes = inXML.SelectNodes(tXPathString)
    
    If Err.Number <> 0 Then
        outError = "Unknown Error > Object(" & Err.Source & "): " & Err.Description
        On Error GoTo 0
        Exit Function
    End If
    
    On Error GoTo 0
    
    If tNodes.Length = 0 Then
        outError = "Error! Nodes not found with tXPathString: [" & tXPathString & "]"
        Exit Function
    End If
    
    Set fGetGTPNodes = tNodes
    
End Function

Public Sub fXMLASender()
Dim tLogTag
Dim tGTPNode, tRootNode
Dim tGTPID, tActiveAreaList, tActiveAreaCount, tSenderCommand, tTraderID, tTraderName, tErrorText, tGTPNodes
Dim tTimeZones As TSendTimeZoneList, tDebugMode, tTicker, tTickerSave
Dim tSendGTP As TSendGTPItem

'00 // Подготовка и инициация
    tDebugMode = False
    'If tDebugMode Then: tTicker = GetTickCount
    
    tLogTag = fGetLogTag("fXMLASender")
    uCDebugPrint tLogTag, 0, "Авторассылка макетов 80020\80040 запущена."
    
    If Not fLocalInit Then: Exit Sub
    
    tTraderID = fGetTraderInfo("id")
    tTraderName = fGetTraderInfo("name")
    uCDebugPrint tLogTag, 0, "Инициация рассылки для участника " & tTraderID & " (" & tTraderName & ")"

'01 // Определим периоды рассылки по часовым поясам и ценовым зонам
    If Not (fGetSendPeriod(tTimeZones)) Then: Exit Sub

'02 // Определим корневую ноду для поиска ГТП
    'Set tRootNode = gXMLBasis.XML.SelectSingleNode("//trader[@id='" & tTraderID & "']")
    'проверим есть ли такая нода
    'If tRootNode Is Nothing Then
    '    uADebugPrint tLogTag, "Ошибка! Код участника " & gTraderInfo.ID & " (" & gTraderInfo.Name & ") не найден в BASIS!"
    '    Exit Sub
    'End If
    'проверим есть ли ГТП в этой ноде
    'If tRootNode.ChildNodes.Length = 0 Then
    '    uADebugPrint tLogTag, "Ошибка! Участник " & gTraderInfo.ID & " (" & gTraderInfo.Name & ") не содержит ГТП в BASIS!"
    '    Exit Sub
    'End If
    Set tGTPNodes = fGetGTPNodes(gXMLBasis.XML, tTraderID, tErrorText)
    If tGTPNodes Is Nothing Then
        uCDebugPrint tLogTag, 2, tErrorText
        Exit Sub
    End If

'03 // Перебор по ГТП
    For Each tGTPNode In tGTPNodes 'tRootNode.ChildNodes
        'If tGTPNode.NodeName = "gtp" Then
            ''Debug.Print "1: " & GetTickCount
            If tDebugMode Then: tTicker = GetTickCount
            
            tGTPID = UCase(tGTPNode.GetAttribute("id"))
            If tGTPNode.ChildNodes.Length = 0 Then
                uCDebugPrint tLogTag, 1, "Предупреждение! ГТП " & tGTPID & " не содержит данных!"
            Else
                'Найти связанные ноды в BD80020 по TZ
                tSendGTP.ID = tGTPID
                Set tSendGTP.Node = tGTPNode
                If tDebugMode Then
                    tTickerSave = GetTickCount - tTicker
                    tTicker = GetTickCount
                    ''Debug.Print "GTP=" & tGTPID & " Ticks=" & tTickerSave
                End If
                
                fGetActiveAreaList tSendGTP, tTimeZones
                ''Debug.Print "3: " & GetTickCount
                fAssignSendList tSendGTP '<<
                ''Debug.Print "4: " & GetTickCount
                tSenderCommand = fCreateSenderCommand(tSendGTP, tTraderID)
                'If tGTPID = "PBELKA13" Then: uADebugPrint tLogTag, "SenderCommand=" & tSenderCommand
                'uADebugPrint tLogTag, "SenderCommand=" & tSenderCommand
                ''Debug.Print "5: " & GetTickCount
                fSendIt tSenderCommand
                ''Debug.Print "6: " & GetTickCount
                If tDebugMode Then
                    tTickerSave = GetTickCount - tTicker
                    tTicker = GetTickCount
                    'Debug.Print "6: Ticks=" & tTickerSave
                End If
            End If
            ''Debug.Print "NEXT GTP"
        'End If
    Next
        
'02 // Определим общий период рассылки, для сканирования БД макетов
    'уберем часы, минуты и секунды
'    tStartDate = CDate(Fix(tStartDate))
'    tEndDate = CDate(Fix(tEndDate))
'    uADebugPrint tLogTag, "Итоговый период рассылки определен: " & tStartDateFormated & " - " & tEndDateFormated
'03 // Произведем сканирование БД макетов для каждого дня из определенного раннее периода
    'tXPath = "/message/trader[@id='" & tTraderID & "']/year[@id='" & tYear & "']/month[@id='" & tMonth & "']/day[@id='" & tDay & "']/aiis[@id='" & tAIISCode & "']/area[@id='" & tAreaID & "' and @class='80020']"
    'Set tNodes = gXML80020DB.XML.SelectNodes("")
    uADebugPrint tLogTag, "Over"
End Sub

Public Sub fTestSub()
Dim tString, tElements
    'nothing to test
    'fInit                                    'глобальные переменные
    'If Not fXMLSmartUpdate Then: Exit Sub   'выбор конфигурации базиса
    'MsgBox fWorkDayShift("20180412", -9)
    'MsgBox Format(Now(), "YYYYMMDDhhmmss")
    tString = "A;" 'vbNullString
    tElements = Split(tString, ";")
    'Debug.Print "UBound=" & UBound(tElements)
End Sub

