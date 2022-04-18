Attribute VB_Name = "Configurator"
'ROBOT CONFIG RULER
Option Explicit

Private Const cnstModuleName = "CFG"
Private Const cnstModuleVersion = 1
Private Const cnstModuleDate = "18-01-2019"

Public Type TTraderInfo
    ID As Variant
    Name As Variant
    INN As Variant
End Type

Public Type TSubjectInfo
    ID As Variant
    ParentID As Variant
    Name As String
    ParentName As String
    LocalUTC As Variant
    TradeZoneUTC As Variant
    TradeZoneID As Variant
    TradeMode As Byte
    TimeZoneID As Variant
    Comment As String
    IsReady As Boolean
End Type

Public Type TXSDFile
    Active As Boolean
    Tag As String
    Name As String
    XML As Object
    Path As String
End Type

Public Type TXMLDataBaseFile
    Name As String
    XML As Object
    LastSaveRebuild As Boolean
    Path As String
    ClassTag As String
    Active As Boolean
    Version As Variant
End Type

Public Type TXMLConfigFile
    Active As Boolean
    ClassTag As String
    Name As String
    XML As Object
    ModificationDate As Variant
End Type

Public Type TXML80020Config
    Active As Boolean
    Path As TPathList
    XSD20V2 As TXSDFile
    XSD40V2 As TXSDFile
End Type

Public Type TDataSource
    ArchType As String
    Name As String
    AccessLimit As Boolean
    AccessCurrent As Integer
    Loaded As Boolean
End Type

Public Type TDataSourceList
    Item() As TDataSource
    Count As Variant
End Type

Public gMainInit As TXMLConfigFile
Public gXML80020CFG As TXML80020Config
Public gXMLCalcRoute As TXMLConfigFile
'Public gXMLCalcRoute2 As TXMLConfigFile
Public gTraderInfo As TTraderInfo
Public gXMLBasis As TXMLConfigFile
Public gXMLConverter As TXMLConfigFile
Public gXMLFrame As TXMLConfigFile
Public gXMLCalendar As TXMLConfigFile
Public gXMLDictionary As TXMLConfigFile
Public gXML80020DB As TXMLDataBaseFile
Public gXML30308DB As TXMLDataBaseFile
Public gMailScanDB As TXMLDataBaseFile
Public gXMLCredentials As TXMLConfigFile
Public gF63DB As TXMLDataBaseFile
Public gCalcDB As TXMLDataBaseFile
Public gBRForecastDB As TXMLDataBaseFile
Public gXSDForecast As TXSDFile
Public gFSO, gWShell, gExcel 'объекты
Public gConfigPath, gDataPath
Public gDataSourceList As TDataSourceList
Public gLocalUTC As Variant
Private gOutlookApp As New Outlook.Application
Private gOutlookNameSpace As Object
Private gOutlookRootFolder As Outlook.MAPIFolder
Private gMainAccount As Account
'Public gMainAccountName
'Public gMainDataSourceList
Private gMainMailAccount

Public Function fGetOutlookRootFolder() As Outlook.MAPIFolder
    Set fGetOutlookRootFolder = gOutlookRootFolder
End Function

Public Function fGetMailAccount() As Account
    Set fGetMailAccount = gMainAccount
End Function

Public Function fGetTraderInfo(inField)
    Select Case LCase(inField)
        Case "id":
            fGetTraderInfo = gTraderInfo.ID
        Case "inn":
            fGetTraderInfo = gTraderInfo.INN
        Case "name":
            fGetTraderInfo = gTraderInfo.Name
        Case Else:
            fGetTraderInfo = vbNullString
    End Select
End Function

Private Function fGetLogTag(inTagText)
    fGetLogTag = cnstModuleName & "." & inTagText
End Function

Public Sub fRebuildXML80020DB()
    If Not gXML80020DB.Active Then: Exit Sub
    fSaveXMLDB gXML80020DB, False, , True, True, "Rebuild CALL!"
End Sub

Public Sub fRebuildF63DB()
    If Not gF63DB.Active Then
        If Not fConfiguratorInit Then: Exit Sub
        If Not fXMLSmartUpdate("F63DB") Then: Exit Sub
    End If
    If Not gF63DB.Active Then: Exit Sub
    fSaveXMLDB gF63DB, False, , True, True, "Rebuild CALL!"
End Sub

Public Sub fRebuildCalcDB()
    If Not gCalcDB.Active Then: Exit Sub
    fSaveXMLDB gCalcDB, False, , True, True, "Rebuild CALL!"
End Sub

Public Sub fRebuildBRForecast()
    If Not gBRForecastDB.Active Then
        If Not fConfiguratorInit Then: Exit Sub
        If Not fXMLSmartUpdate("BRFORECAST") Then: Exit Sub
    End If
    fSaveXMLDB gBRForecastDB, False, , True, True, "Rebuild CALL!"
End Sub

Public Sub fSaveXMLChanges(inXML, inPath, Optional fTimeStampEnabled As Boolean = False, Optional inUTF8BomEnabled = False, Optional inUseRebuild = False, Optional inShowDebugInfo = True)
    Dim tNode, tTextFile, tXMLText, tXMLBufText, tUTF8BomMarker, tAddBomMarker, tIndex, tEncodeExists, tVersionExists
    Dim tTickStart, tBOMAddingTicks, tSaveTicks, tRebuildReadTicks, tRebuildRewriteTicks, tRebuildReSaveTicks
    
    ' 01 // Prepare
    tUTF8BomMarker = "п»ї"
    tAddBomMarker = False
    If inShowDebugInfo Then: tTickStart = GetTickCount
    
    ' 02 // UTF8 with BOM
    If inUTF8BomEnabled Then
        tEncodeExists = vbNullString
        tVersionExists = vbNullString
        
        Set tNode = inXML.FirstChild
        
        If tNode.Attributes.Length > 0 Then
            For tIndex = 0 To tNode.Attributes.Length - 1
                If tNode.Attributes(tIndex).NodeName = "encode" Then: tEncodeExists = UCase(tNode.Attributes(tIndex).NodeValue)
                If tNode.Attributes(tIndex).NodeName = "version" Then: tVersionExists = tNode.Attributes(tIndex).NodeValue
            Next
            
            If tVersionExists = "1.0" Then
                If tEncodeExists = vbNullString Or tEncodeExists = "UTF-8" Then: tAddBomMarker = True
            End If
        End If
    End If
    
    If inShowDebugInfo Then: tBOMAddingTicks = GetTickCount - tTickStart
    
    ' 02 // XML Save
    Set tNode = inXML.DocumentElement 'root
    If fTimeStampEnabled Then: tNode.SetAttribute "releasestamp", fGetTimeStamp()
    inXML.Save (inPath)
    
    If inShowDebugInfo Then
        tSaveTicks = GetTickCount - (tTickStart + tBOMAddingTicks)
        tRebuildReadTicks = 0
        tRebuildRewriteTicks = 0
        tRebuildReSaveTicks = 0
    End If
    
    ' IF USING REBUILD
    If inUseRebuild Then
        
        ' 03 // Reading XML as TEXT
        Set tTextFile = gFSO.OpenTextFile(inPath, 1)
        tXMLText = tTextFile.ReadAll
        tTextFile.Close
        
        If inShowDebugInfo Then: tRebuildReadTicks = GetTickCount - (tTickStart + tSaveTicks + tBOMAddingTicks)
        
        ' 04 // Rewriting XML as TEXT with modifications of ><
        Set tTextFile = gFSO.OpenTextFile(inPath, 2, True)
        tXMLText = Replace(tXMLText, "><", "> <")
        tTextFile.Write tXMLText
        tTextFile.Close
        
        If inShowDebugInfo Then: tRebuildRewriteTicks = GetTickCount - (tTickStart + tSaveTicks + tBOMAddingTicks + tRebuildReadTicks)
        
        ' 05 // ReSave XML
        inXML.Load (inPath) 'RESAVE-READ
        inXML.Save (inPath) 'RESAVE-SAVE
        
        If inShowDebugInfo Then: tRebuildReSaveTicks = GetTickCount - (tTickStart + tSaveTicks + tBOMAddingTicks + tRebuildReadTicks + tRebuildRewriteTicks)
    End If
    
    ' 06 // bom
    If tAddBomMarker Then
        Set tTextFile = gFSO.OpenTextFile(inPath, 1)
        tXMLText = tTextFile.ReadAll
        tTextFile.Close
        Set tTextFile = gFSO.OpenTextFile(inPath, 2, True)
        tXMLText = tUTF8BomMarker & tXMLText
        tTextFile.Write tXMLText
        tTextFile.Close
    End If
    
    If inShowDebugInfo Then
        tBOMAddingTicks = tBOMAddingTicks + GetTickCount - (tTickStart + tSaveTicks + tBOMAddingTicks + tRebuildReadTicks + tRebuildRewriteTicks + tRebuildReSaveTicks)
        Debug.Print "fSaveXMLChanges timers (in ms): TOTAL=" & GetTickCount - tTickStart & " / BOMADD=" & tBOMAddingTicks & " / SAVE=" & tSaveTicks & " / RBLDREAD=" & tRebuildReadTicks & " / RBLDRW=" & tRebuildRewriteTicks & " / RBLDReSAVE=" & tRebuildReSaveTicks
    End If
End Sub

Private Function fFolderCreator(inFolderPath, inCreateNewFolder, outTextError)
    fFolderCreator = False
    'work
    If inCreateNewFolder Then
        If Not (uFolderCreate(inFolderPath)) Then
            outTextError = "Не удалось создать папку > " & inFolderPath
            Exit Function
        End If
    Else
        If Not (uFileExists(inFolderPath)) Then
            outTextError = "Не удалось обнаружить папку > " & inFolderPath
            Exit Function
        End If
    End If
    'over
    fFolderCreator = True
End Function

Public Function fOpenXML80020(outXMLObject, inFilePath, inValidate, outVersion, outClass, outStructureStatus, outErrorText)
Dim tLogTag, tXMLDoc, tClass, tVersion, tRootNode, tError, tValue, tClassTagA20, tClassTagA40, tClassLock
    tLogTag = "OpenXML80020"
    fOpenXML80020 = 0
    outVersion = 0
    outClass = vbNullString
    outErrorText = vbNullString
    tClassTagA20 = "80020"
    tClassTagA40 = "80040"
    outStructureStatus = -1 'unknown
    Set outXMLObject = Nothing
'01 // Работа с путём файла
    If Not gFSO.FileExists(inFilePath) Then
        fOpenXML80020 = 1
        outErrorText = "Не обнаружен файл: " & inFilePath
        uCDebugPrint tLogTag, 2, outErrorText
        Exit Function
    End If
'02 // Загрузка целевого XML для распознания
    Set tXMLDoc = CreateObject("Msxml2.DOMDocument.6.0")
    tXMLDoc.ASync = False
    tXMLDoc.Load inFilePath
'03 // Проверим итоги парсинга
    If tXMLDoc.parseError.ErrorCode = 0 Then 'Parsed?
'04 // Поиск первичных признаков класса
        tClass = "unknown"
        tVersion = "unknown"
        Set tRootNode = tXMLDoc.DocumentElement
        If LCase(tRootNode.NodeName) = "message" Then
            'class
            tValue = UCase(tRootNode.GetAttribute("class"))
            If Not IsNull(tValue) Then
                If tValue = tClassTagA20 Or tValue = tClassTagA40 Then
                    tClass = tValue
                    outClass = tValue
                End If
            End If
            'version
            tValue = UCase(tRootNode.GetAttribute("version"))
            If Not IsNull(tValue) Then
                tVersion = tValue
            End If
        End If
        'resumer
        'a20
        tClassLock = False
        If tClass = tClassTagA20 Then
            Select Case tVersion
                Case "2", 2: tClassLock = True
            End Select
        End If
        'a40
        If tClass = tClassTagA40 Then
            Select Case tVersion
                Case "2", 2: tClassLock = True
            End Select
        End If
        'unknown class
        If Not tClassLock Then
            fOpenXML80020 = 3
            outErrorText = "Ошибка парсинга: объет не является допустимым классом [" & tClass & "] версией [" & tVersion & "]!"
            uCDebugPrint tLogTag, 2, outErrorText
            Set tXMLDoc = Nothing
            Exit Function
        End If
    Else
        fOpenXML80020 = 2
        outErrorText = "Ошибка парсинга" & tXMLDoc.parseError.ErrorCode & " [LINE:" & tXMLDoc.parseError.Line & "/POS:" & tXMLDoc.parseError.LinePos & "]: " & tXMLDoc.parseError.Reason
        uCDebugPrint tLogTag, 2, outErrorText
        Set tXMLDoc = Nothing
        Exit Function
    End If
'05 // Сверка по схеме
    If inValidate Then
        'A20
        If tClass = tClassTagA20 Then
            If (Not gXML80020CFG.Active) Or (gXML80020CFG.XSD20V2.XML Is Nothing) Then
                fOpenXML80020 = 4
                outErrorText = "Ошибка проверки по XSD: XSD не загружен!"
                uCDebugPrint tLogTag, 2, outErrorText
                Set tXMLDoc = Nothing
                Exit Function
            End If
            'привязка схемы
            Set tXMLDoc.Schemas = gXML80020CFG.XSD20V2.XML
        End If
        'A40
        If tClass = tClassTagA40 Then
            If (Not gXML80020CFG.Active) Or (gXML80020CFG.XSD40V2.XML Is Nothing) Then
                fOpenXML80020 = 4
                outErrorText = "Ошибка проверки по XSD: XSD не загружен!"
                uCDebugPrint tLogTag, 2, outErrorText
                Set tXMLDoc = Nothing
                Exit Function
            End If
            'привязка схемы
            Set tXMLDoc.Schemas = gXML80020CFG.XSD40V2.XML
        End If
        'привязалась ли схема
        If Not IsNull(tXMLDoc.Schemas) Then
            Set tError = tXMLDoc.Validate()
            If tError.ErrorCode <> 0 Then
                outStructureStatus = 1 'damaged
                fOpenXML80020 = 6
                outErrorText = "Ошибка проверки по XSD: " & tError.ErrorCode & " [LINE:" & tError.Line & "]: " & tError.Reason
                uCDebugPrint tLogTag, 2, outErrorText
                Set tXMLDoc = Nothing
                Exit Function
            Else
                outStructureStatus = 0 'normal
            End If
        'XML без схемы по умолчанию считаются правильными (не безопасно, но что поделать)
        Else
            fOpenXML80020 = 5
            outErrorText = "Ошибка проверки по XSD: XSD не привязалась!"
            uCDebugPrint tLogTag, 2, outErrorText
            Set tXMLDoc = Nothing
            Exit Function
        End If
    End If
'06 // Завершение
    Set outXMLObject = tXMLDoc
    outVersion = tVersion
    outClass = tClass
    Set tXMLDoc = Nothing
End Function

'Строит путь для сброса макета 80020 в виде >> КорневаяПапка \ ГГГГ-ММ \ КодГТП
Public Function fBuildM80020DropFolder(inParentFolder, outFolder, inYear, inMonth, inGTPID, inCreateFolder, outTextError)
    fBuildM80020DropFolder = False
    outTextError = vbNullString
    'Check
    If Not (uFileExists(inParentFolder)) Then
        outTextError = "Путь не существует > " & inParentFolder
        Exit Function
    End If
    'Level 1 - Date route
    outFolder = inParentFolder & "\" & inYear & "-" & Format(inMonth, "00")
    If Not fFolderCreator(outFolder, inCreateFolder, outTextError) Then: Exit Function
    'Level 2 - GTPID route
    outFolder = outFolder & "\" & inGTPID
    If Not fFolderCreator(outFolder, inCreateFolder, outTextError) Then: Exit Function
    'Over
    fBuildM80020DropFolder = True
End Function

Public Function fXMLDBManualRebuilder()
    Dim tLogTag
    
    tLogTag = fGetLogTag("fSaveXMLDB")
    uCDebugPrint tLogTag, 0, "Manual XMLDB request..."
    
    fSaveXMLDB gMailScanDB, False, , True, True, "fXMLDBManualRebuilder excecuted!"
    fSaveXMLDB gXML80020DB, False, , True, True, "fXMLDBManualRebuilder excecuted!"
    fSaveXMLDB gXML30308DB, False, , True, True, "fXMLDBManualRebuilder excecuted!"
    
    uCDebugPrint tLogTag, 0, "Manual XMLDB request complete!"
    'fSaveXMLDB gF63DB, False, , True, True, "fXMLDBManualRebuilder excecuted!"
    'fSaveXMLDB gCalcDB, False, , True, True, "fXMLDBManualRebuilder excecuted!"
    'fSaveXMLDB gBRForecastDB, False, , True, True, "fXMLDBManualRebuilder excecuted!"
End Function

' fSaveXMLDB - save predefined XML DB structure to file
Public Sub fSaveXMLDB(inXML As TXMLDataBaseFile, Optional inSilent = True, Optional inUTF8BomEnabled = False, Optional inUseRebuild = False, Optional inShowDebugInfo = False, Optional inComment = vbNullString)
    Dim tLogTag, tComment
    tLogTag = fGetLogTag("fSaveXMLDB")
    inXML.LastSaveRebuild = inUseRebuild
    If inXML.Active Then
        fSaveXMLChanges inXML.XML, inXML.Path, True, inUTF8BomEnabled, inUseRebuild, inShowDebugInfo
        If Not inSilent Then
            If inComment <> vbNullString Then: tComment = " Причина: " & inComment
            uCDebugPrint tLogTag, 0, "XMLBD <" & inXML.ClassTag & "> сохранен!" & tComment
        End If
    Else
        uCDebugPrint tLogTag, 1, "XMLBD <" & inXML.ClassTag & "> недоступен!"
    End If
End Sub

' fReloadXMLDB - reload predefined XML DB structure from file (rollback changes which was not saved - main use reason)
Public Sub fReloadXMLDB(inXML As TXMLDataBaseFile, Optional inSilent = True, Optional inComment = vbNullString)
Dim tLogTag, tComment
    tLogTag = fGetLogTag("fReloadXMLDB")
    If inXML.Active Then: inXML.XML.Load inXML.Path
    If Not inSilent Then
        If inComment <> vbNullString Then: tComment = " Причина: " & inComment
        uCDebugPrint tLogTag, 0, "XMLBD <" & inXML.ClassTag & "> перезагружен!" & tComment
    End If
End Sub

Public Function fConfiguratorStop()
    On Error Resume Next
        If Not gExcel Is Nothing Then
            'fExcelControl 1, 1, 1, 1
            gExcel.Quit
        End If
    On Error GoTo 0
    Set gFSO = Nothing
    Set gExcel = Nothing
    Set gWShell = Nothing
End Function

Public Function fGetMainMailAccount(inMailAccountName, Optional inReInit = True) As Outlook.MAPIFolder
    Dim tLogTag, tErrorText
    
    'prep
    Set fGetMainMailAccount = Nothing
    tLogTag = "fGetMainMailAccount"
    
    uCDebugPrint tLogTag, 0, "Requesting mailbox: " & inMailAccountName
    
    'check and get
    If Not fIsMainMailAccountAvailable(tErrorText) Then
        uCDebugPrint tLogTag, 2, "Mail account not available!"
        
        If Not inReInit Then: Exit Function
        
        'try reinit
        If Not fReInitMainMailAccount() Then
            uCDebugPrint tLogTag, 2, "ReInit account [" & fGetMainMailAccountAsString & "] failed!"
            Exit Function
        End If
            
        'last check
        If Not fIsMainMailAccountAvailable(tErrorText) Then
            uCDebugPrint tLogTag, 2, "Mail account still not available, but ReInit account [" & fGetMainMailAccountAsString & "] processed!??"
            Exit Function
        End If
            
        'finalize
        uCDebugPrint tLogTag, 1, "ReInit account [" & fGetMainMailAccountAsString & "] processed!"
    End If
    
    'check names
    '.DeliveryStore.GetDefaultFolder(olFolderInbox)
    If LCase(gMainAccount.DisplayName) <> LCase(inMailAccountName) Then
        uCDebugPrint tLogTag, 2, "Account name check failed! gMainAccount.DisplayName=[" & gMainAccount.DisplayName & "]; inMailAccountName=[" & inMailAccountName & "]; fGetMainMailAccountAsString=[" & fGetMainMailAccountAsString & "]"
        Exit Function
    End If
    
    'return MAPI folder
    Set fGetMainMailAccount = gOutlookRootFolder
End Function

'Outlook.MAPIFolder
Public Function fIsMainMailAccountAvailable(outErrorText) As Boolean
    Dim tValue, tSubLogTag
    
    fIsMainMailAccountAvailable = False
    tSubLogTag = "[fIsMainMailAccountAvailable] "
    outErrorText = vbNullString
    
    On Error Resume Next
        tValue = gOutlookRootFolder.Items.Count
        If Err.Number <> 0 Then
            outErrorText = tSubLogTag & "Main mail account not available! Desc: " & Err.Description
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
    On Error GoTo 0
    
    fIsMainMailAccountAvailable = True
End Function

Public Function fReInitMainMailAccount()
    fReInitMainMailAccount = fInitMainMailAccount(gMainMailAccount)
End Function

Public Function fInitMainMailAccount(inAccountName)
    Dim tLogTag, tAccount, tMainAccountExists, tTempAccountName
    
    ' Инициирует: gMainAccount, gOutlookRootFolder, gOutlookNameSpace, gMainAccountName
    
    ' 01 // Подготовка переменных
    fInitMainMailAccount = False
    tLogTag = fGetLogTag("INITMAIL")
    
    tTempAccountName = LCase(inAccountName)
    
    ' 02 // Просмотр доступных аккаунтов почты
    For Each tAccount In Application.Session.Accounts
        If LCase(tAccount.DisplayName) = tTempAccountName Then
            tMainAccountExists = True
            Set gMainAccount = tAccount
            Exit For
        End If
    Next
    
    ' 03 // Задание глобальным переменным значений по выбранному аккаунту
    If tMainAccountExists Then
        uCDebugPrint tLogTag, 0, "Основной аккаунт определен как <" & gMainAccount.DisplayName & ">."
        
        'safe mode // 2021-08-12
        On Error Resume Next
            Set gOutlookNameSpace = gMainAccount.DeliveryStore
            Set gOutlookRootFolder = gOutlookNameSpace.GetDefaultFolder(olFolderInbox)
            
            'safe check
            If Err.Number <> 0 Then
                uCDebugPrint tLogTag, 2, "Попытка доступа к каталогам неудачна! Возможно нет связи с почтовым сервером."
                uCDebugPrint tLogTag, 2, "Ошибка (" & Err.Number & ") в " & Err.Source & ": " & Err.Description
                On Error GoTo 0
                Exit Function
            End If
        On Error GoTo 0
    Else
        uCDebugPrint tLogTag, 2, "Основной аккаунт <" & tTempAccountName & "> не найден! Проверьте настройки."
        Exit Function
    End If
    
    ' 04 // Успешный выход
    fInitMainMailAccount = True
End Function

'
Public Function fGetDataSourceIndexByTag(inTag)
    Dim tIndex
    
    'On Error Resume Next
    
    fGetDataSourceIndexByTag = -1
    'If Not gDataSourceList.Count > 0 Then: Exit Function
    'tTagValue = UCase(inTag)
        
    For tIndex = 0 To gDataSourceList.Count
        If gDataSourceList.Item(tIndex).Name = inTag Then
            fGetDataSourceIndexByTag = tIndex
            Exit For
        End If
    Next
        
    'On Error GoTo 0
End Function

Public Function fGetLocalUTC(inDefault)
    Dim tItem, tBiasValue, tItemCounter
    
    fGetLocalUTC = inDefault
    
    tItemCounter = 0
    For Each tItem In GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select * from Win32_TimeZone")
        tBiasValue = tItem.Bias
        tItemCounter = tItemCounter + 1
    Next
    
    If tItemCounter > 0 Then: fGetLocalUTC = Fix(tBiasValue / 60)
End Function

Private Function fGetHomeBasedFolder(inSubFolderName)
    fGetHomeBasedFolder = Environ("HOMEPATH")
    If inSubFolderName <> vbNullString Then: fGetHomeBasedFolder = fGetHomeBasedFolder & "\" & inSubFolderName
    If Not gFSO.FolderExists(fGetHomeBasedFolder) Then: fGetHomeBasedFolder = vbNullString
End Function

Private Function fAddSourceObject_CFG(outSourceObject As TXMLConfigFile, inAccessLimit, inTagName, inFileName)
    gDataSourceList.Count = gDataSourceList.Count + 1
    ReDim Preserve gDataSourceList.Item(gDataSourceList.Count)
    
    With gDataSourceList.Item(gDataSourceList.Count)
        .AccessCurrent = 0
        .AccessLimit = inAccessLimit
        .ArchType = "CFG"
        .Loaded = False
        .Name = UCase(inTagName)
        '---
        outSourceObject.ClassTag = .Name
        outSourceObject.Name = inFileName
        outSourceObject.Active = False
    End With
End Function

Private Function fAddSourceObject_DB(outSourceObject As TXMLDataBaseFile, inAccessLimit, inTagName, inFileName)
    gDataSourceList.Count = gDataSourceList.Count + 1
    ReDim Preserve gDataSourceList.Item(gDataSourceList.Count)
    
    With gDataSourceList.Item(gDataSourceList.Count)
        .AccessCurrent = 0
        .AccessLimit = inAccessLimit
        .ArchType = "DB"
        .Loaded = False
        .Name = UCase(inTagName)
        '---
        outSourceObject.ClassTag = .Name
        outSourceObject.Name = inFileName
        outSourceObject.Active = False
    End With
End Function

Private Function fAddSourceObject_XSD(outSourceObject As TXSDFile, inAccessLimit, inTagName, inFileName)
    gDataSourceList.Count = gDataSourceList.Count + 1
    ReDim Preserve gDataSourceList.Item(gDataSourceList.Count)
    
    With gDataSourceList.Item(gDataSourceList.Count)
        .AccessCurrent = 0
        .AccessLimit = inAccessLimit
        .ArchType = "XSD"
        .Loaded = False
        .Name = UCase(inTagName)
        '---
        outSourceObject.Tag = .Name
        outSourceObject.Name = inFileName
        outSourceObject.Active = False
    End With
End Function

Public Function fGetMainMailAccountAsString()
    Dim tLogTag, tNode, tValue, tXPathString
    
    tLogTag = fGetLogTag("fGetMainMailAccount")
    fGetMainMailAccountAsString = vbNullString
    
    If Not gMainInit.Active Then
        uCDebugPrint tLogTag, 2, "Файл конфигурации <" & gMainInit.ClassTag & "> не загружен! [" & gMainInit.Name & "]"
        Exit Function
    End If
    
    tXPathString = "//trader/mail"
    Set tNode = gMainInit.XML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        uCDebugPrint tLogTag, 2, "Файл конфигурации <" & gMainInit.ClassTag & "> не содержит ноды почты! [" & gMainInit.Name & "] XPath=" & tXPathString
        Exit Function
    End If
    
    fGetMainMailAccountAsString = Trim(tNode.Text)
End Function

Public Sub fConfigCall_Test()
    uCDebugPrint "TEST", 0, "##CALLING-START"
    fConfiguratorInit True
    uCDebugPrint "TEST", 0, "##CALLING-OVER"
End Sub

Public Function fConfiguratorInit(Optional inForceInit As Boolean = False)
    Dim tLogTag, tErrorText
    
    fConfiguratorInit = False
    tLogTag = fGetLogTag("INI")
    
    If inForceInit Or IsEmpty(gFSO) Then
        'variables
        gLocalUTC = fGetLocalUTC(4)
        uD2SInit
        
        'objects
        Set gFSO = CreateObject("Scripting.FileSystemObject")
        Set gWShell = CreateObject("WScript.Shell")
        Set gExcel = CreateObject("Excel.Application")
        
        'paths
        gConfigPath = fGetHomeBasedFolder("GTPCFG")
        If gConfigPath = vbNullString Then
            uCDebugPrint tLogTag, 2, "Не удалось найти папку конфигурации <gConfigPath> по пути: " & gConfigPath
            Exit Function
        End If
                
        gDataPath = fGetHomeBasedFolder("Data")
        If gDataPath = vbNullString Then
            uCDebugPrint tLogTag, 2, "Не удалось найти папку данных <gDataPath> по пути: " & gDataPath
            Exit Function
        End If
        
        uCDebugPrint tLogTag, 0, "Папка конфигурации <gConfigPath> определена по пути: " & gConfigPath
        uCDebugPrint tLogTag, 0, "Папка данных <gDataPath> определена по пути: " & gDataPath
        
        'конфиги и ресурсы
        gDataSourceList.Count = -1
        
        fXML80020CFGInit gDataPath
        
        fAddSourceObject_CFG gMainInit, False, "INIT", "Init.xml"
        fAddSourceObject_CFG gXMLBasis, False, "BASIS", "Basis.xml"
        fAddSourceObject_CFG gXMLConverter, False, "CONVERTER", "Converter.xml"
        fAddSourceObject_CFG gXMLFrame, False, "FRAME", "Frame.xml"
        fAddSourceObject_CFG gXMLCalendar, False, "CALENDAR", "Calendar.xml"
        fAddSourceObject_DB gXML80020DB, True, "R80020DB", "R80020.xml"
        fAddSourceObject_DB gMailScanDB, True, "MAILSCAN", "MailScan.xml"
        fAddSourceObject_XSD gXML80020CFG.XSD20V2, False, "XSD80020V2", "M80020V2.xsd"
        fAddSourceObject_XSD gXML80020CFG.XSD40V2, False, "XSD80040V2", "M80040V2.xsd"
        fAddSourceObject_XSD gXSDForecast, False, "XSDBRFORECAST", "BRForecast.xsd"
        fAddSourceObject_CFG gXMLDictionary, False, "DICTIONARY", "Dictionary.xml"
        fAddSourceObject_DB gBRForecastDB, True, "BRFORECAST", "BRForecastDB.xml"
        fAddSourceObject_DB gXML30308DB, True, "R30308DB", "R30308.xml"
        fAddSourceObject_DB gCalcDB, True, "CALCDB", "CalcDB.xml"
        fAddSourceObject_DB gF63DB, True, "F63DB", "F63DB.xml"
        fAddSourceObject_CFG gXMLCredentials, False, "CREDENTIALS", "Credentials.xml"
        fAddSourceObject_CFG gXMLCalcRoute, False, "CALCROUTE", "CalcRoute.xml"
        
        'попытка загрузки основного конфига
        If Not fXMLSmartUpdate("INIT") Then
            uCDebugPrint tLogTag, 2, "Не удалось загрузить основную конфигурацию <" & gMainInit.ClassTag & ">!"
            Exit Function
        End If
                
        'почтовые аккаунты
        gMainMailAccount = fGetMainMailAccountAsString()
        If gMainMailAccount = vbNullString Then
            uCDebugPrint tLogTag, 2, "Не удалось прочитать основной почтовый ящик из основной конфигурации <" & gMainInit.ClassTag & ">!"
            Exit Function
        End If
        
        If Not fInitMainMailAccount(gMainMailAccount) Then
            uCDebugPrint tLogTag, 2, "Основной аккаунт почты <" & gMainMailAccount & "> не удалось определить! Проверьте настройки."
            Exit Function
        End If

        'trader-info
        If Not fReadTraderInfo(gTraderInfo, tErrorText) Then
            uCDebugPrint tLogTag, 2, "Не удалось прочитать параметры из основной конфигурации <" & gMainInit.ClassTag & ">!"
            Exit Function
        End If
        uCDebugPrint tLogTag, 0, "Инициализирована конфигурация для <" & gTraderInfo.Name & "> код торговца <" & gTraderInfo.ID & "> ИНН[" & gTraderInfo.INN & "]."
    End If
    fConfiguratorInit = True
End Function

Private Function fReadTraderInfo(outTraderInfo As TTraderInfo, outErrorText)
    Dim tXPathString, tNode, tValue
    'prepare
    fReadTraderInfo = False
    outErrorText = vbNullString
    
    'default
    outTraderInfo.ID = vbNullString
    outTraderInfo.INN = vbNullString
    outTraderInfo.Name = vbNullString
    
    'check cfg
    If Not gMainInit.Active Then
        outErrorText = "Файл конфигурации <" & gMainInit.ClassTag & "> не загружен! [" & gMainInit.Name & "]"
        Exit Function
    End If
    
    'CODE
    tXPathString = "//trader/code"
    Set tNode = gMainInit.XML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        outErrorText = "Файл конфигурации <" & gMainInit.ClassTag & "> не содержит искомой ноды! [" & gMainInit.Name & "] XPath=" & tXPathString
        Exit Function
    End If
    outTraderInfo.ID = Trim(tNode.Text)
    
    'INN
    tXPathString = "//trader/inn"
    Set tNode = gMainInit.XML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        outErrorText = "Файл конфигурации <" & gMainInit.ClassTag & "> не содержит искомой ноды! [" & gMainInit.Name & "] XPath=" & tXPathString
        Exit Function
    End If
    outTraderInfo.INN = Trim(tNode.Text)
    
    'Name
    tXPathString = "//trader/name"
    Set tNode = gMainInit.XML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        outErrorText = "Файл конфигурации <" & gMainInit.ClassTag & "> не содержит искомой ноды! [" & gMainInit.Name & "] XPath=" & tXPathString
        Exit Function
    End If
    outTraderInfo.Name = Trim(tNode.Text)
    
    'fin
    fReadTraderInfo = True
End Function

Public Function fXSDLoader(inXSD As TXSDFile, inPath, Optional inSilentMode = True) As Boolean
Dim tTempXMLDoc, tXSDPath, tLogTag
    tLogTag = fGetLogTag("XSDLoader")
    If inXSD.XML Is Nothing Or Not (inXSD.Active) Then
        'temp xml objecth gDataSourceList.Item(gDataSourceList.Count)
        inXSD.Active = True
        tXSDPath = inPath & "\" & inXSD.Name
        uCDebugPrint tLogTag, 0, "Загрузка XSD-схемы " & inXSD.Tag & "(" & inXSD.Name & "); Путь [" & tXSDPath & "]", inSilentMode
        If uFileExists(tXSDPath) Then 'File Exists?
            Set tTempXMLDoc = CreateObject("Msxml2.DOMDocument.6.0")
            tTempXMLDoc.ASync = False
            tTempXMLDoc.Load tXSDPath
            If tTempXMLDoc.parseError.ErrorCode = 0 Then 'Parsed?
                Set inXSD.XML = CreateObject("Msxml2.XMLSchemaCache.6.0")
                inXSD.XML.Add "", tTempXMLDoc
                inXSD.Path = tXSDPath
                uCDebugPrint tLogTag, 0, "Загрузка XSD-схемы " & inXSD.Tag & "(" & inXSD.Name & ") успешна."
            Else
                uCDebugPrint tLogTag, 0, "Загрузка XSD-схемы " & inXSD.Tag & "(" & inXSD.Name & ") неудачна! Ошибка парсинга!"
                inXSD.Active = False
                Set inXSD.XML = Nothing
            End If
        Else
            uCDebugPrint tLogTag, 0, "Загрузка XSD-схемы " & inXSD.Tag & "(" & inXSD.Name & ") неудачна! Файл не найден!"
            inXSD.Active = False
            Set inXSD.XML = Nothing
        End If
    End If
    fXSDLoader = inXSD.Active
End Function

Public Function fXMLConfigLoader(inXMLConfig As TXMLConfigFile, inPath, Optional inSilentMode = True) As Boolean
Dim tTempXMLDoc, tXMLPath, tNode, tValue, tActivePath, tTempModificationDate, tLogTag
    'ФАЗА ПОДГОТОВКИ
    tLogTag = fGetLogTag("XMLCFGLoader")
    tActivePath = vbNullString
    tXMLPath = inPath & "\" & inXMLConfig.Name
    uCDebugPrint tLogTag, 0, "Загрузка конфига " & inXMLConfig.ClassTag & "(" & inXMLConfig.Name & "); Путь [" & tXMLPath & "]", inSilentMode
    
    'ФАЗА 1 >> Чтение файла
    If uFileExists(tXMLPath) Then 'File Exists?
        Set tTempXMLDoc = CreateObject("Msxml2.DOMDocument.6.0")
        tTempXMLDoc.ASync = False
        tTempXMLDoc.Load tXMLPath
        If tTempXMLDoc.parseError.ErrorCode = 0 Then 'Parsed?
            Set tNode = tTempXMLDoc.DocumentElement 'root
            tValue = tNode.NodeName
            If tValue = "message" Then 'message?
                tValue = UCase(tNode.GetAttribute("class"))
                If tValue = inXMLConfig.ClassTag Then 'message class is correct?
                    tValue = UCase(tNode.GetAttribute("releasestamp"))
                    If fCheckTimeStamp(tValue) Then 'release stamp correct?
                        tActivePath = tXMLPath
                        tTempModificationDate = CLngLng(tValue)
                    Else
                        uCDebugPrint tLogTag, 0, "Чтение конфига " & inXMLConfig.ClassTag & "(" & inXMLConfig.Name & ") > ReleaseStamp не верен!"
                    End If
                Else
                    uCDebugPrint tLogTag, 0, "Чтение конфига " & inXMLConfig.ClassTag & "(" & inXMLConfig.Name & ") > ClassTag не верен!"
                End If
            Else
                uCDebugPrint tLogTag, 0, "Чтение конфига " & inXMLConfig.ClassTag & "(" & inXMLConfig.Name & ") > корневой блок не [message]!"
            End If
        Else
            uCDebugPrint tLogTag, 0, "Чтение конфига " & inXMLConfig.ClassTag & "(" & inXMLConfig.Name & ") > ошибка парсинга > " & tTempXMLDoc.parseError.Reason
        End If
    End If
    
    'ФАЗА 2 >> Загрузка файла
    If tActivePath = vbNullString Then
        uCDebugPrint tLogTag, 0, "Чтение конфига " & inXMLConfig.ClassTag & "(" & inXMLConfig.Name & ") > Нет доступных файлов!"
    Else
        If IsEmpty(inXMLConfig.XML) Or (tTempModificationDate > inXMLConfig.ModificationDate) Or Not (inXMLConfig.Active) Then
            'uCDebugPrint tLogTag, 0, "Обновленный путь > " & tActivePath
            uCDebugPrint tLogTag, 0, "Обновление или перезагрузка конфига " & inXMLConfig.ClassTag & "(" & inXMLConfig.Name & ")! Обновленный штамп [" & tTempModificationDate & "]"
            Set inXMLConfig.XML = CreateObject("Msxml2.DOMDocument.6.0")
            inXMLConfig.XML.ASync = False
            inXMLConfig.XML.Load (tActivePath)
            inXMLConfig.ModificationDate = tTempModificationDate
        Else
            uCDebugPrint tLogTag, 0, "Нет обновлений конфига " & inXMLConfig.ClassTag & "(" & inXMLConfig.Name & ")." & inXMLConfig.ClassTag, inSilentMode
        End If
    End If
    
    'ФАЗА 3 >> Управление флагом
    If inXMLConfig.XML Is Nothing Then
        inXMLConfig.Active = False
    Else
        inXMLConfig.Active = True
    End If
    
    fXMLConfigLoader = inXMLConfig.Active
End Function

Public Function fXMLDBLoader(inXMLDB As TXMLDataBaseFile, inPath, Optional inSilentMode = True) As Boolean
Dim tActivePath, tXMLPath, tTempXMLDoc, tNode, tValue, tVersion, tLogTag
    'If inXMLDB.XML Is Nothing Or Not (inXMLDB.Active) Then
        tLogTag = fGetLogTag("XMLBDLoader")
        tActivePath = vbNullString
        tVersion = 0
        tXMLPath = inPath & "\" & inXMLDB.Name
        inXMLDB.LastSaveRebuild = False
        uCDebugPrint tLogTag, 0, "Загрузка БД " & inXMLDB.ClassTag & "(" & inXMLDB.Name & "); Путь [" & tXMLPath & "]", inSilentMode
        'uDebugPrint "XMLBDLoader: Поиск " & inXMLDB.ClassTag & " по пути > " & tXMLPath
        'LOCK FILE
        If uFileExists(tXMLPath) Then 'File Exists?
            Set tTempXMLDoc = CreateObject("Msxml2.DOMDocument.6.0")
            tTempXMLDoc.ASync = False
            tTempXMLDoc.Load tXMLPath
            If tTempXMLDoc.parseError.ErrorCode = 0 Then 'Parsed?
                Set tNode = tTempXMLDoc.DocumentElement
                tValue = tNode.NodeName
                If tValue = "message" Then 'message?
                    tValue = UCase(tNode.GetAttribute("class"))
                    If tValue = inXMLDB.ClassTag Then 'message class is basis?
                        tVersion = UCase(tNode.GetAttribute("version"))
                        tActivePath = tXMLPath
                    Else
                        uCDebugPrint tLogTag, 0, "Загрузка БД " & inXMLDB.ClassTag & "(" & inXMLDB.Name & ") > Неожиданный класс!"
                    End If
                Else
                    uCDebugPrint tLogTag, 0, "Загрузка БД " & inXMLDB.ClassTag & "(" & inXMLDB.Name & ") > Корневой блок [message] не обнаружен!"
                End If
            Else
                uCDebugPrint tLogTag, 0, "Загрузка БД " & inXMLDB.ClassTag & "(" & inXMLDB.Name & ") > Ошибка парсинга!"
            End If
        End If
        'UPDATE
        If tActivePath = vbNullString Then
            uCDebugPrint tLogTag, 0, "Загрузка БД " & inXMLDB.ClassTag & "(" & inXMLDB.Name & ") неудачна! Файл БД отсутствует - создание нового!"
            fXMLDBCreator inXMLDB, inPath
            If inXMLDB.XML Is Nothing Then
                uCDebugPrint tLogTag, 0, "Не удалось создать!"
                inXMLDB.Active = False
                inXMLDB.Path = vbNullString
                Set inXMLDB.XML = Nothing
            Else
                uCDebugPrint tLogTag, 0, "Успешно создан!"
                inXMLDB.Path = tXMLPath
                inXMLDB.Active = True
            End If
        Else
            uCDebugPrint tLogTag, 0, "Загрузка БД " & inXMLDB.ClassTag & "(" & inXMLDB.Name & ") успешна!"
            Set inXMLDB.XML = CreateObject("Msxml2.DOMDocument.6.0")
            inXMLDB.XML.ASync = False
            inXMLDB.XML.Load (tActivePath)
            inXMLDB.Version = tVersion
            inXMLDB.Path = tActivePath
            inXMLDB.Active = True
        End If
    'End If
    fXMLDBLoader = inXMLDB.Active
End Function

Public Function fXMLSmartUpdate(Optional inLoadList As String = vbNullString)
    Dim tResult, tLoadItems, tTempList, tIndex
    Dim tLogTag
    
    ' 01 // Подготовка
    tLogTag = fGetLogTag("XMLSMARTUPD")
    
    ' 02 // Разбивка входящего списка
    tTempList = inLoadList
    If tTempList = vbNullString Then
        tTempList = "BASIS,CONVERTER,FRAME,CALENDAR,R80020DB,MAILSCAN,XSD80020V2"
        uCDebugPrint tLogTag, 1, "Не подан список тэгов! Сформирован список тэгов по умолчанию - " & tTempList
    End If
    uCDebugPrint tLogTag, 0, "Тэги загружаемых источников - " & tTempList
    tLoadItems = Split(tTempList, ",")
    tResult = True
    
    ' 03 // Загрзука требуемых по списку тэгов источников
    ' LIST: BASIS,CONVERTER,FRAME,CALENDAR,R80020DB,MAILSCAN,XSD80020V2,XSDBRFORECAST,DICTIONARY,BRFORECAST,R30308DB,CALCROUTE,CALCDB,XSD80040V2,F63DB,CREDENTIALS
    For tIndex = 0 To UBound(tLoadItems)
        Select Case UCase(tLoadItems(tIndex))
            Case gMainInit.ClassTag: tResult = tResult And fXMLConfigLoader(gMainInit, gConfigPath) 'INIT
            Case gXMLBasis.ClassTag: tResult = tResult And fXMLConfigLoader(gXMLBasis, gConfigPath) 'BASIS
            Case gXMLConverter.ClassTag: tResult = tResult And fXMLConfigLoader(gXMLConverter, gConfigPath) 'CONVERTER
            Case gXMLFrame.ClassTag: tResult = tResult And fXMLConfigLoader(gXMLFrame, gConfigPath) 'FRAME
            Case gXMLCalendar.ClassTag: tResult = tResult And fXMLConfigLoader(gXMLCalendar, gConfigPath) 'CALENDAR
            Case gXML80020DB.ClassTag: tResult = tResult And fXMLDBLoader(gXML80020DB, gConfigPath) 'R80020DB
            Case gMailScanDB.ClassTag: tResult = tResult And fXMLDBLoader(gMailScanDB, gConfigPath) 'MAILSCAN
            Case gXML80020CFG.XSD20V2.Tag: tResult = tResult And fXSDLoader(gXML80020CFG.XSD20V2, gConfigPath) 'XSD80020V2
            Case gXSDForecast.Tag: tResult = tResult And fXSDLoader(gXSDForecast, gConfigPath) 'XSDBRFORECAST
            Case gXMLDictionary.ClassTag: tResult = tResult And fXMLConfigLoader(gXMLDictionary, gConfigPath) 'DICTIONARY
            Case gBRForecastDB.ClassTag: tResult = tResult And fXMLDBLoader(gBRForecastDB, gConfigPath) 'BRFORECAST
            Case gXML30308DB.ClassTag: tResult = tResult And fXMLDBLoader(gXML30308DB, gConfigPath) 'R30308DB
            Case gCalcDB.ClassTag: tResult = tResult And fXMLDBLoader(gCalcDB, gConfigPath) 'CALCDB
            Case gXML80020CFG.XSD40V2.Tag: tResult = tResult And fXSDLoader(gXML80020CFG.XSD40V2, gConfigPath) 'XSD80040V2
            Case gF63DB.ClassTag: tResult = tResult And fXMLDBLoader(gF63DB, gConfigPath) 'F63DB
            Case gXMLCredentials.ClassTag: tResult = tResult And fXMLConfigLoader(gXMLCredentials, gConfigPath) 'CREDENTIALS
            Case gXMLCalcRoute.ClassTag: tResult = tResult And fXMLConfigLoader(gXMLCalcRoute, gConfigPath) 'CALCROUTE
            Case Else: uCDebugPrint tLogTag, 1, "Внимание! Тэг <" & UCase(tLoadItems(tIndex)) & "> не имеет доступного источника!"
        End Select
    Next

    ' 04 // Обработка результатов загрузки источников
    fXMLSmartUpdate = tResult
    If tResult Then
        uCDebugPrint tLogTag, 0, "Завершено успешно."
    Else
        uCDebugPrint tLogTag, 1, "Завершено с ошибками!"
    End If
End Function

Private Sub fXMLDBCreator(inXMLDB As TXMLDataBaseFile, inPath)
Dim tFilePath, tRoot, tAttr, tIntro, tComment, tTextFile, tText
    '00 // Path resolve
    tFilePath = inPath & "\" & inXMLDB.Name
    Set inXMLDB.XML = Nothing
    '01 // Delete old file if exists
    If uFileExists(tFilePath) Then
        If Not (uDeleteFile(tFilePath)) Then: Exit Sub
    End If
    '02 // Assign XML object to file path
    Set inXMLDB.XML = CreateObject("Microsoft.XMLDOM")
    inXMLDB.XML.ASync = False
    inXMLDB.XML.Load (tFilePath)
    
    '03 // Creating base structure
    Select Case inXMLDB.ClassTag
        Case "R80020DB":
            inXMLDB.Version = 1
            Set tRoot = inXMLDB.XML.CreateElement("message")
            inXMLDB.XML.AppendChild tRoot
            tRoot.SetAttribute "class", inXMLDB.ClassTag
            tRoot.SetAttribute "version", inXMLDB.Version
            tRoot.SetAttribute "releasestamp", fGetTimeStamp()
        Case "R30308DB":
            inXMLDB.Version = 1
            Set tRoot = inXMLDB.XML.CreateElement("message")
            inXMLDB.XML.AppendChild tRoot
            tRoot.SetAttribute "class", inXMLDB.ClassTag
            tRoot.SetAttribute "version", inXMLDB.Version
            tRoot.SetAttribute "releasestamp", fGetTimeStamp()
         Case "MAILSCAN":
            inXMLDB.Version = 1
            Set tRoot = inXMLDB.XML.CreateElement("message")
            inXMLDB.XML.AppendChild tRoot
            tRoot.SetAttribute "class", inXMLDB.ClassTag
            tRoot.SetAttribute "version", inXMLDB.Version
            tRoot.SetAttribute "releasestamp", fGetTimeStamp()
        Case "BRFORECAST":
            inXMLDB.Version = 1
            Set tRoot = inXMLDB.XML.CreateElement("message")
            inXMLDB.XML.AppendChild tRoot
            tRoot.SetAttribute "class", inXMLDB.ClassTag
            tRoot.SetAttribute "version", inXMLDB.Version
            tRoot.SetAttribute "releasestamp", fGetTimeStamp()
        Case "CALCDB":
            inXMLDB.Version = 1
            Set tRoot = inXMLDB.XML.CreateElement("message")
            inXMLDB.XML.AppendChild tRoot
            tRoot.SetAttribute "class", inXMLDB.ClassTag
            tRoot.SetAttribute "version", inXMLDB.Version
            tRoot.SetAttribute "releasestamp", fGetTimeStamp()
        Case "F63DB":
            inXMLDB.Version = 1
            Set tRoot = inXMLDB.XML.CreateElement("message")
            inXMLDB.XML.AppendChild tRoot
            tRoot.SetAttribute "class", inXMLDB.ClassTag
            tRoot.SetAttribute "version", inXMLDB.Version
            tRoot.SetAttribute "releasestamp", fGetTimeStamp()
    End Select
    
    '04 // Комментарий
    Set tComment = inXMLDB.XML.CreateComment("Сформировано " & Now() & " Outlook Interceptor Tool ")
    inXMLDB.XML.InsertBefore tComment, inXMLDB.XML.ChildNodes(0)
    
    '05 // Processing Instruction
    Set tIntro = inXMLDB.XML.CreateProcessingInstruction("xml", "version='1.0' encoding='Windows-1251' standalone='yes'")
    inXMLDB.XML.InsertBefore tIntro, inXMLDB.XML.ChildNodes(0)
    
    '06 // Save XML
    inXMLDB.XML.Save (tFilePath)
    
    '07 // Реорганизация XML для удобочитаемости в NotePad++
    Set tTextFile = gFSO.OpenTextFile(tFilePath, 1)
    tText = tTextFile.ReadAll
    tTextFile.Close
    Set tTextFile = gFSO.OpenTextFile(tFilePath, 2, True)
    tText = Replace(tText, "><", "> <")
    tTextFile.Write tText
    tTextFile.Close
    
    '08 // Сохранение изменений в XML
    inXMLDB.XML.Load (tFilePath)
    inXMLDB.XML.Save (tFilePath)
End Sub

Private Sub fXML80020CFGInit(inRootPath)
    With gXML80020CFG
        .Active = True
        'Data apply
        If IsEmpty(inRootPath) Then
            uDebugPrint "INI80020: inRootPath пуст > ERROR"
            Exit Sub
        End If
        .Path.Root = inRootPath & "\80020"
        .Path.Incoming = .Path.Root & "\Incoming"
        .Path.Processed = .Path.Root & "\Processed"
        .Path.Done = .Path.Root & "\Done"
        '.XSD.Tag = "XSD80020"
        '.XSD.Name = "80020.xsd"
        '.XSD.Path = vbNullString 'inXSDPath & "\" & .XSD.Name
        '.XSD.Active = False
        'Set .XSD.XML = Nothing
        'uDebugPrint "INI80020: .XSD.Path > " & .XSD.Path
        'Data check
        If Not (uFolderCreate(.Path.Root)) Then 'ROOT
            uDebugPrint "INI80020: .Path.Root > " & .Path.Root & " > ERROR"
            .Path.Root = vbNullString
            .Active = False
        Else
            uDebugPrint "INI80020: .Path.Root > " & .Path.Root & " > OK"
        End If
        If Not (uFolderCreate(.Path.Incoming)) Then 'INCOMING
            uDebugPrint "INI80020: .Path.Incoming > " & .Path.Incoming & " > ERROR"
            .Path.Incoming = vbNullString
            .Active = False
        Else
            uDebugPrint "INI80020: .Path.Incoming > " & .Path.Incoming & " > OK"
        End If
        If Not (uFolderCreate(.Path.Processed)) Then 'PROCESSED
            uDebugPrint "INI80020: .Path.Processed > " & .Path.Processed & " > ERROR"
            .Path.Processed = vbNullString
            .Active = False
        Else
            uDebugPrint "INI80020: .Path.Processed > " & .Path.Processed & " > OK"
        End If
        If Not (uFolderCreate(.Path.Done)) Then 'DONE
            uDebugPrint "INI80020: .Path.Done > " & .Path.Done & " > ERROR"
            .Path.Done = vbNullString
            .Active = False
        Else
            uDebugPrint "INI80020: .Path.Done > " & .Path.Done & " > OK"
        End If
        'fin
        uDebugPrint "INI80020: .Active > " & .Active
    End With
End Sub


Private Function fCheckTimeStamp(inValue) As Boolean
Dim tValue, tYear, tMonth, tDay
    'PREP
    fCheckTimeStamp = False
    'GET
    If Len(inValue) <> 14 Then: Exit Function
    'sec
    tValue = Right(inValue, 2)
    If Not IsNumeric(tValue) Then: Exit Function
    tValue = CInt(tValue)
    If tValue < 0 Or tValue > 59 Then: Exit Function
    'min
    tValue = Mid(inValue, 11, 2)
    If Not IsNumeric(tValue) Then: Exit Function
    tValue = CInt(tValue)
    If tValue < 0 Or tValue > 59 Then: Exit Function
    'hour
    tValue = Mid(inValue, 9, 2)
    If Not IsNumeric(tValue) Then: Exit Function
    tValue = CInt(tValue)
    If tValue < 0 Or tValue > 24 Then: Exit Function
    'day
    tValue = Mid(inValue, 7, 2)
    If Not IsNumeric(tValue) Then: Exit Function
    tValue = CInt(tValue)
    If tValue < 1 Or tValue > 31 Then: Exit Function
    tDay = tValue
    'month
    tValue = Mid(inValue, 5, 2)
    If Not IsNumeric(tValue) Then: Exit Function
    tValue = CInt(tValue)
    If tValue < 1 Or tValue > 12 Then: Exit Function
    tMonth = tValue
    'year
    tValue = Left(inValue, 4)
    If Not IsNumeric(tValue) Then: Exit Function
    tValue = CInt(tValue)
    If tValue < 2010 Or tValue > 2025 Then: Exit Function
    tYear = tValue
    'logic check
    If uDaysPerMonth(tMonth, tYear) < tDay Then: Exit Function
    'ove
    fCheckTimeStamp = True
End Function

Public Function IsTimeStamp(inText, inYear, inMonth, inDay) As Boolean
Dim tDay, tMonth, tYear
    IsTimeStamp = False
    inYear = 0
    inMonth = 0
    inDay = 0
    's1
    If Not (IsNumeric(inText)) Then: Exit Function
    If Len(inText) <> 6 And Len(inText) <> 8 Then: Exit Function
    's2
    tYear = CInt(Left(inText, 4))
    tMonth = CInt(Mid(inText, 5, 2))
    If Len(inText) = 8 Then: tDay = CInt(Right(inText, 2))
    's3
    If tYear < 2000 Or tYear > 2100 Then: Exit Function
    If tMonth < 1 Or tMonth > 12 Then: Exit Function
    If Len(inText) = 8 Then
        If tDay < 1 Or tDay > uDaysPerMonth(tMonth, tYear) Then: Exit Function
        inDay = Format(tDay, "00")
    End If
    inYear = Format(tYear, "0000")
    inMonth = Format(tMonth, "00")
    IsTimeStamp = True
End Function

Public Function IsTimeStampEqual(inText, inDate) As Boolean
    IsTimeStampEqual = False
    's1 chk
    If Not (IsNumeric(inText)) Then: Exit Function
    If Len(inText) <> 6 And Len(inText) <> 8 Then: Exit Function
    If Not (IsDate(inDate)) Then: Exit Function
    's2 cmp
    If CInt(Left(inText, 4)) <> Year(inDate) Then: Exit Function
    If CInt(Mid(inText, 5, 2)) <> Month(inDate) Then: Exit Function
    If Len(inText) = 8 Then
        If CInt(Right(inText, 2)) <> Day(inDate) Then: Exit Function
    End If
    'end
    IsTimeStampEqual = True
End Function

Public Function fGetTimeStamp()
Dim tNow
    tNow = Now() '20171017000000
    fGetTimeStamp = Format(Year(tNow), "0000") & Format(Month(tNow), "00") & Format(Day(tNow), "00") & Format(Hour(tNow), "00") & Format(Minute(tNow), "00") & Format(Second(tNow), "00")
    'Format(, "YYYYMMDDHsMsNs")
    'fin
End Function

Public Sub fExcelControl(Optional inScreen As Integer = 0, Optional inAlerts As Integer = 0, Optional inCalculation As Integer = 0, Optional inEvents As Integer = 0)
    If Not gExcel Is Nothing Then: Exit Sub
    With gExcel
    '=Screen
        If inScreen = 1 Then
            .ScreenUpdating = True
        ElseIf inScreen = -1 Then
            .ScreenUpdating = False
        End If
    '=Alerts
        If inAlerts = 1 Then
            .DisplayAlerts = True
        ElseIf inAlerts = -1 Then
            .DisplayAlerts = False
        End If
    '=Calculation
        If inCalculation = 1 Then
            .Calculation = -4105 'xlCalculationAutomatic
        ElseIf inCalculation = -1 Then
            .Calculation = -4135 'xlCalculationManual
        End If
    '=Events
        If inEvents = 1 Then
            .EnableEvents = True
        ElseIf inEvents = -1 Then
            .EnableEvents = False
        End If
    End With
End Sub

'---- DATA EXTRACTORS ----

' DICTIONARY // Subject Info
Public Function fDictionaryGetSubjectInfo(inSubjectID, outSubjectInfo As TSubjectInfo)
Dim tInternalSubjectID, tXPathString, tNode, tValue, tTempValue, tHasParentSubject, tInternalParentSubjectID
    fDictionaryGetSubjectInfo = False
    tHasParentSubject = False
    tInternalParentSubjectID = vbNullString
' object prepare
    With outSubjectInfo
        .IsReady = False
        .Comment = vbNullString
        .ID = 0
        .ParentID = 0
        .Name = "<NO DATA>"
        .ParentName = "<NO DATA>"
        .LocalUTC = 0
        .TradeZoneUTC = 0
        .TimeZoneID = 0
        .TradeZoneID = 0
    ' test 0
        If inSubjectID = vbNullString Then
            .Comment = "Переменная inSubjectID пуста!"
            Exit Function
        End If
    ' test 1
        If Not IsNumeric(inSubjectID) Then
            .Comment = "Переменная inSubjectID должна быть числом!"
            Exit Function
        End If
        tInternalSubjectID = CLng(inSubjectID)
    ' test 2
        If tInternalSubjectID <= 0 Then
            .Comment = "Переменная inSubjectID должна быть больше нуля!"
            Exit Function
        End If
        'реконфигурация
        If tInternalSubjectID >= 100 Then
            tInternalSubjectID = Format(tInternalSubjectID, "00000")
            tInternalParentSubjectID = Left(tInternalSubjectID, 2)
            tHasParentSubject = True
        Else
            tInternalSubjectID = Format(tInternalSubjectID, "00")
        End If
    'test 3
        If Not gXMLDictionary.Active Then
            .Comment = "Конфиг DICTIONARY не загружен!"
            Exit Function
        End If
    'test 4
        If gXMLDictionary.XML Is Nothing Then
            .Comment = "Конфиг DICTIONARY не привязан к XML!"
            Exit Function
        End If
    'extracting data
        'найдем родительский субъект
        If tHasParentSubject Then
            'parent subject id
            tTempValue = tInternalParentSubjectID
            tXPathString = "//subjects/subject[@id='" & tInternalSubjectID & "']"
            Set tNode = gXMLDictionary.XML.SelectSingleNode(tXPathString)
            If tNode Is Nothing Then
                .Comment = "В DICTIONARY не найден родительский субъект с кодом <" & tTempValue & ">!"
                Exit Function
            End If
            .ParentID = tInternalParentSubjectID
            'parent subject name
            tValue = tNode.GetAttribute("name")
            If IsNull(tValue) Then
                .Comment = "В DICTIONARY субъект с кодом <" & tTempValue & "> не содержит аттрибута <name>!"
                Exit Function
            End If
            .ParentName = tValue
        End If
        'текущий субъект
    'subject id
        tTempValue = tInternalSubjectID
        tXPathString = "//subjects/subject[@id='" & tInternalSubjectID & "']"
        Set tNode = gXMLDictionary.XML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            .Comment = "В DICTIONARY не найден субъект с кодом <" & tTempValue & ">!"
            Exit Function
        End If
        .ID = tInternalSubjectID
    'subject name
        tValue = tNode.GetAttribute("name")
        If IsNull(tValue) Then
            .Comment = "В DICTIONARY субъект с кодом <" & tTempValue & "> не содержит аттрибута <name>!"
            Exit Function
        End If
        .Name = tValue
    'parent resolver
        If Not tHasParentSubject Then
            .ParentID = .ID
            .ParentName = .Name
        End If
    'subject local time in utc
        tValue = tNode.GetAttribute("utc")
        If IsNull(tValue) Then
            .Comment = "В DICTIONARY субъект с кодом <" & tTempValue & "> не содержит аттрибута <utc>!"
            Exit Function
        End If
        If tValue = vbNullString Then: tValue = 0
        If Not IsNumeric(tValue) Then
            .Comment = "В DICTIONARY субъект с кодом <" & tTempValue & "> содержит нечисловое значение [" & tValue & "] аттрибута <utc>!"
            Exit Function
        End If
        If Abs(tValue) > 12 Then
            .Comment = "В DICTIONARY субъект с кодом <" & tTempValue & "> содержит неверное значение [" & tValue & "] аттрибута <utc>!"
            Exit Function
        End If
        .LocalUTC = tValue
    'subject trade zone
        tValue = tNode.GetAttribute("tradezone")
        If IsNull(tValue) Then
            .Comment = "В DICTIONARY субъект с кодом <" & tTempValue & "> не содержит аттрибута <tradezone>!"
            Exit Function
        End If
        If Not IsNumeric(tValue) Then
            .Comment = "В DICTIONARY субъект с кодом <" & tTempValue & "> содержит нечисловое значение [" & tValue & "] аттрибута <tradezone>!"
            Exit Function
        End If
        tValue = Format(tValue, "00")
        tXPathString = "//tradezones/tradezone[@id='" & tValue & "']"
        Set tNode = gXMLDictionary.XML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            .Comment = "В DICTIONARY не найдена торговая зона с кодом <" & tValue & ">!"
            Exit Function
        End If
        .TradeZoneID = tValue
        tTempValue = tValue
    'trade zone trademode
        tValue = tNode.GetAttribute("trademode")
        If IsNull(tValue) Then
            .Comment = "В DICTIONARY торговая зона с кодом <" & tTempValue & "> не содержит аттрибута <trademode>!"
            Exit Function
        End If
        If Not IsNumeric(tValue) Then
            .Comment = "В DICTIONARY торговая зона с кодом <" & tTempValue & "> содержит нечисловое значение [" & tValue & "] аттрибута <trademode>!"
            Exit Function
        End If
        tValue = CLng(tValue)
        If Not (tValue = 0 Or tValue = 1) Then
            .Comment = "В DICTIONARY торговая зона с кодом <" & tTempValue & "> содержит неверное значение [" & tValue & "] аттрибута <trademode>!"
            Exit Function
        End If
        .TradeMode = tValue
    'trade zone timezone
        tValue = tNode.GetAttribute("timezone")
        
        If IsNull(tValue) Then
            .Comment = "В DICTIONARY торговая зона с кодом <" & tTempValue & "> не содержит аттрибута <timezone>!"
            Exit Function
        End If
        If Not IsNumeric(tValue) Then
            .Comment = "В DICTIONARY торговая зона с кодом <" & tTempValue & "> содержит нечисловое значение [" & tValue & "] аттрибута <timezone>!"
            Exit Function
        End If
        tValue = CLng(tValue)
        tXPathString = "//timezones/timezone[@id='" & tValue & "']"
        Set tNode = gXMLDictionary.XML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            .Comment = "В DICTIONARY не найдена временная зона с кодом <" & tValue & ">!"
            Exit Function
        End If
        .TimeZoneID = tValue
        tTempValue = tValue
    'timezone utc
        tValue = tNode.GetAttribute("utc")
        If IsNull(tValue) Then
            .Comment = "В DICTIONARY временная зона с кодом <" & tTempValue & "> не содержит аттрибута <utc>!"
            Exit Function
        End If
        If tValue = vbNullString Then: tValue = 0
        If Not IsNumeric(tValue) Then
            .Comment = "В DICTIONARY временная зона с кодом <" & tTempValue & "> содержит нечисловое значение [" & tValue & "] аттрибута <utc>!"
            Exit Function
        End If
        If Abs(tValue) > 12 Then
            .Comment = "В DICTIONARY временная зона с кодом <" & tTempValue & "> содержит неверное значение [" & tValue & "] аттрибута <utc>!"
            Exit Function
        End If
        .TradeZoneUTC = tValue
    'over
        .IsReady = True
    End With
    fDictionaryGetSubjectInfo = True
End Function

' BASIS // GTP Settings
Public Function fBasisGetGTPSettings(inGTPID, inParamName, outValue, outComment, Optional inTraderID = vbNullString)
Dim tXPathString, tInternalGTPID, tRootNode, tInternalParamName, tNode, tSubNodeName, tValue
    fBasisGetGTPSettings = False
    outValue = 0
    outComment = vbNullString
    If inTraderID = vbNullString Then: inTraderID = gTraderInfo.ID
    'test 0
    If inGTPID = vbNullString Then
        outComment = "Переменная inGTPID пуста!"
        Exit Function
    End If
    tInternalGTPID = UCase(inGTPID)
    'test 1
    If inParamName = vbNullString Then
        outComment = "Переменная inParamName пуста!"
        Exit Function
    End If
    tInternalParamName = LCase(inParamName)
    'test 2
    If Not gXMLBasis.Active Then
        outComment = "Конфиг BASIS не загружен!"
        Exit Function
    End If
    'test 3
    If gXMLBasis.XML Is Nothing Then
        outComment = "Конфиг BASIS не привязан к XML!"
        Exit Function
    End If
    'lock the gtp node
    tXPathString = "//trader[@id='" & inTraderID & "']/gtp[@id='" & tInternalGTPID & "']" 'BY GTP ID
    Set tRootNode = gXMLBasis.XML.SelectSingleNode(tXPathString)
    If tRootNode Is Nothing Then
        tXPathString = "//trader[@id='" & inTraderID & "']/gtp[@aiiscode='" & tInternalGTPID & "']" 'BY AIIS CODE
        Set tRootNode = gXMLBasis.XML.SelectSingleNode(tXPathString)
        If tRootNode Is Nothing Then
            outComment = "В BASIS не найдена ГТП <" & inTraderID & "/" & tInternalGTPID & ">!"
            Exit Function
        End If
    End If
    'lock sub node
    tSubNodeName = "settings"
    Set tNode = fGetChildNodeByName(tRootNode, tSubNodeName)
    If tNode Is Nothing Then
        outComment = "В конфиге BASIS ГТП <" & inTraderID & "/" & tInternalGTPID & "> не имеет дочернего объекта <" & tSubNodeName & ">!"
        Exit Function
    End If
    'get attr
    tValue = tNode.GetAttribute(tInternalParamName)
    If IsNull(tValue) Then
        outComment = "В конфиге BASIS ГТП <" & inTraderID & "/" & tInternalGTPID & "> не имеет аттрибута <" & tInternalParamName & "> дочернего объекта <" & tSubNodeName & ">!"
        Exit Function
    End If
    'got it
    outValue = tValue
    fBasisGetGTPSettings = True
End Function

Public Function fGetUTCForm(inNumber)
    fGetUTCForm = "UTC??"
    If IsEmpty(inNumber) Then: Exit Function
    If IsNull(inNumber) Then: Exit Function
    If Not IsNumeric(inNumber) Then: Exit Function
    If Abs(inNumber) > 12 Then: Exit Function
    If inNumber >= 0 Then
        fGetUTCForm = "UTC+" & inNumber
    Else
        fGetUTCForm = "UTC" & inNumber
    End If
End Function

' U-01 Чтение аттрибута из конфига. Возвращает найденную ноду в outNode; outAttributeValue - значение аттрибута.
Public Function fGetAttributeCFG(inXMLCFG As TXMLConfigFile, inXPathString, inAttributeName, outAttributeValue, outNode, outErrorText, Optional inTargetNode = Nothing)
Dim tRootNode, tTempNode, tTempValue
    fGetAttributeCFG = False
    outErrorText = vbNullString
    Set outNode = Nothing
' 01 // Определим конервую ноду для XPath
    If inTargetNode Is Nothing Then
        Set tRootNode = inXMLCFG.XML
    Else
        Set tRootNode = inTargetNode
    End If
' 02 // Проверка состояния корневой ноды
    If tRootNode Is Nothing Then
        outErrorText = "fGetAttributeCFG > Корневая нода не инициирована! Искомый аттрибут [" & inAttributeName & "], XPath [" & inXPathString & "]"
        Exit Function
    End If
' 03 // Поиск ноды по XPath
    If inXPathString <> vbNullString Then
        Set tTempNode = tRootNode.SelectSingleNode(inXPathString)
    Else
        Set tTempNode = tRootNode 'при отсутствии XPath - искать аттрибут в родительской ноде
    End If
    
    If tTempNode Is Nothing Then
        outErrorText = "fGetAttributeCFG > В конфиге " & inXMLCFG.ClassTag & " не удалось ноду XPath [" & inXPathString & "]!"
        Exit Function
    End If
' 04 // Чтение аттрибута
    tTempValue = tTempNode.GetAttribute(inAttributeName) 'проичтаем аттрибут
    If IsNull(tTempValue) Then
        outErrorText = "fGetAttributeCFG > Не удалось получить аттрибут <@" & inAttributeName & "> ноды XPath [" & inXPathString & "] конфига " & inXMLCFG.ClassTag & "!"
        Exit Function
    End If
' 05 // Завершение
    outAttributeValue = tTempValue
    Set outNode = tTempNode
    Set tRootNode = Nothing
    Set tTempNode = Nothing
    fGetAttributeCFG = True
End Function
