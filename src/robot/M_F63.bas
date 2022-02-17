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
Private Const cnstSendDelay = 35 'кажется 30 секунд там между подачами

Private gLocalInit

Private Function fGetLogTag(inTagText)
    fGetLogTag = cnstModuleName & "." & inTagText
End Function

Private Function fLocalInit(Optional inForceInit As Boolean = False)
Dim tLogTag
    fLocalInit = False
    tLogTag = fGetLogTag("F63INI")
    uCDebugPrint tLogTag, 0, "Инициализация > Имя: " & cnstModuleName & "; Версия: " & cnstModuleVersion & "; Редакция: " & cnstModuleDate
    If inForceInit Or Not gLocalInit Then
        ' // variables
        gLocalInit = False
        ' // objects
        ' // paths
        'gForecastDataPath = gDataPath & "\BRForecast"
        'If Not (gFSO.FolderExists(gForecastDataPath)) Then
        '    If Not (uFolderCreate(gForecastDataPath)) Then
        '        uADebugPrint tLogTag, "Не удалось найти папку данных <gForecastDataPath> по пути: " & gForecastDataPath
        '        gDataPath = vbNullString
        '        Exit Function
        '    Else
        '        uADebugPrint tLogTag, "Создана папка данных <gForecastDataPath> по пути: " & gForecastDataPath
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

' F-01 // Базовая рабочая функция по перебору нод BASIS
Private Sub fF63ExchangeScan(inLocalUTC, inTraderID, Optional inGTPCode = vbNullString, Optional inForceIgnoreGate = False, Optional inForceIgnoreTimeLimits = False, Optional inForceUncommIgnore = False, Optional inForceDayShift = 0, Optional inForceDaysToReport = 0)
Dim tXPathString, tNode, tNodes, tLogTag
' 00 // Предопределения и первичная проверка
    tLogTag = fGetLogTag("F63EXSCN")
    If Not gXMLBasis.Active Then: Exit Sub
    If Not gXMLCalendar.Active Then: Exit Sub
    If Not gF63DB.Active Then: Exit Sub
' 01 // Поиск нод зайдествованных в работе с F63 в BASIS
    If inGTPCode = vbNullString Then
        tXPathString = "//trader[@id='" & inTraderID & "']/gtp/exchange/item[(@id='F63' and @enabled='1')]"
    Else
        tXPathString = "//trader[@id='" & inTraderID & "']/gtp[@id='" & inGTPCode & "']/exchange/item[(@id='F63' and @enabled='1')]"
    End If
    
    Set tNodes = gXMLBasis.XML.SelectNodes(tXPathString)
    
    If inGTPCode <> vbNullString Then: uCDebugPrint tLogTag, 1, "Всего нод найдено - " & tNodes.Length
    
    If tNodes.Length = 0 Then: Exit Sub 'нет элементов для работы

' 02 // Обработка найденных нод
    For Each tNode In tNodes
        fF63ExchangeItem tNode, inLocalUTC, inForceIgnoreGate, inForceIgnoreTimeLimits, Fix(inForceDayShift), Fix(inForceDaysToReport), inForceUncommIgnore
    Next
End Sub

' F-02 // Базовая функция обработки найденных в BASIS нод
Private Sub fF63ExchangeItem(inNode, inLocalUTC, inForceIgnoreGate, inForceIgnoreTimeLimits, inForceDayShift, inForceDaysToReport, inForceUncommIgnore) ', inTraderID)
Dim tLogTag, tGTPID, tTimeZoneUTC, tErrorText, tActiveDays, tIndex, tCurrentDay, tFirstDay, tCurrentDateTime, tCalcValue, tSectionList, tIsUpdated, tCalcReady, tCreateNewNode, tXMLChanged
Dim tCurrentReportNode, tTempValue, tNeedToSend, tLogin, tPassword, tGateIgnore, tUncommIgnore, tTraderID
' 00 // Предопределения
    tLogTag = fGetLogTag("fF63ExchangeItem")
    
' 01 // Чтение параметров входящей ноды
    If Not fReadBasisData(inNode, tTraderID, tGTPID, tTimeZoneUTC, tLogin, tPassword, tGateIgnore, tUncommIgnore, tErrorText) Then
        uCDebugPrint tLogTag, 2, "fReadBasisData > " & tErrorText
        Exit Sub
    End If
    
    If inForceIgnoreGate Then
        uCDebugPrint tLogTag, 1, "Игнорирование ворот приема ПАК ЭНЕРГИИ включено аварийно! inForceIgnoreGate = True"
        tGateIgnore = True
    End If
    
    If inForceIgnoreTimeLimits Then
        uCDebugPrint tLogTag, 1, "Игнорирование ворот подачи УЧАСТНИКОМ включено аварийно! inForceIgnoreTimeLimits = True"
    End If
    
    If inForceUncommIgnore Then
        uCDebugPrint tLogTag, 1, "Игнорирование некоммерческой информации при расчёте потребления включено аварийно! inForceUncommIgnore = True"
    End If
    
' 02 // Определение временных рамок действия
    If Not fIsActivePeriod(inNode, inForceIgnoreTimeLimits, tTimeZoneUTC, inLocalUTC, tCurrentDateTime, tErrorText) Then
        If tErrorText <> vbNullString Then: uCDebugPrint tLogTag, 2, tErrorText 'Тихий выход, если вне периода время находится
        'uCDebugPrint tLogTag, 0, "Out Period"
        Exit Sub
    End If
    uCDebugPrint tLogTag, 0, "Подача данных по ГТП " & tGTPID
    
' 03 // Определение дней за которые можно подать данные (по календарю)
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
        
        uCDebugPrint tLogTag, 1, "Производится аварийное смещение на [" & Fix(inForceDayShift) & "] суток для начала периода."
        uCDebugPrint tLogTag, 1, "Дата начала [" & Format(tFirstDay, "YYYY-MM-DD") & "]; Текущая дата [" & Format(Fix(tCurrentDateTime), "YYYY-MM-DD") & "]; Количество дней для подачи [" & tActiveDays & "]"
    Else
        uCDebugPrint tLogTag, 2, "inForceDayShift (" & inForceDayShift & ") не должен быть больше за пределами целых значений [-5:+5]!"
        Exit Sub
    End If
        
    tCurrentDay = tFirstDay
    'uCDebugPrint tLogTag, 0, "tFirstDay=" & Format(tFirstDay, "YYYYMMDD")
    
' 04 // Определим список сечений для расчёта
    If Not fGetSectionList(inNode, tSectionList, cnstElementSplitter, tErrorText) Then
        uCDebugPrint tLogTag, 2, "Ошибка получения списка сечений по ГТП " & tGTPID & ": " & tErrorText
        Exit Sub
    End If
    uCDebugPrint tLogTag, 0, "Параметры подачи данных: tActiveDays=" & tActiveDays & "; tSectionList=" & tSectionList & "; tGateIgnore=" & tGateIgnore & "; tUncommIgnore=" & tUncommIgnore
    
' 05 // Для каждого дня в списке необходимо проверить статус расчёта и отправки данных
    For tIndex = 1 To tActiveDays
' 06 // Получим статус расчёта
        If Not fCalculateDayValue(tTraderID, tGTPID, tSectionList, cnstElementSplitter, tCurrentDay, tUncommIgnore, tCalcValue, tIsUpdated, tErrorText) Then
            uCDebugPrint tLogTag, 2, tErrorText
            tCalcReady = False 'расчёт неудачен
        Else
            'tCalcValue = Round(tCalcValue, 0) 'v1
            tCalcValue = -tCalcValue 'v2 Потребление это отрицательная величина в понимании алгоритма относительно общего баланс системы (минус - забрал; плюс - отдал)
            uCDebugPrint tLogTag, 0, "CALCVAL=" & tCalcValue & "; ISUPDATED=" & tIsUpdated
            tCalcReady = True 'расчёт удачен
        End If
' 07 // извлечем из БД последний зарегистрированный отчет по этой ГТП на этот день
        If Not fGetCurrentReportF63DB(gF63DB, tTraderID, tGTPID, tCurrentDay, tCurrentReportNode, tErrorText) Then
            uCDebugPrint tLogTag, 2, tErrorText
            Exit Sub
        End If
' 08 // действия с текущим отчетом (если объём не читается или не равен расчётному, то создаём новый report)
        tCreateNewNode = True 'tCalcReady 'если расчёт удачен - то предварительно мы готовы создать новый отчет, если неудачен - создание отчета не возможно
        tXMLChanged = False
        If Not tCurrentReportNode Is Nothing Then
            tTempValue = tCurrentReportNode.GetAttribute("calcstatus")
            If Not IsNull(tTempValue) Then
                '4 возможных ситуации: 1. X -> X; 2. X -> V; 3. V -> V; 4. V -> X
                'Вариант 1 (XX) не должна инициировать создание новой записи, во всех остальных случаях это необходимо (сомнителен разве что 4й варивант)
                If tTempValue = "0" And Not tCalcReady Then
                    tCreateNewNode = False
                'Вариант 3 (VV) обновление данных - следует посмотреть не изменился ли расчёт
                ElseIf tTempValue = "1" And tCalcReady Then
                    tTempValue = tCurrentReportNode.GetAttribute("value")
                    If Not IsNull(tTempValue) Then
                        If IsNumeric(tTempValue) Then
                            tTempValue = CDec(tTempValue) - tCalcValue
                            If tTempValue = 0 Then: tCreateNewNode = False
                        End If
                    End If
                End If
                'Вариант 2 и 4 пропускаются как безусловные к созданию новой записи
            End If
        End If
' 09 // новая нода report
        If tCreateNewNode Then
            If Not fCreateNewReportF63DB(gF63DB, tTraderID, tGTPID, tCurrentDay, tCalcValue, tCalcReady, tCurrentReportNode, tErrorText) Then
                fReloadXMLDB gF63DB, False 'RollBack ANY Changes
                uCDebugPrint tLogTag, 2, tErrorText
                Exit Sub
            End If
            tXMLChanged = True
        End If
' 10 // проверка состояния текущего отчета, для принятия решения о необходимости отправки
        If Not tCurrentReportNode Is Nothing Then
            tNeedToSend = False
            tTempValue = tCurrentReportNode.GetAttribute("sent")
            If IsNull(tTempValue) Then
                tNeedToSend = True
            ElseIf tTempValue = vbNullString Then
                tNeedToSend = True
            End If
            If tNeedToSend Then 'если калькуляци была с ошибками - остановить отправку
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
        'сохранение данных в БД
        If tXMLChanged Then: fSaveXMLDB gF63DB, False
        'к следующему дню
        tCurrentDay = tCurrentDay + 1
    Next
    'uCDebugPrint tLogTag, 0, "tWorkDays=" & tWorkDays
    'uCDebugPrint tLogTag, 2, tGTPID & "  " & tTimeZoneUTC

End Sub

Private Function fReportSending(inReportNode, inLogin, inPassword, inGTPID, inDate, inXMLChanged, inGateIgnore, outErrorText)
Dim tEnergyAPI As New CEnergyAPI
Dim tNode, tCalcValue, tSentTry, tLogTag, tIsSendFailed
' 00 // Предопределения
    fReportSending = False
    outErrorText = vbNullString
    tLogTag = fGetLogTag("F63REPORTSEND")
' 01 // Извлечение данных из ноды report
    tCalcValue = inReportNode.GetAttribute("value")
    tSentTry = inReportNode.GetAttribute("senttrycount")
    If IsNumeric(tSentTry) Then
        tSentTry = CDec(tSentTry) + 1
    Else
        tSentTry = 1
    End If
    uCDebugPrint tLogTag, 1, "Попытка #" & tSentTry & " отправки данных по ГТП " & inGTPID & " на дату " & Format(inDate, "DD.MM.YYYY")
    inXMLChanged = True
    inReportNode.SetAttribute "senttrycount", tSentTry
    inReportNode.SetAttribute "senttry", Format(Now(), "YYYYMMDD hh:mm:ss")
' 02 // Отправка данных через API Energy2010
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
' 03 // По результатам работы API вынести решение
    If Not tIsSendFailed Then
        inReportNode.SetAttribute "sent", Format(Now(), "YYYYMMDD hh:mm:ss")
        inReportNode.SetAttribute "errortext", vbNullString
        uCDebugPrint tLogTag, 0, "Данные отправлены!"
    Else
        inReportNode.SetAttribute "errortext", tEnergyAPI.ErrorText
        uCDebugPrint tLogTag, 1, "Данные не отправлены!"
    End If
' XX // Завершение
    Set tEnergyAPI = Nothing
    fReportSending = True
End Function

Private Function fCreateNewReportF63DB(inXMLDB As TXMLDataBaseFile, inTraderID, inGTPID, inDate, inCalcValue, inCalcStatus, outReportNode, outErrorText)
Dim tXPathString, tRootNode, tDayNode, tNode, tLastReportIndex, tYear, tMonth, tDay
' 00 // Предопределения
    fCreateNewReportF63DB = False
    Set outReportNode = Nothing
    outErrorText = vbNullString
' 00 // Проверка версии и наличия XML
    If inXMLDB.XML Is Nothing Then
        outErrorText = "БД <" & inXMLDB.ClassTag & "> не готова!"
        Exit Function
    End If
    'версия так же определяет первичный запрос
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
            'Чтение индекса последнего отчета
            tLastReportIndex = tDayNode.GetAttribute("lastreport")
            If Not IsNull(tLastReportIndex) Then
                If Not IsNumeric(tLastReportIndex) Then
                    outErrorText = "Не удалось прочитать аттрибут <lastreport>, он оказался нечисловым <" & tLastReportIndex & ">!"
                    Exit Function
                End If
            Else
                tLastReportIndex = 0
            End If
            'Создание нового отчета
            tLastReportIndex = tLastReportIndex + 1 'индекс назначаем следующим
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
            'Передача готовой ноды в результат
            Set outReportNode = tNode
        Case Else:
            outErrorText = "Версия <" & inXMLDB.Version & "> БД <" & inXMLDB.ClassTag & "> не имеет обработчика! Аномалия!"
            Exit Function
    End Select
' XX // Завершение
    fCreateNewReportF63DB = True
End Function

Private Function fGetCurrentReportF63DB(inXMLDB As TXMLDataBaseFile, inTraderID, inGTPID, inDate, outReportNode, outErrorText)
Dim tXPathString, tRootNode, tDayNode, tLastReportIndex
' 00 // Предопределения
    fGetCurrentReportF63DB = False
    Set outReportNode = Nothing
    outErrorText = vbNullString
' 00 // Проверка версии и наличия XML
    If inXMLDB.XML Is Nothing Then
        outErrorText = "БД <" & inXMLDB.ClassTag & "> не готова!"
        Exit Function
    End If
    'версия так же определяет первичный запрос
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
                            outErrorText = "Обнаружены проблемы в БД <" & inXMLDB.ClassTag & ">! Нода <report> с индексом <" & tLastReportIndex & "> не найдена! XPath <" & tXPathString & ">"
                            Exit Function
                        End If
                    Else
                        outErrorText = "Обнаружены проблемы в БД <" & inXMLDB.ClassTag & ">! Индекс <" & tLastReportIndex & "> для ноды <report> указан не числом! XPath <" & tXPathString & ">"
                        Exit Function
                    End If
                End If
            End If
        Case Else:
            outErrorText = "Версия <" & inXMLDB.Version & "> БД <" & inXMLDB.ClassTag & "> не имеет обработчика! Аномалия!"
            Exit Function
    End Select
' XX // Завершение
    fGetCurrentReportF63DB = True
End Function

'Получение списка сечений с активной(боевой) версией из любой ноды входящей в состав ноды ГТП
Private Function fGetSectionList(inNode, outSectionList, inSplitter, outErrorText)
Dim tXPathString, tNode, tSectionID, tNodes, tSectionListCount, tIndex
Dim tSectionList()
' 00 // Предопределения
    fGetSectionList = False
    outErrorText = vbNullString
    outSectionList = vbNullString
' 01 // Переход на родительскую ноду ГТП
    tXPathString = "ancestor::gtp"
    Set tNode = inNode.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        outErrorText = "Не удалось получить родительскую ноду по XPath <" & tXPathString & ">!"
        Exit Function
    End If
' 02 // Поиск дочерних нод АКТИВНЫХ сечений
    tXPathString = "descendant::section[version[@status='active']]"
    Set tNodes = tNode.SelectNodes(tXPathString)
    If tNodes.Length = 0 Then
        outErrorText = "Не удалось получить ноды активных версий сечений по XPath <" & tXPathString & ">!"
        Exit Function
    End If
' 03 // Сбор сечений и проверка на уникальность версий для каждого
    tSectionListCount = -1
    For Each tNode In tNodes
        tSectionID = tNode.GetAttribute("id")
        If IsNull(tSectionID) Then
            outErrorText = "Не удалось получить аттрибут <id> ноды сечения!"
            Exit Function
        End If
        'Поиск SetctionID в списке
        For tIndex = 0 To tSectionListCount
            If tSectionList(tIndex) = tSectionID Then
                outErrorText = "Обнаружено несколько активных версий сечения <" & tSectionID & ">! Аномалия!"
                Exit Function
            End If
        Next
        'Внесение в список
        tSectionListCount = tSectionListCount + 1
        ReDim Preserve tSectionList(tSectionListCount)
        tSectionList(tSectionListCount) = tSectionID
        'подготовка результата функции
        If outSectionList = vbNullString Then
            outSectionList = tSectionID
        Else
            outSectionList = outSectionList & cnstElementSplitter & tSectionID
        End If
    Next
' 04 // Проверка состояния ответа
    If outSectionList = vbNullString Or tSectionListCount = -1 Then
        outErrorText = "При чтении активных версий сечений возникли проблемы, не удалось прочитать список!"
        Exit Function
    End If
' XX // Завершение
    fGetSectionList = True
End Function

Private Function fGetActiveSectionVersion(inTraderID, inGTPID, inSection)
Dim tXPathString, tNodes, tTempValue, tLogTag
    tLogTag = fGetLogTag("GETACTIVESECTION")
    fGetActiveSectionVersion = 0 'version zero - mean error
    If Not gXMLBasis.Active Then: Exit Function
    
    tXPathString = "//trader[@id='" & inTraderID & "']/gtp[@id='" & inGTPID & "']/section[@id='" & inSection & "']/version[@status='active']"
    Set tNodes = gXMLBasis.XML.SelectNodes(tXPathString)
    If tNodes.Length <> 1 Then 'аномальное количество элементов активных сечений
        uCDebugPrint tLogTag, 2, "Количество версий активного сечения должно быть 1 [найдено - " & tNodes.Length & "]; tXPathString=[" & tXPathString & "]"
        Exit Function
    End If
    
    tTempValue = tNodes(0).GetAttribute("id")
    If IsNull(tTempValue) Then 'неудача чтения
        uCDebugPrint tLogTag, 2, "Ошибка ID версии активного сечения; tXPathString=[" & tXPathString & "]"
        Exit Function
    End If
    
    uCDebugPrint tLogTag, 0, "Версия активного сечечния определена как [" & tTempValue & "]; tXPathString=[" & tXPathString & "]"
    fGetActiveSectionVersion = tTempValue
End Function

Private Function fCalculateDayValue(inTraderID, inGTPID, inSectionList, inSplitter, inDate, inUncomIgnore, outValue, outIsUpdated, outErrorText)
Dim tLogTag, tIsUpdated, tSections, tSection, tResultDateStart, tResultDateEnd, tErrorText, tError, tStatusLine, tActiveSectionVersionID
Dim tResult()
' 00 // Предопределения
    tLogTag = fGetLogTag("CALCDAYVALUE")
    fCalculateDayValue = False
    outErrorText = vbNullString
    outValue = 0
    outIsUpdated = False
' 01 // Чтение данных
    tSections = Split(inSectionList, inSplitter)
    For Each tSection In tSections
        'Внешний запрос к модулю CALCROUTE
        tActiveSectionVersionID = fGetActiveSectionVersion(inTraderID, inGTPID, tSection)
        'tError = fGetFactCalculation(inTraderID, inGTPID, tSection, tActiveSectionVersionID, "FULL", 0, "T", 0, "d", inDate, inDate, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText)
        tError = fGetFactCalculation(inTraderID, inGTPID, tSection, tActiveSectionVersionID, "FULL", 0, True, "T", 1, "d", inDate, inDate, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText, inUncomIgnore)
        If tError <> 0 Then
            outErrorText = tErrorText
            uCDebugPrint tLogTag, 1, tErrorText
            Exit Function
        End If
        'Проверка tStatusLine > содержит 2 цифры, в которых обозначен результат обработки макетов 80020 и 80040, где 1 - макеты есть и всё посчитано, 2 - какая-то ошибка расчёта этого типа макетов
        'If Len(tStatusLine) <> 2 Then 'v1
        If Len(tStatusLine) <> 3 Then 'v2
            tErrorText = "#E1# Данные недостоверны, проверьте всё! [" & tStatusLine & "]"
            outErrorText = tErrorText
            uCDebugPrint tLogTag, 1, tErrorText
            Exit Function
        End If
        'If Left(tStatusLine, 1) <> "1" Then 'v1
        If Left(tStatusLine, 1) <> "0" Then 'v2 problem mark with "1"
            tErrorText = "#E2# Данные недостоверны, проверьте всё! [" & tStatusLine & "]"
            outErrorText = tErrorText
            uCDebugPrint tLogTag, 1, tErrorText
            Exit Function
        End If
        If Not IsNumeric(tResult(0)) Then 'v2
            tErrorText = "#E3# Рассчёт неудался! RESULT=" & tResult(0)
            outErrorText = tErrorText
            uCDebugPrint tLogTag, 1, tErrorText
            Exit Function
        End If
        outValue = outValue + tResult(0)
        outIsUpdated = outIsUpdated Or tIsUpdated
    Next
' XX // Завершение
    fCalculateDayValue = True
End Function

Private Function fIsActivePeriod(inNode, inForceIgnoreTimeLimits, inTimeZoneUTC, inLocalUTC, outCurrentDateTime, outErrorText)
Dim tCurrentTime, tStartTime, tEndTime, tShiftTime
' 00 // Предопределения
    fIsActivePeriod = False
    outErrorText = vbNullString
    outCurrentDateTime = 0
    
' 01 // Проверка часового пояса
    'Проверка на значение
    If Not IsNumeric(inTimeZoneUTC) Then
        outErrorText = "Параметр часового пояса должен быть цифровым значением в пределах [-12..+12], а является <" & inTimeZoneUTC & ">!"
        Exit Function
    End If
    inTimeZoneUTC = CDec(inTimeZoneUTC)
    'Проверка на вхождение в допустимые значения
    If Not (Abs(inTimeZoneUTC) <= 12) Then
        outErrorText = "Параметр часового пояса должен быть цифровым значением в пределах [-12..+12], а является <" & inTimeZoneUTC & ">!"
        Exit Function
    End If
    
' 02 // Чтение допустимого периода
    If Not inForceIgnoreTimeLimits Then
        If Not fExtractTimeFromHourText(inNode.GetAttribute("start"), tStartTime, outErrorText) Then
            outErrorText = "Нода F63 аттрибут @start > " & outErrorText
            Exit Function
        End If
        If Not fExtractTimeFromHourText(inNode.GetAttribute("end"), tEndTime, outErrorText) Then
            outErrorText = "Нода F63 аттрибут @end > " & outErrorText
            Exit Function
        End If
        'Логика границ периода
        If tEndTime < tStartTime Then
            outErrorText = "Ошибка логики, время старта периода <" & Format(tStartTime, "hhmm") & "> превышает время его окончания <" & Format(tEndTime, "hhmm") & ">!"
            Exit Function
        End If
    End If
    
' 03 // Текущее время и его коррекция
    tShiftTime = (-inLocalUTC + inTimeZoneUTC) / 24
    tCurrentTime = Time() + tShiftTime
    outCurrentDateTime = Now() + tShiftTime
    
' 04 // Сравнение текущего времени и его допустимых пределов
    'Debug.Print "START=" & Format(tStartTime, "hhmm") & " NOW=" & Format(tCurrentTime, "hhmm") & " END=" & Format(tEndTime, "hhmm")
    If Not inForceIgnoreTimeLimits Then
        If Not (tCurrentTime > tStartTime And tCurrentTime < tEndTime) Then
            Exit Function
        End If
    End If
    
' XX // Завершение
    fIsActivePeriod = True
End Function

Private Function fExtractTimeFromHourText(inValue, outValue, outErrorText)
Dim tHours, tMinutes
' 00 // Предопределения
    fExtractTimeFromHourText = False
    outErrorText = vbNullString
    outValue = -1
' 01 // Проверки
    'Удалось ли прочитать?
    If IsNull(inValue) Then
        outErrorText = "Не удалось прочитать входящий параметр для конвертации в часы!"
        Exit Function
    End If
    'Длина параметра в символах
    If Len(inValue) <> 4 Then
        outErrorText = "Нарушен синтаксис представления входящего параметра <" & inValue & ">, а должен быть [ЧЧММ]!"
        Exit Function
    End If
    'Цифровое ли значение?
    If Not IsNumeric(inValue) Then
        outErrorText = "Нарушен синтаксис представления входящего параметра <" & inValue & ">, а должен быть [ЧЧММ]!"
        Exit Function
    End If
' 02 // Чтение параметра
    tHours = CDec(Left(inValue, 2))
    tMinutes = CDec(Right(inValue, 2))
    'Проверка часа
    If tHours < 0 Or tHours > 24 Then
        outErrorText = "Нарушен синтаксис представления входящего параметра <" & inValue & ">, а должен быть [ЧЧММ] (ЧЧ 00-23)!"
        Exit Function
    End If
    'Проверка минут
    If tMinutes < 0 Or tMinutes > 59 Then
        outErrorText = "Нарушен синтаксис представления входящего параметра <" & inValue & ">, а должен быть [ЧЧММ] (ММ 00-59)!"
        Exit Function
    End If
' XX // Завершение
    outValue = tHours / 24 + tMinutes / 1440 'приведем к общему формату времени
    fExtractTimeFromHourText = True
End Function

Private Function fReadBasisData(inNode, outTraderID, outGTPID, outTimeZoneUTC, outLogin, outPassword, outGateIgnore, outUncommIgnore, outErrorText)
Dim tNode, tXPathString, tTempNode, tValue, tBasisGTPNode
' 00 // Предопределения
    fReadBasisData = False
    outErrorText = vbNullString
    outGTPID = vbNullString
    outTimeZoneUTC = -100
' 01 // Чтение логина для сервиса Energy2010 из BASIS
    tXPathString = "self::node()"
    If Not fGetAttributeCFG(gXMLBasis, tXPathString, "login", outLogin, tNode, outErrorText, inNode) Then: Exit Function
' 02 // Чтение признака игнорирования гейта
    If Not fGetAttributeCFG(gXMLBasis, tXPathString, "gateignore", outGateIgnore, tNode, outErrorText, inNode) Then: Exit Function
    If outGateIgnore = "1" Then
        outGateIgnore = True
    Else
        outGateIgnore = False
    End If
' 03 // Чтение признака игнорирования некоммерческой информации в расчётах по макетам 80020/80040
    tValue = fGetAttributeCFG(gXMLBasis, tXPathString, "uncommignore", outUncommIgnore, tNode, outErrorText, inNode)
    If outUncommIgnore = "1" Then
        outUncommIgnore = True
    Else
        outUncommIgnore = False
    End If
' 04 // Чтение кода ГТП из BASIS
    tXPathString = "ancestor::gtp"
    If Not fGetAttributeCFG(gXMLBasis, tXPathString, "id", outGTPID, tBasisGTPNode, outErrorText, inNode) Then: Exit Function
' 05 // Чтение кода Торговца из BASIS
    tXPathString = "ancestor::trader"
    If Not fGetAttributeCFG(gXMLBasis, tXPathString, "id", outTraderID, tNode, outErrorText, inNode) Then: Exit Function
' 06 // Чтение пароля для сервиса Energy2010 из BASIS
    tXPathString = "//trader[@id='" & outTraderID & "']/service[@id='soenergy2010']/item[@login='" & outLogin & "']"
    If Not fGetAttributeCFG(gXMLCredentials, tXPathString, "password", outPassword, tNode, outErrorText) Then: Exit Function
' 07 // Чтение кода ценовой зоны выбранной ГТП из BASIS и Dictionary
    ' .01 // Чтение кода региона из BASIS
    tXPathString = "child::settings"
    If Not fGetAttributeCFG(gXMLBasis, tXPathString, "subjectid", tValue, tTempNode, outErrorText, tBasisGTPNode) Then: Exit Function
    ' .02 // Чтение кода ценовой зоны
    tXPathString = "//subjects/subject[@id='" & tValue & "']"
    If Not fGetAttributeCFG(gXMLDictionary, tXPathString, "tradezone", tValue, tTempNode, outErrorText) Then: Exit Function
    ' .03 // Чтение кода часового пояса ценовой зоны
    tXPathString = "//tradezones/tradezone[@id='" & tValue & "']"
    If Not fGetAttributeCFG(gXMLDictionary, tXPathString, "timezone", tValue, tTempNode, outErrorText) Then: Exit Function
    ' .04 // Чтение часового пояса
    tXPathString = "//timezones/timezone[@id='" & tValue & "']"
    If Not fGetAttributeCFG(gXMLDictionary, tXPathString, "utc", outTimeZoneUTC, tTempNode, outErrorText) Then: Exit Function
' XX // Завершение
    fReadBasisData = True
End Function




