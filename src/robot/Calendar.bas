Attribute VB_Name = "Calendar"
'Инструментарий для работы с КАЛЕНДАРЁМ
Option Explicit

Private Const cnstModuleName = "CALENDAR"
Private Const cnstModuleVersion = 1
Private Const cnstModuleDate = "17-07-2018"

Private Const cnstMinYear = 2010
Private Const cnstMaxYear = 2040

Public Type TSOPeakHours
    Loaded As Boolean
    Reason As String
    ZoneID As Byte
    Year As Integer
    Month As Byte
    Hours(23) As Byte
    PairPartA() As Byte
    PairPartB() As Byte
    PairCount As Byte
    WorkDay As Boolean
End Type

Private Function fIsDateStamp(inDate, outDate)
Dim tDay, tMonth, tYear
    fIsDateStamp = False
    outDate = -1
    If Not (Len(inDate) = 8 And IsNumeric(inDate)) Then: Exit Function
    tYear = CDec(Left(inDate, 4))
    tMonth = CDec(Mid(inDate, 5, 2))
    tDay = CDec(Right(inDate, 2))
    If tYear < cnstMinYear Or tYear > cnstMaxYear Then: Exit Function
    If tMonth < 1 Or tMonth > 12 Then: Exit Function
    If tDay < 1 Or tDay > uDaysPerMonth(tMonth, tYear) Then: Exit Function
    outDate = DateSerial(tYear, tMonth, tDay)
    fIsDateStamp = True
End Function

Public Sub TestShifter()
Dim tErrorText, tDate, tValue
    If Not fConfiguratorInit Then: Exit Sub
    If Not fXMLSmartUpdate("3") Then: Exit Sub
    tDate = Date
    If Not fWorkDayShiftAdv(tDate, -256, 1, tValue, tErrorText) Then
        uDebugPrint tErrorText
    Else
        uDebugPrint tValue
    End If
End Sub

Private Function fGetShiftedMonthNode(inYear, inMonth, inShift, outNode, outErrorText)
Dim tShiftedMonth, tShiftedYear, tXPathString
    fGetShiftedMonthNode = False
    If Not gXMLCalendar.Active Or gXMLCalendar.XML Is Nothing Then
        outErrorText = "Конфиг CALENDAR не загружен!"
        Exit Function
    End If
    tShiftedMonth = CDec(inYear) * 12 + CDec(inMonth) + CDec(inShift)
    tShiftedYear = Fix(tShiftedMonth / 12)
    tShiftedMonth = tShiftedMonth - tShiftedYear * 12
    If tShiftedMonth = 0 Then
        tShiftedMonth = 12
        tShiftedYear = tShiftedYear - 1
    End If
    tXPathString = "//year[@id='" & Format(tShiftedYear, "0000") & "']/month[@id='" & Format(tShiftedMonth, "00") & "']/workdays"
    Set outNode = gXMLCalendar.XML.SelectSingleNode(tXPathString)
    If outNode Is Nothing Then
        outErrorText = "Возможно, календарь не заполнен! Не удалось найти месяц по XPath <" & tXPathString & ">."
        Exit Function
    End If
    fGetShiftedMonthNode = True
End Function

Public Function fWorkDayShiftAdv(inDate, inShift, inOutMode, outValue, outErrorText)
Dim tTempDate, tCurrentDateNode, tYear, tMonth, tDay, tXPathString, tSiblingNodes, tNode, tShiftCounter, tTargetNode, tActiveMonthNode, tMonthShift, tStartIndex, tEndIndex, tIndexA, tIndexB, tStep, tIndex, tFirstRun
'00 // Подготовка
    fWorkDayShiftAdv = False
    outValue = 0
    outErrorText = vbNullString
'00 // Проверка доступности каленадрика
    If Not gXMLCalendar.Active Or gXMLCalendar.XML Is Nothing Then
        outErrorText = "Конфиг CALENDAR не загружен!"
        Exit Function
    End If
'00 // Приведение входящей даты в нормальный вид
    If Not fIsDateStamp(inDate, tTempDate) Then
        If Not IsDate(inDate) Then
            outErrorText = "Входной параметр даты <" & inDate & "> не является датой!"
            Exit Function
        Else
            tTempDate = inDate
        End If
    End If
'00 // Поиск текущей даты в календарике
    tYear = Format(Year(tTempDate), "0000")
    tMonth = Format(Month(tTempDate), "00")
    tDay = Format(Day(tTempDate), "00")
    tXPathString = "//year[@id='" & tYear & "']/month[@id='" & tMonth & "']/workdays/day[@id='" & tDay & "']"
    Set tCurrentDateNode = gXMLCalendar.XML.SelectSingleNode(tXPathString)
    If tCurrentDateNode Is Nothing Then
        outErrorText = "Возможно, календарь не заполнен! Не удалось найти день по XPath <" & tXPathString & ">."
        Exit Function
    End If
'00 // Нода в которой будет храниться результат
    Set tTargetNode = Nothing
'00 // Для нулевого смещения своя обработка
    If inShift = 0 Then
        outErrorText = "Входным параметром смещения не может быть <0>!"
        Exit Function
'00 // Для смещения в + или - своя обработка
    Else
        tShiftCounter = Abs(inShift)
        tFirstRun = True
        'направляющие слайдера
        If inShift > 0 Then
            tStep = 1
            tStartIndex = CDec(tDay)        'индекс от нуля, поэтому -1, а берем следующий день поэтому +1; = 0 учтем это в StartIndex
        Else
            tStep = -1
            tStartIndex = CDec(tDay) - 2    'индекс от нуля, поэтому -1, а берем предыдущий день поэтому -1; = -2 учтем это в StartIndex
        End If
        'основной цикл поиска (без ограничителя)
        tMonthShift = 0
        Do While tShiftCounter > 0
            'найдем ноду месяца смещенную относительно исходной даты на tMonthShift
            If Not fGetShiftedMonthNode(tYear, tMonth, tMonthShift, tActiveMonthNode, outErrorText) Then
                'autoerror texting
                Exit Function
            End If
            tEndIndex = tActiveMonthNode.ChildNodes.Length - 1
            'двусторонний слайдер
            If tFirstRun Then 'для первого прохода надо отталкиваться от фиксированной даты текущего месяца (исходной даты)
                If tStep > 0 Then
                    tIndexA = tStartIndex
                    tIndexB = tEndIndex
                Else
                    tIndexA = tStartIndex
                    tIndexB = 0
                End If
                tFirstRun = False
            Else 'все месяцы вне исходного следует обходить от начала до конца
                If tStep > 0 Then
                    tIndexA = 0
                    tIndexB = tEndIndex
                Else
                    tIndexA = tEndIndex
                    tIndexB = 0
                End If
            End If
            'переборка месяца с учетом двусторонннего слайдера - tTargetNode сохранит результат
            For tIndex = tIndexA To tIndexB Step tStep
                If tActiveMonthNode.ChildNodes(tIndex).GetAttribute("workday") = "1" Then
                    tShiftCounter = tShiftCounter - 1
                    If tShiftCounter = 0 Then
                        Set tTargetNode = tActiveMonthNode.ChildNodes(tIndex)
                        Exit For
                    End If
                End If
            Next
            tMonthShift = tMonthShift + tStep 'tStep направление определяет
        Loop
    End If
'00 // Проверка на результат
    If tTargetNode Is Nothing Then
        outErrorText = "Логическая аномалия! Поиск дня произведен, но нода не извлечена!"
        Exit Function
    End If
'00 // Подготовка результата
    tDay = CDec(tTargetNode.GetAttribute("id")) 'day
    Set tTargetNode = tTargetNode.ParentNode.ParentNode
    tMonth = CDec(tTargetNode.GetAttribute("id")) 'month
    Set tTargetNode = tTargetNode.ParentNode
    tYear = CDec(tTargetNode.GetAttribute("id")) 'year
    outValue = DateSerial(tYear, tMonth, tDay)
    'режим вывода
    If inOutMode = 1 Then
        outValue = Format(outValue, "YYYYMMDD")
    End If
'XX // Завершение
    fWorkDayShiftAdv = True
End Function

'PP01 // Возвращает ДАТУ отступающую на inShift РАБОЧИХ дней от inDate \\ ФОРМАТ входящей и исходящей ДАТЫ - ГГГГММДД
Public Function fWorkDayShiftX(inDate, inShift) As Variant
Dim tYearA, tMonthA, tYearB, tMonthB, tYearC, tMonthC, tDay, tNodeA, tNodeB, tNodeC, tIndex, tTargetWorkDay, tNumericDay, tWorkDaysListCount, tSorted, tValue
Dim WorkDays()
'00 // Подготовка
    fWorkDayShift = 0
    If inShift = 0 Then: Exit Function
    If Not gXMLCalendar.Active Then: Exit Function 'если календарик не загружен - выход
'01 // Извлечение даты
    If Not (IsTimeStamp(inDate, tYearB, tMonthB, tDay)) Then
        uDebugPrint "CWDS: Не удалось определить дату из [" & inDate & "]."
        Exit Function
    End If
    If tDay = 0 Then
        uDebugPrint "CWDS: Не удалось определить дату из [" & inDate & "]."
        Exit Function
    End If
'02 // Поиск предыдущего месяца по календарю [-1]
    tMonthA = Format(tMonthB - 1, "00")
    tYearA = tYearB
    If tMonthA < 1 Then
        tMonthA = Format(12, "00")
        tYearA = tYearB - 1
    End If
    Set tNodeA = gXMLCalendar.XML.SelectNodes("//year[@id='" & tYearA & "']/month[@id=" & tMonthA & "]/workdays/day")
    If tNodeA.Length = 0 Then
        uDebugPrint "CWDS: Календарь не заполнен или иная ошибка (искомый год - " & tYearA & "; искомый месяц - " & tMonthA & ")."
        Exit Function
    End If
'03 // Поиск года и месяца даты по календарю [X]
    Set tNodeB = gXMLCalendar.XML.SelectNodes("//year[@id='" & tYearB & "']/month[@id=" & tMonthB & "]/workdays/day")
    If tNodeB.Length = 0 Then
        uDebugPrint "CWDS: Календарь не заполнен или иная ошибка (искомый год - " & tYearB & "; искомый месяц - " & tMonthB & ")."
        Exit Function
    End If
'04 // Поиск следующего месяца по календарю [+1]
    tMonthC = Format(tMonthB + 1, "00")
    tYearC = tYearB
    If tMonthC > 12 Then
        tMonthC = Format(1, "00")
        tYearC = tYearB + 1
    End If
    Set tNodeC = gXMLCalendar.XML.SelectNodes("//year[@id='" & tYearC & "']/month[@id=" & tMonthC & "]/workdays/day")
    If tNodeC.Length = 0 Then
        uDebugPrint "CWDS: Календарь не заполнен или иная ошибка (искомый год - " & tYearC & "; искомый месяц - " & tMonthC & ")."
        Exit Function
    End If
'05 // Создание списка
    tWorkDaysListCount = tNodeA.Length + tNodeB.Length + tNodeC.Length - 1
    ReDim WorkDays(tWorkDaysListCount)
    For tIndex = 0 To tNodeA.Length - 1
        WorkDays(tIndex) = CLng(tYearA & tMonthA & tNodeA(tIndex).Text)
    Next
    For tIndex = 0 To tNodeB.Length - 1
        WorkDays(tNodeA.Length + tIndex) = CLng(tYearB & tMonthB & tNodeB(tIndex).Text)
    Next
    For tIndex = 0 To tNodeC.Length - 1
        WorkDays(tNodeA.Length + tNodeB.Length + tIndex) = CLng(tYearC & tMonthC & tNodeC(tIndex).Text)
    Next
'06 // Сортировка списка
    tSorted = False
    Do While Not (tSorted)
        tSorted = True
        For tIndex = 0 To tWorkDaysListCount - 1
            If (WorkDays(tIndex) > WorkDays(tIndex + 1) And inShift > 0) Or (WorkDays(tIndex) < WorkDays(tIndex + 1) And inShift < 0) Then
                tSorted = False
                tValue = WorkDays(tIndex)
                WorkDays(tIndex) = WorkDays(tIndex + 1)
                WorkDays(tIndex + 1) = tValue
            End If
        Next
    Loop
'07 // Определим день сдвига
    tNumericDay = CLng(inDate)
    tTargetWorkDay = Abs(inShift)
    For tIndex = 0 To tWorkDaysListCount
        If (WorkDays(tIndex) > tNumericDay And inShift > 0) Or (WorkDays(tIndex) < tNumericDay And inShift < 0) Then
            tTargetWorkDay = tTargetWorkDay - 1
            If tTargetWorkDay = 0 Then
                tTargetWorkDay = WorkDays(tIndex)
                Exit For
            End If
        End If
    Next
    '07 // Сравним текущую дату и дату сдвига для определения вхождения
    'tNumericDay = CLng(Format(Now(), "YYYYMMDD"))
    fWorkDayShift = tTargetWorkDay
End Function

'PP02 // Получение часов СО по коду региона inRegionID, году inYear и месяцу inMonth
Public Function fGetSOPeakHoursByZone(inZoneID, inYear, inMonth) As TSOPeakHours
Dim tNode, tXPathString, tPart1A, tPart1B, tPart2A, tPart2B, tIndex, tZoneID, tYear, tMonth
'00 // Подготовка объекта
    fGetSOPeakHoursByZone.Loaded = False
    fGetSOPeakHoursByZone.Reason = "Не прочитан"
    If Not gXMLCalendar.Active Then 'если календарик не загружен - выход
        fGetSOPeakHoursByZone.Reason = "Нет источника"
        Exit Function
    End If
    If Not gXMLDictionary.Active Then 'если календарик не загружен - выход
        fGetSOPeakHoursByZone.Reason = "Нет источника"
        Exit Function
    End If
'01 // Проверка входных данных
    If Not IsNumeric(inYear) Then
        fGetSOPeakHoursByZone.Reason = "Год <inYear> указан не цифрой"
        Exit Function
    End If
    If inYear < cnstMinYear Or inYear > cnstMaxYear Then
        fGetSOPeakHoursByZone.Reason = "Год <inYear> указан вне диапазона " & cnstMinYear & "-" & cnstMaxYear
        Exit Function
    End If
    If Not IsNumeric(inMonth) Then
        fGetSOPeakHoursByZone.Reason = "Месяц <inMonth> указан не цифрой"
        Exit Function
    End If
    If inMonth < 1 Or inMonth > 12 Then
        fGetSOPeakHoursByZone.Reason = "Месяц <inMonth> указан вне диапазона 01-12"
        Exit Function
    End If
    If Not IsNumeric(inZoneID) Then
        fGetSOPeakHoursByZone.Reason = "Код региона <inZoneID> указан не цифрой"
        Exit Function
    End If
'02 // Поиск данных в КАЛЕНДАРЕ
    tZoneID = Format(inZoneID, "00")
    tYear = Format(inYear, "0000")
    tMonth = Format(inMonth, "00")
    tXPathString = "//year[@id='" & tYear & "']/month[@id=" & tMonth & "]/sopower/tradezone[@id='" & tZoneID & "']"
    Set tNode = gXMLCalendar.XML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        fGetSOPeakHoursByZone.Reason = "В календаре нет таких данных"
        Exit Function
    End If
    fGetSOPeakHoursByZone.ZoneID = tZoneID
    fGetSOPeakHoursByZone.Year = tYear
    fGetSOPeakHoursByZone.Month = tMonth
    fGetSOPeakHoursByZone.WorkDay = True
'03 // Чтение данных из КАЛЕНДАРЯ
    tPart1A = tNode.GetAttribute("starthour1")
    tPart1B = tNode.GetAttribute("endhour1")
    tPart2A = tNode.GetAttribute("starthour2")
    tPart2B = tNode.GetAttribute("endhour2")
'04 // Проверка данных из КАЛЕНДАРЯ <- нужен ли этот блок если вносить данные будет автоматика? О_О
    '1A - начало пары 1
    If IsNull(tPart1A) Then
        tPart1A = 0
    ElseIf IsNumeric(tPart1A) Then
        If tPart1A < 1 And tPart1A > 24 Then
            tPart1A = -1
        Else
            tPart1A = CInt(tPart1A)
        End If
    ElseIf tPart1A = vbNullString Then
        tPart1A = 0
    Else
        tPart1A = -1
    End If
    '1B - конец пары 1
    If IsNull(tPart1B) Then
        tPart1B = 0
    ElseIf IsNumeric(tPart1B) Then
        If tPart1B < 1 And tPart1B > 24 Then
            tPart1B = -1
        Else
            tPart1B = CInt(tPart1B)
        End If
    ElseIf tPart1B = vbNullString Then
        tPart1B = 0
    Else
        tPart1B = -1
    End If
    '2A - начало пары 2
    If IsNull(tPart2A) Then
        tPart2A = 0
    ElseIf IsNumeric(tPart2A) Then
        If tPart2A < 1 And tPart2A > 24 Then
            tPart2A = -1
        Else
            tPart2A = CInt(tPart2A)
        End If
    ElseIf tPart2A = vbNullString Then
        tPart2A = 0
    Else
        tPart2A = -1
    End If
    '2B - конец пары 2
    If IsNull(tPart2B) Then
        tPart2B = 0
    ElseIf IsNumeric(tPart2B) Then
        If tPart2B < 1 And tPart2B > 24 Then
            tPart2B = -1
        Else
            tPart2B = CInt(tPart2B)
        End If
    ElseIf tPart2B = vbNullString Then
        tPart2B = 0
    Else
        tPart2B = -1
    End If
    'Проверка пары А
    If tPart1A = 0 Or tPart1B = 0 Or tPart1A = -1 Or tPart1B = -1 Or tPart1A > tPart1B Then
        fGetSOPeakHoursByZone.Reason = "Ошибка в паре 1[" & tPart1A & ":" & tPart1B & "]"
        Exit Function
    End If
    'Проверка пары B
    If (tPart2A = 0 And tPart2B <> 0) Or (tPart2A <> 0 And tPart2B = 0) Or tPart2A > tPart2B Then
        fGetSOPeakHoursByZone.Reason = "Ошибка в паре 2[" & tPart2A & ":" & tPart2B & "] "
        Exit Function
    End If
    'Проверка перекщений пар
    If tPart2A > 0 Then
        If tPart1A > tPart2A Or tPart1B > tPart2A Then
            fGetSOPeakHoursByZone.Reason = "Ошибка в паре 1 и 2 - перекрещение 1[" & tPart1A & ":" & tPart1B & "] 2[" & tPart2A & ":" & tPart2B & "]"
            Exit Function
        End If
    End If
'05 // Создание матрицы часовок
    'заготовка
    For tIndex = 0 To 23
        fGetSOPeakHoursByZone.Hours(tIndex) = 0
    Next
    'наложение пары 1
    For tIndex = tPart1A To tPart1B
        fGetSOPeakHoursByZone.Hours(tIndex - 1) = 1
    Next
    fGetSOPeakHoursByZone.PairCount = 0
    ReDim fGetSOPeakHoursByZone.PairPartA(fGetSOPeakHoursByZone.PairCount)
    ReDim fGetSOPeakHoursByZone.PairPartB(fGetSOPeakHoursByZone.PairCount)
    fGetSOPeakHoursByZone.PairPartA(fGetSOPeakHoursByZone.PairCount) = tPart1A
    fGetSOPeakHoursByZone.PairPartB(fGetSOPeakHoursByZone.PairCount) = tPart1B
    'наложение пары 2 (если есть)
    If tPart2A > 0 Then
        For tIndex = tPart2A To tPart2B
            fGetSOPeakHoursByZone.Hours(tIndex - 1) = 1
        Next
        fGetSOPeakHoursByZone.PairCount = 1
        ReDim Preserve fGetSOPeakHoursByZone.PairPartA(fGetSOPeakHoursByZone.PairCount)
        ReDim Preserve fGetSOPeakHoursByZone.PairPartB(fGetSOPeakHoursByZone.PairCount)
        fGetSOPeakHoursByZone.PairPartA(fGetSOPeakHoursByZone.PairCount) = tPart2A
        fGetSOPeakHoursByZone.PairPartB(fGetSOPeakHoursByZone.PairCount) = tPart2B
    End If
'06 // Завершение
    fGetSOPeakHoursByZone.Loaded = True
    fGetSOPeakHoursByZone.Reason = vbNullString
End Function

Public Function fIsWorkDay(inYear, inMonth, inDay)
Dim tXPathString, tNode, tMonth, tDay
    fIsWorkDay = False
    If Not gXMLCalendar.Active Then: Exit Function
    tMonth = Format(inMonth, "00")
    tDay = Format(inDay, "00")
    tXPathString = "//year[@id='" & inYear & "']/month[@id='" & tMonth & "']/workdays/day[(@id='" & tDay & "' and @workday='1')]"
    Set tNode = gXMLCalendar.XML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then: Exit Function
    fIsWorkDay = True
End Function

Public Function fIsOnDutyDay(inDate)
Dim tXPathString, tNode, tYear, tMonth, tDay
    fIsOnDutyDay = False
    If Not gXMLCalendar.Active Then: Exit Function
    If Not IsDate(inDate) Then: Exit Function
    tYear = Format(Year(inDate), "0000")
    tMonth = Format(Month(inDate), "00")
    tDay = Format(Day(inDate), "00")
    tXPathString = "//year[@id='" & tYear & "']/month[@id='" & tMonth & "']/workdays/day[(@id='" & tDay & "' and @onduty='1')]"
    Set tNode = gXMLCalendar.XML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then: Exit Function
    fIsOnDutyDay = True
End Function

'PP03 // Получение часов СО по коду региона inRegionID, году inYear и месяцу inMonth на определенный день inDay
Public Function fGetDaySOPeakHoursByZone(inZoneID, inYear, inMonth, inDay) As TSOPeakHours
Dim tIndex, tDay
'00 // Подготовка объекта
    fGetDaySOPeakHoursByZone = fGetSOPeakHoursByZone(inZoneID, inYear, inMonth)
    If fGetDaySOPeakHoursByZone.Loaded Then
        'tDay = Format(inMonth, "00")
        If Not fIsWorkDay(inYear, inMonth, inDay) Then
            For tIndex = 0 To 23
                fGetDaySOPeakHoursByZone.Hours(tIndex) = 0
            Next
            fGetDaySOPeakHoursByZone.Reason = "Не рабочий день"
            fGetDaySOPeakHoursByZone.WorkDay = False
        End If
    End If
End Function

