Attribute VB_Name = "TimerUnit"
'Модуль управления таймерами и автоматикой TimerUnit
Option Explicit

Private Const cnstMainFlowTimerIndex = 1
Private Const cnstMainFlowTag = "MAINFLOW"
Private Const cnstMainFlowInterval = 60
Private Const cnstModuleName = "TimerUnit"
Private Const cnstModuleVersion = "003"
Private Const cnstModuleDate = "26-06-2020"

Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongLong, ByVal nIDEvent As LongLong, ByVal uElapse As LongLong, ByVal lpTimerfunc As LongLong) As LongLong
Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongLong, ByVal nIDEvent As LongLong) As LongLong

Type TTimerConfig
    Enabled As Boolean
    Active As Boolean
    Called As Boolean
    Mode As Byte '0 - periodic (time - interval in minutes)
    FixValue As Variant
    TempValue As Variant
    CallCount As LongLong
    DataSourceList As String
    TimeOut As Variant
    FixTimeOut As Variant
    Description As String
End Type

Type TTimerControl
    Timers() As TTimerConfig
    TimersCount As Variant
    CallCount As LongLong
    Active As Boolean
End Type

Dim gTimerID(2) As LongLong 'Need a timer ID to eventually turn off the timer. If the timer ID <> 0 then the timer is running
Dim gTimeControl As TTimerControl

Public Sub fTest()
Dim tA, tB
    uDebugPrint "START"
    tA = GetTickCount
    tB = tA + 1000
    While GetTickCount < tB
    Wend
    uDebugPrint "OVER"
End Sub

' for external stops
'Public Sub fMainFlowTimerStop()
'    fDeactivateTimer (0)
'End Sub

Private Function fGetLogTag(inTagText)
    fGetLogTag = cnstMainFlowTag & "." & inTagText
End Function

' Deactivate TIMER by INDEX
Private Sub fDeactivateTimer(inIndex)
    Dim tSuccess As LongLong             '<~ Corrected here
    Dim tIndex, tLogTag
    
    ' 01 // Prepare
    tLogTag = cnstMainFlowTag & ".TIMERDEACTIVE"
    
    ' 02 // Get timer by index
    For tIndex = 1 To UBound(gTimerID)
        If inIndex = tIndex Or inIndex = 0 Then 'if inIndex = 0 - stop ALL timers
            If gTimerID(tIndex) <> 0 Then
                tSuccess = KillTimer(0, gTimerID(tIndex))
                If tSuccess = 0 Then
                    uCDebugPrint tLogTag, 2, "Таймер T" & tIndex & " не удалось деактивировать!"
                Else
                    gTimerID(tIndex) = 0
                End If
            End If
        End If
    Next
End Sub

' Create and start TIMER by INDEX
Private Sub fActivateTimer(ByVal inSeconds As Long, inIndex)
  
  inSeconds = inSeconds * 1000 'The SetTimer call accepts milliseconds, so convert to seconds
  If gTimerID(inIndex) <> 0 Then: Call fDeactivateTimer(inIndex) 'Check to see if timer is running before call to SetTimer
  
  Select Case inIndex
    Case 1: gTimerID(inIndex) = SetTimer(0, 0, inSeconds, AddressOf fTriggerTimerT1)
    Case 2: gTimerID(inIndex) = SetTimer(0, 0, inSeconds, AddressOf fTriggerTimerT2)
  End Select
  
  If gTimerID(inIndex) = 0 Then
    uDebugPrint "TIM: T" & inIndex & " Failed to activate!"
  End If
End Sub

Public Sub fTriggerTimerT1(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idevent As Long, ByVal Systime As Long)
    'fDeactivateTimer (1)
    'fXML80020Reprocessor
    fMainFlowTimerCall
End Sub

Public Sub fTriggerTimerT2(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idevent As Long, ByVal Systime As Long)
    fDeactivateTimer (2)
    'fXMLASender
End Sub

' MAIN FLOW INIT
Private Sub fMainFlowTimerInit()
    gTimeControl.Active = False
    gTimeControl.TimersCount = -1 'drop to default
    
    '= TIMER 1 == [MailScanner]
    'BASIS, CONVERTER,FRAME,CALENDAR,R80020DB,MAILSCAN, XSD80020V2,XSDFORECAST,DICTIONARY,BRFORECAST,R30308DB ,CALCROUTE,CALCDB,XSD80040V2,F63DB,CREDENTIALS
    gTimeControl.TimersCount = gTimeControl.TimersCount + 1
    ReDim Preserve gTimeControl.Timers(gTimeControl.TimersCount)
    With gTimeControl.Timers(gTimeControl.TimersCount)
        .Enabled = True
        .Active = False
        .Called = False
        .Mode = 0
        .FixValue = 5 'minutes
        .TempValue = .FixValue
        .FixTimeOut = .FixValue
        .TimeOut = 0
        .CallCount = 0
        .DataSourceList = "BASIS,CONVERTER,FRAME,CALENDAR,R80020DB,MAILSCAN,DICTIONARY,R30308DB,XSD80020V2,XSD80040V2,CREDENTIALS" '"0,1,2,3,4,5,6,8,10"
        .Description = "Сканер почтовых сообщений"
    End With
    
    '= TIMER 2 == [ASender]
    gTimeControl.TimersCount = gTimeControl.TimersCount + 1
    ReDim Preserve gTimeControl.Timers(gTimeControl.TimersCount)
    With gTimeControl.Timers(gTimeControl.TimersCount)
        .Enabled = True
        .Active = False
        .Called = False
        .Mode = 0
        .FixValue = 10 'minutes
        .TempValue = .FixValue
        .FixTimeOut = .FixValue
        .TimeOut = 0
        .CallCount = 0
        .DataSourceList = "BASIS,CALENDAR,R80020DB,CREDENTIALS" ' "0,3,4"
        .Description = "Автоматический посылатель макетов 80020/80040"
    End With
    
    '= TIMER 3 == [BRForecast]
    gTimeControl.TimersCount = gTimeControl.TimersCount + 1
    ReDim Preserve gTimeControl.Timers(gTimeControl.TimersCount)
    With gTimeControl.Timers(gTimeControl.TimersCount)
        .Enabled = True
        .Active = False
        .Called = False
        .Mode = 0
        .FixValue = 10 'minutes
        .TempValue = .FixValue
        .FixTimeOut = .FixValue
        .TimeOut = 0
        .CallCount = 0
        .DataSourceList = "BASIS,CALENDAR,XSDBRFORECAST,DICTIONARY,BRFORECAST" ' "0,3,7,8,9"
        .Description = "Отчет прогноза часов пик АТС"
    End With
    
    '= TIMER 4 == [F63]
    gTimeControl.TimersCount = gTimeControl.TimersCount + 1
    ReDim Preserve gTimeControl.Timers(gTimeControl.TimersCount)
    With gTimeControl.Timers(gTimeControl.TimersCount)
        .Enabled = True
        .Active = False
        .Called = False
        .Mode = 0
        .FixValue = 10 'minutes
        .TempValue = .FixValue
        .FixTimeOut = .FixValue
        .TimeOut = 0
        .CallCount = 0
        .DataSourceList = "BASIS,FRAME,CALENDAR,R80020DB,XSD80020V2,DICTIONARY,CALCROUTE,CALCDB,XSD80040V2,F63DB,CREDENTIALS" '"0,2,3,4,6,8,12,13,14,15,16" 'v2 11 -> 16
        .Description = "Рассылка данных (форма 63) в систему Энергия СО"
    End With
    
    '= TIMER 5 == [CFGDrop]
    gTimeControl.TimersCount = gTimeControl.TimersCount + 1
    ReDim Preserve gTimeControl.Timers(gTimeControl.TimersCount)
    With gTimeControl.Timers(gTimeControl.TimersCount)
        .Enabled = True
        .Active = False
        .Called = False
        .Mode = 0
        .FixValue = 10 'minutes
        .TempValue = 5
        .FixTimeOut = 5
        .TimeOut = 0
        .CallCount = 0
        .DataSourceList = "BASIS,CALENDAR,DICTIONARY" '"0,3,8"
        .Description = "Копирование конфигов папку общего доступа"
    End With
    
    ' = DONE
    gTimeControl.Active = True
    gTimeControl.CallCount = 0
End Sub


'TIMER FINISHER
'0 - normal
'1 - on external error
Private Sub fMainFlowSubTimerFinisher(inTimerIndex, Optional inMode = 0)
    
    ' 01 // Preventer
    If Not gTimeControl.Active Then: Exit Sub
    If inTimerIndex < 0 Or inTimerIndex > gTimeControl.TimersCount Then: Exit Sub
    
    ' 02 // Finisher
    With gTimeControl.Timers(inTimerIndex)
        
        If Not .Enabled Then: Exit Sub
        If Not .Called Then: Exit Sub
        
        .Active = False
        .Called = False
        
        Select Case .Mode
            Case 0:
                .TempValue = .FixValue
                .TimeOut = 0
        End Select
        
        Select Case inMode
            Case 0: fDataSourceRelease .DataSourceList
        End Select
    End With
End Sub

'Запуск таймера основного потока
Public Sub fMainFlowTimerStart()
    Dim tLogTag
        
    tLogTag = fGetLogTag("MAINFLOW_START")
    
    If Not (gTimeControl.Active) Then: fMainFlowTimerInit
        
    If gTimeControl.Active Then
        If gTimerID(cnstMainFlowTimerIndex) = 0 Then
            fActivateTimer cnstMainFlowInterval, cnstMainFlowTimerIndex
            If gTimerID(cnstMainFlowTimerIndex) <> 0 Then: uCDebugPrint tLogTag, 0, "Активирован таймер основного потока! [Интервал TICK = " & cnstMainFlowInterval & " сек]"
        Else
            uCDebugPrint tLogTag, 0, "Таймер основного потока уже активен!"
        End If
    Else
        uCDebugPrint tLogTag, 2, "Не удалось подготовить таймеры!"
    End If
End Sub

'Остановка таймера основного потока
Public Sub fMainFlowTimerStop()
    Dim tLogTag
    tLogTag = fGetLogTag("MAINFLOW_STOP")
    fDeactivateTimer (cnstMainFlowTimerIndex) 'обращение к таймеру с целью остановки
    If gTimerID(cnstMainFlowTimerIndex) = 0 Then
        gTimeControl.Active = False
        uCDebugPrint tLogTag, 0, "Остановлен таймер основного потока!"
    End If
End Sub

'Вызов основного потока (АВТОМАТИКА ПО ТАЙМЕРУ)
Public Sub fMainFlowTimerCall()
Dim tTimerString, tReportString, tDebug, tIndex, tState, tLogTag

    tLogTag = fGetLogTag("MAINFLOW_CALL")
    If Not (gTimeControl.Active) Then: fMainFlowTimerInit
    If gTimeControl.Active Then
        tDebug = True
        tReportString = "Основной поток вызван в автоматическом режиме (TICK=" & gTimeControl.CallCount & ")"
        If tDebug Then
            tTimerString = vbNullString
            For tIndex = 0 To gTimeControl.TimersCount
                If gTimeControl.Timers(tIndex).Active Then
                    tState = "A"
                Else
                    tState = "x"
                End If
                tTimerString = tTimerString & tState
            Next
            tTimerString = " TIMERS: " & tTimerString
            tReportString = tReportString & tTimerString
        End If
        uCDebugPrint tLogTag, 0, tReportString
        fMainFlow
    Else
        uCDebugPrint tLogTag, 2, "Не удалось вызывать основной поток в автоматическом режиме!"
    End If
End Sub

Public Sub fMainFlowManualCall_MailScanner()
    fManualTimerCall 0 'для ручных запусков через ленту или еще как
End Sub

Public Sub fMainFlowManualCall_ASender()
    fManualTimerCall 1 'для ручных запусков через ленту или еще как
End Sub

Public Sub fMainFlowManualCall_BRForecast()
    fManualTimerCall 2 'для ручных запусков через ленту или еще как
End Sub

Public Sub fMainFlowManualCall_F63()
    fManualTimerCall 3 'для ручных запусков через ленту или еще как
End Sub

Public Sub fMainFlowManualCall_CFGDrop()
    fManualTimerCall 4 'для ручных запусков через ленту или еще как
End Sub

' ПРОВЕРКА КОЛЛИЗИЙ РЕЗЕРВИРОВАНИЯ ресурсов
Private Function fDataSourceCollisionCheck(inDataListString)
    Dim tElements, tElement, tIndex, tLogTag
    
    ' 01 // Подготовка
    tLogTag = fGetLogTag("fDataSourceCollisionCheck")
    fDataSourceCollisionCheck = True
    tElements = Split(UCase(inDataListString), ",")
    
    ' 02 // Перебор списка элементов
    For Each tElement In tElements
    
        tIndex = fGetDataSourceIndexByTag(tElement)
        
        If tIndex <> -1 Then
            ' Если есть ограничение AccessLimit и ресурс уже используется другой операцией - то КОЛЛИЗИЯ РЕЗЕРВИРОВАНИЯ
            If gDataSourceList.Item(tIndex).AccessLimit And gDataSourceList.Item(tIndex).AccessCurrent > 0 Then
                fDataSourceCollisionCheck = False
                Exit Function
            End If
        Else
            uCDebugPrint tLogTag, 1, "SOURCE not FOUND in list > " & tElement
        End If
        
    Next
End Function

' РЕЗЕРВИРОВАНИЕ ресурсов для активации таймера
Private Sub fDataSourceUse(inDataListString)
    Dim tElements, tElement, tIndex, tLogTag
    
    ' 01 // Подготовка
    tLogTag = fGetLogTag("fDataSourceUse")
    tElements = Split(UCase(inDataListString), ",")
    
    ' 02 // Перебор списка элементов
    For Each tElement In tElements
        
        'Перебор по DataSourceList
        'tLock = False
        tIndex = fGetDataSourceIndexByTag(tElement)
        
        If tIndex <> -1 Then
            gDataSourceList.Item(tIndex).AccessCurrent = gDataSourceList.Item(tIndex).AccessCurrent + 1
        Else
            uCDebugPrint tLogTag, 1, "SOURCE not FOUND in list > " & tElement
        End If
            
    Next
    
    ' 03 // Показ актуального состояния резервирования ресурсов
    fDataSourceShow
End Sub

'ПОКАЗ АКТУАЛЬНОГО состояния РЕЗЕРВИРОВАНИЯ ресурсов
Private Sub fDataSourceShow()
Dim tIndex, tResultString, tLogTag
    
    tLogTag = fGetLogTag("fDataSourceShow")
    tResultString = vbNullString
    
    For tIndex = 0 To gDataSourceList.Count
        If tResultString = vbNullString Then
            tResultString = gDataSourceList.Item(tIndex).AccessCurrent
        Else
            tResultString = tResultString & ":" & gDataSourceList.Item(tIndex).AccessCurrent
        End If
    Next
    
    uCDebugPrint tLogTag, 0, "STATUS > " & tResultString
End Sub

' СНЯТИЕ РЕЗЕРВИРОВАНИЯ ресурсов
Private Sub fDataSourceRelease(inDataListString)
Dim tElements, tElement, tIndex, tLogTag
    
    ' 01 // Подготовка
    tLogTag = fGetLogTag("fDataSourceRelease")
    tElements = Split(UCase(inDataListString), ",")
    
    ' 02 // Перебор списка элементов
    For Each tElement In tElements
    
        tIndex = fGetDataSourceIndexByTag(tElement)
        
        If tIndex <> -1 Then
            gDataSourceList.Item(tIndex).AccessCurrent = gDataSourceList.Item(tIndex).AccessCurrent - 1
            If gDataSourceList.Item(tIndex).AccessCurrent < 0 Then
                gDataSourceList.Item(tIndex).AccessCurrent = 0
                uCDebugPrint tLogTag, 1, "Аномалия по освобождению от использования источника <INX:" & tIndex & "; TAG:" & tElement & "> !"
            End If
        Else
            uCDebugPrint tLogTag, 1, "SOURCE not FOUND in list > " & tElement
        End If
        
    Next
    
    ' 03 // Показ актуального состояния резервирования ресурсов
    fDataSourceShow
End Sub

'ОСНОВНОЙ БЛОК обработки псевдотаймеров
Private Sub fMainFlow(Optional inSingleMode = False)
    Dim tIndex, tInit, tLogTag, tInternalTickIndex, tTimeConsume, tCallMode
    
    tLogTag = fGetLogTag("MAINFLOW_TICK")
    
    If Not fConfiguratorInit Then
        uCDebugPrint tLogTag, 2, "Основной поток не смог инициализировать конфигурации!"
        Exit Sub
    End If
    
    If inSingleMode Then
        uCDebugPrint tLogTag, 2, "Внеочередной вызов обработки таймеров!"
        tCallMode = "РУЧНОЙ"
    Else
        gTimeControl.CallCount = gTimeControl.CallCount + 1
        tCallMode = "АВТО"
    End If
    
    tInternalTickIndex = gTimeControl.CallCount
    
' 01 // Действия с таймерами для выявления необходимости реакции
    For tIndex = 0 To gTimeControl.TimersCount
        With gTimeControl.Timers(tIndex)
            If .Enabled Then
                Select Case .Mode
                    Case 0:
                        'TIMEOUT if ACTIVE
                        If .Active Then
                            If Not inSingleMode Then: .TimeOut = .TimeOut + 1
                            If .TimeOut >= .FixTimeOut Then
                                uCDebugPrint tLogTag, 1, "Таймер <" & tIndex & "> деактивируется по таймауту [" & .TimeOut & "/" & .FixTimeOut & "]!"
                                .TimeOut = .FixTimeOut
                                fMainFlowSubTimerFinisher tIndex
                            End If
                        'TIMER if NOT ACTIVE
                        Else
                            If Not inSingleMode Then: .TempValue = .TempValue - 1
                            If .TempValue <= 0 Then
                                .TempValue = 0 'to prevent infinity
                                .Called = True
                            End If
                        End If
                End Select
            End If
        End With
    Next
    
' 02 // OnTimer развёртка
    For tIndex = 0 To gTimeControl.TimersCount
        With gTimeControl.Timers(tIndex)
            If .Enabled And .Called And Not .Active Then  'если вызван
                If fDataSourceCollisionCheck(.DataSourceList) Then
                    If fXMLSmartUpdate(.DataSourceList) Then
                        tTimeConsume = GetTickCount
                        fDataSourceUse .DataSourceList
                        .Active = True 'флаг активности
                        .TimeOut = 0
                        .CallCount = .CallCount + 1
                        fDirectCall tIndex 'нужен обработчик результата работы таймера (если копятся ошибки и таймауты - необходимо этот таймер заблокировать, чтобы не подвергать всю систему крашу, или инициировать ребут системы)
                        fMainFlowSubTimerFinisher tIndex
                        
                        'timecontrol
                        tTimeConsume = GetTickCount - tTimeConsume
                        uCDebugPrint tLogTag, 0, "Тик [#" & tInternalTickIndex - 1 & "] таймер <" & tIndex & "> [ВЫЗОВ: " & tCallMode & "] исполнялся: " & tTimeConsume / 1000 & " сек."
                        
                        'overload preventer
                        If tInternalTickIndex <> gTimeControl.CallCount Then
                            uCDebugPrint tLogTag, 1, "Тик [#" & tInternalTickIndex - 1 & "] исполнялся больше чем время тика (" & cnstMainFlowInterval & " сек); последний таймер <" & tIndex & "> исполенный тиком."
                            Exit For
                        End If
                        
                    Else
                        fMainFlowSubTimerFinisher tIndex, 1
                        uCDebugPrint tLogTag, 1, "Таймер <" & tIndex & "> не смог получить требуемых ресурсов! Отмена вызова до следующей итерации таймера!"
                    End If
                Else
                    uCDebugPrint tLogTag, 1, "Таймер <" & tIndex & "> получил коллизию запроса на ресурс! Отмена вызова до следующей итерации основного потока!"
                    .Called = False
                End If
            End If
        End With
    Next
' 03 // Поток завершен
End Sub

Private Sub fDirectCall(inIndex)
    Select Case inIndex
        Case 0: fGetMail 'interceptor module \\ Проверка почты
        Case 1: fXMLASender 'interceptor module \\ Рассылка макетов
        Case 2: fBRForecastMain 'BRForecast module \\ Прогноз часов пик
        Case 3: fForm63Main 'F63 module \\ Данные в систему Энергия СО
        Case 4: fCFGDropMain 'CFGDrop module \\ Копирование конфигов
    End Select
End Sub

Private Sub fManualTimerCall(inTimerIndex)
    Dim tLogTag
    
    tLogTag = fGetLogTag("MANUALCALL")
    uCDebugPrint tLogTag, 0, "Ручной вызов таймера <" & inTimerIndex & ">"
    If Not (gTimeControl.Active) Then: fMainFlowTimerInit
    If gTimeControl.Active Then
        uCDebugPrint tLogTag, 0, "Вызов таймера <" & inTimerIndex & "> [" & gTimeControl.Timers(inTimerIndex).Description & "] [Статус: " & gTimeControl.Timers(inTimerIndex).Active & "]"
        gTimeControl.Timers(inTimerIndex).TempValue = 0 'триггернет запрос выполнения
        fMainFlow (True)
        uCDebugPrint tLogTag, 0, "Ручной вызов таймера <" & inTimerIndex & "> успешен!"
    Else
        uCDebugPrint tLogTag, 1, "Ручной вызов таймера <" & inTimerIndex & "> не удался!"
    End If
End Sub
