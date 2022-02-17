Attribute VB_Name = "TimerUnit"
'������ ���������� ��������� � ����������� TimerUnit
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
                    uCDebugPrint tLogTag, 2, "������ T" & tIndex & " �� ������� ��������������!"
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
        .Description = "������ �������� ���������"
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
        .Description = "�������������� ���������� ������� 80020/80040"
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
        .Description = "����� �������� ����� ��� ���"
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
        .Description = "�������� ������ (����� 63) � ������� ������� ��"
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
        .Description = "����������� �������� ����� ������ �������"
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

'������ ������� ��������� ������
Public Sub fMainFlowTimerStart()
    Dim tLogTag
        
    tLogTag = fGetLogTag("MAINFLOW_START")
    
    If Not (gTimeControl.Active) Then: fMainFlowTimerInit
        
    If gTimeControl.Active Then
        If gTimerID(cnstMainFlowTimerIndex) = 0 Then
            fActivateTimer cnstMainFlowInterval, cnstMainFlowTimerIndex
            If gTimerID(cnstMainFlowTimerIndex) <> 0 Then: uCDebugPrint tLogTag, 0, "����������� ������ ��������� ������! [�������� TICK = " & cnstMainFlowInterval & " ���]"
        Else
            uCDebugPrint tLogTag, 0, "������ ��������� ������ ��� �������!"
        End If
    Else
        uCDebugPrint tLogTag, 2, "�� ������� ����������� �������!"
    End If
End Sub

'��������� ������� ��������� ������
Public Sub fMainFlowTimerStop()
    Dim tLogTag
    tLogTag = fGetLogTag("MAINFLOW_STOP")
    fDeactivateTimer (cnstMainFlowTimerIndex) '��������� � ������� � ����� ���������
    If gTimerID(cnstMainFlowTimerIndex) = 0 Then
        gTimeControl.Active = False
        uCDebugPrint tLogTag, 0, "���������� ������ ��������� ������!"
    End If
End Sub

'����� ��������� ������ (���������� �� �������)
Public Sub fMainFlowTimerCall()
Dim tTimerString, tReportString, tDebug, tIndex, tState, tLogTag

    tLogTag = fGetLogTag("MAINFLOW_CALL")
    If Not (gTimeControl.Active) Then: fMainFlowTimerInit
    If gTimeControl.Active Then
        tDebug = True
        tReportString = "�������� ����� ������ � �������������� ������ (TICK=" & gTimeControl.CallCount & ")"
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
        uCDebugPrint tLogTag, 2, "�� ������� �������� �������� ����� � �������������� ������!"
    End If
End Sub

Public Sub fMainFlowManualCall_MailScanner()
    fManualTimerCall 0 '��� ������ �������� ����� ����� ��� ��� ���
End Sub

Public Sub fMainFlowManualCall_ASender()
    fManualTimerCall 1 '��� ������ �������� ����� ����� ��� ��� ���
End Sub

Public Sub fMainFlowManualCall_BRForecast()
    fManualTimerCall 2 '��� ������ �������� ����� ����� ��� ��� ���
End Sub

Public Sub fMainFlowManualCall_F63()
    fManualTimerCall 3 '��� ������ �������� ����� ����� ��� ��� ���
End Sub

Public Sub fMainFlowManualCall_CFGDrop()
    fManualTimerCall 4 '��� ������ �������� ����� ����� ��� ��� ���
End Sub

' �������� �������� �������������� ��������
Private Function fDataSourceCollisionCheck(inDataListString)
    Dim tElements, tElement, tIndex, tLogTag
    
    ' 01 // ����������
    tLogTag = fGetLogTag("fDataSourceCollisionCheck")
    fDataSourceCollisionCheck = True
    tElements = Split(UCase(inDataListString), ",")
    
    ' 02 // ������� ������ ���������
    For Each tElement In tElements
    
        tIndex = fGetDataSourceIndexByTag(tElement)
        
        If tIndex <> -1 Then
            ' ���� ���� ����������� AccessLimit � ������ ��� ������������ ������ ��������� - �� �������� ��������������
            If gDataSourceList.Item(tIndex).AccessLimit And gDataSourceList.Item(tIndex).AccessCurrent > 0 Then
                fDataSourceCollisionCheck = False
                Exit Function
            End If
        Else
            uCDebugPrint tLogTag, 1, "SOURCE not FOUND in list > " & tElement
        End If
        
    Next
End Function

' �������������� �������� ��� ��������� �������
Private Sub fDataSourceUse(inDataListString)
    Dim tElements, tElement, tIndex, tLogTag
    
    ' 01 // ����������
    tLogTag = fGetLogTag("fDataSourceUse")
    tElements = Split(UCase(inDataListString), ",")
    
    ' 02 // ������� ������ ���������
    For Each tElement In tElements
        
        '������� �� DataSourceList
        'tLock = False
        tIndex = fGetDataSourceIndexByTag(tElement)
        
        If tIndex <> -1 Then
            gDataSourceList.Item(tIndex).AccessCurrent = gDataSourceList.Item(tIndex).AccessCurrent + 1
        Else
            uCDebugPrint tLogTag, 1, "SOURCE not FOUND in list > " & tElement
        End If
            
    Next
    
    ' 03 // ����� ����������� ��������� �������������� ��������
    fDataSourceShow
End Sub

'����� ����������� ��������� �������������� ��������
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

' ������ �������������� ��������
Private Sub fDataSourceRelease(inDataListString)
Dim tElements, tElement, tIndex, tLogTag
    
    ' 01 // ����������
    tLogTag = fGetLogTag("fDataSourceRelease")
    tElements = Split(UCase(inDataListString), ",")
    
    ' 02 // ������� ������ ���������
    For Each tElement In tElements
    
        tIndex = fGetDataSourceIndexByTag(tElement)
        
        If tIndex <> -1 Then
            gDataSourceList.Item(tIndex).AccessCurrent = gDataSourceList.Item(tIndex).AccessCurrent - 1
            If gDataSourceList.Item(tIndex).AccessCurrent < 0 Then
                gDataSourceList.Item(tIndex).AccessCurrent = 0
                uCDebugPrint tLogTag, 1, "�������� �� ������������ �� ������������� ��������� <INX:" & tIndex & "; TAG:" & tElement & "> !"
            End If
        Else
            uCDebugPrint tLogTag, 1, "SOURCE not FOUND in list > " & tElement
        End If
        
    Next
    
    ' 03 // ����� ����������� ��������� �������������� ��������
    fDataSourceShow
End Sub

'�������� ���� ��������� ��������������
Private Sub fMainFlow(Optional inSingleMode = False)
    Dim tIndex, tInit, tLogTag, tInternalTickIndex, tTimeConsume, tCallMode
    
    tLogTag = fGetLogTag("MAINFLOW_TICK")
    
    If Not fConfiguratorInit Then
        uCDebugPrint tLogTag, 2, "�������� ����� �� ���� ���������������� ������������!"
        Exit Sub
    End If
    
    If inSingleMode Then
        uCDebugPrint tLogTag, 2, "������������ ����� ��������� ��������!"
        tCallMode = "������"
    Else
        gTimeControl.CallCount = gTimeControl.CallCount + 1
        tCallMode = "����"
    End If
    
    tInternalTickIndex = gTimeControl.CallCount
    
' 01 // �������� � ��������� ��� ��������� ������������� �������
    For tIndex = 0 To gTimeControl.TimersCount
        With gTimeControl.Timers(tIndex)
            If .Enabled Then
                Select Case .Mode
                    Case 0:
                        'TIMEOUT if ACTIVE
                        If .Active Then
                            If Not inSingleMode Then: .TimeOut = .TimeOut + 1
                            If .TimeOut >= .FixTimeOut Then
                                uCDebugPrint tLogTag, 1, "������ <" & tIndex & "> �������������� �� �������� [" & .TimeOut & "/" & .FixTimeOut & "]!"
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
    
' 02 // OnTimer ��������
    For tIndex = 0 To gTimeControl.TimersCount
        With gTimeControl.Timers(tIndex)
            If .Enabled And .Called And Not .Active Then  '���� ������
                If fDataSourceCollisionCheck(.DataSourceList) Then
                    If fXMLSmartUpdate(.DataSourceList) Then
                        tTimeConsume = GetTickCount
                        fDataSourceUse .DataSourceList
                        .Active = True '���� ����������
                        .TimeOut = 0
                        .CallCount = .CallCount + 1
                        fDirectCall tIndex '����� ���������� ���������� ������ ������� (���� ������� ������ � �������� - ���������� ���� ������ �������������, ����� �� ���������� ��� ������� �����, ��� ������������ ����� �������)
                        fMainFlowSubTimerFinisher tIndex
                        
                        'timecontrol
                        tTimeConsume = GetTickCount - tTimeConsume
                        uCDebugPrint tLogTag, 0, "��� [#" & tInternalTickIndex - 1 & "] ������ <" & tIndex & "> [�����: " & tCallMode & "] ����������: " & tTimeConsume / 1000 & " ���."
                        
                        'overload preventer
                        If tInternalTickIndex <> gTimeControl.CallCount Then
                            uCDebugPrint tLogTag, 1, "��� [#" & tInternalTickIndex - 1 & "] ���������� ������ ��� ����� ���� (" & cnstMainFlowInterval & " ���); ��������� ������ <" & tIndex & "> ���������� �����."
                            Exit For
                        End If
                        
                    Else
                        fMainFlowSubTimerFinisher tIndex, 1
                        uCDebugPrint tLogTag, 1, "������ <" & tIndex & "> �� ���� �������� ��������� ��������! ������ ������ �� ��������� �������� �������!"
                    End If
                Else
                    uCDebugPrint tLogTag, 1, "������ <" & tIndex & "> ������� �������� ������� �� ������! ������ ������ �� ��������� �������� ��������� ������!"
                    .Called = False
                End If
            End If
        End With
    Next
' 03 // ����� ��������
End Sub

Private Sub fDirectCall(inIndex)
    Select Case inIndex
        Case 0: fGetMail 'interceptor module \\ �������� �����
        Case 1: fXMLASender 'interceptor module \\ �������� �������
        Case 2: fBRForecastMain 'BRForecast module \\ ������� ����� ���
        Case 3: fForm63Main 'F63 module \\ ������ � ������� ������� ��
        Case 4: fCFGDropMain 'CFGDrop module \\ ����������� ��������
    End Select
End Sub

Private Sub fManualTimerCall(inTimerIndex)
    Dim tLogTag
    
    tLogTag = fGetLogTag("MANUALCALL")
    uCDebugPrint tLogTag, 0, "������ ����� ������� <" & inTimerIndex & ">"
    If Not (gTimeControl.Active) Then: fMainFlowTimerInit
    If gTimeControl.Active Then
        uCDebugPrint tLogTag, 0, "����� ������� <" & inTimerIndex & "> [" & gTimeControl.Timers(inTimerIndex).Description & "] [������: " & gTimeControl.Timers(inTimerIndex).Active & "]"
        gTimeControl.Timers(inTimerIndex).TempValue = 0 '���������� ������ ����������
        fMainFlow (True)
        uCDebugPrint tLogTag, 0, "������ ����� ������� <" & inTimerIndex & "> �������!"
    Else
        uCDebugPrint tLogTag, 1, "������ ����� ������� <" & inTimerIndex & "> �� ������!"
    End If
End Sub
