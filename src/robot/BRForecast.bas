Attribute VB_Name = "BRForecast"
'BR FORECAST version 1 (20-07-2018)
Option Explicit

Const cnstModuleName = "BRForecast"
Const cnstModuleVersion = 1
Const cnstModuleDate = "20-07-2018"

Dim gForecastDataPath, gDownloadPath, gLocalInit, gStoragePath

Private Function fGetLogTag(inTagText)
    fGetLogTag = cnstModuleName & "." & inTagText
End Function

Private Function fLocalInit(Optional inForceInit As Boolean = False)
Dim tLogTag
    fLocalInit = False
    tLogTag = "BRINI"
    uADebugPrint tLogTag, "�������������"
    If inForceInit Or Not gLocalInit Then
        'variables
        gLocalInit = False
        'objects
        'paths
        gForecastDataPath = gDataPath & "\BRForecast"
        If Not (gFSO.FolderExists(gForecastDataPath)) Then
            If Not (uFolderCreate(gForecastDataPath)) Then
                uADebugPrint tLogTag, "�� ������� ����� ����� ������ <gForecastDataPath> �� ����: " & gForecastDataPath
                gDataPath = vbNullString
                Exit Function
            Else
                uADebugPrint tLogTag, "������� ����� ������ <gForecastDataPath> �� ����: " & gForecastDataPath
            End If
        End If
        gDownloadPath = gForecastDataPath & "\Temp"
        If Not (gFSO.FolderExists(gDownloadPath)) Then
            If Not (uFolderCreate(gDownloadPath)) Then
                uADebugPrint tLogTag, "�� ������� ����� ����� ������ <gDownloadPath> �� ����: " & gDownloadPath
                gDownloadPath = vbNullString
                Exit Function
            Else
                uADebugPrint tLogTag, "������� ����� ������ <gDownloadPath> �� ����: " & gDownloadPath
            End If
        End If
        gStoragePath = gForecastDataPath & "\Storage"
        If Not (gFSO.FolderExists(gStoragePath)) Then
            If Not (uFolderCreate(gStoragePath)) Then
                uADebugPrint tLogTag, "�� ������� ����� ����� ������ <gStoragePath> �� ����: " & gStoragePath
                gDownloadPath = vbNullString
                Exit Function
            Else
                uADebugPrint tLogTag, "������� ����� ������ <gStoragePath> �� ����: " & gStoragePath
            End If
        End If
        'xsd forecast
        'If Not fXSDLoader(gXSDForecast, gConfigPath) Then
        '    uADebugPrint tLogTag, "�� ������� ��������� <XSD BRForecast> �� ����: " & gXSDForecast.Name
        '    Exit Function
        'End If
        gLocalInit = True
    End If
    fLocalInit = True
    'uADebugPrint tLogTag, "OK"
End Function

Public Sub fTestingSub()
Dim tExcel, tWorkBook, tOutMail, tRange, tSheetIndex, tFullPath, tRow, tShape, tIndex
    If Not fConfiguratorInit Then: Exit Sub
    If Not fLocalInit Then: Exit Sub
    Set tExcel = CreateObject("Excel.Application")
    uDebugPrint "EXCEL things Test"
    Set tWorkBook = tExcel.Workbooks.Add
    'tWorkBook.Activate
    tSheetIndex = 1
    tFullPath = gForecastDataPath & "\TEMP.jpg"
    tWorkBook.WorkSheets(tSheetIndex).Activate
    For tRow = 2 To 25
        tWorkBook.WorkSheets(tSheetIndex).Cells(tRow, 1).Value = tRow - 1
        tWorkBook.WorkSheets(tSheetIndex).Cells(tRow, 2).Value = Fix(Rnd() * 20)
    Next
    'tWorkBook.Worksheets(tSheetIndex).Cells(1, 1).Value = "SIMBA"
    Set tRange = tWorkBook.WorkSheets(tSheetIndex).Range("$A$1:$B$25")
    Set tShape = tWorkBook.WorkSheets(tSheetIndex).Shapes.AddChart2(305, xlColumnStacked)
    tShape.Chart.SetSourceData Source:=tRange
    tShape.Chart.HasLegend = False
    For tIndex = 4 To 17
        With tShape.Chart.FullSeriesCollection(1).Points(tIndex).Format
            '.Format.Fill.Visible = msoTrue
            .Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent1
            '.ForeColor.TintAndShade = 0
            .Fill.ForeColor.Brightness = -0.25
            '.Transparency = 0
            '.Solid
            If tIndex = 10 Then
                '.Line.Visible = msoTrue
                '.Line.Weight = 2.5
                .Fill.ForeColor.RGB = RGB(255, 0, 0)
                '.Line.DashStyle = msoLineSysDot
                '.Line.Transparency = 0
            ElseIf tIndex = 13 Then
                .Fill.ForeColor.RGB = RGB(200, 50, 50)
            End If
        End With
    Next
    'tShape.CopyPicture
    tShape.Chart.Export tFullPath, "JPG"
    'With tWorkBook.Worksheets(tSheetIndex).ChartObjects.Add(tRange.Left, tRange.Top, tRange.Width, tRange.Height)
    '    uDeleteFile tFullPath
    '    .Activate
    '    .Chart.Paste
    '    .Chart.Export tFullPath, "JPG"
    'End With
    tShape.Delete
    'tWorkBook.Worksheets(tSheetIndex).ChartObjects(tWorkBook.Worksheets(tSheetIndex).ChartObjects.Count).Delete
    Set tRange = Nothing
    
    Set tOutMail = Outlook.Application.CreateItem(0)
    On Error Resume Next
        'tDateMark = "����� �� ����: " & inDate & vbCrLf & vbCrLf
        With tOutMail
            .To = "haustov@izhenergy.ru"
            .CC = ""
            .BCC = ""
            .Subject = "ROBOT: TEST"
            .Attachments.Add tFullPath, olByValue, 0

            'Then we add an html <img src=''> link to this image
            'Note than you can customize width and height - not mandatory
            .HTMLBody = "Hello"
            .HTMLBody = .HTMLBody & "<br><B>������</B><br><img src='cid:TEMP.jpg'><br><br>�����</font></span>"
            '.HTMLBody = "Test<br><img src=" & "'" & tFullPath & "'>"
            '.Body = "TEST"
            'You can add a file like this
            '.Attachments.Add tFullPath
            '.Send   'or use .Display
            .Display
        End With
    On Error GoTo 0
    Set tOutMail = Nothing
    
    Set tWorkBook = Nothing
    Set tExcel = Nothing
End Sub

' MANUAL SENDER
Public Sub fBRForecastDirect()
Dim tWorkDate, tLogTag, tGTPCode
    If Not fConfiguratorInit Then: Exit Sub
    If Not fXMLSmartUpdate("BASIS,CALENDAR,XSDBRFORECAST,DICTIONARY,BRFORECAST") Then: Exit Sub '-12 CALC DB
    If Not fLocalInit Then: Exit Sub
    tLogTag = "BRForecastDirect"
    tWorkDate = DateSerial("2021", "01", "13")
    tGTPCode = "PBELKA12"
    uADebugPrint tLogTag, "������ ������ �������� ��� ��� " & tGTPCode & " �� ���� " & Format(tWorkDate, "YYYY-MM-DD") & "..."
    fBRExchangeScan tWorkDate, tGTPCode, True, True, True
End Sub

Public Sub fBRForecastMain()
Dim tSOHours As TSOPeakHours
Dim tFilePath
    'uDebugPrint "BR Forecast Test"
    'If Not fConfiguratorInit Then: Exit Sub
    If Not fLocalInit Then: Exit Sub
    'If Not fXMLSmartUpdate("0,3,7,8,9") Then: Exit Sub
    'tFilePath = fGetReportFileByDate(DateSerial(2018, 7, 18))
    'uADebugPrint "TEST DOWNLOAD", "LOADED=" & tFilePath
    'tSOHours = fGetSOPeakHoursByZone(1, 2018, 1)
    'uADebugPrint "TEST SOGET", "LOADED=" & tSOHours.Loaded & " REASON=" & tSOHours.Reason
    'uADebugPrint "TEST", "ISWORKDAY=" & fIsWorkDay("2018", "07", "22")
    fBRExchangeScan
End Sub

Private Function fGetReportToList(inNode)
Dim tNode, tAddress, tEnabled
' // 00 - ���������������
    fGetReportToList = vbNullString
    If inNode Is Nothing Then: Exit Function
' // 00 - ���������������
    For Each tNode In inNode.ChildNodes
        If LCase(tNode.NodeName) = "reportto" Then
            tEnabled = tNode.GetAttribute("enabled")
            If tEnabled = "1" Then
                tAddress = tNode.GetAttribute("address")
                If Not IsNull(tAddress) Then
                    If InStr(tAddress, "@") > 0 Then 'simple check
                        If fGetReportToList = vbNullString Then
                            fGetReportToList = tAddress
                        Else
                            fGetReportToList = fGetReportToList & ";" & tAddress
                        End If
                    End If
                End If
            End If
        End If
    Next
End Function

Private Function fGetSubjectForecast(inFilePath, inSubjectID) 'UNSAFE CALL
Dim tXMLDoc, tXPathString, tIndex, tResultString, tNodes, tLogTag, tNode, tHour, tValue, tIsOK
Dim tValues(23) As Variant
    tLogTag = "GETSUBJECTFORE"
    fGetSubjectForecast = vbNullString
    tResultString = vbNullString
    Set tXMLDoc = CreateObject("Msxml2.DOMDocument.6.0")
    tXMLDoc.ASync = False
    tXMLDoc.Load inFilePath
    tXPathString = "//row[sub_rf_id='" & inSubjectID & "']"
    Set tNodes = tXMLDoc.SelectNodes(tXPathString)
    If tNodes.Length <> 24 Then
        uADebugPrint tLogTag, "�� ������� �������� ������ �������� ��� ������� <" & inSubjectID & ">; �������� �������� " & tNodes.Length & " �� 24!"
        Set tXMLDoc = Nothing
        Exit Function
    End If
    'clear
    For tIndex = 0 To 23
        tValues(tIndex) = -1
    Next
    'fill
    For Each tNode In tNodes
        tHour = fGetChildNodeByName(tNode, "hour").Text
        If IsNumeric(tHour) Then
            If tHour >= 0 And tHour <= 23 Then
                If tValues(tHour) = -1 Then
                    tValues(tHour) = Replace(fGetChildNodeByName(tNode, "cons_value").Text, ".", ",")
                    If Not IsNumeric(tValues(tHour)) Then: tValues(tHour) = -1
                Else
                    Exit For 'override?
                End If
            Else
                Exit For 'not hour
            End If
        Else
            Exit For 'not numeric hour
        End If
    Next
    'check and result
    tIsOK = True
    For tIndex = 0 To 23
        If tValues(tIndex) = -1 Then
            tIsOK = False
            uADebugPrint tLogTag, "�� ������� �������� ����� �������� ��� ������� <" & inSubjectID & "> �� ��� <" & tIndex & ">!"
            Exit For
        Else
            If tResultString = vbNullString Then
                tResultString = tValues(tIndex)
            Else
                tResultString = tResultString & ";" & tValues(tIndex)
            End If
        End If
    Next
    If tIsOK Then: fGetSubjectForecast = tResultString
    Set tXMLDoc = Nothing
End Function

'Private Function fFormReport(inFilePath, inSubjectID, inSubjectName, inTraderID, inGTPID, inZoneID, inTimeZoneUTC, inReportDate, inReportToDate, inWasChanged)
Private Function fFormReport(inFilePath, inSubjectInfo As TSubjectInfo, inTraderID, inGTPID, inReportDate, inReportToDate, inWasChanged)
Dim tXMLForecastReport, tYear, tMonth, tDay, tLogTag, tFixedSubjectID, tForecastValueString, tForecastElements, tIndex, tIsSorted, tTempValue, tHoursToShow, tHoursShowCount
Dim tFileName, tSubFolderPath, tFullPath, tRootFolder, tXMLReport, tCommentString, tNode, tChildNode, tDBNode, tDBMonthNode, tPairs, tPairsRead, tPeaks
Dim tSOHours As TSOPeakHours
Dim tForecastHours(23) As Variant
Dim tForecastMaxHours(23) As Variant
Dim tForecastMaxHoursIndex(23) As Variant
' // 00 - ���������������
    'Set inReportNode = Nothing
    fFormReport = False
    tLogTag = "FORMREPORT"
    inWasChanged = False
    tHoursToShow = 3
    tHoursShowCount = 0
' // 01 - ����������� ����, ���� � �������
    If Not IsDate(inReportDate) Then: Exit Function
    If Not IsDate(inReportToDate) Then: Exit Function
    If Not inSubjectInfo.IsReady Then: Exit Function
    tYear = Format(Year(inReportToDate), "0000")
    tMonth = Format(Month(inReportToDate), "00")
    tDay = Format(Day(inReportToDate), "00")
    tFixedSubjectID = CLng(inSubjectInfo.ParentID)
    uADebugPrint tLogTag, "REPORT �� ���� <" & tYear & "-" & tMonth & "-" & tDay & ">!"
' // 02 - ��������� ����� ��
    tSOHours = fGetDaySOPeakHoursByZone(inSubjectInfo.TradeZoneID, tYear, tMonth, tDay)
    If Not tSOHours.Loaded Then
        uADebugPrint tLogTag, "�� ������� �������� ���� �� ��� ���� <" & inSubjectInfo.TradeZoneID & "> �� ���� <" & tYear & "-" & tMonth & "-" & tDay & ">!"
        Exit Function
    End If
    'MAKE A MARK FOR HOLIDAY to DB
    If Not tSOHours.WorkDay Then
        uADebugPrint tLogTag, "���� <" & tYear & "-" & tMonth & "-" & tDay & "> �� �������� ������� ���, ������������� ������ ����������!"
        Set tDBNode = fGetRecordByDate(inTraderID, inGTPID, inReportToDate, inReportDate, True)
        If tDBNode Is Nothing Then
            uADebugPrint tLogTag, "�� ������� ������� ������ � ���� ������!"
            Exit Function
        End If
        tDBNode.SetAttribute "workday", 0
        fFormReport = True
        Exit Function
    End If
' // 03 - ��������� ��������� ����������� �� �������
    tForecastValueString = fGetSubjectForecast(inFilePath, tFixedSubjectID)
    If tForecastValueString = vbNullString Then
        uADebugPrint tLogTag, "�� ������� �������� ������ �������� ��� ������� <" & tFixedSubjectID & "> �� ���� <" & tYear & "-" & tMonth & "-" & tDay & ">!"
        Exit Function
    End If
    tForecastElements = Split(tForecastValueString, ";")
    If UBound(tForecastElements) <> 23 Then
        uADebugPrint tLogTag, "������ �������� ��� ������� <" & tFixedSubjectID & "> �� ���� <" & tYear & "-" & tMonth & "-" & tDay & "> ����� �������� ��� ������!"
        Exit Function
    End If
    For tIndex = 0 To 23
        tForecastHours(tIndex) = CDbl(tForecastElements(tIndex)) 'UNSAFE
    Next
' // 04 - ��������� ����� ��� �� �� �������
    For tIndex = 0 To 23
        tForecastMaxHoursIndex(tIndex) = tIndex 'shadow
        If tSOHours.Hours(tIndex) = 1 Then
            tForecastMaxHours(tIndex) = tForecastHours(tIndex)
        Else
            tForecastMaxHours(tIndex) = -1
        End If
    Next
' // 05 - ����������� � ��������� ����������
    tIsSorted = False
    While Not tIsSorted
        tIsSorted = True
        For tIndex = 0 To 22
            If tForecastMaxHours(tIndex) < tForecastMaxHours(tIndex + 1) Then
                'value sort
                tTempValue = tForecastMaxHours(tIndex)
                tForecastMaxHours(tIndex) = tForecastMaxHours(tIndex + 1)
                tForecastMaxHours(tIndex + 1) = tTempValue
                'index shadow moves
                tTempValue = tForecastMaxHoursIndex(tIndex)
                tForecastMaxHoursIndex(tIndex) = tForecastMaxHoursIndex(tIndex + 1)
                tForecastMaxHoursIndex(tIndex + 1) = tTempValue
                tIsSorted = False
            End If
        Next
    Wend
    'For tIndex = 0 To tHoursToShow - 1
    '    If tForecastMaxHours(tIndex) <> -1 Then
    '        tHoursShowCount = tHoursShowCount + 1
    '        'uADebugPrint tLogTag, "MaxHour#" & tIndex & " Hour #" & tForecastMaxHoursIndex(tIndex) + 1 & " - Value=" & tForecastMaxHours(tIndex)
    '    End If
    'Next
    'If tHoursShowCount = 0 Then
    '    uADebugPrint tLogTag, "��������������! ��� ������� ����� �� �� ���� <" & tYear & "-" & tMonth & "-" & tDay & ">!"
    'End If
' // 07 - ���������� ���� � ����� ��� ������
    tRootFolder = gForecastDataPath
    tFileName = inGTPID & "_FORECAST_" & Format(inReportToDate, "YYYYMMDD") & "_" & Format(inReportDate, "YYYYMMDD") & ".xml"
    tSubFolderPath = fSubFolderGet(tRootFolder, "Reports\" & Format(inReportToDate, "YYYY") & "\" & Format(inReportToDate, "MM") & "\" & inGTPID)
    If Not gFSO.FolderExists(tSubFolderPath) Then
        uADebugPrint tLogTag, "�� ������� ������� ����� ��� �������� ������!"
        Exit Function
    End If
    tFullPath = tSubFolderPath & "\" & tFileName
    uADebugPrint tLogTag, "tFullPath > " & tFullPath
    If Not uDeleteFile(tFullPath) Then
        uADebugPrint tLogTag, "�� ������� ������� ������ ���� ������ > " & tFileName
        Exit Function
    End If
' // 08 - ���������� ������ ������
    tCommentString = "������� ������� ��� Outlook " & cnstModuleName & " ������ " & cnstModuleVersion & " �� " & cnstModuleDate
    fBlankXMLForecastReportCreate tXMLReport, inReportToDate, inSubjectInfo.Name, inSubjectInfo.TradeZoneUTC, tCommentString
    If tXMLReport Is Nothing Then
        uADebugPrint tLogTag, "�� ������� ������� ������ ������!"
        Exit Function
    End If
' // 09 - ���������� ������ ������
    '<range from="04:00" to="08:00"/>
    tPairs = vbNullString
    If tSOHours.WorkDay Then
        Set tNode = fGetChildNodeByName(tXMLReport.DocumentElement, "planned-peak-hours")
        For tIndex = 0 To tSOHours.PairCount
            Set tChildNode = tNode.AppendChild(tXMLReport.CreateElement("range"))
            tChildNode.SetAttribute "from", fMakeHourFormat(tSOHours.PairPartA(tIndex) - 1, 0)
            tChildNode.SetAttribute "to", fMakeHourFormat(tSOHours.PairPartB(tIndex), 0)
            'record for DB
            If tPairs = vbNullString Then
                tPairs = tSOHours.PairPartA(tIndex) & ":" & tSOHours.PairPartB(tIndex)
            Else
                tPairs = tPairs & ";" & tSOHours.PairPartA(tIndex) & ":" & tSOHours.PairPartB(tIndex)
            End If
        Next
    End If
    '<hour from="16:00" to="17:00" power-order="1"/>
    tPeaks = vbNullString
    Set tNode = fGetChildNodeByName(tXMLReport.DocumentElement, "region-forecast-max-hours")
    For tIndex = 0 To tHoursToShow - 1
        If tForecastMaxHours(tIndex) <> -1 Then
            tHoursShowCount = tHoursShowCount + 1
            Set tChildNode = tNode.AppendChild(tXMLReport.CreateElement("hour"))
            tChildNode.SetAttribute "from", fMakeHourFormat(tForecastMaxHoursIndex(tIndex), 0)
            tChildNode.SetAttribute "to", fMakeHourFormat(tForecastMaxHoursIndex(tIndex) + 1, 0)
            tChildNode.SetAttribute "power-order", tIndex + 1
            If tPeaks = vbNullString Then
                tPeaks = tForecastMaxHoursIndex(tIndex) + 1
            Else
                tPeaks = tPeaks & ";" & tForecastMaxHoursIndex(tIndex) + 1
            End If
        End If
    Next
    fSaveXMLChanges tXMLReport, tFullPath, , True
' // 10 - �������� � ���� ������ ���������� �� ������
    Set tDBNode = fGetRecordByDate(inTraderID, inGTPID, inReportToDate, inReportDate, True)
    If tDBNode Is Nothing Then
        uADebugPrint tLogTag, "�� ������� ������� ������ � ���� ������!"
        Exit Function
    End If
    tDBNode.SetAttribute "workday", 1
    Set tChildNode = tDBNode.AppendChild(gBRForecastDB.XML.CreateElement("file"))
    tChildNode.SetAttribute "name", tFileName
    Set tChildNode = tDBNode.AppendChild(gBRForecastDB.XML.CreateElement("forecast"))
    tChildNode.SetAttribute "value", tForecastValueString
    Set tChildNode = tDBNode.AppendChild(gBRForecastDB.XML.CreateElement("peaks"))
    tChildNode.SetAttribute "value", tPeaks
    '��� �������� ����� �� ��������� �� � �������� ���� (����� ������� ����������?)
    If tPairs <> vbNullString Then
        Set tDBMonthNode = tDBNode.ParentNode.ParentNode
        tPairsRead = tDBMonthNode.GetAttribute("sopairs")
        If IsNull(tPairsRead) Or tPairsRead = vbNullString Then: tDBMonthNode.SetAttribute "sopairs", tPairs
    End If
' // 11 - �����
    fFormReport = True
End Function

Private Function fGetRecordByDate(inTraderID, inGTPID, inReportToDate, inReportDate, Optional inForceCreateRecord = False)
Dim tNode, tRoot, tYear, tMonth, tDay, tCreatedDate, tTraderID, tGTPID, tLogTag, tXPathString
    Set fGetRecordByDate = Nothing
    tLogTag = "BRFORECREATOR"
    If Not gBRForecastDB.Active Then: Exit Function 'preventer
    tYear = Format(Year(inReportToDate), "0000")
    tMonth = Format(Month(inReportToDate), "00")
    tDay = Format(Day(inReportToDate), "00")
    tCreatedDate = Format(inReportDate, "YYYYMMDD")
    tTraderID = UCase(inTraderID)
    tGTPID = UCase(inGTPID)
'Selector
    tXPathString = "//trader[@id='" & tTraderID & "']/gtp[@id='" & tGTPID & "']/year[@id='" & tYear & "']/month[@id='" & tMonth & "']/day[@id='" & tDay & "']/report[@created='" & tCreatedDate & "']"
    Set tNode = gBRForecastDB.XML.SelectSingleNode(tXPathString)
    If Not tNode Is Nothing Then
        Set fGetRecordByDate = tNode
        Exit Function
    End If
'Creator
    If Not inForceCreateRecord Then: Exit Function
    'TRADER
    Set tRoot = gBRForecastDB.XML.DocumentElement
    tXPathString = "//trader[@id='" & tTraderID & "']"
    Set tNode = gBRForecastDB.XML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        Set tNode = tRoot.AppendChild(gBRForecastDB.XML.CreateElement("trader"))
        tNode.SetAttribute "id", tTraderID
        'check
        Set tNode = gBRForecastDB.XML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            uADebugPrint tLogTag, "������ �������� � " & gBRForecastDB.ClassTag & " ����� <trader>; ���� - " & tXPathString
            Exit Function
        End If
    End If
    'GTP
    Set tRoot = tNode
    tXPathString = tXPathString & "/gtp[@id='" & tGTPID & "']"
    Set tNode = gBRForecastDB.XML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        Set tNode = tRoot.AppendChild(gBRForecastDB.XML.CreateElement("gtp"))
        tNode.SetAttribute "id", tGTPID
        'check
        Set tNode = gBRForecastDB.XML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            uADebugPrint tLogTag, "������ �������� � " & gBRForecastDB.ClassTag & " ����� <gtp>; ���� - " & tXPathString
            Exit Function
        End If
    End If
    'YEAR
    Set tRoot = tNode
    tXPathString = tXPathString & "/year[@id='" & tYear & "']"
    Set tNode = gBRForecastDB.XML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        Set tNode = tRoot.AppendChild(gBRForecastDB.XML.CreateElement("year"))
        tNode.SetAttribute "id", tYear
        'check
        Set tNode = gBRForecastDB.XML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            uADebugPrint tLogTag, "������ �������� � " & gBRForecastDB.ClassTag & " ����� <year>; ���� - " & tXPathString
            Exit Function
        End If
    End If
    'MONTH
    Set tRoot = tNode
    tXPathString = tXPathString & "/month[@id='" & tMonth & "']"
    Set tNode = gBRForecastDB.XML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        Set tNode = tRoot.AppendChild(gBRForecastDB.XML.CreateElement("month"))
        tNode.SetAttribute "id", tMonth
        'check
        Set tNode = gBRForecastDB.XML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            uADebugPrint tLogTag, "������ �������� � " & gBRForecastDB.ClassTag & " ����� <month>; ���� - " & tXPathString
            Exit Function
        End If
    End If
    'DAY
    Set tRoot = tNode
    tXPathString = tXPathString & "/day[@id='" & tDay & "']"
    Set tNode = gBRForecastDB.XML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        Set tNode = tRoot.AppendChild(gBRForecastDB.XML.CreateElement("day"))
        tNode.SetAttribute "id", tDay
        'check
        Set tNode = gBRForecastDB.XML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            uADebugPrint tLogTag, "������ �������� � " & gBRForecastDB.ClassTag & " ����� <day>; ���� - " & tXPathString
            Exit Function
        End If
    End If
    'REPORT
    Set tRoot = tNode
    tXPathString = tXPathString & "/report[@created='" & tCreatedDate & "']"
    Set tNode = gBRForecastDB.XML.SelectSingleNode(tXPathString)
    If tNode Is Nothing Then
        Set tNode = tRoot.AppendChild(gBRForecastDB.XML.CreateElement("report"))
        tNode.SetAttribute "created", tCreatedDate
        'check
        Set tNode = gBRForecastDB.XML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            uADebugPrint tLogTag, "������ �������� � " & gBRForecastDB.ClassTag & " ����� <report>; ���� - " & tXPathString
            Exit Function
        End If
    End If
'EXIT
    Set fGetRecordByDate = tNode
End Function

Private Function fMakeHourFormat(inHour, inMinutes)
    fMakeHourFormat = 0
    If Not IsNumeric(inHour) Then: Exit Function
    If Not IsNumeric(inMinutes) Then: Exit Function
    If inHour < 0 Or inHour > 23 Then: Exit Function
    If inMinutes < 0 Or inMinutes > 59 Then: Exit Function
    fMakeHourFormat = Format(inHour, "00") & ":" & Format(inMinutes, "00")
End Function

Private Function fBRExchangeItem(inNode, inLocalUTC, inDate, Optional inIgnoreTimeGate = False, Optional inIngnoreDateShift = False, Optional inSendOverride = False)
Dim tLogTag, tReportToString, tNode, tTempNode
Dim tGTPID, tSubjectID, tSubjectName, tTradeZoneID, tTimeZoneID, tTimeZoneUTC, tSubjectNode, tTargetUTC, tShiftHour, tCorrectDate, tOnTimeStart, tOnTimeEnd, tCurrentTime, tOnTimeTrigger
Dim tHour, tMinute, tTargetDate, tShift, tDepth, tFileIndex, tFileListElements, tDepthLim, tCurrentLocalDate, tChangeTrigger, tDBNode, tWasSent
Dim tFileList, tComment, tTraderID
Dim TSubjectInfo As TSubjectInfo
' // 00 - ���������������
    tLogTag = "BREXSCN"
    tDepthLim = 3
' // 01 - ������ ���������
    tReportToString = fGetReportToList(inNode)
    If tReportToString = vbNullString Then: Exit Function 'empty list
' // 02 - ������� ���� ��� � ��������� ���
    Set tNode = inNode.ParentNode.ParentNode
    If tNode Is Nothing Then
        uADebugPrint tLogTag, "�������������� ������ ��� ��������� ���� <gtp> ������� BASIS!"
        Exit Function
    End If
    tGTPID = tNode.GetAttribute("id")
    If IsNull(tGTPID) Then
        uADebugPrint tLogTag, "�������������� ������ ��� ��������� ��������� <id> ���� <gtp> ������� BASIS!"
        Exit Function
    End If
' // 03 - ������� ���� �������� ������� ����������� ��� � � ID
    Set tNode = tNode.ParentNode
    If tNode Is Nothing Then
        uADebugPrint tLogTag, "�������������� ������ ��� ��������� ���� <trader> ������� BASIS!"
        Exit Function
    End If
    tTraderID = tNode.GetAttribute("id")
    If IsNull(tTraderID) Then
        uADebugPrint tLogTag, "�������������� ������ ��� ��������� ��������� <id> ���� <trader> ������� BASIS!"
        Exit Function
    End If
' // 04 - ������� ��� �������� ���
    If Not fBasisGetGTPSettings(tGTPID, "subjectid", tSubjectID, tComment, tTraderID) Then
        uADebugPrint tLogTag, tComment
        Exit Function
    End If
' // 05 - ������� ������ ��������
    If Not fDictionaryGetSubjectInfo(tSubjectID, TSubjectInfo) Then
        uADebugPrint tLogTag, TSubjectInfo.Comment
        Exit Function
    End If
' // 06 - �������� ��������� ���� �� ������
    If TSubjectInfo.TradeMode = 0 Then '��������� �� ������������ ��������� ����
        uADebugPrint tLogTag, "������� ���� <" & TSubjectInfo.TradeZoneID & "> (������� <" & TSubjectInfo.ID & ":" & TSubjectInfo.Name & ">) �������� ��������� � �� ����� ���� ������������ ��� ��������!"
        Exit Function
    End If
' // 07 - ������� ������ ��������
    tOnTimeStart = inNode.GetAttribute("start")
    If Len(tOnTimeStart) <> 4 And Not IsNumeric(tOnTimeStart) Then
        uADebugPrint tLogTag, "�������� <start> ���� <gtp[@id='" & tGTPID & "']/exchange/item[@id='BRFORECAST']> �������� ���������� ��������! �������� ������ ����"
        Exit Function
    End If
    tHour = Left(tOnTimeStart, 2)
    tMinute = Right(tOnTimeStart, 2)
    If tHour > 23 Or tMinute > 59 Then
        uADebugPrint tLogTag, "�������� <start> ���� <gtp[@id='" & tGTPID & "']/exchange/item[@id='BRFORECAST']> �������� ���������� �������� [" & tOnTimeStart & "]! �������� ������ ����"
        Exit Function
    End If
' // 08 - ������� ����� ��������
    tOnTimeEnd = inNode.GetAttribute("end")
    If Len(tOnTimeEnd) <> 4 And Not IsNumeric(tOnTimeEnd) Then
        uADebugPrint tLogTag, "�������� <end> ���� <gtp[@id='" & tGTPID & "']/exchange/item[@id='BRFORECAST']> �������� ���������� ��������! �������� ������ ����"
        Exit Function
    End If
    tHour = Left(tOnTimeEnd, 2)
    tMinute = Right(tOnTimeEnd, 2)
    If tHour > 23 Or tMinute > 59 Then
        uADebugPrint tLogTag, "�������� <end> ���� <gtp[@id='" & tGTPID & "']/exchange/item[@id='BRFORECAST']> �������� ���������� �������� [" & tOnTimeEnd & "]! �������� ������ ����"
        Exit Function
    End If
    If tOnTimeStart > tOnTimeEnd Then
        uADebugPrint tLogTag, "�������� ������! �������� <end> ������ ��������� <start> ���� <gtp[@id='" & tGTPID & "']/exchange/item[@id='BRFORECAST']>!"
        Exit Function
    End If
' // 09 - ��������� ��������� ��������
    tCurrentLocalDate = Now()
    tShiftHour = TSubjectInfo.LocalUTC - inLocalUTC
    tCorrectDate = inDate + (1 / 24) * tShiftHour
    tCurrentTime = Format(Hour(tCorrectDate), "00") & Format(Minute(tCorrectDate), "00")
    tOnTimeTrigger = False
    ' //
    If inIgnoreTimeGate Then
        tOnTimeTrigger = True
        uADebugPrint tLogTag, "��������� ������ ������ ����� ���������������! [inIgnoreTimeGate = True]"
    Else
        If tCurrentTime >= tOnTimeStart And tCurrentTime <= tOnTimeEnd Then: tOnTimeTrigger = True
    End If
    
' // 10 - ��������� ��������� ������������ �������
    If tOnTimeTrigger Then
        
        If inIngnoreDateShift Then
            uADebugPrint tLogTag, "�������� <shift> ���� <gtp[@id='" & tGTPID & "']/exchange/item[@id='BRFORECAST']> ����� �������������� >> [tShift = 0]"
            tShift = 0
            tCurrentLocalDate = tCorrectDate
        Else
            tShift = inNode.GetAttribute("shift")
        End If
        
        If Not IsNumeric(tShift) Then
            uADebugPrint tLogTag, "�������� <shift> ���� <gtp[@id='" & tGTPID & "']/exchange/item[@id='BRFORECAST']> �������� ���������� ��������! ��������� 0-" & tDepthLim
            Exit Function
        End If
        tShift = CInt(tShift)
        If tShift < 0 Or tShift > tDepthLim Then
            uADebugPrint tLogTag, "�������� <shift> ���� <gtp[@id='" & tGTPID & "']/exchange/item[@id='BRFORECAST']> �������� ���������� �������� [" & tShift & "]! ��������� 0-" & tDepthLim
            Exit Function
        End If
        tShift = tShift + Fix(tCorrectDate) - Fix(tCurrentLocalDate)
        If tShift < 0 Or tShift > tDepthLim Then
            uADebugPrint tLogTag, "�������� <shift> ���� <gtp[@id='" & tGTPID & "']/exchange/item[@id='BRFORECAST']> ����� ��������� �������� ���������� �������� [" & tShift & "]! ��������� 0-" & tDepthLim
            Exit Function
        End If
        tDepth = inNode.GetAttribute("depth")
        If Not IsNumeric(tDepth) Then
            uADebugPrint tLogTag, "�������� <depth> ���� <gtp[@id='" & tGTPID & "']/exchange/item[@id='BRFORECAST']> �������� ���������� ��������! ��������� 1-" & tDepthLim
            Exit Function
        End If
        tDepth = CInt(tDepth)
        If tDepth < 1 Or tDepth > tDepthLim Then
            uADebugPrint tLogTag, "�������� <depth> ���� <gtp[@id='" & tGTPID & "']/exchange/item[@id='BRFORECAST']> �������� ���������� �������� [" & tDepth & "]! ��������� 1-" & tDepthLim
            Exit Function
        End If
' // 09 - ���������� � ����������� ����������� ������
        fBRForecastDownloader tFileList, tCurrentLocalDate, tDepthLim
        tFileListElements = Split(tFileList, ";")
        If UBound(tFileListElements) >= 0 Then
            For tFileIndex = tShift To tShift + tDepth - 1
                If tFileIndex <= tDepthLim Then
                    If tFileListElements(tFileIndex) <> vbNullString Then
                        '01 \\ ������� ������ � ���� ������ tSubjectInfo
                        Set tDBNode = fGetRecordByDate(tTraderID, tGTPID, tCurrentLocalDate + tFileIndex, tCorrectDate)
                        '02 \\ ���� ������ ��� - ��������� ����� � ������ ������
                        If tDBNode Is Nothing Then
                            If fFormReport(tFileListElements(tFileIndex), TSubjectInfo, tTraderID, tGTPID, tCorrectDate, tCurrentLocalDate + tFileIndex, tChangeTrigger) Then
                                fSaveXMLDB gBRForecastDB, False 'ACCEPT CHANGES
                                Set tDBNode = fGetRecordByDate(tTraderID, tGTPID, tCurrentLocalDate + tFileIndex, tCorrectDate) '������� ����� ������
                            Else
                                If tChangeTrigger Then: fReloadXMLDB gBRForecastDB, False 'DISCARD CHANGES
                            End If
                        End If
                        '03 \\ ���� � ����� ������ ����������
                        tWasSent = fBRForecastSender(tDBNode, tReportToString, inSendOverride)
                    End If
                End If
            Next
        Else
            uADebugPrint tLogTag, "�� ������� �������� ������! ������ ����������."
        End If
        uADebugPrint tLogTag, "TARGETUTC=" & TSubjectInfo.LocalUTC & " \ GTP=" & tGTPID & " \ SUBJECT=" & TSubjectInfo.ID & " \ DATE=" & tCorrectDate & " \ START=" & tOnTimeStart & " \ END=" & tOnTimeEnd & " \ TRIGGER=" & tOnTimeTrigger & " \ SHIFT=" & tShift & " \ DEPTH=" & tDepth & " \ SENT=" & tWasSent
        'uADebugPrint tLogTag, "FILELIST=" & tFileList
    End If
End Function

Private Function fMailListAdjustBasic(inMailString, inSentString)
Dim tFinalString, tMailElements, tSentElements, tMailAddress, tPosA, tSentAddress, tPosB, tMailElement, tSentElement, tAlreadySent
    fMailListAdjustBasic = vbNullString
    If IsNull(inSentString) Then: inSentString = vbNullString
    inMailString = LCase(inMailString)
    inSentString = LCase(inSentString)
    tMailElements = Split(inMailString, ";")
    tSentElements = Split(inSentString, ";")
    If UBound(tSentElements) < 0 Then
        fMailListAdjustBasic = inMailString
        Exit Function 'nothing to adjust - all items should be sent
    End If
    If UBound(tMailElements) < 0 Then: Exit Function 'nothing to adjust - no items on input
    tFinalString = vbNullString
    For Each tMailElement In tMailElements
        tMailAddress = tMailElement
        tAlreadySent = False
        For Each tSentElement In tSentElements
            tSentAddress = tSentElement
            If tSentAddress = tMailAddress Then
                tAlreadySent = True
                Exit For
            End If
        Next
        If Not tAlreadySent Then
            uAddToList tFinalString, tMailElement
        End If
    Next
    fMailListAdjustBasic = tFinalString
End Function

Private Function fBRForecastSender(inDBNode, inReportToString, Optional inSendOverride = False)
Dim tWorkDay, tNode, tMailNode, tXMLWasChange, tFileName, tGTPID, tYear, tMonth, tDay, tLogTag, tMailToString, tSentString, tCurrentMailToString, tPairs, tSubFolderPath, tRootFolder, tFullPath, tMailToElements, tMailToElement
Dim tPeaks, tForecastValueString, tSubjectID, tSubjectName, tXPathString, tWasSent, tTradeZoneID, tTimeZoneID, tBaseUTC, tLocalUTC
Dim tComment
Dim TSubjectInfo As TSubjectInfo
    fBRForecastSender = False
    If inDBNode Is Nothing Then: Exit Function
    If inReportToString = vbNullString Then: Exit Function
    tLogTag = "BRFORESENDER"
    
    tWorkDay = inDBNode.GetAttribute("workday")
    If tWorkDay <> "1" Or tWorkDay <> 1 Then: Exit Function
    
    tXMLWasChange = False
    Set tMailNode = fGetChildNodeByName(inDBNode, "mail")
    If tMailNode Is Nothing Then
        tXMLWasChange = True
        Set tMailNode = fGetChildNodeByName(inDBNode, "mail", True)
        If tMailNode Is Nothing Then
            fReloadXMLDB gBRForecastDB, False 'is it really needed?
            uADebugPrint tLogTag, "�� ������� ������� ������ <mail>!"
            Exit Function
        End If
    End If
    
    tCurrentMailToString = tMailNode.GetAttribute("mailto")
    If IsNull(tCurrentMailToString) Then: tCurrentMailToString = vbNullString
    
    tSentString = tMailNode.GetAttribute("sent")
    If IsNull(tSentString) Then: tSentString = vbNullString
    
    tMailToString = inReportToString
    If Not inSendOverride Then
        tMailToString = fMailListAdjustBasic(tMailToString, tSentString) '������� �� ������ �������� ��� ���������
    Else
        tXMLWasChange = True
        tSentString = vbNullString
        tMailNode.SetAttribute "sent", vbNullString 'send override
    End If
    
    If tCurrentMailToString <> tMailToString Then '���� ������ �������� ��������� - ���� ��� ������������� � ��
        tXMLWasChange = True
        tMailNode.SetAttribute "mailto", tMailToString
    End If
    '���� ������ ������ �������� �� ��������� ������ - �����
    If tMailToString = vbNullString Then
        If tXMLWasChange Then: fSaveXMLDB gBRForecastDB, False
        Exit Function
    End If
    '���� �������� ���� ������ �������� �� ����
    '������ ���������� �� ������ - �.�. ���� �� ?���� ���� �� ������? � �.�.
    Set tNode = inDBNode.ParentNode
    tDay = tNode.GetAttribute("id")
    Set tNode = tNode.ParentNode
    tMonth = tNode.GetAttribute("id")
    tPairs = tNode.GetAttribute("sopairs")
    Set tNode = tNode.ParentNode
    tYear = tNode.GetAttribute("id")
    Set tNode = tNode.ParentNode
    tGTPID = tNode.GetAttribute("id")
    Set tNode = fGetChildNodeByName(inDBNode, "file")
    tFileName = tNode.GetAttribute("name")
    tRootFolder = gForecastDataPath
    tSubFolderPath = tRootFolder & "\Reports\" & tYear & "\" & tMonth & "\" & tGTPID
    tFullPath = tSubFolderPath & "\" & tFileName
    If Not gFSO.FileExists(tFullPath) Then
        uADebugPrint tLogTag, "�� ������� ���������� ���� ������ <" & tFileName & "> �� ����: " & tSubFolderPath
        If tXMLWasChange Then: fReloadXMLDB gBRForecastDB, False 'DISCARD CHANGES
        Exit Function
    End If
    Set tNode = fGetChildNodeByName(inDBNode, "peaks")
    tPeaks = tNode.GetAttribute("value")
    Set tNode = fGetChildNodeByName(inDBNode, "forecast")
    tForecastValueString = tNode.GetAttribute("value")
    '� ������ ������ ����
    If fBasisGetGTPSettings(tGTPID, "subjectid", tSubjectID, tComment) Then
        If Not fDictionaryGetSubjectInfo(tSubjectID, TSubjectInfo) Then
            uADebugPrint tLogTag, TSubjectInfo.Comment
        End If
    Else
        uADebugPrint tLogTag, tComment
    End If
    '���� ����, ������ ����� ��� ��������� �� ������
    tWasSent = False
    'tMailToElements = Split(tMailToString, ";")
    'For Each tMailToElement In tMailToElements
    If fForecastReportSend(tMailToString, tFullPath, tYear, tMonth, tDay, tGTPID, tPeaks, tPairs, tForecastValueString, TSubjectInfo) Then
        uAddToList tSentString, tMailToString
        tMailToString = vbNullString
        'tMailToString = fMailListAdjustBasic(tMailToString, tSentString)
        tMailNode.SetAttribute "mailto", tMailToString
        tMailNode.SetAttribute "sent", tSentString
        tXMLWasChange = True
        tWasSent = True
    End If
    'Next
    'drop changes to XML
    If tXMLWasChange Then: fSaveXMLDB gBRForecastDB
    fBRForecastSender = tWasSent
    'tXMLWasChange = False
    'Set tMailNode = fGetChildNodeByName(inDBNode, "mail")
End Function

Private Function fForecastReportSend(inAddressList, inAttachmentPath, inYear, inMonth, inDay, inGTPID, inPeaks, inPairs, inForecastValueString, inSubjectInfo As TSubjectInfo)
Dim tLogTag, tHeader, tAutoSign, tPicturePath, tPictureCode, tCIDCode
Dim tPeakShifted, tPeak, tPeakArray, tImageAttachment
Dim tOutMail As Outlook.MailItem
Dim tPAccessor As Outlook.PropertyAccessor
Const PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

    tLogTag = "BRFORESEND"
    fForecastReportSend = True
    
    ' IMAGE-GRAPH ADD
    tPicturePath = fGetChartPicture(inForecastValueString, inPairs, inPeaks)
    If tPicturePath <> vbNullString Then
        tCIDCode = uGetFileName(tPicturePath)
        tPictureCode = "<img src=""cid:" & tCIDCode & """>"
    Else
        uADebugPrint tLogTag, "��������� ������! �� ������� ������������ ������!"
        fForecastReportSend = False
        Exit Function
    End If
    
    ' CREATE NEW MAILITEM
    'Set tOutMail = Nothing
    Set tOutMail = Outlook.Application.CreateItem(olMailItem) 'Outlook.Application.CreateItem(0)
    'tHeader = "������� ����� �� " & inYear & "-" & inMonth & "-" & inDay & " �� �������� " & inSubjectInfo.Name & " (" & inGTPID & ")"
    tHeader = "������� ����� �� �������� " & inSubjectInfo.Name & " (" & inGTPID & ")"
    On Error Resume Next
        tAutoSign = vbCrLf & vbCrLf & "// ������ ��������� ������������ ������������� � ����� ������������� �������������� ��������"
        
        'timeshifting
        tPeakShifted = vbNullString
        tPeakArray = Split(inPeaks, ";")
        For Each tPeak In tPeakArray
            If tPeakShifted = vbNullString Then
                tPeakShifted = tPeak - inSubjectInfo.TradeZoneUTC + inSubjectInfo.LocalUTC
            Else
                tPeakShifted = tPeakShifted & ";" & tPeak - inSubjectInfo.TradeZoneUTC + inSubjectInfo.LocalUTC
            End If
        Next
        
        'internal
        With tOutMail
            .SendUsingAccount = gMainAccount
            .To = inAddressList
            .CC = ""
            .BCC = ""
            .Subject = "ROBOT: " & tHeader
            
            'attachments
            If tPicturePath <> vbNullString Then
                .Attachments.Add (tPicturePath)
                Set tPAccessor = .Attachments.Item(.Attachments.Count).PropertyAccessor
                tPAccessor.SetProperty PR_ATTACH_CONTENT_ID, tCIDCode
            End If
            .Attachments.Add inAttachmentPath
            
            'html body
            .HTMLBody = "������ ���������� ����� ��� �� <B>" & inYear & "-" & inMonth & "-" & inDay & "</B> �� �������� " & inSubjectInfo.Name & " (" & inGTPID & ")<br><br>"
            '.HTMLBody = .HTMLBody & "������� ���� ������: <B>" & fGetUTCForm(inSubjectInfo.TradeZoneUTC) & "</B><br>"
            .HTMLBody = .HTMLBody & "������� ���� ��������: <B>" & fGetUTCForm(inSubjectInfo.LocalUTC) & "</B><br>"
            '.HTMLBody = .HTMLBody & "��������� ���� ��� (�����): <B>" & Replace(inPeaks, ";", ", ") & "</B><br>"
            .HTMLBody = .HTMLBody & "��������� ���� ��� (�������): <B>" & Replace(tPeakShifted, ";", ", ") & "</B><br><br>"
            .HTMLBody = .HTMLBody & tPictureCode & "<br><br>" & tAutoSign & "</font></span>"
            .Send   'or use .Display
        End With
        
        'err check
        If Err.Number <> 0 Then
            fForecastReportSend = False
            uADebugPrint tLogTag, "ERROR > " & Err.Description
        End If
    On Error GoTo 0
    
    Set tOutMail = Nothing
    Set tPAccessor = Nothing
    uADebugPrint tLogTag, "SENDING from <" & gMainAccount & "> to <" & inAddressList & ">! RESULT = " & fForecastReportSend
End Function

Private Function fGetChartPicture(inForecastValueString, inPairs, inPeaks)
Dim tWorkBook, tIndex, tSubIndex, tSheetIndex, tFullPath, tForecastElements, tPeakElements, tRange, tShape, tPairsElements, tSubPairElements, tLogTag
Dim tSOColorRGB, tPeakColorRGB, tMainPeakColorRGB, tIsMainPeakPass
    fGetChartPicture = vbNullString
    tLogTag = "GETCHARTPIC"
    tSheetIndex = 1
    tForecastElements = Split(inForecastValueString, ";")
    tPeakElements = Split(inPeaks, ";")
    tPairsElements = Split(inPairs, ";")
    If UBound(tForecastElements) <> 23 Then
        uCDebugPrint tLogTag, 2, "������ ������! �� ����� �� 24 ��������, � - " & UBound(tForecastElements) <> 23
        Exit Function
    End If
    'Debug.Print "1: " & GetTickCount
    If gExcel Is Nothing Then
        uCDebugPrint tLogTag, 2, "������ EXCEL �������� �� ��������! ������������ ������� ����������!"
        Exit Function
    End If
    'Debug.Print "2: " & GetTickCount
    On Error Resume Next
        Set tWorkBook = gExcel.Workbooks.Add
        If Err.Number <> 0 Then
            uCDebugPrint tLogTag, 2, "������ EXCEL �� ���� ������� WORKBOOK! ������������ ������� ����������!"
            Exit Function
        End If
    'On Error GoTo 0
    'Debug.Print "3: " & GetTickCount
    'disable controls
        fExcelControl -1, -1, -1, -1
        'working
        tFullPath = gForecastDataPath & "\TEMP.jpg"
        'drop forecast values
        tWorkBook.WorkSheets(tSheetIndex).Activate
        For tIndex = 2 To 25
            tWorkBook.WorkSheets(tSheetIndex).Cells(tIndex, 1).Value = tIndex - 1
            tWorkBook.WorkSheets(tSheetIndex).Cells(tIndex, 2).Value = CDbl(tForecastElements(tIndex - 2))
        Next
        If Err.Number <> 0 Then
            uCDebugPrint tLogTag, 2, "�� ������� ������ ������ �������� �� ���� EXCEL! ��������: " & Err.Description
            Exit Function
        End If
        'Debug.Print "4: " & GetTickCount
        'create a chart
        Set tRange = tWorkBook.WorkSheets(tSheetIndex).Range("$A$1:$B$25")
        Set tShape = tWorkBook.WorkSheets(tSheetIndex).Shapes.AddChart2(305, xlColumnStacked)
        tShape.Chart.SetSourceData Source:=tRange
        tShape.Chart.HasLegend = False
        If Err.Number <> 0 Then
            uCDebugPrint tLogTag, 2, "�� ������� ������� ������ �� ����� EXCEL! ��������: " & Err.Description
            Exit Function
        End If
        'Debug.Print "5: " & GetTickCount
        tSOColorRGB = RGB(0, 0, 255)
        tPeakColorRGB = RGB(255, 0, 0)
        tMainPeakColorRGB = RGB(150, 0, 0)
        'fill with so hours
        With tShape.Chart.FullSeriesCollection(1)
            For tIndex = 0 To UBound(tPairsElements)
                tSubPairElements = Split(tPairsElements(tIndex), ":")
                For tSubIndex = CInt(tSubPairElements(0)) To CInt(tSubPairElements(1))
                    'With tShape.Chart.FullSeriesCollection(1).Points(CInt(tSubIndex)).Format.Fill
                    '    .ForeColor.RGB = tSOColorRGB '.ObjectThemeColor = msoThemeColorAccent1
                    '    .ForeColor.Brightness = -0.25
                    'End With
                    .Points(CInt(tSubIndex)).Format.Fill.ForeColor.RGB = tSOColorRGB
                    .Points(CInt(tSubIndex)).Format.Fill.ForeColor.Brightness = -0.25
                Next
            Next
            'Debug.Print "6: " & GetTickCount
            'fill with so peak hours
            tIsMainPeakPass = False
            For Each tIndex In tPeakElements
                'With tShape.Chart.FullSeriesCollection(1).Points(CInt(tIndex)).Format
                '    .Fill.ForeColor.RGB = tPeakColorRGB
                '    .Fill.ForeColor.Brightness = -0.25
                'End With
                If Not tIsMainPeakPass Then
                    .Points(CInt(tIndex)).Format.Fill.ForeColor.RGB = tMainPeakColorRGB
                    .Points(CInt(tIndex)).Format.Fill.ForeColor.Brightness = -0.25
                    tIsMainPeakPass = True
                Else
                    .Points(CInt(tIndex)).Format.Fill.ForeColor.RGB = tPeakColorRGB
                    .Points(CInt(tIndex)).Format.Fill.ForeColor.Brightness = -0.25
                End If
            Next
        End With
        If Err.Number <> 0 Then
            uCDebugPrint tLogTag, 2, "�� ������� ������� ����� �� ������ ����� EXCEL! ��������: " & Err.Description
            Exit Function
        End If
        'restore controls
        fExcelControl 1, 1, 1, 1
        'Debug.Print "7: " & GetTickCount
        'get a picture
        tShape.Chart.Export tFullPath, "JPG"
        If Err.Number <> 0 Then
            uCDebugPrint tLogTag, 2, "�� �������������� ������ ����� EXCEL � JPG! ��������: " & Err.Description
            Exit Function
        End If
        'kill objects
        'tShape.Delete '���� ����� �� ����� ������� ��� ����������, ����� ������� ������ � �����?
        Set tRange = Nothing
        Set tShape = Nothing
        'Debug.Print "8: " & GetTickCount
        'close book and kill it
        tWorkBook.Close SaveChanges:=False
        'tExcel.Quit '�������, ����� ����� ������� ��� ��� ������? ���� gFSO � �� �������� ������ ���
        Set tWorkBook = Nothing
        'Debug.Print "9: " & GetTickCount
        'Set tExcel = Nothing
        'result
        If gFSO.FileExists(tFullPath) Then: fGetChartPicture = tFullPath
End Function

Private Sub fBRForecastDownloader(inFileList, inCurrentLocalDate, inDepthLimit)
Dim tDayIndex, tYear, tMonth, tDay, tDeepFolder, tNewFilePath, tFilePath, tFileName, tFullPath, tReasonString, tLogTag, tIndex, tDepth
    inFileList = vbNullString
    'tCurrentLocalDate = Now()
    tDepth = inDepthLimit
    tIndex = 0
    tLogTag = "BRFOREDLDER"
    For tDayIndex = inCurrentLocalDate To inCurrentLocalDate + tDepth
        tYear = Format(Year(tDayIndex), "0000")
        tMonth = Format(Month(tDayIndex), "00")
        tDay = Format(Day(tDayIndex), "00")
        'folders
        tFilePath = gStoragePath & "\" & tYear & "\" & tMonth
        tFileName = "BRForecast_" & tYear & tMonth & tDay & "_" & Format(inCurrentLocalDate, "YYYYMMDD") & ".xml"
        tFullPath = tFilePath & "\" & tFileName
        If Not gFSO.FileExists(tFullPath) Then
            tFullPath = vbNullString
            tNewFilePath = fGetReportFileByDate(tDayIndex)
            '���� �� ������ � ����� ��������
            If tNewFilePath <> vbNullString Then
                tDeepFolder = gStoragePath & "\" & tYear
                If uFolderCreate(tDeepFolder) Then
                    tDeepFolder = tDeepFolder & "\" & tMonth
                    If uFolderCreate(tDeepFolder) Then
                        'file
                        tDeepFolder = tDeepFolder & "\" & tFileName
                        On Error Resume Next
                            gFSO.MoveFile tNewFilePath, tDeepFolder
                        On Error GoTo 0
                        If gFSO.FileExists(tDeepFolder) Then: tFullPath = tDeepFolder
                    End If
                End If
            End If
        End If
        'schema check
        If tFullPath <> vbNullString Then
            If Not fBRForecastReportCheck(tFullPath, tReasonString, True) Then
                uADebugPrint tLogTag, "��������� �������� �����-������ <" & tFileName & ">!"
                tFullPath = vbNullString
            End If
        Else
            uADebugPrint tLogTag, "�� ������� �������� ����-����� <" & tFileName & ">!"
        End If
        'result former
        If tFullPath <> vbNullString Then
            inFileList = inFileList & tFullPath
            If tIndex < tDepth Then
                 inFileList = inFileList & ";"
            End If
        End If
        tIndex = tIndex + 1
    Next
End Sub

Private Sub fBRExchangeScan(Optional inWorkDate = 0, Optional inGTPCode = vbNullString, Optional inIgnoreTimeGate = False, Optional inIngnoreDateShift = False, Optional inSendOverride = False)
Dim tXPathString, tNode, tNodes, tLogTag, tCurrentDate, tGTPCodeElement
' // 00 - ���������������
    tLogTag = "BREXSCN"
    If Not gXMLBasis.Active Then: Exit Sub
' // 00 - ���������������
    tGTPCodeElement = "gtp"
    If inGTPCode <> vbNullString Then: tGTPCodeElement = tGTPCodeElement & "[@id='" & UCase(inGTPCode) & "']"
    tXPathString = "//trader[@id='" & gTraderInfo.ID & "']/" & tGTPCodeElement & "/exchange/item[(@id='BRFORECAST' and @enabled='1')]"
    Set tNodes = gXMLBasis.XML.SelectNodes(tXPathString)
    'uADebugPrint tLogTag, "NODES=" & tNodes.Length
    If tNodes.Length = 0 Then: Exit Sub '��� ��������� ��� ������
' // 00 - ���������������
    If inWorkDate = 0 Then
        tCurrentDate = Now
    Else
        tCurrentDate = inWorkDate
    End If
' // 00 - ���������������
    For Each tNode In tNodes
        fBRExchangeItem tNode, gLocalUTC, tCurrentDate, inIgnoreTimeGate, inIngnoreDateShift, inSendOverride
    Next
' // 00 - ���������������
' // 00 - ���������������
End Sub

Private Function fBRForecastReportCheck(inReportPath, inReasonString, Optional inKillOnFail = False)
Dim tXMLDoc, tResult
    fBRForecastReportCheck = False
    inReasonString = vbNullString
    Set tXMLDoc = CreateObject("Msxml2.DOMDocument.6.0")
    tXMLDoc.ASync = False
    tXMLDoc.Load inReportPath
    Set tXMLDoc.Schemas = gXSDForecast.XML
    Set tResult = tXMLDoc.Validate()
    If tResult <> 0 Then
        inReasonString = tResult.Reason
        If inKillOnFail Then
            If Not (uDeleteFile(inReportPath)) Then: uADebugPrint "BRFOREXSD", "�� ������� ������� ������������ ����-�����!"
        End If
        Exit Function
    End If
    fBRForecastReportCheck = True
End Function

Private Function fGetReportFileByDate(inDate)
Dim tLogTag, tYear, tMonth, tDay, tLinkDatePart, tLinkString, tResultPath, tFileName, tReasonString
' // 00 - ���������������
    fGetReportFileByDate = vbNullString
    tLogTag = "BRGETREPORT"
' // 01 - ���� �� �� �����
    If Not IsDate(inDate) Then
        uADebugPrint tLogTag, "�������� ������ ����!"
        Exit Function
    End If
' // 02 - ������ ������ ������ �� ����
    'tLinkString = "http://br.so-ups.ru/Public/Export/Xml/ForecastConsumSubRf.aspx?"
    tYear = Format(Year(inDate), "0000")
    tMonth = Format(Month(inDate), "00")
    tDay = Format(Day(inDate), "00")
    'tLinkDatePart = "&date=" & tDay & "." & tMonth & "." & tYear
    'tLinkString = tLinkString & tLinkDatePart
' // 03 - ������� ���� �� ��������� ������
    tFileName = "BRForecast_" & tYear & tMonth & tDay & "_" & Format(Now(), "YYYYMMDD") & ".xml"
    'tResultPath = fDownloadFile(tLinkString, gDownloadPath, tFileName)
    tResultPath = fDownloadFileByAPI(inDate, gDownloadPath, tFileName)
    If tResultPath = vbNullString Then
        uADebugPrint tLogTag, "�� ������� ������� ����-�����!"
        Exit Function
    End If
' // 04 - �������� ��������� ����� �� XSD �����
    If Not fBRForecastReportCheck(tResultPath, tReasonString, True) Then
        uADebugPrint tLogTag, "������! ����-����� <" & tFileName & "> � ���������� ���������: " & tReasonString
        uADebugPrint tLogTag, "�� ������� ������� ����-�����!"
        Exit Function
    End If
' // 05 - �� ������
    uADebugPrint tLogTag, "������� �������� ����-����� <" & tFileName & ">"
    fGetReportFileByDate = tResultPath
End Function

Private Function fDownloadFileByAPI(inDate, inDropFolder, inFileName)
Dim tLogTag, tFileFullPath, tStream, tFileWasDownloaded, tHTTP, tServiceURL, tServiceAction, tSOAP, tServiceActionRepsonse, tNode, tXML
'00 // ���������������
    tLogTag = fGetLogTag("DWNLDAPI")
    fDownloadFileByAPI = vbNullString
    
'00 // �������� ������� ����� �������� ��� ����������� �����
    If Not (gFSO.FolderExists(inDropFolder)) Then
        uCDebugPrint tLogTag, 2, "����� ��� ���������� ����� �� ����������!"
        Exit Function
    End If
    
'00 // ����� ��� ������������� ����� � ����� �� ������ � ����� �������� (������� ������� ���� ��� ��� �������)
    tFileFullPath = inDropFolder & "\" & inFileName
    If gFSO.FileExists(tFileFullPath) Then
        If Not (uDeleteFile(tFileFullPath)) Then
            uCDebugPrint tLogTag, 2, "�� ������� ������� ������ ���� ������ �������� �������������� ������� �����! ��� ����� - <" & inFileName & ">"
            Exit Function
        End If
    End If
    
'00 // ���������� ������ ��� ������������ SOAP ������� �� ������ WCF
'      http://br.so-ups.ru:8090/PublicApi/PublicApiService.svc?wsdl - ��� �������� ����� ����� ����
'      https://br.so-ups.ru/Public/Docs/DocView?id=710117d7-501c-4f53-924a-62fae5ed95c1&path=DocList&month&year=2019&intension&doc=%5Bobject%20Object%5D - �������� ����� � API
    Set tHTTP = CreateObject("MSXML2.XMLHTTP")
    Set tXML = CreateObject("MSXML2.DOMDocument.6.0")
    tServiceURL = "http://br.so-ups.ru:8090/PublicApi/PublicApiService.svc"
    tServiceAction = "GetVsvgoConsumingSubRfData"
    tServiceActionRepsonse = "GetVsvgoConsumingSubRfDataResult"
        
'00 // ��������� SOAP ���������
    tSOAP = "<?xml version=""1.0"" encoding=""utf-8""?>"
    tSOAP = tSOAP & "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">"
    tSOAP = tSOAP & "  <soap:Body>"
    tSOAP = tSOAP & "    <" & tServiceAction & " xmlns=""http://www.armd.ru/soft/dssi/SOEES/SBR/Web/Api/PublicApi"">"
    'tSOAP = tSOAP & "      <date>2019-03-01T00:00:00</date>"
    tSOAP = tSOAP & "      <date>" & Format(inDate, "YYYY-MM-DD") & "</date>"
    tSOAP = tSOAP & "      <returnType>Xml</returnType>"
    'tSOAP = tSOAP & "      <strCurrency>USD</strCurrency>"
    'tSOAP = tSOAP & "      <intRank>1</intRank>"
    tSOAP = tSOAP & "    </" & tServiceAction & ">"
    tSOAP = tSOAP & "  </soap:Body>"
    tSOAP = tSOAP & "</soap:Envelope>"
    
    On Error Resume Next
        tFileWasDownloaded = False
        
        With tHTTP
        
        '00 // ������������� ����������
            .Open "POST", tServiceURL, False
            .SetRequestHeader "Content-Type", "text/xml; charset=utf-8"
            .SetRequestHeader "soapAction", "http://www.armd.ru/soft/dssi/SOEES/SBR/Web/Api/PublicApi/IPublicApiService/" & tServiceAction
            
        '00 // �������� SOAP �� ����������
            .Send tSOAP
            
        '00 // ��������� ������ �������
            If .Status = 200 Then
            
                '��������� ����� � ������ XML (� ��������� ���� ���)
                tXML.SetProperty "SelectionNamespaces", "xmlns:x='http://schemas.xmlsoap.org/soap/envelope/' " & "xmlns:m='http://tempuri.org/'" & " xmlns:n='http://www.w3.org/2001/XMLSchema-instance'" & " xmlns:so='http://www.armd.ru/soft/dssi/SOEES/SBR/Web/Api/PublicApi'"
                tXML.LoadXML .ResponseText
                
                '������� ������� ���� � �������
                Set tNode = tXML.SelectSingleNode("//so:" & tServiceActionRepsonse)
                
                '���� ���� ��� - �� ���-�� �� ���
                If tNode Is Nothing Then
                    uCDebugPrint tLogTag, 2, "�� ������� �������� ����� �� �������� API - <" & tServiceAction & ">!"
                Else
                    '���� ���� ������� - �������� � ���������� � ��������� ����
                    tXML.LoadXML tNode.Text
                    tXML.Save tFileFullPath
                    tFileWasDownloaded = True
                End If
            '���� �� ������� ������ ��������
            Else
                uCDebugPrint tLogTag, 2, "�� ������� �������� ����� �� �������� API - <" & tServiceAction & ">, ������ HTTP �������� �� �������� <" & .Status & ">!"
            End If
        End With
        
        '��������� ������
        If Err.Number <> 0 Then
            uCDebugPrint tLogTag, 2, "�������������� ������ ��� ���������� �����: " & Err.Description
            Set tHTTP = Nothing
            Set tXML = Nothing
            Exit Function
        End If
        
        '������� �������
        Set tXML = Nothing
    On Error GoTo 0
    
'00 // �������� ����������� �����
    If Not gFSO.FileExists(tFileFullPath) Then
        uADebugPrint tLogTag, "�� ������� ������� ����! ������� �� ���� � XML - " & tFileWasDownloaded & "; ������ HTTP - " & tHTTP.Status
        Set tHTTP = Nothing
        Exit Function
    End If
    
'00 // ����������
    fDownloadFileByAPI = tFileFullPath
    Set tHTTP = Nothing
End Function


Private Function fDownloadFile(inLinkString, inDropFolder, inFileName)
Dim tLogTag, tFileFullPath, tStream, tStreamWasActive, tHTTP
    tLogTag = "DWNLD"
    fDownloadFile = vbNullString
    tStreamWasActive = False
    If Not (gFSO.FolderExists(inDropFolder)) Then
        uADebugPrint tLogTag, "����� ��� ���������� ����� �� ����������!"
        Exit Function
    End If
    tFileFullPath = inDropFolder & "\" & inFileName
    If gFSO.FileExists(tFileFullPath) Then
        If Not (uDeleteFile(tFileFullPath)) Then
            uADebugPrint tLogTag, "�� ������� ������� ������ ���� ������ �������� �������������� ������� �����! ��� ����� - <" & inFileName & ">"
            Exit Function
        End If
    End If
    'work it
    Set tHTTP = CreateObject("WinHttp.WinHttpRequest.5.1") '������ HTTP
    On Error Resume Next
        tHTTP.Open "GET", inLinkString, False
        'objHTTP.setProxy 2, "proxy.belkam.com:8090", ""
        'objHTTP.SetCredentials "yahaustov", "27u6as", 1
        'objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        tHTTP.SetTimeouts 10, 8, 8, 20
        tHTTP.Send
        'uADebugPrint tLogTag, "������ ������� �� ������� HTTP - " & tHTTP.Status
        If tHTTP.Status = 200 Then
            tStreamWasActive = True
            Set tStream = CreateObject("ADODB.Stream")
            With tStream
                .Type = 1 'adTypeBinary
                .Open
                .Write tHTTP.ResponseBody
                .SaveToFile tFileFullPath
                .Close
            End With
            Set tStream = Nothing
        End If
        '
        If Err.Number <> 0 Then
            uADebugPrint tLogTag, "�������������� ������ ��� ���������� �����: " & Err.Description
            Set tHTTP = Nothing
            Exit Function
        End If
    On Error GoTo 0
    'result
    If Not gFSO.FileExists(tFileFullPath) Then
        uADebugPrint tLogTag, "�� ������� ������� ����! ���������� ������ - " & tStreamWasActive & "; ������ HTTP - " & tHTTP.Status
        Set tHTTP = Nothing
        Exit Function
    End If
    fDownloadFile = tFileFullPath
    Set tHTTP = Nothing
    'tHTTP.WaitForResponse
    'tResult = tHTTP.ResponseText
    'tResultFileName = gDataPath & "\" & "result.txt"
    'uDebugPrint tResultFileName
    'Set tTextFile = gFSO.OpenTextFile(tResultFileName, 2, True)
    'tTextFile.WriteLine tResult
    'tTextFile.Close
End Function

Private Sub fBlankXMLForecastReportCreate(inXML, inReportToDate, inSubjectName, inTimeZoneUTC, inCommentString)
Dim tCurrentTime, tRoot, tNode, tRecord, tComment, tIntro, tLogTag, tDateString, tTimeZoneString
'00 // ����������
    tLogTag = "BLANKFORECAST"
    Set inXML = Nothing
        
'01 // ���������� XML
    Set inXML = CreateObject("Msxml2.DOMDocument.6.0")
    'inXML.ASync = False
    'inXML.Load (tFilePath)
    
'02 // ��������� ���� ������ MESSAGE
    Set tRoot = inXML.CreateElement("data")
    inXML.AppendChild tRoot
    tDateString = Format(inReportToDate, "DD.MM.YYYY")
    tTimeZoneString = "GMT" '???
    If inTimeZoneUTC >= 0 Then
        tTimeZoneString = tTimeZoneString & "+" & inTimeZoneUTC
    Else
        tTimeZoneString = tTimeZoneString & inTimeZoneUTC
    End If
    tRoot.SetAttribute "region", inSubjectName 'SUBJECT NAME
    tRoot.SetAttribute "date", tDateString 'REPORT TO DATE
    tRoot.SetAttribute "timezone", tTimeZoneString 'TIMEZONE
    
'03 // ���� ������� DATETIME
    Set tNode = tRoot.AppendChild(inXML.CreateElement("planned-peak-hours"))
    
'04 // ���� ����������� SENDER
    Set tNode = tRoot.AppendChild(inXML.CreateElement("region-forecast-max-hours"))
    
'05 // �����������
    Set tComment = inXML.CreateComment(inCommentString)
    inXML.InsertBefore tComment, inXML.FirstChild
    
'06 // ���������� ������������ XML
    Set tIntro = inXML.CreateProcessingInstruction("xml", "version='1.0' encoding='UTF-8' standalone='yes'")
    inXML.InsertBefore tIntro, inXML.FirstChild
    
'07 // ������ ���������� �������
    'fSaveXMLChanges inXML, tFilePath
End Sub

Private Function fGetTimeZoneByID(inTimeZoneID)
Dim tXPathString
    Set fGetTimeZoneByID = Nothing
    If Not gXMLDictionary.Active Then: Exit Function
    tXPathString = "//timezones/timezone[@id='" & inTimeZoneID & "']"
    Set fGetTimeZoneByID = gXMLDictionary.XML.SelectSingleNode(tXPathString)
End Function

Private Function fGetSubjectByID(inSubjectID)
Dim tXPathString, tSubjectID
    Set fGetSubjectByID = Nothing
    If Not gXMLDictionary.Active Then: Exit Function
    If inSubjectID < 10 Then 'fix
        tSubjectID = Format(inSubjectID, "00")
    Else
        tSubjectID = CStr(inSubjectID)
    End If
    tXPathString = "//subjects/subject[@id='" & tSubjectID & "']"
    Set fGetSubjectByID = gXMLDictionary.XML.SelectSingleNode(tXPathString)
End Function

Private Function fGetTradeZoneByID(inTradeZoneID)
Dim tXPathString, tTradeZoneID
    Set fGetTradeZoneByID = Nothing
    If Not gXMLDictionary.Active Then: Exit Function
    If inTradeZoneID < 10 Then 'fix
        tTradeZoneID = Format(inTradeZoneID, "00")
    Else
        tTradeZoneID = CStr(inTradeZoneID)
    End If
    tXPathString = "//tradezones/tradezone[@id='" & tTradeZoneID & "']"
    Set fGetTradeZoneByID = gXMLDictionary.XML.SelectSingleNode(tXPathString)
End Function

Private Function fSubFolderGet(inRootFolder, inSubFolderString)
Dim tSubFolderElements, tSubFolderElement, tTempPath
    fSubFolderGet = vbNullString
    If Not gFSO.FolderExists(inRootFolder) Then: Exit Function
    'quick test
    If gFSO.FolderExists(inRootFolder & "\" & inSubFolderString) Then
        fSubFolderGet = inRootFolder & "\" & inSubFolderString
        Exit Function
    End If
    tTempPath = inRootFolder
    tSubFolderElements = Split(inSubFolderString, "\")
    For Each tSubFolderElement In tSubFolderElements
        If tSubFolderElement = vbNullString Then: Exit Function
        tTempPath = tTempPath & "\" & tSubFolderElement
        If Not uFolderCreate(tTempPath) Then: Exit Function
    Next
    fSubFolderGet = tTempPath
End Function
