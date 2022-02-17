Attribute VB_Name = "CalcRoute"
'CALCROUTE
'�������������� ��� ������ � ���ר��� �����������
Option Explicit

Private Const cnstModuleName = "CALCROUTE"
Private Const cnstModuleShortName = "CR"
Private Const cnstModuleVersion = 2
Private Const cnstModuleDate = "17-09-2019"

Private Const cnstMainDelimiter = ";"
Private Const cnstInsideDelimiter = ":"
Private Const cnstTypicalInterval = 48

Private Function fGetLogTag(inTagText)
    fGetLogTag = cnstModuleShortName & "." & inTagText
End Function

'### �������� � ������������ ������� ���������� ���� (��������� 0 ��� ��� ������)
'###
'### IN
'### inIntervalType     - ��� ��������� ������� (s, h, d, m -- �����������, ����, ���, ������);
'### inStartDate        - ���� ������ �������;
'### inStartDate       - ���� ����� �������;
'###
'### OUT
'### outStartDate       - ���� ������ ������� ���������������;
'### outEndDate         - ���� ����� ������� ���������������;
'### outIntervalCount   - ���������� ����������� ������� ��������������� � ������;
'### outErrorText       - ����� ������ (���� ����);
Private Function fIntervalAdapter(inIntervalType, inStartDate, inEndDate, outStartDate, outEndDate, outIntervalCount, outErrorText)

    ' 00 // ���������� ������
    fIntervalAdapter = 0
    outErrorText = vbNullString
    outIntervalCount = 0
    outStartDate = 0
    outEndDate = 0
    
    ' 01 // �������� ������� ���� ������ �������
    If Not IsDate(inStartDate) Then
        fIntervalAdapter = "I#01"
        outErrorText = "������ ��������� ������� �������� ���������� (inStartDate [" & inStartDate & "] �� �������� �����)!"
        Exit Function
    End If
    
    ' 02 // �������� ������� ���� ����� �������
    If Not IsDate(inEndDate) Then
        fIntervalAdapter = "I#02"
        outErrorText = "����� ��������� ������� �������� ���������� (inEndDate [" & inEndDate & "] �� �������� �����)!"
        Exit Function
    End If
    
    ' 03 // ������ � ������������ ���� ������ � ����� ������� � ����������� � ��������� ����� ���������
    Select Case inIntervalType
        Case "s", "S":
            outStartDate = Fix(inStartDate * 48) / 48
            outEndDate = Fix(inEndDate * 48) / 48
        Case "h", "H":
            outStartDate = Fix(inStartDate * 24) / 24
            outEndDate = Fix(inEndDate * 24) / 24 + 1 / 48 'hour extender
        Case "d", "D":
            outStartDate = Fix(inStartDate)
            outEndDate = Fix(inEndDate) + 47 / 48 'day extender
        Case "m", "M":
            outStartDate = Fix(DateSerial(Year(inStartDate), Month(inStartDate), 1))
            outEndDate = Fix(DateSerial(Year(inEndDate), Month(inEndDate), 1))
            outEndDate = DateAdd("m", 1, outEndDate) - 1 / 48
        Case Else:
            fIntervalAdapter = "I#03"
            outErrorText = "��� ��������� inIntervalType [" & inIntervalType & "] �� ���������!"
            Exit Function
    End Select
    
    ' 04 // �������� ��������������� ��� ������ � ����� �� �� �������
    If outStartDate = 0 Or outEndDate = 0 Then
        fIntervalAdapter = "I#04"
        outErrorText = "������� ��������� ��������� �������� (outStartDate [" & outStartDate & "] .. outEndDate [" & outEndDate & "])!"
        Exit Function
    End If
    
    ' 05 // �������� ��������������� ��� ������ � ����� �� �� ���������� (���� ������ ������ ���� ������ ���� �����)
    If outStartDate > outEndDate Then
        fIntervalAdapter = "I#05"
        outErrorText = "������� ������ ��������� ��������� ������ ������� ����� (outStartDate [" & outStartDate & "] > outEndDate [" & outEndDate & "])!"
        Exit Function
    End If
    
    ' 06 // ��������� ������� ������� ���������� ����� ������������ ��� ������� (������� �������� - �����������)
    outIntervalCount = Round((outEndDate - outStartDate) * 48) + 1
End Function

Private Sub fAddElementToFormula(inBase, inElement, Optional inOperator = vbNullString)
    If inElement = vbNullString Then: Exit Sub
    If inBase = vbNullString Then
        inBase = inElement
    Else
        inBase = inBase & cnstMainDelimiter & inElement
        If inOperator <> vbNullString Then: inBase = inBase & cnstMainDelimiter & inOperator
    End If
End Sub

Private Sub fAddElementToTextArray(inBase, inElement, Optional inDelimiter = cnstMainDelimiter)
    If inElement = vbNullString Then: Exit Sub
    If inBase = vbNullString Then
        inBase = inElement
    Else
        inBase = inBase & inDelimiter & inElement
    End If
End Sub

Private Function fExtractFormulaNode(inDirection, inMode, inNode)
    Dim tXPathString, tNode, tValue, tTempValue
    
    fExtractFormulaNode = vbNullString
    
    If inNode Is Nothing Then: Exit Function
    
    'DIRECTION
    If inDirection = "R" Then
        tXPathString = "child::formula[@direction='receive']"
    ElseIf inDirection = "S" Then
        tXPathString = "child::formula[@direction='send']"
    Else
        Exit Function
    End If
    
    'READMODE
    Select Case inMode
        
        'MAIN
        Case "M":
            On Error Resume Next
                Set tNode = inNode.SelectSingleNode(tXPathString & "/main")
                tValue = tNode.Text
                
                If Err.Number = 0 Then: fExtractFormulaNode = tValue
            On Error GoTo 0
        
        'LOSSES
        Case "L":
            On Error Resume Next
                Set tNode = inNode.SelectSingleNode(tXPathString & "/losses")
                tValue = tNode.Text
                
                If Err.Number = 0 Then: fExtractFormulaNode = tValue
            On Error GoTo 0
            
        'FULL
        Case "F":
            On Error Resume Next
                'MAIN
                Set tNode = inNode.SelectSingleNode(tXPathString & "/main")
                tValue = tNode.Text
                
                'LOSSES
                Set tNode = inNode.SelectSingleNode(tXPathString & "/losses")
                If Not tNode Is Nothing Then
                    fAddElementToFormula tValue, tNode.Text
                    fAddElementToFormula tValue, "+"
                End If
                
                'COEF
                If tValue <> vbNullString Then
                    Set tNode = inNode.SelectSingleNode(tXPathString)
                    tTempValue = tNode.GetAttribute("coefficient")
                    If (tTempValue <> "1") Then
                        fAddElementToFormula tValue, tTempValue
                        fAddElementToFormula tValue, "*"
                    End If
                End If
                
                If Err.Number = 0 Then: fExtractFormulaNode = tValue
            On Error GoTo 0
        Case Else: Exit Function
    End Select
    
    Set tNode = Nothing
End Function

Private Function fReadTPAUPFormula(inReadMode, inTPAUPNode, outFormula, inFormulaOwnerMode, outErrorText)
    Dim tRFormula, tSFormula, tFormula

    ' 00 // ���������� ������
    fReadTPAUPFormula = 0
    outFormula = vbNullString
    outErrorText = vbNullString
    
    ' 01 // �������� ����
    If inTPAUPNode Is Nothing Then
        fReadTPAUPFormula = "E#01"
        outErrorText = "������� ���� inTPAUPNode ����������!"
        Exit Function
    End If
    
    tFormula = vbNullString
    Select Case inReadMode
        
        '## FULL
        Case "F":
            tSFormula = fExtractFormulaNode("S", "F", inTPAUPNode)
            tRFormula = fExtractFormulaNode("R", "F", inTPAUPNode)
            
            fAddElementToFormula tFormula, tSFormula
            fAddElementToFormula tFormula, tRFormula
            
            If tSFormula <> vbNullString And tRFormula <> vbNullString Then: fAddElementToFormula tFormula, "+"
        
        '## SEND
        Case "F": tFormula = fExtractFormulaNode("S", "F", inTPAUPNode)
        
        'SEND MAIN
        Case "SM": tFormula = fExtractFormulaNode("S", "M", inTPAUPNode)
        
        'SEND LOSSES
        Case "SL": tFormula = fExtractFormulaNode("S", "L", inTPAUPNode)
        
        '## RECIEVE
        Case "F": tFormula = fExtractFormulaNode("R", "F", inTPAUPNode)
        
        'RECIEVE MAIN
        Case "RM": tFormula = fExtractFormulaNode("R", "M", inTPAUPNode)
        
        'RECIEVE LOSSES
        Case "RL": tFormula = fExtractFormulaNode("R", "L", inTPAUPNode)
        
        'UNKNOWN
        Case Else:
            fReadTPAUPFormula = "E#02"
            outErrorText = "����� ������ ������� ����������! inReadMode=" & inReadMode
            Exit Function
    End Select
    
    'apply OWNERs reversing
    If tFormula <> vbNullString Then
        If Not inFormulaOwnerMode Then
            fAddElementToFormula tFormula, "-1"
            fAddElementToFormula tFormula, "*"
        End If
    End If
        
    'fin
    Debug.Print tFormula
    outFormula = tFormula
End Function

' inFormulaOwnerMode - controling FROM whom we reading formula.. if current TRADER is asking for OWN formula then TRUE, if asking for OTHER TRADER's formula - FALSE
' if inFormulaOwnerMode is FALSE - it'll INVERT formula
Private Function fExtractFullFormulaByORItem(inORNode, inFormulaNodeA, inFormulaNodeB, outFormula, outHasReserve, inFormulaOwnerMode, outErrorText)
    Dim tErrorText, tXPathString, tRRNodeItem, tRRFormula
    Dim tRRFormulaOwnerMode
    
    ' 00 // ���������� ������
    fExtractFullFormulaByORItem = 0
    outFormula = vbNullString
    outErrorText = vbNullString
    outHasReserve = False
    
    ' 01 // ������ ������� ���������������
    If fReadTPAUPFormula("F", inORNode, outFormula, inFormulaOwnerMode, tErrorText) <> 0 Then
        fExtractFullFormulaByORItem = "EFFOI#01"
        outErrorText = "������ ������� �� �������! #ERR > " & tErrorText
        Exit Function
    End If
                
    ' 02 // ������� ��������������� ������ � ����� ������ �������� � �� ���� ������
    If outFormula = vbNullString Then
        fExtractFullFormulaByORItem = "EFFOI#02"
        outErrorText = "������ ������� ��������� ������ �� �������! #ERR > " & tErrorText
        Exit Function
    End If
                
    ' 03 // ����� ����������������, ���� �� ������������ � ��������� (��� ������� ����� ������ - ������� ����������� �� � ��� ��� ������� ���������� ��� ����� �� ���������)
    tXPathString = "child::tp-aup[(@tp-method='rr' and @id-tp-aup='" & inORNode.GetAttribute("id-tp-aup") & "')]"
    Set tRRNodeItem = inFormulaNodeA.SelectSingleNode(tXPathString)
    tRRFormulaOwnerMode = True 'NODE A
    If tRRNodeItem Is Nothing Then
        If Not inFormulaNodeB Is Nothing Then
            Set tRRNodeItem = inFormulaNodeB.SelectSingleNode(tXPathString)
            tRRFormulaOwnerMode = False 'NODE B
        End If
    End If
                
    ' 04 // ���� ��������� ����� ������
    If Not tRRNodeItem Is Nothing Then
        If fReadTPAUPFormula("F", tRRNodeItem, tRRFormula, tRRFormulaOwnerMode, tErrorText) <> 0 Then
            fExtractFullFormulaByORItem = "EFFOI#03"
            outErrorText = "������ ������� �� �������! #ERR > " & tErrorText
            Exit Function
        End If
        
        '���� ���� ���� � �������� ������� - ������ ������ ���� � �������
        If tRRFormula = vbNullString Then
            fExtractFullFormulaByORItem = "EFFOI#04"
            outErrorText = "������ ������� ���������� ������ �� �������! #ERR > " & tErrorText
            Exit Function
        Else
            fAddElementToFormula outFormula, tRRFormula
            fAddElementToFormula outFormula, "R"
            outHasReserve = True
        End If
    End If
End Function


'### ��������� ������� �� ����������
Private Function fExtractFormula(inTraderID, inGTPID, inSectionID, inSectionVersion, inCalcMode, inCalcElementID, outVersionNodeA, outVersionNodeB, outFormula, outErrorText)
    Dim tXPathString, tFormulaNodeA, tFormulaNodeB, tCalcMode, tORNodeItemsA, tORNodeItemsB, tIndex, tRRNodeItem, tFormulaOR, tFormulaRR, tErrorText, tHasReserve
    Dim tTotalFormula()
    Dim tLogTag
    
    ' 00 // ���������� ������
    fExtractFormula = 0
    outErrorText = vbNullString
    Set outVersionNodeA = Nothing   'MAIN
    Set outVersionNodeB = Nothing   'LINKED
    tLogTag = fGetLogTag("fExtractFormula")
    Erase tTotalFormula
    Erase outFormula
    
    ' 01 // �������� ����������� ������ ��� ������ �������� [BASIS, CALCROUTE]
    If Not gXMLBasis.Active Then
        fExtractFormula = "E#01"
        outErrorText = "������ BASIS �� �������� ��� ������!"
        Exit Function
    End If
    
    If Not gXMLCalcRoute.Active Then
        fExtractFormula = "E#02"
        outErrorText = "������ CALCROUTE �� �������� ��� ������!"
        Exit Function
    End If
        
    ' 02 // ����� �������� ���� ������ ������� �� TraderCode, GTPCode, SectionCode, SectionVersion [BASIS]
    tXPathString = "//trader[@id='" & inTraderID & "']/gtp[@id='" & inGTPID & "']/section[@id='" & inSectionID & "']/version[@id='" & inSectionVersion & "']"
    Set outVersionNodeA = gXMLBasis.XML.SelectSingleNode(tXPathString)
    If outVersionNodeA Is Nothing Then
        fExtractFormula = "E#03"
        outErrorText = "������ BASIS �� ����� ����������� ���� [Main]> " & tXPathString
        Exit Function
    End If
    
    ' 03 // ����� ������� ���� ������ ������� �� GTPCode, SectionCode, SectionVersion [BASIS]
    tXPathString = "//trader/gtp[@id='" & inSectionID & "']/section[@id='" & inGTPID & "']/version[@id='" & inSectionVersion & "']"
    Set outVersionNodeB = gXMLBasis.XML.SelectSingleNode(tXPathString)
    
    ' 04 // ����������� ������ ���������� �������
    If inCalcMode = vbNullString Then
        fExtractFormula = "E#04"
        outErrorText = "����� ������� �� �����! inCalcMode"
        Exit Function
    End If
    
    'Debug.Print "#001 - " & inCalcMode
    
    If Len(inCalcMode) > 0 Then
        tCalcMode = UCase(Left(inCalcMode, 1))
        
        'Debug.Print "#002 - " & tCalcMode
        ' 04.01 // ������ ������
        If tCalcMode = "F" Then
            'Debug.Print "#003 - INC"
            tXPathString = "//trader[@id='" & inTraderID & "']/gtp[@id='" & inGTPID & "']/section[@id='" & inSectionID & "']/version[@id='" & inSectionVersion & "']"
            Set tFormulaNodeA = gXMLCalcRoute.XML.SelectSingleNode(tXPathString)
            If tFormulaNodeA Is Nothing Then
                fExtractFormula = "E#05"
                outErrorText = "������ CALCROUTE �� ����� ����������� �������� ���� (��� ������ ������) > " & tXPathString
                Exit Function
            End If
            
            'if dual detected
            If Not outVersionNodeB Is Nothing Then
                tXPathString = "//trader/gtp[@id='" & inSectionID & "']/section[@id='" & inGTPID & "']/version[@id='" & inSectionVersion & "']"
                Set tFormulaNodeB = gXMLCalcRoute.XML.SelectSingleNode(tXPathString)
                If tFormulaNodeB Is Nothing Then
                    fExtractFormula = "E#06"
                    outErrorText = "������ CALCROUTE �� ����� ����������� ������� ���� (��� ������ ������) > " & tXPathString
                    Exit Function
                End If
            Else
                Set tFormulaNodeB = Nothing
            End If
            
            ' ���������� ������� ��� <tp-aup>
            ' 04.01.OR // ������� ��������� ������ �������
            tXPathString = "child::tp-aup[@tp-method='or']"
            Set tORNodeItemsA = tFormulaNodeA.SelectNodes(tXPathString)
            If Not outVersionNodeB Is Nothing Then 'dual?
                Set tORNodeItemsB = tFormulaNodeB.SelectNodes(tXPathString)
                uCDebugPrint tLogTag, 0, "OR Items: A=" & tORNodeItemsA.Length & " B=" & tORNodeItemsB.Length
            Else
                Set tORNodeItemsB = Nothing
                uCDebugPrint tLogTag, 0, "OR Items: A=" & tORNodeItemsA.Length & " B=0"
            End If
            
            ' 04.01.SCAN&BUILD // ������ �������
            ' A-SCAN&BUILD
            For tIndex = 0 To tORNodeItemsA.Length - 1
                If fExtractFullFormulaByORItem(tORNodeItemsA(tIndex), tFormulaNodeA, tFormulaNodeB, tFormulaOR, tHasReserve, True, tErrorText) <> 0 Then
                    fExtractFormula = "E#07"
                    outErrorText = "������ ������� �� �������� OR (A-LVL) �� �������! #ERR > " & tErrorText
                    Exit Function
                End If
                
                uCDebugPrint tLogTag, 0, "A[" & tIndex & ":hasR=" & tHasReserve & "]=" & tFormulaOR
                'fAddElementToFormula tTotalFormula, tFormulaOR, "+"
                fExtendDynamicArray tTotalFormula
                tTotalFormula(UBound(tTotalFormula)) = tFormulaOR
            Next
            
            ' B-SCAN&BUILD
            If Not tORNodeItemsB Is Nothing Then
                For tIndex = 0 To tORNodeItemsB.Length - 1
                    If fExtractFullFormulaByORItem(tORNodeItemsB(tIndex), tFormulaNodeA, tFormulaNodeB, tFormulaOR, tHasReserve, False, tErrorText) <> 0 Then
                        fExtractFormula = "E#08"
                        outErrorText = "������ ������� �� �������� OR (B-LVL) �� �������! #ERR > " & tErrorText
                        Exit Function
                    End If
                    
                    uCDebugPrint tLogTag, 0, "B[" & tIndex & ":hasR=" & tHasReserve & "]=" & tFormulaOR
                    'fAddElementToFormula tTotalFormula, tFormulaOR, "+"
                    fExtendDynamicArray tTotalFormula
                    tTotalFormula(UBound(tTotalFormula)) = tFormulaOR
                Next
            End If
        
        ' 04.02 // �������� ������
        ElseIf tCalcMode = "P-DISABLED" Then 'disabled
        
        ' 04.03 // ����������� �����
        Else
            fExtractFormula = "E#08"
            outErrorText = "����� ������� ����� ������������! tCalcMode=" & tCalcMode
            Exit Function
        End If
    Else
        fExtractFormula = "E#09"
        outErrorText = "����� ������� �� �����! inCalcMode"
        Exit Function
    End If
      
    'inCalcElementID
    'P3
    'tXPathString = "//trader[@id='" & inTraderID & "']/gtp[@id='" & inGTPID & "']/section[@id='" & inSectionID & "']/version[@id='" & inSectionVersion & "']/formula"
    'Set tFormulaNode = gXMLCalcRoute2.XML.SelectSingleNode(tXPathString)
    If IsArrayEmpty(tTotalFormula) Then
        fExtractFormula = "E#10"
        outErrorText = "������ CALCROUTE �� ����� ����������� ������� > " & tXPathString
        Exit Function
    End If
    
    For Each tIndex In tTotalFormula
        Debug.Print tIndex
    Next
    
    'Debug.Print "F=" & tTotalFormula
    outFormula = tTotalFormula
End Function

'Private Function fAddFormulaElement(inFormula, inElement)
'    If inFormula <> vbNullString Then
'        inFormula = inFormula & cnstMainDelimiter & inElement
'    Else
'        inFormula = inElement
'    End If
'End Function

'### �������� ������� MPC � ������ �������
Private Function fGetMeasuringChannelElement(inElement, outMeasuringChannelList)
    Dim tIndex
    
    fGetMeasuringChannelElement = -1
    
    '���� ����� ��� ������ � ������ - �� ������� ������ ���������� ������ ������������ ������ ��� ���������
    If Not IsArrayEmpty(outMeasuringChannelList) Then
        For tIndex = 0 To UBound(outMeasuringChannelList)
            If inElement = outMeasuringChannelList(tIndex) Then
                fGetMeasuringChannelElement = tIndex
                Exit Function
            End If
        Next
        
        tIndex = UBound(outMeasuringChannelList) + 1
    Else
        tIndex = 0
    End If
    
    '������� �� ��� ������ - ������ ����� � ���������� ������ ��� ���������
    ReDim Preserve outMeasuringChannelList(tIndex)
    outMeasuringChannelList(tIndex) = inElement
    
    fGetMeasuringChannelElement = tIndex
End Function

Private Function fScanFrameForMPC(inMPCode, inMPChannelCode, outAreaCode, outFileClass, inAreaCode20, inAreaCode40)
    Dim tTempFileClass, tAreaCodes, tAreaCode, tNode, tXPathString, tAreaElements
    
    fScanFrameForMPC = False
    
    '80020 Areas
    tTempFileClass = "80020"
    tAreaCodes = Split(inAreaCode20, cnstMainDelimiter)
    
    For Each tAreaCode In tAreaCodes
        If tAreaCode <> vbNullString Then
            tAreaElements = Split(tAreaCode, cnstInsideDelimiter)
            tXPathString = "//trader[@id='" & tAreaElements(0) & "']/gtp/area[(@id='" & tAreaElements(1) & "' and @type='1')]/measuringpoint[@code='" & inMPCode & "']/measuringchannel[@code='" & inMPChannelCode & "']"
            Set tNode = gXMLFrame.XML.SelectSingleNode(tXPathString)
            If Not tNode Is Nothing Then
                outAreaCode = tAreaElements(1)
                outFileClass = tTempFileClass
                fScanFrameForMPC = True
                Exit Function
            End If
        End If
    Next
    
    '80040 Areas
    tTempFileClass = "80040"
    tAreaCodes = Split(inAreaCode40, cnstMainDelimiter)
    
    For Each tAreaCode In tAreaCodes
        If tAreaCode <> vbNullString Then
            tAreaElements = Split(tAreaCode, cnstInsideDelimiter)
            tXPathString = "//trader[@id='" & tAreaElements(0) & "']/gtp/area[(@id='" & tAreaElements(1) & "' and @type='0')]/measuringpoint[@code='" & inMPCode & "']/measuringchannel[@code='" & inMPChannelCode & "']"
            Set tNode = gXMLFrame.XML.SelectSingleNode(tXPathString)
            If Not tNode Is Nothing Then
                outAreaCode = tAreaElements(1)
                outFileClass = tTempFileClass
                fScanFrameForMPC = True
                Exit Function
            End If
        End If
    Next
End Function


Private Function fExtractMeasuringChannelList(inVersionNodeA, inVersionNodeB, outFormula, outMeasuringChannelList, outErrorText)
    Dim tFormulaReBuild
    Dim tFormulaElements, tElement, tSubElements, tIndex
    Dim tMeasuringChannelCount, tFormulaIndex, tAreaCode20, tAreaCode40
    Dim tAreaCode, tFileClass

    '00 // Data prepare
    fExtractMeasuringChannelList = 0
    outErrorText = vbNullString
     
    'outMeasuringChannelCount = -1 IsArrayEmpty
    Erase outMeasuringChannelList
    'Debug.Print "IS_EMPTY:" & IsEmpty(outMeasuringChannelList) & " NULL:" & IsNull(outMeasuringChannelList) & " IS_ARRAY_EMPTY:" & IsArrayEmpty(outMeasuringChannelList)
    'Debug.Print UBound(outMeasuringChannelList)
    
    tAreaCode20 = vbNullString
    tAreaCode40 = vbNullString
    
    '01 // Is FRAME available?
    If Not gXMLFrame.Active Then
        fExtractMeasuringChannelList = "M#01"
        outErrorText = "������ FRAME �� �������� ��� ������!"
        Exit Function
    End If
        
    '02 // Extract ChannelList from Formula
    For tFormulaIndex = 0 To UBound(outFormula)
        
        tFormulaReBuild = vbNullString
        tFormulaElements = Split(outFormula(tFormulaIndex), cnstMainDelimiter)
        
        For Each tElement In tFormulaElements
            tSubElements = Split(tElement, cnstInsideDelimiter)
            If tSubElements(0) = "MPC" Then
                tMeasuringChannelCount = fGetMeasuringChannelElement(tElement, outMeasuringChannelList)
                fAddElementToFormula tFormulaReBuild, "DI" & cnstInsideDelimiter & tMeasuringChannelCount ' Format [DI:DataIndex] as link for datasource in postprocessing routes
                'Debug.Print "MPC = " & tElement
            Else
                fAddElementToFormula tFormulaReBuild, tElement
            End If
        Next
    
        outFormula(tFormulaIndex) = tFormulaReBuild
    Next
    
    'Erase outMeasuringChannelList
    If IsArrayEmpty(outMeasuringChannelList) Then: Exit Function 'soft exit - no data channels in formula
                        
    '03 // Extract AreaCodes from VersionNode
    fExtractAreaIDFromVersionNode inVersionNodeA, tAreaCode20, tAreaCode40
    fExtractAreaIDFromVersionNode inVersionNodeB, tAreaCode20, tAreaCode40
    
    'Debug.Print "ACode20=" & tAreaCode20 & " ACode40=" & tAreaCode40
    
    For tIndex = 0 To UBound(outMeasuringChannelList)
        tSubElements = Split(outMeasuringChannelList(tIndex), cnstInsideDelimiter)
        If Not fScanFrameForMPC(tSubElements(1), tSubElements(2), tAreaCode, tFileClass, tAreaCode20, tAreaCode40) Then
            fExtractMeasuringChannelList = "M#02"
            outErrorText = "������ FRAME �� �������� ��������� ������ ��:����� - " & outMeasuringChannelList(tIndex) & "!"
            Exit Function
        End If
        outMeasuringChannelList(tIndex) = outMeasuringChannelList(tIndex) & cnstInsideDelimiter & tAreaCode & cnstInsideDelimiter & tFileClass
        'Debug.Print "MPCC-" & tIndex & " > " & outMeasuringChannelList(tIndex)
        'MPC:562130003113201:04:5600004801:80020
    Next
    
    'If outAreaID20 = vbNullString Then
    '    fExtractMeasuringChannelList = "M#02"
    '    outErrorText = "�� ������� ���������� ����� AREA � ��������� VersionNode (BASIS)!"
    '    Exit Function
    'End If
        
    '04 // Return rebuilded formula
    'outFormula = tFormulaReBuild
End Function

'### UTILITY - Extract AREA ID from Version nodes [BASIS]
Private Sub fExtractAreaIDFromVersionNode(inVersionNode, outAreaID20, outAreaID40)
    Dim tAreaNode, tAreaNodes, tAreaType, tTraderNode, tTraderCode
    
    ' 00 // Precheck
    If inVersionNode Is Nothing Then: Exit Sub
    
    ' 01 // Select AREA childs
    Set tAreaNodes = inVersionNode.SelectNodes("child::area") 'ancestor
    Set tTraderNode = inVersionNode.SelectSingleNode("ancestor::trader")
    tTraderCode = tTraderNode.GetAttribute("id")
    
    ' 02 // Parse AREA childs
    For Each tAreaNode In tAreaNodes
        tAreaType = tAreaNode.GetAttribute("type")
        Select Case tAreaType
            Case "1", 1: fAddElementToTextArray outAreaID20, tTraderCode & cnstInsideDelimiter & tAreaNode.GetAttribute("id")
            Case "0", 0: fAddElementToTextArray outAreaID40, tTraderCode & cnstInsideDelimiter & tAreaNode.GetAttribute("id")
        End Select
    Next
    
    ' 03 // Quit
    Set tTraderNode = Nothing
    Set tAreaNodes = Nothing
End Sub

Private Function fGetAttribute(inXML, inXPathString, inAttributeName, outValue, outTextError)
    Dim tNode, tValue
    
    ' 01 // Prepare DATA
    fGetAttribute = False
    outValue = vbNullString
    outTextError = vbNullString
    
    ' 02 // Get NODE
    Set tNode = inXML.SelectSingleNode(inXPathString)
    If tNode Is Nothing Then
        outTextError = "�� ������� ���������� ���� XPath = " & inXPathString & "!"
        Exit Function
    End If
    
    ' 03 // Get Attribute
    tValue = tNode.GetAttribute(inAttributeName)
    If IsNull(tValue) Then
        outTextError = "�� ������� ���������� �������� @" & inAttributeName & " ���� XPath = " & inXPathString & "!"
        Exit Function
    End If
    
    ' 04 // Over
    fGetAttribute = True
    outValue = tValue
    Set tNode = Nothing
End Function

Private Sub fExtendDynamicArray(inArray)
    If IsArrayEmpty(inArray) Then
        ReDim inArray(0)
    Else
        ReDim Preserve inArray(UBound(inArray) + 1)
    End If
End Sub

Private Function fGetXML800X0HeaderLine(tTempFullPath, tValue)
    Dim tXMLFile, tFile, tElements
    
    fGetXML800X0HeaderLine = False
    tValue = vbNullString
    
    If Not gFSO.FileExists(tTempFullPath) Then: Exit Function
    Set tFile = gFSO.GetFile(tTempFullPath)
    
    tElements = Split(tFile.Name, ".")
    If UBound(tElements) <> 1 Then: Exit Function
    
    tElements = Split(tElements(0), "_")
    If UBound(tElements) <> 4 Then: Exit Function
    
    fGetXML800X0HeaderLine = True
    'tValue = tTempFullPath & cnstInsideDelimiter & tFile.Name & cnstInsideDelimiter & tElements(0)
    tValue = "FR" & cnstInsideDelimiter & tTempFullPath & cnstInsideDelimiter & tFile.Name & cnstInsideDelimiter & tElements(3)
End Function

Private Function fGetM800X0FullFilePath(inVersionNode, inM80020HomeDir, inM80040HomeDir, inStartDate, inEndDate, outM800X0Array, outErrorText)
    Dim tCurrentDate, tGTPCode, tTraderCode, tArea20Code, tArea40Code, tNode, tRootNode, tXPathString, tHomeDir, tTraderINN
    Dim tYear, tMonth, tDay, tTempFullPath, tValue
    
    ' 00 // ���������� ��������� ������
    fGetM800X0FullFilePath = False
    outErrorText = vbNullString
    
    ' 01 // ���� �� ���� �� �����
    If inVersionNode Is Nothing Then
        outErrorText = "VersionNode �� ����������!"
        Exit Function
    End If
    
    ' 02 // ������ ������ �� VersionNode
    'tXPathString = "ancestor::message/descendant::aup-deliverypoints/aup-deliverypoint[(@id-tp-aup='" & outTPID & "' and @aiiscode='" & outTPAIISCode & "' and @trader-code='" & outTPTraderID & "')]/*[contains(name(),'" & outTPDirection & "')]"
    If Not fGetAttribute(inVersionNode, "ancestor::trader", "id", tTraderCode, outErrorText) Then: Exit Function
    If Not fGetAttribute(inVersionNode, "ancestor::trader", "inn", tTraderINN, outErrorText) Then: Exit Function
    If Not fGetAttribute(inVersionNode, "ancestor::gtp", "id", tGTPCode, outErrorText) Then: Exit Function
    If Not fGetAttribute(inVersionNode, "child::area[@type='1']", "id", tArea20Code, outErrorText) Then: Exit Function
    fGetAttribute inVersionNode, "child::area[@type='0']", "id", tArea40Code, outErrorText
    
    ' 03 // ���� �� ���� ���� ������� ���������� � ������ [inStartDate .. inEndDate]
    For tCurrentDate = inStartDate To inEndDate
        
        ' 03.01 // ������� �������� ����
        tYear = Year(tCurrentDate)
        tMonth = Format(Month(tCurrentDate), "00")
        tDay = Format(Day(tCurrentDate), "00")
        
        ' 03.02 // ����� ����� � ��80020
        tXPathString = "//trader[@id='" & tTraderCode & "']/year[@id='" & tYear & "']/month[@id='" & tMonth & "']/day[@id='" & tDay & "']/aiis[@gtpid='" & tGTPCode & "']"
        Set tRootNode = gXML80020DB.XML.SelectSingleNode(tXPathString)
        
        'If Not tRootNode Is Nothing Then
            
            'M80020
            If tArea20Code <> vbNullString Then
                
                ' read DB for RECORD data
                If tRootNode Is Nothing Then
                    Set tNode = Nothing
                Else
                    Set tNode = tRootNode.SelectSingleNode("child::area[@id='" & tArea20Code & "']/outfile")
                End If
                
                If Not tNode Is Nothing Then
                    If Not fBuildM80020DropFolder(inM80020HomeDir, tHomeDir, tYear, tMonth, tGTPCode, False, outErrorText) Then: Exit Function
                    tTempFullPath = tHomeDir & "\" & tNode.Text
                    If Not fGetXML800X0HeaderLine(tTempFullPath, tValue) Then
                        tValue = "NF" & cnstInsideDelimiter & "LinkIsDead"
                    End If
                Else
                    tValue = "NF" & cnstInsideDelimiter & "NotRegistredInDB"
                End If
                
                ' Complete RECORD
                tValue = tYear & tMonth & tDay & cnstInsideDelimiter & tTraderINN & cnstInsideDelimiter & tArea20Code & cnstInsideDelimiter & "80020" & cnstInsideDelimiter & tValue
                
                ' Add RECORD
                fExtendDynamicArray outM800X0Array
                outM800X0Array(UBound(outM800X0Array)) = tValue
                'Debug.Print "FileM20[" & UBound(outM800X0Array) & "] > " & outM800X0Array(UBound(outM800X0Array))
            End If
            
            'M80040
            If tArea40Code <> vbNullString Then
                
                ' read DB for RECORD data
                If tRootNode Is Nothing Then
                    Set tNode = Nothing
                Else
                    Set tNode = tRootNode.SelectSingleNode("child::area[@id='" & tArea40Code & "']/outfile")
                End If
                
                If Not tNode Is Nothing Then
                    If Not fBuildM80020DropFolder(inM80040HomeDir, tHomeDir, tYear, tMonth, tGTPCode, False, outErrorText) Then: Exit Function
                    tTempFullPath = tHomeDir & "\" & tNode.Text
                    If Not fGetXML800X0HeaderLine(tTempFullPath, tValue) Then
                        tValue = "NF" & cnstInsideDelimiter & "LinkIsDead"
                    End If
                Else
                    tValue = "NF" & cnstInsideDelimiter & "NotRegistredInDB"
                End If
                
                ' Complete RECORD
                tValue = tYear & tMonth & tDay & cnstInsideDelimiter & tTraderINN & cnstInsideDelimiter & tArea40Code & cnstInsideDelimiter & "80040" & cnstInsideDelimiter & tValue
                
                ' Add RECORD
                fExtendDynamicArray outM800X0Array
                outM800X0Array(UBound(outM800X0Array)) = tValue
                'Debug.Print "FileM40[" & UBound(outM800X0Array) & "] > " & outM800X0Array(UBound(outM800X0Array))
            End If
            
        'End If
    Next
    
    ' 04 // Over
    fGetM800X0FullFilePath = True
    
End Function

Private Function fDataFileListBuilder(inVersionNodeA, inVersionNodeB, inStartDate, inEndDate, inM80020HomeDir, inM80040HomeDir, out800X0FileList, outErrorText)
    Dim tDayIndex, tYear, tMonth, tDay, tOutFolder, tErrorText, tNode, tXPathString, tIndex, tFuncName
    
    ' 00 // Prepare
    fDataFileListBuilder = 0
    outErrorText = vbNullString
    tFuncName = "fDataFileListBuilder:: "
    
    Erase out800X0FileList
    
    '01 // DateChk
    If Not IsDate(inStartDate) Then
        fDataFileListBuilder = "F#01"
        outErrorText = tFuncName & "���� ������ ������� (" & inStartDate & ") �� �������� �����!"
        Exit Function
    End If
    If Not IsDate(inStartDate) Then
        fDataFileListBuilder = "F#02"
        outErrorText = tFuncName & "���� ���������� ������� (" & inEndDate & ") �� �������� �����!"
        Exit Function
    End If
    
    '02 // R80020 Chk
    If Not gXML80020DB.Active Then
        fDataFileListBuilder = "F#03"
        outErrorText = tFuncName & "���� ������ R80020 �� ������ � ������!"
        Exit Function
    End If
    
    '03 // 80020CFG Chk
    If Not gXML80020CFG.Active Then
        fDataFileListBuilder = "F#04"
        outErrorText = tFuncName & "������������ ��� 80020 �� �������!"
        Exit Function
    End If
    'outEngagedDayCount = Fix(inEndDate - inStartDate)
    
    '04 // GetFileNames
    If Not fGetM800X0FullFilePath(inVersionNodeA, inM80020HomeDir, inM80040HomeDir, inStartDate, inEndDate, out800X0FileList, outErrorText) Then
        fDataFileListBuilder = "F#04"
        outErrorText = tFuncName & outErrorText
        Exit Function
    End If
    
    If Not inVersionNodeB Is Nothing Then
        If Not fGetM800X0FullFilePath(inVersionNodeB, inM80020HomeDir, inM80040HomeDir, inStartDate, inEndDate, out800X0FileList, outErrorText) Then
            fDataFileListBuilder = "F#05"
            outErrorText = tFuncName & outErrorText
            Exit Function
        End If
    End If
    
End Function

'��������� ������ ������� �� ������
Private Function fRawDataExctractor(outRawDataBlock, outTimeWasted, inMeasuringChannelList, inMeasuringChannelCount, inA20FileList, inA40FileList, inFileListCount, outErrorText)
Dim tTimeIndex, tDayIndex, tChannelIndex, tMaxTimeIndex, tXMLData, tVersion, tErrorText, tStructureStatus
    fRawDataExctractor = 0
    outErrorText = vbNullString
    outTimeWasted = GetTickCount
    ' 01 // ���������� ����� ������
    tMaxTimeIndex = inFileListCount * 48
    ReDim outRawDataBlock(tMaxTimeIndex, inMeasuringChannelCount)
    For tTimeIndex = 1 To tMaxTimeIndex
        For tChannelIndex = 0 To inMeasuringChannelCount
            outRawDataBlock(tTimeIndex, tChannelIndex) = 0
        Next
    Next
    ' 02 // ���������� ����� ������
    For tDayIndex = 0 To inFileListCount
        If fOpenXML80020(tXMLData, inA20FileList(tDayIndex), False, tVersion, tStructureStatus, tErrorText) = 0 Then
            'fRawExtract tXMLData, tDayIndex, inMeasuringChannelList, inMeasuringChannelCount
        End If
    Next
    outTimeWasted = outTimeWasted - GetTickCount
End Function

Private Function fGetCalcDBNode(outNode, inXML, inTraderID, inGTPID, inSectionID, inSectionVersion, inYear, inMonth, inDay, inCalcMode, inStatusLine, inForceErase, inIsXMLChanged, tErrorText)
Dim tXPathString, tCreateNewNode, tRootNode, tNode, tTraderID, tGTPID, tSectionID, tCalcMode, tKillCurrentNode, tIndex
' 00 // ����������
    fGetCalcDBNode = False
    tErrorText = vbNullString
    Set outNode = Nothing
    tCreateNewNode = False
    tKillCurrentNode = inForceErase
' 01 // ���������� � ����� ������� ���������� �������
    tTraderID = UCase(inTraderID)
    tGTPID = UCase(inGTPID)
    tSectionID = UCase(inSectionID)
    tCalcMode = UCase(inCalcMode)
' 02 // �������� ������� ����
    tXPathString = "//trader[@id='" & tTraderID & "']/gtp[@id='" & tGTPID & "']/section[@id='" & tSectionID & "']/version[@id='" & inSectionVersion & "']/year[@id='" & inYear & "']/month[@id='" & inMonth & "']/day[@id='" & inDay & "']/calculation[@mode='" & tCalcMode & "']"
    Set outNode = inXML.SelectSingleNode(tXPathString)
' 03 // ��������� ��������� ����
    If outNode Is Nothing Then
        'Debug.Print "No node found " & inYear & inMonth & inDay
        tCreateNewNode = True
    Else
' 04 // �������� ��������� ���� � ����� ��� � ��� ������
        If outNode.ChildNodes.Length <> 48 Then
            tKillCurrentNode = True
            'Debug.Print "No node found " & inYear & inMonth & inDay
        End If
        'clearing?
        If tKillCurrentNode Then
            Set tNode = outNode.ParentNode.RemoveChild(outNode)
            tCreateNewNode = True
        End If
    End If
' 05 // �������� ����� ����, ���� ����������
    If tCreateNewNode Then
        '05.01 / trader level
        Set tRootNode = inXML.DocumentElement
        tXPathString = "//trader[@id='" & tTraderID & "']"
        Set tNode = inXML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("trader"))
            tNode.SetAttribute "id", tTraderID
        End If
        '05.02 / gtp level
        Set tRootNode = tNode
        tXPathString = tXPathString & "/gtp[@id='" & tGTPID & "']"
        Set tNode = inXML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("gtp"))
            tNode.SetAttribute "id", tGTPID
        End If
        '05.03 / section level
        Set tRootNode = tNode
        tXPathString = tXPathString & "/section[@id='" & tSectionID & "']"
        Set tNode = inXML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("section"))
            tNode.SetAttribute "id", tSectionID
        End If
        '05.04 / version level
        Set tRootNode = tNode
        tXPathString = tXPathString & "/version[@id='" & inSectionVersion & "']"
        Set tNode = inXML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("version"))
            tNode.SetAttribute "id", inSectionVersion
        End If
        '05.05 / year level
        Set tRootNode = tNode
        tXPathString = tXPathString & "/year[@id='" & inYear & "']"
        Set tNode = inXML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("year"))
            tNode.SetAttribute "id", inYear
        End If
        '05.06 / month level
        Set tRootNode = tNode
        tXPathString = tXPathString & "/month[@id='" & inMonth & "']"
        Set tNode = inXML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("month"))
            tNode.SetAttribute "id", inMonth
        End If
        '05.07 / day level
        Set tRootNode = tNode
        tXPathString = tXPathString & "/day[@id='" & inDay & "']"
        Set tNode = inXML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("day"))
            tNode.SetAttribute "id", inDay
        End If
        '05.07 / calculation level
        Set tRootNode = tNode
        tXPathString = tXPathString & "/calculation[@mode='" & tCalcMode & "']"
        Set tNode = inXML.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(inXML.CreateElement("calculation"))
            tNode.SetAttribute "mode", tCalcMode
            tNode.SetAttribute "status", inStatusLine
            'items prepare
            Set tRootNode = tNode
            For tIndex = 0 To 47
                Set tNode = tRootNode.AppendChild(inXML.CreateElement("item"))
                tNode.SetAttribute "id", Format(tIndex, "00")
                tNode.SetAttribute "raw", 0
            Next
        End If
        'Finalizer
        inIsXMLChanged = True
        Set outNode = tRootNode
        Set tRootNode = Nothing
    End If
    Set tNode = Nothing
    fGetCalcDBNode = True
End Function

Private Function IsArrayEmpty(inDynamicArray)
    Dim tIndex
    IsArrayEmpty = True
    On Error Resume Next
        tIndex = UBound(inDynamicArray)
        If Err.Number = 0 Then: IsArrayEmpty = False
    On Error GoTo 0
End Function

'get MeasuringChannelList element index by AREA, MP and CHANNEL code
Private Function fGetMeasuringChannelIndex(inMeasuringChannelList, inAreaCode, inMPCode, inCHCode)
    Dim tIndex, tElements
    
    fGetMeasuringChannelIndex = -1
    If IsArrayEmpty(inMeasuringChannelList) Then: Exit Function
    
    'MPC:562130003113201:04:5600004801:80020
    For tIndex = 0 To UBound(inMeasuringChannelList)
        tElements = Split(inMeasuringChannelList(tIndex), cnstInsideDelimiter)
        If inAreaCode = tElements(3) And inMPCode = tElements(1) And inCHCode = tElements(2) Then
            fGetMeasuringChannelIndex = tIndex
            Exit Function
        End If
    Next
End Function

'Private Sub fExtractDataFromXML(inXML, inClass, inVersion, outDataBlock, outStatus, inMeasuringChannelList, inMeasuringChannelCount, outErrorText)
'inXML, inClass, inVersion, outDataBlock, inMeasuringChannelList, outErrorText
Private Sub fExtractDataFromXML(inXML, inClass, inVersion, outDataBlock, inMeasuringChannelList, Optional inIgnoreUncommon = False)
    Dim tXPathString, tMPNodes, tMPNode, tMPCode, tCHNode, tCHCode, tChannelIndex, tPeriodNode, tPeriodStart, tIndex, tValue, tStatus
    Dim tAreaCode, tAreaNode, tAreaNodes, tMeasuringChannelListIndex, tNode, tErrorText, tLogTag
    
    ' 01 // ���������� ������
    tLogTag = cnstModuleShortName & ".DATAEXTXML"
    tErrorText = vbNullString
    
    ' 02 // �������� ���������� ������� � ������ XML ���� 800X0
    If inClass = "80020" Then
        If Not ((inVersion = "1") Or (inVersion = "2")) Then
            tErrorText = "������ > ������������ ������ (" & inVersion & ") XML ������ [" & inClass & "]!"
            uCDebugPrint tLogTag, 1, tErrorText
            Exit Sub
        End If
    ElseIf inClass = "80040" Then
        If Not ((inVersion = "1") Or (inVersion = "2")) Then
            tErrorText = "������ > ������������ ������ (" & inVersion & ") XML ������ [" & inClass & "]!"
            uCDebugPrint tLogTag, 1, tErrorText
            Exit Sub
        End If
    Else
        tErrorText = "������ > ������������ ����� XML = " & inClass
        uCDebugPrint tLogTag, 1, tErrorText
        Exit Sub
    End If
    
    ' 03 // ����� AREA ���
    tXPathString = "//area"
    Set tAreaNodes = inXML.SelectNodes(tXPathString)
    
    '###ERROR exp OFF###
    On Error Resume Next
    
    ' 04 // ������� AREA ���
        For Each tAreaNode In tAreaNodes
                    
            '��������� ��� AREA
            tXPathString = "child::inn"
            Set tNode = tAreaNode.SelectSingleNode(tXPathString)
            tAreaCode = tNode.Text
        
    ' 05 // ����� MeasurePoint ���
            tXPathString = "child::measuringpoint"
            Set tMPNodes = tAreaNode.SelectNodes(tXPathString)
            
    ' 06 // ������� MeasurePoint ���
            For Each tMPNode In tMPNodes
            
                tMPCode = tMPNode.GetAttribute("code")
                
    ' 07 // ������� MP-CHANNEL ���
                For Each tCHNode In tMPNode.ChildNodes
                
                    tCHCode = tCHNode.GetAttribute("code")
                    
    ' 08 // ����� � ������ ����������� � ������� ����� ������ (� 800�0 ����� �������������� ��������� ������������; �.�. � ��������� ������������ �� ��� ������ ������ ��������)
                    tMeasuringChannelListIndex = fGetMeasuringChannelIndex(inMeasuringChannelList, tAreaCode, tMPCode, tCHCode)
                    If tMeasuringChannelListIndex <> -1 Then
                        
    ' 09 // ������� ����������� ������ ������ ������
                        For Each tPeriodNode In tCHNode.ChildNodes
                                                    
                            '��������� ������ �����������
                            tPeriodStart = tPeriodNode.GetAttribute("start")
                            tIndex = CInt(Left(tPeriodStart, 2)) * 2 'very critical
                            If Right(tPeriodStart, 2) = "30" Then: tIndex = tIndex + 1
                            
                            '��������� �������� ����������� � ��������� ���������
                            If outDataBlock(tIndex, tMeasuringChannelListIndex) = "NF2" Or outDataBlock(tIndex, tMeasuringChannelListIndex) = "NF4" Then
                                Set tNode = tPeriodNode.FirstChild
                                tValue = tNode.Text
                                tStatus = tNode.GetAttribute("status")
                                
                                If IsNull(tStatus) Then: tStatus = "0"
                                
                                If inIgnoreUncommon And tStatus = "1" Then: tStatus = "0" 'ignore uncommon flag
                                
                                If tStatus = "0" Then
                                    If IsNumeric(tValue) Then
                                        If CDec(tValue) = Fix(tValue) Then
                                            outDataBlock(tIndex, tMeasuringChannelListIndex) = Fix(tValue)
                                        Else
                                            outDataBlock(tIndex, tMeasuringChannelListIndex) = "UA" & cnstInsideDelimiter & tValue 'not fixed trigger
                                            'Debug.Print "A=" & tAreaCode; "; MP=" & tMPCode & "[" & tCHCode & "]; VALUE=" & tValue & "; E=" & outDataBlock(tIndex, tMeasuringChannelListIndex)
                                        End If
                                    Else
                                        outDataBlock(tIndex, tMeasuringChannelListIndex) = "UN" & cnstInsideDelimiter & tValue 'unnumeric trigger
                                        'Debug.Print "A=" & tAreaCode; "; MP=" & tMPCode & "[" & tCHCode & "]; VALUE=" & tValue & "; E=" & outDataBlock(tIndex, tMeasuringChannelListIndex)
                                    End If
                                Else
                                    outDataBlock(tIndex, tMeasuringChannelListIndex) = "UC" & cnstInsideDelimiter & tStatus & cnstInsideDelimiter & tValue 'uncommon trigger
                                    'Debug.Print "A=" & tAreaCode; "; MP=" & tMPCode & "[" & tCHCode & "]; VALUE=" & tValue & "; E=" & outDataBlock(tIndex, tMeasuringChannelListIndex)
                                End If
                            End If
                            
                        Next
                    
                    End If
                Next
            Next
            
        Next
    
    ' ERROR exp ON
    On Error GoTo 0
    'Debug.Print "KK"
    
    'inClass, inVersion - �� ������������ ����, ���� �������� ����� ������ ��� ������� - ���� ����� ���������
    '������� ���
     ' <<<<<
   
End Sub

'fPrepareDataBlock tDataBlock, tWorkFileList, inMeasuringChannelList, tA20Reading, tA40Reading, tTempStatusLine, tErrorText
Private Function fPrepareDataBlock(outDataBlock, inWorkFileList, inMeasuringChannelList, Optional inIgnoreUncommon = False)
Dim tTimeIndex, tVersion, tStructureStatus, tErrorText, tXMLData, tLogTag, tClass
Dim tFileElements, tFile, tElements, tFileClass, tFullFilePath, tChannelCount, tChannelIndex

    ' 00 // ����������
    fPrepareDataBlock = True
    tLogTag = cnstModuleShortName & ".PREPDATA"
    tErrorText = vbNullString
    'outErrorText = vbNullString
    tChannelCount = UBound(inMeasuringChannelList)
    
    ' 01 // ����� ���� ������, ����������� NF (NoFile)
    ReDim outDataBlock(47, tChannelCount)
    
    For tChannelIndex = 0 To tChannelCount
        
        tClass = Split(inMeasuringChannelList(tChannelIndex), cnstInsideDelimiter)
        If tClass(4) = "80040" Then
            tClass = "NF4" 'no file 80040
        Else
            tClass = "NF2" 'no file 80020
        End If
        
        For tTimeIndex = 0 To 47
            outDataBlock(tTimeIndex, tChannelIndex) = tClass
        Next
    Next
    
    tFileElements = Split(inWorkFileList, cnstMainDelimiter)
    
    '20190925:1834024515:5600004806:80020:FR:\Users\haustov\Data\80020\Processed\2019-09\PMAREM11\80020_1834024515_20190925_4_5600004800.xml:80020_1834024515_20190925_4_5600004800.xml:4
    '20190925:1834024515:5600004805:80040:NF:NotRegistredInDB
    'uCDebugPrint tLogTag, 0, "Entering OK."
    
    For Each tFile In tFileElements
        tElements = Split(tFile, cnstInsideDelimiter)
        
        If UBound(tElements) < 5 Then
            uCDebugPrint tLogTag, 1, "������ > ���� ������ ����� ����� ����������� ���������: " & tFile
            fPrepareDataBlock = False
        End If
        
        tFileClass = tElements(3)

        If tElements(4) = "FR" Then
            tFullFilePath = tElements(5)
        
            'READING BLOCK (80040 ���� ��������)
            If fOpenXML80020(tXMLData, tFullFilePath, False, tVersion, tClass, tStructureStatus, tErrorText) <> 0 Then
                uCDebugPrint tLogTag, 1, "������ > " & tErrorText
            Else
                fExtractDataFromXML tXMLData, tClass, tVersion, outDataBlock, inMeasuringChannelList, inIgnoreUncommon
            End If
            
        End If
    Next
End Function

Private Function fGetDataIndex(inElement, outIndex, inMaxIndex, outErrorText)
    Dim tElements, tLogTag

    fGetDataIndex = False
    outErrorText = vbNullString
    tElements = Split(inElement, cnstInsideDelimiter)
    tLogTag = "fGetDataIndex > "
    outIndex = -1
    
    ' CHECK #1
    If UBound(tElements) <> 1 Then
        outErrorText = tLogTag & "inElement = <" & inElement & "> �� ��������� DI!"
        Exit Function
    End If
    
    ' CHECK #2
    If tElements(0) <> "DI" Then
        outErrorText = tLogTag & "inElement = <" & inElement & "> �� ��������� DI!"
        Exit Function
    End If
    
    ' CHECK #3
    If Not IsNumeric(tElements(1)) Then
        outErrorText = tLogTag & "inElement = <" & inElement & "> �� ��������� DI! ���������� ������!"
        Exit Function
    End If
    
    outIndex = CInt(tElements(1))
    
    ' CHECK #4
    If outIndex < 0 Or outIndex > inMaxIndex Then
        outErrorText = tLogTag & "inElement = <" & inElement & "> �� ��������� DI! ���������� ������ - ��� ����� [0.." & inMaxIndex & "]!"
        outIndex = -1
        Exit Function
    End If
    
    ' RETURN: OK
    fGetDataIndex = True
    
End Function


Private Function fForcedSelect(inBase, inM80040Ignore)
    fForcedSelect = inBase
    
    If inBase = "NF4" Then
        If inM80040Ignore Then: fForcedSelect = 0
        'outNF4Mark = True
    End If
    
    'If inBase = "NF2" Then: outNF2Mark = True
End Function

' ������ ����� ������ ����� (48 �����������) tDataBlock (� ��������� ���� � inMeasuringChannelList) �� ������� inFormula � ���������� ���������� � tResultBlock
' outStatusLine - ���������� �������� ( ABC - ��� A ��� ������������ RMode; ��� B ��� ���������� ������ �� M80020; ��� C ��� ���������� ������ �� M80040; ��������� �������� 0 ��� 1)
Private Function fCalculateDataBlock(inDataBlock, outResultBlock, outResultTemetryBlock, inFormulaList, inMeasuringChannelList, inM80040Ignore, outErrorText)
    Dim tFormula, tFormulaElements, tElementToDataIndexLinks, tErrorText, tIndex, tTempFormula, tElement, tFormulaElementsCount, tElementIndex, tHourIndex, tSubIndex, tFormulaIndex
    Dim tAffixValue, tElements, tChannelIndex, tFormulaCheck, tRMode, tRModeCheck, tSum
    Dim tResultValueA(1)
    Dim tResultValueB(1)
    Dim tTemetryValue(1, 2)
    
    ' 00 // ���������� ������
    fCalculateDataBlock = False
    outErrorText = vbNullString
    tFormulaIndex = -1
    tSum = 0
    
    ' 01 // ��������������� �������� ������� ������
    If IsArrayEmpty(inMeasuringChannelList) Then
        outErrorText = "CDB#01 > ������! ������ ������ ������� <inMeasuringChannelList> �������� ����!"
        Exit Function
    End If
    
    If IsArrayEmpty(inFormulaList) Then
        outErrorText = "CDB#02 > ������! ������ ������ ������ <inFormulaList> �������� ����!"
        Exit Function
    End If
        
    ' 02 // ������ � ��������� � ������
    For Each tFormula In inFormulaList
        tFormulaIndex = tFormulaIndex + 1
        'Debug.Print "CALC [" & tFormulaIndex & "]> " & tFormula
           
    ' 02.01 // ����� ������� ����������� �� ��������
        tFormulaElements = Split(tFormula, cnstMainDelimiter)
        tSubIndex = 0
        tHourIndex = 0
        tSum = 0
        
    ' 02.02 // ���������� ������� ������ �� ������� ������
        tFormulaElementsCount = UBound(tFormulaElements)
        ReDim tElementToDataIndexLinks(tFormulaElementsCount)
        
    ' 02.03 // ���� ������� ������� �������� ��������� ������ - ���������� ������� ������ ������ ������������� � inMeasuringChannelList � ���������� ��� � ������ ������ tElementToDataIndexLinks
        For tElementIndex = 0 To tFormulaElementsCount
            tElement = tFormulaElements(tElementIndex)
            If Left(tElement, 2) = "DI" Then
                If Not fGetDataIndex(tElement, tIndex, UBound(inMeasuringChannelList), tErrorText) Then
                    outErrorText = "CDB#03 > ������! > " & tErrorText
                    Exit Function
                End If
                tElementToDataIndexLinks(tElementIndex) = tIndex
            Else
                tElementToDataIndexLinks(tElementIndex) = -1
            End If
        Next
        
    ' 02.04 // ��� ������ ����������� ������ ���������� ������� ������� ��� ������
        For tIndex = 0 To 47
            
            '���� ������ ���������� ��� ���������, �� ��� ������ ������� ��� ����������� ����� [���������]
            If IsNumeric(outResultBlock(tIndex)) Then
                
            ' 02.04.01 // ���������� ��������� �������
                tTempFormula = vbNullString
                
            ' 02.04.02 // ������ ��������� ������� � ������� tDataBlock ��� ������� �����������
                For tElementIndex = 0 To tFormulaElementsCount
                    tElement = tFormulaElements(tElementIndex)
                    If tElementToDataIndexLinks(tElementIndex) = -1 Then
                        fAddElementToTextArray tTempFormula, tElement
                    Else
                        fAddElementToTextArray tTempFormula, fForcedSelect(inDataBlock(tIndex, tElementToDataIndexLinks(tElementIndex)), inM80040Ignore)
                    End If
                Next
                              
            ' 02.04.03 // ������ �������������� ������� ��� ������� �����������
                If Not fGetFormulaCalculation(tTempFormula, tResultValueA(tSubIndex), tResultValueB(tSubIndex), tRModeCheck) Then
                    outErrorText = "CDB#04 [F-" & tFormulaIndex & ":H-" & tIndex & "][" & tTempFormula & "]> " & tResultValueA(tSubIndex)
                    Exit Function
                End If
                
            ' 02.04.04 // ��������� ������ �������������� R-Mode
                If tIndex = 0 Then
                    tRMode = tRModeCheck
                ElseIf tRMode <> tRModeCheck Then
                    outErrorText = "CDB#05 [F-" & tFormulaIndex & ":H-" & tIndex & "][" & tTempFormula & "]> �������� ������ ������� R-Mode [tIndex=" & tIndex & "] �������� R-Mode=[" & tRMode & "], � ������� R-Mode=[" & tRModeCheck & "]!"
                    Exit Function
                End If
                
            Else
                tResultValueA(tSubIndex) = "CE"
                tResultValueB(tSubIndex) = "CE"
            End If
            
            ' 02.04.04 // ������� ���������� ������� � ����� ���� ����������
            'outResultBlock(tIndex) = CDec(tResultValue)
            'tSum = tSum + outResultBlock(tIndex, 0)
            
            ' 02.04.05 // ������ ���� ��� ����� ������� ������� �������
            tSubIndex = tSubIndex + 1
            If tSubIndex = 2 Then '������ ���������� ������ ��� (����� �������� �������� R-Mode)
                'R-Mode ON
                If tRMode Then
                    If IsNumeric(tResultValueA(0)) And IsNumeric(tResultValueA(1)) Then 'A-R
                        fFormResult outResultBlock(tIndex - 1), tResultValueA(0) '#1
                        fFormResult outResultBlock(tIndex), tResultValueA(1) '#2
                    
                    ElseIf IsNumeric(tResultValueB(0)) And IsNumeric(tResultValueB(1)) Then 'B-R
                        fFormResult outResultBlock(tIndex - 1), tResultValueB(0) '#1
                        fFormResult outResultBlock(tIndex), tResultValueB(1) '#2
                        
                        'used RMode
                        outResultTemetryBlock(tIndex - 1) = True
                        outResultTemetryBlock(tIndex) = True
                    
                    Else 'FAIL-R
                        outResultBlock(tIndex - 1) = "RF" '#1
                        outResultBlock(tIndex) = "RF" '#2
                    End If
                    
                'R-Mode OFF
                Else
                    fFormResult outResultBlock(tIndex - 1), tResultValueA(0)
                    fFormResult outResultBlock(tIndex), tResultValueA(1)
                End If
                
                tHourIndex = tHourIndex + 1
                tSubIndex = 0
            End If
        Next
        
        
           
    Next
    
    'DEBUG
        'For tIndex = 0 To 47
        '    If IsNumeric(tSum) Then
        '        If IsNumeric(outResultBlock(tIndex)) Then
        '            tSum = tSum + outResultBlock(tIndex)
        '        Else
        '            tSum = outResultBlock(tIndex)
        '            Exit For
        '        End If
        '    End If
        'Next
        
        
        'Debug.Print "CALC [" & tFormulaIndex & "]> " & tSum
    
    
    fCalculateDataBlock = True
    
        
End Function

Private Sub fFormResult(inBaseA, inOperandA, Optional inOperandFail = vbNullString)
    If IsNumeric(inBaseA) Then
        If IsNumeric(inOperandA) Then
            inBaseA = inBaseA + CDec(inOperandA)
        ElseIf inOperandFail <> vbNullString Then
            inBaseA = inOperandFail
        Else
            inBaseA = inOperandA
        End If
    End If
End Sub

Private Function fGetWorkFileList(inM800X0FileList, inDateStamp, outWorkFileList)
    Dim tElements, tIndex
    
    fGetWorkFileList = False
    outWorkFileList = vbNullString
    
    If IsArrayEmpty(inM800X0FileList) Then: Exit Function
    'Debug.Print "IX#1:" & inDateStamp
    
    For tIndex = 0 To UBound(inM800X0FileList)
        tElements = Split(inM800X0FileList(tIndex), cnstInsideDelimiter)
        'Debug.Print "IX#2:II:" & tElements(0) & " EQ? " & inDateStamp
        If tElements(0) = inDateStamp Then
            If outWorkFileList = vbNullString Then
                outWorkFileList = inM800X0FileList(tIndex)
            Else
                outWorkFileList = outWorkFileList & cnstMainDelimiter & inM800X0FileList(tIndex)
            End If
        End If
    Next
    
    
    'Debug.Print "IX#3:" & outWorkFileList
    If outWorkFileList <> vbNullString Then: fGetWorkFileList = True
End Function

Private Function fGetBasicCalcDBNode(inBasisNode, inCalcDBXML, outIsCalcDBChanged)
    Dim tLogTag, tNode, tXPathString, tRootNode, tValue, tBasicXPathString

    ' 01 // Prepare
    Set fGetBasicCalcDBNode = Nothing
    'outIsCalcDBChanged = False
    tLogTag = fGetLogTag("GETBASICDBNODE")
    
    ' 02 // Preventer
    If inBasisNode Is Nothing Or inCalcDBXML Is Nothing Then
        uCDebugPrint tLogTag, 0, "�������� XML ��������� ����������! ��������!"
        Exit Function
    End If
    
    ' 03 // Form XPathString
    Set tRootNode = inCalcDBXML.DocumentElement 'CALCDB
    
    ' 03.01 // Trader
    tXPathString = "ancestor::trader"
    Set tNode = inBasisNode.SelectSingleNode(tXPathString) 'BASIS
    If tNode Is Nothing Then
        uCDebugPrint tLogTag, 0, "���� ��������� �����������! BASIS XPath > " & tXPathString
        Exit Function
    End If
    tValue = tNode.GetAttribute("id")
    
    tBasicXPathString = "//trader[@id='" & tValue & "']"
    Set tNode = inCalcDBXML.SelectSingleNode(tBasicXPathString) 'CALCDB
    If tNode Is Nothing Then
        Set tNode = tRootNode.AppendChild(inCalcDBXML.CreateElement("trader"))
        tNode.SetAttribute "id", tValue
        outIsCalcDBChanged = True
    End If
    Set tRootNode = tNode 'CALCDB
    
    ' 03.02 // GTP
    tXPathString = "ancestor::gtp"
    Set tNode = inBasisNode.SelectSingleNode(tXPathString) 'BASIS
    If tNode Is Nothing Then
        uCDebugPrint tLogTag, 0, "���� ��������� �����������! BASIS XPath > " & tXPathString
        Exit Function
    End If
    tValue = tNode.GetAttribute("id")
    
    tBasicXPathString = tBasicXPathString & "/gtp[@id='" & tValue & "']"
    Set tNode = inCalcDBXML.SelectSingleNode(tBasicXPathString) 'CALCDB
    If tNode Is Nothing Then
        Set tNode = tRootNode.AppendChild(inCalcDBXML.CreateElement("gtp"))
        tNode.SetAttribute "id", tValue
        outIsCalcDBChanged = True
    End If
    Set tRootNode = tNode 'CALCDB
    
    ' 03.03 // Section
    tXPathString = "ancestor::section"
    Set tNode = inBasisNode.SelectSingleNode(tXPathString) 'BASIS
    If tNode Is Nothing Then
        uCDebugPrint tLogTag, 0, "���� ��������� �����������! BASIS XPath > " & tXPathString
        Exit Function
    End If
    tValue = tNode.GetAttribute("id")
    
    tBasicXPathString = tBasicXPathString & "/section[@id='" & tValue & "']"
    Set tNode = inCalcDBXML.SelectSingleNode(tBasicXPathString) 'CALCDB
    If tNode Is Nothing Then
        Set tNode = tRootNode.AppendChild(inCalcDBXML.CreateElement("section"))
        tNode.SetAttribute "id", tValue
        outIsCalcDBChanged = True
    End If
    Set tRootNode = tNode 'CALCDB

    ' 03.04 // Version
    tValue = inBasisNode.GetAttribute("id") 'BASIS
    
    tBasicXPathString = tBasicXPathString & "/version[@id='" & tValue & "']"
    Set tNode = inCalcDBXML.SelectSingleNode(tBasicXPathString) 'CALCDB
    If tNode Is Nothing Then
        Set tNode = tRootNode.AppendChild(inCalcDBXML.CreateElement("version"))
        tNode.SetAttribute "id", tValue
        outIsCalcDBChanged = True
    End If
    Set tRootNode = tNode 'ROOT-NODE CALCDB

    ' 04 // Check this Basic NODE
    If tRootNode Is Nothing Then
        uCDebugPrint tLogTag, 0, "�������� ���� ��������� �����������! CALCDB XPath > " & tBasicXPathString
        Exit Function
    End If
    
    ' 05 // Over
    Set fGetBasicCalcDBNode = tRootNode
    Set tRootNode = Nothing
    Set tNode = Nothing
End Function

Private Function fGetWorkCalcDBNode(inCalcDBBasicNode, inMode, inYear, inMonth, inDay, outIsXMLChanged)
    Dim tNode, tXPathString, tRootNode
    
    Set fGetWorkCalcDBNode = Nothing
    If inCalcDBBasicNode Is Nothing Then: Exit Function
    
    ' 02 // QuickScan
    tXPathString = "child::year[@id='" & inYear & "']/month[@id='" & inMonth & "']/day[@id='" & inDay & "']/calculation[@mode='" & inMode & "']"
    Set tNode = inCalcDBBasicNode.SelectSingleNode(tXPathString)
    
    ' 03 // If QScan failed > create node
    If tNode Is Nothing Then
        Set tRootNode = inCalcDBBasicNode
        outIsXMLChanged = True
        
        'YEAR
        tXPathString = "child::year[@id='" & inYear & "']"
        Set tNode = tRootNode.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(tRootNode.OwnerDocument.CreateElement("year"))
            tNode.SetAttribute "id", inYear
        End If
        Set tRootNode = tNode
        
        'MONTH
        tXPathString = "child::month[@id='" & inMonth & "']"
        Set tNode = tRootNode.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(tRootNode.OwnerDocument.CreateElement("month"))
            tNode.SetAttribute "id", inMonth
        End If
        Set tRootNode = tNode
        
        'DAY
        tXPathString = "child::day[@id='" & inDay & "']"
        Set tNode = tRootNode.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(tRootNode.OwnerDocument.CreateElement("day"))
            tNode.SetAttribute "id", inDay
        End If
        Set tRootNode = tNode
        
        'CALCULATION MODE
        tXPathString = "child::calculation[@mode='" & inMode & "']"
        Set tNode = tRootNode.SelectSingleNode(tXPathString)
        If tNode Is Nothing Then
            Set tNode = tRootNode.AppendChild(tRootNode.OwnerDocument.CreateElement("calculation"))
            tNode.SetAttribute "mode", inMode
        End If
    End If
    
    ' 04 // Over
    Set fGetWorkCalcDBNode = tNode
    Set tNode = Nothing
    Set tRootNode = Nothing
End Function

Private Function fCheckWorkCalcDBNode(inCalcDBWorkNode, inWorkFileList)
    Dim tNode, tXPathString, tValueA, tValueB, tIndexA, tIndexB, tFileInfoA, tFileInfoB, tGotIt
    
    ' 01 // Prepare
    fCheckWorkCalcDBNode = False
    If inCalcDBWorkNode Is Nothing Then: Exit Function
    
    ' 02 // Childs count error
    If inCalcDBWorkNode.ChildNodes.Length <> 48 Then: Exit Function
    
    ' 03 // Childs count and statement error
    tXPathString = "child::item[(@id and @rawdata and @telemetry)]"
    Set tNode = inCalcDBWorkNode.SelectNodes(tXPathString)
    If tNode.Length <> 48 Then: Exit Function
    
    ' 04 // Check raw-data-file-list
    tValueA = inCalcDBWorkNode.GetAttribute("raw-data-file-list")
    If IsNull(tValueA) Then: Exit Function
    
    '20191016:1834024515:7300003804:80020:FR:\Users\haustov\Data\80020\Processed\2019-10\PBELKAM2\80020_1834024515_20191016_1_7300003800.xml:80020_1834024515_20191016_1_7300003800.xml:1;20191016:1834024515:7300003803:80040:NF:NotRegistredInDB
    tValueA = Split(tValueA, cnstMainDelimiter)
    tValueB = Split(inWorkFileList, cnstMainDelimiter)
    If UBound(tValueA) <> UBound(tValueB) Then: Exit Function
    
    For tIndexA = 0 To UBound(tValueA)
        
        tFileInfoA = Split(tValueA(tIndexA), cnstInsideDelimiter)
        tGotIt = False
        
        For tIndexB = 0 To UBound(tValueB)
            tFileInfoB = Split(tValueB(tIndexB), cnstInsideDelimiter)
            If tFileInfoB(2) = tFileInfoA(0) And tFileInfoB(3) = tFileInfoA(1) And tFileInfoB(4) = tFileInfoA(2) Then
                If tFileInfoB(4) = "FR" Then
                    If tFileInfoB(6) = tFileInfoA(3) Then
                        tGotIt = True
                        Exit For
                    End If
                Else
                    tGotIt = True
                    Exit For
                End If
            End If
        Next
        
        'Debug.Print "FLIST-CHK[" & tIndexA & "]=" & tGotIt
        If Not tGotIt Then: Exit Function
    Next
    
    ' 05 // All is fine
    fCheckWorkCalcDBNode = True
End Function

Private Sub fPrepareWorkCalcDBNode(inCalcDBWorkNode)

    While inCalcDBWorkNode.ChildNodes.Length > 0
        inCalcDBWorkNode.RemoveChild inCalcDBWorkNode.LastChild
    Wend

End Sub

Private Sub fSetDataToWorkCalcDBNode(inCalcDBWorkNode, inResultBlock, inResultTemetryBlock, inWorkFileList)
    Dim tIndex, tNode, tValue, tValueElements, tTempString, tValueList, tRawDataFileList
    
    For tIndex = 0 To 47
        Set tNode = inCalcDBWorkNode.AppendChild(inCalcDBWorkNode.OwnerDocument.CreateElement("item"))
        tNode.SetAttribute "id", Format(tIndex, "00")
        tNode.SetAttribute "telemetry", inResultTemetryBlock(tIndex)
        tNode.SetAttribute "rawdata", inResultBlock(tIndex)
    Next
    
    Set tNode = Nothing
    
    '20191016:1834024515:7300003804:80020:FR:\Users\haustov\Data\80020\Processed\2019-10\PBELKAM2\80020_1834024515_20191016_1_7300003800.xml:80020_1834024515_20191016_1_7300003800.xml:1;20191016:1834024515:7300003803:80040:NF:NotRegistredInDB
    tRawDataFileList = vbNullString
    tValueList = Split(inWorkFileList, cnstMainDelimiter)
    For Each tValue In tValueList
        tValueElements = Split(tValue, cnstInsideDelimiter)
        tTempString = tValueElements(2) & cnstInsideDelimiter & tValueElements(3) & cnstInsideDelimiter & tValueElements(4)
        If tValueElements(4) = "FR" Then: tTempString = tTempString & cnstInsideDelimiter & tValueElements(6)
        
        If tRawDataFileList = vbNullString Then
            tRawDataFileList = tTempString
        Else
            tRawDataFileList = tRawDataFileList & cnstMainDelimiter & tTempString
        End If
    Next
    
    inCalcDBWorkNode.SetAttribute "raw-data-file-list", tRawDataFileList
End Sub
                
Private Sub fGetDataFromWorkCalcDBNode(inCalcDBWorkNode, inResultBlock, inResultTemetryBlock)
    Dim tIndex, tNode ', tTempValue
    
    For Each tNode In inCalcDBWorkNode.ChildNodes
        tIndex = CInt(tNode.GetAttribute("id"))
        inResultTemetryBlock(tIndex) = tNode.GetAttribute("telemetry")
        'tTempValue = Replace(tNode.getAttribute("rawdata"), ".", ",")
        'Debug.Print "VAL=" & tTempValue
        inResultBlock(tIndex) = Replace(tNode.GetAttribute("rawdata"), ".", ",")
    Next
End Sub

'Private Function fDataCalc(inTraderID, inGTPID, inSectionID, inSectionVersion, inAreaID20, inAreaID40, inStartDay, inEndDay, inCalcMode, inFormula, inMeasuringChannelList, inMeasuringChannelCount, inA20FileList, inA40FileList, inFileListCount, outResult, outTimeWasted, outIsUpdated, outStatusLine, outErrorText)
Private Function fDataCalc(inVersionNodeA, inVersionNodeB, inStartDay, inEndDay, inCalcMode, inFormula, inMeasuringChannelList, inM800X0FileList, inM80040Ignore, outResult, outTimeWasted, outIsUpdated, outErrorText, Optional inIgnoreUncommon = False)
    Dim tDayIndex, tIndex, tFileIndex, tXPathString, tYear, tMonth, tDay, tErrorText, tIsXMLChanged, tRollBackXML
    Dim tCalcDBNode
    Dim tDataBlock()
    Dim tResultBlock()
    Dim tResultTemetryBlock()
    Dim tLogTag, tTempStatusLine
    Dim tDayCount, tCurrentDay, tCalcDBActive, tWorkFileList, tDateStamp, tA20Code, tA40Code, tTraderINN, tNeedToCalc, tCalcDBMode
    Dim tCalcDBBasicNode, tCalcDBWorkNode, tCalcMode
    'Dim tTimeWST

    ' 00 // ����������
    tLogTag = fGetLogTag("DATACALC")
    fDataCalc = 0
    outIsUpdated = False
    outErrorText = vbNullString
    outTimeWasted = GetTickCount
    
    ReDim tResultBlock(47)
    ReDim tResultTemetryBlock(47)
    
    'CalcDB ������������ ������ � ������ ������� ������� � ��� ��������� CalcDB XML
    tCalcDBActive = False
    tRollBackXML = False
    'Debug.Print "CALCDB.STATE=" & gCalcDB.Active
    If ((UCase(inCalcMode) = "F") Or (UCase(inCalcMode) = "FULL")) And gCalcDB.Active Then
        tCalcMode = "FULL"
        Set tCalcDBBasicNode = fGetBasicCalcDBNode(inVersionNodeA, gCalcDB.XML, tIsXMLChanged)
        If Not tCalcDBBasicNode Is Nothing Then
            tCalcDBActive = True
            'Debug.Print "ACTIVE"
        ElseIf tIsXMLChanged Then
            tRollBackXML = True
            'Debug.Print "ROLLBACK"
        End If
    End If
    
    ' 02 // ���������� ����� ������ �����������
    tDayCount = Fix(inEndDay - inStartDay) '���������� ������� ���� � ��������� ������� [� ����]
    ReDim outResult((tDayCount + 1) * 48 - 1, 1) '������� ������ ������ ��� ��������� ��������

    ' 03 // ��� ������� ��� ������������ ������
    For tCurrentDay = inStartDay To inEndDay
    
    ' #.00 // DayIndex
        tDayIndex = Fix(tCurrentDay - inStartDay) 'from 0 to DayCount-1
    
    ' #.01 // ���������� ������ � ������� ����
        tYear = Format(Year(tCurrentDay), "0000")
        tMonth = Format(Month(tCurrentDay), "00")
        tDay = Format(Day(tCurrentDay), "00")
        tDateStamp = tYear & tMonth & tDay
        
    ' #.02 // ���������� ������������ ����� ����������
        For tIndex = 0 To 47
           tResultBlock(tIndex) = 0
           tResultTemetryBlock(tIndex) = 0
        Next
                
    ' #.03 // ��������� ��������� ������ �� ������� ����
        'tTimeWST = GetTickCount
        If fGetWorkFileList(inM800X0FileList, tDateStamp, tWorkFileList) Then
            'Debug.Print "IN#2:" & tWorkFileList
            tNeedToCalc = True
            '���� ����� ������ � CalcDB �������
            If tCalcDBActive And Not tRollBackXML Then
                '���������� ������� ���� �� ��������� ���� (���� ���� �� ���������� - ������)
                Set tCalcDBWorkNode = fGetWorkCalcDBNode(tCalcDBBasicNode, tCalcMode, tYear, tMonth, tDay, tIsXMLChanged)
                If tCalcDBWorkNode Is Nothing Then
                    '���� �� ������� �������� ����, �� ���� ��������� �� ����� ����� �������� �� CALCDB
                    If tIsXMLChanged Then: tRollBackXML = True
                Else
                    '���� ���� ��������, ���������� ������ ������� �� � �������������
                    '�������� ��������� ����� ���� -����� ����; -������ ����; -����� ����� ������
                    If fCheckWorkCalcDBNode(tCalcDBWorkNode, tWorkFileList) Then
                        tNeedToCalc = False '�������� �� CalcDB ��� �������
                        'Debug.Print "CALCDB-W/NOCALC"
                    Else
                        '���������� ���� (�������� �������� ���������)
                        tIsXMLChanged = True
                        fPrepareWorkCalcDBNode tCalcDBWorkNode
                        'Debug.Print "CALCDB-W/CALC"
                    End If
                End If
                '�������� ���� �� ��� ����������� ������ ��� ������ �������?
                '���� ����, �� ��������� �� ���� ����� ������ ��� ��������?
                '���� ����� ����� - ���� �����������; ���� ����� ������ - ��������� ��� ����������� ������
            End If
            
            'Debug.Print "CALC = " & tNeedToCalc & " \ CALCDBActive = " & tCalcDBActive & " \ ROLLBACK = " & tRollBackXML
            
            If tNeedToCalc Then
                fPrepareDataBlock tDataBlock, tWorkFileList, inMeasuringChannelList, inIgnoreUncommon
                If Not fCalculateDataBlock(tDataBlock, tResultBlock, tResultTemetryBlock, inFormula, inMeasuringChannelList, inM80040Ignore, tErrorText) Then
                    fDataCalc = "DC#02"
                    outErrorText = "������ �� ������: " & tErrorText
                    Exit Function
                End If
                outIsUpdated = True '��� ������ ���������� ������
                '���� ��� �������� ���������� �������� ��� ������ � CalcDB
                '���������� XML ���������� ����� (������ � ����� �����, ����� �������� �� ������������������)
                If tCalcDBActive And Not tRollBackXML Then: fSetDataToWorkCalcDBNode tCalcDBWorkNode, tResultBlock, tResultTemetryBlock, tWorkFileList
            Else
                fGetDataFromWorkCalcDBNode tCalcDBWorkNode, tResultBlock, tResultTemetryBlock
            End If
        
        Else
        
            uCDebugPrint tLogTag, 0, "�� ������� ������� ������ �� ���� <" & tDateStamp & ">!"
            
            For tIndex = 0 To 47
                tResultBlock(tIndex) = "NF2"
                tResultTemetryBlock(tIndex) = 0
            Next
            
        End If
        'tTimeWST = GetTickCount - tTimeWST
        'uCDebugPrint tLogTag, 0, "TimeWasted[DAY:" & tDayIndex & "]=" & Format(tTimeWST / 1000, "0.00") & "s"
        
    ' #.04 // �������� ������ � ����� ���� ����������
        For tIndex = 0 To 47
            outResult(tDayIndex * 48 + tIndex, 0) = tResultBlock(tIndex)
            outResult(tDayIndex * 48 + tIndex, 1) = tResultTemetryBlock(tIndex)
        Next

    Next
    
    ' 03 // ���������� ��������� ��� CalcDB
    If gCalcDB.Active Then
        If tRollBackXML Then
            fReloadXMLDB gCalcDB, False
        ElseIf tIsXMLChanged Then
            fSaveXMLDB gCalcDB, False
            
        End If
    End If
    
    ' 04 // ����� ������
    outTimeWasted = GetTickCount - outTimeWasted
End Function

Private Sub fStatusLineAdjust(outBaseStatusLine, inStatus, inNF20Mark, inNF40Mark, inRModeMark)
    Dim tR1, tR2, tR3
    
    'Income
    If VarType(outBaseStatusLine) <> vbString Then: outBaseStatusLine = "000"
    If outBaseStatusLine = vbNullString Then: outBaseStatusLine = "000"
    
    'Split
    tR1 = Mid(outBaseStatusLine, 1, 1)
    tR2 = Mid(outBaseStatusLine, 2, 1)
    tR3 = Mid(outBaseStatusLine, 3, 1)
    
    'R1 - NF20, UC
    If tR1 = "0" And (inNF20Mark Or inStatus = "NF2" Or inStatus = "UC") Then: tR1 = "1"
    'R2 - NF40
    If tR2 = "0" And (inNF40Mark Or inStatus = "NF4") Then: tR2 = "1"
    'R3 - RMode
    If tR3 = "0" And inRModeMark Then: tR3 = "1"
    
    'Merge
    outBaseStatusLine = tR1 & tR2 & tR3
End Sub

'�������������� ��������� ����������
Private Function fResultPrepare(inInternalResult, inStartDate, inEndDate, inResultInterval, inRoundMode, outResult, outStatusLine, inM80040Ignore, outErrorText)
Dim tOutResultCount, tIndexPerInterval, tFirstIndex, tLastIndex, tIndexCount, tCurrentIndex, tOutResultIndex, tPrevValue, tCurrentValue, tIndexDate, tHalfHourDateValue, tLogTag, tTotalSum, tRoundTrunc, tTotalSumRounded

    ' 01 // ���������� ������
    tLogTag = cnstModuleShortName & ".RESPREP"
    fResultPrepare = 0
    tTotalSum = 0
    tOutResultCount = -1
    tHalfHourDateValue = 1 / 48
    outErrorText = vbNullString
    outStatusLine = "000"
    
    ' 02 // ��������� ������� ������� (��������� ������� �� ������ �� 48 �����������, ������� �� ����� ���� � ����� ������� � ������ �� ������������)
    tIndexCount = UBound(inInternalResult)                                      '����� �������� �������� ������� ��������
    tFirstIndex = Round((inStartDate - Fix(inStartDate)) * 48)                  '��� ��������� ������ (����������� � ������ ������ ����������)
    tLastIndex = Round(tIndexCount - 47 + (inEndDate - Fix(inEndDate)) * 48)    '��� �������� ������ (����������� � ������ ������ ����������)
    uCDebugPrint tLogTag, 0, "INDEX_INFO tFirstIndex=" & tFirstIndex & "; tLastIndex=" & tLastIndex & "; tIndexCount=" & tIndexCount + 1
    
    ' 03 // ��������� �� ��������
    tIndexDate = inStartDate + tFirstIndex * tHalfHourDateValue
    tOutResultIndex = -1
    tCurrentValue = -1
    tPrevValue = -1
    
    For tCurrentIndex = tFirstIndex To tLastIndex
            
        '��� ����������� ������ ������ �������� ���������� �� ����� (�.�. ����� ���� �����, ������, ������ � �.�.)
        '��� ��������� ����������� �� �����������, � ���� ������ ���������� �� ���������� - �� ��������� �������� ����������� ��� ���� ��������
        Select Case inResultInterval
            Case "s", "S": '�����������
                tCurrentValue = tCurrentIndex '��� ������, �.�. ������ ��� ������� ���� ����� �������
            Case "h", "H": '����
                tCurrentValue = Fix(tIndexDate * 24)
            Case "d", "D": '���
                tCurrentValue = Fix(tIndexDate)
            Case "m", "M": '������
                tCurrentValue = Month(tIndexDate)
            Case "y", "Y": '����
                tCurrentValue = Year(tIndexDate)
            Case "t", "T": '�������
                tCurrentValue = 0 '������ ������ ������
            Case Else:
                fResultPrepare = "RP#01"
                outErrorText = "��� ��������� inResultInterval [" & inResultInterval & "] �� ���������!"
                Exit Function
        End Select
        
        '��� ��������� ������ ������� ����������, ������ ����� ������� ����������
        If tCurrentValue <> tPrevValue Then
            tPrevValue = tCurrentValue '������� ������������� �������� �������� ��� ������ ��������
            
            '���������� �������� ������ �������� ��� ������ ����������
            tOutResultIndex = tOutResultIndex + 1
            ReDim Preserve outResult(tOutResultIndex)
            outResult(tOutResultIndex) = 0
        End If
        
        '���� ������ ����������
        If IsNumeric(outResult(tOutResultIndex)) Then
            If IsNumeric(inInternalResult(tCurrentIndex, 0)) Then
                outResult(tOutResultIndex) = outResult(tOutResultIndex) + inInternalResult(tCurrentIndex, 0)
                tTotalSum = tTotalSum + inInternalResult(tCurrentIndex, 0)
            Else
                outResult(tOutResultIndex) = inInternalResult(tCurrentIndex, 0)
            End If
        End If
        
        '������ ���������
        fStatusLineAdjust outStatusLine, inInternalResult(tCurrentIndex, 0), False, inM80040Ignore, inInternalResult(tCurrentIndex, 1)
        
        '��������� ����������� � ��������� �� ���������
        tIndexDate = tIndexDate + tHalfHourDateValue 'next halfhour
    Next
    
    '��������� ��������
    If tOutResultIndex < 0 Then
        fResultPrepare = "RP#02"
        outErrorText = "��������� �� ����� ���������! ��������!"
        Exit Function
    End If
    
    '��������������: ����������
    ' inRoundMode           - ���������� ������� (0 - �� ���������; 1 - �� ���������� ��� (����������� �������); 2 - �������������� ����������)
    tTotalSumRounded = 0
    Select Case inRoundMode
        Case "1", 1:
            tRoundTrunc = 0
            For tCurrentIndex = 0 To tOutResultIndex
                If IsNumeric(outResult(tCurrentIndex)) Then
                    tPrevValue = Round(outResult(tCurrentIndex) + tRoundTrunc, 0)
                    tRoundTrunc = tRoundTrunc + (outResult(tCurrentIndex) - tPrevValue)
                    outResult(tCurrentIndex) = tPrevValue
                    tTotalSumRounded = tTotalSumRounded + outResult(tCurrentIndex)
                End If
            Next
        Case "2", 2:
            For tCurrentIndex = 0 To tOutResultIndex
                If IsNumeric(outResult(tCurrentIndex)) Then
                    outResult(tCurrentIndex) = Round(outResult(tCurrentIndex), 0)
                    tTotalSumRounded = tTotalSumRounded + outResult(tCurrentIndex)
                End If
            Next
    End Select
    
    If inRoundMode <> 0 Then
        uCDebugPrint tLogTag, 0, "RESULT_INFO CalcResult=" & tTotalSumRounded & " (Real=" & tTotalSum & "); inRoundMode=" & inRoundMode & "; tOutResultIndex=" & tOutResultIndex & "; outStatusLine=" & outStatusLine
    Else
        uCDebugPrint tLogTag, 0, "RESULT_INFO CalcResult=" & tTotalSum & "; inRoundMode=" & inRoundMode & "; tOutResultIndex=" & tOutResultIndex & "; outStatusLine=" & outStatusLine
    End If
End Function

'PP00 // �������� ������� ������� �������
'   0           - ���� ������ ������ (������� ���������� � outResult)
'   ���������   - ���� ������ �� ������ (��� ������ � outError)
'��������� � ��������� ������:
' inTraderID            - ��� �������� ���
' inGTPID               - ��� ��� ��������
' inSectionID           - ��� ��� ��������
' inSectionVersion      - ������ �������
'��������� �������:
' inCalcMode            - ����� ������ ������� (�� ���� ����� ������� �� �������)
' inCalcElementID       - ����� ��������� ����� (�.�. ���� ������� �����, �� ���������� ���������� ��� �������������� ��������)
' inResultInterval      - ��� �������������� ������� (�������, �� �����, �� ���� � �.�.)
' inRoundMode           - ���������� ������� (0 - �� ���������; 1 - �� ���������� ��� (����������� �������); 2 - �������������� ����������)
'��������� ���������� ��������� ������:
' inIntervalType        - ��� ��������� (����, ���, ������ � �.�.)
' inStartDate           - ������ ���������
' inEndDate             - ����� ���������
' ������� ��������, ��� ��� ��������� ������������ ������ � ����� ��������� �������������, ���� ������� ����� 01� 13� � ������ 01� 15� ������ ���,
' � ��� ��������� "���", �� ����� �������� ���� ���� ������� (48 �����������)
'���������� ������:
' outResult             - ������� ����������� ������
' outError              - ����� �� ������� ������
Public Function fGetFactCalculation(inTraderID, inGTPID, inSectionID, inSectionVersion, inCalcMode, inCalcElementID, inM80040Ignore, inResultInterval, inRoundMode, inIntervalType, inStartDate, inEndDate, outResultDateStart, outResultDateEnd, outResult, outIsUpdated, outStatusLine, outError, Optional inIgnoreUncommon = False)
Dim tIndex, tStartDate, tEndDate, tIntevalCount, tErrorID, tVersionNodeA, tVersionNodeB, tStartDay, tEndDay
Dim tFormula()
Dim tRawDataBlock()
Dim tM80020HomeDir, tM80040HomeDir
Dim tTimeWasted
Dim tLogTag
Dim tFileListCount
Dim tM800X0FileList()
Dim tMeasuringChannelCount, tFormulaElements
Dim tMeasuringChannelList()
Dim tInternalResultCount
Dim tInternalResult()
    ' 00 \\ ����������
    tLogTag = cnstModuleShortName & ".FACTCALC"
    outResultDateStart = 0
    outResultDateEnd = 0
    uCDebugPrint tLogTag, 0, "������ �������! TRADER=" & inTraderID & "; SECTION=" & inGTPID & "-" & inSectionID & " [v" & inSectionVersion & "]; CALCMODE=" & inCalcMode & "; ROUNDMODE=" & inRoundMode & "; IGNORE_UC=" & inIgnoreUncommon
    fGetFactCalculation = 0
    'outStatusLine = "00"
    Erase outResult
    'outResult = Empty
    outIsUpdated = False
    tErrorID = 0
    outError = vbNullString
    Erase tFormula
    tMeasuringChannelCount = -1
    tInternalResultCount = -1
    
    tM80020HomeDir = gXML80020CFG.Path.Processed
    tM80040HomeDir = vbNullString
    Erase tM800X0FileList()
    'Set tVersionNode = Nothing
    ' 01 \\ �������� ������� ������ ������������ ��� ������
    
    ' 02 \\ �������� ��������� � ���������� ��� � ������������ ����
    tErrorID = fIntervalAdapter(inIntervalType, inStartDate, inEndDate, tStartDate, tEndDate, tIntevalCount, outError)
    If tErrorID <> 0 Then
        uCDebugPrint tLogTag, 2, "������ [" & tErrorID & "] > " & outError
        fGetFactCalculation = tErrorID
        Exit Function
    End If
    tStartDay = CDate(Fix(tStartDate)) 'day fixed date
    tEndDay = CDate(Fix(tEndDate)) 'day fixed date
    outResultDateStart = tStartDate
    outResultDateEnd = tEndDate
    uCDebugPrint tLogTag, 0, "S=" & Format(tStartDate, "DD.MM.YYYY hh:nn:ss") & "; E=" & Format(tEndDate, "DD.MM.YYYY hh:nn:ss") & "; INT[" & inIntervalType & "]=" & tIntevalCount & "halfs; DAYS=[" & tEndDay - tStartDay + 1 & "]"
    
    ' 03 \\ ����� ������� � BASIS � ������� ����������� � ���� ������� (��������� ����� �������) [BASIS, CALCROUTE]
    tErrorID = fExtractFormula(inTraderID, inGTPID, inSectionID, inSectionVersion, inCalcMode, inCalcElementID, tVersionNodeA, tVersionNodeB, tFormula, outError)
    If tErrorID <> 0 Then
        uCDebugPrint tLogTag, 2, "������ [" & tErrorID & "] > " & outError
        fGetFactCalculation = tErrorID
        Exit Function
    End If
    
    'Debug.Print "FO = " & tFormula
    
    ' 04 \\ ���������� �� ������� ��������� ������ (�� ��� �� 80020 ������� ����) � ������ �� ��������� �� ������� ���������� ��������� �� FRAME [BASIS, FRAME]
    tErrorID = fExtractMeasuringChannelList(tVersionNodeA, tVersionNodeB, tFormula, tMeasuringChannelList, outError)
    If tErrorID <> 0 Then
        uCDebugPrint tLogTag, 2, "������ [" & tErrorID & "] > " & outError
        fGetFactCalculation = tErrorID
        Exit Function
    End If
    
    'Debug.Print "FI = " & tFormula
    If Not IsArrayEmpty(tMeasuringChannelList) Then
        uCDebugPrint tLogTag, 0, "CHANNELS=" & UBound(tMeasuringChannelList) + 1
    Else
        uCDebugPrint tLogTag, 0, "CHANNELS=NOTHING_FOUND"
    End If
    
    ' 05 \\ ���������� ������ ������ � ������� ��� �������
    tErrorID = fDataFileListBuilder(tVersionNodeA, tVersionNodeB, tStartDay, tEndDay, tM80020HomeDir, tM80040HomeDir, tM800X0FileList, outError)
    If tErrorID <> 0 Then
        uCDebugPrint tLogTag, 2, "������ [" & tErrorID & "] > " & outError
        fGetFactCalculation = tErrorID
        Exit Function
    End If
    
    ' 06 \\ ������ ������
    'tErrorID = fDataCalc(inTraderID, inGTPID, inSectionID, inSectionVersion, tAreaID20, tAreaID40, tStartDay, tEndDay, inCalcMode, tFormula, tMeasuringChannelList, tMeasuringChannelCount, tA20FileList, tA40FileList, tFileListCount, tInternalResult, tTimeWasted, outIsUpdated, outStatusLine, outError)
    tErrorID = fDataCalc(tVersionNodeA, tVersionNodeB, tStartDay, tEndDay, inCalcMode, tFormula, tMeasuringChannelList, tM800X0FileList, inM80040Ignore, tInternalResult, tTimeWasted, outIsUpdated, outError, inIgnoreUncommon)
    If tErrorID <> 0 Then
        uCDebugPrint tLogTag, 2, "������ [" & tErrorID & "] > " & outError
        fGetFactCalculation = tErrorID
        Exit Function
    End If
    tInternalResultCount = UBound(tInternalResult)
    
    ' 07 \\ ������������� � ����������� ����������� ������� ������ � ������������� �����
    tErrorID = fResultPrepare(tInternalResult, tStartDate, tEndDate, inResultInterval, inRoundMode, outResult, outStatusLine, inM80040Ignore, outError)
    If tErrorID <> 0 Then
        uCDebugPrint tLogTag, 2, "������ [" & tErrorID & "] > " & outError
        fGetFactCalculation = tErrorID
        Exit Function
    End If

    ' XX \\ ����������
    uCDebugPrint tLogTag, 0, "TimeWasted_on_CalcData=" & Format(tTimeWasted / 1000, "0.00") & "s; tInternalResultCount=" & tInternalResultCount & "; outIsUpdated=" & outIsUpdated & "; StatusLine=" & outStatusLine
End Function

Public Sub fGTPCalcRoute(inTrader, inGTP, inStartDate, inEndDate)
    If Not fConfiguratorInit Then: Exit Sub
    If Not fXMLSmartUpdate("0,2,4,6,16,13,12") Then: Exit Sub '-12 CALC DB
End Sub

Public Sub fTestCalc()
Dim tCalcResult, tErrorText, tDateStart, tDateEnd, tIsUpdated, tResultDateStart, tResultDateEnd, tStatusLine, tSleepTime, tIndex
Dim tResult()
Dim tTextFile, tFilePath
    'BASIS, CONVERTER,FRAME,CALENDAR,R80020DB,MAILSCAN, XSD80020V2,XSDFORECAST,DICTIONARY,BRFORECAST,R30308DB ,CALCROUTE,CALCDB,XSD80040V2,F63DB,CREDENTIALS
    If Not fConfiguratorInit Then: Exit Sub
    If Not fXMLSmartUpdate("BASIS,FRAME,R80020DB,XSD80020V2,CALCROUTE,XSD80040V2,CALCDB") Then: Exit Sub
    tDateStart = DateSerial(2022, 1, 2)  ' + TimeSerial(1, 0, 0)
    tDateEnd = DateSerial(2022, 1, 2) 'DateSerial(2019, 1, 16) + TimeSerial(1, 0, 0)
    'If fGetFactCalculation("BELKAMKO", "PBELKAM5", "PKOMIENE", "2", "FULL", 0, True, "T", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    'If fGetFactCalculation("BELKAMKO", "PBELKAM5", "PKOMIENE", "2", "FULL", 0, True, "T", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    'If fGetFactCalculation("BELKAMKO", "PBELKAM8", "PUTEREK5", "1", "FULL", 0, True, "h", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText, True) <> 0 Then
    'If fGetFactCalculation("BELKAMKO", "PTNEFT17", "PRNENER7", "7", "FULL", 0, True, "h", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    'If fGetFactCalculation("BELKAMKO", "PTNEFT17", "PUDMURTE", "5", "FULL", 0, True, "h", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    'If fGetFactCalculation("BELKAMKO", "PBELKA11", "PORENBEN", "1", "FULL", 0, True, "h", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    'If fGetFactCalculation("BELKAMKO", "PMAREM11", "PORENBEN", "6", "FULL", 0, True, "T", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    If fGetFactCalculation("BELKAMKO", "PMAREM11", "PORENBE6", "7", "FULL", 0, True, "T", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    'If fGetFactCalculation("BELKAMKO", "PBELKAM1", "PPENZAEN", "3", "FULL", 0, True, "d", 1, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    
    'If fGetFactCalculation("BELKAMKO", "PBELKA19", "PVOLGOGE", "1", "FULL", 0, True, "h", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    'If fGetFactCalculation("BELKAMKO", "PBELKA19", "PSARATEN", "2", "FULL", 0, True, "h", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    
    'If fGetFactCalculation("BELKAMKO", "PBELKAM4", "PTUMENEN", "1", "FULL", 0, True, "h", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    'If fGetFactCalculation("VOLGSTGK", "PORENBE6", "PMAREM11", "7", "FULL", 0, True, "T", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    'If fGetFactCalculation("BELKAMKO", "PBELKA20", "PSARATEN", "1", "FULL", 1, True, "d", 1, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    'If fGetFactCalculation("BELKAMKO", "PBELKA13", "PKUBANEN", "1", "FULL", 0, True, "T", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    'If fGetFactCalculation("BELKAMKO", "PBELKA13", "PNESKR11", "1", "FULL", 0, True, "T", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
    'If fGetFactCalculation("BELKAMKO", "PMAREM11", "PORENBE6", "6", "FULL", 0, False, "T", 0, "d", tDateStart, tDateEnd, tResultDateStart, tResultDateEnd, tResult, tIsUpdated, tStatusLine, tErrorText) <> 0 Then
        uDebugPrint tErrorText
    Else
        'Debug.Print "UPDATED=" & tIsUpdated & "; START=" & Format(tResultDateStart, "DD.MM.YYYY hh:nn:ss") & "; END=" & Format(tResultDateEnd, "DD.MM.YYYY hh:nn:ss")
        'tSleepTime = 5
        'uDebugPrint "SLEEP TEST = " & tSleepTime
        'uSleep tSleepTime
        uDebugPrint "RESULT[Length=" & UBound(tResult) + 1 & "] >>> "
        'For tIndex = 0 To UBound(tResult)
        '   uDebugPrint "RESULT[" & tIndex & "]=" & tResult(tIndex)
        'Next
                
        'gDataPath = Environ("HOMEPATH") & "\Data"
        If Not (gFSO.FolderExists(gDataPath)) Then
            uDebugPrint "�� ������� ����� ����� ������ <gDataPath> �� ����: " & gDataPath
        Else
            tFilePath = gDataPath & "\CalcRoute_DataResult.txt"
            Set tTextFile = gFSO.OpenTextFile(tFilePath, 2, True)
            For tIndex = 0 To UBound(tResult)
                tTextFile.WriteLine tResult(tIndex)
            Next
            tTextFile.Close
            uDebugPrint "��������� ���������� � �����: " & tFilePath
        End If
        
    End If
End Sub

'PP01 // ��������� �������� �� �������� ����������
Private Function fIsOperator(inValue)
    fIsOperator = False
    Select Case inValue
        Case "+", "-", "*", "/", "^": fIsOperator = True 'math
        Case "=", "?", ">", "AND": fIsOperator = True 'logical
    End Select
End Function

Private Function fGetArgumentErrorCheck(inArgumentStack, inUpperArgumentIndex, inArgumentsUsed, outErrorIndex)
    Dim tIndex, tCheckIndex
    
    fGetArgumentErrorCheck = False
    outErrorIndex = vbNullString
    
    For tIndex = 1 To inArgumentsUsed
        tCheckIndex = inUpperArgumentIndex - (tIndex - 1)
        If inArgumentStack(tCheckIndex) = "UC" Or inArgumentStack(tCheckIndex) = "NF2" Or inArgumentStack(tCheckIndex) = "NF4" Then
            outErrorIndex = inArgumentStack(tCheckIndex)
            fGetArgumentErrorCheck = True
            Exit Function
        End If
    Next
End Function

'PP02 // ���������� �������� ��� �����������
Private Function fMakeOperation(inOperator, outArgumentStack, outUpperArgumentIndex, outReport)
    Dim tErrorMode, tErrorIndex, tArgumentsUsed, tArgumentAIndex, tArgumentBIndex, tArgumentCIndex
    
    ' 01 // ���������������
    fMakeOperation = False
    outReport = vbNullString
    
    ' 02 // ��������� ���������� ����������� ���������� ��� ��������
    Select Case inOperator
        Case "+": tArgumentsUsed = 2
        Case "-": tArgumentsUsed = 2
        Case "*": tArgumentsUsed = 2
        Case "/": tArgumentsUsed = 2
        Case "^": tArgumentsUsed = 2
        Case "=": tArgumentsUsed = 2
        Case ">": tArgumentsUsed = 2
        Case "AND": tArgumentsUsed = 2
        Case "?": tArgumentsUsed = 3
        Case Else:
            outReport = "�������� ����������!  (" & inOperator & ")!"
            Exit Function
    End Select
    
    ' 03 // �������� ���������� ��������� ���������� � ���������� ����������� ��� ��������
    If outUpperArgumentIndex < (tArgumentsUsed - 1) Then
        outReport = "��������� ���������� ���������� �������� " & tArgumentsUsed & ", � ��������� � ����� ��������� ����� " & outUpperArgumentIndex + 1 & "!  (" & inOperator & ")!"
        Exit Function
    End If
    
    ' 04 // ������ ��������
    tArgumentAIndex = outUpperArgumentIndex - (tArgumentsUsed - 1)
    tArgumentBIndex = outUpperArgumentIndex - (tArgumentsUsed - 2)
    tArgumentCIndex = outUpperArgumentIndex - (tArgumentsUsed - 3)
        
    ' 05 // �������� ���������� �� ������ ������������������
    tErrorMode = fGetArgumentErrorCheck(outArgumentStack, outUpperArgumentIndex, tArgumentsUsed, tErrorIndex)
    
    If Not tErrorMode Then
    
    ' 06 // �������� ���������� �� ������ ������������
        Select Case inOperator
            
            ' ��� 2� �������� ����������
            Case "+", "-", "*", "/", "^", "=", ">":
                If Not IsNumeric(outArgumentStack(tArgumentAIndex)) Then
                    outReport = "�������� � �� �����! (" & outArgumentStack(tArgumentAIndex) & ")!"
                    Exit Function
                End If
                
                If Not IsNumeric(outArgumentStack(tArgumentBIndex)) Then
                    outReport = "�������� B �� �����! (" & outArgumentStack(tArgumentBIndex) & ")!"
                    Exit Function
                End If
                
                outArgumentStack(tArgumentAIndex) = CDbl(outArgumentStack(tArgumentAIndex))
                outArgumentStack(tArgumentBIndex) = CDbl(outArgumentStack(tArgumentBIndex))
            
            ' ��� 3� ���������� (��������� IF)
            Case "?":
                If VarType(outArgumentStack(tArgumentAIndex)) <> vbBoolean Then '�������� � ��� ��������� ������� (������ ����������)
                    outReport = "�������� � ��������� ��������� IF �� ������ ��������! (" & outArgumentStack(tArgumentAIndex) & ")!"
                    Exit Function
                End If '��������� 2 ��������� (B � �) ����� ���� ��� ������ (������� ��� ������, � �������������� ����������� ����� ���������)
                
            Case "AND":
                If VarType(outArgumentStack(tArgumentAIndex)) <> vbBoolean Then '�������� � ��� ��������� ������� (������ ����������)
                    outReport = "�������� � ��������� ��������� AND �� ������ ��������! (" & outArgumentStack(tArgumentAIndex) & ")!"
                    Exit Function
                End If
                
                If VarType(outArgumentStack(tArgumentBIndex)) <> vbBoolean Then '�������� B ��� ��������� ������� (������ ����������)
                    outReport = "�������� B ��������� ��������� AND �� ������ ��������! (" & outArgumentStack(tArgumentBIndex) & ")!"
                    Exit Function
                End If
                
        End Select

    ' 07 // ���������� ��������
        Select Case inOperator
            Case "+": outArgumentStack(tArgumentAIndex) = outArgumentStack(tArgumentAIndex) + outArgumentStack(tArgumentBIndex)
            Case "-": outArgumentStack(tArgumentAIndex) = outArgumentStack(tArgumentAIndex) - outArgumentStack(tArgumentBIndex)
            Case "*": outArgumentStack(tArgumentAIndex) = outArgumentStack(tArgumentAIndex) * outArgumentStack(tArgumentBIndex)
            Case "/": outArgumentStack(tArgumentAIndex) = outArgumentStack(tArgumentAIndex) / outArgumentStack(tArgumentBIndex)
            Case "^": outArgumentStack(tArgumentAIndex) = outArgumentStack(tArgumentAIndex) ^ outArgumentStack(tArgumentBIndex)
            Case "=": outArgumentStack(tArgumentAIndex) = (outArgumentStack(tArgumentAIndex) = outArgumentStack(tArgumentBIndex))
            Case ">": outArgumentStack(tArgumentAIndex) = (outArgumentStack(tArgumentAIndex) > outArgumentStack(tArgumentBIndex))
            Case "AND": outArgumentStack(tArgumentAIndex) = (outArgumentStack(tArgumentAIndex) And outArgumentStack(tArgumentBIndex))
            Case "?":
                If outArgumentStack(tArgumentAIndex) Then 'Condition is ARG A
                    outArgumentStack(tArgumentAIndex) = outArgumentStack(tArgumentBIndex) 'TRUE = Arg B
                Else
                    outArgumentStack(tArgumentAIndex) = outArgumentStack(tArgumentCIndex) 'FALSE = Arg C
                End If
        End Select
        
    Else
        outArgumentStack(tArgumentAIndex) = tErrorIndex
    End If
    
    ' 08 // ���������� � ��������� ��������� ������� �� ������� ���������� �������� (�������� �)
    outUpperArgumentIndex = tArgumentAIndex
    fMakeOperation = True
End Function
    
'PP03 // ���������� ���ר� ������� � �������������� ����������
Private Function fGetFormulaCalculation(inFormula, outResultA, outResultB, outRMode)
Dim tFormulaElements, tFormulaElement, tUpperArgumentIndex, tFormulaElementsCount, tIndex, tErrInfo, tFormulaElementIndex, tArgumentA, tArgumentB, tArgumentC, tArgumentsUsed
Dim tArgumentStack()
    
    ' 01 // �������� ������ � �������� �������
    fGetFormulaCalculation = False
    tFormulaElements = Split(inFormula, cnstMainDelimiter)
    tFormulaElementsCount = UBound(tFormulaElements)
    outResultA = 0
    outResultB = 0
    outRMode = False
    
    ' 02  // ��������� ��������
    tUpperArgumentIndex = -1
    tFormulaElementIndex = -1
    
    ' 03  // ������� ��������� �������
    For Each tFormulaElement In tFormulaElements
        
        tFormulaElementIndex = tFormulaElementIndex + 1
        
        '��������
        If fIsOperator(tFormulaElement) Then
            If tUpperArgumentIndex >= 1 Then '��� �������� ���������� �� ����� 2� ��������� � ����� (�������� ���������� ��� ����� ������ ��������� ���������� �����)
                
                '���� ���������� �� ���� ������� �������, � ��������� �������� ���������� �� ������� �����
                If Not fMakeOperation(tFormulaElement, tArgumentStack, tUpperArgumentIndex, tErrInfo) Then
                    outResultA = "������ (������� #" & tFormulaElementIndex & ")! �������� �� ���������! [" & tErrInfo & "]" 'ERROR in formula
                    Exit Function
                End If
            Else
                outResultA = "������ (������� #" & tFormulaElementIndex & " [" & tFormulaElement & "])! ���������� ���������� � ����� ��� �������� ������ ����!" 'ERROR in formula
                Exit Function
            End If
        '��������
        ElseIf IsNumeric(tFormulaElement) Or (tFormulaElement = "UC" Or tFormulaElement = "NF2" Or tFormulaElement = "NF4") Then '������� � ����
            tUpperArgumentIndex = tUpperArgumentIndex + 1
            ReDim Preserve tArgumentStack(tUpperArgumentIndex)
            tArgumentStack(tUpperArgumentIndex) = tFormulaElement
        ElseIf tFormulaElement = "R" Then
            'Debug.Print "R - MODE! [STACK_SIZE=" & tUpperArgumentIndex & "]; STACK[0]=" & tArgumentStack(0) & "; STACK[1]=" & tArgumentStack(tUpperArgumentIndex)
            If tUpperArgumentIndex <> 1 Then
                outResultA = "������ (������� #" & tFormulaElementIndex & " [" & tFormulaElement & "])! ���������� ���������� � ����� ��� �������� �� ����� ���� (" & tUpperArgumentIndex & ")! �������� R-Mode �� ���������!" 'ERROR in formula
                Exit Function
            End If
            outRMode = True
        '����������
        Else
            outResultA = "������ (������� #" & tFormulaElementIndex & " [" & tFormulaElement & "])! ������� �� �������!" 'ERROR in formula
            Exit Function
        End If
    Next
    '02 // �����
    If tUpperArgumentIndex = 0 Then
        outResultA = tArgumentStack(tUpperArgumentIndex)
        fGetFormulaCalculation = True
    ElseIf outRMode Then
        outResultA = tArgumentStack(0)
        outResultB = tArgumentStack(1)
        fGetFormulaCalculation = True
    Else
        outResultA = "������! �������� ������ �������, ���� �� ������! [STACK_SIZE=" & tUpperArgumentIndex & "]"
    End If
End Function
