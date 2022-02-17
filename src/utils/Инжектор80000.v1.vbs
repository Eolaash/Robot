'Проект "Инжектор80000" v001 от 23.05.2018
'
'ОПИСАНИЕ:
Option Explicit

Const cCurrentVersion = 1
Const cCurrentScript = "Инжектор80000"

Dim gScriptFileName, gFSO, gWSO, gScriptPath, gXMLFrame, gXMLBasis, gXMLFilePathA, gXMLFilePathB, gXMLBasisPathLock, gXMLFramePathLock
Dim tFile, tFilePath
Dim tCurrentRow, tAREAReaded, tAREAInjected

'func 1
Private Function fGetFileExtension(inFileName)
Dim tPos
	fGetFileExtension = vbNullString
	tPos = InStrRev(inFileName, ".")
	If tPos > 0 Then
		fGetFileExtension = UCase(Right(inFileName, Len(inFileName) - tPos))
	End If
End Function

'func 2
Private Function fGetFileName(inFileName)
Dim tPos
	fGetFileName = vbNullString
	tPos = InStrRev(inFileName, ".")
	If tPos > 1 Then
		fGetFileName = Left(inFileName, tPos - 1)
	End If
End Function

'func 3
Private Function fGetPeriod(inText, inYear, inMonth)
Dim tYear, tMonth
	'prep
	fGetPeriod = False
	inYear = 0
	inMonth = 0
	'chk 1
	If Len(inText) <> 6 Then: Exit Function	
	If Not IsNumeric(inText) Then: Exit Function	
	tYear = CInt(Left(inText, 4))
	tMonth = CInt(Right(inText, 2))
	'chk 2
	If tYear < 2000 Or tYear > 2100 Then: Exit Function
	If tMonth < 1 Or tMonth > 12 Then: Exit Function
	'fin
	fGetPeriod = True
	inYear = tYear
	inMonth = tMonth
End Function

'func 4
Private Function fGetTraderID(inText, inTraderID)
	'prep
	fGetTraderID = False
	If Len(inText) <> 8 Then: Exit Function
	'fin
	inTraderID = inText
	fGetTraderID = True
End Function

'func 5
Private Function fGetSubjectID(inText, inSubjectID)
Dim tSubjectID
	'prep
	fGetSubjectID = False
	If Not IsNumeric(inText) Then: Exit Function
	tSubjectID = CInt(inText)
	If tSubjectID < 1 or tSubjectID > 99 Then: Exit Function
	'fin
	inSubjectID = tSubjectID
	fGetSubjectID = True
End Function

Private Function fXMLFrameConfigCreate(inXMLObject, inDropPath)
Dim tFilePath, tRoot, tComment, tIntro, tTextFile, tText, tValue
	fXMLFrameConfigCreate = False
	
	'01 // Resolve File Operations
	tFilePath = inDropPath & "\" & "Frame.xml"
	WScript.Echo tFilePath
	If gFSO.FileExists(tFilePath) Then
		If gFSO.FileExists(tFilePath & ".bak") Then: Exit Function
		gFSO.MoveFile tFilePath, tFilePath & ".bak"
	End If
	If gFSO.FileExists(tFilePath) Then: Exit Function
	WScript.Echo "P1 Over"
	
	'02 // RootNode
	Set inXMLObject = CreateObject("Msxml2.DOMDocument.6.0")
	Set tRoot = inXMLObject.CreateElement("message")
	inXMLObject.AppendChild tRoot
	tValue = "FRAME"
	tRoot.SetAttribute "class", tValue
	tValue = "1"
	tRoot.SetAttribute "version", tValue
	tValue = fGetTimeStamp()
	tRoot.SetAttribute "releasestamp", tValue
	
	'03 // Комментарий
    Set tComment = inXMLObject.CreateComment("Сформировано " & Now() & " " & cCurrentScript & " v" & cCurrentVersion)
    inXMLObject.InsertBefore tComment, inXMLObject.ChildNodes(0)
    
    '04 // Processing Instruction
    Set tIntro = inXMLObject.CreateProcessingInstruction("xml", "version='1.0' encoding='Windows-1251' standalone='yes'")
    inXMLObject.InsertBefore tIntro, inXMLObject.ChildNodes(0)
    
    '05 // Save XML
    inXMLObject.Save (tFilePath)
    
    '06 // Реорганизация XML для удобочитаемости в NotePad++
    Set tTextFile = gFSO.OpenTextFile(tFilePath, 1)
    tText = tTextFile.ReadAll
    tTextFile.Close
    Set tTextFile = gFSO.OpenTextFile(tFilePath, 2, True)
    tText = Replace(tText, "><", "> <")
    tTextFile.Write tText
    tTextFile.Close
    
    '07 // Сохранение изменений в XML
    inXMLObject.Load (tFilePath)
    inXMLObject.Save (tFilePath)
	inDropPath = tFilePath
	fXMLFrameConfigCreate = True
End Function

'func 6
Private Function fGetXMLConfig(inPathList, inXMLObject, inFileName, inClassTag)
Dim tPathList, tLock, tIndex, tFileName, tFilePath, tTempXML, tNode, tValue
	fGetXMLConfig = False
	tPathList = Split(inPathList, ";")
	inPathList = vbNullString
	Set tTempXML = CreateObject("Msxml2.DOMDocument.6.0")
	tTempXML.ASync = False
	tFileName = inFileName
	tIndex = 0
	tLock = False
	'scan
	Do While Not tLock
		If UBound(tPathList) < tIndex Then: Exit Do
		'file path forming
		tFilePath = tPathList(tIndex)
		If Right(tFilePath, 1) <> "\" Then: tFilePath = tFilePath & "\"
		tFilePath = tFilePath & tFileName
		'check if file exist		
		If gFSO.FileExists(tFilePath) Then			
			tTempXML.Load tFilePath
			WScript.Echo TypeName(tTempXML)
			If tTempXML.parseError.ErrorCode = 0 Then 'Parsed?
				Set tNode = tTempXML.DocumentElement 'root
                tValue = tNode.NodeName
                If tValue = "message" Then 'message?
					tValue = UCase(tNode.getAttribute("class"))
                    If tValue = inClassTag Then 'message class is FRAME?
						tValue = tNode.getAttribute("releasestamp")
                        If fCheckTimeStamp(tValue) Then 'release stamp correct?
                            tLock = True							
                        End If
					End If
				End If
			End If
		End If
		tIndex = tIndex + 1
	Loop	
	'fin
	If Not (IsEmpty(tTempXML)) Then: Set tTempXML = Nothing 'release object
	If tLock Then
		'WScript.Echo "LOCK > " & tFilePath
		Set inXMLObject = CreateObject("Msxml2.DOMDocument.6.0")
		inXMLObject.ASync = False
		inXMLObject.Load tFilePath
		inPathList = tFilePath
		fGetXMLConfig = True		
	Else
		'WScript.Echo "Ошибка! XML файл конфигурации " & inClassTag & " не найден!"
	End If	
End Function

'func 7
Private Function fCheckTimeStamp(inValue)
Dim tValue, tYear, tMonth, tDay
    'PREP
    fCheckTimeStamp = False
    'GET
    If Len(inValue) <> 14 or Not IsNumeric(inValue) Then: Exit Function	
    'sec
    tValue = CInt(Right(inValue, 2))    
    If tValue < 0 Or tValue > 59 Then: Exit Function
    'min
    tValue = CInt(Mid(inValue, 11, 2))    
    If tValue < 0 Or tValue > 59 Then: Exit Function
    'hour
    tValue = CInt(Mid(inValue, 9, 2))    
    If tValue < 0 Or tValue > 24 Then: Exit Function
    'day
    tValue = CInt(Mid(inValue, 7, 2))    
    If tValue < 1 Or tValue > 31 Then: Exit Function
    tDay = tValue
    'month
    tValue = CInt(Mid(inValue, 5, 2))    
    If tValue < 1 Or tValue > 12 Then: Exit Function
    tMonth = tValue
    'year
    tValue = CInt(Left(inValue, 4))
    If tValue < 2010 Or tValue > 2025 Then: Exit Function
    tYear = tValue
    'logic check
    If fDaysPerMonth(tMonth, tYear) < tDay Then: Exit Function
    'over
    fCheckTimeStamp = True
End Function

'func 8
Private Function fDaysPerMonth(inMonth, inYear)
    fDaysPerMonth = 0
    Select Case LCase(inMonth)
        Case "январь", 1:       fDaysPerMonth = 31
        Case "февраль", 2:
            If (inYear Mod 4) = 0 Then
                                fDaysPerMonth = 29
            Else
                                fDaysPerMonth = 28
            End If
        Case "март", 3:         fDaysPerMonth = 31
        Case "апрель", 4:       fDaysPerMonth = 30
        Case "май", 5:          fDaysPerMonth = 31
        Case "июнь", 6:         fDaysPerMonth = 30
        Case "июль", 7:         fDaysPerMonth = 31
        Case "август", 8:       fDaysPerMonth = 31
        Case "сентябрь", 9:     fDaysPerMonth = 30
        Case "октябрь", 10:     fDaysPerMonth = 31
        Case "ноябрь", 11:      fDaysPerMonth = 30
        Case "декабрь", 12:     fDaysPerMonth = 31
    End Select
    If inYear <= 0 Then: fDaysPerMonth = 0
End Function

Private Function fInjectATSHours(inData)
Dim tBasisNode, tTargetNode, tTempNode, tDataElements, tYear, tMonth, tSubjectID, tTraderID, tGTPList, tDataElementCount, tIndexA, tIndexB, tCloneNode, tStartIndex, tData, tLock, tGTPListNode, tDaysNode, tGTPListElementNode, tDaysElementNode
	fInjectATSHours = False
	'precheck
	tDataElements = Split(inData, ";")
	If UBound(tDataElements) < 6 Then: Exit Function
	tDataElementCount = tDataElements(UBound(tDataElements))
	If Not IsNumeric(tDataElementCount) Then: Exit Function
	tDataElementCount = CInt(tDataElementCount)
	If UBound(tDataElements) <> 5 + tDataElementCount Then: Exit Function
	'recognize
	tSubjectID = tDataElements(0)
	tTraderID = tDataElements(1)
	tGTPList = Split(tDataElements(2), ",")
	tYear = tDataElements(3)
	tMonth = tDataElements(4)
	tStartIndex = 5
	'zero fix
	If tMonth < 10 Then: tMonth = "0" & tMonth	
	If tSubjectID < 10 Then: tSubjectID = "0" & tSubjectID	
	'basis check
	Set tBasisNode = gXML.SelectNodes("//year[@id='" & tYear & "']/month[@id='" & tMonth & "']/workdays/day")
	If tBasisNode.Length = 0 Then
		WScript.Echo "Ошибка!" & vbCrLf & "Календарь не подготовлен для внесения данных на период " & tYear & tMonth & "!"
		Exit Function
	ElseIf tBasisNode.Length <> tDataElementCount Then
		WScript.Echo "Ошибка!" & vbCrLf & "Календарь на период " & tYear & tMonth & " [" & tBasisNode.Length & "] расходится со значениями полученными для " & tSubjectID & "_" & tTraderID & " [" & tDataElementCount & "]!"
		Exit Function
	End If
	'precheck on workdays sync
	For tIndexA = 0 to tDataElementCount - 1
		tData = Split(tDataElements(tStartIndex + tIndexA), ":")
		tLock = False
		For tIndexB = 0	to tBasisNode.Length - 1
			If CInt(tData(0)) = CInt(tBasisNode(tIndexB).Text) Then 'risky cint of not numeric
				tLock = True
				Exit For
			End If
		Next
		If Not tLock Then
			WScript.Echo "Ошибка!" & vbCrLf & "День [" & tData(0) & "] часа пик не был найден в рабочих днях периода " & tYear & tMonth & "!"
			Exit Function
		End If
	Next	
	'locking injection root node 
	Set tTargetNode = gXML.SelectNodes("//year[@id='" & tYear & "']/month[@id='" & tMonth & "']/atspower")
	If tTargetNode.Length = 0 Then 'creating
		Set tTargetNode = gXML.SelectSingleNode("//year[@id='" & tYear & "']/month[@id='" & tMonth & "']")
		If tTargetNode Is Nothing Then 'not ready
			WScript.Echo "Ошибка!" & vbCrLf & "Календарь не подготовлен для внесения данных на период " & tYear & tMonth & "!"
			Exit Function
		End If
		Set tTempNode = tTargetNode.AppendChild(gXML.CreateElement("atspower"))
		Set tTargetNode = gXML.SelectNodes("//year[@id='" & tYear & "']/month[@id='" & tMonth & "']/atspower")
		If tTargetNode.Length = 0 Then 'unknown error
			WScript.Echo "Ошибка!" & vbCrLf & "Не удаётся добавить блок ""atspower"" к " & tYear & tMonth & "!"
			Exit Function
		End If
	ElseIf tTargetNode.Length > 1 Then 'corrupted
		WScript.Echo "Ошибка!" & vbCrLf & "Календарь на период " & tYear & tMonth & " содержит нарушения структуры: количество блоков ""atspower"" должно быть 1, а составило " & tTargetNode.Length & "."
		Exit Function
	End If
	'lock subject
	Set tTempNode = gXML.SelectNodes("//year[@id=" & tYear & "]/month[@id=" & tMonth & "]/atspower/subject[@id='" & tSubjectID & "' and @traderid='" & tTraderID & "']")
	If tTempNode.Length = 0 Then
		Set tTempNode = tTargetNode(0).AppendChild(gXML.CreateElement("subject"))
		tTempNode.SetAttribute "id", tSubjectID
		tTempNode.SetAttribute "traderid", tTraderID
		Set tTempNode = gXML.SelectNodes("//year[@id=" & tYear & "]/month[@id=" & tMonth & "]/atspower/subject[@id='" & tSubjectID & "' and @traderid='" & tTraderID & "']")
		If tTempNode.Length = 0 Then
			WScript.Echo "Ошибка!" & vbCrLf & "Не удаётся добавить блок ""atspower/subject[@id='" & tSubjectID & "' and @traderid='" & tTraderID & "']"" к " & tYear & tMonth & "!"
			Exit Function
		End If
	ElseIf tTargetNode.Length > 1 Then 'corrupted
		WScript.Echo "Ошибка!" & vbCrLf & "Календарь на период " & tYear & tMonth & " содержит нарушения структуры: количество блоков ""atspower/subject[@id='" & tSubjectID & "' and @traderid='" & tTraderID & "']"" должно быть 1, а составило " & tTargetNode.Length & "."
		Exit Function
	End If
	'clear subject
	While tTempNode(0).ChildNodes.Length > 0
		tTempNode(0).RemoveChild tTempNode(0).LastChild
	Wend
	'make gtplist
	Set tGTPListNode = tTempNode(0).AppendChild(gXML.CreateElement("gtplist"))
	For tIndex = 0 to UBound(tGTPList)
		Set tGTPListElementNode = tGTPListNode.AppendChild(gXML.CreateElement("gtp"))
		tGTPListElementNode.Text = tGTPList(tIndex)
	Next
	'make pick hour list
	Set tDaysNode = tTempNode(0).AppendChild(gXML.CreateElement("days"))
	For tIndex = 0 to tDataElementCount - 1
		Set tDaysElementNode = tDaysNode.AppendChild(gXML.CreateElement("day"))
		tData = Split(tDataElements(tStartIndex + tIndex), ":")
		If tData(0) < 10 Then
			tDaysElementNode.SetAttribute "id", "0" & tData(0)
		Else
			tDaysElementNode.SetAttribute "id", tData(0)
		End If
		tDaysElementNode.Text = tData(1)		
	Next
	'tIndex = 5
	'WScript.Echo "OK"
	fInjectATSHours = True
End Function

Private Function fGetTimeStamp()
Dim tNow, tResult, tTemp
	tNow = Now() '20171017000000
	'year
	tResult = Year(tNow)
	'month
	tTemp = Month(tNow)
	If tTemp < 10 Then: tTemp = "0" & tTemp
	tResult = tResult & tTemp
	'day
	tTemp = Day(tNow)
	If tTemp < 10 Then: tTemp = "0" & tTemp
	tResult = tResult & tTemp
	'hour
	tTemp = Hour(tNow)
	If tTemp < 10 Then: tTemp = "0" & tTemp
	tResult = tResult & tTemp
	'min
	tTemp = Minute(tNow)
	If tTemp < 10 Then: tTemp = "0" & tTemp
	tResult = tResult & tTemp
	'sec
	tTemp = Second(tNow)
	If tTemp < 10 Then: tTemp = "0" & tTemp
	tResult = tResult & tTemp
	'fin
	fGetTimeStamp = tResult
End Function

'sub 1
Private Sub fQuitScript()
	'destroy objects
	Set gFSO = Nothing	
	Set gWSO = Nothing
	Set gXMLBasis = Nothing
	Set gXMLFrame = Nothing
	'quit
	WScript.Quit
End Sub

'sub 2
Private Sub fSaveXMLConfigChanges(inFilePath, inXMLObject)
Dim tNode, tValue, tTextFile, tXMLText, tXMLBufText
	Set tNode = inXMLObject.DocumentElement 'root
	tValue = fGetTimeStamp()
	tNode.SetAttribute "releasestamp", tValue
	inXMLObject.Save (inFilePath)
	'p2
	Set tTextFile = gFSO.OpenTextFile(inFilePath, 1)		
	tXMLText = tTextFile.ReadAll	
	tTextFile.Close
	'p3
	Set tTextFile = gFSO.OpenTextFile(inFilePath, 2, True)	
	tXMLText = Replace(tXMLText,"><","> <")
	tTextFile.Write tXMLText
	tTextFile.Close
	'p4
	inXMLObject.Load(inFilePath) 'RESAVE-READ
	inXMLObject.Save(inFilePath) 'RESAVE-SAVE
End Sub

Private Function fAdoptAreaType(inAreaTypeValue)
	If inAreaTypeValue = "1" Then
		fAdoptAreaType = 1
	ElseIf inAreaTypeValue = "2" Then
		fAdoptAreaType = 0
	Else				
		WScript.Echo "ANOMALY! [AREA TYPE] // fAdoptAreaType func"
		WScript.Quit
	End If
End Function

Private Function fSelectVersionByNodes(inAreaNodes)
	Dim tVersion, tXPathString, tNode, tResultString, tValue, tIndex
	Dim tVersionList, tSelectedVersion, tVersionLock, tAreaText
	
	'default
	fSelectVersionByNodes = -1
	tResultString = vbNullString
	tIndex = 0

	'check?
	If TypeName(inAreaNodes) <> "IXMLDOMSelection" Then: Exit Function 'ANOMALY
	If inAreaNodes.Length <= 1 Then: Exit Function 'ANOMALY
	
	'areatext
	tValue = inAreaNodes(tIndex).GetAttribute("id")
	If IsNull(tValue) Then: Exit Function
	tAreaText = tValue
	tValue = inAreaNodes(tIndex).GetAttribute("type")
	If IsNull(tValue) Then: Exit Function
	tAreaText = tAreaText & "[TYPE=" & tValue & "]"
	
	'form header	
	tXPathString = "ancestor::section"
	Set tNode = inAreaNodes(tIndex).SelectSingleNode(tXPathString)
	If tNode Is Nothing Then: Exit Function
	tValue = tNode.GetAttribute("id")
	If IsNull(tValue) Then: Exit Function
	
	tResultString = tValue 'Section ID
	fSelectVersionByNodes = -2
	
	tXPathString = "ancestor::gtp"
	Set tNode = inAreaNodes(tIndex).SelectSingleNode(tXPathString)
	If tNode Is Nothing Then: Exit Function
	tValue = tNode.GetAttribute("id")
	If IsNull(tValue) Then: Exit Function
	
	tResultString = tValue & "-" & tResultString & ">>>" 'GTP ID
	fSelectVersionByNodes = -3
	
	ReDim tVersionList(inAreaNodes.Length - 1)
	
	'form list
	For tIndex = 0 To inAreaNodes.Length - 1
		tValue = inAreaNodes(tIndex).parentNode.GetAttribute("id")
		If IsNull(tValue) Then: Exit Function
		tVersionList(tIndex) = tValue
		tValue = inAreaNodes(tIndex).parentNode.GetAttribute("status")
		If IsNull(tValue) Then: Exit Function
		tResultString = tResultString & vbCrLf & tVersionList(tIndex) & " (" & tValue & ")"		
	Next
	
	fSelectVersionByNodes = -4
	
	'selecting
	tSelectedVersion = tVersionList(inAreaNodes.Length - 1) 'default last one
	tSelectedVersion = InputBox("ВНИМАНИЕ! Было найдено несколько версий для данного кода AREA " & tAreaText & "! " & vbCrLf & "Выберите версию перетока " & tResultString, "Задайте номер версии из списка", tSelectedVersion)
	If Not(IsNumeric(tSelectedVersion)) Then
		WScript.Echo "Необходимо выбирать варианты из списка и только!"
		Exit Function
	Else
		tVersionLock = -1
		For tIndex = 0 To inAreaNodes.Length - 1
			If Fix(tVersionList(tIndex)) = Fix(tSelectedVersion) Then
				tVersionLock = tIndex
				Exit For
			End If
		Next
		If tVersionLock = -1 Then
			WScript.Echo "Необходимо выбирать варианты из списка и только!"
			Exit Function
		End If
	End If
	
	fSelectVersionByNodes = tIndex	
End Function

Private Sub fFileDataExtract(inFile, inXMLFrame, inXMLBasis)
Dim tXMLFile, tNode, tValue, tVersion, tTraderID, tTraderINN, tAIISCode, tAIISNodes, tAREACode, tAREATimeZone, tAIISNode, tAREANode, tAREAType, tFNode, tRoot, tNewAREANode, tBTraderNode, tBAIISNode, tBAREANode, tMPNodes, tMPNode, tNewMPNode, tCHNode, tNewCHNode, tMPChannelList, tXPathString, tSNode, tTempValue
Dim tSectionVersion
	' 01 \\ Avoid non-format
	If LCase(Right(inFile.Name, 4)) <> ".xml" Then: Exit Sub
	If LCase(Left(inFile.Name, 6)) <> "80000_" Then: Exit Sub
	' 02 \\
	tVersion = 0
	tTraderID = vbNullString
	tTraderINN = vbNullString
	Set tXMLFile = CreateObject("Msxml2.DOMDocument.6.0")
	tXMLFile.ASync = False
	tXMLFile.Load inFile.Path
	If tXMLFile.parseError.ErrorCode = 0 Then 'Parsed?
		Set tNode = tXMLFile.DocumentElement 'root
		tValue = tNode.NodeName
		If tValue = "message" Then 'message?
			tValue = tNode.getAttribute("class")
			If tValue = "80000" Then
				tVersion = tNode.getAttribute("version")
				Set tNode = tXMLFile.SelectSingleNode("//organization")
				tTraderINN = tNode.getAttribute("inn")				
			End If
		End If
	End If
	If Not(tVersion = "1" or tVersion = "3" or tVersion="4") Then 'AVAILABLE READERS
		Set tNode = Nothing
		Set tXMLFile = Nothing
		WScript.Echo "Версия макета 80000 [" & tVersion & "] не поддерживается!"
		Exit Sub
	End If
	' 03 \\ Lock TRADER
	Set tBTraderNode = inXMLBasis.SelectSingleNode("//trader[@inn='" & tTraderINN & "']")
	If tBTraderNode Is Nothing Then
		Set tNode = Nothing
		Set tXMLFile = Nothing
		Exit Sub
	End If	
	tTraderID = tBTraderNode.getAttribute("id")
	WScript.Echo "TRADER=" & tTraderID
	' 00 \\
	Set tAIISNodes = tXMLFile.SelectNodes("//aiis")
	For Each tAIISNode In tAIISNodes
		tAIISCode = tAIISNode.getAttribute("aiiscode")
		If tVersion = "4" Then: tAIISCode = tAIISNode.getAttribute("ats-code")
		Set tBAIISNode = inXMLBasis.SelectSingleNode("//trader[@inn='" & tTraderINN & "']/gtp[@aiiscode='" & tAIISCode & "']")
		If Not(tBAIISNode Is Nothing) Then			
			For Each tAREANode In tAIISNode.ChildNodes				
				If tAREANode.NodeName = "area" Then
					'get area params
					Select Case tVersion
						Case "1", "2", "3":
							tAREACode = tAREANode.getAttribute("areacode")
							tAREAType = tAREANode.getAttribute("grouptype")
						Case "4":
							tAREACode = tAREANode.getAttribute("ats-code")
							tAREAType = tAREANode.getAttribute("group-type")
							tAREATimeZone = tAREANode.getAttribute("time-zone")
					End Select
					tAREAReaded = tAREAReaded + 1
					tAREAType = fAdoptAreaType(tAREAType)
					'scan for this AREA in BASIS
					Set tBAREANode = inXMLBasis.SelectNodes("//trader[@inn='" & tTraderINN & "']/gtp[@aiiscode='" & tAIISCode & "']/section/version/area[@id='" & tAREACode & "']")
					
					'hotfix 2021-05-14					
					If tBAREANode.Length > 1 Then
						tSectionVersion = fSelectVersionByNodes(tBAREANode)
						If tSectionVersion < 0 Or tSectionVersion > tBAREANode.Length - 1 Then
							WScript.Echo "ANOMALY! [get SECTION VERSION(" & tSectionVersion & ")]"
							WScript.Quit
						End If
						Set tBAREANode = tBAREANode(tSectionVersion)
					ElseIf tBAREANode.Length = 1 Then
						Set tBAREANode = tBAREANode(0)
					Else
						Set tBAREANode = Nothing
					End If
					
					'WScript.Quit
					
					'if AREA exists in BASIS					
					If Not(tBAREANode Is Nothing) Then
						'scan for this AREA in FRAME
						'WScript.Echo "IN"
						Set tFNode = inXMLFrame.SelectSingleNode("//trader[@inn='" & tTraderINN & "']/gtp[@aiiscode='" & tAIISCode & "']/area[@id='" & tAREACode & "']")
						If tFNode Is Nothing Then
							'WScript.Echo "AREA " & tAREACode & "[" & tAREAType & "] can be injected!"
							' P1 // TRADER inject
							Set tFNode = inXMLFrame.SelectSingleNode("//trader[@inn='" & tTraderINN & "']")
							If tFNode Is Nothing Then
								Set tFNode = inXMLFrame.DocumentElement 'ROOT
								Set tFNode = tFNode.AppendChild(inXMLFrame.CreateElement("trader"))
								tFNode.SetAttribute "id", tBTraderNode.getAttribute("id")
								tFNode.SetAttribute "name", tBTraderNode.getAttribute("name")
								tFNode.SetAttribute "inn", tBTraderNode.getAttribute("inn")
							End If
							' P2 // AIIS inject
							Set tFNode = inXMLFrame.SelectSingleNode("//trader[@inn='" & tTraderINN & "']/gtp[@aiiscode='" & tAIISCode & "']")
							If tFNode Is Nothing Then
								Set tFNode = inXMLFrame.SelectSingleNode("//trader[@inn='" & tTraderINN & "']") 'ROOT
								If tFNode Is Nothing Then
									WScript.Echo "ANOMALY! [inject AIIS]"
									WScript.Quit
								End If
								Set tFNode = tFNode.AppendChild(inXMLFrame.CreateElement("gtp"))
								tFNode.SetAttribute "id", tBAIISNode.getAttribute("id")
								tFNode.SetAttribute "aiiscode", tBAIISNode.getAttribute("aiiscode")
							End If
							' P3 // AREA inject
							Set tFNode = inXMLFrame.SelectSingleNode("//trader[@inn='" & tTraderINN & "']/gtp[@aiiscode='" & tAIISCode & "']/area[@id='" & tAREACode & "']")
							If tFNode Is Nothing Then
								Set tFNode = inXMLFrame.SelectSingleNode("//trader[@inn='" & tTraderINN & "']/gtp[@aiiscode='" & tAIISCode & "']") 'ROOT
								If tFNode Is Nothing Then
									WScript.Echo "ANOMALY! [inject AREA]"
									WScript.Quit
								End If
								Set tFNode = tFNode.AppendChild(inXMLFrame.CreateElement("area"))
								tFNode.SetAttribute "id", tBAREANode.getAttribute("id")
								tFNode.SetAttribute "type", tAREAType
								If tVersion = "4" Then: tFNode.SetAttribute "time-zone", tAREATimeZone 'v4
							End If
							' P4 // Get DEST and SOURCE by version
							Set tMPChannelList = Nothing
							Set tFNode = inXMLFrame.SelectSingleNode("//trader[@inn='" & tTraderINN & "']/gtp[@aiiscode='" & tAIISCode & "']/area[@id='" & tAREACode & "']") 'DEST
							Select Case tVersion
								Case "1": Set tMPNodes = tXMLFile.SelectNodes("//aiis[@aiiscode='" & tAIISCode & "']/area[@areacode='" & tAREACode & "']/measuringpoint") 'SOURCE
								Case "3": Set tMPNodes = tXMLFile.SelectNodes("//aiis[@aiiscode='" & tAIISCode & "']/area[@areacode='" & tAREACode & "']/measuringpoints/measuringpoint") 'SOURCE
								Case "4": 
									Set tMPChannelList = tXMLFile.SelectNodes("//aiis[@ats-code='" & tAIISCode & "']/area[@ats-code='" & tAREACode & "']/measuring-channels-list/measuring-channel") 'SOURCE-LIST
									If tMPChannelList Is Nothing Then
										WScript.Echo "ANOMALY! [80000 version 4] tMPChannelList is NOTHING"
										WScript.Quit
									ElseIf tMPChannelList.Length = 0 Then
										WScript.Echo "ANOMALY! [80000 version 4] tMPChannelList is EMPTY"
										WScript.Quit
									End If
									Set tMPNodes = tXMLFile.SelectNodes("//dictionaries/measuring-points/measuring-point") 'SOURCE-DICT
								Case Else:	Set tMPNodes = Nothing							
							End Select							
							If tMPNodes Is Nothing Then
								WScript.Echo "ANOMALY! [80000 version unknown]"
								WScript.Quit
							End If
							If tMPNodes.Length = 0 Then
								WScript.Echo "ANOMALY! [No MPoints located] V:" & tVersion
								WScript.Quit
							End If
							WScript.Echo tMPChannelList.Length & ":" & tMPNodes.Length
							'WScript.Quit
							'P5  // MPoints inject
							'WScript.Echo tMPNodes.Length & " V:" & tVersion
							For Each tMPNode In tMPNodes
								If tMPNode.NodeName = "measuringpoint" Then
									Set tNewMPNode = tFNode.AppendChild(inXMLFrame.CreateElement("measuringpoint"))
									tNewMPNode.SetAttribute "code", tMPNode.getAttribute("code")
									tNewMPNode.SetAttribute "name", tMPNode.getAttribute("name")
									tNewMPNode.SetAttribute "objectname", tMPNode.getAttribute("objectname")
									tValue = tMPNode.getAttribute("voltage")
									If tValue <> 0 Then: tNewMPNode.SetAttribute "voltage", tValue
									For Each tCHNode In tMPNode.ChildNodes
										If tCHNode.NodeName = "measuringchannel" Then
											Set tNewCHNode = tNewMPNode.AppendChild(inXMLFrame.CreateElement("measuringchannel"))
											'WScript.Echo "ANOMALY! [] File:" & inFile.Name
											tNewCHNode.SetAttribute "code", tCHNode.getAttribute("code")
											tNewCHNode.SetAttribute "devicemodel", tCHNode.getAttribute("devicemodel")
											tNewCHNode.SetAttribute "integral", tCHNode.getAttribute("is_integral_device")
											tNewCHNode.SetAttribute "period", tCHNode.getAttribute("period")
										End If
									Next
								End If
								'v4
								If tMPNode.NodeName = "measuring-point" Then
									Set tNewMPNode = tFNode.AppendChild(inXMLFrame.CreateElement("measuringpoint"))
									tNewMPNode.SetAttribute "code", tMPNode.getAttribute("ats-code")									
									tNewMPNode.SetAttribute "voltage", tMPNode.getAttribute("point-voltage")
									tNewMPNode.SetAttribute "pointtype", tMPNode.getAttribute("measuring-point-type")
									'get name node
									tXPathString = "child::name"
									Set tSNode = tMPNode.SelectSingleNode(tXPathString)
									tTempValue = tSNode.getAttribute("location-description")
									If IsNull(tTempValue) Then: tTempValue = tSNode.getAttribute("connection-name")
									tNewMPNode.SetAttribute "name", tTempValue
									tNewMPNode.SetAttribute "objectname", tSNode.getAttribute("power-object-name")
									'get measuring device
									tXPathString = "child::measuring-device"
									Set tSNode = tMPNode.SelectSingleNode(tXPathString)
									tNewMPNode.SetAttribute "deviceid", tSNode.getAttribute("device-modification-code")
									'get channels
									For Each tCHNode In tSNode.ChildNodes
										If tCHNode.NodeName = "measuring-channel" Then
											Set tNewCHNode = tNewMPNode.AppendChild(inXMLFrame.CreateElement("measuringchannel"))											
											tNewCHNode.SetAttribute "code", tCHNode.getAttribute("ats-code")
											tNewCHNode.SetAttribute "period", tCHNode.getAttribute("period")
										End If
									Next
								End If
							Next
							' PX // SAVE
							tAREAInjected = tAREAInjected + 1
							fSaveXMLConfigChanges gXMLFramePathLock, gXMLFrame
						End If						
					End If
				End If
			Next
		End If
	Next
	' 00 \\
	' 00 \\
End Sub

'======= // START
tAREAReaded = 0
tAREAInjected = 0

Set gFSO = CreateObject("Scripting.FileSystemObject")
gScriptFileName = Wscript.ScriptName
gScriptPath = gFSO.GetParentFolderName(WScript.ScriptFullName)

Set gWSO = CreateObject("WScript.Shell")

gXMLFilePathA = gWSO.ExpandEnvironmentStrings("%HOMEPATH%") & "\GTPCFG"
gXMLBasisPathLock = gXMLFilePathA
gXMLFramePathLock = gXMLFilePathA
If Not fGetXMLConfig(gXMLBasisPathLock, gXMLBasis, "Basis.xml", "BASIS") Then 
	WScript.Echo "Не удалось загрузить XML конфиг BASIS (или он не найден)!"
	fQuitScript
End If
If Not fGetXMLConfig(gXMLFramePathLock, gXMLFrame, "Frame.xml", "FRAME") Then 
	gXMLFramePathLock = gXMLFilePathA
	If Not(fXMLFrameConfigCreate(gXMLFrame, gXMLFramePathLock)) Then 
		WScript.Echo "Не удалось загрузить XML конфиг FRAME (или он не найден)!"
		fQuitScript
	End If
End If

If WScript.Arguments.Length > 0 Then
	For Each tFilePath in WScript.Arguments
		If gFSO.FileExists(tFilePath) Then
			Set tFile = gFSO.GetFile(tFilePath)
			fFileDataExtract tFile, gXMLFrame, gXMLBasis
		End If
	Next
Else
	For Each tFile in gFSO.GetFolder(gScriptPath).Files
		fFileDataExtract tFile, gXMLFrame, gXMLBasis
	Next
End If


'statistic
'tAREAReaded = 0
'tAREAInjected = 0



WScript.Echo "Прочитано:" & vbTab & tAREAReaded & vbCrLf & "Инъекций:" & vbTab & tAREAInjected
fQuitScript