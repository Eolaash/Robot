Attribute VB_Name = "M_CFGDrop"
'M_CFGDrop version 1 (25-01-2019)
Option Explicit

Private Const cnstModuleName = "M_CFGDrop"
Private Const cnstModuleVersion = 1
Private Const cnstModuleDate = "25-01-2019"

Dim gLocalInit
Dim gDropSubFolder
Dim gDropMainPathList()

Private Function fGetLogTag(inTagText)
    fGetLogTag = cnstModuleName & "." & inTagText
End Function

Private Function fLocalInit(Optional inForceInit As Boolean = False)
Dim tLogTag
    fLocalInit = False
    tLogTag = fGetLogTag("CFGDropINI")
    uCDebugPrint tLogTag, 0, "Инициализация > Имя: " & cnstModuleName & "; Версия: " & cnstModuleVersion & "; Редакция: " & cnstModuleDate
    If inForceInit Or Not gLocalInit Then
        '//some actions
        ReDim gDropMainPathList(2)
        gDropMainPathList(0) = "Z:"
        gDropMainPathList(1) = "\\srv1\share\common"
        gDropMainPathList(2) = "\\192.168.32.4\share\common"
        gDropSubFolder = "Config"
    End If
    fLocalInit = True
End Function

Public Sub fCFGDropMain()
Dim tDropPath
    If Not fLocalInit Then: Exit Sub
    If Not fGetDropPath(tDropPath) Then: Exit Sub
    fDropFileProcess tDropPath
End Sub

Private Sub fDropFileProcess(inDropPath)
    fXMLDropper inDropPath, gXMLBasis
    fXMLDropper inDropPath, gXMLCalendar
    fXMLDropper inDropPath, gXMLDictionary
End Sub

Private Sub fXMLDropper(inDropPath, inCFG As TXMLConfigFile)
Dim tLogTag, tTargetPath, tXMLDoc, tXPathString, tNodes, tNode, tText, tCommentMD5, tTargetMD5, tOriginalMD5, tComment
' 00 // Предопределения
    tLogTag = fGetLogTag("XMLDropper")
' 00 // Проверки состояния источника
    If Not inCFG.Active Then: Exit Sub
    If inCFG.XML Is Nothing Then: Exit Sub
' 00 // Поиск принимающего файла
    Set tXMLDoc = CreateObject("Msxml2.DOMDocument.6.0")
    tXMLDoc.ASync = False
    tTargetPath = inDropPath & "\" & inCFG.Name
    If uFileExists(tTargetPath) Then
        tXMLDoc.Load tTargetPath
        If tXMLDoc.parseError.ErrorCode = 0 Then
            tXPathString = "//comment()"
            Set tNodes = tXMLDoc.SelectNodes(tXPathString)
            For Each tNode In tNodes
                tText = tNode.Text
                tCommentMD5 = fExtractMD5fromText(tText)
                tTargetMD5 = fGetHashString(tXMLDoc)
                If tTargetMD5 = tCommentMD5 Then
                    tOriginalMD5 = fGetHashString(inCFG.XML)
                    If tOriginalMD5 = tTargetMD5 Then
                        Set tXMLDoc = Nothing
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If
' 00 // Производим копирование данных
    If uDeleteFile(tTargetPath) Then
        tXMLDoc.LoadXML (inCFG.XML.XML)
        tCommentMD5 = "MD5:" & fGetHashString(tXMLDoc)
        Set tComment = tXMLDoc.CreateComment(tCommentMD5)
        tXMLDoc.InsertBefore tComment, tXMLDoc.DocumentElement
        tXMLDoc.Save tTargetPath
        uCDebugPrint tLogTag, 0, "Произведено обновление сопряженного конфига " & inCFG.ClassTag & ": " & tTargetPath
    Else
        uCDebugPrint tLogTag, 0, "Неудача при обновлении сопряженного конфига " & inCFG.ClassTag & ": " & tTargetPath
    End If
    Set tXMLDoc = Nothing
End Sub

Private Function fGetDropPath(outDropPath)
Dim tPath, tSubFolders, tSubFolder, tTotalPath, tIsPathOK
    fGetDropPath = True
    outDropPath = vbNullString
    tSubFolders = Split(gDropSubFolder, "\")
    For Each tPath In gDropMainPathList
        If uFileExists(tPath) Then
            tIsPathOK = True
            tTotalPath = tPath
            'fix
            If Right(tTotalPath, 1) = "\" Then: tTotalPath = Left(tTotalPath, Len(tTotalPath) - 1)
            For Each tSubFolder In tSubFolders
                tTotalPath = tTotalPath & "\" & tSubFolder
                If Not uFolderCreate(tTotalPath) Then
                    tIsPathOK = False
                    Exit For
                End If
            Next
            'result
            If tIsPathOK Then
                outDropPath = tTotalPath
                Exit Function
            End If
        End If
    Next
    fGetDropPath = False
End Function
