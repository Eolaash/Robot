Attribute VB_Name = "XMLUtils"
'XML UTILS module V1
'18.07.2018
Option Explicit

Public Function fGetChildNodeByName(inRootNode, inChildName, Optional inForceCreate As Boolean = False)
Dim tChildCount, tIndex, tXML
    Set fGetChildNodeByName = Nothing
    If inRootNode Is Nothing Then: Exit Function
    tChildCount = inRootNode.ChildNodes.Length
    For tIndex = 0 To tChildCount - 1
        If inRootNode.ChildNodes(tIndex).BaseName = inChildName Then
            Set fGetChildNodeByName = inRootNode.ChildNodes(tIndex)
            Exit Function
        End If
    Next
    'Forced child creation if not found
    If inForceCreate Then
        Set tXML = CreateObject("Msxml2.DOMDocument.6.0")
        Set fGetChildNodeByName = inRootNode.AppendChild(tXML.CreateElement(inChildName))
        Set tXML = Nothing
    End If
End Function


Public Function fGetNodeSafe(inXML, Optional inIndex = 0, Optional inMultiNodes = False, Optional inXPathString = vbNullString, Optional inSilent = False)
    Dim tTempNode, tTypeName, tXMLBaseNode
    
    Set fGetNodeSafe = Nothing
    tTypeName = TypeName(inXML)
    
    If Not (tTypeName = "IXMLDOMSelection" Or tTypeName = "IXMLDOMElement" Or tTypeName = "DOMDocument60") Then: Exit Function
    
    On Error Resume Next
        If tTypeName = "IXMLDOMSelection" Then
            Set tXMLBaseNode = inXML(inIndex)
        Else
            Set tXMLBaseNode = inXML
        End If
    
        If inXPathString <> vbNullString Then
            If inMultiNodes Then
                Set tTempNode = tXMLBaseNode.SelectNodes(inXPathString)
            Else
                Set tTempNode = tXMLBaseNode.SelectSingleNode(inXPathString)
            End If
        Else
            Set tTempNode = tXMLBaseNode
        End If
        
        If Err.Number <> 0 Then
            Set tTempNode = Nothing
            Set tXMLBaseNode = Nothing
            If Not inSilent Then: Debug.Print "fGetNodeSafe function error: object(" & Err.Source & "): " & Err.Description
            Exit Function
        End If
        
        Set fGetNodeSafe = tTempNode
        Set tTempNode = Nothing
        Set tXMLBaseNode = Nothing
    On Error GoTo 0
    
End Function

Public Function fGetAttr(inNode, inAttributeName, outValue, Optional inDefaultValue = vbNullString)
    Dim tValue
    
    fGetAttr = False
    outValue = inDefaultValue
    
    On Error Resume Next
    
        tValue = inNode.GetAttribute(inAttributeName) 'проичтаем аттрибут
        If IsNull(tValue) Then: Exit Function
        If Err.Number <> 0 Then: Exit Function
        
    On Error GoTo 0
    
    outValue = tValue
    fGetAttr = True
End Function

Public Sub fRemoveChilds(inNode)
    While inNode.ChildNodes.Length > 0
        inNode.RemoveChild inNode.LastChild
    Wend
End Sub

Public Sub fAX_TestA()
Dim tResult
    If Not fConfiguratorInit Then: Exit Sub
    If Not fXMLSmartUpdate("2,3") Then: Exit Sub
    tResult = fGetHashString(gXMLFrame.XML)
    Debug.Print "MD5: " & tResult
End Sub

Public Sub fAX_TestB()
Dim tEnergyObject As New CEnergyAPI
    'Set tEnergyObject = New CEnergyAPI
    Debug.Print "================="
    tEnergyObject.PrintLog = True
    If tEnergyObject.IsActive Then
        tEnergyObject.SetCredentials "engtpp1005", "uygWcBLg"
        If Not tEnergyObject.SendRequest(2, DateSerial(2019, 1, 17), 1, "PBELKAM6:1225") Then
        'If Not tEnergyObject.SendRequest(1, DateSerial(2019, 1, 1), 4) Then
        'If Not tEnergyObject.SendRequest(1, DateSerial(2018, 12, 2), 4) Then
        'If Not tEnergyObject.SendRequest(0) Then
            Debug.Print tEnergyObject.ErrorText
        Else
            'MsgBox tEnergyObject.ResponseText
            If Not tEnergyObject.ReadResponse Then
                Debug.Print tEnergyObject.ResponseSOAPText
                Debug.Print "FAIL"
            End If
        End If
    Else
        Debug.Print tEnergyObject.ErrorText
    End If
    Set tEnergyObject = Nothing
End Sub

Public Function fExtractMD5fromText(inText, Optional inSplitter = ":")
Dim tElements
    fExtractMD5fromText = vbNullString
    If inText = vbNullString Then: Exit Function
    tElements = Split(inText, inSplitter)
    If UBound(tElements) <> 1 Then: Exit Function
    If tElements(0) <> "MD5" Then: Exit Function
    fExtractMD5fromText = tElements(1)
End Function

'Extracting MD5 HASH for ROOT NODE
Public Function fGetHashString(inXML)
Dim tEncodning, tMD5, tByteString, tRootNode, tByteIndex, tHi, tLo, tBuf, tCharHi, tCharLo
    fGetHashString = vbNullString
    If inXML Is Nothing Then: Exit Function
    Set tRootNode = inXML.DocumentElement
    Set tEncodning = CreateObject("System.Text.UTF8Encoding")
    Set tMD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    'Debug.Print "1LN: " & Len(tRootNode.XML)
    tByteString = tMD5.ComputeHash_2(tEncodning.GetBytes_4(tRootNode.XML))
    'Debug.Print "2LN: " & Len(tByteString)
    For tByteIndex = 1 To LenB(tByteString)
        tBuf = AscB(MidB(tByteString, tByteIndex, 1))
        tLo = tBuf Mod 16: tHi = (tBuf - tLo) / 16
        If tHi > 9 Then tCharHi = Chr(Asc("a") + tHi - 10) Else tCharHi = Chr(Asc("0") + tHi)
        If tLo > 9 Then tCharLo = Chr(Asc("a") + tLo - 10) Else tCharLo = Chr(Asc("0") + tLo)
        fGetHashString = fGetHashString & tCharHi & tCharLo
    Next
    Set tEncodning = Nothing
    Set tMD5 = Nothing
End Function
