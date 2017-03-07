VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omQASPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------
' Class Module  omQASPro
' File:         Class1.cls
' Author:       IVW
' Date:         10-10-2001
' Purpose:      Retrieves an address or list of possible addresses from
'               the third party QAS
'
' History:
'   AQR             AUTHOR      DESCRIPTION
'   ---             ------      -----------
'   None            IVW         01/10/2001  Created.
'   None            IVW         02/10/2001  Added FLATNUMBER search.
'   None            IVW         10/10/2001  Added NAME
'   None            IVW         22/10/2001  Returned explicit XML fields.
'                                           DO NOT CHANGES QASADDRESS.INI now.
'   None            IVW         14/12/2001  Added new attribute PICKLIST. If set 'No' or 'NO' or 'N' then a picklist is returned, it returns "NOT FOUND" instead.
'   None            IVW         17/12/2001  Added new attribute VALIDATIONTYPE. If set to 1 it will ensure that a postcode AND House Name AND/OR House No have been entered. If it doesn't have this it returns an error. Add new types as you feel fit.
'   SYS4190         INR         07/03/2002  In the call to OneAddressItemPickList, can't always guarentee that the second item is a "|st:"
'------------------------------------------------------------------------
' Note
' This is generated as .EXE because the QAS Server API software
' is not threadsafe.
Option Explicit
Private Const INVALID_XML_ERROR = vbObjectError + 1
Private Const QAS_ERROR = vbObjectError + 2
Private Const INVALID_SEARCH_ERROR = vbObjectError + 3
' use enumerated values in case new search XML are requested later, also more readable
Public Enum AddressArray
    FLATNUMBER = 0
    HOUSENUMBER
    HOUSENAME
    ADDRESSLINE1
    ADDRESSLINE2
    ADDRESSLINE3
    ADDRESSLINE4
    COUNTRY
    POSTCODE
End Enum
Public Function FindAddressWithKey(ByVal p_strXML As String) As String
Dim lStartUpErrorCode As Long
Dim l_xmlDom As FreeThreadedDOMDocument40
Dim strError As String
Dim strKey As String
Dim strBuffer As String
Dim bContinue As Boolean
Dim lSearchErr As Long
Dim bPickList As Boolean
On Error GoTo RequestKeyErrortrap
    bPickList = True
    Set l_xmlDom = New FreeThreadedDOMDocument40
    l_xmlDom.async = False
    l_xmlDom.setProperty "NewParser", True
    l_xmlDom.validateOnParse = False
    l_xmlDom.loadXML p_strXML
    If l_xmlDom.parseError.errorCode <> 0 Then
        Err.Raise l_xmlDom.parseError.errorCode, "RequestKey", "The XML string could not be parsed - parse error = " & l_xmlDom.parseError.reason
    Else
        bContinue = ExtractKeyElements(l_xmlDom, strKey)
        If (bContinue) Then
            bContinue = CreateResponseNode(l_xmlDom)
        End If
    End If
    If (bContinue) Then
        lStartUpErrorCode = StartUp
        If (lStartUpErrorCode < 0) Then
            Err.Raise QAS_ERROR, "RequestKey", "Unable to Start up QAS Address Pro - QASRVEA.DLL - Error(" & Str(lStartUpErrorCode) & ")"
        End If
        bContinue = SearchQASPro(strKey, "", strBuffer, l_xmlDom)
        Call ShutDown
        If (bContinue) Then
            StoreAsXML strBuffer, bPickList, l_xmlDom
        End If
    End If
    FindAddressWithKey = l_xmlDom.xml
    Set l_xmlDom = Nothing
Exit Function:
RequestKeyErrortrap:
    CreateErrorXML l_xmlDom, Err.Number, Err.Source, Err.Description
    FindAddressWithKey = l_xmlDom.xml
    Set l_xmlDom = Nothing
End Function
Private Function CreateResponseNode(ByRef p_xmlDom As FreeThreadedDOMDocument40) As Boolean
Dim l_xmlResponseElementNode As IXMLDOMElement
Dim l_xmlNode As IXMLDOMNode
Dim l_xmlNode3 As IXMLDOMNode
On Error GoTo CreateResponseNodeErrortrap
    
    Set l_xmlResponseElementNode = p_xmlDom.createElement("RESPONSE")
    Set l_xmlNode = p_xmlDom.selectSingleNode("/REQUEST")
    If (l_xmlNode Is Nothing) Then
        Err.Raise INVALID_XML_ERROR, "CreateResponseNode", "No Request element present in Input XML"
    End If
    Set l_xmlNode3 = p_xmlDom.replaceChild(l_xmlResponseElementNode, l_xmlNode)
Set l_xmlResponseElementNode = Nothing
Set l_xmlNode = Nothing
Set l_xmlNode3 = Nothing
CreateResponseNode = True
Exit Function:
CreateResponseNodeErrortrap:
CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description
Set l_xmlResponseElementNode = Nothing
Set l_xmlNode = Nothing
Set l_xmlNode3 = Nothing
CreateResponseNode = False
End Function
Public Function FindAddress(ByVal p_strXML As String) As String
Dim lStartUpErrorCode As Long
Dim l_xmlDom As FreeThreadedDOMDocument40
Dim strError As String
Dim strBuffer As String
Dim bContinue As Boolean
Dim bPickList As Boolean
Dim strAddressElements(11) As String ' should be sized to all max search elements.
bContinue = False
On Error GoTo RequestErrortrap
    bPickList = True
    Set l_xmlDom = New FreeThreadedDOMDocument40
    l_xmlDom.async = False
        
    l_xmlDom.setProperty "NewParser", True
    l_xmlDom.validateOnParse = False
    l_xmlDom.loadXML p_strXML
    If l_xmlDom.parseError.errorCode <> 0 Then
        Err.Raise l_xmlDom.parseError.errorCode, "Request", "The XML string could not be parsed - XML parse error = " & l_xmlDom.parseError.reason
    Else
        bContinue = ExtractAddressElements(l_xmlDom, strAddressElements, bPickList)
        If (bContinue) Then
            bContinue = ValidateAddressElements(l_xmlDom, strAddressElements)
        End If
        If (bContinue) Then
            bContinue = CreateResponseNode(l_xmlDom)
        End If
    End If
    If (bContinue) Then
        
        lStartUpErrorCode = StartUp
        If (lStartUpErrorCode < 0) Then
            Err.Raise QAS_ERROR, "Request", "Unable to Start up QAS Address Pro - QASRVEA.DLL - Error(" & Str(lStartUpErrorCode) & ")"
        End If
        bContinue = SearchQASPro("", strAddressElements, strBuffer, l_xmlDom)
        Call ShutDown
        If (bContinue) Then
            StoreAsXML strBuffer, bPickList, l_xmlDom
        End If
    End If
    FindAddress = l_xmlDom.xml
    Set l_xmlDom = Nothing
Exit Function:
RequestErrortrap:
    CreateErrorXML l_xmlDom, Err.Number, Err.Source, Err.Description
    FindAddress = l_xmlDom.xml
    Set l_xmlDom = Nothing
End Function
Private Function ValidateAddressElements(ByRef p_xmlDom As FreeThreadedDOMDocument40, _
    ByRef p_strAddressLines As Variant) As Boolean
Dim l_strNodeValue As String
Dim l_xmlNode As IXMLDOMNode
Dim l_xmlAttribute As IXMLDOMAttribute
Dim bOk As Boolean
bOk = True ' assume everything is okay
On Error GoTo ValidateAddressElementsError
    
    Set l_xmlNode = p_xmlDom.selectSingleNode("/REQUEST/ADDRESS")
    If (l_xmlNode Is Nothing) Then
        Err.Raise INVALID_XML_ERROR, "ExtractAddressElements", "ADDRESS Element missing from REQUEST XML"
    End If
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("VALIDATIONTYPE")
    If Not l_xmlAttribute Is Nothing Then
        Select Case l_xmlAttribute.Text
            Case "1"
                ' must have a post code and/or a house name and number.
                If (p_strAddressLines(8) = "") Then
                     bOk = False
                ElseIf (p_strAddressLines(1) = "" And p_strAddressLines(2) = "") Then
                     bOk = False
                End If
        End Select
        If (bOk = False) Then
            Err.Raise QAS_ERROR, "ValidateAddressElements", "Error in XML - Insuffucient Search Criteria"
        End If
    End If
    Set l_xmlNode = Nothing
    Set l_xmlAttribute = Nothing
    ValidateAddressElements = bOk
Exit Function:
ValidateAddressElementsError:
    
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description
    Set l_xmlNode = Nothing
    Set l_xmlAttribute = Nothing
    ValidateAddressElements = False
End Function
Private Function ExtractAddressElements(ByRef p_xmlDom As FreeThreadedDOMDocument40, _
    ByRef p_strAddressLines As Variant, _
    ByRef p_bPickList) As Boolean
            
    'read input XML string.
Dim l_strNodeValue As String
Dim l_xmlNode As IXMLDOMNode
Dim l_xmlAttribute As IXMLDOMAttribute
On Error GoTo ExtractAddressElements
    
    Set l_xmlNode = p_xmlDom.selectSingleNode("/REQUEST/ADDRESS")
    If (l_xmlNode Is Nothing) Then
        Err.Raise INVALID_XML_ERROR, "ExtractAddressElements", "ADDRESS Element missing from REQUEST XML"
    End If
       
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("PICKLIST")
    If Not l_xmlAttribute Is Nothing Then
        Select Case l_xmlAttribute.Text
            Case "N", "No", "NO", "no", "nO", "F", "False", "FALSE"
                p_bPickList = False
            End Select
    End If
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("FLATNUMBER")
    If Not l_xmlAttribute Is Nothing Then
        p_strAddressLines(0) = l_xmlAttribute.Text
    Else
        p_strAddressLines(0) = ""
    End If
    Trim$ (p_strAddressLines(0))
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("HOUSENUMBER")
    If Not l_xmlAttribute Is Nothing Then
        p_strAddressLines(1) = l_xmlAttribute.Text
    Else
        p_strAddressLines(1) = ""
    End If
    Trim$ (p_strAddressLines(1))
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("NAME")
    If Not l_xmlAttribute Is Nothing Then
        p_strAddressLines(2) = l_xmlAttribute.Text
    Else
        p_strAddressLines(2) = ""
    End If
    Trim$ (p_strAddressLines(2))
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("ADDRESSLINE1")
    If Not l_xmlAttribute Is Nothing Then
        p_strAddressLines(3) = l_xmlAttribute.Text
    Else
        p_strAddressLines(3) = ""
    End If
    Trim$ (p_strAddressLines(3))
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("ADDRESSLINE2")
    If Not l_xmlAttribute Is Nothing Then
        p_strAddressLines(4) = l_xmlAttribute.Text
    Else
        p_strAddressLines(4) = ""
    End If
    Trim$ (p_strAddressLines(4))
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("ADDRESSLINE3")
    If Not l_xmlAttribute Is Nothing Then
        p_strAddressLines(5) = l_xmlAttribute.Text
    Else
        p_strAddressLines(5) = ""
    End If
    Trim$ (p_strAddressLines(5))
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("ADDRESSLINE4")
    If Not l_xmlAttribute Is Nothing Then
        p_strAddressLines(6) = l_xmlAttribute.Text
    Else
        p_strAddressLines(6) = ""
    End If
    Trim$ (p_strAddressLines(6))
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("COUNTRY")
    If Not l_xmlAttribute Is Nothing Then
        p_strAddressLines(7) = l_xmlAttribute.Text
    Else
        p_strAddressLines(7) = ""
    End If
    Trim$ (p_strAddressLines(7))
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("POSTCODE")
    If Not l_xmlAttribute Is Nothing Then
        p_strAddressLines(8) = l_xmlAttribute.Text
    Else
        p_strAddressLines(8) = ""
    End If
    Trim$ (p_strAddressLines(8))
    Set l_xmlNode = Nothing
    Set l_xmlAttribute = Nothing
    ExtractAddressElements = True
Exit Function:
ExtractAddressElements:
    
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description
    Set l_xmlNode = Nothing
    Set l_xmlAttribute = Nothing
    ExtractAddressElements = False
End Function
Private Function CreateErrorXML(ByRef p_xmlDom As FreeThreadedDOMDocument40, ByRef lNumber As Long, ByRef strSource As String, ByRef strDesc As String)
    
Dim l_xmlResponseElement As IXMLDOMElement
Dim l_xmlRequestNode As IXMLDOMNode
Dim l_xmlResponseNode As IXMLDOMNode
Dim l_xmlErrorElement As IXMLDOMElement
Dim l_xmlErrorElementNode As IXMLDOMElement
Dim l_xmlWorkingNode As IXMLDOMNode
Dim l_OldNode As IXMLDOMNode
On Error GoTo CreateErrorXMLErrortrap
    ' firstly, a catch all for ensuring no errors created thus far on response xml
    Set l_OldNode = p_xmlDom.selectSingleNode("/RESPONSE/ERROR")
    If (l_OldNode Is Nothing) Then
        Set l_xmlResponseNode = p_xmlDom.selectSingleNode("/RESPONSE")
        If (l_xmlResponseNode Is Nothing) Then
            
            Set l_xmlResponseElement = p_xmlDom.createElement("RESPONSE")
            Set l_xmlRequestNode = p_xmlDom.selectSingleNode("/REQUEST")
            If (l_xmlRequestNode Is Nothing) Then
                p_xmlDom.appendChild l_xmlResponseElement
            Else
                Set l_OldNode = p_xmlDom.replaceChild(l_xmlResponseElement, l_xmlRequestNode)
            End If
        Else
            Set l_xmlResponseElement = l_xmlResponseNode.selectSingleNode("/RESPONSE")
        End If
        l_xmlResponseElement.setAttribute "TYPE", "APPERR"
        Set l_xmlErrorElement = p_xmlDom.createElement("ERROR")
        Set l_xmlWorkingNode = l_xmlResponseElement.appendChild(l_xmlErrorElement)
        Set l_xmlErrorElementNode = p_xmlDom.createElement("NUMBER")
        l_xmlErrorElementNode.Text = Str(lNumber)
        l_xmlWorkingNode.appendChild l_xmlErrorElementNode
        Set l_xmlErrorElementNode = p_xmlDom.createElement("SOURCE")
        l_xmlErrorElementNode.Text = strSource
        l_xmlWorkingNode.appendChild l_xmlErrorElementNode
        Set l_xmlErrorElementNode = p_xmlDom.createElement("DESCRIPTION")
        l_xmlErrorElementNode.Text = strDesc
        l_xmlWorkingNode.appendChild l_xmlErrorElementNode
    End If
    Set l_xmlResponseElement = Nothing
    Set l_xmlRequestNode = Nothing
    Set l_xmlResponseNode = Nothing
    Set l_xmlErrorElement = Nothing
    Set l_xmlErrorElementNode = Nothing
    Set l_xmlWorkingNode = Nothing
    Set l_OldNode = Nothing
Exit Function:
CreateErrorXMLErrortrap:
        
    Set l_xmlResponseElement = Nothing
    Set l_xmlRequestNode = Nothing
    Set l_xmlResponseNode = Nothing
    Set l_xmlErrorElement = Nothing
    Set l_xmlErrorElementNode = Nothing
    Set l_xmlWorkingNode = Nothing
    Set l_OldNode = Nothing
        
End Function
Private Function StoreAsXML(ByRef strBuffer As String, ByVal bPickList As Boolean, ByRef p_xmlDom As FreeThreadedDOMDocument40) As Boolean
    
Dim lIndex As Long
Dim lLastIndex As Long
Dim lLineIndex As Long
Dim lLastLineIndex As Long
Dim sResponseType As String
Dim sError As String
Dim lNoItems As Long
Dim lCount As Long
Dim sLine As String
Dim sPickHeader As String
Dim sOutLine As String
Dim lElementCount As Long
Dim lElementNo As Long
Dim l_xmlResponseElementNode As IXMLDOMElement
Dim l_xmlElementNode As IXMLDOMElement
Dim l_xmlElementNode2 As IXMLDOMElement
Dim sKey As String
Dim str_xmlAttribute As String
Dim iLineCount As Integer
Dim lErrorNo As Integer
On Error GoTo StoreAsXMLErrortrap
    
    If Len(strBuffer) = 0 Then
        Exit Function
    End If
    lIndex = 0
    lLastIndex = lIndex + 1
    lIndex = InStr(lLastIndex, strBuffer, "|")
    sResponseType = Mid$(strBuffer, lLastIndex, lIndex - lLastIndex)
    Set l_xmlResponseElementNode = p_xmlDom.selectSingleNode("/RESPONSE")
    Select Case sResponseType
    Case "ERROR"
        lLastIndex = lIndex + 1
        lIndex = InStr(lLastIndex, strBuffer, "|")
        sError = Mid$(strBuffer, lLastIndex, lIndex - lLastIndex)
        lLastIndex = lIndex + 1
        lIndex = InStr(lLastIndex, strBuffer, "|")
        lErrorNo = Val(Mid$(strBuffer, lLastIndex, lIndex - lLastIndex))
        If (lErrorNo <> qaerr_SRVNOADDRESSMATCH) Then
            Err.Raise QAS_ERROR, "StoreAsXML", sError
        Else
            Set l_xmlElementNode = p_xmlDom.createElement("ADDRESSDATA")
            l_xmlElementNode.setAttribute "QASRETURN", "NOMATCH"
            l_xmlResponseElementNode.setAttribute "TYPE", "SUCCESS"
            l_xmlResponseElementNode.appendChild l_xmlElementNode
        End If
    Case "PICKLIST"
        Select Case bPickList
            
        Case False
            
        ' if PICKLIST is set to 'No' in the REQUEST XML, do not return a pick list - ever. flag as not found
        ' this was a specific requirement for Omiga Phase I because if they send a
        ' they do not use picklists. just assume it has not been found.
            
            Set l_xmlElementNode = p_xmlDom.createElement("ADDRESSDATA")
            l_xmlElementNode.setAttribute "QASRETURN", "NOMATCH"
            l_xmlResponseElementNode.setAttribute "TYPE", "SUCCESS"
            l_xmlResponseElementNode.appendChild l_xmlElementNode
        Case True
            l_xmlResponseElementNode.setAttribute "TYPE", "SUCCESS"
            lLastIndex = lIndex + 1
            lIndex = InStr(lLastIndex, strBuffer, "|")
            sPickHeader = Mid$(strBuffer, lLastIndex, lIndex - lLastIndex)
            lLastIndex = lIndex + 1
            lIndex = InStr(lLastIndex, strBuffer, "|")
            lNoItems = Val(Mid$(strBuffer, lLastIndex, lIndex - lLastIndex))
            Set l_xmlElementNode = p_xmlDom.createElement("ADDRESSDATA")
            l_xmlElementNode.setAttribute "QASRETURN", "PICKLIST"
            l_xmlElementNode.setAttribute "LISTLINES", lNoItems
            l_xmlElementNode.setAttribute "LISTHEADER", sPickHeader
            l_xmlResponseElementNode.appendChild l_xmlElementNode
            lCount = 0
            lElementNo = 0
            ' Store results in XML.
            Do
                lLastIndex = lIndex + 1
                lIndex = InStr(lLastIndex, strBuffer, "|")
                sLine = Mid$(strBuffer, lLastIndex, lIndex - lLastIndex)
                ' Only do this for the first line in each list
                If lElementNo = 0 Then
                    lLineIndex = 0
                    ' Count the number of elements in the first line
                    Do
                        lLastLineIndex = lLineIndex + 1
                        lLineIndex = InStr(lLastLineIndex, sLine, "~")
                        lElementNo = lElementNo + 1
                    Loop While lLineIndex > 0
                End If
                lElementCount = 0
                lLineIndex = 0
                sOutLine = ""
                ' Loop through all ~ seperated elements
                Do
                    Select Case lElementCount
                        Case 0
                        Case 1, 2, 3, 4
                            sOutLine = sOutLine & ","
                    End Select
                    lLastLineIndex = lLineIndex + 1
                    lLineIndex = InStr(lLastLineIndex, sLine, "~")
                    If lLineIndex > 0 Then
                        sOutLine = sOutLine & Mid$(sLine, lLastLineIndex, lLineIndex - lLastLineIndex)
                    End If
                    lElementCount = lElementCount + 1
                Loop While lElementCount < lElementNo
                sOutLine = sOutLine & Right$(sLine, Len(sLine) - lLastLineIndex + 1)
                Set l_xmlElementNode2 = p_xmlDom.createElement("ITEM")
                l_xmlElementNode2.setAttribute "TEXT", sOutLine
                lLastIndex = lIndex + 1
                lIndex = InStr(lLastIndex, strBuffer, "|")
                sKey = Mid$(strBuffer, lLastIndex, lIndex - lLastIndex)
                l_xmlElementNode2.setAttribute "KEY", sKey
                lCount = lCount + 1
                l_xmlElementNode.appendChild l_xmlElementNode2
            Loop While lIndex > lLastIndex And lCount < lNoItems
        End Select
    Case "ADDRESS"
        l_xmlResponseElementNode.setAttribute "TYPE", "SUCCESS"
        lLastIndex = lIndex + 1
        lIndex = InStr(lLastIndex, strBuffer, "|")
        lNoItems = Val(Mid$(strBuffer, lLastIndex, lIndex - lLastIndex))
        lCount = 0
        Set l_xmlElementNode = p_xmlDom.createElement("ADDRESSDATA")
        l_xmlElementNode.setAttribute "QASRETURN", "ADDRESS"
        l_xmlResponseElementNode.appendChild l_xmlElementNode
        Do
            lLastIndex = lIndex + 1
            lIndex = InStr(lLastIndex, strBuffer, "|")
            str_xmlAttribute = ""
            lCount = lCount + 1
            If lIndex > 0 Then
                
                ' select the address item from the returned buffer
                sOutLine = Mid$(strBuffer, lLastIndex, lIndex - lLastIndex)
                        
                ' special lines...
'Line1=W60 ORGA
'Line2=W10 SUBP
'Line3=W60 NAME
'Line4=W10 NUMB
'Line5=W50 DEPT
'Line6=W50 THOR
'Line7=W50 DDPL
'Line8=W50 DEPL
'Line9=W50 TOWN
'Line10=W50 COUN
'Line11=W8 XPOST
'Line12=20 POBO
                
                
                Select Case lCount
                    Case 1
                        str_xmlAttribute = "ORGANISATIONNAME"
                    Case 2
                        str_xmlAttribute = "SUBBUILDINGNAME"
                    Case 3
                        str_xmlAttribute = "BUILDINGORHOUSENAME"
                    Case 4
                        str_xmlAttribute = "BUILDINGORHOUSENUMBER"
                    Case 5
                        str_xmlAttribute = "DEPENDENTTHOROUGHFARENAME"
                    Case 6
                        str_xmlAttribute = "THOROUGHFARENAME"
                    Case 7
                        str_xmlAttribute = "DOUBLEDEPENDENTLOCALITY"
                    Case 8
                        str_xmlAttribute = "DEPENDENTLOCALITY"
                    Case 9
                        str_xmlAttribute = "POSTTOWN"
                    Case 10
                        str_xmlAttribute = "COUNTY"
                    Case 11
                        str_xmlAttribute = "POSTCODE"
                    Case 12
                        str_xmlAttribute = "POBOXDETAILS"
                End Select
                                    
                l_xmlElementNode.setAttribute str_xmlAttribute, sOutLine
            End If
        Loop While lIndex >= lLastIndex And lCount < lNoItems
        l_xmlResponseElementNode.appendChild l_xmlElementNode
    End Select
    Set l_xmlResponseElementNode = Nothing
    Set l_xmlElementNode = Nothing
    Set l_xmlElementNode2 = Nothing
Exit Function:
StoreAsXMLErrortrap:
    
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description
    Set l_xmlResponseElementNode = Nothing
    Set l_xmlElementNode = Nothing
    Set l_xmlElementNode2 = Nothing
End Function
Private Function StartUp() As Long
    
    Dim errorCode As Long
    errorCode = QAServer_Startup(MakeCString(""), MakeCString(""))
    If (errorCode = qaerr_SRVALREADYSTARTED) Then
        errorCode = 0
    End If
        
    StartUp = errorCode
End Function
Private Function SearchQASPro(ByRef p_strKey As String, _
ByRef p_strElements As Variant, _
ByRef p_strBuffer As String, _
ByRef p_xmlDom As FreeThreadedDOMDocument40) As Boolean
Dim sSearch As String
Dim eCode As Long
Dim lResultSize As Long
Dim glBufferSize As Long
Dim strKey As String
Dim bKeySearch As Boolean
Dim bSearchLoop As Boolean
bKeySearch = False
bSearchLoop = True
On Error GoTo SearchQASProErrortrap
    If (Len(p_strKey) <> 0) Then
        sSearch = p_strKey
        bKeySearch = True
    Else
        sSearch = "sr:"
        If (Len(p_strElements(0)) <> 0) Then ' flat #
            sSearch = sSearch & "FLAT" & p_strElements(0)
        End If
        If (Len(p_strElements(1)) <> 0) Then ' house #
            If (Len(sSearch) > 3) Then
                sSearch = sSearch & ","
            End If
            sSearch = sSearch & p_strElements(1)
        End If
        If (Len(p_strElements(2)) <> 0) Then ' house name
            If (Len(sSearch) > 3) Then
                sSearch = sSearch & ","
            End If
            sSearch = sSearch & p_strElements(2)
        End If
        If (Len(p_strElements(3)) <> 0) Then ' address line 1
            If (Len(sSearch) > 3) Then
                sSearch = sSearch & ","
            End If
            sSearch = sSearch & p_strElements(3)
        End If
        If (Len(p_strElements(4)) <> 0) Then ' address line 2
            If (Len(sSearch) > 3) Then
                sSearch = sSearch & ","
            End If
            sSearch = sSearch & p_strElements(4)
        End If
        If (Len(p_strElements(5)) <> 0) Then ' address line 3
            If (Len(sSearch) > 3) Then
                sSearch = sSearch & ","
            End If
            sSearch = sSearch & p_strElements(5)
        End If
        If (Len(p_strElements(6)) <> 0) Then ' address line 4
            If (Len(sSearch) > 3) Then
                sSearch = sSearch & ","
            End If
            sSearch = sSearch & p_strElements(6)
        End If
        If (Len(p_strElements(7)) <> 0) Then ' country
            If (Len(sSearch) > 3) Then
                sSearch = sSearch & ","
            End If
            sSearch = sSearch & p_strElements(7)
        End If
        If (Len(p_strElements(8)) <> 0) Then ' post code
            If (Len(sSearch) > 3) Then
                sSearch = sSearch & ","
            End If
            sSearch = sSearch & p_strElements(8)
        End If
        ' sSearch = "sr:1,B69 2JP"
    End If
    If (Len(sSearch) = 3) Then
        Err.Raise INVALID_SEARCH_ERROR, "SearchQASPro", "No search criteria have been entered."
    End If
    While bSearchLoop
        bSearchLoop = False
        eCode = QAServer_Request(MakeCString(sSearch), "")
        If eCode < 0 Then
            Err.Raise QAS_ERROR, "QAServer_Request", "Error search QAS - Error" & Str(eCode) & ")"
        Else
            ' How much response is there ??
            lResultSize = QAServer_ResponseSize()
            ' Check we've got a big enough buffer
            If lResultSize > glBufferSize Then
                glBufferSize = lResultSize
            End If
                    
            ' Clear the buffer
            p_strBuffer = Space(glBufferSize)
                 
            ' Okay, lets grab it
            eCode = QAServer_Response(p_strBuffer, glBufferSize)
            Call QAServer_FreeResponse("")
        End If
        ' IVW - Required : If picklist contains one picklist item with an address key then go back
        ' and get the full address. There does exist a setting called AutoStepIn to do this
        ' in QAS, however it does not work in the server API
        If (OneAddressItemPickList(p_strBuffer, strKey)) Then
            sSearch = strKey
            bSearchLoop = True
        End If
    Wend
      
    SearchQASPro = True
            
Exit Function:
SearchQASProErrortrap:
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description
    SearchQASPro = False
        
End Function
' Private Function OneAddressItemPickList(ByRef p_strBuffer As String, ByRef p_strKey As String) As Boolean
'
' 1) Sometimes QAS returns a picklist of one item that contains an address key - 'ad:'
' 2) Sometimes QAS returns a picklist of two items. The first is an address, the
'    second is a step in key for the post code. The first is a 100% match, the
'    second is a 50% match. If this happens we need to go and get the full address
'    for the 100% match address key.
'
Private Function OneAddressItemPickList(ByRef p_strBuffer As String, ByRef p_strKey As String) As Boolean
Dim lIndex As Long
Dim lLastIndex As Long
Dim sResponseType As String
Dim sBlah As String
Dim lNoItems As Long
Dim lStartPipe1 As Long
Dim lStartPipe2 As Long
Dim lStartPipe3 As Long
Dim lStartPipe4 As Long
Dim strKey2 As String
Dim strPercent1 As String
Dim strPercent2 As String
Dim iNoItems As Integer
    iNoItems = 0
    OneAddressItemPickList = False
    If Len(p_strBuffer) = 0 Then
        Exit Function
    End If
    lIndex = 0
    ' locate beginning of the address key
    lStartPipe1 = InStr(1, p_strBuffer, "|ad:", vbTextCompare)
    ' locate the type of response recieved from QAS
    lLastIndex = lIndex + 1
    lIndex = InStr(lLastIndex, p_strBuffer, "|")
    sResponseType = Mid$(p_strBuffer, lLastIndex, lIndex - lLastIndex)
    If (sResponseType = "PICKLIST") Then
        ' the next two blocks of code allow us to locate the amount of list items returned
        lLastIndex = lIndex + 1
        lIndex = InStr(lLastIndex, p_strBuffer, "|")
        sBlah = Mid$(p_strBuffer, lLastIndex, lIndex - lLastIndex)
        lLastIndex = lIndex + 1
        lIndex = InStr(lLastIndex, p_strBuffer, "|")
        lNoItems = Val(Mid$(p_strBuffer, lLastIndex, lIndex - lLastIndex))
        'extract the address key from the response block.
        ' only do this if one list item of type address has been returned
        If (lNoItems = 1 And lStartPipe1 > 0) Then
            
            lStartPipe2 = InStr(lStartPipe1 + 1, p_strBuffer, "|", vbTextCompare)
            p_strKey = Mid$(p_strBuffer, lStartPipe1 + 1, lStartPipe2 - lStartPipe1 - 1)
            OneAddressItemPickList = True
        End If
        'sometimes QAS returns a picklist of the same address!
        'we do not want this. if one address is a 100% match and the other
        'address is a 50% match then assume the first one is the one we need.
                
        If (lNoItems = 2 And lStartPipe1 > 0) Then
                    
            lStartPipe2 = InStr(lStartPipe1 + 1, p_strBuffer, "|", vbTextCompare)
            p_strKey = Mid$(p_strBuffer, lStartPipe1 + 1, lStartPipe2 - lStartPipe1 - 1)
                
            ' locate beginning of the 2nd address key
            lStartPipe3 = InStr(1, p_strBuffer, "|st:", vbTextCompare)
            lStartPipe4 = InStr(lStartPipe3 + 1, p_strBuffer, "|", vbTextCompare)
            strKey2 = Mid$(p_strBuffer, lStartPipe3 + 1, lStartPipe4 - lStartPipe3 - 1)
            'INR SYS4190 Can't always guarentee that the second item is a "|st:"
            'Take BD12 8AX 19A as an example
            If lStartPipe3 > 0 Then
                ' we now have the 2 QAS address keys
                strPercent1 = Mid$(p_strBuffer, lStartPipe1 - 4, 4)
                strPercent2 = Mid$(p_strBuffer, lStartPipe3 - 3, 3)
                ' if the keys contain the same post code AND the percentage match is 100% for
                ' key 1 and 50% for key 2 then it has erronously returned a PICKLIST
                ' of the same address.
                If (strPercent1 = "100%" And strPercent2 = "50%") Then
                    OneAddressItemPickList = True
                End If
                lStartPipe2 = InStr(1, p_strKey, ":", vbTextCompare)
                lStartPipe4 = InStr(lStartPipe2 + 1, p_strKey, ":", vbTextCompare)
                strPercent1 = Mid$(p_strKey, lStartPipe2 + 1, lStartPipe4 - lStartPipe2 - 1)
                    
                lStartPipe1 = InStr(1, p_strKey, ":", vbTextCompare)
                lStartPipe3 = InStr(lStartPipe1 + 1, p_strKey, ":", vbTextCompare)
                strPercent2 = Mid$(p_strKey, lStartPipe1 + 1, lStartPipe3 - lStartPipe1 - 1)
                If (strPercent1 = strPercent2) Then
                    OneAddressItemPickList = True
                End If
            End If
                
        End If
    End If
End Function
Public Function ShutDown()
    Call QAServer_Shutdown
End Function
Function MakeCString(ByVal arg As String) As String
    MakeCString = arg & Chr$(0)
End Function
Private Function ExtractKeyElements(ByRef p_xmlDom As FreeThreadedDOMDocument40, _
    ByRef p_strKey As String) As Boolean
        
'read input XML string.
Dim l_strNodeValue As String
Dim l_xmlNode As IXMLDOMNode
Dim l_xmlAttribute As IXMLDOMAttribute
On Error GoTo ExtractKeyElements
    
    Set l_xmlNode = p_xmlDom.selectSingleNode("/REQUEST/ADDRESSKEY")
    If (l_xmlNode Is Nothing) Then
        Err.Raise INVALID_XML_ERROR, "ExtractKeyElements", "ADDRESSKEY Element missing from REQUEST XML"
    End If
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("KEY")
    If Not l_xmlAttribute Is Nothing Then
        p_strKey = l_xmlAttribute.Text
    Else
        p_strKey = ""
    End If
    Trim$ (p_strKey)
    Set l_xmlNode = Nothing
    Set l_xmlAttribute = Nothing
    ExtractKeyElements = True
Exit Function:
ExtractKeyElements:
    
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description
    Set l_xmlNode = Nothing
    Set l_xmlAttribute = Nothing
    ExtractKeyElements = False
End Function
