VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omGBNatReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------
' Class Module  omGBNatReg
' File:         omGBNatReg.cls
' Author:       INR
' Date:         13-11-2002
' Purpose:      Retrieves an address or list of possible addresses from
'               the third party GB National Register Data Files

'
' History:
'   AQR                 AUTHOR      DESCRIPTION
'   ---                 ------      -----------
'   BMIDS00960/CC021    INR     13/11/2002   Created.
'   BMIDS01029          INR     20/11/2002   Limit House Name to 40 chars. and deal
'                                               with &'s in the XML.
'   BM0306              MV      12/02/2003   Amended FindAddress()
'   BM0306              MV      24/02/2003   Amended FindAddress() and FindAddressWithKey()
'   BM0306              MV      27/02/2003   Amended FindAddress() and FindAddressWithKey()
'------------------------------------------------------------------------

' Note
' This is generated as .EXE because the API software
' is not threadsafe.

Option Explicit

Private Const INVALID_XML_ERROR = vbObjectError + 1
Private Const GB_ERROR = vbObjectError + 2
Private Const INVALID_SEARCH_ERROR = vbObjectError + 3

Public Function FindAddressWithKey(ByVal p_strXML As String) As String

Dim nStartUpErrorCode As Integer
Dim nShutdownErrorCode As Integer
Dim l_xmlDom As FreeThreadedDOMDocument
Dim strError As String
Dim strKey As String
Dim strDataBuffer As String
Dim bContinue As Boolean
Dim bPickList As Boolean
Dim nSearchReturn As Integer


On Error GoTo RequestKeyErrortrap

    bPickList = True
    
    Set l_xmlDom = New FreeThreadedDOMDocument

    l_xmlDom.async = False

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
    
        nStartUpErrorCode = StartUp(strError)
        
        'MV - 12/02/2003 - BM0306
        'Start
        If nStartUpErrorCode = 1715 Or nStartUpErrorCode = 0 Then
        
            bContinue = SearchGBAddress(strKey, "", strDataBuffer, l_xmlDom, nSearchReturn)

            If (bContinue) Then
                StoreAsXML strDataBuffer, bPickList, l_xmlDom, nSearchReturn
            End If
    
            nShutdownErrorCode = ShutDown()
        
            If Not nShutdownErrorCode = 0 Then
                Err.Raise GB_ERROR, "GBClose Error", "Unable to Shutdown GB Address - Error(" & Str(nShutdownErrorCode) & ")"
            End If
        
        ElseIf Not nStartUpErrorCode <> 0 Then
            
            Err.Raise GB_ERROR, "RequestKey", "Unable to Start up GB Address - Error(" & Str(nStartUpErrorCode) & ")" & strError
        
        End If
        'End
    
    End If

    FindAddressWithKey = l_xmlDom.xml

    Set l_xmlDom = Nothing

Exit Function
RequestKeyErrortrap:
    nShutdownErrorCode = ShutDown()

    CreateErrorXML l_xmlDom, Err.Number, Err.Source, Err.Description
    FindAddressWithKey = l_xmlDom.xml
    
    Set l_xmlDom = Nothing

End Function

Private Function CreateResponseNode(ByRef p_xmlDom As FreeThreadedDOMDocument) As Boolean

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

Dim nStartUpErrorCode As Long
Dim nShutdownErrorCode As Long
Dim l_xmlDom As FreeThreadedDOMDocument
Dim strError As String
Dim strDataBuffer As String
Dim bContinue As Boolean
Dim bPickList As Boolean
Dim strAddressElements(11) As String ' should be sized to all max search elements.
Dim nSearchReturn As Integer

bContinue = False

On Error GoTo RequestErrortrap

    bPickList = True

    Set l_xmlDom = New FreeThreadedDOMDocument

    l_xmlDom.async = False
        
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
        
        nStartUpErrorCode = StartUp(strError)
        
        'MV - 12/02/2003 - BM0306
        'Start
        If nStartUpErrorCode = 1715 Or nStartUpErrorCode = 0 Then

            bContinue = SearchGBAddress("", strAddressElements, _
                                strDataBuffer, l_xmlDom, nSearchReturn)
            
            If (bContinue) Then
                StoreAsXML strDataBuffer, bPickList, l_xmlDom, nSearchReturn
            End If
            
            nShutdownErrorCode = ShutDown()
            
            If Not nShutdownErrorCode = 0 Then
                Err.Raise GB_ERROR, "GBClose Error", "Unable to Shutdown GB Address - Error(" & Str(nShutdownErrorCode) & ")"
            End If
    
        ElseIf Not nStartUpErrorCode <> 0 Then
            
            Err.Raise GB_ERROR, "RequestKey", "Unable to Start up GB Address - Error(" & Str(nStartUpErrorCode) & ")" & strError
        
        End If
        'End
    End If

    FindAddress = l_xmlDom.xml

    Set l_xmlDom = Nothing

Exit Function

RequestErrortrap:
    nShutdownErrorCode = ShutDown()

    CreateErrorXML l_xmlDom, Err.Number, Err.Source, Err.Description
    FindAddress = l_xmlDom.xml
    
    Set l_xmlDom = Nothing

End Function

Private Function ValidateAddressElements(ByRef p_xmlDom As FreeThreadedDOMDocument, _
    ByRef p_strAddressLines As Variant) As Boolean

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
            Err.Raise GB_ERROR, "ValidateAddressElements", "Error in XML - Insuffucient Search Criteria"
        End If
    End If
    
    
    Set l_xmlNode = Nothing
    Set l_xmlAttribute = Nothing
    
    ValidateAddressElements = bOk
    
Exit Function
ValidateAddressElementsError:
    
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description
    
    Set l_xmlNode = Nothing
    Set l_xmlAttribute = Nothing

    ValidateAddressElements = False
    
End Function

Private Function ExtractAddressElements(ByRef p_xmlDom As FreeThreadedDOMDocument, _
    ByRef p_strAddressLines As Variant, _
    ByRef p_bPickList) As Boolean
            
    'read input XML string.
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
    
Exit Function
ExtractAddressElements:
    
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description
    
    Set l_xmlNode = Nothing
    Set l_xmlAttribute = Nothing

    ExtractAddressElements = False
    
End Function


Private Function CreateErrorXML(ByRef p_xmlDom As FreeThreadedDOMDocument, ByRef lNumber As Long, ByRef strSource As String, ByRef strDesc As String)
    
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
    
Exit Function

CreateErrorXMLErrortrap:
        
    Set l_xmlResponseElement = Nothing
    Set l_xmlRequestNode = Nothing
    Set l_xmlResponseNode = Nothing
    Set l_xmlErrorElement = Nothing
    Set l_xmlErrorElementNode = Nothing
    Set l_xmlWorkingNode = Nothing
    Set l_OldNode = Nothing
        
End Function

Private Function StoreAsXML(ByRef strDataBuffer As String, _
                ByVal bPickList As Boolean, _
                ByRef p_xmlDom As FreeThreadedDOMDocument, _
                ByVal nSearchReturn As Integer) As Boolean
    

On Error GoTo StoreAsXMLErrortrap

Const cstrFunctionName As String = "StoreAsXML"
Dim objGeneralAssist As New GeneralAssist

Dim xmlDoc As FreeThreadedDOMDocument
Dim vxmlRequest As IXMLDOMElement, xmlAddresses As IXMLDOMElement, xmlAddress As IXMLDOMElement
Dim xmlAddressList As IXMLDOMNodeList
Dim strPickList As String
Dim str_xmlAttribute As String
Dim l_xmlResponseElementNode As IXMLDOMElement
Dim l_xmlElementNode As IXMLDOMElement
Dim l_xmlElementNode2 As IXMLDOMElement


If Len(strDataBuffer) = 0 Then
    Exit Function
End If

'Get Addresses
Set xmlDoc = xmlLoad(strDataBuffer, cstrFunctionName)
Set xmlAddresses = xmlDoc.selectSingleNode("/GBRESPONSE")
strPickList = xmlGetMandatoryAttributeText(xmlAddresses, "PICKLIST")

Dim nPosition As Integer
Dim sKey As String
Dim nNoItems As Integer
Dim nCurrItem As Integer
Dim nPossibleNumItems As Integer
Dim sPickHeader As String
Dim strDisplayText As String

Dim strOrg As String
Dim strSubBuild As String
Dim strBuildname As String
Dim strBuildnum As String
Dim strThoro As String
Dim strDepThoro As String
Dim strDoubleDepLoc As String
Dim strDepLocality As String
Dim strCounty As String
Dim strTown As String
Dim strPCode As String
Dim strPOBox As String
'Dim bHasNumber As Boolean
'
'bHasNumber = False

    Set l_xmlResponseElementNode = p_xmlDom.selectSingleNode("/RESPONSE")


    If (StrComp(strPickList, "N") = 0) Then
'        ' Generated address from postcode - no problems

        l_xmlResponseElementNode.setAttribute "TYPE", "SUCCESS"
        Set l_xmlElementNode = p_xmlDom.createElement("ADDRESSDATA")
        l_xmlElementNode.setAttribute "QASRETURN", "ADDRESS"

        l_xmlResponseElementNode.appendChild l_xmlElementNode

        Set xmlAddress = xmlGetMandatoryNode(xmlAddresses, "/GBRESPONSE/GBResult")
            'elements available for selection
            '  DEPT,ORGN,POBX,SUBB,BNAM,BNUM,DPTH,THOR,DDLO,DPLO,TOWN,CNTY,PCOD
            strOrg = xmlGetAttributeText(xmlAddress, "ORGN")
            If (Len(strOrg) > 0) Then
                str_xmlAttribute = "ORGANISATIONNAME"
                strOrg = objGeneralAssist.ConvertToMixedCase(strOrg)
                l_xmlElementNode.setAttribute str_xmlAttribute, strOrg
            End If

            strSubBuild = xmlGetAttributeText(xmlAddress, "SUBB")
            If (Len(strSubBuild) > 0) Then
                'If subbuilding contains the word FLAT, remove it.
                nPosition = InStr(strSubBuild, "FLAT") + 4
                If nPosition > 4 Then
                    strSubBuild = Trim(Mid$(strSubBuild, nPosition, Len(strSubBuild) - nPosition + 1))
                End If
                str_xmlAttribute = "SUBBUILDINGNAME"
                l_xmlElementNode.setAttribute str_xmlAttribute, strSubBuild
            End If

            strBuildname = xmlGetAttributeText(xmlAddress, "BNAM")
            If (Len(strOrg) > 0) Or (Len(strBuildname) > 0) Then
                'if we have an organisation name, display it.
                str_xmlAttribute = "BUILDINGORHOUSENAME"
                If (Len(strBuildname) > 0) Then
                    strBuildname = strBuildname & " " & strOrg
                Else
                    strBuildname = strOrg
                End If
                'BMIDS01029
                If (Len(strBuildname) > 40) Then
                    strBuildname = Left(strBuildname, 40)
                End If

                strBuildname = objGeneralAssist.ConvertToMixedCase(strBuildname)
                l_xmlElementNode.setAttribute str_xmlAttribute, strBuildname
            End If
            
            strBuildnum = xmlGetAttributeText(xmlAddress, "BNUM")
            If (Len(strBuildnum) > 0) Then
                str_xmlAttribute = "BUILDINGORHOUSENUMBER"
                strBuildnum = objGeneralAssist.ConvertToMixedCase(strBuildnum)
                l_xmlElementNode.setAttribute str_xmlAttribute, strBuildnum
            End If
            
            strThoro = xmlGetAttributeText(xmlAddress, "THOR")
            If (Len(strThoro) > 0) Then
                str_xmlAttribute = "THOROUGHFARENAME"
                strThoro = objGeneralAssist.ConvertToMixedCase(strThoro)
                l_xmlElementNode.setAttribute str_xmlAttribute, strThoro
            End If
            'elements available for selection
            '  DEPT,ORGN,POBX,SUBB,BNAM,BNUM,DPTH,THOR,DDLO,DPLO,TOWN,CNTY,PCOD
            
            strDepThoro = xmlGetAttributeText(xmlAddress, "DPTH")
            If (Len(strThoro) > 0) Then
                str_xmlAttribute = "DEPENDENTTHOROUGHFARENAME"
                strDepThoro = objGeneralAssist.ConvertToMixedCase(strDepThoro)
                l_xmlElementNode.setAttribute str_xmlAttribute, strDepThoro
            End If
            
            strDoubleDepLoc = xmlGetAttributeText(xmlAddress, "DDLO")
            If (Len(strThoro) > 0) Then
                str_xmlAttribute = "DOUBLEDEPENDENTLOCALITY"
                strDoubleDepLoc = objGeneralAssist.ConvertToMixedCase(strDoubleDepLoc)
                l_xmlElementNode.setAttribute str_xmlAttribute, strDoubleDepLoc
    
            End If
            
            strDepLocality = xmlGetAttributeText(xmlAddress, "DPLO")
            If (Len(strThoro) > 0) Then
                str_xmlAttribute = "DEPENDENTLOCALITY"
                strDepLocality = objGeneralAssist.ConvertToMixedCase(strDepLocality)
                l_xmlElementNode.setAttribute str_xmlAttribute, strDepLocality
            End If
            
            strTown = xmlGetAttributeText(xmlAddress, "TOWN")
            If (Len(strTown) > 0) Then
                str_xmlAttribute = "POSTTOWN"
                strTown = objGeneralAssist.ConvertToMixedCase(strTown)
                l_xmlElementNode.setAttribute str_xmlAttribute, strTown
            End If
            
            strCounty = xmlGetAttributeText(xmlAddress, "CNTY")
            If (Len(strCounty) > 0) Then
                str_xmlAttribute = "COUNTY"
                strCounty = objGeneralAssist.ConvertToMixedCase(strCounty)
                l_xmlElementNode.setAttribute str_xmlAttribute, strCounty
    
            End If
            
            strPCode = xmlGetAttributeText(xmlAddress, "PCOD")
            If (Len(strPCode) > 0) Then
                str_xmlAttribute = "POSTCODE"
                l_xmlElementNode.setAttribute str_xmlAttribute, strPCode
            End If
            
            strPOBox = xmlGetAttributeText(xmlAddress, "POBX")
            If (Len(strPCode) > 0) Then
                str_xmlAttribute = "POBOXDETAILS"
                l_xmlElementNode.setAttribute str_xmlAttribute, strPOBox
    
            End If

        l_xmlResponseElementNode.appendChild l_xmlElementNode


    Else
        ' A Building list has been returned - display list and get a choice from user

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
        
            nPossibleNumItems = 0
            nNoItems = 0

            l_xmlResponseElementNode.setAttribute "TYPE", "SUCCESS"

            Set l_xmlElementNode = p_xmlDom.createElement("ADDRESSDATA")
            l_xmlElementNode.setAttribute "QASRETURN", "PICKLIST"
            l_xmlElementNode.setAttribute "LISTLINES", ""
            l_xmlElementNode.setAttribute "LISTHEADER", sPickHeader

            l_xmlResponseElementNode.appendChild l_xmlElementNode

            ' Go through the returned items, and
            ' place in the xml to be returned
            Set xmlAddressList = xmlGetMandatoryNodeList(xmlAddresses, ".//GBResult")
            nPossibleNumItems = xmlAddressList.length
            
            If nPossibleNumItems > 0 Then
               For nCurrItem = 1 To nPossibleNumItems
                    Set xmlAddress = xmlAddressList.Item(nCurrItem - 1)
    
                '  DEPT,ORGN,POBX,SUBB,BNAM,BNUM,DPTH,THOR,DDLO,DPLO,TOWN,CNTY,PCOD
                    strPCode = xmlGetAttributeText(xmlAddress, "PCOD")
                    strThoro = xmlGetAttributeText(xmlAddress, "THOR")
                    strBuildnum = xmlGetAttributeText(xmlAddress, "BNUM")
                    strBuildname = xmlGetAttributeText(xmlAddress, "BNAM")
                    strSubBuild = xmlGetAttributeText(xmlAddress, "SUBB")
                    strOrg = xmlGetAttributeText(xmlAddress, "ORGN")
                    
                    Set l_xmlElementNode2 = p_xmlDom.createElement("ITEM")
                    'what elements are available to display to the user
                    strDisplayText = strThoro
                    If (Len(strBuildnum) > 0) Then
                        strDisplayText = strDisplayText & "  " & strBuildnum
                    End If
                    If (Len(strBuildname) > 0) Then
                        strDisplayText = strDisplayText & "  " & strBuildname
                    End If
                    If (Len(strSubBuild) > 0) Then
                        strDisplayText = strDisplayText & "  " & strSubBuild
                    End If
                    If (Len(strOrg) > 0) Then
                        strDisplayText = strDisplayText & "  " & strOrg
                    End If
                    
                    l_xmlElementNode2.setAttribute "TEXT", strDisplayText
    
                    'Make up the key with all possible elements we might need to
                    'pinpoint the address
                    sKey = "ad:" & strPCode & ":" & strThoro & ":" & strBuildnum & ":" _
                            & strBuildname & ":" & strSubBuild & ":" & strOrg
                    l_xmlElementNode2.setAttribute "KEY", sKey
    
                    l_xmlElementNode.appendChild l_xmlElementNode2
                    nNoItems = nNoItems + 1
                Next
            End If
            
            l_xmlElementNode.setAttribute "LISTLINES", nNoItems

        End Select

    End If

    Set objGeneralAssist = Nothing
    Set xmlDoc = Nothing
    Set vxmlRequest = Nothing
    Set xmlAddresses = Nothing
    Set xmlAddress = Nothing
    Set xmlAddressList = Nothing
    Set l_xmlResponseElementNode = Nothing
    Set l_xmlElementNode = Nothing
    Set l_xmlElementNode2 = Nothing

Exit Function
StoreAsXMLErrortrap:
    
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description
    
    Set objGeneralAssist = Nothing
    Set xmlDoc = Nothing
    Set vxmlRequest = Nothing
    Set xmlAddresses = Nothing
    Set xmlAddress = Nothing
    Set xmlAddressList = Nothing
    Set l_xmlResponseElementNode = Nothing
    Set l_xmlElementNode = Nothing
    Set l_xmlElementNode2 = Nothing
    
End Function

Private Function StartUp(ByRef strErrMsg As String) As Integer

On Error GoTo StartupErrortrap
Dim sQuery As String * 100
Dim strNatRegisterDataPath As String
    
    Dim objGlobalParamDO As GlobalParameterDO
    Set objGlobalParamDO = New GlobalParameterDO

    'Get the GBNatRegisterDataLocation GlobalParameter Value from GlobalParameter DB Table
    strNatRegisterDataPath = objGlobalParamDO.GetCurrentParameterByType("GBNatRegisterDataLocation", "STRING")
    
    'Check for the Directory Existence
    If Dir(strNatRegisterDataPath, vbDirectory) = "" Then
        'Doesn't exist, data is in the wrong place
        Err.Raise GB_ERROR, "Startup", "Data directory specified in the GlobalParameter: GBNatRegisterDataLocation, does not exist."
    End If
    
    ' Clear contents of buffer
    szBuffr = Space(15000)
    ' Clear contents of reply
    szReply = Space(10)
    ' Clear contents of handle
    szHandl = Space(10)
    
    '  INIT('c:\NATREG')
    sQuery = "INIT('" & strNatRegisterDataPath & "')"
    
    ' Perform tools call
    GBIMAPI sQuery, szBuffr, szReply, szHandl
    
    Set objGlobalParamDO = Nothing
    
    StartUp = szReply
Exit Function

StartupErrortrap:
    Set objGlobalParamDO = Nothing

    ShutDown
    Err.Raise GB_ERROR, "Startup: ", "Startup error" & Err.Number & Err.Source & Err.Description
        
End Function

Private Function SearchGBAddress(ByRef p_strKey As String, _
ByRef p_strElements As Variant, _
ByRef p_strFieldText As String, _
ByRef p_xmlDom As FreeThreadedDOMDocument, _
ByRef nSearchReturn As Integer) As Boolean

Dim szTemp As String * 64
Dim szTemp2 As String * 6
Dim nSize, nFieldCount As Integer
Dim sAddress, sQuery As String

Dim colFieldIds As Collection
Set colFieldIds = New Collection
Dim nStartPos As Integer
Dim nCurrPos As Integer
Dim nFlatPos As Integer
Dim strFlatCompare As String
Dim nNumRecords As Integer

Dim strOrg As String
Dim strSubBuild As String
Dim strBuildname As String
Dim strBuildnum As String
Dim strThoroFare As String
Dim strPCode As String

On Error GoTo SearchGBAddressErrortrap

nNumRecords = 0
    
    If (Len(p_strKey) <> 0) Then
        '   FORMAT for p_strKey is :
        '   "ad:strPCode:strThoro:strBuildnum:strBuildname:strSubBuild:strOrg"
        '   Not all of the information will always be available
        '   Possible fields for adding to select:  ORGN,SUBB,BNAM,BNUM,THOR,PCOD
    
        'Get the Postcode
        nStartPos = InStr(p_strKey, ":") + 1
        nCurrPos = InStr(nStartPos, p_strKey, ":")
    
        If ((nStartPos = 4) And (Len(p_strKey) > 4)) Then
            strPCode = Mid$(p_strKey, nStartPos, nCurrPos - nStartPos)
        Else
            'invalid key format
            Err.Raise INVALID_SEARCH_ERROR, " SearchGBAddress ", "Using invalid key in search: " & p_strKey
        End If
        sQuery = "SELECT * FROM NR WHERE WALK = '" & strPCode & "'"
    
        'Get the thorofare
        'once we get to here we have a valid postcode and hence a valid search key, so don't error.
        nStartPos = InStr(nStartPos, p_strKey, ":") + 1
        nCurrPos = InStr(nStartPos, p_strKey, ":")
    
        If ((nStartPos > 4) And (nCurrPos > nStartPos)) Then
            strThoroFare = Mid$(p_strKey, nStartPos, nCurrPos - nStartPos)
        End If
        If (Len(strThoroFare) > 0) Then
            sQuery = sQuery & " AND THOR = '" & strThoroFare & "'"
        End If
        
        'Get the building number
        nStartPos = InStr(nStartPos, p_strKey, ":") + 1
        nCurrPos = InStr(nStartPos, p_strKey, ":")
    
        If ((nStartPos > 4) And (nCurrPos > nStartPos)) Then
            strBuildnum = Mid$(p_strKey, nStartPos, nCurrPos - nStartPos)
        End If
        If (Len(strBuildnum) > 0) Then
            sQuery = sQuery & " AND BNUM = '" & strBuildnum & "'"
        End If
        
        'Get the Building Name
        nStartPos = InStr(nStartPos, p_strKey, ":") + 1
        nCurrPos = InStr(nStartPos, p_strKey, ":")
    
        If ((nStartPos > 4) And (nCurrPos > nStartPos)) Then
            strBuildname = Mid$(p_strKey, nStartPos, nCurrPos - nStartPos)
        End If
        If (Len(strBuildname) > 0) Then
            sQuery = sQuery & " AND BNAM = '" & strBuildname & "'"
        End If
    
        'Get the SubBuilding number
        nStartPos = InStr(nStartPos, p_strKey, ":") + 1
        nCurrPos = InStr(nStartPos, p_strKey, ":")
    
        If ((nStartPos > 4) And (nCurrPos > nStartPos)) Then
            strSubBuild = Mid$(p_strKey, nStartPos, nCurrPos - nStartPos)
        End If
        If (Len(strSubBuild) > 0) Then
            sQuery = sQuery & " AND SUBB = '" & strSubBuild & "'"
        End If
        
        'Get the Organisation
        nStartPos = InStr(nCurrPos, p_strKey, ":") + 1
        nCurrPos = Len(p_strKey) + 1
        
        If (nCurrPos > nStartPos) Then
            strOrg = Mid$(p_strKey, nStartPos, nCurrPos - nStartPos)
        End If
    
        If (Len(strOrg) > 0) Then
            sQuery = sQuery & " AND ORGN = '" & strOrg & "'"
        End If
    
    Else
        'Queries are of the Format:
        'However, doesn't seem to like anything except select * from NR
        '  sQuery = "SELECT * FROM NR WHERE WALK = 'WR9 7NB'"
        '  sQuery = "SELECT DEPT,ORGN,POBX,SUBB,BNAM,BNUM,DPTH,THOR,DDLO,DPLO,TOWN,CNTY,PCOD FROM NR WHERE WALK = 'WR9 7NB'"
           sQuery = "SELECT * FROM NR WHERE "
        
        If (Len(p_strElements(8)) <> 0) Then ' post code
            sQuery = sQuery & " WALK = '" & p_strElements(8) & "'"
        End If
    
        If (Len(p_strElements(3)) <> 0) Then ' street
            sQuery = sQuery & " AND THOR = '" & p_strElements(3) & "'"
        End If
    
        'Building can consist of multiple comma seperated elements
        If (Len(p_strElements(0)) <> 0) Then ' flat #
    
            strFlatCompare = UCase(p_strElements(0))
            nFlatPos = InStr(strFlatCompare, "FLAT")
    
            If nFlatPos > 0 Then
                sQuery = sQuery & " AND SUBB = '" & p_strElements(0) & "'"
            Else
                'Subbuilding may or may not be a flat, check for both
                sQuery = sQuery & " AND (SUBB = '" & "FLAT " & p_strElements(0) & "'" & " OR SUBB = '" & p_strElements(0) & "')"
            End If
        End If
        
    
        If (Len(p_strElements(1)) <> 0) Then ' house #
            sQuery = sQuery & " AND BNUM = '" & p_strElements(1) & "'"
        End If
        
    
        If (Len(p_strElements(2)) <> 0) Then ' house name
        'House Name here could be a house name or an organisation name
            sQuery = sQuery & " AND (BNAM = '" & p_strElements(2) & "' OR ORGN = '" & p_strElements(2) & "')"
        End If
        
    End If


    ' Clear contents of buffer
    szBuffr = Space(15000)
    ' Clear contents of reply
    szReply = Space(10)

    GBIMAPI sQuery, szBuffr, szReply, szHandl
    
    szReply = Space(10)
    szTemp = Space(64)
    
    GBIMAPI "GETRETURNSIZE", szTemp, szReply, szHandl
    szTemp2 = Left$(szTemp, 5)
    nSize = szTemp2
    
    ' Clear contents of buffer
    szBuffr = Space(nSize)
    ' Clear contents of reply
    szReply = Space(10)
    sQuery = "GETALLFIELDS"
    GBIMAPI sQuery, szBuffr, szReply, szHandl
  
    'What are the available fields returned
    'No Guarentee between diff postcode calls or software updates as to the format
    'of the fields returned, so we have to look and align the data returned for
    'every time we have a new PCode.
    Set colFieldIds = GetFieldIdents(szBuffr, nFieldCount)
    
    ' Clear contents of buffer
    szBuffr = Space(nSize)
    ' Clear contents of reply
    szReply = Space(10)
    sQuery = "GETREC(1)"
    GBIMAPI sQuery, szBuffr, szReply, szHandl
    
    If szReply = 0 Then
        p_strFieldText = p_strFieldText & GetFieldTextAsStrXML(szBuffr, nFieldCount, colFieldIds)
    Else
        Err.Raise GB_ERROR, ": Address not Found", ", GB Error Number: " & szReply
    End If

    While szReply = 0
        nNumRecords = nNumRecords + 1
        ' Clear contents of buffer
        szBuffr = Space(nSize)
        ' Clear contents of reply
        szReply = Space(10)
        sQuery = "GETNEXT"
        GBIMAPI sQuery, szBuffr, szReply, szHandl
        p_strFieldText = p_strFieldText & GetFieldTextAsStrXML(szBuffr, nFieldCount, colFieldIds)
    Wend

    If nNumRecords = 1 Then
        'We have a PickList
        p_strFieldText = "<GBRESPONSE PICKLIST=""N"" >" & p_strFieldText & "</GBRESPONSE>"
    Else
        'We have an address
        p_strFieldText = "<GBRESPONSE PICKLIST=""Y"" >" & p_strFieldText & "</GBRESPONSE>"
    End If

    Set colFieldIds = Nothing

SearchGBAddress = True
            
Exit Function

SearchGBAddressErrortrap:
    ShutDown
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description
    SearchGBAddress = False
        
End Function

Public Function ShutDown() As Integer

    Dim sQuery As String * 100
    Dim nReply As Integer
    
    ' Clear contents of buffer
    szBuffr = Space(15000)
    ' Clear contents of reply
    szReply = Space(10)
    
    'copy in the IMQL string for the command
    sQuery = "TERM"
    ' use the handle we stored from our INIT call
    'which is stored in the pszHandle buffer
    ' make the Information Management Toolkit call
    nReply = GBIMAPI(sQuery, szBuffr, szReply, szHandl)
  
    ' Clear contents of handle
    szHandl = Space(10)
    ' Clear contents of buffer
    szBuffr = Space(15000)
    ' Clear contents of reply
    szReply = Space(10)

ShutDown = nReply

End Function

Private Function ExtractKeyElements(ByRef p_xmlDom As FreeThreadedDOMDocument, _
    ByRef p_strKey As String) As Boolean
        
'read input XML string.
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
    
Exit Function
ExtractKeyElements:
    
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description
    
    Set l_xmlNode = Nothing
    Set l_xmlAttribute = Nothing

    ExtractKeyElements = False
    
End Function


Public Function GetFieldIdents(ByVal strBuffer As String, _
                                    ByRef nNumFields As Integer) As Collection
Dim colFieldIds As Collection
Set colFieldIds = New Collection
Dim nStartPos, npos As Integer
Dim strField As String

On Error GoTo GetFieldIdentsErrortrap
  
nStartPos = 1
nNumFields = 0

'Buffer should be in the format:
'DEPT|ORGN|POBX|SUBB|BNAM|BNUM|DPTH|THOR|DDLO|DPLO|TOWN|CNTY|PCOD
npos = InStr(nStartPos, szBuffr, "|")
    
While npos > 0
    nNumFields = nNumFields + 1
    ' copy from the start position
    strField = Mid$(szBuffr, nStartPos, npos - nStartPos)
    colFieldIds.Add strField, CStr(nNumFields)
    nStartPos = npos + 1
    npos = InStr(nStartPos, szBuffr, "|")

Wend

'Check there are no further field names after the last seperator
npos = InStr(nStartPos, szBuffr, Chr(0))
If npos > 0 Then
    If npos > nStartPos Then
        'we have a further field to get
        nNumFields = nNumFields + 1
        ' copy from the start position
        strField = Mid$(szBuffr, nStartPos, npos - nStartPos)
        colFieldIds.Add strField, CStr(nNumFields)
    End If
End If

Set GetFieldIdents = colFieldIds
Set colFieldIds = Nothing

Exit Function
    
GetFieldIdentsErrortrap:

    
End Function

Public Function GetFieldTextAsStrXML(ByVal strBuffer As String, _
                                    ByRef nNumFields As Integer, _
                                    ByVal colFieldText As Collection) As String
                                
Dim nDataLen, nCurPos, nIdentCount As Integer
Dim strXML, strId, strText, strXmlOut As String
Dim bHadResult As Boolean

On Error GoTo GetFieldTextErrortrap

bHadResult = True
nDataLen = 64
nCurPos = 1

strXmlOut = "<GBResult"

For nIdentCount = 1 To nNumFields
    
    strText = Trim$(Mid$(strBuffer, nCurPos, nDataLen))

    
    'Any text returned
    If (Len(strText) > 0) Then
        'BMIDS01029 &'s and xml don't mix
        Dim npos As Integer
        Dim strTemp As String
        npos = 0
        npos = InStr(strText, "&")
        If (npos > 0) Then
            strTemp = Replace(strText, "&", "&amp;")
            strText = strTemp
        End If
        'set up our attributes
        strId = colFieldText.Item(nIdentCount)
        strXML = strXML & " " & strId & " = "
        strXML = strXML & """" & strText & """"
        strXmlOut = strXmlOut & strXML
        bHadResult = True
    End If
    strXML = ""
    nCurPos = nCurPos + nDataLen

Next

If (bHadResult = True) Then
    strXmlOut = strXmlOut & ">" & "</GBResult>"
Else
    strXmlOut = ""
End If

GetFieldTextAsStrXML = strXmlOut

Exit Function
    
GetFieldTextErrortrap:
    
End Function
