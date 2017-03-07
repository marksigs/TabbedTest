VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omGBAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------
' Class Module  omGBAddress
' File:         omGBAddress.cls
' Author:       INR
' Date:         30-10-2002
' Purpose:      Retrieves an address or list of possible addresses from
'               the third party GB ADDRESS MANAGER FOR WINDOWS

'
' History:
'   AQR             AUTHOR      DESCRIPTION
'   ---             ------      -----------
'   CC021            INR        30/10/2002  Created. NOTE: Within the debug environment, the gbclose()
'                               followed by gbopen()which occurs for picklist selection on a (flat &
'                               housename) or (house number & housename) combination seems to corrupt
'                               the data returned. Loss of (flat & housename) or (house number & housename
'                               Whereas, using the same string returned from the picklist
'                               in a direct call (i.e. no picklist is generated) works OK.
'                               This problem seems to go away when running release exe's/dll's.
'                               Release I was running against where I had this problem was Y2M08.
'
' BMIDS00881        INR         08/11/02 Check to see if flat has already been entered.
' BMIDS00881        INR         08/11/02 Set up GBGetNext buffer properly.
'------------------------------------------------------------------------

' Note
' This is generated as .EXE because the API software
' is not threadsafe.

Option Explicit
Private Const MAXADDRESSSIZE = 1000
Private Const MAXDATABUFF = 500
Private Const Error_buf_size = 1000

Private Const INVALID_XML_ERROR = vbObjectError + 1
Private Const GB_ERROR = vbObjectError + 2
Private Const INVALID_SEARCH_ERROR = vbObjectError + 3

Public Function FindAddressWithKey(ByVal p_strXML As String) As String

Dim nStartUpErrorCode As Integer
Dim nShutdownErrorCode As Integer
Dim l_xmlDom As FreeThreadedDOMDocument
Dim strError As String
Dim strKey As String
Dim strBuffer As String
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
        
        If (nStartUpErrorCode < 0) Then
            Err.Raise GB_ERROR, "RequestKey", "Unable to Start up GB Address - Error(" & Str(nStartUpErrorCode) & ")" & strError
        End If

        bContinue = SearchGBAddress(strKey, "", strBuffer, l_xmlDom, nSearchReturn)

        If (bContinue) Then
            StoreAsXML strBuffer, bPickList, l_xmlDom, nSearchReturn
        End If
    
        nShutdownErrorCode = ShutDown()
        
        If Not nShutdownErrorCode = 0 Then
            ' IMPORTANT - we have closed the toolkit, so don't try and call
            ' GBGetError after GBClose - instead just check the reply code
            Err.Raise GB_ERROR, "GBClose Error", "Unable to Shutdown GB Address - Error(" & Str(nShutdownErrorCode) & ")"
        End If
        
    
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
Dim strBuffer As String
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
    
        If Not nStartUpErrorCode = 0 Then
            Err.Raise GB_ERROR, "RequestKey", "Unable to Start up GB Address - Error(" & Str(nStartUpErrorCode) & ")" & strError
        End If

        bContinue = SearchGBAddress("", strAddressElements, strBuffer, l_xmlDom, nSearchReturn)
        
        If (bContinue) Then
            StoreAsXML strBuffer, bPickList, l_xmlDom, nSearchReturn
        End If
        
        nShutdownErrorCode = ShutDown()
        
        If Not nShutdownErrorCode = 0 Then
            ' IMPORTANT - we have closed the toolkit, so don't try and call
            ' GBGetError after GBClose - instead just check the reply code
            Err.Raise GB_ERROR, "GBClose Error", "Unable to Shutdown GB Address - Error(" & Str(nShutdownErrorCode) & ")"
        End If
        
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

Private Function StoreAsXML(ByRef strBuffer As String, _
                ByVal bPickList As Boolean, _
                ByRef p_xmlDom As FreeThreadedDOMDocument, _
                ByVal nSearchReturn As Integer) As Boolean
    

On Error GoTo StoreAsXMLErrortrap

Dim objGeneralAssist As New GeneralAssist
Dim szError_buf As String
Dim strData_buf As String
Dim str_xmlAttribute As String

Dim l_xmlResponseElementNode As IXMLDOMElement
Dim l_xmlElementNode As IXMLDOMElement
Dim l_xmlElementNode2 As IXMLDOMElement

Dim nError_no As Integer
Dim nReply As Integer
Dim nStart As Integer
Dim nPosition As Integer
Dim strError_msg As String
Dim sBuilding As String
Dim sKey As String
Dim lNoItems As Long
Dim sPickHeader As String
Dim strThoro_buf As String
Dim strPCode_buf As String
Dim strData As String

Dim strOrg As String
Dim strSubBuild As String
Dim strBuildname As String
Dim strBuildnum As String
Dim strThoro As String
Dim strDepThoro As String
Dim strDoubleDepLoc As String
Dim strDepLocality As String
Dim strLocality As String
Dim strCounty As String
Dim strTown As String
Dim strPCode As String
Dim strPOBox As String
Dim bHasNumber As Boolean

bHasNumber = False

' define the new line char as carriage return (ASCII 10) and line feed (ASCII 13)
NL = Chr(13) + Chr(10)
    
    If Len(strBuffer) = 0 Then
        Exit Function
    End If
    
    Set l_xmlResponseElementNode = p_xmlDom.selectSingleNode("/RESPONSE")

    
    If nSearchReturn >= 0 And nSearchReturn < 100 Then
        ' Generated address from postcode - no problems
        ' NOTE GB Mail mixed case functionality only works for the GBGetAddress call,
        ' despite the manual giving the impression that it should work. This is not
        ' a bug. It just doesn't have the functionality.

        l_xmlResponseElementNode.setAttribute "TYPE", "SUCCESS"
        Set l_xmlElementNode = p_xmlDom.createElement("ADDRESSDATA")
        l_xmlElementNode.setAttribute "QASRETURN", "ADDRESS"

        l_xmlResponseElementNode.appendChild l_xmlElementNode
        
        strData_buf = String(MAXDATABUFF, 0)
        nReply = GBGetOrganisation(strData_buf, MAXDATABUFF)
        'If ORGANISATIONNAME is returned, convert it to mixed case.
        nPosition = InStr(strData_buf, Chr(0))
        If nPosition > 0 Then
            ' copy from the start position
            strOrg = Left$(strData_buf, nPosition - 1)
            str_xmlAttribute = "ORGANISATIONNAME"
            strOrg = objGeneralAssist.ConvertToMixedCase(strOrg)
            l_xmlElementNode.setAttribute str_xmlAttribute, strOrg
        
        End If

        strData_buf = String(MAXDATABUFF, 0)
        nReply = GBGetBuilding(strData_buf, MAXDATABUFF)
        Dim strLeft As String
        strLeft = Left(strData_buf, 1)
        If (InStr(1, "0123456789", strLeft, vbTextCompare) > 0) Then
            str_xmlAttribute = "BUILDINGORHOUSENUMBER"
            bHasNumber = True

            nPosition = InStr(strData_buf, Chr(0))
            If nPosition > 0 Then
                ' copy from the start position
                strBuildnum = Left$(strData_buf, nPosition - 1)
                l_xmlElementNode.setAttribute str_xmlAttribute, strBuildnum
            
            End If
            If (Len(strOrg) > 0) Then
                str_xmlAttribute = "BUILDINGORHOUSENAME"
                l_xmlElementNode.setAttribute str_xmlAttribute, strOrg
            End If
            

        Else
            str_xmlAttribute = "BUILDINGORHOUSENAME"
                        
            'If BUILDINGNAME is returned, convert it to mixed case.
            nPosition = InStr(strData_buf, Chr(0))
            If nPosition > 0 Then
                ' copy from the start position
                strBuildname = Left$(strData_buf, nPosition - 1)
                'convert to mixedcase
                strBuildname = objGeneralAssist.ConvertToMixedCase(strBuildname)
                l_xmlElementNode.setAttribute str_xmlAttribute, strBuildname
            
            End If

        End If
        
        strData_buf = String(MAXDATABUFF, 0)
        nReply = GBGetSubBuilding(strData_buf, MAXDATABUFF)

        'If SUBBuilding is returned prefixed with "FLAT", remove it.
        nPosition = InStr(strData_buf, Chr(0))
        If nPosition <> 0 Then
            ' copy from the start position
            strSubBuild = Left$(strData_buf, nPosition - 1)
            nPosition = InStr(strSubBuild, "FLAT") + 4
            If nPosition > 4 Then
                strSubBuild = Trim(Mid$(strSubBuild, nPosition, Len(strSubBuild) - nPosition + 1))
                str_xmlAttribute = "SUBBUILDINGNAME"
                l_xmlElementNode.setAttribute str_xmlAttribute, strSubBuild
            Else
                'subbuilding is used for housenumber where we have a housename and housenumber.
                'i.e. NOT a flat.
                'FLAT is actually returned in the description from GBMail where it means FLAT.
                'So if FLAT is not found, we don't already have a house number AND
                ' there is something in subBuilding, stick this into BUILDINGORHOUSENUMBER
                If ((Not bHasNumber) And (Len(strSubBuild) > 0)) Then
                    str_xmlAttribute = "BUILDINGORHOUSENUMBER"
                    l_xmlElementNode.setAttribute str_xmlAttribute, strSubBuild
                End If
            End If
        End If
        
        strData_buf = String(MAXDATABUFF, 0)
        nReply = GBGetThorofare(strData_buf, MAXDATABUFF)
        'If THOROUGHFARENAME is returned, convert it to mixed case.
        nPosition = InStr(strData_buf, Chr(0))
        If nPosition > 0 Then
            ' copy from the start position
            strThoro = Left$(strData_buf, nPosition - 1)
            str_xmlAttribute = "THOROUGHFARENAME"
            strThoro = objGeneralAssist.ConvertToMixedCase(strThoro)
            l_xmlElementNode.setAttribute str_xmlAttribute, strThoro
        
        End If
        
        strData_buf = String(MAXDATABUFF, 0)
        nReply = GBGetDepThorofare(strData_buf, MAXDATABUFF)
        'If DEPENDENTTHOROUGHFARENAME is returned, convert it to mixed case.
        nPosition = InStr(strData_buf, Chr(0))
        If nPosition > 0 Then
            ' copy from the start position
            strDepThoro = Left$(strData_buf, nPosition - 1)
            str_xmlAttribute = "DEPENDENTTHOROUGHFARENAME"
            strDepThoro = objGeneralAssist.ConvertToMixedCase(strDepThoro)
            l_xmlElementNode.setAttribute str_xmlAttribute, strDepThoro
        
        End If
        
        strData_buf = String(MAXDATABUFF, 0)
        nReply = GBGetDoubleLocality(strData_buf, MAXDATABUFF)
        'If DOUBLEDEPENDENTLOCALITY is returned, convert it to mixed case.
        nPosition = InStr(strData_buf, Chr(0))
        If nPosition > 0 Then
            ' copy from the start position
            strDoubleDepLoc = Left$(strData_buf, nPosition - 1)
            str_xmlAttribute = "DOUBLEDEPENDENTLOCALITY"
            strDoubleDepLoc = objGeneralAssist.ConvertToMixedCase(strDoubleDepLoc)
            l_xmlElementNode.setAttribute str_xmlAttribute, strDoubleDepLoc
        
        End If
        
        strData_buf = String(MAXDATABUFF, 0)
        nReply = GBGetDepLocality(strData_buf, MAXDATABUFF)
        'If DEPENDENTLOCALITY is returned, convert it to mixed case.
        nPosition = InStr(strData_buf, Chr(0))
        If nPosition > 0 Then
            ' copy from the start position
            strDepLocality = Left$(strData_buf, nPosition - 1)
            str_xmlAttribute = "DEPENDENTLOCALITY"
            strDepLocality = objGeneralAssist.ConvertToMixedCase(strDepLocality)
            l_xmlElementNode.setAttribute str_xmlAttribute, strDepLocality
        
        End If
        
        strData_buf = String(MAXDATABUFF, 0)
        nReply = GBGetLocality(strData_buf, MAXDATABUFF)
        str_xmlAttribute = "POSTTOWN"
        'If Locality is returned, convert it to mixed case.
        nPosition = InStr(strData_buf, Chr(0))
        If nPosition > 0 Then
            ' copy from the start position
            strTown = Left$(strData_buf, nPosition - 1)
            str_xmlAttribute = "POSTTOWN"
            strTown = objGeneralAssist.ConvertToMixedCase(strTown)
            l_xmlElementNode.setAttribute str_xmlAttribute, strTown
        
        End If
       
        strData_buf = String(MAXDATABUFF, 0)
        nReply = GBGetCounty(strData_buf, MAXDATABUFF)
        'If COUNTY is returned, convert it to mixed case.
        nPosition = InStr(strData_buf, Chr(0))
        If nPosition > 0 Then
            ' copy from the start position
            strCounty = Left$(strData_buf, nPosition - 1)
            str_xmlAttribute = "COUNTY"
            strCounty = objGeneralAssist.ConvertToMixedCase(strCounty)
            l_xmlElementNode.setAttribute str_xmlAttribute, strCounty
        
        End If
        
        strData_buf = String(MAXDATABUFF, 0)
        nReply = GBGetPostcode(strData_buf, MAXDATABUFF)
        'If POSTCODE is returned
        nPosition = InStr(strData_buf, Chr(0))
        If nPosition > 0 Then
            ' copy from the start position
            strPCode = Left$(strData_buf, nPosition - 1)
            str_xmlAttribute = "POSTCODE"
            l_xmlElementNode.setAttribute str_xmlAttribute, strPCode
        
        End If
        
        strData_buf = String(MAXDATABUFF, 0)
        nReply = GBGetPOBox(strData_buf, MAXDATABUFF)
        'If POBOXDETAILS is returned,
        nPosition = InStr(strData_buf, Chr(0))
        If nPosition > 0 Then
            ' copy from the start position
            strPOBox = Left$(strData_buf, nPosition - 1)
            str_xmlAttribute = "POBOXDETAILS"
            l_xmlElementNode.setAttribute str_xmlAttribute, strPOBox
        
        End If
        
        l_xmlResponseElementNode.appendChild l_xmlElementNode
        
        
    ElseIf nSearchReturn >= 100 And nSearchReturn < 200 Then
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

            lNoItems = 0
            
            strThoro_buf = String(1024, 0)
            ' get the thorofare name
            nReply = GBGetThorofare(strThoro_buf, User_buf_size)
            
            ' check for and display errors in the main form
            If nReply <> 0 Then
                Do
                    szError_buf = String(400, 0)
                    nError_no = GBGetError(szError_buf, Error_buf_size)
                    strError_msg = strError_msg & Str(nError_no) & ": " & szError_buf
                Loop Until nError_no = 0
                Err.Raise GB_ERROR, "StoreAsXML", "Failed to get Thorofare" & strError_msg
            Else
                'just want the info
                nPosition = InStr(strThoro_buf, Chr(0))
                If nPosition <> 0 Then
                    ' copy from the start position
                    strThoro = Left$(strThoro_buf, nPosition - 1)
                End If
            End If
            
            strPCode_buf = String(1024, 0)
            ' get the Postcode
            nReply = GBGetPostcode(strPCode_buf, User_buf_size)
            
            ' check for and display errors in the main form
            If nReply <> 0 Then
                Do
                    strError_msg = strError_msg & NL & "GBGetPostCode: " & Str(nError_no) & " " & szError_buf
                    szError_buf = String(400, 0)
                    nError_no = GBGetError(szError_buf, Error_buf_size)
                    strError_msg = strError_msg & Str(nError_no) & ": " & szError_buf
                Loop Until nError_no = 0
                Err.Raise GB_ERROR, "StoreAsXML", "Failed to get PostCode" & strError_msg
            Else
                'just want the info
                nPosition = InStr(strPCode_buf, Chr(0))
                If nPosition <> 0 Then
                    ' copy from the start position
                    strPCode = Left$(strPCode_buf, nPosition - 1)
                End If

            End If
            
            
            l_xmlResponseElementNode.setAttribute "TYPE", "SUCCESS"

            Set l_xmlElementNode = p_xmlDom.createElement("ADDRESSDATA")
            l_xmlElementNode.setAttribute "QASRETURN", "PICKLIST"
            l_xmlElementNode.setAttribute "LISTLINES", ""
            l_xmlElementNode.setAttribute "LISTHEADER", sPickHeader

            l_xmlResponseElementNode.appendChild l_xmlElementNode
            
            ' cut up the returned string into individual items, and
            ' place in the xml to be returned
            
            nStart = 1
            nPosition = 1
            While Not nPosition = 0 ' Instr will return zero when it reaches the end
                ' find NL in the string
                nPosition = InStr(nStart, strBuffer, NL)
                If nPosition <> 0 Then
                    ' this is not the last item
                    ' copy from the start position to the position of the NL
                    sBuilding = Mid$(strBuffer, nStart, nPosition - nStart)

                    Set l_xmlElementNode2 = p_xmlDom.createElement("ITEM")
                    l_xmlElementNode2.setAttribute "TEXT", strThoro & "  " & sBuilding
    
                    sKey = "ad:" & strPCode & ":" & strThoro & ":" & sBuilding
                    l_xmlElementNode2.setAttribute "KEY", sKey
    
                    l_xmlElementNode.appendChild l_xmlElementNode2

                    ' move the start position to past the NL
                    nStart = nPosition + 2
                    lNoItems = lNoItems + 1
                Else
                    ' this is the last item
                    ' copy from the start position to the end of the string
                    
                    sBuilding = Mid$(strBuffer, nStart, Len(strBuffer) - nStart)
                    
                    Set l_xmlElementNode2 = p_xmlDom.createElement("ITEM")
                    l_xmlElementNode2.setAttribute "TEXT", strThoro & "  " & sBuilding
    
                    sKey = "ad:" & strPCode & ":" & strThoro & ":" & sBuilding
                    l_xmlElementNode2.setAttribute "KEY", sKey
    
                    l_xmlElementNode.appendChild l_xmlElementNode2
                    lNoItems = lNoItems + 1
                    
                End If
                
            Wend
            l_xmlElementNode.setAttribute "LISTLINES", lNoItems

            
        End Select
        
    ElseIf nSearchReturn >= 200 And nSearchReturn < 300 Then
        ' A thorofare list has been returned - display list and get a choice from user

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


            lNoItems = 0
            
            strPCode_buf = String(1024, 0)
            ' get the Postcode
            nReply = GBGetPostcode(strPCode_buf, User_buf_size)
            
            ' check for and display errors in the main form
            If nReply <> 0 Then
                Do
                    strError_msg = strError_msg & NL & "GBGetPostCode: " & Str(nError_no) & " " & szError_buf
                    szError_buf = String(400, 0)
                    nError_no = GBGetError(szError_buf, Error_buf_size)
                    strError_msg = strError_msg & Str(nError_no) & ": " & szError_buf
                Loop Until nError_no = 0
                Err.Raise GB_ERROR, "StoreAsXML", "Failed to get PostCode" & strError_msg
            Else
                'just want the info
                nPosition = InStr(strPCode_buf, Chr(0))
                If nPosition <> 0 Then
                    ' copy from the start position
                    strPCode = Left$(strPCode_buf, nPosition - 1)
                End If

            End If
            
            
            l_xmlResponseElementNode.setAttribute "TYPE", "SUCCESS"

            Set l_xmlElementNode = p_xmlDom.createElement("ADDRESSDATA")
            l_xmlElementNode.setAttribute "QASRETURN", "PICKLIST"
            l_xmlElementNode.setAttribute "LISTLINES", ""
            l_xmlElementNode.setAttribute "LISTHEADER", sPickHeader

            l_xmlResponseElementNode.appendChild l_xmlElementNode
            
            ' cut up the returned string into individual items, and
            ' place in the xml to be returned
            
            nStart = 1
            nPosition = 1
            While Not nPosition = 0 ' Instr will return zero when it reaches the end
                ' find NL in the string
                nPosition = InStr(nStart, strBuffer, NL)
                If nPosition <> 0 Then
                    ' this is not the last item
                    ' copy from the start position to the position of the NL
                    strThoro = Mid$(strBuffer, nStart, nPosition - nStart)

                    Set l_xmlElementNode2 = p_xmlDom.createElement("ITEM")
                    l_xmlElementNode2.setAttribute "TEXT", strThoro
    
                    sKey = "st:" & strPCode & ":" & strThoro
                    l_xmlElementNode2.setAttribute "KEY", sKey
    
                    l_xmlElementNode.appendChild l_xmlElementNode2

                    ' move the start position to past the NL
                    nStart = nPosition + 2
                    lNoItems = lNoItems + 1
                Else
                    ' this is the last item
                    ' copy from the start position to the end of the string
                    
                    sBuilding = Mid$(strBuffer, nStart, Len(strBuffer) - nStart)
                    
                    Set l_xmlElementNode2 = p_xmlDom.createElement("ITEM")
                    l_xmlElementNode2.setAttribute "TEXT", strThoro
    
                    sKey = "st:" & strPCode & ":" & strThoro
                    l_xmlElementNode2.setAttribute "KEY", sKey
    
                    l_xmlElementNode.appendChild l_xmlElementNode2
                    lNoItems = lNoItems + 1
                    
                End If
                
            Wend
            l_xmlElementNode.setAttribute "LISTLINES", lNoItems
        End Select

    End If
    
    Set objGeneralAssist = Nothing
    Set l_xmlResponseElementNode = Nothing
    Set l_xmlElementNode = Nothing
    Set l_xmlElementNode2 = Nothing

Exit Function
StoreAsXMLErrortrap:
    
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description
    
    Set objGeneralAssist = Nothing
    Set l_xmlResponseElementNode = Nothing
    Set l_xmlElementNode = Nothing
    Set l_xmlElementNode2 = Nothing
    
End Function

Private Function StartUp(ByRef strErrMsg As String) As Integer
    
Dim nReply As Integer
Dim nError_no As Integer
Dim strError_msg As String
Dim szError_buf As String

On Error GoTo StartupErrortrap

' define the new line char as carriage return (ASCII 10) and line feed (ASCII 13)
NL = Chr(13) + Chr(10)

nReply = GBOpen()

' If the error code is not zero, and the error is not fatal then pop errors off the
' stack until it is empty
If nReply <> 0 And nReply > -700 Then
    Do
        ' set up the error buffer as a string of nulls
        szError_buf = String(Error_buf_size, 0)
        nError_no = GBGetError(szError_buf, Error_buf_size)

        ' push each error messsage into a buffer
        strError_msg = strError_msg & NL & "GBOpen: " & Str(nError_no) & " " & szError_buf

    Loop Until nError_no = 0

    ' return the contents of the error buffer
    strErrMsg = strError_msg
End If

' If there is a fatal error, we cannot call the error handler - show GBOpens error message,
' and exit

If nReply < -700 Then
    strError_msg = " A fatal error has occurred in GBOpen, and the Error handler could not be initialised. The error number is " & Str(nReply)
    strErrMsg = strError_msg
    nReply = GBClose()
    End
End If

' Only those configuration options that differ from the default config are set here.
' force delimiters to correct values so any config file changes won't affect our call.
nReply = GBSetInDelimiter("<CR><LF>")
nReply = GBSetOutDelimiter("<CR><LF>")
nReply = GBSetOutTerminator("<NULL>")
        
strErrMsg = strError_msg
StartUp = nReply
Exit Function

StartupErrortrap:
    ShutDown
    Err.Raise GB_ERROR, "Startup", "Startup error" & Err.Number & Err.Source & Err.Description
        
End Function

Private Function SearchGBAddress(ByRef p_strKey As String, _
ByRef p_strElements As Variant, _
ByRef p_strBuffer As String, _
ByRef p_xmlDom As FreeThreadedDOMDocument, _
ByRef nSearchReturn As Integer) As Boolean

Dim sSearch As String
Dim sMultiData As String
Dim bFlat As Boolean
Dim bHouseName As Boolean
Dim szError_buf As String
Dim strError_msg As String
Dim nError_no As Integer

Dim nStartPos As Integer
Dim nCurrPos As Integer
Dim nFlatPos As Integer
Dim strFlatCompare As String
Dim strPCode As String
Dim strThoroFare As String
Dim strBuilding As String

On Error GoTo SearchGBAddressErrortrap

' define the new line char as carriage return (ASCII 10) and line feed (ASCII 13)
NL = Chr(13) + Chr(10)

bFlat = False
bHouseName = False

If (Len(p_strKey) <> 0) Then
'   FORMAT for p_strKey is "ad:postcode:thorofare:building for a building list"
'                           OR "st:postcode:thorofare"
'   FORMAT for sSearch is "postcodeNLthorofareNLbuilding"

    nStartPos = InStr(p_strKey, ":") + 1
    nCurrPos = InStr(nStartPos, p_strKey, ":")

    If ((nStartPos = 4) And (Len(p_strKey) > 4)) Then
        strPCode = Mid$(p_strKey, nStartPos, nCurrPos - nStartPos)
    Else
        'invalid key format
        Err.Raise INVALID_SEARCH_ERROR, " SearchGBAddress ", "Using invalid key in search: " & p_strKey
    End If
    sSearch = strPCode & NL

    nStartPos = InStr(nCurrPos, p_strKey, ":") + 1
    nCurrPos = InStr(nStartPos, p_strKey, ":")
    If nCurrPos = 0 Then
        'Could be that we only have a postcode & thorofare
        nCurrPos = Len(p_strKey) + 1
    End If
    'once we get to here we have a valid postcode and hence a valid search key, so don't error.
    If ((nStartPos > 4) And (nCurrPos > nStartPos)) Then
        strThoroFare = Mid$(p_strKey, nStartPos, nCurrPos - nStartPos)
        If (StrComp(strThoroFare, "<NOSTREET>", vbTextCompare) = 0) Then
            'No Thorfare available for searching
            strThoroFare = ""
        End If
    End If
    
    sSearch = sSearch & strThoroFare & NL
   
    'once we get to here we have at least a valid postcode and hence a valid search key, so don't error.
    If (nCurrPos < Len(p_strKey)) Then
        strBuilding = Mid$(p_strKey, nCurrPos + 1, Len(p_strKey) - nCurrPos)
    End If
    sSearch = sSearch & strBuilding

Else
'        FORMAT for sSearch is "postcodeNLthorofareNLbuilding"
    sSearch = ""

    If (Len(p_strElements(8)) <> 0) Then ' post code
        sSearch = sSearch & p_strElements(8)
    End If
    sSearch = sSearch & NL
    
    If (Len(p_strElements(3)) <> 0) Then ' street
        sSearch = sSearch & p_strElements(3)
    End If
    sSearch = sSearch & NL

    'Building can consist of multiple comma seperated elements
    If (Len(p_strElements(0)) <> 0) Then ' flat #
    
        strFlatCompare = UCase(p_strElements(0))
        nFlatPos = InStr(strFlatCompare, "FLAT")
        'BMIDS00881
        If nFlatPos > 0 Then
            sSearch = sSearch & p_strElements(0)
        Else
            sSearch = sSearch & "FLAT " & p_strElements(0)
        End If
        bFlat = True
    End If

    If (Len(p_strElements(1)) <> 0) Then ' house #
        If (bFlat) Then
            sSearch = sSearch & ","
        End If
        sSearch = sSearch & p_strElements(1)
        bHouseName = True
    End If
    
    If (Len(p_strElements(2)) <> 0) Then ' house name
        If (bHouseName Or bFlat) Then
            sSearch = sSearch & ","
        End If
        sSearch = sSearch & p_strElements(2)
    End If

End If

If (Len(sSearch) = 2) Then
    Err.Raise INVALID_SEARCH_ERROR, "SearchGBAddress", "No search criteria have been entered."
End If

' This next line is required as Visual Basic always
' converts fixed length strings to variant length strings
' before passing them to a DLL. We therefore ensure the
' string is the correct length by stuffing it with nulls
p_strBuffer = sSearch & String(MAXADDRESSSIZE - Len(sSearch), 0)

nSearchReturn = GBGetAddress(p_strBuffer, MAXADDRESSSIZE)

If nSearchReturn < 0 Then
    ' There is an error - further calls to GBGetAddress will not
    ' return valid information
    ' Pop all the errors off the error stack
    While nSearchReturn <> 0
        szError_buf = String(Error_buf_size, 0)
        nSearchReturn = GBGetError(szError_buf, Error_buf_size)

        If nSearchReturn <> 0 Then
            Dim nPos As Integer
            nPos = InStr(szError_buf, Chr(0))
            If nPos <> 0 Then
                ' this is not the last item
                ' copy from the start position
                szError_buf = Left$(szError_buf, nPos)
            End If

            ' concat all errors into a buffer
            strError_msg = strError_msg & NL & "GBGetAddress: " & Str(nError_no) & " " & szError_buf
        End If
    Wend
    
    Err.Raise GB_ERROR, " SearchGBAddress: ", strError_msg
    
End If
nError_no = nSearchReturn
sMultiData = p_strBuffer
While (nError_no = 501)

    szError_buf = String(Error_buf_size, 0)

    'The reply of 501 will only ever be returned in conjunction with
    'another reply code. Pop the more response off the error stack
    nError_no = GBGetError(szError_buf, Error_buf_size)

    ' set nSearchReturn to next error on stack, which will be
    ' either building or thorofare list
    nSearchReturn = GBGetError(szError_buf, Error_buf_size)
    
    '502 indicates that the call has returned the last part of the generated data
    'BMIDS00881
    p_strBuffer = sSearch & String(MAXADDRESSSIZE - Len(sSearch), 0)

    nError_no = GBGetNext(p_strBuffer, MAXADDRESSSIZE)
    sMultiData = sMultiData & p_strBuffer
    
Wend

p_strBuffer = sMultiData
SearchGBAddress = True
            
Exit Function

SearchGBAddressErrortrap:
    ShutDown
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description
    SearchGBAddress = False
        
End Function

Public Function ShutDown() As Integer

Dim nReply As Integer
Dim szError_text As String

' Close the HLT

nReply = GBClose()

' IMPORTANT - we have closed the toolkit, so don't try and call
' GBGetError after GBClose - instead just return the reply code
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

