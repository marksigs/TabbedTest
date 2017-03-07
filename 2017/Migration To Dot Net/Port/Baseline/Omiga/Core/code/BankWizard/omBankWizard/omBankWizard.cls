VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omBankWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------
' Class Module  : omBankWizard
' File:         : omBankWizard.cls
' Author:       : Ian Woodbridge(IVW)
' Date:         : 12 October 2001
' Purpose:      : Retrieve Bank Details from Eiger Bank Wizard Software
' History:
'   AQR                 AUTHOR                      DESCRIPTION
'   ---                 ------                      -----------
'   ---                 Ian Woodbridge              Initial version
'------------------------------------------------------------------------
' Note
' This is generated as .EXE because the Eiger Bank Wizard software
' is not threadsafe.
Option Explicit
Private Const INVALID_XML_ERROR = vbObjectError + 1
Public Function CheckDetails(ByVal p_strXML As String) As String
Dim l_xmlDom As FreeThreadedDOMDocument40
Dim strError As String
Dim strSortCode As String
Dim strAccountNumber As String
Dim strBuffer As String
Dim bContinue As Boolean
Dim lSearchErr As Long
Dim UKDetails As UKDetailsType
On Error GoTo CheckDetailsErrortrap
    Set l_xmlDom = New FreeThreadedDOMDocument40
    l_xmlDom.async = False
    l_xmlDom.setProperty "NewParser", True
    l_xmlDom.validateOnParse = False
    l_xmlDom.loadXML p_strXML
    If l_xmlDom.parseError.errorCode <> 0 Then
        Err.Raise l_xmlDom.parseError.errorCode, "CheckDetails", "The XML string could not be parsed - parse error = " & l_xmlDom.parseError.reason
    Else
        bContinue = ExtractBankElements(l_xmlDom, strSortCode, strAccountNumber)
        If (bContinue) Then
            bContinue = CreateResponseNode(l_xmlDom)
        End If
        If (bContinue) Then
            bContinue = GetBranchDetails(UKDetails, l_xmlDom, strSortCode, strAccountNumber)
            If (bContinue) Then
                StoreBankDetailsAsXML UKDetails, l_xmlDom, strAccountNumber
            End If
        End If
    End If
    CheckDetails = l_xmlDom.xml
Exit Function:
CheckDetailsErrortrap:
    CreateErrorXML l_xmlDom, Err.Number, Err.Source, Err.Description, "APPERR"
    CheckDetails = l_xmlDom.xml
    Set l_xmlDom = Nothing
End Function
Private Function StoreBankDetailsAsXML(ByRef p_UKDetailsType As UKDetailsType, _
ByRef p_xmlDom As FreeThreadedDOMDocument40, _
ByRef strAccountNo As String) As Boolean
Dim l_xmlResponseElementNode As IXMLDOMElement
Dim l_xmlElementNode As IXMLDOMElement
Dim bContinue As Boolean
On Error GoTo StoreBankDetailsAsXMLErrorTrap:
    Set l_xmlResponseElementNode = p_xmlDom.selectSingleNode("/RESPONSE")
    l_xmlResponseElementNode.setAttribute "TYPE", "SUCCESS" ' assume SUCCESS
    With p_UKDetailsType
        ' check no errors exist in bank details response.
        ' note the special case for BWBADAN, this exists
        ' because if we don't REQUEST an account number the
        ' bank wizard dll returns an error despite correctly
        ' returning bank info for the given sort code.
        bContinue = False ' only continue if all error checks pass
        ' ERRORS - if error exists in Bank Wizard return code then raise BANKWIZARDERROR
        If (.Result And BWNOINI) = BWNOINI Then
            CreateErrorXML p_xmlDom, BWNOINI, "UKCheckDetails", "BWNOINI: BANK WIZARD has not been initialised", "BANKWIZARDERROR"
        ElseIf (.Result And BWBADAN And Len(Trim$(strAccountNo)) > 0) = BWBADAN Then ' only error on this if an account # was entered
            CreateErrorXML p_xmlDom, BWBADAN, "UKCheckDetails", "BWBADAN: Account number format incorrect", "BANKWIZARDERROR"
        ElseIf (.Result And BWBADSC) = BWBADSC Then
            CreateErrorXML p_xmlDom, BWBADSC, "UKCheckDetails", "BWBADSC: Sort-code format incorrect", "BANKWIZARDERROR"
        ElseIf (.Result And BWNOBBD) = BWNOBBD Then
            CreateErrorXML p_xmlDom, BWNOBBD, "UKCheckDetails", "BWNOBBD: BANK WIZARD has failed to open the Branch Database", "BANKWIZARDERROR"
        ElseIf (.Result And BWNOMEM) = BWNOMEM Then
            CreateErrorXML p_xmlDom, BWNOMEM, "UKCheckDetails", "BWNOMEM: Insufficient memory to initialise BANK WIZARD", "BANKWIZARDERROR"
        ElseIf (.Result And BWNOMOD) = BWNOMOD Then
            CreateErrorXML p_xmlDom, BWNOMOD, "UKCheckDetails", "BWNOMOD: The Database Table File UKMOD.TBL could not be found or opened", "BANKWIZARDERROR"
        ElseIf (.Result And BWNOSUB) = BWNOSUB Then
            CreateErrorXML p_xmlDom, BWNOSUB, "UKCheckDetails", "BWNOSUB: The Database Table File UKSUB.TBL could not be found or opened", "BANKWIZARDERROR"
        ElseIf (.Result And BWNOOFS) = BWNOOFS Then
            CreateErrorXML p_xmlDom, BWNOOFS, "UKCheckDetails", "BWNOOFS: The Database Table File UKOFS.TBL could not be found or opened", "BANKWIZARDERROR"
        ElseIf (.Result And BWNOLIC) = BWNOLIC Then
            CreateErrorXML p_xmlDom, BWNOLIC, "UKCheckDetails", "BWNOLIC: The Licence Key File ACCESS.KEY could not be found or opened", "BANKWIZARDERROR"
        ElseIf (.Result And BWEXLIC) = BWEXLIC Then
            CreateErrorXML p_xmlDom, BWEXLIC, "UKCheckDetails", "BWEXLIC: Your licence to use BANK WIZARD has expired -  BANK WIZARD will no longer function", "BANKWIZARDERROR"
        ElseIf (.Result And BWNOUNQ) = BWNOUNQ Then
            CreateErrorXML p_xmlDom, BWNOUNQ, "UKCheckDetails", "BWNOUNQ: The Database Table File UKUNIQ.TBL could not be found or opened ", "BANKWIZARDERROR"
        ElseIf (.Result And BWSCNBR) = BWSCNBR Then
            CreateErrorXML p_xmlDom, BWSCNBR, "UKCheckDetails", "BWSCNBR: Sort-code not allocated to any bank branch", "BANKWIZARDERROR"
        ElseIf (.Result And BWMODFL) = BWMODFL Then
            CreateErrorXML p_xmlDom, BWMODFL, "UKCheckDetails", "BWMODFL: Modulus check failed", "BANKWIZARDERROR"
        ElseIf (.Result And BWNOTRN) = BWNOTRN Then
            CreateErrorXML p_xmlDom, BWNOTRN, "UKCheckDetails", "BWNOTRN: The Database Table File UKTRAN.TBL could not be found or opened", "BANKWIZARDERROR"
        ElseIf (.Result And BWCLOSE) = BWCLOSE Then
            CreateErrorXML p_xmlDom, BWCLOSE, "UKCheckDetails", "BWCLOSE: This branch has closed", "BANKWIZARDERROR"
        ElseIf (.Result And BWCLOSR) = BWCLOSR Then
            CreateErrorXML p_xmlDom, BWCLOSR, "UKCheckDetails", "BWCLOSR: This branch has closed - see redirected to sort-code field", "BANKWIZARDERROR"
        Else
            ' Return all warnings with the bank details, warnings are not fatal and
            ' may contain additional bank information
            bContinue = True ' no errors exist so continue
            ' test ..........        .Result = BWBSDIS + BWSCNIR
            ' WARNINGS - if a warning exists in Bank Wizard then raise a BANKWIZARDWARNING
            If ((.Result And BWSCNIR) And bContinue) = BWSCNIR Then
                bContinue = CreateWarningXML(p_xmlDom, "Modulus Checking cannot be applied to this bank/branch ", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BWBYFOR) And bContinue) = BWBYFOR Then
                bContinue = CreateWarningXML(p_xmlDom, "Account is a Foreign Currency Account", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BWTXDIS) And bContinue) = BWTXDIS Then
                bContinue = CreateWarningXML(p_xmlDom, "This branch will not support Direct Debit transactions", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BWXSCRB) And bContinue) = BWXSCRB Then
                bContinue = CreateWarningXML(p_xmlDom, "Account details are not in standard form and were transposed", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BWDBDIS) And bContinue) = BWDBDIS Then
                bContinue = CreateWarningXML(p_xmlDom, "This account will not support Direct Debit transactions", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BWROLLN) And bContinue) = BWROLLN Then
                bContinue = CreateWarningXML(p_xmlDom, "Head Office collection account. Requires a Roll account number", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BWCRDIS) And bContinue) = BWCRDIS Then
                bContinue = CreateWarningXML(p_xmlDom, "This branch will not support Direct Credit transactions", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BW30DAY) And bContinue) = BW30DAY Then
                bContinue = CreateWarningXML(p_xmlDom, "Your licence to use BANK WIZARD will expire in less than 30 days", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BW2MONT) And bContinue) = BW2MONT Then
                bContinue = CreateWarningXML(p_xmlDom, "This branch is due to close", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BW2MONR) And bContinue) = BW2MONR Then
                bContinue = CreateWarningXML(p_xmlDom, "This branch is due to close - see redirected to sort-code field", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BWCUDIS) And bContinue) = BWCUDIS Then
                bContinue = CreateWarningXML(p_xmlDom, "This branch will not support Claims for Unpaid Cheque transactions", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BWPRDIS) And bContinue) = BWPRDIS Then
                bContinue = CreateWarningXML(p_xmlDom, "This branch will not support Life Assurance Premiums transactions", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BWBSDIS) And bContinue) = BWBSDIS Then
                bContinue = CreateWarningXML(p_xmlDom, "This branch will not support Building Society Credit transactions", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BWDVDIS) And bContinue) = BWDVDIS Then
                bContinue = CreateWarningXML(p_xmlDom, "This branch will not support Dividend Interest Payment transactions", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If ((.Result And BWAUDIS) And bContinue) = BWAUDIS Then
                bContinue = CreateWarningXML(p_xmlDom, "This branch will not support AUDDIS transactions", l_xmlResponseElementNode, "BANKWIZARDWARNING")
            End If
            If (bContinue) Then
                Set l_xmlElementNode = p_xmlDom.createElement("BANKDATA")
                l_xmlElementNode.setAttribute "SORTCODE", Trim$(Left$(.SortCode, .SortCodeLen))
                l_xmlElementNode.setAttribute "ACCOUNTNO", Trim$(Left$(.AccountNo, .AccountNoLen))
                l_xmlElementNode.setAttribute "ACCOUNTCODE", Trim$(Left$(.AccCode, .AccCodeLen))
                l_xmlElementNode.setAttribute "RESULT", Trim$(CStr(.Result))
                l_xmlElementNode.setAttribute "SUBBRANCHSUFFIX", Trim$(Left$(.SubBranchSuffix, .SubBranchSuffixLen))
                l_xmlElementNode.setAttribute "BRANCHTITLE", Trim$(Left$(.BranchTitle, .BranchTitleLen))
                l_xmlElementNode.setAttribute "SHORTNAME", Trim$(Left$(.ShortName, .ShortNameLen))
                l_xmlElementNode.setAttribute "DELETEDDATE", Trim$(Left$(.DeletedDate, .DeletedDateLen))
                l_xmlElementNode.setAttribute "BRANCHNAMEPLACE", Trim$(Left$(.BranchNamePlace, .BranchNamePlaceLen))
                l_xmlElementNode.setAttribute "ADDRESS1", Trim$(Left$(.Address1, .Address1Len))
                l_xmlElementNode.setAttribute "ADDRESS2", Trim$(Left$(.Address2, .Address2Len))
                l_xmlElementNode.setAttribute "ADDRESS3", Trim$(Left$(.Address3, .Address3Len))
                l_xmlElementNode.setAttribute "ADDRESS4", Trim$(Left$(.Address4, .Address4Len))
                l_xmlElementNode.setAttribute "TOWN", Trim$(Left$(.Town, .TownLen))
                l_xmlElementNode.setAttribute "COUNTY", Trim$(Left$(.County, .CountyLen))
                'l_xmlElementNode.setAttribute "POSTCODE1", Trim$(Left$(.PostCodeMajor, .PostCodeMajorLen))
                'l_xmlElementNode.setAttribute "POSTCODE2", Trim$(Left$(.PostCodeMinor, .PostCodeMinorLen))
                l_xmlElementNode.setAttribute "POSTCODE", Trim$(Left$(.PostCodeMajor, .PostCodeMajorLen)) + " " + Trim$(Left$(.PostCodeMinor, .PostCodeMinorLen))
                l_xmlElementNode.setAttribute "STDCODE", Trim$(Left$(.TelAreaCode, .TelAreaCodeLen))
                l_xmlElementNode.setAttribute "TELEPHONENUMBER", Trim$(Left$(.TelNumber, .TelNumberLen))
                l_xmlResponseElementNode.appendChild l_xmlElementNode
            End If
        End If
    End With
    Set l_xmlResponseElementNode = Nothing
    Set l_xmlElementNode = Nothing
    If (bContinue) Then
        StoreBankDetailsAsXML = True
    Else
        StoreBankDetailsAsXML = False
    End If
Exit Function:
StoreBankDetailsAsXMLErrorTrap:
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description, "APPERR"
    Set l_xmlResponseElementNode = Nothing
    Set l_xmlElementNode = Nothing
    StoreBankDetailsAsXML = False
End Function
Private Function CreateWarningXML(ByRef p_xmlDom As FreeThreadedDOMDocument40, _
ByRef p_strWarningMessage, _
ByRef p_ResponseXML As IXMLDOMElement, _
ByRef p_strWarningType)
Dim l_xmlWarningNode As IXMLDOMElement
Dim l_xmlSubNode As IXMLDOMElement
On Error GoTo CreateWarningXMLErrorTrap:
    p_ResponseXML.setAttribute "TYPE", p_strWarningType
    Set l_xmlWarningNode = p_xmlDom.createElement("MESSAGE")
    Set l_xmlSubNode = p_xmlDom.createElement("MESSAGETEXT")
    l_xmlSubNode.Text = p_strWarningMessage
    l_xmlWarningNode.appendChild l_xmlSubNode
    Set l_xmlSubNode = p_xmlDom.createElement("MESSAGETYPE")
    l_xmlSubNode.Text = "WARNING"
    l_xmlWarningNode.appendChild l_xmlSubNode
    p_ResponseXML.appendChild l_xmlWarningNode
    Set l_xmlWarningNode = Nothing
    Set l_xmlSubNode = Nothing
    CreateWarningXML = True
Exit Function:
CreateWarningXMLErrorTrap:
    
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description, "APPERR"
    Set l_xmlWarningNode = Nothing
    Set l_xmlSubNode = Nothing
    CreateWarningXML = False
End Function
Private Function GetBranchDetails(ByRef p_UKDetailsType As UKDetailsType, _
ByRef p_xmlDom As FreeThreadedDOMDocument40, _
ByRef strSortCode As String, _
ByRef strAccountNo As String) As Boolean
Dim ILen          As Integer
Dim mlngIndex     As Long
  
On Error GoTo GetBranchDetailsErrorTrap:
    With p_UKDetailsType
        .SortCode = String$(9, ".")
        .AccountNo = String$(13, ".")
        .AccCode = String$(2, ".")
        .SubBranchSuffix = String$(3, ".")
        .BranchTitle = String$(28, ".")
        .ShortName = String$(21, ".")
        .DeletedDate = String$(11, ".")
        .BACSStatus = String$(2, ".")
        .DateClosedBACS = String$(11, ".")
        .RedirFlag = String$(2, ".")
        .RedirSC = String$(7, ".")
        .DR = String$(3, ".")
        .CR = String$(3, ".")
        .CU = String$(3, ".")
        .PR = String$(3, ".")
        .BS = String$(3, ".")
        .DV = String$(3, ".")
        .AU = String$(3, ".")
        .SterlingStatus = String$(2, ".")
        .EuroStatus = String$(2, ".")
        .BranchTypeInd = String$(2, ".")
        .BranchNamePlace = String$(36, ".")
        .Address1 = String$(36, ".")
        .Address2 = String$(36, ".")
        .Address3 = String$(36, ".")
        .Address4 = String$(36, ".")
        .Town = String$(31, ".")
        .County = String$(31, ".")
        .PostCodeMajor = String$(5, ".")
        .PostCodeMinor = String$(5, ".")
        .TelAreaCode = String$(11, ".")
        .TelNumber = String$(9, ".")
        .FaxAreaCode = String$(11, ".")
        .FaxNumber = String$(9, ".")
            
' Intialise inputs for call to UKCheckDetails
' Ensure that input values are at least as long
' as the values to be returned - C cannot
' dynamically lengthen the strings
        ILen = Len(strSortCode)
        If ILen < 6 Then
            .SortCode = strSortCode + Space$(6 - ILen)
            .SortCodeLen = 6
        Else
            .SortCode = strSortCode
            .SortCodeLen = ILen
        End If
        ILen = Len(strAccountNo)
        If ILen < 8 Then
            .AccountNo = strAccountNo + Space$(8 - ILen)
            .AccountNoLen = 8
        Else
            .AccountNo = strAccountNo
            .AccountNoLen = ILen
        End If
        .InFlags = MIXEDCASE + BRCHECK + BRDETAILS
    End With
    UKCheckDetails p_UKDetailsType, mlngIndex
    GetBranchDetails = True
        
Exit Function:
GetBranchDetailsErrorTrap:
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description, "APPERR"
    GetBranchDetails = False
End Function
Private Function CreateErrorXML(ByRef p_xmlDom As FreeThreadedDOMDocument40, _
ByRef lNumber As Long, _
ByRef strSource As String, _
ByRef strDesc As String, _
ByRef p_strErrorType) As Boolean
    
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
        l_xmlResponseElement.setAttribute "TYPE", p_strErrorType
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
    CreateErrorXML = True
Exit Function:
CreateErrorXMLErrortrap:
        
    Set l_xmlResponseElement = Nothing
    Set l_xmlRequestNode = Nothing
    Set l_xmlResponseNode = Nothing
    Set l_xmlErrorElement = Nothing
    Set l_xmlErrorElementNode = Nothing
    Set l_xmlWorkingNode = Nothing
    Set l_OldNode = Nothing
    CreateErrorXML = False
        
End Function
Private Function CreateResponseNode(ByRef p_xmlDom As FreeThreadedDOMDocument40) As Boolean
Dim l_xmlResponseElementNode As IXMLDOMElement
Dim l_xmlNode As IXMLDOMNode
Dim l_xmlOldNode As IXMLDOMNode
On Error GoTo CreateResponseNodeErrortrap
    
    Set l_xmlResponseElementNode = p_xmlDom.createElement("RESPONSE")
    Set l_xmlNode = p_xmlDom.selectSingleNode("/REQUEST")
    If (l_xmlNode Is Nothing) Then
        Err.Raise INVALID_XML_ERROR, "CreateResponseNode", "No Request element present in Input XML"
    End If
    Set l_xmlOldNode = p_xmlDom.replaceChild(l_xmlResponseElementNode, l_xmlNode)
    Set l_xmlResponseElementNode = Nothing
    Set l_xmlNode = Nothing
    Set l_xmlOldNode = Nothing
    CreateResponseNode = True
Exit Function:
CreateResponseNodeErrortrap:
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description, "APPERR"
    Set l_xmlResponseElementNode = Nothing
    Set l_xmlNode = Nothing
    Set l_xmlOldNode = Nothing
    CreateResponseNode = False
End Function
Private Function ExtractBankElements(ByRef p_xmlDom As FreeThreadedDOMDocument40, _
ByRef p_strSortCode As String, _
ByRef p_strAccountNumber As String) As Boolean
            
'read input XML string.
Dim l_strNodeValue As String
Dim l_xmlNode As IXMLDOMNode
Dim l_xmlAttribute As IXMLDOMAttribute
On Error GoTo ExtractAddressElements
    
    Set l_xmlNode = p_xmlDom.selectSingleNode("/REQUEST/BANKDETAILS")
    If (l_xmlNode Is Nothing) Then
        Err.Raise INVALID_XML_ERROR, "ExtractBankElements", "BANKDETAILS Element missing from REQUEST XML"
    End If
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("SORTCODE")
    If Not l_xmlAttribute Is Nothing Then
        p_strSortCode = l_xmlAttribute.Text
    Else
        p_strSortCode = ""
    End If
    Trim$ (p_strSortCode)
    Set l_xmlAttribute = l_xmlNode.Attributes.getNamedItem("ACCOUNTNUMBER")
    If Not l_xmlAttribute Is Nothing Then
        p_strAccountNumber = l_xmlAttribute.Text
    Else
        p_strAccountNumber = ""
    End If
    Trim$ (p_strAccountNumber)
    Set l_xmlNode = Nothing
    Set l_xmlAttribute = Nothing
    ExtractBankElements = True
Exit Function:
ExtractAddressElements:
    
    CreateErrorXML p_xmlDom, Err.Number, Err.Source, Err.Description, "APPERR"
    Set l_xmlNode = Nothing
    Set l_xmlAttribute = Nothing
    ExtractBankElements = False
End Function
