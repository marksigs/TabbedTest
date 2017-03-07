VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "QuickQuoteBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      QuickQuoteBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Quick Quote Business Object
'
'Dependencies:  CustomerDO, ApplicationDO, QuickQuoteApplicationDetailsDO,
'               ComboDO, PaymentProtestcionSubQuoteBO, MortgageSubQuoteDO
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    11/11/99    Created
'DJP    09/02/00    Added CalculateAndSaveBCSubQuote
'DJP    28/02/00    Added FindMortgageProducts
'DJP    29/02/00    Added GetDefaultsForNewLoanComponent
'DJP    29/02/00    Added GetLoanCompositionDetails
'IW     23/05/00    SYS0774 DISTRIBUTIONCHANNELID S/B/ CHANNELID
'JLD    26/05/00    SYS0501 in getHighestEarners make sure call to quickQuoteApplicantDetailsDO.FindList
'                   gets all applicants for the application.
'APS    06/06/00    SYS0786 - Life cover quote now not passing in "Joint" regardless of the
'                   number of applcants
'MS     08/07/00    SYS0824 - UK residency NOT a criterion for PP Eligibility
'PSC    30/06/00    SYS1012 - Call ICostModelAssitBO.ProcessAddedOneOffCosts rather than
'                             IQuickQuoteTxBO.ProcessAddedOneOffCosts
'DJP    13/09/00    SYS1031 - Change FindMortgageProducts to correctly handle self employed applicant
'CL     07/05/02    SYS4510  Modifications to class initialise
'CL     10/05/02    SYS4510  Remove class initialize & class terminate
'SG     28/05/02    SYS4767 MSMS to Core integration
'------------------------------------------------------------------------------------------
Option Explicit
' Enables us to have per class variables .e.g., error assist, XML assist etc.
Implements ObjectControl
Implements IQuickQuoteBO
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Private m_clsConvertAssist As ConvertAssist
Private m_objContext As ObjectContext
 
 '=============================================
 'Constant Declaration Section
 '=============================================
Private Enum ApplicantType
    ApplicantOne = 1
    ApplicantTwo = 2
End Enum

 '=============================================
 'Variable Declaration Section
 '=============================================

Private Const cstrTABLE_NAME = "QUICKQUOTE"
'SG 28/05/02 SYS4767 Added function

'------------------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'MO     02/07/2002    Public function ValidateManualAdjustmentPercent added as it was missed
'                   from the Core upgrade, so presumably when they add it to the core we
'                   remove it
'MV     18/07/2002  BMIDS00179 - CoreUpgrade Rollback - Modified GetValidatedQuotationDetails()
'GD     18/07/2002  BMIDS00045 - Remove Calculations from CalculateAndSavePPCosts
'MV     12/08/2002  BMIDS00311 - Core Ref : SYS1705 ;Modified QuickQuoteBO_CalculateMortgageCosts
'MV     02/09/2002  BMIDS00353  - Amended IQuickQuoteBO_GetLoanCompositionDetails
'GD     03/09/2002  BMIDS00256  - Amended GetValidatedQuotationDetails
'GD     12/09/02    BMIDS00431 - Do not calculate the LTV
'JD     19/07/04    BMIDS0763   - pass applicationDate to GetLoanCompositionDetails
'HMA    21/02/05    BMIDS959    - Correct ValidateManualAdjustmentPercent.
'------------------------------------------------------------------------------------------

Private Sub Class_Initialize()
    If m_objErrAssist Is Nothing Then
        Set m_objErrAssist = New ErrAssist
    End If

    If m_objXmlAssist Is Nothing Then
        Set m_objXmlAssist = New XMLAssist
    End If

    If m_clsConvertAssist Is Nothing Then
        Set m_clsConvertAssist = New ConvertAssist
    End If

End Sub

Private Sub Class_Terminate()
    Set m_objErrAssist = Nothing
End Sub

Private Function IQuickQuoteBO_CalculateAndSaveBCSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CalculateAndSaveBCSubQuoteVbErr
    
    Const strFunctionName As String = "IQuickQuoteBO_CalculateAndSaveBCSubQuote"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuickQuoteTxBO As IQuickQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIQuickQuoteTxBO = New QuickQuoteTxBO
    Else
        Set objIQuickQuoteTxBO = m_objContext.CreateInstance(App.Title & ".QuickQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIQuickQuoteTxBO.CalculateAndSaveBCSubQuote(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IQuickQuoteBO_CalculateAndSaveBCSubQuote = xmlTempResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CalculateAndSaveBCSubQuoteExit:

    Set objIQuickQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CalculateAndSaveBCSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_CalculateAndSaveBCSubQuote = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalculateAndSaveBCSubQuoteExit

End Function

Private Function IQuickQuoteBO_CreateNewQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CreateNewQuotationVbErr
    
    Const strFunctionName As String = "IQuickQuoteBO_CreateNewQuotation"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuickQuoteTxBO As IQuickQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem

    If m_objContext Is Nothing Then
        Set objIQuickQuoteTxBO = New QuickQuoteTxBO
    Else
        Set objIQuickQuoteTxBO = m_objContext.CreateInstance(App.Title & ".QuickQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIQuickQuoteTxBO.CreateNewQuotation(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IQuickQuoteBO_CreateNewQuotation = xmlTempResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateNewQuotationExit:

    Set objIQuickQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateNewQuotationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_CreateNewQuotation = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewQuotationExit

End Function

Private Function IQuickQuoteBO_ProcessAddedOneOffCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ProcessAddedOneOffCostsVbErr
    
    Const strFunctionName As String = "IQuickQuoteBO_ProcessAddedOneOffCosts"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICostModelAssistBO As ICostModelAssistBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If
    
    ' PSC 30/06/00 SYS1012
    Set xmlTempResponseNode = objICostModelAssistBO.ProcessAddedOneOffCosts(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IQuickQuoteBO_ProcessAddedOneOffCosts = xmlTempResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ProcessAddedOneOffCostsExit:

    Set objICostModelAssistBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ProcessAddedOneOffCostsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_ProcessAddedOneOffCosts = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ProcessAddedOneOffCostsExit
End Function

Private Function IQuickQuoteBO_ReinstateQuotation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ReinstateQuotationVbErr
    
    Const strFunctionName As String = "IQuickQuoteBO_ReinstateQuotation"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuickQuoteTxBO As IQuickQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIQuickQuoteTxBO = New QuickQuoteTxBO
    Else
        Set objIQuickQuoteTxBO = m_objContext.CreateInstance(App.Title & ".QuickQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIQuickQuoteTxBO.ReinstateQuotation(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    ' If every thing is successful, add message 217 as Warning and send it back
    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 217
    Set IQuickQuoteBO_ReinstateQuotation = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ReinstateQuotationExit:

    Set objIQuickQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ReinstateQuotationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_ReinstateQuotation = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ReinstateQuotationExit

End Function

Private Function IQuickQuoteBO_ResetMortgageSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

On Error GoTo ResetMortgageSubQuoteVbErr
    
    Const strFunctionName As String = "IQuickQuoteBO_ResetMortgageSubQuote"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuickQuoteTxBO As IQuickQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIQuickQuoteTxBO = New QuickQuoteTxBO
    Else
        Set objIQuickQuoteTxBO = m_objContext.CreateInstance(App.Title & ".QuickQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIQuickQuoteTxBO.ResetMortgageSubQuote(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    Set IQuickQuoteBO_ResetMortgageSubQuote = xmlTempResponseNode
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ResetMortgageSubQuoteExit:

    Set objIQuickQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ResetMortgageSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_ResetMortgageSubQuote = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ResetMortgageSubQuoteExit

End Function

Private Function IQuickQuoteBO_UpdateMortgageSubQuote(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Updates Mortgage Subquote.
'
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>
'   <MORTGAGESUBQUOTE>
'       <APPLICATIONNUMBER>B00003239</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'   </MORTGAGESUBQUOTE>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageSubQuoteVbErr

    Const strFunctionName As String = "IQuickQuoteBO_UpdateMortgageSubQuote"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objICostModelAssistBO As ICostModelAssistBO
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If

    ' call common costModelAssist method
    Set xmlTempResponseNode = objICostModelAssistBO.UpdateMortgageSubQuote(vxmlRequest)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    Set IQuickQuoteBO_UpdateMortgageSubQuote = xmlResponseElem
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateMortgageSubQuoteExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateMortgageSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_UpdateMortgageSubQuote = m_objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageSubQuoteExit
End Function

Private Function IQuickQuoteBO_ValidateManualAdjustmentPercent(ByVal vxmlRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
On Error GoTo ValidateManualAdjustmentVbErr
    
    Const strFunctionName As String = "IQuickQuoteBO_ValidateManualAdjustmentPercent"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlCompetency As New FreeThreadedDOMDocument40
    xmlCompetency.validateOnParse = False
    xmlCompetency.setProperty "NewParser", True
    Dim xmlCompElem As IXMLDOMElement
    Dim dblUserDecreaseLimitPercent As Double
    Dim dblUserIncreaseLimitPercent As Double
    Dim dblRateAdjustment As Double
    Dim xmlProduct As New FreeThreadedDOMDocument40
    xmlProduct.validateOnParse = False
    xmlProduct.setProperty "NewParser", True
    Dim xmlProductElem As IXMLDOMElement
    Dim xmlProductNode As IXMLDOMNode
    Dim xmlTempElement As IXMLDOMElement
    Dim xmlKeyElement As IXMLDOMElement
    Dim xmlProdResponse As New FreeThreadedDOMDocument40
    xmlProdResponse.validateOnParse = False
    xmlProdResponse.setProperty "NewParser", True
    Dim xmlProdNode As IXMLDOMNode
        
    Dim objIOrganisationDO As IOrganisationDO
    Dim objIMortgageProductBO As MortgageProductBO
    Dim objConvAssist As New ConvertAssist
    
    If m_objContext Is Nothing Then
        Set objIOrganisationDO = New OrganisationDO
        Set objIMortgageProductBO = New MortgageProductBO
    Else
        Set objIOrganisationDO = m_objContext.CreateInstance(gstrORGANISATION_COMPONENT & ".OrganisationDO")
        Set objIMortgageProductBO = m_objContext.CreateInstance(gstrMORTGAGEPRODUCT & ".MortgageProductBO")
    End If
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    'Validate users adjustment limits
    Set xmlCompElem = xmlCompetency.createElement("REQUEST")
    xmlCompElem.setAttribute "USERID", vxmlRequest.getAttribute("USERID")
    xmlCompetency.appendChild xmlCompElem
    Set xmlTempResponseNode = xmlCompElem.appendChild(xmlOut.createElement("USERID"))
    xmlTempResponseNode.Text = vxmlRequest.getAttribute("USERID")
    
    'Set xmlTempResponseNode = objIOrganisationDO.GetCurrentCompetency(xmlCompElem)
    Set xmlTempResponseNode = objIOrganisationDO.ValidateUserMandateLevel(xmlCompElem)
    'm_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    dblRateAdjustment = objConvAssist.CSafeDbl(vxmlRequest.getElementsByTagName("MANUALADJUSTMENTPERCENT").Item(0).Text)
    dblUserDecreaseLimitPercent = objConvAssist.CSafeDbl(xmlTempResponseNode.selectSingleNode("USERDECREASELIMITPERCENT").Text)
    dblUserIncreaseLimitPercent = objConvAssist.CSafeDbl(xmlTempResponseNode.selectSingleNode("USERINCREASELIMITPERCENT").Text)
    
    ' JLD MSMS0058
    If dblRateAdjustment < 0 Then
        If Abs(dblRateAdjustment) > dblUserDecreaseLimitPercent Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 6001
        End If
    Else
        If Abs(dblRateAdjustment) > dblUserIncreaseLimitPercent Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 6001
        End If
    End If
    
    'Validate product parameter levels
    Set xmlProductElem = xmlProduct.createElement("REQUEST")
    xmlProduct.appendChild xmlProductElem
    Set xmlKeyElement = xmlProduct.createElement("MORTGAGEPRODUCT")
    xmlProductElem.appendChild xmlKeyElement
    Set xmlTempElement = xmlProduct.createElement("MORTGAGEPRODUCTCODE")
    xmlTempElement.Text = vxmlRequest.getElementsByTagName("MORTGAGEPRODUCTCODE").Item(0).Text
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("STARTDATE")
    xmlTempElement.Text = vxmlRequest.getElementsByTagName("STARTDATE").Item(0).Text
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("AMOUNTREQUESTED")
    If Not vxmlRequest.getElementsByTagName("AMOUNTREQUESTED").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("AMOUNTREQUESTED").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("PURCHASEPRICE")
    If Not vxmlRequest.getElementsByTagName("PURCHASEPRICE").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("PURCHASEPRICE").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("LTV")
    If Not vxmlRequest.getElementsByTagName("LTV").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("LTV").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("TYPEOFAPPLICATION")
    If Not vxmlRequest.getElementsByTagName("TYPEOFAPPLICATION").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("TYPEOFAPPLICATION").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("LOCATION")
    If Not vxmlRequest.getElementsByTagName("LOCATION").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("LOCATION").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("TYPEOFVALUATION")
    If Not vxmlRequest.getElementsByTagName("TYPEOFVALUATION").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("TYPEOFVALUATION").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    Set xmlTempElement = xmlProduct.createElement("FEETYPE")
    If Not vxmlRequest.getElementsByTagName("FEETYPE").Item(0) Is Nothing Then
        xmlTempElement.Text = vxmlRequest.getElementsByTagName("FEETYPE").Item(0).Text
    Else
        xmlTempElement.Text = "1"
    End If
    xmlKeyElement.appendChild xmlTempElement
    
    xmlProdResponse.loadXML (objIMortgageProductBO.GetFullProductDetails(xmlProduct.xml))
    m_objErrAssist.CheckResponse xmlProdResponse.documentElement, xmlResponseElem, True
    
    Set xmlProdNode = xmlProdResponse.documentElement.selectSingleNode(".//INTERESTRATETYPELIST/INTERESTRATETYPE[INTERESTRATETYPESEQUENCENUMBER='1']")
    If Not xmlProdNode Is Nothing Then
        If xmlProdNode.selectSingleNode("RATETYPE").Text = "F" Or _
           xmlProdNode.selectSingleNode("RATETYPE").Text = "B" Or _
           xmlProdNode.selectSingleNode("RATETYPE").Text = "D" Then
            Dim dblDiffLowerLimit As Double
            Dim dblDiffUpperLimit As Double
            dblDiffLowerLimit = objConvAssist.CSafeDbl(xmlProdResponse.documentElement.selectSingleNode(".//DIFFERENTIALLOWERLIMIT").Text)
            dblDiffUpperLimit = objConvAssist.CSafeDbl(xmlProdResponse.documentElement.selectSingleNode(".//DIFFERENTIALUPPERLIMIT").Text)
            
            'BMIDS959 Correct limit check
            If dblRateAdjustment < 0 Then
                If Abs(dblRateAdjustment) > dblDiffLowerLimit Then
                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 6002
                End If
            Else
                If Abs(dblRateAdjustment) > dblDiffUpperLimit Then
                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 6002
                End If
            End If
        Else
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 6003
        End If
    End If
    
    ' Check that the resulting resolved rate does not go beneath the decrease limit percent
    Dim dblResolvedRate As Double
    Dim dblProdLimit As Double
    Dim strProdLimit As String
    strProdLimit = xmlProdResponse.documentElement.selectSingleNode(".//MANUALDECREASELIMITPERCENT").Text
    dblProdLimit = objConvAssist.CSafeDbl(strProdLimit)
    dblResolvedRate = objConvAssist.CSafeDbl(vxmlRequest.getElementsByTagName("RESOLVEDRATE").Item(0).Text)
    If dblResolvedRate < dblProdLimit Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 6006, "", strProdLimit
    End If
    
    Set IQuickQuoteBO_ValidateManualAdjustmentPercent = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ValidateManualAdjustmentExit:
    
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlCompetency = Nothing
    Set xmlCompElem = Nothing
    Set xmlProduct = Nothing
    Set xmlProductElem = Nothing
    Set xmlProductNode = Nothing
    Set xmlTempElement = Nothing
    Set xmlKeyElement = Nothing
    Set xmlProdResponse = Nothing
    Set xmlProdNode = Nothing

    Exit Function

ValidateManualAdjustmentVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_ValidateManualAdjustmentPercent = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateManualAdjustmentExit

End Function

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    If m_objErrAssist Is Nothing Then
        Set m_objErrAssist = New ErrAssist
    End If

    If m_objXmlAssist Is Nothing Then
        Set m_objXmlAssist = New XMLAssist
    End If

    If m_clsConvertAssist Is Nothing Then
        Set m_clsConvertAssist = New ConvertAssist
    End If

    Set m_objContext = GetObjectContext()
End Sub

Private Sub ObjectControl_Deactivate()
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
End Sub

Public Function CalculateRecommendedPPCover(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calculates the recommended payment protection cover
'
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CalculateRecommendedPPCoverVbErr
    
    Const strFunctionName As String = "CalculateRecommendedPPCover"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objPPSubQuoteBO As PaymentProtectionSubQuoteBO
    Dim objIApplicationDO As IApplicationDO
    Dim objIMtgSubQuoteDO As IMortgageSubQuoteDO
    Dim objComboDO As ComboDO
    Dim xmlData As IXMLDOMElement
    Dim xmlPPSubQuoteDoc As New FreeThreadedDOMDocument40
    xmlPPSubQuoteDoc.validateOnParse = False
    xmlPPSubQuoteDoc.setProperty "NewParser", True
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objPPSubQuoteBO = New PaymentProtectionSubQuoteBO
        Set objIApplicationDO = New ApplicationDO
        Set objIMtgSubQuoteDO = New MortgageSubQuoteDO
        Set objComboDO = New ComboDO
    Else
        Set objPPSubQuoteBO = objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".PaymentProtectionSubQuoteBO")
        Set objIApplicationDO = objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
        Set objIMtgSubQuoteDO = objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteDO")
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
    
    Dim strApplicationNo As String
    Dim strApplicationFFNo As String
    Dim strMtgSubQuoteNo As String
    Dim strTypeofApplication As String
    Dim strTagName As String
    
    strTagName = "APPLICATIONNUMBER"
    strApplicationNo = objXmlAssist.GetTagValue(xmlRequestNode, strTagName)
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    strApplicationFFNo = objXmlAssist.GetTagValue(xmlRequestNode, strTagName)
    
    strTagName = "MORTGAGESUBQUOTENUMBER"
    strMtgSubQuoteNo = objXmlAssist.GetTagValue(xmlRequestNode, strTagName)
    
    strTagName = "TYPEOFAPPLICATION"
    strTypeofApplication = objXmlAssist.GetTagValue(xmlRequestNode, strTagName)
    
    If Len(strApplicationNo) = 0 Or Len(strApplicationFFNo) = 0 Or _
       Len(strMtgSubQuoteNo) = 0 Or Len(strTypeofApplication) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "APPLICATIONNUMBER = " & strApplicationNo & _
                                "; APPLICATIONFACTFINDNUMBER = " & strApplicationFFNo & _
                                "; MORTGAGESUBQUOTENUMBER = " & strMtgSubQuoteNo & _
                                "; TYPEOFAPPLICATION = " & strTypeofApplication

    End If

    Dim strData As String
    Dim strNewRepayments As String
    Dim strExistingRepayments As String
    Dim xmlSubQuoteElem As IXMLDOMElement
    
    objXmlAssist.ChangeNodeName xmlRequestNode, cstrTABLE_NAME, "MORTGAGESUBQUOTE"
    Set xmlSubQuoteElem = objIMtgSubQuoteDO.GetData(xmlRequestNode)
    
    m_objXmlAssist.CheckNode xmlSubQuoteElem, TypeName(Me), strFunctionName
    
    strNewRepayments = objXmlAssist.GetTagValue(xmlSubQuoteElem, _
                                                "TOTALNETMONTHLYCOST")

    ' If a further advance get the existing repayments
    If objComboDO.IsFurtherAdvance(strTypeofApplication) Then
        objXmlAssist.ChangeNodeName xmlRequestNode, "MORTGAGESUBQUOTE", "APPLICATION"
        Set xmlData = objIApplicationDO.GetApplicationData(xmlRequestNode)
    
        strExistingRepayments = objXmlAssist.GetTagValue(xmlData, _
                                                         "TOTALMONTHLYREPAYMENT")
    End If
    
    strTagName = "PAYMENTPROTECTIONSUBQUOTE"
    Set xmlTableElem = xmlPPSubQuoteDoc.createElement(strTagName)
    xmlPPSubQuoteDoc.appendChild xmlTableElem
    
    strTagName = "NEWREPAYMENTS"
    Set xmlElement = xmlPPSubQuoteDoc.createElement(strTagName)
    xmlElement.Text = strNewRepayments
    xmlTableElem.appendChild xmlElement
    
    strTagName = "EXISTINGREPAYMENTS"
    Set xmlElement = xmlPPSubQuoteDoc.createElement(strTagName)
    xmlElement.Text = strExistingRepayments
    xmlTableElem.appendChild xmlElement
            
    strData = objPPSubQuoteBO.CalculateRecommendedCover(xmlPPSubQuoteDoc.xml)
    objErrAssist.CheckXMLResponse strData, True
    
    CalculateRecommendedPPCover = strData

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CalculateRecommendedPPCoverExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlRequestNode = Nothing
    Set objXmlAssist = Nothing
    Set objPPSubQuoteBO = Nothing
    Set objIApplicationDO = Nothing
    Set objIMtgSubQuoteDO = Nothing
    Set objComboDO = Nothing
    Set xmlPPSubQuoteDoc = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlData = Nothing

    Exit Function

CalculateRecommendedPPCoverVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateRecommendedPPCover = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalculateRecommendedPPCoverExit
    
End Function

Public Function CheckApplicantPPEligibility(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calculates the recommended payment protection cover
'
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CheckApplicantPPEligibilityVbErr
    
    Const strFunctionName As String = "CheckApplicantPPEligibility"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objPPSubQuoteBO As PaymentProtectionSubQuoteBO
    Dim objIApplicationDO As IApplicationDO
    Dim xmlData As IXMLDOMElement
    Dim xmlPPSubQuoteDoc As New FreeThreadedDOMDocument40
    xmlPPSubQuoteDoc.validateOnParse = False
    xmlPPSubQuoteDoc.setProperty "NewParser", True
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlCust1RoleElem As IXMLDOMElement
    Dim xmlCust2RoleElem As IXMLDOMElement
        
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objPPSubQuoteBO = New PaymentProtectionSubQuoteBO
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objPPSubQuoteBO = objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".PaymentProtectionSubQuoteBO")
        Set objIApplicationDO = objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
    End If
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
    
    Dim strApplicationNo As String
    Dim strApplicationFFNo As String
    Dim strTagName As String
    
    strTagName = "APPLICATIONNUMBER"
    strApplicationNo = objXmlAssist.GetTagValue(xmlRequestNode, strTagName)
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    strApplicationFFNo = objXmlAssist.GetTagValue(xmlRequestNode, strTagName)
    
    If Len(strApplicationNo) = 0 Or Len(strApplicationFFNo) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "APPLICATIONNUMBER = " & strApplicationNo & _
                                "; APPLICATIONFACTFINDNUMBER = " & strApplicationFFNo

    End If
        
    Dim strData As String
    
    objXmlAssist.ChangeNodeName xmlRequestNode, cstrTABLE_NAME, "APPLICATION"
    Set xmlData = objIApplicationDO.FindMainApplicants(xmlRequestNode)
        
    Dim strCust1No As String
    Dim strCust1VerNo As String
'    Dim strCust1UKResident As String
    Dim strCust1EmpStatus As String
    Dim strCust2No As String
    Dim strCust2VerNo As String
'    Dim strCust2UKResident As String
    Dim strCust2EmpStatus As String
    
    Set xmlCust1RoleElem = xmlData.getElementsByTagName("CUSTOMERROLE").Item(0)

'    MS SYS0824 UK residency NOT a criterion for PP Eligibility
'    GetResidencyAndEmploymentStatus xmlCust1RoleElem, _
'                                    strCust1UKResident, _
'                                    strCust1EmpStatus
    GetResidencyAndEmploymentStatus xmlCust1RoleElem, _
                                    "", _
                                    strCust1EmpStatus
    
    
    Set xmlCust2RoleElem = xmlData.getElementsByTagName("CUSTOMERROLE").Item(1)
    
    If Not xmlCust2RoleElem Is Nothing Then

'    MS SYS0824 UK residency NOT a criterion for PP Eligibility

'        GetResidencyAndEmploymentStatus xmlCust2RoleElem, _
'                                        strCust2UKResident, _
'                                        strCust2EmpStatus
        GetResidencyAndEmploymentStatus xmlCust2RoleElem, _
                                        "", _
                                        strCust2EmpStatus
    
    End If
    
    
    strTagName = "PAYMENTPROTECTIONSUBQUOTE"
    Set xmlTableElem = xmlPPSubQuoteDoc.createElement(strTagName)
    xmlPPSubQuoteDoc.appendChild xmlTableElem
        
'    MS SYS0824 UK residency NOT a criterion for PP Eligibility
'    strTagName = "APPLICANT1UK"
'    Set xmlElement = xmlPPSubQuoteDoc.createElement(strTagName)
'    xmlElement.Text = strCust1UKResident
'    xmlTableElem.appendChild xmlElement
    
    strTagName = "APPLICANT1EMPLOYMENTSTATUS"
    Set xmlElement = xmlPPSubQuoteDoc.createElement(strTagName)
    xmlElement.Text = strCust1EmpStatus
    xmlTableElem.appendChild xmlElement
    
'    MS SYS0824 UK residency NOT a criterion for PP Eligibility
'    strTagName = "APPLICANT2UK"
'    Set xmlElement = xmlPPSubQuoteDoc.createElement(strTagName)
'    xmlElement.Text = strCust2UKResident
'    xmlTableElem.appendChild xmlElement
    
    strTagName = "APPLICANT2EMPLOYMENTSTATUS"
    Set xmlElement = xmlPPSubQuoteDoc.createElement(strTagName)
    xmlElement.Text = strCust2EmpStatus
    xmlTableElem.appendChild xmlElement
            
    strData = objPPSubQuoteBO.CheckApplicantEligibility(xmlPPSubQuoteDoc.xml)
    objErrAssist.CheckXMLResponse strData, True
        
    CheckApplicantPPEligibility = strData
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CheckApplicantPPEligibilityExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set objErrAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objXmlAssist = Nothing
    Set objPPSubQuoteBO = Nothing
    Set objIApplicationDO = Nothing
    Set xmlData = Nothing
    Set xmlPPSubQuoteDoc = Nothing
    Set xmlTableElem = Nothing
    Set xmlElement = Nothing
    Set xmlCust1RoleElem = Nothing
    Set xmlCust2RoleElem = Nothing

    Exit Function

CheckApplicantPPEligibilityVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CheckApplicantPPEligibility = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CheckApplicantPPEligibilityExit
    
End Function

Private Sub GetResidencyAndEmploymentStatus(ByVal vxmlCustomerRoleElem As IXMLDOMElement, _
                                            ByRef rstrUKResident As String, _
                                            ByRef rstrEmploymentStatus As String)
' header ----------------------------------------------------------------------------------
' description:  Gets the residency and employment status for a customer
' pass:         vxmlCustomerRoleElem  xml element for the customer role
'               rstrUKResident        whether the customer is UK resident
'               rstrEmploymentStatus  employment status of the customer
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo GetResidencyAndEmploymentStatusVbErr

    Const strFunctionName As String = "GetResidencyAndEmploymentStatus"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    'Dim objICustomerDO As ICustomerDO
    Dim objComboDO As ComboDO
    Dim objQQApplicantDetailsDO As QuickQuoteApplicantDetailsDO
    Dim xmlDataDoc As FreeThreadedDOMDocument40
    Dim xmlData As IXMLDOMElement

    If objContext Is Nothing Then
        'Set objICustomerDO = New CustomerDO
        Set objQQApplicantDetailsDO = New QuickQuoteApplicantDetailsDO
        Set objComboDO = New ComboDO
    Else
        'Set objICustomerDO = objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerDO")
        Set objQQApplicantDetailsDO = objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsDO")
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    'objXmlAssist.ChangeNodeName vxmlCustomerRoleElem, "CUSTOMERROLE", "CUSTOMER"

    Dim strData As String
    Dim strTagName As String
    
    'Set xmlData = objICustomerDO.GetPersonalDetails(vxmlCustomerRoleElem)
    
    'strTagName = "NORMALLYRESIDENTINUK"
    'rstrUKResident = objXmlAssist.GetTagValue(xmlData, strTagName)
    
    objXmlAssist.ChangeNodeName vxmlCustomerRoleElem, "CUSTOMERROLE", "QUICKQUOTEAPPLICANTDETAILS"
    
    strData = objQQApplicantDetailsDO.GetData(vxmlCustomerRoleElem.xml)
    Set xmlDataDoc = objXmlAssist.load(strData, TypeName(Me), strFunctionName)
    
    Dim strEmpStatusValue As String
    
'    strTagName = "UKTAXPAYER"
'    rstrUKResident = objXmlAssist.GetTagValue(xmlDataDoc.documentElement, strTagName)
    
    strTagName = "EMPLOYMENTSTATUS"
    strEmpStatusValue = objXmlAssist.GetTagValue(xmlDataDoc.documentElement, strTagName)
        
    rstrEmploymentStatus = objComboDO.GetFirstComboValidation("EmploymentStatus", _
                                                              strEmpStatusValue)
    
    objXmlAssist.ChangeNodeName vxmlCustomerRoleElem, _
                                "QUICKQUOTEAPPLICANTDETAILS", _
                                "CUSTOMERROLE"
   
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    'Set objICustomerDO = Nothing
    Set objComboDO = Nothing
    Set objQQApplicantDetailsDO = Nothing
    Set xmlDataDoc = Nothing
    Set xmlData = Nothing

    Exit Sub

GetResidencyAndEmploymentStatusVbErr:
    
    Set objXmlAssist = Nothing
    'Set objICustomerDO = Nothing
    Set objComboDO = Nothing
    Set objQQApplicantDetailsDO = Nothing
    Set xmlDataDoc = Nothing
    Set xmlData = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description


End Sub

Private Function GetDateOfBirth(ByVal sCustomerNumber As String, ByVal sCustomerVersion As String, ByVal xmlResp As IXMLDOMElement, objICustomerDO As ICustomerDO) As String
    Dim strTagName As String
    Dim strDateOfBirth As String
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlCustNoElem As IXMLDOMElement
    Dim xmlCustVerNoElem As IXMLDOMElement
    Dim xmlParentNode As IXMLDOMNode
    Dim xmlCustomerDoc As New FreeThreadedDOMDocument40
    xmlCustomerDoc.validateOnParse = False
    xmlCustomerDoc.setProperty "NewParser", True
    Dim xmlCustomerResp As IXMLDOMElement
    
Const strFunctionName As String = "GetDateOfBirth"
    On Error GoTo GetDataOfBirthErr
    
    Set xmlParentNode = xmlCustomerDoc.appendChild(xmlCustomerDoc.createElement("REQUEST"))
    
    strTagName = "CUSTOMER"
    Set xmlParentNode = xmlParentNode.appendChild(xmlCustomerDoc.createElement(strTagName))
    
    strTagName = "CUSTOMERNUMBER"
    Set xmlCustNoElem = xmlParentNode.appendChild(xmlCustomerDoc.createElement(strTagName))
    
    strTagName = "CUSTOMERVERSIONNUMBER"
    Set xmlCustVerNoElem = xmlParentNode.appendChild(xmlCustomerDoc.createElement(strTagName))
    
    xmlCustNoElem.Text = sCustomerNumber
    xmlCustVerNoElem.Text = sCustomerVersion
    
    ' Call GetPersonalDetails
    Set xmlCustomerResp = objICustomerDO.GetPersonalDetails(xmlCustomerDoc.documentElement)
    
    strTagName = "DATEOFBIRTH"
    
    ' Find the date of birth in the XML returned
    strDateOfBirth = m_objXmlAssist.GetTagValue(xmlCustomerResp, strTagName)

    ' Is it valid?
    If IsDate(CDate(strDateOfBirth)) = False Then
       m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInValidKeyValue, _
                                "Invalid date of birth:" & strDateOfBirth
        
    End If

    GetDateOfBirth = strDateOfBirth

GetDataOfBirthExit:
    Exit Function

GetDataOfBirthErr:
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Public Function CalculateAndSavePPCosts(ByVal vstrXMLRequest As String) As String

' header ----------------------------------------------------------------------------------
' description:  Calculates the payment protection and saves the results to the payment
'               protection sub quote
'               GD BMIDS00045 - Removed Calculation
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CalculateAndSavePPCostsVbErr
    
    Const strFunctionName As String = "CalculateAndSavePPCosts"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim xmlRequestElement As IXMLDOMElement

    Dim objPPSubQuoteBO As PaymentProtectionSubQuoteBO
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objPPSubQuoteBO = New PaymentProtectionSubQuoteBO
    Else
        Set objPPSubQuoteBO = objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".PaymentProtectionSubQuoteBO")
    End If
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlRequestElement = xmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
        
    If xmlRequestElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
    
    Dim strApplicationNo As String
    Dim strApplicationFFNo As String
    Dim strPPSubQuoteNo As String
    Dim strCoverType  As String
    Dim strPPReferenceNumber As String
    Dim strTotalPPMonthlyCost As String
    Dim strTagName As String
    
    strTagName = "APPLICATIONNUMBER"
    strApplicationNo = m_objXmlAssist.GetTagValue(xmlRequestElement, strTagName)
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    strApplicationFFNo = m_objXmlAssist.GetTagValue(xmlRequestElement, strTagName)
    
    strTagName = "PPSUBQUOTENUMBER"
    strPPSubQuoteNo = m_objXmlAssist.GetTagValue(xmlRequestElement, strTagName)
    
    'Cover Type
    strTagName = "COVERTYPE"
    strCoverType = m_objXmlAssist.GetTagValue(xmlRequestElement, strTagName)
    'Reference Number
    strTagName = "PPREFERENCENUMBER"
    strPPReferenceNumber = m_objXmlAssist.GetTagValue(xmlRequestElement, strTagName)
    'Monthly Cost
    strTagName = "TOTALPPMONTHLYCOST"
    strTotalPPMonthlyCost = m_objXmlAssist.GetTagValue(xmlRequestElement, strTagName)
    
    If Len(strApplicationNo) = 0 Or Len(strApplicationFFNo) = 0 Or Len(strPPSubQuoteNo) = 0 _
    Or Len(strCoverType) = 0 Or Len(strPPReferenceNumber) = 0 _
    Or Len(strTotalPPMonthlyCost) = 0 Then
       
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "APPLICATIONNUMBER = " & strApplicationNo & _
                                "; APPLICATIONFACTFINDNUMBER = " & strApplicationFFNo & _
                                "; PPSUBQUOTENUMBER = " & strPPSubQuoteNo & _
                                "; COVERTYPE = " & strCoverType & _
                                "; PPREFERENCENUMBER = " & strPPReferenceNumber & _
                                "; TOTALPPMONTHLYCOST = " & strTotalPPMonthlyCost

    End If

    Dim strData As String
    m_objXmlAssist.ChangeNodeName xmlRequestElement, _
                                cstrTABLE_NAME, _
                                "PAYMENTPROTECTIONSUBQUOTE"
    strData = objPPSubQuoteBO.SaveDetails(xmlIn.xml)
    m_objErrAssist.CheckXMLResponse strData, True, xmlResponseElem

    CalculateAndSavePPCosts = xmlOut.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
CalculateAndSavePPCostsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlResponseNode = Nothing
    Set xmlRequestElement = Nothing
    Set objPPSubQuoteBO = Nothing

    Exit Function

CalculateAndSavePPCostsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateAndSavePPCosts = m_objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalculateAndSavePPCostsExit
    
End Function


Public Function GetValidatedQuotationDetails(ByVal vstrXMLRequest As String) As String
    
    On Error GoTo GetValidatedQuotationDetailsVbErr

    Const strFunctionName As String = "GetValidatedQuotationDetails"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseNode As IXMLDOMNode
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objCostModelAssistBO As CostModelAssistBO
    Dim objICostModelAssistDO As ICostModelAssistDO
    Dim objQuotationBO As QuotationBO
    Dim objMSQBO As MortgageSubQuoteBO
    Dim blnValidMortgageSubQuote As Boolean
    Dim blnValidBCSubQuote As Boolean
    Dim blnValidPPSubQuote As Boolean
    
    blnValidMortgageSubQuote = False
    blnValidBCSubQuote = False
    blnValidPPSubQuote = False
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlResponseNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objICostModelAssistDO = New CostModelAssistDO
        Set objCostModelAssistBO = New CostModelAssistBO
        Set objQuotationBO = New QuotationBO
        Set objMSQBO = New MortgageSubQuoteBO
    Else
        Set objICostModelAssistDO = objContext.CreateInstance(App.Title & ".CostModelAssistDO")
        Set objCostModelAssistBO = objContext.CreateInstance(App.Title & ".CostModelAssistBO")
        Set objQuotationBO = objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
        Set objMSQBO = objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
    End If
    
    'Check the root node
    Set xmlRequestNode = xmlIn.getElementsByTagName("BASICQUOTATIONDETAILS").Item(0)
    
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "BASICQUOTATIONDETAILS" & " tag not found"
    End If

    Dim xmlRootNode As IXMLDOMNode
    Dim xmlReturnNode As IXMLDOMNode
    Dim lngError As Long
    
    Set xmlReturnNode = objICostModelAssistDO.GetDataForSubQuoteValidation(xmlRequestNode)
    'Data returned in xmlReturnNode check if any warnings
    
    Dim strActiveQuoteNo As String
    strActiveQuoteNo = objXmlAssist.GetTagValue(xmlRequestNode, "ACTIVEQUOTENUMBER")
    If Len(strActiveQuoteNo) <> 0 Then
        Dim xmlElement As IXMLDOMElement
        objXmlAssist.ChangeNodeName xmlRequestNode, "BASICQUOTATIONDETAILS", "QUICKQUOTEAPPLICANTDETAILS"
        'xmlResponseElem gets used in the next method
        Set xmlElement = xmlRequestNode.ownerDocument.createElement("ISSELFEMPLOYED")
    
        xmlElement.Text = IIf(objCostModelAssistBO.IsSelfEmployed(xmlRequestNode), 1, 0)
        'Integer value returned here true or false
        '
        Set xmlElement = xmlRequestNode.appendChild(xmlElement)
        
        'MV - 18/07/2002 - BMIDS00179 - Core Upgrade Rollback
        'SG 28/05/02 SYS4767 START
        'JLD MSMS0034 add in tag to indicate that we are in Quick Quote to GetValidatedQuotationDetails
        'Set xmlElement = xmlRequestNode.ownerDocument.createElement("ISQUICKQUOTE")
        'xmlElement.Text = "Y"
        'Set xmlElement = xmlRequestNode.appendChild(xmlElement)
        'SG 28/05/02 SYS4767 END
        
        objXmlAssist.ChangeNodeName xmlRequestNode, "QUICKQUOTEAPPLICANTDETAILS", "QUOTATION"
    
        Dim xmlReturn As FreeThreadedDOMDocument40
        Dim xmlTemp As IXMLDOMElement
        Dim strXMLResponse As String
        strXMLResponse = objQuotationBO.GetValidatedQuotationDetails(xmlRequestNode.ownerDocument.xml)
        Set xmlReturn = objXmlAssist.load(strXMLResponse, TypeName(Me), strFunctionName)
        lngError = objErrAssist.CheckResponse(xmlReturn.documentElement, xmlResponseNode, True)
        'Establish the validity of each of the subquotes
        Set xmlTemp = xmlReturn.selectSingleNode(".//VALIDMORTGAGESUBQUOTE")
        blnValidMortgageSubQuote = False
        If Not (xmlTemp Is Nothing) Then
            blnValidMortgageSubQuote = m_clsConvertAssist.CSafeBool(xmlTemp.Text)
        End If
        
        blnValidBCSubQuote = True
        Set xmlTemp = xmlReturn.selectSingleNode(".//VALIDBCSUBQUOTE")
        If Not (xmlTemp Is Nothing) Then
            blnValidBCSubQuote = m_clsConvertAssist.CSafeBool(xmlTemp.Text)
        End If
        
        blnValidPPSubQuote = True
        Set xmlTemp = xmlReturn.selectSingleNode(".//VALIDPPSUBQUOTE")
        If Not (xmlTemp Is Nothing) Then
            blnValidPPSubQuote = m_clsConvertAssist.CSafeBool(xmlTemp.Text)
        End If

        objXmlAssist.ChangeNodeName xmlRequestNode, "QUOTATION", "QUOTATIONDETAILS"
        Set xmlResponseNode = xmlResponseNode.appendChild(xmlRequestNode)
        objXmlAssist.AttachResponseData xmlResponseNode, xmlReturn.documentElement
    Else
     
        xmlResponseNode.appendChild xmlRequestNode
    
    End If


    'Check that all 3 are valid (MSQ, BC and PP)
    If blnValidMortgageSubQuote And blnValidBCSubQuote And blnValidPPSubQuote Then
    
        'GD BMIDS00256
        'Dim xmlCostModelLTV As IXMLDOMElement
        Dim xmlCostModelLTVRequest As IXMLDOMElement
        Dim xmlAppNumElem As IXMLDOMElement
        Dim xmlAppFFNumElem As IXMLDOMElement
        Dim xmlAmountRequestElem As IXMLDOMElement
        Dim xmlCostModelLTVResponse As IXMLDOMElement
        Dim xmlTempRequest As IXMLDOMElement
        Set xmlTempRequest = xmlIn.documentElement.cloneNode(False)
        
        ' Call IQuickQuoteBO_CalcCostModelLTV to Calculate the New LTV
        'GD BMIDS00431 12/09/02 - do nothing if AMOUNTREQUESTED is Zero or absent
        Set xmlAmountRequestElem = xmlOut.selectSingleNode(".//AMOUNTREQUESTED")
        If Not (xmlAmountRequestElem Is Nothing) Then
            If m_clsConvertAssist.CSafeDbl(xmlAmountRequestElem.Text) <> 0 Then
                Set xmlCostModelLTVRequest = xmlOut.createElement("LTV")
                Set xmlAppNumElem = xmlOut.selectSingleNode(".//APPLICATIONNUMBER")
                Set xmlAppFFNumElem = xmlOut.selectSingleNode(".//APPLICATIONFACTFINDNUMBER")
                xmlCostModelLTVRequest.appendChild xmlAppNumElem
                xmlCostModelLTVRequest.appendChild xmlAppFFNumElem
                xmlCostModelLTVRequest.appendChild xmlAmountRequestElem
                xmlTempRequest.appendChild xmlCostModelLTVRequest
                Set xmlCostModelLTVResponse = IQuickQuoteBO_CalcCostModelLTV(xmlTempRequest)
                'validate response
                objErrAssist.CheckResponse xmlCostModelLTVResponse, xmlOut.documentElement, True
                
                ' Call MortgageSubQuoteBO.Update to update the MORTGAGESUBQUOTE table with the new LTV
                ' - use string call across components
                Dim strMSQResponse As String
                Dim strMSQRequest As String
                
                strMSQRequest = "<REQUEST><MORTGAGESUBQUOTE>"
                strMSQRequest = strMSQRequest & xmlAppNumElem.xml
                strMSQRequest = strMSQRequest & xmlAppFFNumElem.xml
                strMSQRequest = strMSQRequest & xmlOut.selectSingleNode(".//MORTGAGESUBQUOTENUMBER").xml
                strMSQRequest = strMSQRequest & xmlCostModelLTVResponse.selectSingleNode(".//LTV").xml
                strMSQRequest = strMSQRequest & "</MORTGAGESUBQUOTE></REQUEST>"
                strMSQResponse = objMSQBO.Update(strMSQRequest)
                'Validate response
                objErrAssist.CheckXMLResponse strMSQResponse, True, xmlOut.documentElement
            'Else
            
                ' else ... MORTGAGESUBQUOTE table IS NOT updated with the new LTV
                'clean up objects used
                Set xmlAppNumElem = Nothing
                Set xmlAppFFNumElem = Nothing
                Set xmlCostModelLTVResponse = Nothing
                Set xmlAmountRequestElem = Nothing
                Set xmlCostModelLTVRequest = Nothing
                Set xmlTempRequest = Nothing
            End If
        End If
    End If
    
    GetValidatedQuotationDetails = xmlOut.xml

GetValidatedQuotationDetailsExit:

    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objICostModelAssistDO = Nothing
    Set objCostModelAssistBO = Nothing
    Set objQuotationBO = Nothing
    Set objMSQBO = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing

    Exit Function

GetValidatedQuotationDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
 
    Set objErrAssist = Nothing

    GetValidatedQuotationDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume GetValidatedQuotationDetailsExit

End Function

Public Function CalculateMaximumBorrowing(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instance of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'                               Takes a list of customers 1 to n and calcultes the 2 highest
'                               earners
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'                               returns the highest earners of the customers passed in as
'                               CustomerA and CustomerB
'------------------------------------------------------------------------------------------
On Error GoTo CalculateMaximumBorrowingVbErr
    
    Const strFunctionName As String = "CalculateMaximumBorrowing"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True

    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
    Dim objICostModelAssistBO As ICostModelAssistBO
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Dim xmlTemp As New FreeThreadedDOMDocument40
    xmlTemp.validateOnParse = False
    xmlTemp.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlTemp.createElement("MAXBORROWINGLIST")
    Set xmlDataNode = xmlTemp.appendChild(xmlResponseElem)

    If objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If
                
    Set xmlRequestNode = xmlIn.getElementsByTagName("CUSTOMERLIST").Item(0)
    
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "CUSTOMERLIST tag not found"
    End If
    Dim strHighestEarners As String
    
    Dim objReturnNode As IXMLDOMNode
    
    strHighestEarners = GetHighestEarners(xmlIn.xml)
    
    ' APS 03/03/00 - Re-raise error from GetHighestEarners method
    objErrAssist.CheckXMLResponse strHighestEarners, True
                                
    Set objReturnNode = objICostModelAssistBO.GetBorrowingAndLenderName(strHighestEarners)
                                                                                                                                                
    CalculateMaximumBorrowing = objReturnNode.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CalculateMaximumBorrowingExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    
    Exit Function

CalculateMaximumBorrowingVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateMaximumBorrowing = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume CalculateMaximumBorrowingExit

End Function

Private Function GetHighestEarners(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instance of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'                               Takes a list of customers 1 to n and calcultes the 2 highest
'                               earners
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'                               returns the highest earners of the customers passed in as
'                               CustomerA and CustomerB
'------------------------------------------------------------------------------------------
On Error GoTo GetHighestEarnersVbErr
    
    Const strFunctionName As String = "GetHighestEarners"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlData As New FreeThreadedDOMDocument40
    xmlData.validateOnParse = False
    xmlData.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objQuickQuoteApplicantDetailsDO As QuickQuoteApplicantDetailsDO
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objQuickQuoteApplicantDetailsDO = New QuickQuoteApplicantDetailsDO
    Else
        Set objQuickQuoteApplicantDetailsDO = objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsDO")
    End If
                
    Set xmlRequestNode = xmlIn.getElementsByTagName("CUSTOMERLIST").Item(0)
    
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "CUSTOMERLIST tag not found"
    End If
    Dim strXMLData As String
    
    
    'objXmlAssist.ChangeNodeName xmlIn.documentElement, "CUSTOMERLIST", "QUICKQUOTEAPPLICANTDETAILS"
    'Find all customers for this application.
    Dim xmlFindDoc As FreeThreadedDOMDocument40
    Set xmlFindDoc = New FreeThreadedDOMDocument40
    xmlFindDoc.validateOnParse = False
    xmlFindDoc.setProperty "NewParser", True
    Dim xmlReqNode As IXMLDOMNode
    Dim xmlFindNode As IXMLDOMNode
    Dim xmlTmpNode As IXMLDOMNode
    
    Set xmlReqNode = xmlFindDoc.appendChild(xmlFindDoc.createElement("REQUEST"))
    Set xmlFindNode = xmlFindDoc.createElement("QUICKQUOTEAPPLICANTDETAILS")
    Set xmlTmpNode = xmlFindDoc.createElement("APPLICATIONNUMBER")
    xmlTmpNode.Text = xmlIn.getElementsByTagName("APPLICATIONNUMBER").Item(0).Text
    xmlFindNode.appendChild xmlTmpNode
    Set xmlTmpNode = xmlFindDoc.createElement("APPLICATIONFACTFINDNUMBER")
    xmlTmpNode.Text = xmlIn.getElementsByTagName("APPLICATIONFACTFINDNUMBER").Item(0).Text
    xmlFindNode.appendChild xmlTmpNode
    xmlReqNode.appendChild xmlFindNode
    
    'strXMLData = objQuickQuoteApplicantDetailsDO.FindList(xmlIn.xml)
    strXMLData = objQuickQuoteApplicantDetailsDO.FindList(xmlFindDoc.xml)
    
    Set xmlData = objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
    
    Set xmlResponseNode = xmlDataNode.appendChild(xmlData.documentElement)
            
    Dim xmlCustomerNodelist As IXMLDOMNodeList
                    
    Set xmlCustomerNodelist = xmlOut.getElementsByTagName("QUICKQUOTEAPPLICANTDETAILS")
    
    Dim nloop As Integer
    
    Dim xmlElement As IXMLDOMElement
    
    Dim lngMaximumAmount As Long
    Dim lngLowerAmount As Long
    Dim intMaximumNode As Integer
    Dim intLowerNode As Integer
    
    lngMaximumAmount = 0
    lngLowerAmount = 0
    
    'get all the customers
    For nloop = 0 To (xmlCustomerNodelist.length - 1)
                
        Set xmlElement = xmlCustomerNodelist.Item(nloop)
                                    
        'do the calculation
        Dim lngAmount As Long
                
        lngAmount = objXmlAssist.GetTagValue(xmlElement, "ANNUALINCOMEORNETPROFIT")
        'New value greater than current max
        If lngAmount >= lngMaximumAmount Then
            'so set lower amount at previous max
            lngLowerAmount = lngMaximumAmount
            intLowerNode = intMaximumNode
            'then set max to current
            lngMaximumAmount = lngAmount
            intMaximumNode = nloop
            
        ElseIf lngAmount >= lngLowerAmount Then
            'new value greater than previous so set to this value
            lngLowerAmount = lngAmount
            intLowerNode = nloop
        End If
                
    Next
    
    If lngMaximumAmount > 0 Then
    
        Set xmlElement = xmlCustomerNodelist.Item(intMaximumNode)
        objXmlAssist.ChangeNodeName xmlElement, "QUICKQUOTEAPPLICANTDETAILS", "CUSTOMERA"
        
    End If

    If lngLowerAmount > 0 Then
    
        Set xmlElement = xmlCustomerNodelist.Item(intLowerNode)
        objXmlAssist.ChangeNodeName xmlElement, "QUICKQUOTEAPPLICANTDETAILS", "CUSTOMERB"
        
    End If
                
    objXmlAssist.ChangeNodeName xmlOut.documentElement, "ANNUALINCOMEORNETPROFIT", "TOTALGROSSEARNEDINCOME"
                
    GetHighestEarners = xmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetHighestEarnersExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlData = Nothing
    Set xmlOut = Nothing
    Set xmlFindDoc = Nothing
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objQuickQuoteApplicantDetailsDO = Nothing
    
    Exit Function

GetHighestEarnersVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetHighestEarners = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    '   go to clean-up section
    Resume GetHighestEarnersExit

End Function

Private Function IQuickQuoteBO_FindMortgageProducts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

    Dim objICostModelAssistBO As ICostModelAssistBO
    Dim objCostModelAssistDO As ICostModelAssistDO
    Dim xmlResponseRoot As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IQuickQuoteBO_FindMortgageProducts"
    Dim objApplicantDetailsDO As IQuickQuoteApplicantDetailsDO
    Dim xmlQuoteNodeResp As IXMLDOMNode
        
On Error GoTo FindMortgageProductsVbErr
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
        Set objApplicantDetailsDO = New QuickQuoteApplicantDetailsDO
        Set objCostModelAssistDO = New ICostModelAssistDO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(gstrQUICK_QUOTE & ".CostModelAssistBO")
        Set objApplicantDetailsDO = m_objContext.CreateInstance(gstrQUICK_QUOTE & ".QuickQuoteApplicantDetailsDO")
        Set objCostModelAssistDO = m_objContext.CreateInstance(gstrQUICK_QUOTE & ".CostModelAssistDO")
    End If
    
    ' DJP SYS1031
    ' Need to call QuickQuoteApplicantDetailsDO GetMainSelfEmployment to get whether or not either of the
    ' applicants are self employed.
    Dim bSelfEmployed As Boolean
    Dim xmlSelfEmployed As IXMLDOMNode
    Dim xmlDOMDoc As FreeThreadedDOMDocument40
    
    Set xmlDOMDoc = vxmlRequest.ownerDocument
    bSelfEmployed = objApplicantDetailsDO.GetMainSelfEmployment(vxmlRequest)

    Set xmlSelfEmployed = vxmlRequest.appendChild(xmlDOMDoc.createElement("APPLICANTSELFEMPLOYED"))
    xmlSelfEmployed.Text = CStr(bSelfEmployed)
    
    Set IQuickQuoteBO_FindMortgageProducts = objICostModelAssistBO.FindMortgageProducts(vxmlRequest)
    vxmlRequest.removeChild xmlSelfEmployed
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindMortgageProductsExit:

    '   clean-up section, release object references

    Exit Function

FindMortgageProductsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_FindMortgageProducts = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' Append
    ' go to clean-up section
    Resume FindMortgageProductsExit

End Function

Public Function FindMortgageProducts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Establishes the default values in readiness for a new loan component to be
'               setup.
'
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>
'   <MORTGAGEPRODUCTREQUEST>
'        <SEARCHCONTEXT>TEST01</SEARCHCONTEXT>
'        <CHANNELID>1</CHANNELID>
'        <MORTGAGESUBQUOTENUMBER>1</MORTGAGESUBQUOTENUMBER>
'        <APPLICATIONNUMBER></APPLICATIONNUMBER>
'        <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'        <MORTGAGESUBQUOTENUMBER></MORTGAGESUBQUOTENUMBER>
'        <PURPOSEOFLOAN></PURPOSEOFLOAN>
'        <TERMINYEARS></TERMINYEARS>
'        <TERMINMONTHS></TERMINMONTHS>
'        <AMOUNTREQUESTED></AMOUNTREQUESTED>
'   </MORTGAGEPRODUCTREQUEST>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "FindMortgageProducts"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))

    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DomDocument based method
    
    Set xmlTempResponseNode = IQuickQuoteBO_FindMortgageProducts(xmlIn.documentElement)

    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    m_objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    FindMortgageProducts = xmlResponseElem.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindMortgageProducts = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CreateExit
End Function

Public Function GetDefaultsForNewLoanComponent(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Establishes the default values in readiness for a new loan component to be
'               setup.
'
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>"
'   <LOANCOMPONENTDETAILS>
'      <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>"
'        <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>"
'        <MORTGAGESUBQUOTENUMBER>1</MORTGAGESUBQUOTENUMBER>
'        <AMOUNTREQUESTED>100000</AMOUNTREQUESTED>"
'      </LOANCOMPONENTDETAILS>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "GetDefaultsForNewLoanComponent"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuickQuoteBO_GetDefaultsForNewLoanComponent(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    m_objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetDefaultsForNewLoanComponent = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetDefaultsForNewLoanComponent = m_objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IQuickQuoteBO_GetDefaultsForNewLoanComponent(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  Establishes the default values in readiness for a new loan component to be
'               setup.
'
' pass:         vxmlRequest     xml Request data stream
'
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
    Const strFunctionName As String = "IQuickQuoteBO_GetDefaultsForNewLoanComponent"
    Dim strNodeName As String
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseRoot As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlGetDefaultsResp As IXMLDOMNode
    Dim objICostModelAssistBO As ICostModelAssistBO

    On Error GoTo GetDefaultsForNewLoanComponentErr
    
    ' Create default repsonse block
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If
    
    ' Call the cost model assist BO
    Set xmlGetDefaultsResp = objICostModelAssistBO.GetDefaultsForNewLoanComponent(vxmlRequest)
    
    m_objErrAssist.CheckResponse xmlGetDefaultsResp, xmlResponseRoot, True
    
    m_objXmlAssist.AttachResponseData xmlResponseRoot, xmlGetDefaultsResp
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IQuickQuoteBO_GetDefaultsForNewLoanComponent = xmlResponseRoot

GetDefaultsForNewLoanComponentExit:

    '   clean-up section, release object references

    Exit Function

GetDefaultsForNewLoanComponentErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_GetDefaultsForNewLoanComponent = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' Append
    ' go to clean-up section
    Resume GetDefaultsForNewLoanComponentExit
    
End Function

Public Function GetLoanCompositionDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the loan composition details for a given mortgage subquote
'
' pass:         vstrXMLRequest  xml Request as follows:
'
'    <REQUEST>
'      <LOANCOMPOSITION>
'        <MORTGAGESUBQUOTENUMBER>2</MORTGAGESUBQUOTENUMBER>
'        <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>
'        <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'      </LOANCOMPOSITION>
'   </REQUEST>"

' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "GetLoanCompositionDetails"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuickQuoteBO_GetLoanCompositionDetails(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    m_objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetLoanCompositionDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetLoanCompositionDetails = m_objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IQuickQuoteBO_GetLoanCompositionDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
        
            
    On Error GoTo GetLoanCompositionDetailsErr
    
    Const strFunctionName As String = "IQuickQuoteBO_GetLoanCompositionDetails"
    
    Dim strAppNo As String
    Dim strFactFindNo As String
    Dim strMortgageSubQuote As String
    Dim strQQLocation As String
    Dim strQQValuationType As String
    Dim strLocation As String
    Dim sTypeOfValutation As String
    Dim strPurchasePrice As String
    Dim strApplicationType As String
    Dim strAppDate As String
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlAppDataElem As IXMLDOMNode
    Dim xmlAppResponseElem As IXMLDOMNode
    Dim xmlSubQuoteResp As IXMLDOMElement
    Dim xmlDOMSubQuoteReq As FreeThreadedDOMDocument40
    Dim xmlSubQuoteRoot As IXMLDOMElement
    Dim xmlSubQuoteElem As IXMLDOMElement
    Dim xmlSubQuoteItem As IXMLDOMElement
    Dim xmlFactFindNode As IXMLDOMNode
    Dim xmlDOMParameter As FreeThreadedDOMDocument40
    Dim xmlDomAppData As New FreeThreadedDOMDocument40
    xmlDomAppData.validateOnParse = False
    xmlDomAppData.setProperty "NewParser", True
    Dim xmlAppDataItem As IXMLDOMElement
    
    Dim objIApplicationDO As IApplicationDO
    Dim objIMortgageSubQuoteBO As IMortgageSubQuoteBO
    Dim objGlobalParameterBO As GlobalParameterBO
    Dim objComboDO As ComboDO
    
    Dim strNodeName As String
    strNodeName = "LOANCOMPOSITION"
    
    ' Create default repsonse block
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
        Set objIMortgageSubQuoteBO = New MortgageSubQuoteBO
        Set objGlobalParameterBO = New GlobalParameterBO
        Set objComboDO = New ComboDO
    Else
        Set objIApplicationDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
        Set objIMortgageSubQuoteBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".MortgageSubQuoteBO")
        Set objGlobalParameterBO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterBO")
        Set objComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    ' Check the quotation is not already recommended
    ' Need to get the application data by calling GetApplicationData
    ' The tag for this function is "QUOTATION", so need to rename it to "APPLICATION"
    
    ' Validate the node passed in
    m_objXmlAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    
    ' Find the parent node, just to make sure it's there
    Set xmlRequestElem = m_objXmlAssist.GetNode(vxmlRequest, "/REQUEST/" & strNodeName, True)
    
    strAppNo = m_objXmlAssist.GetTagValue(xmlRequestElem, "APPLICATIONNUMBER")
    strFactFindNo = m_objXmlAssist.GetTagValue(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuote = m_objXmlAssist.GetTagValue(xmlRequestElem, "MORTGAGESUBQUOTENUMBER")
    
    If Len(strAppNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONNUMBER is empty"
    End If
    
    If Len(strMortgageSubQuote) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "MORTGAGESUBQUOTENUMBER is empty"
    End If
    
    If Len(strFactFindNo) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "FACTFINDNUMBER is empty"
    End If
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Need to build the request to pass to GetApplicationData
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set xmlAppDataElem = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    
    m_objXmlAssist.CheckNode xmlAppDataElem, TypeName(Me), strFunctionName
    
    Set xmlDomAppData.documentElement = xmlAppDataElem
    Set xmlAppDataElem = xmlAppDataElem.appendChild(xmlDomAppData.createElement("APPLICATION"))
    
    m_objXmlAssist.CheckNode xmlAppDataElem, TypeName(Me), strFunctionName
    
    ' Add the application number
    Set xmlAppDataItem = xmlAppDataElem.appendChild(xmlDomAppData.createElement("APPLICATIONNUMBER"))
    m_objXmlAssist.CheckNode xmlAppDataItem, TypeName(Me), strFunctionName
    xmlAppDataItem.Text = strAppNo
    
    ' Add the application fact find number
    Set xmlAppDataItem = xmlAppDataElem.appendChild(xmlDomAppData.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlAppDataItem.Text = strFactFindNo
    
    ' Get the ApplicationData
    Set xmlAppResponseElem = objIApplicationDO.GetApplicationData(xmlDomAppData.documentElement)
    ' JD BMIDS763 get the applicationdate
    strAppDate = m_objXmlAssist.GetElementText(xmlAppResponseElem, ".//APPLICATIONDATE")
    
    Set xmlFactFindNode = m_objXmlAssist.GetNode(xmlAppResponseElem, "//APPLICATIONFACTFIND")
    strPurchasePrice = m_objXmlAssist.GetTagValue(xmlFactFindNode, "PURCHASEPRICEORESTIMATEDVALUE")
    strApplicationType = m_objXmlAssist.GetTagValue(xmlFactFindNode, "TYPEOFAPPLICATION")
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get Location and Valuation Type
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' QQ Location
    strQQLocation = objGlobalParameterBO.GetCurrentParameter("QQLocation")
    Set xmlDOMParameter = m_objXmlAssist.load(strQQLocation, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDOMParameter.documentElement, xmlResponseElem, True
    strQQLocation = m_objXmlAssist.GetNodeValue(xmlDOMParameter.documentElement, ".//STRING", True)
    
    strLocation = objComboDO.GetComboValueId("PropertyLocation", strQQLocation)
    Set xmlDOMParameter = m_objXmlAssist.load(strLocation, TypeName(Me), strFunctionName)
    strLocation = m_objXmlAssist.GetNodeValue(xmlDOMParameter.documentElement, ".//VALUEID", True)
    
    ' QQ Valuation Type
    strQQValuationType = objGlobalParameterBO.GetCurrentParameter("QQValuationType")
    Set xmlDOMParameter = m_objXmlAssist.load(strQQValuationType, TypeName(Me), strFunctionName)
    m_objErrAssist.CheckResponse xmlDOMParameter.documentElement, xmlResponseElem, True
    
    strQQValuationType = m_objXmlAssist.GetNodeValue(xmlDOMParameter.documentElement, ".//STRING", True)
    sTypeOfValutation = objComboDO.GetComboValueId("ValuationType", strQQValuationType)
    Set xmlDOMParameter = m_objXmlAssist.load(sTypeOfValutation, TypeName(Me), strFunctionName)
    sTypeOfValutation = m_objXmlAssist.GetNodeValue(xmlDOMParameter.documentElement, ".//VALUEID", True)
    
    If Len(sTypeOfValutation) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "TYPEOFVALUATION is empty"
    End If
    
    If Len(strLocation) = 0 Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "PROPERTYLOCATION is empty"
    End If
    
    ' Get Loan Composition Details
    Set xmlSubQuoteRoot = m_objXmlAssist.GetRequestNode(vxmlRequest.ownerDocument)
    Set xmlDOMSubQuoteReq = xmlSubQuoteRoot.ownerDocument
    
    Set xmlSubQuoteElem = xmlSubQuoteRoot.appendChild(xmlDOMSubQuoteReq.createElement("MORTGAGESUBQUOTE"))
    
    ' Now add the application number
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("APPLICATIONNUMBER"))
    xmlSubQuoteItem.Text = strAppNo
        
    '... and application fact find number
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlSubQuoteItem.Text = strFactFindNo
    
    'JD BMIDS763 and application date
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("APPLICATIONDATE"))
    xmlSubQuoteItem.Text = strAppDate
        
    ' Mortgage sub quote number
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("MORTGAGESUBQUOTENUMBER"))
    xmlSubQuoteItem.Text = strMortgageSubQuote
        
    ' Location
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("LOCATION"))
    xmlSubQuoteItem.Text = strLocation
    
    ' Type of Valuation
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("TYPEOFVALUATION"))
    xmlSubQuoteItem.Text = sTypeOfValutation
    
    ' Purchase Price
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("PURCHASEPRICE"))
    xmlSubQuoteItem.Text = strPurchasePrice
    
    ' Type of Application
    Set xmlSubQuoteItem = xmlSubQuoteElem.appendChild(xmlDOMSubQuoteReq.createElement("TYPEOFAPPLICATION"))
    xmlSubQuoteItem.Text = strApplicationType
    
    Set xmlSubQuoteResp = objIMortgageSubQuoteBO.GetLoanCompositionDetails(xmlSubQuoteRoot)
    m_objErrAssist.CheckResponse xmlSubQuoteResp, xmlResponseElem, True
    
    ' Now build the XML to be returned from this method.
    ' Need to return Mortgage sub quote, loan component, mortgage product details, non panel mortgage product
    ' details, and incentives. These are returned in MortgageSubQuoteXML, LoanComponentXML, MortgageProductXML,
    ' NonPanelMortgageProductXML and IncentivesXML - all returned from GetLoanCompositionDetails
    
    'MV - 02/09/2002  - BMIDS00353
    Set xmlFactFindNode = m_objXmlAssist.GetNode(xmlAppResponseElem, "//APPLICATION")
    xmlSubQuoteResp.appendChild xmlFactFindNode.cloneNode(True)
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlSubQuoteResp
    
    Set IQuickQuoteBO_GetLoanCompositionDetails = xmlResponseElem
    
GetLoanCompositionDetailsExit:

    '   clean-up section, release object references

    Exit Function

GetLoanCompositionDetailsErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_GetLoanCompositionDetails = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' Append
    ' go to clean-up section
    Resume GetLoanCompositionDetailsExit
    
End Function

Public Function CalculateMortgageCosts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calculates and saves the mortgage costs
'
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' <REQUEST>
' <MORTGAGECOSTS>
'   <CONTEXT>Cost Modelling</CONTEXT>
'   <QUOTATIONNUMBER>1</QUOTATIONNUMBER>
'   <LIFESUBQUOTENUMBER>1</LIFESUBQUOTENUMBER>
'   <MORTGAGESUBQUOTENUMBER>2</MORTGAGESUBQUOTENUMBER>
'   <CUSTOMERLIST>
'       <CUSTOMER>
'           <CUSTOMERNUMBER>1333</CUSTOMERNUMBER>
'           <CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>
'       </CUSTOMER>
'       <CUSTOMER>
'           <CUSTOMERNUMBER>1325</CUSTOMERNUMBER>
'           <CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>
'       </CUSTOMER>
'   </CUSTOMERLIST>
'   <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>
'   "<APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
' </MORTGAGECOSTS>
' </REQUEST>

' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CalculateMortgageCosts"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuickQuoteBO_CalculateMortgageCosts(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    m_objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    CalculateMortgageCosts = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateMortgageCosts = m_objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IQuickQuoteBO_CalculateMortgageCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo CalculateMortgageCostsErr
    
    Const strFunctionName As String = "IQuickQuoteBO_CalculateMortgageCosts"
    
    Dim strNodeName As String
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlAppDataElem As IXMLDOMNode
    Dim xmlAppResponseElem As IXMLDOMNode
    
    Dim objIApplicationDO As IApplicationDO
    Dim objApplicationBO As ApplicationBO
    Dim objIQuotationBO As IQuotationBO
    Dim objIMortgageSubQuoteBO As IMortgageSubQuoteBO
    Dim objICostModelAssistBO As ICostModelAssistBO
    Dim objIQuickQuoteTxBO As IQuickQuoteTxBO
    
    strNodeName = "MORTGAGECOSTS"
    
    ' Validate the node passed in
    m_objXmlAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    
    ' Find the parent node, just to make sure it's there
    Set xmlRequestElem = m_objXmlAssist.GetNode(vxmlRequest, "/REQUEST/" & strNodeName, True)
    
    ' Create default repsonse block
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
        Set objIQuickQuoteTxBO = New IQuickQuoteTxBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
        Set objIQuickQuoteTxBO = m_objContext.CreateInstance(App.Title & ".QuickQuoteTxBO")
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get loan composition details - specific to Quick Quote
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xmlLoanCompReq As IXMLDOMNode
    Dim xmlLoanCompReqElem As IXMLDOMNode
    Dim xmlLoanCompItem As IXMLDOMNode
    Dim xmlLoanCompResp As IXMLDOMNode
    Dim xmlDOMDoc As FreeThreadedDOMDocument40
    Dim xmlAppNo As IXMLDOMNode
    Dim xmlAppDate As IXMLDOMNode
    Dim xmlAppFactFindNo As IXMLDOMNode
    Dim xmlMortgageSubQuoteNumber As IXMLDOMNode
    Dim xmlLifeSubQuoteNumber As IXMLDOMNode
    
    Set xmlLoanCompReq = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlDOMDoc = xmlLoanCompReq.ownerDocument
    
    Set xmlAppNo = m_objXmlAssist.GetNode(xmlRequestElem, "APPLICATIONNUMBER")
    Set xmlAppDate = m_objXmlAssist.GetNode(xmlRequestElem, "APPLICATIONDATE")
    Set xmlAppFactFindNo = m_objXmlAssist.GetNode(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    Set xmlMortgageSubQuoteNumber = m_objXmlAssist.GetNode(xmlRequestElem, "MORTGAGESUBQUOTENUMBER")
    Set xmlLifeSubQuoteNumber = m_objXmlAssist.GetNode(xmlRequestElem, "LIFESUBQUOTENUMBER")
    Set xmlLoanCompReqElem = xmlLoanCompReq.appendChild(xmlDOMDoc.createElement("LOANCOMPOSITION"))
    
    xmlLoanCompReqElem.appendChild xmlAppNo.cloneNode(True)
    xmlLoanCompReqElem.appendChild xmlAppDate.cloneNode(True) 'JD BMIDS763
    xmlLoanCompReqElem.appendChild xmlAppFactFindNo.cloneNode(True)
    xmlLoanCompReqElem.appendChild xmlMortgageSubQuoteNumber.cloneNode(True)
    
    Set xmlLoanCompResp = IQuickQuoteBO_GetLoanCompositionDetails(xmlLoanCompReq)
    m_objErrAssist.CheckResponse xmlLoanCompResp, xmlResponseElem, True

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Call the common method for CalculateMortgageCosts
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xmlMortgageCostsResp As IXMLDOMNode
    
    Set xmlMortgageCostsResp = objIQuickQuoteTxBO.CalculateMortgageCosts(vxmlRequest, _
                                                                            xmlLoanCompResp, _
                                                                            xmlAppNo, _
                                                                            xmlAppFactFindNo, _
                                                                            xmlLifeSubQuoteNumber, _
                                                                            xmlMortgageSubQuoteNumber)
                                                                            
    m_objErrAssist.CheckResponse xmlMortgageCostsResp, xmlResponseElem, True
        
    Set IQuickQuoteBO_CalculateMortgageCosts = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalculateMortgageCostsExit:

    '   clean-up section, release object references
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestElem = Nothing
    Set xmlAppDataElem = Nothing
    Set xmlAppResponseElem = Nothing
    
    Set objIApplicationDO = Nothing
    Set objApplicationBO = Nothing
    Set objIQuotationBO = Nothing
    Set objIMortgageSubQuoteBO = Nothing
    Set objICostModelAssistBO = Nothing
    Set objIQuickQuoteTxBO = Nothing
    
    Exit Function

CalculateMortgageCostsErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_CalculateMortgageCosts = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalculateMortgageCostsExit
End Function

Private Function IQuickQuoteBO_CalculateAndSaveLifeCoverCosts(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    Dim xmlResponseRoot As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Const strFunctionName As String = "IQuickQuoteBO_CalculateAndSaveLifeCoverCosts"
    Dim objICostModelAssistBO As ICostModelAssistBO

On Error GoTo CalculateAndSaveLifeCoverCostsVbErr
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(gstrQUICK_QUOTE & ".CostModelAssistBO")
    End If
    
    Set xmlRequestElem = m_objXmlAssist.GetNode(vxmlRequest, "LIFECOVERCOSTS")
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Build the Life Cover calc input interface
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xmlLifeCoverRespNode As IXMLDOMNode
    Dim xmlLifeCoverReqNode As IXMLDOMNode
    
    Set xmlLifeCoverRespNode = IQuickQuoteBO_BuildLifeCoverCalcInputInterface(vxmlRequest)
    m_objErrAssist.CheckResponse xmlLifeCoverRespNode, xmlResponseRoot, True

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Calculate the premiums, then save the life cover details
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xmlLifeCoverReq As IXMLDOMNode
    Dim xmlLifeCoverResp As IXMLDOMNode
    Dim xmlLifeCoverElem As IXMLDOMNode
    Dim xmlAppNo As IXMLDOMNode
    Dim xmlAppFactFindNo As IXMLDOMNode
    Dim xmlSubQuote As IXMLDOMNode
    
    Set xmlLifeCoverReq = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    
    Set xmlLifeCoverElem = m_objXmlAssist.GetNode(xmlLifeCoverRespNode, "LIFECOVER").cloneNode(True)
    Set xmlAppNo = m_objXmlAssist.GetNode(xmlRequestElem, "APPLICATIONNUMBER")
    Set xmlAppFactFindNo = m_objXmlAssist.GetNode(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    Set xmlSubQuote = m_objXmlAssist.GetNode(xmlRequestElem, "LIFESUBQUOTENUMBER")
    
    xmlLifeCoverElem.appendChild xmlAppNo.cloneNode(True)
    xmlLifeCoverElem.appendChild xmlAppFactFindNo.cloneNode(True)
    xmlLifeCoverElem.appendChild xmlSubQuote.cloneNode(True)

    
    ' Need to modify the customer XML
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim xmlCustomer As IXMLDOMNode
    Dim xmlItem As IXMLDOMNode
    Dim intCustomerNumber As Integer

    Set xmlCustomerList = xmlRequestElem.getElementsByTagName("CUSTOMER")
    intCustomerNumber = 1

    For Each xmlCustomer In xmlCustomerList
        Set xmlItem = m_objXmlAssist.GetNode(xmlCustomer, "CUSTOMERNUMBER")
        m_objXmlAssist.ChangeNodeName xmlItem, "CUSTOMERNUMBER", "CUSTOMERNUMBER" & intCustomerNumber

        Set xmlItem = m_objXmlAssist.GetNode(xmlCustomer, "CUSTOMERVERSIONNUMBER")
        m_objXmlAssist.ChangeNodeName xmlItem, "CUSTOMERVERSIONNUMBER", "CUSTOMERVERSIONNUMBER" & intCustomerNumber
        intCustomerNumber = intCustomerNumber + 1
    Next
    xmlLifeCoverElem.appendChild m_objXmlAssist.GetNode(xmlRequestElem, "CUSTOMERLIST")
    xmlLifeCoverReq.appendChild xmlLifeCoverElem.cloneNode(True)
    Set xmlLifeCoverResp = objICostModelAssistBO.CalcAndSaveLifeCover(xmlLifeCoverReq)
    m_objErrAssist.CheckResponse xmlLifeCoverResp, xmlResponseRoot, True
    Set IQuickQuoteBO_CalculateAndSaveLifeCoverCosts = xmlResponseRoot

CalculateAndSaveLifeCoverCostsExit:

    '   clean-up section, release object references

    Exit Function

CalculateAndSaveLifeCoverCostsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_CalculateAndSaveLifeCoverCosts = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' Append
    ' go to clean-up section
    Resume CalculateAndSaveLifeCoverCostsExit

End Function

Public Function CalculateAndSaveLifeCoverCosts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'
' pass:         vstrXMLRequest  containing values of
'               ApplicationNumber, ApplicationFactFindNumber, LifeSubQuoteNumber, Customer1Number,
'               Customer2Number, Customer1VersionNumber, Customer2VersionNumber,
'               LoanDetailsXML,
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CalculateAndSaveLifeCoverCostsVbErr

    Const strFunctionName As String = "CalculateAndSaveLifeCoverCosts"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))

    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DomDocument based method
    
    Set xmlTempResponseNode = IQuickQuoteBO_CalculateAndSaveLifeCoverCosts(xmlIn.documentElement)

    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    m_objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    CalculateAndSaveLifeCoverCosts = xmlResponseElem.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalculateAndSaveLifeCoverCostsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CalculateAndSaveLifeCoverCostsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateAndSaveLifeCoverCosts = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CalculateAndSaveLifeCoverCostsExit
End Function

Public Function BuildLifeCoverCalcInputInterface(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calculate the Term and Endowment Needs and Terms excluding non-ported loan
'               components
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>"
'   <LIFECOVER>
'       <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>"
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>"
'       <LOANCOMPONENTLIST>
'           <LOANCOMPONENT>
'           </LOANCOMPONENT>
'       </LOANCOMPONENTLIST>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER></CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER></CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'       </CUSTOMERLIST>
'   </LIFECOVER>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "BuildLifeCoverCalcInputInterface"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuickQuoteBO_BuildLifeCoverCalcInputInterface(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    BuildLifeCoverCalcInputInterface = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    BuildLifeCoverCalcInputInterface = m_objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IQuickQuoteBO_BuildLifeCoverCalcInputInterface(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' description:  Calculate the Term and Endowment Needs and Terms excluding non-ported loan
'               components
' pass:         vxmlRequest  xml Request data stream
' <REQUEST>"
'   <LIFECOVERCOSTS>
'       <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>"
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>"
'       <LOANCOMPONENTLIST>
'           <LOANCOMPONENT>
'           </LOANCOMPONENT>
'       </LOANCOMPONENTLIST>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER></CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER></CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'       </CUSTOMERLIST>
'   </LIFECOVERCOSTS>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
    Const strFunctionName As String = "IQuickQuoteBO_BuildLifeCoverCalcInputInterface"
    Dim strNodeName As String
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDOMDoc As FreeThreadedDOMDocument40
    Dim xmlResponseRoot As IXMLDOMElement
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlResponseItem As IXMLDOMElement
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlGetDefaultsResp As IXMLDOMNode
    Dim xmlAppNo As IXMLDOMNode
    Dim xmlAppFactFindNo As IXMLDOMNode
    Dim objICostModelAssistBO As ICostModelAssistBO
    Dim objQQApplicantDetailsDO As QuickQuoteApplicantDetailsDO
    Dim objIApplicationDO As IApplicationDO
    Dim objIComboDO As ComboDO
    Dim objICustomerDO As ICustomerDO
    
    On Error GoTo BuildLifeCoverCalcInputInterfaceErr
        
    strNodeName = "LIFECOVERCOSTS"
    
    ' Create default repsonse block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"
    Set xmlResponseElem = xmlResponseRoot.appendChild(xmlOut.createElement("LIFECOVER"))
    Set xmlRequestElem = m_objXmlAssist.GetNode(vxmlRequest, "LIFECOVERCOSTS")
    
    Set xmlAppNo = m_objXmlAssist.GetNode(xmlRequestElem, "APPLICATIONNUMBER")
    Set xmlAppFactFindNo = m_objXmlAssist.GetNode(xmlRequestElem, "APPLICATIONFACTFINDNUMBER")
    
    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
        Set objQQApplicantDetailsDO = New QuickQuoteApplicantDetailsDO
        Set objIApplicationDO = New ApplicationDO
        Set objIComboDO = New ComboDO
        Set objICustomerDO = New CustomerDO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
        Set objQQApplicantDetailsDO = m_objContext.CreateInstance(App.Title & ".QuickQuoteApplicantDetailsDO")
        Set objIApplicationDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
        Set objIComboDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
        Set objICustomerDO = m_objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerDO")
    End If
    
    ' Call the cost model assist BO
    Set xmlGetDefaultsResp = objICostModelAssistBO.BuildLifeCoverCalcInputInterface(vxmlRequest)
    m_objErrAssist.CheckResponse xmlGetDefaultsResp, xmlResponseRoot, True
    
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlGetDefaultsResp
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get Customer Details - Build up a list
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim xmlCustomerResp As IXMLDOMNode
    Dim xmlCustomerReq As IXMLDOMElement
    Dim xmlCustomerNode As IXMLDOMElement
    Dim xmlCustomerList As IXMLDOMNodeList
    Dim xmlCustomerRespList As IXMLDOMElement
    Dim xmlCustomerRespNode As IXMLDOMElement
    Dim xmlCustomerDoc As FreeThreadedDOMDocument40
    Dim xmlDOMApplicantDetails As FreeThreadedDOMDocument40
    Dim strApplicantDetailsResp As String
    Dim xmlQQApplicantElem As IXMLDOMElement
    Dim xmlCustomerItem As IXMLDOMNode
    Dim strCustomerItem As String
    
    Set xmlCustomerDoc = New FreeThreadedDOMDocument40
    xmlCustomerDoc.validateOnParse = False
    xmlCustomerDoc.setProperty "NewParser", True
    Set xmlCustomerList = xmlRequestElem.getElementsByTagName("CUSTOMER")
    
    Set xmlCustomerRespList = xmlCustomerDoc.createElement("CUSTOMERLIST")
    strNodeName = "QUICKQUOTEAPPLICANTDETAILS"
    
    Dim intCustomerNumber As Integer
    Dim strGender As String
    For Each xmlCustomerNode In xmlCustomerList
        ' Setup the output XML
        Set xmlCustomerRespNode = xmlCustomerRespList.appendChild(xmlCustomerDoc.createElement("CUSTOMER"))
        
        Set xmlCustomerDoc.documentElement = xmlCustomerDoc.createElement(strNodeName)
        Set xmlQQApplicantElem = xmlCustomerDoc.documentElement
        
        xmlQQApplicantElem.appendChild xmlCustomerNode.cloneNode(True)
        xmlQQApplicantElem.appendChild xmlAppNo.cloneNode(True)
        xmlQQApplicantElem.appendChild xmlAppFactFindNo.cloneNode(True)
        
        ' First, get the smoker and health status
        strApplicantDetailsResp = objQQApplicantDetailsDO.GetData(xmlQQApplicantElem.xml)
        Set xmlCustomerDoc = m_objXmlAssist.load(strApplicantDetailsResp, TypeName(Me), strFunctionName)
        
        ' Health
        strCustomerItem = m_objXmlAssist.GetNodeValue(xmlCustomerDoc.documentElement, "GOODHEALTH")
        Set xmlCustomerItem = xmlCustomerRespNode.appendChild(xmlCustomerDoc.createElement("GOODHEALTH"))
        xmlCustomerItem.Text = strCustomerItem
        
        ' Smoker status
        strCustomerItem = m_objXmlAssist.GetNodeValue(xmlCustomerDoc.documentElement, "SMOKER")
        Set xmlCustomerItem = xmlCustomerRespNode.appendChild(xmlCustomerDoc.createElement("SMOKERSTATUS"))
        xmlCustomerItem.Text = strCustomerItem
        
        ' Now get age and gender
        Set xmlCustomerResp = objICustomerDO.GetPersonalDetails(xmlCustomerNode)
        
        Set xmlCustomerItem = m_objXmlAssist.GetNode(xmlCustomerResp, ".//AGE")
        xmlCustomerRespNode.appendChild xmlCustomerItem.cloneNode(True)

        Set xmlCustomerItem = m_objXmlAssist.GetNode(xmlCustomerResp, ".//GENDER")
        If Not xmlCustomerItem Is Nothing Then
            
            ' APS 06/06/00 SYS0786 - Life cover quote now not passing in "Joint"
            ' regardless of the number of applcants
            If (intCustomerNumber = 0) Then
                strGender = xmlCustomerItem.Text
            End If
        
            xmlCustomerRespNode.appendChild xmlCustomerItem.cloneNode(True)
        End If
        
        xmlCustomerRespList.appendChild xmlCustomerRespNode
        intCustomerNumber = intCustomerNumber + 1
    Next
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Determine the relationship of customer 1 and customer 2, if more than
    ' one customer
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim strLifeCoverRelationship As String
    
    ' APS 06/06/00 SYS0786 - Life cover quote now not passing in "Joint"
    ' regardless of the number of applcants
    If xmlCustomerList.length > 1 Then
        strLifeCoverRelationship = "Joint"
    Else
        strLifeCoverRelationship = strGender
    End If
    
    Set xmlResponseItem = xmlResponseElem.appendChild(xmlOut.createElement("RELATIONSHIP"))
    xmlResponseItem.Text = strLifeCoverRelationship
    
    ' Call CostModelAssist to build the xml from the customer version XML
    Dim xmlLifeCoverElem As IXMLDOMElement
    Set xmlLifeCoverElem = objICostModelAssistBO.BuildLifeCoverInterfaceXML(xmlCustomerRespList)
    m_objErrAssist.CheckResponse xmlLifeCoverElem, xmlResponseRoot, True

    ' All done
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlLifeCoverElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
        
    
    Set IQuickQuoteBO_BuildLifeCoverCalcInputInterface = xmlResponseRoot

BuildLifeCoverCalcInputInterfaceExit:

    '   clean-up section, release object references

    Exit Function

BuildLifeCoverCalcInputInterfaceErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_BuildLifeCoverCalcInputInterface = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' Append
    ' go to clean-up section
    Resume BuildLifeCoverCalcInputInterfaceExit
    
End Function

Public Function CalcCostModelLTV(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Establishes the default values in readiness for a new loan component to be
'               setup.
'
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>
'   <LTV>
'       <APPLICATIONNUMBER>TEST01</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'       <AMOUNTREQUESTED>70000</AMOUNTREQUESTED>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER>1234</CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER>1234</CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'       </CUSTOMERLIST>
'   </LTV>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "CalcCostModelLTV"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))

    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Delegate to DomDocument based method
    
    Set xmlTempResponseNode = IQuickQuoteBO_CalcCostModelLTV(xmlIn.documentElement)

    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    m_objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    CalcCostModelLTV = xmlResponseElem.xml

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalcCostModelLTV = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Resume CreateExit
End Function

Private Function IQuickQuoteBO_CalcCostModelLTV(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
    Dim objICostModelAssistBO As ICostModelAssistBO
    Dim xmlResponseRoot As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlCalcLTVNode As IXMLDOMNode
    
Const strFunctionName As String = "IQuickQuoteBO_CalcCostModelLTV"

On Error GoTo CalcCostModelLTVVbErr
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseRoot = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    xmlResponseRoot.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Delegate to the assist object
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set xmlCalcLTVNode = objICostModelAssistBO.CalcCostModelLTV(vxmlRequest)
    m_objErrAssist.CheckResponse xmlCalcLTVNode, xmlResponseRoot, True
    m_objXmlAssist.AttachResponseData xmlResponseRoot, xmlCalcLTVNode
    
    Set IQuickQuoteBO_CalcCostModelLTV = xmlResponseRoot
    ' All done
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalcCostModelLTVExit:

    '   clean-up section, release object references

    Exit Function

CalcCostModelLTVVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseRoot
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_CalcCostModelLTV = m_objErrAssist.CreateErrorResponseEx

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' Append
    ' go to clean-up section
    Resume CalcCostModelLTVExit

End Function

Private Function IQuickQuoteBO_FindStoredQuoteDetails( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindStoredQuoteDetailsVbErr
    
    Const strFunctionName As String = "IQuickQuoteBO_FindStoredQuoteDetails"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
        
    Dim objICostModelAssistBO As ICostModelAssistBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objICostModelAssistBO = New CostModelAssistBO
    Else
        Set objICostModelAssistBO = m_objContext.CreateInstance(App.Title & ".CostModelAssistBO")
    End If
                
    Set xmlTempResponseNode = objICostModelAssistBO.FindStoredQuoteDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
            
    Set IQuickQuoteBO_FindStoredQuoteDetails = xmlTempResponseNode
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindStoredQuoteDetailsExit:
    
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindStoredQuoteDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_FindStoredQuoteDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume FindStoredQuoteDetailsExit

End Function

Public Function FindStoredQuoteDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindStoredQuoteDetailsVbErr

    Const strFunctionName As String = "FindStoredQuoteDetails"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IQuickQuoteBO_FindStoredQuoteDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindStoredQuoteDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindStoredQuoteDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindStoredQuoteDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindStoredQuoteDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume FindStoredQuoteDetailsExit

End Function

Private Function IQuickQuoteBO_RecostMortgageComponentAndLifeCover(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
On Error GoTo RecostMortgageComponentAndLifeCoverVbErr
    
    Const strFunctionName As String = "IQuickQuoteBO_RecostMortgageComponentAndLifeCover"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIQuickQuoteTxBO As IQuickQuoteTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objIQuickQuoteTxBO = New QuickQuoteTxBO
    Else
        Set objIQuickQuoteTxBO = m_objContext.CreateInstance(App.Title & ".QuickQuoteTxBO")
    End If
    
    Set xmlTempResponseNode = objIQuickQuoteTxBO.RecostMortgageComponentAndLifeCover(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IQuickQuoteBO_RecostMortgageComponentAndLifeCover = xmlTempResponseNode

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
RecostMortgageComponentAndLifeCoverExit:

    Set objIQuickQuoteTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

RecostMortgageComponentAndLifeCoverVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IQuickQuoteBO_RecostMortgageComponentAndLifeCover = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RecostMortgageComponentAndLifeCoverExit
End Function

Public Function RecostMortgageComponentAndLifeCover(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  RecostMortgageComponentAndLifeCover a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo RecostMortgageComponentAndLifeCoverVbErr

    Const strFunctionName As String = "RecostMortgageComponentAndLifeCover"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuickQuoteBO_RecostMortgageComponentAndLifeCover(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    RecostMortgageComponentAndLifeCover = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

RecostMortgageComponentAndLifeCoverExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

RecostMortgageComponentAndLifeCoverVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    RecostMortgageComponentAndLifeCover = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume RecostMortgageComponentAndLifeCoverExit
    
End Function

Public Function CreateNewQuotation(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Creates a new quotation record as active quotation
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Request format
'<REQUEST....>
'   <QUICKQUOTE>
'       <APPLICATIONNUMBER>..</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>..</APPLICATIONFACTFINDNUMBER>
'       <ACTIVEQUOTENUMBER>..</ACTIVEQUOTENUMBER>
'   </QUICKQUOTE>
'</REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewQuotationVbErr

    Const strFunctionName As String = "CreateNewQuotation"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuickQuoteBO_CreateNewQuotation(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateNewQuotation = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateNewQuotationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateNewQuotationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateNewQuotation = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewQuotationExit
    
End Function

Public Function ReinstateQuotation(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description: This reinstates a stored quotation but, after checking the quotation is still
'              valid
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Request format
'<REQUEST....>
'   <QUICKQUOTE>
'       <APPLICATIONNUMBER>..</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>..</APPLICATIONFACTFINDNUMBER>
'       <QUOTATIONNUMBER>..</QUOTATIONNUMBER>
'   </QUICKQUOTE>
'</REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo ReinstateQuotationVbErr

    Const strFunctionName As String = "ReinstateQuotation"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuickQuoteBO_ReinstateQuotation(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ReinstateQuotation = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ReinstateQuotationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ReinstateQuotationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ReinstateQuotation = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ReinstateQuotationExit
    
End Function

Public Function ResetMortgageSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description: This method resets an un-calculated mortgageSubQuote.
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Request format
'<REQUEST....>
'   <QUICKQUOTE>
'       <APPLICATIONNUMBER>..</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>..</APPLICATIONFACTFINDNUMBER>
'       <MORTGAGESUBQUOTENUMBER>..</MORTGAGESUBQUOTENUMBER>
'       <AMOUNTREQUESTED>..</AMOUNTREQUESTED>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER>..</CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER>..</CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'       <CUSTOMERLIST>
'   </QUICKQUOTE>
'</REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo ResetMortgageSubQuoteVbErr

    Const strFunctionName As String = "ResetMortgageSubQuote"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuickQuoteBO_ResetMortgageSubQuote(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ResetMortgageSubQuote = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ResetMortgageSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ResetMortgageSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ResetMortgageSubQuote = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ResetMortgageSubQuoteExit
    
End Function

Public Function UpdateMortgageSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Updates Mortgage Subquote.
'
' pass:         vstrXMLRequest  xml Request data stream
' <REQUEST>
'   <MORTGAGESUBQUOTE>
'       <APPLICATIONNUMBER>B00003239</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'   </MORTGAGESUBQUOTE>
' </REQUEST>
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageSubQuoteVbErr

    Const strFunctionName As String = "UpdateMortgageSubQuote"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.appendChild(xmlOut.createElement("RESPONSE"))
    
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuickQuoteBO_UpdateMortgageSubQuote(xmlIn.documentElement)
    
    ' Check response
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    ' Add the response to our repsonse, removing "RESPONSE"
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    UpdateMortgageSubQuote = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateMortgageSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateMortgageSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateMortgageSubQuote = m_objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateMortgageSubQuoteExit
End Function

Public Function CalculateAndSaveBCSubQuote(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Request format
'<REQUEST....>
'   <APPLICATIONQUOTE>
'       <APPLICATIONNUMBER>..</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>..</APPLICATIONFACTFINDNUMBER>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER>..</CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER>..</CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'       <CUSTOMERLIST>
'       <BUILDINGANDCONTENTS>
'           .................
'       </BUILDINGANDCONTENTS>
'   </APPLICATIONQUOTE>
'</REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo CalculateAndSaveBCSubQuoteVbErr

    Const strFunctionName As String = "CalculateAndSaveBCSubQuote"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuickQuoteBO_CalculateAndSaveBCSubQuote(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CalculateAndSaveBCSubQuote = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalculateAndSaveBCSubQuoteExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CalculateAndSaveBCSubQuoteVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalculateAndSaveBCSubQuote = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CalculateAndSaveBCSubQuoteExit

End Function

Public Function ProcessAddedOneOffCosts(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
' Request format
'<REQUEST....>
'   <QUICKQUOTE>
'       <APPLICATIONNUMBER>..</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>..</APPLICATIONFACTFINDNUMBER>
'       <CUSTOMERLIST>
'           <CUSTOMER>
'               <CUSTOMERNUMBER>..</CUSTOMERNUMBER>
'               <CUSTOMERVERSIONNUMBER>..</CUSTOMERVERSIONNUMBER>
'           </CUSTOMER>
'       <CUSTOMERLIST>
'       <ONEOFFCOST>
'           <COMBOVALIDATIONTYPE>MIG</COMBOVALIDATIONTYPE>
'           <AMOUNT>50</AMOUNT>
'           <ADDEDTOLOAN>1</ADDEDTOLOAN>
'       </ONEOFFCOST>
'   </QUICKQUOTE>
'</REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo ProcessAddedOneOffCostsVbErr

    Const strFunctionName As String = "ProcessAddedOneOffCosts"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IQuickQuoteBO_ProcessAddedOneOffCosts(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ProcessAddedOneOffCosts = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ProcessAddedOneOffCostsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ProcessAddedOneOffCostsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ProcessAddedOneOffCosts = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ProcessAddedOneOffCostsExit

End Function

Public Function ValidateManualAdjustmentPercent(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Checks the user and product parameters to validate the adjustment.
' pass:         vstrXmlRequest  xml Request data stream
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ValidateManualAdjustmentVbErr

    Const strFunctionName As String = "ValidateManualAdjustmentPercent"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
    Set xmlTempResponseNode = IQuickQuoteBO_ValidateManualAdjustmentPercent(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ValidateManualAdjustmentPercent = xmlTempResponseNode.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

ValidateManualAdjustmentExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

ValidateManualAdjustmentVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    ValidateManualAdjustmentPercent = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume ValidateManualAdjustmentExit
    
End Function

