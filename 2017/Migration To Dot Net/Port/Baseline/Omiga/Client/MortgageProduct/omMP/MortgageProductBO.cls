VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MortgageProductBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      MortgageProductBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Mortgage Product Business Object which 'supports transactions' only.
'               Methods which do not require transaction support reside in this
'               class. Any methods that require transactions will be delegated to
'               MortgageProductTxBO

'Dependencies:  MortgageProductDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    14/09/99    Created
'PSC    28/02/2000  Move getting of default number of records to DO
'PSC    11/04/00    AQR SYS0613: Check that LTV is present if context is Quick Quote or
'                                Cost Modelling
'MS     21/07/00    performance rework.
'APS    18/09/00    GetFullProductDetails VB error
'MV     27/11/00    Added Function FindProductNames
'DRC    04/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'MDC    10/12/01    SYS3440 - Rate Change amendments
'DRC    08/03/02    SYS4176 Added Function GetInterestRateForRateChange
'DB     28/05/02    SYS4767 MSMS to Core Integration
'------------------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'GD     10/09/02    BMIDS00313 - APWP2 - BM004 Added FindProductConditionsList
'DPF    23/10/02    CPWP1 - Amendment made to GetIncentivesForProduct
'MDC    01/11/2002  BMIDS00654 - Added GetMortgageProductIncomeMultiplierSet
'MDC    09/01/2003  BM0230 - Failed batch runs
'JD     21/06/04    BMIDS765 CC076 added GetMortgageProductRentalIncomeRateSet
'------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'PSC    18/06/2006  MAR1687 Add debug capability to GetFullProductDetails
'GHun   28/06/2006  MAR1878 improve performance of product searches
'------------------------------------------------------------------------------

 Option Explicit
Implements IMortgageProductBO
Implements ObjectControl
 
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Private m_objIDOAssist As IDOAssist
Private Const cstrTABLE_NAME = "MORTGAGEPRODUCT"

Private Function IMortgageProductBO_FindProductConditionsList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindProductConditionsListVbErr
    
    Const strFunctionName As String = "IMortgageProductBO_FindProductConditionsList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objIMortgageProductDO As IMortgageProductDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIMortgageProductDO = New MortgageProductDO
    Else
        Set objIMortgageProductDO = objContext.CreateInstance(App.Title & ".MortgageProductDO")
    End If

    Set xmlRequestNode = vxmlRequest.selectSingleNode("MORTGAGEPRODUCTLIST")

    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "MORTGAGEPRODUCTLIST" & " tag not found"
    End If
    
        
    Set xmlDataNode = objIMortgageProductDO.FindProductConditionsList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set IMortgageProductBO_FindProductConditionsList = xmlResponseElem.cloneNode(True)

FindProductConditionsListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageProductDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindProductConditionsListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IMortgageProductBO_FindProductConditionsList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindProductConditionsListExit
End Function

'BMIDS00654 MDC 01/11/2002
Private Function IMortgageProductBO_GetMortgageProductIncomeMultiplierSet(ByVal vxmlRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode

On Error GoTo IMortgageProductBO_GetMortgageProductIncomeMultiplierSetVbErr
    
Const strFunctionName As String = "IMortgageProductBO_GetMortgageProductIncomeMultiplierSet"

Dim objContext As ObjectContext
Set objContext = GetObjectContext()

Dim objErrAssist As ErrAssist
Set objErrAssist = New ErrAssist
Dim objXmlAssist As XMLAssist
Set objXmlAssist = New XMLAssist
Dim xmlOut As New FreeThreadedDOMDocument40
xmlOut.validateOnParse = False
xmlOut.setProperty "NewParser", True
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode

Dim objIMortgageProductDO As IMortgageProductDO

Dim strTempValue As String



    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIMortgageProductDO = New MortgageProductDO
    Else
        Set objIMortgageProductDO = objContext.CreateInstance(App.Title & ".MortgageProductDO")
    End If
        
    'Validate request
    strTempValue = objXmlAssist.GetMandatoryElementText(vxmlRequest, "MORTGAGEPRODUCT/MORTGAGEPRODUCTCODE")
    strTempValue = objXmlAssist.GetMandatoryElementText(vxmlRequest, "MORTGAGEPRODUCT/STARTDATE")
    
    Set xmlTempResponseNode = objIMortgageProductDO.GetMortgageProductIncomeMultiplierSet(vxmlRequest)
    xmlResponseElem.appendChild xmlTempResponseNode
    
    Set IMortgageProductBO_GetMortgageProductIncomeMultiplierSet = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

IMortgageProductBO_GetMortgageProductIncomeMultiplierSetExit:

    Set xmlResponseElem = Nothing

    Set xmlOut = Nothing

    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageProductDO = Nothing
    
    Exit Function

IMortgageProductBO_GetMortgageProductIncomeMultiplierSetVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.loadXML m_objErrAssist.FormatMessageNode
        'Set objXmlNode =
        xmlResponseElem.appendChild objXmlDoc.documentElement
        Set objXmlDoc = Nothing
        Resume Next
    End If

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IMortgageProductBO_GetMortgageProductIncomeMultiplierSet = objErrAssist.CreateErrorResponseEx
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume IMortgageProductBO_GetMortgageProductIncomeMultiplierSetExit

End Function
'BMIDS00654 MDC 01/11/2002 - End

Private Function IMortgageProductBO_GetMortgageProductLanguage(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Retrieves data from MortgageProductLanguage Table
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
 On Error GoTo GetMortgageProductLanguageVbErr

Const strFunctionName As String = "IMortgageProductBO_GetMortgageProductLanguage"

Dim objXmlAssist As New XMLAssist
Dim objErrAssist As New ErrAssist
Dim xmlRequestNode As IXMLDOMNode
Dim xmlDataNode As IXMLDOMNode
Dim xmlOut As New FreeThreadedDOMDocument40
xmlOut.validateOnParse = False
xmlOut.setProperty "NewParser", True
Dim xmlResponseElem As IXMLDOMElement

Dim objIMortgageProductDO As IMortgageProductDO

Const TABLE_NAME As String = "MORTGAGEPRODUCTLANGUAGE"

    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIMortgageProductDO = New MortgageProductDO
    Else
        Set objIMortgageProductDO = m_objContext.CreateInstance(App.Title & ".MortgageProductDO")
    End If
                
    If vxmlRequest.nodeName = TABLE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(TABLE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, TABLE_NAME & " tag not found"
    End If
    
    Set xmlDataNode = objIMortgageProductDO.GetMortgageProductLanguage(vxmlRequest)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IMortgageProductBO_GetMortgageProductLanguage = xmlResponseElem.cloneNode(True)

GetMortgageProductLanguageExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageProductDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetMortgageProductLanguageVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IMortgageProductBO_GetMortgageProductLanguage = objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetMortgageProductLanguageExit

End Function

Private Function IMortgageProductBO_GetMortgageProductRentalIncomeRateSet(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Retrieves rental income rate set data given product code and application data
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
 On Error GoTo GetMortgageProductRentalIncomeVbErr

Const strFunctionName As String = "IMortgageProductBO_GetMortgageProductRentalIncomeRateSet"

Dim objXmlAssist As New XMLAssist
Dim objErrAssist As New ErrAssist
Dim objSqlAssist As New SQLAssist
Dim xmlRequestNode As IXMLDOMNode
Dim xmlDataNode As IXMLDOMNode
Dim xmlRetNode As IXMLDOMNode
Dim xmlOut As New FreeThreadedDOMDocument40
xmlOut.validateOnParse = False
xmlOut.setProperty "NewParser", True
Dim xmlResponseElem As IXMLDOMElement

Dim objIMortgageProductDO As IMortgageProductDO

Const NODE_NAME As String = "LOANCOMPONENT"

    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIMortgageProductDO = New MortgageProductDO
    Else
        Set objIMortgageProductDO = m_objContext.CreateInstance(App.Title & ".MortgageProductDO")
    End If
                
    If vxmlRequest.nodeName = NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, NODE_NAME & " tag not found"
    End If
    
    'Get the mortgage product record from the product code and startdate passed in
    Dim xmlMortProdDoc As New FreeThreadedDOMDocument40
    xmlMortProdDoc.validateOnParse = False
    xmlMortProdDoc.setProperty "NewParser", True
    Dim xmlMortProdReqElem As IXMLDOMElement
    Dim xmlTempElem As IXMLDOMElement
    Dim xmlReqNode As IXMLDOMNode
    Dim xmlMortProdNode As IXMLDOMNode
    Dim xmlNewNode As IXMLDOMNode
    Dim strCode As String
    Dim strDate As String
    Dim strRateSet As String

    Set xmlNewNode = objXmlAssist.GetMandatoryNode(xmlRequestNode, "MORTGAGEPRODUCTCODE")
    strCode = xmlNewNode.Text
    Set xmlNewNode = objXmlAssist.GetMandatoryNode(xmlRequestNode, "MORTGAGEPRODUCTSTARTDATE")
    strDate = xmlNewNode.Text

    Set xmlMortProdReqElem = xmlMortProdDoc.createElement("REQUEST")
    Set xmlReqNode = xmlMortProdDoc.appendChild(xmlMortProdReqElem)
    Set xmlTempElem = xmlMortProdDoc.createElement("MORTGAGEPRODUCT")
    Set xmlMortProdNode = xmlReqNode.appendChild(xmlTempElem)
    Set xmlTempElem = xmlMortProdDoc.createElement("MORTGAGEPRODUCTCODE")
    Set xmlNewNode = xmlMortProdNode.appendChild(xmlTempElem)
    xmlNewNode.Text = strCode
    Set xmlTempElem = xmlMortProdDoc.createElement("STARTDATE")
    Set xmlNewNode = xmlMortProdNode.appendChild(xmlTempElem)
    xmlNewNode.Text = strDate

    Set xmlDataNode = objIMortgageProductDO.GetData(xmlMortProdReqElem)
    Set xmlNewNode = objXmlAssist.GetMandatoryNode(xmlDataNode, "RENTALINCOMERATESET")
    strRateSet = xmlNewNode.Text
    If strRateSet = "" Then
        Set IMortgageProductBO_GetMortgageProductRentalIncomeRateSet = xmlResponseElem.cloneNode(True)
        GoTo GetMortgageProductRentalIncomeExit
    End If

    'Set up new XML to pass to DO from Mortgage product record and input XML
    Dim xmlRateSetDoc As New FreeThreadedDOMDocument40
    xmlRateSetDoc.validateOnParse = False
    xmlRateSetDoc.setProperty "NewParser", True
    Dim xmlRateSetReqElem As IXMLDOMElement
    Dim xmlRateSetNode As IXMLDOMNode
    Dim xmlTmpNode As IXMLDOMNode

    Set xmlRateSetReqElem = xmlRateSetDoc.createElement("REQUEST")
    Set xmlReqNode = xmlRateSetDoc.appendChild(xmlRateSetReqElem)
    Set xmlTempElem = xmlRateSetDoc.createElement("RENTALINCOMERATESETBAND")
    Set xmlRateSetNode = xmlReqNode.appendChild(xmlTempElem)
    Set xmlTempElem = xmlRateSetDoc.createElement("RENTALINCOMERATESET")
    Set xmlNewNode = xmlRateSetNode.appendChild(xmlTempElem)
    xmlNewNode.Text = strRateSet
    Set xmlTempElem = xmlRateSetDoc.createElement("STARTDATE")
    Set xmlNewNode = xmlRateSetNode.appendChild(xmlTempElem)
    Set xmlTmpNode = objXmlAssist.GetMandatoryNode(xmlRequestNode, "APPLICATIONDATE")
    xmlNewNode.Text = xmlTmpNode.Text
    Set xmlTempElem = xmlRateSetDoc.createElement("LOANAMOUNT")
    Set xmlNewNode = xmlRateSetNode.appendChild(xmlTempElem)
    Set xmlTmpNode = objXmlAssist.GetMandatoryNode(xmlRequestNode, "LOANAMOUNT")
    xmlNewNode.Text = xmlTmpNode.Text
    Set xmlTempElem = xmlRateSetDoc.createElement("LTV")
    Set xmlNewNode = xmlRateSetNode.appendChild(xmlTempElem)
    Set xmlTmpNode = objXmlAssist.GetMandatoryNode(xmlRequestNode, "LTV")
    xmlNewNode.Text = xmlTmpNode.Text

    Set xmlRetNode = objIMortgageProductDO.GetMortgageProductRentalIncomeRateSet(xmlRateSetReqElem)
    If Not xmlRetNode Is Nothing Then
        xmlResponseElem.appendChild xmlRetNode
    End If
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IMortgageProductBO_GetMortgageProductRentalIncomeRateSet = xmlResponseElem.cloneNode(True)

GetMortgageProductRentalIncomeExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageProductDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRateSetDoc = Nothing
    Set xmlMortProdDoc = Nothing
    Set objSqlAssist = Nothing

    Exit Function

GetMortgageProductRentalIncomeVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IMortgageProductBO_GetMortgageProductRentalIncomeRateSet = objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetMortgageProductRentalIncomeExit

End Function

Private Function IMortgageProductBO_GetRateForBaseRateSet(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Retrieves BaseRate data for the given Base Rate Set
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
 On Error GoTo GetRateForBaseRateSetVbErr

Const strFunctionName As String = "IMortgageProductBO_GetRateForBaseRateSet"

Dim objXmlAssist As New XMLAssist
Dim objErrAssist As New ErrAssist
Dim xmlRequestNode As IXMLDOMNode
Dim xmlDataNode As IXMLDOMNode
Dim xmlOut As New FreeThreadedDOMDocument40
xmlOut.validateOnParse = False
xmlOut.setProperty "NewParser", True
Dim xmlResponseElem As IXMLDOMElement

Dim objIMortgageProductDO As IMortgageProductDO

Const TABLE_NAME As String = "BASERATESET"

    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIMortgageProductDO = New MortgageProductDO
    Else
        Set objIMortgageProductDO = m_objContext.CreateInstance(App.Title & ".MortgageProductDO")
    End If
                
    If vxmlRequest.nodeName = TABLE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(TABLE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, TABLE_NAME & " tag not found"
    End If
    
    'SYS3440 MDC 10/12/2001.
'    Set xmlDataNode = objIMortgageProductDO.GetRateForBaseRateSet(xmlRequestNode)
    Set xmlDataNode = objIMortgageProductDO.GetRateForBaseRateSet(vxmlRequest)
    xmlResponseElem.appendChild xmlDataNode
    'SYS3440 End
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IMortgageProductBO_GetRateForBaseRateSet = xmlResponseElem.cloneNode(True)

GetRateForBaseRateSetExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageProductDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetRateForBaseRateSetVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IMortgageProductBO_GetRateForBaseRateSet = objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetRateForBaseRateSetExit

End Function

Private Function IMortgageProductBO_GetInterestRateForRateChange(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Calculates the New & Old Rate for a specific Loan Component
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
 On Error GoTo GetInterestRateForRateChangeVbErr

Const strFunctionName As String = "IMortgageProductBO_GetInterestRateForRateChange"

Dim objXmlAssist As New XMLAssist
Dim objErrAssist As New ErrAssist
Dim objConvAssist As New ConvertAssist
Dim xmlRequestNode As IXMLDOMNode
Dim xmlBaseRateElem As IXMLDOMElement
Dim xmlTempElem As IXMLDOMElement
Dim xmlDataNode As IXMLDOMNode
Dim xmlNewNode As IXMLDOMNode
Dim xmlOut As New FreeThreadedDOMDocument40
xmlOut.validateOnParse = False
xmlOut.setProperty "NewParser", True
Dim xmlIn As New FreeThreadedDOMDocument40
xmlIn.validateOnParse = False
xmlIn.setProperty "NewParser", True
Dim xmlRateChangeElem As IXMLDOMElement
Dim xmlInterestRateElem As IXMLDOMElement
Dim xmlResponseElem As IXMLDOMElement
Dim dblOldInterestRate As Double
Dim dblNewInterestRate As Double
Dim dblRateDifference As Double
Dim dblDiscountRate As Double
Dim dblFlooredRate As Double
Dim dblCeilingRate As Double
Dim strNewStartDate As String
Dim lngErr As Long
Dim strRateType As String
Dim blnContinue As Boolean
Dim objIMortgageProductDO As IMortgageProductDO
'BM0230 MDC 08/01/2003
Dim xmlAtttrib As IXMLDOMAttribute
Dim blnBaseRateApplied As Boolean
'BM0230 MDC 08/01/2003 - End

Const TABLE_NAME As String = "BASERATESET"

    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIMortgageProductDO = New MortgageProductDO
    Else
        Set objIMortgageProductDO = m_objContext.CreateInstance(App.Title & ".MortgageProductDO")
    End If
                
    If vxmlRequest.nodeName = TABLE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(TABLE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, TABLE_NAME & " tag not found"
    End If
    
    '
    'Get the previous (i.e. current) base interest rate
    '
    'BM0230 MDC 08/01/2003
    Set xmlBaseRateElem = vxmlRequest.firstChild
    Set xmlAtttrib = vxmlRequest.Attributes.getNamedItem("OLDRATE")
    If Not xmlAtttrib Is Nothing Then
        'Old rate from previous failed batch run to be used as new base rate has been applied
        dblOldInterestRate = objConvAssist.CSafeDbl(xmlAtttrib.Text)
    Else
        'Get the previous rate from the base rate table
        Set xmlDataNode = objIMortgageProductDO.GetCurrentRateForBaseRateSet(xmlBaseRateElem)
        dblOldInterestRate = objConvAssist.CSafeDbl(objXmlAssist.GetElementText(xmlDataNode, "BASEINTERESTRATE"))
    End If
    'BM0230 MDC 08/01/2003 - End
    
    '
    'Get the new Base Interest Rate and the Start Date
    '
    'need to append a boolean for rate change
    Set xmlRateChangeElem = xmlIn.createElement("RATECHANGE")
    xmlRateChangeElem.Text = "1"
    xmlBaseRateElem.appendChild xmlRateChangeElem
    '
    'Call to Data object method
    On Error Resume Next
    Set xmlDataNode = objIMortgageProductDO.GetRateForBaseRateSet(xmlBaseRateElem)
    '
    'Check for record not found
    '
    lngErr = objErrAssist.GetOmigaErrorNumber(Err.Number)
    If lngErr = oeRecordNotFound Then
       blnContinue = False
    Else
       blnContinue = True
    End If
    On Error GoTo GetInterestRateForRateChangeVbErr
    If Not blnContinue Then
        Set xmlTempElem = xmlOut.createElement("NEWRATE")
            Set xmlNewNode = xmlResponseElem.appendChild(xmlTempElem)
            xmlNewNode.Text = "Error"
            Set xmlTempElem = xmlOut.createElement("OLDRATE")
            Set xmlNewNode = xmlResponseElem.appendChild(xmlTempElem)
            xmlNewNode.Text = "Error"
            Set xmlTempElem = xmlOut.createElement("NEWRATESTARTDATE")
            Set xmlNewNode = xmlResponseElem.appendChild(xmlTempElem)
            xmlNewNode.Text = "Error"
            Set IMortgageProductBO_GetInterestRateForRateChange = xmlResponseElem.cloneNode(True)
        
    Else
    On Error GoTo GetInterestRateForRateChangeVbErr
        dblNewInterestRate = objConvAssist.CSafeDbl(objXmlAssist.GetElementText(xmlDataNode, "BASEINTERESTRATE"))
        strNewStartDate = objXmlAssist.GetElementText(xmlDataNode, "BASERATESTARTDATE")

    
        '
        ' Get the latest Rate Difference from the BaseRateBand table for a specific BaseRateSet
        '
        Set xmlDataNode = objIMortgageProductDO.GetRateDifference(xmlBaseRateElem)
        dblRateDifference = objConvAssist.CSafeDbl(objXmlAssist.GetElementText(xmlDataNode, "RATEDIFFERENCE"))
        '
        ' Apply the Rate Difference to both the old and the new rate
        '
        'BM0230 MDC 08/01/2003 - Do not apply adjustments on failed batch runs
        If xmlAtttrib Is Nothing Then
            dblOldInterestRate = dblOldInterestRate + dblRateDifference
        End If
        'BM0230 MDC 08/01/2003 - End
        dblNewInterestRate = dblNewInterestRate + dblRateDifference
        '
        ' Get the interest rate type & sequence number
        '
    
    
        Set xmlNewNode = vxmlRequest.cloneNode(True)
        
        objXmlAssist.ChangeNodeName xmlNewNode, "LOANCOMPONENT", "INTERESTRATETYPE"
        Set xmlDataNode = objIMortgageProductDO.GetInterestRateType(xmlNewNode)
        
        ' TODO note the literal is being returned - we should really be doing a
        ' get combo here - but since Supervisor doesn't use combo values
        ' this is how we are doing this until supervisor is updated
        
        strRateType = objXmlAssist.GetElementText(xmlDataNode, "RATETYPE")
        
        ' Get the discount rate where rate type is discount or capped
        
       If ((strRateType = "D") Or (strRateType = "C")) Then
        dblDiscountRate = objConvAssist.CSafeDbl(objXmlAssist.GetElementText(xmlDataNode, "RATE"))
        'BM0230 MDC 08/01/2003 - Do not apply adjustments on failed batch runs
        If xmlAtttrib Is Nothing Then
            dblOldInterestRate = dblOldInterestRate - dblDiscountRate
        End If
        'BM0230 MDC 08/01/2003 - End
        dblNewInterestRate = dblNewInterestRate - dblDiscountRate
       End If
    ' apply capping
       If (strRateType = "C") Then
        dblFlooredRate = objConvAssist.CSafeDbl(objXmlAssist.GetElementText(xmlDataNode, "FLOOREDRATE"))
        dblCeilingRate = objConvAssist.CSafeDbl(objXmlAssist.GetElementText(xmlDataNode, "CEILINGRATE"))
    ' adjust old rate
        If dblOldInterestRate < dblFlooredRate Then
          dblOldInterestRate = dblFlooredRate
        Else
           If dblOldInterestRate > dblCeilingRate Then
             dblOldInterestRate = dblCeilingRate
           End If
        End If
    ' adjust new rate
        If dblNewInterestRate < dblFlooredRate Then
          dblNewInterestRate = dblFlooredRate
        Else
           If dblNewInterestRate > dblCeilingRate Then
             dblNewInterestRate = dblCeilingRate
           End If
        End If
       End If
     
        
        
        If Not m_objContext Is Nothing Then
            m_objContext.SetComplete
        End If
    '
    '    Now put results into response
    '
        Set xmlTempElem = xmlOut.createElement("NEWRATE")
        Set xmlNewNode = xmlResponseElem.appendChild(xmlTempElem)
        xmlNewNode.Text = dblNewInterestRate
        Set xmlTempElem = xmlOut.createElement("OLDRATE")
        Set xmlNewNode = xmlResponseElem.appendChild(xmlTempElem)
        xmlNewNode.Text = dblOldInterestRate
        Set xmlTempElem = xmlOut.createElement("NEWRATESTARTDATE")
        Set xmlNewNode = xmlResponseElem.appendChild(xmlTempElem)
        xmlNewNode.Text = strNewStartDate
        
        Set IMortgageProductBO_GetInterestRateForRateChange = xmlResponseElem.cloneNode(True)
    End If
GetInterestRateForRateChangeExit:

    
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objConvAssist = Nothing
    Set objIMortgageProductDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempElem = Nothing
    Set xmlNewNode = Nothing
    Set xmlAtttrib = Nothing    'BM0230 MDC 08/01/2003
    

    Exit Function

GetInterestRateForRateChangeVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IMortgageProductBO_GetInterestRateForRateChange = objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetInterestRateForRateChangeExit

End Function

Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = True
    
End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing

End Sub

Public Function GetRateForBaseRateSet(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Retrieves BaseRate for the BaseRateSet
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
 On Error GoTo GetRateForBaseRateSetVbErr

Const strFunctionName As String = "GetRateForBaseRateSet"

Dim xmlIn As FreeThreadedDOMDocument40
Dim objErrAssist As New ErrAssist
Dim objXmlAssist As New XMLAssist
Dim xmlOut As New FreeThreadedDOMDocument40
xmlOut.validateOnParse = False
xmlOut.setProperty "NewParser", True
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageProductBO_GetRateForBaseRateSet(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetRateForBaseRateSet = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetRateForBaseRateSetExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetRateForBaseRateSetVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetRateForBaseRateSet = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetRateForBaseRateSetExit

End Function
Public Function GetMortgageProductRentalIncomeRateSet(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Retrieves rental income rate set
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
 On Error GoTo GetMortgageProductRentalIncomeVbErr

Const strFunctionName As String = "GetMortgageProductRentalIncomeRateSet"

Dim xmlIn As FreeThreadedDOMDocument40
Dim objErrAssist As New ErrAssist
Dim objXmlAssist As New XMLAssist
Dim xmlOut As New FreeThreadedDOMDocument40
xmlOut.validateOnParse = False
xmlOut.setProperty "NewParser", True
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageProductBO_GetMortgageProductRentalIncomeRateSet(xmlIn.documentElement)
    
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    
    GetMortgageProductRentalIncomeRateSet = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetMortgageProductRentalIncomeExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetMortgageProductRentalIncomeVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetMortgageProductRentalIncomeRateSet = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetMortgageProductRentalIncomeExit

End Function

Public Function GetInterestRateForRateChange(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Calculates the New & Old Rate for a specific Loan Component
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
 'On Error GoTo GetInterestRateForRateChangeVbErr

Const strFunctionName As String = "GetInterestRateForRateChange"

Dim xmlIn As FreeThreadedDOMDocument40
Dim objErrAssist As New ErrAssist
Dim objXmlAssist As New XMLAssist
Dim xmlOut As New FreeThreadedDOMDocument40
xmlOut.validateOnParse = False
xmlOut.setProperty "NewParser", True
Dim xmlResponseElem As IXMLDOMElement
Dim xmlTempResponseNode As IXMLDOMNode

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageProductBO_GetInterestRateForRateChange(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetInterestRateForRateChange = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetInterestRateForRateChangeExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetInterestRateForRateChangeVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetInterestRateForRateChange = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetInterestRateForRateChangeExit

End Function

Public Function FindProductList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Retrieves all instances of mortgage products that match the required
'               search criteria
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
 On Error GoTo FindProductListVbErr

    Const strFunctionName As String = "FindProductList"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As ErrAssist
    Dim objXmlAssist As XMLAssist
    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    Set objErrAssist = New ErrAssist
    Set objXmlAssist = New XMLAssist
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageProductBO_FindProductList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindProductList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindProductListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindProductListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindProductList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindProductListExit

End Function

Private Sub ValidateFindProductList(ByVal vxmlRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Validates the input request for FindProductList
'
' pass:         vxmlRequest     xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo ValidateFindProductListVbErr
    
    Const strFunctionName As String = "ValidateFindProductList"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As ErrAssist
    Dim objXmlAssist As XMLAssist
    Dim objConvAssist As ConvertAssist
    Dim objGlobalParamDO As GlobalParameterDO
    
    Set objErrAssist = New ErrAssist
    Set objXmlAssist = New XMLAssist
    Set objConvAssist = New ConvertAssist
    
    ' Check valid amount requested
    Dim strTagValue As String
    Dim lngAmtRequested As Long
        
    strTagValue = objXmlAssist.GetTagValue(vxmlRequest, "AMOUNTREQUESTED")
    lngAmtRequested = objConvAssist.CSafeLng(strTagValue)
    
    'SG 21/03/02 SYS4202
    Dim strMortgageProductCode As String
    strMortgageProductCode = objXmlAssist.GetTagValue(vxmlRequest, "MORTGAGEPRODUCTCODE")
    'SG 21/03/02 SYS4202
    
    'DB SYS4767 - MSMS Integration
    If lngAmtRequested <= 0 Then
        'SG 21/03/02 SYS4202
        'If we have a product code then don't error because stored proc wont use Amount param
        If Len(strMortgageProductCode) = 0 Then     'MAR1878 GHun
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 131
        End If
        'SG 21/03/02 SYS4202
    End If
    'DB End
    
    ' Check filtering criteria
    Dim strProductsByGroup As String
    Dim strProductGroup As String
    Dim blnProductsByGroup As Boolean
    
    strTagValue = objXmlAssist.GetTagValue(vxmlRequest, "PRODUCTSBYGROUP")
    blnProductsByGroup = objConvAssist.CSafeBool(strTagValue)
    
    strProductGroup = objXmlAssist.GetTagValue(vxmlRequest, "PRODUCTGROUP")
        
    'DB SYS4767 - MSMS Integration
    If blnProductsByGroup = True And Len(strProductGroup) = 0 Then
    'SG 21/03/02 SYS4202
        'If we have a product code then don't error because stored proc wont use Product Group param
        If Len(strMortgageProductCode) = 0 Then     'MAR1878 GHun
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 132
        End If
    'SG 21/03/02 SYS4202
    End If
    'DB End
    
    ' Check term is in range
    Dim intMinimumTerm As Integer
    Dim intMaximumTerm As Integer
    Dim intTermMonths As Integer
    Dim intTermYears As Integer
    Dim intTerm As Integer
     
    strTagValue = objXmlAssist.GetTagValue(vxmlRequest, "TERMINMONTHS")
    intTermMonths = objConvAssist.CSafeInt(strTagValue)
    strTagValue = objXmlAssist.GetTagValue(vxmlRequest, "TERMINYEARS")
    intTermYears = objConvAssist.CSafeInt(strTagValue)

    intTerm = (intTermYears * 12) + intTermMonths
    
    If objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
    Else
        Set objGlobalParamDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    End If
    
    ' Get Minimum Term and convert to months
    intMinimumTerm = objGlobalParamDO.GetCurrentParameterByType("MinimumTerm", "AMOUNT")
    intMinimumTerm = intMinimumTerm * 12
    
    ' Get Maximum Term and convert to months
    intMaximumTerm = objGlobalParamDO.GetCurrentParameterByType("MaximumTerm", "AMOUNT")
    intMaximumTerm = intMaximumTerm * 12
        
    'DB SYS4767 - MSMS Integration
    If intTerm < intMinimumTerm Or intTerm > intMaximumTerm Then
    'SG 21/03/02 SYS4202
        'If we have a product code then don't error because stored proc wont use Product Group param
        If Len(strMortgageProductCode) = 0 Then     'MAR1878 GHun
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 130
        End If
    'SG 21/03/02 SYS4202
    End If
    'DB End
    
    ' PSC 11/04/00 AQR SYS0613 - Start
    ' Check LTV is present if Cost Modelling or Quick Quote
    Dim strSearchContext As String
    
    strSearchContext = objXmlAssist.GetTagValue(vxmlRequest, "SEARCHCONTEXT")
    
    If StrComp(strSearchContext, "Quick Quote", vbTextCompare) = 0 Or _
       StrComp(strSearchContext, "Cost Modelling", vbTextCompare) = 0 Then
    
        Dim dblLTV As Double
        
        strTagValue = objXmlAssist.GetTagValue(vxmlRequest, "LTV")
        dblLTV = objConvAssist.CSafeDbl(strTagValue)
        
        'DB SYS4767 - MSMS Integration
        If dblLTV <= 0 Then
        'SG 21/03/02 SYS4202
            'If we have a product code then don't error because stored proc wont use LTV param
            If Len(strMortgageProductCode) = 0 Then     'MAR1878 GHun
                objErrAssist.ThrowError TypeName(Me), strFunctionName, 281
            End If
        'SG 21/03/02 SYS4202
        End If
        'DB End
    End If
    ' PSC 11/04/00 AQR SYS0613 - End
        
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objGlobalParamDO = Nothing

    Exit Sub
    
ValidateFindProductListVbErr:

    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objGlobalParamDO = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Public Function FindProductNames(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Retrieves all instances of mortgage products that match the required
'               search criteria
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
 On Error GoTo FindProductNamesVbErr

    Const strFunctionName As String = "FindProductNames"

    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageProductBO_FindProductNames(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindProductNames = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

FindProductNamesExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindProductNamesVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindProductNames = objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindProductNamesExit

End Function

Private Sub AddDefaults(ByVal vxmlRequest As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:  Add default values for Language, Country of Origin and Max Records
'
' pass:         vxmlRequest     xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo AddDefaultsVbErr
    
    Const strFunctionName As String = "AddDefaults"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objErrAssist As ErrAssist
    Dim objXmlAssist As XMLAssist
    Dim objConvAssist As ConvertAssist
    Dim objGlobalParamDO As GlobalParameterDO
    Dim objComboDO As ComboDO
    Dim xmlElement As IXMLDOMElement
    Dim strTagName As String
            
    ' Add the country of origin  if not already set
    Dim strCountryOfOrigin As String
    
    Set objErrAssist = New ErrAssist
    Set objXmlAssist = New XMLAssist
    Set objConvAssist = New ConvertAssist
    
    strTagName = "COUNTRYOFORIGIN"
    Set xmlElement = vxmlRequest.getElementsByTagName(strTagName).Item(0)
    
    If xmlElement Is Nothing Then
        Set xmlElement = vxmlRequest.ownerDocument.createElement(strTagName)
        vxmlRequest.appendChild xmlElement
    End If
    
    If objContext Is Nothing Then
        Set objGlobalParamDO = New GlobalParameterDO
        Set objComboDO = New ComboDO
    Else
        Set objGlobalParamDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    If Len(xmlElement.Text) = 0 Then
        strCountryOfOrigin = objComboDO.GetFirstComboValueId("CountryOfOrigin", "UK")
        xmlElement.Text = strCountryOfOrigin
    End If
   
    ' Add the language  if not already set
    Dim strLanguage As String
    
    strTagName = "LANGUAGE"
    Set xmlElement = vxmlRequest.getElementsByTagName(strTagName).Item(0)
    
    If xmlElement Is Nothing Then
        Set xmlElement = vxmlRequest.ownerDocument.createElement(strTagName)
        vxmlRequest.appendChild xmlElement
    End If
    
    If Len(xmlElement.Text) = 0 Then
        strLanguage = objComboDO.GetFirstComboValueId("MortgageProductLanguage", "Eng")
        xmlElement.Text = strLanguage
    End If
    
    ' Add the LTV if not already set
    Dim strLTV As String
    
    strTagName = "LTV"
    Set xmlElement = vxmlRequest.getElementsByTagName(strTagName).Item(0)
        
    If xmlElement Is Nothing Then
        Set xmlElement = vxmlRequest.ownerDocument.createElement(strTagName)
        vxmlRequest.appendChild xmlElement
    End If
    
    If Len(xmlElement.Text) = 0 Then
        strLTV = objGlobalParamDO.GetCurrentParameterByType("DefaultLTV", "PERCENTAGE")
        xmlElement.Text = strLTV
    End If
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objGlobalParamDO = Nothing
    Set xmlElement = Nothing
    Set objComboDO = Nothing

    Exit Sub

AddDefaultsVbErr:
    
    Set objXmlAssist = Nothing
    Set objConvAssist = Nothing
    Set objGlobalParamDO = Nothing
    Set xmlElement = Nothing
    Set objComboDO = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function GetIncentivesForProduct(ByVal vstrXMLRequest As String) As String

' header ----------------------------------------------------------------------------------
' description:  Get the data associated wit ha mortgage product
'
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetIncentivesForProductVbErr
    
    Const strFunctionName As String = "GetIncentivesForProduct"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlData As New FreeThreadedDOMDocument40
    objXmlData.validateOnParse = False
    objXmlData.setProperty "NewParser", True
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlRequestNode As IXMLDOMNode
    Dim objXmlDataNode As IXMLDOMNode
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
    Dim objMortgageProductDO As MortgageProductDO

    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    Set objXmlDataNode = objXmlOut.appendChild(objXmlResponseElem)
    objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Dim blnRaiseError As Boolean
    blnRaiseError = True
    
    If objContext Is Nothing Then
        Set objMortgageProductDO = New MortgageProductDO
    Else
        Set objMortgageProductDO = objContext.CreateInstance(App.Title & ".MortgageProductDO")
    End If
        
    Set objXmlRequestNode = objXmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    
    Dim strXMLData As String
    
    If objXmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrTABLE_NAME & " tag not found"
    End If
    
    ' Get the inclusive incentives
    blnRaiseError = False
    strXMLData = objMortgageProductDO.GetInclusiveIncentivesForProduct(objXmlRequestNode.xml)
    

    If Len(strXMLData) > 0 Then
        Set objXmlData = objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlDataNode.appendChild(objXmlData.documentElement)
    End If
    
    'DPF 21/10/2002 - CPWP1 blank response string
    strXMLData = ""
    
    ' Get the exclusive incentives
    strXMLData = objMortgageProductDO.GetExclusiveIncentivesForProduct(objXmlRequestNode.xml)
    
    If Len(strXMLData) > 0 Then
        Set objXmlData = objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlDataNode.appendChild(objXmlData.documentElement)
    End If
          
    GetIncentivesForProduct = objXmlOut.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetIncentivesForProductExit:
    
    Set objXmlNode = Nothing
    Set objXmlDataNode = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlIn = Nothing
    Set objXmlData = Nothing
    Set objXmlOut = Nothing
    Set objXmlAssist = Nothing
    Set objXmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objMortgageProductDO = Nothing
    
    Exit Function

GetIncentivesForProductVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning objXmlResponseElem
        Resume Next
    End If

    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        If Not blnRaiseError Then
            Resume Next
        End If
    End If

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetIncentivesForProduct = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume GetIncentivesForProductExit

End Function

Private Function IMortgageProductBO_FindProductList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Retrieves all instances of mortgage products that match the required
'               search criteria
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindProductListVbErr
    
    Const strFunctionName As String = "IMortgageProductBO_FindProductList"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objXmlAssist As XMLAssist
    Dim objErrAssist As ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objIMortgageProductDO As IMortgageProductDO
    
    Set objXmlAssist = New XMLAssist
    Set objErrAssist = New ErrAssist
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIMortgageProductDO = New MortgageProductDO
    Else
        Set objIMortgageProductDO = objContext.CreateInstance(App.Title & ".MortgageProductDO")
    End If
                
    If vxmlRequest.nodeName = cstrTABLE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
    
    ValidateFindProductList xmlRequestNode
    AddDefaults xmlRequestNode
        
    Set xmlDataNode = objIMortgageProductDO.FindProductList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set IMortgageProductBO_FindProductList = xmlResponseElem.cloneNode(True)

FindProductListExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageProductDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindProductListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IMortgageProductBO_FindProductList = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindProductListExit

End Function

Public Function GetCapitalAndInterestRepaymentValueId() As String
' header ----------------------------------------------------------------------------------
'description:
'   Return the combo values for the entries found to be interest only repayments
'pass:
'
' return:       Capital and Interest Repayment Combo Value
'------------------------------------------------------------------------------------------
On Error GoTo GetCapitalAndInterestRepaymentValueIdVbErr

    Const strFunctionName As String = "GetCapitalAndInterestRepaymentValueId"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objComboDO As ComboDO
    Dim objXmlResponse As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If

    Dim strResponse As String
    strResponse = objComboDO.GetComboValueId("RepaymentType", "C")
    Set objXmlResponse = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)

    GetCapitalAndInterestRepaymentValueId = _
                    objXmlAssist.GetTagValue(objXmlResponse.documentElement, "VALUEID")

    Set objComboDO = Nothing
    Set objXmlResponse = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

GetCapitalAndInterestRepaymentValueIdVbErr:

    Set objComboDO = Nothing
    Set objXmlResponse = Nothing
    Set objXmlAssist = Nothing

    objErrAssist.AddToErrSource (strFunctionName)

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function GetPartAndPartRepaymentValueId() As String
' header ----------------------------------------------------------------------------------
'description:
'   Return the combo values for the entries found to be interest only repayments
'pass:
'
' return:       Part and Part Repayment Combo Value
'------------------------------------------------------------------------------------------
On Error GoTo GetPartAndPartRepaymentValueIdVbErr

    Const strFunctionName As String = "GetPartAndPartRepaymentValueId"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objComboDO As ComboDO
    Dim objXmlResponse As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If

    Dim strResponse As String
    strResponse = objComboDO.GetComboValueId("RepaymentType", "P")
    Set objXmlResponse = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)

    GetPartAndPartRepaymentValueId = _
                    objXmlAssist.GetTagValue(objXmlResponse.documentElement, "VALUEID")

    Set objComboDO = Nothing
    Set objXmlResponse = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

GetPartAndPartRepaymentValueIdVbErr:

    Set objComboDO = Nothing
    Set objXmlResponse = Nothing
    Set objXmlAssist = Nothing

    objErrAssist.AddToErrSource (strFunctionName)

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function GetInterestOnlyRepaymentValueId() As String
' header ----------------------------------------------------------------------------------
' description:
'   Return the combo values for the entries found to be interest only repayments
' pass:
'
' return:       Interest Only Repayment Combo Value
'------------------------------------------------------------------------------------------
On Error GoTo GetInterestOnlyRepaymentValueIdVbErr

    Const strFunctionName As String = "GetInterestOnlyRepaymentValueId"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim objComboDO As ComboDO
    Dim objXmlResponse As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist

    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If

    Dim strResponse As String
    strResponse = objComboDO.GetComboValueId("RepaymentType", "I")
    Set objXmlResponse = objXmlAssist.load(strResponse, TypeName(Me), strFunctionName)

    GetInterestOnlyRepaymentValueId = _
                    objXmlAssist.GetTagValue(objXmlResponse.documentElement, "VALUEID")

    Set objComboDO = Nothing
    Set objXmlResponse = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing

    Exit Function

GetInterestOnlyRepaymentValueIdVbErr:

    Set objComboDO = Nothing
    Set objXmlResponse = Nothing
    Set objXmlAssist = Nothing

    objErrAssist.AddToErrSource (strFunctionName)

    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function GetFullProductDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Retrieves all instances of mortgage products that match the required
'               search criteria
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
 On Error GoTo GetFullProductDetailsVbErr

    Const strFunctionName As String = "GetFullProductDetails"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    'PSC 18/06/2006 MAR1687 - Start
    Dim strSavePath As String
    Dim strFileName As String
    'PSC 18/06/2006 MAR1687 - End
    
    strSavePath = GetomMPDebugPath()
    
   Dim objXmlAssist As XMLAssist
    Dim objErrAssist As ErrAssist
   Set objXmlAssist = New XMLAssist
    Set objErrAssist = New ErrAssist
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    'PSC 18/06/2006 MAR1687 - Start
    If Len(strSavePath) > 0 Then
        strFileName = strSavePath & "GetFullProductDetailsRequest" & "-" & Format(Now, "yyyymmdd_hhnnss") & ".xml"
        SaveNodeToFile xmlIn.documentElement, strFileName
    End If
    'PSC 18/06/2006 MAR1687 - End
                                       
        ' Run validation checks
    Set xmlTempResponseNode = IMortgageProductBO_Validate(xmlIn.documentElement, bomGetFullProductDetails)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageProductBO_GetFullProductDetails(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetFullProductDetails = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetFullProductDetailsExit:

    'PSC 18/06/2006 MAR1687 - Start
    If Len(strSavePath) > 0 Then
        strFileName = strSavePath & "GetFullProductDetailsResponse" & "-" & Format(Now, "yyyymmdd_hhnnss") & ".xml"
        SaveNodeToFile xmlOut.documentElement, strFileName
    End If
    'PSC 18/06/2006 MAR1687 - End
    
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetFullProductDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetFullProductDetails = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetFullProductDetailsExit

End Function

Public Function GetMortgageProductLanguage(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Retrieves all instances of mortgage products that match the required
'               search criteria
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
 On Error GoTo GetMortgageProductLanguageVbErr

    Const strFunctionName As String = "GetMortgageProductLanguage"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
   Dim objXmlAssist As XMLAssist
    Dim objErrAssist As ErrAssist
   Set objXmlAssist = New XMLAssist
    Set objErrAssist = New ErrAssist
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
        ' Run validation checks
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageProductBO_GetMortgageProductLanguage(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetMortgageProductLanguage = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetMortgageProductLanguageExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetMortgageProductLanguageVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetMortgageProductLanguage = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetMortgageProductLanguageExit

End Function

Private Function IMortgageProductBO_FindProductNames(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Retrieves all instances of mortgage products that match the required
'               search criteria
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindProductNamesVbErr
    
    Const strFunctionName As String = "IMortgageProductBO_FindProductNames"


    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objIMortgageProductDO As IMortgageProductDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objIMortgageProductDO = New MortgageProductDO
    Else
        Set objIMortgageProductDO = m_objContext.CreateInstance(App.Title & ".MortgageProductDO")
    End If
                
    If vxmlRequest.nodeName = "MORTGAGEPRODUCT" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("MORTGAGEPRODUCT").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "MORTGAGEPRODUCT tag not found"
    End If
    
    'ValidateFindProductList xmlRequestNode
    
    'AddDefaults xmlRequestNode
        
    Set xmlDataNode = objIMortgageProductDO.FindProductNames(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IMortgageProductBO_FindProductNames = xmlResponseElem.cloneNode(True)

FindProductNamesExit:

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageProductDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindProductNamesVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IMortgageProductBO_FindProductNames = objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindProductNamesExit

End Function

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    Set m_objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")

End Sub

Private Function IMortgageProductBO_GetFullProductDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  Get the data associated wit ha mortgage product
'
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo IMortgageProductBO_GetFullProductDetailsVbErr
    
    Const strFunctionName As String = "IMortgageProductBO_GetFullProductDetails"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    
    Dim objIMortgageProductDO As IMortgageProductDO

    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIMortgageProductDO = New MortgageProductDO
    Else
        Set objIMortgageProductDO = objContext.CreateInstance(App.Title & ".MortgageProductDO")
    End If
        
    Set xmlTempResponseNode = objIMortgageProductDO.GetFullProductDetails(vxmlRequest)
    xmlResponseElem.appendChild xmlTempResponseNode
    
    Set IMortgageProductBO_GetFullProductDetails = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

IMortgageProductBO_GetFullProductDetailsExit:

    Set xmlResponseElem = Nothing

    Set xmlOut = Nothing

    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageProductDO = Nothing
    
    Exit Function

IMortgageProductBO_GetFullProductDetailsVbErr:

    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.validateOnParse = False
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.loadXML m_objErrAssist.FormatMessageNode
        'Set objXmlNode =
        xmlResponseElem.appendChild objXmlDoc.documentElement
        Set objXmlDoc = Nothing
        Resume Next
    End If

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IMortgageProductBO_GetFullProductDetails = objErrAssist.CreateErrorResponseEx
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    '   go to clean-up section
    Resume IMortgageProductBO_GetFullProductDetailsExit

End Function

Private Function IMortgageProductBO_Validate(ByVal vxmlRequest As IXMLDOMElement, _
                                    ByVal veboMethod As IOmiga4BOMethod) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------

On Error GoTo ValidateVbErr

    Const strFunctionName As String = "IMortgageProductBO_Validate"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objIMortgageProductDO As IMortgageProductDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    


    
    If m_objContext Is Nothing Then
        Set objIMortgageProductDO = New MortgageProductDO
    Else
        Set objIMortgageProductDO = m_objContext.CreateInstance(App.Title & ".MortgageProductDO")
    End If
                
    If vxmlRequest.nodeName = cstrTABLE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrTABLE_NAME).Item(0)
    End If
    
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
     
'    objIMortgageProductDO.AddDerivedData xmlRequestNode
            
    ' TODO ************************************************************************
    '   add hard coded validation here, & raise error if it fails
    ' *****************************************************************************
    Select Case veboMethod
    
    Case bomGetFullProductDetails
        'No validation
    Case Else
    
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            omiga4MissingOrInvalidEnumForValidation, "Mortgage Product BO"
    End Select
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set IMortgageProductBO_Validate = xmlResponseElem
    
ValidateExit:

    Set xmlRequestNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIMortgageProductDO = Nothing
    
    Exit Function

ValidateVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IMortgageProductBO_Validate = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateExit

End Function

Public Function FindProductConditionsList(ByVal vstrXMLRequest As String) As String

' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
'REQUEST :-

'<REQUEST OPERATION="FindProductConditionsList">
'   <MORTGAGEPRODUCTLIST>
'      <MORTGAGEPRODUCT MORTGAGEPRODUCTCODE="swp2" STARTDATE="01/01/2001"/>
'      <MORTGAGEPRODUCT MORTGAGEPRODUCTCODE="12345" STARTDATE="13/06/2002 09:00:00"/>
'      <MORTGAGEPRODUCT MORTGAGEPRODUCTCODE="somecode" STARTDATE="01/01/2001"/>
'   </MORTGAGEPRODUCTLIST>
'</REQUEST>

'RESPONSE :-
'<RESPONSE TYPE="SUCCESS">
'   <CONDITIONSLIST>
'      <CONDITIONS>
'         <CONDITIONREFERENCE>515</CONDITIONREFERENCE>
'         <CONDITIONNAME>GD Test ConditionName5</CONDITIONNAME>
'         <CONDITIONDESCRIPTION>GD Test Condition Description5ertertert</CONDITIONDESCRIPTION>
'         <CONDITIONTYPE TEXT="Product Special Condition">30</CONDITIONTYPE>
'         <EDITABLEIND>0</EDITABLEIND>
'         <FREEFORMATIND>0</FREEFORMATIND>
'         <CONDITIONIND>0</CONDITIONIND>
'         <CHANNELID>1</CHANNELID>
'         <RULEREFERENCE/>
'         <DELETEFLAG>0</DELETEFLAG>
'      </CONDITIONS>
'      <CONDITIONS>
'         <CONDITIONREFERENCE>516</CONDITIONREFERENCE>
'         <CONDITIONNAME>GD Test ConditionName6</CONDITIONNAME>
'         <CONDITIONDESCRIPTION>GD Test Condition Description6</CONDITIONDESCRIPTION>
'         <CONDITIONTYPE TEXT="Product Special Condition">30</CONDITIONTYPE>
'         <EDITABLEIND>0</EDITABLEIND>
'         <FREEFORMATIND>0</FREEFORMATIND>
'         <CONDITIONIND>0</CONDITIONIND>
'         <CHANNELID>7</CHANNELID>
'         <RULEREFERENCE/>
'         <DELETEFLAG>0</DELETEFLAG>
'      </CONDITIONS>
'   </CONDITIONSLIST>
'</RESPONSE>
 
 On Error GoTo FindProductConditionsListVbErr

    Const strFunctionName As String = "FindProductConditionsList"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageProductBO_FindProductConditionsList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindProductConditionsList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindProductConditionsListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindProductConditionsListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindProductConditionsList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindProductConditionsListExit

End Function

'BMIDS00654 MDC 01/11/2002
Public Function GetMortgageProductIncomeMultiplierSet(ByVal vstrXMLRequest As String) As String

 On Error GoTo GetMortgageProductIncomeMultiplierSetVbErr

    Const strFunctionName As String = "GetMortgageProductIncomeMultiplierSet"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
   Dim objXmlAssist As XMLAssist
    Dim objErrAssist As ErrAssist
   Set objXmlAssist = New XMLAssist
    Set objErrAssist = New ErrAssist
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageProductBO_GetMortgageProductIncomeMultiplierSet(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetMortgageProductIncomeMultiplierSet = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetMortgageProductIncomeMultiplierSetExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetMortgageProductIncomeMultiplierSetVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetMortgageProductIncomeMultiplierSet = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetMortgageProductIncomeMultiplierSetExit

End Function
'BMIDS00654 MDC 01/11/2002 - End
'PSC 18/06/2006 MAR1687 - Start
Private Function GetomMPDebugPath() As String
  
    Const strRegSection As String = "HKLM\SOFTWARE\Omiga4\System Configuration\"
    
    Dim objWshShell     As Object
    Dim strDebugPath    As String
       
    ' Ignore any errors that occur
    On Error Resume Next
    
    Set objWshShell = CreateObject("WScript.Shell")
    
    ' Read the path to save admin debugging info to from the registry
    strDebugPath = Trim(objWshShell.RegRead(strRegSection & "omMPDebugPath"))
    
    If Len(strDebugPath) > 0 Then
    
        If Right(strDebugPath, 1) <> "\" Then
            strDebugPath = strDebugPath & "\"
        End If
        
        If Len(Dir(strDebugPath, vbDirectory)) = 0 Then
            'Create the folder if it does not exist
            MkDir strDebugPath
            If Len(Dir(strDebugPath, vbDirectory)) = 0 Then
                ' The output folder does not exist and cannot be created
                strDebugPath = vbNullString
            End If
        End If
    End If
    
    GetomMPDebugPath = strDebugPath
    
    'Clear any errors that may have occurred as they can be ignored
    Err.Clear
End Function

Private Sub SaveNodeToFile(ByVal vxmlNode As IXMLDOMNode, ByVal vstrFileName As String)

    Dim xmlDoc      As FreeThreadedDOMDocument40
    
    'Ignore any errors that occur
    On Error Resume Next
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.appendChild vxmlNode.cloneNode(True)
    
    xmlDoc.Save vstrFileName
    
    Set xmlDoc = Nothing
End Sub
'PSC 18/06/2006 MAR1687 - End



