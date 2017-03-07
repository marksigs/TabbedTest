VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "MortgageSubQuoteDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      MortgageSubQuoteDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for MortgageSubQuote Data Object

'               Public Function GetData(vstrXMLRequest As String) As String
'
'Dependencies:
'               Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    10/11/99    Created
'MCS    04/01/2000  Interfacing added and Findloan componentlist added
'JLD    07/03/2000  Fixes in FindLoanComponentList
'JLD    17/04/00    Change to FindSubQuoteDetails view and data returned from it.
'PSC    07/07/00    SYS1126 - Amend AmountRequested to TotalLoanAmount
'IVW    09/11/00    Amended to disconnect record sets more efficiently
'SR     18/06/01    SYS2377 - modified method 'TotalIncentiveForAllComponents'
'AD     04/10/01    Removed the MoveFirst as it's a ForwardOnly. SYS2718.
'------------------------------------------------------------------------------------------
'--- BMIDS specific history-----------------------------------------------------------------
'DPF    19/07/02    CMWP3 - Added new method IMortgageSubQuoteDO_CreateLoanComponentRedemptionFee()
'GHun   31/10/2003  BMIDS622 Clean up and memory leak fix
'GHun   20/04/2004  BMIDS736 Core49 Alpha+ integration
'MC     12/06/2004  bmids763 - applicationdate used.
'GHun   22/06/2004  BMIDS782 Removed unused methods and Added SaveIllustrationResults
'SR    16/08/2004   BMIDS815 - Modified method 'FindLoanComponentList'
'-------------------------------------------------------------------------------------------
' MARS History
'GHun   29/01/2006  MAR1130 Changed GetLoanComponentProductParameters
'AS     22/03/2006  MAR1461 PERFORMANCE - Error INSERT INTO MORTGAGEONEOFFCOST. BBG1486 and BBG1313 Amend IMortgageSubQuoteDO_GetDataForOneOffCostsDetails()
'AS     23/03/2006  MAR1495 PERFORMANCE - Apply BBG1984 Amend IMortgageSubQuoteDO_TotalIncentivesForAllComponents()
'JD     07/05/2006  MAR1703 Amend reservemortgageproduct to use stored procedure
'-------------------------------------------------------------------------------------------
'Epsom History
'GHun   12/02/2007  EP2_954 Changed SaveIllustrationResults for additional borrowing
'-------------------------------------------------------------------------------------------

Option Explicit
Implements IMortgageSubQuoteDO

'BMIDS622
Implements ObjectControl

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
'BMIDS622 End

Private Sub IMortgageSubQuoteDO_Create(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_Create"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    Dim objIDOAssist As IDOAssist
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageSubQuoteData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IMortgageSubQuoteDO_CreateLoanComponent(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateLoanComponentVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_CreateLoanComponent"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    Dim objIDOAssist As IDOAssist
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateLoanComponentVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IMortgageSubQuoteDO_CreateLoanComponentRedemptionFee(ByVal vxmlTableElement As MSXML2.IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateLoanComponentRedemptionFeeVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_CreateLoanComponentRedemptionFee"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    Set xmlClassDefDoc = New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    
    Dim objIDOAssist As IDOAssist
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentRedemptionFeeData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateLoanComponentRedemptionFeeVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description
End Sub

'BMIDS782 GHun
'BMIDS736 GHun No longer used, but cannot be totally removed due to OOSS
'Private Sub IMortgageSubQuoteDO_CreateLoanComponentResults(ByVal vxmlTableElement As IXMLDOMElement)
'' header ----------------------------------------------------------------------------------
'' description:
''   create an instance of the persistant data associated with this data object
''   for each set of data in the request
'' pass:
''   vxmlxmlTableElement  xml element containing the request
'' return:                n/a
''------------------------------------------------------------------------------------------
'On Error GoTo CreateLoanComponentResultsVbErr
'
'    Const strFunctionName As String = "CreateLoanComponentResults"
'
'    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
'    Dim objIomCMClassDef As IomCMClassDef
'
'    Dim objIDOAssist As IDOAssist
'
'    If m_objContext Is Nothing Then
'        Set objIDOAssist = New DOAssist
'    Else
'        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
'    End If
'
'    Set objIomCMClassDef = New omCMClassDef
'    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentResultsData
'    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
'
'    Set objIomCMClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
'    Set objIDOAssist = Nothing
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
'
'    Exit Sub
'
'CreateLoanComponentResultsVbErr:
'
'    Set objIomCMClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
'    Set objIDOAssist = Nothing
'
'    m_objErrAssist.AddToErrSource (strFunctionName)
'
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
'    End If
'
'    If Not m_objContext Is Nothing Then
'        If m_objErrAssist.IsSystemError Then
'            m_objContext.SetAbort
'        Else
'            m_objContext.SetComplete
'        End If
'    End If
'
'    '   re-raise error for business object to interpret as appropriate
'    Err.Raise Err.number, Err.Source, Err.Description
'End Sub
'BMIDS736
'BMIDS782 End

Private Sub IMortgageSubQuoteDO_CreateMortgageOneOffCost(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageOneOffCostVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_CreateMortgageOneOffCost"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    Dim objIDOAssist As IDOAssist
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageOneOffCostData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateMortgageOneOffCostVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IMortgageSubQuoteDO_CreateNonPanelMortgageProduct(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateNonPanelMortgageProductVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_CreateNonPanelMortgageProduct"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    Dim objIDOAssist As IDOAssist
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadNonPanelMortgageProductData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateNonPanelMortgageProductVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IMortgageSubQuoteDO_CreateMortgageIncentive(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageIncentiveVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_CreateMortgageIncentive"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    Dim objIDOAssist As IDOAssist
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageIncentiveData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateMortgageIncentiveVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IMortgageSubQuoteDO_CreateOneOffCosts(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateOneOffCostsVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_CreateOneOffCosts"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    Dim objIDOAssist As IDOAssist
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageOneOffCostData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateOneOffCostsVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IMortgageSubQuoteDO_DeleteAllLoanComponent(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllLoanComponentVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_DeleteAllLoanComponent"

    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strMortgageSubQuoteNumber As String
    
    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGESUBQUOTENUMBER")
    
    If Len(strApplicationNumber) = 0 Or _
        Len(strApplicationFactFindNumber) = 0 Or _
        Len(strMortgageSubQuoteNumber) = 0 Then
            
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
        oeMissingPrimaryTag, _
        ",APPLICATIONNUMBER =" & strApplicationNumber & _
        ",APPLICATIONFACTFINDNUMBER =" & strApplicationFactFindNumber & _
        ",MORTGAGESUBQUOTENUMBER =" & strMortgageSubQuoteNumber & "Expected"
    
    End If
                
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentData
    objIDOAssist.DeleteAll vxmlTableElement, xmlClassDefDoc
 
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

DeleteAllLoanComponentVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

'BMIDS782 GHun
'BMIDS736 GHun No longer used but cannot be totally removed due to OOSS
'Private Sub IMortgageSubQuoteDO_DeleteAllLoanComponentResults(ByVal vxmlTableElement As IXMLDOMElement)
'' header ----------------------------------------------------------------------------------
'' description:
''   Delete all instances of the persistant data associated with this
''   data object that match the key values specified
'' pass:
''   vxmlxmlTableElement  xml element containing the request
'' return:                n/a
''------------------------------------------------------------------------------------------
'On Error GoTo DeleteAllLoanComponentResultsVbErr
'
'    Const strFunctionName As String = "DeleteAllLoanComponentResults"
'
'    Dim objIDOAssist As IDOAssist
'    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
'    Dim objIomCMClassDef As IomCMClassDef
'
'    If m_objContext Is Nothing Then
'        Set objIDOAssist = New DOAssist
'    Else
'        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
'    End If
'
'    Dim strApplicationNumber As String
'    Dim strApplicationFactFindNumber As String
'    Dim strMortgageSubQuoteNumber As String
'
'    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
'    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
'    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGESUBQUOTENUMBER")
'
'    If Len(strApplicationNumber) = 0 Or _
'        Len(strApplicationFactFindNumber) = 0 Or _
'        Len(strMortgageSubQuoteNumber) = 0 Then
'
'        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
'        oeMissingPrimaryTag, _
'        ",APPLICATIONNUMBER =" & strApplicationNumber & _
'        ",APPLICATIONFACTFINDNUMBER =" & strApplicationFactFindNumber & _
'        ",MORTGAGESUBQUOTENUMBER =" & strMortgageSubQuoteNumber & "Expected"
'
'    End If
'
'    Set objIomCMClassDef = New omCMClassDef
'    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentResultsData
'    objIDOAssist.DeleteAll vxmlTableElement, xmlClassDefDoc
'
'    Set objIomCMClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
'    Set objIDOAssist = Nothing
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
'
'    Exit Sub
'
'DeleteAllLoanComponentResultsVbErr:
'
'    Set objIomCMClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
'    Set objIDOAssist = Nothing
'
'    m_objErrAssist.AddToErrSource (strFunctionName)
'
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
'    End If
'
'    If Not m_objContext Is Nothing Then
'        If m_objErrAssist.IsSystemError Then
'            m_objContext.SetAbort
'        Else
'            m_objContext.SetComplete
'        End If
'    End If
'
'    '   re-raise error for business object to interpret as appropriate
'    Err.Raise Err.number, Err.Source, Err.Description
'End Sub
'BMIDS736 End
'BMIDS782 End

Private Sub IMortgageSubQuoteDO_DeleteAllMortgageIncentive(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllMortgageIncentiveVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_DeleteAllMortgageIncentive"

    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strMortgageSubQuoteNumber As String
    
    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGESUBQUOTENUMBER")
    
    If Len(strApplicationNumber) = 0 Or _
        Len(strApplicationFactFindNumber) = 0 Or _
        Len(strMortgageSubQuoteNumber) = 0 Then
            
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
        oeMissingPrimaryTag, _
        ",APPLICATIONNUMBER =" & strApplicationNumber & _
        ",APPLICATIONFACTFINDNUMBER =" & strApplicationFactFindNumber & _
        ",MORTGAGESUBQUOTENUMBER =" & strMortgageSubQuoteNumber & "Expected"
    
    End If
                                                
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageIncentiveData
    objIDOAssist.DeleteAll vxmlTableElement, xmlClassDefDoc
 
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

DeleteAllMortgageIncentiveVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IMortgageSubQuoteDO_DeleteAllMortgageOneOffCost(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllMortgageOneOffCostVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_DeleteAllMortgageOneOffCost"

    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageOneOffCostData
    objIDOAssist.DeleteAll vxmlTableElement, xmlClassDefDoc
 
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

DeleteAllMortgageOneOffCostVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IMortgageSubQuoteDO_DeleteAllNonPanelMortgageProduct(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllNonPanelMortgageProductVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_DeleteAllNonPanelMortgageProduct"

    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strMortgageSubQuoteNumber As String
    
    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGESUBQUOTENUMBER")
    
    If Len(strApplicationNumber) = 0 Or _
        Len(strApplicationFactFindNumber) = 0 Or _
        Len(strMortgageSubQuoteNumber) = 0 Then
            
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
        oeMissingPrimaryTag, _
        ",APPLICATIONNUMBER =" & strApplicationNumber & _
        ",APPLICATIONFACTFINDNUMBER =" & strApplicationFactFindNumber & _
        ",MORTGAGESUBQUOTENUMBER =" & strMortgageSubQuoteNumber & "Expected"
    
    End If
                                
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadNonPanelMortgageProductData
    objIDOAssist.DeleteAll vxmlTableElement, xmlClassDefDoc
 
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

DeleteAllNonPanelMortgageProductVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IMortgageSubQuoteDO_DeleteLoanComponent(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLoanComponentVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_DeleteLoanComponent"

    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    Dim blnRaiseError As Boolean
    
    blnRaiseError = False
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    'BMIDS736 GHun
    'm_objXmlAssist.ChangeNodeName vxmlTableElement, "LOANCOMPONENT", "LOANCOMPONENTRESULTS"
    'IMortgageSubQuoteDO_DeleteLoanComponentResults vxmlTableElement
    m_objXmlAssist.ChangeNodeName vxmlTableElement, "LOANCOMPONENT", "LOANCOMPONENTBALANCESCHEDULE"
    IMortgageSubQuoteDO_DeleteAllLoanComponentBalanceSchedules vxmlTableElement
    m_objXmlAssist.ChangeNodeName vxmlTableElement, "LOANCOMPONENTBALANCESCHEDULE", "LOANCOMPONENTPAYMENTSCHEDULE"
    IMortgageSubQuoteDO_DeleteAllLoanComponentPaymentSchedules vxmlTableElement
    'BMIDS736 End
   
    'IMortgageSubQuoteDO_DeleteMirasDetails vxmlTableElement
    
    'This deleteallMortgageIncentive only checks 3 of the 4 primary keys but
    'the previous method has all ready checked that all 4 keys exist so the
    'next call will only delete one record
    
    'BMIDS736 GHun
    'm_objXmlAssist.ChangeNodeName vxmlTableElement, "LOANCOMPONENTRESULTS", "MORTGAGEINCENTIVE"
    m_objXmlAssist.ChangeNodeName vxmlTableElement, "LOANCOMPONENTPAYMENTSCHEDULE", "MORTGAGEINCENTIVE"
    'BMIDS736 End
    
    IMortgageSubQuoteDO_DeleteAllMortgageIncentive vxmlTableElement
    
    m_objXmlAssist.ChangeNodeName vxmlTableElement, "MORTGAGEINCENTIVE", "NONPANELMORTGAGEPRODUCT"
    
    IMortgageSubQuoteDO_DeleteNonPanelMortgageProduct vxmlTableElement
                        
    m_objXmlAssist.ChangeNodeName vxmlTableElement, "NONPANELMORTGAGEPRODUCT", "LOANCOMPONENT"
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentData
    blnRaiseError = True
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

DeleteLoanComponentVbErr:

    If m_objErrAssist.GetOmigaErrorNumber(Err.number) = oeNoRowsAffected Or _
        m_objErrAssist.GetOmigaErrorNumber(Err.number) = oeNoRowsAffectedByDeleteAll Then
        If Not blnRaiseError Then
            Resume Next
        End If
    End If

    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

'BMIDS782 GHun
'BMIDS736 GHun No longer used but cannot be totally removed due to OOSS
'Private Sub IMortgageSubQuoteDO_DeleteLoanComponentResults(ByVal vxmlTableElement As IXMLDOMElement)
'' header ----------------------------------------------------------------------------------
'' description:
''   Delete a single instance of the persistant data associated with this
''   data object
'' pass:
''   vxmlxmlTableElement  xml element containing the request
'' return:                n/a
''------------------------------------------------------------------------------------------
'On Error GoTo DeleteLoanComponentResultsVbErr
'
'    Const strFunctionName As String = "DeleteLoanComponentResults"
'
'    Dim objIDOAssist As IDOAssist
'    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
'    Dim objIomCMClassDef As IomCMClassDef
'
'    If m_objContext Is Nothing Then
'        Set objIDOAssist = New DOAssist
'    Else
'        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
'    End If
'
'    Set objIomCMClassDef = New omCMClassDef
'    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentResultsData
'    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
'
'    Set objIomCMClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
'    Set objIDOAssist = Nothing
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
'
'    Exit Sub
'
'DeleteLoanComponentResultsVbErr:
'
'    Set objIomCMClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
'    Set objIDOAssist = Nothing
'
'    m_objErrAssist.AddToErrSource (strFunctionName)
'
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
'    End If
'
'    If Not m_objContext Is Nothing Then
'        If m_objErrAssist.IsSystemError Then
'            m_objContext.SetAbort
'        Else
'            m_objContext.SetComplete
'        End If
'    End If
'
'    '   re-raise error for business object to interpret as appropriate
'    Err.Raise Err.number, Err.Source, Err.Description
'End Sub
'BMIDS736 End
'BMIDS782 End

Private Sub IMortgageSubQuoteDO_DeleteNonPanelMortgageProduct(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteNonPanelMortgageProductVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_DeleteNonPanelMortgageProduct"

    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadNonPanelMortgageProductData
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

DeleteNonPanelMortgageProductVbErr:

    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Function IMortgageSubQuoteDO_FindActiveLoanComponentsList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:
'                   xml node that has been attached
'------------------------------------------------------------------------------------------
On Error GoTo FindActiveLoanComponentsListVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_FindActiveLoanComponentsList"

    Dim objDOAssist As DOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    If m_objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadActiveLoanComponentsViewData
            
    Set IMortgageSubQuoteDO_FindActiveLoanComponentsList = objDOAssist.FindListEx(vxmlTableElement, xmlClassDefDoc)

    Set xmlClassDefDoc = Nothing
    Set objDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

FindActiveLoanComponentsListVbErr:
    
    Set xmlClassDefDoc = Nothing
    Set objDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IMortgageSubQuoteDO_FindLoanComponentList(ByVal vxmlTableElement As _
                            IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
'   vxmlInNode      xml node to which data should be attached
' return:
'                   xml node that has been attached
'------------------------------------------------------------------------------------------
On Error GoTo FindLoanComponentListVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_FindLoanComponentList"

    Dim objDOAssist         As DOAssist
    Dim xmlClassDefDoc      As FreeThreadedDOMDocument40
    Dim objIomCMClassDef    As IomCMClassDef
    Dim xmlNewElement       As IXMLDOMElement
    Dim xmlTempDoc          As FreeThreadedDOMDocument40
    Dim tempTopElem         As IXMLDOMElement
    Dim tempElem            As IXMLDOMElement
    Dim xmlTempNewElement   As IXMLDOMElement
    Dim strPortingData As String   'SR 16/08/2004:BMIDS815
    
    If m_objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set xmlNewElement = vxmlTableElement.cloneNode(True)
    
    Set xmlTempDoc = New FreeThreadedDOMDocument40
    xmlTempDoc.validateOnParse = False
    xmlTempDoc.setProperty "NewParser", True
    Set tempTopElem = xmlTempDoc.createElement("LOANCOMPONENT")
    xmlTempDoc.appendChild tempTopElem
    
    Set tempElem = xmlTempDoc.createElement("APPLICATIONNUMBER")
    Set xmlTempNewElement = xmlNewElement.getElementsByTagName("APPLICATIONNUMBER").Item(0)
    tempElem.Text = xmlTempNewElement.Text
    tempTopElem.appendChild tempElem
    
    Set tempElem = xmlTempDoc.createElement("APPLICATIONFACTFINDNUMBER")
    Set xmlTempNewElement = xmlNewElement.getElementsByTagName("APPLICATIONFACTFINDNUMBER").Item(0)
    tempElem.Text = xmlTempNewElement.Text
    tempTopElem.appendChild tempElem
    
    Set tempElem = xmlTempDoc.createElement("MORTGAGESUBQUOTENUMBER")
    Set xmlTempNewElement = xmlNewElement.getElementsByTagName("MORTGAGESUBQUOTENUMBER").Item(0)
    tempElem.Text = xmlTempNewElement.Text
    tempTopElem.appendChild tempElem
    
    Set objIomCMClassDef = New omCMClassDef
        'SR 16/08/2004:BMIDS815
    strPortingData = m_objXmlAssist.GetTagValue(vxmlTableElement, "PORTINGDATAONLY")
    If strPortingData = "1" Then
        Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentData(LoanComponentPortingData)
    Else
        Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentData
    End If 'SR 16/08/2004:BMIDS815 - End
            
    'Set IMortgageSubQuoteDO_FindLoanComponentList = objDOAssist.FindListEx(vxmlTableElement, xmlClassDefDoc)
    Set IMortgageSubQuoteDO_FindLoanComponentList = objDOAssist.FindListEx(xmlTempDoc.documentElement, xmlClassDefDoc)

    Set xmlClassDefDoc = Nothing
    Set xmlTempDoc = Nothing
    Set objDOAssist = Nothing
    Set objIomCMClassDef = Nothing
    Set xmlNewElement = Nothing
    Set tempTopElem = Nothing
    Set tempElem = Nothing
    Set xmlTempNewElement = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

FindLoanComponentListVbErr:
    
    Set xmlClassDefDoc = Nothing
    Set xmlTempDoc = Nothing
    Set objDOAssist = Nothing
    Set objIomCMClassDef = Nothing
    Set xmlNewElement = Nothing
    Set tempTopElem = Nothing
    Set tempElem = Nothing
    Set xmlTempNewElement = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IMortgageSubQuoteDO_FindMortgageIncentiveList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindMortgageIncentiveListVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_FindMortgageIncentiveList"

    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageIncentiveData
    Set IMortgageSubQuoteDO_FindMortgageIncentiveList = objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)

    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

FindMortgageIncentiveListVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        'BMIDS622 Memory leak fix: MSQDO is not released when SetComplete is called here
        'Else
        '    m_objContext.SetComplete
        'BMIDS622 End
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IMortgageSubQuoteDO_FindMortgageOneOffCostList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindMortgageOneOffCostListVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_FindMortgageOneOffCostList"

    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageOneOffCostData
    Set IMortgageSubQuoteDO_FindMortgageOneOffCostList = objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)

    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

FindMortgageOneOffCostListVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IMortgageSubQuoteDO_FindNonPanelMortgageProductList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindNonPanelMortgageProductListVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_FindNonPanelMortgageProductList"

    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadNonPanelMortgageProductData
    Set IMortgageSubQuoteDO_FindNonPanelMortgageProductList = objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)

    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

FindNonPanelMortgageProductListVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        'BMIDS622 Memory leak fix: MSQDO is not released when SetComplete is called here
        'Else
        '    m_objContext.SetComplete
        'BMIDS622 End
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IMortgageSubQuoteDO_FindSubQuoteDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlRequest     xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
' return:
'   GetDataEx       string containing XML data stream representation of
'                   data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo FindSubQuoteDetailsVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_FindSubQuoteDetails"

    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                        
    Dim strInApplicationNumber As String
    Dim strInApplicationFactFindNumber As String
    Dim strInActiveMortgageSubQuoteNumber As String

    'Need to get the values for these fields out of the XML passed in
    strInApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strInApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strInActiveMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "ACTIVEQUOTENUMBER")


    If Len(strInApplicationNumber) = 0 Or _
        Len(strInApplicationFactFindNumber) = 0 Or _
        Len(strInActiveMortgageSubQuoteNumber) = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4MissingPrimaryTag, _
        "APPLICATIONNUMBER =" & strInApplicationNumber & _
        ",APPLICATIONFACTFINDNUMBER =" & strInApplicationFactFindNumber & _
        ",ACTIVEQUOTENUMBER =" & strInActiveMortgageSubQuoteNumber & "Expected"
    End If
                                
    Dim strSQL As String
            
    Dim objSQLAssist As SQLAssist
    Set objSQLAssist = New SQLAssist
    ' build the full select statement
    'Using a view
    strSQL = "SELECT * FROM " & "FindSubQuoteDetails" & " WHERE " _
            & "APPLICATIONNUMBER = " & objSQLAssist.FormatString(strInApplicationNumber) _
            & "AND APPLICATIONFACTFINDNUMBER = " & strInApplicationFactFindNumber _
            & " ORDER BY MORTGAGESUBQUOTENUMBER"
            
    'IVW - Disconnecting record sets
    Dim cmd As ADODB.Command
    Dim conn As ADODB.Connection
    Dim rstThisRecSet As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    Set conn = New ADODB.Connection
    Set rstThisRecSet = New ADODB.Recordset

    cmd.CommandText = strSQL
    
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.CursorType = adOpenForwardOnly
    rstThisRecSet.LockType = adLockReadOnly
    
    conn.Open objIADOAssist.GetConnStr
    cmd.ActiveConnection = conn
            
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
    Set rstThisRecSet.ActiveConnection = Nothing
    
    'IVW
            
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim objXMLTempNode As IXMLDOMNode
    Dim xmlRootNode As IXMLDOMNode
    Dim objXMLTempElement As IXMLDOMElement
    
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    Set objXMLTempNode = xmlOutDoc.createElement("MORTGAGESUBQUOTEDETAILS")
    Set xmlRootNode = xmlOutDoc.appendChild(objXMLTempNode)
    
    Dim objXMLActiveQuoteNode As IXMLDOMNode
        
    Set objXMLActiveQuoteNode = xmlOutDoc.createElement("ACTIVEQUOTE")

    Dim objIomCMClassDef As IomCMClassDef
    Dim strCurrentMortgageSubQuoteNumber As String
    Dim strTempMortgageSubQuoteNumber As String

    Set objIomCMClassDef = New omCMClassDef
                 
    Dim objXMLSubQuoteNode As IXMLDOMNode
    Dim objXMLLoanComponentNode As IXMLDOMNode
    Dim objXMLMortgageProductNode As IXMLDOMNode
    Dim objXMLNonPanelMortgageNode  As IXMLDOMNode
                 
    Dim objConvertAssist As ConvertAssist
    Set objConvertAssist = New ConvertAssist
        
    Dim xmlClassDefLoanComponentDom As FreeThreadedDOMDocument40
    
    Set xmlClassDefLoanComponentDom = objIomCMClassDef.LoadLoanComponentData(LoanComponentView)
    
    Dim xmlClassDefMortgageSubQuoteDom As FreeThreadedDOMDocument40
                 
    Set xmlClassDefMortgageSubQuoteDom = objIomCMClassDef.LoadMortgageSubQuoteData
                 
    Dim xmlClassDefNonPanelMortgageProductDom As FreeThreadedDOMDocument40
    
    Set xmlClassDefNonPanelMortgageProductDom = objIomCMClassDef.LoadNonPanelMortgageProductData
    
    Dim xmlClassMortgageProductDefDom As FreeThreadedDOMDocument40
        
    Dim objIomMPClassDef As IomMPClassDef
    Set objIomMPClassDef = New omMPClassDef
    'Set xmlClassMortgageProductDefDom = objIomMPClassDef.LoadMortgageProductData
    Set xmlClassMortgageProductDefDom = objIomMPClassDef.LoadQuotationSummaryViewData
    
    Set objIomMPClassDef = Nothing
    
    If rstThisRecSet.EOF Then
        
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound, _
                "No mortgagesubquote record found"
    
    End If
        
    Dim blnAdd As Boolean

    blnAdd = False
    
    ' AD 04/10/01 Removed the MoveFirst as it's a ForwardOnly. SYS2718.

    ' loop through the record set
    While Not rstThisRecSet.EOF
                    
        'find the loan component first
        Set objXMLLoanComponentNode = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlClassDefLoanComponentDom)
        'Get the mortgagesubquote number for this loan component
        strTempMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(objXMLLoanComponentNode, "MORTGAGESUBQUOTENUMBER")
        
        If Len(strCurrentMortgageSubQuoteNumber) = 0 Then
            'we assume that this must be the first time in here so
            'set the MortgageSubQuoteNumber to this one
            strCurrentMortgageSubQuoteNumber = strTempMortgageSubQuoteNumber
            'and set the flag to add a MortgageSubQuote record
            blnAdd = True
            
        Else
            'We have an existing subquotenumber
            If StrComp(strCurrentMortgageSubQuoteNumber, strTempMortgageSubQuoteNumber, vbTextCompare) = 0 Then
                'same sub quote so just add the loan component to it
                Set objXMLTempNode = objXMLSubQuoteNode.appendChild(objXMLLoanComponentNode)
                'and set flag so we don't add a MortgageSubQuote record
                blnAdd = False
                
            Else
                
                'set the strCurrentMortgageSubQuoteNumber to the new one
                strCurrentMortgageSubQuoteNumber = strTempMortgageSubQuoteNumber
                'and set the flag to add a MortgageSubQuote record
                blnAdd = True
            End If
        End If
                    
        If blnAdd Then
        
            'new subquote needed so create one and add the loan component to it
            Set objXMLTempNode = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlClassDefMortgageSubQuoteDom)
            'append the Mortgage sub quote data to the root node
            Set objXMLSubQuoteNode = xmlRootNode.appendChild(objXMLTempNode)
            'Is this the active mortgagesubquote ??
            If StrComp(strCurrentMortgageSubQuoteNumber, strInActiveMortgageSubQuoteNumber, vbTextCompare) = 0 Then
                'it is so append active quote to it
                Set objXMLTempNode = objXMLSubQuoteNode.appendChild(objXMLActiveQuoteNode)
            End If
            
            'and append the current loan component to the new sub quote xml
            Set objXMLTempNode = objXMLSubQuoteNode.appendChild(objXMLLoanComponentNode)
                            
        End If
        'Get the mortgageproduct data out
        Set objXMLMortgageProductNode = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlClassMortgageProductDefDom)
                    
        If objXMLTempNode Is Nothing Then
              
            ' no record raise errror
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound, _
                "No mortgageproduct or nonpanelmortgageproduct record found"
        
        Else
              
            'append this mortgageproduct to this loan component
            Set objXMLTempElement = objXMLLoanComponentNode.appendChild(objXMLMortgageProductNode)
            
            'check if its a panel lender
            If objConvertAssist.CSafeBool(m_objXmlAssist.GetTagValue(objXMLTempNode, "NONPANELLENDEROPTION")) Then
            
                'get the non panel mortgage product out
                
                Set objXMLNonPanelMortgageNode = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlClassDefNonPanelMortgageProductDom)
                
                If objXMLTempNode Is Nothing Then
                  
                    ' no record raise errror
                    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound, _
                        "No mortgageproduct or nonpanelmortgageproduct record found"
                    
                Else
                    'append the NonPanelMortgageProduct
                        
                    Set objXMLTempElement = objXMLMortgageProductNode.appendChild(objXMLNonPanelMortgageNode)
            
                End If

            End If
        
        End If
            
        rstThisRecSet.MoveNext
            
    Wend

    Set IMortgageSubQuoteDO_FindSubQuoteDetails = xmlOutDoc.documentElement
    
    Set cmd = Nothing
    conn.Close
    Set conn = Nothing

    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
    
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomCMClassDef = Nothing
    Set objSQLAssist = Nothing
    Set xmlOutDoc = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

FindSubQuoteDetailsVbErr:
    
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomCMClassDef = Nothing
    Set objSQLAssist = Nothing
    Set xmlOutDoc = Nothing
    
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If
    
    ' Close the connection if still open
    If conn.State = adStateOpen Then
        conn.Close
    End If

    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IMortgageSubQuoteDO_GetData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_GetData"
       
    Dim objIDOAssist As IDOAssist
    Dim xmlomCMClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlomCMClassDefDoc = objIomCMClassDef.LoadMortgageSubQuoteData
    'Dim obj As IXMLDOMNode

    'Set obj = objIDOAssist.GetData(vxmlTableElement, xmlomCMClassDefDoc)
    'Set IMortgageSubQuoteDO_GetData = obj
    Set IMortgageSubQuoteDO_GetData = objIDOAssist.GetData(vxmlTableElement, xmlomCMClassDefDoc)
    Set objIomCMClassDef = Nothing
    Set xmlomCMClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetDataVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlomCMClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IMortgageSubQuoteDO_GetLoanComponentData(ByVal vxmlTableElement As _
                            IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetLoanComponentDataVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_GetLoanComponentData"
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    Dim strPortingData As String
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    'SR 16/08/2004:BMIDS815
    strPortingData = m_objXmlAssist.GetTagValue(vxmlTableElement, "PORTINGDATAONLY")
    If strPortingData = "1" Then
        Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentData(LoanComponentPortingData)
    Else
        Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentData
    End If  'SR 16/08/2004:BMIDS815 - End
    
    Set IMortgageSubQuoteDO_GetLoanComponentData = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetLoanComponentDataVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IMortgageSubQuoteDO_GetLoanComponentProductParameters(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetData         string containing XML data stream representation of
'                   data retrieved
'   vstrClassDef    xml Class Def to parse vstrXMLRequest against
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetLoanComponentProductParametersVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_GetLoanComponentProductParameters"

    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objIomMPClassDef As IomMPClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    Dim strSQL As String
    
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strMortgageSubQuoteNumber As String
    
    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGESUBQUOTENUMBER")
    
    If Len(strApplicationNumber) = 0 Or _
        Len(strApplicationFactFindNumber) = 0 Or _
        Len(strMortgageSubQuoteNumber) = 0 Then
            
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
        oeMissingPrimaryTag, _
        ",APPLICATIONNUMBER =" & strApplicationNumber & _
        ",APPLICATIONFACTFINDNUMBER =" & strApplicationFactFindNumber & _
        ",MORTGAGESUBQUOTENUMBER =" & strMortgageSubQuoteNumber & "Expected"
    
    End If
    
    Set objIomMPClassDef = New omMPClassDef
        
    ' build the full select statement
    Dim objSQLAssist As SQLAssist
    Set objSQLAssist = New SQLAssist
           
    'MAR1130 GHun Treat NULL Ported Loan as not ported
    strSQL = "SELECT DISTINCT MPP.COMPULSORYBC, MPP.COMPULSORYPP FROM LOANCOMPONENT LC" & _
        " INNER JOIN MORTGAGEPRODUCTPARAMETERS MPP" & _
        " ON LC.MORTGAGEPRODUCTCODE = MPP.MORTGAGEPRODUCTCODE" & _
        " AND LC.STARTDATE = MPP.STARTDATE" & _
        " WHERE LC.APPLICATIONNUMBER = " & objSQLAssist.FormatString(strApplicationNumber) & _
        " AND LC.APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber & _
        " AND LC.MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber & _
        " AND (LC.PORTEDLOAN IS NULL OR LC.PORTEDLOAN = 0)"
           
   'IVW - Disconnecting record sets
    
    Dim cmd As ADODB.Command
    Dim conn As ADODB.Connection
    Dim rstThisRecSet As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    Set conn = New ADODB.Connection
    Set rstThisRecSet = New ADODB.Recordset

    cmd.CommandText = strSQL
    
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.CursorType = adOpenForwardOnly
    rstThisRecSet.LockType = adLockReadOnly
    
    conn.Open objIADOAssist.GetConnStr
    cmd.ActiveConnection = conn

    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
    Set rstThisRecSet.ActiveConnection = Nothing
    
 
    'IVW
                            
    Set xmlClassDefDoc = objIomMPClassDef.LoadMortgageProductParametersData(ValidateCompulsoryProducts)
    
    Dim xmlReturnElement As IXMLDOMElement
    
    'Need to loop round all the records
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlRootNode As IXMLDOMNode
    Dim xmlTempElement As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlTempElement = xmlOut.createElement("PRODUCTPARAMETERSLIST")
    Set xmlRootNode = xmlOut.appendChild(xmlTempElement)
    
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, 226
    End If
    
        ' convert record set to xml
    While Not rstThisRecSet.EOF
        
        Set xmlReturnElement = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlClassDefDoc)
        xmlRootNode.appendChild xmlReturnElement
    
        rstThisRecSet.MoveNext
        
    Wend
    
    Set IMortgageSubQuoteDO_GetLoanComponentProductParameters = xmlOut.documentElement
    
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomMPClassDef = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing

    rstThisRecSet.Close
    conn.Close
    Set conn = Nothing
    Set rstThisRecSet = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

GetLoanComponentProductParametersVbErr:
    
    Set rstThisRecSet = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomMPClassDef = Nothing
    Set objSQLAssist = Nothing
        
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If

    ' Close the connection if still open
    If conn.State = adStateOpen Then
        conn.Close
    End If
    
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

'BMIDS782 GHun
'BMIDS736 GHun No longer used but cannot be totally removed due to OOSS
'Private Function IMortgageSubQuoteDO_GetLoanComponentResultsData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
'' header ----------------------------------------------------------------------------------
'' description:
''   Get the data for a single instance of the persistant data associated with
''   this data object
'' pass:
''   vxmlxmlTableElement  xml element containing the request
'' return:                xml node containing retrieved data
''------------------------------------------------------------------------------------------
'On Error GoTo GetLoanComponentResultsDataVbErr
'
'    Const strFunctionName As String = "GetLoanComponentResultsData"
'
'    Dim objIDOAssist As IDOAssist
'    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
'    Dim objIomCMClassDef As IomCMClassDef
'
'    If m_objContext Is Nothing Then
'        Set objIDOAssist = New DOAssist
'    Else
'        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
'    End If
'
'    Set objIomCMClassDef = New omCMClassDef
'    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentResultsData
'    Set IMortgageSubQuoteDO_GetLoanComponentResultsData = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
'
'    Set objIomCMClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
'    Set objIDOAssist = Nothing
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
'
'    Exit Function
'
'GetLoanComponentResultsDataVbErr:
'
'    Set objIomCMClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
'    Set objIDOAssist = Nothing
'
'    m_objErrAssist.AddToErrSource (strFunctionName)
'
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
'    End If
'
'    If Not m_objContext Is Nothing Then
'        If m_objErrAssist.IsSystemError Then
'            m_objContext.SetAbort
'        Else
'            m_objContext.SetComplete
'        End If
'    End If
'
'    '   re-raise error for business object to interpret as appropriate
'    Err.Raise Err.number, Err.Source, Err.Description
'End Function
'BMIDS736 End
'BMIDS782 End

Private Function IMortgageSubQuoteDO_GetLowestLoanComponent(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetData         string containing XML data stream representation of
'                   data retrieved
'   vstrClassDef    xml Class Def to parse vstrXMLRequest against
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetLowestLoanComponentVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_GetLowestLoanComponent"

    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objIomCMClassDef As IomCMClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    Dim strSQL As String
    
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    Dim strMortgageSubQuoteNumber As String
    
    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGESUBQUOTENUMBER")
    
    If Len(strApplicationNumber) = 0 Or _
        Len(strApplicationFactFindNumber) = 0 Or _
        Len(strMortgageSubQuoteNumber) = 0 Then
            
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
        oeMissingPrimaryTag, _
        ",APPLICATIONNUMBER =" & strApplicationNumber & _
        ",APPLICATIONFACTFINDNUMBER =" & strApplicationFactFindNumber & _
        ",MORTGAGESUBQUOTENUMBER =" & strMortgageSubQuoteNumber & "Expected"
    
    End If
    
    Set objIomCMClassDef = New omCMClassDef
        
    ' build the full select statement
    Dim objSQLAssist As SQLAssist
    Set objSQLAssist = New SQLAssist
           
    strSQL = "SELECT * FROM LOANCOMPONENT WHERE " & _
                "APPLICATIONNUMBER = " & objSQLAssist.FormatString(strApplicationNumber) & _
                " AND APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber & _
                " AND MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber & _
                " AND LOANCOMPONENTSEQUENCENUMBER = (SELECT MIN(LOANCOMPONENTSEQUENCENUMBER) " & _
                " FROM LOANCOMPONENT WHERE  " & _
                " APPLICATIONNUMBER = " & objSQLAssist.FormatString(strApplicationNumber) & _
                " AND APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber & _
                " AND MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber & ")"

    'IVW - Disconnecting record sets
    
    Dim cmd As ADODB.Command
    Dim conn As ADODB.Connection
    Dim rstThisRecSet As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    Set conn = New ADODB.Connection
    Set rstThisRecSet = New ADODB.Recordset

    cmd.CommandText = strSQL
    
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.CursorType = adOpenForwardOnly
    rstThisRecSet.LockType = adLockReadOnly
    
    conn.Open objIADOAssist.GetConnStr
    cmd.ActiveConnection = conn
            
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
    Set rstThisRecSet.ActiveConnection = Nothing
    
 
    'IVW
            
    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentData()
    
    Dim xmlReturnElement As IXMLDOMElement
    
    If Not rstThisRecSet.EOF Then
        ' convert record set to xml
        
        Set xmlReturnElement = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlClassDefDoc)
    Else
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
    
    Set IMortgageSubQuoteDO_GetLowestLoanComponent = xmlReturnElement
    
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomCMClassDef = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing

    rstThisRecSet.Close
    Set rstThisRecSet = Nothing

    conn.Close
    Set conn = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

GetLowestLoanComponentVbErr:
    
    Set rstThisRecSet = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomCMClassDef = Nothing
    
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If

    ' Close the connection if still open
    If conn.State = adStateOpen Then
        conn.Close
    End If
    
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IMortgageSubQuoteDO_GetMortgageCalculatorDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageCalculatorDetailsVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_GetMortgageCalculatorDetails"
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageCalculatorDetailsData
    Set IMortgageSubQuoteDO_GetMortgageCalculatorDetails = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetMortgageCalculatorDetailsVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IMortgageSubQuoteDO_GetNonPanelMortgageProductData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetNonPanelMortgageProductDataVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_GetNonPanelMortgageProductData"
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadNonPanelMortgageProductData
    Set IMortgageSubQuoteDO_GetNonPanelMortgageProductData = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetNonPanelMortgageProductDataVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IMortgageSubQuoteDO_GetPortedLender(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetData         string containing XML data stream representation of
'                   data retrieved
'   vstrClassDef    xml Class Def to parse vstrXMLRequest against
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetPortedLenderVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_GetPortedLender"

    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objIomCMClassDef As IomCMClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    Dim strSQL As String
    
    Dim strMortgageLoanGuid As String
    
    strMortgageLoanGuid = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGELOANGUID")

    If Len(strMortgageLoanGuid) = 0 Then
            
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
        oeMissingPrimaryTag, "MORTGAGELOANGUID tag empty or not found"
    
    End If
    
    Set objIomCMClassDef = New omCMClassDef
        
    ' build the full select statement
    Dim objSQLAssist As SQLAssist
    Set objSQLAssist = New SQLAssist
    
    strSQL = "SELECT MORTGAGELENDERDIRECTORY.ORGANISATIONID FROM " & _
                " MORTGAGELOAN , ACCOUNT , MORTGAGELENDERDIRECTORY " & _
                " WHERE MORTGAGELOAN.MORTGAGELOANGUID = " & _
                objSQLAssist.FormatGuid(strMortgageLoanGuid) & _
                " AND MORTGAGELOAN.ACCOUNTGUID = ACCOUNT.ACCOUNTGUID " & _
                " AND ACCOUNT.DIRECTORYGUID = MORTGAGELENDERDIRECTORY.DIRECTORYGUID "
                                
    'IVW - Disconnecting record sets
    
    Dim cmd As ADODB.Command
    Dim conn As ADODB.Connection
    Dim rstThisRecSet As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    Set conn = New ADODB.Connection
    Set rstThisRecSet = New ADODB.Recordset

    cmd.CommandText = strSQL
    
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.CursorType = adOpenForwardOnly
    rstThisRecSet.LockType = adLockReadOnly
    
    conn.Open objIADOAssist.GetConnStr
    cmd.ActiveConnection = conn
            
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
    Set rstThisRecSet.ActiveConnection = Nothing
    
 
    'IVW
            
    Set xmlClassDefDoc = objIomCMClassDef.LoadPortedLenderData()
    
    Dim xmlReturnElement As IXMLDOMElement
    
    If Not rstThisRecSet.EOF Then
        ' convert record set to xml
        
        Set xmlReturnElement = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlClassDefDoc)
    Else
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
        
    Set IMortgageSubQuoteDO_GetPortedLender = xmlReturnElement
    
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomCMClassDef = Nothing
    Set objSQLAssist = Nothing
    
    Set cmd = Nothing
    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
    conn.Close
    Set conn = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

GetPortedLenderVbErr:
    
    Set rstThisRecSet = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomCMClassDef = Nothing
    Set objSQLAssist = Nothing
    
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If

    ' Close the connection if still open
    If conn.State = adStateOpen Then
        conn.Close
    End If
    
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Sub IMortgageSubQuoteDO_ReserveMortgageProduct(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
'   MAR1703 use stored procedure USP_RESERVEMORTGAGEPRODUCT
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo ReserveMortgageProductVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_ReserveMortgageProduct"
   
    Dim objIADOAssist                   As IADOAssist
    Dim adoConnection   As ADODB.Connection
    Dim adoCommand      As ADODB.Command
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")

    Set adoConnection = New ADODB.Connection
                                    
    'Create the parameters and call the stored proc
    Set adoCommand = New ADODB.Command
    With adoCommand
        .CommandType = adCmdStoredProc
        .CommandText = "USP_RESERVEMORTGAGEPRODUCT"
        'Application Number
        .Parameters.Append .CreateParameter("@p_AppNumber", adBSTR, adParamInput, 12, strApplicationNumber)
        'Application FactFind Number
        .Parameters.Append .CreateParameter("@p_AppFFNumber", adInteger, adParamInput, , strApplicationFactFindNumber)
    End With

    adoConnection.Open objIADOAssist.GetConnStr
    adoCommand.ActiveConnection = adoConnection
    adoCommand.Execute
    
    'Close the database connection
    Set adoCommand.ActiveConnection = Nothing
    adoConnection.Close

    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

ReserveMortgageProductVbErr:
    
    ' Close the connection if still open
    If adoConnection.State = adStateOpen Then
        adoConnection.Close
    End If
    
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set objIADOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description
End Sub
Private Sub IMortgageSubQuoteDO_ReserveMortgageProduct_OLD(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
' MAR1703 Method changed to use stored procedure. This mehtod no longer called
'------------------------------------------------------------------------------------------
On Error GoTo ReserveMortgageProductVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_ReserveMortgageProduct"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomMPClassDef As IomMPClassDef
    
    Dim objIDOAssist As IDOAssist
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomMPClassDef = New omMPClassDef
    Set xmlClassDefDoc = objIomMPClassDef.LoadReservedProductData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIomMPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

ReserveMortgageProductVbErr:
    
    Set objIomMPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IMortgageSubQuoteDO_ResetMortgageSubQuote(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo ResetMortgageSubQuoteVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_ResetMortgageSubQuote"

    Dim xmlTempNode As IXMLDOMElement
           
    Set xmlTempNode = vxmlTableElement.cloneNode(True)
    
    'BMIDS736 GHun
    'm_objXmlAssist.ChangeNodeName xmlTempNode, "MORTGAGESUBQUOTE", "LOANCOMPONENTRESULTS"
    'IMortgageSubQuoteDO_DeleteAllLoanComponentResults xmlTempNode
    m_objXmlAssist.ChangeNodeName xmlTempNode, "MORTGAGESUBQUOTE", "LOANCOMPONENTBALANCESCHEDULE"
    IMortgageSubQuoteDO_DeleteAllLoanComponentBalanceSchedules xmlTempNode
    m_objXmlAssist.ChangeNodeName xmlTempNode, "LOANCOMPONENTBALANCESCHEDULE", "LOANCOMPONENTPAYMENTSCHEDULE"
    IMortgageSubQuoteDO_DeleteAllLoanComponentPaymentSchedules xmlTempNode
    'IMortgageSubQuoteDO_DeleteAllMIRASDetails (vxmlTableElement)
    m_objXmlAssist.ChangeNodeName xmlTempNode, "LOANCOMPONENTPAYMENTSCHEDULE", "MORTGAGEINCENTIVE"
    'BMIDS736 End
    
    IMortgageSubQuoteDO_DeleteAllMortgageIncentive xmlTempNode
    
    m_objXmlAssist.ChangeNodeName xmlTempNode, "MORTGAGEINCENTIVE", "NONPANELMORTGAGEPRODUCT"
    
    IMortgageSubQuoteDO_DeleteAllNonPanelMortgageProduct xmlTempNode
    
    m_objXmlAssist.ChangeNodeName xmlTempNode, "NONPANELMORTGAGEPRODUCT", "LOANCOMPONENT"
    
    IMortgageSubQuoteDO_DeleteAllLoanComponent xmlTempNode
     
    Exit Sub

ResetMortgageSubQuoteVbErr:
    
    If m_objErrAssist.GetOmigaErrorNumber(Err.number) = oeNoRowsAffected Or _
        m_objErrAssist.GetOmigaErrorNumber(Err.number) = oeNoRowsAffectedByDeleteAll Then
        Resume Next
    
    End If

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IMortgageSubQuoteDO_SaveIncentiveAgainstLoanComponent(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo SaveIncentiveAgainstLoanComponentVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_SaveIncentiveAgainstLoanComponent"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    Dim objIDOAssist As IDOAssist
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageIncentiveData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

SaveIncentiveAgainstLoanComponentVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

'AS     23/03/2006  MAR1495 PERFORMANCE - Apply BBG1984 Amend IMortgageSubQuoteDO_TotalIncentivesForAllComponents()
Private Function IMortgageSubQuoteDO_TotalIncentivesForAllComponents(ByVal vxmlTableElement As IXMLDOMElement) As Double
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetData         string containing XML data stream representation of
'                   data retrieved
'   vstrClassDef    xml Class Def to parse vstrXMLRequest against
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo TotalIncentivesForAllComponentsVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_TotalIncentivesForAllComponents"

    Dim objIADOAssist                   As IADOAssist
    Dim adoConnection   As ADODB.Connection
    Dim adoCommand      As ADODB.Command
    Dim strApplicationNumber            As String
    Dim strApplicationFactFindNumber    As String
    '*=[MC]BMIDS763
    Dim strApplicationDate              As String
    '*section end
    Dim strMortgageSubQuoteNumber       As String
    Dim strDetails                      As String
    Dim dblTotalIncentives              As Double
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strApplicationDate = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONDATE")
    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGESUBQUOTENUMBER")
                                                                                                          
    If Len(strApplicationNumber) = 0 Or Len(strApplicationFactFindNumber) = 0 Or _
        Len(strMortgageSubQuoteNumber) = 0 Then
                
        strDetails = "APPLICATIONNUMBER = " & strApplicationNumber
        strDetails = strDetails & "; APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber
        strDetails = strDetails & "; MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber
        
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                strDetails
    End If
                                    
    Set adoConnection = New ADODB.Connection
                                    
    'Create the parameters and call the stored proc
    Set adoCommand = New ADODB.Command
    With adoCommand
        .CommandType = adCmdStoredProc
        .CommandText = "USP_TOTALINCENTIVEFORALLCOMPONENTS"
        'Application Number
        .Parameters.Append .CreateParameter("@pApplicationNumber", adBSTR, adParamInput, 12, strApplicationNumber)
        'Application FactFind Number
        .Parameters.Append .CreateParameter("@pApplicationFactFindNumber", adInteger, adParamInput, , strApplicationFactFindNumber)
        'MortgageSubQuote Number
        .Parameters.Append .CreateParameter("@pMortgageSubQuoteNumber", adInteger, adParamInput, , strMortgageSubQuoteNumber)
        'Total Incentive Amount
        .Parameters.Append .CreateParameter("@pTotalIncentiveAmount", adInteger, adParamOutput)
    End With

    adoConnection.Open objIADOAssist.GetConnStr
    adoCommand.ActiveConnection = adoConnection
    adoCommand.Execute
    
    dblTotalIncentives = 0
    
    If Not IsNull(adoCommand.Parameters("@pTotalIncentiveAmount").Value) Then
        dblTotalIncentives = CDbl(adoCommand.Parameters("@pTotalIncentiveAmount").Value)
    End If
    
    'Close the database connection
    Set adoCommand.ActiveConnection = Nothing
    adoConnection.Close

    IMortgageSubQuoteDO_TotalIncentivesForAllComponents = dblTotalIncentives
    
    Set objIADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

TotalIncentivesForAllComponentsVbErr:
    
    ' Close the connection if still open
    If adoConnection.State = adStateOpen Then
        adoConnection.Close
    End If
    
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set objIADOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Sub IMortgageSubQuoteDO_Update(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_Update"
        
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    Dim objIDOAssist As IDOAssist

    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageSubQuoteData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

UpdateVbErr:
        
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IMortgageSubQuoteDO_UpdateLoanComponent(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLoanComponentVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_UpdateLoanComponent"
        
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    Dim objIDOAssist As IDOAssist
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

UpdateLoanComponentVbErr:
        
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

'BMIDS782 GHun
'BMIDS736 GHun No longer used but cannot be totally removed due to OOSS
'Private Sub IMortgageSubQuoteDO_UpdateLoanComponentResults(ByVal vxmlTableElement As IXMLDOMElement)
'' header ----------------------------------------------------------------------------------
'' description:
''   Update a single instance of the persistant data associated with this
''   data object
'' pass:
''   vxmlxmlTableElement  xml element containing the request
'' return:                n/a
''------------------------------------------------------------------------------------------
'On Error GoTo UpdateLoanComponentResultsVbErr
'
'    Const strFunctionName As String = "UpdateLoanComponentResults"
'
'    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
'    Dim objIomCMClassDef As IomCMClassDef
'    Dim objIDOAssist As IDOAssist
'
'    If m_objContext Is Nothing Then
'        Set objIDOAssist = New DOAssist
'    Else
'        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
'    End If
'
'    Set objIomCMClassDef = New omCMClassDef
'    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentResultsData
'    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
'
'    Set objIomCMClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
'    Set objIDOAssist = Nothing
'
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
'
'    Exit Sub
'
'UpdateLoanComponentResultsVbErr:
'
'    Set objIomCMClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
'    Set objIDOAssist = Nothing
'
'    m_objErrAssist.AddToErrSource (strFunctionName)
'
'    If m_objErrAssist.IsSystemError = True Then
'        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
'    End If
'
'    If Not m_objContext Is Nothing Then
'        If m_objErrAssist.IsSystemError Then
'            m_objContext.SetAbort
'        Else
'            m_objContext.SetComplete
'        End If
'    End If
'
'    '   re-raise error for business object to interpret as appropriate
'    Err.Raise Err.number, Err.Source, Err.Description
'End Sub
'BMIDS736 End
'BMIDS782 End

Private Function IMortgageSubQuoteDO_GetDataForOneOffCostsDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetData         string containing XML data stream representation of
'                   data retrieved
'   vstrClassDef    xml Class Def to parse vstrXMLRequest against
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetDataForOneOffCostsDetailsVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_GetDataForOneOffCostsDetails"

    Dim objIADOAssist                       As IADOAssist
    Dim objIDOAssist                        As IDOAssist
    Dim objIomCMClassDef                    As IomCMClassDef
    Dim objIomOrgClassDef                   As IOmOrgClassDef
    Dim xmlOutDoc                           As FreeThreadedDOMDocument40
    Dim xmlRootNode                         As IXMLDOMNode
    Dim objXMLTempNode                      As IXMLDOMNode
    Dim strSQL                              As String
    Dim objSQLAssist                        As SQLAssist
    Dim strApplicationNumber                As String
    Dim strApplicationFactFindNumber        As String
    Dim strMortgageSubQuoteNumber           As String
    Dim strDetails                          As String
    Dim cmd                                 As ADODB.Command
    Dim conn                                As ADODB.Connection
    Dim rstThisRecSet                       As ADODB.Recordset
    Dim xmlMortgageSubQuoteClassDefDoc      As FreeThreadedDOMDocument40
    Dim xmlMortgageOneOffCostClassDefDoc    As FreeThreadedDOMDocument40
    Dim xmlMortgagLenderClassDefDoc         As FreeThreadedDOMDocument40
    Dim objXMLMortgageSubQuoteNode          As IXMLDOMNode
    Dim objXMLMortgageLenderNode            As IXMLDOMNode
    Dim objXMLMortgageOneOffCostListNode    As IXMLDOMNode
    Dim objXMLMortgageOnOffCostNode         As IXMLDOMNode
    
    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
    Set objXMLTempNode = xmlOutDoc.createElement("ONEOFFCOST")
    Set xmlRootNode = xmlOutDoc.appendChild(objXMLTempNode)
                
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                    
    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, _
                                                  "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, _
                                                     "MORTGAGESUBQUOTENUMBER")
                                                     
    Set objSQLAssist = New SQLAssist
                                                     
    If Len(strApplicationNumber) = 0 Or Len(strApplicationFactFindNumber) = 0 Or _
        Len(strMortgageSubQuoteNumber) = 0 Then
        
        strDetails = "APPLICATIONNUMBER = " & objSQLAssist.FormatString(strApplicationNumber)
        strDetails = strDetails & "; APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber
        strDetails = strDetails & "; MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNumber
        
        m_objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                strDetails
    End If
                
    'AS 22/03/2006 MAR1461 Start
    Set conn = New ADODB.Connection
    Set rstThisRecSet = New ADODB.Recordset
               
    'This change prevents SQL Server lock timeouts under load
    Set cmd = New ADODB.Command
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = " USP_GetCalculateMortgageCostsData"
        .Parameters.Append .CreateParameter("@p_APPLICATIONNUMBER", adVarChar, adParamInput, Len(strApplicationNumber), strApplicationNumber)
        .Parameters.Append .CreateParameter("@p_APPLICATIONFACTFINDNUMBER", adInteger, adParamInput, , strApplicationFactFindNumber)
        .Parameters.Append .CreateParameter("@p_MORTGAGESUBQUOTENUMBER", adInteger, adParamInput, , strMortgageSubQuoteNumber)
    End With
               
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.CursorType = adOpenForwardOnly
    rstThisRecSet.LockType = adLockReadOnly
    
    conn.Open objIADOAssist.GetConnStr
    cmd.ActiveConnection = conn
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
    Set rstThisRecSet.ActiveConnection = Nothing
    'AS 22/03/2006 MAR1461 End
                    
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
        
    Set objIomCMClassDef = New omCMClassDef
    Set objIomOrgClassDef = New omOrgClassDef
    
    Set xmlMortgageSubQuoteClassDefDoc = objIomCMClassDef.LoadMortgageSubQuoteData
    Set xmlMortgageOneOffCostClassDefDoc = objIomCMClassDef.LoadMortgageOneOffCostData
    Set xmlMortgagLenderClassDefDoc = objIomOrgClassDef.LoadMortgageLenderData
    
    ' loop through the record set
    While Not rstThisRecSet.EOF
                    
        If objXMLMortgageSubQuoteNode Is Nothing Then
            'only want one mortgagesubquote record
            Set objXMLMortgageSubQuoteNode = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlMortgageSubQuoteClassDefDoc)
            'append the Mortgage sub quote data to the root node
            xmlRootNode.appendChild objXMLMortgageSubQuoteNode

        End If
        
        If objXMLMortgageLenderNode Is Nothing Then
            'only want one mortgagelender record
            Set objXMLMortgageLenderNode = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlMortgagLenderClassDefDoc)
            'append the Mortgage lender data to the root node
            xmlRootNode.appendChild objXMLMortgageLenderNode

        End If
                
        If objXMLMortgageOneOffCostListNode Is Nothing Then
            'first time in so create a list element
            Set objXMLTempNode = xmlOutDoc.createElement("MORTGAGEONEOFFCOSTLIST")
            'and set the one off cost node to it
            Set objXMLMortgageOneOffCostListNode = xmlRootNode.appendChild(objXMLTempNode)
        
        End If
        
        Set objXMLMortgageOnOffCostNode = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlMortgageOneOffCostClassDefDoc)
        'append the Mortgage lender data to the root node
        objXMLMortgageOneOffCostListNode.appendChild objXMLMortgageOnOffCostNode
                    
        rstThisRecSet.MoveNext
            
    Wend
                    
    Set IMortgageSubQuoteDO_GetDataForOneOffCostsDetails = xmlOutDoc.documentElement
    
    Set cmd = Nothing
    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
    conn.Close
    Set conn = Nothing
    
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomCMClassDef = Nothing
    Set objIomOrgClassDef = Nothing
    Set objSQLAssist = Nothing
    Set xmlOutDoc = Nothing
    Set xmlRootNode = Nothing
    Set objXMLTempNode = Nothing
    Set xmlMortgageSubQuoteClassDefDoc = Nothing
    Set xmlMortgageOneOffCostClassDefDoc = Nothing
    Set xmlMortgagLenderClassDefDoc = Nothing
    Set objXMLMortgageSubQuoteNode = Nothing
    Set objXMLMortgageLenderNode = Nothing
    Set objXMLMortgageOneOffCostListNode = Nothing
    Set objXMLMortgageOnOffCostNode = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

GetDataForOneOffCostsDetailsVbErr:
    
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIomCMClassDef = Nothing
    Set objIomOrgClassDef = Nothing
    Set objSQLAssist = Nothing
    Set xmlOutDoc = Nothing
    Set xmlRootNode = Nothing
    Set objXMLTempNode = Nothing
    Set xmlMortgageSubQuoteClassDefDoc = Nothing
    Set xmlMortgageOneOffCostClassDefDoc = Nothing
    Set xmlMortgagLenderClassDefDoc = Nothing
    Set objXMLMortgageSubQuoteNode = Nothing
    Set objXMLMortgageLenderNode = Nothing
    Set objXMLMortgageOneOffCostListNode = Nothing
    Set objXMLMortgageOnOffCostNode = Nothing
    
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If

    ' Close the connection if still open
    If conn.State = adStateOpen Then
        conn.Close
    End If
    
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Sub IMortgageSubQuoteDO_UpdateTotalLoanAmount(ByVal vXMLRequest As IXMLDOMElement, _
            ByVal lngLoanAmount As Long, Optional ByVal blnNonPanel As Boolean = True)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateTotalLoanAmountVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_UpdateTotalLoanAmount"
        
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist As IADOAssist

    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageSubQuoteData

    Dim strSQL As String        ' SQL insert statement
    Dim strAndSeparated  As String
    
    Dim strCommaSeparated As String
    Dim strTable As String

    objIDOAssist.BuildSQLString vXMLRequest, xmlClassDefDoc, strTable, sftAndSeparated, , , strAndSeparated
    
    ' PSC 07/07/00 SYS1126
    strCommaSeparated = "TOTALLOANAMOUNT = TOTALLOANAMOUNT - " & lngLoanAmount
    
    'now add the other values
    If Not blnNonPanel Then
    
        strCommaSeparated = strCommaSeparated & _
                ",  NONPANELLENDERNAME = "" , NONPANELLENDERSELECTD = 0" 'FALSE

    
    End If
        
    strSQL = "UPDATE " & strTable & " SET " & strCommaSeparated & " where " & strAndSeparated

    objIADOAssist.executeSQLCommand strSQL
    
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

UpdateTotalLoanAmountVbErr:
        
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Function IMortgageSubQuoteDO_GetAmountRequested(ByVal vxmlTableElement As IXMLDOMElement) As Long
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetAmountRequestedVbErr

    Const strFunctionName As String = "IMortgageSubQuoteDO_GetAmountRequested"
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    Dim objConvertAssist As ConvertAssist
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadMortgageSubQuoteData(MortgageSubQuoteAmountRequested)
    
    Set objConvertAssist = New ConvertAssist
    
    IMortgageSubQuoteDO_GetAmountRequested = _
                            objConvertAssist.CSafeLng( _
                            m_objXmlAssist.GetTagValue( _
                            objIDOAssist.GetData( _
                            vxmlTableElement, xmlClassDefDoc), "AMOUNTREQUESTED"))
    

    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objConvertAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
       
    Exit Function

GetAmountRequestedVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objConvertAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

'BMIDS736 GHun
Private Sub IMortgageSubQuoteDO_CreateLoanComponentBalanceSchedule(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateLoanComponentBalanceScheduleVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteDO_CreateLoanComponentBalanceSchedule"
    Dim xmlClassDefDoc    As FreeThreadedDOMDocument40
    Dim objIomCMClassDef  As IomCMClassDef
    Dim objIDOAssist      As IDOAssist
                    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentBalanceScheduleData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub

CreateLoanComponentBalanceScheduleVbErr:
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description
End Sub
'BMIDS736 End

'BMIDS736 GHun
Private Sub IMortgageSubQuoteDO_CreateLoanComponentPaymentSchedule(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateLoanComponentPaymentScheduleVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteDO_CreateLoanComponentPaymentSchedule"
    Dim xmlClassDefDoc    As FreeThreadedDOMDocument40
    Dim objIomCMClassDef  As IomCMClassDef
    Dim objIDOAssist      As IDOAssist
                    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentPaymentScheduleData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub

CreateLoanComponentPaymentScheduleVbErr:
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description
End Sub
'BMIDS736 End

'BMIDS736 GHun
Private Sub IMortgageSubQuoteDO_DeleteAllLoanComponentBalanceSchedules(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vxmlTableElement  xml element containing the request
' return:             n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllLoanComponentBalanceSchedulesVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteDO_DeleteAllLoanComponentBalanceSchedules"
           
    Dim objIDOAssist                    As IDOAssist
    Dim xmlClassDefDoc                  As FreeThreadedDOMDocument40
    Dim objIomCMClassDef                As IomCMClassDef
    Dim strApplicationNumber            As String
    Dim strApplicationFactFindNumber    As String
    Dim strMortgageSubQuoteNumber       As String
    'Dim strLoanComponentSeqNumber       As String
                            
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGESUBQUOTENUMBER")
    'GHun LoanComponentSeqNumber is not mandatory when called from ResetMortgageSubQuote
    'strLoanComponentSeqNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "LOANCOMPONENTSEQUENCENUMBER")
    
    If Len(strApplicationNumber) = 0 Or _
        Len(strApplicationFactFindNumber) = 0 Or _
        Len(strMortgageSubQuoteNumber) = 0 Then
            
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
        oeMissingPrimaryTag, _
        ",APPLICATIONNUMBER=" & strApplicationNumber & _
        ",APPLICATIONFACTFINDNUMBER=" & strApplicationFactFindNumber & _
        ",MORTGAGESUBQUOTENUMBER=" & strMortgageSubQuoteNumber & " expected"
    End If
                                
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentBalanceScheduleData
    objIDOAssist.DeleteAll vxmlTableElement, xmlClassDefDoc
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
    
DeleteAllLoanComponentBalanceSchedulesVbErr:
    If m_objErrAssist.GetOmigaErrorNumber(Err.number) = oeNoRowsAffectedByDeleteAll Then
        Resume Next
    End If
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description
End Sub
'BMIDS736 End

'BMIDS736 GHun
Private Sub IMortgageSubQuoteDO_DeleteAllLoanComponentPaymentSchedules(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vxmlTableElement  xml element containing the request
' return:             n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllLoanComponentPaymentSchedulesVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteDO_DeleteAllLoanComponentPaymentSchedules"
           
    Dim objIDOAssist                    As IDOAssist
    Dim xmlClassDefDoc                  As FreeThreadedDOMDocument40
    Dim objIomCMClassDef                As IomCMClassDef
    Dim strApplicationNumber            As String
    Dim strApplicationFactFindNumber    As String
    Dim strMortgageSubQuoteNumber       As String
    'Dim strLoanComponentSeqNumber       As String
                            
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGESUBQUOTENUMBER")
    'GHun LoanComponentSeqNumber is not mandatory when called from ResetMortgageSubQuote
    'strLoanComponentSeqNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "LOANCOMPONENTSEQUENCENUMBER")
    
    If Len(strApplicationNumber) = 0 Or _
        Len(strApplicationFactFindNumber) = 0 Or _
        Len(strMortgageSubQuoteNumber) = 0 Then
            
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
        oeMissingPrimaryTag, _
        ",APPLICATIONNUMBER=" & strApplicationNumber & _
        ",APPLICATIONFACTFINDNUMBER=" & strApplicationFactFindNumber & _
        ",MORTGAGESUBQUOTENUMBER=" & strMortgageSubQuoteNumber & " expected"
    End If
                                
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentPaymentScheduleData
    objIDOAssist.DeleteAll vxmlTableElement, xmlClassDefDoc
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
    
DeleteAllLoanComponentPaymentSchedulesVbErr:
    If m_objErrAssist.GetOmigaErrorNumber(Err.number) = oeNoRowsAffectedByDeleteAll Then
        Resume Next
    End If
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description
End Sub
'BMIDS736 End

'BMIDS736 GHun
Private Sub IMortgageSubQuoteDO_DeleteAllLoanComponentRedemptionFees(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vxmlTableElement  xml element containing the request
' return:             n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllLoanComponentRedemptionFeesVbErr
    
    Const strFunctionName As String = "IMortgageSubQuoteDO_DeleteAllLoanComponentRedemptionFees"
           
    Dim objIDOAssist                    As IDOAssist
    Dim xmlClassDefDoc                  As FreeThreadedDOMDocument40
    Dim objIomCMClassDef                As IomCMClassDef
    Dim strApplicationNumber            As String
    Dim strApplicationFactFindNumber    As String
    Dim strMortgageSubQuoteNumber       As String
    Dim strLoanComponentSeqNumber       As String
                            
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    strApplicationNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "MORTGAGESUBQUOTENUMBER")
    strLoanComponentSeqNumber = m_objXmlAssist.GetTagValue(vxmlTableElement, "LOANCOMPONENTSEQUENCENUMBER")
    
    If Len(strApplicationNumber) = 0 Or _
        Len(strApplicationFactFindNumber) = 0 Or _
        Len(strMortgageSubQuoteNumber) = 0 Or _
        Len(strLoanComponentSeqNumber) = 0 Then
            
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
        oeMissingPrimaryTag, _
        ",APPLICATIONNUMBER=" & strApplicationNumber & _
        ",APPLICATIONFACTFINDNUMBER=" & strApplicationFactFindNumber & _
        ",MORTGAGESUBQUOTENUMBER=" & strMortgageSubQuoteNumber & _
        ",LOANCOMPONENTSEQUENCENUMBER=" & strLoanComponentSeqNumber & "Expected"
    End If
                                
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadLoanComponentRedemptionFeeData
    objIDOAssist.DeleteAll vxmlTableElement, xmlClassDefDoc
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
    
DeleteAllLoanComponentRedemptionFeesVbErr:
    If m_objErrAssist.GetOmigaErrorNumber(Err.number) = oeNoRowsAffectedByDeleteAll Then
        Resume Next
    End If
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description
End Sub
'BMIDS736 End

'BMIDS782 GHun
Private Sub IMortgageSubQuoteDO_SaveIllustrationResults(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Calls a stored proc to save LoanComponent, LoanComponentPaymentSchedule, LoanComponentBalanceSchedule
'   and LoanComponentRedemptionFees all in one go.
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
    Const cstrFunctionName      As String = "IMortgageSubQuoteDO_SaveIllustrationResults"
    Dim adoConnection           As ADODB.Connection
    Dim adoCommand              As ADODB.Command
    Dim intAdditionalBorrowing  As Integer  'EP2_954 GHun
    
On Error GoTo SaveIllustrationResultsVbErr
    
    'EP2_954 GHun
    If Not vxmlTableElement.selectSingleNode("MORTGAGELOANPAYMENTSCHEDULE") Is Nothing Then
        intAdditionalBorrowing = 1
    Else
        intAdditionalBorrowing = 0
    End If
    'EP2_954 End
    
    Set adoConnection = New ADODB.Connection
    With adoConnection
        .ConnectionString = adoGetDbConnectString
        .Open
    End With
                      
    'Create the parameters and call the stored proc
    Set adoCommand = New ADODB.Command
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandType = adCmdStoredProc
        .CommandText = "USP_SAVEILLUSTRATIONRESULTS"
        .Parameters.Append .CreateParameter("@p_XMLDoc", adChar, adParamInput, Len(vxmlTableElement.xml), vxmlTableElement.xml)
        .Parameters.Append .CreateParameter("@p_IsAdditionalBorrowing", adBoolean, adParamInput, , intAdditionalBorrowing)  'EP2_954 GHun
        .Execute , , adExecuteNoRecords
    End With
    
    'Close the database connection
    adoConnection.Close
    Set adoCommand.ActiveConnection = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

SaveIllustrationResultsVbErr:
    
    Set adoConnection = Nothing
    Set adoCommand = Nothing
        
    m_objErrAssist.AddToErrSource (cstrFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description
End Sub
'BMIDS782 End

'BMIDS622 Added to conform to standards
Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
End Sub
'BMIDS622 End
