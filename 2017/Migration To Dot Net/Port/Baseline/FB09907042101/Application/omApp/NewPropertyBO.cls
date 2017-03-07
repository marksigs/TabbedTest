VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "NewPropertyBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      NewPropertyBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' RF    07/01/00    Created
' RF    15/02/00    Added method GetValuationTypeAndLocation.
' IW    03/05/00    Added Validation for DATELEASESTARTED
' MC    11/05/00    Vendor Details - Remove Directory processing
' BG    26/05/00    Added Method SaveNewPropertyAndAddress
' MC    30/05/00    Added GetPropertyUnknownDetails method
' SR    02/06/00    Added method 'SaveNewPropertyAndDeposit'
' IVW   06/06/00    Added GetNewPropertyAndAddressDetails
' BG    12/06/00    Added GetPropertyUnknownDetails
' SR    13/06/00    New methods 'GetFullNewPropertyDetails' and 'SaveNewProperty'
' SR    15/06/00    Added GetFullNewPropertyDetails
' MC    14/07/00    SYS1176 - Added SaveNewPropertyAddress method
' LD    04/09/00    Ammendments for NT4/W2000 dual-compile
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------
'BMIDS History:
'
' Prog  Date        Description
' GD    04/07/02    BMIDS00165 - Added CalcExcessRentalIncome
' GD    19/07/02    BMIDS00217 - Amended CalcExcessRentalIncome to raise an error if either of the 2 required global parameters are empty
' GD                             Altered 'If' logic in CalExcessRentalIncome to cater for Non - Buy To Let.
' GHun  11/11/2002  BMIDS00755 - Amended CalcExcessRentalIncome to use a different method for getting
'                                mortgage loan records.
' BS    13/05/2003  BM0310 Amended CalcExcessRentalIncome to return all ValidationTypes not just first one
'                               and allow for BTL that is not a Further Advance or Secured Personal Loan
' GHun  02/07/2003  BMIDS591 CC040 Amended CalcExcessRentalIncome
' JD    22/06/04    BMIDS765 CC076 CalcExcessRentalIncome now calls method to get the buyToLetRate
' JD    28/06/04    BMIDS765 CC076 Removed deactivation of m_obj* objects on exit from GetRentalIncomeRates
' GHun  23/09/2004  BMIDS887 Added SaveNewPropertyAndUpdateFactFind
'------------------------------------------------------------------------------------------
'Epsom History
'
' Prog  Date        Description
' GHun  26/10/2006  EP2_9 Changed CalcExcessRentalIncome to comment out BM specific code
'------------------------------------------------------------------------------------------

Option Explicit

Implements INewPropertyBO
Implements ObjectControl
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required

Private Const cstrROOT_NODE_NAME As String = "NEWPROPERTY"

 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

Public Function DeleteNewPropertyAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   business object.
' pass:
'   vstrXmlRequest  xml Request data stream containing data to be persisted
' return:
'   xml Response data stream containing results of operation
'   either: TYPE="SUCCESS"
'   or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteNewPropertyAddressVbErr

    Const strFunctionName As String = "DeleteNewPropertyAddress"
   
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_DeleteNewPropertyAddress( _
        xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteNewPropertyAddress = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DeleteNewPropertyAddressExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteNewPropertyAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteNewPropertyAddress = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    Resume DeleteNewPropertyAddressExit

End Function

Public Function CreateNewPropertyAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Create a single instance of the persistant data associated with this
'               business object
' pass:
'   vstrXmlRequest  xml Request data stream containing data to be persisted
' return:
'   xml Response data stream containing results of operation
'   either: TYPE="SUCCESS"
'   or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPropertyAddressVbErr

    Const strFunctionName As String = "CreateNewPropertyAddress"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = INewPropertyBO_ValidateNewPropertyAddress(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript( _
        App.Title, TypeName(Me), strFunctionName, xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_CreateNewPropertyAddress(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateNewPropertyAddress = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateNewPropertyAddressExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateNewPropertyAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateNewPropertyAddress = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    Resume CreateNewPropertyAddressExit

End Function

Public Function GetNewPropertyGeneral(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get a subset of the data on the NEWPROPERTY table.
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetNewPropertyGeneralVbErr

    Const strFunctionName As String = "GetNewPropertyGeneral"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = INewPropertyBO_GetNewPropertyGeneral( _
        xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetNewPropertyGeneral = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetNewPropertyGeneralExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetNewPropertyGeneralVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetNewPropertyGeneral = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume GetNewPropertyGeneralExit

End Function

Public Function GetFullNewPropertyDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetFullNewPropertyDetailsVbErr

    Const strFunctionName As String = "GetFullNewPropertyDetails"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = INewPropertyBO_GetFullNewPropertyDetails( _
                                                    xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetFullNewPropertyDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetFullNewPropertyDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetFullNewPropertyDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetFullNewPropertyDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume GetFullNewPropertyDetailsExit

End Function

Public Function GetNewPropertyAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get NEWPROPERTYADDRESS and ADDRESS.
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetNewPropertyAddressVbErr

    Const strFunctionName As String = "GetNewPropertyAddress"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = INewPropertyBO_GetNewPropertyAddress( _
        xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetNewPropertyAddress = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetNewPropertyAddressExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetNewPropertyAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetNewPropertyAddress = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume GetNewPropertyAddressExit

End Function

Public Function GetRentalIncomeRates(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get rental income rate associated with the mortgage product on this application
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetRentalIncomeRatesVbErr

    Const strFunctionName As String = "GetRentalIncomeRates"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = INewPropertyBO_GetRentalIncomeRates( _
        xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetRentalIncomeRates = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetRentalIncomeRatesExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetRentalIncomeRatesVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetRentalIncomeRates = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume GetRentalIncomeRatesExit

End Function

Public Function GetNewPropertyDescription(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get a subset of the data on the NEWPROPERTY table, plus from some child tables.
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetNewPropertyDescriptionVbErr

    Const strFunctionName As String = "GetNewPropertyDescription"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = INewPropertyBO_GetNewPropertyDescription( _
        xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetNewPropertyDescription = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetNewPropertyDescriptionExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetNewPropertyDescriptionVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetNewPropertyDescription = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume GetNewPropertyDescriptionExit

End Function

Public Function GetNewPropertyAndAddressDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get a subset of the data on the NEWPROPERTY table, plus from some child tables.
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetNewPropertyAndAddressDetailsVBErr

    Const strFunctionName As String = "GetNewPropertyAndAddressDetails"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = INewPropertyBO_GetNewPropertyAndAddressDetails( _
        xmlIn.documentElement)
        
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetNewPropertyAndAddressDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetNewPropertyAndAddressDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetNewPropertyAndAddressDetailsVBErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetNewPropertyAndAddressDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume GetNewPropertyAndAddressDetailsExit

End Function

Private Function INewPropertyBO_CreateNewPropertyAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Create a single instance of the persistant data associated with this
'   business object. Invokes appropriate transactioning (Tx) business object.
' pass:
'   vxmlRequest  xml Request data stream containing data to be persisted
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPropertyAddressVbErr
    
    Const strFunctionName As String = "INewPropertyBO_CreateNewPropertyAddress"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
        
    Set xmlTempResponseNode = objINewPropertyTxBO.CreateNewPropertyAddress(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set INewPropertyBO_CreateNewPropertyAddress = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateNewPropertyAddressExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateNewPropertyAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_CreateNewPropertyAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    Resume CreateNewPropertyAddressExit

End Function

Private Function INewPropertyBO_DeleteNewPropertyAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   business object. Invokes appropriate transactioning (Tx) business object
' pass:
'   vxmlRequest
'       xml Request data stream containing data to be persisted
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteNewPropertyAddressVbErr
    
    Const strFunctionName As String = "INewPropertyBO_DeleteNewPropertyAddress"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
           
    '   call Business Transaction Object DeleteNewPropertyAddress function
    Set xmlTempResponseNode = objINewPropertyTxBO.DeleteNewPropertyAddress(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set INewPropertyBO_DeleteNewPropertyAddress = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DeleteNewPropertyAddressExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteNewPropertyAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_DeleteNewPropertyAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    Resume DeleteNewPropertyAddressExit

End Function

Private Function INewPropertyBO_GetFullNewPropertyDetails( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetDetailsVbErr
    
    Const strFunctionName As String = "INewPropertyBO_GetFullNewPropertyDetails"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objINewPropertyDO.GetFullNewPropertyDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set INewPropertyBO_GetFullNewPropertyDetails = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objINewPropertyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_GetFullNewPropertyDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetDetailsExit

End Function

Private Function INewPropertyBO_GetNewPropertyAndAddressDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetNewPropertyAndAddressDetailsVBErr
    
    Const strFunctionName As String = "INewPropertyBO_GetNewPropertyAndAddressDetails"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    Dim strTableName As String
    
    strTableName = "NEWPROPERTYANDADDRESSDETAILS"
                
    If vxmlRequest.nodeName = strTableName Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(strTableName).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, strTableName & " tag not found"
    End If
        
    Set xmlDataNode = objINewPropertyDO.GetNewPropertyAndAddressDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set INewPropertyBO_GetNewPropertyAndAddressDetails = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetNewPropertyAndAddressDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objINewPropertyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetNewPropertyAndAddressDetailsVBErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_GetNewPropertyAndAddressDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume GetNewPropertyAndAddressDetailsExit

End Function

Private Function INewPropertyBO_GetNewPropertyDescription( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetNewPropertyDescriptionVbErr
    
    Const strFunctionName As String = "INewPropertyBO_GetNewPropertyDescription"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    Dim strTableName As String
    
    strTableName = "NEWPROPERTY"
                
    If vxmlRequest.nodeName = strTableName Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(strTableName).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, strTableName & " tag not found"
    End If
        
    Set xmlDataNode = objINewPropertyDO.GetNewPropertyDescription(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set INewPropertyBO_GetNewPropertyDescription = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetNewPropertyDescriptionExit:
    
    Set xmlRequestNode = Nothing
    Set objINewPropertyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetNewPropertyDescriptionVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_GetNewPropertyDescription = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume GetNewPropertyDescriptionExit

End Function

Private Function INewPropertyBO_GetNewPropertyGeneral( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetNewPropertyGeneralVbErr
    
    Const strFunctionName As String = "INewPropertyBO_GetNewPropertyGeneral"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
                
    Dim strTableName As String
    
    strTableName = "NEWPROPERTY"
    
    If vxmlRequest.nodeName = strTableName Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(strTableName).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, strTableName & " tag not found"
    End If
        
    Set xmlDataNode = objINewPropertyDO.GetNewPropertyGeneral(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set INewPropertyBO_GetNewPropertyGeneral = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetNewPropertyGeneralExit:
    
    Set xmlRequestNode = Nothing
    Set objINewPropertyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetNewPropertyGeneralVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_GetNewPropertyGeneral = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume GetNewPropertyGeneralExit

End Function

Private Function INewPropertyBO_GetNewPropertyAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetNewPropertyAddressVbErr
    
    Const strFunctionName As String = "INewPropertyBO_GetNewPropertyAddress"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
                
    Dim strTableName As String
    
    strTableName = "NEWPROPERTYADDRESS"
    
    If vxmlRequest.nodeName = strTableName Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(strTableName).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, strTableName & " tag not found"
    End If
        
    Set xmlDataNode = objINewPropertyDO.GetNewPropertyAddress(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set INewPropertyBO_GetNewPropertyAddress = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetNewPropertyAddressExit:
    
    Set xmlRequestNode = Nothing
    Set objINewPropertyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetNewPropertyAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_GetNewPropertyAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume GetNewPropertyAddressExit

End Function

Private Function INewPropertyBO_GetPropertyUnknownDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a subset of the data on the NEWPROPERTY table.
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetPropertyUnknownDetailsVbErr
    
    Const strFunctionName As String = "INewPropertyBO_GetPropertyUnknownDetails"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objINewPropertyDO.GetPropertyUnknownDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set INewPropertyBO_GetPropertyUnknownDetails = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetPropertyUnknownDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objINewPropertyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetPropertyUnknownDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_GetPropertyUnknownDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetPropertyUnknownDetailsExit

End Function

Private Function INewPropertyBO_GetRentalIncomeRates(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
On Error GoTo GetRentalIncomeRatesVbErr
    
    Const strFunctionName As String = "INewPropertyBO_GetRentalIncomeRates"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMElement
    Dim xmlTempApplicationElem As IXMLDOMElement
    Dim xmlTempApplicationNumberElem As IXMLDOMElement
    Dim xmlTempApplicationFFNumber As IXMLDOMElement
    Dim xmlTempResponse As IXMLDOMElement
    
    Dim objIApplicationBO As IApplicationBO
    Dim objQuotationBO As Object
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlTempQuotationNumber As IXMLDOMElement
    Dim strResponse As String

    Dim xmlTempLoanCompElem As IXMLDOMElement
    Dim xmlTempProdCodeElem As IXMLDOMElement
    Dim xmlTempProdStartDateElem As IXMLDOMElement
    Dim xmlTempAppDate As IXMLDOMElement
    Dim xmlTempLoanAmount As IXMLDOMElement
    Dim xmlTempLTV As IXMLDOMElement

    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim strAppDate As String
    Dim strQuotationNumber As String
    Dim strLoanAmount As String
    Dim strLTV As String
    Dim strProductCode As String
    Dim strProductStartDate As String
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If

    ' get application numbers
    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(xmlRequestNode, ".//APPLICATIONNUMBER")
    strApplicationFFNumber = m_objXmlAssist.GetMandatoryElementText(xmlRequestNode, ".//APPLICATIONFACTFINDNUMBER")
    
    ' get the application details to find the application date
    Set objIApplicationBO = New ApplicationBO
    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTempApplicationElem = xmlOut.createElement("APPLICATION")
    Set xmlTempApplicationNumberElem = xmlOut.createElement("APPLICATIONNUMBER")
    Set xmlTempApplicationFFNumber = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    
    xmlTempApplicationNumberElem.Text = strApplicationNumber
    xmlTempApplicationFFNumber.Text = strApplicationFFNumber
    xmlTempApplicationElem.appendChild xmlTempApplicationNumberElem
    xmlTempApplicationElem.appendChild xmlTempApplicationFFNumber
    xmlTempRequest.appendChild xmlTempApplicationElem
    Set xmlTempResponse = objIApplicationBO.GetApplicationData(xmlTempRequest)
    'Check for error
    m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
    'Get ActiveQuotationNumber
    strQuotationNumber = m_objXmlAssist.GetMandatoryElementText(xmlTempResponse, ".//ACTIVEQUOTENUMBER")
    'Get ApplicationDate
    strAppDate = m_objXmlAssist.GetElementText(xmlTempResponse, ".//APPLICATIONDATE")
    'Destroy all objects used in this call
    Set xmlTempRequest = Nothing
    Set xmlTempApplicationNumberElem = Nothing
    Set xmlTempApplicationFFNumber = Nothing
    Set xmlTempApplicationElem = Nothing
    Set objIApplicationBO = Nothing

    'get the quotation details
    
    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTempApplicationElem = xmlOut.createElement("APPLICATION")
    Set xmlTempApplicationNumberElem = xmlOut.createElement("APPLICATIONNUMBER")
    Set xmlTempApplicationFFNumber = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    
    xmlTempApplicationNumberElem.Text = strApplicationNumber
    xmlTempApplicationFFNumber.Text = strApplicationFFNumber
    xmlTempApplicationElem.appendChild xmlTempApplicationNumberElem
    xmlTempRequest.appendChild xmlTempApplicationElem
    xmlTempApplicationElem.appendChild xmlTempApplicationFFNumber
    Set objQuotationBO = m_objContext.CreateInstance("omAQ.ApplicationQuoteBO")
    strResponse = objQuotationBO.GetAcceptedOrActiveQuoteData(xmlTempRequest.xml)
    errCheckXMLResponse strResponse
    Set xmlDoc = xmlLoad(strResponse, strFunctionName)
    
    strLoanAmount = xmlDoc.selectSingleNode("//MORTGAGESUBQUOTE/AMOUNTREQUESTED").Text
    strLTV = xmlDoc.selectSingleNode("//MORTGAGESUBQUOTE/LTV").Text

    Set xmlTempRequest = Nothing
    Set xmlTempApplicationNumberElem = Nothing
    Set xmlTempApplicationFFNumber = Nothing
    Set xmlTempApplicationElem = Nothing
    Set objQuotationBO = Nothing
    
    'need to call GetQuotationSummary too
    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTempApplicationElem = xmlOut.createElement("QUOTATION")
    Set xmlTempApplicationNumberElem = xmlOut.createElement("APPLICATIONNUMBER")
    Set xmlTempApplicationFFNumber = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    Set xmlTempQuotationNumber = xmlOut.createElement("QUOTATIONNUMBER")
    
    xmlTempApplicationNumberElem.Text = strApplicationNumber
    xmlTempApplicationFFNumber.Text = strApplicationFFNumber
    xmlTempQuotationNumber.Text = strQuotationNumber
    xmlTempApplicationElem.appendChild xmlTempApplicationNumberElem
    xmlTempRequest.appendChild xmlTempApplicationElem
    xmlTempApplicationElem.appendChild xmlTempApplicationFFNumber
    xmlTempApplicationElem.appendChild xmlTempQuotationNumber
    Set objQuotationBO = m_objContext.CreateInstance("omCM.QuotationBO")
    strResponse = objQuotationBO.GetQuotationSummary(xmlTempRequest.xml)
    errCheckXMLResponse strResponse
    Set xmlDoc = xmlLoad(strResponse, strFunctionName)
    
    strProductCode = xmlDoc.selectSingleNode("//QUOTATIONSUMMARY/LOANCOMPONENTLIST/LOANCOMPONENT/MORTGAGEPRODUCTCODE").Text
    strProductStartDate = xmlDoc.selectSingleNode("//QUOTATIONSUMMARY/LOANCOMPONENTLIST/LOANCOMPONENT/STARTDATE").Text
    
    Set xmlTempRequest = Nothing
    Set xmlTempApplicationNumberElem = Nothing
    Set xmlTempApplicationFFNumber = Nothing
    Set xmlTempApplicationElem = Nothing
    Set objQuotationBO = Nothing
    
    ' build the xml request to call GetRentalIncome
    
    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTempLoanCompElem = xmlOut.createElement("LOANCOMPONENT")
    Set xmlTempProdCodeElem = xmlOut.createElement("MORTGAGEPRODUCTCODE")
    Set xmlTempProdStartDateElem = xmlOut.createElement("MORTGAGEPRODUCTSTARTDATE")
    Set xmlTempAppDate = xmlOut.createElement("APPLICATIONDATE")
    Set xmlTempLoanAmount = xmlOut.createElement("LOANAMOUNT")
    Set xmlTempLTV = xmlOut.createElement("LTV")
    
    xmlTempProdCodeElem.Text = strProductCode
    xmlTempProdStartDateElem.Text = strProductStartDate
    xmlTempAppDate.Text = strAppDate
    xmlTempLoanAmount.Text = strLoanAmount
    xmlTempLTV.Text = strLTV
    
    xmlTempLoanCompElem.appendChild xmlTempProdCodeElem
    xmlTempLoanCompElem.appendChild xmlTempProdStartDateElem
    xmlTempLoanCompElem.appendChild xmlTempAppDate
    xmlTempLoanCompElem.appendChild xmlTempLoanAmount
    xmlTempLoanCompElem.appendChild xmlTempLTV
    xmlTempRequest.appendChild xmlTempLoanCompElem
    
    Dim objMortgageProductBO As Object
    Set objMortgageProductBO = m_objContext.CreateInstance("omMP.MortgageProductBO")
    strResponse = objMortgageProductBO.GetMortgageProductRentalIncomeRateSet(xmlTempRequest.xml)
    errCheckXMLResponse strResponse
    Set xmlDoc = xmlLoad(strResponse, strFunctionName)
    
    Set INewPropertyBO_GetRentalIncomeRates = xmlDoc.documentElement.cloneNode(True)
    
    Set xmlTempRequest = Nothing
    Set xmlTempLoanCompElem = Nothing
    Set xmlTempProdCodeElem = Nothing
    Set xmlTempProdStartDateElem = Nothing
    Set xmlTempAppDate = Nothing
    Set xmlTempLoanAmount = Nothing
    Set xmlTempLTV = Nothing
    Set objMortgageProductBO = Nothing
    
GetRentalIncomeRatesExit:
    
    Set xmlRequestNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objIApplicationBO = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempApplicationNumberElem = Nothing
    Set xmlTempApplicationFFNumber = Nothing
    Set xmlTempApplicationElem = Nothing
    Set xmlTempResponse = Nothing
    Exit Function

GetRentalIncomeRatesVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_GetRentalIncomeRates = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetRentalIncomeRatesExit
End Function

Private Function INewPropertyBO_SaveNewProperty( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveNewPropertyVbErr
    
    Const strFunctionName As String = "INewPropertyBO_SaveNewProperty"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
    
    Set xmlTempResponseNode = objINewPropertyTxBO.SaveNewProperty(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set INewPropertyBO_SaveNewProperty = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveNewPropertyExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

SaveNewPropertyVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_SaveNewProperty = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNewPropertyExit

End Function

Private Function INewPropertyBO_SaveNewPropertyAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveNewPropertyAddress a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveNewPropertyAddressVbErr
    
    Const strFunctionName As String = "INewPropertyBO_SaveNewPropertyAddress"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New ApplicationTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
    
    Set xmlTempResponseNode = objINewPropertyTxBO.SaveNewPropertyAddress(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set INewPropertyBO_SaveNewPropertyAddress = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveNewPropertyAddressExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

SaveNewPropertyAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_SaveNewPropertyAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNewPropertyAddressExit

End Function

Private Function INewPropertyBO_SaveNewPropertyAndAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveNewPropertyAndAddressVbErr
    
    Const strFunctionName As String = "INewPropertyBO_SaveNewPropertyAndAddress"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
    
    Set xmlTempResponseNode = objINewPropertyTxBO.SaveNewPropertyAndAddress(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set INewPropertyBO_SaveNewPropertyAndAddress = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveNewPropertyAndAddressExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

SaveNewPropertyAndAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_SaveNewPropertyAndAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNewPropertyAndAddressExit

End Function

Private Function INewPropertyBO_SaveNewPropertyAndDeposit( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveNewPropertyAndDepositVbErr
    
    Const strFunctionName As String = "INewPropertyBO_SaveNewPropertyAndDeposit"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
    
    Set xmlTempResponseNode = objINewPropertyTxBO.SaveNewPropertyAndDeposit(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set INewPropertyBO_SaveNewPropertyAndDeposit = xmlResponseElem

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveNewPropertyAndDepositExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

SaveNewPropertyAndDepositVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_SaveNewPropertyAndDeposit = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNewPropertyAndDepositExit

End Function

Private Function INewPropertyBO_Validate(ByVal vxmlRequest As IXMLDOMElement, ByVal vbomMethod As INewPropertyBOMethod) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------

On Error GoTo ValidateVbErr

    Const strFunctionName As String = "INewPropertyBO_Validate"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set INewPropertyBO_Validate = xmlResponseElem
    
ValidateExit:

    Set xmlRequestNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
        
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

    Set INewPropertyBO_Validate = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateExit

End Function

Private Function INewPropertyBO_ValidateNewPropertyAddress( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   performs hard coded validation for this object
' pass:
'   vxmlRequest  xml Request data stream containing details of action to be
'   performed and data required. This is the full request block as received from the client
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateNewPropertyAddressVbErr

    Const strFunctionName As String = "INewPropertyBO_ValidateNewPropertyAddress"

    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
        
    ' Get node to add derived data to and apply validation to
    If vxmlRequest.nodeName = "UPDATE" Or vxmlRequest.nodeName = "CREATE" Then
        Set xmlRequestTypeNode = vxmlRequest
    ElseIf vxmlRequest.getElementsByTagName("UPDATE").length > 0 Then
        Set xmlRequestTypeNode = vxmlRequest.getElementsByTagName("UPDATE").Item(0)
    ElseIf vxmlRequest.getElementsByTagName("CREATE").length > 0 Then
        Set xmlRequestTypeNode = vxmlRequest.getElementsByTagName("CREATE").Item(0)
    End If
        
    If Not xmlRequestTypeNode Is Nothing Then
        Set xmlDataNode = xmlRequestTypeNode.firstChild

        ' Add in derived data
        'If Not xmlDataNode Is Nothing Then
        '    objINewPropertyDO.AddDerivedData xmlDataNode
        'End If
    End If
    
' TODO ************************************************************************
'   add hard coded validation here, & raise error if it fails
' *****************************************************************************

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set INewPropertyBO_ValidateNewPropertyAddress = xmlResponseElem.cloneNode(True)
    
ValidateNewPropertyAddressExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set objINewPropertyDO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateNewPropertyAddressVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_ValidateNewPropertyAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume ValidateNewPropertyAddressExit

End Function

Public Function GetVendorDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetVendorDetailsVbErr

    Const strFunctionName As String = "GetVendorDetails"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = INewPropertyBO_GetVendorDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    m_objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetVendorDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetVendorDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetVendorDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetVendorDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetVendorDetailsExit

End Function

Private Function INewPropertyBO_GetVendorDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetVendorDetailsVbErr
    
    Const strFunctionName As String = "INewPropertyBO_GetVendorDetails"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
    
    Dim strTableName As String
    strTableName = "NEWPROPERTYVENDOR"
                
    If vxmlRequest.nodeName = strTableName Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(strTableName).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, strTableName & " tag not found"
    End If
        
    Set xmlDataNode = objINewPropertyDO.GetVendorDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set INewPropertyBO_GetVendorDetails = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetVendorDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objINewPropertyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetVendorDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_GetVendorDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetVendorDetailsExit
End Function

Private Function INewPropertyBO_DeleteVendorDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVendorDetailsVbErr
    
    Const strFunctionName As String = "INewPropertyBO_DeleteVendorDetails"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
           
    '   call Business Transaction Object DeleteVendorDetails function
    Set xmlTempResponseNode = objINewPropertyTxBO.DeleteVendorDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set INewPropertyBO_DeleteVendorDetails = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DeleteVendorDetailsExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteVendorDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_DeleteVendorDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteVendorDetailsExit
End Function

Public Function DeleteVendorDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVendorDetailsVbErr

    Const strFunctionName As String = "DeleteVendorDetails"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_DeleteVendorDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteVendorDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DeleteVendorDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteVendorDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteVendorDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteVendorDetailsExit

End Function

Public Function UpdateVendorDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVendorDetailsVbErr

    Const strFunctionName As String = "UpdateVendorDetails"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = INewPropertyBO_ValidateVendorDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript( _
        App.Title, TypeName(Me), strFunctionName, xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_UpdateVendorDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateVendorDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateVendorDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

UpdateVendorDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateVendorDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateVendorDetailsExit
    
End Function

Private Function INewPropertyBO_UpdateVendorDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateVendorDetails a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVendorDetailsVbErr
    
    Const strFunctionName As String = "INewPropertyBO_UpdateVendorDetails"
    
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
    
    Set xmlTempResponseNode = objINewPropertyTxBO.UpdateVendorDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set INewPropertyBO_UpdateVendorDetails = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateVendorDetailsExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateVendorDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_UpdateVendorDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateVendorDetailsExit
End Function

Private Function INewPropertyBO_ValidateVendorDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVendorDetailsVbErr

    Const strFunctionName As String = "INewPropertyBO_ValidateVendorDetails"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
       
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set INewPropertyBO_ValidateVendorDetails = xmlResponseElem.cloneNode(True)
    
ValidateVendorDetailsExit:

    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateVendorDetailsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_ValidateVendorDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateVendorDetailsExit

End Function

Public Function CreateVendorDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXmlRequest
'       xml Request data stream containing data to be persisted
' return:
'   xml Response data stream containing results of operation
'   either: TYPE="SUCCESS"
'   or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateVendorDetailsVbErr

    Const strFunctionName As String = "CreateVendorDetails"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = INewPropertyBO_ValidateVendorDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, TypeName(Me), _
        strFunctionName, xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_CreateVendorDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateVendorDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateVendorDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

CreateVendorDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateVendorDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateVendorDetailsExit

End Function

Private Function INewPropertyBO_CreateVendorDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateVendorDetailsVbErr
    
    Const strFunctionName As String = "INewPropertyBO_CreateVendorDetails"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
        
    '   call Business Transaction Object CreateVendorDetails function
    Set xmlTempResponseNode = objINewPropertyTxBO.CreateVendorDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set INewPropertyBO_CreateVendorDetails = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateVendorDetailsExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateVendorDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_CreateVendorDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateVendorDetailsExit
End Function

Public Function CreateNewPropertyDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXmlRequest
'       xml Request data stream containing data to be persisted
' return:
'   xml Response data stream containing results of operation
'   either: TYPE="SUCCESS"
'   or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPropertyDetailsVbErr

    Const strFunctionName As String = "CreateNewPropertyDetails"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Run validation checks
    Set xmlTempResponseNode = INewPropertyBO_ValidateNewPropertyDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, TypeName(Me), _
        strFunctionName, xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_CreateNewPropertyDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateNewPropertyDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateNewPropertyDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

CreateNewPropertyDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateNewPropertyDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewPropertyDetailsExit

End Function

Private Function INewPropertyBO_CreateNewPropertyDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlRequest
'       xml Request data stream containing data to be persisted
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPropertyDetailsVbErr
    
    Const strFunctionName As String = "INewPropertyBO_CreateNewPropertyDetails"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
        
    '   call Business Transaction Object CreateNewPropertyDetails function
    Set xmlTempResponseNode = objINewPropertyTxBO.CreateNewPropertyDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set INewPropertyBO_CreateNewPropertyDetails = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateNewPropertyDetailsExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateNewPropertyDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_CreateNewPropertyDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewPropertyDetailsExit
End Function

Public Function CreateNewProperty(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPropertyVbErr

    Const strFunctionName As String = "CreateNewProperty"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode

    Dim xmlDirectoryGUIDNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    ' Run validation checks
    Set xmlTempResponseNode = INewPropertyBO_ValidateNewProperty(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, TypeName(Me), _
        strFunctionName, xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_CreateNewProperty(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateNewProperty = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CreateNewPropertyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlDirectoryGUIDNode = Nothing

    Exit Function

CreateNewPropertyVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateNewProperty = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewPropertyExit

End Function

Private Function INewPropertyBO_CreateNewProperty( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlRequest  xml Request data stream containing data to be persisted
' return:
'   xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPropertyVbErr
    
    Const strFunctionName As String = "INewPropertyBO_CreateNewProperty"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
        
    '   call Business Transaction Object CreateNewProperty function
    Set xmlTempResponseNode = objINewPropertyTxBO.CreateNewProperty(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set INewPropertyBO_CreateNewProperty = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CreateNewPropertyExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateNewPropertyVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_CreateNewProperty = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateNewPropertyExit
End Function

Private Function INewPropertyBO_ValidateNewPropertyDetails( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateNewPropertyDetailsVbErr

    Const strFunctionName As String = "INewPropertyBO_ValidateNewPropertyDetails"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If (m_objXmlAssist.GetTagValue(vxmlRequest, "DATELEASESTARTED") <> "") Then
        If CDate(m_objXmlAssist.GetTagValue(vxmlRequest, "DATELEASESTARTED")) > Now() Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 284
        End If
    End If

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set INewPropertyBO_ValidateNewPropertyDetails = xmlResponseElem.cloneNode(True)
    
ValidateNewPropertyDetailsExit:

    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateNewPropertyDetailsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_ValidateNewPropertyDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateNewPropertyDetailsExit

End Function

Private Function INewPropertyBO_ValidateNewProperty( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateNewPropertyVbErr

    Const strFunctionName As String = "INewPropertyBO_ValidateNewProperty"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set INewPropertyBO_ValidateNewProperty = xmlResponseElem.cloneNode(True)
    
ValidateNewPropertyExit:

    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateNewPropertyVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_ValidateNewProperty = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume ValidateNewPropertyExit

End Function

Public Function UpdateNewProperty(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNewPropertyVbErr

    Const strFunctionName As String = "UpdateNewProperty"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = INewPropertyBO_ValidateNewProperty(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_UpdateNewProperty(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateNewProperty = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateNewPropertyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateNewPropertyVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateNewProperty = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateNewPropertyExit
    
End Function

Private Function INewPropertyBO_UpdateNewProperty( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNewPropertyVbErr
    
    Const strFunctionName As String = "INewPropertyBO_UpdateNewProperty"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
    
    Set xmlTempResponseNode = objINewPropertyTxBO.UpdateNewProperty(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set INewPropertyBO_UpdateNewProperty = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateNewPropertyExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateNewPropertyVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_UpdateNewProperty = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateNewPropertyExit

End Function

Public Function UpdateNewPropertyDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXmlRequest  xml Request data stream containing data to be persisted
' return:
'   xml Response data stream containing results of operation
'   either: TYPE="SUCCESS"
'   or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNewPropertyDetailsVbErr

    Const strFunctionName As String = "UpdateNewPropertyDetails"
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = INewPropertyBO_ValidateNewPropertyDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript( _
        App.Title, TypeName(Me), strFunctionName, xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_UpdateNewPropertyDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateNewPropertyDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateNewPropertyDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateNewPropertyDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateNewPropertyDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateNewPropertyDetailsExit
    
End Function

Private Function INewPropertyBO_UpdateNewPropertyDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNewPropertyDetailsVbErr
    
    Const strFunctionName As String = "INewPropertyBO_UpdateNewPropertyDetails"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
    
    Set xmlTempResponseNode = objINewPropertyTxBO.UpdateNewPropertyDetails(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set INewPropertyBO_UpdateNewPropertyDetails = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateNewPropertyDetailsExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateNewPropertyDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_UpdateNewPropertyDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateNewPropertyDetailsExit

End Function

Public Function UpdateNewPropertyAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXmlRequest  xml Request data stream containing data to be persisted
' return:
'   xml Response data stream containing results of operation
'   either: TYPE="SUCCESS"
'   or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNewPropertyAddressVbErr

    Const strFunctionName As String = "UpdateNewPropertyAddress"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = INewPropertyBO_ValidateNewPropertyDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript( _
        App.Title, TypeName(Me), strFunctionName, xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_UpdateNewPropertyAddress(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateNewPropertyAddress = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

UpdateNewPropertyAddressExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateNewPropertyAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateNewPropertyAddress = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateNewPropertyAddressExit
    
End Function

Private Function INewPropertyBO_UpdateNewPropertyAddress(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   business object
'   Invokes appropriate transactioning (Tx) business object
' pass:
'   vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNewPropertyAddressVbErr
    
    Const strFunctionName As String = "INewPropertyBO_UpdateNewPropertyAddress"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
    
    Set xmlTempResponseNode = objINewPropertyTxBO.UpdateNewPropertyAddress(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set INewPropertyBO_UpdateNewPropertyAddress = xmlResponseElem.cloneNode(True)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
UpdateNewPropertyAddressExit:

    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateNewPropertyAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_UpdateNewPropertyAddress = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateNewPropertyAddressExit

End Function

Public Function GetValuationTypeAndLocation(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get a subset of the data on the NEWPROPERTY table.
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetValuationTypeAndLocationVbErr

    Const strFunctionName As String = "GetValuationTypeAndLocation"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = INewPropertyBO_GetValuationTypeAndLocation( _
        xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetValuationTypeAndLocation = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetValuationTypeAndLocationExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetValuationTypeAndLocationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetValuationTypeAndLocation = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume GetValuationTypeAndLocationExit

End Function

Private Function INewPropertyBO_GetValuationTypeAndLocation( _
    ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetValuationTypeAndLocationVbErr
    
    Const strFunctionName As String = "INewPropertyBO_GetValuationTypeAndLocation"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    
    Dim objINewPropertyDO As INewPropertyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
    End If
                
    Dim strTableName As String
    
    strTableName = "NEWPROPERTY"
    
    If vxmlRequest.nodeName = strTableName Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(strTableName).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, strTableName & " tag not found"
    End If
        
    Set xmlDataNode = objINewPropertyDO.GetValuationTypeAndLocation(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set INewPropertyBO_GetValuationTypeAndLocation = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetValuationTypeAndLocationExit:
    
    Set xmlRequestNode = Nothing
    Set objINewPropertyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetValuationTypeAndLocationVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_GetValuationTypeAndLocation = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume GetValuationTypeAndLocationExit

End Function

Public Function SaveNewProperty(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveNewPropertyVbErr

    Const strFunctionName As String = "SaveNewProperty"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks - Add here if required
    
        
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_SaveNewProperty(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveNewProperty = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

SaveNewPropertyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveNewPropertyVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveNewProperty = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNewPropertyExit
    
End Function

Public Function SaveNewPropertyAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveNewPropertyAddress a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveNewPropertyAddressVbErr

    Const strFunctionName As String = "SaveNewPropertyAddress"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_SaveNewPropertyAddress(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveNewPropertyAddress = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

SaveNewPropertyAddressExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveNewPropertyAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveNewPropertyAddress = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNewPropertyAddressExit
    
End Function

Public Function SaveNewPropertyAndAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveNewPropertyAndAddressVbErr

    Const strFunctionName As String = "SaveNewPropertyAndAddress"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
' TODO ************************************************************************
'   replace bomUpdate with bomMethodName
' *****************************************************************************
    Set xmlTempResponseNode = INewPropertyBO_Validate(xmlIn.documentElement, bomSaveNewPropertyAndAddress)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_SaveNewPropertyAndAddress(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveNewPropertyAndAddress = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

SaveNewPropertyAndAddressExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveNewPropertyAndAddressVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveNewPropertyAndAddress = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNewPropertyAndAddressExit
    
End Function

Public Function SaveNewPropertyAndDeposit(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Update a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'<REQUEST ACTION=""CREATE"">
'   <NEWPROPERTY>
'       <APPLICATIONNUMBER>C00040576</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'       <ANYOTHERRESIDENTSINDICATOR>1</ANYOTHERRESIDENTSINDICATOR>
'       <APPLICANTTOOCCUPYINDICATOR>1</APPLICANTTOOCCUPYINDICATOR>
'       <CURRENTLYOCCUPIEDINDICATOR>1</CURRENTLYOCCUPIEDINDICATOR>
'       .................
'       .................
'
'       <NEWPROPERTYDEPOSIT>
'           <APPLICATIONNUMBER>C00040576</APPLICATIONNUMBER>
'           <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'           <NPDEPOSITSEQUENCENUMBER>1</NPDEPOSITSEQUENCENUMBER>
'           <SOURCEOFFUNDING></SOURCEOFFUNDING>
'           <AMOUNT>1500</AMOUNT>
'       </NEWPROPERTYDEPOSIT>
'   </NEWPROPERTY>
'</REQUEST>
'------------------------------------------------------------------------------------------
On Error GoTo SaveNewPropertyAndDepositVbErr

    Const strFunctionName As String = "SaveNewPropertyAndDeposit"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = INewPropertyBO_Validate(xmlIn.documentElement, bomSaveNewPropertyAndDeposit)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, TypeName(Me), _
                                                    strFunctionName, xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = INewPropertyBO_SaveNewPropertyAndDeposit(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    SaveNewPropertyAndDeposit = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

SaveNewPropertyAndDepositExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

SaveNewPropertyAndDepositVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveNewPropertyAndDeposit = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveNewPropertyAndDepositExit
    
End Function

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

Public Function GetPropertyUnknownDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get a subset of the data on the NEWPROPERTY table.
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo GetPropertyUnknownDetailsVbErr

    Const strFunctionName As String = "GetPropertyUnknownDetails"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = INewPropertyBO_GetPropertyUnknownDetails( _
        xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetPropertyUnknownDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetPropertyUnknownDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetPropertyUnknownDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetPropertyUnknownDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume GetPropertyUnknownDetailsExit

End Function

Public Function CalcExcessRentalIncome(ByVal vstrXMLRequest As String) As String

'---------------------------------------------REQUEST FORMAT :-  -------------------------------------------

'<REQUEST USERID="GregD" UNITID="Unit1" MACHINEID="CH006346" CHANNELID="1" ADMINSYSTEMSTATE="null" USERAUTHORITYLEVEL="99" OPERATION="SEARCH" ACTION="SEARCH">
'   <NEWPROPERTY>
'      <RENTALDETAILS>
'         <APPLICATIONNUMBER>100003069</APPLICATIONNUMBER>
'         <APPLICATIONFACTFINDNUMBER>1</APPLICATIONFACTFINDNUMBER>
'         <CUSTOMERNUMBER>CUST0001</CUSTOMERNUMBER>
'         <CUSTOMERVERSIONNUMBER>1</CUSTOMERVERSIONNUMBER>
'         <REQUESTEDLOANAMOUNT>100000</REQUESTEDLOANAMOUNT>
'         <MONTHLYRENTALINCOME>375</MONTHLYRENTALINCOME>
'      </RENTALDETAILS>
'   </NEWPROPERTY>
'</REQUEST>



'GD BMIDS00165 - New Method

On Error GoTo CalcExcessRentalIncomeVbErr

    Const strFunctionName As String = "CalcExcessRentalIncome"

    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = INewPropertyBO_CalcExcessRentalIncome( _
                                                    xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    CalcExcessRentalIncome = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

CalcExcessRentalIncomeExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CalcExcessRentalIncomeVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CalcExcessRentalIncome = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Resume CalcExcessRentalIncomeExit

End Function

Private Function INewPropertyBO_CalcExcessRentalIncome( _
            ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode

'GD BMIDS00165 - New Method

On Error GoTo CalcExcessRentalIncomeVbErr
    
    Const strFunctionName As String = "INewPropertyBO_CalcExcessRentalIncome"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlTempRequest As IXMLDOMElement
    Dim xmlTempApplicationElem As IXMLDOMElement
    Dim xmlTempApplicationNumberElem As IXMLDOMElement
    Dim xmlTempApplicationFFNumber As IXMLDOMElement
    Dim xmlTempResponse As IXMLDOMElement
    
    Dim objGlobalParameterDO As GlobalParameterDO
    Dim objIApplicationBO As IApplicationBO
    Dim objCFBO As ICustomerFinancialBO
    Dim xmlCustList As IXMLDOMElement
    
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim strCustomerNumber As String
    Dim strCustomerVersionNumber As String
    Dim strRequestedLoanAmount As String
    Dim strMonthlyRentalIncome As String
    Dim strTypeOfApplication As String
    Dim strSpecialScheme As String
    Dim dblRequestedLoanAmount As Double
    Dim dblMonthlyRentalIncome As Double
    Dim dblTotLoansNotRedeemed As Double
    Dim dblMonthlyPayment As Double
    Dim dblExcessMonthlyIncome As Double
    Dim dblPercentageOfRentalIncome1 As Double
    Dim dblBuyToLetBaseRate As Double

    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        'Set objINewPropertyDO = New NewPropertyDO

        Set objGlobalParameterDO = New GlobalParameterDO
        Set objCFBO = New CustomerFinancialBO
    Else
        'Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")

        Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
        Set objCFBO = m_objContext.CreateInstance(gstrCUSTOMER_FINANCIAL & ".CustomerFinancialBO")
    End If
    
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If

    '-------------------------------------Check for mandatory elements------------------------------------
    Set xmlTempNode = m_objXmlAssist.GetMandatoryNode(xmlRequestNode, "RENTALDETAILS")
    Set xmlTempNode = Nothing
    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(xmlRequestNode, ".//APPLICATIONNUMBER")
    strApplicationFFNumber = m_objXmlAssist.GetMandatoryElementText(xmlRequestNode, ".//APPLICATIONFACTFINDNUMBER")
    strCustomerNumber = m_objXmlAssist.GetMandatoryElementText(xmlRequestNode, ".//CUSTOMERNUMBER")
    strCustomerVersionNumber = m_objXmlAssist.GetMandatoryElementText(xmlRequestNode, ".//CUSTOMERVERSIONNUMBER")
    strRequestedLoanAmount = m_objXmlAssist.GetMandatoryElementText(xmlRequestNode, ".//REQUESTEDLOANAMOUNT")
    strMonthlyRentalIncome = m_objXmlAssist.GetMandatoryElementText(xmlRequestNode, ".//MONTHLYRENTALINCOME")
    
    '-------------------------------------Convert Numerics to Doubles-------------------------------------
    dblRequestedLoanAmount = CSafeDbl(strRequestedLoanAmount)
    dblMonthlyRentalIncome = CSafeDbl(strMonthlyRentalIncome)
    
    '----------------Initialise TotLoansNotRedeemed, MonthlyPayment and ExcessMonthlyRentalIncome---------
    dblTotLoansNotRedeemed = 0
    dblMonthlyPayment = 0
    dblExcessMonthlyIncome = 0

    '--Get the value of PercentageOfRentalIncome1 and BuyToLetBaseRate from global Parameter for Excess Rental Amount calculation
    'GD BMIDS000217 - Handle errors in NULL or NOT FOUND GlobalParameters
    'JD BMIDS765 BuyToLetBaseRate now held in the database NOT as a global parameter
    Dim lngErrorNumber As Long
    'On Error Resume Next
    Dim xmlRentalIncomeNode As IXMLDOMNode
    Dim xmlRentalRateNode As IXMLDOMNode
    Set xmlRentalIncomeNode = INewPropertyBO_GetRentalIncomeRates(vxmlRequest)
    'dblBuyToLetBaseRate = objGlobalParameterDO.GetCurrentParameterByType("BuyToLetRate", "PERCENTAGE")
    Set xmlRentalRateNode = xmlRentalIncomeNode.selectSingleNode("//RENTALINCOMERATESETBAND/RENTALINCOMEINTERESTRATE")
    If Not xmlRentalRateNode Is Nothing Then
        dblBuyToLetBaseRate = CDbl(xmlRentalRateNode.Text)
    Else
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingBuyToLetGlobalParameter ', "'BuyToLetRate' Globalparameter Blah Blah Blah"
    End If
    'Save error number
    'lngErrorNumber = Err.Number
    'On Error GoTo CalcExcessRentalIncomeVbErr
    'If lngErrorNumber = 94 Or m_objErrAssist.GetOmigaErrorNumber(lngErrorNumber) = oeRecordNotFound Then 'Invalid Use of Null OR RECORD NOT FOUND
    '    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingBuyToLetGlobalParameter ', "'BuyToLetRate' Globalparameter Blah Blah Blah"
    'End If
    
    On Error Resume Next
    dblPercentageOfRentalIncome1 = objGlobalParameterDO.GetCurrentParameterByType("PercentageOfRentalIncome1", "PERCENTAGE")
    'Save error number
    lngErrorNumber = Err.Number
    On Error GoTo CalcExcessRentalIncomeVbErr
    If lngErrorNumber = 94 Or m_objErrAssist.GetOmigaErrorNumber(lngErrorNumber) = oeRecordNotFound Then 'Invalid Use of Null OR RECORD NOT FOUND
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPercentageRentalIncome1GlobalParameter ', "'PercentageOfRentalIncome1' Globalparameter Blah Blah Blah"
    End If

    '--Call GetApplicationData to get Type of Application and Special Scheme Type for calculation of the Excess Rental Amount
    'Create all objects for this call
    Set objIApplicationBO = New ApplicationBO
    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTempApplicationElem = xmlOut.createElement("APPLICATION")
    Set xmlTempApplicationNumberElem = xmlOut.createElement("APPLICATIONNUMBER")
    Set xmlTempApplicationFFNumber = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    
    xmlTempApplicationNumberElem.Text = strApplicationNumber
    xmlTempApplicationFFNumber.Text = strApplicationFFNumber
    xmlTempApplicationElem.appendChild xmlTempApplicationNumberElem
    xmlTempRequest.appendChild xmlTempApplicationElem
    Set xmlTempResponse = objIApplicationBO.GetApplicationData(xmlTempRequest)
    'Check for error
    m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True
    'Get TypeOfApplication from xml returned from call to GetApplicationData
    strTypeOfApplication = m_objXmlAssist.GetMandatoryElementText(xmlTempResponse, ".//TYPEOFAPPLICATION")
    'Get SpecialScheme from xml returned from call to GetApplicationData
    strSpecialScheme = m_objXmlAssist.GetElementText(xmlTempResponse, ".//SPECIALSCHEME")
    'Destroy all objects used in this call
    Set xmlTempRequest = Nothing
    Set xmlTempApplicationNumberElem = Nothing
    Set xmlTempApplicationFFNumber = Nothing
    Set xmlTempApplicationElem = Nothing
    Set objIApplicationBO = Nothing

    'BMIDS00755 Changed to call GetMtgAcctDataForFurtherAdvance instead
    ''--Call GetLoansNotRedeemedTotals to get all the loans in the existing account tables(s) where the redemption status is ‘NOT TO BE REDEEMED’ per customer
    ''Create all objects for this call
    'Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    'Set xmlCustList = m_objXmlAssist.GetMandatoryNode(vxmlRequest, ".//CUSTOMERLIST")
    'xmlTempRequest.appendChild xmlCustList
    'Set xmlTempResponse = objCFBO.GetLoansNotRedeemedTotals(xmlTempRequest)
    
    Set xmlTempRequest = m_objXmlAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTempApplicationElem = xmlOut.createElement("MORTGAGEACCOUNT")
    Set xmlTempApplicationNumberElem = xmlOut.createElement("APPLICATIONNUMBER")
    Set xmlTempApplicationFFNumber = xmlOut.createElement("APPLICATIONFACTFINDNUMBER")
    
    xmlTempApplicationNumberElem.Text = strApplicationNumber
    xmlTempApplicationElem.appendChild xmlTempApplicationNumberElem
    xmlTempApplicationFFNumber.Text = strApplicationFFNumber
    xmlTempApplicationElem.appendChild xmlTempApplicationFFNumber
    
    'EP2_9 GHun redundant
    ''BMIDS591 GHun 02/07/2003
    'Set xmlTempNode = xmlOut.createElement("GETBMONLYNOTTOBEREDEEMEDLOANS")
    'xmlTempNode.Text = "0"
    'xmlTempApplicationElem.appendChild xmlTempNode
    ''BMIDS591 End
    'EP2_9 End
    
    xmlTempRequest.appendChild xmlTempApplicationElem
    Set xmlTempResponse = objCFBO.GetMtgAcctDataForFurtherAdv(xmlTempRequest)
    'BMIDS00755 End
    
    'Check for error
    m_objErrAssist.CheckResponse xmlTempResponse, xmlResponseElem, True

    '------------------------Add all the total loans not redeemed of all customers------------------------
    'Get LOANSNOTREDEEMED from from MORTGAGELOAN element
    'BMIDS00755
    'dblTotLoansNotRedeemed = CSafeDbl(m_objXmlAssist.GetElementText(xmlTempResponse, ".//LOANSNOTREDEEMED"))
    dblTotLoansNotRedeemed = CSafeDbl(m_objXmlAssist.GetElementText(xmlTempResponse, ".//TOTALOUTSTANDINGBALANCE"))
    'BMIDS00755 End
    
    'Destroy all objects for this call
    Set xmlTempRequest = Nothing
    Set xmlCustList = Nothing
    Set xmlTempResponse = Nothing
    
    ' If the special Scheme is ‘BuyToLet’ (Response from ApplicationBo.GetApplicationData) AND Application Type (Response from ApplicationBo.GetApplicationData) is Further Advance or Secured Personal Loan
    
    'BS BM0310 13/05/03
    'If GetValidationTypeForValueID("SpecialSchemes", CSafeInt(strSpecialScheme)) = "BTL" Then
    If IsValidationTypeInValidationList("SpecialSchemes", "BTL", CSafeInt(strSpecialScheme)) Then
        'if Further Advance or Secured Personal Loans
        If (IsValidationTypeInValidationList("TypeOfMortgage", "F", CSafeInt(strTypeOfApplication)) _
            And IsValidationTypeInValidationList("TypeOfMortgage", "M", CSafeInt(strTypeOfApplication))) _
            Or _
            (IsValidationTypeInValidationList("TypeOfMortgage", "F", CSafeInt(strTypeOfApplication)) _
            And IsValidationTypeInValidationList("TypeOfMortgage", "S", CSafeInt(strTypeOfApplication))) Then
                dblMonthlyPayment = (((dblRequestedLoanAmount + dblTotLoansNotRedeemed) * dblBuyToLetBaseRate) / 100) / 12
            'BS BM0310 13/05/03
            Else
                dblMonthlyPayment = ((dblRequestedLoanAmount * dblBuyToLetBaseRate) / 100) / 12
        End If
    Else
            
        dblMonthlyPayment = ((dblRequestedLoanAmount * dblBuyToLetBaseRate) / 100) / 12
        
    End If
    ' Calculate Excess Monthly Rental Income using MonthlyRentalIncome (Request), Monthly Payment (Calculated above) and PercentageOfRentalIncome1 (global parameter)
    dblExcessMonthlyIncome = dblMonthlyRentalIncome - (dblMonthlyPayment * (dblPercentageOfRentalIncome1 / 100))
    
    ' Round now after all calculations have been done, and return value to the GUI
    dblExcessMonthlyIncome = Round(dblExcessMonthlyIncome, 2)
    
    Set xmlTempRequest = xmlOut.createElement("RENTALDETAILS")
    Set xmlDataNode = xmlOut.createElement("EXCESSMONTHLYRENTALINCOME")
    xmlDataNode.Text = CStr(dblExcessMonthlyIncome)
    xmlTempRequest.appendChild xmlDataNode
    xmlResponseElem.appendChild xmlTempRequest
        
    Set INewPropertyBO_CalcExcessRentalIncome = xmlResponseElem.cloneNode(True)
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
CalcExcessRentalIncomeExit:
    
    Set xmlRequestNode = Nothing
    'Set objINewPropertyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Set objIApplicationBO = Nothing
    Set xmlTempNode = Nothing
    Set xmlTempRequest = Nothing
    Set xmlTempApplicationNumberElem = Nothing
    Set xmlTempApplicationFFNumber = Nothing
    Set xmlTempApplicationElem = Nothing
    Set xmlTempResponse = Nothing
    Set objCFBO = Nothing
    Set xmlCustList = Nothing
    Exit Function

CalcExcessRentalIncomeVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set INewPropertyBO_CalcExcessRentalIncome = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume CalcExcessRentalIncomeExit

End Function

'BMIDS887 GHun
Private Function INewPropertyBO_SaveNewPropertyAndUpdateFactFind(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
   
    Const cstrFunctionName As String = "SaveNewPropertyAndUpdateFactFind"
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objINewPropertyTxBO As INewPropertyTxBO
    
On Error GoTo SaveNewPropertyAndUpdateFactFindVbErr
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If m_objContext Is Nothing Then
        Set objINewPropertyTxBO = New NewPropertyTxBO
    Else
        Set objINewPropertyTxBO = m_objContext.CreateInstance(App.Title & ".NewPropertyTxBO")
    End If
    
    Set xmlTempResponseNode = objINewPropertyTxBO.SaveNewPropertyAndUpdateFactFind(vxmlRequest)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    Set INewPropertyBO_SaveNewPropertyAndUpdateFactFind = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
SaveNewPropertyAndUpdateFactFindExit:
    Set objINewPropertyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Exit Function

SaveNewPropertyAndUpdateFactFindVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    Set INewPropertyBO_SaveNewPropertyAndUpdateFactFind = m_objErrAssist.CreateErrorResponseEx()
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    ' go to clean-up section
    Resume SaveNewPropertyAndUpdateFactFindExit

End Function
'BMIDS887 End

'BMIDS887 GHun
Public Function SaveNewPropertyAndUpdateFactFind(ByVal vstrXMLRequest As String) As String

    Const cstrFunctionName As String = "SaveNewPropertyAndUpdateFactFind"
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
On Error GoTo SaveNewPropertyAndUpdateFactFindVbErr

    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), cstrFunctionName)
                                                                                            
    ' Delegate to DOMDocument based method
    Set xmlTempResponseNode = INewPropertyBO_SaveNewPropertyAndUpdateFactFind(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    SaveNewPropertyAndUpdateFactFind = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

SaveNewPropertyAndUpdateFactFindExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Exit Function

SaveNewPropertyAndUpdateFactFindVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource cstrFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), cstrFunctionName, Err.Number, Err.Description
    End If
    
    SaveNewPropertyAndUpdateFactFind = m_objErrAssist.CreateErrorResponse
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    ' go to clean-up section
    Resume SaveNewPropertyAndUpdateFactFindExit
End Function
'BMIDS887 End

