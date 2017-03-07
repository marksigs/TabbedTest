VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ThirdPartyBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'ThirdPartyBO' with actual class name
' *****************************************************************************

'Workfile:      ThirdPartyBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for ThirdParty Business Object which 'supports transactions' only
'               i.e. any code that does not require transaction support will reside in this
'               first level (public interface) module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level) module
'               (ThirdPartyTxBO in this example)
'               contains (implied) standard interface:
'               Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'               Private Sub Validate(vstrXMLRequest As String)

'Dependencies:  List any other dependent components
'               e.g. ThirdPartyTxBO, ThirdPartyDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'IK     30/06/99    Created
'RF     30/09/99    Applied changes raised by code review of 30/09/99, including:
'                   removed AnonInterfaceFunction
'                   removed calls to Validate
'                   improved error handling
'RF     04/10/99    Added profiling
'PSC    21/10/99    Amend warning processing
'PSC    11/04/01    SYS2188 Add GetPaymentDetails
'DRC    05/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------

Option Explicit

Implements IThirdPartyBO
Implements ObjectControl

Private Const cstrROOT_NODE_NAME = "VA_PANELVALUERLIST"
Private Const cstrPAYMENTDETAILS_NAME = "PAYMENTDETAILS"
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist

'GD ADDED
'------------------------------------------------------------------------------------------
'BMIDS History:
'
'Prog   Date        Description
'PSC    03/12/2002  BM0105 - Add GetDirectoryDetails
'JD     05/10/2004  BMIDS897 added GetDirectoryGUID
'------------------------------------------------------------------------------------------

Private Const cstrROOT_NANODE_NAME = "NAMEANDADDRESSDIRECTORY"

Private Function IThirdPartyBO_FindPanelValuerList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_FindPanelValuerList"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
        
' TODO ************************************************************************
'   Remove oeNotImplemented error when the function has been implemented
' *****************************************************************************
'    m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeNotImplemented

' TODO ************************************************************************
'   replace Omiga4DO reference with actual DO required
' *****************************************************************************
    Dim objIThirdPartyDO As IThirdPartyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objIThirdPartyDO = m_objContext.CreateInstance(App.Title & ".ThirdPartyDO")
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIThirdPartyDO.FindPanelValuerList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    
    Set IThirdPartyBO_FindPanelValuerList = xmlResponseElem

FindListExit:

    Set xmlRequestNode = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IThirdPartyBO_FindPanelValuerList = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit

End Function

Private Function IThirdPartyBO_CreateDirectory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateDirectory a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateDirectoryVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_CreateDirectory"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIThirdPartyTxBO As IThirdPartyTxBO
    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseElem.setAttribute "OPERATION", "CREATE"

    If objContext Is Nothing Then
        Set objIThirdPartyTxBO = New ThirdPartyTxBO
    Else
        Set objIThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ThirdPartyTxBO")
    End If
        
    '   call Business Transaction Object CreateDirectory function
    Set xmlTempResponseNode = objIThirdPartyTxBO.CreateDirectory(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
    If Not (xmlKeysNode Is Nothing) Then
        xmlResponseElem.appendChild xmlKeysNode
    End If
   
    Set IThirdPartyBO_CreateDirectory = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objErrAssist = Nothing
    Set objIThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateDirectoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    Set objIThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_CreateDirectory = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IThirdPartyBO_CreateThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateThirdParty a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateThirdPartyVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_CreateThirdParty"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIThirdPartyTxBO As IThirdPartyTxBO
    Dim xmlKeysNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseElem.setAttribute "OPERATION", "CREATE"

    If objContext Is Nothing Then
        Set objIThirdPartyTxBO = New ThirdPartyTxBO
    Else
        Set objIThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ThirdPartyTxBO")
    End If
        
    '   call Business Transaction Object CreateThirdParty function
    Set xmlTempResponseNode = objIThirdPartyTxBO.CreateThirdParty(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
    If Not (xmlKeysNode Is Nothing) Then
        xmlResponseElem.appendChild xmlKeysNode
    End If

    Set IThirdPartyBO_CreateThirdParty = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objErrAssist = Nothing
    Set objIThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    Exit Function

CreateThirdPartyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    Set objIThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_CreateThirdParty = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IThirdPartyBO_CreateMortgageLenderDirectory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  CreateMortgageLenderDirectory a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageLenderDirectoryVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_CreateMortgageLenderDirectory"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIThirdPartyTxBO As IThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    If objContext Is Nothing Then
        Set objIThirdPartyTxBO = New ThirdPartyTxBO
    Else
        Set objIThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ThirdPartyTxBO")
    End If
        
    '   call Business Transaction Object CreateMortgageLenderDirectory function
    Set xmlTempResponseNode = objIThirdPartyTxBO.CreateMortgageLenderDirectory(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
   
    Set IThirdPartyBO_CreateMortgageLenderDirectory = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objErrAssist = Nothing
    Set objIThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

CreateMortgageLenderDirectoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    Set objIThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_CreateMortgageLenderDirectory = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IThirdPartyBO_DeleteThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  DeleteThirdParty a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo DeleteThirdPartyVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_DeleteThirdParty"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIThirdPartyTxBO As IThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseElem.setAttribute "OPERATION", "DELETE"
    
    If objContext Is Nothing Then
        Set objIThirdPartyTxBO = New ThirdPartyTxBO
    Else
        Set objIThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ThirdPartyTxBO")
    End If
           
    '   call Business Transaction Object DeleteThirdParty function
    Set xmlTempResponseNode = objIThirdPartyTxBO.DeleteThirdParty(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IThirdPartyBO_DeleteThirdParty = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objErrAssist = Nothing
    Set objIThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteThirdPartyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objIThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
       
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_DeleteThirdParty = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IThirdPartyBO_FindDirectoryList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindDirectoryListVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_FindDirectoryList"


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
    Dim objIThirdPartyDO As IThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIThirdPartyDO = New ThirdPartyDO
    Else
        Set objIThirdPartyDO = objContext.CreateInstance(App.Title & ".ThirdPartyDO")
    End If
                
    If vxmlRequest.nodeName = "NAMEANDADDRESSDIRECTORY" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("NAMEANDADDRESSDIRECTORY").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "NAMEANDADDRESSDIRECTORY tag not found"
    End If
        
    Set xmlDataNode = objIThirdPartyDO.FindDirectoryList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set IThirdPartyBO_FindDirectoryList = xmlResponseElem.cloneNode(True)
    objXmlAssist.WriteXMLToFile xmlResponseElem.xml

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindDirectoryListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_FindDirectoryList = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IThirdPartyBO_FindDirectorySummaryList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindDirectorySummaryListVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_FindDirectorySummaryList"


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
    Dim objIThirdPartyDO As IThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIThirdPartyDO = New ThirdPartyDO
    Else
        Set objIThirdPartyDO = objContext.CreateInstance(App.Title & ".ThirdPartyDO")
    End If
                
    If vxmlRequest.nodeName = "NAMEANDADDRESSDIRECTORY" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("NAMEANDADDRESSDIRECTORY").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "NAMEANDADDRESSDIRECTORY tag not found"
    End If
        
    Set xmlDataNode = objIThirdPartyDO.FindDirectorySummaryList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set IThirdPartyBO_FindDirectorySummaryList = xmlResponseElem.cloneNode(True)

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindDirectorySummaryListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_FindDirectorySummaryList = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IThirdPartyBO_FindMainMortgageLenderList() As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindMainMortgageLenderListVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_FindMainMortgageLenderList"


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
    Dim objIThirdPartyDO As IThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIThirdPartyDO = New ThirdPartyDO
    Else
        Set objIThirdPartyDO = objContext.CreateInstance(App.Title & ".ThirdPartyDO")
    End If
        
    Set xmlDataNode = objIThirdPartyDO.FindMainMortgageLenderList()
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set IThirdPartyBO_FindMainMortgageLenderList = xmlResponseElem.cloneNode(True)

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindMainMortgageLenderListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_FindMainMortgageLenderList = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IThirdPartyBO_FindThirdPartyList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindThirdPartyListVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_FindThirdPartyList"


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
    Dim objIThirdPartyDO As IThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If objContext Is Nothing Then
        Set objIThirdPartyDO = New ThirdPartyDO
    Else
        Set objIThirdPartyDO = objContext.CreateInstance(App.Title & ".ThirdPartyDO")
    End If
                
    If vxmlRequest.nodeName = "THIRDPARTY" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("THIRDPARTY").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "THIRDPARTY tag not found"
    End If
        
    Set xmlDataNode = objIThirdPartyDO.FindThirdPartyList(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set IThirdPartyBO_FindThirdPartyList = xmlResponseElem.cloneNode(True)

    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

FindThirdPartyListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_FindThirdPartyList = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IThirdPartyBO_GetThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetThirdPartyVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_GetThirdParty"


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
    Dim objIThirdPartyDO As IThirdPartyDO
    Dim blnCreateDirectory As Boolean
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIThirdPartyDO = New ThirdPartyDO
    Else
        Set objIThirdPartyDO = objContext.CreateInstance(App.Title & ".ThirdPartyDO")
    End If

    blnCreateDirectory = False
                
    If vxmlRequest.nodeName = "THIRDPARTY" Then
        Set xmlRequestNode = vxmlRequest
    ElseIf vxmlRequest.nodeName = "NAMEANDADDRESSDIRECTORY" Then
        Set xmlRequestNode = vxmlRequest
        blnCreateDirectory = True
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("THIRDPARTY").Item(0)
    
        If xmlRequestNode Is Nothing Then
            Set xmlRequestNode = vxmlRequest.getElementsByTagName("NAMEANDADDRESSDIRECTORY").Item(0)
            blnCreateDirectory = True
        End If
    End If

    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "THIRDPARTY" & " tag not found"
    End If

    If blnCreateDirectory Then
        Set xmlDataNode = objIThirdPartyDO.GetDirectory(xmlRequestNode)
    Else
        Set xmlDataNode = objIThirdPartyDO.GetThirdParty(xmlRequestNode)
    End If

    xmlResponseElem.appendChild xmlDataNode
        
    Set IThirdPartyBO_GetThirdParty = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetThirdPartyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_GetThirdParty = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IThirdPartyBO_GetDirectoryGUID(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetDirectoryGUIDVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_GetDirectoryGUID"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIThirdPartyDO As IThirdPartyDO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If objContext Is Nothing Then
        Set objIThirdPartyDO = New ThirdPartyDO
    Else
        Set objIThirdPartyDO = objContext.CreateInstance(App.Title & ".ThirdPartyDO")
    End If

    Set xmlDataNode = objIThirdPartyDO.GetDirectoryGUID(vxmlRequest)
    
    xmlResponseElem.appendChild xmlDataNode
        
    Set IThirdPartyBO_GetDirectoryGUID = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objErrAssist = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing

    Exit Function

GetDirectoryGUIDVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_GetDirectoryGUID = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IThirdPartyBO_SaveThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  SaveThirdParty a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo SaveThirdPartyVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_SaveThirdParty"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    Dim botOperation As BOOPERATIONTYPE

    Set objIThirdPartyClassDef = New ThirdPartyClassDef

    botOperation = objXmlAssist.DetermineOperationEx(vxmlRequest, objIThirdPartyClassDef.LoadThirdPartyData)

    Select Case botOperation
        Case booCreate: Set IThirdPartyBO_SaveThirdParty = IThirdPartyBO_CreateThirdParty(vxmlRequest)
        Case booUpdate: Set IThirdPartyBO_SaveThirdParty = IThirdPartyBO_UpdateThirdParty(vxmlRequest)
        Case booDelete: Set IThirdPartyBO_SaveThirdParty = IThirdPartyBO_DeleteThirdParty(vxmlRequest)
        Case Else
            ' N/A
    End Select

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIThirdPartyClassDef = Nothing
    
    Exit Function

SaveThirdPartyVbErr:
    Set objIThirdPartyClassDef = Nothing
    Set objXmlAssist = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_SaveThirdParty = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IThirdPartyBO_UpdateThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateThirdParty a single instance of the persistant data associated with this
'               business object
'               invokes appropriate transactioning (Tx) business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateThirdPartyVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_UpdateThirdParty"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim objIThirdPartyTxBO As IThirdPartyTxBO
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseElem.setAttribute "OPERATION", "UPDATE"

    If objContext Is Nothing Then
        Set objIThirdPartyTxBO = New ThirdPartyTxBO
    Else
        Set objIThirdPartyTxBO = objContext.CreateInstance(App.Title & ".ThirdPartyTxBO")
    End If
    
    Set xmlTempResponseNode = objIThirdPartyTxBO.UpdateThirdParty(vxmlRequest)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    Set IThirdPartyBO_UpdateThirdParty = xmlResponseElem.cloneNode(True)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objErrAssist = Nothing
    Set objIThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

UpdateThirdPartyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set objIThirdPartyTxBO = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
        
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_UpdateThirdParty = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IThirdPartyBO_ValidateDirectory(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateDirectoryVbErr

    Const strFunctionName As String = "IThirdPartyBO_ValidateDirectory"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim objComboDO As New ComboDO
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strValue As String
    Dim strXML As String
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '
    ' Company name must be specified
    '
    strValue = objXmlAssist.GetTagValue(vxmlRequest, "COMPANYNAME")
    If strValue = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 259
    End If

    '
    ' Name and address type must be specified
    '
    strValue = objXmlAssist.GetTagValue(vxmlRequest, "NAMEANDADDRESSTYPE")
    If strValue = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 258
    End If

    '
    ' Name and address type must be valid
    '
    strXML = objComboDO.GetComboValue(objXmlAssist.BuildComboValueList("ThirdPartyType", strValue))
    If strXML = "" Then
        ' 'Record not found' error must have occured in the GetComboValue function
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 258
    Else
        Set xmlDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
        If objXmlAssist.GetTagValue(xmlDoc.documentElement, "VALUENAME") = "" Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 258
        End If

        If objXmlAssist.GetTagValue(xmlDoc.documentElement, "VALIDATIONTYPE") = "P" Then
            If objXmlAssist.GetTagValue(vxmlRequest, "PANELID") = "" Then
                objErrAssist.ThrowError TypeName(Me), strFunctionName, 260
            End If
        End If
    End If

    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' call scripting interface
    '-----------------------------------------------------------------------------
    Set IThirdPartyBO_ValidateDirectory = xmlResponseElem.cloneNode(True)

    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateDirectoryVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_ValidateDirectory = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IThirdPartyBO_ValidateThirdParty(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo ValidateThirdPartyVbErr

    Const strFunctionName As String = "IThirdPartyBO_ValidateThirdParty"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlDataNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement

    Dim strValue As String
       
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    '
    ' Third party type must be specified
    '
    strValue = objXmlAssist.GetTagValue(vxmlRequest, "THIRDPARTYTYPE")
    If strValue = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 258
    End If

    '
    ' Company name must be specified
    '
    strValue = objXmlAssist.GetTagValue(vxmlRequest, "COMPANYNAME")
    If strValue = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 259
    End If

    Set IThirdPartyBO_ValidateThirdParty = xmlResponseElem.cloneNode(True)

    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

ValidateThirdPartyVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    Set xmlDataNode = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IThirdPartyBO_ValidateThirdParty = objErrAssist.CreateErrorResponseEx()
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function FindPanelValuerList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindListVbErr

    Const strFunctionName As String = "FindPanelValuerList"
    
       
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
'    Dim xmlList As IXMLDOMElement
'    Dim xmlNewList As IXMLDOMElement
'    Dim iTotal As Integer
'    Dim iNewIndex As Integer
'    Dim iIndex As Integer
'    Dim iNewListTotal As Integer
'
'    Dim xmlCurrentRowElement As IXMLDOMElement
'    Dim xmlCurrentNewRowElement As IXMLDOMElement
'
'    Dim xmlChildToRemove As IXMLDOMNode
'    Dim xmlChildToAdd As IXMLDOMNode
'    Dim xmlNewRowElement As IXMLDOMElement
'    Dim xmlTempNode As IXMLDOMNode
'    Dim xmlTempElement As IXMLDOMElement
'    Dim xmlTempElement1 As IXMLDOMElement
'
'    Dim strCurrentDirectoryGUID As String
'    Dim strNewDirectoryGUID As String
'
'    Dim strValuationTypeText As String
'    Dim strValuationType As String
'
'    Dim blnResult As Boolean
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IThirdPartyBO_FindPanelValuerList(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    FindPanelValuerList = xmlResponseElem.xml
    
    m_objContext.SetComplete

FindListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    
    Exit Function

FindListVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindPanelValuerList = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume FindListExit

End Function

Public Function CreateDirectory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateDirectory a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateDirectoryVbErr

    Const strFunctionName As String = "CreateDirectory"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    Dim xmlKeysNode As IXMLDOMNode
    
    ' CreateDirectory default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseElem.setAttribute "OPERATION", "CREATE"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IThirdPartyBO_ValidateDirectory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IThirdPartyBO_CreateDirectory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True

    Set xmlKeysNode = xmlTempResponseNode.selectSingleNode("GENERATEDKEYS")
    If Not (xmlKeysNode Is Nothing) Then
        xmlResponseElem.appendChild xmlKeysNode
    End If
    
    CreateDirectory = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateDirectoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Set xmlKeysNode = Nothing

    Exit Function

CreateDirectoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateDirectory = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateDirectoryExit

End Function

Public Function CreateMortgageLenderDirectory(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateMortgageLenderDirectory a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageLenderDirectoryVbErr

    Const strFunctionName As String = "CreateMortgageLenderDirectory"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' CreateMortgageLenderDirectory default response block
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
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IThirdPartyBO_CreateMortgageLenderDirectory(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateMortgageLenderDirectory = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateMortgageLenderDirectoryExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateMortgageLenderDirectoryVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateMortgageLenderDirectory = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateMortgageLenderDirectoryExit

End Function

Public Function CreateThirdParty(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  CreateThirdParty a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CreateThirdPartyVbErr

    Const strFunctionName As String = "CreateThirdParty"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIScriptInterface As IScriptInterface
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' CreateThirdParty default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    xmlResponseElem.setAttribute "OPERATION", "CREATE"

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IThirdPartyBO_ValidateThirdParty(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IThirdPartyBO_CreateThirdParty(xmlIn.documentElement)
    Dim lngRet As Long
    lngRet = objErrAssist.CheckXMLResponse(xmlTempResponseNode.xml, True)
    'objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    CreateThirdParty = xmlTempResponseNode.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

CreateThirdPartyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

CreateThirdPartyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    CreateThirdParty = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CreateThirdPartyExit

End Function

Public Function DeleteThirdParty(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  DeleteThirdParty a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DeleteThirdPartyVbErr

    Const strFunctionName As String = "DeleteThirdParty"

    
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
    xmlResponseElem.setAttribute "OPERATION", "DELETE"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IThirdPartyBO_DeleteThirdParty(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
        
    DeleteThirdParty = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

DeleteThirdPartyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

DeleteThirdPartyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    DeleteThirdParty = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DeleteThirdPartyExit

End Function

Public Function FindDirectoryList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindDirectoryListVbErr

    Const strFunctionName As String = "FindDirectoryList"

    
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
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IThirdPartyBO_FindDirectoryList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    FindDirectoryList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindDirectoryListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindDirectoryListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindDirectoryList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindDirectoryListExit

End Function

Public Function FindDirectorySummaryList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindDirectorySummaryListVbErr

    Const strFunctionName As String = "FindDirectorySummaryList"

    
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
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IThirdPartyBO_FindDirectorySummaryList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    FindDirectorySummaryList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindDirectorySummaryListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindDirectorySummaryListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindDirectorySummaryList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindDirectorySummaryListExit

End Function

Public Function FindMainMortgageLenderList() As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindMainMortgageLenderListVbErr

    Const strFunctionName As String = "FindMainMortgageLenderList"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
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
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IThirdPartyBO_FindMainMortgageLenderList()
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    FindMainMortgageLenderList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindMainMortgageLenderListExit:

    '   clean-up section, release object references
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindMainMortgageLenderListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindMainMortgageLenderList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindMainMortgageLenderListExit

End Function

Public Function FindThirdPartyList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindThirdPartyListVbErr

    Const strFunctionName As String = "FindThirdPartyList"

    
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
                                        
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IThirdPartyBO_FindThirdPartyList(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    FindThirdPartyList = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

FindThirdPartyListExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

FindThirdPartyListVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    FindThirdPartyList = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume FindThirdPartyListExit

End Function

Public Function GetThirdParty(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetThirdPartyVbErr

    Const strFunctionName As String = "GetThirdParty"

    
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
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IThirdPartyBO_GetThirdParty(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetThirdParty = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetThirdPartyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetThirdPartyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetThirdParty = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetThirdPartyExit

End Function
Public Function GetDirectoryGUID(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  return the GUID for the nameandaddressdirectory with this companyname
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetDirectoryGUIDVbErr

    Const strFunctionName As String = "GetDirectoryGUID"

    
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
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    Set xmlTempResponseNode = IThirdPartyBO_GetDirectoryGUID(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode

    objXmlAssist.WriteXMLToFile xmlResponseElem.xml
    GetDirectoryGUID = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

GetDirectoryGUIDExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetDirectoryGUIDVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetDirectoryGUID = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetDirectoryGUIDExit

End Function
Public Function SaveThirdParty(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  SaveThirdParty a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo SaveThirdPartyVbErr

    Const strFunctionName As String = "SaveThirdParty"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn  As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlResponseNode As IXMLDOMNode

    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set xmlResponseNode = IThirdPartyBO_SaveThirdParty(xmlIn.documentElement)
    
    SaveThirdParty = xmlResponseNode.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

SaveThirdPartyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlResponseNode = Nothing

    Exit Function

SaveThirdPartyVbErr:
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    SaveThirdParty = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume SaveThirdPartyExit

End Function

Public Function UpdateThirdParty(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  UpdateThirdParty a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo UpdateThirdPartyVbErr

    Const strFunctionName As String = "UpdateThirdParty"

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
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
    xmlResponseElem.setAttribute "OPERATION", "UPDATE"
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Run validation checks
    Set xmlTempResponseNode = IThirdPartyBO_ValidateThirdParty(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    ' Run scripting interface
    Set objIScriptInterface = New ScriptInterface
    Set xmlTempResponseNode = objIScriptInterface.RunScript(App.Title, _
                                                            TypeName(Me), _
                                                            strFunctionName, _
                                                            xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                                                                           
    ' Delegate to DomDocument based method
    Set xmlTempResponseNode = IThirdPartyBO_UpdateThirdParty(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    
    UpdateThirdParty = xmlResponseElem.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

UpdateThirdPartyExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objIScriptInterface = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

UpdateThirdPartyVbErr:

    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    UpdateThirdParty = objErrAssist.CreateErrorResponse

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateThirdPartyExit
    
End Function

Public Function GetPaymentDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the payment details given a directory guid
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetPaymentDetailsVbErr

    Const strFunctionName As String = "GetPaymentDetails"
        
        
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
    Set xmlTempResponseNode = IThirdPartyBO_GetPaymentDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetPaymentDetails = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetPaymentDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetPaymentDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetPaymentDetails = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetPaymentDetailsExit

End Function

Private Function IThirdPartyBO_GetPaymentDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets the payment details given a directory guid
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetPaymentDetailsVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_GetPaymentDetails"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIThirdPartyDO As IThirdPartyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objIThirdPartyDO = m_objContext.CreateInstance(App.Title & ".ThirdPartyDO")
                
    If vxmlRequest.nodeName = cstrPAYMENTDETAILS_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrPAYMENTDETAILS_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrPAYMENTDETAILS_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIThirdPartyDO.GetPaymentDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IThirdPartyBO_GetPaymentDetails = xmlResponseElem
    
    m_objContext.SetComplete
    
GetPaymentDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

GetPaymentDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IThirdPartyBO_GetPaymentDetails = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetPaymentDetailsExit

End Function

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    'Set m_objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False

End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
    'Set m_objIDOAssist = Nothing

End Sub

' PSC 03/12/2002 BM0105 - Start
Public Function GetDirectoryDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetDirectoryDetailsVbErr

    Const strFunctionName As String = "GetDirectoryDetails"
        
        
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
    Set xmlTempResponseNode = IThirdPartyBO_GetDirectoryDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetDirectoryDetails = xmlResponseElem.xml
    
    m_objContext.SetComplete

GetDirectoryDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    
    Exit Function

GetDirectoryDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetDirectoryDetails = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDirectoryDetailsExit

End Function

Private Function IThirdPartyBO_GetDirectoryDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetDirectoryDetailsVbErr
    
    Const strFunctionName As String = "IThirdPartyBO_GetDirectoryDetails"

    
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIThirdPartyDO As IThirdPartyDO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objIThirdPartyDO = m_objContext.CreateInstance(App.Title & ".ThirdPartyDO")
                
    If vxmlRequest.nodeName = cstrROOT_NANODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NANODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NANODE_NAME & " tag not found"
    End If
    
    
        
    Set xmlDataNode = objIThirdPartyDO.GetDirectoryDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IThirdPartyBO_GetDirectoryDetails = xmlResponseElem
    
    m_objContext.SetComplete
    
GetDirectoryDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objIThirdPartyDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    

    Exit Function

GetDirectoryDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IThirdPartyBO_GetDirectoryDetails = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume GetDirectoryDetailsExit

End Function
' PSC 03/12/2002 BM0105 - End

