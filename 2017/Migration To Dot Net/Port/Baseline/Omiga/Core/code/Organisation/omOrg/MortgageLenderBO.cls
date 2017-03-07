VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "MortgageLenderBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      MortgageLenderBO.cls
'Copyright:     Copyright © 2000 Marlborough Stirling
'Description:   Code template for omiga4 Business Object which 'supports transactions' only
'               i.e. any code that does not require transaction support will reside in this
'               first level (public interface) module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level) module
'               (OrganisationTxBO in this example)
'               contains (implied) standard interface:
'               Public Function Create(vstrXMLRequest As String) As String
'               Public Function Update(vstrXMLRequest As String) As String
'               Public Function Delete(vstrXMLRequest As String) As String
'               Public Function GetData(vstrXMLRequest As String) As String
'               Private Sub Validate(vstrXMLRequest As String)
'Dependencies:  MortgageLenderTxBO, MortgageLenderDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' MV    24/11/00    Created.
' CL    28/11/00    Method FindLenderNameAndIncomeMultiples added
' PSC   12/01/01    SYS1753 Added GetMainLenderDirectoryGuid()
' DRC   12/06/01    Method GetLenderDetails added
' DRC   02/10/01    AQR SYS2696 GetLenderDetails modifed
' DRC   05/10/01    SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
' SAB   20/09/05    CORE155 Added GetApplicationMortgageLender
'------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'Baseline specific History:
'
' Prog      Date        AQR         Description
' MV        20/11/2003  CORE6       Added GetPanelLenderDetails(); IMortgageLenderBO_GetPanelLenderDetails()
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------
'BBG Change History:
'---------------------
' Prog      Date        AQR         Description
'
' AW        14/09/04    BBG1416     Added GetLenderTelephoneDetails
' AW        14/09/04    BBG1436     Added GetLenderNameDirectoryDetails
'
'-----------------------------------------------------------------------------------------------
Option Explicit
Implements IMortgageLenderBO
Implements ObjectControl
 
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Private m_objIDOAssist As IDOAssist
Private Const cstrROOT_NODE_NAME As String = "MORTGAGELENDER"
Private Const cstrAllowableIncomeFactors_Name As String = "ALLOWABLEINCOMEFACTORS"
Public Function FindAllowableIncomeFactors(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with the table
'               'AllowableIncomeFactors'
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindAllowableIncomeFactorsVbErr
    Const strFunctionName As String = "FindAllowableIncomeFactors"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageLenderBO_FindAllowableIncomeFactors(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    FindAllowableIncomeFactors = xmlResponseElem.xml
    m_objContext.SetComplete
FindAllowableIncomeFactorsExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Exit Function
FindAllowableIncomeFactorsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    FindAllowableIncomeFactors = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindAllowableIncomeFactorsExit
End Function
Public Function GetPanelLenderDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the data for a lender given a lender DirectoryGUID
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetPanelLenderDetailsVbErr
    Const strFunctionName As String = "GetPanelLenderDetails"
        
    ' <VSA>
    ' Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As vsa_shared
        Set VSA = New vsa_shared
        VSA.Initialize (strFunctionName)
    #End If
    ' </VSA>
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageLenderBO_GetPanelLenderDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    GetPanelLenderDetails = xmlResponseElem.xml
    m_objContext.SetComplete
GetPanelLenderDetailsExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    ' <VSA>
    ' Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    ' </VSA>
    Exit Function
GetPanelLenderDetailsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    GetPanelLenderDetails = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume GetPanelLenderDetailsExit
End Function
Public Function FindLenderNameAndIncomeMultiples(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
 On Error GoTo FindLenderNameAndIncomeMultiplesVbErr
    Const strFunctionName As String = "FindLenderNameAndIncomeMultiples"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageLenderBO_FindLenderNameAndIncomeMultiples(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    FindLenderNameAndIncomeMultiples = xmlResponseElem.xml
    m_objContext.SetComplete
FindLenderNameAndIncomeMultiplesExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Exit Function
FindLenderNameAndIncomeMultiplesVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    FindLenderNameAndIncomeMultiples = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindLenderNameAndIncomeMultiplesExit
End Function
Public Function FindLenderNames(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Retrieves all instances of mortgage products that match the required
'               search criteria
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
 On Error GoTo FindLenderNamesVbErr
    Dim strFunctionName As String
    strFunctionName = "FindLenderNames"
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageLenderBO_FindLenderNames(xmlIn.documentElement)
    objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    FindLenderNames = xmlResponseElem.xml
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
FindLenderNamesExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Exit Function
FindLenderNamesVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    objErrAssist.AddToErrSource strFunctionName
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    FindLenderNames = objErrAssist.CreateErrorResponse
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    ' go to clean-up section
    Resume FindLenderNamesExit
End Function
Public Function GetMainLenderDirectoryGuid(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the directory guid for the main lender given a lender code
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetMainLenderDirectoryGuidVbErr
    Const strFunctionName As String = "GetMainLenderDirectoryGuid"
        
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageLenderBO_GetMainLenderDirectoryGuid(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    GetMainLenderDirectoryGuid = xmlResponseElem.xml
    m_objContext.SetComplete
GetMainLenderDirectoryGuidExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Exit Function
GetMainLenderDirectoryGuidVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    GetMainLenderDirectoryGuid = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume GetMainLenderDirectoryGuidExit
End Function
Public Function GetLenderDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the data for a lender given a lender organisationID
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetLenderDetailsVbErr
    Const strFunctionName As String = "GetLenderDetails"
        
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageLenderBO_GetLenderDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    GetLenderDetails = xmlResponseElem.xml
    m_objContext.SetComplete
GetLenderDetailsExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Exit Function
GetLenderDetailsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    GetLenderDetails = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume GetLenderDetailsExit
End Function
Private Function IMortgageLenderBO_FindAllowableIncomeFactors( _
                        ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
On Error GoTo FindAllowableIncomeFactorsVbErr
    
    Const strFunctionName As String = "IMortgageLenderBO_FindAllowableIncomeFactors"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIMortgageLenderDO As IMortgageLenderDO
       
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set objIMortgageLenderDO = m_objContext.CreateInstance(App.Title & ".MortgageLenderDO")
                
    If vxmlRequest.nodeName = cstrAllowableIncomeFactors_Name Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrAllowableIncomeFactors_Name).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrAllowableIncomeFactors_Name & " tag not found"
    End If
        
    Set xmlDataNode = objIMortgageLenderDO.FindAllowableIncomeFactors(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    Set IMortgageLenderBO_FindAllowableIncomeFactors = xmlResponseElem
FindAllowableIncomeFactorsExit:
    Set xmlRequestNode = Nothing
    Set objIMortgageLenderDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
FindAllowableIncomeFactorsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IMortgageLenderBO_FindAllowableIncomeFactors = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindAllowableIncomeFactorsExit
End Function
Private Function IMortgageLenderBO_FindLenderNameAndIncomeMultiples(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindLenderNameAndIncomeMultiplesVbErr
    
    Const strFunctionName As String = "IMortgageLenderBO_FindLenderNameAndIncomeMultiples"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIMortgageLenderDO As IMortgageLenderDO
        
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set objIMortgageLenderDO = m_objContext.CreateInstance(App.Title & ".MortgageLenderDO")
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIMortgageLenderDO.FindLenderNameAndIncomeMultiples(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    Set IMortgageLenderBO_FindLenderNameAndIncomeMultiples = xmlResponseElem
FindLenderNameAndIncomeMultiplesExit:
    Set xmlRequestNode = Nothing
    Set objIMortgageLenderDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
FindLenderNameAndIncomeMultiplesVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IMortgageLenderBO_FindLenderNameAndIncomeMultiples = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindLenderNameAndIncomeMultiplesExit
End Function
Private Function IMortgageLenderBO_FindLenderNames(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Retrieves all instances of mortgage products that match the required
'               search criteria
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindLenderNamesVbErr
    
    Dim strFunctionName As String
    strFunctionName = "IMortgageLenderBO_FindLenderNames"
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objIMortgageLenderDO As IMortgageLenderDO
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    If m_objContext Is Nothing Then
        Set objIMortgageLenderDO = New MortgageLenderDO
    Else
        Set objIMortgageLenderDO = m_objContext.CreateInstance(App.Title & ".MortgageLenderDO")
    End If
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
    Set xmlDataNode = objIMortgageLenderDO.FindLenderNames(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Set IMortgageLenderBO_FindLenderNames = xmlResponseElem.cloneNode(True)
FindLenderNamesExit:
    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
    Set objErrAssist = Nothing
    Set objIMortgageLenderDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
FindLenderNamesVbErr:
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    objErrAssist.AddToErrSource strFunctionName
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IMortgageLenderBO_FindLenderNames = objErrAssist.CreateErrorResponseEx()
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    ' go to clean-up section
    Resume FindLenderNamesExit
End Function

Private Function IMortgageLenderBO_GetLenderNameDirectoryDetails(ByVal vxmlRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets the telephone, fax details for the main lender given a directory guid
'
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetLenderNameDirectoryDetailsVbErr
    
    Const strFunctionName As String = "IMortgageLenderBO_GetLenderNameDirectoryDetails"
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIMortgageLenderDO As IMortgageLenderDO
        
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objIMortgageLenderDO = m_objContext.CreateInstance(App.Title & ".MortgageLenderDO")
    
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIMortgageLenderDO.GetLenderNameDirectoryDetails(xmlRequestNode)
    m_objErrAssist.CheckResponse xmlDataNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlDataNode
        
    Set IMortgageLenderBO_GetLenderNameDirectoryDetails = xmlResponseElem
    
    m_objContext.SetComplete
    
GetLenderNameDirectoryDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objIMortgageLenderDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function
    
GetLenderNameDirectoryDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IMortgageLenderBO_GetLenderNameDirectoryDetails = m_objErrAssist.CreateErrorResponseEx()
    
    m_objContext.SetAbort
    ' go to clean-up section
    Resume GetLenderNameDirectoryDetailsExit
End Function

' AW        14/09/04    BBG1416
Private Function IMortgageLenderBO_GetLenderTelephoneDetails(ByVal vxmlRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets the telephone, fax details for the main lender given a directory guid
'
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetLenderTelephoneDetailsVbErr
    
    Const strFunctionName As String = "IMortgageLenderBO_GetLenderTelephoneDetails"
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIMortgageLenderDO As IMortgageLenderDO
        
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set objIMortgageLenderDO = m_objContext.CreateInstance(App.Title & ".MortgageLenderDO")
    
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIMortgageLenderDO.GetLenderTelephoneDetails(xmlRequestNode)
    m_objErrAssist.CheckResponse xmlDataNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlDataNode
        
    Set IMortgageLenderBO_GetLenderTelephoneDetails = xmlResponseElem
    
    m_objContext.SetComplete
    
GetLenderTelephoneDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objIMortgageLenderDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function
    
GetLenderTelephoneDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set IMortgageLenderBO_GetLenderTelephoneDetails = m_objErrAssist.CreateErrorResponseEx()
    
    m_objContext.SetAbort
    ' go to clean-up section
    Resume GetLenderTelephoneDetailsExit
End Function

Private Function IMortgageLenderBO_GetMainLenderDirectoryGuid(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets the directory guid for the main lender given a lender code
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetMainLenderDirectoryGuidVbErr
    
    Const strFunctionName As String = "IMortgageLenderBO_GetMainLenderDirectoryGuid"
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIMortgageLenderDO As IMortgageLenderDO
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    Set objIMortgageLenderDO = m_objContext.CreateInstance(App.Title & ".MortgageLenderDO")
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIMortgageLenderDO.GetMainLenderDirectoryGuid(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IMortgageLenderBO_GetMainLenderDirectoryGuid = xmlResponseElem
    m_objContext.SetComplete
GetMainLenderDirectoryGuidExit:
    
    Set xmlRequestNode = Nothing
    Set objIMortgageLenderDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
GetMainLenderDirectoryGuidVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IMortgageLenderBO_GetMainLenderDirectoryGuid = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume GetMainLenderDirectoryGuidExit
End Function
Private Function IMortgageLenderBO_GetLenderDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets the directory guid for the main lender given a lender code
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetLenderDetailsVbErr
    
    Const strFunctionName As String = "IMortgageLenderBO_GetLenderDetails"
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIMortgageLenderDO As IMortgageLenderDO
        
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set objIMortgageLenderDO = m_objContext.CreateInstance(App.Title & ".MortgageLenderDO")
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIMortgageLenderDO.GetLenderDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IMortgageLenderBO_GetLenderDetails = xmlResponseElem
    m_objContext.SetComplete
GetLenderDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objIMortgageLenderDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
GetLenderDetailsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IMortgageLenderBO_GetLenderDetails = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume GetLenderDetailsExit
End Function
Private Function IMortgageLenderBO_GetPanelLenderDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Gets the directory guid for the main lender given a lender code
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetPanelLenderDetailsVbErr
    
    Const strFunctionName As String = "IMortgageLenderBO_GetPanelLenderDetails"
    ' <VSA>
    ' Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As vsa_shared
        Set VSA = New vsa_shared
        VSA.Initialize (strFunctionName)
    #End If
    ' </VSA>
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIMortgageLenderDO As IMortgageLenderDO
        
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set objIMortgageLenderDO = m_objContext.CreateInstance(App.Title & ".MortgageLenderDO")
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    Set xmlDataNode = objIMortgageLenderDO.GetPanelLenderDetails(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IMortgageLenderBO_GetPanelLenderDetails = xmlResponseElem
    m_objContext.SetComplete
GetPanelLenderDetailsExit:
    
    Set xmlRequestNode = Nothing
    Set objIMortgageLenderDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    ' <VSA>
    ' Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    ' </VSA>
    Exit Function
GetPanelLenderDetailsVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IMortgageLenderBO_GetPanelLenderDetails = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume GetPanelLenderDetailsExit
End Function
Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    Set m_objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
    Set m_objIDOAssist = Nothing
End Sub
' AW        14/09/04    BBG1416
Public Function GetLenderTelephoneDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the data for a lender given a lender Directory guid
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetLenderTelephoneDetailsVbErr
    Const strFunctionName As String = "GetLenderTelephoneDetails"
        
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageLenderBO_GetLenderTelephoneDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetLenderTelephoneDetails = xmlResponseElem.xml
    
    m_objContext.SetComplete
    
GetLenderTelephoneDetailsExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function
    
GetLenderTelephoneDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    GetLenderTelephoneDetails = m_objErrAssist.CreateErrorResponse
    
    m_objContext.SetAbort
    ' go to clean-up section
    Resume GetLenderTelephoneDetailsExit
End Function

Public Function GetLenderNameDirectoryDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the data for a lender given a lender Directory guid
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetLenderNameDirectoryDetailsVbErr
    Const strFunctionName As String = "GetLenderNameDirectoryDetails"
        
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = IMortgageLenderBO_GetLenderNameDirectoryDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetLenderNameDirectoryDetails = xmlResponseElem.xml
    
    m_objContext.SetComplete
    
GetLenderNameDirectoryDetailsExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function
    
GetLenderNameDirectoryDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    GetLenderNameDirectoryDetails = m_objErrAssist.CreateErrorResponse
    
    m_objContext.SetAbort
    ' go to clean-up section
    Resume GetLenderNameDirectoryDetailsExit
End Function

' CORE155 - New Method
Public Function GetApplicationMortgageLender(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the data for a lender given a lender organisationID
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo ErrorGetApplicationMortgageLender

    Const strFunctionName As String = "GetApplicationMortgageLender"
        
    Dim xmlIn               As FreeThreadedDOMDocument40
    Dim xmlOut              As FreeThreadedDOMDocument40
    Dim xmlResponseElem     As IXMLDOMElement
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
    Set xmlTempResponseNode = IMortgageLenderBO_GetApplicationMortgageLender(xmlIn.documentElement)
    
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    
    GetApplicationMortgageLender = xmlResponseElem.xml
    
    m_objContext.SetComplete

ExitGetApplicationMortgageLender:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    
    Exit Function

ErrorGetApplicationMortgageLender:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetApplicationMortgageLender = m_objErrAssist.CreateErrorResponse

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume ExitGetApplicationMortgageLender

End Function

Private Function IMortgageLenderBO_GetApplicationMortgageLender(ByVal vxmlRequest As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMNode
    On Error GoTo ErrorGetApplicationMortgageLender
    
    Const strFunctionName As String = "IMortgageLenderBO_GetApplicationMortgageLender"
    
    Dim xmlRequestNode          As IXMLDOMNode
    Dim xmlDataNode             As IXMLDOMNode
    Dim xmlOut                  As FreeThreadedDOMDocument40
    Dim xmlResponseElem         As IXMLDOMElement
    Dim objIMortgageLenderDO    As IMortgageLenderDO
   
        
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    If Not m_objContext Is Nothing Then
        Set objIMortgageLenderDO = m_objContext.CreateInstance(App.Title & ".MortgageLenderDO")
    Else
        Set objIMortgageLenderDO = CreateObject(App.Title & ".MortgageLenderDO")
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
        
    Set xmlDataNode = objIMortgageLenderDO.GetApplicationMortgageLender(xmlRequestNode)
    xmlResponseElem.appendChild xmlDataNode
        
    Set IMortgageLenderBO_GetApplicationMortgageLender = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
ExitGetApplicationMortgageLender:
    Set xmlRequestNode = Nothing
    Set objIMortgageLenderDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function

ErrorGetApplicationMortgageLender:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IMortgageLenderBO_GetApplicationMortgageLender = m_objErrAssist.CreateErrorResponseEx()

    m_objContext.SetComplete
    
    ' go to clean-up section
    Resume ExitGetApplicationMortgageLender

End Function
' CORE155 - End
