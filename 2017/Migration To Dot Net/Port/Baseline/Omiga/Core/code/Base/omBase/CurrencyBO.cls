VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CurrencyBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      CurrencyBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Collects all Currency data for use with the Currency Calculator
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PF     9/4/01      Created
'------------------------------------------------------------------------------------------
Option Explicit
Implements ICurrencyBO
Implements ObjectControl
'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
    
Public Function FindList() As String
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
  
 On Error GoTo FindListVbErr
    Const strFunctionName As String = "FindList"
       
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                                        
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ICurrencyBO_FindList()
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    FindList = xmlResponseElem.xml
    m_objContext.SetComplete
FindListExit:
    '   clean-up section, release object references
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
    FindList = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindListExit
End Function
Private Function ICurrencyBO_FindList() As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get all instances of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    
    Const strFunctionName As String = "ICurrencyBO_FindList"
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objICurrencyDO As ICurrencyDO
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set objICurrencyDO = m_objContext.CreateInstance(App.Title & ".CurrencyDO")
                
    Set xmlDataNode = objICurrencyDO.FindList()
    xmlResponseElem.appendChild xmlDataNode
        
    m_objContext.SetComplete
    Set ICurrencyBO_FindList = xmlResponseElem
FindListExit:
    Set objICurrencyDO = Nothing
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
    Set ICurrencyBO_FindList = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindListExit
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
