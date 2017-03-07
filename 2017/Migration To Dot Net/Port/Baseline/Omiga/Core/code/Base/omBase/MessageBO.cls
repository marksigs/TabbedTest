VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "MessageBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      MessageBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Message Business Object which 'supports transactions' only
'
'Dependencies:  MessageDO
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
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
'Prog  Date     Description
'TK    22/11/04 BBG1821 - Performance related fixes.
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
 '=============================================
 'Constant Declaration Section
 '=============================================
 '=============================================
 'Variable Declaration Section
 '=============================================
' TODO ************************************************************************
'   replace TABLENAME with actual Table Reference e.g. ADDRESS
' *****************************************************************************
Private Const cstrTABLE_NAME = "MESSAGE"
Public Function GetData(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get the data for a single instance of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr
    
    Dim strFunctionName As String
    strFunctionName = "GetData"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlData As New FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
    Dim objMessageDO As MessageDO
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    Set xmlDataNode = xmlOut.appendChild(xmlResponseElem)
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    If m_objContext Is Nothing Then
        Set objMessageDO = New MessageDO
    Else
        Set objMessageDO = m_objContext.CreateInstance(App.Title & ".MessageDO")
    End If
                
    Set xmlRequestNode = xmlIn.getElementsByTagName(cstrTABLE_NAME).Item(0)
    '   call Data Object GetData function
    Dim strXMLData As String
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrTABLE_NAME & " tag not found"
    End If
        
    strXMLData = objMessageDO.GetData(xmlRequestNode.xml)
    Set xmlData = m_objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
    Set xmlResponseNode = xmlDataNode.appendChild(xmlData.documentElement)
    GetData = xmlOut.xml
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
GetDataExit:
    
    Set xmlResponseNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlResponseElem = Nothing
    Set xmlIn = Nothing
    Set xmlData = Nothing
    Set xmlOut = Nothing
'    Set objXmlAssist = Nothing
    Set xmlRequestNode = Nothing
'    Set objErrAssist = Nothing
    Set objMessageDO = Nothing
    Exit Function
GetDataVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    GetData = m_objErrAssist.CreateErrorResponse
        
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    '   go to clean-up section
    Resume GetDataExit
End Function

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = True
End Function
Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
End Sub
