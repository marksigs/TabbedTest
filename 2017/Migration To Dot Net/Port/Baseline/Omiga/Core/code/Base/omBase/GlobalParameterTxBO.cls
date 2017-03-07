VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "GlobalParameterTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'Omiga4TxBO' with actual class name
' *****************************************************************************
'Workfile:      GlobalParameterTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Transactioned Global Parameter Business Object
'               which requires transactions
'               Public Sub Create(vstrXMLRequest As String) As String
'               Public Sub Update(vstrXMLRequest As String) As String
'               Public Sub Delete(vstrXMLRequest As String) As String
'Dependencies:  List any other dependent components
'               e.g. GlobalParameterDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'PSC    21/07/99 Created
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
Public Sub Create(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr
    Dim strFunctionName As String
    strFunctionName = "Create"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
' TODO ************************************************************************
'   replace GlobalParameterDO reference with actual DO required
' *****************************************************************************
    Dim objGlobalParameterDO As GlobalParameterDO
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                
' TODO ************************************************************************
'   replace TABLENAME with actual Table Reference e.g. ADDRESS
' *****************************************************************************
    ' Get a list of nodes that need to be passed down to the data object
    Set objXmlNodeList = objXmlIn.getElementsByTagName("TABLENAME")
    If objXmlNodeList.Length = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                "TABLENAME tag not found"
    End If
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(App.Title & ".GlobalParameterDO")
    End If
    Dim intNodeIndex As Integer
    ' For each relevant node add key and call the data object create
    For intNodeIndex = 0 To objXmlNodeList.Length - 1
        Set objElement = objXmlNodeList.Item(intNodeIndex)
' TODO ************************************************************************
'   Add Key generation if required
' *****************************************************************************
        objGlobalParameterDO.Create objElement.xml
    Next
    Set objGlobalParameterDO = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
CreateVbErr:
   
    Set objGlobalParameterDO = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub Update(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing update data
' return:       n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr
    Dim strFunctionName As String
    strFunctionName = "Update"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    Dim objXmlUpdateElement As IXMLDOMElement
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
' TODO ************************************************************************
'   replace GlobalParameterDO reference with actual DO required
' *****************************************************************************
    Dim objGlobalParameterDO As GlobalParameterDO
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
            
    ' Get the after image to pass to the data object
    Dim blnFound As Boolean
    Dim intIndex As Integer
    Dim intNoOfUpdates As Integer
    blnFound = False
    intIndex = 0
    intNoOfUpdates = objXmlIn.getElementsByTagName("UPDATE").Length
    ' Loop through each Update element to find the one with the after image
    Do While blnFound = False And intIndex < intNoOfUpdates
        Set objXmlUpdateElement = objXmlIn.getElementsByTagName("UPDATE").Item(intIndex)
        If objXmlUpdateElement.getAttribute("TYPE") = "AFTER" Then
            blnFound = True
        End If
        intIndex = intIndex + 1
    Loop
       
    If blnFound = False Then
        m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoAfterImagePresent
    End If
' TODO ************************************************************************
'   replace TABLENAME with actual Table Reference e.g. ADDRESS
' *****************************************************************************
    ' Get a list of nodes that need to be passed down to the data object
    Set objXmlNodeList = objXmlUpdateElement.getElementsByTagName("TABLENAME")
    If objXmlNodeList.Length = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                "TABLENAME tag not found"
    End If
    Dim intNodeIndex As Integer
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(App.Title & ".GlobalParameterDO")
    End If
    ' For each relevant node that we are passed call the data object update
    For intNodeIndex = 0 To objXmlNodeList.Length - 1
        Set objElement = objXmlNodeList.Item(intNodeIndex)
        objGlobalParameterDO.Update objElement.xml
    Next
    Set objXmlUpdateElement = Nothing
    Set objGlobalParameterDO = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
UpdateVbErr:
   
    Set objXmlUpdateElement = Nothing
    Set objGlobalParameterDO = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub Delete(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  Delete a single instance of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform delete
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance to be deleted
' return:       n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr
    Dim strFunctionName As String
    strFunctionName = "Delete"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objXmlAssist As New XMLAssist
'    Dim objErrAssist As New ErrAssist
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
' TODO ************************************************************************
'   replace GlobalParameterDO reference with actual DO required
' *****************************************************************************
    Dim objGlobalParameterDO As GlobalParameterDO
    Set objXmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
            
' TODO ************************************************************************
'   replace TABLENAME with actual Table Reference e.g. ADDRESS
' *****************************************************************************
    ' Get a list of nodes that need to be passed down to the data object
    Set objXmlNodeList = objXmlIn.getElementsByTagName("TABLENAME")
    If objXmlNodeList.Length = 0 Then
        m_objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                "TABLENAME tag not found"
    End If
    Dim intNodeIndex As Integer
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(App.Title & ".GlobalParameterDO")
    End If
    ' For each relevant node that we are passed call the data object delete
    For intNodeIndex = 0 To objXmlNodeList.Length - 1
        Set objElement = objXmlNodeList.Item(intNodeIndex)
        objGlobalParameterDO.Delete objElement.xml
    Next
    Set objGlobalParameterDO = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
DeleteVbErr:
   
    Set objGlobalParameterDO = Nothing
'    Set objXmlAssist = Nothing
'    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub DeleteAll(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  Delete the data for all instances of the persistant data associated with
'               this data object
' pass:         vstrXMLRequest  xml Request data stream containing data to which identifies
'                               instance of the persistant data to be retrieved
'
' return:       n/a
'
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr
    
    Dim strFunctionName As String
    strFunctionName = "DeleteAll"
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'    Dim objErrAssist As New ErrAssist
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    m_objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented
' TODO ************************************************************************
'   replace GlobalParameterDO reference with actual DO required
' *****************************************************************************
    Dim objGlobalParameterDO As GlobalParameterDO
    If m_objContext Is Nothing Then
        Set objGlobalParameterDO = New GlobalParameterDO
    Else
        Set objGlobalParameterDO = m_objContext.CreateInstance(App.Title & ".GlobalParameterDO")
    End If
        
    ' Pass through to the data object
    objGlobalParameterDO.DeleteAll (vstrXMLRequest)
    Set objGlobalParameterDO = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    Exit Sub
DeleteAllVbErr:
    Set objGlobalParameterDO = Nothing
'    Set objErrAssist = Nothing
    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

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
