VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "QuickQuotePersonalDebtsTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      QuickQuotePersonalDebtsTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Business Object which requires transactions
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String) As String
'               Public Sub Update(vstrXMLRequest As String) As String
'               Public Sub Delete(vstrXMLRequest As String) As String

'Dependencies:  QuickQuotePersonalDebtsDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     01/08/99    Created
'APS    01/03/00    SYS0324 - DeleteAll error not handled
'------------------------------------------------------------------------------------------

Option Explicit
 
 '=============================================
 'Constant Declaration Section
 '=============================================

 '=============================================
 'Variable Declaration Section
 '=============================================

Private Const cstrTABLE_NAME = "QUICKQUOTEPERSONALDEBTS"
 
 
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

    Const strFunctionName As String = "Create"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    Dim objQuickQuotePersonalDebtsDO As QuickQuotePersonalDebtsDO

    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                
    ' Get a list of nodes that need to be passed down to the data object
    Set objXmlNodeList = objXmlIn.getElementsByTagName(cstrTABLE_NAME)
    
    If objXmlNodeList.length = 0 Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, _
            omiga4missingprimarytag, cstrTABLE_NAME & " tag not found"
    End If
   
    If objContext Is Nothing Then
        Set objQuickQuotePersonalDebtsDO = New QuickQuotePersonalDebtsDO
    Else
        Set objQuickQuotePersonalDebtsDO = _
            objContext.CreateInstance(App.Title & ".QuickQuotePersonalDebtsDO")
    End If
    
    Dim intNodeIndex As Integer
    
    ' For each relevant node add key and call the data object create
    For intNodeIndex = 0 To objXmlNodeList.length - 1
        Set objElement = objXmlNodeList.Item(intNodeIndex)
        
        objQuickQuotePersonalDebtsDO.Create objElement.xml
    Next
 
    Set objQuickQuotePersonalDebtsDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

CreateVbErr:
   
    Set objQuickQuotePersonalDebtsDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
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

    Const strFunctionName As String = "Update"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
    Dim objXmlUpdateElement As IXMLDOMElement

' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented

' TODO ************************************************************************
'   replace QuickQuotePersonalDebtsDO reference with actual DO required
' *****************************************************************************
    Dim objQuickQuotePersonalDebtsDO As QuickQuotePersonalDebtsDO
    
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
            
    ' Get the after image to pass to the data object
    Dim blnFound As Boolean
    Dim intIndex As Integer
    Dim intNoOfUpdates As Integer
    
    blnFound = False
    intIndex = 0
    intNoOfUpdates = objXmlIn.getElementsByTagName("UPDATE").length
    
    ' Loop through each Update element to find the one with the after image
    Do While blnFound = False And intIndex < intNoOfUpdates
        Set objXmlUpdateElement = objXmlIn.getElementsByTagName("UPDATE").Item(intIndex)
        
        If objXmlUpdateElement.getAttribute("TYPE") = "AFTER" Then
            blnFound = True
        End If
        
        intIndex = intIndex + 1
    Loop
       
    If blnFound = False Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NoAfterImagePresent
    End If

    ' Get a list of nodes that need to be passed down to the data object
    Set objXmlNodeList = objXmlUpdateElement.getElementsByTagName(cstrTABLE_NAME)
    
    If objXmlNodeList.length = 0 Then
        objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4missingprimarytag, _
                                cstrTABLE_NAME & " tag not found"
    End If
    
    Dim intNodeIndex As Integer
    
    If objContext Is Nothing Then
        Set objQuickQuotePersonalDebtsDO = New QuickQuotePersonalDebtsDO
    Else
        Set objQuickQuotePersonalDebtsDO = objContext.CreateInstance(App.Title & ".QuickQuotePersonalDebtsDO")
    End If
    
    ' For each relevant node that we are passed call the data object update
    For intNodeIndex = 0 To objXmlNodeList.length - 1
        Set objElement = objXmlNodeList.Item(intNodeIndex)
        objQuickQuotePersonalDebtsDO.Update objElement.xml
    Next
 
    Set objXmlUpdateElement = Nothing
    Set objQuickQuotePersonalDebtsDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
   
    Exit Sub

UpdateVbErr:
   
    Set objXmlUpdateElement = Nothing
    Set objQuickQuotePersonalDebtsDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
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

    Const strFunctionName As String = "Delete"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
' TODO ************************************************************************
'   Remove omiga4NotImplemented error when the function has been implemented
' *****************************************************************************
    objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4NotImplemented

' TODO ************************************************************************
'   replace QuickQuotePersonalDebtsDO reference with actual DO required
' *****************************************************************************
    Dim objQuickQuotePersonalDebtsDO As QuickQuotePersonalDebtsDO
   
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
            
    ' Get a list of nodes that need to be passed down to the data object
    Set objXmlNodeList = objXmlIn.getElementsByTagName(cstrTABLE_NAME)
    
    If objXmlNodeList.length = 0 Then
        objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4missingprimarytag, _
                                cstrTABLE_NAME & " tag not found"
    End If
    
    Dim intNodeIndex As Integer

    If objContext Is Nothing Then
        Set objQuickQuotePersonalDebtsDO = New QuickQuotePersonalDebtsDO
    Else
        Set objQuickQuotePersonalDebtsDO = objContext.CreateInstance(App.Title & ".QuickQuotePersonalDebtsDO")
    End If
    
    ' For each relevant node that we are passed call the data object delete
    For intNodeIndex = 0 To objXmlNodeList.length - 1
        Set objElement = objXmlNodeList.Item(intNodeIndex)
        objQuickQuotePersonalDebtsDO.Delete objElement.xml
    Next
 
    Set objQuickQuotePersonalDebtsDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteVbErr:
   
    Set objQuickQuotePersonalDebtsDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlIn = Nothing
    Set objXmlNodeList = Nothing
    Set objElement = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
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
' return:       n/a
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllVbErr
    
    Const strFunctionName As String = "DeleteAll"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objErrAssist As New ErrAssist
    Dim objQuickQuotePersonalDebtsDO As QuickQuotePersonalDebtsDO

    If objContext Is Nothing Then
        Set objQuickQuotePersonalDebtsDO = New QuickQuotePersonalDebtsDO
    Else
        Set objQuickQuotePersonalDebtsDO = _
            objContext.CreateInstance(App.Title & ".QuickQuotePersonalDebtsDO")
    End If
        
    ' Pass through to the data object
    objQuickQuotePersonalDebtsDO.DeleteAll (vstrXMLRequest)
    
    Set objQuickQuotePersonalDebtsDO = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub
    
DeleteAllVbErr:

    ' APS 01/03/00 SYS0324 - DeleteAll error not handled
    ' As DOAssist.Delete raises the omiga4NoRowsAffected we need
    ' functionality here to resume next because it is at this stage that we set
    ' abort
    If Err.Number = omiga4NoRowsAffectedByDeleteAll Then
        Resume Next
    End If
    
    Set objQuickQuotePersonalDebtsDO = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
        objContext.SetComplete
    End If
    
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub



