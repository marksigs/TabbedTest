VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "AddressTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AddressTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Public Sub Create(vstrXMLRequest As String) As String
'               Public Sub Update(vstrXMLRequest As String) As String
'               Public Sub Delete(vstrXMLRequest As String) As String

'Dependencies:  AddressDO, AddressBO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SR     02/11/99    Created
'------------------------------------------------------------------------------------------

Option Explicit
 
Public Function Create(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  create multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing data to be persisted
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "Create"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXMLAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlElement As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
   
    Dim objAddressDO As AddressDO
    Dim strXML As String

    Set xmlIn = objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = xmlIn.getElementsByTagName("ADDRESS")
    
    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "ADDRESS" & " tag not found"
    End If
   
    If objContext Is Nothing Then
        Set objAddressDO = New AddressDO
    Else
        Set objAddressDO = objContext.CreateInstance(App.Title & ".AddressDO")
    End If
    
       
    Dim intNodeIndex As Integer
    
    ' For each relevant node add key and call the data object create
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set xmlElement = xmlOut.createElement("GENERATEDKEYS")
        Set xmlNode = xmlOut.appendChild(xmlElement)

        strXML = objAddressDO.Create(xmlNodeList.Item(intNodeIndex).xml)
        xmlNode.appendChild objXMLAssist.load(strXML, TypeName(Me), strFunctionName).documentElement
    Next

    Create = xmlOut.xml
 
    Set objAddressDO = Nothing
    Set objXMLAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

CreateVbErr:
   
    Set objAddressDO = Nothing
    Set objXMLAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Sub Update(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  Update multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vstrXMLRequest  xml Request data stream containing update data
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "Update"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXMLAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement
        
    Dim objAddressDO As AddressDO
    
    Set xmlIn = objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
            
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = xmlIn.getElementsByTagName("ADDRESS")
    
    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "ADDRESS" & " tag not found"
    End If
    
    Dim intNodeIndex As Integer
    
    If objContext Is Nothing Then
        Set objAddressDO = New AddressDO
    Else
        Set objAddressDO = objContext.CreateInstance(App.Title & ".AddressDO")
    End If
    
    ' For each relevant node that we are passed call the data object update
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        objAddressDO.Update objElement.xml
    Next
     
    Set objAddressDO = Nothing
    Set objXMLAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
   
    Exit Sub

UpdateVbErr:
  
    Set objAddressDO = Nothing
    Set objXMLAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
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
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "Delete"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objXMLAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objElement As IXMLDOMElement

    Dim objAddressDO As AddressDO
   
    Set xmlIn = objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
            
    ' Get a list of nodes that need to be passed down to the data object
    Set xmlNodeList = xmlIn.getElementsByTagName("ADDRESS")
    
    If xmlNodeList.length = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "ADDRESS" & " tag not found"
    End If
    
    Dim intNodeIndex As Integer

    If objContext Is Nothing Then
        Set objAddressDO = New AddressDO
    Else
        Set objAddressDO = objContext.CreateInstance(App.Title & ".AddressDO")
    End If
    
    ' For each relevant node that we are passed call the data object delete
    For intNodeIndex = 0 To xmlNodeList.length - 1
        Set objElement = xmlNodeList.Item(intNodeIndex)
        objAddressDO.Delete objElement.xml
    Next
 
    Set objAddressDO = Nothing
    Set objXMLAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteVbErr:
   
    Set objAddressDO = Nothing
    Set objXMLAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNodeList = Nothing
    Set objElement = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
        
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


