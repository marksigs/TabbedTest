VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 4  'RequiresNewTransaction
END
Attribute VB_Name = "ApplicationNTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ApplicationNTxBO.cls
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:   Code template for Application Business Object interface which requires new transactions
'
'Dependencies:  List any other dependent components
'               none

'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: Requires new transaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' IK    20/06/02    SYS2372 CREATED
' IK    20/06/02    SYS2372 Added GetNextApplicationNumber
'
' BMids History:
' Prog  Date        Description
' DM    20/11/2002  BMIDS01022 Changed transaction mode to 1 - NoTransactions
' GHun  25/11/2002  BMIDS01043 Changed transaction mode back to 4 - Requires new transaction
'                              Moved GetNextApplicationNumber to ApplicationNSBO, as it should not support transactions after BMIDS01022 change
'                              Added UpdateApplication which requires a new transaction
'----------------------------------------------------------------------------------------------------------------
Option Explicit

Implements IApplicationNTxBO

'BMIDS001043
'Private Function IApplicationNTxBO_GetNextApplicationNumber(ByVal vstrChannelID As String) As String
'
'    Dim objDO As IApplicationDO
'    Set objDO = GetObjectContext.CreateInstance(App.Title & ".ApplicationDO")
'    IApplicationNTxBO_GetNextApplicationNumber = objDO.GetNextApplicationNumber(vstrChannelID)
'    Set objDO = Nothing
'
'End Function
'BMIDS001043 End

'BMIDS001043
Private Function IApplicationNTxBO_UpdateApplication(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  UpdateApplication multiple instances of the persistant data associated with this
'               business object
'               invokes appropriate Data Object(s) to perform persistance
' pass:         vxmlRequest  xml Request element
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationVbErr

    Const strFunctionName As String = "IApplicationNTxBO_UpdateApplication"

    Dim objContext          As ObjectContext
    Dim objXmlAssist        As XMLAssist
    Dim objErrAssist        As ErrAssist
    Dim xmlNodeList         As IXMLDOMNodeList
    Dim xmlElement          As IXMLDOMElement
    Dim xmlOut              As FreeThreadedDOMDocument40
    Dim xmlResponseElem     As IXMLDOMElement
    Dim objIApplicationDO   As IApplicationDO
    
    Set objContext = GetObjectContext()
    Set objXmlAssist = New XMLAssist
    Set objErrAssist = New ErrAssist
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' get a list of elements that are to be processed
    If vxmlRequest.nodeName = "APPLICATION" Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlNodeList = vxmlRequest.getElementsByTagName("APPLICATION")
    End If
    
    If xmlElement Is Nothing Then
        If xmlNodeList.length = 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "APPLICATION" & " tag not found"
        End If
    End If
   
    If objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = objContext.CreateInstance(App.Title & ".ApplicationDO")
    End If
    
    If Not xmlElement Is Nothing Then
        objIApplicationDO.UpdateApplication xmlElement
    Else
        ' For each relevant node call the data object UpdateApplication
        For Each xmlElement In xmlNodeList
            objIApplicationDO.UpdateApplication xmlElement
        Next
    End If
    
    Set IApplicationNTxBO_UpdateApplication = xmlResponseElem.cloneNode(True)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
UpdateApplicationExit:
    
    Set objIApplicationDO = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNodeList = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

UpdateApplicationVbErr:
    
    If objErrAssist.IsWarning = True Then
        objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IApplicationNTxBO_UpdateApplication = objErrAssist.CreateErrorResponseEx()

    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume UpdateApplicationExit

End Function
'BMIDS001043 End

