VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LockProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         : LockProcessing
' Description   : Contains all code to handle lock processing. This involves talking to the database
'                 directly and also going via the middle tier.
' Change history
' Prog      Date        Description
' DJP       26/06/01    SQL Server port, use Supervisor version of SQLAssist instead
' IK        07/03/03    BM0314 Added remove (DMS) Document Locks
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_enumLockType As LockType
Private m_objXMLAssist As XMLAssistSP
Private Sub RemoveCustomerLock(vApplicationNumber As Variant)
    On Error GoTo Failed
    Dim objXMLLock As New FreeThreadedDOMDocument
    Dim objListElement As IXMLDOMElement
    Dim objItem As IXMLDOMElement
    Dim objAppNo As IXMLDOMElement
    Dim objXMLResp As New FreeThreadedDOMDocument
    Dim nResponse As Integer
    Dim clsOmiga4 As New Omiga4Support
    Set objListElement = objXMLLock.createElement("REQUEST")
    objXMLLock.appendChild objListElement

    Set objItem = objXMLLock.createElement("CUSTOMERLOCK")
    objListElement.appendChild objItem

    Set objAppNo = objXMLLock.createElement("CUSTOMERNUMBER")
    objAppNo.Text = CStr(vApplicationNumber)
    objItem.appendChild objAppNo

    Dim xmlResponseNode As IXMLDOMNode
    Dim sResponse As String
    
    sResponse = clsOmiga4.RunASP(objListElement.xml, "DeleteCustomerLock.asp")
    
    ' Check the return code
    ' DJP SQL Server port, use Supervisor version instead
    g_clsXMLAssist.CheckXMLResponse sResponse, True

    ' Check the return code
    nResponse = g_clsXMLAssist.CheckXMLResponse(sResponse, True)

    If nResponse <> 0 Then
        g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub RemoveApplicationLock(vApplicationNumber As Variant)
    On Error GoTo Failed
    Dim objXMLLock As New FreeThreadedDOMDocument
    Dim objListElement As IXMLDOMElement
    Dim objAppNo As IXMLDOMElement
    Dim objXMLResp As New FreeThreadedDOMDocument
    Dim nResponse As Integer
    Dim clsOmiga4 As New Omiga4Support
    Set objListElement = objXMLLock.createElement("REQUEST")
    objXMLLock.appendChild objListElement

    Set objAppNo = objXMLLock.createElement("APPLICATIONNUMBER")
    objAppNo.Text = CStr(vApplicationNumber)
    objListElement.appendChild objAppNo

    Dim xmlResponseNode As IXMLDOMNode
    Dim sResponse As String
    
    sResponse = clsOmiga4.RunASP(objXMLLock.xml, "UnlockApplicationAndCustomers.asp")
    
    ' Check the return code
    g_clsXMLAssist.CheckXMLResponse sResponse, True

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

' ik_bm0314
' Public Sub RemoveLock(vApplicationNumber As Variant)
'    On Error GoTo Failed
'    Dim clsApplicationLock As ApplicationLockTable
'    Dim clsCustomerLock As CustomerLockTable
'
'    If Not IsEmpty(vApplicationNumber) And Not IsNull(vApplicationNumber) Then
'
'        Select Case m_enumLockType
'        Case Application
'            Set clsApplicationLock = New ApplicationLockTable
'
'            clsApplicationLock.DeleteApplicationAndCustomers CStr(vApplicationNumber)
'            'RemoveApplicationLock vApplicationNumber
'        Case Customer
'            'RemoveCustomerLock vApplicationNumber
'            Set clsCustomerLock = New CustomerLockTable
'            clsCustomerLock.DeleteCustomer CStr(vApplicationNumber)
'        End Select
'
'    End If
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
'End Sub

Public Sub RemoveLock(ByRef fcolKeys As Collection)
    On Error GoTo Failed
    Dim clsApplicationLock As ApplicationLockTable
    Dim clsCustomerLock As CustomerLockTable
    Dim clsDocumentLock As DocumentLock
    
'    If Not IsEmpty(vApplicationNumber) And Not IsNull(vApplicationNumber) Then

        Select Case m_enumLockType
        Case Application
            'RemoveApplicationLock vApplicationNumber
            Set clsApplicationLock = New ApplicationLockTable
            clsApplicationLock.DeleteApplicationAndCustomers CStr(fcolKeys(1))
            
        Case Customer
            'RemoveCustomerLock vApplicationNumber
            Set clsCustomerLock = New CustomerLockTable
            clsCustomerLock.DeleteCustomer CStr(fcolKeys(1))
            
        Case Document
            Set clsDocumentLock = New DocumentLock
            clsDocumentLock.DeleteDocumentLock fcolKeys
        
        End Select

'    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
' ik_bm0314_ends

Public Sub SetLockType(enumType As LockType)
    m_enumLockType = enumType
End Sub
Public Sub CreateLock(vAppNo As Variant, vUnitID As Variant, vUserID As Variant, vchannelID As Variant)
    On Error GoTo Failed
    ' First get the application Fact Find Number
    Dim sFactFind As String
    Dim clsFactFind As New AppFactFindTable
    
Const strFunctionName As String = "CreateLock"
    sFactFind = clsFactFind.GetNewestFactFind(CStr(vAppNo))

    If Len(sFactFind) > 0 Then
        If CLng(sFactFind) > 0 Then
            Dim objXMLLock As New FreeThreadedDOMDocument
            Dim objListElement As IXMLDOMElement
            Dim objItem As IXMLDOMElement
            Dim objXMLResp As New FreeThreadedDOMDocument
            Dim nResponse As Integer
            Dim sChannelID As String
            Dim clsOmiga4 As New Omiga4Support
            
            Set objListElement = objXMLLock.createElement("REQUEST")
            objXMLLock.appendChild objListElement
            objListElement.setAttribute "USERID", CStr(vUserID)
            objListElement.setAttribute "UNITID", CStr(vUnitID)
            objListElement.setAttribute "CHANNELID", CStr(vchannelID)
    
            Set objItem = objXMLLock.createElement("APPLICATIONNUMBER")
            objItem.Text = CStr(vAppNo)
            objListElement.appendChild objItem
    
            Set objItem = objXMLLock.createElement("APPLICATIONFACTFINDNUMBER")
            objItem.Text = sFactFind
            objListElement.appendChild objItem
    
            Dim sResponse As String
            Dim xmlDOMResp As FreeThreadedDOMDocument
            Dim sReadOnly As String
            sResponse = clsOmiga4.RunASP(objXMLLock.xml, "LockCustomersForApplication.asp")
            
            ' Check the return code
            g_clsXMLAssist.CheckXMLResponse sResponse, True
            
            Set xmlDOMResp = m_objXMLAssist.Load(sResponse, TypeName(Me), strFunctionName)
            sReadOnly = g_clsXMLAssist.GetAttributeFromNode(xmlDOMResp.documentElement, "READONLY")
        
            If Len(sReadOnly) > 0 Then
                If CBool(sReadOnly) = True Then
                    MsgBox "Application is already locked", vbInformation
                End If
            End If
        
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate Fact Find number for application number"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate Fact Find number for application number"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub Class_Initialize()
    Set m_objXMLAssist = New XMLAssistSP
End Sub
Private Sub Class_Terminate()
    Set m_objXMLAssist = Nothing
End Sub
