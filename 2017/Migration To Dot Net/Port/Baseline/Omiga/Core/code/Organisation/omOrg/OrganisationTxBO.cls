VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "OrganisationTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      OrganisationTxBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Code template for omiga4 Business Object which requires transactions
'               contains (implied) standard interface:
'               Public Sub Create(vstrXMLRequest As String) As String
'               Public Sub Update(vstrXMLRequest As String) As String
'               Public Sub Delete(vstrXMLRequest As String) As String
'Dependencies:  List any other dependent components
'               e.g. OrganisationDO
'Issues:        not part of public interface
'               Instancing:         MultiUse
'               MTSTransactionMode: RequiresTransactions
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'RF     24/09/99 Created
'RF     11/11/99 Use ErrAssist.ThrowError rather than RaiseError.
'                Remove obsolete, unimplemented functions.
'MC     17/05/00 SYS0166 - Add method UpdateChangePasswordIndicator
'CL     7/11/00  CORE00005 Modifications to meet new standards
'------------------------------------------------------------------------------------------
Option Explicit
Implements IOrganisationTxBO
Implements ObjectControl
 '=============================================
 'Constant Declaration Section
 '=============================================
 '=============================================
 'Variable Declaration Section
 '=============================================
 Private m_objContext As ObjectContext
 Private m_objErrAssist As ErrAssist
 Private m_objXmlAssist As XMLAssist
 Private Const cstrPASSWORDROOT As String = "PASSWORD"
 Private Const cstrOMIGAUUSERROOT As String = "OMIGAUSER"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   BMIDS
'
'   AW     13/08/2002  BM029 - Added IOrganisationTxBO_UpdateUnitChequeNo
'
Private Function IOrganisationTxBO_CreatePassword(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreatePasswordVbErr
    Const strFunctionName As String = "IOrganisationTxBO_CreatePassword"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
       
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrPASSWORDROOT Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrPASSWORDROOT).Item(0)
    End If
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrPASSWORDROOT & " tag not found"
    End If
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
    objIOrganisationDO.CreatePassword xmlElement
    Set IOrganisationTxBO_CreatePassword = xmlResponseElem
    m_objContext.SetComplete
CreatePasswordExit:
    Set objIOrganisationDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
CreatePasswordVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationTxBO_CreatePassword = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetAbort
    ' go to clean-up section
    Resume CreatePasswordExit
     
End Function
Private Function IOrganisationTxBO_UpdateChangePasswordIndicator(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
'                    <REQUEST>
'                       <OMIGAUSER>
'                           <USERID></USERID>
'                           <CHANGEPASSWORDINDICATOR></CHANGEPASSWORDINDICATOR>
'                       </OMIGAUSER>
'                    </REQUEST>
'
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateChangePasswordIndicatorVbErr
    Const strFunctionName As String = "IOrganisationTxBO_UpdateChangePasswordIndicator"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                
    ' Check if the element passed in is the one with the table name else
    ' look below the node that is passed in
    If vxmlRequest.nodeName = cstrOMIGAUUSERROOT Then
        Set xmlElement = vxmlRequest
    Else
        Set xmlElement = vxmlRequest.getElementsByTagName(cstrOMIGAUUSERROOT).Item(0)
    End If
    If xmlElement Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  oeMissingPrimaryTag, _
                                  cstrOMIGAUUSERROOT & " tag not found"
    End If
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
    objIOrganisationDO.UpdateChangePasswordIndicator xmlElement
    Set IOrganisationTxBO_UpdateChangePasswordIndicator = xmlResponseElem
    m_objContext.SetComplete
UpdateChangePasswordIndicatorExit:
    
    Set objIOrganisationDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
UpdateChangePasswordIndicatorVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationTxBO_UpdateChangePasswordIndicator = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetAbort
    ' go to clean-up section
    Resume UpdateChangePasswordIndicatorExit
End Function
Private Function IOrganisationTxBO_UpdateUnitChequeNo(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance(s) of the persistant data to be updated
'                    <REQUEST>
'                       <OMIGAUSER>
'                           <USERID></USERID>
'                           <CHANGEPASSWORDINDICATOR></CHANGEPASSWORDINDICATOR>
'                       </OMIGAUSER>
'                    </REQUEST>
'
' return:       n/a
'   AW     13/08/2002  BM029
'------------------------------------------------------------------------------------------
On Error GoTo UpdateUnitChequeNoVbErr
    Const strFunctionName As String = "IOrganisationTxBO_UpdateUnitChequeNo"
    If m_objContext Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMTSNotFound
    End If
    Dim xmlElement As IXMLDOMElement
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objIOrganisationDO As IOrganisationDO
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set objIOrganisationDO = m_objContext.CreateInstance(App.Title & ".OrganisationDO")
    Set xmlElement = vxmlRequest
    objIOrganisationDO.UpdateUnitChequeNo xmlElement
    Set IOrganisationTxBO_UpdateUnitChequeNo = xmlResponseElem
    m_objContext.SetComplete
UpdateUnitChequeNoExit:
    
    Set objIOrganisationDO = Nothing
    Set xmlElement = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
UpdateUnitChequeNoVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set IOrganisationTxBO_UpdateUnitChequeNo = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetAbort
    ' go to clean-up section
    Resume UpdateUnitChequeNoExit
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
