VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CostModelAssistDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ICostModelAssistDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Component to link cost modelling to AiP
'
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' MCS   13/01/2000  Created
' DJP   16/03/2000  Added GetNewPropertyDetails and GetApplicationData
' PSC   07/04/2000  Create objects correctly under MTS in GetApplicationData
' LD    07/11/2000  Explicity close recordsets
' LD    07/11/2000  Explicity destroy command objects
' AD    04/10/2001  Remove MoveFirst as the cursor was ForwardOnly. SYS2718.
' CL    07/05/2002  SYS4510  Modifications to class Initialize
' CL    10/05/02    SYS4510  Remove class initialize & class terminate
'------------------------------------------------------------------------------------------
'Mars History
'
' GHun  27/06/2006  MAR1878 Changed GetApplicationData and GetDataForSubQuoteValidation for performance
'------------------------------------------------------------------------------------------
Option Explicit
Implements ICostModelAssistDO
Implements ObjectControl

' Private member variables
Private m_objErrAssist As ErrAssist
Private m_objXMLAssist As XMLAssist
Private m_objConvertAssist As ConvertAssist
Private m_objContext As ObjectContext
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    If m_objErrAssist Is Nothing Then
        Set m_objErrAssist = New ErrAssist
    End If

    If m_objXMLAssist Is Nothing Then
        Set m_objXMLAssist = New XMLAssist
    End If

    If m_objConvertAssist Is Nothing Then
        Set m_objConvertAssist = New ConvertAssist
    End If

    Set m_objContext = GetObjectContext()
End Sub

Private Sub ObjectControl_Deactivate()
    Set m_objErrAssist = Nothing
    Set m_objXMLAssist = Nothing
    Set m_objContext = Nothing
End Sub

Private Function LoadDataForSubQuoteValidation() As FreeThreadedDOMDocument40
'------------------------------------------------------------------------------------------
'
'
'
' *****************************************************************************
    Const strFunctionName As String = "LoadDataForSubQuoteValidation"
    
    Dim strXML As String
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objXmlAssist As New XMLAssist
       
    strXML = strXML & "<TABLENAME>EMPLOYMENT"
    strXML = strXML & "<OTHERS>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACTIVEQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RECOMMENDEDQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCEPTEDQUOTENUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CUSTOMERNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CUSTOMERVERSIONNUMBER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MEMBEROFSTAFF<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set xmlClassDefDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set LoadDataForSubQuoteValidation = xmlClassDefDoc
    
    Set objXmlAssist = Nothing
    Set xmlClassDefDoc = Nothing

End Function

Private Function ICostModelAssistDO_GetDataForSubQuoteValidation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDataForSubQuoteValidationVbErr

    Const strFunctionName As String = "ICostModelAssistDO_GetDataForSubQuoteValidation"
      
    Dim objErrAssist As ErrAssist
    Dim objSQLAssist As SQLAssist
    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As XMLAssist
    Dim xmlRootNode As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                                
    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = xmlCreateDOMObject()
    
    Dim xmlResponseElem As IXMLDOMElement
    Set objErrAssist = New ErrAssist
    Set objSQLAssist = New SQLAssist
    Set objXmlAssist = New XMLAssist
                                
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
                                                                
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                                    
    'validate that we have a BASICQUOTATIONDETAILS passed in
    If vxmlRequest.nodeName = "BASICQUOTATIONDETAILS" Then
        Set xmlRootNode = vxmlRequest
    Else
        Set xmlRootNode = vxmlRequest.getElementsByTagName("BASICQUOTATIONDETAILS").Item(0)
    End If
    
    If xmlRootNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "BASICQUOTATIONDETAILS tag not found"
    End If

    Dim strApplicationNumber As String
    'get the application number out
    strApplicationNumber = objXmlAssist.GetTagValue(xmlRootNode, "APPLICATIONNUMBER")
    
    If Len(strApplicationNumber) = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "APPLICATIONNUMBER not found"
    End If
    
    Dim strSQL As String
    
    'MAR1878 GHun Improved performance of query and eliminated potential join errors
    strSQL = "SELECT CR.CUSTOMERNUMBER, CR.CUSTOMERVERSIONNUMBER, CV.MEMBEROFSTAFF," _
        & " AFF.APPLICATIONFACTFINDNUMBER , AFF.ACTIVEQUOTENUMBER, AFF.ACCEPTEDQUOTENUMBER, AFF.RECOMMENDEDQUOTENUMBER" _
        & " FROM CustomerRole CR" _
        & " INNER JOIN CustomerVersion CV" _
        & " ON CV.CustomerNumber = CR.CustomerNumber" _
        & " AND CV.CustomerVersionNumber = CR.CustomerVersionNumber" _
        & " INNER JOIN ApplicationFactFind AFF" _
        & " ON AFF.ApplicationNumber = CR.ApplicationNumber" _
        & " AND AFF.ApplicationFactFindNumber = CR.ApplicationFactFindNumber" _
        & " WHERE AFF.ApplicationNumber = " & objSQLAssist.FormatString(strApplicationNumber) _
        & " AND AFF.ApplicationFactFindNumber = 1" _
        & " AND CR.CustomerVersionNumber =  (SELECT MAX(CR2.CustomerVersionNumber)" _
        & " FROM CustomerRole CR2" _
        & " WHERE CR2.ApplicationNumber = AFF.ApplicationNumber" _
        & " AND CR2.ApplicationFactFindNumber = AFF.ApplicationFactFindNumber" _
        & " AND CR2.CustomerNumber = CR.CustomerNumber)"
           
    Dim cmd As ADODB.Command
    Set cmd = New ADODB.Command
    
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL
    
    Dim rstThisRecSet As ADODB.Recordset
    Set rstThisRecSet = New ADODB.Recordset

    Set rstThisRecSet = cmd.Execute(1, , adCmdText)
                        
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
    End If
                           
    Dim objClassDef As FreeThreadedDOMDocument40
    Set objClassDef = LoadDataForSubQuoteValidation
    
    Dim xmltempdoc As FreeThreadedDOMDocument40
    Set xmltempdoc = xmlCreateDOMObject()
    
    Dim xmlTempEmploymentElem As IXMLDOMElement
    Dim xmlEmploymentElem As IXMLDOMElement
    
    Set xmlEmploymentElem = xmltempdoc.createElement("EMPLOYMENTLIST")
    Set xmlTempEmploymentElem = xmltempdoc.appendChild(xmlEmploymentElem)

    Dim nloop As Integer
    nloop = 0
    Dim xmlTempElement As IXMLDOMElement

    Dim blnMemberOfStaff As Boolean
    Dim objConvertAssist As ConvertAssist
    Set objConvertAssist = New ConvertAssist
    
    ' AD 04/10/01 Remove MoveFirst as it's a ForwardOnly cursor. SYS 2718.
    blnMemberOfStaff = False
    ' loop through the record set
    While Not rstThisRecSet.EOF

        Set xmlEmploymentElem = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, objClassDef)
        
        Set xmlTempElement = xmlTempEmploymentElem.appendChild(xmlEmploymentElem)
        
        If nloop = 0 Then 'first time round the loop only
            Set xmlTempElement = xmlEmploymentElem.getElementsByTagName("APPLICATIONFACTFINDNUMBER").Item(nloop)
            Set xmlTempElement = xmlRootNode.appendChild(xmlTempElement.cloneNode(True))
            Set xmlTempElement = xmlEmploymentElem.getElementsByTagName("ACTIVEQUOTENUMBER").Item(nloop)
            Set xmlTempElement = xmlRootNode.appendChild(xmlTempElement.cloneNode(True))
        
        End If
               
        If blnMemberOfStaff = False Then
            
            If objConvertAssist.CSafeBool(objXmlAssist.GetTagValue(xmlEmploymentElem, "MEMBEROFSTAFF")) Then

                blnMemberOfStaff = True
            
            End If
        
        End If
        
        nloop = nloop + 1 'increment loop counter
        rstThisRecSet.MoveNext
        
    Wend
    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
    Set cmd = Nothing

    Set xmlTempElement = xmlRootNode.ownerDocument.createElement("ISMEMBEROFSTAFF")
        
    xmlTempElement.Text = Abs(blnMemberOfStaff)
    Set xmlTempElement = xmlRootNode.appendChild(xmlTempElement)
    xmlResponseElem.appendChild xmlRootNode.cloneNode(True)
    Set ICostModelAssistDO_GetDataForSubQuoteValidation = xmltempdoc.documentElement
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objSQLAssist = Nothing
    Set objConvertAssist = Nothing
    
    Exit Function

GetDataForSubQuoteValidationVbErr:
    
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If
    Set rstThisRecSet = Nothing

    Set cmd = Nothing
'    Set objIClassDef = Nothing
'    Set xmlClassDefDoc = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'Private Function ICostModelAssistDO_GetDataForSubQuoteValidation(ByVal vxmlRequest As _
'                                                        IXMLDOMElement) As IXMLDOMNode
'' header ----------------------------------------------------------------------------------
'' description:
''   Get the data for a single instance of the persistant data associated with
''   this data object
'' pass:
''   vstrXMLRequest  xml Request data stream containing data to which identifies
''                   the instance of the persistant data to be retrieved
'' return:
''   GetData         string containing XML data stream representation of
''                   data retrieved
'' Raise Errors: if record not found, raise ApplicationRecordNotFound
''------------------------------------------------------------------------------------------
'On Error GoTo GetDataForSubQuoteValidationVbErr
'
'    Const strFunctionName As String = "ICostModelAssistDO_GetDataForSubQuoteValidation"
'
'
'    Dim objErrAssist As New ErrAssist
'    Dim objSQLAssist As New SQLAssist
'    Dim objIDOAssist As IDOAssist
'    Dim objIADOAssist As IADOAssist
'    Dim objXmlAssist As New XMLAssist
'    Dim xmlRootNode As IXMLDOMNode
'
'    Dim objContext As ObjectContext
'    Set objContext = GetObjectContext()
'
'    Dim xmlOut As New DOMDocument
'    Dim xmlResponseElem As IXMLDOMElement
'
'    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
'    xmlOut.appendChild xmlResponseElem
'    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
'
'    If objContext Is Nothing Then
'        Set objIADOAssist = New ADOAssist
'        Set objIDOAssist = New DOAssist
'    Else
'        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
'        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
'    End If
'
'    'validate that we have a BASICQUOTATIONDETAILS passed in
'    If vxmlRequest.nodeName = "BASICQUOTATIONDETAILS" Then
'        Set xmlRootNode = vxmlRequest
'    Else
'        Set xmlRootNode = vxmlRequest.getElementsByTagName("BASICQUOTATIONDETAILS").Item(0)
'    End If
'
'    If xmlRootNode Is Nothing Then
'        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
'            oeMissingPrimaryTag, "BASICQUOTATIONDETAILS tag not found"
'    End If
'
'    Dim strApplicationNumber As String
'    'get the application number out
'    strApplicationNumber = objXmlAssist.GetTagValue(xmlRootNode, "APPLICATIONNUMBER")
'
'    If Len(strApplicationNumber) = 0 Then
'        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
'            oeMissingPrimaryTag, "APPLICATIONNUMBER not found"
'    End If
'
'    Dim strSQL As String
'
'    strSQL = "SELECT DISTINCT"
'    strSQL = strSQL & " CUSTOMERROLE.CUSTOMERNUMBER,CUSTOMERROLE.CUSTOMERVERSIONNUMBER,"
'    strSQL = strSQL & " CUSTOMERVERSION.MEMBEROFSTAFF,"
'    strSQL = strSQL & " APPLICATIONFACTFIND.APPLICATIONFACTFINDNUMBER,"
'    strSQL = strSQL & " APPLICATIONFACTFIND.ACTIVEQUOTENUMBER"
'    strSQL = strSQL & " FROM"
'    strSQL = strSQL & " CUSTOMERROLE,CUSTOMERVERSION,APPLICATIONFACTFIND"
'    strSQL = strSQL & " WHERE"
'    strSQL = strSQL & " APPLICATIONFACTFIND.APPLICATIONNUMBER = "
'    strSQL = strSQL & objSQLAssist.FormatString(strApplicationNumber)
'    strSQL = strSQL & " AND APPLICATIONFACTFIND.APPLICATIONNUMBER ="
'    strSQL = strSQL & " CUSTOMERROLE.APPLICATIONNUMBER"
'    strSQL = strSQL & " AND APPLICATIONFACTFIND.APPLICATIONFACTFINDNUMBER ="
'    strSQL = strSQL & " CUSTOMERROLE.APPLICATIONFACTFINDNUMBER "
'    strSQL = strSQL & " AND CUSTOMERROLE.CUSTOMERVERSIONNUMBER = "
'    strSQL = strSQL & " (SELECT MAX(CUSTOMERVERSIONNUMBER) FROM CUSTOMERROLE WHERE " & _
'                        "APPLICATIONNUMBER = APPLICATIONFACTFIND.APPLICATIONNUMBER AND " & _
'                        "APPLICATIONFACTFINDNUMBER = APPLICATIONFACTFIND.APPLICATIONFACTFINDNUMBER)"
'    strSQL = strSQL & " AND CUSTOMERROLE.CUSTOMERNUMBER = CUSTOMERVERSION.CUSTOMERNUMBER"
'    strSQL = strSQL & " AND APPLICATIONFACTFIND.APPLICATIONFACTFINDNUMBER ="
'    strSQL = strSQL & " (SELECT MAX(APPLICATIONFACTFINDNUMBER) FROM APPLICATIONFACTFIND"
'    strSQL = strSQL & " WHERE APPLICATIONNUMBER = "
'    strSQL = strSQL & objSQLAssist.FormatString(strApplicationNumber) & ")"
'
'    Dim Cmd As New ADODB.Command
'
'    Cmd.ActiveConnection = objIADOAssist.GetConnStr
'    Cmd.CommandText = strSQL
'
'    Dim rstThisRecSet As New ADODB.Recordset
'
'    Set rstThisRecSet = Cmd.Execute(1, , adCmdText)
'
'    If rstThisRecSet.EOF Then
'        ' raise application error to be interpreted by calling object
'        objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4RecordNotFound
'    End If
'
'    Dim objClassDef As DOMDocument
'    Dim strLoadData As String
'
'    Set objClassDef = LoadDataForSubQuoteValidation
'
'    Dim xmltempdoc As New DOMDocument
'    Dim xmlTempEmploymentElem As IXMLDOMElement
'    Dim xmlEmploymentElem As IXMLDOMElement
'
'    Set xmlEmploymentElem = xmltempdoc.createElement("EMPLOYMENTLIST")
'    Set xmlTempEmploymentElem = xmltempdoc.appendChild(xmlEmploymentElem)
'
'
'
'    Dim nloop As Integer
'    nloop = 0
'    Dim xmlTempElement As IXMLDOMElement
'
'    Dim blnMemberOfStaff As Boolean
'    Dim objConvertAssist As New ConvertAssist
'    rstThisRecSet.MoveFirst
'    blnMemberOfStaff = False
'    ' loop through the record set
'    While Not rstThisRecSet.EOF
'
'
'        Set xmlEmploymentElem = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, objClassDef)
'
'        Set xmlTempElement = xmlTempEmploymentElem.appendChild(xmlEmploymentElem)
'
'        If nloop = 0 Then 'first time round the loop only
'            Set xmlTempElement = xmlEmploymentElem.getElementsByTagName("APPLICATIONFACTFINDNUMBER").Item(nloop)
'            Set xmlTempElement = xmlRootNode.appendChild(xmlTempElement.cloneNode(True))
'            Set xmlTempElement = xmlEmploymentElem.getElementsByTagName("ACTIVEQUOTENUMBER").Item(nloop)
'            Set xmlTempElement = xmlRootNode.appendChild(xmlTempElement.cloneNode(True))
'
'        End If
'
'        If blnMemberOfStaff = False Then
'
'            If objConvertAssist.CSafeBool(objXmlAssist.GetTagValue(xmlEmploymentElem, "MEMBEROFSTAFF")) Then
'
'                blnMemberOfStaff = True
'
'            End If
'
'        End If
'
'        nloop = nloop + 1 'increment loop counter
'        rstThisRecSet.MoveNext
'
'    Wend
'
'    Set xmlTempElement = xmlRootNode.ownerDocument.createElement("ISMEMBEROFSTAFF")
'
'    xmlTempElement.Text = Abs(blnMemberOfStaff)
'    Set xmlTempElement = xmlRootNode.appendChild(xmlTempElement)
'    xmlResponseElem.appendChild xmlRootNode
'    Set ICostModelAssistDO_GetDataForSubQuoteValidation = xmltempdoc.documentElement
'
'
'    If Not objContext Is Nothing Then
'        objContext.SetComplete
'    End If
'
'
'GetDataForSubQuoteValidationExit:
'
'    Set objSQLAssist = Nothing
'    Set objIDOAssist = Nothing
'    Set objIADOAssist = Nothing
'    Set objXmlAssist = Nothing
'    Set xmlRootNode = Nothing
'    Set objErrAssist = Nothing
'    Set xmltempdoc = Nothing
'    Set objConvertAssist = Nothing
'    Set rstThisRecSet = Nothing
'
'    Exit Function
'
'GetDataForSubQuoteValidationVbErr:
'
'    If objErrAssist.IsWarning = True Then
'        objErrAssist.AddWarning xmlResponseElem
'        Resume Next
'    End If
'
'    objErrAssist.AddToErrSource strFunctionName
'
'    If objErrAssist.IsSystemError = True Then
'        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
'    End If
'
'    Set ICostModelAssistDO_GetDataForSubQuoteValidation = objErrAssist.CreateErrorResponseEx
'
'    If Not objContext Is Nothing Then
'        objContext.SetAbort
'    End If
'
'    '   go to clean-up section
'    Resume GetDataForSubQuoteValidationExit
'
'End Function

Private Function ICostModelAssistDO_GetApplicationData(vxmlRequest As IXMLDOMNode) As IXMLDOMNode
    
On Error GoTo GetApplicationDataErr

    Dim xmlAppDataRoot As IXMLDOMNode
    Dim objApplicationDO As IApplicationDO
    Dim domDoc As FreeThreadedDOMDocument40
    Const strFunctionName As String = "ICostModelAssistDO_GetApplicationData"
    Dim xmlAppNo As IXMLDOMNode
    Dim xmlFactFindNo As IXMLDOMNode
        
    m_objXMLAssist.CheckNode vxmlRequest, TypeName(Me), strFunctionName
    Set domDoc = xmlCreateDOMObject()
    
    Set xmlAppDataRoot = domDoc.appendChild(domDoc.createElement("APPLICATION"))
    Set xmlAppNo = m_objXMLAssist.GetNode(vxmlRequest, ".//APPLICATIONNUMBER")
    Set xmlFactFindNo = m_objXMLAssist.GetNode(vxmlRequest, ".//APPLICATIONFACTFINDNUMBER")
    
    xmlAppDataRoot.appendChild xmlAppNo.cloneNode(True)
    xmlAppDataRoot.appendChild xmlFactFindNo.cloneNode(True)
    
    'MAR1878 GHun No need to call GetApplicationData twice
    'Set xmlAppData = objApplicationDO.GetApplicationData(xmlAppDataRoot)
    'MAR1878 End
    
    If m_objContext Is Nothing Then
        Set objApplicationDO = New ApplicationDO
    Else
        Set objApplicationDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
    End If
    
    Set ICostModelAssistDO_GetApplicationData = objApplicationDO.GetApplicationData(xmlAppDataRoot)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetApplicationDataExit:

    '   clean-up section, release object references
    Set xmlAppDataRoot = Nothing
    Set objApplicationDO = Nothing
    Set domDoc = Nothing
    Set xmlAppNo = Nothing
    Set xmlFactFindNo = Nothing

    Exit Function

GetApplicationDataErr:

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ICostModelAssistDO_GetNewPropertyDetails(strAppNo As String, strFactFindNo As String) As IXMLDOMNode
' description:  Common method which obtains new property details. Called at present from
'               QQ and AQ.
' pass:         vxmlRequest  xml Request data stream
' return:                       xml Response data stream
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
On Error GoTo GetNewPropertyDetailsErr

    Dim xmlNewPropertyResp As IXMLDOMElement
    Dim xmlNewPropertyElem As IXMLDOMElement
    Dim xmlNewPropertyItem As IXMLDOMElement
    Dim xmlDOMNewProperty As FreeThreadedDOMDocument40
    Dim objIApplicationDO As IApplicationDO
    
Const strFunctionName As String = "ICostModelAssistDO_GetNewPropertyDetails"
    
    If m_objContext Is Nothing Then
        Set objIApplicationDO = New ApplicationDO
    Else
        Set objIApplicationDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationDO")
    End If
    
    Set xmlDOMNewProperty = New FreeThreadedDOMDocument40
    xmlDOMNewProperty.validateOnParse = False
    xmlDOMNewProperty.setProperty "NewParser", True
    
    Set xmlNewPropertyElem = xmlDOMNewProperty.appendChild(xmlDOMNewProperty.createElement("LOANPROPERTYDETAILS"))
    
    Set xmlNewPropertyItem = xmlNewPropertyElem.appendChild(xmlDOMNewProperty.createElement("APPLICATIONNUMBER"))
    xmlNewPropertyItem.Text = strAppNo
    
    Set xmlNewPropertyItem = xmlNewPropertyElem.appendChild(xmlDOMNewProperty.createElement("APPLICATIONFACTFINDNUMBER"))
    xmlNewPropertyItem.Text = strFactFindNo
    
    Set ICostModelAssistDO_GetNewPropertyDetails = objIApplicationDO.GetLoanPropertyDetails(xmlNewPropertyElem)

    Set xmlNewPropertyResp = Nothing
    Set xmlNewPropertyElem = Nothing
    Set xmlNewPropertyItem = Nothing
    Set xmlDOMNewProperty = Nothing
    Set objIApplicationDO = Nothing
    
    Exit Function
    
GetNewPropertyDetailsErr:

    Set xmlNewPropertyResp = Nothing
    Set xmlNewPropertyElem = Nothing
    Set xmlNewPropertyItem = Nothing
    Set xmlDOMNewProperty = Nothing
    Set objIApplicationDO = Nothing
    
    m_objErrAssist.AddToErrSource strFunctionName
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Function ICostModelAssistDO_GetValuationTypeAndLocation(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
On Error GoTo GetValuationTypeAndLocationVbErr
    Const strFunctionName As String = "ICostModelAssistDO_GetValuationTypeAndLocation"
    Dim xmlDOMNewProperty  As FreeThreadedDOMDocument40
    Dim xmlNewPropertyElem As IXMLDOMElement
    Dim xmlAppNo As IXMLDOMNode
    Dim xmlAppFactFindNo As IXMLDOMNode
    Dim objINewPropertyDO As omApp.INewPropertyDO
        
    If m_objContext Is Nothing Then
        Set objINewPropertyDO = New omApp.NewPropertyDO
    Else
        Set objINewPropertyDO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".NewPropertyDO")
    End If

    Set xmlAppNo = m_objXMLAssist.GetNode(vxmlRequest, "//APPLICATIONNUMBER")
    Set xmlAppFactFindNo = m_objXMLAssist.GetNode(vxmlRequest, "//APPLICATIONFACTFINDNUMBER")
    
    Set xmlDOMNewProperty = New FreeThreadedDOMDocument40
    xmlDOMNewProperty.validateOnParse = False
    xmlDOMNewProperty.setProperty "NewParser", True
    Set xmlNewPropertyElem = xmlDOMNewProperty.appendChild(xmlDOMNewProperty.createElement("NEWPROPERTY"))
    
    xmlNewPropertyElem.appendChild xmlAppNo.cloneNode(True)
    xmlNewPropertyElem.appendChild xmlAppFactFindNo.cloneNode(True)

    Set ICostModelAssistDO_GetValuationTypeAndLocation = objINewPropertyDO.GetValuationTypeAndLocation(xmlDOMNewProperty.documentElement)

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

GetValuationTypeAndLocationVbErr:
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function ICostModelAssistDO_CreateNewQuotation( _
                    ByVal vxmlRequest As IXMLDOMElement) As String
'----------------------------------------------------------------------------------
'Description : Creates a new quotation record and updates the new quotation number as
'              as active quote in ApplicationFactFind
'Pass        : vxmlRequest - containing ApplicationNumber, AFFNumber, A
'----------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "ICostModelAssistDO_CreateNewQuotation"

    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDoc As FreeThreadedDOMDocument40 ' Document to load the string response
    Dim xmlNewRequest As IXMLDOMNode
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement, xmlTableElem As IXMLDOMElement
    Dim xmlApplNumberElem As IXMLDOMElement, xmlAFFNumberElem As IXMLDOMElement
    Dim strResponse As String, strNewQuoteNumber As String
    
    Dim objQuotationBO As QuotationBO
    '*** FIX ME----------------------------------------------------------------------------
    'Refernce to omCR (which has ApplicationBO componrnt) is still existing. The reference
    'to omCR has to be removed; the methods in omCR have been moved to either to omApp or
    'omCR
    '-------------------------------------------------------------------------------------
    Dim objApplicationBO As omApp.ApplicationBO
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem

    If m_objContext Is Nothing Then
        Set objQuotationBO = New QuotationBO
        Set objApplicationBO = New omApp.ApplicationBO
    Else
        Set objQuotationBO = m_objContext.CreateInstance(gstrCOST_MODEL_COMPONENT & ".QuotationBO")
        Set objApplicationBO = m_objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationBO")
    End If

    ' Build the 'Request' and create a quotation
    Dim strTag As String
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("QUOTATION")
    xmlNewRequest.appendChild xmlTableElem
    
    strTag = "APPLICATIONNUMBER"
    Set xmlApplNumberElem = xmlOut.createElement(strTag)
    xmlApplNumberElem.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableElem.appendChild xmlApplNumberElem
    
    strTag = "APPLICATIONFACTFINDNUMBER"
    Set xmlAFFNumberElem = xmlOut.createElement(strTag)
    xmlAFFNumberElem.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableElem.appendChild xmlAFFNumberElem
    
    strTag = "ACTIVEQUOTENUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = vxmlRequest.selectSingleNode(".//" & strTag).Text
    xmlTableElem.appendChild xmlElement
    
    strResponse = objQuotationBO.CreateNewQuotation(xmlNewRequest.xml)
    Set xmlDoc = m_objXMLAssist.load(strResponse, TypeName(Me), strFunctionName)
    m_objErrAssist.ReRaiseResponseError xmlDoc.documentElement
    
    strNewQuoteNumber = xmlDoc.selectSingleNode("//QUOTATIONNUMBER").Text
    
    ' Build the request and update ApplicationFactFind
    Set xmlNewRequest = m_objXMLAssist.GetRequestNodeEx(vxmlRequest)
    Set xmlTableElem = xmlOut.createElement("APPLICATIONFACTFIND")
    xmlNewRequest.appendChild xmlTableElem
    
    xmlTableElem.appendChild xmlApplNumberElem
    xmlTableElem.appendChild xmlAFFNumberElem
    
    strTag = "ACTIVEQUOTENUMBER"
    Set xmlElement = xmlOut.createElement(strTag)
    xmlElement.Text = strNewQuoteNumber
    xmlTableElem.appendChild xmlElement
    
    strResponse = objApplicationBO.UpdateApplicationFactFind(xmlNewRequest.xml)
    m_objErrAssist.CheckXMLResponse strResponse, True
    
    ICostModelAssistDO_CreateNewQuotation = strNewQuoteNumber

    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlResponseElem = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    Set xmlApplNumberElem = Nothing
    Set xmlAFFNumberElem = Nothing

    Set objQuotationBO = Nothing
    Set objApplicationBO = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

CreateVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    Set xmlOut = Nothing
    Set xmlDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlResponseElem = Nothing
    Set xmlElement = Nothing
    Set xmlTableElem = Nothing
    Set xmlApplNumberElem = Nothing
    Set xmlAFFNumberElem = Nothing

    Set objQuotationBO = Nothing
    Set objApplicationBO = Nothing

    m_objErrAssist.AddToErrSource strFunctionName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If

    ' go to clean-up section
    Err.Raise Err.Number, Err.Source, Err.Description

End Function



