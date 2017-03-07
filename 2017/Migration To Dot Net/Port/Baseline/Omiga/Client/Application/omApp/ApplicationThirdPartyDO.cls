VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ApplicationThirdPartyDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ApplicationThirdPartyDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Code template for ApplicationThirdParty Data Object
'
'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'AD     11/01/00    Created
'SR     22/08/00    Modified mehtod 'UpdateBankBuildingSoc' to cater to deletion of ThirdParty
'                   and saving the Bank details
'SR     23/08/00    Modified method 'UpdateLegalRep' to cater to saving LegalRep details with
'                   no ThirdParty data
'JR     20/04/01    SYS2048 - Commented-out SetAbort/SetComplete in GetLegalRep method
'------------------------------------------------------------------------------------------

Option Explicit

' TODO ************************************************************************
'   Add interface definition class and amend this to implement it
' *****************************************************************************
Implements IApplicationThirdPartyDO


'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    11/12/2002  BM0094 - Legal Rep Contact Details
'MDC    10/01/2003  BM0244 - Reapply missing change from BM0094
'GHun   12/10/2004  BMIDS786 Changed UpdateLegalRep to create ContactDetails if they don't exist
'------------------------------------------------------------------------------------------

Private Function GetAppThirdPartyData(ByVal vxmlApplicationNumberNode As IXMLDOMNode, _
                                      ByVal vxmlApplicationFactFindNumberNode As IXMLDOMNode, _
                                      ByVal vstrTableName As String, _
                                      ByVal vxmlClassDefDoc As FreeThreadedDOMDocument40, _
                                      Optional ByVal vblnFindList As Boolean = False) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetAppThirdPartyDataVbErr

    Const strFunctionName As String = "GetAppThirdPartyData"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlRequest As IXMLDOMNode, _
        xmlReturn  As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    ' Build the request XML
    Set xmlRequest = vxmlApplicationNumberNode.ownerDocument.createElement(vstrTableName)
    xmlRequest.appendChild vxmlApplicationNumberNode.cloneNode(True)
    xmlRequest.appendChild vxmlApplicationFactFindNumberNode.cloneNode(True)

    If vblnFindList Then
        Set xmlReturn = objIDOAssist.FindList(xmlRequest, vxmlClassDefDoc)
        For Each xmlNode In xmlReturn.childNodes
            objTPAssist.GetThirdParty xmlNode
        Next xmlNode
    Else
        Set xmlReturn = objIDOAssist.GetData(xmlRequest, vxmlClassDefDoc)
        objTPAssist.GetThirdParty xmlReturn
    End If

GetAppThirdPartyDataExit:
    Set GetAppThirdPartyData = xmlReturn

    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlRequest = Nothing
    Set xmlReturn = Nothing
    Set xmlNode = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetAppThirdPartyDataVbErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        ' No data found, so return Nothing
        Set xmlReturn = Nothing
        Resume GetAppThirdPartyDataExit
    End If
    
    Set objIDOAssist = Nothing
    Set xmlRequest = Nothing
    Set xmlReturn = Nothing
    Set xmlNode = Nothing
    Set objTPAssist = Nothing
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub IApplicationThirdPartyDO_CreateArchitect(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateArchitect an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateArchitectVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_CreateArchitect"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef

    Dim objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Maintain THIRDPARTY if necessary
    '
    If (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTY")) <> "") And _
       (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID")) = "") Then
        ' THIRDPARTY details exist in the XML, so maintain the appropriate THIRDPARTY record
        objTPAssist.MaintainThirdParty vxmlTableElement
    End If

    '
    ' Maintain main table
    '
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadArchitectData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateArchitectVbErr:
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub IApplicationThirdPartyDO_CreateBankBuildingSociety(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateBankBuildingSociety an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateBankBuildingSocietyVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_CreateBankBuildingSociety"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef

    Dim objXmlAssist As New XMLAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist
   
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Maintain THIRDPARTY if necessary
    '
    If (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTY")) <> "") And _
       (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID")) = "") Then
        ' THIRDPARTY details exist in the XML, so maintain the appropriate THIRDPARTY record
        objTPAssist.MaintainThirdParty vxmlTableElement
    End If

    '
    ' Maintain main table
    '
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadBankBuildingSocietyData
    objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "BANKACCOUNTSEQUENCENUMBER"
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateBankBuildingSocietyVbErr:
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub IApplicationThirdPartyDO_CreateBuilder(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateBuilder an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateBuilderVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_CreateBuilder"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef

    Dim objXmlAssist As New XMLAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Maintain THIRDPARTY if necessary
    '
    If (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTY")) <> "") And _
       (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID")) = "") Then
        ' THIRDPARTY details exist in the XML, so maintain the appropriate THIRDPARTY record
        objTPAssist.MaintainThirdParty vxmlTableElement
    End If

    '
    ' Maintain main table
    '
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadBuilderData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateBuilderVbErr:
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IApplicationThirdPartyDO_CreateEstateAgent(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateEstateAgent an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateEstateAgentVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_CreateEstateAgent"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef

    Dim objXmlAssist As New XMLAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Maintain THIRDPARTY if necessary
    '
    If (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTY")) <> "") And _
       (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID")) = "") Then
        ' THIRDPARTY details exist in the XML, so maintain the appropriate THIRDPARTY record
        objTPAssist.MaintainThirdParty vxmlTableElement
    End If

    '
    ' Maintain main table
    '
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadEstateAgentData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateEstateAgentVbErr:
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IApplicationThirdPartyDO_CreateLegalRep(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateLegalRep an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateLegalRepVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_CreateLegalRep"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef

    Dim objXmlAssist As New XMLAssist
    
    'BM0094 MDC 11/12/2002 - Save Contact Details
    Dim objIContactBO As IContactBO
    Dim xmlContactDetails As IXMLDOMElement
    Dim xmlResponse As IXMLDOMElement
    Dim strContactDetailsGUID As String
    'BM0094 MDC 11/12/2002 - End
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIContactBO = New ContactBO   'BM0094 MDC 11/12/2002
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIContactBO = objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ContactBO")  'BM0094 MDC 11/12/2002
    End If

    '
    ' Maintain THIRDPARTY if necessary
    '
    If (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTY")) <> "") And _
       (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID")) = "") Then
        ' THIRDPARTY details exist in the XML, so maintain the appropriate THIRDPARTY record
        objTPAssist.MaintainThirdParty vxmlTableElement
    End If

    'BM0094 MDC 11/12/2002 - Save Contact Details
    Set xmlContactDetails = objXmlAssist.GetNode(vxmlTableElement, "CONTACTDETAILS", False)
    If Not xmlContactDetails Is Nothing Then
        Set xmlResponse = objIContactBO.CreateContact(xmlContactDetails)
        objErrAssist.CheckResponse xmlResponse, xmlResponse, True
        strContactDetailsGUID = objXmlAssist.GetMandatoryElementText(xmlResponse, "GENERATEDKEYS/CONTACTDETAILSGUID")
        objXmlAssist.SetMandatoryChildText vxmlTableElement, "CONTACTDETAILSGUID", strContactDetailsGUID
        'objXmlAssist.CopyNode "CONTACTDETAILSGUID", "CONTACTDETAILSGUID", xmlResponse, vxmlTableElement, False
    End If
    'BM0094 MDC 11/12/2002 - End
    
    '
    ' Maintain main table
    '
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadLegalRepData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    'BM0094 MDC 11/12/2002
    Set objIContactBO = Nothing
    Set xmlContactDetails = Nothing
    Set xmlResponse = Nothing
    'BM0094 MDC 11/12/2002 - End

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateLegalRepVbErr:
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    'BM0094 MDC 11/12/2002
    Set objIContactBO = Nothing
    Set xmlContactDetails = Nothing
    Set xmlResponse = Nothing
    'BM0094 MDC 11/12/2002 - End
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IApplicationThirdPartyDO_DeleteArchitect(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteArchitect a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteArchitectVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_DeleteArchitect"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef

    Dim xmlData As IXMLDOMNode

    Dim strThirdPartyGUID As String
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadArchitectData()

    ' Retrieve THIRDPARTYGUID
    Set xmlData = IApplicationThirdPartyDO_GetArchitect(vxmlTableElement)
    strThirdPartyGUID = objXmlAssist.GetTagValue(xmlData, "THIRDPARTYGUID")

    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    If strThirdPartyGUID <> "" Then
        ' Delete the thirdparty details
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    End If

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteArchitectVbErr:

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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
    
    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IApplicationThirdPartyDO_DeleteBankBuildingSociety(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteBankBuildingSociety a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBankBuildingSocietyVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_DeleteBankBuildingSociety"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef

    Dim xmlData As IXMLDOMNode

    Dim strThirdPartyGUID As String
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadBankBuildingSocietyData()

    ' Retrieve THIRDPARTYGUID
    Set xmlData = IApplicationThirdPartyDO_GetBankBuildingSociety(vxmlTableElement)
    strThirdPartyGUID = objXmlAssist.GetTagValue(xmlData, "THIRDPARTYGUID")

    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    If strThirdPartyGUID <> "" Then
        ' Delete the thirdparty details
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    End If

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteBankBuildingSocietyVbErr:

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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
    
    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IApplicationThirdPartyDO_DeleteBuilder(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteBuilder a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteBuilderVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_DeleteBuilder"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef

    Dim xmlData As IXMLDOMNode

    Dim strThirdPartyGUID As String
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadBuilderData()

    ' Retrieve THIRDPARTYGUID
    Set xmlData = IApplicationThirdPartyDO_GetBuilder(vxmlTableElement)
    strThirdPartyGUID = objXmlAssist.GetTagValue(xmlData, "THIRDPARTYGUID")

    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    If strThirdPartyGUID <> "" Then
        ' Delete the thirdparty details
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    End If

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteBuilderVbErr:

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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
    
    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IApplicationThirdPartyDO_DeleteEstateAgent(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteEstateAgent a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteEstateAgentVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_DeleteEstateAgent"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef

    Dim xmlData As IXMLDOMNode

    Dim strThirdPartyGUID As String
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadEstateAgentData()

    ' Retrieve THIRDPARTYGUID
    Set xmlData = IApplicationThirdPartyDO_GetEstateAgent(vxmlTableElement)
    strThirdPartyGUID = objXmlAssist.GetTagValue(xmlData, "THIRDPARTYGUID")

    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    If strThirdPartyGUID <> "" Then
        ' Delete the thirdparty details
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    End If

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteEstateAgentVbErr:

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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
    
    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IApplicationThirdPartyDO_DeleteLegalRep(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteLegalRep a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLegalRepVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_DeleteLegalRep"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef

    Dim xmlData As IXMLDOMNode

    Dim strThirdPartyGUID As String
    
    'BM0094 MDC 11/12/2002
    Dim strContactDetailsGUID As String
    Dim objIContactBO As IContactBO
    Dim xmlResponse As IXMLDOMElement
    Dim xmlRequest As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlDocument As FreeThreadedDOMDocument40
    'BM0094 MDC 11/12/2002 - End
    
    If objContext Is Nothing Then
        Set objIContactBO = New ContactBO   'BM0094 MDC 11/12/2002
        Set objIDOAssist = New DOAssist
    Else
        Set objIContactBO = objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ContactBO") 'BM0094 MDC 11/12/2002
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadLegalRepData()

    ' Retrieve THIRDPARTYGUID
    Set xmlData = IApplicationThirdPartyDO_GetLegalRep(vxmlTableElement)
    strThirdPartyGUID = objXmlAssist.GetTagValue(xmlData, "THIRDPARTYGUID")

    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    'BM0094 MDC 11/12/2002
    'Get ContactDetailsGuid (if one exists)
    strContactDetailsGUID = objXmlAssist.GetElementText(xmlData, "CONTACTDETAILS/CONTACTDETAILSGUID") '.GetTagValue(xmlData, "CONTACTDETAILSGUID")
    If Len(Trim$(strContactDetailsGUID)) > 0 Then
        'Delete Contact Details
        Set xmlDocument = New FreeThreadedDOMDocument40
        xmlDocument.validateOnParse = False
        xmlDocument.setProperty "NewParser", True
        Set xmlRequest = xmlDocument.createElement("CONTACTDETAILS")
        Set xmlTempNode = xmlDocument.createElement("CONTACTDETAILSGUID")
        xmlTempNode.Text = strContactDetailsGUID
        xmlRequest.appendChild xmlTempNode
        
        Set xmlResponse = objIContactBO.DeleteContact(xmlRequest)
        objErrAssist.CheckResponse xmlResponse, xmlResponse, True
    End If
    'BM0094 MDC 11/12/2002 - End
    
    If strThirdPartyGUID <> "" Then
        ' Delete the thirdparty details
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    End If

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    'BM0094 MDC 11/12/2002
    Set xmlResponse = Nothing
    Set xmlRequest = Nothing
    Set xmlTempNode = Nothing
    Set xmlDocument = Nothing
    'BM0094 MDC 11/12/2002 - End
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteLegalRepVbErr:

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    'BM0094 MDC 11/12/2002
    Set xmlResponse = Nothing
    Set xmlRequest = Nothing
    Set xmlTempNode = Nothing
    Set xmlDocument = Nothing
    'BM0094 MDC 11/12/2002 - End
    
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
    
    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function IApplicationThirdPartyDO_FindApplicationThirdPartyList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindApplicationThirdPartyListVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_FindApplicationThirdPartyList"

    Dim objErrAssist As New ErrAssist
    
    Const strAppTP As String = "APPLICATIONTHIRDPARTY"

    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef

    Dim xmlRequest As IXMLDOMNode, _
        xmlReturn  As IXMLDOMNode
    Dim xmlApplicationNumberNode         As IXMLDOMNode, _
        xmlApplicationFactFindNumberNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlAppTPNode     As IXMLDOMElement, _
        xmlThisAppTPNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    
    'BM0094 MDC 11/12/2002
    Dim objIContactBO As IContactBO
    Dim objXmlAssist As XMLAssist
    Dim xmlContactDetailsGuid As IXMLDOMElement
    Dim xmlContactDetails As IXMLDOMElement
    Dim xmlResponse As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    'BM0094 MDC 11/12/2002 - End
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIContactBO = New ContactBO   'BM0094 MDC 12/12/2002
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIContactBO = objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ContactBO")  'BM0094 MDC 12/12/2002
    End If

    '
    ' Initialise
    '
    xmlOut.appendChild xmlOut.createElement("APPLICATIONTHIRDPARTYLIST")
    Set objIAppTPClassDef = New AppTPClassDef

    ' Find the nodes in the request XML corresponding to the APPLICATIONNUMBER and APPLICATIONFACTFINDNUMBER
    Set xmlApplicationNumberNode = vxmlTableElement.selectSingleNode("APPLICATIONNUMBER")
    Set xmlApplicationFactFindNumberNode = vxmlTableElement.selectSingleNode("APPLICATIONFACTFINDNUMBER")

    If xmlApplicationNumberNode Is Nothing Or xmlApplicationFactFindNumberNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
    End If

    '
    ' Main
    '
    ' Architect
    Set xmlReturn = GetAppThirdPartyData(xmlApplicationNumberNode, xmlApplicationFactFindNumberNode, _
                        "APPLICATIONARCHITECT", objIAppTPClassDef.LoadArchitectData)
    If Not (xmlReturn Is Nothing) Then
        ' Add the architect as an applicationthirdparty to the return XML
        Set xmlAppTPNode = xmlOut.documentElement.appendChild(xmlOut.createElement(strAppTP))
        xmlAppTPNode.setAttribute "TABLE", "APPLICATIONARCHITECT"
        For Each xmlNode In xmlReturn.childNodes
            xmlAppTPNode.appendChild xmlNode
        Next xmlNode
    End If

    ' BankBuildingSociety
    Set xmlReturn = GetAppThirdPartyData(xmlApplicationNumberNode, xmlApplicationFactFindNumberNode, _
                        "APPLICATIONBANKBUILDINGSOC", objIAppTPClassDef.LoadBankBuildingSocietyData, vblnFindList:=True)
    If Not (xmlReturn Is Nothing) Then
        ' Add each bank as an applicationthirdparty to the return XML
        For Each xmlThisAppTPNode In xmlReturn.childNodes
            Set xmlAppTPNode = xmlOut.documentElement.appendChild(xmlOut.createElement(strAppTP))
            xmlAppTPNode.setAttribute "TABLE", "APPLICATIONBANKBUILDINGSOC"
            For Each xmlNode In xmlThisAppTPNode.childNodes
                xmlAppTPNode.appendChild xmlNode
            Next xmlNode
        Next xmlThisAppTPNode
    End If

    ' Builder
    Set xmlReturn = GetAppThirdPartyData(xmlApplicationNumberNode, xmlApplicationFactFindNumberNode, _
                        "APPLICATIONBUILDER", objIAppTPClassDef.LoadBuilderData)
    If Not (xmlReturn Is Nothing) Then
        ' Add the builder as an applicationthirdparty to the return XML
        Set xmlAppTPNode = xmlOut.documentElement.appendChild(xmlOut.createElement(strAppTP))
        xmlAppTPNode.setAttribute "TABLE", "APPLICATIONBUILDER"
        For Each xmlNode In xmlReturn.childNodes
            xmlAppTPNode.appendChild xmlNode
        Next xmlNode
    End If

    ' EstateAgent
    Set xmlReturn = GetAppThirdPartyData(xmlApplicationNumberNode, xmlApplicationFactFindNumberNode, _
                        "APPLICATIONESTATEAGENT", objIAppTPClassDef.LoadEstateAgentData)
    If Not (xmlReturn Is Nothing) Then
        ' Add the estateagent as an applicationthirdparty to the return XML
        Set xmlAppTPNode = xmlOut.documentElement.appendChild(xmlOut.createElement(strAppTP))
        xmlAppTPNode.setAttribute "TABLE", "APPLICATIONESTATEAGENT"
        For Each xmlNode In xmlReturn.childNodes
            xmlAppTPNode.appendChild xmlNode
        Next xmlNode
    End If

    ' LegalRep
    Set xmlReturn = GetAppThirdPartyData(xmlApplicationNumberNode, xmlApplicationFactFindNumberNode, _
                        "APPLICATIONLEGALREP", objIAppTPClassDef.LoadLegalRepData)
    If Not (xmlReturn Is Nothing) Then
        'BM0094 MDC 12/12/2002
        Set objXmlAssist = New XMLAssist
        Set xmlContactDetailsGuid = objXmlAssist.GetNode(xmlReturn, "CONTACTDETAILSGUID", False)
        If Not xmlContactDetailsGuid Is Nothing Then
            If xmlContactDetailsGuid.Text <> "" Then
                'Get Contact Details...
                Set xmlDoc = New FreeThreadedDOMDocument40
                xmlDoc.validateOnParse = False
                xmlDoc.setProperty "NewParser", True
                Set xmlRequest = xmlDoc.createElement("REQUEST")
                Set xmlContactDetails = xmlDoc.createElement("CONTACTDETAILS")
                xmlRequest.appendChild xmlContactDetails
                xmlContactDetails.appendChild xmlContactDetailsGuid
                Set xmlResponse = objIContactBO.GetContact(xmlRequest)
                objErrAssist.CheckResponse xmlResponse, xmlReturn, True
                
                '.. and append them to the Legal Rep
                Set xmlContactDetails = objXmlAssist.GetMandatoryNode(xmlResponse, "CONTACTDETAILS")
                Set xmlNode = objXmlAssist.GetNode(xmlReturn, ".//CONTACTDETAILS")
                If Not xmlNode Is Nothing Then
                    xmlReturn.insertBefore xmlContactDetails, xmlNode.parentNode
                Else
                    xmlReturn.appendChild xmlContactDetails
                End If
            End If
        End If
        'BM0094 MDC 12/12/2002 - End
        
        ' Add the legalrep as an applicationthirdparty to the return XML
        Set xmlAppTPNode = xmlOut.documentElement.appendChild(xmlOut.createElement(strAppTP))
        xmlAppTPNode.setAttribute "TABLE", "APPLICATIONLEGALREP"
        For Each xmlNode In xmlReturn.childNodes
            xmlAppTPNode.appendChild xmlNode
        Next xmlNode
    End If

    Set IApplicationThirdPartyDO_FindApplicationThirdPartyList = xmlOut.documentElement

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set xmlRequest = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    Set xmlAppTPNode = Nothing
    Set xmlThisAppTPNode = Nothing
    'BM0094 MDC 12/12/2002
    Set objIContactBO = Nothing
    Set objXmlAssist = Nothing
    Set xmlContactDetailsGuid = Nothing
    Set xmlContactDetails = Nothing
    Set xmlResponse = Nothing
    Set xmlDoc = Nothing
    'BM0094 MDC 12/12/2002 - End
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindApplicationThirdPartyListVbErr:
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set xmlRequest = Nothing
    Set xmlOut = Nothing
    Set xmlNode = Nothing
    Set xmlAppTPNode = Nothing
    Set xmlThisAppTPNode = Nothing
    'BM0094 MDC 12/12/2002
    Set objIContactBO = Nothing
    Set objXmlAssist = Nothing
    Set xmlContactDetailsGuid = Nothing
    Set xmlContactDetails = Nothing
    Set xmlResponse = Nothing
    Set xmlDoc = Nothing
    'BM0094 MDC 12/12/2002 - End
    
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

    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IApplicationThirdPartyDO_FindBankBuildingSocietyList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindBankBuildingSocietyListVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_FindBankBuildingSocietyList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef

    Dim xmlReturn As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadBankBuildingSocietyData
    Set xmlReturn = objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)

    For Each xmlNode In xmlReturn.childNodes
        objTPAssist.GetThirdParty xmlNode
    Next xmlNode

    Set IApplicationThirdPartyDO_FindBankBuildingSocietyList = xmlReturn

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set xmlNode = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindBankBuildingSocietyListVbErr:
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlNode = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
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

    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IApplicationThirdPartyDO_GetArchitect(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetArchitectVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_GetArchitect"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadArchitectData()
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    objTPAssist.GetThirdParty xmlReturn

    Set IApplicationThirdPartyDO_GetArchitect = xmlReturn

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetArchitectVbErr:
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function IApplicationThirdPartyDO_GetBankBuildingSociety(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetBankBuildingSocietyVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_GetBankBuildingSociety"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadBankBuildingSocietyData()
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    objTPAssist.GetThirdParty xmlReturn

    Set IApplicationThirdPartyDO_GetBankBuildingSociety = xmlReturn

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetBankBuildingSocietyVbErr:
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function IApplicationThirdPartyDO_GetBuilder(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetBuilderVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_GetBuilder"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadBuilderData()
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    objTPAssist.GetThirdParty xmlReturn

    Set IApplicationThirdPartyDO_GetBuilder = xmlReturn

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetBuilderVbErr:
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function IApplicationThirdPartyDO_GetEstateAgent(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetEstateAgentVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_GetEstateAgent"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadEstateAgentData()
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    objTPAssist.GetThirdParty xmlReturn

    Set IApplicationThirdPartyDO_GetEstateAgent = xmlReturn

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetEstateAgentVbErr:
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    
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
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function IApplicationThirdPartyDO_GetLegalRep(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetLegalRepVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_GetLegalRep"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef
    Dim xmlReturn As IXMLDOMNode
    
    'BM0094 MDC 11/12/2002
    Dim objIContactBO As IContactBO
    Dim objXmlAssist As XMLAssist
    Dim xmlContactDetailsGuid As IXMLDOMElement
    Dim xmlContactDetails As IXMLDOMElement
    Dim xmlRequest As IXMLDOMElement
    Dim xmlResponse As IXMLDOMElement
    Dim xmlDoc As FreeThreadedDOMDocument40
    'BM0094 MDC 11/12/2002 - End
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIContactBO = New ContactBO   'BM0094 MDC 11/12/2002
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIContactBO = objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ContactBO")  'BM0094 MDC 11/12/2002
    End If
    
    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadLegalRepData()
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    objTPAssist.GetThirdParty xmlReturn

    'BM0094 MDC 11/12/2002
    Set objXmlAssist = New XMLAssist
    Set xmlContactDetailsGuid = objXmlAssist.GetNode(xmlReturn, "CONTACTDETAILSGUID", False)
    If Not xmlContactDetailsGuid Is Nothing Then
        If xmlContactDetailsGuid.Text <> "" Then
            'Get Contact Details...
            Set xmlDoc = New FreeThreadedDOMDocument40
            xmlDoc.validateOnParse = False
            xmlDoc.setProperty "NewParser", True
            Set xmlRequest = xmlDoc.createElement("REQUEST")
            Set xmlContactDetails = xmlDoc.createElement("CONTACTDETAILS")
            xmlRequest.appendChild xmlContactDetails
            xmlContactDetails.appendChild xmlContactDetailsGuid
            Set xmlResponse = objIContactBO.GetContact(xmlRequest)
            objErrAssist.CheckResponse xmlResponse, xmlReturn, True
            
            '.. and append them to the Legal Rep
            Set xmlContactDetails = objXmlAssist.GetMandatoryNode(xmlResponse, "CONTACTDETAILS")
            xmlReturn.appendChild xmlContactDetails
        End If
    End If
    'BM0094 MDC 11/12/2002 - End
    
    Set IApplicationThirdPartyDO_GetLegalRep = xmlReturn

    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    'BM0094 MDC 11/12/2002
    Set objIContactBO = Nothing
    Set objXmlAssist = Nothing
    Set xmlContactDetailsGuid = Nothing
    Set xmlContactDetails = Nothing
    Set xmlRequest = Nothing
    Set xmlResponse = Nothing
    Set xmlDoc = Nothing
    'BM0094 MDC 11/12/2002 - End
    
' JR SYS2048
'    If Not objContext Is Nothing Then
'        objContext.SetComplete
'    End If
    
    Exit Function

GetLegalRepVbErr:
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objTPAssist = Nothing
    'BM0094 MDC 11/12/2002
    Set objIContactBO = Nothing
    Set objXmlAssist = Nothing
    Set xmlContactDetailsGuid = Nothing
    Set xmlContactDetails = Nothing
    Set xmlRequest = Nothing
    Set xmlResponse = Nothing
    Set xmlDoc = Nothing
    'BM0094 MDC 11/12/2002 - End
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
' JR SYS2048
'    If Not objContext Is Nothing Then
'        If objErrAssist.IsSystemError Then
'            objContext.SetAbort
'        Else
'            objContext.SetComplete
'        End If
'    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub IApplicationThirdPartyDO_UpdateArchitect(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateArchitect a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateArchitectVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_UpdateArchitect"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist


    Dim xmlThirdPartyGUIDNode As IXMLDOMNode
    Dim strThirdPartyGUID As String, _
        strDirectoryGUID  As String

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadArchitectData()

    '
    ' Maintain the associated THIRDPARTY record if appropriate
    '
    strThirdPartyGUID = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTYGUID"))
    strDirectoryGUID = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID"))

    If (strThirdPartyGUID <> "") And (strDirectoryGUID <> "") Then
        ' Cannot have both a DIRECTORY and THIRDPARTY GUID. Remove the THIRDPARTYGUID
        Set xmlThirdPartyGUIDNode = vxmlTableElement.selectSingleNode("THIRDPARTYGUID")
        If Not (xmlThirdPartyGUIDNode Is Nothing) Then
            xmlThirdPartyGUIDNode.Text = ""
        End If

        ' Update the primary table with the null THIRDPARTYGUID
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        ' Delete from the THIRDPARTY table
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    Else
        objTPAssist.MaintainThirdParty vxmlTableElement
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    End If
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateArchitectVbErr:
        
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub IApplicationThirdPartyDO_UpdateBankBuildingSociety(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateBankBuildingSociety a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBankBuildingSocietyVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_UpdateBankBuildingSociety"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlThirdPartyGUIDNode As IXMLDOMNode
    Dim strThirdPartyGUID As String, _
        strDirectoryGUID  As String

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadBankBuildingSocietyData()

    Dim botOperation As BOOPERATIONTYPE
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    
    Set objIThirdPartyClassDef = New ThirdPartyClassDef
    botOperation = objXmlAssist.DetermineOperationEx(vxmlTableElement, objIThirdPartyClassDef.LoadThirdPartyData)

    '
    ' Maintain the associated THIRDPARTY record if appropriate
    '
    strThirdPartyGUID = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTYGUID"))
    strDirectoryGUID = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID"))

    If (strThirdPartyGUID <> "") And (strDirectoryGUID <> "") Then
        ' Cannot have both a DIRECTORY and THIRDPARTY GUID. Remove the THIRDPARTYGUID
        Set xmlThirdPartyGUIDNode = vxmlTableElement.selectSingleNode("THIRDPARTYGUID")
        If Not (xmlThirdPartyGUIDNode Is Nothing) Then
            xmlThirdPartyGUIDNode.Text = ""
        End If

        ' Update the primary table with the null THIRDPARTYGUID
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        ' Delete from the THIRDPARTY table
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    Else
        ' SR 22/08/00 : If the ThirdParty is to be deleted, first remove the reference of TPGuid
        '               in main table and save it and then delete the ThirdParty record.
        If botOperation = booDelete Then
            Set xmlThirdPartyGUIDNode = vxmlTableElement.selectSingleNode("THIRDPARTYGUID")
            xmlThirdPartyGUIDNode.Text = ""
            
            objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
            objTPAssist.DeleteThirdParty strThirdPartyGUID
        Else
            objTPAssist.MaintainThirdParty vxmlTableElement
            objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        End If
    End If
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    Set objIThirdPartyClassDef = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateBankBuildingSocietyVbErr:
        
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub IApplicationThirdPartyDO_UpdateBuilder(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateBuilder a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateBuilderVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_UpdateBuilder"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlThirdPartyGUIDNode As IXMLDOMNode
    Dim strThirdPartyGUID As String, _
        strDirectoryGUID  As String

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadBuilderData()

    '
    ' Maintain the associated THIRDPARTY record if appropriate
    '
    strThirdPartyGUID = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTYGUID"))
    strDirectoryGUID = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID"))

    If (strThirdPartyGUID <> "") And (strDirectoryGUID <> "") Then
        ' Cannot have both a DIRECTORY and THIRDPARTY GUID. Remove the THIRDPARTYGUID
        Set xmlThirdPartyGUIDNode = vxmlTableElement.selectSingleNode("THIRDPARTYGUID")
        If Not (xmlThirdPartyGUIDNode Is Nothing) Then
            xmlThirdPartyGUIDNode.Text = ""
        End If

        ' Update the primary table with the null THIRDPARTYGUID
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        ' Delete from the THIRDPARTY table
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    Else
        objTPAssist.MaintainThirdParty vxmlTableElement
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    End If
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateBuilderVbErr:
        
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IApplicationThirdPartyDO_UpdateEstateAgent(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateEstateAgent a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateEstateAgentVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_UpdateEstateAgent"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist, _
        objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlThirdPartyGUIDNode As IXMLDOMNode
    Dim strThirdPartyGUID As String, _
        strDirectoryGUID  As String

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadEstateAgentData()

    '
    ' Maintain the associated THIRDPARTY record if appropriate
    '
    strThirdPartyGUID = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTYGUID"))
    strDirectoryGUID = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID"))

    If (strThirdPartyGUID <> "") And (strDirectoryGUID <> "") Then
        ' Cannot have both a DIRECTORY and THIRDPARTY GUID. Remove the THIRDPARTYGUID
        Set xmlThirdPartyGUIDNode = vxmlTableElement.selectSingleNode("THIRDPARTYGUID")
        If Not (xmlThirdPartyGUIDNode Is Nothing) Then
            xmlThirdPartyGUIDNode.Text = ""
        End If

        ' Update the primary table with the null THIRDPARTYGUID
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        ' Delete from the THIRDPARTY table
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    Else
        objTPAssist.MaintainThirdParty vxmlTableElement
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    End If
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateEstateAgentVbErr:
        
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    
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

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IApplicationThirdPartyDO_UpdateLegalRep(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateLegalRep a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLegalRepVbErr

    Const strFunctionName As String = "IApplicationThirdPartyDO_UpdateLegalRep"
        
    Dim objErrAssist As ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIAppTPClassDef As IAppTPClassDef
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As XMLAssist
    Dim objTPAssist As IThirdPartyAssist
    Set objTPAssist = New ThirdPartyAssist

    Dim xmlThirdPartyGUIDNode As IXMLDOMNode
    Dim strThirdPartyGUID As String, _
        strDirectoryGUID  As String

    Dim botOperation As BOOPERATIONTYPE
    Dim objIThirdPartyClassDef As IThirdPartyClassDef
    
    'BM0094 MDC 11/12/2002
    Dim xmlContactDetails As IXMLDOMElement
    Dim xmlResponse As IXMLDOMElement
    Dim objIContactBO As IContactBO
    'BM0094 MDC 11/12/2002 - End
    
    Dim strContactDetailsGUID As String 'BMIDS786 GHun
    
    Set objErrAssist = New ErrAssist
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIContactBO = New ContactBO       'BM0094 MDC 11/12/2002
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIContactBO = objContext.CreateInstance(gstrTHIRDPARTY_COMPONENT & ".ContactBO") 'BM0094 MDC 11/12/2002
    End If

    Set objIAppTPClassDef = New AppTPClassDef
    Set xmlClassDefDoc = objIAppTPClassDef.LoadLegalRepData()

    '
    ' Maintain the associated THIRDPARTY record if appropriate
    '
    Set objXmlAssist = New XMLAssist
    strThirdPartyGUID = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTYGUID"))
    strDirectoryGUID = Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID"))

    'BM0094 MDC 11/12/2002 - Update Contact Details
    Set xmlContactDetails = objXmlAssist.GetNode(vxmlTableElement, "CONTACTDETAILS", False)
    If Not xmlContactDetails Is Nothing Then
        'BMIDS786 GHun If the ContactDetailsGUID does not exist then create the contact details
        strContactDetailsGUID = objXmlAssist.GetNodeValue(xmlContactDetails, "CONTACTDETAILSGUID", False, False)
        If Len(strContactDetailsGUID) > 0 Then
        'BMIDS786 End
            Set xmlResponse = objIContactBO.UpdateContact(xmlContactDetails)
            objErrAssist.CheckResponse xmlResponse, xmlResponse, True
        'BMIDS786 GHun
        Else
            Set xmlResponse = objIContactBO.CreateContact(xmlContactDetails)
            objErrAssist.CheckResponse xmlResponse, xmlResponse, True
            'Save the returned CONTACTDETAILSGUID on ApplicationLegalRep
            strContactDetailsGUID = objXmlAssist.GetNodeValue(xmlResponse, ".//CONTACTDETAILSGUID")
            Set xmlContactDetails = objXmlAssist.GetNode(vxmlTableElement, "CONTACTDETAILSGUID")
            xmlContactDetails.Text = strContactDetailsGUID
        End If
        'BMIDS786 End
    End If
    'BM0094 MDC 11/12/2002 - End
    
    If (Len(strThirdPartyGUID) > 0) And (Len(strDirectoryGUID) > 0) Then
        ' Cannot have both a DIRECTORY and THIRDPARTY GUID. Remove the THIRDPARTYGUID
        Set xmlThirdPartyGUIDNode = vxmlTableElement.selectSingleNode("THIRDPARTYGUID")
        If Not (xmlThirdPartyGUIDNode Is Nothing) Then
            xmlThirdPartyGUIDNode.Text = ""
        End If

        ' Update the primary table with the null THIRDPARTYGUID
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        ' Delete from the THIRDPARTY table
        objTPAssist.DeleteThirdParty strThirdPartyGUID
    Else
        ' SR 23/08/00 SYS1457 : Delete the ThirdParty record and update LegalRep record accordingly
        Set objIThirdPartyClassDef = New ThirdPartyClassDef
        botOperation = objXmlAssist.DetermineOperationEx(vxmlTableElement, objIThirdPartyClassDef.LoadThirdPartyData)
        
        If botOperation = booDelete Then
            Set xmlThirdPartyGUIDNode = vxmlTableElement.selectSingleNode("THIRDPARTYGUID")
            xmlThirdPartyGUIDNode.Text = ""
            
            objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
            objTPAssist.DeleteThirdParty strThirdPartyGUID
        Else
            objTPAssist.MaintainThirdParty vxmlTableElement
            objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
        End If
    End If
    
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    Set objIThirdPartyClassDef = Nothing
    'BM0094 MDC 11/12/2002
    Set xmlContactDetails = Nothing
    Set xmlResponse = Nothing
    Set objIContactBO = Nothing
    'BM0094 MDC 11/12/2002 - End
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateLegalRepVbErr:
        
    Set objIAppTPClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objTPAssist = Nothing
    Set objIThirdPartyClassDef = Nothing
    'BM0094 MDC 11/12/2002
    Set xmlContactDetails = Nothing
    Set xmlResponse = Nothing
    Set objIContactBO = Nothing
    'BM0094 MDC 11/12/2002 - End
    
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

    Set objErrAssist = Nothing
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

