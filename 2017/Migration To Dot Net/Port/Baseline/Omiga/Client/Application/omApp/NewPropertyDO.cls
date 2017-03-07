VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "NewPropertyDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      NewPropertyDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date        Description
' RF    07/01/00    Created
' RF    19/01/00    Use ThirdPartyAssist rather than copied code.
' RF    19/01/00    Use IThirdPartyAssist.
' RF    15/02/00    Added method GetValuationTypeAndLocation.
' RF    21/02/00    Handle oeNoRowsAffectedByDeleteAll in UpdateNewPropertyDetails.
' SR    02/06/00    Added methods 'CreateNewPropertyDeposit' and 'UpdatePropertyDeposit'
' BG    12/06/00    Added GetPropertyUnknownDetails
' SR    15/06/00    Added GetFullNewPropertyDetails
' MC    19/06/00    Create SharedOwnership details for Import
' IVW   24/06/00    Amended code to take account of Ashley's code comments
' MC    14/07/00    SYS1176 Amended Create and Update NewPropertyAddress
' LD    07/11/00    Explicity close recordsets
' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'MV     03/09/2002  BMIDS00353  Amended INewPropertyDO_GetValuationTypeAndLocation()
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'MARS Specific History:
'
'Prog   Date        AQR         Description
'MV     24/10/2005  MAR109      Amended INewPropertyDO_GetValuationTypeAndLocation()
'------------------------------------------------------------------------------------------

Option Explicit

Implements INewPropertyDO
Private Function INewPropertyDO_GetFullNewPropertyDetails( _
            ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' retrieves of information from the tables NEWPROPERTY, NEWPROPERTYVENDORM, THIRDPARTY
' ADDRESS
' pass:
'   vxmlxmlTableElement
'       xml element containing the request
' return:
'       xml node containing retrieved data
'--------------------------------------------------------------------------------------------
'** FIX ME : This currently works for fetching only record from the tables NewPropertDeposit
'            and NewPropertyRoomType (Both have one - many relation with NewProperty). If
'            all the records from these tables are required, modify the method.

On Error GoTo GetFullNewPropertyDetailsVbErr

    Const strFunctionName As String = "INewPropertyDO_GetFullNewPropertyDetails"
    
    Dim objErrAssist As ErrAssist
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As XMLAssist
    
    Set objXmlAssist = New XMLAssist
    Set objErrAssist = New ErrAssist
    
    Dim xmlNodeToPromote As IXMLDOMNode
    Dim xmlTableElem As IXMLDOMNode
    Dim xmlThirdPartyNode As IXMLDOMNode
    Dim xmlAddressNode As IXMLDOMNode
    Dim xmlFullNewProperty As IXMLDOMNode
    Dim xmlNewPropertyVendor As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim objIApplClassDef As IApplicationClassDef
    
    Dim strApplNumber As String, intAFFNumber As Integer
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    '--------------------------------------------------------------------------------
    ' Build request and fetch data from view 'NEWPROPERTYTHIRDPARTYDETAILS'
    '--------------------------------------------------------------------------------
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlTableElem = xmlDoc.createElement("NEWPROPERTYTHIRDPARTYDETAILS")
    xmlTableElem.appendChild vxmlTableElement.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
    xmlTableElem.appendChild vxmlTableElement.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
    
    Set objIApplClassDef = New ApplicationClassDef
    ' SR : 20/07/00 - Do not raise error if no record is found
    Err.Clear
    On Error Resume Next
    Set xmlNewPropertyVendor = objIDOAssist.GetData(xmlTableElem, _
                                    objIApplClassDef.LoadNewPropertyThirdPartyData)
    
    objErrAssist.SaveErr
    On Error GoTo GetFullNewPropertyDetailsVbErr
    If objErrAssist.LastErr <> 0 Then
        If objErrAssist.GetOmigaErrorNumber(objErrAssist.LastErr) <> oeRecordNotFound Then
            objErrAssist.ReRaise
        End If
    Else  ' record is found
        ' Promote the NewProperty node as parent in the hierarchy
        Set xmlNodeToPromote = xmlNewPropertyVendor.selectSingleNode(".//NEWPROPERTYVENDOR")
        objXmlAssist.PromoteNode xmlNewPropertyVendor, xmlNodeToPromote, False, True
        
        Set xmlAddressNode = xmlNewPropertyVendor.selectSingleNode(".//ADDRESS")
        Set xmlThirdPartyNode = xmlNewPropertyVendor.selectSingleNode(".//THIRDPARTY")
        xmlThirdPartyNode.appendChild xmlAddressNode
           
    End If
    '----------------------------------------------------------------------------------------
    'Fetch data from 'FullNewPropertyDetails' and set them in the hierarchy
    '----------------------------------------------------------------------------------------
    Set xmlTableElem = xmlDoc.createElement("FULLNEWPROPERTYDETAILS")
    xmlTableElem.appendChild vxmlTableElement.selectSingleNode(".//APPLICATIONNUMBER").cloneNode(True)
    xmlTableElem.appendChild vxmlTableElement.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").cloneNode(True)
    
    Set xmlNode = vxmlTableElement.selectSingleNode("NPDEPOSITSEQUENCENUMBER")
    If Not xmlNode Is Nothing Then
        xmlTableElem.appendChild xmlNode.cloneNode(True)
    End If
    
    Set xmlNode = vxmlTableElement.selectSingleNode("ROOMTYPE")
    If Not xmlNode Is Nothing Then
        xmlTableElem.appendChild xmlNode.cloneNode(True)
    End If
    
    Set xmlFullNewProperty = objIDOAssist.GetData(xmlTableElem, objIApplClassDef.LoadFullNewPropertyDetails)
    
    '-----------------------------------------------------------------------------------
    'Append all the data in hierarchy and return
    '---------------------------------------------------------------------------------
    If Not xmlFullNewProperty Is Nothing Then
        Set xmlNodeToPromote = xmlFullNewProperty.selectSingleNode(".//NEWPROPERTYADDRESS")
        Set xmlAddressNode = xmlFullNewProperty.selectSingleNode(".//ADDRESS")
        xmlNodeToPromote.appendChild xmlAddressNode
        
        Set xmlNodeToPromote = xmlFullNewProperty.selectSingleNode(".//NEWPROPERTY")
        objXmlAssist.PromoteNode xmlFullNewProperty, xmlNodeToPromote, True, True
        
        If Not xmlNewPropertyVendor Is Nothing Then
            xmlFullNewProperty.appendChild xmlNewPropertyVendor.selectSingleNode(".//NEWPROPERTYVENDOR")
        End If

    End If
               
    Set INewPropertyDO_GetFullNewPropertyDetails = xmlFullNewProperty
    
    Exit Function
    
GetFullNewPropertyDetailsVbErr:
   

    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlNodeToPromote = Nothing
    Set xmlTableElem = Nothing
    Set xmlThirdPartyNode = Nothing
    Set xmlAddressNode = Nothing
    Set xmlFullNewProperty = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set objIApplClassDef = Nothing
        
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description
    

End Function

Private Function INewPropertyDO_GetNewPropertyAndAddressDetails( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'
' retrieves a subset of information from the tables NEWPROPERTYLEASEHOLD,
' NEWPROPERTY, NEWPROPERTYROOMTYPE *,ADDRESS,NEWPROPERTYADDRESS
' pass:
'   vxmlxmlTableElement
'       xml element containing the request
' return:
'       xml node containing retrieved data
'------------------------------------------------------------------------------------------

On Error GoTo GetNewPropertyAndAddressDetailsVBErr

    Const strFunctionName As String = "INewPropertyDO_GetNewPropertyAndAddressDetails"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim objSqlAssist As New SQLAssist
    Dim xmlReturnNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlKeyNode As IXMLDOMNode
    Dim xmlListNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Dim strSQL, strAppNo, strAppFFNo As String
    
    strAppNo = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    
    If Len(strAppNo) = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "APPLICATIONNUMBER must be specified"
    End If
    
    strAppFFNo = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    
    If Len(strAppFFNo) = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "APPLICATIONFACTFINDNUMBER must be specified"
    End If
    
    ' GETNEWPROPERTYANDADDRESSDETAILS is a view
    
    strSQL = "select * from NEWPROPERTYANDADDRESSDETAILS "
    strSQL = strSQL & "NEWPROPERTY WHERE "
    strSQL = strSQL & "APPLICATIONNUMBER = "
    strSQL = strSQL & objSqlAssist.FormatString(strAppNo)
    strSQL = strSQL & " AND APPLICATIONFACTFINDNUMBER = "
    strSQL = strSQL & objSqlAssist.FormatString(strAppFFNo)
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL

    Set rst = cmd.Execute(, , adCmdText)

    If rst.EOF Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    
    '------------------------------------------------------------------------------------------
    ' first get the single NEWPROPERTY record description details out of the recordset
    '------------------------------------------------------------------------------------------

    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyData(NPDNewPropertyAndAddressDetails)
        
    ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    ' rst.MoveFirst
    Set xmlReturnNode = objIDOAssist.GetXMLFromRecordSet(rst, xmlClassDefDoc)
    
    '------------------------------------------------------------------------------------------
    ' Room Type records
    '------------------------------------------------------------------------------------------
    
    Set xmlListNode = xmlReturnNode.appendChild(xmlReturnNode.ownerDocument.createElement("NEWPROPERTYROOMTYPELIST"))
    
    ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    'rst.MoveFirst
    
    Do While Not rst.EOF
    
        Set xmlClassDefDoc = objIClassDef.LoadNewPropertyRoomTypeData()
        objIDOAssist.GetXMLFromRecordSet rst, xmlClassDefDoc, xmlListNode
        
        rst.MoveNext
    Loop
    rst.Close
    
    Set INewPropertyDO_GetNewPropertyAndAddressDetails = xmlReturnNode

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlKeyNode = Nothing
    Set xmlListNode = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set objSqlAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlReturnNode = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetNewPropertyAndAddressDetailsVBErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlKeyNode = Nothing
    Set xmlListNode = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set objSqlAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlReturnNode = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub INewPropertyDO_CreateNewPropertyDeposit( _
                        ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlxmlTableElement
'       xml element containing the request
' return: n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "INewPropertyDO_CreateNewPropertyDeposit"
   
    Dim objErrAssist As New ErrAssist
   
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyDepositData()
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Function INewPropertyDO_GetPropertyUnknownDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get a subset of the data on the NEWPROPERTY table.
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetPropertyUnknownDetailsVbErr

    Const strFunctionName As String = "INewPropertyDO_GetPropertyUnknownDetails"
      
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadPropertyUnknownDetailsData()
    Set INewPropertyDO_GetPropertyUnknownDetails = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    ' TODO ************************************************************************
    '   Add Any derived data here
    ' *****************************************************************************
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetPropertyUnknownDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
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

Private Function INewPropertyDO_GetValuationTypeAndLocation( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get a subset of the data on the NEWPROPERTY table.
' pass:
'   vxmlxmlTableElement
'       xml element containing the request
' return:
'       xml node containing retrieved data
'------------------------------------------------------------------------------------------
    On Error GoTo GetValuationTypeAndLocationVbErr

    Const strFunctionName As String = "INewPropertyDO_GetValuationTypeAndLocation"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim xmlNewPropertyData As IXMLDOMNode
    Dim xmlSharedOwnerShipPercentage As IXMLDOMNode
    Dim objIClassDef As IApplicationClassDef
    Dim objXmlAssist As New XMLAssist
    Dim blnRecordNotFound  As Boolean
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyData(NPDValuationTypeAndLocation)
    Set xmlNewPropertyData = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    ' SHAREDOWNERSHIPDETAILS
    objXmlAssist.ChangeNodeName vxmlTableElement, "NEWPROPERTY", "SHAREDOWNERSHIPDETAILS"
    Set xmlClassDefDoc = objIClassDef.LoadSharedOwnershipDetailsData()
    Set xmlSharedOwnerShipPercentage = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    If Not blnRecordNotFound Then
        xmlNewPropertyData.appendChild xmlSharedOwnerShipPercentage
    End If
    
    Set INewPropertyDO_GetValuationTypeAndLocation = xmlNewPropertyData
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetValuationTypeAndLocationVbErr:
    
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        blnRecordNotFound = True
        Resume Next
    End If
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function INewPropertyDO_GetNewPropertyGeneral( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get a subset of the data on the NEWPROPERTY table.
' pass:
'   vxmlxmlTableElement
'       xml element containing the request
' return:
'       xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetNewPropertyGeneralVbErr

    Const strFunctionName As String = "INewPropertyDO_GetNewPropertyGeneral"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyData(NPDGeneralDetails)
    Set INewPropertyDO_GetNewPropertyGeneral = _
        objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetNewPropertyGeneralVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function INewPropertyDO_GetNewPropertyDescription( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get a subset of the data on the NEWPROPERTY table, plus all the data from child tables
'   NEWPROPERTYROOMTYPE, NEWPROPERTYLEASEHOLD and NEWPROPERTYDEPOSIT.
' pass:
'   vxmlxmlTableElement
'       xml element containing the request
' return:
'       xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetNewPropertyDescriptionVbErr

    Const strFunctionName As String = "INewPropertyDO_GetNewPropertyDescription"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim cmd As ADODB.Command
    Dim rst As ADODB.Recordset
    Dim objSqlAssist As New SQLAssist
    Dim xmlReturnNode As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode
    Dim xmlKeyNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Dim strSQL, strAppNo, strAppFFNo As String
    
    strAppNo = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    If Len(strAppNo) = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "APPLICATIONNUMBER must be specified"
    End If
    
    strAppFFNo = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    If Len(strAppNo) = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "APPLICATIONFACTFINDNUMBER must be specified"
    End If
    
    ' GETNEWPROPERTYDESCRIPTION is a view
    strSQL = "select * from GETNEWPROPERTYDESCRIPTION "
    strSQL = strSQL & "NEWPROPERTYLEASEHOLD WHERE "
    strSQL = strSQL & "APPLICATIONNUMBER = "
    strSQL = strSQL & objSqlAssist.FormatString(strAppNo)
    strSQL = strSQL & " AND APPLICATIONFACTFINDNUMBER = "
    strSQL = strSQL & objSqlAssist.FormatString(strAppFFNo)
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL

    Set rst = cmd.Execute(, , adCmdText)

    If rst.EOF Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    
    '------------------------------------------------------------------------------------------
    ' first get the single NEWPROPERTY record description details out of the recordset
    '------------------------------------------------------------------------------------------

    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyData(NPDDescriptionDetails)
    
    ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    ' rst.MoveFirst
    Set xmlReturnNode = objIDOAssist.GetXMLFromRecordSet(rst, xmlClassDefDoc)

    '------------------------------------------------------------------------------------------
    ' now append the child records
    '------------------------------------------------------------------------------------------
    
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyLeaseholdData()
    objIDOAssist.GetXMLFromRecordSet rst, xmlClassDefDoc, xmlReturnNode
    
    ' DM 16/10/01 SYS2718 MoveFirst Causes SQL Error on default forward only cursor.
    ' rst.MoveFirst
    Do While Not rst.EOF
        Set xmlClassDefDoc = objIClassDef.LoadNewPropertyRoomTypeData()
        objIDOAssist.GetXMLFromRecordSet rst, xmlClassDefDoc, xmlReturnNode
        
        Set xmlClassDefDoc = objIClassDef.LoadNewPropertyDepositData()
        objIDOAssist.GetXMLFromRecordSet rst, xmlClassDefDoc, xmlReturnNode
        
        rst.MoveNext
    Loop
    rst.Close

    ' remove duplicate child nodes
    objXmlAssist.RemoveDuplicates xmlReturnNode

    objXmlAssist.GroupNodesIntoList xmlReturnNode, "NEWPROPERTYDEPOSIT"
    objXmlAssist.GroupNodesIntoList xmlReturnNode, "NEWPROPERTYROOMTYPE"
    
    '------------------------------------------------------------------------------------------
    ' Mark all 'empty' DEPOSIT and ROOMTYPE records as such by setting their keys to null
    '------------------------------------------------------------------------------------------
    Set xmlNodeList = xmlReturnNode.selectNodes(".//NEWPROPERTYDEPOSIT")
    For Each xmlNode In xmlNodeList
        If objXmlAssist.GetTagValue(xmlNode, "NPDEPOSITSEQUENCENUMBER") = "" Then
            Set xmlKeyNode = xmlNode.selectSingleNode("APPLICATIONNUMBER")
            If Not (xmlKeyNode Is Nothing) Then xmlKeyNode.Text = ""
            Set xmlKeyNode = xmlNode.selectSingleNode("APPLICATIONFACTFINDNUMBER")
            If Not (xmlKeyNode Is Nothing) Then xmlKeyNode.Text = ""
        End If
    Next xmlNode

    Set xmlNodeList = xmlReturnNode.selectNodes(".//NEWPROPERTYROOMTYPE")
    For Each xmlNode In xmlNodeList
        If objXmlAssist.GetTagValue(xmlNode, "ROOMTYPE") = "" Then
            Set xmlKeyNode = xmlNode.selectSingleNode("APPLICATIONNUMBER")
            If Not (xmlKeyNode Is Nothing) Then xmlKeyNode.Text = ""
            Set xmlKeyNode = xmlNode.selectSingleNode("APPLICATIONFACTFINDNUMBER")
            If Not (xmlKeyNode Is Nothing) Then xmlKeyNode.Text = ""
        End If
    Next xmlNode

    Set INewPropertyDO_GetNewPropertyDescription = xmlReturnNode

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlKeyNode = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set objSqlAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlReturnNode = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetNewPropertyDescriptionVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlKeyNode = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set objSqlAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlReturnNode = Nothing
    Set xmlNode = Nothing
    Set xmlNodeList = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function INewPropertyDO_GetNewPropertyAddress( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get NEWPROPERTYADDRESS and ADDRESS.
' pass:
'   vxmlxmlTableElement
'       xml element containing the request
' return:
'       xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetNewPropertyAddressVbErr

    Const strFunctionName As String = "INewPropertyDO_GetNewPropertyAddress"
    
    Dim objErrAssist As New ErrAssist
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim xmlReturnNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyAddressDetailsData()
    
    Set xmlReturnNode = objIDOAssist.GetData( _
        vxmlTableElement, xmlClassDefDoc, , "NEWPROPERTYADDRESS")
        
    objXmlAssist.PromoteNode xmlReturnNode, _
        xmlReturnNode.selectSingleNode("NEWPROPERTYADDRESS")
        
    Set INewPropertyDO_GetNewPropertyAddress = xmlReturnNode

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturnNode = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetNewPropertyAddressVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturnNode = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub INewPropertyDO_CreateNewPropertyAddress( _
    ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Create NEWPROPERTYADDRESS and ADDRESS records. Also a NEWPROPERTY record, if one
'   does not already exist.
' pass:
'   vxmlxmlTableElement
'       xml element containing the request
' return: n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPropertyAddressVbErr

    Const strFunctionName As String = "INewPropertyDO_CreateNewPropertyAddress"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist As IADOAssist
    Dim objIAddressDO As IAddressDO
    Dim xmlElem As IXMLDOMElement
    Dim xmlElemTemp As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    Dim xmlKeysNode As IXMLDOMNode
    Dim xmlAddressNode As IXMLDOMNode
    Dim xmlAddressGUIDNode As IXMLDOMNode
    Dim strAddressGUID As String
    Dim sAddressGUID As String
    

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
        Set objIAddressDO = New AddressDO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIAddressDO = objContext.CreateInstance(App.Title & ".AddressDO")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    
    '------------------------------------------------------------------------------------------
    ' Save the Address record
    '------------------------------------------------------------------------------------------
    Set xmlKeysNode = vxmlTableElement.selectSingleNode("//NEWPROPERTYADDRESS")
    Set xmlAddressNode = vxmlTableElement.selectSingleNode("//NEWPROPERTYADDRESS/ADDRESS")
    Set xmlAddressGUIDNode = vxmlTableElement.selectSingleNode("ADDRESSGUID")
    If Not (xmlAddressGUIDNode Is Nothing) Then
        vxmlTableElement.removeChild xmlAddressGUIDNode
    End If
    If Not (xmlAddressNode Is Nothing) Then
        objIAddressDO.CreateAddress xmlAddressNode, xmlKeysNode
    End If
    
    '------------------------------------------------------------------------------------------
    ' Create the NEWPROPERTY record, if required
    '------------------------------------------------------------------------------------------
    
    Dim strTable As String
    Dim strFieldValuePair As String
    
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyData
    
    Set xmlElemTemp = vxmlTableElement.cloneNode(True)
    objXmlAssist.ChangeNodeName xmlElemTemp, "NEWPROPERTYADDRESS", "NEWPROPERTY"
    objIDOAssist.BuildSQLString _
        xmlElemTemp, xmlClassDefDoc, strTable, , , , strFieldValuePair
    If objIADOAssist.CheckRecordExists(strTable, strFieldValuePair) = False Then
        INewPropertyDO_CreateNewProperty xmlElemTemp
    End If
        
    '------------------------------------------------------------------------------------------
    ' Create the NEWPROPERTYADDRESS record
    '------------------------------------------------------------------------------------------
    
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyAddressData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlAddressGUIDNode = Nothing
    Set xmlAddressNode = Nothing
    Set xmlKeysNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateNewPropertyAddressVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set xmlAddressGUIDNode = Nothing
    Set xmlAddressNode = Nothing
    Set xmlKeysNode = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub INewPropertyDO_DeleteNewPropertyAddress(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete NewPropertyAddress record and associated Address record.
' pass:
'   vxmlxmlTableElement
'       xml element containing the request
' return: n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteNewPropertyAddressVbErr

    Const strFunctionName As String = "INewPropertyDO_DeleteNewPropertyAddress"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
           
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim objIAddressDO As IAddressDO
    Dim xmlNodeList As IXMLDOMNodeList
    Dim xmlNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim lngError As Long
   
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIAddressDO = New AddressDO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIAddressDO = objContext.CreateInstance(App.Title & ".AddressDO")
    End If
            
    '------------------------------------------------------------------------------------------
    ' Delete the NEWPROPERTYADDRESS record
    '------------------------------------------------------------------------------------------
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyAddressData()
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    '------------------------------------------------------------------------------------------
    ' Delete the NEWPROPERTYROOMTYPE record
    '------------------------------------------------------------------------------------------
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyRoomTypeData()
    Dim TempNode As FreeThreadedDOMDocument40
    Dim xmlApplicationNumber As IXMLDOMNode
    Dim xmlAppFactNumber As IXMLDOMNode
    Dim xmlRoomTypeDoc As FreeThreadedDOMDocument40
    
 
    Set xmlApplicationNumber = vxmlTableElement.selectSingleNode("//APPLICATIONNUMBER")
    Set xmlAppFactNumber = vxmlTableElement.selectSingleNode("//APPLICATIONFACTFINDNUMBER")
    
    Set xmlRoomTypeDoc = New FreeThreadedDOMDocument40
    xmlRoomTypeDoc.validateOnParse = False
    xmlRoomTypeDoc.setProperty "NewParser", True
    
    'Dim xmlRoomType As IXMLDOMElement
    
    Dim xmlRoomTypeElement As IXMLDOMElement
    Set xmlRoomTypeElement = xmlRoomTypeDoc.createElement("NEWPROPERTYROOMTYPE")
    xmlRoomTypeElement.appendChild xmlApplicationNumber
    xmlRoomTypeElement.appendChild xmlAppFactNumber
    On Error Resume Next
    objIDOAssist.DeleteAll xmlRoomTypeElement, xmlClassDefDoc
    
    If Err.Number <> 0 Then
        If objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeNoRowsAffectedByDeleteAll Then
            objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        End If
    End If
    '------------------------------------------------------------------------------------------
    ' Delete the NEWPROPERTYDEPOSIT record
    '------------------------------------------------------------------------------------------
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyDepositData
   
    Dim xmlDepositDoc As FreeThreadedDOMDocument40
    Dim xmlDeposit As IXMLDOMElement
    Set xmlDepositDoc = New FreeThreadedDOMDocument40
    xmlDepositDoc.validateOnParse = False
    xmlDepositDoc.setProperty "NewParser", True
    Set xmlDeposit = xmlDepositDoc.createElement("NEWPROPERTYDEPOSIT")
    
    xmlDeposit.appendChild xmlApplicationNumber
    xmlDeposit.appendChild xmlAppFactNumber
    On Error Resume Next
    objIDOAssist.DeleteAll xmlDeposit, xmlClassDefDoc
    
    If Err.Number <> 0 Then
        If objErrAssist.GetOmigaErrorNumber(Err.Number) <> oeNoRowsAffectedByDeleteAll Then
            objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
        End If
    End If
    '------------------------------------------------------------------------------------------
    ' Delete the NEWPROPERTYLEASEHOLD record
    '------------------------------------------------------------------------------------------
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyLeaseholdData
    
    Dim xmlLeaseholdDoc As FreeThreadedDOMDocument40
    Dim xmlLeasehold As IXMLDOMElement
    
    Set xmlLeaseholdDoc = New FreeThreadedDOMDocument40
    xmlLeaseholdDoc.validateOnParse = False
    xmlLeaseholdDoc.setProperty "NewParser", True
    Set xmlLeasehold = xmlDepositDoc.createElement("NEWPROPERTYLEASEHOLD")
    
    xmlLeasehold.appendChild xmlApplicationNumber
    xmlLeasehold.appendChild xmlAppFactNumber
    
    On Error Resume Next
    objIDOAssist.DeleteAll xmlLeasehold, xmlClassDefDoc
    
    If Err.Number <> 0 Then
        lngError = objErrAssist.GetOmigaErrorNumber(Err.Number)
        If lngError <> oeNoRowsAffectedByDeleteAll Then
            If lngError <> oeRecordNotFound Then
                objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
            End If
        End If
    End If
    
    lngError = 0
    
    '------------------------------------------------------------------------------------------
    ' Delete the NEWPROPERTYVENDOR record
    '------------------------------------------------------------------------------------------
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyThirdPartyData
    
    Dim xml3rdPartyDoc As FreeThreadedDOMDocument40
    Dim xml3rdParty As IXMLDOMElement
    
    Set xml3rdPartyDoc = New FreeThreadedDOMDocument40
    xml3rdPartyDoc.validateOnParse = False
    xml3rdPartyDoc.setProperty "NewParser", True
    Set xml3rdParty = xmlDepositDoc.createElement("NEWPROPERTYVENDOR")
    
    xml3rdParty.appendChild xmlApplicationNumber
    xml3rdParty.appendChild xmlAppFactNumber
    On Error Resume Next
    Call INewPropertyDO_DeleteVendorDetails(xml3rdParty)
    
    If Err.Number <> 0 Then
        lngError = objErrAssist.GetOmigaErrorNumber(Err.Number)
        If (lngError <> oeNoRowsAffectedByDeleteAll) Then
            If (lngError <> oeRecordNotFound) Then
                objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
            End If
        End If
    End If
    '------------------------------------------------------------------------------------------
    ' Delete the ADDRESS record
    '------------------------------------------------------------------------------------------
    
    If Len(objXmlAssist.GetTagValue(vxmlTableElement, "ADDRESSGUID")) > 0 Then
        Set xmlNode = vxmlTableElement.cloneNode(True)
        objXmlAssist.ChangeNodeName xmlNode, "NEWPROPERTYADDRESS", "ADDRESS"
        objIAddressDO.DeleteAddress xmlNode
    End If
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteNewPropertyAddressVbErr:

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function INewPropertyDO_GetVendorDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetVendorDetailsVbErr

    Const strFunctionName As String = "INewPropertyDO_GetVendorDetails"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim xmlReturn As IXMLDOMNode
    Dim objITPAssist As IThirdPartyAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objITPAssist = New ThirdPartyAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objITPAssist = objContext.CreateInstance(App.Title & ".ThirdPartyAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadVendorData()
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    ' add third party details
    objITPAssist.GetThirdParty xmlReturn

    Set INewPropertyDO_GetVendorDetails = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objITPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetVendorDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objITPAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub INewPropertyDO_DeleteVendorDetails( _
    ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVendorDetailsVbErr

    Const strFunctionName As String = "INewPropertyDO_DeleteVendorDetails"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim objITPAssist As IThirdPartyAssist
    Dim xmlData As IXMLDOMNode

    Dim strThirdPartyGUID As String
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objITPAssist = New ThirdPartyAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objITPAssist = objContext.CreateInstance(App.Title & ".ThirdPartyAssist")
    End If
            
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadVendorData()

    ' Retrieve THIRDPARTYGUID
    Set xmlData = INewPropertyDO_GetVendorDetails(vxmlTableElement)
    strThirdPartyGUID = objXmlAssist.GetTagValue(xmlData, "THIRDPARTYGUID")

    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    If strThirdPartyGUID <> "" Then
        ' Delete the thirdparty details
        objITPAssist.DeleteThirdParty strThirdPartyGUID
    End If

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objITPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteVendorDetailsVbErr:

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objITPAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub INewPropertyDO_UpdateNewPropertyDeposit( _
                ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "INewPropertyDO_UpdateNewPropertyDeposit"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyDepositData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub INewPropertyDO_UpdateVendorDetails( _
    ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVendorDetailsVbErr

    Const strFunctionName As String = "INewPropertyDO_UpdateVendorDetails"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim objITPAssist As IThirdPartyAssist
    Dim xmlThirdPartyGUIDNode As IXMLDOMNode
    Dim strThirdPartyGUID As String, _
        strDirectoryGUID  As String
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objITPAssist = New ThirdPartyAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objITPAssist = objContext.CreateInstance(App.Title & ".ThirdPartyAssist")
    End If

    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadVendorData()

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
        objITPAssist.DeleteThirdParty strThirdPartyGUID
    Else
        objITPAssist.MaintainThirdParty vxmlTableElement
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    End If
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objITPAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateVendorDetailsVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objITPAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub INewPropertyDO_CreateVendorDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVendorDetailsVbErr

    Const strFunctionName As String = "INewPropertyDO_CreateVendorDetails"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim xmlElemTemp As IXMLDOMElement
    Dim xmlElem As IXMLDOMElement
    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
    Dim objITPAssist As IThirdPartyAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    Set objIClassDef = New ApplicationClassDef
    
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
        Set objITPAssist = New ThirdPartyAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objITPAssist = objContext.CreateInstance(App.Title & ".ThirdPartyAssist")
    End If
    
    '------------------------------------------------------------------------------------------
    ' Create the NEWPROPERTY record, if required
    '------------------------------------------------------------------------------------------
    
    Dim strTable As String
    Dim strFieldValuePair As String
    
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyData
    Set xmlElemTemp = vxmlTableElement.cloneNode(True)
    objXmlAssist.ChangeNodeName xmlElemTemp, "NEWPROPERTYVENDOR", "NEWPROPERTY"
    objIDOAssist.BuildSQLString _
        xmlElemTemp, xmlClassDefDoc, strTable, , , , strFieldValuePair
    If objIADOAssist.CheckRecordExists(strTable, strFieldValuePair) = False Then
        INewPropertyDO_CreateNewProperty xmlElemTemp
    End If

    '------------------------------------------------------------------------------------------
    ' Maintain THIRDPARTY if necessary
    '------------------------------------------------------------------------------------------
    If (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "THIRDPARTY")) <> "") And _
       (Trim$(objXmlAssist.GetTagValue(vxmlTableElement, "DIRECTORYGUID")) = "") Then
        ' THIRDPARTY details exist in the XML, so maintain the appropriate THIRDPARTY record
        objITPAssist.MaintainThirdParty vxmlTableElement
    End If

    '------------------------------------------------------------------------------------------
    ' Maintain main table
    '------------------------------------------------------------------------------------------
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadVendorData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    'Set objINewPropertyBO = Nothing
    Set objITPAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateVendorDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objITPAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub INewPropertyDO_CreateNewPropertyDetails( _
    ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlxmlTableElement
'       xml element containing the request
' return: n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPropertyDetailsVbErr

    Const strFunctionName As String = "INewPropertyDO_CreateNewPropertyDetails"
   
    Dim objErrAssist As New ErrAssist
   
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    Dim xmlElem As IXMLDOMElement
    Dim xmlElemTemp As IXMLDOMElement
    Dim xmlTableNodeList As IXMLDOMNodeList
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlSequenceNode As IXMLDOMNode
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    
    Dim strTableName As String
    
    '------------------------------------------------------------------------------------------
    ' Create NEWPROPERTY record
    '------------------------------------------------------------------------------------------
    
    INewPropertyDO_CreateNewProperty vxmlTableElement
    
    '------------------------------------------------------------------------------------------
    ' Create NEWPROPERTYLEASEHOLD record
    '------------------------------------------------------------------------------------------
    
    If IsLeasehold(vxmlTableElement) Then
        strTableName = "NEWPROPERTYLEASEHOLD"
        
        Set xmlTableNodeList = vxmlTableElement.getElementsByTagName(strTableName)
        If (xmlTableNodeList.length) > 0 Then
            Set xmlClassDefDoc = objIClassDef.LoadNewPropertyLeaseholdData()
            objIDOAssist.Create xmlTableNodeList.Item(0), xmlClassDefDoc
        End If
    End If
    
    '------------------------------------------------------------------------------------------
    ' Create NEWPROPERTYDEPOSIT records;
    ' n.b. must cope with sequence number
    '------------------------------------------------------------------------------------------
    
    strTableName = "NEWPROPERTYDEPOSIT"
    
    Set xmlTableNodeList = vxmlTableElement.getElementsByTagName(strTableName)
    If (xmlTableNodeList.length) > 0 Then
        CreateNewPropertyDeposit xmlTableNodeList
    End If
   
    '------------------------------------------------------------------------------------------
    ' Create NEWPROPERTYROOMTYPE records
    '------------------------------------------------------------------------------------------
    
    strTableName = "NEWPROPERTYROOMTYPE"
    
    Set xmlTableNodeList = vxmlTableElement.getElementsByTagName(strTableName)
    If (xmlTableNodeList.length) > 0 Then
        Set xmlClassDefDoc = objIClassDef.LoadNewPropertyRoomTypeData()
        For Each xmlTableNode In xmlTableNodeList
            If objXmlAssist.GetTagValue(xmlTableNode, "APPLICATIONNUMBER") <> "" Then
                objIDOAssist.Create xmlTableNode, xmlClassDefDoc
            End If
        Next xmlTableNode
    End If
        
    '------------------------------------------------------------------------------------------
    ' Create SHAREDOWNERSHIPDETAILS records
    ' Added for Import. Shared Ownership record not created if tag not found
    ' in the incoming XML
    '------------------------------------------------------------------------------------------
    strTableName = "SHAREDOWNERSHIPDETAILS"
    Set xmlTableNodeList = vxmlTableElement.getElementsByTagName(strTableName)
    If (xmlTableNodeList.length) > 0 Then
        Set xmlClassDefDoc = objIClassDef.LoadSharedOwnershipDetailsData()
        For Each xmlTableNode In xmlTableNodeList
            If objXmlAssist.GetTagValue(xmlTableNode, "APPLICATIONNUMBER") <> "" Then
                objIDOAssist.Create xmlTableNode, xmlClassDefDoc
            End If
        Next xmlTableNode
    End If
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlTableNodeList = Nothing
    Set objXmlAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateNewPropertyDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    'Set objINewPropertyBO = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub INewPropertyDO_CreateNewProperty(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlxmlTableElement
'       xml element containing the request
' return: n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPropertyVbErr

    Const strFunctionName As String = "INewPropertyDO_CreateNewProperty"
   
    Dim objErrAssist As New ErrAssist
   
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyData()
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateNewPropertyVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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

Private Sub INewPropertyDO_UpdateNewProperty(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNewPropertyVbErr

    Const strFunctionName As String = "INewPropertyDO_UpdateNewProperty"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateNewPropertyVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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


Private Sub INewPropertyDO_UpdateNewPropertyRoomDetails( _
    ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update NEWPROPERTYROOMTYPE
' pass:
'   vxmlxmlTableElement
'       xml element containing the request
' return: n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNewPropertyRoomDetailsVbErr

    Const strFunctionName As String = "INewPropertyDO_UpdateNewPropertyRoomDetails"
        
    Dim objErrAssist As ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objXmlAssist As XMLAssist
    Dim blnDoCreate As Boolean

    Dim strApplicationNumber         As String, _
        strApplicationFactFindNumber As String

    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    Set xmlClassDefDoc = New FreeThreadedDOMDocument40
    Set objXmlAssist = New XMLAssist
    Set objErrAssist = New ErrAssist
    
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    
    strApplicationNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")

    If (strApplicationNumber = "") Or (strApplicationFactFindNumber = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "Expected APPLICATIONNUMBER " & _
            "and APPLICATIONFACTFINDNUMBER tags"
    End If
    
    ' Build the request block to execute DeleteAll by
    xmlRequestDoc.appendChild xmlRequestDoc.createElement("NEWPROPERTYROOMTYPE")
    xmlRequestDoc.documentElement.appendChild(xmlRequestDoc.createElement("APPLICATIONNUMBER")).Text = strApplicationNumber
    xmlRequestDoc.documentElement.appendChild(xmlRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")).Text = strApplicationFactFindNumber
    
    '------------------------------------------------------------------------------------------
    ' NEWPROPERTYROOMTYPE
    '------------------------------------------------------------------------------------------
    
    ' delete all NEWPROPERTYROOMTYPE using keys from NEWPROPERTY
    
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyRoomTypeData()
    
    Err.Clear
    On Error Resume Next
    objIDOAssist.DeleteAll xmlRequestDoc.documentElement, xmlClassDefDoc
    objErrAssist.SaveErr
    On Error GoTo UpdateNewPropertyRoomDetailsVbErr
    If objErrAssist.LastErr <> 0 And _
        objErrAssist.GetOmigaErrorNumber(objErrAssist.LastErr) <> oeNoRowsAffectedByDeleteAll Then
        objErrAssist.ReRaise
    End If
      
    ' create new NEWPROPERTYROOMTYPE records if there are any
    
    Set xmlNodeList = vxmlTableElement.getElementsByTagName("NEWPROPERTYROOMTYPE")
    If (xmlNodeList.length) > 0 Then
        For Each xmlNode In xmlNodeList
            If objXmlAssist.GetTagValue(xmlNode, "APPLICATIONNUMBER") <> "" Then
                objIDOAssist.Create xmlNode, xmlClassDefDoc
            End If
        Next xmlNode
    End If
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeNoRowsAffected Then
        Resume Next
    End If
    GoTo UpdateNewPropertyRoomDetailsVbErr

UpdateErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeNoRowsAffected Then
        blnDoCreate = True
        Resume Next
    End If

UpdateNewPropertyRoomDetailsVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Private Sub INewPropertyDO_UpdateNewPropertyDetails( _
    ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update NEWPROPERTY, NEWPROPERTYROOMTYPE, NEWPROPERTYDEPOSIT, NEWPROPERTYLEASEHOLD.
' pass:
'   vxmlxmlTableElement
'       xml element containing the request
' return: n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNewPropertyDetailsVbErr

    Const strFunctionName As String = "INewPropertyDO_UpdateNewPropertyDetails"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As New FreeThreadedDOMDocument40
    xmlClassDefDoc.validateOnParse = False
    xmlClassDefDoc.setProperty "NewParser", True
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    Dim xmlRequestDoc As New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    Dim xmlNode As IXMLDOMNode
    Dim xmlNodeList As IXMLDOMNodeList
    Dim objXmlAssist As New XMLAssist
    Dim blnDoCreate As Boolean
    Dim xmlNewPropertyLeasehold As IXMLDOMNode

    Dim strApplicationNumber         As String, _
        strApplicationFactFindNumber As String

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    
    strApplicationNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")

    If (strApplicationNumber = "") Or (strApplicationFactFindNumber = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeInvalidParameter, "Expected APPLICATIONNUMBER " & _
            "and APPLICATIONFACTFINDNUMBER tags"
    End If

    '------------------------------------------------------------------------------------------
    ' Update NEWPROPERTY
    '------------------------------------------------------------------------------------------
    
    INewPropertyDO_UpdateNewProperty vxmlTableElement

    ' Build the request block to execute DeleteAll by
    xmlRequestDoc.appendChild xmlRequestDoc.createElement("NEWPROPERTYROOMTYPE")
    xmlRequestDoc.documentElement.appendChild(xmlRequestDoc.createElement("APPLICATIONNUMBER")).Text = strApplicationNumber
    xmlRequestDoc.documentElement.appendChild(xmlRequestDoc.createElement("APPLICATIONFACTFINDNUMBER")).Text = strApplicationFactFindNumber

    '------------------------------------------------------------------------------------------
    ' NEWPROPERTYLEASEHOLD
    '------------------------------------------------------------------------------------------
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyLeaseholdData()
    Set xmlNewPropertyLeasehold = vxmlTableElement.selectSingleNode("NEWPROPERTYLEASEHOLD")
          
    If IsLeasehold(vxmlTableElement) Then
        ' Leasehold record is required - try to update an existing record first
        blnDoCreate = False
        On Error GoTo UpdateErr
        objIDOAssist.Update xmlNewPropertyLeasehold, xmlClassDefDoc
        On Error GoTo UpdateNewPropertyDetailsVbErr

        If blnDoCreate Then
            objIDOAssist.Create xmlNewPropertyLeasehold, xmlClassDefDoc
        End If
    Else
        ' NO leasehold record required - delete it if it does exist
        On Error GoTo DeleteErr
        objIDOAssist.Delete xmlNewPropertyLeasehold, xmlClassDefDoc
        On Error GoTo UpdateNewPropertyDetailsVbErr
    End If
    
    '------------------------------------------------------------------------------------------
    ' NEWPROPERTYROOMTYPE
    '------------------------------------------------------------------------------------------
    
    ' delete all NEWPROPERTYROOMTYPE using keys from NEWPROPERTY
    
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyRoomTypeData()
    
    Err.Clear
    On Error Resume Next
    objIDOAssist.DeleteAll xmlRequestDoc.documentElement, xmlClassDefDoc
    objErrAssist.SaveErr
    On Error GoTo UpdateNewPropertyDetailsVbErr
    If objErrAssist.LastErr <> 0 And _
        objErrAssist.GetOmigaErrorNumber(objErrAssist.LastErr) <> oeNoRowsAffectedByDeleteAll Then
        objErrAssist.ReRaise
    End If
      
    ' create new NEWPROPERTYROOMTYPE records if there are any
    
    Set xmlNodeList = vxmlTableElement.getElementsByTagName("NEWPROPERTYROOMTYPE")
    If (xmlNodeList.length) > 0 Then
        For Each xmlNode In xmlNodeList
            If objXmlAssist.GetTagValue(xmlNode, "APPLICATIONNUMBER") <> "" Then
                objIDOAssist.Create xmlNode, xmlClassDefDoc
            End If
        Next xmlNode
    End If
        
    '------------------------------------------------------------------------------------------
    ' NEWPROPERTYDEPOSIT
    '------------------------------------------------------------------------------------------
    
    ' delete all NEWPROPERTYDEPOSIT using keys from NEWPROPERTY
    
    objXmlAssist.ChangeNodeName xmlRequestDoc, "NEWPROPERTYROOMTYPE", "NEWPROPERTYDEPOSIT"
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyDepositData()
    
    Err.Clear
    On Error Resume Next
    objIDOAssist.DeleteAll xmlRequestDoc.documentElement, xmlClassDefDoc
    objErrAssist.SaveErr
    On Error GoTo UpdateNewPropertyDetailsVbErr
    If objErrAssist.LastErr <> 0 And _
        objErrAssist.GetOmigaErrorNumber(objErrAssist.LastErr) <> oeNoRowsAffectedByDeleteAll Then
        objErrAssist.ReRaise
    End If
      
    ' create new NEWPROPERTYDEPOSIT records if there are any
    
    Set xmlNodeList = vxmlTableElement.getElementsByTagName("NEWPROPERTYDEPOSIT")
    If (xmlNodeList.length) > 0 Then
        CreateNewPropertyDeposit xmlNodeList
    End If
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNewPropertyLeasehold = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeNoRowsAffected Then
        Resume Next
    End If
    GoTo UpdateNewPropertyDetailsVbErr

UpdateErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeNoRowsAffected Then
        blnDoCreate = True
        Resume Next
    End If

UpdateNewPropertyDetailsVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlNewPropertyLeasehold = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function IsLeasehold(ByVal vxmlTableElement As IXMLDOMElement) As Boolean
' header ----------------------------------------------------------------------------------
' description:
'   Return true if vxmlTableElement corresponds to leasehold or feudal.
' pass:
' return:
'------------------------------------------------------------------------------------------
On Error GoTo IsLeaseholdVbErr

    Const strFunctionName As String = "IsLeasehold"
        
    Dim objErrAssist As New ErrAssist
    Dim blnIsLeasehold As Boolean
    Dim xmlNode As IXMLDOMNode
    Dim objComboDO As ComboDO
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    If objContext Is Nothing Then
        Set objComboDO = New ComboDO
    Else
        Set objComboDO = objContext.CreateInstance(gstrBASE_COMPONENT & ".ComboDO")
    End If
    
    blnIsLeasehold = False
    
    Dim strTenureType As String
    
    Set xmlNode = vxmlTableElement.selectSingleNode("TENURETYPE")
    
    If Not xmlNode Is Nothing Then
        blnIsLeasehold = objComboDO.IsItemInValidation("PropertyTenure", Trim$(xmlNode.Text), "L")
    End If
    
    IsLeasehold = blnIsLeasehold
    
    Set objComboDO = Nothing
    
    Exit Function
    
IsLeaseholdVbErr:
    
    Set objComboDO = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub INewPropertyDO_UpdateNewPropertyAddress( _
    ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update NewPropertyAddress record and associated Address record.
' pass:
'   vxmlTableElement
'       xml element containing the request
' return: n/a
'------------------------------------------------------------------------------------------
On Error GoTo INewPropertyDO_UpdateNewPropertyAddressVbErr

    Const strFunctionName As String = "INewPropertyDO_UpdateNewPropertyAddress"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As New IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    Dim objIAddressBO As IAddressBO
    Dim objXmlAssist As New XMLAssist
    Dim xmlNode As IXMLDOMNode
    Dim xmlAddressNode As IXMLDOMNode
    Dim sAddressGUID As String
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIAddressBO = New AddressBO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIAddressBO = objContext.CreateInstance(App.Title & ".AddressBO")
    End If
    
    '------------------------------------------------------------------------------------------
    ' Update the ADDRESS record
    '------------------------------------------------------------------------------------------
    
    Set xmlAddressNode = vxmlTableElement.selectSingleNode(".//ADDRESS")
    If Not (xmlAddressNode Is Nothing) Then
        objIAddressBO.UpdateAddress xmlAddressNode
    Else
        sAddressGUID = objXmlAssist.GetElementText(vxmlTableElement, ".//ADDRESSGUID")
        If Len(Trim$(sAddressGUID)) > 0 Then
            objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "ADDRESS tag missing"
        End If
    End If
    
    '------------------------------------------------------------------------------------------
    ' Update the NEWPROPERTYADDRESS record
    '------------------------------------------------------------------------------------------
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyAddressData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objIAddressBO = Nothing
    Set objXmlAssist = Nothing
    Set xmlAddressNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

INewPropertyDO_UpdateNewPropertyAddressVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIAddressBO = Nothing
    Set objXmlAssist = Nothing
    Set xmlAddressNode = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Sub CreateNewPropertyDeposit(vxmlTableNodeList As IXMLDOMNodeList)
' header ----------------------------------------------------------------------------------
' description:
'   Create NEWPROPERTYDEPOSIT records; n.b. must cope with sequence number.
' pass:
'   vxmlTableElement
'       xml element containing the request
' return: n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewPropertyDepositVbErr
        
    Const strFunctionName As String = "CreateNewPropertyDeposit"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As New IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlTableNode As IXMLDOMNode
    Dim xmlSequenceNode As IXMLDOMNode
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Dim lngSeqNo As Long
    
    Dim strSequenceField, strTableName As String
    strSequenceField = "NPDEPOSITSEQUENCENUMBER"
    strTableName = "NEWPROPERTYDEPOSIT"
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyDepositData()
    
    lngSeqNo = objIDOAssist.GetNextSequenceNumber( _
        vxmlTableNodeList.Item(0), xmlClassDefDoc, strTableName, strSequenceField)
        
    For Each xmlTableNode In vxmlTableNodeList
        If objXmlAssist.GetTagValue(xmlTableNode, "APPLICATIONNUMBER") <> "" Then
            Set xmlSequenceNode = xmlTableNode.selectSingleNode(strSequenceField)
            If xmlSequenceNode Is Nothing Then
                ' No sequence number tag exists yet, so create one
                Set xmlSequenceNode = xmlTableNode.appendChild( _
                    xmlTableNode.ownerDocument.createElement(strSequenceField))
            End If
        
            xmlSequenceNode.Text = CStr(lngSeqNo)
                
            lngSeqNo = lngSeqNo + 1
        
            objIDOAssist.Create xmlTableNode, xmlClassDefDoc
        End If
    Next xmlTableNode
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    Exit Sub

CreateNewPropertyDepositVbErr:

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
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
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

