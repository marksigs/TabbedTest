VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CaseTrackingDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omCaseTrackBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Case Tracking Data Object
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'DJP    11/04/01    SYS1927 Created
'LD     11/06/01    SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Implements ICaseTrackingDO

Private gobjContext As ObjectContext



'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'MV     19/08/2002  BMIDS00343  Core Upgrade Ref AQR: SYS5299 Modified ICaseTrackingDO_GetCaseTrackingAllData
'PSC    23/08/2002  BMIDS00343  Change GetXMLForCaseTracking to cope with null values
'DRC    23/01/2003  BM0266      Change GetXMLForCaseTracking so that XML format complies with the DTD
'------------------------------------------------------------------------------------------

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Private Sub ICaseTrackingDO_GetCaseTrackingAllData(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve case tracking data for the application passed in
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       APPLICATION node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetCaseTrackingAllDataExit

    Const strFunctionName As String = "ICaseTrackingDO_GetCaseTrackingAllData"

    Dim strAppNo As String
    Dim xmlCaseTrackNode As IXMLDOMNode
    Dim rst As ADODB.Recordset
    Dim adoCommand As ADODB.Command
    Dim adoParameter As ADODB.Parameter
    ' validate REQUEST & select CASETASK node
    Set xmlCaseTrackNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/APPLICATION")
    
    strAppNo = xmlGetMandatoryAttributeText(xmlCaseTrackNode, "APPLICATIONNUMBER")

    ''''''''''''''''''''''''''''''

    Set adoCommand = New ADODB.Command
    Dim strMaxRecords As String
    Dim strSQL As String
    strMaxRecords = "5000"
    
    'MV - 19/08/2002 - BMIDS00343 - Core Upgrade
    'SYS5299 - SQLServer call
    strSQL = "{CALL USP_CASETRACKDETAILS(?)}"
    
    adoCommand.CommandType = adCmdText
    adoCommand.CommandText = strSQL

    Set adoParameter = adoCommand.CreateParameter(, adBSTR, adParamInput, Len(strAppNo), strAppNo)
    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter
    ''''''''''''''''''''''''''''''
    
    Set rst = executeGetRecordSet(adoCommand)
    
    If rst Is Nothing Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
    Dim strTmp As String
    Dim xmlNode As IXMLDOMNode
    
    Set xmlNode = GetXMLForCaseTracking(vxmlResponseNode.ownerDocument, rst)
    vxmlResponseNode.appendChild xmlNode
    
GetCaseTrackingAllDataExit:

    Set xmlCaseTrackNode = Nothing

    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Function GetXMLForCaseTracking(domDoc As FreeThreadedDOMDocument40, rst As ADODB.Recordset) As IXMLDOMNode
    On Error GoTo GetXMLForCaseTrackingError
    Dim strGroup As String
    Dim strGroupText As String
    Dim strAreaStatus      As Variant
    Dim strGroupStatus As String
    Const strFunctionName As String = "GetXMLForCaseTracking"
    Dim strBusinessArea As String
    Dim strFreeFormatText As Variant
    Dim strBusinessAreaText As String
    Dim strPreviousBusinessArea As String

    Dim bFreeFormat As Boolean
    Dim dtCompletionDate As Variant
    
    Dim xmlNode As IXMLDOMElement
    Dim xmlReturnNode As IXMLDOMElement
    Dim xmlBusinessAreaNode As IXMLDOMElement
    Dim xmlGroupNode As IXMLDOMElement
    
    rst.MoveFirst
    
    Set xmlReturnNode = domDoc.createElement("CASETRACKSTATUS")
    bFreeFormat = False
    
    Do While Not rst.EOF
    
        ' PSC 23/08/2002 BMIDS00343
        strBusinessArea = rst("OT_CT_BUS_AREA").Value & ""
        
        If Not bFreeFormat Then
            bFreeFormat = True
            
            ' PSC 23/08/2002 BMIDS00343 - Start
            strFreeFormatText = rst("OT_CT_FREE_FORMAT_TEXT").Value & ""
            
'            If IsNull(strFreeFormatText) Then
'                strFreeFormatText = ""
'            End If
            ' PSC 23/08/2002 BMIDS00343 - End
            
            xmlReturnNode.setAttribute "FREEFORMATTEXT", strFreeFormatText
        End If
        
        If strBusinessArea <> strPreviousBusinessArea Then
            ' PSC 23/08/2002 BMIDS00343
            strBusinessAreaText = rst("OT_CT_BUS_AREA_TEXT").Value & ""
            Set xmlBusinessAreaNode = xmlReturnNode.appendChild(domDoc.createElement("BUSINESSAREA"))
            xmlBusinessAreaNode.setAttribute "BUSINESSAREA", strBusinessArea
            xmlBusinessAreaNode.setAttribute "BUSINESSAREA_TEXT", strBusinessAreaText
        End If

        ' PSC 23/08/2002 BMIDS00343
        strAreaStatus = rst("OT_CT_AREA_STATUS").Value & ""
        
        ' PSC 23/08/2002 BMIDS00343
        If Len(strAreaStatus) > 0 Then
            xmlBusinessAreaNode.setAttribute "AREASTATUS", strAreaStatus
        End If
        
        ' Add the Groups
        ' PSC 23/08/2002 BMIDS00343
        strGroup = rst("OT_CT_GROUP_NAME").Value & ""
        Set xmlGroupNode = xmlBusinessAreaNode.appendChild(domDoc.createElement("GROUP"))
        xmlGroupNode.setAttribute "GROUPNAME", strGroup
    
        ' Status
        ' PSC 23/08/2002 BMIDS00343
        strGroupStatus = rst("OT_CT_STATUS").Value & ""
        xmlGroupNode.setAttribute "STATUS", strGroupStatus
        
        ' Text
        ' PSC 23/08/2002 BMIDS00343
        strGroupText = rst("OT_CT_TEXT").Value & ""
        'DRC BM0266- changed to TEXT from Status
        xmlGroupNode.setAttribute "TEXT", strGroupText
        'xmlGroupNode.setAttribute "STATUS", strGroupText
        strPreviousBusinessArea = strBusinessArea
        
        'DRC - BM0266 Dates go with Groups not Areas
        dtCompletionDate = rst("OT_CT_COMPLETION_DATE").Value
        If Not IsNull(dtCompletionDate) Then
            xmlGroupNode.setAttribute "COMPLETIONDATE", CStr(dtCompletionDate)
      '     xmlBusinessAreaNode.setAttribute "COMPLETIONDATE", CStr(dtCompletionDate)
        End If
     
      
        rst.MoveNext
    Loop
    
    Set GetXMLForCaseTracking = xmlReturnNode
    
    Exit Function

GetXMLForCaseTrackingError:
    errCheckError strFunctionName, TypeName(Me)

End Function


