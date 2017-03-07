VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CaseTrackingBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      omCaseTrackBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Case Tracking Business Object
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'DJP    11/04/01    SYS1927 Created
'------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl
Implements ICaseTrackingBO
Private gobjContext As ObjectContext

' Constants
Private Const cstrValuationComplete = 30
'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'MV     19/08/2002  BMIDS00343  Core Upgrade Ref AQR: SYS5299 Modified ICaseTrackingBO_FindApplicationsForCaseTracking
'MO     24/10/2002  BMIDS00652  Included the changes to intermediaries into Case tracking.
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

Private Sub ICaseTrackingBO_FindApplicationsForCaseTracking(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve Applications eligible for case tracking
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

    On Error GoTo FindApplicationsForCaseTrackingExit

    Const strFunctionName As String = "ICaseTrackingBO_FindApplicationsForCaseTracking"

    Dim strSource As String
    Dim strSchema As String
    Dim xmlCaseTrackNode As IXMLDOMNode

    ' validate REQUEST & select CASETASK node
    ' AQR SYS5299 - customers OR intermediaries
    
    'MO 24/10/2002 BMIDS00652 - Start
    
    Dim strIntroducerId As String
    Dim strIntroducerMCCB As String
    Dim strFilter As String
    
    'Set xmlCaseTrackNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/CUSTOMER|INTERMEDIARY")
    Set xmlCaseTrackNode = xmlGetMandatoryNode(vxmlRequestNode, "/REQUEST/CUSTOMER|INTRODUCER")
    'strSource = xmlGetAttributeText(xmlCaseTrackNode, "CUSTOMERNUMBER")
    
    If xmlCaseTrackNode.nodeName = "CUSTOMER" Then
        strSchema = "FINDAPPFORCASETRACK"
        strSource = xmlGetAttributeText(xmlCaseTrackNode, "CUSTOMERNUMBER")
        If Len(strSource) = 0 Then
            errThrowError strFunctionName, oeMissingPrimaryTag, "@CUSTOMERNUMBER must be present"
        End If
    Else
        strSchema = "FINDAPPFORCASETRACKINTERMED"
        
        strIntroducerId = xmlGetAttributeText(xmlCaseTrackNode, "INTRODUCERID")
        strIntroducerMCCB = xmlGetAttributeText(xmlCaseTrackNode, "INTRODUCERMCCB")
        
        If Len(strIntroducerId) = 0 And Len(strIntroducerMCCB) = 0 Then
            errThrowError strFunctionName, oeMissingPrimaryTag, "@INTRODUCERID or @INTRODUCERMCCB must be present"
        End If
        
        'create the filter to search the id's and mccb's for each level
        If Len(strIntroducerId) > 0 Then
            strFilter = "(INTRODUCERIDLEVEL1 = '" & strIntroducerId & "' OR INTRODUCERIDLEVEL2 = '" & strIntroducerId & "' OR INTRODUCERIDLEVEL3 = '" & strIntroducerId & "')"
        End If
        
        If Len(strIntroducerMCCB) > 0 Then
            'if a filter has already been produced bung an and on it
            If Len(strFilter) > 0 Then
                strFilter = strFilter & " AND "
            End If
            strFilter = "(INTRODUCERMCCBLEVEL1 = '" & strIntroducerMCCB & "' OR INTRODUCERMCCBLEVEL2 = '" & strIntroducerMCCB & "' OR INTRODUCERMCCBLEVEL3= '" & strIntroducerMCCB & "')"
        End If
    End If
    
    'If Len(strSource) > 0 Then
    '    strSchema = "FINDAPPFORCASETRACK"
    'Else
    '    strSource = xmlGetAttributeText(xmlCaseTrackNode, "INTERMEDIARYGUID")
    '
    '    If (Len(strSource) = 0) Then
    '        errThrowError strFunctionName, oeMissingPrimaryTag, "CUSTOMERNUMBER or INTERMEDIARYGUID must be present"
    '    End If
    '
    'End If
    
    'adoGetAsXML vxmlRequestParentNode, vxmlResponseNode, "APPLICATIONREVIEWHISTORY"
    'adoGetAsXML xmlCaseTrackNode, vxmlResponseNode, strSchema
    
    adoGetAsXML xmlCaseTrackNode, vxmlResponseNode, strSchema, strFilter
    
    'MO 24/10/2002 BMIDS00652 - End
    
    ' SYS1839 Add RecordNotFound
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If


    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

FindApplicationsForCaseTrackingExit:

    Set xmlCaseTrackNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub ICaseTrackingBO_GetCaseTrackingAllData(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
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
    
    Const strFunctionName As String = "ICaseTrackingBO_GetCaseTrackingAllData"
    Dim objCaseTrackingDO As ICaseTrackingDO
    
    Dim xmlApplicationNode As IXMLDOMNode
    
    Set xmlApplicationNode = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
 
    'Set xmlThisRequestDoc = New DOMDocument

    If objCaseTrackingDO Is Nothing Then
        Set objCaseTrackingDO = gobjContext.CreateInstance(App.Title & ".CaseTrackingDO")
    End If
  
    objCaseTrackingDO.GetCaseTrackingAllData xmlApplicationNode, vxmlResponseNode
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetCaseTrackingAllDataExit:
    Set objCaseTrackingDO = Nothing
    Set xmlApplicationNode = Nothing

    errCheckError strFunctionName, TypeName(Me)
End Sub

