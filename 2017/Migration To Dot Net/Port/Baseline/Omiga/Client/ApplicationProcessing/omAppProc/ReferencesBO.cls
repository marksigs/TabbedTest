VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "referencesBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ReferencesBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling
'Description:   Reference Checks Business Object.
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'SA    16/05/01    SYS2256 Plugging in of rules component for Employers/PrevEmployers
'                   Lenders/PrevLenders/Accountants/Landlords/prevlandlords refs.
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl
Implements IReferencesBO

Private gobjContext As ObjectContext

Private gobjTxBO As IReferencesTxBO
Private gobjDO As IReferencesDO


'---------------------------------------------------------------------------------------------
'BMIDS Specific History:

'Prog   Date        AQR         Description
'MV     14/08/2002  BMIDS00333  Modified GetApplicationData() - Core Ref AQR: SYS4919
'----------------------------------------------------------------------------------------------

Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Private Sub IReferencesBO_CreateEmployersReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Creates specified EmployersReference entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo CreateEmployersReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_CreateEmployersReference"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "EMPLOYERSREFERENCE"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.CreateEmployersReference vxmlRequestNode, vxmlResponseNode
    
CreateEmployersReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_CreateAccountantsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified AccountantsReference entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo CreateAccountantsReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_CreateAccountantsRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "ACCOUNTANTREF"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.CreateAccountantsReference vxmlRequestNode, vxmlResponseNode
    
CreateAccountantsReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_CreateLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified LandlordsReference entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo CreateLandlordsReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_CreateLandlordsRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "LANDLORDSREF"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.CreateLandlordsReference vxmlRequestNode, vxmlResponseNode
    
CreateLandlordsReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_CreateCurrEmployersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified CurrentEmployersRef entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo CreateCurrEmployersRefExit
    
    Const strFunctionName As String = "IReferencesBO_CreateCurrEmployersRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CURRENTEMPLOYERSREF"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.CreateCurrEmployersRef vxmlRequestNode, vxmlResponseNode
    
CreateCurrEmployersRefExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_CreatePrevEmployersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified PreviousEmployersRef entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo CreatePrevEmployersRefExit
    
    Const strFunctionName As String = "IReferencesBO_CreatePrevEmployersRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "PREVIOUSEMPLOYERSREF"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.CreatePrevEmployersRef vxmlRequestNode, vxmlResponseNode
    
CreatePrevEmployersRefExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_CreateLendersReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified LendersReference entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo CreateLendersReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_CreateLendersReference"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "LENDERSREFERENCE"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.CreateLendersReference vxmlRequestNode, vxmlResponseNode
    
CreateLendersReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_CreateCurrLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified CurrentLandlordsRef entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo CreateCurrLandlordsRefExit
    
    Const strFunctionName As String = "IReferencesBO_CreateCurrLandlordsRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CURRENTLANDLORDSREF"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.CreateCurrLandlordsRef vxmlRequestNode, vxmlResponseNode
    
CreateCurrLandlordsRefExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_CreatePrevLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified PreviousLandlordsRef entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo CreatePrevLandlordsRefExit
    
    Const strFunctionName As String = "IReferencesBO_CreatePrevLandlordsRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "PREVIOUSLANDLORDSREF"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.CreatePrevLandlordsRef vxmlRequestNode, vxmlResponseNode
    
CreatePrevLandlordsRefExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_CreateCurrLendersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified CurrentLendersRef entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo CreateCurrLendersRefExit
    
    Const strFunctionName As String = "IReferencesBO_CreateCurrLendersRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "CURRENTLENDERSREF"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.CreateCurrLendersRef vxmlRequestNode, vxmlResponseNode
    
CreateCurrLendersRefExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_CreatePrevLendersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified PreviousLendersRef entities.
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------

    On Error GoTo CreatePrevLendersRefExit
    
    Const strFunctionName As String = "IReferencesBO_CreatePrevLendersRef"
    
    ' validate REQUEST
    xmlGetMandatoryNode vxmlRequestNode, "PREVIOUSLENDERSREF"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.CreatePrevLendersRef vxmlRequestNode, vxmlResponseNode
    
CreatePrevLendersRefExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_GetEmployersReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified EMPLOYERSREFERENCE entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       EMPLOYERSREFERENCE node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetEmployersReferenceExit

    Const strFunctionName As String = "IReferencesBO_GetEmployersReference"

    Dim xmlEmployersRefNode As IXMLDOMNode

    ' validate REQUEST & select CASETASK node
    Set xmlEmployersRefNode = xmlGetMandatoryNode(vxmlRequestNode, "EMPLOYERSREFERENCE")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    End If

    gobjDO.GetEmployersReference xmlEmployersRefNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetEmployersReferenceExit:

    Set xmlEmployersRefNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_GetCurrEmployersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified EMPLOYERSREFERENCE and CurrentEmployersRef entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       CURRENTEMPLOYERSREF node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetCurrEmployersRefExit

    Const strFunctionName As String = "IReferencesBO_GetCurrEmployersRef"

    Dim xmlEmployersRefNode As IXMLDOMNode

    ' validate REQUEST & select CASETASK node
    Set xmlEmployersRefNode = xmlGetMandatoryNode(vxmlRequestNode, "CURRENTEMPLOYERSREF")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    End If

    gobjDO.GetCurrEmployersRef xmlEmployersRefNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetCurrEmployersRefExit:

    Set xmlEmployersRefNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_GetPrevEmployersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified EMPLOYERSREFERENCE and PreviousEmployersRef entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       PREVIOUSEMPLOYERSREF node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetPrevEmployersRefExit

    Const strFunctionName As String = "IReferencesBO_GetPrevEmployersRef"

    Dim xmlEmployersRefNode As IXMLDOMNode

    ' validate REQUEST & select PREVIOUSEMPLOYERSREF node
    Set xmlEmployersRefNode = xmlGetMandatoryNode(vxmlRequestNode, "PREVIOUSEMPLOYERSREF")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    End If

    gobjDO.GetPrevEmployersRef xmlEmployersRefNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetPrevEmployersRefExit:

    Set xmlEmployersRefNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_GetAccountantsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified ACCOUNTANTSREFERENCE entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       ACCOUNTANTSREFERENCE node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetAccountantsReferenceExit

    Const strFunctionName As String = "IReferencesBO_GetAccountantsRef"

    Dim xmlAccountantsRefNode As IXMLDOMNode

    Set xmlAccountantsRefNode = xmlGetMandatoryNode(vxmlRequestNode, "ACCOUNTANT")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    End If

    gobjDO.GetAccountantsRef xmlAccountantsRefNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetAccountantsReferenceExit:

    Set xmlAccountantsRefNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_GetLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified LANDLORDSREFERENCE entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       LANDLORDSREFERENCE node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetLandlordsReferenceExit

    Const strFunctionName As String = "IReferencesBO_GetLandlordsRef"

    Dim xmlLandlordsRefNode As IXMLDOMNode

    Set xmlLandlordsRefNode = xmlGetMandatoryNode(vxmlRequestNode, "LANDLORDSREF")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    End If

    gobjDO.GetLandlordsRef xmlLandlordsRefNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetLandlordsReferenceExit:

    Set xmlLandlordsRefNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_GetCurrLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified LANDLORDSREFERENCE and CurrentLandlordsRef entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       CURRENTLANDLORDSREF node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetCurrLandlordsRefExit

    Const strFunctionName As String = "IReferencesBO_GetCurrLandlordsRef"

    Dim xmlLandlordsRefNode As IXMLDOMNode

    Set xmlLandlordsRefNode = xmlGetMandatoryNode(vxmlRequestNode, "CURRENTLANDLORDSREF")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    End If

    gobjDO.GetCurrLandlordsRef xmlLandlordsRefNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetCurrLandlordsRefExit:

    Set xmlLandlordsRefNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_GetPrevLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified LANDLORDSREFERENCE and PreviousLandlordsRef entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       PREVIOUSLANDLORDSREF node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetPrevLandlordsRefExit

    Const strFunctionName As String = "IReferencesBO_GetPrevLandlordsRef"

    Dim xmlLandlordsRefNode As IXMLDOMNode

    Set xmlLandlordsRefNode = xmlGetMandatoryNode(vxmlRequestNode, "PREVIOUSLANDLORDSREF")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    End If

    gobjDO.GetPrevLandlordsRef xmlLandlordsRefNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetPrevLandlordsRefExit:

    Set xmlLandlordsRefNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_GetLendersReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified LENDERSREFERENCE entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       LENDERSREFERENCE node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetLendersReferenceExit

    Const strFunctionName As String = "IReferencesBO_GetLendersReference"

    Dim xmlLendersRefNode As IXMLDOMNode

    ' validate REQUEST & select CASETASK node
    Set xmlLendersRefNode = xmlGetMandatoryNode(vxmlRequestNode, "LENDERSREFERENCE")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    End If

    gobjDO.GetLendersReference xmlLendersRefNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetLendersReferenceExit:

    Set xmlLendersRefNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_GetCurrLendersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified LENDERSREFERENCE and CurrentLendersRef entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       CURRENTLENDERSREF node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetCurrLendersRefExit

    Const strFunctionName As String = "IReferencesBO_GetCurrLendersRef"

    Dim xmlLendersRefNode As IXMLDOMNode

    ' validate REQUEST & select CASETASK node
    Set xmlLendersRefNode = xmlGetMandatoryNode(vxmlRequestNode, "CURRENTLENDERSREF")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    End If

    gobjDO.GetCurrLendersRef xmlLendersRefNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetCurrLendersRefExit:

    Set xmlLendersRefNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_GetPrevLendersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified LENDERSREFERENCE and PreviousLendersRef entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       PREVIOUSLENDERSREF node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetPrevLendersRefExit

    Const strFunctionName As String = "IReferencesBO_GetPrevLendersRef"

    Dim xmlLendersRefNode As IXMLDOMNode

    Set xmlLendersRefNode = xmlGetMandatoryNode(vxmlRequestNode, "PREVIOUSLENDERSREF")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    End If

    gobjDO.GetPrevLendersRef xmlLendersRefNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetPrevLendersRefExit:

    Set xmlLendersRefNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_UpdateEmployersReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   update specified EMPLOYERSREFERENCE entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------
    
    On Error GoTo UpdateEmployersReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_UpdateEmployersReference"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.UpdateEmployersReference vxmlRequestNode, vxmlResponseNode
    
UpdateEmployersReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesBO_UpdateLendersReference( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   update specified LENDERSREFERENCE entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------
    
    On Error GoTo UpdateLendersReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_UpdateLendersReference"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.UpdateLendersReference vxmlRequestNode, vxmlResponseNode
    
UpdateLendersReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesBO_UpdateCurrEmployersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   update specified EMPLOYERSREFERENCE and CURRENTEMPLOYERSREF entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------
    
    On Error GoTo UpdateCurrEmployersReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_UpdateCurrEmployersRef"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.UpdateCurrEmployersRef vxmlRequestNode, vxmlResponseNode
    
UpdateCurrEmployersReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesBO_UpdatePrevEmployersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   update specified EMPLOYERSREFERENCE and PREVIOUSEMPLOYERSREF entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------
    
    On Error GoTo UpdatePrevEmployersReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_UpdatePrevEmployersRef"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.UpdatePrevEmployersRef vxmlRequestNode, vxmlResponseNode
    
UpdatePrevEmployersReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesBO_UpdateAccountantsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   update specified ACCOUNTANTSREFERENCE entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------
    
    On Error GoTo UpdateAccountantsReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_UpdateAccountantsRef"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.UpdateAccountantsReference vxmlRequestNode, vxmlResponseNode
    
UpdateAccountantsReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesBO_UpdateLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   update specified LANDLORDSREFERENCE entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------
    
    On Error GoTo UpdateLandlordsReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_UpdateLandlordsRef"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.UpdateLandlordsReference vxmlRequestNode, vxmlResponseNode
    
UpdateLandlordsReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesBO_UpdateCurrLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   update specified LANDLORDSREFERENCE and CURRENTLANDLORDSREF entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------
    
    On Error GoTo UpdateCurrLandlordsReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_UpdateCurrLandlordsRef"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.UpdateCurrLandlordsRef vxmlRequestNode, vxmlResponseNode
    
UpdateCurrLandlordsReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesBO_UpdatePrevLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   update specified LANDLORDSREFERENCE and PREVIOUSLANDLORDSREF entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------
    
    On Error GoTo UpdatePrevLandlordsReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_UpdatePrevLandlordsRef"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.UpdatePrevLandlordsRef vxmlRequestNode, vxmlResponseNode
    
UpdatePrevLandlordsReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesBO_UpdateCurrLendersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   update specified LENDERSREFERENCE and CURRENTLENDERSREF entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------
    
    On Error GoTo UpdateCurrLendersReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_UpdateCurrLendersRef"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.UpdateCurrLendersRef vxmlRequestNode, vxmlResponseNode
    
UpdateCurrLendersReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesBO_UpdatePrevLendersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   update specified LENDERSREFERENCE and PREVIOUSLENDERSREF entity,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'------------------------------------------------------------------------------------------
    
    On Error GoTo UpdatePrevLendersReferenceExit
    
    Const strFunctionName As String = "IReferencesBO_UpdatePrevLendersRef"
    
    If gobjTxBO Is Nothing Then
        Set gobjTxBO = gobjContext.CreateInstance("omAppProc.ReferencesTxBO")
    End If
    
    gobjTxBO.UpdatePrevLendersRef vxmlRequestNode, vxmlResponseNode
    
UpdatePrevLendersReferenceExit:
    
    errCheckError strFunctionName, TypeName(Me)
End Sub

Private Sub IReferencesBO_GetEmpDetailsForRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified entity from the EMPLOYEDDETAILSFORREF view,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       EMPLOYMENTDETAIL node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetEmpDetailsExit

    Const strFunctionName As String = "IReferencesBO_GetEmpDetailsForRef"

    Dim xmlEmpDetailsNode As IXMLDOMNode

    Set xmlEmpDetailsNode = xmlGetMandatoryNode(vxmlRequestNode, "EMPDETAILSFORREF")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    End If

    gobjDO.GetEmpDetailsForRef xmlEmpDetailsNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetEmpDetailsExit:

    Set xmlEmpDetailsNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_GetLoanDetailsForRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified entity from the LOANDETAILSFORREF view,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       LOANDETAIL node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetLoanDetailsExit

    Const strFunctionName As String = "IReferencesBO_GetLoanDetailsForRef"

    Dim xmlLoanDetailsNode As IXMLDOMNode

    Set xmlLoanDetailsNode = xmlGetMandatoryNode(vxmlRequestNode, "LOANDETAIL")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    End If

    gobjDO.GetLoanDetailsForRef xmlLoanDetailsNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetLoanDetailsExit:

    Set xmlLoanDetailsNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_GetTenancyDetailsForRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Retrieve specified entity from the TENANCYDETAILSFORREF view,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       TENANCYDETAIL node appended to vxmlResponseNode on exit
' exceptions:
'       oeRecordNotFound
'------------------------------------------------------------------------------------------

    On Error GoTo GetTenancyDetailsExit

    Const strFunctionName As String = "IReferencesBO_GetTenancyDetailsForRef"

    Dim xmlTenancyDetailsNode As IXMLDOMNode

    Set xmlTenancyDetailsNode = xmlGetMandatoryNode(vxmlRequestNode, "TENANCYDETAIL")

    If gobjDO Is Nothing Then
        Set gobjDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    End If

    gobjDO.GetTenancyDetailsForRef xmlTenancyDetailsNode, vxmlResponseNode

    ' if RESPONSE nodes has no child nodes,
    ' then no records found matching REQUEST criteria
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If

GetTenancyDetailsExit:

    Set xmlTenancyDetailsNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_ValidateCurrEmployersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Validate the CurrEmployersRef using APRulesBO,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       APPSTATUS node appended to vxmlResponseNode on exit
'------------------------------------------------------------------------------------------

    On Error GoTo ValidateCurrEmployersRefExit

    Const strFunctionName As String = "IReferencesBO_ValidateCurrEmployersRef"
    
    CheckMandatoryAttributes "ValidateCurrEmployersRef", vxmlRequestNode

    RunRefRules vxmlRequestNode, vxmlResponseNode, _
                    "RunCurrEmployersRefRules", "CURRENTEMPLOYERSREF", _
                        "30", "CURRENTEMPLOYERREFRULE"
    
ValidateCurrEmployersRefExit:

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_ValidatePrevEmployersRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Validate the PrevEmployersRef using APRulesBO,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       APPSTATUS node appended to vxmlResponseNode on exit
'------------------------------------------------------------------------------------------
    
    On Error GoTo ValidatePrevEmployersRefExit

    Const strFunctionName As String = "IReferencesBO_ValidatePrevEmployersRef"
    
    CheckMandatoryAttributes "ValidatePrevEmployersRef", vxmlRequestNode

    RunRefRules vxmlRequestNode, vxmlResponseNode, _
                        "RunPrevEmployersRefRules", _
                        "VALPREVEMP", "40", "PREVIOUSEMPLOYERREFRULE"
       
ValidatePrevEmployersRefExit:

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_ValidateCurrLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByRef vxmlResponseNode As IXMLDOMNode)
' header----------------------------------------------------------------------------------
' description:
'   Validate the CurrLandlordsRef using APRulesBO,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       APPSTATUS node appended to vxmlResponseNode on exit
'------------------------------------------------------------------------------------------
    
    On Error GoTo ValidateCurrLandlordsRefExit

    Const strFunctionName As String = "IReferencesBO_ValidateCurrLandlordsRef"
    
    CheckMandatoryAttributes "ValidateCurrLAndlordsRef", vxmlRequestNode
    
    RunRefRules vxmlRequestNode, vxmlResponseNode, _
                        "RunCurrLandlordRefRules", _
                        "VALIDATELANDLORDREF", "70", "CURRENTLANDLORDEFRULE"
           
    
ValidateCurrLandlordsRefExit:
            
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_ValidatePrevLandlordsRef( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Validate the PrevLandlordsRef using APRulesBO,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       APPSTATUS node appended to vxmlResponseNode on exit
'------------------------------------------------------------------------------------------
    
    On Error GoTo ValidatePrevLandlordsRefExit

    Const strFunctionName As String = "IReferencesBO_ValidatePrevLandlordsRef"
    
    CheckMandatoryAttributes "ValidatePrevLAndlordsRef", vxmlRequestNode
    
    RunRefRules vxmlRequestNode, vxmlResponseNode, _
                        "RunPrevLandlordRefRules", _
                        "VALIDATEPREVLANDREF", "70", "PREVIOUSLANDLORDEFRULE"
    
ValidatePrevLandlordsRefExit:

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_ValidateCurrLendersRef(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Validate the CurrLendersRef using APRulesBO,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       APPSTATUS node appended to vxmlResponseNode on exit
'------------------------------------------------------------------------------------------
    
    On Error GoTo ValidateCurrLendersRefExit

    Const strFunctionName As String = "IReferencesBO_ValidateCurrLendersRef"
    
    CheckMandatoryAttributes "ValidateCurrLendersRef", vxmlRequestParentNode
    
    RunRefRules vxmlRequestParentNode, vxmlResponseNode, _
            "RunCurrLendersRefRules", "VALLENDREF", "50", "CURRENTLENDERREFRULE"
    
ValidateCurrLendersRefExit:

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_ValidatePrevLendersRef(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Validate the PrevLendersRef using APRulesBO,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       APPSTATUS node appended to vxmlResponseNode on exit
'------------------------------------------------------------------------------------------

    On Error GoTo ValidatePrevLendersRefExit

    Const strFunctionName As String = "IReferencesBO_ValidatePrevLendersRef"
    
    CheckMandatoryAttributes "ValidatePrevLendersRef", vxmlRequestParentNode
    
    RunRefRules vxmlRequestParentNode, vxmlResponseNode, _
        "RunPrevLendersRefRules", "VALPREVLENDREF", "60", _
        "PREVIOUSLENDERREFRULE"
    
ValidatePrevLendersRefExit:

    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IReferencesBO_ValidateAccountantsRef(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Validate the AccountantsRef using APRulesBO,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
'       APPSTATUS node appended to vxmlResponseNode on exit
'------------------------------------------------------------------------------------------
    
    On Error GoTo ValidateAccountantsRefExit

    Const strFunctionName As String = "IReferencesBO_ValidateAccountantsRef"
    
    CheckMandatoryAttributes "ValidateAccountantsRef", vxmlRequestParentNode
    
    RunRefRules vxmlRequestParentNode, vxmlResponseNode, _
            "RunAccountantsRefRules", "VALACCTREF", "20", "ACCOUNTANTSREFRULE"
   
ValidateAccountantsRefExit:
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub RunRefRules( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode, _
    ByVal strOperation As String, _
    ByVal strRequestNodeToSelect As String, _
    ByVal strReviewReason As String, _
    ByVal strResponseNodeToSelect)
' header ----------------------------------------------------------------------------------
' description:
'   Run Rules using APRulesBO,
' pass:
'   vxmlRequestNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
'   strRequestNodeToSelect
'       Name of Requestnode to select
'   strOperation
'       Reference to validate
'   strReviewReason
'       Application Review Reason
'   strResponseNodeToSelect
'       Name of Responsenode to select
' return:
'       none
'       APPSTATUS node appended to vxmlResponseNode on exit
'------------------------------------------------------------------------------------------
    On Error GoTo RunRefRulesExit
    Const strFunctionName As String = "RunRefRules"

    Dim xmlThisRequestDoc As FreeThreadedDOMDocument40, _
        xmlThisResponseDoc As FreeThreadedDOMDocument40, _
        xmlThisRequestNode As IXMLDOMNode, _
        xmlApplicationNode As IXMLDOMNode, _
        xmlThisElem As IXMLDOMElement, _
        xmlResponseList As IXMLDOMNodeList, _
        xmlThisNode As IXMLDOMNode
        
        
    Dim objAPRules As Object
    
    Set xmlThisRequestDoc = New FreeThreadedDOMDocument40
    xmlThisRequestDoc.validateOnParse = False
    xmlThisRequestDoc.setProperty "NewParser", True
    xmlThisRequestDoc.async = False
    
    Set xmlThisRequestNode = xmlThisRequestDoc.appendChild(vxmlRequestNode.cloneNode(False))
    ' set up the name of the reference to validate
    xmlThisRequestNode.Attributes.getNamedItem("OPERATION").Text = strOperation
    
    ' Now we need to add on the attributes from the Application Node
    Set xmlApplicationNode = vxmlRequestNode.selectSingleNode(strRequestNodeToSelect)
    xmlCopyAttribute xmlApplicationNode, xmlThisRequestNode, "CUSTOMERNUMBER"
    xmlCopyAttribute xmlApplicationNode, xmlThisRequestNode, "CUSTOMERVERSIONNUMBER"
    xmlCopyAttribute xmlApplicationNode, xmlThisRequestNode, "EMPLOYMENTSEQUENCENUMBER"
    xmlCopyAttribute xmlApplicationNode, xmlThisRequestNode, "CUSTOMERADDRESSSEQUENCENUMBER"
    xmlCopyAttribute xmlApplicationNode, xmlThisRequestNode, "ACCOUNTGUID"   ' JR - SYS2256 for Curr/Prev LendersRefRules
    
    ' Add on any Application Data
    Set xmlApplicationNode = vxmlRequestNode.selectSingleNode("CASETASK")
    GetApplicationData xmlApplicationNode, xmlThisRequestNode
    
    Set objAPRules = gobjContext.CreateInstance("omAPRules.APRulesBO")
    Set xmlThisResponseDoc = New FreeThreadedDOMDocument40
    xmlThisResponseDoc.validateOnParse = False
    xmlThisResponseDoc.setProperty "NewParser", True
    xmlThisResponseDoc.async = False
    xmlThisResponseDoc.loadXML objAPRules.RunRules(xmlThisRequestDoc.xml)
    Set objAPRules = Nothing
     
    'Copy attrib for response node
    xmlCopyAttribute xmlThisResponseDoc.selectSingleNode("RESPONSE"), vxmlResponseNode, "TYPE"
     
    'Set xmlThisElem = xmlThisResponseDoc.selectSingleNode(".//" & strResponseNodeToSelect)
    Set xmlResponseList = xmlThisResponseDoc.selectNodes(".//" & strResponseNodeToSelect)
    If Not xmlResponseList Is Nothing Then
        For Each xmlThisNode In xmlResponseList
            vxmlResponseNode.appendChild xmlThisNode
        Next
        ' Add on Application Node to use to create App Review History (back in omAppProc)
        vxmlResponseNode.appendChild xmlApplicationNode
    End If
    

RunRefRulesExit:
    Set objAPRules = Nothing
        
    Set xmlThisRequestNode = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlThisElem = Nothing
    Set xmlThisRequestDoc = Nothing
    Set xmlThisResponseDoc = Nothing

    errCheckError strFunctionName, TypeName(Me)
    
End Sub

Private Sub GetApplicationData(ByVal vxmlApplicationSrceNode As IXMLDOMNode, ByVal vxmlParentNode As IXMLDOMNode)
    
    On Error GoTo GetApplicationDataExit
    Const cstrFunctionName As String = "GetApplicationData"
    
    Dim xmlRequestDoc As FreeThreadedDOMDocument40
    
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim xmlAttrib As IXMLDOMAttribute
    
    Dim objRB As Object
    
    Set xmlRequestDoc = New FreeThreadedDOMDocument40
    xmlRequestDoc.validateOnParse = False
    xmlRequestDoc.setProperty "NewParser", True
    xmlRequestDoc.async = False
    
    Set xmlElem = xmlRequestDoc.createElement("REQUEST")
    Set xmlNode = xmlRequestDoc.appendChild(xmlElem)
    
    Set xmlAttrib = xmlRequestDoc.createAttribute("COMBOLOOKUP")
    xmlAttrib.Text = "NO"
    xmlNode.Attributes.setNamedItem xmlAttrib
    
    'MV - 14/08/2002 - BMIDS00333 - Core Upgrade
    Set xmlAttrib = xmlRequestDoc.createAttribute("RB_TEMPLATE")
    xmlAttrib.Text = "APRefsRBTemplate"
    xmlNode.Attributes.setNamedItem xmlAttrib
    
    Set xmlElem = xmlRequestDoc.createElement("APPLICATION")
    
    xmlElem.setAttribute "_SCHEMA_", "APPLICATION"
    
    xmlElem.setAttribute _
        "APPLICATIONNUMBER", _
        vxmlApplicationSrceNode.selectSingleNode("@CASEID").Text
    
    If vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONFACTFINDNUMBER") Is Nothing Then
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
    Else
        xmlElem.setAttribute _
            "APPLICATIONFACTFINDNUMBER", _
            vxmlApplicationSrceNode.selectSingleNode("@APPLICATIONFACTFINDNUMBER").Text
    End If
    
    xmlNode.appendChild xmlElem
    
    Debug.Print "GetApplicationData request"
    Debug.Print xmlRequestDoc.xml
    
    Set objRB = gobjContext.CreateInstance("omRB.OmRequestDO")
    xmlRequestDoc.loadXML objRB.OmDataRequest(xmlRequestDoc.xml)
    Set objRB = Nothing
    
    Debug.Print "GetApplicationData response"
    Debug.Print xmlRequestDoc.xml
    
    errCheckXMLResponseNode xmlRequestDoc.selectSingleNode("RESPONSE"), , True
    
    If xmlRequestDoc.selectSingleNode("RESPONSE/APPLICATION") Is Nothing Then
        errThrowError _
            "GetApplicationData", _
            oeXMLMissingElement, _
            "No APPLICATION detail"
    End If
    
    'MV - 14/08/2002 - BMIDS00333 - Core Upgrade
    For Each xmlNode In xmlRequestDoc.selectSingleNode("RESPONSE").childNodes
        vxmlParentNode.appendChild xmlNode.cloneNode(True)
    Next
    
GetApplicationDataExit:
    
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlAttrib = Nothing
    Set xmlRequestDoc = Nothing
    
    errCheckError cstrFunctionName

End Sub

Private Sub CheckMandatoryAttributes(ByVal cMethodName As String, ByVal vxmlRequestNode As IXMLDOMNode)

Dim xmlIn As IXMLDOMNode

On Error GoTo CheckMandatoryAttributesExit
Const cstrFunctionName As String = "CheckMandatoryAttributes"

' Select correct node
Select Case UCase$(cMethodName)
    Case "VALIDATECURRLANDLORDSREF"
        Set xmlIn = xmlGetMandatoryNode(vxmlRequestNode, "VALIDATELANDLORDREF")
    Case "VALIDATEPREVLANDLORDSREF"
        Set xmlIn = xmlGetMandatoryNode(vxmlRequestNode, "VALIDATEPREVLANDREF")
    Case "VALIDATECURREMPLOYERSREF"
        Set xmlIn = xmlGetMandatoryNode(vxmlRequestNode, "CURRENTEMPLOYERSREF")
    Case "VALIDATEPREVEMPLOYERSREF"
        Set xmlIn = xmlGetMandatoryNode(vxmlRequestNode, "VALPREVEMP")
    Case "VALIDATEPREVLENDERSREF"
        Set xmlIn = xmlGetMandatoryNode(vxmlRequestNode, "VALPREVLENDREF")
    Case "VALIDATECURRLENDERSREF"
        Set xmlIn = xmlGetMandatoryNode(vxmlRequestNode, "VALLENDREF")
    Case "VALIDATEACCOUNTANTSREF"
        Set xmlIn = xmlGetMandatoryNode(vxmlRequestNode, "VALACCTREF")
          
End Select
' Check mandatory attributes
Select Case UCase$(cMethodName)
    Case "VALIDATECURRLANDLORDSREF", "VALIDATEPREVLANDLORDSREF"
        xmlCheckMandatoryAttribute xmlIn, "CUSTOMERNUMBER"
        xmlCheckMandatoryAttribute xmlIn, "CUSTOMERVERSIONNUMBER"
        xmlCheckMandatoryAttribute xmlIn, "CUSTOMERADDRESSSEQUENCENUMBER"
    
    Case "VALIDATECURREMPLOYERSREF", "VALIDATEPREVEMPLOYERSREF", "VALIDATEACCOUNTANTSREF"
        xmlCheckMandatoryAttribute xmlIn, "CUSTOMERNUMBER"
        xmlCheckMandatoryAttribute xmlIn, "CUSTOMERVERSIONNUMBER"
        xmlCheckMandatoryAttribute xmlIn, "EMPLOYMENTSEQUENCENUMBER"
    
    Case "VALIDATEPREVLENDERSREF", "VALIDATECURRLENDERSREF"
        xmlCheckMandatoryAttribute xmlIn, "ACCOUNTGUID"
        xmlCheckMandatoryAttribute xmlIn, "CUSTOMERNUMBER"
        xmlCheckMandatoryAttribute xmlIn, "CUSTOMERVERSIONNUMBER"
        
End Select
    
' Now select Casetask node and check mandatory attribs.
If UCase$(cMethodName) <> "VALIDATEPREVLENDERSREF" And _
    UCase$(cMethodName) <> "VALIDATECURRLENDERSREF" Then

    Set xmlIn = xmlGetMandatoryNode(vxmlRequestNode, "CASETASK")
        
    xmlCheckMandatoryAttribute xmlIn, "SOURCEAPPLICATION"
    xmlCheckMandatoryAttribute xmlIn, "APPLICATIONFACTFINDNUMBER"
    xmlCheckMandatoryAttribute xmlIn, "ACTIVITYID"
    xmlCheckMandatoryAttribute xmlIn, "CASEID"
    xmlCheckMandatoryAttribute xmlIn, "TASKID"
    xmlCheckMandatoryAttribute xmlIn, "TASKINSTANCE"
    xmlCheckMandatoryAttribute xmlIn, "CASETASKNAME"
    
End If

CheckMandatoryAttributesExit:

    errCheckError cstrFunctionName


End Sub

