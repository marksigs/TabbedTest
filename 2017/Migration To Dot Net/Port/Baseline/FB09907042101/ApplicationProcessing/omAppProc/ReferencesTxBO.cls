VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "ReferencesTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IReferencesTxBO
Implements ObjectControl

Private gobjContext As ObjectContext
Private Sub ObjectControl_Activate()
    Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set gobjContext = Nothing
End Sub

Private Sub IReferencesTxBO_CreateEmployersReference( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified CreateEmployersReference entities.
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

    On Error GoTo CreateEmployersReferenceVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_CreateEmployersReference"
    
    Dim xmlEmployersRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlEmployersRefList = vxmlRequestParentNode.selectNodes("EMPLOYERSREFERENCE")
        
    If xmlEmployersRefList.length > 0 Then
        ' call CreateEmployersReference if EMPLOYERSREFERENCE node in request
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlEmployersRefList
            objDO.CreateEmployersReference xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
CreateEmployersReferenceExit:
    
    Set objDO = Nothing
    Set xmlEmployersRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateEmployersReferenceVbErr:

    gobjContext.SetAbort
    
    GoTo CreateEmployersReferenceExit

End Sub

Private Sub IReferencesTxBO_CreateAccountantsReference( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
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

    On Error GoTo CreateAccountantsReferenceVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_CreateAccountantsReference"
    
    Dim xmlAccountantsRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlAccountantsRefList = vxmlRequestParentNode.selectNodes("ACCOUNTANTREF")
        
    If xmlAccountantsRefList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlAccountantsRefList
            objDO.CreateAccountantsReference xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
CreateAccountantsReferenceExit:
    
    Set objDO = Nothing
    Set xmlAccountantsRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateAccountantsReferenceVbErr:

    gobjContext.SetAbort
    
    GoTo CreateAccountantsReferenceExit

End Sub

Private Sub IReferencesTxBO_CreateLandlordsReference( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
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

    On Error GoTo CreateLandlordsReferenceVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_CreateLandlordsReference"
    
    Dim xmlLandlordsRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlLandlordsRefList = vxmlRequestParentNode.selectNodes("LANDLORDSREF")
        
    If xmlLandlordsRefList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlLandlordsRefList
            objDO.CreateLandlordsReference xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
CreateLandlordsReferenceExit:
    
    Set objDO = Nothing
    Set xmlLandlordsRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateLandlordsReferenceVbErr:

    gobjContext.SetAbort
    
    GoTo CreateLandlordsReferenceExit

End Sub

Private Sub IReferencesTxBO_CreateCurrEmployersRef( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

' header ----------------------------------------------------------------------------------
' description:
'   Creates specified CreateCurrEmployersRef entities.
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

    On Error GoTo CreateCurrEmployersRefVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_CreateCurrEmployersRef"
    
    Dim xmlEmployersRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlEmployersRefList = vxmlRequestParentNode.selectNodes("CURRENTEMPLOYERSREF")
        
    If xmlEmployersRefList.length > 0 Then
        ' call CreateCurrEmployersRef if CURRENTEMPLOYERSREF node in request
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlEmployersRefList
            objDO.CreateCurrEmployersRef xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
CreateCurrEmployersRefExit:
    
    Set objDO = Nothing
    Set xmlEmployersRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateCurrEmployersRefVbErr:

    gobjContext.SetAbort
    
    GoTo CreateCurrEmployersRefExit

End Sub

Private Sub IReferencesTxBO_CreatePrevEmployersRef( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
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

    On Error GoTo CreatePrevEmployersRefVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_CreatePrevEmployersRef"
    
    Dim xmlEmployersRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlEmployersRefList = vxmlRequestParentNode.selectNodes("PREVIOUSEMPLOYERSREF")
        
    If xmlEmployersRefList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlEmployersRefList
            objDO.CreatePrevEmployersRef xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
CreatePrevEmployersRefExit:
    
    Set objDO = Nothing
    Set xmlEmployersRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreatePrevEmployersRefVbErr:

    gobjContext.SetAbort
    
    GoTo CreatePrevEmployersRefExit

End Sub

Private Sub IReferencesTxBO_CreateLendersReference( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
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

    On Error GoTo CreateLendersReferenceVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_CreateLendersReference"
    
    Dim xmlLendersRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlLendersRefList = vxmlRequestParentNode.selectNodes("LENDERSREFERENCE")
        
    If xmlLendersRefList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlLendersRefList
            objDO.CreateLendersReference xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
CreateLendersReferenceExit:
    
    Set objDO = Nothing
    Set xmlLendersRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateLendersReferenceVbErr:

    gobjContext.SetAbort
    
    GoTo CreateLendersReferenceExit

End Sub

Private Sub IReferencesTxBO_CreateCurrLandlordsRef( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
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

    On Error GoTo CreateCurrLandlordsRefVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_CreateCurrLandlordsRef"
    
    Dim xmlLandlordsRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlLandlordsRefList = vxmlRequestParentNode.selectNodes("CURRENTLANDLORDSREF")
        
    If xmlLandlordsRefList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlLandlordsRefList
            objDO.CreateCurrLandlordsRef xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
CreateCurrLandlordsRefExit:
    
    Set objDO = Nothing
    Set xmlLandlordsRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateCurrLandlordsRefVbErr:

    gobjContext.SetAbort
    
    GoTo CreateCurrLandlordsRefExit

End Sub

Private Sub IReferencesTxBO_CreatePrevLandlordsRef( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
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

    On Error GoTo CreatePrevLandlordsRefVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_CreatePrevLandlordsRef"
    
    Dim xmlLandlordsRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlLandlordsRefList = vxmlRequestParentNode.selectNodes("PREVIOUSLANDLORDSREF")
        
    If xmlLandlordsRefList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlLandlordsRefList
            objDO.CreatePrevLandlordsRef xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
CreatePrevLandlordsRefExit:
    
    Set objDO = Nothing
    Set xmlLandlordsRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreatePrevLandlordsRefVbErr:

    gobjContext.SetAbort
    
    GoTo CreatePrevLandlordsRefExit

End Sub

Private Sub IReferencesTxBO_CreateCurrLendersRef( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
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

    On Error GoTo CreateCurrLendersRefVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_CreateCurrLendersRef"
    
    Dim xmlLendersRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlLendersRefList = vxmlRequestParentNode.selectNodes("CURRENTLENDERSREF")
        
    If xmlLendersRefList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlLendersRefList
            objDO.CreateCurrLendersRef xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
CreateCurrLendersRefExit:
    
    Set objDO = Nothing
    Set xmlLendersRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateCurrLendersRefVbErr:

    gobjContext.SetAbort
    
    GoTo CreateCurrLendersRefExit

End Sub

Private Sub IReferencesTxBO_CreatePrevLendersRef( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
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

    On Error GoTo CreatePrevLendersRefVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_CreatePrevLendersRef"
    
    Dim xmlLendersRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlLendersRefList = vxmlRequestParentNode.selectNodes("PREVIOUSLENDERSREF")
        
    If xmlLendersRefList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlLendersRefList
            objDO.CreatePrevLendersRef xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
CreatePrevLendersRefExit:
    
    Set objDO = Nothing
    Set xmlLendersRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreatePrevLendersRefVbErr:

    gobjContext.SetAbort
    
    GoTo CreatePrevLendersRefExit

End Sub

Private Sub IReferencesTxBO_UpdateEmployersReference( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateEmployersReferenceVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_UpdateEmployersReference"
    
    Dim xmlEmployersReferenceList As IXMLDOMNodeList
    Dim xmlEmployersReferenceNode As IXMLDOMNode
    
    Dim objDO As IReferencesDO
    
    Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    
    Set xmlEmployersReferenceList = vxmlRequestParentNode.selectNodes("EMPLOYERSREFERENCE")
    
    For Each xmlEmployersReferenceNode In xmlEmployersReferenceList
        objDO.UpdateEmployersReference xmlEmployersReferenceNode, vxmlResponseNode
    Next
    
    gobjContext.SetComplete
    
UpdateEmployersReferenceExit:
    
    Set objDO = Nothing
    Set xmlEmployersReferenceList = Nothing
    Set xmlEmployersReferenceNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdateEmployersReferenceVbErr:

    gobjContext.SetAbort

    GoTo UpdateEmployersReferenceExit
End Sub

Private Sub IReferencesTxBO_UpdateAccountantsReference( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateAccountantsReferenceVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_UpdateAccountantsReference"
    
    Dim xmlAccountantsReferenceList As IXMLDOMNodeList
    Dim xmlAccountantsReferenceNode As IXMLDOMNode
    
    Dim objDO As IReferencesDO
    
    Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    
    Set xmlAccountantsReferenceList = vxmlRequestParentNode.selectNodes("ACCOUNTANTREF")
    
    For Each xmlAccountantsReferenceNode In xmlAccountantsReferenceList
        objDO.UpdateAccountantsReference xmlAccountantsReferenceNode, vxmlResponseNode
    Next
    
    gobjContext.SetComplete
    
UpdateAccountantsReferenceExit:
    
    Set objDO = Nothing
    Set xmlAccountantsReferenceList = Nothing
    Set xmlAccountantsReferenceNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdateAccountantsReferenceVbErr:

    gobjContext.SetAbort

    GoTo UpdateAccountantsReferenceExit
End Sub

Private Sub IReferencesTxBO_UpdateLandlordsReference( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateLandlordsReferenceVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_UpdateLandlordsReference"
    
    Dim xmlLandlordsReferenceList As IXMLDOMNodeList
    Dim xmlLandlordsReferenceNode As IXMLDOMNode
    
    Dim objDO As IReferencesDO
    
    Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    
    Set xmlLandlordsReferenceList = vxmlRequestParentNode.selectNodes("LANDLORDSREF")
    
    For Each xmlLandlordsReferenceNode In xmlLandlordsReferenceList
        objDO.UpdateLandlordsReference xmlLandlordsReferenceNode, vxmlResponseNode
    Next
    
    gobjContext.SetComplete
    
UpdateLandlordsReferenceExit:
    
    Set objDO = Nothing
    Set xmlLandlordsReferenceList = Nothing
    Set xmlLandlordsReferenceNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdateLandlordsReferenceVbErr:

    gobjContext.SetAbort

    GoTo UpdateLandlordsReferenceExit
End Sub

Private Sub IReferencesTxBO_UpdateCurrLandlordsRef( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo UpdateCurrLandlordsRefVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_UpdateCurrLandlordsRef"
    
    Dim xmlLandlordsRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlLandlordsRefList = vxmlRequestParentNode.selectNodes("CURRENTLANDLORDSREF")
        
    If xmlLandlordsRefList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlLandlordsRefList
            objDO.UpdateCurrLandlordsRef xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
UpdateCurrLandlordsRefExit:
    
    Set objDO = Nothing
    Set xmlLandlordsRefList = Nothing
    Set xmlRefNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdateCurrLandlordsRefVbErr:

    gobjContext.SetAbort
    
    GoTo UpdateCurrLandlordsRefExit

End Sub

Private Sub IReferencesTxBO_UpdatePrevLandlordsRef( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo UpdatePrevLandlordsRefVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_UpdatePrevLandlordsRef"
    
    Dim xmlLandlordsRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlLandlordsRefList = vxmlRequestParentNode.selectNodes("PREVIOUSLANDLORDSREF")
        
    If xmlLandlordsRefList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlLandlordsRefList
            objDO.UpdatePrevLandlordsRef xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
UpdatePrevLandlordsRefExit:
    
    Set objDO = Nothing
    Set xmlLandlordsRefList = Nothing
    Set xmlRefNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdatePrevLandlordsRefVbErr:

    gobjContext.SetAbort
    
    GoTo UpdatePrevLandlordsRefExit

End Sub

Private Sub IReferencesTxBO_UpdateCurrEmployersRef( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo UpdateCurrEmployersRefVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_UpdateCurrEmployersRef"
    
    Dim xmlEmployersRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlEmployersRefList = vxmlRequestParentNode.selectNodes("CURRENTEMPLOYERSREF")
        
    If xmlEmployersRefList.length > 0 Then
        ' call UpdateCurrEmployersRef if CURRENTEMPLOYERSREF node in request
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlEmployersRefList
            objDO.UpdateCurrEmployersRef xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
UpdateCurrEmployersRefExit:
    
    Set objDO = Nothing
    Set xmlEmployersRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdateCurrEmployersRefVbErr:

    gobjContext.SetAbort
    
    GoTo UpdateCurrEmployersRefExit

End Sub

Private Sub IReferencesTxBO_UpdatePrevEmployersRef( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo UpdatePrevEmployersRefVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_UpdatePrevEmployersRef"
    
    Dim xmlEmployersRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlEmployersRefList = vxmlRequestParentNode.selectNodes("PREVIOUSEMPLOYERSREF")
        
    If xmlEmployersRefList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlEmployersRefList
            objDO.UpdatePrevEmployersRef xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
UpdatePrevEmployersRefExit:
    
    Set objDO = Nothing
    Set xmlEmployersRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdatePrevEmployersRefVbErr:

    gobjContext.SetAbort
    
    GoTo UpdatePrevEmployersRefExit

End Sub

Private Sub IReferencesTxBO_UpdateLendersReference( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo UpdateLendersReferenceVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_UpdateLendersReference"
    
    Dim xmlLendersReferenceList As IXMLDOMNodeList
    Dim xmlLendersReferenceNode As IXMLDOMNode
    
    Dim objDO As IReferencesDO
    
    Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
    
    Set xmlLendersReferenceList = vxmlRequestParentNode.selectNodes("LENDERSREFERENCE")
    
    For Each xmlLendersReferenceNode In xmlLendersReferenceList
        objDO.UpdateLendersReference xmlLendersReferenceNode, vxmlResponseNode
    Next
    
    gobjContext.SetComplete
    
UpdateLendersReferenceExit:
    
    Set objDO = Nothing
    Set xmlLendersReferenceList = Nothing
    Set xmlLendersReferenceNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdateLendersReferenceVbErr:

    gobjContext.SetAbort

    GoTo UpdateLendersReferenceExit
End Sub

Private Sub IReferencesTxBO_UpdateCurrLendersRef( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo UpdateCurrLendersRefVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_UpdateCurrLendersRef"
    
    Dim xmlLendersRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlLendersRefList = vxmlRequestParentNode.selectNodes("CURRENTLENDERSREF")
        
    If xmlLendersRefList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlLendersRefList
            objDO.UpdateCurrLendersRef xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
UpdateCurrLendersRefExit:
    
    Set objDO = Nothing
    Set xmlLendersRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdateCurrLendersRefVbErr:

    gobjContext.SetAbort
    
    GoTo UpdateCurrLendersRefExit

End Sub

Private Sub IReferencesTxBO_UpdatePrevLendersRef( _
    ByVal vxmlRequestParentNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)

    On Error GoTo UpdatePrevLendersRefVbErr
    
    Const strFunctionName As String = "IReferencesTxBO_UpdatePrevLendersRef"
    
    Dim xmlLendersRefList As IXMLDOMNodeList
    Dim objDO As IReferencesDO
    Dim xmlRefNode As IXMLDOMNode
    
    Set xmlLendersRefList = vxmlRequestParentNode.selectNodes("PREVIOUSLENDERSREF")
        
    If xmlLendersRefList.length > 0 Then
        Set objDO = gobjContext.CreateInstance("omAppProc.ReferencesDO")
        For Each xmlRefNode In xmlLendersRefList
            objDO.UpdatePrevLendersRef xmlRefNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
UpdatePrevLendersRefExit:
    
    Set objDO = Nothing
    Set xmlLendersRefList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdatePrevLendersRefVbErr:

    gobjContext.SetAbort
    
    GoTo UpdatePrevLendersRefExit

End Sub


