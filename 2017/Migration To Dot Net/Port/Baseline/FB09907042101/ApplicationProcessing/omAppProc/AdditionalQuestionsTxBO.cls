VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "AdditionalQuestionsTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IAdditionalQuestionsTxBO
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

Private Sub IAdditionalQuestionsTxBO_CreateAppAdditionalQuestions(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo CreateAppAdditionalQuestionsVbErr
    
    Const strFunctionName As String = "IAdditionalQuestionsTxBO_CreateAppAdditionalQuestions"
    
    Dim xmlAppQuestionsList As IXMLDOMNodeList
    Dim xmlAppQuestionsNode As IXMLDOMNode
    Dim objDO As IAdditionalQuestionsDO
    
    Set xmlAppQuestionsList = vxmlRequestParentNode.selectNodes("APPLNADDITIONALQUESTIONS")
        
    If xmlAppQuestionsList.length > 0 Then
        ' call CreateAppAdditionalQuestions if APPLNADDITIONALQUESTIONS node in request
        Set objDO = gobjContext.CreateInstance("omAppProc.AdditionalQuestionsDO")
        For Each xmlAppQuestionsNode In xmlAppQuestionsList
            objDO.CreateAppAdditionalQuestions xmlAppQuestionsNode, vxmlResponseNode
        Next
    End If
    
    gobjContext.SetComplete
    
CreateAppAdditionalQuestionsExit:
    
    Set objDO = Nothing
    Set xmlAppQuestionsList = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateAppAdditionalQuestionsVbErr:

    gobjContext.SetAbort
    
    GoTo CreateAppAdditionalQuestionsExit

End Sub

Private Sub IAdditionalQuestionsTxBO_UpdateAppAdditionalQuestions(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    On Error GoTo UpdateAppAdditionalQuestionsVbErr
    
    Const strFunctionName As String = "IAdditionalQuestionsTxBO_UpdateAppAdditionalQuestions"
    
    Dim xmlAppAdditionalQuestionsList As IXMLDOMNodeList
    Dim xmlAppAdditionalQuestionsNode As IXMLDOMNode
    
    Dim objDO As IAdditionalQuestionsDO
    
    Set objDO = gobjContext.CreateInstance("omAppProc.AdditionalQuestionsDO")
    
    Set xmlAppAdditionalQuestionsList = vxmlRequestParentNode.selectNodes("APPLNADDITIONALQUESTIONS")
    
    For Each xmlAppAdditionalQuestionsNode In xmlAppAdditionalQuestionsList
        objDO.UpdateAppAdditionalQuestions xmlAppAdditionalQuestionsNode, vxmlResponseNode
    Next
    
    gobjContext.SetComplete
    
UpdateAppAdditionalQuestionsExit:
    
    Set objDO = Nothing
    Set xmlAppAdditionalQuestionsList = Nothing
    Set xmlAppAdditionalQuestionsNode = Nothing
    
    errCheckError strFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdateAppAdditionalQuestionsVbErr:

    gobjContext.SetAbort

    GoTo UpdateAppAdditionalQuestionsExit

End Sub

