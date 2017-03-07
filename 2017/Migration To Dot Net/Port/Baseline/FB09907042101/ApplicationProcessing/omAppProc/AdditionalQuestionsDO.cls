VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AdditionalQuestionsDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AdditionalQuestionsDO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Additional Questions Data Object.
'

'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'JR     26/03/01    SYS2048 Created
'                   Added FindAdditionalQuestionList
'                   Added FindAppAdditionalQuestionList
'                   Added CreateAppAdditionalQuestions
'                   Added UpdateAppAdditionalQuestions
'------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl
Implements IAdditionalQuestionsDO

Private gobjContext As ObjectContext



'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        Description
'MV     14/06/2002  BMIDS00013 - Created IAdditionalQuestionsDO_GetAdditionalQuestionsList
'------------------------------------------------------------------------------------------

Private Sub IAdditionalQuestionsDO_CreateAppAdditionalQuestions(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo CreateAppAdditionalQuestionsExit
    
    Const strFunctionName As String = "IAdditionalQuestionsDO_CreateAppAdditionalQuestions"
    
    adoCreateFromNode vxmlRequestNode, "APPLNADDITIONALQUESTIONS"
    
CreateAppAdditionalQuestionsExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub IAdditionalQuestionsDO_FindAdditionalQuestionList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo FindAdditionalQuestionListExit
    
    Const cstrFunctionName As String = "IAdditionalQuestionsDO_FindAdditionalQuestionList"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "ADDITIONALQUESTIONS"
    
FindAdditionalQuestionListExit:
    
    gobjContext.SetComplete
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IAdditionalQuestionsDO_FindAppAdditionalQuestionList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo FindAppAdditionalQuestionListExit
    
    Const cstrFunctionName As String = "IAdditionalQuestionsDO_FindAppAdditionalQuestionList"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "APPLNADDITIONALQUESTIONS"
    
FindAppAdditionalQuestionListExit:
    
    gobjContext.SetComplete
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IAdditionalQuestionsDO_GetAdditionalQuestionsList(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
    
    On Error GoTo GetAdditionalQuestionsListExit
    
    Const strFunctionName As String = "IAdditionalQuestionsDO_GetAdditionalQuestionsList"
    
    adoGetAsXML vxmlRequestNode, vxmlResponseNode, "GETADDITIONALQUESTIONSLIST"
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError strFunctionName, oeRecordNotFound
    End If
    
GetAdditionalQuestionsListExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)
    
End Sub


Private Sub IAdditionalQuestionsDO_UpdateAppAdditionalQuestions(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
On Error GoTo UpdateAppAdditionalQuestionsExit
    
    Const strFunctionName As String = "IAdditionalQuestionsDO_UpdateAppAdditionalQuestions"
    
    adoUpdateFromNode vxmlRequestNode, "APPLNADDITIONALQUESTIONS"
    
UpdateAppAdditionalQuestionsExit:
    
    gobjContext.SetComplete
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

Private Sub ObjectControl_Activate()
 Set gobjContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
  Set gobjContext = Nothing
End Sub

