VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "AppSummaryTxBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AppSummaryTxBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Application Summary Transactioned Business Object.
'

'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    26/02/01    SYS1879 Created
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl
Implements IAppSummaryTxBO

Private m_objContext As ObjectContext


Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub

Private Sub IAppSummaryTxBO_CreateApplicationReviewHistory(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Creates specified APPLICATIONREVIEWHISTORY entities.
' pass:
'   vxmlRequestParentNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------
    On Error GoTo CreateApplicationReviewHistoryVbErr
    
    Const cstrFunctionName As String = "IAppSummaryTxBO_CreateApplicationReviewHistory"
    
    Dim xmlAppreviewHistoryList As IXMLDOMNodeList
    Dim objAppSummaryDO As IAppSummaryDO
    
    Set xmlAppreviewHistoryList = vxmlRequestParentNode.selectNodes("APPLICATIONREVIEWHISTORY")
        
    If xmlAppreviewHistoryList.length > 0 Then
        Set objAppSummaryDO = m_objContext.CreateInstance(App.Title & ".AppSummaryDO")
        objAppSummaryDO.CreateApplicationReviewHistory xmlAppreviewHistoryList, vxmlResponseNode
    End If
    
    m_objContext.SetComplete
    
CreateApplicationReviewHistoryExit:
    
    Set objAppSummaryDO = Nothing
    Set xmlAppreviewHistoryList = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
    
CreateApplicationReviewHistoryVbErr:

    m_objContext.SetAbort
    
    GoTo CreateApplicationReviewHistoryExit

End Sub

Private Sub IAppSummaryTxBO_UpdateApplicationReviewHistory(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Updates specified APPLICATIONREVIEWHISTORY entities.
' pass:
'   vxmlRequestParentNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       none
' exceptions:
'       none
'------------------------------------------------------------------------------------------
    On Error GoTo UpdateApplicationReviewHistoryVbErr
    
    Const cstrFunctionName As String = "IAppSummaryTxBO_UpdateApplicationReviewHistory"
    
    Dim xmlAppreviewHistoryList As IXMLDOMNodeList
    Dim objAppSummaryDO As IAppSummaryDO
    
    Set xmlAppreviewHistoryList = vxmlRequestParentNode.selectNodes("APPLICATIONREVIEWHISTORY")
        
    If xmlAppreviewHistoryList.length > 0 Then
        Set objAppSummaryDO = m_objContext.CreateInstance(App.Title & ".AppSummaryDO")
        objAppSummaryDO.UpdateApplicationReviewHistory xmlAppreviewHistoryList, vxmlResponseNode
    End If
    
    m_objContext.SetComplete
    
UpdateApplicationReviewHistoryExit:
    
    Set objAppSummaryDO = Nothing
    Set xmlAppreviewHistoryList = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
    
    Exit Sub
    
UpdateApplicationReviewHistoryVbErr:

    m_objContext.SetAbort
    
    GoTo UpdateApplicationReviewHistoryExit

End Sub

