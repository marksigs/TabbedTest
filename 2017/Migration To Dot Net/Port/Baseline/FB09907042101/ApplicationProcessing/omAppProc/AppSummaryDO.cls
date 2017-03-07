VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AppSummaryDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AppSummaryDO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Application Summary Data Object.
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
Implements IAppSummaryDO

Private m_objContext As ObjectContext

Private Sub IAppSummaryDO_CreateApplicationReviewHistory(ByVal vxmlRequestNodeList As IXMLDOMNodeList, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Creates APPLICATIONREVIEWHISTORY entities
' pass:
'   vxmlRequestNodeList
'       List of nodes that need to be created
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       n/a
' exceptions:
'       n/a
'------------------------------------------------------------------------------------------
    On Error GoTo CreateApplicationReviewHistoryExit
    
    Const cstrFunctionName As String = "IAppSummaryDO_CreateApplicationReviewHistory"
    
    adoCreateFromNodeList vxmlRequestNodeList, "APPLICATIONREVIEWHISTORY"
    
CreateApplicationReviewHistoryExit:
    
    m_objContext.SetComplete
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub IAppSummaryDO_FindApplicationReviewHistoryList(ByVal vxmlRequestParentNode As IXMLDOMNode, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve a list of APPLICATIONREVIEWHISTORY entities
' pass:
'   vxmlRequestParentNode
'       XML REQUEST node
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       APPLICATIONREVIEWHISTORY nodes appended to vxmlResponseNode on exit
' exceptions:
'       n/a
'------------------------------------------------------------------------------------------
    On Error GoTo FindApplicationReviewHistoryListExit
    
    Const cstrFunctionName As String = "IAppSummaryDO_FindApplicationReviewHistoryList"
    
    adoGetAsXML vxmlRequestParentNode, vxmlResponseNode, "APPLICATIONREVIEWHISTORY"
    
    If vxmlResponseNode.hasChildNodes = False Then
        errThrowError cstrFunctionName, oeRecordNotFound
    End If
    
FindApplicationReviewHistoryListExit:
    
    m_objContext.SetComplete
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub


Private Sub IAppSummaryDO_UpdateApplicationReviewHistory(ByVal vxmlRequestNodeList As IXMLDOMNodeList, ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Updates APPLICATIONREVIEWHISTORY entities
' pass:
'   vxmlRequestNodeList
'       List of nodes that need to be updated
'   vxmlResponseNode
'       XML RESPONSE node
' return:
'       n/a
' exceptions:
'       n/a
'------------------------------------------------------------------------------------------
    On Error GoTo UpdateApplicationReviewHistoryExit
    
    Const cstrFunctionName As String = "IAppSummaryDO_UpdateApplicationReviewHistory"
    
    adoUpdateFromNodeList vxmlRequestNodeList, "APPLICATIONREVIEWHISTORY"
    
UpdateApplicationReviewHistoryExit:
    
    m_objContext.SetComplete
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub

