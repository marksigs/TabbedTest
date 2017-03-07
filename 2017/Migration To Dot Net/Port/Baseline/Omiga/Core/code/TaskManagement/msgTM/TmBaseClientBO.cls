VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "TmBaseClientBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
Public Function TmBaseRequest(ByVal vstrXmlIn As String) As String
    Dim objTm As MsgTm.IMsgTmBaseBO
    Set objTm = gobjContext.CreateInstance("MsgTm.MsgTmBO")
    TmBaseRequest = objTm.TmBaseRequest(vstrXmlIn)
    Set objTm = Nothing
End Function
