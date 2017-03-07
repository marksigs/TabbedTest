Attribute VB_Name = "MsgTmGlobals"
Option Explicit
Public Enum TMERROR
    ' (msgTm) Task Management specific errors
    oeTmNoNextStage = 4800
    oeTmStageNotApplicable = 4801
    oeTmMandatoryTasksOutstanding = 4802
    oeTmNotExceptionStage = 4803
    oeTmNoStageDetail = 4804
    oeTmNoTaskDetail = 4805
    oeTmNoStageAuthority = 4806
    oeTmNoTaskAuthority = 4807
    oeTmNoStageTaskAuthority = 4825 'CORE230 GHun
End Enum
Public Sub Main()
    ' adoAssist
    adoLoadSchema
    adoBuildDbConnectionString
End Sub
