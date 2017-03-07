Attribute VB_Name = "omTmGlobals"
Option Explicit
 
 '=============================================
 'Enumeration Declaration Section
 '=============================================
Public Enum TASKSTATUS
    omiga4TASKSTATUSUndefined = 0
    omiga4TASKSTATUSIncomplete = 10
    omiga4TASKSTATUSPending = 20
    omiga4TASKSTATUSNotApplicable = 30
    omiga4TASKSTATUSComplete = 40
    omiga4TASKSTATUSCarriedForward = 50
    omiga4TASKSTATUSChasedUp = 60
    omiga4TASKSTATUSCancelled = 70
End Enum
' keep in step with MsgTmGlobals
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
    oeTmHunterInterfaceFailed = 7001    'BMIDS00025 MDC 10/06/2002
    oeTmBureauDataImportFailed = 7004    'BMIDS00336 MDC 22/08/2002
    oeTmCCOKBureauImportFailed = 7005    'BMIDS00336 MDC 22/08/2002
    oeTmAutomaticTaskDidntComplete = 7011   'BMIDS01076 MO 25/11/2002
    oeTmTransactFullDecisionFailed = 10004 ' TK 28/07/2004 E2EM00000113
End Enum
'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        Description
'MDC    10/06/2002  BMIDS00025 - Hunter Interface
'MDC    22/08/2002  BMIDS00336 - CCWP1 BM062 Credit Check and Bureau Download
'MO     25/11/2002  BMIDS01076 - Auto task failed error
'------------------------------------------------------------------------------------------
Public Sub Main()
    ' adoAssist
    adoBuildDbConnectionString
End Sub
