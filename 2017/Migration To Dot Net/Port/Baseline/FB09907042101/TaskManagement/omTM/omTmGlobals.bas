Attribute VB_Name = "omTmGlobals"
'Workfile:      omTMGlobals.bas
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Task Manager Global functions and declarations.

'Dependencies:  Add any other dependent components
'
'-------------------------------------------------------------------------------------------------------
'MARS Specific History:
'
'Prog   Date        Description
'PSC    24/08/2005  MAR32 Add adoLoadSchema add TaskIds for Task Automation Service
'GHun   23/01/2006  MAR1084 Added oeTmAddressTargetingNotSupported
'GHun   13/03/2006  MAR1300 moved gstrHUNTERINTERFACE_COMPONENT definition to StdData
'------------------------------------------------------------------------------------------

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
    omiga4TASKSTATUSTASInProgress = 80  ' PSC 24/08/2005 MAR32
    omiga4TASKSTATUSTASFailed = 90      ' PSC 24/08/2005 MAR32
    omiga4TASKSTATUSTASRetry = 100      ' PSC 24/08/2005 MAR32
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
    oeTmAddressTargetingNotSupported = 8540 'MAR1084 GHun
End Enum

Public Sub Main()
    ' adoAssist
    adoLoadSchema   'PSC 24/08/2005 MAR32
    adoBuildDbConnectionString
End Sub

