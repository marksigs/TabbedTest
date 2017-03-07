Attribute VB_Name = "StdData"
Option Explicit
'Workfile:      StdData.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Standard data
'Dependencies:
'Issues:        This module should be common to all Omiga4 projects but not
'               shared between them, i.e. each component should have its own
'               version.
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'MCS    08/11/99 Created
'PSC    05/01/01 Added Administration System Interface, omAdmin
'PSC    07/03/01 SYS1879 Add Application Processing, omAppProc
'PSC    17/04/01 SYS2188 Add Omiga Task Management, omTM and
'                         Batch Scheduler, omBatch
'MDC    05/06/01 Add Admin Rules (omAdminRules) constant
'SR     06/09/01 New constants for 'ReportOnTitle' and 'RateChange'
'LD     03/10/01 New constant gstrFVS_COMPONENT
'DM     23/11/01 New constant gstrMESSAGE_QUEUE_COMPONENT
'DR     27/02/02 New constant gstrPRINTMANAGER_COMPONENT
'------------------------------------------------------------------------------------------
'MARS History
'
'Prog   Date        Description
'JD     14/10/2005  added constant gstrHOMETRACK_COMPONENT
'GHun   13/03/2006  MAR1300 Merged with core version
'PSC    03/04/2006  MAR1573 Added gstrCRUD_COMPONENT
'------------------------------------------------------------------------------------------
' old refs, do not use...here for backwards compatability.
Public Const gstrCUSTOMEREMPLOYMENT_COMPONENT As String = "omCE"
Public Const gstrCUSTOMER_FINANCIAL_COMPONENT As String = "omCF"
Public Const gstrCOMPONENT_NAME As String = "omCR"
Public Const gstrSUBMISSION_COMPONENT As String = "omSub"
Public Const gstrCREDITCHECK_COMPONENT As String = "omCC"
Public Const gstrRISKASSESSMENT_COMPONENT As String = "omRA"
Public Const gstrDOWNLOAD_COMPONENT As String = "om4To3"
'new refs, programmers use these...
Public Const gstrOMIGA3_MANAGER_COMPONENT As String = "Om3Manager"
Public Const gstrOMIGA4ToOMIGA3DOWNLOAD = "om4to3"
Public Const gstrAPPLICATION_COMPONENT As String = "omApp"
Public Const cstrTABLE_NAME As String = "ADDRESS"
Public Const gstrAPPLICATIONQUOTE As String = "omAQ"
Public Const gstrAUDIT_COMPONENT As String = "omAU"
Public Const gstrBASE_COMPONENT As String = "omBase"
Public Const gstrBUILDINGSANDCONTENTS As String = "omBC"
Public Const gstrCREDIT_CHECK As String = "omCC"
Public Const gstrCOST_MODEL_COMPONENT As String = "omCM"
Public Const gstrCOMPLETENESS_RULES_COMP As String = "omCompRules"
' PSC 03/04/2006 MAR1573
Public Const gstrCRUD_COMPONENT As String = "omCRUD"
Public Const gstrCUSTOMER_EMPLOYMENT As String = "omCE"
Public Const gstrCUSTOMER_FINANCIAL As String = "omCF"
Public Const gstrCUSTOMER_COMPONENT As String = "omCust"
Public Const gstrDPS_COMPONENT As String = "omDPS"
Public Const gstrEXPERIAN_COMPONENT As String = "omExp"
Public Const gstrIMPORT_COMPONENT As String = "OmImp"
Public Const gstrLIFECOVER_COMPONENT As String = "omLC"
Public Const gstrMORTGAGEPRODUCT As String = "omMP"
Public Const gstrORGANISATION_COMPONENT As String = "omOrg"
Public Const gstrPP_COMPONENT As String = "omPP"
Public Const gstrQUICK_QUOTE As String = "omQQ"
Public Const gstrRISK_ASSESSMENT As String = "omRA"
Public Const gstrRISK_ASSESSMENT_RULES_COMP As String = "omRARules"
Public Const gstrSubmission As String = "omSub"
Public Const gstrREQUEST_BROKER_COMPONENT = "omRB"
Public Const gstrTHIRDPARTY_COMPONENT As String = "omTP"
Public Const gstrADMIN_INTERFACE As String = "omAdmin"
Public Const gstrPAYMENTPROCESSING As String = "omPayProc"
Public Const gstrAPPLICATIONPROCESSING As String = "omAppProc"
Public Const gstrMsgTm_COMPONENT As String = "MsgTm"
Public Const gstrPRINT_COMPONENT As String = "omPrint"
Public Const gstrTASKMANAGEMENT_COMPONENT As String = "omTM"
Public Const gstrBATCH_SCHEDULER_COMPONENT As String = "omBatch"
Public Const gstrADMINRULES As String = "omAdminRules"
Public Const gstrREPORTONTITLE_COMPONENT = "omROT"
Public Const gstrRATECHANGE As String = "omRC"
Public Const gstrMESSAGE_QUEUE_COMPONENT As String = "omMessageQueue"
Public Const gstrPRINTMANAGER_COMPONENT As String = "omPM"
Public Const gstrHUNTERINTERFACE_COMPONENT As String = "omHI"
'JD MAR41 new object hometrack
Public Const gstrHOMETRACK_COMPONENT As String = "omHomeTrack"
Public Const gstrAIP_COMPONENT As String = "omAIP"
Public Const gstrCOMP_CHECK_COMPONENT As String = "omComp"
Public Const gstrDECISION_MANAGER_COMPONENT As String = "omDM"
Public Const gstrHOME_USE_COMPONENT As String = "omHU"
Public Const gstrINTERMEDIARY As String = "omIM"
Public Const gstrPAF As String = "omPAF"
Public Const gstrFVS_COMPONENT = "omFVS"
Public Const gstrMSGSPM As String = "MTxSpm.SharedPropertyGroupManager.1"
Public Const gstrCRM_COMPONENT  As String = "omCRMLock"
#If TIMINGS Then
Public Const TIMEFORMAT = "000.000"
#End If

Public Function test()
#If GENERIC_SQL Then
    Dim intA As Integer
#Else
    Dim intB As Integer
#End If
End Function

'GHun Extra test cases
Public Function test2()
    #If GENERIC_SQL Then
        Debug.Print "This should remain 1"
    #Else
        Debug.Print "This should be removed 1"
    #End If
    #If GENERIC_SQL Then
        Debug.Print "This should remain 2"
    #End If
    
    #If Not GENERIC_SQL Then
        Debug.Print "This should be removed 3"
    #End If
    #If Not GENERIC_SQL Then
        Debug.Print "This should be removed 4"
    #Else
        Debug.Print "This should remain 4"
    #End If
    
    #If Win32 Then
        Debug.Print "This should remain 5"
    #Else
        Debug.Print "This should be removed 5"
    #End If
    
    #If USING_VSA Then
        Debug.Print "This should be removed 6"
    #End If
    
    #If MIN_BUILD Then
        Debug.Print "This should be removed 7"
    #Else
        Debug.Print "This should remain 7"
    #End If
    #If Not MIN_BUILD Then
        Debug.Print "This should remain 8"
    #Else
        Debug.Print "This should be removed 8"
    #End If
    
End Function

Public Sub Main()
End Sub
'GHun End

