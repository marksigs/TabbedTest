VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SupervisorObjectCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   SupervisorObjectCopy
' Description   :   Class that handles copying of Supervisor items, such as coutry, dist
'                   channel, lenders etc. to between different databases.
' Change history
' Prog      Date        Description
' DJP       31/01/01    Added copying of BankHoliday to CopyCountry.
' DJP       22/06/01    SQL Server port.
' DJP       07/09/01    Check tables exist before attempting a promote.
' DJP       21/10/01    SYS2829 - Handle GUIDS correctly during promote.
' STB       02/01/02    SYS2551 - Telephone numbers are now promoted as part of
'                       Contacts.
' STB       07/01/02    SYS3581 - ThirdParties now use the CopyContact() routine.
' STB       29/01/02    SYS2957 - Security access is not promoted.
' STB       04/02/02    SYS3857 - Base Rates can now be promoted.
' DJP       24/02/02    SYS4121 - Client promotion.
' DJP       26/02/02    SYS4188 - Base Rate promotion.
' STB       13/05/02    SYS4417 - Added AllowableIncomeFactors.
' CL        28/05/02    SYS4766 Merge MSMS & CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS History
' Prog      Date        Description
' GHun      16/08/2005  MAR45 Apply BBG1370 and BBG1054 (Promotion of printing templates)
' GHun      14/10/2005  MAR202 Promote Packs
' HMA       16/01/2006  MAR524 Promote TaskInterfaceMessage and TaskInterfaceSubMessage with Tasks.
' HMA       31/01/2006  MAR1120 Changes for stage promotion.
' HMA       31/01/2006  MAR967  Changes for pack promotion.
' PSC       07/03/2006  MAR1298 Change CopyObject to copy records correctly
' HMA       21/03/2006  MAR1397 Change CompareKeys for guids. Correct CopyBusinessGroup.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' EPSOM History
' Prog      Date        Description
' TW        09/10/2006  EP2_7 Add promotion of Additional Borrowing Fee and Credit Limit Increase Fee
' TW        17/10/2006  EP2_15 Add promotion for MortgageClubNetworkAssociation
' TW        18/11/2006  EP2_132 ECR20/21 Further promotions for FSA related changes
' TW        23/11/2006  EP2_172 Change control EP2_5 - E2CR16 changes related to Introducer/Product Exclusives
' TW        11/12/2006  EP2_20 Add promotion of Transfer of Equity Fee
' TW        14/12/2006  EP2_518 Added handling for Default Procuration Fees, Loan Amount Adjustments, LTV Amount Adjustments
' TW        20/12/2006  EP2_604 Potential error in Promotion of Users
' TW        27/03/2007  EP2_1101 - Promotion inconsistencies
' TW        04/04/2007  EP2_2242 - Email_COT_Received’  is not being triggered - Promotion problem
' TW        05/04/2007  EP2_2292 - Create New Packager Firm
' TW        14/05/2007  VR216 - Display up to 3 phone numbers for Introducers in Supervisor
'------------------------------------------------------------------------------
Option Explicit

'Private data
Private m_sTargetDatabase As String
Private m_sSourceDatabase As String
Private m_clsProgress As MSGProgress

' Enums
Private Enum BandedTableType
    typeBandedSetTable = 1
    typeBandedTable = 2
End Enum

Private Enum PromoteStageType
    enumStageAll = 1
    EnumSingleStage
    EnumNone
End Enum

Private Sub CopyObjectNonPrimaryKey(clsSource As TableAccess, clsTarget As TableAccess, rs As ADODB.Recordset)
' TW 27/03/2007 EP2_1101 - New Routine
Dim colValues As Collection
    
    On Error GoTo Failed:
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            Do While Not rs.EOF
                Set colValues = New Collection
                colValues.Add rs(0)
                clsSource.SetKeyMatchValues colValues
                clsTarget.SetKeyMatchValues colValues
                CopyRecords clsSource, clsTarget, m_sTargetDatabase, True
                rs.MoveNext
            Loop
        End If
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog      Date        Description
' GD        27/05/02    AQR : BMIDS00016 ; Supervisor Promotions
' AW        29/07/02    BMIDS00237          Do not delete target when promoting an edit
' DB        03/01/03    BM0141              Ensure tasklist is promoted along with the businessgroup
' DJP       06/03/03    BM0423              Add call to ValidateProduct on Product promotion.
' MV        04/04/2003  BM0402              Amended CopyComboGroup();Created CreateComboValidationAuditRecords()
'                                           CreateComboValueAuditRecords()
' BS        30/04/2003  BM0498  Check for EOF in CreateComboValidationAuditRecords and CreateComboValueAuditRecords
' [MC]      25/04/2004  BMIDS763    Data Promotion code added for three new sets
'                                   (Product switch fee, Insurance Fee set, TT FeeSet)
' JD        12/07/2004  BMIDS775    Add promote of mortgage product additionalFeatures
' JD        12/07/2004  BMIDS765    add promote of rental income rate sets.
' HMA       09/12/2004  BMIDS959    Remove MortgageProductBands table
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function CreateComboValueAuditRecords(clsSrcTableAccess As TableAccess, strTargetDatabase As String) As Boolean
    
    On Error GoTo Failed

    Dim sTable As String
    Dim sTargetDatabase  As String
    
    Dim colMatchValues  As New Collection
    Dim clsTable As TableAccess
    Dim clsComboValueAuditTable As ComboValueAuditTable
    
    Dim rsSrc As ADODB.Recordset
    Set rsSrc = New ADODB.Recordset
    
    sTargetDatabase = GetTargetDatabase()
    Set colMatchValues = clsSrcTableAccess.GetKeyMatchValues()
    
    sTable = clsSrcTableAccess.GetTable()
    
    If Not g_clsDataAccess.DoesTableExist(sTable) Then
        g_clsErrorHandling.RaiseError errGeneralError, "Source Database: " & sTable & " table does not exist"
    End If
    
    If Not g_clsDataAccess.DoesTableExist(sTable, sTargetDatabase) Then
        g_clsErrorHandling.RaiseError errGeneralError, "Target Database: " & sTable & " table does not exist"
    End If
    
    
    ' Source
    Set rsSrc = clsSrcTableAccess.GetTableData()
    
    'BS BM0498 30/04/03
    'Check for EOF
    If Not rsSrc.EOF Then
    'BS BM0498 End 30/04/03
    
        rsSrc.MoveFirst
        
        
        Set clsComboValueAuditTable = New ComboValueAuditTable
        Set clsTable = clsComboValueAuditTable
        
        clsTable.SetDatabase GetTargetDatabase()
        
        Do While Not rsSrc.EOF
        
            g_clsFormProcessing.CreateNewRecord clsComboValueAuditTable
            
            clsComboValueAuditTable.SetGroupName rsSrc.fields("GROUPNAME").Value
            clsComboValueAuditTable.SetValueID rsSrc.fields("VALUEID").Value
            clsComboValueAuditTable.SetAuditDate Now()
            clsComboValueAuditTable.SetValueName rsSrc.fields("VALUENAME").Value
            clsComboValueAuditTable.SetChangeUser g_sSupervisorUser
            clsComboValueAuditTable.SetOperation "C"
            
            TableAccess(clsComboValueAuditTable).Update
            
            rsSrc.MoveNext
        
        Loop
    
    'BS BM0498 30/04/03
    End If
    'BS BM0498 End 30/04/03
    
    CreateComboValueAuditRecords = True
    
    Exit Function

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function



Private Function CreateComboValidationAuditRecords(clsSrcTableAccess As TableAccess, strTargetDatabase As String) As Boolean
    
    On Error GoTo Failed

    Dim sTable As String
    Dim sTargetDatabase  As String
    
    Dim colMatchValues  As New Collection
    Dim clsTable As TableAccess
    Dim clsComboValidationAuditTable As ComboValidationAuditTable
    
    Dim rsSrc As ADODB.Recordset
    Set rsSrc = New ADODB.Recordset
    
    sTargetDatabase = GetTargetDatabase()
    Set colMatchValues = clsSrcTableAccess.GetKeyMatchValues()
    
    sTable = clsSrcTableAccess.GetTable()
    
    If Not g_clsDataAccess.DoesTableExist(sTable) Then
        g_clsErrorHandling.RaiseError errGeneralError, "Source Database: " & sTable & " table does not exist"
    End If
    
    If Not g_clsDataAccess.DoesTableExist(sTable, sTargetDatabase) Then
        g_clsErrorHandling.RaiseError errGeneralError, "Target Database: " & sTable & " table does not exist"
    End If
    
    
    ' Source
    Set rsSrc = clsSrcTableAccess.GetTableData()
    
    'BS BM0498 30/04/03
    'Check for EOF
    If Not rsSrc.EOF Then
    'BS BM0498 End 30/04/03
    
        rsSrc.MoveFirst
        
        
        Set clsComboValidationAuditTable = New ComboValidationAuditTable
        Set clsTable = clsComboValidationAuditTable
        
        clsTable.SetDatabase GetTargetDatabase()
        
        Do While Not rsSrc.EOF
        
            g_clsFormProcessing.CreateNewRecord clsComboValidationAuditTable
            
            clsComboValidationAuditTable.SetGroupName rsSrc.fields("GROUPNAME").Value
            clsComboValidationAuditTable.SetValueID rsSrc.fields("VALUEID").Value
            clsComboValidationAuditTable.SetAuditDate Now()
            clsComboValidationAuditTable.SetValidationType rsSrc.fields("VALIDATIONTYPE").Value
            clsComboValidationAuditTable.SetChangeUser g_sSupervisorUser
            clsComboValidationAuditTable.SetOperation "C"
            
            TableAccess(clsComboValidationAuditTable).Update
            
            rsSrc.MoveNext
        
        Loop
    
    'BS BM0498 30/04/03
    End If
    'BS BM0498 End 30/04/03
    CreateComboValidationAuditRecords = True
    
    Exit Function

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Function










Private Sub DeleteUnusedEntries(clsSource As TableAccess, clsTarget As TableAccess, vval)
' TW 27/03/2007 EP2_1101 - New routine
' This routine deletes entries from the Target Database table which are not in the Source Database table

Dim blnDelete As Boolean
Dim X As Integer

Dim strDelete As String

Dim rsSource As ADODB.Recordset
Dim rsTarget As ADODB.Recordset
        
    On Error GoTo Failed:
' 1 Get the records on the Source Database
    Set rsSource = clsSource.GetRecordSet()
    If rsSource Is Nothing Then
        Exit Sub
    End If
    If rsSource.RecordCount = 0 Then
        Exit Sub
    End If

' 2 Get the records on the Target Database
    Set rsTarget = clsTarget.GetRecordSet()
    If rsTarget Is Nothing Then
        Exit Sub
    End If
    If rsTarget.RecordCount = 0 Then
        Exit Sub
    End If
    If rsSource.RecordCount = rsTarget.RecordCount Then
        Exit Sub
    End If
    
 ' 3 Delete any records that appear in the target but not in the source
    rsSource.MoveFirst
    rsTarget.MoveFirst
    Do While Not rsSource.EOF And Not rsTarget.EOF
        blnDelete = False
        Do While Not rsTarget.EOF
            strDelete = ""
            For X = 0 To rsTarget.fields.Count - 1
                strDelete = strDelete & rsTarget.fields(X).Name & " = " & rsTarget.fields(X).Value
                If X < rsTarget.fields.Count - 1 Then
                    strDelete = strDelete & ", "
                End If
            Next X
            If rsSource.EOF Then
                blnDelete = True
            Else
                For X = 0 To rsTarget.fields.Count - 1
                    If Format$(rsTarget(X)) <> Format$(rsSource(X)) Then
                        blnDelete = True
                    End If
                Next X
            End If
            If blnDelete Then
                rsTarget.Delete
                WriteLogEvent "Deleted from target database table '" & clsSource.GetTable() & "' - Record with values (" & strDelete & ")"
            Else
                rsSource.MoveNext
            End If
            rsTarget.MoveNext
            If blnDelete Then
                blnDelete = False
                Exit Do
            End If
        Loop
    Loop

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub


Public Sub PromoteMortgageClubNetworkAssociation(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New MortgageClubNetAssocTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteMortgageNatureOfLoanClass(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New MortProdNatureOfLoanTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub PromoteMortgageIncomeStatusClass(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New MortProdIncomeStatusTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteProcFeeDefault(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New ProcFeeDefaultTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub PromoteMortgageProductProductClass(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New MortProdProdClassTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteFeeAdjByLoan(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New MortProdProdClassTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub PromoteProcFeeAdjByLoan(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New ProcFeeAdjByLoanTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub PromoteProcFeeAdjByLTV(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New ProcFeeAdjByLTVTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteAppointments(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New AppointmentsTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub PromoteMortgageProductExclusivity(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New MortProdExclusivityTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteFirmNetAssoc(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New FirmClubNetAssocTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteARFirm(clsSrcARFirm As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcARFirm.SetDatabase g_clsDataAccess.GetConnectionKey()
' TW 27/03/2007 EP2_1101
'    CopyRecords clsSrcARFirm, New ARFirmTable, sTargetDatabase
    CopyARFirm clsSrcARFirm, sTargetDatabase
' TW 27/03/2007 EP2_1101 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteFirmTradingName(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New FirmTradingNameTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteFirmPermissions(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New FirmPermissionsTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub PromoteIntroducerFirm(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New IntroducerFirmTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteIntroducer(clsSrcIntroducer As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcIntroducer.SetDatabase g_clsDataAccess.GetConnectionKey()
' TW 27/03/2007 EP2_1101
'    CopyRecords clsSrcIntroducer, New IntroducerTable, sTargetDatabase
    CopyIntroducer clsSrcIntroducer, sTargetDatabase
' TW 27/03/2007 EP2_1101 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromotePrincipalFirm(clsSrcPrincipalFirm As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcPrincipalFirm.SetDatabase g_clsDataAccess.GetConnectionKey()
' TW 27/03/2007 EP2_1101
'    CopyRecords clsSrcPrincipalFirm, New PrincipalFirmTable, sTargetDatabase
    CopyPrincipalFirm clsSrcPrincipalFirm, sTargetDatabase
' TW 27/03/2007 EP2_1101 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



Public Sub SetProgress(clsProgress As MSGProgress)
    Set m_clsProgress = clsProgress
End Sub

Public Function GetTargetDatabase() As String
    On Error GoTo Failed
    
    If Len(m_sTargetDatabase) > 0 Then
        GetTargetDatabase = m_sTargetDatabase
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "GetTargetDatabase: Key is empty"
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub SetTargetDatabase(sTarget As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTarget
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub OpenTargetDatabase(sTargetDatabase As String)
    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    m_sSourceDatabase = g_clsDataAccess.GetActiveDatabaseName()
    'g_clsDataAccess.OpenDatabase GetTargetDatabase()

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteThirdParty(clsSrcDirectory As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    CopyThirdParty clsSrcDirectory, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteMessages(clsSrcMessage As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcMessage.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcMessage, New ErrorMessageTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteWorkingHours(clsSrcWorkingHours As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcWorkingHours.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyBandedFee clsSrcWorkingHours, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteRegion(clsSrcRegion As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcRegion.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcRegion, New RegionTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteBandedParameter(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New BandedGlobalParametersTable, sTargetDatabase, True
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteGlobalParameter(clsSrcParameter As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New FixedParametersTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteCompetency(clsSrcCompetency As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcCompetency.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcCompetency, New CompetencyTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteUser(clsSrcUser As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase

    clsSrcUser.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyUser clsSrcUser, sTargetDatabase

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteUnit(clsSrcUnit As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase

    clsSrcUnit.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyUnit clsSrcUnit, sTargetDatabase

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteDepartment(clsSrcDepartment As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcDepartment.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyDepartment clsSrcDepartment, sTargetDatabase
            
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub PromoteBandedFee(clsSrcBandedFee As TableAccess, sTargetDatabase As String, Optional bDeleteTarget As Boolean = True)

    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcBandedFee.SetDatabase g_clsDataAccess.GetConnectionKey()
    '   AW  29/07/02    BMIDS00237
    CopyBandedFee clsSrcBandedFee, sTargetDatabase, bDeleteTarget
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub


Public Sub PromoteBaseRateSet(clsSrcBandedFee As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcBandedFee.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyBandedFee clsSrcBandedFee, sTargetDatabase, False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub

Public Sub PromoteDistributionChannel(clsSrcDistChannel As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    
    
    clsSrcDistChannel.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyDistributionChannel clsSrcDistChannel, sTargetDatabase

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteTask(clsSrcTask As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcTask.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyTask clsSrcTask, sTargetDatabase

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteStage(clsSrcStage As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcStage.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyStage clsSrcStage, sTargetDatabase

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteActivity(clsSrcActivity As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcActivity.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyActivity clsSrcActivity, sTargetDatabase

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromotePayProtProducts(clsTableAccess As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    
    
    clsTableAccess.SetDatabase g_clsDataAccess.GetConnectionKey()
    
    CopyRecords clsTableAccess, New PayProtProductsTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteBAndCProduct(clsTableAccess As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    
    
    clsTableAccess.SetDatabase g_clsDataAccess.GetConnectionKey()
    
    CopyRecords clsTableAccess, New BuildingAndContentsTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteLifeCoverRates(clsTableAccess As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    
    
    clsTableAccess.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyLifeCoverRates clsTableAccess, sTargetDatabase

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteCountry(clsTableAccess As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    
    
    clsTableAccess.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyCountry clsTableAccess, sTargetDatabase

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteLender(clsTableAccess As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    
    
    clsTableAccess.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyLender clsTableAccess, sTargetDatabase

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteComboGroup(clsTableAccess As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    
    
    clsTableAccess.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyComboGroup clsTableAccess, sTargetDatabase

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteProduct(clsTableAccess As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    m_sTargetDatabase = sTargetDatabase
    
    
    
    CopyProduct clsTableAccess, sTargetDatabase
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyRecords(clsSrcTable As TableAccess, clsTargetTable As TableAccess, sTargetDatabase As String, Optional bDeleteTarget As Boolean)
    On Error GoTo Failed
    
    Dim colMatchValues As Collection
    
    Set colMatchValues = clsSrcTable.GetKeyMatchValues()
    
    clsTargetTable.SetDatabase sTargetDatabase
    clsTargetTable.SetKeyMatchValues colMatchValues
    
    If bDeleteTarget Then
        clsTargetTable.DeleteRecords
    End If
    
    CopyObject clsSrcTable, clsTargetTable
        
' TW 27/03/2007 EP2_1101
    If bDeleteTarget Then
        DeleteUnusedEntries clsSrcTable, clsTargetTable, colMatchValues(1)
    End If
' TW 27/03/2007 EP2_1101 End
    clsTargetTable.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyThirdParty(clsSrcDirectory As NameAndAddressDirTable, sTargetDatabase As String)
    On Error GoTo Failed
    Dim vDirectoryGUID As Variant
    
    GetDatabaseObject clsSrcDirectory
    vDirectoryGUID = clsSrcDirectory.GetDirectoryGUID()
    CopyDirectory vDirectoryGUID, sTargetDatabase ' for name and address directory

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyLifeCoverRates(clsSrcLifeCover As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    Dim clsTargetLifeCover As LifeCoverRatesTable
    Dim clsTargetTableAccess As TableAccess
    Dim colMatchValues As Collection
    
    Set colMatchValues = clsSrcLifeCover.GetKeyMatchValues()
    
    Set clsTargetLifeCover = New LifeCoverRatesTable
    Set clsTargetTableAccess = clsTargetLifeCover
    clsTargetTableAccess.SetDatabase sTargetDatabase
    clsTargetTableAccess.SetKeyMatchValues colMatchValues
    
    CopyObject clsSrcLifeCover, clsTargetLifeCover
    clsTargetTableAccess.Update
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyDistributionChannel(clsSrcDistChannel As DistributionChannelTable, sTargetDatabase As String)
    On Error GoTo Failed
    ' Things to be copied here are the distribution channel itself, plus the
    ' NonWorkingDay table
    
    ' First, distribution channel
    Dim clsTargetDistChannel As DistributionChannelTable
    Dim clsSrcTableAccess As TableAccess
    Dim clsTargetTableAccess As TableAccess
    Dim colMatchValues As Collection
    Dim sChannelID As String
    Dim rsSrc As ADODB.Recordset
    
    Set clsTargetDistChannel = New DistributionChannelTable
    Set clsTargetTableAccess = clsTargetDistChannel
    
    Set clsSrcTableAccess = clsSrcDistChannel
    Set colMatchValues = clsSrcTableAccess.GetKeyMatchValues
    
    clsTargetTableAccess.SetKeyMatchValues colMatchValues
    clsTargetTableAccess.SetDatabase sTargetDatabase
    
    Set rsSrc = clsSrcTableAccess.GetTableData()
    
    If clsSrcTableAccess.RecordCount() > 0 Then
        sChannelID = clsSrcDistChannel.GetChannelID()
        CopyObject clsSrcDistChannel, clsTargetDistChannel
        
        clsTargetTableAccess.Update
        
        ' Now, non working day
        Dim clsSrcNonWorkingDay As NonWorkingDayTable
        Dim clsTargetNonWorkingDay As NonWorkingDayTable
    
        Set clsSrcNonWorkingDay = New NonWorkingDayTable
        Set clsSrcTableAccess = clsSrcNonWorkingDay
        clsSrcTableAccess.SetKeyMatchValues colMatchValues
        
        Set clsTargetNonWorkingDay = New NonWorkingDayTable
        Set clsTargetTableAccess = clsTargetNonWorkingDay
        clsTargetTableAccess.SetKeyMatchValues colMatchValues
        clsTargetTableAccess.SetDatabase sTargetDatabase
        clsTargetTableAccess.DeleteRecords sTargetDatabase
        
        CopyObject clsSrcNonWorkingDay, clsTargetNonWorkingDay, True
        clsTargetTableAccess.Update
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyIntermediaryProcFee(clsSrcIntermediary As IntermediaryTable, sTargetDatabase As String)
    On Error GoTo Failed
    
    Dim colMatchValues As Collection
    Dim clsSrcProcFeeType As IntermediaryProcFeeTable
    Dim clsSrcProcFeeSplit As IntProcFeeSplitTable
    Dim clsSrcProcFeeSplitByInt As IntProcFeeSplitForIntTable
    
    Set clsSrcProcFeeType = New IntermediaryProcFeeTable
    Set clsSrcProcFeeSplit = New IntProcFeeSplitTable
    Set clsSrcProcFeeSplitByInt = New IntProcFeeSplitForIntTable
    
    Set colMatchValues = TableAccess(clsSrcIntermediary).GetKeyMatchValues
    
    
    'First the proc fee type
    TableAccess(clsSrcProcFeeType).SetKeyMatchValues colMatchValues
    CopyRecords clsSrcProcFeeType, New IntermediaryProcFeeTable, sTargetDatabase
    
    'Next the proc fee split
    TableAccess(clsSrcProcFeeSplit).SetKeyMatchValues colMatchValues
    CopyRecords clsSrcProcFeeSplit, New IntProcFeeSplitTable, sTargetDatabase
    
    'Next the Proc fee split by Intermediary
    TableAccess(clsSrcProcFeeSplitByInt).SetKeyMatchValues colMatchValues
    CopyRecords clsSrcProcFeeSplitByInt, New IntProcFeeSplitForIntTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyIntermediary(clsSrcIntermediary As IntermediaryTable, sTargetDatabase As String)
    On Error GoTo Failed

    Dim clsTargetIntermediary As IntermediaryTable
    Dim clsIntermediary As Intermediary
    Dim clsSrcIntOrg As IntermediaryOrganisation
    Dim clsSrcIntInd As IntermediaryIndividual
    Dim clsSrcIntBankAcc As IntBankAccountTable
    Dim colMatchValues As Collection
    Dim vContactGUID As Variant
    Dim vAddressGUID As Variant
    Dim sIntType As String
    Dim sIndIntTypeID As String
    Dim clsSrcCorrespondence As IntCorrespondenceTable
    Dim clsSrcIntInsurance As IntCrossSellingTable
    Dim clsSrcIntDailyReport As IntermediaryReportDaysTable
    Dim clsSrcIntProcFeeTarget As IntermediaryTargetTable
    Dim bPromoteProcFee As Boolean
    
    Set clsSrcIntProcFeeTarget = New IntermediaryTargetTable
    Set clsSrcIntDailyReport = New IntermediaryReportDaysTable
    Set clsSrcIntInsurance = New IntCrossSellingTable
    Set clsSrcCorrespondence = New IntCorrespondenceTable
    Set clsIntermediary = New Intermediary
    Set clsSrcIntBankAcc = New IntBankAccountTable
    Set clsTargetIntermediary = New IntermediaryTable
    Set clsSrcIntOrg = New IntermediaryOrganisation
    Set clsSrcIntInd = New IntermediaryIndividual
    
    TableAccess(clsSrcIntermediary).GetTableData
    Set colMatchValues = TableAccess(clsSrcIntermediary).GetKeyMatchValues
    'First Copy the Contact/Address details
    
    vAddressGUID = clsSrcIntermediary.GetAddressGUID
    vContactGUID = clsSrcIntermediary.GetContactDetailsGUID
    sIntType = clsSrcIntermediary.GetIntermediaryType
    sIndIntTypeID = clsIntermediary.GetValueIDForIntType(INTERMEDIARIES_INDIVIDUAL)
    
    'Contact Details
    CopyContactDetailsTable vContactGUID, sTargetDatabase
    
    'Address Details
    CopyAddressTable vAddressGUID, sTargetDatabase
    
    'Next copy the intermediary itself
    CopyRecords clsSrcIntermediary, New IntermediaryTable, sTargetDatabase

    'Next copy the Intermediary Name (either from Organisation or Individual table)
    If sIntType = sIndIntTypeID Then
        'Individual
        TableAccess(clsSrcIntInd).SetKeyMatchValues colMatchValues
    
        CopyRecords clsSrcIntInd, New IntermediaryIndividual, sTargetDatabase
        bPromoteProcFee = False
    Else
        'Organisation
        TableAccess(clsSrcIntOrg).SetKeyMatchValues colMatchValues
    
        CopyRecords clsSrcIntOrg, New IntermediaryOrganisation, sTargetDatabase
        bPromoteProcFee = True
    End If
    
    'Next Copy IntermediaryBankAccount details
    TableAccess(clsSrcIntBankAcc).SetKeyMatchValues colMatchValues
    CopyRecords clsSrcIntBankAcc, New IntBankAccountTable, sTargetDatabase

    'Now the Cross Selling Details
    TableAccess(clsSrcIntInsurance).SetKeyMatchValues colMatchValues
    CopyRecords clsSrcIntInsurance, New IntCrossSellingTable, sTargetDatabase
    
    'Correspondence
    TableAccess(clsSrcCorrespondence).SetKeyMatchValues colMatchValues
    CopyRecords clsSrcCorrespondence, New IntCorrespondenceTable, sTargetDatabase
    
    'Daily Report Days
    TableAccess(clsSrcIntDailyReport).SetKeyMatchValues colMatchValues
    CopyRecords clsSrcIntDailyReport, New IntermediaryReportDaysTable, sTargetDatabase
    
    'Proc Fee Target
    TableAccess(clsSrcIntProcFeeTarget).SetKeyMatchValues colMatchValues
    CopyRecords clsSrcIntProcFeeTarget, New IntermediaryTargetTable, sTargetDatabase
    
    If bPromoteProcFee Then
        CopyIntermediaryProcFee clsSrcIntermediary, sTargetDatabase
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyTask(clsSrcTask As TaskTable, sTargetDatabase As String, Optional bGetSourceData As Boolean = True)
    On Error GoTo Failed
    ' Things to be copied here are the Task itself, plus the
    ' TaskPriority table
    
    ' First, the Task
    
    Dim clsTargetTask As TaskTable
    Dim clsSrcTableAccess As TableAccess
    Dim clsTargetTableAccess As TableAccess
    Dim colMatchValues As Collection
    Dim sTaskID As String
    Dim rsSrc As ADODB.Recordset
    
    Set clsTargetTask = New TaskTable
    Set clsTargetTableAccess = clsTargetTask
    
    Set clsSrcTableAccess = clsSrcTask
    Set colMatchValues = clsSrcTableAccess.GetKeyMatchValues
    
    clsTargetTableAccess.SetKeyMatchValues colMatchValues
    clsTargetTableAccess.SetDatabase sTargetDatabase
    
    If bGetSourceData Then
        Set rsSrc = clsSrcTableAccess.GetTableData()
    Else
        Set rsSrc = clsSrcTableAccess.GetRecordSet
    End If
        
    If clsSrcTableAccess.RecordCount() > 0 Then
        sTaskID = clsSrcTask.GetTaskID()
        CopyObject clsSrcTask, clsTargetTask
        
        clsTargetTableAccess.Update
        
        ' Now, TaskPriority
        Dim clsSrcTaskPriority As TaskPriority
        Dim clsTargetTaskPriority As TaskPriority
    
        Set clsSrcTaskPriority = New TaskPriority
        Set clsSrcTableAccess = clsSrcTaskPriority
        clsSrcTableAccess.SetKeyMatchValues colMatchValues
        
        Set clsTargetTaskPriority = New TaskPriority
        Set clsTargetTableAccess = clsTargetTaskPriority
        
        clsTargetTableAccess.SetKeyMatchValues colMatchValues
        clsTargetTableAccess.SetDatabase sTargetDatabase
        clsTargetTableAccess.DeleteRecords sTargetDatabase
        
        CopyObject clsSrcTaskPriority, clsTargetTaskPriority, True
        clsTargetTableAccess.Update
        
        'MAR524 Promote TaskInterfaceMessage and TaskInterfaceSubMessage tables.
        Dim clsSrcTaskSubInterface As TaskInterfaceSubMessageTable
        Dim clsSrcTaskInterface As TaskInterfaceMessageTable
        Dim clsTargetTaskSubInterface As TaskInterfaceSubMessageTable
        Dim clsTargetTaskInterface As TaskInterfaceMessageTable
    
        Set clsSrcTaskSubInterface = New TaskInterfaceSubMessageTable
        Set clsSrcTaskInterface = New TaskInterfaceMessageTable
        
        Set clsTargetTaskSubInterface = New TaskInterfaceSubMessageTable
        Set clsTargetTaskInterface = New TaskInterfaceMessageTable
        
        'Delete messages and sub messages
        Set clsTargetTableAccess = clsTargetTaskSubInterface
        clsTargetTableAccess.SetKeyMatchValues colMatchValues
        clsTargetTableAccess.SetDatabase sTargetDatabase
        clsTargetTableAccess.DeleteRecords sTargetDatabase
        
        Set clsTargetTableAccess = clsTargetTaskInterface
        clsTargetTableAccess.SetKeyMatchValues colMatchValues
        clsTargetTableAccess.SetDatabase sTargetDatabase
        clsTargetTableAccess.DeleteRecords sTargetDatabase

        
        'Copy messages
        'Task Interface Message
        Set clsSrcTableAccess = clsSrcTaskInterface
        clsSrcTableAccess.SetKeyMatchValues colMatchValues
        
        Set clsTargetTableAccess = clsTargetTaskInterface
        clsTargetTableAccess.SetKeyMatchValues colMatchValues
                
        CopyObject clsSrcTaskInterface, clsTargetTaskInterface, True
        clsTargetTableAccess.Update
                
        'Task Interface Sub Message
        Set clsSrcTableAccess = clsSrcTaskSubInterface
        clsSrcTableAccess.SetKeyMatchValues colMatchValues
        
        Set clsTargetTableAccess = clsTargetTaskSubInterface
        clsTargetTableAccess.SetKeyMatchValues colMatchValues
                
        CopyObject clsSrcTaskSubInterface, clsTargetTaskSubInterface, True
        clsTargetTableAccess.Update
        
' TW 04/04/2007 EP2_2242 - Need to promote Linked Tasks
        Dim clsLinkedTask As LinkedTaskTable
        Dim clsTargetLinkedTask As LinkedTaskTable
        Dim clsTempSrcTask As TaskTable
        Dim colMatchFields As Collection
        Dim rs As ADODB.Recordset
        
        Set clsLinkedTask = New LinkedTaskTable
        Set clsTargetLinkedTask = New LinkedTaskTable
        
        Set colMatchFields = New Collection
        colMatchFields.Add "TASKID"
        
        ' Delete any existing LINKEDTASK table entries for this task from the target database
        Set clsTargetTableAccess = clsTargetLinkedTask
        clsTargetTableAccess.SetKeyMatchFields colMatchFields
        clsTargetTableAccess.SetKeyMatchValues colMatchValues
        clsTargetTableAccess.SetDatabase sTargetDatabase
        clsTargetTableAccess.DeleteRecords sTargetDatabase
        
        ' Promote any linked tasks necessary to the targtet database
        Set rs = g_clsDataAccess.ExecuteCommand("SELECT * FROM LINKEDTASK WHERE TASKID = '" & colMatchValues(1) & "'")
        If Not rs Is Nothing Then
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    Set clsTempSrcTask = New TaskTable
                    Set clsTargetTask = New TaskTable
                    Set colMatchValues = New Collection
                    colMatchValues.Add rs!LINKEDTASKID
                    Set clsSrcTableAccess = clsTempSrcTask
                    Set clsTargetTableAccess = clsTargetTask
                    
                    clsSrcTableAccess.SetKeyMatchValues colMatchValues
                    clsTargetTableAccess.SetKeyMatchValues colMatchValues
                    
                    ' Copy any tasks referred to by the linked task record
                    CopyObject clsTempSrcTask, clsTargetTask
                    clsTargetTableAccess.Update
                    
                    Set clsSrcTableAccess = clsLinkedTask
                    Set clsTargetTableAccess = clsTargetLinkedTask
                    
                    Set colMatchValues = New Collection
                    colMatchValues.Add rs!TASKID
                    
                    clsSrcTableAccess.SetKeyMatchFields colMatchFields
                    clsSrcTableAccess.SetKeyMatchValues colMatchValues
                    
                    clsTargetTableAccess.SetKeyMatchFields colMatchFields
                    clsTargetTableAccess.SetKeyMatchValues colMatchValues
                    clsTargetTableAccess.SetDatabase sTargetDatabase
                    
                    ' Copy the linked task
                    CopyObject clsSrcTableAccess, clsTargetTableAccess
                    clsTargetTableAccess.Update
                    
                    rs.MoveNext
                Loop
            End If
        End If
        
' TW 04/04/2007 EP2_2242 End
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyBusinessGroup(clsSrcGroup As BusinessGroupTable, sTargetDatabase As String)
    On Error GoTo Failed
    ' Things to be copied here are the Group, and the associated tasks
    
    ' First, the Group
    Dim clsTargetGroup As BusinessGroupTable
    Dim clsSrcTableAccess As TableAccess
    Dim clsTargetTableAccess As TableAccess
    Dim colMatchValues As Collection
    Dim vGroupID As Variant
    Dim rsSrc As ADODB.Recordset
    
    Set clsTargetGroup = New BusinessGroupTable
    Set clsTargetTableAccess = clsTargetGroup
    
    Set clsSrcTableAccess = clsSrcGroup
    Set colMatchValues = clsSrcTableAccess.GetKeyMatchValues
    
    clsTargetTableAccess.SetKeyMatchValues colMatchValues
    clsTargetTableAccess.SetDatabase sTargetDatabase
    
    Set rsSrc = clsSrcTableAccess.GetTableData(POPULATE_KEYS)
        
    If clsSrcTableAccess.RecordCount() > 0 Then
        ' DJP SQL Server port
        vGroupID = clsSrcGroup.GetGroupID
        CopyObject clsSrcGroup, clsTargetGroup
        
        clsTargetTableAccess.Update
        
        ' Now, The tasks
        Dim clsTaskLink As TaskLinkTable
        Dim clsTask As TaskTable
        Dim clsTargetTaskLink As TaskLinkTable
        Dim nTaskCnt As Long
        Dim colValues As Collection
        Dim sTaskID As String
        Dim iSrcCount As Integer      ' MAR1397
        Dim iTargetCount As Integer   ' MAR1397
        Dim i As Integer              ' MAR1397
        Dim sTaskIDs As String        ' MAR1397
        Dim sDeleteSQL As String      ' MAR1397
        
        Set colValues = New Collection
        Set clsTaskLink = New TaskLinkTable
        Set clsTask = New TaskTable
        Set clsTargetTaskLink = New TaskLinkTable
        
        TableAccess(clsTargetTaskLink).SetDatabase sTargetDatabase
               
        'DB BM0141 03/01/03
        vGroupID = g_clsSQLAssistSP.ByteArrayToGuidString(CStr(vGroupID))
        vGroupID = g_clsSQLAssistSP.FormatGuid(CStr(vGroupID))
        'DB End
        clsTaskLink.GetTasksForGroup CStr(vGroupID)
        iSrcCount = TableAccess(clsTaskLink).RecordCount    ' Source
        If iSrcCount > 0 Then
            TableAccess(clsTaskLink).MoveFirst
            For nTaskCnt = 1 To iSrcCount
                sTaskID = clsTaskLink.GetTaskID
                colValues.Add sTaskID
                TableAccess(clsTask).SetKeyMatchValues colValues
                TableAccess(clsTask).GetTableData POPULATE_KEYS
                CopyTask clsTask, sTargetDatabase, False
                colValues.Remove (1)
                TableAccess(clsTaskLink).MoveNext
            Next
        End If
          
        'Now update the tasklinkTable
        clsTargetTaskLink.GetTasksForGroup CStr(vGroupID)
            
        'MAR1397  Delete all the linked tasks first
        iTargetCount = TableAccess(clsTargetTaskLink).RecordCount
            
        sTaskIDs = ""
        If iTargetCount > 0 Then
            For i = 1 To iTargetCount
                If Len(sTaskIDs) > 0 Then
                    TableAccess(clsTargetTaskLink).MoveNext
                    sTaskIDs = sTaskIDs & " , " & "'" & clsTargetTaskLink.GetTaskID & "'"
                Else
                    TableAccess(clsTargetTaskLink).MoveFirst
                    sTaskIDs = "'" & clsTargetTaskLink.GetTaskID & "'"
                End If
            Next i
                
            TableAccess(clsTargetTaskLink).MoveFirst

            sDeleteSQL = "Delete from TaskLink where GroupID = " & vGroupID & " AND TaskID in ( " & sTaskIDs & ")"
            g_clsDataAccess.ExecuteCommand sDeleteSQL, sTargetDatabase
            
            TableAccess(clsTargetTaskLink).GetTableData POPULATE_EMPTY

        End If

        If (iSrcCount > 0) Then
            TableAccess(clsTaskLink).MoveFirst
            CopyObject clsTaskLink, clsTargetTaskLink, , False, False
        End If
        
        'Save Task Link Records
        TableAccess(clsTargetTaskLink).Update
                        
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub CopyPrintingTemplate(clsSrcPrintTemplate As PrintingTemplateTable, sTargetDatabase As String)
    On Error GoTo Failed
    ' Things to be copied here are the Template, and the associated stages
    
    Dim clsTargetPrintTemplate As PrintingTemplateTable
    Dim colMatchValues As Collection
    Dim sTemplateID As String
    Dim rsSrc As ADODB.Recordset
    'MAR45 GHun
    Dim colValues As Collection
    Dim sDPSTemplateID As String
    Dim clsSrcDocumentLocations As TemplateTable
    'MAR45 End
    Set clsTargetPrintTemplate = New PrintingTemplateTable
    
    Set colMatchValues = TableAccess(clsSrcPrintTemplate).GetKeyMatchValues
    
    TableAccess(clsTargetPrintTemplate).SetKeyMatchValues colMatchValues
    TableAccess(clsTargetPrintTemplate).SetDatabase sTargetDatabase
    
    Set rsSrc = TableAccess(clsSrcPrintTemplate).GetTableData(POPULATE_KEYS)
        
    If TableAccess(clsSrcPrintTemplate).RecordCount() > 0 Then
        sTemplateID = clsSrcPrintTemplate.GetHostTemplateID
        
        'MAR45 GHun
        sDPSTemplateID = clsSrcPrintTemplate.GetDPSTemplateID
        
        'Promote Print document locations first
        Set clsSrcDocumentLocations = New TemplateTable
        Set colValues = New Collection
        colValues.Add sDPSTemplateID
        
        TableAccess(clsSrcDocumentLocations).SetKeyMatchValues colValues
        TableAccess(clsSrcDocumentLocations).GetTableData
        CopyRecords clsSrcDocumentLocations, New TemplateTable, sTargetDatabase, False
        
        Set colValues = Nothing
        'MAR45 End
        
        CopyObject clsSrcPrintTemplate, clsTargetPrintTemplate, , False
        TableAccess(clsTargetPrintTemplate).Update
        
        ' Now, The Stages
        Dim clsStage As StageTable
        'Dim nThisStage As Long
        'Dim sStage As String
        Dim clsAvailableTemplates As AvailableTemplatesTable
        Dim clsTargetAvailableTemplate As AvailableTemplatesTable
        
        Set colValues = New Collection
        Set clsStage = New StageTable
        Set clsAvailableTemplates = New AvailableTemplatesTable
        Set clsTargetAvailableTemplate = New AvailableTemplatesTable
        
        'MAR45 GHun
        'TableAccess(clsTargetAvailableTemplate).SetDatabase sTargetDatabase
        TableAccess(clsAvailableTemplates).SetKeyMatchValues colMatchValues
        TableAccess(clsAvailableTemplates).GetTableData POPULATE_KEYS
        
        CopyRecords clsAvailableTemplates, clsTargetAvailableTemplate, sTargetDatabase, True
        TableAccess(clsTargetAvailableTemplate).Update
        
        ''Are there any stages to promote?
        'If TableAccess(clsAvailableTemplates).RecordCount > 0 Then
        '    TableAccess(clsAvailableTemplates).MoveFirst
        '    For nThisStage = 1 To TableAccess(clsAvailableTemplates).RecordCount
        '        sStage = clsAvailableTemplates.GetStageID
        '        colValues.Add sStage
        '        TableAccess(clsStage).SetKeyMatchValues colValues
        '        TableAccess(clsStage).GetTableData POPULATE_KEYS
        '        CopyStage clsStage, sTargetDatabase, , EnumSingleStage
        '        colValues.Remove (1)
        '        TableAccess(clsAvailableTemplates).MoveNext
        '    Next
        '
        '    'Now update the AvailableTemplates table
        '    clsTargetAvailableTemplate.GetSelectedStages sTemplateID
        '    TableAccess(clsAvailableTemplates).MoveFirst
        '    CopyObject clsAvailableTemplates, clsTargetAvailableTemplate, , False, False
        '    TableAccess(clsTargetAvailableTemplate).Update
        'End If
        'MAR45 End
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyActivity(clsSrcActivity As ActivityTable, sTargetDatabase As String, Optional bPromoteStages As Boolean = True)
    On Error GoTo Failed
    
    'MAR1120. Do not promote stages associated with the activity (unnecessary)
    
    ' First, the Activity
    Dim clsTargetActivity As ActivityTable
    Dim clsSrcTableAccess As TableAccess
    Dim clsTargetTableAccess As TableAccess
    Dim colMatchValues As Collection
    Dim sActivityID As String
    Dim rsSrc As ADODB.Recordset
    
    Set clsTargetActivity = New ActivityTable
    Set clsTargetTableAccess = clsTargetActivity
    
    Set clsSrcTableAccess = clsSrcActivity
    Set colMatchValues = clsSrcTableAccess.GetKeyMatchValues
    
    clsTargetTableAccess.SetKeyMatchValues colMatchValues
    clsTargetTableAccess.SetDatabase sTargetDatabase
    
    Set rsSrc = clsSrcTableAccess.GetTableData()
        
    sActivityID = clsSrcActivity.GetActivityID()
    CopyObject clsSrcActivity, clsTargetActivity
    
    clsTargetTableAccess.Update
    
    'MAR1120
'    'Promote Stages
'    clsStage.GetStagesForActivity sActivityID, enumStageAllIncDelete, True
'
'    'Are there any Stages to promote?
'    If TableAccess(clsStage).RecordCount > 0 And bPromoteStages Then
'        CopyStage clsStage, sTargetDatabase, True, enumStageAll
'    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyStage(clsSrcStage As StageTable, sTargetDatabase As String, Optional bGetTableData As Boolean = True, Optional nStageType As PromoteStageType = EnumSingleStage)
    On Error GoTo Failed
    ' Things to be copied here are the Stage itself, plus the
    ' any associated tasks and Activities
    
    Dim clsTargetStage As StageTable
    Dim clsSrcTableAccess As TableAccess
    Dim clsTargetTableAccess As TableAccess
    Dim clsActivity As ActivityTable
    Dim colActivityValues As Collection
    Dim colMatchValues As Collection
    Dim sStageID As String
    Dim vActivityID As Variant
    Dim rsSrc As ADODB.Recordset
    Dim nStageTaskCnt As Long
    Dim iCount As Integer       ' MAR1120
    Dim sTaskIDs As String      ' MAR1120
    Dim i As Integer            ' MAR1120
    Dim sDeleteSQL As String    ' MAR1120
    
    Set clsTargetStage = New StageTable
    Set clsTargetTableAccess = clsTargetStage
    
    Set colMatchValues = New Collection
    Set clsActivity = New ActivityTable
    Set colActivityValues = New Collection
    Set clsSrcTableAccess = clsSrcStage
    
    clsTargetTableAccess.SetDatabase sTargetDatabase
        
    'Promote Activity for Stage
    If bGetTableData And nStageType <> EnumNone Then
        If nStageType = EnumSingleStage Then
            Set colMatchValues = clsSrcTableAccess.GetKeyMatchValues
            clsTargetTableAccess.SetKeyMatchValues colMatchValues
            clsTargetTableAccess.GetTableData
            Set rsSrc = clsSrcTableAccess.GetTableData
            vActivityID = clsSrcStage.GetActivityID
        Else
            vActivityID = clsSrcStage.GetActivityID
            clsSrcStage.GetStagesForActivity vActivityID, enumStageAllIncDelete, True
            clsTargetStage.GetStagesForActivity vActivityID, enumStageAllIncDelete, True
            Set rsSrc = TableAccess(clsSrcStage).GetRecordSet
        End If
    Else
        Set rsSrc = clsSrcTableAccess.GetRecordSet
    End If
    
    'Are there any Stages to promote?
    If rsSrc.RecordCount() > 0 Then
        sStageID = clsSrcStage.GetStageID()
        vActivityID = clsSrcStage.GetActivityID
        
        'Is there an Activity associated with this Stage?
        If Not IsNull(vActivityID) And nStageType <> EnumNone Then
            If Len(CStr(vActivityID)) > 0 Then
                'Yes
                colActivityValues.Add vActivityID
                TableAccess(clsActivity).SetKeyMatchValues colActivityValues
                CopyActivity clsActivity, sTargetDatabase, False
            End If
        End If
        
        TableAccess(clsTargetStage).SetDatabase sTargetDatabase
        
        'Right, when you promote a activity, with stages that have been DEAssociated witht the stage (ie moved to the left)
        'the the two record sets are not equal therefore a UNIQUE constraint is ?????. So set all Activiytids on the stage table = NULL
        CopyObject clsSrcStage, clsTargetStage, , False, False
        
        TableAccess(clsTargetStage).Update
        
        'Now StageTask
        Dim clsSrcStageTask As StageTaskTable
        Dim clsTargetStageTask As StageTaskTable
        Dim sTaskID As String
        
        Set clsTargetStageTask = New StageTaskTable
        Set clsSrcStageTask = New StageTaskTable
        
        TableAccess(clsSrcStage).MoveFirst
        
        'MAR1120 Changes to prevent duplicate key errors (as per BBG changes)
        For nStageTaskCnt = 1 To TableAccess(clsSrcStage).RecordCount
            
            sStageID = clsSrcStage.GetStageID
            
            'Populate Source with the stagetask records for current stageid
            clsSrcStageTask.GetTasksForStage sStageID, True
            TableAccess(clsTargetStageTask).SetDatabase sTargetDatabase
                    
            'First delete these tasks from the target so that we don't get any key
            'violations when copying tasks across.
            iCount = TableAccess(clsSrcStageTask).RecordCount
            sTaskIDs = ""
            If iCount > 0 Then
                For i = 1 To iCount
                    If Len(sTaskIDs) > 0 Then
                        TableAccess(clsSrcStageTask).MoveNext
                        sTaskIDs = sTaskIDs & " , " & "'" & clsSrcStageTask.GetTaskID & "'"
                    Else
                        TableAccess(clsSrcStageTask).MoveFirst
                        sTaskIDs = "'" & clsSrcStageTask.GetTaskID & "'"
                    End If
                Next i
            
                TableAccess(clsSrcStageTask).MoveFirst
                
                sDeleteSQL = "Delete from stageTask where StageID = '" & sStageID & "' AND TaskID in ( " & sTaskIDs & ")"
                g_clsDataAccess.ExecuteCommand sDeleteSQL, sTargetDatabase
                    
                TableAccess(clsTargetStageTask).GetTableData POPULATE_EMPTY
                
                CopyObject clsSrcStageTask, clsTargetStageTask, , False, False
                        
                'Save Stage Task Records
                TableAccess(clsTargetStageTask).Update
            End If
                                 
            ' Now, Tasks
            Dim clsSrcTask As TaskTable
            Dim col As Collection
            Dim nCnt As Long
            Set col = New Collection
            Set clsSrcTask = New TaskTable
            
            If TableAccess(clsSrcStageTask).RecordCount Then
                TableAccess(clsSrcStageTask).MoveFirst
                
                For nCnt = 1 To TableAccess(clsSrcStageTask).RecordCount
                    sTaskID = clsSrcStageTask.GetTaskID
                    col.Add sTaskID
                    TableAccess(clsSrcTask).SetKeyMatchValues col
                    If Not bGetTableData Then
                        TableAccess(clsSrcTask).GetTableData
                    End If
                    'Save the Task Records
                    CopyTask clsSrcTask, sTargetDatabase, bGetTableData
                    col.Remove (1)
                    TableAccess(clsSrcStageTask).MoveNext
                Next
            End If
            TableAccess(clsSrcStage).MoveNext
        Next
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub GetBandedTable(clsSrcTableAccess As TableAccess, clsBandedTable As TableAccess, tableType As BandedTableType)
    On Error GoTo Failed
    Dim sObject As String
    Dim clsAdminFee As AdminFeeTable
    Dim clsValuationFee As ValuationFeeTable
    Dim clsBaseRate As BaseRateTable
    Dim clsPayProtRates As PayProtRatesTable
    Dim sAdminFeeTable As String
    Dim sValuationFeeTable As String
    
' TW 09/10/2006 EP2_7
    Dim sAdditionalBorrowingFeeTable As String
    Dim sCreditLimitIncreaseFeeTable As String
    
    Dim clsAdditionalBorrowingFeeTable As AdditionalBorrowingFeeTable
    Dim clsCreditLimitIncreaseFeeTable As CreditLimitIncreaseFeeTable
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
    Dim sTransferOfEquityFeeTable As String

    Dim clsTranferOfEquityFeeTable As TransferOfEquityFeeTable
' TW 11/12/2006 EP2_20 End

    '*=[MC]BMIDS763 - NEW FEESETS PROMOTION ENABLED
    Dim clsTTFeeSetTable As TTFeeBand
    Dim clsIAFeeSetTable As InsuranceAdminFeeBand
    Dim clsPSFeeSetTable As ProductSwitchFeeTable
    'SECTION END
    
    Dim sTTFeeTable As String
    Dim sIAFeeTable As String
    Dim sPSFeeTable As String
    
    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions START
    Dim sRedemptionFeeTable As String
    Dim sMPMIGRateSetTable As String
    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions END
    
    Dim sBaseRateTable As String
    Dim sPayProtRates As String
    Dim sWorkingHoursTable As String
    Dim clsTableAccess As TableAccess
    Dim clsWorkingHours As WorkingHoursTable
    Dim sFunctionName As String
    
    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions START
    Dim clsRedemptionFee As RedemptionFeeTable
    Dim clsMPMIGRateSet As MPMigRateTable
    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions END
    
    sFunctionName = "GetBandedTable"
    
' TW 09/10/2006 EP2_7
    Set clsAdditionalBorrowingFeeTable = New AdditionalBorrowingFeeTable
    Set clsCreditLimitIncreaseFeeTable = New CreditLimitIncreaseFeeTable
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
    Set clsTranferOfEquityFeeTable = New TransferOfEquityFeeTable
' TW 11/12/2006 EP2_20 End
    
    Set clsAdminFee = New AdminFeeTable
    Set clsValuationFee = New ValuationFeeTable
    
    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions START
    Set clsRedemptionFee = New RedemptionFeeTable
    Set clsMPMIGRateSet = New MPMigRateTable
    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions END
    
    '*=[MC]BMIDS763 - NEW FEESETS
    Set clsTTFeeSetTable = New TTFeeBand
    Set clsIAFeeSetTable = New InsuranceAdminFeeBand
    Set clsPSFeeSetTable = New ProductSwitchFeeTable
    'end
    
    Set clsBaseRate = New BaseRateTable
    Set clsPayProtRates = New PayProtRatesTable
    Set clsWorkingHours = New WorkingHoursTable
    
    Set clsTableAccess = clsWorkingHours
    sWorkingHoursTable = TypeName(clsTableAccess)

' TW 09/10/2006 EP2_7
    Set clsTableAccess = clsAdditionalBorrowingFeeTable
    sAdditionalBorrowingFeeTable = TypeName(clsTableAccess)
    
    Set clsTableAccess = clsCreditLimitIncreaseFeeTable
    sCreditLimitIncreaseFeeTable = TypeName(clsTableAccess)
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
    Set clsTableAccess = clsTranferOfEquityFeeTable
    sTransferOfEquityFeeTable = TypeName(clsTableAccess)
' TW 11/12/2006 EP2_20 End
    
    Set clsTableAccess = clsAdminFee
    sAdminFeeTable = TypeName(clsTableAccess)
    '*=[MC]BMIDS763 - FEESETS
    Set clsTableAccess = clsTTFeeSetTable
    sTTFeeTable = TypeName(clsTableAccess)
    
    Set clsTableAccess = clsIAFeeSetTable
    sIAFeeTable = TypeName(clsTableAccess)
    
    Set clsTableAccess = clsPSFeeSetTable
    sPSFeeTable = TypeName(clsTableAccess)
    'SECTION END
    
    Set clsTableAccess = clsValuationFee
    sValuationFeeTable = TypeName(clsTableAccess)
    
    
    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions START
    Set clsTableAccess = clsRedemptionFee
    sRedemptionFeeTable = TypeName(clsTableAccess)
    Set clsTableAccess = clsMPMIGRateSet
    sMPMIGRateSetTable = TypeName(clsTableAccess)
    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions END
    
    'JD BMIDS765 promotion of rentalincomeratesetbands
    Dim clsRentalIncome As RentalIncomeRateSetBandTable
    Set clsRentalIncome = New RentalIncomeRateSetBandTable
    Set clsTableAccess = clsRentalIncome
    Dim sRentalIncomeTable As String
    sRentalIncomeTable = TypeName(clsTableAccess)
    
    Set clsTableAccess = clsBaseRate
    sBaseRateTable = TypeName(clsTableAccess)
    
    Set clsTableAccess = clsPayProtRates
    sPayProtRates = TypeName(clsTableAccess)
    
    sObject = TypeName(clsSrcTableAccess)

    Select Case sObject
        Case sWorkingHoursTable
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New WorkingHourTypeTable
            Else
                Set clsBandedTable = New WorkingHoursTable
            End If
        
' TW 09/10/2006 EP2_7
        Case sAdditionalBorrowingFeeTable
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New AddnlBorrFeeSetTable
            Else
                Set clsBandedTable = New AdditionalBorrowingFeeTable
            End If
            
        Case sCreditLimitIncreaseFeeTable
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New CreditLimitIncFeeSetTable
            Else
                Set clsBandedTable = New CreditLimitIncreaseFeeTable
            End If

' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
        Case sTransferOfEquityFeeTable
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New TransferOfEquityFeeSetTable
            Else
                Set clsBandedTable = New TransferOfEquityFeeTable
            End If
' TW 11/12/2006 EP2_20 End
        
        Case sAdminFeeTable
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New AdminFeeSetTable
            Else
                Set clsBandedTable = New AdminFeeTable
            End If
        '*=[MC]BMIDS763 - FEE SET UPDATES
        Case sTTFeeTable
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New TTFeeSetTable
            Else
                Set clsBandedTable = New TTFeeBand
            End If
        
        Case sIAFeeTable
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New InsuranceAdminFeeSetTable
            Else
                Set clsBandedTable = New InsuranceAdminFeeBand
            End If
        
        Case sPSFeeTable
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New ProductSwitchFeeSetTable
            Else
                Set clsBandedTable = New ProductSwitchFeeTable
            End If
        '*=[MC]BMIDS763 - SECTION END
        
        'JD BMIDS765 added rental income rate sets
        Case sRentalIncomeTable
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New RentalIncomeRateSetTable
            Else
                Set clsBandedTable = New RentalIncomeRateSetBandTable
            End If
            
        Case sValuationFeeTable
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New ValuationFeeSetTable
            Else
                Set clsBandedTable = New ValuationFeeTable
            End If
        
        'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions START
        Case sRedemptionFeeTable
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New RedemptionFeeSetTable
            Else
                Set clsBandedTable = New RedemptionFeeTable
            End If

        Case sMPMIGRateSetTable
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New MPMigRateSetTable
            Else
                Set clsBandedTable = New MPMigRateTable
            End If
        'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions END
        
        Case sPayProtRates
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New PayProtRatesSetTable
            Else
                Set clsBandedTable = New PayProtRatesTable
            End If
        
        Case sBaseRateTable
            If tableType = typeBandedSetTable Then
                Set clsBandedTable = New BaseRateSetTable
            Else
                Set clsBandedTable = New BaseRateTable
            End If
            
        Case Else
            g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Unknown object " + sObject
    End Select

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub CopyBandedFee(clsSrcBandedFee As TableAccess, sTargetDatabase As String, Optional bRemoveTarget As Boolean = True)
    On Error GoTo Failed
    
    ' First the Banded Fee Set
    Dim clsSrcBandedFeeSet As TableAccess
    Dim clsTargetBandedFeeSet As TableAccess
    Dim colMatchValues As Collection
    Dim colFeeSetValues As Collection
    Dim sFeeSet As String
    
    Const nFeeSetKey As Integer = 1
    
    GetBandedTable clsSrcBandedFee, clsTargetBandedFeeSet, typeBandedSetTable
    GetBandedTable clsSrcBandedFee, clsSrcBandedFeeSet, typeBandedSetTable
    
    clsTargetBandedFeeSet.SetDatabase sTargetDatabase
    
    Set colMatchValues = clsSrcBandedFee.GetKeyMatchValues()
    
    ' The feeset is the first key
    sFeeSet = colMatchValues(nFeeSetKey)
    
    If Len(sFeeSet) > 0 Then
        Set colFeeSetValues = New Collection
        colFeeSetValues.Add sFeeSet
        clsSrcBandedFeeSet.SetKeyMatchValues colFeeSetValues
        clsTargetBandedFeeSet.SetKeyMatchValues colFeeSetValues
        
        CopyObject clsSrcBandedFeeSet, clsTargetBandedFeeSet
        
        clsTargetBandedFeeSet.Update
    
        ' Now the Banded Fee table
        Dim clsTargetBandedFee As TableAccess

        GetBandedTable clsSrcBandedFee, clsTargetBandedFee, typeBandedTable

        clsTargetBandedFee.SetDatabase sTargetDatabase
        clsTargetBandedFee.SetKeyMatchValues colMatchValues
        
        ' DJP SYS4188, don't just delete the records using a SQL Statement because we'll case a trigger to fire and complain
        ' about Base Ratesbeing used by a mortgage product. Instead, allow the caller to decide.
        If bRemoveTarget Then
            clsTargetBandedFee.DeleteRecords sTargetDatabase
        End If
        
        CopyObject clsSrcBandedFee, clsTargetBandedFee
        
' TW 27/03/2007 EP2_1101
        DeleteUnusedEntries clsSrcBandedFee, clsTargetBandedFee, colMatchValues(1)
' TW 27/03/2007 EP2_1101 End

        clsTargetBandedFee.Update
    
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub CopyCountry(clsSrcCountry As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    Dim clsTargetCountry As CountryTable
    Dim clsTargetTableAccess As TableAccess
    Dim colMatchValues As Collection
    
    Set clsTargetCountry = New CountryTable
    Set clsTargetTableAccess = clsTargetCountry
    
    Set colMatchValues = clsSrcCountry.GetKeyMatchValues
    clsTargetTableAccess.SetKeyMatchValues colMatchValues
    clsTargetTableAccess.SetDatabase sTargetDatabase
    
    CopyObject clsSrcCountry, clsTargetCountry
    clsTargetTableAccess.Update
    
    ' Now copy Bank Holiday entries, if present
    Dim clsSrcBankHoliday As TableAccess
    Dim clsTargetBankHoliday As TableAccess
    
    Set clsSrcBankHoliday = New BankHolidayTable
    Set clsTargetBankHoliday = New BankHolidayTable
    
    clsSrcBankHoliday.SetKeyMatchValues colMatchValues
    clsTargetBankHoliday.SetKeyMatchValues colMatchValues
    
    clsTargetBankHoliday.SetDatabase sTargetDatabase
    clsTargetBankHoliday.DeleteRecords
    
    CopyObject clsSrcBankHoliday, clsTargetBankHoliday, True
    clsTargetBankHoliday.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyIntroducer(clsIntroducer As IntroducerTable, sTargetDatabase As String)
' TW 27/03/2007 EP2_1101 - New routine
Dim sIntroducerId As String
Dim sARFirmId As String
Dim sPRINCIPALFIRMID As String
Dim clsTableAccess As TableAccess
Dim colMatchValues As Collection
Dim iManagement As ISecurityManager
Dim clsConnection As SupervisorConnection

Dim clsARFirm As TableAccess
Dim clsIntroducerFirm As TableAccess
Dim clsPrincipalFirm As TableAccess

Dim clsTargetARFirm As TableAccess
Dim clsTargetIntroducerFirm As TableAccess
Dim clsTargetIntroducer As TableAccess
Dim clsTargetPrincipalFirm As TableAccess

Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
    
' TW 14/05/2007 VR216
Dim clsUserContactTable As OrgUserContactTable
Dim clsSrcContactTable As TableAccess
Dim clsTrgContactTable As TableAccess
Dim clsSrcTelephoneTable As TableAccess
Dim clsTrgTelephoneTable As TableAccess
Dim clsSrcUserContactTable As TableAccess
Dim clsTrgUserContactTable As TableAccess

Dim sUserId As String
Dim vContactGUID As Variant
' TW 14/05/2007 VR216 End
    
    On Error GoTo Failed
    
    Set clsTableAccess = clsIntroducer
    GetDatabaseObject clsIntroducer
    
    sIntroducerId = clsIntroducer.GetIntroducerID()
' TW 14/05/2007 VR216
    sUserId = clsIntroducer.GetUserID
' TW 14/05/2007 VR216 End
    
    Set colMatchValues = clsTableAccess.GetKeyMatchValues
    
    ' First, Principal Firm
    Set clsTargetIntroducer = New IntroducerTable
    clsTargetIntroducer.SetKeyMatchValues colMatchValues
    clsTargetIntroducer.SetDatabase sTargetDatabase
    
    CopyObject clsIntroducer, clsTargetIntroducer, , False
    clsTargetIntroducer.Update
    
    ' Next, IntroducerFirm
    Set clsIntroducerFirm = New IntroducerFirmTable
    Set clsTargetIntroducerFirm = New IntroducerFirmTable

    Set rs = g_clsDataAccess.ExecuteCommand("SELECT * FROM IntroducerFirm WHERE IntroducerID = '" & colMatchValues(1) & "'")
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            Do While Not rs.EOF
                sARFirmId = Format$(rs!ARFIRMID)
                If Len(sARFirmId) > 0 Then
' If linked to an AR Firm, promote it if necessary
                    Set clsARFirm = New ARFirmTable
                    Set clsTargetARFirm = New ARFirmTable
                    
                    Set colMatchValues = New Collection
                    colMatchValues.Add sARFirmId
                    clsARFirm.SetKeyMatchValues colMatchValues
                    clsTargetARFirm.SetKeyMatchValues colMatchValues
                    CopyObject clsARFirm, clsTargetARFirm
                    clsTargetARFirm.Update
                    
                End If
                sPRINCIPALFIRMID = Format$(rs!PRINCIPALFIRMID)
                If Len(sPRINCIPALFIRMID) > 0 Then
' If linked to a Principal Firm, promote it if necessary
                    Set clsPrincipalFirm = New PrincipalFirmTable
                    Set clsTargetPrincipalFirm = New PrincipalFirmTable
                    Set colMatchValues = New Collection
                    colMatchValues.Add sPRINCIPALFIRMID
                    clsPrincipalFirm.SetKeyMatchValues colMatchValues
                    clsTargetPrincipalFirm.SetKeyMatchValues colMatchValues
                    CopyObject clsPrincipalFirm, clsTargetPrincipalFirm
                    clsTargetPrincipalFirm.Update
                
                End If
                rs.MoveNext
            Loop
            g_clsDataAccess.ExecuteCommand "DELETE FROM IntroducerFirm where INTRODUCERID = '" & sIntroducerId & "'", m_sTargetDatabase
            CopyObjectNonPrimaryKey clsIntroducerFirm, clsTargetIntroducerFirm, rs
        End If
    End If
    
' TW 14/05/2007 VR216
    'Deal with any telephone numbers
    
    Set clsUserContactTable = New OrgUserContactTable
    
    Set colMatchValues = New Collection
    colMatchValues.Add sUserId
    TableAccess(clsUserContactTable).SetKeyMatchValues colMatchValues
    TableAccess(clsUserContactTable).GetTableData
    
    If TableAccess(clsUserContactTable).RecordCount > 0 Then
        
        vContactGUID = clsUserContactTable.GetContactDetailsGUID
        Set colMatchValues = New Collection
        colMatchValues.Add vContactGUID
        
        'Contact Details.
        Set clsSrcContactTable = New ContactDetailsTable
        Set clsTrgContactTable = New ContactDetailsTable
        
        clsSrcContactTable.SetKeyMatchValues colMatchValues
        clsTrgContactTable.SetKeyMatchValues colMatchValues
        clsTrgContactTable.SetDatabase sTargetDatabase
        
        CopyObject clsSrcContactTable, clsTrgContactTable
        clsTrgContactTable.Update
    
        'Contact Telephone Details.
        Set clsSrcTelephoneTable = New ContactDetailsTelephoneTable
        Set clsTrgTelephoneTable = New ContactDetailsTelephoneTable
        
        clsSrcTelephoneTable.SetKeyMatchValues colMatchValues
        clsTrgTelephoneTable.SetKeyMatchValues colMatchValues
        clsTrgTelephoneTable.SetDatabase sTargetDatabase
        
        CopyObject clsSrcTelephoneTable, clsTrgTelephoneTable
        clsTrgTelephoneTable.Update
    
        'User-to-Contact mapping.
        Set clsSrcUserContactTable = New OrgUserContactTable
        Set clsTrgUserContactTable = New OrgUserContactTable
        
        clsSrcUserContactTable.SetKeyMatchValues colMatchValues
        clsTrgUserContactTable.SetKeyMatchValues colMatchValues
        clsTrgUserContactTable.SetDatabase sTargetDatabase
        
        CopyObject clsSrcUserContactTable, clsTrgUserContactTable
        clsTrgUserContactTable.Update
    End If
' TW 14/05/2007 VR216 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub CopyARFirm(clsARFirm As ARFirmTable, sTargetDatabase As String)
' TW 27/03/2007 EP2_1101 - New routine
Dim sARFirmId As String
Dim clsTableAccess As TableAccess
Dim colMatchValues As Collection
Dim iManagement As ISecurityManager
Dim clsConnection As SupervisorConnection

Dim clsAppointments As TableAccess
Dim clsFirmPermissions As TableAccess
Dim clsTradingName As TableAccess

Dim clsTargetAppointments As TableAccess
Dim clsTargetFirmPermissions As TableAccess
Dim clsTargetARFirm As TableAccess
Dim clsTargetTradingName As TableAccess

Dim rs As ADODB.Recordset
    
    On Error GoTo Failed
    
    Set clsTableAccess = clsARFirm
    GetDatabaseObject clsARFirm
    
    sARFirmId = clsARFirm.GetARFirmID()
    
    Set colMatchValues = clsTableAccess.GetKeyMatchValues
    
    ' First, Principal Firm
    Set clsTargetARFirm = New ARFirmTable
    clsTargetARFirm.SetKeyMatchValues colMatchValues
    clsTargetARFirm.SetDatabase sTargetDatabase
    
    CopyObject clsARFirm, clsTargetARFirm, , False
    clsTargetARFirm.Update
    
    ' Next, Firm Activities
    g_clsDataAccess.ExecuteCommand "DELETE FROM FIRMPERMISSIONS where ARFIRMID = '" & colMatchValues(1) & "'", m_sTargetDatabase
    Set clsFirmPermissions = New FirmPermissionsTable
    Set clsTargetFirmPermissions = New FirmPermissionsTable
    
    Set rs = g_clsDataAccess.ExecuteCommand("SELECT * FROM FIRMPERMISSIONS WHERE ARFirmID = '" & colMatchValues(1) & "'")
    CopyObjectNonPrimaryKey clsFirmPermissions, clsTargetFirmPermissions, rs
    
    ' Next, Trading Names
    g_clsDataAccess.ExecuteCommand "DELETE FROM FIRMTRADINGNAME where ARFIRMID = '" & colMatchValues(1) & "'", m_sTargetDatabase
    Set clsTradingName = New FirmTradingNameTable
    Set clsTargetTradingName = New FirmTradingNameTable
    
    Set rs = g_clsDataAccess.ExecuteCommand("SELECT * FROM FIRMTRADINGNAME WHERE ARFirmID = '" & colMatchValues(1) & "'")
    CopyObjectNonPrimaryKey clsTradingName, clsTargetTradingName, rs
    
    ' Next, Appointments
    g_clsDataAccess.ExecuteCommand "DELETE FROM APPOINTMENTS where ARFIRMID = '" & colMatchValues(1) & "'", m_sTargetDatabase
    Set clsAppointments = New AppointmentsTable
    Set clsTargetAppointments = New AppointmentsTable
    
    Set rs = g_clsDataAccess.ExecuteCommand("SELECT * FROM APPOINTMENTS WHERE ARFirmID = '" & colMatchValues(1) & "'")
    CopyObjectNonPrimaryKey clsAppointments, clsTargetAppointments, rs
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyPrincipalFirm(clsPrincipalFirm As PrincipalFirmTable, sTargetDatabase As String)
' TW 27/03/2007 EP2_1101 - New routine
Dim sPRINCIPALFIRMID As String
Dim sClubAssociationId As String
Dim clsTableAccess As TableAccess
Dim colMatchValues As Collection
Dim iManagement As ISecurityManager
Dim clsConnection As SupervisorConnection

Dim clsFirmClubNetworkAssociation As TableAccess
Dim clsFirmPermissions As TableAccess
Dim clsMortgageClubNetAssocTable As TableAccess
Dim clsTradingName As TableAccess

Dim clsTargetFirmClubNetworkAssociation As TableAccess
Dim clsTargetFirmPermissions As TableAccess
Dim clsTargetMortgageClubNetAssocTable As TableAccess
Dim clsTargetPrincipalFirm As TableAccess
Dim clsTargetTradingName As TableAccess

' TW 05/04/2007 EP2_2292
Dim clsIntroducer As TableAccess
Dim clsUser As TableAccess

Dim strPackagerIndicator As String
Dim strUserId As String
' TW 05/04/2007 EP2_2292 End

Dim rs As ADODB.Recordset
    
    On Error GoTo Failed
    
    Set clsTableAccess = clsPrincipalFirm
    GetDatabaseObject clsPrincipalFirm
    
    sPRINCIPALFIRMID = clsPrincipalFirm.GetPrincipalFirmID()
' TW 05/04/2007 EP2_2292
    strPackagerIndicator = clsPrincipalFirm.GetPackagerIndicator()
    strUserId = "U" & clsPrincipalFirm.GetUnitId()
' TW 05/04/2007 EP2_2292 End
    
    Set colMatchValues = clsTableAccess.GetKeyMatchValues
    
    ' First, Principal Firm
    Set clsTargetPrincipalFirm = New PrincipalFirmTable
    clsTargetPrincipalFirm.SetKeyMatchValues colMatchValues
    clsTargetPrincipalFirm.SetDatabase sTargetDatabase
    
    CopyObject clsPrincipalFirm, clsTargetPrincipalFirm, , False
    clsTargetPrincipalFirm.Update
    
    ' Next, Firm Activities
    g_clsDataAccess.ExecuteCommand "DELETE FROM FIRMPERMISSIONS where PRINCIPALFIRMID = '" & colMatchValues(1) & "'", m_sTargetDatabase
    Set clsFirmPermissions = New FirmPermissionsTable
    Set clsTargetFirmPermissions = New FirmPermissionsTable
    
    Set rs = g_clsDataAccess.ExecuteCommand("SELECT * FROM FIRMPERMISSIONS WHERE PRINCIPALFIRMID = '" & colMatchValues(1) & "'")
    CopyObjectNonPrimaryKey clsFirmPermissions, clsTargetFirmPermissions, rs
    
    ' Next, Trading Names
    g_clsDataAccess.ExecuteCommand "DELETE FROM FIRMTRADINGNAME where PRINCIPALFIRMID = '" & colMatchValues(1) & "'", m_sTargetDatabase
    Set clsTradingName = New FirmTradingNameTable
    Set clsTargetTradingName = New FirmTradingNameTable
    
    Set rs = g_clsDataAccess.ExecuteCommand("SELECT * FROM FIRMTRADINGNAME WHERE PRINCIPALFIRMID = '" & colMatchValues(1) & "'")
    CopyObjectNonPrimaryKey clsTradingName, clsTargetTradingName, rs
    
    ' Next, Associations
    g_clsDataAccess.ExecuteCommand "DELETE FROM FIRMCLUBNETWORKASSOCIATION WHERE ISNULL(CLUBNETWORKASSOCIATIONID, '') != '' AND PRINCIPALFIRMID = '" & colMatchValues(1) & "'", m_sTargetDatabase
    Set clsFirmClubNetworkAssociation = New FirmClubNetAssocTable
    Set clsTargetFirmClubNetworkAssociation = New FirmClubNetAssocTable
    
    Set rs = g_clsDataAccess.ExecuteCommand("SELECT * FROM FIRMCLUBNETWORKASSOCIATION WHERE ISNULL(CLUBNETWORKASSOCIATIONID, '') != '' AND PRINCIPALFIRMID = '" & colMatchValues(1) & "'")
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            rs.MoveFirst
            Do While Not rs.EOF
                sClubAssociationId = Format$(rs!CLUBNETWORKASSOCIATIONID)
                If Len(sClubAssociationId) > 0 Then
' If linked to a Mortgage Club or Association, promote it if necessary
                    Set clsMortgageClubNetAssocTable = New MortgageClubNetAssocTable
                    Set clsTargetMortgageClubNetAssocTable = New MortgageClubNetAssocTable
                    Set colMatchValues = New Collection
                    colMatchValues.Add sClubAssociationId
                    clsMortgageClubNetAssocTable.SetKeyMatchValues colMatchValues
                    clsTargetMortgageClubNetAssocTable.SetKeyMatchValues colMatchValues
                    CopyObject clsMortgageClubNetAssocTable, clsTargetMortgageClubNetAssocTable
                    clsTargetMortgageClubNetAssocTable.Update
                
                End If
                rs.MoveNext
            Loop
            CopyObjectNonPrimaryKey clsFirmClubNetworkAssociation, clsTargetFirmClubNetworkAssociation, rs
        End If
    End If
' TW 05/04/2007 EP2_2292
    If strPackagerIndicator = "1" Then
        ' Next, User etc
        Set clsUser = New OmigaUserTable
        Set colMatchValues = New Collection
        colMatchValues.Add strUserId
        clsUser.SetKeyMatchValues colMatchValues
        CopyUser clsUser, sTargetDatabase
        
        ' Next, Introducer etc
        Set clsIntroducer = New IntroducerTable
        Set colMatchValues = New Collection
        colMatchValues.Add strUserId
        clsIntroducer.SetKeyMatchValues colMatchValues
        CopyIntroducer clsIntroducer, sTargetDatabase
    End If
' TW 05/04/2007 EP2_2292 End
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyUser(clsSrcUser As OmigaUserTable, sTargetDatabase As String)
        
    Dim sUserId As String
    Dim clsTableAccess As TableAccess
    Dim colMatchValues As Collection
    Dim iManagement As ISecurityManager
    Dim clsConnection As SupervisorConnection
    
    On Error GoTo Failed
    
    
    Set clsTableAccess = clsSrcUser
    GetDatabaseObject clsSrcUser
    
    sUserId = clsSrcUser.GetUserID()
    
' TW 20/12/2006 EP2_604
' Check to see if there are any USERHISTORY records on the target database for this user
Dim rs As ADODB.Recordset
Dim sCommand As String
        
    sCommand = "SELECT COUNT(1) FROM USERHISTORY WHERE USERID = '" & sUserId & "'"
    Set rs = g_clsDataAccess.ExecuteCommand(sCommand, m_sTargetDatabase)
    If rs(0) > 0 Then
        MsgBox "User '" & sUserId & "' cannot be promoted. There is USERHISTORY on the target database for this user", vbExclamation
        Exit Sub
    End If
' TW 20/12/2006 EP2_604 End
    
    
    Set colMatchValues = clsTableAccess.GetKeyMatchValues
    
    ' Omiga User
    Dim clsTargetUser As TableAccess
    
    Set clsTargetUser = New OmigaUserTable
    clsTargetUser.SetKeyMatchValues colMatchValues
    clsTargetUser.SetDatabase sTargetDatabase
    
    CopyObject clsSrcUser, clsTargetUser, , False
    clsTargetUser.Update
    
    ' First, user role
    Dim clsSrcUserRole As TableAccess
    Dim clsTargetUserRole As TableAccess
    
    Set clsSrcUserRole = New UserRoleTable
    Set clsTargetUserRole = New UserRoleTable
    
    clsSrcUserRole.SetKeyMatchValues colMatchValues
    clsTargetUserRole.SetKeyMatchValues colMatchValues
    
    clsTargetUserRole.SetDatabase sTargetDatabase
    clsTargetUserRole.DeleteRecords
    
    CopyObject clsSrcUserRole, clsTargetUserRole, True
    clsTargetUserRole.Update
    
    ' Organisation User
    Dim clsSrcOrgUser As TableAccess
    Dim clsTargetOrgUser As TableAccess
    
    Set clsSrcOrgUser = New OrgUserTable
    Set clsTargetOrgUser = New OrgUserTable
    
    clsSrcOrgUser.SetKeyMatchValues colMatchValues
    clsTargetOrgUser.SetKeyMatchValues colMatchValues
    clsTargetOrgUser.SetDatabase sTargetDatabase
    
    CopyObject clsSrcOrgUser, clsTargetOrgUser
    clsTargetOrgUser.Update
    
    ' Password
    Dim clsSrcPassword As TableAccess
    Dim clsTargetPassword As TableAccess
    
    Set clsSrcPassword = New PasswordTable
    Set clsTargetPassword = New PasswordTable
    
    clsSrcPassword.SetKeyMatchValues colMatchValues
    clsTargetPassword.SetKeyMatchValues colMatchValues
    clsTargetPassword.SetDatabase sTargetDatabase
    
    CopyObject clsSrcPassword, clsTargetPassword
    clsTargetPassword.Update
    
    ' User Competency
    Dim clsSrcCompetency As TableAccess
    Dim clsTargetCompetency As TableAccess
    
    Set clsSrcCompetency = New UserCompetencyTable
    Set clsTargetCompetency = New UserCompetencyTable
    
    clsSrcCompetency.SetKeyMatchValues colMatchValues
    clsTargetCompetency.SetKeyMatchValues colMatchValues
    clsTargetCompetency.SetDatabase sTargetDatabase
    clsTargetCompetency.DeleteRecords sTargetDatabase
    
    CopyObject clsSrcCompetency, clsTargetCompetency
    clsTargetCompetency.Update
    
    ' Qualification
    Dim clsSrcQualification As TableAccess
    Dim clsTargetQualification As TableAccess
    
    Set clsSrcQualification = New QualificationTable
    Set clsTargetQualification = New QualificationTable
    
    clsSrcQualification.SetKeyMatchValues colMatchValues
    clsTargetQualification.SetKeyMatchValues colMatchValues
    clsTargetQualification.SetDatabase sTargetDatabase
    
    CopyObject clsSrcQualification, clsTargetQualification
    clsTargetQualification.Update
    
    'If we have a User-to-Contact mapping.
    Dim clsUserContactTable As OrgUserContactTable
    
    Set clsUserContactTable = New OrgUserContactTable
    TableAccess(clsUserContactTable).SetKeyMatchValues colMatchValues
    TableAccess(clsUserContactTable).GetTableData
    
    If TableAccess(clsUserContactTable).RecordCount > 0 Then
        'Build a contact GUID keys collection.
        Dim vContactGUID As Variant
        
        vContactGUID = clsUserContactTable.GetContactDetailsGUID
        Set colMatchValues = New Collection
        colMatchValues.Add vContactGUID
        
        'Contact Details.
        Dim clsSrcContactTable As TableAccess
        Dim clsTrgContactTable As TableAccess
        
        Set clsSrcContactTable = New ContactDetailsTable
        Set clsTrgContactTable = New ContactDetailsTable
        
        clsSrcContactTable.SetKeyMatchValues colMatchValues
        clsTrgContactTable.SetKeyMatchValues colMatchValues
        clsTrgContactTable.SetDatabase sTargetDatabase
        
        CopyObject clsSrcContactTable, clsTrgContactTable
        clsTrgContactTable.Update
    
        'Contact Telephone Details.
        Dim clsSrcTelephoneTable As TableAccess
        Dim clsTrgTelephoneTable As TableAccess
        
        Set clsSrcTelephoneTable = New ContactDetailsTelephoneTable
        Set clsTrgTelephoneTable = New ContactDetailsTelephoneTable
        
        clsSrcTelephoneTable.SetKeyMatchValues colMatchValues
        clsTrgTelephoneTable.SetKeyMatchValues colMatchValues
        clsTrgTelephoneTable.SetDatabase sTargetDatabase
        
        CopyObject clsSrcTelephoneTable, clsTrgTelephoneTable
        clsTrgTelephoneTable.Update
    
        'Set the key match collection back to the UserID.
        Set clsTableAccess = clsSrcUser
        Set colMatchValues = clsTableAccess.GetKeyMatchValues
    
        'User-to-Contact mapping.
        Dim clsSrcUserContactTable As TableAccess
        Dim clsTrgUserContactTable As TableAccess
        
        Set clsSrcUserContactTable = New OrgUserContactTable
        Set clsTrgUserContactTable = New OrgUserContactTable
        
        clsSrcUserContactTable.SetKeyMatchValues colMatchValues
        clsTrgUserContactTable.SetKeyMatchValues colMatchValues
        clsTrgUserContactTable.SetDatabase sTargetDatabase
        
        CopyObject clsSrcUserContactTable, clsTrgUserContactTable
        clsTrgUserContactTable.Update
    End If
    
    'Cast an interface onto the security management component.
    Set iManagement = g_clsSecurityMgr
    
    'Get a reference to the target databases ADO connection.
    g_clsDataAccess.GetSupervisorConnection clsConnection, sTargetDatabase
    
    'Ask the security component to copy the user's rights from one DB to the Other.
    iManagement.PromoteAgent sUserId, AgentTypeUser, clsConnection.GetConnection()
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub CopyUnit(clsSrcUnit As UnitTable, sTargetDatabase As String)
    On Error GoTo Failed
    Dim clsTargetUnit As TableAccess
    Dim colMatchValues As Collection
    Dim clsSrcTableAccess As TableAccess
    Dim clsTargetTableAccess As TableAccess
    Dim sUnitID As String
    Dim vAddressGUID As Variant
    
    ' First copy the UNIT table itself
    Set clsTargetUnit = New UnitTable
    Set clsSrcTableAccess = clsSrcUnit
    Set colMatchValues = clsSrcTableAccess.GetKeyMatchValues()
    clsTargetUnit.SetKeyMatchValues colMatchValues
    clsTargetUnit.SetDatabase sTargetDatabase
    
    GetDatabaseObject clsSrcUnit
    
    sUnitID = clsSrcUnit.GetUnitId()
    vAddressGUID = clsSrcUnit.GetAddressGUID()
    
    CopyObject clsSrcUnit, clsTargetUnit
    
    ' Can't update just yet as we need address, for some reason.
    CopyAddressTable vAddressGUID, sTargetDatabase
    
    ' Can now though...
    clsTargetUnit.Update
    
    ' Contact details
    Dim clsSrcUnitContactDetails As UnitContactDetailsTable
    Dim clsTargetUnitContactDetails As UnitContactDetailsTable
    Dim vContactDetailsGUID As Variant
    
    Set clsSrcUnitContactDetails = New UnitContactDetailsTable
    Set clsSrcTableAccess = clsSrcUnitContactDetails
    
    Set clsTargetUnitContactDetails = New UnitContactDetailsTable
    Set clsTargetTableAccess = clsTargetUnitContactDetails
    
    clsSrcTableAccess.SetKeyMatchValues colMatchValues
    clsTargetTableAccess.SetKeyMatchValues colMatchValues
    
    ' Get the source Unit Contact Details table data
    GetDatabaseObject clsSrcUnitContactDetails
    
    ' Get the GUID
    vContactDetailsGUID = clsSrcUnitContactDetails.GetContactDetailsGUID()

    Set clsTargetTableAccess = clsTargetUnitContactDetails
    clsTargetTableAccess.DeleteRecords sTargetDatabase
    
    
    ' Delete the target Unit Contact Details table data
'    clsTargetTableAccess.DeleteRecords sTargetDatabase
    CopyObject clsSrcUnitContactDetails, clsTargetUnitContactDetails, True, False

    CopyContactDetailsTable vContactDetailsGUID, sTargetDatabase
    
    clsTargetTableAccess.Update
    
'    ' Now the region table
'    Dim clsSrcRegion As RegionTable
'    Dim clsTargetRegion As RegionTable
'
'    Set clsSrcRegion = New RegionTable
'    Set clsTargetRegion = New RegionTable
'
'    Set clsTargetTableAccess = clsTargetRegion
'    clsTargetTableAccess.SetDatabase sTargetDatabase
'    clsTargetRegion.GetRegionsFromUnit sUnitID
'
'    clsSrcRegion.GetRegionsFromUnit sUnitID
    
    
    ' Need the Unit Region table
    Dim clsSrcUnitRegion As UnitRegionTable
    Dim clsTargetUnitRegion As UnitRegionTable
    
    Set clsSrcUnitRegion = New UnitRegionTable
    Set clsTargetUnitRegion = New UnitRegionTable
    
    Set clsSrcTableAccess = clsSrcUnitRegion
    clsSrcTableAccess.SetKeyMatchValues colMatchValues
    
    Set clsTargetTableAccess = clsTargetUnitRegion
    clsTargetTableAccess.SetKeyMatchValues colMatchValues
    clsTargetTableAccess.SetDatabase sTargetDatabase
    
    ' Delete target database unit regions
    clsTargetTableAccess.DeleteRecords sTargetDatabase
    
'    Set clsTargetTableAccess = clsTargetRegion
'    clsTargetTableAccess.SetDatabase sTargetDatabase
'    clsTargetTableAccess.DeleteRecords sTargetDatabase
    'clsTargetTableAccess.DeleteAllRows
    'clsTargetTableAccess.Update
    
'    ' Copy the Region table
'    CopyObject clsSrcRegion, clsTargetRegion, True, False
'    Set clsTargetTableAccess = clsTargetRegion
'    clsTargetTableAccess.Update
    
    ' Copy Unit Regions
    CopyObject clsSrcUnitRegion, clsTargetUnitRegion, True
    Set clsTargetTableAccess = clsTargetUnitRegion
    clsTargetTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyContactDetailsTable(vContactDetailsGUID As Variant, sTargetDatabase As String)
    
    Dim sFunctionName As String
    Dim clsSrcTelephoneTable As TableAccess
    Dim clsDstTelephoneTable As TableAccess
    Dim clsSrcContactDetails As TableAccess
    Dim colContactDetailsValues As Collection
    Dim clsTargetContactDetails As TableAccess
    
    On Error GoTo Failed
    
    sFunctionName = "CopyContactDetailsTable"
    
    'Only try and copy data if a contact GUID is specified.
    If Not IsNull(vContactDetailsGUID) Then
        If Len(vContactDetailsGUID) > 0 Then
            'Create the source, destination tables and a keys collection.
            Set colContactDetailsValues = New Collection
            Set clsSrcContactDetails = New ContactDetailsTable
            Set clsTargetContactDetails = New ContactDetailsTable
            
            'Add the primary key into the collection.
            colContactDetailsValues.Add vContactDetailsGUID
            
            'Assign the same keys collection to both source and destination tables.
            clsSrcContactDetails.SetKeyMatchValues colContactDetailsValues
            clsTargetContactDetails.SetKeyMatchValues colContactDetailsValues
            
            'Copy the source record to the target database.
            CopyObject clsSrcContactDetails, clsTargetContactDetails
            
            'Update the target table(s).
            clsTargetContactDetails.Update
            
            
            'Create the source and destination telephone tables.
            Set clsSrcTelephoneTable = New ContactDetailsTelephoneTable
            Set clsDstTelephoneTable = New ContactDetailsTelephoneTable
            
            'Assign the same keys collection to the telephone source and destination tables.
            clsSrcTelephoneTable.SetKeyMatchValues colContactDetailsValues
            clsDstTelephoneTable.SetKeyMatchValues colContactDetailsValues
            
            'Point the target table to the target database.
            clsTargetContactDetails.SetDatabase sTargetDatabase
            clsDstTelephoneTable.SetDatabase sTargetDatabase
                        
            'Copy the telephone number(s) for this contact.
            CopyObject clsSrcTelephoneTable, clsDstTelephoneTable
            
            'Update the target table(s).
            clsDstTelephoneTable.Update
        Else
            g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": ContactDetails GUID is empty"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": ContactDetails GUID is empty"
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyAddressTable(vAddressGUID As Variant, sTargetDatabase As String)
    On Error GoTo Failed
    Dim clsSrcAddress As TableAccess
    Dim clsTargetAddress As TableAccess
    Dim colAddressValues As Collection
    Dim sFunctionName As String
    
    sFunctionName = "CopyAddressTable"
    
    If Not IsNull(vAddressGUID) Then
        If Len(vAddressGUID) > 0 Then
            Set colAddressValues = New Collection
            
            Set clsSrcAddress = New AddressTable
            Set clsTargetAddress = New AddressTable
            colAddressValues.Add vAddressGUID
            
            clsSrcAddress.SetKeyMatchValues colAddressValues
            clsTargetAddress.SetKeyMatchValues colAddressValues
            clsTargetAddress.SetDatabase sTargetDatabase
            CopyObject clsSrcAddress, clsTargetAddress
            
            clsTargetAddress.Update
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub

Private Sub CopyDepartment(clsSrcDepartment As DepartmentTable, sTargetDatabase As String)
    On Error GoTo Failed
    Dim vAddressGUID As Variant
    Dim clsTargetDepartment As TableAccess
    Dim clsTableAccess As TableAccess
    Dim colMatchValues As Collection
    
    Set clsTargetDepartment = New DepartmentTable
    GetDatabaseObject clsSrcDepartment
    Set clsTableAccess = clsSrcDepartment
    Set colMatchValues = clsTableAccess.GetKeyMatchValues
    
    clsTargetDepartment.SetKeyMatchValues colMatchValues
    clsTargetDepartment.SetDatabase sTargetDatabase
    
    ' First, get the address guid
    vAddressGUID = clsSrcDepartment.GetAddressGUID()
    
    ' Copy the department table
    CopyObject clsSrcDepartment, clsTargetDepartment, False, False
    

    ' Copy the Address table
    CopyAddressTable vAddressGUID, sTargetDatabase
    
    clsTargetDepartment.Update
    
    ' Need the get the deparment contact details
    Dim clsSrcDeptContact As DepartmentContactTable
    Dim clsTargetDeptContact As TableAccess
    Dim vContactDetailsGUID As Variant
    
    Set clsSrcDeptContact = New DepartmentContactTable
    Set clsTargetDeptContact = New DepartmentContactTable
    
    Set clsTableAccess = clsSrcDepartment
    Set colMatchValues = clsTableAccess.GetKeyMatchValues

    Set clsTableAccess = clsSrcDeptContact
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    clsTargetDeptContact.SetKeyMatchValues colMatchValues
    
    GetDatabaseObject clsSrcDeptContact
    vContactDetailsGUID = clsSrcDeptContact.GetContactDetailsGUID()
    clsTargetDeptContact.SetDatabase sTargetDatabase
    
    CopyObject clsSrcDeptContact, clsTargetDeptContact
    
    CopyContactDetailsTable vContactDetailsGUID, sTargetDatabase
    clsTargetDeptContact.Update

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyComboGroup(clsSrcCombo As TableAccess, sTargetDatabase As String)
    
    On Error GoTo Failed
    Dim clsSrcComboValue As ComboValueTable
    Dim clsSrcComboValidation As ComboValidationTable
    Dim clsDestComboValidation As ComboValidationTable
    Dim clsDestComboValue As ComboValueTable
    Dim clsDestComboGroup As ComboValueGroupTable
    Dim clsDestCombo As TableAccess
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    Dim colMatchKeys As Collection
    Set colMatchValues = clsSrcCombo.GetKeyMatchValues()
    
    Set colMatchKeys = clsSrcCombo.GetKeyMatchFields()
    Set clsSrcComboValue = New ComboValueTable
    Set clsTableAccess = clsSrcComboValue
    clsTableAccess.SetDatabase m_sSourceDatabase
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    ' Just want to search on the one item, the group name
    clsTableAccess.SetKeyMatchFields colMatchKeys
    
    Set clsSrcComboValidation = New ComboValidationTable
    Set clsTableAccess = clsSrcComboValidation
    clsTableAccess.SetDatabase m_sSourceDatabase
    clsTableAccess.SetKeyMatchValues colMatchValues
    ' Just want to search on the one item, the group name
    clsTableAccess.SetKeyMatchFields colMatchKeys

    Set clsDestComboValidation = New ComboValidationTable
    Set clsDestComboValue = New ComboValueTable
    Set clsDestComboGroup = New ComboValueGroupTable
    
    ' First, delete the current record, if it's there:
    Set clsDestCombo = clsDestComboGroup
    
    clsDestCombo.SetDatabase sTargetDatabase
    clsDestCombo.SetKeyMatchValues colMatchValues
    clsDestComboGroup.DeleteComboGroup sTargetDatabase
    
    ' Then add the new one
    CopyObject clsSrcCombo, clsDestComboGroup, True
    CopyObject clsSrcComboValue, clsDestComboValue, True
    CopyObject clsSrcComboValidation, clsDestComboValidation, True
    
    CreateComboValueAuditRecords clsSrcComboValue, sTargetDatabase
    CreateComboValidationAuditRecords clsSrcComboValidation, sTargetDatabase
    
    Set clsDestCombo = clsDestComboGroup
    clsDestCombo.Update
    
    Set clsDestCombo = clsDestComboValue
    clsDestCombo.Update

    Set clsDestCombo = clsDestComboValidation
    clsDestCombo.Update

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub CopyProductTable(clsSrcProduct As TableAccess, clsTargetProduct As TableAccess, sTargetDatabase As String, Optional bDeleteFirst As Boolean = False)
    On Error GoTo Failed
    Dim colKeyMatchValues As Collection
    Dim colKeyMatchFields As Collection
    Dim sFunctionName As String
    
    sFunctionName = "CopyProductTable"
    
    Set colKeyMatchValues = clsSrcProduct.GetKeyMatchValues()
    Set colKeyMatchFields = clsSrcProduct.GetKeyMatchFields()
    
    If Not colKeyMatchValues Is Nothing Then
        If colKeyMatchValues.Count = colKeyMatchFields.Count Then
            
            clsTargetProduct.SetKeyMatchValues colKeyMatchValues
            clsTargetProduct.SetDatabase sTargetDatabase
            
            If bDeleteFirst Then
                clsTargetProduct.DeleteRecords sTargetDatabase
            End If
            
            CopyObject clsSrcProduct, clsTargetProduct, bDeleteFirst
            clsTargetProduct.Update
        Else
            g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": No key values"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": No key values"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub CopyProduct(clsSrcProduct As TableAccess, sTargetDatabase As String)

    On Error GoTo Failed
    Dim colKeyMatchValues As Collection
    
    Set colKeyMatchValues = clsSrcProduct.GetKeyMatchValues()
    
    ' Product table
    Dim clsTargetMortProd As TableAccess
    Set clsTargetMortProd = New MortgageProductTable
    
    CopyProductTable clsSrcProduct, clsTargetMortProd, sTargetDatabase
    
    ' Mortgage product parameters table
    Dim clsSrcProdParams As TableAccess
    Dim clsTargetProdParams As TableAccess
    
    Set clsSrcProdParams = New MortProdParamsTable
    Set clsTargetProdParams = New MortProdParamsTable
    clsSrcProdParams.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcProdParams, clsTargetProdParams, sTargetDatabase
        
    
    ' Channel eligibility
    Dim clsSrcChannelElig As TableAccess
    Dim clsTargetChannelElig As TableAccess
    
    Set clsSrcChannelElig = New MortProdChannelEligTable
    Set clsTargetChannelElig = New MortProdChannelEligTable
    clsSrcChannelElig.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcChannelElig, clsTargetChannelElig, sTargetDatabase, True
    
    ' Property Location
    Dim clsSrcPropLocation As TableAccess
    Dim clsTargetPropLocation As TableAccess
    
    Set clsSrcPropLocation = New MortProdPropLocTable
    Set clsTargetPropLocation = New MortProdPropLocTable
    clsSrcPropLocation.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcPropLocation, clsTargetPropLocation, sTargetDatabase, True
    
    
    ' Type of buyer
    Dim clsSrcTypeOfBuyer As TableAccess
    Dim clsTargetTypeOfBuyer As TableAccess
    
    Set clsSrcTypeOfBuyer = New MortProdTypeOfBuyerTable
    Set clsTargetTypeOfBuyer = New MortProdTypeOfBuyerTable
    clsSrcTypeOfBuyer.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcTypeOfBuyer, clsTargetTypeOfBuyer, sTargetDatabase, True
        
    ' Purpose of loan
    Dim clsSrcPurpOfLoan As TableAccess
    Dim clsTargetPurpOfLoan As TableAccess
    
    Set clsSrcPurpOfLoan = New MortProdPurpOfLoanTable
    Set clsTargetPurpOfLoan = New MortProdPurpOfLoanTable
    clsSrcPurpOfLoan.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcPurpOfLoan, clsTargetPurpOfLoan, sTargetDatabase, True
    
    ' Mortgage product language
    Dim clsSrcProductLanguage As TableAccess
    Dim clsTargetProductLanguage As TableAccess
    
    Set clsSrcProductLanguage = New MortProdLanguageTable
    Set clsTargetProductLanguage = New MortProdLanguageTable
    clsSrcProductLanguage.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcProductLanguage, clsTargetProductLanguage, sTargetDatabase, True
        
    ' BMIDS959  Remove Mortgage product bands
       
    ' Other fees
    Dim clsSrcOtherFees As TableAccess
    Dim clsTargetOtherFees As TableAccess
    
    Set clsSrcOtherFees = New OtherFeeTable
    Set clsTargetOtherFees = New OtherFeeTable
    clsSrcOtherFees.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcOtherFees, clsTargetOtherFees, sTargetDatabase, True
    
    ' Typical APR's
    Dim clsSrcTypicalAPR As TableAccess
    Dim clsTargetTypicalAPR As TableAccess
    
    Set clsSrcTypicalAPR = New MortProdTypicalAPRTable
    Set clsTargetTypicalAPR = New MortProdTypicalAPRTable
    clsSrcTypicalAPR.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcTypicalAPR, clsTargetTypicalAPR, sTargetDatabase, True
    
    ' Type of application eligibility
    Dim clsSrcAppElig As TableAccess
    Dim clsTargetAppElig As TableAccess
    
    Set clsSrcAppElig = New MortProdTypeofAppEligTable
    Set clsTargetAppElig = New MortProdTypeofAppEligTable
    clsSrcAppElig.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcAppElig, clsTargetAppElig, sTargetDatabase, True
    
    ' Employment eligibility
    Dim clsSrcEmpElig As TableAccess
    Dim clsTargetEmpElig As TableAccess
    
    Set clsSrcEmpElig = New MortProdEmpEligTable
    Set clsTargetEmpElig = New MortProdEmpEligTable
    clsSrcEmpElig.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcEmpElig, clsTargetEmpElig, sTargetDatabase, True
    
    ' Special Group
    Dim clsSrcSpecialGroup As TableAccess
    Dim clsTargetSpecialGroup As TableAccess
    
    Set clsSrcSpecialGroup = New MortProdSpecialGroupTable
    Set clsTargetSpecialGroup = New MortProdSpecialGroupTable
    clsSrcSpecialGroup.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcSpecialGroup, clsTargetSpecialGroup, sTargetDatabase, True
    
    ' JD BMIDS775 additionalFeatures
    Dim clsSrcAddFea As TableAccess
    Dim clsTargetAddFea As TableAccess
    
    Set clsSrcAddFea = New MortProdAddtFeaTable
    Set clsTargetAddFea = New MortProdAddtFeaTable
    clsSrcAddFea.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcAddFea, clsTargetAddFea, sTargetDatabase, True

    
'GD BMIDS BMIDS00016 ; promotions start
    'MORTGAGEPRODUCTCONDITION, CONDITIONS table.
    Dim clsSrcCond As TableAccess
    Dim clsTargetCond As TableAccess
    Set clsSrcCond = New MortProdProdCondTable
    Set clsTargetCond = New MortProdProdCondTable
    clsSrcCond.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcCond, clsTargetCond, sTargetDatabase, True


'GD BMIDS BMIDS00016 ; promotions end

' TW 30/11/2006 EP2_253
    Dim clsSrcMortProdIncomeStatus As TableAccess
    Dim clsTargetMortProdIncomeStatus As TableAccess
    Set clsSrcMortProdIncomeStatus = New MortProdIncomeStatusTable
    Set clsTargetMortProdIncomeStatus = New MortProdIncomeStatusTable
    
    clsSrcMortProdIncomeStatus.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcMortProdIncomeStatus, clsTargetMortProdIncomeStatus, sTargetDatabase, True

    Dim clsSrcMortProdNatureOfLoan As TableAccess
    Dim clsTargetMortProdNatureOfLoan As TableAccess

    Set clsSrcMortProdNatureOfLoan = New MortProdNatureOfLoanTable
    Set clsTargetMortProdNatureOfLoan = New MortProdNatureOfLoanTable
    
    clsSrcMortProdNatureOfLoan.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcMortProdNatureOfLoan, clsTargetMortProdNatureOfLoan, sTargetDatabase, True

    Dim clsSrcMortProdProdClass As TableAccess
    Dim clsTargetMortProdProdClass As TableAccess
    Set clsSrcMortProdProdClass = New MortProdProdClassTable
    Set clsTargetMortProdProdClass = New MortProdProdClassTable
    
    clsSrcMortProdProdClass.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcMortProdProdClass, clsTargetMortProdProdClass, sTargetDatabase, True
' TW 30/11/2006 EP2_253 End

    
    ' Interest Rates
    Dim clsSrcIntRates As TableAccess
    Dim clsTargetIntRates As TableAccess
    Dim colKeyMatchFields As Collection
    
    Set clsSrcIntRates = New MortProdIntRatesTable
    Set clsTargetIntRates = New MortProdIntRatesTable
    Set colKeyMatchFields = clsSrcProduct.GetKeyMatchFields()
    clsSrcIntRates.SetKeyMatchFields colKeyMatchFields
    clsTargetIntRates.SetKeyMatchFields colKeyMatchFields
    
    clsSrcIntRates.SetKeyMatchValues colKeyMatchValues
    
    CopyProductTable clsSrcIntRates, clsTargetIntRates, sTargetDatabase, True
    
    ' Inclusive Incentives
    CopyIncentive colKeyMatchValues, sTargetDatabase

'    ' DJP SQL Server port - missing Product Grouping so added
'    ' Product Grouping
'    Dim clsSrcProductGrouping As TableAccess
'    Dim clsTargetProductGrouping As TableAccess
'
'    Set clsSrcProductGrouping = New ProductGroupingTable
'    Set clsTargetProductGrouping = New ProductGroupingTable
'    clsSrcProductGrouping.SetKeyMatchValues colKeyMatchValues
'
'    CopyProductTable clsSrcProductGrouping, clsTargetProductGrouping, sTargetDatabase, True
 
    Dim clsMortProd As MortgageProductTable
    Set clsMortProd = clsSrcProduct
    
    ' BM0423
    clsMortProd.ValidateProduct sTargetDatabase
 
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



Private Sub CopyIncentive(colKeyMatchValues As Collection, sTargetDatabase As String)
'GD BMIDS00016
    On Error GoTo Failed
    Dim clsSrcInclusive As MortProdIncIncentiveTable
    Dim clsTargetInclusive As MortProdIncIncentiveTable
    Dim clsSrcExclusive As MortProdIncIncentiveTable
    Dim clsTargetExclusive As MortProdIncIncentiveTable
    Dim clsTargetIncIncentive As MortProdIncentiveTable
    Dim clsTableAccess As TableAccess

    Set clsSrcInclusive = New MortProdIncIncentiveTable
    Set clsTargetInclusive = New MortProdIncIncentiveTable
    Set clsSrcExclusive = New MortProdIncIncentiveTable
    Set clsTargetExclusive = New MortProdIncIncentiveTable
    Set clsTargetIncIncentive = New MortProdIncentiveTable
    
    Set clsTableAccess = clsTargetIncIncentive
    clsTableAccess.SetKeyMatchValues colKeyMatchValues
    
    clsTargetIncIncentive.DeleteIncentives sTargetDatabase
    clsSrcInclusive.SetType Inclusive
    clsTargetInclusive.SetType Inclusive
    
    clsSrcExclusive.SetType Exclusive
    clsTargetExclusive.SetType Exclusive
    
    Set clsTableAccess = clsSrcExclusive
    clsTableAccess.SetKeyMatchValues colKeyMatchValues
    
    Set clsTableAccess = clsSrcInclusive
    clsTableAccess.SetKeyMatchValues colKeyMatchValues
    
    ' Copy inclusive and exclusive tables
    CopyObject clsSrcInclusive, clsTargetInclusive, True
    CopyObject clsSrcExclusive, clsTargetExclusive, True
    
    ' Get Incentive table records for inclusive and exclusive
    Dim clsSrcIncIncentive As MortProdIncentiveTable

    Dim clsSrcExIncentive As MortProdIncentiveTable
    Dim clsTargetExIncentive As MortProdIncentiveTable
    
    Set clsSrcIncIncentive = New MortProdIncentiveTable
    Set clsSrcExIncentive = New MortProdIncentiveTable
    Set clsTargetExIncentive = New MortProdIncentiveTable

    ' Target Inclusive incentives
    Set clsTableAccess = clsTargetIncIncentive
    clsTableAccess.SetDatabase sTargetDatabase
    
    'GD BMIDS00016 clsTargetIncIncentive.GetIncentives Inclusive, colKeyMatchValues, True
    clsTargetIncIncentive.GetIncentivesForPromotions Inclusive, colKeyMatchValues, True
    
    ' Source Inclusive incentives
    'GD BMIDS00016 clsSrcIncIncentive.GetIncentives Inclusive, colKeyMatchValues, True
    clsSrcIncIncentive.GetIncentivesForPromotions Inclusive, colKeyMatchValues, True
    CopyObject clsSrcIncIncentive, clsTargetIncIncentive, True, False
    
    ' Target Exclusive incentives
    Set clsTableAccess = clsTargetExIncentive
    clsTableAccess.SetDatabase sTargetDatabase
    'GD BMIDS00016 clsTargetExIncentive.GetIncentives Exclusive, colKeyMatchValues, True
    clsTargetExIncentive.GetIncentivesForPromotions Exclusive, colKeyMatchValues, True
    
    ' Source Exclusive incentives
    'GD BMIDS00016 clsSrcExIncentive.GetIncentives Exclusive, colKeyMatchValues, True
    clsSrcExIncentive.GetIncentivesForPromotions Exclusive, colKeyMatchValues, True
    
    
    CopyObject clsSrcExIncentive, clsTargetExIncentive, True, False
    
    Set clsTableAccess = clsTargetIncIncentive
    clsTableAccess.Update

    Set clsTableAccess = clsTargetExIncentive
    clsTableAccess.Update
    
    Set clsTableAccess = clsTargetInclusive
    clsTableAccess.Update
    
    Set clsTableAccess = clsTargetExclusive
    clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Function GetLenderRateset(clsSrcLender As MortgageLendersTable)
    On Error GoTo Failed
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = clsSrcLender
    Set rs = clsTableAccess.GetTableData()
    
    If clsTableAccess.RecordCount() = 1 Then
        GetLenderRateset = clsSrcLender.GetRateSet()
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub CopyDirectory(ByVal vDirectoryGUID As Variant, ByVal sTargetDatabase As String)
    
    Dim vAddressGUID As Variant
    Dim vContactGUID As Variant
    Dim clsSrcTableAccess As TableAccess
    Dim colDirectoryKeyValues As Collection
    Dim clsTargetTableAccess As TableAccess
    Dim clsSrcDirectory As NameAndAddressDirTable
    Dim clsTargetDirectory As NameAndAddressDirTable
    Dim clsTargetDirectoryTableAccess As TableAccess

    On Error GoTo Failed

    If Len(vDirectoryGUID) > 0 Then
        ' Next, the NameAndAddressDirectory
        Set colDirectoryKeyValues = New Collection
        colDirectoryKeyValues.Add vDirectoryGUID

        Set clsSrcDirectory = New NameAndAddressDirTable
        Set clsSrcTableAccess = clsSrcDirectory
        clsSrcTableAccess.SetKeyMatchValues colDirectoryKeyValues

        Set clsTargetDirectory = New NameAndAddressDirTable
        Set clsTargetDirectoryTableAccess = clsTargetDirectory
        clsTargetDirectoryTableAccess.SetKeyMatchValues colDirectoryKeyValues
        clsTargetDirectoryTableAccess.SetDatabase sTargetDatabase

        GetDatabaseObject clsSrcDirectory
        'Set rs = clsSrcTableAccess.GetTableData()

        If clsSrcTableAccess.RecordCount() > 0 Then
            vAddressGUID = clsSrcDirectory.GetAddressGUID()
            vContactGUID = clsSrcDirectory.GetContactDetailsGUID()

            CopyObject clsSrcDirectory, clsTargetDirectory, , False
        End If

        If Not IsNull(vAddressGUID) Then
            If Len(vAddressGUID) > 0 Then
                ' Now just need to do address and contact details
                Dim clsSrcAddress As AddressTable
                Dim clsTargetAddress As AddressTable
                Dim colAddressValues As Collection
    
                Set colAddressValues = New Collection
                colAddressValues.Add vAddressGUID
                Set clsSrcAddress = New AddressTable
                Set clsSrcTableAccess = clsSrcAddress
                clsSrcTableAccess.SetKeyMatchValues colAddressValues
    
                Set clsTargetAddress = New AddressTable
                Set clsTargetTableAccess = clsTargetAddress
                clsTargetTableAccess.SetKeyMatchValues colAddressValues
                clsTargetTableAccess.SetDatabase sTargetDatabase
    
                CopyObject clsSrcAddress, clsTargetAddress
                clsTargetTableAccess.Update
            End If
        End If
        
        If Not IsNull(vContactGUID) Then
            If Len(vContactGUID) > 0 Then
                'Copy the contact details and telephone records.
                CopyContactDetailsTable vContactGUID, sTargetDatabase
            End If
        End If
        
        clsTargetDirectoryTableAccess.Update
        
        ' Panel
        Dim clsSrcPanel As TableAccess
        Dim clsTargetPanel As TableAccess
        
        Set clsSrcPanel = New PanelTable
        Set clsTargetPanel = New PanelTable
        
        clsSrcPanel.SetKeyMatchValues colDirectoryKeyValues
        clsTargetPanel.SetKeyMatchValues colDirectoryKeyValues
        clsTargetPanel.SetDatabase sTargetDatabase
            
        CopyObject clsSrcPanel, clsTargetPanel
        clsTargetPanel.Update
        
        ' Panel Valuer
        Dim clsSrcPanelValuer As TableAccess
        Dim clsTargetPanelValuer As TableAccess
        
        Set clsSrcPanelValuer = New PanelValuerTable
        Set clsTargetPanelValuer = New PanelValuerTable
        
        clsSrcPanelValuer.SetKeyMatchValues colDirectoryKeyValues
        clsTargetPanelValuer.SetKeyMatchValues colDirectoryKeyValues
        clsTargetPanelValuer.SetDatabase sTargetDatabase
            
        CopyObject clsSrcPanelValuer, clsTargetPanelValuer
        clsTargetPanelValuer.Update

        ' Panel Valuer Type
        Dim clsSrcPanelValuerType As TableAccess
        Dim clsTargetPanelValuerType As TableAccess
        
        Set clsSrcPanelValuerType = New PanelValuerTypeTable
        Set clsTargetPanelValuerType = New PanelValuerTypeTable
        
        clsSrcPanelValuerType.SetKeyMatchValues colDirectoryKeyValues
        clsTargetPanelValuerType.SetKeyMatchValues colDirectoryKeyValues
        clsTargetPanelValuerType.SetDatabase sTargetDatabase
            
        CopyObject clsSrcPanelValuerType, clsTargetPanelValuerType
        clsTargetPanelValuerType.Update
        
        ' Panel Legal Rep
        Dim clsSrcPanelLegalRep As TableAccess
        Dim clsTargetPanelLegalRep As TableAccess
        
        Set clsSrcPanelLegalRep = New PanelLegalRepTable
        Set clsTargetPanelLegalRep = New PanelLegalRepTable
        
        clsSrcPanelLegalRep.SetKeyMatchValues colDirectoryKeyValues
        clsTargetPanelLegalRep.SetKeyMatchValues colDirectoryKeyValues
        clsTargetPanelLegalRep.SetDatabase sTargetDatabase
            
        CopyObject clsSrcPanelLegalRep, clsTargetPanelLegalRep
        clsTargetPanelLegalRep.Update
        
        ' Panel Bank Account
        Dim clsSrcPanelBank As TableAccess
        Dim clsTargetPanelBank As TableAccess
        
        Set clsSrcPanelBank = New PanelBankAccountTable
        Set clsTargetPanelBank = New PanelBankAccountTable
        
        clsSrcPanelBank.SetKeyMatchValues colDirectoryKeyValues
        clsTargetPanelBank.SetKeyMatchValues colDirectoryKeyValues
        clsTargetPanelBank.SetDatabase sTargetDatabase
        clsTargetPanelBank.DeleteRecords sTargetDatabase
        CopyObject clsSrcPanelBank, clsTargetPanelBank, True
        clsTargetPanelBank.Update
        
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyLenderDirectory(colMatchValues As Collection, sTargetDatabase As String)
    On Error GoTo Failed
    Dim clsSrcLenderDirectory As MortLenderDirTable
    Dim clsTargetLenderDirectory As MortLenderDirTable
    Dim clsTargetLender As TableAccess
    Dim clsSrcTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim vDirectoryGUID  As Variant
    
    ' First, copy the Lender Directory
    Set clsSrcLenderDirectory = New MortLenderDirTable
    Set clsSrcTableAccess = clsSrcLenderDirectory
    clsSrcTableAccess.SetKeyMatchValues colMatchValues
    
    Set rs = clsSrcTableAccess.GetTableData()
    
    If clsSrcTableAccess.RecordCount > 0 Then
        vDirectoryGUID = clsSrcLenderDirectory.GetDirectoryGUID()
        
        If Not IsNull(vDirectoryGUID) Then
            If Len(vDirectoryGUID) > 0 Then
                CopyDirectory vDirectoryGUID, sTargetDatabase
            End If
        End If
        
        Set clsTargetLenderDirectory = New MortLenderDirTable
        Set clsTargetLender = clsTargetLenderDirectory
        clsTargetLender.SetKeyMatchValues colMatchValues
        clsTargetLender.SetDatabase sTargetDatabase
    
        CopyObject clsSrcLenderDirectory, clsTargetLenderDirectory
        TableAccess(clsTargetLenderDirectory).Update
        clsSrcTableAccess.CloseRecordSet
    End If
    
'    If Not IsNull(vDirectoryGUID) Then
'        If Len(vDirectoryGUID) > 0 Then
'            CopyDirectory vDirectoryGUID, sTargetDatabase
'        End If
'    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyMIGRates(sRateSet As String, sTargetDatabase As String)
    On Error GoTo Failed
    Dim clsSrcMIGRateSet As MIGRateSetsTable
    Dim clsTargetMIGRateSet As MIGRateSetsTable
    Dim clsSrcTableAccess As TableAccess
    Dim clsTargetTableAccess As TableAccess
    Dim colMatchValues As Collection
    
    ' First, copy the MIGRateSet table
    Set colMatchValues = New Collection
    colMatchValues.Add sRateSet
    
    Set clsSrcMIGRateSet = New MIGRateSetsTable
    Set clsSrcTableAccess = clsSrcMIGRateSet
    clsSrcTableAccess.SetKeyMatchValues colMatchValues
    
    Set clsTargetMIGRateSet = New MIGRateSetsTable
    Set clsTargetTableAccess = clsTargetMIGRateSet
    clsTargetTableAccess.SetKeyMatchValues colMatchValues
    clsTargetTableAccess.SetDatabase sTargetDatabase

    CopyObject clsSrcMIGRateSet, clsTargetMIGRateSet
    
    ' Update the target MIG Rate Set table
    clsTargetTableAccess.Update
    
    ' Now the MIG Rate table itself
    Dim clsSrcMIGRates As MIGRateTable
    Dim clsTargetMIGRates As MIGRateTable
    
    Set clsSrcMIGRates = New MIGRateTable
    Set clsTargetMIGRates = New MIGRateTable
    Set clsTargetTableAccess = clsTargetMIGRates
    clsTargetTableAccess.SetDatabase sTargetDatabase
    
    ' Delete what's there first
    clsTargetMIGRates.DeleteMIGRates sRateSet, sTargetDatabase
    
    ' Make sure we search only on rateset
    clsSrcMIGRates.SetAllRatesFromRateSet sRateSet
        
    ' Do the copy
    CopyObject clsSrcMIGRates, clsTargetMIGRates, True
    
    ' Update the target MIG Rate table
    clsTargetTableAccess.Update
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyLegalFees(colMatchFields As Collection, _
                          colMatchValues As Collection, _
                          sTargetDatabase As String)
    On Error GoTo Failed
    Dim clsSrcTableAccess As TableAccess
    Dim clsTargetTableAccess As TableAccess
    Dim clsSrcLegalFeeSet As LegalFeeSetTable
    Dim clsTargetLegalFeeSet As LegalFeeSetTable
    Dim rs As ADODB.Recordset
    Dim sFeeSet As String

    Set clsSrcLegalFeeSet = New LegalFeeSetTable
    Set clsSrcTableAccess = clsSrcLegalFeeSet

    Set clsTargetLegalFeeSet = New LegalFeeSetTable
    Set clsTargetTableAccess = clsTargetLegalFeeSet
    clsTargetTableAccess.SetDatabase sTargetDatabase
    clsTargetTableAccess.SetKeyMatchFields colMatchFields
    clsTargetTableAccess.SetKeyMatchValues colMatchValues
    
    ' Set the source up with the key to locate the record
    clsSrcTableAccess.SetKeyMatchFields colMatchFields
    clsSrcTableAccess.SetKeyMatchValues colMatchValues

    Set rs = clsSrcTableAccess.GetTableData()

    If clsSrcTableAccess.RecordCount() > 0 Then
        sFeeSet = clsSrcLegalFeeSet.GetFeeSet()
    End If

    If Len(sFeeSet) > 0 Then
        ' We have a feeset, so copy it
        CopyObject clsSrcLegalFeeSet, clsTargetLegalFeeSet
        
        ' Update the target legal fee set table
        clsTargetTableAccess.Update
        
        ' Now the legal fee table itself
        Dim clsSrcLegalFee As LegalFeesTable
        Dim clsTargetLegalFee As LegalFeesTable

        Set clsSrcLegalFee = New LegalFeesTable
        Set clsSrcTableAccess = clsSrcLegalFee
        
        Set clsTargetLegalFee = New LegalFeesTable
        Set clsTargetTableAccess = clsTargetLegalFee
    
        clsSrcTableAccess.SetKeyMatchFields colMatchFields
        clsSrcTableAccess.SetKeyMatchValues colMatchValues
        
        clsTargetTableAccess.SetDatabase sTargetDatabase
        
        ' Delete all the legal fees that match these from the destination database
        clsTargetLegalFee.DeleteLegalFees sFeeSet, sTargetDatabase
        
        clsSrcLegalFee.SetAllFeesFromFeeSet sFeeSet
        ' Having just delete our stuff, don't bother checking it to see if it's there,
        ' just create a new record
        CopyObject clsSrcLegalFee, clsTargetLegalFee, True

       ' Update the target legal fee table
        clsTargetTableAccess.Update
 
    End If
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyLender(clsSrcLender As MortgageLendersTable, sTargetDatabase As String)
    On Error GoTo Failed
    Dim clsTargetLender As MortgageLendersTable
    Dim clsTableAccess As TableAccess
    Dim colMatchValues As Collection
    Dim sRateSet As String
    Dim vOrganisationID As Variant
    Dim rs As ADODB.Recordset
    Dim sFunctionName As String
    Dim colKeys As Collection
    
    sFunctionName = "CopyLender"
    
    Set clsTableAccess = clsSrcLender
    Set colMatchValues = clsTableAccess.GetKeyMatchValues()

    ' DJP SQL Server port. Need to convert the OrganisationID into a raw byte format
    vOrganisationID = colMatchValues(1)
    vOrganisationID = g_clsSQLAssistSP.GuidStringToByteArray(vOrganisationID)
    colMatchValues.Remove (1)
    colMatchValues.Add vOrganisationID
    
    Set colKeys = clsTableAccess.GetKeyMatchFields()
    
    Set rs = clsTableAccess.GetTableData()
    
    If clsTableAccess.RecordCount() <> 1 Then
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": No Lender records found"
    End If
    
    sRateSet = clsSrcLender.GetRateSet()
    vOrganisationID = clsSrcLender.GetOrganisationID()
    
    ' Copy the Lender table itself
    Set clsTargetLender = New MortgageLendersTable
    Set clsTableAccess = clsTargetLender
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.SetDatabase sTargetDatabase

    CopyObject clsSrcLender, clsTargetLender
    
    ' MIGRateSet
    CopyMIGRates sRateSet, sTargetDatabase

    ' Update the Morgage Lender table first
    Set clsTableAccess = clsTargetLender
    clsTableAccess.Update
    
    ' LenderOtherFee
    Dim clsSrcLenderOtherFee As LenderOtherFeesTable
    Dim clsTargetLenderOtherFee As LenderOtherFeesTable
    
    Set clsSrcLenderOtherFee = New LenderOtherFeesTable
    Set clsTargetLenderOtherFee = New LenderOtherFeesTable
    
    Set clsTableAccess = clsSrcLenderOtherFee
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    Set clsTableAccess = clsTargetLenderOtherFee
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.SetDatabase sTargetDatabase
    clsTableAccess.DeleteRecords sTargetDatabase
    
    CopyObject clsSrcLenderOtherFee, clsTargetLenderOtherFee, True
    Set clsTableAccess = clsTargetLenderOtherFee
    clsTableAccess.Update
    
    ' LegalFees
    CopyLegalFees colKeys, colMatchValues, sTargetDatabase
    
    ' MortgageLenderDirectory
    CopyLenderDirectory colMatchValues, sTargetDatabase

    ' MortgageLenderParameters
    Dim clsSrcParameters As MortLenderParamsTable
    Dim clsTargetParameters As MortLenderParamsTable
    
    Set clsSrcParameters = New MortLenderParamsTable
    Set clsTargetParameters = New MortLenderParamsTable
    
    Set clsTableAccess = clsSrcParameters
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    Set clsTableAccess = clsTargetParameters
    clsTableAccess.SetKeyMatchValues colMatchValues
    clsTableAccess.SetDatabase sTargetDatabase
    
    CopyObject clsSrcParameters, clsTargetParameters
    Set clsTableAccess = clsTargetParameters
    clsTableAccess.Update
    
    ' DJP SQL Server port - missing ledger codes so added
    Dim clsSrcLedgerCodes As TableAccess
    Dim clsTargetLedgerCodes As TableAccess
    
    Set clsSrcLedgerCodes = New LedgerCodesTable
    Set clsTargetLedgerCodes = New LedgerCodesTable
    
    If g_clsDataAccess.DoesTableExist(TableAccess(clsSrcLedgerCodes).GetTable) Then
        clsSrcLedgerCodes.SetKeyMatchValues colMatchValues
        clsTargetLedgerCodes.SetKeyMatchValues colMatchValues
        clsTargetLedgerCodes.SetDatabase sTargetDatabase
        
        CopyObject clsSrcLedgerCodes, clsTargetLedgerCodes
        
        clsTargetLedgerCodes.Update
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub CopyObject(clsSrcTableAccess As TableAccess, clsTargetTableAccess As TableAccess, Optional bCreate As Boolean = False, Optional bGetSrcData As Boolean = True, Optional bGetTargetData As Boolean = True)
    Dim nRows As Long
    Dim sTable As String
    Dim nFields As Long
    Dim nThisRow As Long
    Dim nThisField As Long
    Dim nTotalFields As Long
    Dim rowData As Variant
    Dim sTargetDatabase As String
    Dim rsSrc As ADODB.Recordset
    Dim rsDest As ADODB.Recordset
    Dim colMatchValues As Collection
    
    ' PSC 07/03/2006 MAR1298 - Start
    Dim intResult As Integer
    Dim sTargetField As String
    Dim colKeys As Collection
    Dim rsTemp As ADODB.Recordset
    Dim thisField As ADODB.field
    Dim tempCols As Collection
    Dim tempValues As Collection
    ' PSC 07/03/2006 MAR1298 - End

    On Error GoTo Failed

    sTargetDatabase = GetTargetDatabase()
    Set colMatchValues = clsSrcTableAccess.GetKeyMatchValues()
    clsTargetTableAccess.SetDatabase GetTargetDatabase()
    
    sTable = clsSrcTableAccess.GetTable()
    
    ' DJP 07/09/01 Check tables exist before attempting a promote
    If Not g_clsDataAccess.DoesTableExist(sTable) Then
        g_clsErrorHandling.RaiseError errGeneralError, "Source Database: " & sTable & " table does not exist"
    End If
    
    If Not g_clsDataAccess.DoesTableExist(sTable, sTargetDatabase) Then
        g_clsErrorHandling.RaiseError errGeneralError, "Target Database: " & sTable & " table does not exist"
    End If
    
    ' Source
    If bGetSrcData Then
        ' PSC 07/03/2006 MAR1298
        Set rsSrc = clsSrcTableAccess.GetTableData(POPULATE_KEYS_ORDERBY_PRIMARY)
    Else
        Set rsSrc = clsSrcTableAccess.GetRecordSet()
    End If
    
    nTotalFields = rsSrc.fields.Count * rsSrc.RecordCount
    
    If nTotalFields > 0 Then
        'm_clsProgress.StartProgress nTotalFields, frmManageUpdates
    End If

    '.StartProgress nTotalFields, m_sObjectName, frmManageUpdates
    ' Destination - if this record exists on the target, get it. If not, create
    ' one
    
    If bCreate = False Then
        If bGetTargetData Then
            ' PSC 07/03/2006 MAR1298 - Start
            ' Store the key values as the get set to null when using POPULATE_EMPTY
            Set tempCols = clsTargetTableAccess.GetKeyMatchFields()
            Set tempValues = clsTargetTableAccess.GetKeyMatchValues()
            
            Set rsTemp = clsTargetTableAccess.GetTableData(POPULATE_EMPTY)
            rsTemp.Close
            Set rsTemp = g_clsDataAccess.GetTableData(sTable, rsTemp.Source, GetTargetDatabase())
            
            clsTargetTableAccess.SetKeyMatchFields tempCols
            clsTargetTableAccess.SetKeyMatchValues tempValues
            Set rsDest = clsTargetTableAccess.GetTableData(POPULATE_KEYS_ORDERBY_PRIMARY)
            ' PSC 07/03/2006 MAR1298 - End
        Else
            Set rsDest = clsTargetTableAccess.GetRecordSet()
        End If
        
        If rsDest.RecordCount() = 0 Then
            bCreate = True
        End If
    Else
        Set rsDest = clsTargetTableAccess.GetTableData(POPULATE_EMPTY)
    End If
    
    If clsSrcTableAccess.RecordCount() > 0 Then
        ' PSC 07/03/2006 MAR1298 - Start
        ' If we've got the source data and are either creating a target or have got
        ' the target the data will be in primary key order otherwise use old method
        If bGetSrcData And (bCreate Or bGetTargetData) Then
            ' If creating then just add the new records into the empty target
            If bCreate Then
                rsSrc.MoveFirst
                
                While Not rsSrc.EOF
                    rsDest.AddNew
                    
                    For Each thisField In rsSrc.fields
                        sTargetField = thisField.Name
                        rsDest(sTargetField).Value = thisField.Value
                    Next
          
                    rsSrc.MoveNext
                Wend
            Else
                ' Merge the two recordsets
                ' Get the Primary Key definitions
                Set colKeys = g_clsDataAccess.GetTablePrimaryKeys(clsTargetTableAccess.GetTable())
                
                rsSrc.MoveFirst
                rsDest.MoveFirst
                           
                ' Go thorough source and target
                While Not rsSrc.EOF Or Not rsDest.EOF
                    
                    intResult = 0
                                     
                   ' Compare the primary keys of the current records
                    If Not rsSrc.EOF And Not rsDest.EOF Then
                        intResult = CompareKeys(rsSrc, rsDest, colKeys)
                    End If
                   
                    Select Case True
                        ' Source is less than target or we have reached the end of target
                        Case intResult = -1 Or rsDest.EOF
                            ' Add record into temporary recordset so we don't lose our place in the target
                            rsTemp.AddNew
                            
                            For Each thisField In rsSrc.fields
                                sTargetField = thisField.Name
                                rsTemp(sTargetField).Value = thisField.Value
                            Next
                
                            rsSrc.MoveNext
                       ' Source is greater than target or we have reached the end of source
                        Case intResult = 1 Or rsSrc.EOF
                            ' Leave the target record as is and move to the next row
                            rsDest.MoveNext
                        ' Source and target match
                        Case Else
                            For Each thisField In rsSrc.fields
                                sTargetField = thisField.Name
                                rsDest(sTargetField).Value = thisField.Value
                            Next
                            rsSrc.MoveNext
                            rsDest.MoveNext
                    End Select
                Wend
                            
                ' Copy any new rows across to the target
                If rsTemp.RecordCount > 0 Then
                    rsTemp.MoveFirst
                
                    While Not rsTemp.EOF
                        rsDest.AddNew
                        For Each thisField In rsTemp.fields
                            sTargetField = thisField.Name
                            rsDest(sTargetField).Value = thisField.Value
                        Next
                        rsTemp.MoveNext
                    Wend
                    
                End If
            End If
        Else
            ' Old Stuff left in as in other cases the data is not in primary
            ' key order so can't do the new processing
            rowData = rsSrc.GetRows()
            nRows = UBound(rowData, 2)
            
            ValidateCopy rsSrc, rsDest
            
            For nThisRow = 0 To nRows
                nFields = UBound(rowData, 1) + 1
                ' Copy the existing fields
    
                If bCreate Then
                    rsDest.AddNew
                End If
                
                If Not rsDest.EOF Then
                    For nThisField = 0 To nFields - 1
                        ' DJP SQL Server port
                        sTargetField = rsSrc(nThisField).Name
                        rsDest(sTargetField).Value = rowData(nThisField, nThisRow)
                    Next
                    
                    If Not bCreate Then
                        clsTargetTableAccess.MoveNext
                        If rsDest.EOF And rsDest.RecordCount < rsSrc.RecordCount Then
                            bCreate = True
                        End If
                    End If
                End If
            Next
        End If
        ' PSC 07/03/2006 MAR1298 - End
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub ValidateCopy(rsSrc As ADODB.Recordset, rsDest As ADODB.Recordset)
    On Error GoTo Failed
    Dim nSrcCount As Long
    Dim nDestCount As Long
    Dim sMessage As String
    Dim srcField As field
    Dim destField As field
    Dim tmpField As field
    Dim nDiff As Integer
    Dim colErrorFields As Collection
    Dim vName As Variant
    Const diffSrc As Integer = 1
    Const diffDest As Integer = 2

    If Not rsSrc Is Nothing And Not rsDest Is Nothing Then
        nSrcCount = rsSrc.fields.Count
        nDestCount = rsDest.fields.Count
        
        nDiff = diffSrc
        Set colErrorFields = New Collection
        
        If nSrcCount <> nDestCount Then
            If nSrcCount < nDestCount Then
                nDiff = diffDest
                
                Set tmpField = srcField
                Set srcField = destField
                Set destField = tmpField
            End If
        End If
        On Error Resume Next
        
        For Each srcField In rsSrc.fields
            Set destField = rsDest.fields(srcField.Name)
            If Err.Number <> 0 Then
                Err.Clear
                colErrorFields.Add srcField.Name
            End If
        Next
        
        If colErrorFields.Count > 0 Then
            On Error GoTo Failed
            
            sMessage = "The following fields don't exist on the "
            
            If nDiff = diffSrc Then
                sMessage = sMessage & " Destination database: " & vbCr & vbLf
            Else
                sMessage = sMessage & " Source database: "
            End If
            
            For Each vName In colErrorFields
                sMessage = sMessage & CStr(vName) & vbCr & vbLf
            Next
                        
            g_clsErrorHandling.RaiseError errGeneralError, sMessage
        End If
    Else
        g_clsErrorHandling.RaiseError errRecordSetEmpty
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyObjectKeys(clsDestTable As TableAccess, colKeyValues As Collection)
    On Error GoTo Failed
    Dim colKeyFields As Collection
    Dim sField As Variant
    Dim rsDest As ADODB.Recordset
    Dim nThisItem As Integer
    Dim sFunctionName As String
    
    sFunctionName = "CopyObjectKeys"
    Set colKeyFields = clsDestTable.GetKeyMatchFields()
    
    If colKeyFields.Count <> colKeyValues.Count Then
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Key count doesn't match value count"
    End If
    
    Set rsDest = clsDestTable.GetRecordSet()
    
    nThisItem = 1
    
    For Each sField In colKeyFields
        rsDest(sField) = colKeyValues(nThisItem)
        nThisItem = nThisItem + 1
    Next
    


    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub GetDatabaseObject(clsTableAccess As TableAccess, Optional clsSupervisorUpdateDetails As SupervisorUpdateDetails)
    On Error GoTo Failed
    Dim colKeyValues As Collection
    Dim rsDatabaseObject As ADODB.Recordset
    Const sFunctionName As String = "GetDatabaseObject"
    
    If Not clsSupervisorUpdateDetails Is Nothing Then
        Set colKeyValues = clsSupervisorUpdateDetails.GetKeyMatchValues()  'GetSelectedObjectKeyValues()
        clsTableAccess.SetKeyMatchValues colKeyValues
    End If
    
    Set rsDatabaseObject = clsTableAccess.GetTableData()
        
    ' Should be one
'    If clsTableAccess.RecordCount() <> 1 Then
'        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Record not found"
'
'    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteQuestion(clsSrcQuestion As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    
    clsSrcQuestion.SetDatabase g_clsDataAccess.GetConnectionKey()
    
    CopyRecords clsSrcQuestion, New AdditionalQuestionsTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteCondition(clsSrcCondition As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    
    clsSrcCondition.SetDatabase g_clsDataAccess.GetConnectionKey()
    
    CopyRecords clsSrcCondition, New ConditionsTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteIntermediary(clsSrcIntermediary As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcIntermediary.SetDatabase g_clsDataAccess.GetConnectionKey()
    
    CopyIntermediary clsSrcIntermediary, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromotePrintingTemplate(clsSrcPrintTemplate As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcPrintTemplate.SetDatabase g_clsDataAccess.GetConnectionKey()
    
    CopyPrintingTemplate clsSrcPrintTemplate, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteCurrency(clsSrcCurrency As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcCurrency.SetDatabase g_clsDataAccess.GetConnectionKey()
    
    CopyRecords clsSrcCurrency, New CurrencyTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteBusinessGroup(clsSrcBusinessGroup As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcBusinessGroup.SetDatabase g_clsDataAccess.GetConnectionKey()
    
    CopyBusinessGroup clsSrcBusinessGroup, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub PromoteBaseRate(ByRef clsSrcBaseRateTable As TableAccess, ByVal sTargetDatabase As String)
    
    Dim clsTrgBaseRateTable As TableAccess

    On Error GoTo Failed
    
    'Store the target database.
    m_sTargetDatabase = sTargetDatabase
    
    'Create a target table.
    Set clsTrgBaseRateTable = New RateTable
    
    'Ensure the tables are pointing to the correct databases.
    clsSrcBaseRateTable.SetDatabase g_clsDataAccess.GetConnectionKey()
    clsTrgBaseRateTable.SetDatabase sTargetDatabase
    
    'Copy the keys collection from source to destination.
    clsTrgBaseRateTable.SetKeyMatchValues clsSrcBaseRateTable.GetKeyMatchValues
    
    'Copy the Base Rate data.
    CopyObject clsSrcBaseRateTable, clsTrgBaseRateTable, False, True, True
    
    'Update the target table.
    clsTrgBaseRateTable.Update
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions
Public Sub PromoteIncomeMultipleSet(clsSrcParameter As TableAccess, sTargetDatabase As String)

    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New IncomeMultipleSetTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub PromoteMortgageProductCondition(clsSrcParameter As TableAccess, sTargetDatabase As String)
'GD BMIDS00016
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcParameter.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyRecords clsSrcParameter, New MortProdProdCondTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub CopyMortgageProductCondition(clsSrc As TableAccess, clsTarget As TableAccess, sTargetDatabase As String, Optional bDeleteFirst As Boolean = False)
'GD BMIDS00016
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim sProductCode As String
    Dim sStartDate As String
    Dim rsTarget As ADODB.Recordset
    Dim rsDelTarget As ADODB.Recordset
    
    'Source variables defn
    Dim clsSrcCond As TableAccess 'CONDITIONS

    'Target variables defn
    Dim clsTargetCond As TableAccess 'CONDITIONS

    'Source variable instantiation
    Set clsSrcCond = New ConditionsTable
    
    'Target variable instantiation
    Set clsTargetCond = New ConditionsTable
    Set clsTarget = New MortProdProdCondTable
    
    Set colMatchValues = clsSrc.GetKeyMatchValues()
    sProductCode = colMatchValues.Item(PRODUCT_CODE_KEY)
    sStartDate = colMatchValues.Item(PRODUCT_START_DATE_KEY)
    'Get All The MORTGAGEPRODUCTCONDITIONS associated with this product FROM SOURCE
    clsSrc.GetTableDataFromSearch "SELECT * FROM MORTGAGEPRODUCTCONDITION WHERE MORTGAGEPRODUCTCODE = '" & sProductCode & "' AND STARTDATE = '" & sStartDate & "'"

    If bDeleteFirst Then
        'Get any existing records on target and remove them.

        'Get all MORTGAGEPRODUCTCONDITIONS for this product from target
        Set rsTarget = g_clsDataAccess.GetTableData("MORTGAGEPRODUCTCONDITION", "SELECT * FROM MORTGAGEPRODUCTCONDITION WHERE MORTGAGEPRODUCTCODE = '" & sProductCode & "' AND STARTDATE = '" & sStartDate & "'", sTargetDatabase)
        If rsTarget.RecordCount > 0 Then
            'Delete MORTGAGEPRODUCTCONDITIONS for this product from Target
            Set rsDelTarget = g_clsDataAccess.ExecuteCommand("DELETE FROM MORTGAGEPRODUCTCONDITION WHERE MORTGAGEPRODUCTCODE = '" & sProductCode & "' AND STARTDATE = '" & sStartDate & "'", sTargetDatabase)
        End If
    End If
    CopyObject clsSrc, clsTarget
    clsTarget.Update
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub



Public Sub PromoteIncomefactor(ByRef clsSrcIncomeFactorTable As TableAccess, ByVal sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    clsSrcIncomeFactorTable.SetDatabase g_clsDataAccess.GetConnectionKey()
    CopyIncomeFactor clsSrcIncomeFactorTable, sTargetDatabase

    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub CopyIncomeFactor(ByRef clsSrcIncomeFactor As TableAccess, ByVal sTargetDatabase As String)
        
    Dim colMatchValues As Collection
    Dim clsTrgIncomeFactor As AllowableIncomeFactorTable
    
    On Error GoTo Failed
    
    Set colMatchValues = clsSrcIncomeFactor.GetKeyMatchValues()
    Set clsTrgIncomeFactor = New AllowableIncomeFactorTable
        
    TableAccess(clsTrgIncomeFactor).SetDatabase sTargetDatabase
    TableAccess(clsTrgIncomeFactor).SetKeyMatchValues colMatchValues
    
    CopyObject clsSrcIncomeFactor, clsTrgIncomeFactor
    TableAccess(clsTrgIncomeFactor).Update
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'MAR45 GHun
Public Sub PromoteDocumentLocations(clsSrcPrintingLocations As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcPrintingLocations.SetDatabase g_clsDataAccess.GetConnectionKey()
    
    CopyRecords clsSrcPrintingLocations, New TemplateTable, sTargetDatabase
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'MAR45 End

'MAR202 GHun
Public Sub PromotePackControl(clsSrcPrintPack As TableAccess, sTargetDatabase As String)
    On Error GoTo Failed
    
    m_sTargetDatabase = sTargetDatabase
    
    clsSrcPrintPack.SetDatabase g_clsDataAccess.GetConnectionKey()
    
    CopyPack clsSrcPrintPack, sTargetDatabase        ' MAR967

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
   
End Sub
'MAR202 End

'MAR967  Add code to promote pack
Private Sub CopyPack(clsSrcPack As PackControlTable, sTargetDatabase As String, Optional bGetSourceData As Boolean = True)
    On Error GoTo Failed
    ' Things to be copied here are the Pack Control and Pack Member tables and the
    ' templates for this pack.
    
    ' First, the Pack Control
    
    Dim clsTargetPack As PackControlTable
    Dim clsSrcTableAccess As TableAccess
    Dim clsTargetTableAccess As TableAccess
    Dim colMatchValues As Collection
    Dim rsSrc As ADODB.Recordset
    
    Dim nPackCount As Integer
    Dim iCount As Integer
    Dim sPackIDs As String
    Dim i As Integer
    Dim sDeleteSQL As String
    Dim sPackControlID As String
    
    Set clsTargetPack = New PackControlTable
    Set clsTargetTableAccess = clsTargetPack
    
    Set clsSrcTableAccess = clsSrcPack
    Set colMatchValues = clsSrcTableAccess.GetKeyMatchValues
    
    clsTargetTableAccess.SetKeyMatchValues colMatchValues
    clsTargetTableAccess.SetDatabase sTargetDatabase
    
    If bGetSourceData Then
        Set rsSrc = clsSrcTableAccess.GetTableData()
    Else
        Set rsSrc = clsSrcTableAccess.GetRecordSet
    End If
    
    If clsSrcTableAccess.RecordCount() > 0 Then
        
        CopyObject clsSrcPack, clsTargetPack
        clsTargetTableAccess.Update
        
        'Now promote the templates for this pack
        Dim clsSrcPackMember As PackMemberTable
        Dim clsTargetPackMember As PackMemberTable
        Dim clsSrcTemplate As PrintingTemplateTable

        Dim sTemplateID As String
        
        Set clsTargetPackMember = New PackMemberTable
        Set clsSrcPackMember = New PackMemberTable
        Set clsSrcTemplate = New PrintingTemplateTable
        
        TableAccess(clsSrcPack).MoveFirst
        
        For nPackCount = 1 To TableAccess(clsSrcPack).RecordCount
            
            sPackControlID = clsSrcPack.GetPackControlNumber
            
            'Populate Source with the PackMember records for current PackControlId
            clsSrcPackMember.GetMembersForPack sPackControlID
            TableAccess(clsTargetPackMember).SetDatabase sTargetDatabase
                    
            'Promote the templates for this pack.
            Dim col As Collection
            Dim nCnt As Long
            Set col = New Collection
            
            If TableAccess(clsSrcPackMember).RecordCount Then
                TableAccess(clsSrcPackMember).MoveFirst
                
                For nCnt = 1 To TableAccess(clsSrcPackMember).RecordCount
                    sTemplateID = clsSrcPackMember.GetHostTemplateID
                    col.Add sTemplateID
                    TableAccess(clsSrcTemplate).SetKeyMatchValues col
                    If Not bGetSourceData Then
                        TableAccess(clsSrcTemplate).GetTableData
                    End If
                    'Save the Template Records
                    CopyPrintingTemplate clsSrcTemplate, sTargetDatabase
                    col.Remove (1)
                    TableAccess(clsSrcPackMember).MoveNext
                Next
            End If
                    
            'Now promote the PackMember records for this pack.
                    
            'First delete these PackMember entries from the target so that we don't get any key
            'violations when copying records across.
            iCount = TableAccess(clsSrcPackMember).RecordCount
            sPackIDs = ""
            If iCount > 0 Then
                For i = 1 To iCount
                    If Len(sPackIDs) > 0 Then
                        TableAccess(clsSrcPackMember).MoveNext
                        sPackIDs = sPackIDs & " , " & "'" & clsSrcPackMember.GetHostTemplateID & "'"
                    Else
                        TableAccess(clsSrcPackMember).MoveFirst
                        sPackIDs = "'" & clsSrcPackMember.GetHostTemplateID & "'"
                    End If
                Next i
            
                TableAccess(clsSrcPackMember).MoveFirst
                
                sDeleteSQL = "Delete from PackMember where PackControlNumber = '" & sPackControlID & "'"
                g_clsDataAccess.ExecuteCommand sDeleteSQL, sTargetDatabase
                    
                TableAccess(clsTargetPackMember).GetTableData POPULATE_EMPTY
                
                CopyObject clsSrcPackMember, clsTargetPackMember, , False, False
                        
                'Save Pack Member Records
                TableAccess(clsTargetPackMember).Update
            End If
                                 
            TableAccess(clsSrcPack).MoveNext
        Next
    
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

' PSC 07/03/2006 MAR1298 - Start
Private Function CompareKeys(ByVal rsSource As ADODB.Recordset, ByVal rsDest As ADODB.Recordset, ByVal colKeys As Collection) As Integer
    On Error GoTo Failed
    
    Dim intIndex As Integer
    Dim intResult As Integer
    Dim strKey As String
    Dim strSourceGuid As String     'MAR1397
    Dim strDestGuid As String       'MAR1397
    intIndex = 1
    
    While intResult = 0 And intIndex <= colKeys.Count
        
        strKey = colKeys.Item(intIndex)
        
        'MAR1397 Convert guids
        If (rsSource(strKey).Type = adBinary) Then
            
            strSourceGuid = g_clsSQLAssistSP.GuidToString(CStr(rsSource(strKey)))
            strDestGuid = g_clsSQLAssistSP.GuidToString(CStr(rsDest(strKey)))
            
            Select Case True
                Case (StrComp(strSourceGuid, strDestGuid) < 0)
                    intResult = -1
                Case (StrComp(strSourceGuid, strDestGuid) > 0)
                    intResult = 1
                Case Else
                    intResult = 0
            End Select
        Else
        
            Select Case True
                Case rsSource(strKey) < rsDest(strKey)
                    intResult = -1
                Case rsSource(strKey) > rsDest(strKey)
                    intResult = 1
                Case Else
                    intResult = 0
            End Select
        End If

        intIndex = intIndex + 1
    Wend
    
    CompareKeys = intResult
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
' PSC 07/03/2006 MAR1298 - End





