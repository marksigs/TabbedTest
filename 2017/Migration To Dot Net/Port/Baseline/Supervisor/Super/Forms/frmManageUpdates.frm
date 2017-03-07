VERSION 5.00
Object = "{71648807-648A-434B-83A9-7A7F4CA61E39}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmManageUpdates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Promotions"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "frmManageUpdates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   8565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdReadLog 
      Caption         =   "&Read Log"
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   60
      Width           =   1275
   End
   Begin MSGOCX.MSGComboBox cboAvailableUpdates 
      Height          =   315
      Left            =   1800
      TabIndex        =   12
      Top             =   960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Text            =   ""
   End
   Begin VB.CommandButton cmdResetForUser 
      Caption         =   "Reset my history"
      Height          =   375
      Left            =   6540
      TabIndex        =   11
      Top             =   60
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset all history"
      Height          =   375
      Left            =   6540
      TabIndex        =   8
      Top             =   510
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdPromote 
      Caption         =   "&Promote"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4500
      Width           =   1335
   End
   Begin VB.CommandButton cmdDone 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   4500
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtManagement 
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   60
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Enabled         =   0   'False
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
   End
   Begin MSGOCX.MSGListView lvUpdates 
      Height          =   2235
      Left            =   120
      TabIndex        =   3
      Top             =   1500
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3942
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      AllowColumnReorder=   0   'False
   End
   Begin MSGOCX.MSGComboBox cboTarget 
      Height          =   315
      Left            =   1800
      TabIndex        =   6
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Text            =   ""
   End
   Begin VB.Label lblLabel 
      Caption         =   "Status:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   4020
      Width           =   735
   End
   Begin VB.Label lblTitle 
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   4020
      Width           =   6795
   End
   Begin VB.Label lblLabel 
      Caption         =   "Target Database"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   1455
   End
   Begin VB.Label lblLabel 
      Caption         =   "Available Updates"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblLabel 
      Caption         =   "Active Database:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmManageUpdates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmManageUpdates
' Description   :   Allows changes made in Supervisor to be promoted to any
'                   defined database
'
' Change history
' Prog      Date        Description
' DJP       31/01/01    Removed GetChannelDependencies, and removed calls from
'                       CheckDepartments and CheckPayProtRates
' AA        01/02/01    Added functionality to enabled the promotion of
'                       Questions & Conditions.
' DJP       22/06/01    SQL Server port
' STB       26/11/01    SYS2278 - Commit each group of promotions.
' STB       20/12/01    SYS2550 - Intermediaries.
' STB       02/01/02    SYS2551 - Intermediary dependancies now also checks for
'                       Mortgage Products and Lenders.
' STB       04/01/02    SYS3581 - Promotion fixes.
' STB       21/01/02    SYS2957 - Supervisor Security Enhancement (see
'                       PopulateAvailableUpdates)
' STB       04/02/02    SYS3857 - Base Rates can now be promoted. Also
'                       corrected the 'Text' property is read-only error.
' DJP       13/02/02    SYS4052 Display promotions if promotion created when user not
'                       logged in.
' STB       22/02/02    SYS4091 Altered TelephoneUsage combo to use the -
'                       ContactTelephoneUsage combo.
' DJP       24/02/02    SYS4142 Client variant for promotions.
' DJP       26/02/02    SYS4121 Client variant for promotions/dependencies.
' DJP       27/02/02    SYS4188 Make BaseRates promote correctly when used as a
'                       dependency (Mortgage Product)
' STB       28/03/02    SYS4315 Dependancy order fixed for organisation.
' STB       13/05/02    SYS4417 Added AllowableIncomeFactors.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change history
' Prog      Date        Description
' GD        27/05/02    AQR : BMIDS00016 ; Supervisor Promotions
' CL        25/09/02    AQR : BMIDS00087 ; Modify promote to remove promoted items on muli-select error
' CL        27/09/02    AQR : BMIDS00064 ; Modify promote to include transaction on source database
' DB        06/01/03    AQR : SYS5457    ; Modify promotion of Intermediaries.
' DB        07/01/03    AQR : BM0010     ; Modify promotions where there are multiple dependencies.
' DJP       06/03/03    AQR : BM0423     ; When promoting, make sure the user is on the target database.
' DJP       11/03/03    AQR : BM0316     ; Added logging for promotions.
' DJP       11/03/03    AQR : BM0316     ; Added logging for promotions, ability to read log file.
' DJP       13/03/03    AQR : BM0458     ; Lender dependency not picked up for Mortgage Products.
' BS        10/04/03    BM0500 Amend CheckProductDependencies to get BaseRateSet from InterestRateType table
'                                               and resolve promotion problems in ExectuteDependencies
' BS        28/04/03    BM00498 Add a check for combo dependencies when promoting Tasks (new method CheckTask).
' BS        15/03/03    BM0500 Add combogroups to MortgageProduct promotion
' [MC]      25/04/2004  BMIDS763    Data Promotion code added for three new sets
'                                   (Product switch fee, Insurance Fee set, TT FeeSet)
' JD        12/07/2004  BMIDS765    Added RentalIncomeRateSets for promotion
'*******************************************************************************************************************
' MARS Change history
' Prog      Date        Description
' GHun      16/08/2005  MAR45 Apply BBG1370 (New screen for print configuration)
' HMA       31/01/2006  MAR967  Allow for Printing Packs
'--------------------------------------------------------------------------------------------------------
' EPSOM Change history
' Prog      Date        Description
' TW        09/10/2006  EP2_7 Added handling for Additional Borrowing Fee and Credit Limit Increase Fee
' TW        17/10/2006  EP2_15 Add promotion for MortgageClubNetworkAssociation
' TW        18/11/2006  EP2_132 ECR20/21 Further promotions for FSA related changes
' TW        23/11/2006  EP2_172 Change control EP2_5 - E2CR16 changes related to Introducer/Product Exclusives
' TW        11/12/2006  EP2_20 WP1 - Loans & Products Supervisor Changes part 3
' TW        14/12/2006  EP2_518 Added handling for Default Procuration Fees, Loan Amount Adjustments, LTV Amount Adjustments
' TW        05/02/2007  EP2_706 - Should  network be mandatory data for ar firms
' TW        23/03/2007  EP2_1942 - Promotion of Income Factors does not work
' SR        21/11/2007  CORE00000313(VR 37)/MAR1968 - Allow UserID to be empty (when using integrated security)
'--------------------------------------------------------------------------------------------------------
Option Explicit

'In-order to add new entites to be promoted the following routines require
'modification (assuming entity [XYZ] has been added): -
'
'   frmManageUpdates.SetupListViewHeaders()
'   frmManageUpdates.Set[XYZ]Headers()
'   frmManageUpdates.CheckDependancies()
'   frmManageUpdates.Check[XYZ]()
'   HandleUpdates.GetObjectDescription()
'   HandleUpdates.GetObjectTableClass()
'   HandleUpdates.PromoteObject()
'   SupervisorObjectCopy.Promote[XYZ]()
'   SupervisorObjectCopy.Copy[XYZ]()
'
'Additionally, Ensure the GetRowOfData() on the base table object has a final
'line with a displayable description. This description will be displayed in the
'first column of the listview on frmManageUpdates.frm.The Collection 's last
'item should have a key with a value of OBJECT_DESCRIPTION_KEY. This
'functionality is only used when the user right-clicks and chooses 'Mark For
'Promotion'.

'Control indexes.
Private Const ACTIVE_DATABASE As Integer = 0

'Return code of the form - indicates success or failure.
Private m_ReturnCode As MSGReturnCode

'Supervisor update table.
Private m_clsSupervisorUpdate As SupervisorUpdateTable

'Stores the current connection's database name.
Private m_sSrcServiceName As String

'Stores the current user id.
Private m_sSrcUserID As String

'Class which wraps-up the actual data-copying work.
Private m_clsPromote As SupervisorObjectCopy

Private m_sUserLoggedOn As String

' TW 18/11/2006 EP2_132
Private strIntroducerFirmSeqNo As String
' TW 18/11/2006 EP2_132 End

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateAvailableUpdates
' Description   :   Populates the available update combobox with a list of all
'                   updates available for the database selected
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateAvailableUpdates()
        
    Dim lIndex As Long
    Dim bEnable As Boolean
    Dim sTargetService As String
    Dim sTartgetUserID As String
    Const strFunctionName As String = "PopulateAvailableUpdates"
    Dim colKeyValues As Collection
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    
    BeginWaitCursor
    
    bEnable = False
    cboAvailableUpdates.SelText = ""
    GetTargetDatabase sTargetService, sTartgetUserID
    
    If Len(sTargetService) > 0 Then  'SR 22/11/2007 : CORE00000313(VR 37)/MAR1968 - Allow UserID to be empty (when using integrated security)
        m_clsSupervisorUpdate.FindTableUpdates m_sSrcServiceName, m_sSrcUserID, sTargetService, sTartgetUserID
        Set clsTableAccess = m_clsSupervisorUpdate
        
        'Create a collection for field values.
        Set colKeyValues = New Collection
        
        'Copy the objectnames into a collection (filtering out those which the
        'user has no permission to see).
        For lIndex = 1 To clsTableAccess.RecordCount
'           'TODO: If we need to restrict access in the future here's where to do it!
'            If g_clsSecurityMgr.HasAccess(g_sSupervisorUser, m_clsSupervisorUpdate.GetObjectName()) = True Then
                colKeyValues.Add m_clsSupervisorUpdate.GetObjectName()
'            End If
            
            'Move onto the next objectname.
            clsTableAccess.MoveNext
        Next lIndex
                
        cboAvailableUpdates.SetListTextFromCollection colKeyValues
        bEnable = True
    End If
    
    cboAvailableUpdates.Enabled = bEnable
    
    ' Clear the listview
    lvUpdates.ListItems.Clear
    EndWaitCursor
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetPackagerHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "id"
    
    colHeaders.Add lvHeaders
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetMortgageClubAssociationHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Name"
    
    colHeaders.Add lvHeaders
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cboAvailableUpdates_Click
' Description   :   Called when the Available Updates combo box is changed.
'                   Populates the listview with available updates
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboAvailableUpdates_Click()

    On Error GoTo Failed

    PopulateUpdateItemDetails

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cboTarget_Click
' Description   :   Called when the list of available target databases is
'                   changed. Populates the list of available updates
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboTarget_Click()
    On Error GoTo Failed
        
    ' BM0423 Reset the userID to the user logged on
    m_sUserLoggedOn = g_sSupervisorUser
    
    ' Enable and populate list of available updates
    PopulateAvailableUpdates
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Promote
' Description   :   Promotes a change from the source database to the selected target database.
'                   For each promotion request there can be any number of dependencies, so those
'                   have to be checked too. Once all dependencies have been found, they have to be promoted
'                   before the object we originally wanted to promote. The list of dependencies are
'                   iterated though and for each dependency, Promote calls itself to a) check for
'                   dependencies of dependencies, and to Promote each dependency from the bottom up.
'                   When control is finally returned to the top level Object, it will be promoted too.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Promote(clsSupervisorUpdateDetails As SupervisorUpdateDetails, colUpdateDetails As Collection) As Boolean
    On Error GoTo Failed

    ' Get the key details required for the copy, and the object name
    Dim sTargetDatabase As String
    Dim sServiceName As String
    Dim sUserID As String
    Dim sCreationDate As String
    Dim nDependencyCount As Integer
    Dim bPromoteDependency As Boolean
    Dim bPromote As Boolean
    Dim enumPromoteType As SupervisorPromoteType
    Dim sFunctionName As String

    sFunctionName = "Promote"
    WriteLogEvent "--- Promote --- Start"

    bPromoteDependency = False
    bPromote = True

    sTargetDatabase = GetTargetDatabaseName()

    ' How many dependencies do we have?
    nDependencyCount = colUpdateDetails.Count

    ' Edit or Delete?

    enumPromoteType = clsSupervisorUpdateDetails.GetOperation()

    ' We only want to deal with dependencies if we're editing
    If enumPromoteType = PromoteEdit Then
        ' Find the dependencies for the Object passed in.
        'DB BM0010 07/01/03
        Set colUpdateDetails = New Collection
        bPromoteDependency = CheckDependencies(clsSupervisorUpdateDetails, colUpdateDetails)
        
        'DB BM0010 07/01/03 - Commented out, not needed.
        'If nDependencyCount = colUpdateDetails.Count Then
        '    bPromoteDependency = False
        'Else
        '    bPromoteDependency = True
        'End If
        'DB End
        
        bPromote = True

        If bPromoteDependency Then
            ' Got all dependencies, so execute them. But not before adding the original object to the list.
            bPromote = ExectuteDependencies(colUpdateDetails, clsSupervisorUpdateDetails, sTargetDatabase)
        End If
    End If

    If bPromote Then
        Dim sObjectName As String
        Dim clsTableAccess As TableAccess
        Dim colKeyMatchValues As Collection
        Dim vInstanceID As Variant
        Dim clsSupervisorConnection As SupervisorConnection

        sObjectName = clsSupervisorUpdateDetails.GetObjectName()

        Set clsTableAccess = g_clsHandleUpdates.GetObjectTableClass(sObjectName)
        Set colKeyMatchValues = clsSupervisorUpdateDetails.GetKeyMatchValues()
        clsTableAccess.SetKeyMatchValues colKeyMatchValues

        WriteLogKeys clsTableAccess

        WriteLogEvent "Promote Source='" & g_clsDataAccess.GetConnectionKey & "' Target='" & sTargetDatabase & "'"
        g_clsHandleUpdates.PromoteObject clsSupervisorUpdateDetails, clsTableAccess, sTargetDatabase

        ' Get the connection
        g_clsDataAccess.GetSupervisorConnection clsSupervisorConnection, sTargetDatabase

        sUserID = clsSupervisorConnection.GetUserID()
        sServiceName = clsSupervisorConnection.GetDatabaseName()
        sCreationDate = clsSupervisorUpdateDetails.GetCreationDate()
        vInstanceID = clsSupervisorUpdateDetails.GetObjectID()

        g_clsHandleUpdates.SavePromotionRequest sServiceName, sUserID, CDate(sCreationDate), vInstanceID, enumPromoteType
        
        ' BM0423 Save the promoted item on the target database
        SaveTargetPromotion clsTableAccess
        WriteLogEvent "Promote - Saved Promotion request on source and target"
    End If

    cmdPromote.Enabled = False
    WriteLogEvent "--- Promote --- End"

    Promote = bPromote

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
' DJP BM0423
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveTargetPromotion
' Description   :   When promoting an item to a target database, make sure that a Change Request
'                   is also created on that database so it can be promoted to another database.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveTargetPromotion(clsTargetTable As TableAccess)
    On Error GoTo Failed
    
    ' Now save the request on the target
    Dim colLine As Collection
    Dim sDescription As String
    Dim sCurrentConnection As String
    Dim clsUser As OmigaUserTable
    
    ' First, see if the user has a login on the target database
    sCurrentConnection = g_clsDataAccess.GetConnectionKey
    g_clsDataAccess.SetActiveConnection GetTargetDatabaseName
    
    Set clsUser = New OmigaUserTable
    clsUser.GetUsers m_sUserLoggedOn
    
    If TableAccess(clsUser).RecordCount = 0 Then

        WriteLogEvent "UserID '" & m_sUserLoggedOn & "' doesn''t exist on target"
        frmPromotionUser.Show vbModal, Me
    
        If frmPromotionUser.GetReturnCode = MSGSuccess Then
            m_sUserLoggedOn = frmPromotionUser.GetUserID
            WriteLogEvent "User has selected '" & m_sUserLoggedOn & " '"
            Unload frmPromotionUser
        
        Else
            WriteLogEvent "User has not selected a user, promotion cancelled"
            Unload frmPromotionUser
            g_clsErrorHandling.RaiseError errGeneralError, "Promotion cannot continue without a valid target User ID"
        End If
    End If
    
    Set colLine = lvUpdates.GetLine
    ' First column is always the title
    sDescription = colLine(1)

    g_clsHandleUpdates.SaveChangeRequest clsTargetTable, , , m_sUserLoggedOn
    g_clsDataAccess.SetActiveConnection sCurrentConnection
    
    WriteLogEvent "SaveTargetPromotion - Successful"
            
    Exit Sub
Failed:
    m_sUserLoggedOn = g_sSupervisorUser
    ' If we have an error, ensure we've got the right database selected
    g_clsErrorHandling.SaveError
    
    If Len(sCurrentConnection) > 0 Then
        g_clsDataAccess.SetActiveConnection sCurrentConnection
    End If
    
    g_clsErrorHandling.RaiseError
End Sub

Private Sub cmdPromote_Click()
    On Error GoTo Failed
    
    ' Get the key details required for the copy, and the object name
    Dim sTargetDatabase As String
    Dim clsSupervisorUpdateDetails As SupervisorUpdateDetails
    Dim colUpdateDetails As Collection
    Dim nResponse As Integer
    Dim bPromoted As Boolean
    
    bPromoted = False
    nResponse = MsgBox("Are you sure you want to promote the selected record(s)?", vbYesNo + vbQuestion)
    
    If nResponse = vbYes Then
        SetLoggingContext ""
        WriteLogEvent "cmdPromote --- Start ('" & g_sSupervisorUser & "')"
        BeginWaitCursor
        sTargetDatabase = GetTargetDatabaseName()
        
        'MAR45 GHun
        If Not (ValidatePromotions()) Then
            Exit Sub
        End If
        'MAR45 End
        
        ' Target database
        g_clsDataAccess.BeginTrans sTargetDatabase
        
        'BMIDS00064 CL
        ' Source database
        g_clsDataAccess.BeginTrans
        'BMIDS00064 CL END
        
        ' Start looping
        Dim lstItem As ListItem
        For Each lstItem In lvUpdates.ListItems
            If lstItem.Selected = True Then
                GetSelectedObjectDetails clsSupervisorUpdateDetails, lstItem.Index
                SetLoggingContext clsSupervisorUpdateDetails.GetObjectName
                

                ' First check to see if we have any dependencies
                Set colUpdateDetails = New Collection
            
                bPromoted = Promote(clsSupervisorUpdateDetails, colUpdateDetails)
            End If
        Next
        
        If bPromoted Then
            Dim nItemCount As Long
            
            
            nItemCount = Me.lvUpdates.ListItems.Count
            
            If nItemCount > 1 Then
                PopulateUpdateItemDetails
            Else
                PopulateAvailableUpdates
            End If
            
            SetUpdateStateAfterPromote

            cmdPromote.Enabled = False
            
            'SYS2278 - STB: Commit the transaction for all of the promoted items.
            g_clsDataAccess.CommitTrans sTargetDatabase
            
            'BMIDS00064 CL
            ' Source database
            g_clsDataAccess.CommitTrans
            'BMIDS00064 CL END
            WriteLogEvent "Promote - transaction committed"
        Else
            g_clsDataAccess.RollbackTrans sTargetDatabase
            
            'BMIDS00064 CL
            ' Source database
            g_clsDataAccess.RollbackTrans
            'BMIDS00064 CL END
            
            'CL BMIDS00087
            PopulateAvailableUpdates
            'END BMIDS00087
        End If
        WriteLogEvent "cmdPromote --- End ('" & g_sSupervisorUser & "')"
        EndWaitCursor
    End If

    CloseLogging
    
    Exit Sub
    
Failed:
    g_clsDataAccess.RollbackTrans sTargetDatabase
    
    'BMIDS00064 CL
    ' Source database
    g_clsDataAccess.RollbackTrans
    'BMIDS00064 CL END
        
    g_clsErrorHandling.DisplayError
    WriteLogEvent "Promote - error is " & Err.DESCRIPTION
    
    'CL BMIDS00087
    PopulateAvailableUpdates
    'END BMIDS00087
    CloseLogging
    
End Sub
Private Sub SetUpdateStateAfterPromote()
    On Error GoTo Failed
    
    If lvUpdates.ListItems.Count = 0 Then
        cboAvailableUpdates.ListIndex = -1
        cboAvailableUpdates.Enabled = (cboAvailableUpdates.ListCount > 0)
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckUserCompetencies
' Description   : Check all of the competencies for the given user and return true if any of
'                 them require promoting.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckUserCompetencies(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
        
    Dim lIndex As Long
    Dim sType As String
    Dim colObjectList As Collection
    Dim bDependencyFound As Boolean
    Dim colMatchValues As Collection
    Dim clsUserCompetency As UserCompetencyTable
    
    On Error GoTo Failed
    
    'Create a usercompetency table object.
    Set clsUserCompetency = New UserCompetencyTable
    
    'Load all the competencies for the user.
    m_clsPromote.GetDatabaseObject clsUserCompetency, clsUpdateDetails
    
    'Only start processing them if there are any.
    If TableAccess(clsUserCompetency).RecordCount() > 0 Then
        TableAccess(clsUserCompetency).MoveFirst
        
        'Create a collection to hold keys collections in.
        Set colObjectList = New Collection
        
        'Iterate through each competency record returned.
        For lIndex = 1 To TableAccess(clsUserCompetency).RecordCount
            'Get the ID value to be added into a collection.
            sType = clsUserCompetency.GetType()
            
            'Create a new collection to hold keys.
            Set colMatchValues = New Collection
            
            'Add the key into the collection.
            colMatchValues.Add sType
            
            'Add the keys collection to the 'master' collection.
            colObjectList.Add colMatchValues
            
            'Move onto the next record
            TableAccess(clsUserCompetency).MoveNext
        Next lIndex
    End If
    
    'Check to see if any of these competencies require promoting.
    bDependencyFound = GetObjectDependencies(COMPETENCIES, colUpdateDetails, colObjectList)
    
    'Return the boolean to the caller.
    CheckUserCompetencies = bDependencyFound
    
    Exit Function

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckPackager
' Description   : Check to see if dependant combogroups, competencies, working hours and
'                 qualifications are pending promotion.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckPackager(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
        
    Dim sWorkingHourType As String
    Dim bDependencyFound As Boolean
    Dim colObjectList As Collection
    Dim colMatchValues As Collection
    Dim clsOmigaUser As OmigaUserTable
    
    On Error GoTo Failed
    
    'Create a user table.
    Set clsOmigaUser = New OmigaUserTable
    
    'Load the correct user record.
    m_clsPromote.GetDatabaseObject clsOmigaUser, clsUpdateDetails
    
    'UserCompentencies.
    bDependencyFound = CheckUserCompetencies(clsUpdateDetails, colUpdateDetails) Or bDependencyFound
    
    'Check for combogroup dependancies.
    Set colObjectList = New Collection
    
    'UserAccessType.
    Set colMatchValues = New Collection
    colMatchValues.Add "UserAccessType"
    colObjectList.Add colMatchValues
    
    'UserQualification.
    Set colMatchValues = New Collection
    colMatchValues.Add "UserQualification"
    colObjectList.Add colMatchValues
    
    'ComptencyType.
    Set colMatchValues = New Collection
    colMatchValues.Add "ComptencyType"
    colObjectList.Add colMatchValues
    
    'UserRole.
    Set colMatchValues = New Collection
    colMatchValues.Add "UserRole"
    colObjectList.Add colMatchValues
    
    'Title.
    Set colMatchValues = New Collection
    colMatchValues.Add "Title"
    colObjectList.Add colMatchValues
    
    'ComboGroups.
    bDependencyFound = GetObjectDependencies(COMBOBOX_ENTRIES, colUpdateDetails, colObjectList) Or bDependencyFound
    
    'Working Hours.
    sWorkingHourType = clsOmigaUser.GetWorkingHourType()
    
    'If there is a working hour record, check to see if it requires promoting.
    If Len(sWorkingHourType) > 0 Then
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        
        colMatchValues.Add sWorkingHourType
        colObjectList.Add colMatchValues
        
        bDependencyFound = GetObjectDependencies(WORKING_HOURS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    'Reset the colMatchValues collection.
    Set colMatchValues = New Collection
    colMatchValues.Add clsOmigaUser.GetUserID
    
    'Units.
    bDependencyFound = GetUserUnitDependencies(colMatchValues, colUpdateDetails) Or bDependencyFound
    
    CheckPackager = bDependencyFound
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckIntroducers
' Description   : Check to see if dependant objects are pending promotion.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckIntroducers(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
Dim clsIntroducer As IntroducerTable
Dim colObjectList As Collection
Dim colIntroducerMatchValues As Collection
Dim strIntroducerID As String
Dim strUserID As String
Dim bDependencyFound As Boolean
    
    On Error GoTo Failed
    
    Set clsIntroducer = New IntroducerTable
    Set colObjectList = New Collection
   
    'Load the correct Introducer record.
    m_clsPromote.GetDatabaseObject clsIntroducer, clsUpdateDetails
    strIntroducerID = clsIntroducer.GetIntroducerID
    
    ' Get User dependencies
    
    strUserID = clsIntroducer.GetUserID
    If Len(strUserID) > 0 Then
        Set colObjectList = New Collection
        Set colIntroducerMatchValues = New Collection
        colIntroducerMatchValues.Add strUserID
        
        colObjectList.Add colIntroducerMatchValues
        bDependencyFound = GetObjectDependencies(USERS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If

    
    CheckIntroducers = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Function CheckFirmClubNetworkAssociation(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
Dim clsFirmClubNetAssoc As FirmClubNetAssocTable
Dim colObjectList As Collection
Dim colMatchValues As Collection
Dim strFirmClubSeqNo As String
Dim strARFirmID As String
Dim strPrincipalFirmID As String
Dim strIntroducerID As String
Dim strClubNetworkAssociationID As String
Dim bDependencyFound As Boolean
    
    On Error GoTo Failed
    
    Set clsFirmClubNetAssoc = New FirmClubNetAssocTable
    Set colObjectList = New Collection
   
    'Load the correct FirmClubNetAssoc record.
    m_clsPromote.GetDatabaseObject clsFirmClubNetAssoc, clsUpdateDetails
    strFirmClubSeqNo = clsFirmClubNetAssoc.GetFirmClubSeqNo
   
    ' First, MortgageClubNetworkAssociation dependencies
    strClubNetworkAssociationID = clsFirmClubNetAssoc.GetClubNetworkAssociationID
    If Len(strClubNetworkAssociationID) > 0 Then
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        colMatchValues.Add strClubNetworkAssociationID
        
        colObjectList.Add colMatchValues
        bDependencyFound = GetObjectDependencies(CLUBSANDASSOCIATIONS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    ' Next, ARFirm dependencies
    strARFirmID = clsFirmClubNetAssoc.GetARFirmID
    If Len(strARFirmID) > 0 Then
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        colMatchValues.Add strARFirmID
        
        colObjectList.Add colMatchValues
        bDependencyFound = GetObjectDependencies(ARFIRMS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    ' Next, PrincipalFirm dependencies
    strPrincipalFirmID = clsFirmClubNetAssoc.GetPrincipalFirmID
    If Len(strPrincipalFirmID) > 0 Then
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        colMatchValues.Add strPrincipalFirmID
        
        colObjectList.Add colMatchValues
        bDependencyFound = GetObjectDependencies(PRINCIPALFIRMPACKAGER, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    CheckFirmClubNetworkAssociation = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Function CheckExclusiveProducts(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
Dim clsMortgageProductExclusivity As MortProdExclusivityTable
Dim colObjectList As Collection
Dim colMatchValues As Collection
Dim strMortgageProductExclusivitySeqNo As String
Dim strARFirmID As String
Dim strPrincipalFirmID As String
Dim strClubNetworkAssociationID As String
Dim strMortgageProductCode As String
Dim bDependencyFound As Boolean
    
    On Error GoTo Failed
    
    Set clsMortgageProductExclusivity = New MortProdExclusivityTable
    Set colObjectList = New Collection
   
    'Load the correct MortgageProductExclusivity record.
    m_clsPromote.GetDatabaseObject clsMortgageProductExclusivity, clsUpdateDetails
    strMortgageProductExclusivitySeqNo = clsMortgageProductExclusivity.GetProductExclusivitySeqNo
    
    ' First, PrincipalFirm dependencies
    strPrincipalFirmID = clsMortgageProductExclusivity.GetPrincipalFirmID
    If Len(strPrincipalFirmID) > 0 Then
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        colMatchValues.Add strPrincipalFirmID
        
        colObjectList.Add colMatchValues
        bDependencyFound = GetObjectDependencies(PRINCIPALFIRMPACKAGER, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    ' Next, ARFirm dependencies
    strARFirmID = clsMortgageProductExclusivity.GetARFirmID
    If Len(strARFirmID) > 0 Then
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        colMatchValues.Add strARFirmID
        
        colObjectList.Add colMatchValues
        bDependencyFound = GetObjectDependencies(ARFIRMS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    ' Next, MortgageClubNetworkAssociation dependencies
    strClubNetworkAssociationID = clsMortgageProductExclusivity.GetClubNetworkAssociationID
    If Len(strClubNetworkAssociationID) > 0 Then
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        colMatchValues.Add strClubNetworkAssociationID
        
        colObjectList.Add colMatchValues
        bDependencyFound = GetObjectDependencies(CLUBSANDASSOCIATIONS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    ' Next, MortgageProduct dependencies
    strMortgageProductCode = clsMortgageProductExclusivity.GetMortgageProductCode
    If Len(strMortgageProductCode) > 0 Then
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        colMatchValues.Add strMortgageProductCode
        
        colObjectList.Add colMatchValues
        bDependencyFound = GetObjectDependencies(MORTGAGE_PRODUCTS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    CheckExclusiveProducts = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Function CheckIntroducerFirm(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
Dim clsIntroducerFirm As IntroducerFirmTable
Dim colObjectList As Collection
Dim colIntroducerFirmMatchValues As Collection
Dim strIntroducerFirmSeqNo As String
Dim strARFirmID As String
Dim strPrincipalFirmID As String
Dim strIntroducerID As String
Dim strUserID As String
Dim bDependencyFound As Boolean
    
    On Error GoTo Failed
    
    Set clsIntroducerFirm = New IntroducerFirmTable
    Set colObjectList = New Collection
   
    'Load the correct IntroducerFirm record.
    m_clsPromote.GetDatabaseObject clsIntroducerFirm, clsUpdateDetails
    strIntroducerFirmSeqNo = clsIntroducerFirm.GetIntroducerFirmSequenceNumber
   
    ' First, User dependencies
    strUserID = clsIntroducerFirm.GetIntroducerID
    If Len(strUserID) > 0 Then
        Set colObjectList = New Collection
        Set colIntroducerFirmMatchValues = New Collection
        colIntroducerFirmMatchValues.Add strUserID
        
        colObjectList.Add colIntroducerFirmMatchValues
        bDependencyFound = GetObjectDependencies(USERS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
   
    ' Next, Introducer dependencies
    strIntroducerID = clsIntroducerFirm.GetIntroducerID
    If Len(strIntroducerID) > 0 Then
        Set colObjectList = New Collection
        Set colIntroducerFirmMatchValues = New Collection
        colIntroducerFirmMatchValues.Add strIntroducerID
        
        colObjectList.Add colIntroducerFirmMatchValues
        bDependencyFound = GetObjectDependencies(INTRODUCER, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    ' Next, ARFirm dependencies
    strARFirmID = clsIntroducerFirm.GetARFirmID
    If Len(strARFirmID) > 0 Then
        Set colObjectList = New Collection
        Set colIntroducerFirmMatchValues = New Collection
        colIntroducerFirmMatchValues.Add strARFirmID
        
        colObjectList.Add colIntroducerFirmMatchValues
        bDependencyFound = GetObjectDependencies(ARFIRMS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    ' Next, PrincipalFirm dependencies
    strPrincipalFirmID = clsIntroducerFirm.GetPrincipalFirmID
    If Len(strPrincipalFirmID) > 0 Then
        Set colObjectList = New Collection
        Set colIntroducerFirmMatchValues = New Collection
        colIntroducerFirmMatchValues.Add strPrincipalFirmID
        
        colObjectList.Add colIntroducerFirmMatchValues
        bDependencyFound = GetObjectDependencies(PRINCIPALFIRMPACKAGER, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    CheckIntroducerFirm = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckUsers
' Description   : Check to see if dependant combogroups, competencies, working hours and
'                 qualifications are pending promotion.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckUsers(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
        
    Dim sWorkingHourType As String
    Dim bDependencyFound As Boolean
    Dim colObjectList As Collection
    Dim colMatchValues As Collection
    Dim clsOmigaUser As OmigaUserTable
    
    On Error GoTo Failed
    
    'Create a user table.
    Set clsOmigaUser = New OmigaUserTable
    
    'Load the correct user record.
    m_clsPromote.GetDatabaseObject clsOmigaUser, clsUpdateDetails
    
    
    'UserCompentencies.
    bDependencyFound = CheckUserCompetencies(clsUpdateDetails, colUpdateDetails) Or bDependencyFound
    
    'Check for combogroup dependancies.
    Set colObjectList = New Collection
    
    'UserAccessType.
    Set colMatchValues = New Collection
    colMatchValues.Add "UserAccessType"
    colObjectList.Add colMatchValues
    
    'UserQualification.
    Set colMatchValues = New Collection
    colMatchValues.Add "UserQualification"
    colObjectList.Add colMatchValues
    
    'ComptencyType.
    Set colMatchValues = New Collection
    colMatchValues.Add "ComptencyType"
    colObjectList.Add colMatchValues
    
    'UserRole.
    Set colMatchValues = New Collection
    colMatchValues.Add "UserRole"
    colObjectList.Add colMatchValues
    
    'Title.
    Set colMatchValues = New Collection
    colMatchValues.Add "Title"
    colObjectList.Add colMatchValues
    
    'ComboGroups.
    bDependencyFound = GetObjectDependencies(COMBOBOX_ENTRIES, colUpdateDetails, colObjectList) Or bDependencyFound
    
    'Working Hours.
    sWorkingHourType = clsOmigaUser.GetWorkingHourType()
    
    'If there is a working hour record, check to see if it requires promoting.
    If Len(sWorkingHourType) > 0 Then
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        
        colMatchValues.Add sWorkingHourType
        colObjectList.Add colMatchValues
        
        bDependencyFound = GetObjectDependencies(WORKING_HOURS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    'Reset the colMatchValues collection.
    Set colMatchValues = New Collection
    colMatchValues.Add clsOmigaUser.GetUserID
    
    'Units.
    bDependencyFound = GetUserUnitDependencies(colMatchValues, colUpdateDetails) Or bDependencyFound
    
    CheckUsers = bDependencyFound
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckBaseRate
' Description   : Check to see if the dependant combogroup is marked for
'                 promotion.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckBaseRate(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean

    Dim bDependencyFound As Boolean
    Dim colObjectList As Collection
    Dim colMatchValues As Collection

    'Check for combogroup dependancies.
    Set colObjectList = New Collection
    
    'UserAccessType.
    Set colMatchValues = New Collection
    colMatchValues.Add "RateType"
    colObjectList.Add colMatchValues
    
    'ComboGroups.
    bDependencyFound = GetObjectDependencies(COMBOBOX_ENTRIES, colUpdateDetails, colObjectList) Or bDependencyFound
    
    CheckBaseRate = bDependencyFound

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckIncomeFactors
' Description   : Check to see if the dependant combogroups are marked for
'                 promotion and the dependant lenders.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckIncomeFactors(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean

    Dim vOrganisationID As Variant
    Dim bDependencyFound As Boolean
    Dim colObjectList As Collection
    Dim colMatchValues As Collection
    Dim clsIncomeFactorTable As AllowableIncomeFactorTable

    'List of dependant objects.
    Set colObjectList = New Collection
    
    'Create our income factor table
    Set clsIncomeFactorTable = New AllowableIncomeFactorTable
    
    'Load the correct user record.
    m_clsPromote.GetDatabaseObject clsIncomeFactorTable, clsUpdateDetails
    
    'Get the Lender key from the i/factor record.
    vOrganisationID = clsIncomeFactorTable.GetOrganisationID
    
    'Get the Lender key into a collection.
    Set colObjectList = New Collection
    Set colMatchValues = New Collection

    colMatchValues.Add vOrganisationID
    colObjectList.Add colMatchValues
    
    bDependencyFound = GetObjectDependencies(LENDERS, colUpdateDetails, colObjectList) Or bDependencyFound

    'Check for combogroup dependancies.
    'UserAccessType.
    Set colMatchValues = New Collection
    colMatchValues.Add "EmploymentStatus"
    colObjectList.Add colMatchValues
    
    Set colMatchValues = New Collection
    colMatchValues.Add "IncomeGroupType"
    colObjectList.Add colMatchValues
    
    Set colMatchValues = New Collection
    colMatchValues.Add "IncomeType"
    colObjectList.Add colMatchValues
    
    'ComboGroups.
    bDependencyFound = GetObjectDependencies(COMBOBOX_ENTRIES, colUpdateDetails, colObjectList) Or bDependencyFound
                
    CheckIncomeFactors = bDependencyFound

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckBaseRateSets
' Description   : Check if the dependant base rate is marked for promotion.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckBaseRateSets(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
    On Error GoTo Failed
    Dim bDependencyFound As Boolean
    Dim sFeeSet As String
    Dim sRateID As String
    Dim colObjectList As Collection
    Dim colMatchValues As Collection
    Dim clsBaseRateTable As BaseRateTable
    Dim clsBaseRateSetsTable As BaseRateSetTable

    Set colMatchValues = New Collection
    'Create a BaseRateSet table.
    Set clsBaseRateTable = New BaseRateTable
    Set clsBaseRateSetsTable = New BaseRateSetTable
    
    'Load the correct record.
    m_clsPromote.GetDatabaseObject clsBaseRateTable, clsUpdateDetails
    
    sFeeSet = clsBaseRateTable.GetFeeSet
    colMatchValues.Add sFeeSet
    TableAccess(clsBaseRateSetsTable).SetKeyMatchValues colMatchValues
    TableAccess(clsBaseRateSetsTable).GetTableData
    
    sRateID = clsBaseRateSetsTable.GetRateID
    
    ' Check the Base Rate.
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    colMatchValues.Add sRateID
    colObjectList.Add colMatchValues
    
    bDependencyFound = GetObjectDependencies(BASE_RATE, colUpdateDetails, colObjectList) Or bDependencyFound
    
    CheckBaseRateSets = bDependencyFound
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetUserUnitDependencies
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetUserUnitDependencies(ByRef colMatchValues As Collection, ByRef colUpdateDetails As Collection) As Boolean
        
    Dim sUnit As String
    Dim rs As ADODB.Recordset
    Dim bDependencyFound As Boolean
    Dim colObjectList As Collection
    Dim clsUserRole As UserRoleTable
    Dim clsTableAccess As TableAccess
    Dim clsUpdateDetails As SupervisorUpdateDetails
    
    On Error GoTo Failed
    
    bDependencyFound = False
    Set clsUserRole = New UserRoleTable
    Set clsTableAccess = clsUserRole
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    m_clsPromote.GetDatabaseObject clsUserRole
    Set clsTableAccess = clsUserRole
    
    If clsTableAccess.RecordCount() > 0 Then
        Set rs = clsTableAccess.GetRecordSet()
                
        Set clsUpdateDetails = New SupervisorUpdateDetails
        clsTableAccess.MoveFirst
        
        Do While Not rs.EOF
            Set colObjectList = New Collection
            Set colMatchValues = New Collection
            
            sUnit = clsUserRole.GetUnitId()
            colMatchValues.Add sUnit
            colObjectList.Add colMatchValues
            
            clsUpdateDetails.SetKeyMatchValues colMatchValues
            
            'Check all the dependant records on this unit.
            bDependencyFound = CheckUnit(clsUpdateDetails, colUpdateDetails) Or bDependencyFound
            
            'Check the unit itself.
            bDependencyFound = GetObjectDependencies(UNITS, colUpdateDetails, colObjectList) Or bDependencyFound
            
            'bDependencyFound = CheckUnitDepartment(colUpdateDetails, colMatchValues) Or bDependencyFound
            
            clsTableAccess.MoveNext
        Loop
    End If
    
    GetUserUnitDependencies = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Private Function CheckARFirm(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
Dim clsARFirm As ARFirmTable
Dim colObjectList As Collection
Dim colMatchValues As Collection
Dim strARFirmID As String
Dim strUnitID As String
Dim bDependencyFound As Boolean
    
    On Error GoTo Failed
    
    Set clsARFirm = New ARFirmTable
    Set colObjectList = New Collection
   
    'Load the correct ARFirm record.
    m_clsPromote.GetDatabaseObject clsARFirm, clsUpdateDetails
    strARFirmID = clsARFirm.GetARFirmID
    
    ' Get Unit dependencies
    
    strUnitID = clsARFirm.GetUnitId
    If Len(strUnitID) > 0 Then
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        colMatchValues.Add strUnitID
        
        colObjectList.Add colMatchValues
        bDependencyFound = GetObjectDependencies(UNITS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If

    CheckARFirm = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Function CheckPrincipalFirm(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
Dim clsPrincipalFirm As PrincipalFirmTable
Dim colObjectList As Collection
Dim colMatchValues As Collection
Dim strPrincipalFirmID As String
Dim strUnitID As String
Dim bDependencyFound As Boolean
    
    On Error GoTo Failed
    
    Set clsPrincipalFirm = New PrincipalFirmTable
    Set colObjectList = New Collection
   
    'Load the correct PrincipalFirm record.
    m_clsPromote.GetDatabaseObject clsPrincipalFirm, clsUpdateDetails
    strPrincipalFirmID = clsPrincipalFirm.GetPrincipalFirmID
    
    ' Get Unit dependencies
    
    strUnitID = clsPrincipalFirm.GetUnitId
    If Len(strUnitID) > 0 Then
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        colMatchValues.Add strUnitID
        
        colObjectList.Add colMatchValues
        bDependencyFound = GetObjectDependencies(UNITS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If

    CheckPrincipalFirm = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Function CheckFirmTradingName(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
Dim clsFirmTradingName As FirmTradingNameTable
Dim colObjectList As Collection
Dim colMatchValues As Collection
Dim strFirmTradingNameSeqNo As String
Dim strARFirmID As String
Dim strPrincipalFirmID As String
Dim bDependencyFound As Boolean
    
    On Error GoTo Failed
    
    Set clsFirmTradingName = New FirmTradingNameTable
    Set colObjectList = New Collection
   
    'Load the correct FirmTradingName record.
    m_clsPromote.GetDatabaseObject clsFirmTradingName, clsUpdateDetails
    strFirmTradingNameSeqNo = clsFirmTradingName.GetFirmTradingNameSeqNo
    
    ' First, ARFirm dependencies
    
    strARFirmID = clsFirmTradingName.GetARFirmID
    If Len(strARFirmID) > 0 Then
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        colMatchValues.Add strARFirmID
        
        colObjectList.Add colMatchValues
        bDependencyFound = GetObjectDependencies(ARFIRMS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    ' Next, PrincipalFirm dependencies
    strPrincipalFirmID = clsFirmTradingName.GetPrincipalFirmID
    If Len(strPrincipalFirmID) > 0 Then
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        colMatchValues.Add strPrincipalFirmID
        
        colObjectList.Add colMatchValues
        bDependencyFound = GetObjectDependencies(PRINCIPALFIRMPACKAGER, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    CheckFirmTradingName = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Private Function CheckAppointments(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
Dim clsAppointments As AppointmentsTable
Dim colObjectList As Collection
Dim colAppointmentsMatchValues As Collection
Dim strAppointmentsID As String
Dim strARFirmID As String
Dim strPrincipalFirmID As String
Dim strActivityID As String
Dim bDependencyFound As Boolean
    
    On Error GoTo Failed
    
    Set clsAppointments = New AppointmentsTable
    Set colObjectList = New Collection
   
    'Load the correct Appointments record.
    m_clsPromote.GetDatabaseObject clsAppointments, clsUpdateDetails
    strAppointmentsID = clsAppointments.GetAppointmentID
    
    ' First, ARFirm dependencies
    
    strARFirmID = clsAppointments.GetARFirmID
    If Len(strARFirmID) > 0 Then
        Set colObjectList = New Collection
        Set colAppointmentsMatchValues = New Collection
        colAppointmentsMatchValues.Add strARFirmID
        
        colObjectList.Add colAppointmentsMatchValues
        bDependencyFound = GetObjectDependencies(ARFIRMS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    ' Next, PrincipalFirm dependencies
    strPrincipalFirmID = clsAppointments.GetPrincipalFirmID
    If Len(strPrincipalFirmID) > 0 Then
        Set colObjectList = New Collection
        Set colAppointmentsMatchValues = New Collection
        colAppointmentsMatchValues.Add strPrincipalFirmID
        
        colObjectList.Add colAppointmentsMatchValues
        bDependencyFound = GetObjectDependencies(PRINCIPALFIRMPACKAGER, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    CheckAppointments = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Function CheckFirmPermissions(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
Dim clsFirmPermissions As FirmPermissionsTable
Dim colObjectList As Collection
Dim colFirmPermissionsMatchValues As Collection
Dim strFirmPermissionsSeqNo As String
Dim strARFirmID As String
Dim strPrincipalFirmID As String
Dim strActivityID As String
Dim bDependencyFound As Boolean
    
    On Error GoTo Failed
    
    Set clsFirmPermissions = New FirmPermissionsTable
    Set colObjectList = New Collection
   
    'Load the correct FirmPermissions record.
    m_clsPromote.GetDatabaseObject clsFirmPermissions, clsUpdateDetails
    strFirmPermissionsSeqNo = clsFirmPermissions.GetFirmPermissionsSeqNo
    
    ' First, ARFirm dependencies
    
    strARFirmID = clsFirmPermissions.GetARFirmID
    If Len(strARFirmID) > 0 Then
        Set colObjectList = New Collection
        Set colFirmPermissionsMatchValues = New Collection
        colFirmPermissionsMatchValues.Add strARFirmID
        
        colObjectList.Add colFirmPermissionsMatchValues
        bDependencyFound = GetObjectDependencies(ARFIRMS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    ' Next, PrincipalFirm dependencies
    strPrincipalFirmID = clsFirmPermissions.GetPrincipalFirmID
    If Len(strPrincipalFirmID) > 0 Then
        Set colObjectList = New Collection
        Set colFirmPermissionsMatchValues = New Collection
        colFirmPermissionsMatchValues.Add strPrincipalFirmID
        
        colObjectList.Add colFirmPermissionsMatchValues
        bDependencyFound = GetObjectDependencies(PRINCIPALFIRMPACKAGER, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    ' Next, ActivityFSA dependencies
    strActivityID = clsFirmPermissions.GetActivityID
    If Len(strActivityID) > 0 Then
        Set colObjectList = New Collection
        Set colFirmPermissionsMatchValues = New Collection
        colFirmPermissionsMatchValues.Add strActivityID
        
        colObjectList.Add colFirmPermissionsMatchValues
        bDependencyFound = GetObjectDependencies(ACTIVITYFSA, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    CheckFirmPermissions = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function




Private Function CheckUnitDepartment(colUpdateDetails As Collection, colMatchValues As Collection) As Boolean
    On Error GoTo Failed
    Dim clsUnit As UnitTable
    Dim clsTableAccess As TableAccess
    Dim colObjectList As Collection
    Dim colDepMatchValues As Collection
    Dim sDepartmentID As String
    Dim bDependencyFound As Boolean
    
    Set clsUnit = New UnitTable
    Set colObjectList = New Collection
    Set colDepMatchValues = New Collection

    Set clsUnit = New UnitTable
    Set clsTableAccess = clsUnit
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    ' First, Department dependencies
    m_clsPromote.GetDatabaseObject clsUnit
    sDepartmentID = clsUnit.GetDepartmentID()
    colDepMatchValues.Add sDepartmentID
    
    colObjectList.Add colDepMatchValues
    
    bDependencyFound = GetObjectDependencies(DEPARTMENTS, colUpdateDetails, colObjectList)
    CheckUnitDepartment = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetRegionDependencies
' Description   : Given a UnitID, this routine will obtain a list of regions assigned to that
'                 Unit and return TRUE if any of them require promoting.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetRegionDependencies(sUnitID As String, colUpdateDetails As Collection) As Boolean
    On Error GoTo Failed
    
    ' Region Dependencies
    Dim sRegionID As String
    Dim bDependencyFound As Boolean
    Dim clsTableAccess As TableAccess
    Dim rsRegion As ADODB.Recordset
    Dim clsRegion As RegionTable
    Dim colDepMatchValues As Collection
    Dim colObjectList As Collection
    
    Set colObjectList = New Collection
    
    bDependencyFound = False
    Set clsRegion = New RegionTable
        
    clsRegion.GetRegionsFromUnit sUnitID
    Set clsTableAccess = clsRegion
    Set rsRegion = clsTableAccess.GetRecordSet()
    
    If clsTableAccess.RecordCount > 0 Then
        clsTableAccess.MoveFirst
        
        Do While Not rsRegion.EOF
            sRegionID = clsRegion.GetRegionID()
            
            Set colDepMatchValues = New Collection
            colDepMatchValues.Add sRegionID
            colObjectList.Add colDepMatchValues
            
            bDependencyFound = GetObjectDependencies(REGIONS, colUpdateDetails, colObjectList)
            
            clsTableAccess.MoveNext
        Loop
    End If
    
    GetRegionDependencies = bDependencyFound
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckUnit
' Description   : Check for dependant combogroups, regions, departments, channels (of those
'                 departments), countries (of those channels) and taskowners.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckUnit(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
        
    Dim sChannelID As String
    Dim sUnitID As String
    Dim clsUnit As UnitTable
    Dim sDepartmentID As String
    Dim sFunctionName As String
    Dim sCountryNumber As String
    Dim bDependencyFound As Boolean
    Dim colObjectList As Collection
    Dim colMatchValues As Collection
    Dim clsChannel As DistributionChannelTable
    
    On Error GoTo Failed
    
    sFunctionName = "CheckUnit"
    
    'Now check for related combo groups.
    bDependencyFound = CheckContactComboDependencies(colUpdateDetails) Or bDependencyFound
    
    'Reset the collection of keys collections.
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    'UserRole.
    colMatchValues.Add "UserRole"
    colObjectList.Add colMatchValues
    
    'Check for dependent combo groups.
    bDependencyFound = GetObjectDependencies(COMBOBOX_ENTRIES, colUpdateDetails, colObjectList) Or bDependencyFound

    
    'Create a unit table.
    Set clsUnit = New UnitTable
    
    'Load the unit record.
    m_clsPromote.GetDatabaseObject clsUnit, clsUpdateDetails
    
    'Get the department and unit IDs.
    sDepartmentID = clsUnit.GetDepartmentID()
    sUnitID = clsUnit.GetUnitId()
    
    'Distribution channel dependencies from the above department.
    Set clsChannel = New DistributionChannelTable
    
    'Load the relevant channel for the DepartmentID.
    clsChannel.GetChannelForDepartmentID sDepartmentID
        
    'Get the Channel ID and CountryID.
    sChannelID = clsChannel.GetChannelID()
    sCountryNumber = clsChannel.GetCountry()
    
    'Reset the collection of keys collections.
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
        
    'Add the CountryID.
    colMatchValues.Add sCountryNumber
    colObjectList.Add colMatchValues
    
    'Check for dependant countries.
    bDependencyFound = GetObjectDependencies(COUNTRIES, colUpdateDetails, colObjectList) Or bDependencyFound
    
    'Check for any dependant UNIT-regions.
    bDependencyFound = GetRegionDependencies(sUnitID, colUpdateDetails) Or bDependencyFound
    
    'Reset the collection of keys collections.
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    'Add the ChannelID.
    colMatchValues.Add sChannelID
    colObjectList.Add colMatchValues
    
    'Check for dependant channels.
    bDependencyFound = GetObjectDependencies(DISTRIBUTION_CHANNELS, colUpdateDetails, colObjectList) Or bDependencyFound
    
    'Reset the collection of keys collections.
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    'Add the DepartmentID.
    colMatchValues.Add sDepartmentID
    colObjectList.Add colMatchValues
    
    'Check for dependant departments.
    bDependencyFound = GetObjectDependencies(DEPARTMENTS, colUpdateDetails, colObjectList) Or bDependencyFound
        
    CheckUnit = bDependencyFound
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckDepartments
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckDepartments(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
        
    Dim sChannelID As String
    Dim sFunctionName As String
    Dim colObjectList As Collection
    Dim bDependencyFound As Boolean
    Dim colMatchValues As Collection
    Dim clsDepartment As DepartmentTable
    
    On Error GoTo Failed
    
    sFunctionName = "CheckDepartments"
        
    Set clsDepartment = New DepartmentTable
    
    'Get the department.
    m_clsPromote.GetDatabaseObject clsDepartment, clsUpdateDetails
    
    'See if the contactdetails combogroups require promoting.
    bDependencyFound = CheckContactComboDependencies(colUpdateDetails)
        
    'Now check for dependant channels.
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    'Get the ChannelID from the department.
    sChannelID = clsDepartment.GetChannelID()
    
    'Build a collection of keys collections.
    colMatchValues.Add sChannelID
    colObjectList.Add colMatchValues
    
    bDependencyFound = GetObjectDependencies(DISTRIBUTION_CHANNELS, colUpdateDetails, colObjectList) Or bDependencyFound
        
    CheckDepartments = bDependencyFound
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckIntermediary
' Description   : Checks if the intermediary has a parent, or lenders or products which are
'                 related to any procuration fee records. True is returned in either case.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckIntermediary(clsUpdateDetails As SupervisorUpdateDetails, colUpdateDetails As Collection) As Boolean
                
    Dim vParentGUID As Variant
    Dim colParentInt As Collection
    Dim colObjectList As Collection
    Dim bDependencyFound As Boolean
    Dim colMatchValues As Collection
    Dim clsIntemediary As IntermediaryTable
    
    On Error GoTo Failed
    
    Set clsIntemediary = New IntermediaryTable
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
      
    'Populate the Intermediary table object from the source data.
    m_clsPromote.GetDatabaseObject clsIntemediary, clsUpdateDetails
    
    'Check for dependant combogroups.
    Set colObjectList = New Collection

    'IntermediaryType.
    Set colMatchValues = New Collection
    colMatchValues.Add "IntermediaryType"
    colObjectList.Add colMatchValues

    'IntermediaryReportFrequency.
    Set colMatchValues = New Collection
    colMatchValues.Add "IntermediaryReportFrequency"
    colObjectList.Add colMatchValues

    'IntermediaryInsuranceType.
    Set colMatchValues = New Collection
    colMatchValues.Add "IntermediaryInsuranceType"
    colObjectList.Add colMatchValues

    'IntermediaryCorrespondenceType.
    Set colMatchValues = New Collection
    colMatchValues.Add "IntermediaryCorrespondenceType"
    colObjectList.Add colMatchValues

    'CommunicationMethod.
    Set colMatchValues = New Collection
    colMatchValues.Add "CommunicationMethod"
    colObjectList.Add colMatchValues

    'PaymentMethod.
    Set colMatchValues = New Collection
    colMatchValues.Add "PaymentMethod"
    colObjectList.Add colMatchValues

    'InsuranceType.
    Set colMatchValues = New Collection
    colMatchValues.Add "InsuranceType"
    colObjectList.Add colMatchValues

    'ComboGroups.
    bDependencyFound = GetObjectDependencies(COMBOBOX_ENTRIES, colUpdateDetails, colObjectList) Or bDependencyFound

    'Contact Details combogroups.
    bDependencyFound = CheckContactComboDependencies(colUpdateDetails) Or bDependencyFound
    
    'Get the parent's GUID (if there is one).
    vParentGUID = clsIntemediary.GetParentIntermediaryGUID
    
    'Loop up through each parent of this intermediary and check to see if each requires
    'promoting (in which case there is a dependancy).
    While Not IsNull(vParentGUID)
        Set colMatchValues = New Collection
        Set colParentInt = New Collection
        
        'Add a string version of the parent's GUID into a wrapper collection
        'which is passed into the GetObjectDependancies check.
        colParentInt.Add vParentGUID
        
        'Add the wrapper collection into the list of intermediaries to check for.
        colObjectList.Add colParentInt
        
        'Add the parent's key into a keys collection.
        colMatchValues.Add vParentGUID
        
        'Load the parent record.
        TableAccess(clsIntemediary).SetKeyMatchValues colMatchValues
        TableAccess(clsIntemediary).GetTableData
        
        'Get the GUID.
        vParentGUID = clsIntemediary.GetParentIntermediaryGUID
    Wend
    
    'Now check for dependancies on all these intermediary GUIDs.
    bDependencyFound = GetObjectDependencies(INTERMEDIARIES, colUpdateDetails, colObjectList) Or bDependencyFound
    
    CheckIntermediary = bDependencyFound
    
    Exit Function

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckStage
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckStage(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
    
    Dim sActivityID As String
    Dim clsStage As StageTable
    Dim sFunctionName As String
    Dim colObjectList As Collection
    Dim bDependencyFound As Boolean
    Dim colMatchValues As Collection
    
    On Error GoTo Failed
    
    sFunctionName = "CheckStage"
        
    'Create a stage table.
    Set clsStage = New StageTable
    
    'Load the correct stage to promote.
    m_clsPromote.GetDatabaseObject clsStage, clsUpdateDetails
    
    'Check for dependant combogroups.
    Set colObjectList = New Collection
    
    'UserRole.
    Set colMatchValues = New Collection
    colMatchValues.Add "UserRole"
    colObjectList.Add colMatchValues
    
    'ComboGroups.
    bDependencyFound = GetObjectDependencies(COMBOBOX_ENTRIES, colUpdateDetails, colObjectList) Or bDependencyFound
        
    'Get the Activity ID.
    sActivityID = clsStage.GetActivityID
    
    'Add into a keys collection.
    colMatchValues.Add sActivityID
    
    'If there is an associated activity.
    If Len(sActivityID) > 0 Then
        'Check for Dependant Activities.
        bDependencyFound = GetStageDependencies(colMatchValues, colUpdateDetails) Or bDependencyFound
    End If
    
    CheckStage = bDependencyFound
    
    Exit Function

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Private Function CheckPayProtRates(clsUpdateDetails As SupervisorUpdateDetails, colUpdateDetails As Collection) As Boolean
    On Error GoTo Failed
    Dim bDependencyFound As Boolean
    Dim sChannelID As String
    Dim clsPayProtRates As PayProtRatesTable
    Dim clsTableAccess As TableAccess
    Dim colMatchValues As Collection
    Dim rsChannels As ADODB.Recordset
    Dim colChannels As Collection
    ' Get the distribution channel dependencies for these rates
    
    Set colChannels = New Collection
    Set colMatchValues = clsUpdateDetails.GetKeyMatchValues()
    Set clsPayProtRates = New PayProtRatesTable
    Set clsTableAccess = clsPayProtRates
    
    clsTableAccess.SetKeyMatchValues colMatchValues
    
    clsTableAccess.GetTableData

    If clsTableAccess.RecordCount() > 0 Then
        Set rsChannels = clsTableAccess.GetRecordSet
                        
        Do While Not rsChannels.EOF
            sChannelID = clsPayProtRates.GetChannelID()
            colChannels.Add sChannelID
            
            clsTableAccess.MoveNext
        Loop
        
        bDependencyFound = GetDistributionChannelDependencies(colChannels, colUpdateDetails)
    End If
    
    CheckPayProtRates = bDependencyFound
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Function CheckPayProtProducts(clsUpdateDetails As SupervisorUpdateDetails, colUpdateDetails As Collection) As Boolean
    On Error GoTo Failed
    Dim sFunctionName As String
    Dim clsPayProtProducts As PayProtProductsTable
    Dim clsPayProtRates As PayProtRatesTable
    Dim clsTableAccess As TableAccess
    Dim sRateNumber As String
    Dim colKeys As Collection
    Dim colObjectList As Collection
    Dim colMatchValues As Collection
    Dim bDependencyFound As Boolean
    
    sFunctionName = "CheckPayProtProducts"
    
    Set clsPayProtProducts = New PayProtProductsTable
    Set clsPayProtRates = New PayProtRatesTable
        
    ' Get the Payment protection products
    m_clsPromote.GetDatabaseObject clsPayProtProducts, clsUpdateDetails
    
    sRateNumber = clsPayProtProducts.GetRateNumber()
    
    ' Any rate dependencies?
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    colMatchValues.Add sRateNumber
    colObjectList.Add colMatchValues
    Set clsTableAccess = clsPayProtProducts
    Set colKeys = clsTableAccess.GetKeyMatchFields()
    
    Set clsTableAccess = clsPayProtRates
    clsTableAccess.SetKeyMatchFields colKeys
    
    bDependencyFound = GetObjectDependencies(PAYMENT_PROTECTION_RATES, colUpdateDetails, colObjectList)
    
    CheckPayProtProducts = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Function CheckDistributionChannels(clsUpdateDetails As SupervisorUpdateDetails, colUpdateDetails As Collection) As Boolean
    On Error GoTo Failed
    ' The only dependency for distribution channels is the country table. If any
    ' records exist for promotion with the same country number as the one the distribution channel
    ' we are promoting, they need to be promoted first.
    
    ' So, first, get the country number
    Dim clsTableAccess As TableAccess
    Dim clsDistChannel As DistributionChannelTable
    Dim clsCountry As CountryTable
    Dim sFunctionName As String
    Dim sCountryNumber As String
    Dim colObjectList As Collection
    Dim colMatchValues As Collection
    Dim bDependencyFound As Boolean
     
    sFunctionName = "CheckDistributionChannels"
    
    Set clsDistChannel = New DistributionChannelTable
    Set clsCountry = New CountryTable
    
    ' Get the distribution channels
    m_clsPromote.GetDatabaseObject clsDistChannel, clsUpdateDetails
    Set clsTableAccess = clsDistChannel
    
    sCountryNumber = clsDistChannel.GetCountry()
    
    ' Now, any Country dependencies?
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    colMatchValues.Add sCountryNumber
    colObjectList.Add colMatchValues
    
    bDependencyFound = GetObjectDependencies(COUNTRIES, colUpdateDetails, colObjectList)
    
    CheckDistributionChannels = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Function CheckProductDependencies(clsUpdateDetails As SupervisorUpdateDetails, colUpdateDetails As Collection) As Boolean
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim clsMortgageProduct As MortgageProductTable
    Dim sFunctionName As String
    Dim sBaseRateSet As String
    Dim sValuationFeeSet As String
    Dim sAdminFeeSet As String
    Dim sIAFeeSet As String
    Dim sTTFeeSet As String
    Dim sProdSwitchFeeSet As String
    
' TW 09/10/2006 EP2_7
    Dim sAdditionalBorrowingFeeSet As String
    Dim sCreditLimitIncreaseFeeSet As String
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
    Dim sTransferOfEquityFeeSet As String
' TW 11/12/2006 EP2_20 End

    Dim colChannels As Collection
    Dim clsDistributionChannels As MortProdChannelEligTable
    Dim sProductCode As String
    Dim sStartDate As String
    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions START
    Dim clsMortgageProductCondition As MortProdProdCondTable
    Dim sRedemptionFeeSet As String
    Dim sMPMIGRateSet As String
    Dim sIncomeMultipleSet As String
    Dim sConditionReference As String
    Dim iIndex As Integer
    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions END
    Dim sLenderOrgID As Variant
    Dim bDependencyFound As Boolean
    'BS BM0500 10/04/03
    Dim clsInterestRates As MortProdIntRatesTable
    
    sFunctionName = "CheckProductDependencies"
    Set clsMortgageProduct = New MortgageProductTable
    Set clsMortgageProductCondition = New MortProdProdCondTable
    ' Get the product
    m_clsPromote.GetDatabaseObject clsMortgageProduct, clsUpdateDetails
    Set clsTableAccess = clsMortgageProduct
    'BS BM0500 10/04/03
    Set clsInterestRates = New MortProdIntRatesTable
        
    ' Get the BASERATESET, VALUATIONFEESET and ADMINISTRATIONFEESET
    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions
    ' AND REDEMPTIONFEESET and MPMIGRATESET and INCOMEMULTIPLESET
    ' & CONDITIONS
    
    'BS BM0500 10/04/03
    'BaseRateSet isn't populated on MortgageProduct table any more. It is held on the
    'InterestRateType table, so MortgageProducts can be associated with multiple BaseRateSets
    'sBaseRateSet = clsMortgageProduct.GetBaseRateFeeSet()
    'BS BM0500 End 10/04/03
    sValuationFeeSet = clsMortgageProduct.GetValuationFeeSet()
    sAdminFeeSet = clsMortgageProduct.GetAdminFeeSet()
    sProductCode = clsMortgageProduct.GetMortgageProductCode()
    sStartDate = clsMortgageProduct.GetStartDate()
    
' TW 09/10/2006 EP2_7
    sAdditionalBorrowingFeeSet = clsMortgageProduct.GetAdditionalBorrowingFeeSet
    sCreditLimitIncreaseFeeSet = clsMortgageProduct.GetCreditLimitIncreaseFeeSet
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
    sTransferOfEquityFeeSet = clsMortgageProduct.GetTransferOfEquityFeeSet
' TW 11/12/2006 EP2_20 End

    '*=[MC]BMIDS763
    sIAFeeSet = clsMortgageProduct.GetIAFeeSet()
    sTTFeeSet = clsMortgageProduct.GetTTFeeSet()
    sProdSwitchFeeSet = clsMortgageProduct.GetProductSwitchFeeSet()
    
    '*=SECTION END
    
    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions START
    sRedemptionFeeSet = clsMortgageProduct.GetRedemptionFeeSet()
    sMPMIGRateSet = clsMortgageProduct.GetMPMigRateSet()
    sIncomeMultipleSet = clsMortgageProduct.GetIncomeMultiplierCode()
    clsMortgageProductCondition.GetMortgageProductCondition sProductCode, sStartDate

    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions END
    
    ' DJP SQL Server port
    sLenderOrgID = clsMortgageProduct.GetOrganisationID()
    
    ' Now check for dependencies
    
    ' First, any combos?
    Dim colObjectList As Collection
    Dim colMatchValues As Collection
    Set colObjectList = New Collection
        
    Set colMatchValues = New Collection
    colMatchValues.Add "TypeOfMortgage"
    colObjectList.Add colMatchValues
    
    Set colMatchValues = New Collection
    colMatchValues.Add "TypeOfBuyerNewLoan"
    colObjectList.Add colMatchValues
    
    Set colMatchValues = New Collection
    colMatchValues.Add "TypeOfBuyerRemortgage"
    colObjectList.Add colMatchValues
    
    Set colMatchValues = New Collection
    colMatchValues.Add "PurposeOfLoan"
    colObjectList.Add colMatchValues
    
    Set colMatchValues = New Collection
    colMatchValues.Add "EmploymentType"
    colObjectList.Add colMatchValues
    
    Set colMatchValues = New Collection
    colMatchValues.Add "ProductBrandName"
    colObjectList.Add colMatchValues
    
    Set colMatchValues = New Collection
    colMatchValues.Add "ProductGroup"
    colObjectList.Add colMatchValues
    
    Set colMatchValues = New Collection
    colMatchValues.Add "ProductLine"
    colObjectList.Add colMatchValues
    
    Set colMatchValues = New Collection
    colMatchValues.Add "CountryOfOrigin"
    colObjectList.Add colMatchValues
    
    'BS BM0500 13/05/03
    Set colMatchValues = New Collection
    colMatchValues.Add "PropertyLocation"
    colObjectList.Add colMatchValues
    
    Set colMatchValues = New Collection
    colMatchValues.Add "EmploymentStatus"
    colObjectList.Add colMatchValues
        
    Set colMatchValues = New Collection
    colMatchValues.Add "PurposeOfLoanSupervisor"
    colObjectList.Add colMatchValues
    'BS BM0500 End 13/05/03
        
    bDependencyFound = GetObjectDependencies(COMBOBOX_ENTRIES, colUpdateDetails, colObjectList) Or bDependencyFound

    ' Now, any Lender updates?
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    colMatchValues.Add sLenderOrgID
    colObjectList.Add colMatchValues
    bDependencyFound = GetObjectDependencies(LENDERS, colUpdateDetails, colObjectList) Or bDependencyFound
    
    ' Next, Distribution Channels
    ' For the channel ID we need to get a list of channels associated with this
    ' product then check for them to be in the list of oustanding updates
    
    Set clsDistributionChannels = New MortProdChannelEligTable
    Set colChannels = clsDistributionChannels.GetChannelsForProduct(sProductCode, sStartDate)
    
    bDependencyFound = GetDistributionChannelDependencies(colChannels, colUpdateDetails) Or bDependencyFound
    
' TW 09/10/2006 EP2_7
    ' Additional Borrowing Fees
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    colMatchValues.Add sAdditionalBorrowingFeeSet
    colObjectList.Add colMatchValues
    bDependencyFound = GetObjectDependencies(ADDITIONAL_BORROWING_FEES, colUpdateDetails, colObjectList) Or bDependencyFound
    
    ' Credit Limit Increase Fees
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    colMatchValues.Add sCreditLimitIncreaseFeeSet
    colObjectList.Add colMatchValues
    bDependencyFound = GetObjectDependencies(CREDIT_LIMIT_INCREASE_FEES, colUpdateDetails, colObjectList) Or bDependencyFound
    
    sAdditionalBorrowingFeeSet = clsMortgageProduct.GetAdditionalBorrowingFeeSet
    sCreditLimitIncreaseFeeSet = clsMortgageProduct.GetCreditLimitIncreaseFeeSet
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
    ' Transfer of Equity Fees
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    colMatchValues.Add sTransferOfEquityFeeSet
    colObjectList.Add colMatchValues
    bDependencyFound = GetObjectDependencies(TRANSFER_OF_EQUITY_FEES, colUpdateDetails, colObjectList) Or bDependencyFound
    
    sTransferOfEquityFeeSet = clsMortgageProduct.GetTransferOfEquityFeeSet
' TW 11/12/2006 EP2_20 End
    
    ' Next, Administration Fees
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    colMatchValues.Add sAdminFeeSet
    colObjectList.Add colMatchValues
    bDependencyFound = GetObjectDependencies(ADMIN_FEES, colUpdateDetails, colObjectList) Or bDependencyFound
    
    '*=[MC]BMIDS763
    'INSURANCE ADMIN FEE SET
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    colMatchValues.Add sIAFeeSet
    colObjectList.Add colMatchValues
    bDependencyFound = GetObjectDependencies(INSURANCE_ADMIN_FEESETS, colUpdateDetails, colObjectList) Or bDependencyFound
    
    'TT FEE SET
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    colMatchValues.Add sTTFeeSet
    colObjectList.Add colMatchValues
    bDependencyFound = GetObjectDependencies(TT_FEESETS, colUpdateDetails, colObjectList) Or bDependencyFound
    
    'PRODUCT SWITCH FEE SETS
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    colMatchValues.Add sProdSwitchFeeSet
    colObjectList.Add colMatchValues
    bDependencyFound = GetObjectDependencies(PRODUCT_SWITCH_FEESETS, colUpdateDetails, colObjectList) Or bDependencyFound
    
    '*=SECTION END
    
    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions START
    
    'Redemption Fee Sets
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    colMatchValues.Add sRedemptionFeeSet
    colObjectList.Add colMatchValues
    bDependencyFound = GetObjectDependencies(REDEM_FEE_SETS, colUpdateDetails, colObjectList) Or bDependencyFound

    'MP MIG Rate Sets
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    colMatchValues.Add sMPMIGRateSet
    colObjectList.Add colMatchValues
    bDependencyFound = GetObjectDependencies(MP_MIG_RATE_SETS, colUpdateDetails, colObjectList) Or bDependencyFound
    
    'Income Multiple Sets
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    colMatchValues.Add sIncomeMultipleSet
    colObjectList.Add colMatchValues
    bDependencyFound = GetObjectDependencies(INCOME_MULTIPLE, colUpdateDetails, colObjectList) Or bDependencyFound

''    'MortgageProductConditions and Conditions
'    Need to check any dependencies of MortgageProductConditions on Conditions
    Set colObjectList = New Collection
    

    If TableAccess(clsMortgageProductCondition).RecordCount() > 0 Then
        TableAccess(clsMortgageProductCondition).MoveFirst
        
        'Create a collection to hold keys collections in.
        Set colObjectList = New Collection
        
        'Iterate through each competency record returned.
        For iIndex = 1 To TableAccess(clsMortgageProductCondition).RecordCount
            'Get the ID value to be added into a collection.
            sConditionReference = clsMortgageProductCondition.GetConditionReference()
            
            'Create a new collection to hold keys.
            Set colMatchValues = New Collection
            
            'Add the key into the collection.
            colMatchValues.Add sConditionReference
            
            'Add the keys collection to the 'master' collection.
            colObjectList.Add colMatchValues
            
            'Move onto the next record
            TableAccess(clsMortgageProductCondition).MoveNext
        Next iIndex
    End If
    
    bDependencyFound = GetObjectDependencies(CONDITIONS, colUpdateDetails, colObjectList) Or bDependencyFound
'

    'GD 27/05/02 AQR : BMIDS00016 ; Supervisor Promotions END
    
    
    ' Valuation Fees
    Set colObjectList = New Collection
    Set colMatchValues = New Collection
    
    colMatchValues.Add sValuationFeeSet
    colObjectList.Add colMatchValues
    bDependencyFound = GetObjectDependencies(VALUATION_FEES, colUpdateDetails, colObjectList) Or bDependencyFound
    
    ' Base Rates
    'BS BM0500 10/04/03
    'The BaseRateSet field isn't populated on MortgageProduct table any more. It is held on the
    'InterestRateType table, so need to check any dependencies of InterestRateTypes on BaseRateSet

    'Set colObjectList = New Collection
    Set colMatchValues = New Collection

    'colMatchValues.Add sBaseRateSet
    'colObjectList.Add colMatchValues
    
    colMatchValues.Add sProductCode
    colMatchValues.Add sStartDate
    
    clsInterestRates.GetInterestRates colMatchValues

    'InterestRateTypes and BaseRateSet
    
    If TableAccess(clsInterestRates).RecordCount() > 0 Then
        TableAccess(clsInterestRates).MoveFirst
        
        'Create a collection to hold keys collections in.
        Set colObjectList = New Collection
        
        'Iterate through each InterestRateType record returned.
        For iIndex = 1 To TableAccess(clsInterestRates).RecordCount
            'Get the ID value to be added into a collection.
            sBaseRateSet = clsInterestRates.GetBaseRateSet
            
            'Create a new collection to hold keys.
            Set colMatchValues = New Collection
            
            'Add the key into the collection.
            colMatchValues.Add sBaseRateSet
            
            'Add the keys collection to the 'master' collection.
            colObjectList.Add colMatchValues
            
            'Move onto the next record
            TableAccess(clsInterestRates).MoveNext
        Next iIndex
    End If

    ' We only have the rate the product is stored against, but BaseRateSet promotions are stored using the Base Rate Set
    ' and Start Date. This method will use the Base Rate Set to find the latest Rate Set for that Set Number, retrieve the
    ' date and see if there are any outstanding promotions.
    'bDependencyFound = CheckBaseRateSetFromRate(sBaseRateSet, clsUpdateDetails, colUpdateDetails) Or bDependencyFound
        
    bDependencyFound = GetObjectDependencies(BASE_RATES, colUpdateDetails, colObjectList) Or bDependencyFound
    'BS BM0500 End 10/04/03

    CheckProductDependencies = bDependencyFound
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Private Function GetDistributionChannelDependencies(colChannels As Collection, colUpdateDetails As Collection) As Boolean
    On Error GoTo Failed
    Dim colMatchValues  As Collection
    Dim colObjectList As Collection
    Dim clsChannelUpdateDetails As SupervisorUpdateDetails
    Dim sChannelID As Variant
    Dim bDependencyFound As Boolean
    
    bDependencyFound = False
    
    Set clsChannelUpdateDetails = New SupervisorUpdateDetails
    Set colObjectList = New Collection
    
    For Each sChannelID In colChannels
        Set colMatchValues = New Collection
        colMatchValues.Add sChannelID
        colObjectList.Add colMatchValues

        clsChannelUpdateDetails.SetKeyMatchValues colMatchValues
        bDependencyFound = CheckDistributionChannels(clsChannelUpdateDetails, colUpdateDetails) Or bDependencyFound

        bDependencyFound = GetObjectDependencies(DISTRIBUTION_CHANNELS, colUpdateDetails, colObjectList) Or bDependencyFound
    Next
    
    GetDistributionChannelDependencies = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Function GetStageDependencies(colActivity As Collection, colUpdateDetails As Collection) As Boolean
    On Error GoTo Failed
    Dim colMatchValues  As Collection
    Dim colObjectList As Collection
    Dim clsActivityUpdateDetails As SupervisorUpdateDetails
    Dim sActivityID As Variant
    Dim bDependencyFound As Boolean
    
    bDependencyFound = False
    
    Set clsActivityUpdateDetails = New SupervisorUpdateDetails
    Set colObjectList = New Collection
    
    For Each sActivityID In colActivity
        Set colMatchValues = New Collection
        colMatchValues.Add sActivityID
        colObjectList.Add colMatchValues

        clsActivityUpdateDetails.SetKeyMatchValues colMatchValues
        bDependencyFound = GetObjectDependencies(TASK_MANAGEMENT_ACTIVITIES, colUpdateDetails, colObjectList) Or bDependencyFound
    Next
    
    GetStageDependencies = bDependencyFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckContactComboDependencies
' Description   : Ensure the combogroups required for contact details aren't pending promotion.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckContactComboDependencies(ByRef colUpdateDetails As Collection) As Boolean

    Dim bDependencyFound As Boolean
    Dim colObjectList As Collection
    Dim colMatchValues As Collection
    
    'Create a list to hold key-collections.
    Set colObjectList = New Collection
    
    'ContactType.
    Set colMatchValues = New Collection
    colMatchValues.Add "ContactType"
    colObjectList.Add colMatchValues
    
    'ContactTitle.
    Set colMatchValues = New Collection
    colMatchValues.Add "ContactTitle"
    colObjectList.Add colMatchValues
    
    'TelephoneUsage.
    Set colMatchValues = New Collection
    colMatchValues.Add "ContactTelephoneUsage"
    colObjectList.Add colMatchValues
    
    'Country.
    Set colMatchValues = New Collection
    colMatchValues.Add "Country"
    colObjectList.Add colMatchValues
    
    'Check for these combogroups.
    bDependencyFound = GetObjectDependencies(COMBOBOX_ENTRIES, colUpdateDetails, colObjectList)
    
    CheckContactComboDependencies = bDependencyFound

End Function


Private Function ExectuteDependencies(colUpdateDetails As Collection, _
                                      clsCurrentObject As SupervisorUpdateDetails, _
                                      sTargetDatabase As String) As Boolean
    On Error GoTo Failed
    Dim colUpdateAreas As Collection
    Dim sTmp As Variant
    Dim sObjectName As String
    Dim nResponse As Integer
    Dim sDependencyList As String
    Dim nThisItem As Integer
    Dim clsUpdateDetails As SupervisorUpdateDetails
    Dim clsTableAccess As TableAccess
    Dim bPromote As Boolean
    'BS BM0500 10/04/03
    Dim colUpdDtl As Collection
        
    'g_clsDataAccess.BeginTrans sTargetDatabase
    Set colUpdateAreas = New Collection

    sTargetDatabase = GetTargetDatabaseName()
    bPromote = True
    
    If colUpdateDetails.Count > 0 Then
        ' Remove duplicates
        For Each clsUpdateDetails In colUpdateDetails
            sObjectName = clsUpdateDetails.GetObjectName()
            On Error Resume Next
            sTmp = colUpdateAreas(sObjectName)
            
            If Err.Number <> 0 Then
                colUpdateAreas.Add sObjectName, sObjectName
            End If
        Next
        
        On Error GoTo Failed
        
        ' Produce a list
        For nThisItem = 1 To colUpdateAreas.Count
            sDependencyList = sDependencyList + colUpdateAreas(nThisItem)
            If nThisItem < colUpdateAreas.Count Then
                sDependencyList = sDependencyList + ", "
            End If
            
        Next
        WriteLogEvent "ExecuteDependencies - " & sDependencyList
        
        'DB BM0010 07/01/03
        'nResponse = MsgBox("The following dependencies exist: " + sDependencyList + ". Promote all dependencies first?", vbOKCancel + vbQuestion)
        nResponse = MsgBox("The following dependencies exist for " & clsCurrentObject.GetObjectName & " : " + sDependencyList + ". Promote all dependencies first?", vbOKCancel + vbQuestion)
        'DB End
        
        If nResponse = vbOK Then
            Dim colKeyMatchValues As Collection
            WriteLogEvent "User said Yes"
            For Each clsUpdateDetails In colUpdateDetails
                sObjectName = clsUpdateDetails.GetObjectName()
                Set clsTableAccess = g_clsHandleUpdates.GetObjectTableClass(sObjectName)
                Set colKeyMatchValues = clsUpdateDetails.GetKeyMatchValues()
                clsTableAccess.SetKeyMatchValues colKeyMatchValues
                
                'BS BM0500 10/04/03
                'colUpdateDetails gets set to new collection in Promote so when processing returns
                'here, Next finds nothing.
                Set colUpdDtl = New Collection
                'Promote clsUpdateDetails, colUpdateDetails
                Promote clsUpdateDetails, colUpdDtl
                'BS BM0500 End 10/04/03
            Next
        Else
            WriteLogEvent "User said No"
            bPromote = False
        End If
    End If
    
    WriteLogEvent "ExectuteDependencies " & bPromote
    
    ExectuteDependencies = bPromote
    Exit Function
Failed:
    g_clsErrorHandling.SaveError
    g_clsErrorHandling.RaiseError
End Function
Public Function GetObjectDependencies(sObjectToCheck As String, colUpdateDetails As Collection, colObjectList As Collection) As Boolean
    
    ' Are there any combo updates outstanding?
    Dim lIndex As Long
    Dim sField As String
    Dim sObjectName As String
    Dim sObjectKey As Variant
    Dim colItemList As Collection
    Dim clsUpdateDetails As SupervisorUpdateDetails
    Dim sUpdateGroup As Variant
    Dim colMatchValues As Collection
    Dim colObjectMatchValues As Collection
    Dim bUpdateFound As Boolean
    Dim bUpdatePending As Boolean
    Dim sInstanceID As String
    
    
    On Error GoTo Failed
    bUpdateFound = False
    Set colItemList = New Collection
    'Loop through all object areas that are outstanding.
    For lIndex = 0 To cboAvailableUpdates.ListCount - 1
        sField = m_clsSupervisorUpdate.GetObjectNameField()
        sObjectName = cboAvailableUpdates.List(lIndex)
        
        'DB SYS5457 06/01/03 - If we are promoting Intermediaries, we need to check the various
        ' different types of Intermediary there may be. The object name passed in will be
        ' the individual type of Intermediary, such as Individual or Lead agent. For this
        ' process to work, we need to change the type to the generic "Intermediary" as
        ' all four types work in the same way.
        If sObjectToCheck = INTERMEDIARIES Then
            GetIntermediaryDependency sObjectName
        End If
        'DB End
                
        If sObjectName = sObjectToCheck Then
            ' Do we have any entries that match the list passed in?
            GetItemList sObjectName, colItemList

            ' Loop through all oustanding combo update requests
            For Each clsUpdateDetails In colItemList
                Set colMatchValues = clsUpdateDetails.GetKeyMatchValues()
                
                ' Then loop through all key values for each update request
                For Each sUpdateGroup In colMatchValues
                    bUpdatePending = False
                    
                    For Each colObjectMatchValues In colObjectList
                        For Each sObjectKey In colObjectMatchValues
                            'Ensure we can handle binary GUIDs.
                            If TypeName(sObjectKey) = "Byte()" Then
                                sObjectKey = g_clsSQLAssistSP.ByteArrayToGuidString(CStr(sObjectKey))
                            End If
                            
                            'Ensure we can handle binary GUIDs.
                            If TypeName(sUpdateGroup) = "Byte()" Then
                                sUpdateGroup = g_clsSQLAssistSP.ByteArrayToGuidString(CStr(sUpdateGroup))
                            End If
                            
                            If sObjectKey = sUpdateGroup Then
                                bUpdatePending = True
                                bUpdateFound = True
                            Else
                                bUpdatePending = False
                            End If
                            
                            If bUpdatePending Then
                                clsUpdateDetails.SetObjectName sObjectToCheck
                                sInstanceID = clsUpdateDetails.GetObjectID()
                                
                                On Error Resume Next
                                colUpdateDetails.Add clsUpdateDetails, sInstanceID
                                Err.Clear
                                
                                On Error GoTo Failed
                                Exit For
                            End If
                        Next
                    Next
                Next
            Next
        End If
    Next lIndex
    
    GetObjectDependencies = bUpdateFound
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function CheckDependencies(clsUpdateDetails As SupervisorUpdateDetails, colUpdateDetails As Collection) As Boolean
    Dim sObject As String
    Dim bDependencyFound As Boolean
    Dim clsDependency As ObjectDependenciesCS
    
    Set clsDependency = New ObjectDependenciesCS
    
    On Error GoTo Failed
    
    bDependencyFound = False
    sObject = clsUpdateDetails.GetObjectName()
    
    Select Case sObject
        Case MORTGAGE_PRODUCTS
            bDependencyFound = CheckProductDependencies(clsUpdateDetails, colUpdateDetails)
        
        Case DISTRIBUTION_CHANNELS
            bDependencyFound = CheckDistributionChannels(clsUpdateDetails, colUpdateDetails)
        
        Case PAYMENT_PROTECTION_PRODUCTS
            bDependencyFound = CheckPayProtProducts(clsUpdateDetails, colUpdateDetails)
    
        Case PAYMENT_PROTECTION_RATES
            bDependencyFound = CheckPayProtRates(clsUpdateDetails, colUpdateDetails)
        
        Case DEPARTMENTS
            bDependencyFound = CheckDepartments(clsUpdateDetails, colUpdateDetails)
    
        Case UNITS
            bDependencyFound = CheckUnit(clsUpdateDetails, colUpdateDetails)
    
        Case USERS
            bDependencyFound = CheckUsers(clsUpdateDetails, colUpdateDetails)
        
        Case TASK_MANAGEMENT_STAGES
            bDependencyFound = CheckStage(clsUpdateDetails, colUpdateDetails)
        
        'DB SYS5457 06/01/03 - Instead of just one type of Intermediary, we now have four.
        Case INDIVIDUAL_DESCRIPTION, LEADAGENT_DESCRIPTION, INTERMEDIARY_COMPANY_DESCRIPTION, ADMIN_CENTRE_DESCRIPTION
            bDependencyFound = CheckIntermediary(clsUpdateDetails, colUpdateDetails)
        'DB End
        Case BASE_RATE
            bDependencyFound = CheckBaseRate(clsUpdateDetails, colUpdateDetails)
            
        Case BASE_RATES
            bDependencyFound = CheckBaseRateSets(clsUpdateDetails, colUpdateDetails)
            
        Case INCOME_FACTORS
            bDependencyFound = CheckIncomeFactors(clsUpdateDetails, colUpdateDetails)
            
        'BS BM0498 28/04/03
        Case TASK_MANAGEMENT_TASKS
            bDependencyFound = CheckTask(clsUpdateDetails, colUpdateDetails)
        'BS BM0498 End 28/04/03
        Case PACKAGER
            bDependencyFound = CheckPackager(clsUpdateDetails, colUpdateDetails)
        Case INTRODUCER
            bDependencyFound = CheckIntroducers(clsUpdateDetails, colUpdateDetails)
' TW 18/11/2006 EP2_132
        Case FIRMPERMISSIONS
            bDependencyFound = CheckFirmPermissions(clsUpdateDetails, colUpdateDetails)
        Case INTRODUCERFIRM
            bDependencyFound = CheckIntroducerFirm(clsUpdateDetails, colUpdateDetails)
        Case NETWORKASSOCIATIONS
            bDependencyFound = CheckFirmClubNetworkAssociation(clsUpdateDetails, colUpdateDetails)
        Case FIRMTRADINGNAME
            bDependencyFound = CheckFirmTradingName(clsUpdateDetails, colUpdateDetails)
        Case PRINCIPALFIRMPACKAGER
            bDependencyFound = CheckPrincipalFirm(clsUpdateDetails, colUpdateDetails)
        Case ARFIRMS
            bDependencyFound = CheckARFirm(clsUpdateDetails, colUpdateDetails)
' TW 18/11/2006 EP2_132 End
' TW 23/11/2006 EP2_172
        Case EXCLUSIVEPRODUCTS
            bDependencyFound = CheckExclusiveProducts(clsUpdateDetails, colUpdateDetails)
' TW 23/11/2006 EP2_172 End
' TW 05/02/2007 EP2_706
        Case APPOINTMENTS
            bDependencyFound = CheckAppointments(clsUpdateDetails, colUpdateDetails)
' TW 05/02/2007 EP2_706 End
    End Select
    
    WriteLogEvent "CheckDependencies - Dependency found = " & bDependencyFound
    
    ' DJP SYS4121 Check for any extra client variant dependencies
        
    bDependencyFound = clsDependency.CheckDependencies(clsUpdateDetails, colUpdateDetails) Or bDependencyFound
    
    CheckDependencies = bDependencyFound
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


Private Sub ResetUpdates(Optional sUserID As String)
    On Error GoTo Failed
    Dim sMessage As String
    Dim nResponse As Integer
    
    sMessage = "Remove ALL Update information"
    If Len(sUserID) > 0 Then
        sMessage = sMessage + " for User " + sUserID
    End If
    
    sMessage = sMessage + "?"
    nResponse = MsgBox(sMessage, vbYesNo + vbQuestion)
    
    If nResponse = vbYes Then
        BeginWaitCursor
        
        m_clsSupervisorUpdate.ResetAllUpdates sUserID
        ClearUpdates
    
        EndWaitCursor
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub cmdReadLog_Click()
    On Error GoTo Failed
    
    frmText.Show vbModal, Me
    Unload frmText
    
    Exit Sub
Failed:
End Sub

Private Sub cmdResetForUser_Click()
    On Error GoTo Failed
    If MsgBox("Clearing the user history will potentially cause problems with promotions. Continue ?", vbYesNo + vbExclamation) = vbYes Then
        ResetUpdates g_sSupervisorUser
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdReset_Click()
    On Error GoTo Failed
    If MsgBox("Clearing the history will potentially cause problems with promotions. Continue ?", vbYesNo + vbExclamation) = vbYes Then
        ResetUpdates
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub ClearUpdates()
    PopulateAvailableUpdates
    PopulateUpdateItemDetails
End Sub


Private Sub Form_Initialize()
    Set m_clsSupervisorUpdate = New SupervisorUpdateTable
    Set m_clsPromote = New SupervisorObjectCopy
End Sub
Private Sub Form_Load()
    On Error GoTo Failed
    SetReturnCode MSGFailure

    ' BM0423 reset the promotion user to the logged on user
    m_sUserLoggedOn = g_sSupervisorUser
    
    m_sSrcServiceName = g_clsDataAccess.GetDatabaseName()
    m_sSrcUserID = g_clsDataAccess.GetDatabaseUserID()

    InitialiseFields
    
    g_clsFormProcessing.PopulateAvailableTargets cboTarget, m_sSrcServiceName, m_sSrcUserID
    
    If cboTarget.ListCount = 0 Then
        MsgBox "At least one other database has to be defined. See Tools -> Database Options to add other databases", vbCritical
        cboTarget.Enabled = False
        cmdReset.Enabled = False
        cmdResetForUser.Enabled = False
    End If
    
    If Len(g_sSupervisorUser) = 0 Then
        lblTitle = "Updates are not possible as you have not logged in to Supervisor"
        cboAvailableUpdates.Enabled = False
        cboTarget.Enabled = False
        cmdResetForUser.Enabled = False
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub CheckUpdatesAvailable(sDatabaseKey As String)
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsSupervisorUpdate
        
    If Not g_clsDataAccess.DoesTableExist(clsTableAccess.GetTable(), sDatabaseKey) Then
        g_clsErrorHandling.RaiseError errGeneralError, "Updating from the current active database is not allowed"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetDefaultProcurationFeeHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Submission Route"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Scheme Description"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Fee Rate"
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetLoanAmountAdjustmentFeeHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Product Category"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Maximum Loan"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Fee Rate"
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetLTVAmountAdjustmentFeeHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Product Category"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Maximum LTV"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Fee Rate"
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetAdminFeeHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Set Number"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Start Date"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetValuationFeeHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Set Number"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Start Date"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetFixParamHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Parameter Name"
    
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Start Date"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetThirdPartyHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 40
    lvHeaders.sName = "Third party Name and Type"
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetErrorMessageHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 30
    lvHeaders.sName = "Error Message Number"
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetWorkingHoursHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 30
    lvHeaders.sName = "Working Hours Type"
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetRegionHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 30
    lvHeaders.sName = "Region ID"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetBandedParamHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 30
    lvHeaders.sName = "Parameter Name"
    
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Start Date"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetCompetencyHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Competency"
    
    colHeaders.Add lvHeaders
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetUserHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "User Code"
    
    colHeaders.Add lvHeaders
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetUnitHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Unit Code"
    
    colHeaders.Add lvHeaders
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetDepartmentHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Department Code"
    
    colHeaders.Add lvHeaders
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetPayProtProductsHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Product Code"
    
    colHeaders.Add lvHeaders
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetPayProtRatesHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Rate Number"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Start Date"
    
    colHeaders.Add lvHeaders
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetLifeCoverHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Life Cover Rates Number"
   
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetBAndCHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Product Number"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetCountryHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Country Number"
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Country Name"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetBaseRateHeaders(ByRef lvHeaders As listViewAccess, ByRef colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Rate ID"
    colHeaders.Add lvHeaders

    lvHeaders.nWidth = 20
    lvHeaders.sName = "Start Date"

    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetBaseRateSetHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Set Number"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Start Date"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetLenderHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Lender Code"
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetDistChannelHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Channel ID"
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetComboGroupHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Group Name"
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetProductHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Product Code"
    colHeaders.Add lvHeaders
    
    lvHeaders.nWidth = 18
    lvHeaders.sName = "Product Start Date"
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'MAR967
Private Sub SetPrintingPackHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Pack ID"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub GetSelectedObjectDetails(clsSupervisorUpdateDetails As SupervisorUpdateDetails, Optional nIndex As Integer = -1)
    On Error GoTo Failed
    Dim colLine As Collection
    
    Set colLine = lvUpdates.GetLine(nIndex, clsSupervisorUpdateDetails)
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function GetSelectedObjectKeyValues() As Collection
    On Error GoTo Failed
    Dim colLine As Collection
    Dim colValues As Collection
    Dim sFunctionName As String
    Dim clsSupervisorUpdate As SupervisorUpdateDetails
    sFunctionName = "GetSelectedObjectKeyValues"

    Set colLine = lvUpdates.GetLine(, clsSupervisorUpdate)
    Set colValues = clsSupervisorUpdate.GetKeyMatchValues()
    If Not colValues Is Nothing Then
        If colValues.Count > 0 Then
            Set GetSelectedObjectKeyValues = colValues
        Else
            g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": No key values found"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": No key values found"
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Function GetSelectedObjectKeys(Optional sObjectName As String = "") As Collection
    On Error GoTo Failed
    Dim colCopyKeys As New Collection
    Dim colKeys As Collection
    Dim clsTableAccess As TableAccess
    Dim sFunctionName As String
    
    sFunctionName = "GetSelectedObjectKeys"
    
    
    If Len(sObjectName) = 0 Then
        sObjectName = GetSelectedObjectName()
    End If
    
    Set clsTableAccess = g_clsHandleUpdates.GetObjectTableClass(sObjectName)
    Set colKeys = clsTableAccess.GetKeyMatchFields()
    
    If colKeys.Count > 0 Then
        Set colCopyKeys = colKeys
    Else
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Key count is 0"
    End If
    
    Set GetSelectedObjectKeys = colCopyKeys
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetupListViewHeaders
' Description   : Based upon the current item type selection, setup the column
'                 headers in the main listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupListViewHeaders(sObjectName As String)
    
    Dim sFunctionName As String
    Dim colHeaders As Collection
    Dim lvHeaders As listViewAccess
    
    On Error GoTo Failed
    
    Set colHeaders = New Collection
    
    sFunctionName = "SetupListViewHeaders"
    lvUpdates.ListItems.Clear
    lvUpdates.ColumnHeaders.Clear

    
    ' Object specific headers
    If Len(sObjectName) > 0 Then
        Select Case sObjectName
            Case MORTGAGE_PRODUCTS
                SetProductHeaders lvHeaders, colHeaders
            
            Case COMBOBOX_ENTRIES
                SetComboGroupHeaders lvHeaders, colHeaders
            
            Case DISTRIBUTION_CHANNELS
                SetDistChannelHeaders lvHeaders, colHeaders
            
            Case LENDERS
                SetLenderHeaders lvHeaders, colHeaders
            
' TW 09/10/2006 EP2_7
            Case ADDITIONAL_BORROWING_FEES, CREDIT_LIMIT_INCREASE_FEES
                SetAdminFeeHeaders lvHeaders, colHeaders
' TW 09/10/2006 EP2_7 End
' TW 11/12/2006 EP2_20
            Case TRANSFER_OF_EQUITY_FEES
                SetAdminFeeHeaders lvHeaders, colHeaders
' TW 11/12/2006 EP2_20 End
' TW 14/12/2006 EP2_518
            Case DEFAULT_PROCURATION_FEES
                SetDefaultProcurationFeeHeaders lvHeaders, colHeaders
            Case LOAN_AMOUNT_ADJUSTMENTS
                SetLoanAmountAdjustmentFeeHeaders lvHeaders, colHeaders
            Case LTV_AMOUNT_ADJUSTMENTS
                SetLTVAmountAdjustmentFeeHeaders lvHeaders, colHeaders
' TW 14/12/2006 EP2_518 End
    
' TW 17/10/2006 EP2_15
'            Case ASSOCIATIONS, CLUBS
'                SetMortgageClubAssociationHeaders lvHeaders, colHeaders
'
'            Case PACKAGERS
'                SetPackagerHeaders lvHeaders, colHeaders
' TW 18/11/2006 EP2_132
'            Case CLUBS, ASSOCIATIONS, PRINCIPALS, PACKAGERS, ARFIRMS, INDIVIDUALS & "_" & PACKAGERS, ARBROKERS, DABROKERS
' TW 23/11/2006 EP2_172
'            Case CLUBSANDASSOCIATIONS, PRINCIPALFIRMPACKAGER, PACKAGERS, ARFIRMS, INDIVIDUALS & "_" & PACKAGERS, ARBROKERS, DABROKERS, INTRODUCER, FIRMTRADINGNAME, FIRMPERMISSIONS, INTRODUCERFIRM, NETWORKASSOCIATIONS
            Case CLUBSANDASSOCIATIONS, PRINCIPALFIRMPACKAGER, PACKAGERS, ARFIRMS, INDIVIDUALS & "_" & PACKAGERS, ARBROKERS, DABROKERS, INTRODUCER, FIRMTRADINGNAME, FIRMPERMISSIONS, INTRODUCERFIRM, NETWORKASSOCIATIONS, EXCLUSIVEPRODUCTS
' TW 23/11/2006 EP2_172 End
' TW 18/11/2006 EP2_132 End
                SetPackagerHeaders lvHeaders, colHeaders
' TW 18/11/2006 EP2_132
'            Case INTRODUCER
'                SetPackagerHeaders lvHeaders, colHeaders
' TW 18/11/2006 EP2_132 End
' TW 17/10/2006 EP2_15 End
            
            Case ADMIN_FEES
                SetAdminFeeHeaders lvHeaders, colHeaders
            
            Case VALUATION_FEES
                SetValuationFeeHeaders lvHeaders, colHeaders
            
            Case BASE_RATE
                SetBaseRateHeaders lvHeaders, colHeaders
            
            Case BASE_RATES
                SetBaseRateSetHeaders lvHeaders, colHeaders
                            
            Case COUNTRIES
                SetCountryHeaders lvHeaders, colHeaders
            
            Case BUILDINGS_AND_CONTENTS_PRODUCTS
                SetBAndCHeaders lvHeaders, colHeaders
            
            Case LIFE_COVER_RATES
                SetLifeCoverHeaders lvHeaders, colHeaders
            
            Case PAYMENT_PROTECTION_RATES
                SetPayProtRatesHeaders lvHeaders, colHeaders
            
            Case PAYMENT_PROTECTION_PRODUCTS
                SetPayProtProductsHeaders lvHeaders, colHeaders
            
            Case DEPARTMENTS
                SetDepartmentHeaders lvHeaders, colHeaders
            
            Case UNITS
                SetUnitHeaders lvHeaders, colHeaders
            
            Case USERS
                SetUserHeaders lvHeaders, colHeaders
            
            Case COMPETENCIES
                SetCompetencyHeaders lvHeaders, colHeaders
            
            Case GLOBAL_PARAM_FIXED
                SetFixParamHeaders lvHeaders, colHeaders
                
            Case GLOBAL_PARAM_BANDED
                SetBandedParamHeaders lvHeaders, colHeaders
            
            Case REGIONS
                SetRegionHeaders lvHeaders, colHeaders
            
            Case WORKING_HOURS
                SetWorkingHoursHeaders lvHeaders, colHeaders
            
            Case ERROR_MESSAGES
                SetErrorMessageHeaders lvHeaders, colHeaders
            
            Case NAMES_AND_ADDRESSES
                SetThirdPartyHeaders lvHeaders, colHeaders
                
            Case ADDITIONAL_QUESTIONS
                SetQuestionHeaders lvHeaders, colHeaders
                
            Case CONDITIONS
                SetConditionHeaders lvHeaders, colHeaders
                
            Case PRINTING_TEMPLATE
                SetPrintingHeaders lvHeaders, colHeaders
            
            'MAR45 GHun
            Case PRINTING_DOCUMENT
                SetPrintingDocsHeaders lvHeaders, colHeaders
            'MAR45 End
                
            Case TASK_MANAGEMENT_TASKS
                SetTaskHeaders lvHeaders, colHeaders
                
            Case TASK_MANAGEMENT_STAGES
                SetStageHeaders lvHeaders, colHeaders
                
            Case TASK_MANAGEMENT_ACTIVITIES
                SetActivityHeaders lvHeaders, colHeaders
                
            Case BUSINESS_GROUPS
                SetBusinessGroupHeaders lvHeaders, colHeaders
                
            Case CURRENCIES
                SetCurrencyHeaders lvHeaders, colHeaders
                
            'DB SYS5457 06/01/03 - Instead of just one type of Intermediary, we now have four.
            Case LEADAGENT_DESCRIPTION, INDIVIDUAL_DESCRIPTION, INTERMEDIARY_COMPANY_DESCRIPTION, ADMIN_CENTRE_DESCRIPTION
                SetIntermediaryHeaders lvHeaders, colHeaders
            'DB End
            Case INCOME_FACTORS
                SetIncomeFactorHeaders lvHeaders, colHeaders
                 
            Case PRODUCT_SWITCH_FEESETS, INSURANCE_ADMIN_FEESETS, TT_FEESETS, RENTAL_INCOME_RATES
                '*=[MC]BMIDS763 PSF,IAF,TTFee sets require same headers as do rental income rates - JD BMIDS765
                ' as adminfeesets, to avoid duplication code, same routine has been called here.
                'NOTE: This call is for new 3 products
                SetAdminFeeHeaders lvHeaders, colHeaders
            
            'MAR967
            Case PRINTING_PACK
                SetPrintingPackHeaders lvHeaders, colHeaders
' TW 05/02/2007 EP2_706
            Case APPOINTMENTS
                SetPackagerHeaders lvHeaders, colHeaders
' TW 05/02/2007 EP2_706 End

            Case Else
                ' DJP SYS4142
                Dim bSuccess As Boolean
                Dim clsPromote As PromoteCS
                
                Set clsPromote = New PromoteCS
                
                bSuccess = clsPromote.SetupListViewHeaders(sObjectName, lvHeaders, colHeaders)
                
                If Not bSuccess Then
                    g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Unknown object :" + sObjectName
                End If
        End Select
        
        ' Common to all
        lvHeaders.nWidth = 27
        lvHeaders.sName = "Last Updated"
        colHeaders.Add lvHeaders
        
        lvHeaders.nWidth = 20
        lvHeaders.sName = "Updated by"
        colHeaders.Add lvHeaders
        
        lvHeaders.nWidth = 20
        lvHeaders.sName = "Update type"
        colHeaders.Add lvHeaders
        
        lvUpdates.AddHeadings colHeaders
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function GetSelectedObjectName() As String
    On Error GoTo Failed
    Dim sFunctionName As String
    
    sFunctionName = "GetSelectedObjectName"
    GetSelectedObjectName = Me.cboAvailableUpdates.SelText
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub GetItemList(sObjectName As String, colItemList As Collection)
    On Error GoTo Failed
    Dim bWildCard As Boolean 'DB SYS5457 06/01/03
    Dim sFunctionName As String
    Dim sTargetService As String
    Dim sTartgetUserID As String
    Dim clsTableAccess As TableAccess
    Dim colListViewKeys As Collection
    Dim rs As ADODB.Recordset
    Dim sKey As String
    ' DJP SQL Server port
    Dim sVal As Variant
    Dim sInstance As String
    Dim colTmp As Collection
    Dim sCreationDate As String
    Dim sDatabaseUserID As String
    Dim colKeyValues As Collection
    Dim colObjectKeys As Collection
    Dim sTmp As Variant
    
    sFunctionName = "PopulateUpdateItemDetails"
    bWildCard = False 'DB SYS5457 06/01/03
    Set colListViewKeys = New Collection
    Set colKeyValues = New Collection
    
    ' What are the keys called?
    Set colObjectKeys = GetSelectedObjectKeys(sObjectName)
    
    GetTargetDatabase sTargetService, sTartgetUserID
    
    'DB SYS5457 06/01/03 - if promoting Intermediaries we need to do a partial search to match all four types.
    If sObjectName = INTERMEDIARIES Then
        bWildCard = True
    End If
    
    m_clsSupervisorUpdate.FindUpdateItemDetails sObjectName, m_sSrcServiceName, m_sSrcUserID, sTargetService, sTartgetUserID, bWildCard
    
    'm_clsSupervisorUpdate.FindUpdateItemDetails sObjectName, m_sSrcServiceName, m_sSrcUserID, sTargetService, sTartgetUserID
    'DB End
    
    Set clsTableAccess = m_clsSupervisorUpdate
    
    If clsTableAccess.RecordCount > 0 Then
    
        ' Copy the object keys for use later
        For Each sTmp In colObjectKeys
            colListViewKeys.Add sTmp
        Next
        
        colListViewKeys.Add "CREATIONDATE"
        colListViewKeys.Add "USERID"
        colListViewKeys.Add "SUPERVISOROBJECTINSTANCEID"
    
        Set rs = clsTableAccess.GetRecordSet()
        
        clsTableAccess.MoveFirst
        
        Dim colRows As New Collection
        Dim colColumns As Collection
        
        Do
            
            'nType = rs("SUPERVISOROBJECTVALUE").Type
            
            ' DJP SQL Server port
            If rs.fields.Item("SUPERVISOROBJECTVALUETYPE").Value = "Byte()" Then
                sVal = g_clsSQLAssistSP.GuidStringToByteArray(rs("SUPERVISOROBJECTVALUE"))
            Else
                sVal = rs("SUPERVISOROBJECTVALUE")
            End If
            
            sKey = rs("SUPERVISOROBJECTKEY")
            sCreationDate = rs("CREATIONDATE")
            sDatabaseUserID = rs("USERID")
            
            ' DJP SQL Server port
            sInstance = g_clsSQLAssistSP.GuidToString(CStr(rs("SUPERVISOROBJECTINSTANCEID")))
            
            ' Does the instance exist in our row collection?
            On Error Resume Next
            Set colTmp = colRows(sInstance)
            
            If Err.Number <> 0 Then
                Set colTmp = New Collection
                colTmp.Add sVal, sKey
                colTmp.Add sCreationDate, "CREATIONDATE"
                colTmp.Add sDatabaseUserID, "USERID"
                colTmp.Add sInstance, "SUPERVISOROBJECTINSTANCEID"
                colRows.Add colTmp, sInstance
            Else
                ' Exists
                colTmp.Add sVal, sKey
                colTmp.Add sCreationDate, "CREATIONDATE"
                colTmp.Add sDatabaseUserID, "USERID"
                colTmp.Add sInstance, "SUPERVISOROBJECTINSTANCEID"
            End If
            On Error GoTo Failed
            
            clsTableAccess.MoveNext
        Loop While rs.EOF = False
        
        Dim sThisKey As Variant
        Dim sThisItem As Variant
        Dim nThisItem As Integer
        Dim updateDetails As SupervisorUpdateDetails
        
        For Each colColumns In colRows
            Set colKeyValues = New Collection
            
            sInstance = ""
            sCreationDate = ""
            nThisItem = 1
            Set updateDetails = New SupervisorUpdateDetails
            
            For Each sThisKey In colListViewKeys
                sThisItem = colColumns(sThisKey)
                                
                ' if the current item index is more than the key count, it's not a key, so
                ' don't add it
                If nThisItem <= colObjectKeys.Count Then
                    colKeyValues.Add sThisItem
                End If
                                
                If sThisKey = "CREATIONDATE" Then
                    sCreationDate = sThisItem
                End If
                
                If sThisKey = "SUPERVISOROBJECTINSTANCEID" Then
                    sInstance = sThisItem
                End If

                nThisItem = nThisItem + 1
            Next
            
            ' Update details
            updateDetails.SetKeyMatchValues colKeyValues
            updateDetails.SetCreationDate sCreationDate
            updateDetails.SetObjectID sInstance
            
            colItemList.Add updateDetails 'colKeyValues
        Next
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Function FindKeyValue(sThisKey As Variant, clsTableAccess As TableAccess, sKeyVal As Variant) As Boolean
    On Error GoTo Failed
    
    Dim rs As ADODB.Recordset
    Dim sVal As Variant
    Dim sKey As Variant
    Dim bFound As Boolean
    
    clsTableAccess.MoveFirst
    Set rs = clsTableAccess.GetRecordSet()
    bFound = False
    
    Do While Not rs.EOF And Not bFound
        sVal = rs("SUPERVISOROBJECTVALUE")
        sKey = rs("SUPERVISOROBJECTKEY")
        
        If sKey = sThisKey Then
        
            If Not IsEmpty(sVal) And Not IsNull(sVal) Then
                sKeyVal = sVal
                bFound = True
            End If
        End If
        clsTableAccess.MoveNext
    Loop
        
    FindKeyValue = bFound
        
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateUpdateItemDetails
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateUpdateItemDetails()
    On Error GoTo Failed
    Dim sObjectName As String
    Dim sFunctionName As String
    Dim sTargetService As String
    Dim sTartgetUserID As String
    Dim clsTableAccess As TableAccess
    Dim colListViewKeys As Collection
    Dim rs As ADODB.Recordset
    Dim sKey As String
    ' DJP SQL Server port
    Dim vval As Variant
    Dim vValueType As Variant
    Dim sInstance As String
    Dim colTmp As Collection
    Dim sCreationDate As String
    Dim sDatabaseUserID As String
    Dim colKeyValues As Collection
    Dim colObjectKeys As Collection
    Dim sTmp As Variant
    Dim sDescription As Variant
    Dim sUserID As String
    Dim enumOperation As SupervisorPromoteType
    Dim sOperation As String
        
    sFunctionName = "PopulateUpdateItemDetails"
    
    BeginWaitCursor
    
    Set colListViewKeys = New Collection
    Set colKeyValues = New Collection
    
    sObjectName = GetSelectedObjectName()
    
    SetupListViewHeaders sObjectName
    
    If Len(sObjectName) > 0 Then
    
        ' What are the keys called?
        Set colObjectKeys = GetSelectedObjectKeys()
        
        
        GetTargetDatabase sTargetService, sTartgetUserID
        
        m_clsSupervisorUpdate.FindUpdateItemDetails sObjectName, m_sSrcServiceName, m_sSrcUserID, sTargetService, sTartgetUserID

        Set clsTableAccess = m_clsSupervisorUpdate
        
        If clsTableAccess.RecordCount > 0 Then
        
            ' Copy the object keys for use later
            For Each sTmp In colObjectKeys
                colListViewKeys.Add sTmp
            Next
            
            colListViewKeys.Add "CREATIONDATE"
            colListViewKeys.Add "DBUSERID"
            colListViewKeys.Add "SUPERVISOROBJECTINSTANCEID"
            colListViewKeys.Add "SUPERVISOROBJECTDESCRIPTION"
            colListViewKeys.Add "SUPERVISOROPERATION"
            
            Set rs = clsTableAccess.GetRecordSet()
            
            clsTableAccess.MoveFirst
            
            Dim colRows As New Collection
            Dim colColumns As Collection
            
            Do
                'We must check the datatype field (to be added) so that GUIDS
                'are not read-back as strings.
                vValueType = rs("SUPERVISOROBJECTVALUETYPE")

                'If the value type is specified.
                If IsNull(vValueType) = False Then
                    'If the key value is a byte array, convert the field value into one.
                    If CStr(vValueType) = "Byte()" Then
                        vval = g_clsSQLAssistSP.GuidStringToByteArray(rs("SUPERVISOROBJECTVALUE"))

                    'Otherwise, just get the value as a string.
                    Else
                        vval = rs("SUPERVISOROBJECTVALUE")
                    End If
                Else
                    vval = rs("SUPERVISOROBJECTVALUE")
                End If
                
                sKey = rs("SUPERVISOROBJECTKEY")
                sCreationDate = rs("CREATIONDATE")
                sDatabaseUserID = m_clsSupervisorUpdate.GetSupervisorUserID
                
                ' DJP SQL Server port
                sInstance = g_clsSQLAssistSP.GuidToString(CStr(rs("SUPERVISOROBJECTINSTANCEID")))
                sDescription = rs("SUPERVISOROBJECTDESCRIPTION")
                enumOperation = rs("SUPERVISOROPERATION")
                
                ' Does the instance exist in our row collection?
                On Error Resume Next
                Set colTmp = colRows(sInstance)
                
                If Err.Number <> 0 Then
                    Set colTmp = New Collection
                    colTmp.Add sDescription, "SUPERVISOROBJECTDESCRIPTION"
                    colTmp.Add vval, sKey
                    colTmp.Add sCreationDate, "CREATIONDATE"
                    colTmp.Add sDatabaseUserID, "DBUSERID"
                    colTmp.Add sInstance, "SUPERVISOROBJECTINSTANCEID"
                    colTmp.Add enumOperation, "SUPERVISOROPERATION"
                    colRows.Add colTmp, sInstance
                Else
                    ' Exists
                    colTmp.Add sDescription, "SUPERVISOROBJECTDESCRIPTION"
                    colTmp.Add vval, sKey
                    colTmp.Add sCreationDate, "CREATIONDATE"
                    colTmp.Add sDatabaseUserID, "DBUSERID"
                    colTmp.Add sInstance, "SUPERVISOROBJECTINSTANCEID"
                    colTmp.Add enumOperation, "SUPERVISOROPERATION"

                End If
                
                clsTableAccess.MoveNext
            Loop While rs.EOF = False
            
            Dim colLine As New Collection
            Dim sThisKey As Variant
            Dim sThisItem As Variant
            Dim nThisItem As Integer
            Dim colDescription As Collection
            Dim updateDetails As SupervisorUpdateDetails
                    
            Set colDescription = New Collection
            
            For Each colColumns In colRows
                Set colLine = New Collection
                Set colKeyValues = New Collection
                Set colDescription = New Collection
                
                sInstance = ""
                sCreationDate = ""
                sUserID = ""
                sDescription = ""
                
                nThisItem = 1
                Set updateDetails = New SupervisorUpdateDetails
    
                For Each sThisKey In colListViewKeys
                    sThisItem = colColumns(sThisKey)
                    
                    ' if the current item index is more than the key count, it's not a key, so
                    ' don't add it
                    If nThisItem <= colObjectKeys.Count Then
                        colKeyValues.Add sThisItem
                    End If
                                    
                    If sThisKey = "SUPERVISOROPERATION" Then
                        enumOperation = sThisItem
                    ElseIf sThisKey = "CREATIONDATE" Then
                        sCreationDate = sThisItem
                    ElseIf sThisKey = "DBUSERID" Then
                        sUserID = sThisItem
                    ElseIf sThisKey = "SUPERVISOROBJECTINSTANCEID" Then
                        sInstance = sThisItem
                    ElseIf sThisKey = "SUPERVISOROBJECTDESCRIPTION" And Len(sThisItem) > 0 Then
                        sDescription = sThisItem
                    ElseIf Len(sDescription) = 0 Then
                        If Not IsNull(sThisItem) Then
                            
                            If IsEmpty(sThisItem) Then
                                If Not (FindKeyValue(sThisKey, clsTableAccess, sThisItem)) Then
                                    g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Key " + sThisKey + " does not have a valid value"
                                End If
                            End If
                            
                            colDescription.Add sThisItem
                        End If
                    End If
    
                    nThisItem = nThisItem + 1
                Next
                If Len(sDescription) = 0 Then
                    For Each sDescription In colDescription
                        colLine.Add sDescription
                    Next
                Else
                    colLine.Add sDescription
                End If
                
                colLine.Add sCreationDate
                colLine.Add sUserID
                sOperation = GetOperation(enumOperation)
                colLine.Add sOperation
                
                ' Update details
                
                'sConcatenatedKey = GetConcatenatedKey(colKeyValues)
                On Error Resume Next
                
                'colListViewAdded.Add "any old thing", sConcatenatedKey
                
                If Err.Number = 0 Then
                    On Error GoTo Failed
                    updateDetails.SetKeyMatchValues colKeyValues
                    updateDetails.SetCreationDate sCreationDate
                    updateDetails.SetObjectID sInstance
                    updateDetails.SetObjectName sObjectName
                    updateDetails.SetOperation enumOperation
' TW 23/03/2007 EP2_1942
' Check that the number of columns defined in colLine = the number of column headers defined in lvUpdates.
' If these are not equal, there is a data error in the SUPERVISOROBJECTINSTANCEDET table.
' The following code is designed to prevent Supervisor from falling over when encountering this situation
' by displaying a 'Key Error' message in the promotions list to put the user on their guard.
                    Dim X As Integer
                    Dim TempCol As Collection
                    Select Case colLine.Count
                        Case lvUpdates.ColumnHeaders.Count 'OK
                        Case Is > lvUpdates.ColumnHeaders.Count ' Need to remove some columns and add in the 'Key Error' message
                            Set TempCol = New Collection
                            TempCol.Add "Key Error"
                            For X = lvUpdates.ColumnHeaders.Count + 1 To colLine.Count
                                TempCol.Add colLine(X)
                            Next X
                            Set colLine = TempCol
                        Case Else ' Need to add columns including the 'Key Error' message
                            Set TempCol = New Collection
                            TempCol.Add "Key Error"
                            For X = 2 To lvUpdates.ColumnHeaders.Count - colLine.Count
                                TempCol.Add "Unknown"
                            Next X
                            For X = 1 To colLine.Count
                                TempCol.Add colLine(X)
                            Next X
                            Set colLine = TempCol
                    End Select
                    If colLine.Count <> lvUpdates.ColumnHeaders.Count Then
                        Debug.Print "colKeyValues count : " & colKeyValues.Count
                        Debug.Print "lvUpdates.ColumnHeaders count : " & lvUpdates.ColumnHeaders.Count
                        Debug.Print "ColLine count : " & colLine.Count
                    End If
' TW 23/03/2007 EP2_1942 End
                    lvUpdates.AddLine colLine, updateDetails 'colKeyValues
                End If
                On Error GoTo Failed
            Next
        Else
            ' Repopulate the available updates
            cboAvailableUpdates.ListIndex = -1
            PopulateAvailableUpdates
        End If
    End If
    
    EndWaitCursor
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub
Private Function GetConcatenatedKey(colKeyValues As Collection) As String
    On Error GoTo Failed
    Dim sKey As Variant
    Dim sContatenatedKey As String
    
    For Each sKey In colKeyValues
        sContatenatedKey = sContatenatedKey + "," + sKey
    Next
    
    GetConcatenatedKey = sContatenatedKey
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Function GetOperation(enumOperation As SupervisorPromoteType) As String
    Dim sFunctionName As String
    
    sFunctionName = "GetOperation"
    
    Select Case enumOperation
        Case PromoteEdit
            GetOperation = "Edit"
        
        Case PromoteDelete
            GetOperation = "Delete"
        
        Case Else
            g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Unknown operation"
    End Select
End Function
Private Function GetTargetDatabaseName() As String
    On Error GoTo Failed
    Dim sKey As String
    Dim sFunctionName As String
    
    sFunctionName = "GetTargetDatabaseName"
    sKey = cboTarget.SelText
    
    GetTargetDatabaseName = sKey

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub GetTargetDatabase(sTargetService As String, sTargetUserID As String)
    On Error GoTo Failed
    Dim sKey As String
    Const strFunctionName As String = "GetTargetDatabase"
    Dim clsSupervisorConnection As SupervisorConnection
        
    sKey = GetTargetDatabaseName()
    
    If Len(sKey) > 0 Then
        g_clsDataAccess.GetSupervisorConnection clsSupervisorConnection, sKey
        
        sTargetService = clsSupervisorConnection.GetDatabaseName()
        sTargetUserID = clsSupervisorConnection.GetUserID()
        'SR 22/11/2007 : CORE00000313(VR 37)/MAR1968 - Allow UserID to be empty (when using integrated security)
        If Len(sTargetService) = 0 Then
            g_clsErrorHandling.RaiseError errGeneralError, strFunctionName + ": Service must be valid"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub InitialiseFields()
    On Error GoTo Failed
    Dim sName As String
    
    sName = g_clsDataAccess.GetActiveDatabaseName()
    
    CheckUpdatesAvailable sName
    
    If Len(sName) > 0 Then
        txtManagement(ACTIVE_DATABASE).Text = sName
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "No active database"
    End If
    
    cboAvailableUpdates.Enabled = False
    cmdPromote.Enabled = False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function


Private Sub cmdDone_Click()
    
    On Error GoTo Failed
    
    SetReturnCode MSGSuccess
    Hide
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Private Sub lvUpdates_ItemClick(ByVal Item As MSComctlLib.IListItem)
    cmdPromote.Enabled = True
End Sub


Private Sub SetQuestionHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Question Reference"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetCurrencyHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Currency"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetIntermediaryHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 45
    'DB SYS5457 06/01/03 Only display the name now, not name and type.
    lvHeaders.sName = "Intermediary Name"
    'lvHeaders.sName = "Intermediary Name and Type"
    'DB End
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetIncomeFactorHeaders(ByRef lvHeaders As listViewAccess, ByRef colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 75
    lvHeaders.sName = "Income Factor"
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetBusinessGroupHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Business Group"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetPrintingHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Printing Template ID"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'MAR45 GHun
Private Sub SetPrintingDocsHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "DPS Template ID"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'MAR45 End

Private Sub SetTaskHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Task ID"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetActivityHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Activity ID"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetStageHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Stage ID"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetConditionHeaders(lvHeaders As listViewAccess, colHeaders As Collection)
    On Error GoTo Failed
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Condition Reference"
    
    colHeaders.Add lvHeaders
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   CheckLenderDependencies
' Description   :   Check for client specific Lender dependencies. We need to check for Base Rates and
'                   Base Rate Sets.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckBaseRateSetFromRate(sBaseRateSet As String, clsUpdateDetails As SupervisorUpdateDetails, colUpdateDetails As Collection) As Boolean
    On Error GoTo Failed
    Dim sStartDate As String
    Dim colObjectList As Collection
    Dim colMatchValues As Collection
    Dim clsBaseRates As BaseRateTable
    Dim bDependencyFound As Boolean
    Dim clsBaseRateDetails As SupervisorUpdateDetails
    Dim colUpdateBaseRate As Collection
    
    bDependencyFound = False
    Set clsBaseRates = New BaseRateTable
    
    If Len(sBaseRateSet) > 0 Then
        ' Get the list of applicable Base Rates
        clsBaseRates.SetFindNewestBaseRateSet
        TableAccess(clsBaseRates).GetTableData POPULATE_FIRST_BAND
        
        ' There should only be one with the Base Rate we need
        TableAccess(clsBaseRates).ApplyFilter "BASERATESET = " & sBaseRateSet
        sStartDate = clsBaseRates.GetStartDate
        
        ' Base Rates
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        
        ' Setup the promotion we are looking for (Base Rate Band with Set number and Start Date)
        colMatchValues.Add sBaseRateSet
        colMatchValues.Add sStartDate
        colObjectList.Add colMatchValues
        
        Set colUpdateBaseRate = New Collection
        Set clsBaseRateDetails = New SupervisorUpdateDetails

        colUpdateBaseRate.Add clsBaseRateDetails
        clsBaseRateDetails.SetKeyMatchValues colMatchValues
        
        ' Now check the Base Rate itself
        bDependencyFound = frmManageUpdates.CheckBaseRateSets(clsBaseRateDetails, colUpdateDetails) Or bDependencyFound
        
        ' Check the dependencies of the Base Rate Band we have, but only if we found a dependency in the first place.
        bDependencyFound = frmManageUpdates.GetObjectDependencies(BASE_RATES, colUpdateDetails, colObjectList) Or bDependencyFound
        
    End If
    
    CheckBaseRateSetFromRate = bDependencyFound
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Public Function CheckMortgageProductCondition(sConditionReference As String, clsUpdateDetails As SupervisorUpdateDetails, colUpdateDetails As Collection) As Boolean
    On Error GoTo Failed
'    Dim sStartDate As String
'    Dim colObjectList As Collection
'    Dim colMatchValues As Collection
'    Dim clsBaseRates As BaseRateTable
'    Dim clsTableAccess As TableAccess
    Dim bDependencyFound As Boolean
'    Dim clsBaseRateDetails As SupervisorUpdateDetails
'    Dim colUpdateBaseRate As Collection
'
'    bDependencyFound = False
'    Set clsBaseRates = New BaseRateTable
'
    If Len(sConditionReference) > 0 Then
'        ' Get the list of applicable Base Rates
'        clsBaseRates.SetFindNewestBaseRateSet
'        TableAccess(clsBaseRates).GetTableData POPULATE_FIRST_BAND
'
'        ' There should only be one with the Base Rate we need
'        TableAccess(clsBaseRates).ApplyFilter "BASERATESET = " & sBaseRateSet
'        sStartDate = clsBaseRates.GetStartDate
'
'        ' Base Rates
'        Set colObjectList = New Collection
'        Set colMatchValues = New Collection
'
'        ' Setup the promotion we are looking for (Base Rate Band with Set number and Start Date)
'        colMatchValues.Add sBaseRateSet
'        colMatchValues.Add sStartDate
'        colObjectList.Add colMatchValues
'
'        Set colUpdateBaseRate = New Collection
'        Set clsBaseRateDetails = New SupervisorUpdateDetails
'
'        colUpdateBaseRate.Add clsBaseRateDetails
'        clsBaseRateDetails.SetKeyMatchValues colMatchValues
'
'        ' Now check the Base Rate itself
'        bDependencyFound = frmManageUpdates.CheckBaseRateSets(clsBaseRateDetails, colUpdateDetails) Or bDependencyFound
'
'        ' Check the dependencies of the Base Rate Band we have, but only if we found a dependency in the first place.
'        bDependencyFound = frmManageUpdates.GetObjectDependencies(BASE_RATES, colUpdateDetails, colObjectList) Or bDependencyFound
'
    End If
'
    CheckMortgageProductCondition = bDependencyFound

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'
'

'DB SYS5457 - Taken from MCAP
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetIntermediaryDependency
' Description   : Given any of the different types of Intermediary passed in, return a generic name
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetIntermediaryDependency(ByRef sObjectName As String)
    On Error GoTo Failed
    
    If sObjectName = INDIVIDUAL_DESCRIPTION Or sObjectName = INTERMEDIARY_COMPANY_DESCRIPTION Or _
       sObjectName = ADMIN_CENTRE_DESCRIPTION Or sObjectName = LEADAGENT_DESCRIPTION Then
    
        sObjectName = INTERMEDIARIES
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'BS BM0498 28/04/03
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CheckTask
' Description   : Check Task for dependencies
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckTask(ByRef clsUpdateDetails As SupervisorUpdateDetails, ByRef colUpdateDetails As Collection) As Boolean
    
    Dim sChasingTask As String
    Dim clsTask As TaskTable
    Dim sFunctionName As String
    Dim colObjectList As Collection
    Dim bDependencyFound As Boolean
    Dim colMatchValues As Collection
    
    On Error GoTo Failed
    
    sFunctionName = "CheckTask"
        
    'Create a Task table.
    Set clsTask = New TaskTable
    
    'Load the correct Task to promote.
    m_clsPromote.GetDatabaseObject clsTask, clsUpdateDetails
    
    'Check for dependant combogroups.
    Set colObjectList = New Collection
    
    'UserRole.
    Set colMatchValues = New Collection
    colMatchValues.Add "UserRole"
    colObjectList.Add colMatchValues
    
    'TaskType.
    Set colMatchValues = New Collection
    colMatchValues.Add "TaskType"
    colObjectList.Add colMatchValues
        
    'TaskContactType.
    Set colMatchValues = New Collection
    colMatchValues.Add "TaskContactType"
    colObjectList.Add colMatchValues
    
    'ComboGroups.
    bDependencyFound = GetObjectDependencies(COMBOBOX_ENTRIES, colUpdateDetails, colObjectList) Or bDependencyFound
    
    'Get the ChasingTask.
    sChasingTask = clsTask.GetChasingTask
    
    'If there is an associated ChasingTask.
    If Len(sChasingTask) > 0 Then
    
        'Reset the collection of keys collections.
        Set colObjectList = New Collection
        Set colMatchValues = New Collection
        
        'Add the ChasingTask.
        colMatchValues.Add sChasingTask
        colObjectList.Add colMatchValues

        'Check for Dependant ChasingTask.
        bDependencyFound = GetObjectDependencies(TASK_MANAGEMENT_TASKS, colUpdateDetails, colObjectList) Or bDependencyFound
    End If
    
    CheckTask = bDependencyFound
    
    Exit Function

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

'BS BM0498 End 28/04/03

'MAR45 End
Private Function ValidatePromotions() As Boolean
    On Error GoTo Failed
    Dim bResult As Boolean
    Dim sMessage As String
    bResult = True
    If cboAvailableUpdates.Text = PRINTING_TEMPLATE Then
        sMessage = "Promotion of Printing Template requires all the associated stages to have been promoted." & vbCrLf _
                                        & "Have you already promoted the stages?"
        If MsgBox(sMessage, vbYesNo, Me.Caption) = vbNo Then
            bResult = False
            sMessage = "Please make sure that stages are promoted before promoting printing templates."
            g_clsErrorHandling.DisplayError sMessage
        End If
    End If
    
    ValidatePromotions = bResult
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Function
'MAR45 End

