VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#2.0#0"; "MSGOCX.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Omiga 4 Supervisor Logon"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   276
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Admin System"
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   3975
   End
   Begin MSGOCX.MSGComboBox cboDatabase 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin MSGOCX.MSGEditBox txtUserID 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   30
   End
   Begin MSGOCX.MSGPasswordEditBox txtPassword 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      Mandatory       =   -1  'True
   End
   Begin MSGOCX.MSGComboBox cboUnit 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.Label lblUnit 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label lblDatabase 
      BackStyle       =   0  'Transparent
      Caption         =   "Database"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label lblUserID 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmLogin
' Description   :   The first form that's displayed to the user when Supervisor
'                   is started. Checks that a database connection can be established,
'                   then invites the user to logon. If successful, loads the main
'                   Supervisor form, frmMain
'
' Prog      Date        Description
' DJP       19/06/01    OK made Default button.
' STB       04/02/02    SYS3327 Ask for UnitID along with username, this is
'                       required for batch processing (as part of the XML
'                       request).
' STB       25/03/02    SYS4312 Added Optimus logon controls.
' STB       05/04/02    SYS4312 ODI logon only enabled if admin system active.
' CL        24/05/02    SYS4736 Incorrect Logon Procedure
' AW        08/07/02    BMIDS00178  Removed ODI admin login
' LD        20/10/05    MAR242 NT Authentication
' RF        28/11/05    MAR353 Windows authentication fails when app server not set
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Underlying table objects.
Private m_clsOmigaUserTable As OmigaUserTable
Private m_clsPasswordTable As PasswordTable

'Return code for the form (Success or Failure).
Private m_ReturnCode As MSGReturnCode


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboDatabase_Click
' Description   : Check that the OmigaUser table exists in the selected
'                 database and hide the login controls if it doesn't.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboDatabase_Click()
    
    Dim sKey As String
    Dim bEnableLogin As Boolean
    
    On Error GoTo Failed
    
    BeginWaitCursor
    
    sKey = cboDatabase.SelText
    
    If Len(sKey) > 0 Then
        'Change the active connection to the selected item.
        g_clsDataAccess.SetActiveConnection sKey
    
        'Does the user table exist?
        bEnableLogin = g_clsDataAccess.DoesTableExist(TableAccess(m_clsOmigaUserTable).GetTable())
        
        'Show/hide the input controls if it does/doesn't.
        ShowSupervisorLogin bEnableLogin
        
        'Refresh the list of Units.
        RefreshListOfUnits
        
    End If
    
    EndWaitCursor
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ShowSupervisorLogin
' Description   : Show or hide the login input controls.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowSupervisorLogin(Optional ByVal bVisible As Boolean = True)
    
    On Error GoTo Failed
    
    txtPassword.Visible = bVisible
    lblPassword.Visible = bVisible
    
    txtUserID.Visible = bVisible
    lblUserID.Visible = bVisible
    
    Exit Sub

Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Initialize
' Description   : Create a User and Password table object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Initialize()
    
    Set m_clsOmigaUserTable = New OmigaUserTable
    Set m_clsPasswordTable = New PasswordTable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Ensure all the valid information has been supplied and
'                 attempt to validate the user's supplied credentials.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    Dim bValid As Boolean
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
        
    'Cast a generic table interface onto the user table.
    Set clsTableAccess = m_clsOmigaUserTable
    
    'Only attempt to login if the table exists.
    If g_clsDataAccess.DoesTableExist(clsTableAccess.GetTable()) Then
        'Ensure all mandatory fields have been supplied.
        bValid = g_clsFormProcessing.DoMandatoryProcessing(Me)
        
        BeginWaitCursor
        
        If bValid Then
            'Does the user have the valid rights?
            m_clsOmigaUserTable.ValidateSuperUser txtUserID.Text
            
            'They have the rights, but is the password correct?
            ValidatePassword
            
            'Store the user's information as global variables.
            StoreUserCredentials
            
            'Set the new database caption on the main form.
            g_clsMainSupport.SetDatabaseCaption
        
            'Inidicate to the caller that the login was sucessful.
            SetReturnCode MSGSuccess
        
            'Hide the form and return control to the caller.
            Hide
        End If
    End If
        
    EndWaitCursor
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidatePassword
' Description   : Unsure that the password supplied matches the UserID.
' History:
' RF        28/11/05    MAR353 Use Windows authentication -
'                       use app server to validate logon.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidatePassword()
    Dim sUserID As String
    Dim sPassword As String
    On Error GoTo Failed

    sPassword = Me.txtPassword.Text
    sUserID = Me.txtUserID.Text

    Dim xmlDoc As New MSXML2.DOMDocument40
    xmlDoc.async = False
    xmlDoc.loadXML ("<REQUEST USERID=""""><OMIGAUSER><USERID></USERID><PASSWORDVALUE></PASSWORDVALUE><UNITID></UNITID><AUDITRECORDTYPE>1</AUDITRECORDTYPE></OMIGAUSER></REQUEST>")
    xmlDoc.selectSingleNode("REQUEST/@USERID").Text = sUserID
    xmlDoc.selectSingleNode("REQUEST/OMIGAUSER/USERID").Text = sUserID
    xmlDoc.selectSingleNode("REQUEST/OMIGAUSER/PASSWORDVALUE").Text = sPassword
    xmlDoc.selectSingleNode("REQUEST/OMIGAUSER/UNITID").Text = cboUnit.GetExtra(cboUnit.ListIndex)

    Dim clsOmiga4 As New Omiga4Support
    Dim sResponse As String
    sResponse = clsOmiga4.RunASP(xmlDoc.xml, "ValidateUserLogon.asp")

    g_clsXMLAssist.CheckXMLResponse sResponse, True

    Exit Sub

Failed:
    txtPassword.SetFocus
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

'    Dim sUserID As String
'    Dim sPassword As String
'    Dim sCurrentPassword As String
'
'    On Error GoTo Failed
'
'    sPassword = Me.txtPassword.Text
'    sUserID = Me.txtUserID.Text
'
'    'Query the table object to get the required password.
'    m_clsPasswordTable.FindPassword sUserID, FindNewest
'
'    'Obtain the current password for the given UserID.
'    sCurrentPassword = m_clsPasswordTable.GetPassword()
'
'    'If the password is incorrect then inform the user.
'    If Encrypt(sPassword) <> sCurrentPassword Then
'        txtPassword.SetFocus
'        g_clsErrorHandling.RaiseError ErrPasswordIncorrect
'    End If
'
'    Exit Sub
'
'Failed:
'
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : The login has been aborted, hide the form and return to the
'                 caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    
    On Error GoTo Failed
    
    'Hide the form and return to the caller.
    Hide
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Load a list of databases from the registry.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed
       
    'The default return code from the form is failure.
    SetReturnCode MSGFailure
    
    BeginWaitCursor
    
    'Populate the database combobox.
    SetDatabaseList
    
    'Set the control visibility state.
    UpdateUnitState
    
    HideSplash
    
    EndWaitCursor
    
    Exit Sub
    
Failed:
    HideSplash
    g_clsErrorHandling.DisplayError
    End
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetDatabaseList
' Description   : Populate and choose a database in the drop-down list.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetDatabaseList()
    On Error GoTo Failed
    Dim sActiveDatabase As String
    
    g_clsFormProcessing.PopulateAvailableTargets Me.cboDatabase
    sActiveDatabase = g_clsDataAccess.GetActiveDatabaseName()
    
    cboDatabase.Text = sActiveDatabase
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtPassword_Validate
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPassword_Validate(Cancel As Boolean)
    
    On Error GoTo Failed
       
    Cancel = Not txtPassword.ValidateData()
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
    Cancel = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtUserID_Validate
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtUserID_Validate(Cancel As Boolean)
    
    On Error GoTo Failed
    
    Cancel = Not txtUserID.ValidateData()
    
    'Refresh the list of Units.
    RefreshListOfUnits
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
    Cancel = True
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetReturnCode
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReturnCode
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : EnableChangeDatabase
' Description   : If the window is called from the connection maintenance
'                 screen, the connection can be fixed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EnableChangeDatabase(Optional ByVal bEnable As Boolean = True)
    
    On Error GoTo Failed
    
    cboDatabase.Enabled = bEnable
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : RefreshListOfUnits
' Description   : Attempt to populate a list of units if the UserID and or
'                 DB alias has altered.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshListOfUnits()
    
    Dim clsUnitTable As UnitTable
    Static sLastUserID As String
    Static sLastDatabase As String
    
    'Only refresh the unit list, if either the UserID or Database has altered.
    If Not ((txtUserID.Text = sLastUserID) And (cboDatabase.SelText = sLastDatabase)) Then
        'Create a unit table to retrieve a list of units.
        Set clsUnitTable = New UnitTable
                
        'Populate the unit table with any Units for which the UserId has a
        'supervisor role with.
        clsUnitTable.GetUnitsForUserID txtUserID.Text
        
        'Populate the combo from the table object.
        PopulateAndSelectUnitCombo clsUnitTable
        
        'Hide the unit combo if there is only one unit.
        UpdateUnitState
    
        'Store the new UserID and Database values.
        sLastUserID = txtUserID.Text
        sLastDatabase = cboDatabase.SelText
    End If
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : UpdateUnitState
' Description   : If the list of units is empty, disable the unit combo.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateUnitState()
    cboUnit.Enabled = (cboUnit.ListCount > 1)
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateAndSelectUnitCombo
' Description   : Populates all the units in the specified table in the combo
'                 list and selects the first one (if there are any).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateAndSelectUnitCombo(ByRef clsUnitTable As UnitTable)

    Dim lIndex As Long
    Dim colUnitIDList As Collection
    Dim colUnitNameList As Collection
    
    'Create the value/text collections.
    Set colUnitIDList = New Collection
    Set colUnitNameList = New Collection
    
    'If we have records then add them into the collections.
    If TableAccess(clsUnitTable).RecordCount > 0 Then
        TableAccess(clsUnitTable).MoveFirst
        
        'Add each unit into both collections.
        For lIndex = 1 To TableAccess(clsUnitTable).RecordCount
            colUnitIDList.Add clsUnitTable.GetUnitID
            colUnitNameList.Add clsUnitTable.GetUnitName
            
            'Move onto the next record.
            TableAccess(clsUnitTable).MoveNext
        Next lIndex
    End If
    
    'Populate the unit combo from these collections.
    cboUnit.SetListTextFromCollection colUnitNameList, colUnitIDList

    'If there is an item, select it.
    If cboUnit.ListCount > 0 Then
        cboUnit.ListIndex = 0
    End If
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : StoreUserCredentials
' Description   : Store the users information as global variables.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub StoreUserCredentials()

    Dim clsUnitTable As UnitTable
    
    g_sSupervisorUser = txtUserID.Text
    g_sUnitID = cboUnit.GetExtra(cboUnit.ListIndex)

    'Create a unit table.
    Set clsUnitTable = New UnitTable
    
    'We must obtain the ChannelID from the Unit.
    g_sChannelID = clsUnitTable.GetChannelIDFromUnitID(g_sUnitID)

End Sub


