VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmDatabaseOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Options"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   Icon            =   "frmDatabaseOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imgTickCross 
      Left            =   360
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabaseOptions.frx":0442
            Key             =   "Cross"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatabaseOptions.frx":059C
            Key             =   "Tick"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   10500
      TabIndex        =   14
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   9180
      TabIndex        =   13
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Frame frameDetails 
      Caption         =   "Database Details"
      Height          =   4995
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   11535
      Begin MSGOCX.MSGPasswordEditBox txtPassword 
         Height          =   315
         Left            =   7380
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
      End
      Begin VB.CommandButton cmdSetActive 
         Caption         =   "&Set Active"
         Height          =   315
         Left            =   10440
         TabIndex        =   12
         Top             =   4260
         Width           =   915
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "Re&place"
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   4320
         Width           =   915
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   315
         Left            =   2220
         TabIndex        =   10
         Top             =   4320
         Width           =   915
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Top             =   4320
         Width           =   915
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "&Test"
         Height          =   315
         Left            =   3240
         TabIndex        =   11
         Top             =   4320
         Width           =   915
      End
      Begin MSGOCX.MSGEditBox txtDetails 
         Height          =   315
         Index           =   0
         Left            =   1860
         TabIndex        =   2
         Top             =   780
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         BackColor       =   16777215
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
      Begin MSGOCX.MSGListView lvConnections 
         Height          =   1995
         Left            =   180
         TabIndex        =   7
         Top             =   2100
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   3519
         Sorted          =   -1  'True
         AllowColumnReorder=   0   'False
      End
      Begin MSGOCX.MSGEditBox txtDetails 
         Height          =   315
         Index           =   4
         Left            =   7380
         TabIndex        =   3
         Top             =   780
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         BackColor       =   16777215
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
      Begin MSGOCX.MSGEditBox txtDetails 
         Height          =   315
         Index           =   3
         Left            =   1860
         TabIndex        =   6
         Top             =   1620
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         BackColor       =   16777215
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
      Begin MSGOCX.MSGEditBox txtDetails 
         Height          =   315
         Index           =   5
         Left            =   7380
         TabIndex        =   1
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         BackColor       =   16777215
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
      Begin MSGOCX.MSGComboBox cboDatabaseTypes 
         Height          =   315
         Left            =   1860
         TabIndex        =   0
         Top             =   360
         Width           =   2955
         _ExtentX        =   5212
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
      Begin MSGOCX.MSGPasswordEditBox txtUserID 
         Height          =   315
         Left            =   1860
         TabIndex        =   4
         Top             =   1200
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
      End
      Begin VB.Label Label5 
         Caption         =   "Database Type"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   420
         Width           =   1395
      End
      Begin VB.Label lblDatabaseServer 
         Caption         =   "Database Server"
         Height          =   255
         Left            =   5880
         TabIndex        =   21
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label lblAppServer 
         Caption         =   "Application Server"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Provider"
         Height          =   195
         Left            =   5880
         TabIndex        =   19
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Password"
         Height          =   195
         Left            =   5880
         TabIndex        =   18
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "User ID"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Database Name"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmDatabaseOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmDatabaseOptions
' Description   : Form that manages all Supervisor database connections. Connections are added,
'                 removed and edited here.
'
' Change history
' Prog      Date        Description
' DJP       04/06/2001  SQL Server port and tidy up.
' STB       21/01/2002  SYS2957 Supervisor Security Enhancement.
' STB       08/07/2002  SYS4529 'ESC' now closes the form and BorderStyle set to 'Fixed Dialog'.
' CL        23/05/2002  SYS4705 Add asterix's to the password displayed on the Database Options screen
' JR        30/05/2002  SYS4816 Enhancement to password security
' SA        17/06/2002  SYS4905 Use normal password not Encrypted password when building connection string.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change History:
' Prog      Date        Description
' DB        14/11/2002  BMIDS00851 Add asterix's to the user id displayed
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS Change History
' RF        29/11/05    MAR734 Supervisor should cope with integrated security
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom Change History
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Constants

' Textbox keys
Private Const SERVICE_NAME As Integer = 0
Private Const USER_ID = 1
Private Const APP_SERVER = 3
Private Const PROVIDER = 4
Private Const DATABASE_SERVER = 5

' General constants
Private Const nImageTick As Integer = 2
Private Const nImageCross As Integer = 1

' Private variables
Private m_ReturnCode As MSGReturnCode
Private m_nActiveConnection As Integer
Private m_sOriginalConnection As String
Private m_clsDatabaseOptions As DatabaseOptions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cboDatabaseTypes_Click
' Description   : Called when the database type is changed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboDatabaseTypes_Click()
    On Error GoTo Failed
    
    DatabaseChange

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetEditDatabaseType
' Description   : Returns the type of the database as selected by the Database combo. This could
'                 be different to the currently selected database from the listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetEditDatabaseType() As enumDatabaseTypes
    On Error GoTo Failed
    Dim nType As enumDatabaseTypes
    
    nType = cboDatabaseTypes.ListIndex
    GetEditDatabaseType = nType
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DatabaseChange
' Description   : Called when the database type is changed.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DatabaseChange()
    On Error GoTo Failed
    Dim nType As enumDatabaseTypes
    Dim bEnableServer As Boolean
    nType = GetEditDatabaseType()
    
    Select Case nType
        Case INDEX_ORACLE
            bEnableServer = False
        
        Case INDEX_SQL_SERVER
            bEnableServer = True
        Case -1
            ' ok - they just haven't selected anything
        Case Else
            g_clsErrorHandling.RaiseError errDatabaseNotSupported, "Database Options"
    End Select
    
    txtDetails(DATABASE_SERVER).Text = ""
    txtDetails(DATABASE_SERVER).Enabled = bEnableServer
    lblDatabaseServer.Enabled = bEnableServer
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Called when form is loaded. Performs all connection initialisation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    On Error GoTo Failed
    BeginWaitCursor
    Set m_clsDatabaseOptions = New DatabaseOptions
    
    ' Default return code
    SetReturnCode MSGFailure
    
    m_nActiveConnection = -1
    Set lvConnections.SmallIcons = imgTickCross
        
    SetupCombos
    InitialiseFields
    SetupListViewHeaders
    PopulateListView
    
    EndWaitCursor
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub SetupCombos()
    On Error GoTo Failed
    
    cboDatabaseTypes.AddItem DATABASE_ORACLE, INDEX_ORACLE
    cboDatabaseTypes.AddItem DATABASE_SQL_SERVER, INDEX_SQL_SERVER

    ' Default Oracle for now
    cboDatabaseTypes.ListIndex = -1
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetupListViewHeaders
' Description   : Called during initialisation to setup the headings the user will see in the
'                 database connection listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupListViewHeaders()
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    
    lvHeaders.nWidth = 12
    lvHeaders.sName = "Database Type"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 12
    lvHeaders.sName = "Database Name"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 12
    lvHeaders.sName = "User ID"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 0
    lvHeaders.sName = "Password"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Application Server"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 12
    lvHeaders.sName = "Provider"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Database Server"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 8
    lvHeaders.sName = "Active"
    headers.Add lvHeaders
           
    lvConnections.AddHeadings headers
    
    ' Load any field resizing the user may have done
    lvConnections.LoadColumnDetails Me.Name
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitialiseFields
' Description   : Sets the default state of all screen fields
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitialiseFields(Optional bEnable As Boolean = False)
    On Error GoTo Failed
    
    cmdSetActive.Enabled = bEnable
    cmdRemove.Enabled = bEnable
    cmdTest.Enabled = bEnable
    cmdReplace.Enabled = bEnable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : InitialiseFields
' Description   : Sets the default state of all screen fields
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateListView()
    Dim bActive As Boolean
    Dim nThisEntry As Integer
    Dim nActiveConnection As Integer
    Dim sGetPassword As String
    Dim Counter As Integer
    Dim sProvider As String
    Dim sPassword As String
    Dim sAppServer As String
    Dim sTargetUserID As String
    Dim sDatabaseType As String
    Dim sTargetService As String
    Dim sDatabaseServer As String
    Dim colRow As Collection
    Dim colConnections As Collection
    Dim clsSupervisorConnection As SupervisorConnection
    'JR SYS4816
    Dim clsPasswordData As PasswordData
    Dim sEncryptPassword As String
    Dim iCounter As Integer
    'End
    
    'DB BMIDS00851
    Dim sEncryptUserID As String
    
    ' Get list of current connections
    Set colConnections = g_clsDataAccess.GetConnectionList()
    
    ' Loop through all connections returned adding them to the listview
    If Not colConnections Is Nothing Then
        nThisEntry = 1
        If colConnections.Count > 0 Then
            For Each clsSupervisorConnection In colConnections
                Set colRow = New Collection
                
                bActive = clsSupervisorConnection.GetIsActive()
                sProvider = clsSupervisorConnection.GetProvider()
                sPassword = clsSupervisorConnection.GetPassword()
                sTargetUserID = clsSupervisorConnection.GetUserID()
                sTargetService = clsSupervisorConnection.GetDatabaseName()
                sAppServer = clsSupervisorConnection.GetAppServer()
                sDatabaseServer = clsSupervisorConnection.GetDatabaseServer()
                sDatabaseType = clsSupervisorConnection.GetDatabaseType()
                
                'JR SYS4816
                Set clsPasswordData = New PasswordData
                clsPasswordData.SetPassword sPassword
                clsPasswordData.SetUserID sTargetUserID ' DB BMIDS00851
                
                For iCounter = 1 To Len(sPassword)
                    sEncryptPassword = sEncryptPassword & "*"
                Next
                'JR End
                
                'DB BMIDS00851
                For iCounter = 1 To Len(sTargetUserID)
                    sEncryptUserID = sEncryptUserID & "*"
                Next
                'DB END
                
                colRow.Add sDatabaseType
                colRow.Add sTargetService
                'colRow.Add sTargetUserID
                colRow.Add sEncryptUserID 'DB BMIDS00851
                'colRow.Add sPassword
                colRow.Add sEncryptPassword ' JR SYS4816
                colRow.Add sAppServer
                colRow.Add sProvider
                colRow.Add sDatabaseServer
                colRow.Add ""
                                               
                If bActive Then
                    nActiveConnection = nThisEntry
                End If
                
                lvConnections.AddLine colRow, clsPasswordData
                
                sEncryptPassword = "" 'JR SYS4816
                sEncryptUserID = "" 'DB BMIDS00851
                nThisEntry = nThisEntry + 1
            Next
        
            If nActiveConnection > 0 Then
                SetActive nActiveConnection
                m_sOriginalConnection = g_clsDataAccess.GetConnectionKey()
            End If
        
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ReplaceItem
' Description   : Replaces the currently selected row with the one passed in, or a blank line
'                 if nothing is passed in
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ReplaceItem(Optional colLine As Collection = Nothing)
    On Error GoTo Failed
    
    ' Remove the current line
    RemoveLine

    If Not colLine Is Nothing Then
        AddRow colLine
    Else
        AddRow
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : RemoveLine
' Description   : Removes the currently selected row in the listview.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveLine()
    On Error GoTo Failed
    
    Dim lstItem As ListItem
    Dim bEnable As Boolean
    Set lstItem = lvConnections.SelectedItem
    
    ' Remove a row from the listview
    If Not lstItem Is Nothing Then
        lvConnections.RemoveLine lstItem.Index
    End If
    
    ' Current selection has now gone, so the fields state.
    InitialiseFields
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetActive
' Description   : Called when the SetActive button is pressed and changes the active connection.
'                 If bLogin is true, the user will be forced to login before they can change the
'                 connection.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetActive(Optional nActiveIndex As Integer = 0, Optional bLogin As Boolean = False)
    On Error GoTo Failed
    Dim bLoginSuccessful As Boolean
    Dim nIndex As Integer
    Dim sDatabase As String
    Dim lstItem As ListItem
    
    BeginWaitCursor
    
    If nActiveIndex = 0 Then
        Set lstItem = lvConnections.SelectedItem
        nIndex = lstItem.Index
    Else
        Set lstItem = lvConnections.ListItems.Item(nActiveIndex)
        nIndex = nActiveIndex
    End If
    
    If Not lstItem Is Nothing Then
        bLoginSuccessful = False
        
        If bLogin Then
            sDatabase = GetActiveKey()
            
            ' Make sure the connection is there
            AddConnection nIndex
            
            If (g_clsDataAccess.SetActiveConnection(sDatabase)) Then
                bLoginSuccessful = g_clsMainSupport.UserLogin(Me, False)
            End If
        Else
            bLogin = False
        End If
        
        If (bLoginSuccessful And bLogin) Or (Not bLogin) Then
            ' We've successfully changed the active connection, so remove the current active connection
            ' and create a new one
            SetAllInactive
    
            Set lstItem = lvConnections.ListItems.Item(nIndex)
            lstItem.ListSubItems(COL_ACTIVE_CONNECTION - 1).ReportIcon = nImageTick
            m_nActiveConnection = nIndex
        End If
    End If
    
    EndWaitCursor
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAllInactive
' Description   : Sets the state of each row in the listview to inactive. This just sets the
'                 active icon to a red cross instead of a tick
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAllInactive()
    On Error GoTo Failed
    Dim lstItem As ListItem
    Dim lstSubItem As ListSubItem
    
    For Each lstItem In Me.lvConnections.ListItems
        For Each lstSubItem In lstItem.ListSubItems
            If lstSubItem.Index = COL_ACTIVE_CONNECTION - 1 Then
                lstSubItem.ReportIcon = nImageCross
            End If
        Next
    Next

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : AddConnection
' Description   : Adds a connection to the DataAccess class for use by Supervisor. Current connection
'                 details are pulled from the listview, added to a SupervisorConnection class, then
'                 appended to a list of connections DataAccess stores.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddConnection(nListIndex As Integer)
    On Error GoTo Failed
    Dim sUserID As String
    Dim sService As String
    Dim sProvider As String
    Dim sPassword As String
    Dim sAppServer As String
    Dim sConnection As String
    Dim sConnectionKey As String
    Dim sDatabaseType As String
    Dim sDatabaseServer As String
    Dim colLine As Collection
    Dim connDetails As ConnectionDetails
    Dim clsConnection As SupervisorConnection
    'JR SYS4816
    Dim clsPasswordData As PasswordData
    Dim NewColLine As Collection
    'End
        
    ' Get the requested connection
    'JR SYS4816
    Set clsPasswordData = New PasswordData
    Set colLine = lvConnections.GetLine(nListIndex, clsPasswordData)
    sPassword = clsPasswordData.GetPassword
    sUserID = clsPasswordData.GetUserID 'DB BMIDS00851
        
    Set NewColLine = clsPasswordData.ReplaceCollectionForPassword(colLine)
    'Set NewColLine = clsPasswordData.ReplaceCollectionForUserID(colLine) 'DB BMIDS00851
    'End
    
    ' Get the details of the connection
    'm_clsDatabaseOptions.GetConnectionDetails colLine, connDetails
    m_clsDatabaseOptions.GetConnectionDetails NewColLine, connDetails
    sConnection = m_clsDatabaseOptions.BuildConnectionString(connDetails)
    
    ' Read details of the connection
    sService = NewColLine(COL_SERVICE_NAME)
    sUserID = NewColLine(COL_USER_ID)
    'sPassword = colLine(COL_PASSWORD)
    sPassword = NewColLine(COL_PASSWORD)
    
    sProvider = NewColLine(COL_PROVIDER)
    sAppServer = NewColLine(COL_APP_SERVER)
    sDatabaseType = NewColLine(COL_DATABASE_TYPE)
    sDatabaseServer = NewColLine(COL_DATABASE_SERVER)
     
    ' Create a new connection, add the connection details, and add to the DataAccess
    Set clsConnection = New SupervisorConnection
    
    clsConnection.SetConnectionString sConnection
    clsConnection.SetUserID sUserID
    clsConnection.SetDatabaseName sService
    clsConnection.SetProvider sProvider
    clsConnection.SetPassword sPassword
    clsConnection.SetAppServer sAppServer
    clsConnection.SetDatabaseServer sDatabaseServer
    clsConnection.SetDatabaseType sDatabaseType
    
    sConnectionKey = GetConnectionKey(sService, sUserID)
    
    ' Add it to the DataAccess class
    g_clsDataAccess.AddConnection clsConnection, sConnectionKey
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetDatabaseTypeString
' Description   : Gets the name of the database based on the type passed in
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetDatabaseTypeString(nType As enumDatabaseTypes) As String
    On Error GoTo Failed
    GetDatabaseTypeString = cboDatabaseTypes.ListText
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetActiveKey
' Description   : Gets the active key for the current active connection.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetActiveKey()
    On Error GoTo Failed
    Dim sUserID As String
    Dim sService As String
    Dim sFunctionName As String
    Dim sConnectionKey As String
    Dim colLine As Collection
    'DB BMIDS00851
    Dim clsPasswordData As PasswordData
    Dim colNewLine As Collection
    
    ' Get the active row
    Set colLine = lvConnections.GetLine(, clsPasswordData)
    Set colNewLine = clsPasswordData.ReplaceCollectionForPassword(colLine)
    
    sService = colLine(COL_SERVICE_NAME)
    sUserID = colNewLine(COL_USER_ID)
    
    'DB BMIDS00851
    Set clsPasswordData = New PasswordData
    clsPasswordData.SetUserID sUserID
    
    sConnectionKey = GetConnectionKey(sService, sUserID)
    
    If Len(sConnectionKey) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Connection Key empty"
    End If

    GetActiveKey = sConnectionKey
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetReturnCode
' Description   : Sets the return code of this form to be returned to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReturnCode
' Description   : Allows the caller to read the return code of this form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateScreenData
' Description   : Called to ensure all data entered is valid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateScreenData()
    On Error GoTo Failed
    
    ' Is there an active connection?
    If m_nActiveConnection = -1 Then
        g_clsErrorHandling.RaiseError errGeneralError, "At least one active connection must be created"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   : Called to save all data entered onto the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    On Error GoTo Failed
    BeginWaitCursor
    ' Need to save the values from the listbox into the registry
    SaveItems
    
    'Force the database connection to open otherwise the Security component will fail.
    g_clsDataAccess.OpenDatabase
    
    'Create a security server to check for user access.
    frmMain.InitialiseSecurity

    ' Only build the parts of the main treeview that are needed (the database we have moved to may not support it)
    g_clsMainSupport.ShowSupervisorObjects
    
    ' Save any field resizing the user may have done
    lvConnections.SaveColumnDetails
    
    EndWaitCursor
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveItems
' Description   : Saves all listview entries to the registry.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveItems()
    On Error GoTo Failed
    Dim bActive As Boolean
    Dim nCount As Long
    Dim nThisLine As Integer
    Dim colLine As Collection
    'JR SYS4816
    Dim clsPasswordData As PasswordData
    Set clsPasswordData = New PasswordData
    Dim NewColLine As Collection
    'End
    
    ' Delete what's there, then add the new entries
    DeleteKey HKEY_LOCAL_MACHINE, REG_CONNECTIONS
    
    ' Clear all connections from DataAccess
    g_clsDataAccess.ClearConnections
    
    nCount = lvConnections.ListItems.Count
    
    For nThisLine = 1 To nCount
        If m_nActiveConnection = nThisLine Then
            bActive = True
        Else
            bActive = False
        End If
        ' Save this row
        Set colLine = lvConnections.GetLine(nThisLine, clsPasswordData)
        Set NewColLine = clsPasswordData.ReplaceCollectionForPassword(colLine)
        'Set NewColLine = clsPasswordData.ReplaceCollectionForUserID(colLine) 'DB BMIDS00851
        m_clsDatabaseOptions.SaveConnection NewColLine, m_nActiveConnection, nThisLine, bActive
    Next
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ClearTextBoxes
' Description   : Clears the text of all text boxes on the screen.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ClearTextBoxes()
    On Error GoTo Failed
    txtDetails(SERVICE_NAME).Text = ""
    'DB BMIDS00851
    txtUserID.Text = ""
    'txtDetails(USER_ID).Text = ""
    'txtDetails(PASSWORD).Text = ""
    txtDetails(PROVIDER).Text = ""

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : AddRow
' Description   : Adds a new row into the listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddRow(Optional colLine As Collection = Nothing)
    On Error GoTo Failed
    Dim colRow As Collection
    Dim iCounter As Integer
    Dim sPassword As String
    Dim clsPasswordData As PasswordData 'JR SYS4816
        
    'DB BMIDS00851
    Dim sUserID As String
        
    If Not colLine Is Nothing Then
        Set colRow = colLine
    Else
        Set colRow = New Collection
        colRow.Add cboDatabaseTypes.Text
        colRow.Add txtDetails(SERVICE_NAME).Text
        'colRow.Add txtDetails(USER_ID).Text
        'sPassword = txtPassword.Text
        
        'JR SYS4816
        Set clsPasswordData = New PasswordData
        clsPasswordData.SetPassword txtPassword.Text
        For iCounter = 1 To Len(txtPassword.Text)
            sPassword = sPassword & "*"
        Next
        'JR End
            
        'DB BMIDS00851
        'Set clsPasswordData = New PasswordData
        clsPasswordData.SetUserID txtUserID.Text
        For iCounter = 1 To Len(txtUserID.Text)
            sUserID = sUserID & "*"
        Next
        'DB END
        
        colRow.Add sUserID 'DB BMIDS00851
        colRow.Add sPassword
        colRow.Add txtDetails(APP_SERVER).Text
        colRow.Add txtDetails(PROVIDER).Text
        colRow.Add txtDetails(DATABASE_SERVER).Text
        colRow.Add ""
    End If
    
    'JR SYS4816 added clsPasswordData
    lvConnections.AddLine colRow, clsPasswordData
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : IsAddDataValid
' Description   : Verifies that all data entered into the textboxes is valid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsAddDataValid()
    On Error GoTo Failed
    Dim bValid As Boolean
    
    bValid = False
    
    ' RF 29/11/05 MAR734
    'If (Len(txtDetails(SERVICE_NAME).Text) > 0) And _
    '   (Len(txtUserID.Text) > 0) And _
    '   (Len(txtPassword.Text) > 0 And _
    '   Len(txtDetails(PROVIDER).Text) > 0) Then
    If (Len(txtDetails(SERVICE_NAME).Text) > 0) And _
       (Len(txtDetails(PROVIDER).Text) > 0) Then
       bValid = True
    End If
    
    IsAddDataValid = bValid
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : TestConnection
' Description   : Opens an ADO connection to make sure the connection details the user enteres
'                 are valid.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TestConnection()
    On Error GoTo Failed
    
    Dim sConnection As String
    Dim connDetails As ConnectionDetails
    Dim adoConn As Connection
    Dim colLine As Collection
    'JR SYS4816
    Dim clsPasswordData As PasswordData
    Dim sPassword As String
    Dim NewColLine As Collection
    'End
    
    Dim sUserID As String 'DB BMIDS00851
    
    BeginWaitCursor
    
    'JR SYS4816
    Set clsPasswordData = New PasswordData
    Set colLine = lvConnections.GetLine(, clsPasswordData)
    sPassword = clsPasswordData.GetPassword
    sUserID = clsPasswordData.GetUserID 'DB BMIDS00851
        
    Set NewColLine = clsPasswordData.ReplaceCollectionForPassword(colLine)
    'Set NewColLine = clsPasswordData.ReplaceCollectionForUserID(colLine) 'DB BMIDS00851
    'End
    
'    m_clsDatabaseOptions.GetConnectionDetails colLine, connDetails
    m_clsDatabaseOptions.GetConnectionDetails NewColLine, connDetails
    sConnection = m_clsDatabaseOptions.BuildConnectionString(connDetails)
    
    Set adoConn = New Connection
    adoConn.Open sConnection
                
    MsgBox "Connection Successful", vbInformation
    
    adoConn.Close
    Set adoConn = Nothing
    
    EndWaitCursor

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Event Handlers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdReplace_Click
' Description   : Called when the user clicks Replace
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdReplace_Click()
    On Error GoTo Failed
    
    ReplaceItem
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdAdd_Click
' Description   : Called when the user clicks Set Active.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSetActive_Click()
    On Error GoTo Failed
    
    SetActive , True
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdAdd_Click
' Description   : Called when the user clicks Add. Need to make sure all data has been entered
'                 correctly, then add a row.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAdd_Click()
    On Error GoTo Failed
    
    If g_clsFormProcessing.DoMandatoryProcessing(Me) Then
        AddRow
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Called when the user clicks Cancel. If there was an active connection when we
'                 came into this screen, restore it as the active connection overriding any active
'                 connections we may have setup here.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    On Error GoTo Failed
    
    If Len(m_sOriginalConnection) > 0 Then
        g_clsDataAccess.SetActiveConnection m_sOriginalConnection
        g_clsMainSupport.SetDatabaseCaption
    End If
    
    Hide
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Called when the user clicks OK. Need to validate the data the user has entered and
'                 if it's ok, save everything.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    
    ValidateScreenData
    SaveScreenData
    SetReturnCode
    
    Hide

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdRemove_Click
' Description   : Called when the user clicks Remove. Removes the selected row from the listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemove_Click()
    On Error GoTo Failed
    
    RemoveLine
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdTest_Click
' Description   : Called when the user clicks Test.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdTest_Click()
    On Error GoTo Failed
    
    TestConnection
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvConnections_DblClick
' Description   : Called when the user double clicks a row on the listview. Just need to make
'                 that row active.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvConnections_DblClick()
    On Error GoTo Failed
    
    SetActive , True
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : lvConnections_ItemClick
' Description   : Called when the user clicks a row on the listview. Need to populate the textboxes
'                 with the entries from the listview row selected.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvConnections_ItemClick(ByVal Item As MSComctlLib.IListItem)
    On Error GoTo Failed
    Dim connDetails As ConnectionDetails
    Dim colLine As Collection
    
    'JR SYS4816
    Dim clsPasswordData As PasswordData
    Dim colNewLine As Collection 'SA SYS4905
    'End
    
    InitialiseFields True
    SetScreenFields
    
    Set colLine = lvConnections.GetLine(, clsPasswordData)
    Set colNewLine = clsPasswordData.ReplaceCollectionForPassword(colLine)   'SA SYS4905
    
    'Set colLine = lvConnections.GetLine(, clsPasswordData) 'DB BMIDS00851
    'Set colNewLine = clsPasswordData.ReplaceCollectionForUserID(colLine) 'DB BMIDS00851
    
    'SYS4905 Use new collection to get connection details
    'm_clsDatabaseOptions.GetConnectionDetails colLine, connDetails
    m_clsDatabaseOptions.GetConnectionDetails colNewLine, connDetails
    
    ' Populate the text boxes with the values
    txtDetails(SERVICE_NAME).Text = connDetails.sService
    txtUserID.Text = connDetails.sUserID
    txtPassword.Text = connDetails.sPassword
    txtDetails(PROVIDER).Text = connDetails.sProvider
    txtDetails(APP_SERVER).Text = connDetails.sAppServer
    txtDetails(DATABASE_SERVER).Text = connDetails.sDatabaseServer
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub SetScreenFields()
    On Error GoTo Failed
    
    Dim connDetails As ConnectionDetails
    Dim colLine As Collection
    
    Set colLine = lvConnections.GetLine()
    
    m_clsDatabaseOptions.GetConnectionDetails colLine, connDetails
    
    ' Populate the text boxes with the values
    txtDetails(SERVICE_NAME).Text = connDetails.sService
    txtUserID.Text = connDetails.sUserID
    txtPassword.Text = connDetails.sPassword
    txtDetails(PROVIDER).Text = connDetails.sProvider
    txtDetails(APP_SERVER).Text = connDetails.sAppServer
    txtDetails(DATABASE_SERVER).Text = connDetails.sDatabaseServer
    cboDatabaseTypes.Text = connDetails.sDatabaseType
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : txtDetails_Validate
' Description   : The Validate event is fired when the user clicks away from one of the textboxes,
'                 or tabs away. We can do per control validation here.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtDetails_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not txtDetails(Index).ValidateData()
End Sub
Private Function GetSelectedItem() As Integer
    On Error GoTo Failed
    Dim nItemIndex As Integer
    Dim colLine As Collection
    Dim lstItem As ListItem
    
    Set lstItem = lvConnections.SelectedItem

    If Not lstItem Is Nothing Then
        nItemIndex = lstItem.Index
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "GetSelectedItem: No selected items"
    End If
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function
Private Sub txtPassword_Validate(Cancel As Boolean)
    
    Cancel = Not txtPassword.ValidateData()

End Sub
