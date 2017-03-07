VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormProcessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class         :   FormProcessing
' Description   :   Contains generic routines to do with forms, such as mandatory
'                   processing
' Change history
' Prog      Date        Description
' DJP       29/11/00    In PopulateFromRecordset ensure we move to the first record.
' CL        20/04/01    Adjusted ClearScreenFields to clear Checkboxes
' DJP       05/06/01    Added EnableFrameFields
' DJP       26/06/01    SQL Server port
' DJP       28/08/01    SYS2564 CheckForDuplicates, only move once the recordset has been set, otherwise we
'                       are moving the wrong one.
' DJP       28/08/01    DJP SYS2564 MCAP Core functionality
' DJP       14/11/01    DJP SYS2996 Make <None> in combos to <Select>, but move to constants.bas
' DJP       20/11/01    SYS2831/SYS2912 Support client variants & SQL Server locking problem. Removed
'                       NextNumber
' DJP       03/12/01    SYS2912 SQL Server locking problem - PopulateFromRecordset shouldn't raise
'                       an error if no records passed in, just do nothing.
' STB       05/12/01    SYS2912 - PopulateComboByValidation() added.
' DJP       02/02/01    SYS2831 Client variants.
' STB       08/03/02    SYS4251 - Added HandleDateEx can now handle null dates (empty masks).
' PB        10/11/06    EP2_13 Changes for new introducer structure
' TW        09/04/2007  EP2_2332 - Error when trying to edit a user's "User Profile"
' AW        24/04/07    EP2_2562 - Remove duplicate calls to CheckApostrophes()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_frmCurrent As Form

' Constants
Private Const cstrWILDCARD As String = "*"
Private Const cstrWILDCARDSQL = "%"
Private Const BLANK_DATE_MASK As String = "__/__/____"

'BMIDS Change history
'Prog      Date         Description
'GD        16/05/2002   BMIDS00009 : Allow data combos to have concatenated contents.

Friend Function ClearScreenFields(frmToClear As Form) As Boolean
    Dim bRet As Boolean
    Dim ctrlCurrent As Control

    bRet = True
    
    On Error Resume Next
    
    For Each ctrlCurrent In frmToClear.Controls
        
        If TypeOf ctrlCurrent Is MSGDataCombo Or TypeOf ctrlCurrent Is MSGComboBox Then
            ctrlCurrent.Text = ""
            ctrlCurrent.ListIndex = -1
        ElseIf TypeOf ctrlCurrent Is MSGEditBox Or TypeOf ctrlCurrent Is MSGTextMulti Then
            ctrlCurrent.Text = ""
        ElseIf TypeOf ctrlCurrent Is CheckBox Then
               ctrlCurrent.Value = False
        End If
    Next
    
    ClearScreenFields = bRet
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
End Function

Friend Function DoMandatoryProcessing(ByRef frmToCheck As Form, Optional bShowError As Boolean = True) As Boolean
    Dim bRet As Boolean
    Dim ctrlCurrent As Control
    Dim ctrlFirst As Control
    Dim nIndex As Integer
    Dim bMand As Boolean
    Dim bFirst As Boolean
    Dim bFailed As Boolean
    bRet = False
    bFirst = True
    
    On Error Resume Next
    bFailed = False
    
    For Each ctrlCurrent In frmToCheck.Controls
        
        If TypeOf ctrlCurrent Is MSGDataCombo Then
        End If
        
        If TypeOf ctrlCurrent Is MSGEditBox Or TypeOf ctrlCurrent Is MSGComboBox Then
            bMand = ctrlCurrent.Mandatory
                    
            If bMand = True Then
                If TypeOf ctrlCurrent Is MSGEditBox Then
    
                End If
                Err.Clear
                bRet = ctrlCurrent.CheckMandatory(bShowError)
                
                If bRet = False And Err.Number = 0 Then
                    bRet = (IsControlVisible(ctrlCurrent) And ctrlCurrent.Visible And ctrlCurrent.Enabled)
                    If bRet = True Then
                        bFailed = True
                        If bFirst = True Then
                            Set ctrlFirst = ctrlCurrent
                            bFirst = False
                        End If
                    End If
                End If
            End If
        
        End If
    Next
    
    If bFailed = True Then
        On Error GoTo Failed
        SetTabFromControl ctrlFirst
        
        If bShowError = True Then
            If Not ctrlFirst Is Nothing Then
                If ctrlFirst.Enabled = True Then
                    ctrlFirst.SetFocus
                    g_clsErrorHandling.RaiseError errMandatoryFieldsRequired
                Else
                    bFailed = False
                End If
            End If
        End If
    
    End If
    
    DoMandatoryProcessing = Not bFailed
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub SetControlFocus(ctrl As Object)
    On Error GoTo Failed
    SetTabFromControl ctrl
    ctrl.SetFocus
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub SetTabFromControl(ctrl As Control)
    On Error GoTo Failed
    Dim bDone As Boolean
    Dim sTag As String
    Dim objParent As Object
    Dim sTab As String
    Dim ssTabControl As SSTab
    bDone = False
    
    Set objParent = ctrl.Container
    
    While bDone = False
        If TypeOf objParent Is Frame Then
            sTag = objParent.Tag

            If sTag Like "Tab#*" Then
                sTab = Right(sTag, Len(sTag) - 3)
                ' So, we have the tab, but we now need the SSTab, should be the container
                
                If TypeOf objParent.Container Is SSTab Then
                    Set ssTabControl = objParent.Container
                                            
                     ssTabControl.Tab = CInt(sTab) - 1
                    bDone = True
                End If
            Else
                Set objParent = objParent.Container
            End If
        ElseIf TypeOf objParent Is Form Then
            bDone = True
        Else
            Set objParent = objParent.Container
        End If
    Wend
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function IsControlVisible(ctrlToCheck As Control) As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bDone As Boolean
    Dim sTag As String
    Dim objParent As Object
    Dim nPos As Variant
    Dim sTab As String
    Dim ssTabControl As SSTab
    
    bDone = False
    bRet = True
    
    Set objParent = ctrlToCheck.Container
    
    While bDone = False
        If TypeOf objParent Is Frame Then
            sTag = objParent.Tag

            If sTag Like "Tab#*" Then
                nPos = InStr(sTag, "Tab")
                                
                sTab = Right(sTag, Len(sTag) - 3)
                ' So, we have the tab, but we now need the SSTab, should be the container
                
                If TypeOf objParent.Container Is SSTab Then
                    Set ssTabControl = objParent.Container
                                            
                    If ssTabControl.TabVisible(CInt(sTab) - 1) = False Then
                        bRet = False
                    End If
                    bDone = True
                End If
            Else
                Set objParent = objParent.Container
            End If
        ElseIf TypeOf objParent Is Form Then
            bDone = True
            bRet = True
        Else
            Set objParent = objParent.Container
        End If
    Wend
    IsControlVisible = bRet
    Exit Function
Failed:
    g_clsErrorHandling.DisplayError
    IsControlVisible = False
End Function

Friend Sub PopulateDBScreen(colTableAccess As Collection, Optional colDataControl As Collection = Nothing, Optional enumPopulateType As PopulateType = POPULATE_KEYS)
    On Error GoTo Failed
    Dim bDone As Boolean
    Dim nTableCount As Integer
    Dim nThisTable As Integer
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    
    If Not colTableAccess Is Nothing Then
        nTableCount = colTableAccess.Count
    
        nThisTable = 1
        bDone = False
        While bDone = False And nThisTable <= nTableCount
            Set clsTableAccess = colTableAccess(nThisTable)
            
            Set rs = clsTableAccess.GetTableData(enumPopulateType)
            
'            If rs.EOF = True Or rs.BOF = True Then
'                g_clsErrorHandling.RaiseError errRecordNotFound
'            End If
            clsTableAccess.ValidateData
            
            If Not colDataControl Is Nothing Then
                If colDataControl.Count >= nThisTable Then
                    Set colDataControl(nThisTable).DataSource = rs
                End If
            End If
            
            If rs.RecordCount > 0 Then
                clsTableAccess.MoveFirst
            End If
            nThisTable = nThisTable + 1
        Wend
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "PopulateDBScreen: Table Class empty"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Sub CopyRecordset(clsTableSrc As TableAccess, clsTableDest As TableAccess)
    Dim nFields As Integer
    Dim nRecords As Integer
    Dim nThisRecord As Long
    Dim nThisField As Long
    Dim rsCopy As ADODB.Recordset
    Dim rsSource As ADODB.Recordset
    Dim vValue As Variant
    Dim vTmp As Variant
    On Error GoTo Failed
    
    Set rsSource = clsTableSrc.GetRecordSet()
    ' Just open it, but don't populate it
    Set rsCopy = clsTableDest.GetTableData(POPULATE_TABLE)
    
    If Not rsCopy Is Nothing And Not rsSource Is Nothing Then
        If rsSource.RecordCount > 0 Then
            rsSource.MoveFirst
            vTmp = rsSource.GetRows()
            ' Arrays start at 0, so it's nFields + 1 and nRecords + 1
            nFields = UBound(vTmp, 1)
            nRecords = UBound(vTmp, 2)
              
                        
            For nThisRecord = 0 To nRecords
                ' Create a new row in the copy
                rsCopy.AddNew
                
                For nThisField = 0 To nFields
                    ' Copy all fields
                    vValue = vTmp(nThisField, nThisRecord)
                    rsCopy(nThisField).Value = vValue
                Next
            Next
            rsCopy.MoveFirst
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "CopyRecordSet: No records to copy"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "CopyRecordSet: Recordset is empty"
    End If
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Friend Sub CreateNewRecord(clsTable As TableAccess)
    On Error GoTo Failed
    
    If g_clsDataAccess.DoesTableExist(clsTable.GetTable()) Then
        clsTable.AddRow
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub Class_Initialize()
    'm_bAdded = False
End Sub

Friend Sub HandleRadioButtons(optYes As OptionButton, optNo As OptionButton, vRadio As Variant, enumOperation As ControlOperation)
    If enumOperation = SET_CONTROL_VALUE Then
        If (Not IsNull(vRadio) And Len(vRadio) > 0) Then
            If (CBool(vRadio) = True) Then
                optYes.Value = True
                optNo.Value = False
            Else
                optYes.Value = False
                optNo.Value = True
            End If
        Else
            optYes.Value = False
            optNo.Value = True
        End If
    Else
        vRadio = CInt(optYes.Value)
    
        If CBool(vRadio) = True Then
            vRadio = 1
        End If
    End If

End Sub

Friend Sub HandleCheckBox(chkBox As CheckBox, vVal As Variant, enumOperation As ControlOperation)
    If enumOperation = SET_CONTROL_VALUE Then
        If (Not IsNull(vVal) And Len(vVal) > 0) Then
            chkBox.Value = CInt(vVal)
        Else
            chkBox.Value = 0
        End If
    Else
        vVal = chkBox.Value
        
        If CBool(vVal) = True Then
            vVal = CInt(1)
        End If
    
    End If

End Sub

Public Sub HandleDate(txtDate As MSGEditBox, vDate As Variant, enumOperation As ControlOperation)
        
    On Error GoTo Error
    
    If (enumOperation = SET_CONTROL_VALUE) Then
        If (Not IsNull(vDate) And Len(vDate) > 0) Then
            'vDate = g_clsValidation.GetFullYear(CStr(vDate))
            vDate = Format(vDate, "Short Date")
            txtDate.Text = vDate
        End If
    Else
        Dim bRet As Boolean
        Dim bNull As Boolean
    
        g_clsValidation.SetReportError False
        bRet = g_clsValidation.ValidateDate(txtDate, bNull, False)
        g_clsValidation.SetReportError True
        
        If (bRet = True) Then
            vDate = txtDate.Text
        Else
            vDate = Null
        End If
    End If
    Exit Sub
Error:
    MsgBox ("Unable to set control: " + txtDate.Name + " to value: " + CStr(vDate) + ". Date is in invalid format")
End Sub

Public Sub HandleDateEx(ByRef txtDate As MSGEditBox, ByRef vDate As Variant, ByVal uOperation As ControlOperation)
    
    Dim bRet As Boolean
    Dim bNull As Boolean
        
    If (uOperation = SET_CONTROL_VALUE) Then
        'If a null date is entered - clear the text control.
        If IsNull(vDate) Then
            txtDate.Text = ""
        ElseIf IsDate(vDate) Then
            'Set the contents of the text control now.
            txtDate.Text = Format(vDate, "Short Date")
        Else
            'Raise an error, the value is not a valid date.
            g_clsErrorHandling.RaiseError errGeneralError, "The value specified is not a date."
        End If
    Else
        If txtDate.Text = BLANK_DATE_MASK Then
            'Return NULL if the control has an empty mask.
            vDate = Null
        ElseIf IsDate(txtDate.Text) Then
            'If the date is valid, return it.
            vDate = Format(txtDate.Text, "Short Date")
        Else
            'If the date is not valid, display a message
            g_clsErrorHandling.DisplayError "The value is not a valid date."
            txtDate.SetFocus
        End If
    End If
    
    Exit Sub
    
End Sub

Public Sub AddLine(lvListView As MSGListView, cLine As Collection)
    Dim nEntry As Long
    Dim nNumEntries As Long
    Dim listLine As ListItem
    
    Set listLine = lvListView.ListItems.Add()
    
    nNumEntries = cLine.Count
    
    For nEntry = 1 To nNumEntries
        If (nEntry = 1) Then
            listLine.Text = cLine(nEntry)
        Else
            listLine.SubItems(nEntry - 1) = cLine(nEntry)
        End If
    Next

End Sub

Public Sub PopulateFromRecordset(lvListView As MSGListView, clsTable As TableAccess, Optional bClear As Boolean = True)
    On Error GoTo Failed
    Dim nThisField As Integer
    Dim nFieldCount As Integer
    Dim sField As String
    Dim sObjectDescription As String
    Dim sExtraValue As String
    Dim colListLine As Collection
    Dim colMatchFields As Collection
    Dim colMatchValues As Collection
    Dim rs As ADODB.Recordset
    Dim vValue As Variant
    Dim nType As DataTypeEnum
    Dim clsPopulateDetails As PopulateDetails
    
    On Error GoTo Failed
    BeginWaitCursor
    
    If bClear = True Then
        lvListView.ListItems.Clear
    End If
    
    Set rs = clsTable.GetRecordSet()
    
    If Not rs Is Nothing Then
        Set colMatchFields = New Collection
        Set colMatchFields = clsTable.GetKeyMatchFields()
        
        nFieldCount = colMatchFields.Count
        
        If rs.RecordCount > 0 Then
            clsTable.MoveFirst
            
            Do While Not rs.EOF
                Set colListLine = New Collection
                Set colMatchValues = New Collection
                
                For nThisField = 1 To nFieldCount
                    sField = colMatchFields(nThisField)
                    
                    If Len(sField) > 0 Then
                        vValue = rs(sField).Value
                        nType = rs(sField).Type
                        
                        ' DJP SQL Server port. SQL Server reports GUID as adBinary not adVarBinary
                        colMatchValues.Add vValue
                    End If
                Next
                
                clsTable.GetRowOfData colListLine
                On Error Resume Next
                sObjectDescription = colListLine(OBJECT_DESCRIPTION_KEY)

                If Err.Number = 0 Then
                    colListLine.Remove OBJECT_DESCRIPTION_KEY
                End If
                Err.Clear
                
                sExtraValue = colListLine(OBJECT_EXTRA_VALUE)
                If Err.Number = 0 Then
                    colListLine.Remove OBJECT_EXTRA_VALUE
                End If
                
                On Error GoTo Failed
                
                Set clsPopulateDetails = New PopulateDetails
                clsPopulateDetails.SetKeyMatchValues colMatchValues
                clsPopulateDetails.SetExtra sExtraValue
                clsPopulateDetails.SetObjectDescription sObjectDescription
                lvListView.AddLine colListLine, clsPopulateDetails
                
                rs.MoveNext
            Loop
        End If
    End If
    
    EndWaitCursor
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

    EndWaitCursor
End Sub


Friend Sub PopulateDataCombo(combo As MSGDataCombo, clsTableAccess As TableAccess)
'GD BMIDS.16/05/2002. AQR : BMIDS00009. Allow data combo to have concatenated contents.
    Dim sField As String
    Dim rs As ADODB.Recordset
    Dim sError As String
    
    sError = "PopulateDataCombo: "
    
    On Error GoTo Failed
    If Not clsTableAccess Is Nothing Then
        
        Set rs = clsTableAccess.GetRecordSet()
        
        If Not rs Is Nothing Then
            'GD BMIDS
            If rs.fields.Count >= 1 Then
                sField = rs.fields(0).Name
                Set combo.RowSource = rs
                
                combo.ListField = sField
            Else
                g_clsErrorHandling.RaiseError errGeneralError, sError + " No records to populate"
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, sError + " Recordset is empty"
        End If

    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Friend Sub PopulateCollectionFromTable(clsTableAccess As TableAccess, colFields As Collection, colText As Collection, colValues As Collection)
    Dim bRet As Boolean
    Dim rs As ADODB.Recordset
    Dim sTextField As String
    Dim sValueField As String
    Dim sValue As String
    Dim vText As Variant
    
    On Error GoTo Failed
    
    bRet = True
    If colFields.Count > 0 Then
        sTextField = colFields(1)
    End If
    
    If colFields.Count > 1 Then
        sValueField = colFields(2)
    End If
    
    Set rs = clsTableAccess.GetRecordSet()
    
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                vText = rs(sTextField)
                
                If Not IsNull(vText) Then
' TW 09/04/2007 EP2_2332 - Deals with a situation where the associated 'Value' field is null.
                    If Len(sValueField) > 0 Then
                        If Not IsNull(rs(sValueField)) Then
                            colText.Add CStr(vText)
' TW 09/04/2007 EP2_2332 End
                            sValue = rs(sValueField)
                            colValues.Add sValue
                        End If
                    Else
                        colText.Add CStr(vText)
' TW 09/04/2007 EP2_2332
                    End If
' TW 09/04/2007 EP2_2332 End
                End If
                rs.MoveNext
            Loop
        Else
            bRet = False
            g_clsErrorHandling.RaiseError errGeneralError, "No records found in " + clsTableAccess.GetTable()
            'MsgBox "No records found in " + clsTableAccess.GetTable()
        End If
    Else
        bRet = False
        g_clsErrorHandling.RaiseError errGeneralError, "No records found in " + clsTableAccess.GetTable()
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub CancelForm(frmForm As Form)
    EndWaitCursor
    Set m_frmCurrent = frmForm
    g_Timer.Enabled = True
End Sub

Public Function GetCurrentForm() As Form
    Set GetCurrentForm = m_frmCurrent
End Function

Public Sub PopulateChannel(cboChannel As MSGComboBox)
    Dim clsChannel As New DistributionChannelTable
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim colFields As New Collection
    
    On Error GoTo Failed
    Set clsTableAccess = clsChannel
    
    Set rs = clsTableAccess.GetTableData(POPULATE_ALL)
    
    If Not rs Is Nothing Then
        If clsTableAccess.RecordCount() > 0 Then
            Set colFields = clsChannel.GetComboFields()
            
            g_clsFormProcessing.PopulateComboFromTable cboChannel, clsTableAccess, colFields
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "No channels exist"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "No channels exist"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SetTabstops(frmToSet As Form)
    Dim C As Control
    Dim nLeft As Long
    Dim objParent As Object
    Dim bDone As Boolean
    
    On Error Resume Next
    For Each C In frmToSet.Controls
        ' Is the control placed directly onto the tab?
        If TypeOf C.Container Is SSTab Then
            'Not all controls have the TabStop property

            C.TabStop = C.Left > 0
        ElseIf TypeOf C.Container Is Frame Then
            ' The control is held inside a frame (one or more)
            Set objParent = C.Container
            Dim bSetTab As Boolean
            bDone = False
            
            ' Loop finding the frame directly beneath the sstab.
            While bDone = False And Not objParent Is Nothing
                If TypeOf objParent Is Frame Then
                    If TypeOf objParent.Container Is SSTab Then
                        bDone = True
                        If TypeOf C Is OptionButton Then
                                Dim a As Boolean
                                a = True
                                Dim b As String
                                Dim d As Integer
                                b = C.Name
                                d = C.Index
'                            MsgBox "option button: " & c.Name
                        End If
                        Err.Clear
                        nLeft = objParent.Left
                        
                        If Err.Number = 0 Then
                            Err.Clear
                            bSetTab = nLeft > 0
                            
                            If bSetTab = True Then
                                ' Need to set tabstop to true, if it was originally
                                If Len(C.Tag) > 0 Then
                                    If CBool(C.Tag) = True Then
                                        bSetTab = True
                                    Else
                                        bSetTab = False
                                    End If
                                ElseIf C.TabStop = True Then
                                    bSetTab = True
                                Else
                                    bSetTab = False
                                End If
                            Else
                                bSetTab = False
                            End If
                            
                            If C.TabStop = True Then
                                C.Tag = True
                            End If
                            
                            C.TabStop = bSetTab
                        End If
                    Else
                        Set objParent = objParent.Container
                    End If
                Else
                    Set objParent = objParent.Container
                End If
            Wend
        End If
    Next
End Sub

Public Sub SetText(ctrl As Object, bEnable As Boolean, Optional sText As String)
    On Error GoTo Failed
    If TypeOf ctrl Is MSGEditBox Then
        Dim objEdit As MSGEditBox
        Set objEdit = ctrl
        
        objEdit.Text = sText
        objEdit.Enabled = bEnable
    
    ElseIf TypeOf ctrl Is Label Then
        ctrl.Enabled = bEnable
    
    ElseIf TypeOf ctrl Is MSGComboBox Then
        Dim objCombo As MSGComboBox
        Set objCombo = ctrl
        
        objCombo.SelText = sText
        objCombo.Enabled = bEnable
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function GetText(ctrl As Object) As String
    On Error GoTo Failed
    Dim sText As String
    If TypeOf ctrl Is MSGEditBox Then
        Dim objEdit As MSGEditBox
        Set objEdit = ctrl
        sText = objEdit.Text
    
    ElseIf TypeOf ctrl Is MSGComboBox Then
        Dim objCombo As MSGComboBox
        Set objCombo = ctrl
        sText = objCombo.SelText
    End If
    
    GetText = sText
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Function FormatWildcardString(ByVal strIn As String) As String
    On Error GoTo Failed
    Dim strOut As String
    Dim nLen As Integer
    Dim nThisChar As Integer
    Dim sThisChar As String
    
    'AW EP2_2562 24/04/07
    'CheckApostrophes() is called via g_clsSQLAssistSP.FormatString()
    'Remove this call which results in a call to CheckApostrophes()
    ''strIn = FormatString(strIn)
    
    nLen = Len(strIn)
    
    For nThisChar = 1 To nLen
        sThisChar = Mid(strIn, nThisChar, 1)
        
        If sThisChar = cstrWILDCARD Then
            sThisChar = cstrWILDCARDSQL
        End If
        strOut = strOut + sThisChar
    Next

    FormatWildcardString = g_clsSQLAssistSP.FormatString(strOut)

    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub PopulateAvailableTargets(cboDatabase As MSGComboBox, Optional sSourceService As String, Optional sSourceUserID As String)
    On Error GoTo Failed
    Dim bDone As Boolean
    Dim colConnections As Collection
    Dim clsSupervisorConnection As SupervisorConnection
    Dim sTargetService As String
    Dim sTargetUserID As String
    Dim sDatabaseKey As String
    'Dim sactivedatabase As String
    
    Set colConnections = g_clsDataAccess.GetConnectionList()
    
    If colConnections.Count > 0 Then
        For Each clsSupervisorConnection In colConnections
            sTargetService = clsSupervisorConnection.GetDatabaseName()
            sTargetUserID = clsSupervisorConnection.GetUserID()
            
            If sTargetService <> sSourceService Or sTargetUserID <> sSourceUserID Then
                sDatabaseKey = GetConnectionKey(sTargetService, sTargetUserID)
                cboDatabase.AddItem sDatabaseKey
            End If
        Next
        'cboDatabase.Enabled = True
    Else
        cboDatabase.Enabled = False
        g_clsErrorHandling.RaiseError errGeneralError, "No database targets defined"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function CheckForDuplicates(clsTableToCheck As TableAccess) As Boolean
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim rsRecord As ADODB.Recordset
    Dim rsRecordClone As ADODB.Recordset
    Dim colMatches As Collection
    Dim sRecord As String
    Dim bDuplicated As Boolean
    Dim colDupFields As Collection
    Dim sFunctionName  As String
    Dim sField As Variant
    Dim sDupFields As String
    
    sFunctionName = "CheckForDuplicates"
    bDuplicated = False
    
    Set colDupFields = clsTableToCheck.GetDuplicateKeys()
    
    If colDupFields Is Nothing Then
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Duplicate key list is empty"
    End If
    
    Set rsRecord = clsTableToCheck.GetRecordSet()
    Set rsRecordClone = rsRecord.Clone
    
    If Not rsRecordClone Is Nothing Then
        If clsTableToCheck.RecordCount > 0 Then

            clsTableToCheck.SetRecordSet rsRecordClone
            ' DJP Only move once the recordset has been set, otherwise we are moving the wrong one.
            clsTableToCheck.MoveFirst
            
            Set colMatches = New Collection
            
            Do While Not rsRecordClone.EOF And Not bDuplicated
                sDupFields = ""
                For Each sField In colDupFields
                    sRecord = rsRecordClone(sField) 'm_clsValuationType.GetValuationTypeText()
                    sDupFields = sDupFields + "," + sRecord
                Next
            
                On Error Resume Next
                colMatches.Add "Anything", sDupFields
                        
                If Err.Number <> 0 Then
                    ' Exists already
                    bDuplicated = True
                End If
                On Error GoTo Failed
                rsRecordClone.MoveNext

            Loop
            clsTableToCheck.SetRecordSet rsRecord
        End If

        rsRecordClone.Close
    
    End If
    
    CheckForDuplicates = bDuplicated
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Friend Sub PopulateCombo(sGroup As String, combo As MSGComboBox, Optional bIncludeOptionNone As Boolean = True)
    Dim bRet As Boolean
    Dim colNames As New Collection
    Dim colValues As New Collection
    Dim clsComboTable As New ComboValueTable
    Dim clsTableAccess As TableAccess
    
    On Error GoTo Failed
    bRet = True
    
    bRet = clsComboTable.GetComboValues(sGroup, colNames, colValues)
    
    If bRet = True Then
        If bIncludeOptionNone Then
            colNames.Add COMBO_NONE
        End If
        colValues.Add colNames.Count
        
        combo.SetListTextFromCollection colNames, colValues

    End If
    
    If bRet = False Then
        g_clsErrorHandling.RaiseError errGeneralError, "Unable to find combo: " + sGroup
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub HandleComboText(cboCombo As Object, vVal As Variant, enumOperation As ControlOperation)
    On Error GoTo Failed
    
    If (enumOperation = SET_CONTROL_VALUE) Then
        If (Not IsNull(vVal) And Len(vVal) > 0) Then
            cboCombo.Text = CStr(vVal)
        End If
    Else
        vVal = cboCombo.SelText
    
        If CStr(vVal) = COMBO_NONE Then
            vVal = ""
        End If
    
    End If
Failed:

End Sub

Public Sub SetComboFromValidation(cboCombo As Object, sGroup As String, sValidation As String)
    On Error GoTo Failed
    Dim clsValidation As ComboValidationTable
    Dim clsTableAccess As TableAccess
    Dim sFunctionName As String
    Dim sValueID As String
    
    sFunctionName = "SetComboFromValidation"
    
    If Len(sValidation) > 0 Then
        Set clsValidation = New ComboValidationTable
        Set clsTableAccess = clsValidation
        
        clsValidation.GetValueFromValidation sGroup, sValidation
        
        If clsTableAccess.RecordCount() > 0 Then
            sValueID = clsValidation.GetValueID()

        End If
    End If
    
    g_clsFormProcessing.HandleComboExtra cboCombo, sValueID, SET_CONTROL_VALUE
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub GetComboValidation(cboCombo As Object, sGroup As String, sValidation As String)
    On Error GoTo Failed
    Dim vIndex As Variant
    Dim sFunctionName As String
    Dim vValueID As Variant
    
    sFunctionName = "GetComboValidation"
    
    ' Get the index
    If Len(sGroup) = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Group name is empty"
    End If
    
    HandleComboExtra cboCombo, vValueID, GET_CONTROL_VALUE
    
    If Not IsEmpty(vValueID) Then
        If Len(CStr(vValueID)) > 0 Then
            Dim clsValidation As ComboValidationTable
            Dim clsValidationTableAccess As TableAccess
            Dim colMatchValues As Collection
            
            Set clsValidation = New ComboValidationTable
            Set clsValidationTableAccess = clsValidation
            Set colMatchValues = New Collection
            
            colMatchValues.Add sGroup
            colMatchValues.Add CStr(vValueID)
            clsValidationTableAccess.SetKeyMatchValues colMatchValues
                    
            clsValidationTableAccess.GetTableData
                
            If clsValidationTableAccess.RecordCount() > 0 Then
                clsValidationTableAccess.MoveFirst
                sValidation = clsValidation.GetValidationType()
            End If
            
            If Len(sValidation) = 0 Then
                g_clsErrorHandling.RaiseError errGeneralError, sFunctionName + ": Unable to obtain Validation for Group " + sGroup + " and ValueID " + CStr(vValueID)
            End If
        End If
    End If
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub HandleComboExtra(cboCombo As Object, vVal As Variant, enumOperation As ControlOperation)
    Dim nIndex As Integer
    Dim sValueID As String
    Dim bFound As Boolean
    
    On Error GoTo Failed
    
    If (enumOperation = SET_CONTROL_VALUE) Then
        If Not IsNull(vVal) Then
            If Len(vVal) > 0 Then
                Dim nCount As Integer
                Dim nThisItem As Integer
                
                nThisItem = 0
                nCount = cboCombo.ListCount
                bFound = False
                
                While bFound = False And nThisItem <= nCount
                    If cboCombo.GetExtra(nThisItem) = CStr(vVal) Then
                        cboCombo.ListIndex = nThisItem
                        bFound = True
                    Else
                        nThisItem = nThisItem + 1
                    End If
                Wend
                If bFound = False Then
                    cboCombo.ListIndex = cboCombo.ListCount - 1
                End If
                
            Else
                cboCombo.ListIndex = cboCombo.ListCount - 1
            End If
        End If
    Else
        nIndex = cboCombo.ListIndex
        
        If nIndex >= 0 Then
            vVal = cboCombo.GetExtra(nIndex)
            
            'If CStr(cboCombo.List(vVal - 1)) = COMBO_NONE Then
            If CStr(cboCombo.List(nIndex)) = COMBO_NONE Then
                Dim sType As String
                sType = UCase(TypeName(vVal))
                
                If sType = "STRING" Or sType = "EMPTY" Then
                    vVal = ""
                End If
            End If
        ElseIf nIndex = -1 Then
            vVal = ""
        End If
        
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub HandleComboIndex(cboCombo As Object, vVal As Variant, enumOperation As ControlOperation)
    On Error GoTo Failed
    Dim a As MSGComboBox

    If (enumOperation = SET_CONTROL_VALUE) Then
        If (Not IsNull(vVal) And Len(vVal) > 0) Then
            cboCombo.ListIndex = CInt(vVal)
        Else
            ' last one is <NONE>
            cboCombo.ListIndex = cboCombo.ListCount - 1
        End If
    Else
        vVal = cboCombo.ListIndex
        If CInt(vVal) < 0 Then
            vVal = Null
        End If
    End If
Failed:
End Sub

Public Sub HandleDataComboText(cboCombo As MSGDataCombo, vVal As Variant, enumOperation As ControlOperation)
    On Error GoTo Failed
    If (enumOperation = SET_CONTROL_VALUE) Then
        If (Not IsNull(vVal) And Len(vVal) > 0) Then
            cboCombo.Text = CStr(vVal)
        End If
    Else
        vVal = cboCombo.SelText
        
        If CStr(vVal) = COMBO_NONE Then
            vVal = ""
        End If
            
    End If
Failed:
End Sub

Friend Sub PopulateComboFromTable(combo As MSGComboBox, clsTableAccess As TableAccess, colFields As Collection)
    Dim colText As New Collection
    Dim colValues As New Collection
    
    On Error GoTo Failed

    PopulateCollectionFromTable clsTableAccess, colFields, colText, colValues
    colText.Add COMBO_NONE
    
    If colValues.Count > 0 Then
        colValues.Add colText.Count
    End If
    
    combo.SetListTextFromCollection colText, colValues
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function DoesSwapValueExist(SwapList As MSGHorizontalSwapList, sValue As String) As Boolean
    On Error GoTo Failed
    Dim nCount As Integer
    Dim nThisItem As Integer
    Dim bFound As Boolean
    Dim sSecondValue As String
    Dim colValue As Collection

    nCount = SwapList.GetSecondCount()

    bFound = False
    nThisItem = 1
    
    While bFound = False And nThisItem <= nCount
        Set colValue = New Collection
        Set colValue = SwapList.GetLineSecond(nThisItem)
        
        If colValue.Count = 1 Then
            sSecondValue = colValue(1)
        
            If sSecondValue = sValue Then
                bFound = True
            End If
        End If
        
        nThisItem = nThisItem + 1
    Wend
    
    DoesSwapValueExist = bFound
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Function GetCollectionFromRecordset(clsTableAccess As TableAccess) As Collection
    Dim sUserID As String
    Dim rs As ADODB.Recordset
    Dim sField As String
    Dim col As Collection
    
    On Error GoTo Failed
    Set col = New Collection
    
    sField = clsTableAccess.GetCollectionField()
    
    
    Set rs = clsTableAccess.GetRecordSet

    If clsTableAccess.RecordCount() > 0 Then
        Do While Not rs.EOF
            col.Add clsTableAccess.GetVal(sField)
            clsTableAccess.MoveNext
        Loop
    
    End If
    
    Set GetCollectionFromRecordset = col
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Function SetDataComboTextFromValueID(sValueID As String, clsComboVal As ComboValidationTable) As String

    On Error GoTo Failed
    Dim sComboValueID As String
    Dim sValueName As String
    Dim rs As ADODB.Recordset
    Dim bExitLoop As Boolean
    
    Set rs = TableAccess(clsComboVal).GetRecordSet
    TableAccess(clsComboVal).ValidateData "SetDataComboTextFromValueID"
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF And Not bExitLoop
            sComboValueID = clsComboVal.GetValueID
            If sValueID = sComboValueID Then
                sValueName = clsComboVal.GetValueName()
                bExitLoop = True
            End If
            rs.MoveNext
        Wend
    End If
    
    SetDataComboTextFromValueID = sValueName
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Function

Public Function DisplayMessage(sText As String, Optional sTitle As String) As Variant

    On Error GoTo Failed
    Dim vResponse As Variant

    If Len(sTitle) = 0 Then
        sTitle = App.Title
    End If
    
    vResponse = MsgBox(sText, vbYesNo, sTitle)
    
    DisplayMessage = vResponse
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub EnableFrameFields(objParent As Frame, Optional bEnable As Boolean = True)
    On Error GoTo Failed
    Dim sName As String
    Dim objThisControl As Control
    Dim objControls As Controls
    
    If Not objParent Is Nothing Then
        sName = objParent.Name
        On Error Resume Next
        
        For Each objThisControl In objParent.Parent.Controls
            If objThisControl.Container.Name = sName Then
                If Err.Number = 0 Then
                    objThisControl.Enabled = bEnable
                Else
                    Err.Clear
                End If
            End If
        Next
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Cannot clear fields - parent is empty"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Function PopulateComboByValidation(ByRef cboComboToPopulate As MSGComboBox, ByVal sGroupName As String, ByVal sValidationType As String, Optional bIncludeOptionNone As Boolean = True) As Boolean

    Dim sSQL As String
    Dim colNames As Collection
    Dim colValues As Collection
    Dim rsComboItems As ADODB.Recordset

    On Error GoTo Failed
    
    'Create our collections.
    Set colNames = New Collection
    Set colValues = New Collection
    
    '"With" to make the follwing SQL more readable.
    With g_clsSQLAssistSP
        'Construct the SQL statement.
        sSQL = "SELECT  COMBOVALUE.VALUENAME, " & _
               "        COMBOVALUE.VALUEID " & _
               "FROM    COMBOVALUE, " & _
               "        COMBOVALIDATION " & _
               "WHERE   COMBOVALUE.GROUPNAME = COMBOVALIDATION.GROUPNAME AND " & _
               "        COMBOVALUE.VALUEID = COMBOVALIDATION.VALUEID AND " & _
               "        COMBOVALUE.GROUPNAME = " & .FormatString(sGroupName) & " AND " & _
               "        COMBOVALIDATION.VALIDATIONTYPE = " & .FormatString(sValidationType)
    End With
    
    'Obtain a recordset of combo items/ids.
    Set rsComboItems = g_clsDataAccess.GetTableData(, sSQL)
    
    'Populate the collections from this recordset.
    Do While Not rsComboItems.EOF
        colNames.Add rsComboItems.fields("ValueName").Value
        colValues.Add rsComboItems.fields("ValueID").Value
    
        rsComboItems.MoveNext
    Loop
    
    If bIncludeOptionNone Then
        colNames.Add COMBO_NONE
        colValues.Add colNames.Count
    End If
    
    'Populate the combo from these collections.
    cboComboToPopulate.SetListTextFromCollection colNames, colValues
    
    'Indicate success.
    PopulateComboByValidation = True
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Public Sub GetListViewKeys(lv As MSGListView, colKeyValues As Collection)
    On Error GoTo Failed
    Dim sSetKey As String
    Dim sStartDateKey As String
    Dim sFeeType As String
    Dim sAppType As String
    Dim nListIndex As Integer
    Dim clsPopulateDetails As PopulateDetails
    
    If Not lv.SelectedItem Is Nothing Then
        nListIndex = lv.SelectedItem.Index
        Set clsPopulateDetails = lv.GetExtra(nListIndex)
        
        If Not clsPopulateDetails Is Nothing Then
            Set colKeyValues = clsPopulateDetails.GetKeyMatchValues()
        
            If colKeyValues.Count = 0 Then
                g_clsErrorHandling.RaiseError errGeneralError, "Unable to obtain keys"
            End If
        Else
            g_clsErrorHandling.RaiseError errGeneralError, " Populate Details object not set"
        End If
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "No listview selection"
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub CloseFindDialogs()
    On Error GoTo Failed
    
    Unload frmFindApplication
    
    Exit Sub
Failed:
End Sub

