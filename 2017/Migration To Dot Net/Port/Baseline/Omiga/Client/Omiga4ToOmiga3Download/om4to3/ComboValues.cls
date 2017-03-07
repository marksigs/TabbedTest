VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "ComboValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      {ThisFileName}
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   A Helper class to assist with translation of combo values
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MH     07-03-00    Created
'LD     07/11/00    Explicity close recordsets
'DM    16/05/2001 AQR SYS2313 Outer join changes.
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
'=============================================
'Constant Declaration Section
'=============================================

'=============================================
'Variable Declaration Section
'=============================================
Private m_colGroupValuesByID As Collection
Private m_colGroupValuesByType As Collection
Private m_colGroupInfo As Collection
Private m_conOmiga4 As Connection
Private m_objContext As ObjectContext
Private m_objErrAssist As New ErrAssist

Public Function GetByValueID(ByVal vstrGroupName As String, ByVal vstrValueID As String) As ComboValue
    On Error GoTo GetByValueIDErr
    
    If vstrValueID = vbNullString Then
        Set GetByValueID = Nothing
    Else
        Set GetByValueID = m_colGroupValuesByID(UCase$(vstrGroupName & "." & vstrValueID))
    End If
    Exit Function

GetByValueIDErr:
    AddGroup vstrGroupName
    Set GetByValueID = m_colGroupValuesByID(UCase$(vstrGroupName & "." & vstrValueID))
    Exit Function
       
       
End Function

Public Function GetByValue(ByVal vstrGroupName As String, ByVal vstrValue As String) As ComboValue
    On Error GoTo GetByValueErr
    If vstrValue = vbNullString Then
        Set GetByValue = Nothing
    Else
        Set GetByValue = m_colGroupValuesByType(UCase$(vstrGroupName & "." & vstrValue))
    End If
    
    Exit Function
    
GetByValueErr:
    AddGroup vstrGroupName
    Set GetByValue = m_colGroupValuesByType(UCase$(vstrGroupName & "." & vstrValue))
    Exit Function
    

End Function

Public Sub CreateWithPreload(ByVal vconDataBase As Connection)
Dim t As Single
    Dim fldGroup As Field
    Dim fldValueName As Field
    Dim fldValue As Field
    Dim fldValueID As Field
    Dim fldOM3ValueID As Field
    Dim rstData As Recordset
    Dim objValue As ComboValue
    
    Set m_colGroupValuesByID = New Collection
    Set m_colGroupValuesByType = New Collection
    Set m_colGroupInfo = New Collection
    
    'DM 16/05/2001 AQR SYS2313
    #If GENERIC_SQL Then
        Dim aComm As ADODB.Command
    #End If
    
    'Load up everything because it is quicker in the long run....
    'DM 16/05/2001 AQR SYS2313
    #If GENERIC_SQL Then
        Set aComm = New ADODB.Command
    
        With aComm
            Set .ActiveConnection = vconDataBase
            .CommandType = adCmdStoredProc
            .CommandText = "usp_CreateWithPreload" ' The stored procedure name
        End With
    
        Set rstData = aComm.Execute
        
        Set aComm = Nothing
        
    #Else
        Set rstData = vconDataBase.Execute("select COMBOVALUE.VALUEID, COMBOVALUE.GROUPNAME, COMBOVALUE.OM3VALUEID, COMBOVALUE.VALUENAME, COMBOVALIDATION.VALIDATIONTYPE FROM COMBOVALUE, COMBOVALIDATION WHERE (COMBOVALIDATION.GROUPNAME (+) = COMBOVALUE.GROUPNAME) AND (COMBOVALIDATION.VALUEID (+) = COMBOVALUE.VALUEID)")
    #End If
    Set fldGroup = rstData.Fields("GROUPNAME")
    Set fldValueName = rstData.Fields("VALUENAME")
    Set fldValueID = rstData.Fields("VALUEID")
    Set fldValue = rstData.Fields("VALIDATIONTYPE")
    Set fldOM3ValueID = rstData.Fields("OM3VALUEID")
    
    
    Do While Not rstData.EOF
        If m_objContext Is Nothing Then
            Set objValue = New ComboValue
        Else
            Set objValue = m_objContext.CreateInstance(App.Title & ".ComboValue")
        End If
        
        objValue.Create fldGroup.Value, fldValueID.Value, _
                SafeNull(fldOM3ValueID.Value), SafeNull(fldValue.Value), SafeNull(fldValueName.Value)
        
        AddValueByID objValue
        AddValueByType objValue
        
        rstData.MoveNext
    Loop
        rstData.Close
        Set rstData = Nothing
End Sub

Public Sub Create(ByVal vconDataBase As Connection)
    Set m_conOmiga4 = vconDataBase
    Set m_colGroupValuesByID = New Collection
    Set m_colGroupValuesByType = New Collection
    Set m_colGroupInfo = New Collection
End Sub

Private Sub AddGroup(ByVal vstrGroupName As String)
' header ----------------------------------------------------------------------------------
' description:
'    Loads up data for a particular group. This method is called when an
'    element in a group was not found.
'
' pass: vstrGroupName       The name of the group to load
'
' return:
'------------------------------------------------------------------------------------------

    Const strMethodName = "AddGroup"
    On Error GoTo AddGroupVBErr
    Dim strValues As String
    Dim fldGroup As Field
    Dim fldValueName As Field
    Dim fldValue As Field
    Dim fldValueID As Field
    Dim fldOM3ValueID As Field
    Dim rstData As Recordset
    Dim objValue As ComboValue
    
    'DM 16/05/2001 AQR SYS2313
    #If GENERIC_SQL Then
        Dim aComm As ADODB.Command
    #End If
      
    'Only get a group that doesn't exist
    
    If Not GroupExists(vstrGroupName) Then
        'DM 16/05/2001 AQR SYS2313
        #If GENERIC_SQL Then
            Set aComm = New ADODB.Command

            With aComm
                Set .ActiveConnection = m_conOmiga4
                .CommandType = adCmdStoredProc
                .CommandText = "usp_AddGroup" ' The stored procedure name
            End With
            
            aComm.Parameters.Append aComm.CreateParameter("p_GroupName", adBSTR, adParamInput, Len(vstrGroupName), vstrGroupName)

            Set rstData = aComm.Execute
            
            Set aComm = Nothing
            
        #Else
            Set rstData = m_conOmiga4.Execute("select COMBOVALUE.VALUEID, COMBOVALUE.GROUPNAME, COMBOVALUE.OM3VALUEID, COMBOVALUE.VALUENAME, COMBOVALIDATION.VALIDATIONTYPE FROM COMBOVALUE, COMBOVALIDATION WHERE (COMBOVALIDATION.GROUPNAME (+) = COMBOVALUE.GROUPNAME) AND (COMBOVALIDATION.VALUEID (+) = COMBOVALUE.VALUEID) AND COMBOVALUE.GROUPNAME='" & vstrGroupName & "'")
        #End If
        Set fldValueName = rstData.Fields("VALUENAME")
        Set fldValueID = rstData.Fields("VALUEID")
        Set fldValue = rstData.Fields("VALIDATIONTYPE")
        Set fldOM3ValueID = rstData.Fields("OM3VALUEID")
        
        strValues = vbNullString
        
        Do While Not rstData.EOF
            If m_objContext Is Nothing Then
                Set objValue = New ComboValue
            Else
                Set objValue = m_objContext.CreateInstance(App.Title & ".ComboValue")
            End If
            
            objValue.Create vstrGroupName, fldValueID.Value, _
                    SafeNull(fldOM3ValueID.Value), SafeNull(fldValue.Value), SafeNull(fldValueName.Value)
            
            AddValueByID objValue
            AddValueByType objValue
            
            strValues = strValues & ", " & objValue.ValueID & "=" & objValue.ValueType
            
            rstData.MoveNext
        Loop
                rstData.Close
                Set rstData = Nothing

        strValues = Mid$(strValues, 3)
        m_colGroupInfo.Add strValues, vstrGroupName
    End If
    Exit Sub

AddGroupVBErr:
          
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Private Function SafeNull(ByVal vvarValue As Variant) As Variant
    If IsNull(vvarValue) Then
        SafeNull = vbNullString
    Else
        SafeNull = vvarValue
    End If
End Function

Private Sub AddValueByID(ByVal vobjCombo As ComboValue)
    On Error GoTo AddValueByIDVBErr
    'Assume that the object can be added ok.
    
    'Dont catalog nulls
    If vobjCombo.ValueID = vbNullString Then Exit Sub
    
    m_colGroupValuesByID.Add vobjCombo, vobjCombo.ValueIDKey
    Exit Sub

AddValueByIDVBErr:
    m_colGroupValuesByID(vobjCombo.ValueIDKey).AddSibling vobjCombo
        
End Sub

Private Sub AddValueByType(ByVal vobjCombo As ComboValue)
    On Error GoTo AddValueByTypeVBErr
    'Assume that the object can be added ok.
    
    'Dont catalog nulls
    If vobjCombo.ValueType = vbNullString Then Exit Sub
    
    m_colGroupValuesByType.Add vobjCombo, vobjCombo.ValueTypeKey
    Exit Sub

AddValueByTypeVBErr:
    m_colGroupValuesByType(vobjCombo.ValueTypeKey).AddSibling vobjCombo
        
End Sub

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_colGroupValuesByID = Nothing
    Set m_colGroupValuesByType = Nothing
    Set m_conOmiga4 = Nothing
End Sub

Public Function GroupInfo(ByVal vstrGroupName As String) As String
    Dim varValues As Variant
    
    On Error Resume Next
    
    varValues = m_colGroupInfo(vstrGroupName)
    If Err.Number = 0 Then
        GroupInfo = vstrGroupName & " Valid Values(" & varValues & ")"
    Else
        GroupInfo = "Combo Group """ & vstrGroupName & """ not found"
    End If

End Function

Private Function GroupExists(ByVal vstrGroupName As String) As Boolean
    Dim varValues As Variant
    
    On Error Resume Next
    varValues = m_colGroupInfo(vstrGroupName)
    GroupExists = (Err.Number = 0)

End Function

