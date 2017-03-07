VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "DataCacheRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      DataCacheRecordset.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   A simple wrapper for the recordset so that its key can be seen
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MH     23/02/00    Created
'------------------------------------------------------------------------------------------

Option Explicit

Implements IDataCacheItem

'=============================================
'Constant Declaration Section
'=============================================

'=============================================
'Variable Declaration Section
'=============================================
Private m_strName As String
Private m_rstRecordset As Recordset
Private m_lngRecordCount As Long
Private m_strSQL As String
Private m_objErrAssist As New ErrAssist

Private Property Get IDataCacheItem_AtEndOfData() As Boolean
    IDataCacheItem_AtEndOfData = m_rstRecordset.EOF
End Property

Private Function IDataCacheItem_Clone(ByVal vstrNewName As String) As IDataCacheItem
    Const strMethodName = "IDataCacheItem_Clone"
    On Error GoTo IDataCacheItem_CloneVBErr
    Dim objContext As ObjectContext
    Dim objNew As DataCacheRecordset
    
    
    Set objContext = GetObjectContext()
    
    If objContext Is Nothing Then
        Set objNew = New DataCacheRecordset
    Else
        Set objNew = objContext.CreateInstance(App.Title & ".DataCacheRecordset")
    End If
    
    objNew.Create vstrNewName, m_rstRecordset
    Set IDataCacheItem_Clone = objNew
    
    Exit Function

IDataCacheItem_CloneVBErr:
          
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Property Get IDataCacheItem_Count() As Long
' header ----------------------------------------------------------------------------------
' description:
'    Returns a Count of records. If the recordset is cursor based then a
'    count will be derived and cached. Otherwise the value will be taken
'    from the recordset directly.
'
' pass:
'
' return:  The number of records
'------------------------------------------------------------------------------------------
On Error GoTo RecordCountVBErr
    
    If m_lngRecordCount < 0 Then
        
        With m_rstRecordset
            If m_rstRecordset.RecordCount >= 0 Then
                m_lngRecordCount = m_rstRecordset.RecordCount
            Else
                If Not m_rstRecordset.BOF Then
                    m_rstRecordset.MoveFirst
                    m_lngRecordCount = 0
                    Do While Not m_rstRecordset.EOF
                        m_lngRecordCount = m_lngRecordCount + 1
                        m_rstRecordset.MoveNext
                    Loop
                    m_rstRecordset.MoveFirst
                Else
                    m_lngRecordCount = 0
                End If
            End If
        End With
        
    End If

    IDataCacheItem_Count = m_lngRecordCount
    Exit Property
    
RecordCountVBErr:
    m_lngRecordCount = 0
    IDataCacheItem_Count = m_lngRecordCount
    
End Property

' description:
' Returns a particular row from the recordset.
' If the recordset supports Absoluteposition then reposition within the recordset
' If it does not then do it the hard way

Private Function IDataCacheItem_Entry(ByVal vlngIndex As Long) As IDataCacheItem
Dim lngRow As Long
    
    If Not m_rstRecordset.BOF Then
        If m_rstRecordset.AbsolutePosition <> adPosUnknown Then
            'Absoluteposition is supported. Hurrah. Return the row
            'If you go past the end/start that is OK.
            On Error Resume Next
            m_rstRecordset.AbsolutePosition = vlngIndex
        Else
            lngRow = 1
            m_rstRecordset.MoveFirst
            Do While lngRow < vlngIndex And Not m_rstRecordset.EOF
                lngRow = lngRow + 1
                m_rstRecordset.MoveNext
            Loop
        End If
    Else
        'No data.  just return me
    End If
    
    Set IDataCacheItem_Entry = SafeRef(Me)
    
End Function

Private Function IDataCacheItem_GetValue(ByVal vstrKeyName As String) As Variant
' description:
'    Retrieves the value of a field.
'
' pass:
'   vstrKeyName        The name of the field to process
'
' return: The field value. Variant is slightly faster and preserves type
'       NB: Nulls will be handled by the caller

    Const strMethodName  As String = "GetValue"
    On Error GoTo GetValueVBErr
       
#If GENERIC_SQL Then
    Dim objSQLAssist As New SQLAssist
#End If

    With m_rstRecordset.Fields(vstrKeyName)
        'It is faster to check for null than to raise an error
        If IsNull(.Value) Then
            IDataCacheItem_GetValue = vbNullString
        Else
            '204 is a Byte Array
            If Not .Type = 204 Then
                IDataCacheItem_GetValue = .Value
            Else
#If GENERIC_SQL Then
                IDataCacheItem_GetValue = objSQLAssist.GuidToString(.Value)
#Else
                IDataCacheItem_GetValue = GuidToString(.Value)
#End If
            End If
        End If
    End With
    
#If GENERIC_SQL Then
    Set objSQLAssist = Nothing
#End If
    
    Exit Function

GetValueVBErr:
    '3021 - BOF or EOF is true
    If Err.Number = 3021 Then
        IDataCacheItem_GetValue = vbNullString
    Else
        m_objErrAssist.AddToErrSource strMethodName
        Err.Description = Err.Description & vbCrLf & "Recordset=" & m_strName & "  Field=" & vstrKeyName
        If m_objErrAssist.IsSystemError = True Then
            m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
        End If
    
        '   re-raise error for business object to interpret as appropriate
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

#If GENERIC_SQL Then
    Set objSQLAssist = Nothing
#End If

End Function

Private Sub IDataCacheItem_GotoNextEntry()
    m_rstRecordset.MoveNext
End Sub

Private Sub IDataCacheItem_GotoFirstEntry()
    m_rstRecordset.MoveFirst
End Sub

Private Property Get IDataCacheItem_HasData() As Boolean
    IDataCacheItem_HasData = Not m_rstRecordset.BOF
End Property

Private Property Get IDataCacheItem_Name() As String
    IDataCacheItem_Name = m_strName
End Property

Private Property Get IDataCacheItem_XML() As String
' header ----------------------------------------------------------------------------------
' description:
'    This method returns an XML representation of a m_rstRecordset. It is in
'    here for debugging purposes - to be used from the Immediate window.
'
' pass:
'
' return:
'------------------------------------------------------------------------------------------
On Error GoTo XMLVBErr
Dim strOut As String
Dim fldCurrent As Field
    
#If GENERIC_SQL Then
    Dim objSQLAssist As New SQLAssist
#End If
    
    If m_strSQL = vbNullString Then
        strOut = "<" & IDataCacheItem_Name & " TYPE=""Recordset"">" & vbCrLf
    Else
        strOut = "<" & IDataCacheItem_Name & " TYPE=""Recordset"" SQL=""" & m_strSQL & """>" & vbCrLf
    End If
    
    If Not m_rstRecordset.BOF Then
        m_rstRecordset.MoveFirst
        
        Do While Not m_rstRecordset.EOF
            strOut = strOut & "  <ITEM>" & vbCrLf
            
            For Each fldCurrent In m_rstRecordset.Fields
                If Not IsNull(fldCurrent.Value) Then
                    If fldCurrent.Type <> 204 Then
                        strOut = strOut & "    <" & fldCurrent.Name & ">" & fldCurrent.Value & "</" & fldCurrent.Name & ">" & vbCrLf
                    Else
#If GENERIC_SQL Then
                        strOut = strOut & "    <" & fldCurrent.Name & ">" & objSQLAssist.GuidToString(fldCurrent.Value) & "</" & fldCurrent.Name & ">" & vbCrLf
#Else
                        strOut = strOut & "    <" & fldCurrent.Name & ">" & GuidToString(fldCurrent.Value) & "</" & fldCurrent.Name & ">" & vbCrLf
#End If
                    End If
                Else
                    strOut = strOut & "    <" & fldCurrent.Name & "/>" & vbCrLf
                End If
            Next fldCurrent
            
            m_rstRecordset.MoveNext
        
            strOut = strOut & "  </ITEM>" & vbCrLf
        Loop
        
        m_rstRecordset.MoveFirst
    Else
        For Each fldCurrent In m_rstRecordset.Fields
            strOut = strOut & "    <" & fldCurrent.Name & "/>" & vbCrLf
        Next fldCurrent
            
    End If

    IDataCacheItem_XML = strOut & "</" & IDataCacheItem_Name & ">" & vbCrLf
    
#If GENERIC_SQL Then
    Set objSQLAssist = Nothing
#End If
    
    Exit Property
    
XMLVBErr:
    IDataCacheItem_XML = strOut & "</" & IDataCacheItem_Name & "> Error"
    
#If GENERIC_SQL Then
    Set objSQLAssist = Nothing
#End If

End Property



Public Sub Create(ByVal vstrName As String, _
            ByVal vrstRecordset As Recordset, Optional ByVal vstrSQL As String)
' header ----------------------------------------------------------------------------------
' description:
'    A create method is used to keep the properties read-only
'
' pass:
'   vstrName                  The name of this recordset. Used as a key
'   vrstRecordset             The Data
'   vstrSQL                   The SQL used to generate this query - stored for diags
'
'------------------------------------------------------------------------------------------' description:
    Const strMethodName = "Create"

    m_strName = vstrName
    Set m_rstRecordset = vrstRecordset
    m_lngRecordCount = -1
    m_strSQL = vstrSQL
    Exit Sub
End Sub

#If Not GENERIC_SQL Then
Private Function GuidToString(ByRef rbytArray() As Byte) As String
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
'------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strGuid As String
    
    strGuid = vbNullString
    
    For i = 0 To 15
        strGuid = strGuid & Right$("00" & Hex$(rbytArray(i)), 2)
    Next i
    
    GuidToString = strGuid

End Function
#End If
