VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "DataCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Workfile:
' Copyright:    Copyright © 1999 Marlborough Stirling
'
'Description:
'   A collection of named ADO recordsets. Used as a data cache
'   for the Data mapping processes


' Dependencies:

' Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'MH    23/02/00 Created
'CL    07/05/02 SYS4510  Modifications to class initialise
'CL    10/05/02 SYS4510  Remove class initialize & class terminate
'------------------------------------------------------------------------------------------
Implements ObjectControl

Option Explicit

'=============================================
'Constant Declaration Section
'=============================================

'=============================================
'Variable Declaration Section
'=============================================
Private m_colDataCacheItems As Collection
Private m_objDefault As IDataCacheItem
Private m_objContext As ObjectContext
Private m_objErrAssist As New ErrAssist

Public Property Get Default() As IDataCacheItem
    Set Default = m_objDefault
End Property

Public Sub SetDefault(ByVal vIndex As Variant)
' description:
'    Sets the default recordset based on what is in the collection already
' return:
'
    Const strMethodName = "SetDefaultRecordset"
    On Error GoTo SetDefaultRecordsetVbErr
    
    Set m_objDefault = Item(vIndex)
    Exit Sub

SetDefaultRecordsetVbErr:
          
    m_objErrAssist.AddToErrSource (strMethodName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Property Get Count() As Long
' description:
'    The standard count method
' return:
'            The number of elements in the collection
    
    Const strMethodName = "Count"
    On Error GoTo CountVbErr
    
    Count = m_colDataCacheItems.Count
    Exit Property

CountVbErr:
          
    m_objErrAssist.AddToErrSource (strMethodName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Property

'Only in here for help in debugging
Public Property Get XML() As String
    Dim strOut As String
    Dim objIDataCache As IDataCacheItem
    
    For Each objIDataCache In m_colDataCacheItems
        strOut = strOut & objIDataCache.XML
    Next objIDataCache
    XML = strOut
    
End Property

Public Function Item(ByVal vIndex As Variant) As IDataCacheItem
' description:
'    Standard Item method
    
    Const strMethodName = "Item"
    On Error GoTo ItemVbErr
    
    'If a bracket is provided then reference a particular row
    If Right$(vIndex, 1) <> ")" Then
        Set Item = m_colDataCacheItems.Item(vIndex)
    Else
        Dim intOpen As Integer
        Dim strIndex As String
        Dim objIItem As IDataCacheItem
        strIndex = vIndex
        intOpen = InStr(strIndex, "(")
        Set objIItem = m_colDataCacheItems.Item(Left$(strIndex, intOpen - 1))
        Set Item = objIItem.Entry(Mid$(strIndex, intOpen + 1, Len(strIndex) - intOpen - 1))
    End If
    Exit Function

ItemVbErr:
    Err.Raise omiga4MissingTableName, TypeName(Me), "Table """ & vIndex & """ not found in cache"

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_colDataCacheItems.[_NewEnum]
End Function

Public Function AddRecordset(ByVal vstrRecordsetName As String, _
            ByVal vrstData As Recordset, _
            Optional ByVal vstrOriginalSQL As String) As IDataCacheItem
' description:
'    Adds a recordset into the collection and makes it the default
' pass:
'    vstrRecordsetName    The name of the recordset to add
'    vrstData             The data
'
    On Error GoTo AddVbErr
    
    Dim objDataCacheItem As DataCacheRecordset
    
    If m_objContext Is Nothing Then
        Set objDataCacheItem = New DataCacheRecordset
    Else
        Set objDataCacheItem = m_objContext.CreateInstance(App.Title & ".DataCacheRecordset")
    End If
    
    objDataCacheItem.Create vstrRecordsetName, vrstData, vstrOriginalSQL
    m_colDataCacheItems.Add objDataCacheItem, vstrRecordsetName
    
    Set m_objDefault = objDataCacheItem
    Set AddRecordset = objDataCacheItem
    
    #If WRITETABLESTOFILE Then
        Dim objIDataCacheItem As IDataCacheItem
        Set objIDataCacheItem = objDataCacheItem
        Open "C:\TABLES.xml" For Append As 1
        Print #1, Now
        Print #1, objIDataCacheItem.XML
        Close 1
    #End If
    Exit Function

AddVbErr:
             
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Public Function AddorReplaceRecordset(ByVal vstrRecordsetName As String, _
                            ByVal vrstData As Recordset, _
                            Optional ByVal vblnChangeDefault As Boolean = True) As IDataCacheItem
' description:
'    Adds a recordset into the collection, or replaces it if it is already there
' pass:
'    vstrRecordsetName    The name of the recordset to add
'    vrstData             The data
'    vblnChangeDefault    The Default recordset is set to be the latest added. This is correct for all but cached data.
 Const strMethodName = "AddorReplace"
    On Error GoTo AddorReplaceVbErr
 
    Dim objDataCacheItem As IDataCacheItem
    
    Set objDataCacheItem = m_objDefault
    
    Set AddorReplaceRecordset = AddRecordset(vstrRecordsetName, vrstData)
    
    If Not vblnChangeDefault Then Set m_objDefault = objDataCacheItem
    Exit Function

AddorReplaceVbErr:
    Remove vstrRecordsetName
    Set AddorReplaceRecordset = AddRecordset(vstrRecordsetName, vrstData)
    If Not vblnChangeDefault Then Set m_objDefault = objDataCacheItem
End Function

Public Function AddXML(ByVal vstrName As String, _
            ByVal vstrXML As String) As IDataCacheItem
' description:
'    Adds a recordset into the collection and makes it the default
' pass:
'    vstrName    The name of the recordset to add
'    vrstData             The data
'
    On Error GoTo AddVbErr
    
    Dim objDataCacheItem As DataCacheXML
    
    If m_objContext Is Nothing Then
        Set objDataCacheItem = New DataCacheXML
    Else
        Set objDataCacheItem = m_objContext.CreateInstance(App.Title & ".DataCacheXML")
    End If
    
    objDataCacheItem.Create vstrName, vstrXML
    m_colDataCacheItems.Add objDataCacheItem, vstrName
    
    Set m_objDefault = objDataCacheItem
    Set AddXML = objDataCacheItem
    
    #If WRITETABLESTOFILE Then
        Dim objIDataCacheItem As IDataCacheItem
        Set objIDataCacheItem = objDataCacheItem
        Open "C:\TABLES.xml" For Append As 1
        Print #1, Now
        Print #1, objIDataCacheItem.XML
        Close 1
    #End If
    Exit Function

AddVbErr:
             
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Sub AddorReplaceConstant(ByVal vstrName As String, ByVal vstrValue As String)
' description:
'    Adds a constant into the collection. if a constants collection already exists then
'    it is replicated. It is expected that constants are defined at the top of a map
'    replicating on change should be faster than replicating on recursion into a table node
' pass:
'    vstrName    The name of the recordset to add
'    vrstData             The data
'
    On Error GoTo AddVbErr
    
    Dim objIDataCacheItem As IDataCacheItem
    Dim objDataCacheItem As DataCacheConstants
    
    On Error Resume Next
    Set objIDataCacheItem = Item("#CONSTANT")
    
    If Err.Number = 0 Then
        Set objDataCacheItem = objIDataCacheItem.Clone(objIDataCacheItem.Name)
        Remove objIDataCacheItem.Name
        Set objIDataCacheItem = objDataCacheItem
        m_colDataCacheItems.Add objIDataCacheItem, objIDataCacheItem.Name
    Else
        If m_objContext Is Nothing Then
            Set objDataCacheItem = New DataCacheConstants
         Else
            Set objDataCacheItem = m_objContext.CreateInstance(App.Title & ".DataCacheConstants")
        End If
        objDataCacheItem.Create
        Set objIDataCacheItem = objDataCacheItem
        m_colDataCacheItems.Add objIDataCacheItem, objIDataCacheItem.Name
    End If
    
    objDataCacheItem.Setvalue vstrName, vstrValue
    Exit Sub

AddVbErr:
             
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function AddAlias(ByVal vstrName As String, _
            ByVal vstrAlias As String) As IDataCacheItem
' description:
'    Adds an Alias of an existing item into the collection and makes it the default
' pass:
'    vstrName    The current name
'    vrstAlias   The new name
'
    On Error GoTo AddVbErr
    
    Dim objIDataCacheItem As IDataCacheItem
    
    Set objIDataCacheItem = Item(vstrName).Clone(vstrAlias)
    
    m_colDataCacheItems.Add objIDataCacheItem, vstrAlias
    Set m_objDefault = objIDataCacheItem
    Set AddAlias = objIDataCacheItem
    
    #If WRITETABLESTOFILE Then
        Open "C:\TABLES.XML" For Append As 1
        Print #1, Now
        Print #1, objIDataCacheItem.XML
        Close 1
    #End If
    Exit Function

AddVbErr:
             
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function




Public Sub Remove(ByVal vIndex As Variant)
' description:
'    A standard remove method
' pass:
'     vindex        The item to be removed
'
    Const strMethodName = "Remove"
    On Error GoTo RemoveVbErr
    
    m_colDataCacheItems.Remove vIndex
    Exit Sub

RemoveVbErr:
          
    m_objErrAssist.AddToErrSource (strMethodName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub Clone(ByVal vcolSource As DataCache)
' description:
'    Copies the recordsets from the source collection and puts them in
'    this collection.

' pass:
'   vcolSource          The collection to copy
'
    Const strMethodName = "Clone"
    On Error GoTo CloneVbErr

    Dim objIDataCacheItem As IDataCacheItem
    
    Set m_colDataCacheItems = New Collection
    
    ' copy the objects over
    For Each objIDataCacheItem In vcolSource
        m_colDataCacheItems.Add objIDataCacheItem, objIDataCacheItem.Name
    Next objIDataCacheItem
    
    If Not vcolSource.Default Is Nothing Then
        SetDefault vcolSource.Default.Name
    Else
        Set m_objDefault = Nothing
    End If
    
    Exit Sub
    
CloneVbErr:
          
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Function GetValue(ByVal vstrFieldName As String) As Variant
' description:
'    Retrieves the value of a field. Field names can be in the form
'    'Field' and 'Table.Field'. If Table is omitted the default table is
'    used. Recordsets are assumed to be correctly positioned.
' pass:
'   vstrFieldName        The name of the field to process
'   vstrFilter           The name of the field to process'
'
' return: The field value or a null string. Variant is slightly faster and preserves type
'
    On Error GoTo GetValueVBErr
    Const strMethodName  As String = "GetValue"
    Dim lngDot As Long

    lngDot = InStr(vstrFieldName, ".")
    
    If lngDot = 0 Then
        'No dot, assume it is just a field name and use the default recordset
        GetValue = m_objDefault.GetValue(vstrFieldName)
    Else
        'Get the named table then use the thing after the dot as the field name
         GetValue = Item(Left$(vstrFieldName, lngDot - 1)).GetValue(Mid$(vstrFieldName, lngDot + 1))
    End If
        
    Exit Function

GetValueVBErr:
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    Resume      ' never called, but in here for debugging
End Function

Public Function GetValueUsingFilter(ByVal vstrFieldName As String, ByVal vstrFilter As String) As Variant
' description:
'    Retrieves the value of a field. Field names can be in the form
'    'Field' and 'Table.Field'. If Table is omitted the default table is
'    used. IDataCacheItems are assumed to be correctly positioned.
'
' pass:
'   vstrFieldName        The name of the field to process
'   vstrFilter           The name of the field to process'
'
' return: The field value or a null string. Variant is slightly faster and preserves type
    
    Const strMethodName  As String = "GetValueUsingFilter"
    Dim lngDot As Long
    Dim objIDataCacheItem As IDataCacheItem
    
    lngDot = InStr(vstrFieldName, ".")
    
    If lngDot > 0 Then
        'Get the named table then use the thing after the dot as the field name
        Set objIDataCacheItem = Item(Left$(vstrFieldName, lngDot - 1))
    Else
        'No dot, assume it is just a field name and use the default IDataCacheItem
        Set objIDataCacheItem = m_objDefault
    End If
        
    If objIDataCacheItem.HasData Then
        
        'Try and look at the current row
        If Not FilterIsTrue(ReplaceFieldsWithData(vstrFilter)) Then
            'Now run through the IDataCacheItem.
            'If the filter does not match then it will fail and raise a vb error.
            objIDataCacheItem.GotoFirstEntry
            Do While Not FilterIsTrue(ReplaceFieldsWithData(vstrFilter)) And Not objIDataCacheItem.AtEndOfData
                objIDataCacheItem.GotoNextEntry
            Loop
        End If
        
        GetValueUsingFilter = GetValue(vstrFieldName)
        
    Else
        GetValueUsingFilter = vbNullString
    End If
        
    Exit Function

GetValueUsingFilterVbErr:
          
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function ReplaceFieldsWithData(ByVal vstrExpr As String) As String
' description:
'    Takes a string expression that may represent SQL or a VB regular
'    expression and replaces all of the field references. Field references
'    must be in the form [Table.Field] or [Field]. The latter will use the default table
'
' pass:
'   vstrExpr            The expression to be parsed
'
    Const strMethodName = "ReplaceFieldsWithData"
    Const strTempStringDelim As String = "¬"
    Const strMyStringDelim As String = "'"
    Const strSQLStringDelim As String = "'"
    
    Dim strExpr As String
    Dim strOut As String
    Dim strFieldName As String
    
    Dim lngFieldStart As Long
    Dim lngFieldEnd As Long
    
    strOut = vbNullString
    
    'SQLStrings are assumed. They will be converted to SQL strigns if required
    strExpr = Replace(vstrExpr, strMyStringDelim, strTempStringDelim)
    
    'All fields are delimited by [] - eg [Table.Field], [Field].
    
    lngFieldStart = InStr(strExpr, "[")
    
    Do While lngFieldStart > 0
        lngFieldEnd = InStr(lngFieldStart + 1, strExpr, "]")
        
        'Grab the bit before the field
        strOut = strOut & Left$(strExpr, lngFieldStart - 1)
        
        'Grab the field name
        strFieldName = Mid$(strExpr, lngFieldStart + 1, lngFieldEnd - lngFieldStart - 1)
        
        strOut = strOut & GetValue(strFieldName)
        
        'remove the processed field from the expression and continue parsing
        strExpr = Mid$(strExpr, lngFieldEnd + 1)
        lngFieldStart = InStr(strExpr, "[")
    Loop
    
    'Grab whatever is left of the expression
    strOut = strOut & strExpr
    
    'Now double up the string delimiters so that it will withstand parsing
    strOut = Replace(strOut, strSQLStringDelim, strSQLStringDelim & strSQLStringDelim)
    
    'Finally, substitute the old string delimiters in
    ReplaceFieldsWithData = Replace(strOut, strTempStringDelim, strSQLStringDelim)
    
    Exit Function

ReplaceFieldsWithDataVbErr:
          
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function FilterIsTrue(ByVal vstrFilter As String) As Boolean
' description:
'    Evaluates a text expression which is a boolean expression
' pass:
'   vstrFilter             The expression to test
'
' return:
'            True if it matches
    Const strMethodName = "FilterIsTrue"
    Dim lngEquals As Long
    
    lngEquals = InStr(vstrFilter, "=")
    
    If lngEquals > 0 Then
        FilterIsTrue = (Left$(vstrFilter, lngEquals - 1) = Mid$(vstrFilter, lngEquals + 1))
    Else
        m_objErrAssist.RaiseError TypeName(Me), strMethodName, omiga4UnspecifiedError, "Filter expression """ & vstrFilter & """ is too complicated"
    End If
    
    Exit Function

FilterIsTrueVbErr:
          
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    '   re-raise error for business object to lngerpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_colDataCacheItems = New Collection
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_colDataCacheItems = Nothing
    Set m_objDefault = Nothing
    Set m_objContext = Nothing
End Sub
