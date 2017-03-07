VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "DataCacheConstants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Workfile:      DataCacheConstants.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:  A class for holding pre-defined values.
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MH     23/02/00    Created
'------------------------------------------------------------------------------------------

Implements IDataCacheItem
'=============================================
'Constant Declaration Section
'=============================================
Private Const m_strName = "#CONSTANT"
'=============================================
'Variable Declaration Section
'=============================================
Private m_colConstants As Collection
Private m_objErrAssist As New ErrAssist
Private m_blnEOD As Boolean

Private Property Get IDataCacheItem_AtEndOfData() As Boolean
    IDataCacheItem_AtEndOfData = m_blnEOD
End Property

Private Function IDataCacheItem_Clone(ByVal vstrNewName As String) As IDataCacheItem
    Const strMethodName = "IDataCacheItem_Clone"
    On Error GoTo IDataCacheItem_CloneVBErr

    Dim objDataCacheConstants As DataCacheConstants
    Dim objContext As ObjectContext

    Set objContext = GetObjectContext()

    If objContext Is Nothing Then
        Set objDataCacheConstants = New DataCacheConstants
    Else
        Set objDataCacheConstants = objContext.CreateInstance(App.Title & ".DataCacheConstants")
    End If

    objDataCacheConstants.Copy m_colConstants
    Set IDataCacheItem_Clone = objDataCacheConstants
    
    Exit Function

IDataCacheItem_CloneVBErr:

    m_objErrAssist.AddToErrSource strMethodName

    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Property Get IDataCacheItem_Count() As Long
    IDataCacheItem_Count = 1
End Property

Private Function IDataCacheItem_Entry(ByVal vlngIndex As Long) As IDataCacheItem
    'Not yet supported for Constants
    Set IDataCacheItem_Entry = Nothing
End Function

Private Function IDataCacheItem_GetValue(ByVal vstrKeyName As String) As Variant
' description:
'    Retrieves the value of a Constant .
'
' pass:
'   vstrKeyName        The name of the Constants element to process
'
    Dim objConstant As DataCacheConstant
    
    On Error GoTo GetValueVBErr
    
    Set objConstant = m_colConstants(vstrKeyName)
    IDataCacheItem_GetValue = objConstant.Value
    Exit Function

GetValueVBErr:
    Err.Raise oeMissingElement, TypeName(Me), "Constant """ & vstrKeyName & """ has not been defined"

End Function

Private Sub IDataCacheItem_GotoNextEntry()
    m_blnEOD = True
End Sub

Private Sub IDataCacheItem_GotoFirstEntry()
    m_blnEOD = False
End Sub

Private Property Get IDataCacheItem_HasData() As Boolean
    IDataCacheItem_HasData = (m_colConstants.Count > 0)
End Property

Private Property Get IDataCacheItem_Name() As String
    IDataCacheItem_Name = m_strName
End Property

Private Property Get IDataCacheItem_XML() As String
' header ----------------------------------------------------------------------------------
' description:
'    This method returns an Constants representation of a m_rstConstants. It is in
'    here for debugging purposes - to be used from the Immediate window.
'
' pass:
'
' return:
'------------------------------------------------------------------------------------------
On Error GoTo ConstantsVBErr
    
    Dim strOut As String
    Dim objDataCacheConstant
    strOut = "<#CONSTANT TYPE=""#CONSTANT"">" & vbCrLf
    
    For Each objDataCacheConstant In m_colConstants
    strOut = strOut & "<CONSTANT>" & vbCrLf & _
            "<NAME>" & objDataCacheConstant.Name & "</NAME>" & vbCrLf & _
            "<VALUE>" & objDataCacheConstant.Value & "</VALUE>" & vbCrLf & _
            "</CONSTANT>" & vbCrLf
    Next objDataCacheConstant
    
    IDataCacheItem_XML = strOut & "</#CONSTANT>" & vbCrLf
    Exit Function
ConstantsVBErr:
    Err.Raise Err.Number, TypeName(Me), Err.Description
    

End Property


Public Sub Create()
' header ----------------------------------------------------------------------------------
' description:
'    A create method is used to keep the properties read-only
'
' pass:
'   vstrName                  The name of this Constants. Used as a key
'   vrstConstants                   The Data
'   vstrSQL                   The SQL used to generate this query - stored for diags
'
'------------------------------------------------------------------------------------------' description:
    Const strMethodName = "Create"
    On Error GoTo CreateVBErr

    Set m_colConstants = New Collection
    m_blnEOD = False
    Exit Sub

CreateVBErr:
          
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub Copy(ByVal vcolConstants As Collection)
' header ----------------------------------------------------------------------------------
' description:
'    A create method is used to keep the properties read-only
'
' pass:
'   vstrName                  The name of this Constants. Used as a key
'   vrstConstants                   The Data
'   vstrSQL                   The SQL used to generate this query - stored for diags
'
'------------------------------------------------------------------------------------------' description:
    Const strMethodName = "Create"
    On Error GoTo CreateVBErr
    Dim objConstant As DataCacheConstant
    
    Set m_colConstants = New Collection
    m_blnEOD = False
    
    For Each objConstant In vcolConstants
        m_colConstants.Add objConstant, objConstant.Name
    Next objConstant
    Exit Sub

CreateVBErr:
          
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Public Sub Setvalue(ByVal vstrName As String, ByVal vstrValue As String)
    On Error Resume Next
    Dim objConstant As DataCacheConstant
    
    Set objConstant = m_colConstants(vstrName)
    
    If objConstant Is Nothing Then
        Set objConstant = New DataCacheConstant
        objConstant.Name = vstrName
        m_colConstants.Add objConstant, vstrName
    End If
    
    objConstant.Value = vstrValue
End Sub
