VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataCacheXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      DataCacheXML.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   A simple wrapper for the XML so that its data can be handled in the
'  same way as a recordset from the Maps.
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
Private m_XML As FreeThreadedDOMDocument40
Private m_objErrAssist As New ErrAssist
Private m_blnEOD As Boolean

Private Property Get IDataCacheItem_AtEndOfData() As Boolean
    IDataCacheItem_AtEndOfData = m_blnEOD
End Property

Private Function IDataCacheItem_Clone(ByVal vstrNewName As String) As IDataCacheItem
    Const strMethodName = "IDataCacheItem_Clone"
    On Error GoTo IDataCacheItem_CloneVBErr

    Dim objDataCacheXML As DataCacheXML
    Dim objContext As ObjectContext
    
    Set objContext = GetObjectContext()
    
    If objContext Is Nothing Then
        Set objDataCacheXML = New DataCacheXML
    Else
        Set objDataCacheXML = objContext.CreateInstance(App.Title & ".DataCacheXML")
    End If
    
    objDataCacheXML.CreateUsingXMLDocument vstrNewName, m_XML
    Set IDataCacheItem_Clone = objDataCacheXML
    
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
    'Not yet supported for XML
    Set IDataCacheItem_Entry = Nothing
End Function

Private Function IDataCacheItem_GetValue(ByVal vstrKeyName As String) As Variant
' description:
'    Retrieves the value of an XML element.
'
' pass:
'   vstrKeyName        The name of the xml element to process
'
    Const strMethodName = "IDataCacheItem_GetValue"
    On Error GoTo IDataCacheItem_GetValueVBErr
    
    Dim xmlNodes As IXMLDOMNodeList
    
    Set xmlNodes = m_XML.getElementsByTagName(vstrKeyName)

    IDataCacheItem_GetValue = xmlNodes.Item(0).Text
    Exit Function

IDataCacheItem_GetValueVBErr:
    Err.Raise oeMissingElement, TypeName(Me), "Cannot find an Element like """ & vstrKeyName & """ in the input XML"
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub IDataCacheItem_GotoNextEntry()
    m_blnEOD = True
End Sub

Private Sub IDataCacheItem_GotoFirstEntry()
    m_blnEOD = False
End Sub

Private Property Get IDataCacheItem_HasData() As Boolean
    IDataCacheItem_HasData = m_XML
End Property

Private Property Get IDataCacheItem_Name() As String
    IDataCacheItem_Name = m_strName
End Property

Private Property Get IDataCacheItem_XML() As String
' header ----------------------------------------------------------------------------------
' description:
'    This method returns an XML representation of a m_rstXML. It is in
'    here for debugging purposes - to be used from the Immediate window.
'
' pass:
'
' return:
'------------------------------------------------------------------------------------------
On Error GoTo XMLVBErr
    
    IDataCacheItem_XML = "<" & IDataCacheItem_Name & " TYPE=""XML"">" & vbCrLf & _
            m_XML.XML & "</" & IDataCacheItem_Name & ">" & vbCrLf
    Exit Property
    
XMLVBErr:
    Err.Raise Err.Number, TypeName(Me), Err.Description
    

End Property



Public Sub Create(ByVal vstrName As String, _
            ByVal vstrXML As String)
' header ----------------------------------------------------------------------------------
' description:
'    A create method is used to keep the properties read-only
'
' pass:
'   vstrName                  The name of this XML. Used as a key
'   vrstXML                   The Data
'   vstrSQL                   The SQL used to generate this query - stored for diags
'
'------------------------------------------------------------------------------------------' description:
    Const strMethodName = "Create"
    On Error GoTo CreateVBErr

    m_strName = vstrName
    
    Set m_XML = New FreeThreadedDOMDocument40
    m_XML.validateOnParse = False
    m_XML.setProperty "NewParser", True
    m_XML.async = True
    m_XML.loadXML vstrXML
    m_blnEOD = False
    Exit Sub

CreateVBErr:
          
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Public Sub CreateUsingXMLDocument(ByVal vstrName As String, _
            ByVal vxmlDocument As FreeThreadedDOMDocument40)
' header ----------------------------------------------------------------------------------
' description:
'    A CreateUsingXMLDocument method is used to keep the properties read-only
'
' pass:
'   vstrName                  The name of this XML. Used as a key
'   vrstXML                   The Data
'   vstrSQL                   The SQL used to generate this query - stored for diags
'
'------------------------------------------------------------------------------------------' description:
    Const strMethodName = "CreateUsingXMLDocument"
    On Error GoTo CreateUsingXMLDocumentVBErr

    m_strName = vstrName
    Set m_XML = vxmlDocument
    Exit Sub

CreateUsingXMLDocumentVBErr:
          
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


