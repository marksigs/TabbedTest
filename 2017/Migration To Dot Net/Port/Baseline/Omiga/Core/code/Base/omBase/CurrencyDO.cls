VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "CurrencyDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      CurrencyDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Collects all Currency data for use with the Currency Calculator
'
'Dependencies:
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PF     9/4/01      Created
'------------------------------------------------------------------------------------------
Option Explicit
Implements ICurrencyDO
Implements ObjectControl
 
'=============================================
'Constant Declaration Section
'=============================================
'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist             ' Remove if not required
Private m_objIDOAssist As IDOAssist
Private Sub ICurrencyDO_AddDerivedData(ByVal vxmlData As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   XML elements must be created for any derived values as specified.
'   Add any derived values to XML. E.g. data type 'double' fields will
'   need to be formatted as strings to required precision & rounding.
' pass:
'   vxmlData          base XML node
'                       as:
'                           <tablename>
'                               <element1>element1 value</element1>
'                               <elementn>elementn value</elementn>
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo AddDerivedDataVBErr
    Const strFunctionName As String = "ICurrencyDO_AddDerivedData"
    Exit Sub
AddDerivedDataVBErr:
    
    
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function ICurrencyDO_FindList() As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr
    Const strFunctionName As String = "ICurrencyDO_FindList"
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomBaseClassDef
    Dim xmlElement As IXMLDOMElement
    Set objIClassDef = New omBaseClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCurrencyData()
    Set xmlElement = xmlClassDefDoc.createElement("CURRENCY")
    Set ICurrencyDO_FindList = m_objIDOAssist.FindList(xmlElement, xmlClassDefDoc)
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlElement = Nothing
    m_objContext.SetComplete
    Exit Function
FindListVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlElement = Nothing
    m_objErrAssist.AddToErrSource (strFunctionName)
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    Set m_objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
    Set m_objIDOAssist = Nothing
End Sub
