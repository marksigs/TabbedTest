VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "MapDefOm4To3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      C:\Projects\dev\Omiga4ToOmiga3Download\om4to3\MapDefom4to3.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Mapping templates for the Omiga4 to Omiga3 download
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MH     23/02/00    Created
'APS    22/09/00    Change to pick up XML map from "...\Omiga 4\XML" directory
'------------------------------------------------------------------------------------------

Option Explicit
Implements IMapDefOm4To3

Private m_objErrAssist As New ErrAssist
' header ----------------------------------------------------------------------------------
' description:
'    Creates an xml document which maps defines an output xml stream in
'    terms of the Omiga4 database and an input specification
'
' pass:  None

' return:                       An XML DOMdocumet map
'------------------------------------------------------------------------------------------
Private Function IMapDefOm4To3_LoadMap(ByVal vstrMapName As String) As FreeThreadedDOMDocument40
    Const strMethodName = "IMapDefOm4To3_LoadMap"
    On Error GoTo IMapDefOm4To3_LoadMapVBErr
    Const strFunctionName As String = "IMapDefOm4To3_LoadMap"
    Dim hFile As Integer
    Dim strFileName As String
    
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    
    'Use VB to look for the file. If it is not there it will create a nice
    ' File not Found message which can be appendend to with the file name.
    strFileName = App.Path & "\" & vstrMapName & ".XML"
    
    ' APS 22/09/00 Change to pick up XML map from "...\Omiga 4\XML" directory
    ' Only do the subsitution once to change DLL -> XML
    strFileName = Replace(strFileName, "DLL", "XML", 1, 1, vbTextCompare)
    
    hFile = FreeFile
    Open strFileName For Input Access Read As hFile
    Close hFile
    
    'If anything now breaks it will be XML related.
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    xmlDoc.load (strFileName)
    Set IMapDefOm4To3_LoadMap = xmlDoc
        
    If xmlDoc.parseError.errorCode <> 0 Then
            m_objErrAssist.ThrowError TypeName(Me), strMethodName, _
            oeXMLParserError, m_objErrAssist.FormatParserError(xmlDoc.parseError)
    End If

    Exit Function
 
IMapDefOm4To3_LoadMapVBErr:
    'Append the filename so that the caller has a clue.
    Err.Description = Err.Description & " File=""" & strFileName & """"
    
    m_objErrAssist.AddToErrSource strMethodName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strMethodName, Err.Number, Err.Description
    End If

    Err.Raise Err.Number, Err.Source, Err.Description

End Function


