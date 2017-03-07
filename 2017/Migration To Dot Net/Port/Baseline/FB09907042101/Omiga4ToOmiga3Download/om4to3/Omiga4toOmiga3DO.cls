VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Omiga4toOmiga3DO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      C:\Projects\dev\Omiga4ToOmiga3Download\om4to3\Omiga4toOmiga3DO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Code template for omiga4 Data Object
'
'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MH     23/02/00    Created
'RF     02/05/00    Amended error handling in IOmiga4toOmiga3DO_GetData.
'LD     04/09/00    Ammendments for NT4/W2000 dual-compile
'------------------------------------------------------------------------------------------
Option Explicit

Implements IOmiga4toOmiga3DO
Implements ObjectControl
 
'=============================================
'Constant Declaration Section
'=============================================

'=============================================
'Variable Declaration Section
'=============================================

Private m_objContext As ObjectContext
Private m_objErrAssist As New ErrAssist

Private Function IOmiga4toOmiga3DO_GetData(ByVal vxmlTableElement As IXMLDOMElement, _
        Optional ByVal vxmlPerformance As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'
'Notes:
'  vxmlPerformance is used to return to the caller performance metrics. It
' is only going to be used when compiled with the TIMINGS flag.
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "IOmiga4toOmiga3DO_GetData"

    Dim conOmiga4 As Connection
    Dim objMapDef As IMapDefOm4To3
    Dim objIADOAssist As IADOAssist
    Dim objIDBXMLWizard As IDBXMLWizard
    Dim xmlTYPEElement As IXMLDOMElement
    Dim xmlMap As FreeThreadedDOMDocument40
    
    #If TIMINGS Then
        Dim t As Single
        t = Timer
    #End If
    
    If m_objContext Is Nothing Then
        Set objMapDef = New MapDefOm4To3
    Else
        Set objMapDef = m_objContext.CreateInstance(App.Title & ".MapDefOm4To3")
    End If
    
    Set xmlTYPEElement = vxmlTableElement.getElementsByTagName("TYPE").Item(0)
    If xmlTYPEElement Is Nothing Then
        'Err.Raise oeMissingElement, TypeName(Me), "An element with the name 'TYPE' was expected but not found"
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingElement, _
            "An element with the name 'TYPE' was expected but not found"
    End If
       
    Set xmlMap = objMapDef.LoadMap(xmlTYPEElement.Text)
    Set objMapDef = Nothing
    
    #If TIMINGS Then
        vxmlPerformance.setAttribute "MAPLOAD", Format$(Timer - t, TIMEFORMAT)
        t = Timer
    #End If
    
    #If WRITETABLESTOFILE Then
        On Error Resume Next
        Kill App.Path & "\TABLES.xml"
        On Error GoTo GetDataVbErr
    #End If
 
   
    'Open the connection - using the helpers and create the required classes
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDBXMLWizard = New DBXMLWizard
    Else
        Set objIDBXMLWizard = m_objContext.CreateInstance(App.Title & ".DBXMLWizard")
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        #If TIMINGS Then
            vxmlPerformance.setAttribute "OTHERBOS", Format$(Timer - t, TIMEFORMAT)
            t = Timer
        #End If
    End If

    Set conOmiga4 = New ADODB.Connection
    Dim strCon As String
    
    strCon = objIADOAssist.GetConnStr
    
    #If TIMINGS Then
        vxmlPerformance.setAttribute "CONSTR", Format$(Timer - t, TIMEFORMAT)
        t = Timer
    #End If
    
    conOmiga4.open strCon
    conOmiga4.CursorLocation = adUseClient
    
    #If TIMINGS Then
        vxmlPerformance.setAttribute "CONNECTION", Format$(Timer - t, TIMEFORMAT)
        t = Timer
    #End If
    
    Set IOmiga4toOmiga3DO_GetData = objIDBXMLWizard.CreateXML(vxmlTableElement, xmlMap, conOmiga4)
    
    #If TIMINGS Then
        vxmlPerformance.setAttribute "OM4TOXML", Format$(Timer - t, TIMEFORMAT)
        t = Timer
    #End If
    
    conOmiga4.Close
    Set conOmiga4 = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function
    
GetDataVbErr:
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not conOmiga4 Is Nothing Then
         conOmiga4.Close
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
        
    'Err.Raise oeCommandFailed, TypeName(Me),  Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function

Private Sub ObjectControl_Deactivate()
    Set m_objContext = Nothing
End Sub


