Attribute VB_Name = "ODILogOffUser"
'Workfile:      ODILogOffUser.bas
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:
'
'Dependencies:
'Issues:        Instancing:
'               MTSTransactionMode:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     24/08/01    Created.
'------------------------------------------------------------------------------------------
Option Explicit

Public Sub LogOffUser( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
    On Error GoTo LogOffUserExit
    
    Const strFunctionName = "LogOffUser"
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If
    
    Dim nodeConverterResponse As IXMLDOMNode
    
    Set nodeConverterResponse = Session_endSession(vobjODITransformerState)
    
    CheckConverterResponse nodeConverterResponse, True
    
    AddExceptionsToResponse nodeConverterResponse, vxmlResponseNode
        
LogOffUserExit:

    Set nodeConverterResponse = Nothing
    
    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Sub




