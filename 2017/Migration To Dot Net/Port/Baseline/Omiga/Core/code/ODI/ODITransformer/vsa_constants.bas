Attribute VB_Name = "vsa_constants"
'Workfile:      VSA_Constants.bas
'Copyright:     Copyright © 2000 Marlborough Stirling
'
'Description:   Omiga support for Visual Studio Analyzer
'
'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: NotAnMTSObject
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    05/01/01    Created
'------------------------------------------------------------------------------------------
Option Explicit

#If USING_VSA Then
    ' Change this if an event source other than OMIGA4_EVENT_SOURCE is required.
    Public Const gstrVSAEventSource = OMIGA4_EVENT_SOURCE
    Public Const gstrVSASessionPrefix As String = "ODI"
#End If

