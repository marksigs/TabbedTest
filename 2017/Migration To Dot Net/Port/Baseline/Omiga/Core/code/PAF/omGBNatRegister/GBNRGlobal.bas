Attribute VB_Name = "GBNRGlobals"
Option Explicit

'------------------------------------------------------------------------
' Class Module  omGBNR
' File:         omGBNR.cls
' Author:       INR
' Date:         13-11-2002
' Purpose:      Holds Globals for omGBNatRegWrapper

'
' History:
'   AQR                 AUTHOR      DESCRIPTION
'   ---                 ------      -----------
'   BMIDS00960/CC021    INR        13/11/2002  Created.
'------------------------------------------------------------------------
Declare Function GBIMAPI Lib "GBNRTI32.DLL" Alias "_GBIMAPI@16" (ByVal Query As String, ByVal Buffr As String, ByVal Reply As String, ByVal Handl As String) As Integer
Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Global Const User_buf_size = 15000             ' The size of the user buffer
Global Const Reply_buf_size = 10             ' The size of the Reply buffer
Global Const Handl_buf_size = 10             ' The size of the Reply buffer

Global szBuffr As String * User_buf_size
Global szReply As String * Reply_buf_size
Global szHandl As String * Reply_buf_size

Public Sub Main()
    adoBuildDbConnectionString
End Sub

