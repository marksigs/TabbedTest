Attribute VB_Name = "KeyPress"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description   - Constants used throughout Supervisor
'
' Prog      Date        Description
' AA        06/03/01    Checks the for the alt key being pressed
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Declare Function GetAsyncKeyState Lib "user32" (ByVal dwMessage As Long) As Integer

