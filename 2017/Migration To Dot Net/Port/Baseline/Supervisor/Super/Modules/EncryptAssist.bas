Attribute VB_Name = "EncryptAssist"
'Workfile:      EncryptAssist.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Encryption and decryption routines. Based on code from Encrypt.dll.
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'RF    27/10/99 Created
'------------------------------------------------------------------------------------------
Option Explicit

Private Const cint_MIN = 10
Private Const cint_MAX = 40


Public Function Encrypt(ByVal vstrIn) As String

    Dim lngInputLen, lngChar As Long
    Dim intIncrement As Integer
    Dim intChar1, intChar2 As Integer
    Dim strOut As String
    
    lngInputLen = Len(vstrIn)
    
    ' initialise the random-number generator
    Rnd -1
    Randomize 1
    
    For lngChar = 1 To lngInputLen
    
        intIncrement = Int(Rnd() * cint_MAX) + cint_MIN
        intChar1 = Asc(Mid(vstrIn, lngChar, 1))
        
        intChar2 = intChar1 + intIncrement
        If intChar2 >= Asc("~") Then
            ' new character exceeds ascii 127 ('~'), thus deduct increment
            ' precede new character with "~" in output string so that Decrypt
            ' knows that the increment should be added and not deducted
            strOut = strOut & "~"
            intChar2 = intChar1 - intIncrement
        End If
        strOut = strOut & Chr(intChar2)
        
    Next

    Encrypt = strOut
        
End Function



