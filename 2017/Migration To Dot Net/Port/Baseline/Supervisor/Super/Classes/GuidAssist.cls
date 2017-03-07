VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuidAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Workfile:      GuidAssist.cls
'`right:     Copyright © 1999 Marlborough Stirling

'Description:   Helper object for creating Global Unique Identifiers
'Dependencies:
'Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'PSC   06/07/99 Created
'------------------------------------------------------------------------------------------

Option Explicit
 
Private Type Guid
    D1       As Long
    D2       As Integer
    D3       As Integer
    D4(8)    As Byte
End Type

Private Declare Function WinCoCreateGuid Lib "OLE32.DLL" Alias "CoCreateGuid" (guidNewGuid As Guid) As Long

Public Function CreateGUID() As String
' -----------------------------------------------------------------------------------------
' description:  Creates a Global Unique Identifier
' return:
'------------------------------------------------------------------------------------------

    Dim guidNewGuid  As Guid
    Dim strBuffer    As String

    Call WinCoCreateGuid(guidNewGuid)

    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D1), 8)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D2), 4)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D3), 4)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(0)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(1)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(2)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(3)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(4)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(5)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(6)), 2)
    strBuffer = PadRight0(strBuffer, Hex$(guidNewGuid.D4(7)), 2)

    CreateGUID = strBuffer

End Function

Private Function PadRight0(ByVal vstrBuffer As String, _
                           ByVal vstrBit As String, _
                           ByVal intLenRequired As Integer, _
                           Optional bHyp As Boolean _
                         ) As String
' -----------------------------------------------------------------------------------------
' description:
' return:
'------------------------------------------------------------------------------------------

    PadRight0 = vstrBuffer & _
                vstrBit & _
                String$(Abs(intLenRequired - Len(vstrBit)), "0")

End Function


