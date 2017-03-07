Attribute VB_Name = "guidAssistEx"
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
