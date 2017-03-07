VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMortgage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************************
'** Module:         IMortgage
'** Created by:     Andy Maggs
'** Date:           24/03/2004
'** Description:    Defines the interface that classes responsible for formatting
'**                 data into the format required for the various KFI document
'**                 templates must support
'********************************************************************************
Option Explicit
Option Compare Text

Public Sub FormatData(ByVal vobjCommon As CommonDataHelper)
End Sub

