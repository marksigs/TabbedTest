Attribute VB_Name = "omAdminGlobals"
Option Explicit

Public Enum ADMINERROR
    ' Rate Change specific errors
    oeAdminCustomerNumbersError = 4909
End Enum

Public Sub Main()
    adoBuildDbConnectionString
End Sub

