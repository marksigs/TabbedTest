Attribute VB_Name = "omHIGlobals"
Option Explicit
'------------------------------------------------------------------------------------------
'BMids History:
'
' Prog   Date           Description
' MDC    18/06/2002     BMIDS00142 IWP1 BM059 - Code review change
' MO     08/11/2002     BMIDS00752 - Hunter fails when there is no quotation
'-------------------------------------------------------------------------------------------------
'BMIDS00142 MDC 18/06/2002
Public Enum HUNTERERROR
    oeMissingApplicantsForApplication = 127
    ' MO - 08/11/2002 - BMIDS00752
    oeNoAcceptedOrActiveQuotation = 7009
End Enum
'BMIDS00142 MDC 18/06/2002 - End
Public Sub Main()
    adoBuildDbConnectionString
End Sub
