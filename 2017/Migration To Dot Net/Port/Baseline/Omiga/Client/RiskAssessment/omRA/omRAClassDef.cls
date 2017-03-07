VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "omRAClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      omRAClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   ClassDef's for Risk Assessment area
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MH     39/03/99    Created
'------------------------------------------------------------------------------------------

Option Explicit

Implements IClassDefomRA

'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
' Prog  Date        Description
' MV    10/06/2002  BMIDS00032 - BM078- Amended IClassDefomRA_LoadRiskAssessmentData with ScoreCardInd
' MV    11/06/2002  BMIDS00032 - BM078 - Removed ScoreCardInd from IClassDefomRA_LoadRiskAssessmentData
'                                        Amended IClassDefomRA_LoadRARuleData with ScoreCardInd
' DB    29/10/2002  BMIDS00712 - CC019 - Added new classdef of LoadPreviousRARuleOverrideData
' MV    29/01/2003  BM0302  Amended IClassDefomRA_LoadPreviousRARuleOverrideData
'------------------------------------------------------------------------------------------

Private Function IClassDefomRA_LoadRARuleData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IClassDefomRA_LoadRARuleData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>" & _
        "RISKASSESSMENTRULE" & _
        "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
        "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
        "<PRIMARYKEY>STAGENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
        "<PRIMARYKEY>RISKASSESSMENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
        "<PRIMARYKEY>RARULENUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
        "<OTHERS>RARULENAME<TYPE>dbdtString</TYPE></OTHERS>" & _
        "<OTHERS>RARULEVALUE<TYPE>dbdtString</TYPE></OTHERS>" & _
        "<OTHERS>RARULERESULT<TYPE>dbdtInt</TYPE></OTHERS>" & _
        "<OTHERS>RARULESCORE<TYPE>dbdtInt</TYPE></OTHERS>" & _
        "<OTHERS>SCORECARDIND<TYPE>dbdtBoolean</TYPE></OTHERS>" & _
     "</TABLENAME>"

    Set IClassDefomRA_LoadRARuleData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing


End Function

Private Function IClassDefomRA_LoadRARuleOverrideData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IClassDefomRA_LoadRARuleOverrideData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>" & _
        "RISKASSESSMENTRULEOVERRIDE" & _
        "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
        "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
        "<PRIMARYKEY>STAGENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
        "<PRIMARYKEY>RISKASSESSMENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
        "<PRIMARYKEY>RARULENUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
        "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>" & _
        "<OTHERS>RAOVERRIDEDATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>" & _
        "<OTHERS>RAOVERRIDEREASON<TYPE>dbdtString</TYPE></OTHERS>" & _
        "<OTHERS>RAOVERRIDEAPPROVALREFERENCE<TYPE>dbdtString</TYPE></OTHERS>" & _
        "<OTHERS>RAOVERRIDEREASONCODE<TYPE>dbdtInt</TYPE></OTHERS>" & _
     "</TABLENAME>"

    Set IClassDefomRA_LoadRARuleOverrideData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function



Private Function IClassDefomRA_LoadRiskAssessmentData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IClassDefomRA_LoadRiskAssessmentData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>" & _
        "RISKASSESSMENT" & _
        "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
        "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
        "<PRIMARYKEY>STAGENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
        "<PRIMARYKEY>RISKASSESSMENTSEQUENCENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
        "<OTHERS>RISKASSESSMENTSCORE<TYPE>dbdtInt</TYPE></OTHERS>" & _
        "<OTHERS>RISKASSESSMENTDECISION<TYPE>dbdtComboId</TYPE><COMBO>RiskAssessmentDecision</COMBO></OTHERS>" & _
        "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>" & _
        "<OTHERS>RISKASSESSMENTDATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>" & _
        "<OTHERS>UNDERWRITERDECISION<TYPE>dbdtComboId</TYPE><COMBO>RiskAssessmentDecision</COMBO></OTHERS>" & _
     "</TABLENAME>"

    Set IClassDefomRA_LoadRiskAssessmentData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IClassDefomRA_LoadPreviousRARuleOverrideData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IClassDefomRA_LoadPreviousRARuleOverrideData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>" & _
        "RISKASSESSMENTRULEOVERRIDE" & _
        "<OTHERS>USERID<TYPE>dbdtString</TYPE></OTHERS>" & _
        "<OTHERS>RAOVERRIDEREASON<TYPE>dbdtString</TYPE></OTHERS>" & _
        "<OTHERS>RAOVERRIDEDATETIME<TYPE>dbdtDateTime</TYPE></OTHERS>" & _
        "<OTHERS>RAOVERRIDEREASONCODE<TYPE>dbdtInt</TYPE></OTHERS>" & _
        "<OTHERS>RAOVERRIDEAPPROVALREFERENCE<TYPE>dbdtString</TYPE></OTHERS>" & _
     "</TABLENAME>"

    Set IClassDefomRA_LoadPreviousRARuleOverrideData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

