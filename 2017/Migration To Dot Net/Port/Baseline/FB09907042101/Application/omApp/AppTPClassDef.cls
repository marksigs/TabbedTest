VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppTPClassDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      AppTPClassDef.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Code template for omiga4 Class Definition
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'AD     11/01/00    Created
'SR     18/08/00    Modified LoadBankBuildingSocietyData. Add new column 'DIRECTDEBITPAPERIND'
'--------------------------------------------------------------------------------------------

Option Explicit

Implements IAppTPClassDef

'------------------------------------------------------------------------------------------
'BMids History:
'
'Prog   Date        Description
'MDC    11/12/2002  BM0094 Legal Rep Contact Details
'MDC    13/01/2003  BM0244 Add DXID and DXLOCATION to ApplicationLegalRep
'GD     21/05/2003  BM0370 Change Type of DXID
'TL     06/09/2005  MAR39 Added "TRANSPOSEDINDICATOR" to "APPLICATIONBANKBUILDINGSOC"
'--------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------
' MARS History
'
'Prog   Date        AQR     Description
'MF     06/09/2005  MAR20   Added field DDEXPLAINEDIND to LoadBankBuildingSocietyData structure
'------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------
' EPSOM History
'
'Prog   Date        AQR     Description
'HMA    22/09/2006  EP2_3   Added field ROLLNUMBER to LoadBankBuildingSocietyData structure
'------------------------------------------------------------------------------------------

Private Function IAppTPClassDef_LoadArchitectData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IAppTPClassDef_LoadArchitectData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONARCHITECT"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>LEVELOFINDEMNITYCOVER<TYPE>dbdtComboId</TYPE><COMBO>IndemnityCoverLevel</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>QUALIFICATIONS<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>RIBAQUALIFIEDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>THIRDPARTYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IAppTPClassDef_LoadArchitectData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IAppTPClassDef_LoadBankBuildingSocietyData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IAppTPClassDef_LoadBankBuildingSocietyData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONBANKBUILDINGSOC"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    'DM     13/11/02    BMIDS00935 added nolock hint
    strXML = strXML & "<PRIMARYKEY>BANKACCOUNTSEQUENCENUMBER<TYPE>dbdtInt</TYPE><SQLNOLOCK>TRUE</SQLNOLOCK></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ACCOUNTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ACCOUNTNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>BANKWIZARDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DDREFERENCE<TYPE>dbdtString</TYPE></OTHERS>"
    '~ MF MAR20 New field Direct Debit explained flag
    strXML = strXML & "<OTHERS>DDEXPLAINEDIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PREFEREDPAYMENTDAY<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PROPOSEDPAYMENTMETHOD<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>REPAYMENTBANKACCOUNTINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TRANSPOSEDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TIMEATBANK<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>THIRDPARTYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SALARYFEDINDICATOR<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DIRECTDEBITPAPERIND<TYPE>dbdtBoolean</TYPE></OTHERS>"
    'strXML = strXML & "<OTHERS>ACCESSAUDITGUID<TYPE>dbdtGUID</TYPE></OTHERS>"
    'EP2_3 Add Roll Number
    strXML = strXML & "<OTHERS>ROLLNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IAppTPClassDef_LoadBankBuildingSocietyData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
End Function

Private Function IAppTPClassDef_LoadBuilderData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IAppTPClassDef_LoadBuilderData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONBUILDER"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>BUILDERPROPERTYDESCRIPTION<TYPE>dbdtString</TYPE><COMBO>IndemnityCoverLevel</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>EXPECTEDCOMPLETIONDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>THIRDPARTYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IAppTPClassDef_LoadBuilderData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing
End Function

Private Function IAppTPClassDef_LoadEstateAgentData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IAppTPClassDef_LoadEstateAgentData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONESTATEAGENT"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>THIRDPARTYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    Set IAppTPClassDef_LoadEstateAgentData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function

Private Function IAppTPClassDef_LoadLegalRepData() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "IAppTPClassDef_LoadLegalRepData"

    Dim strXML As String
    Dim objXmlAssist As New XMLAssist
    
    strXML = strXML & "<TABLENAME>APPLICATIONLEGALREP"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>SEPARATELEGALREPRESENTATIVE<TYPE>dbdtBoolean</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DIRECTORYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>THIRDPARTYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    'BM0094 MDC 11/12/2002
    strXML = strXML & "<OTHERS>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    'BM0094 MDC 11/12/2002 - End
    'BM0244 MDC 10/01/2003
    'GD BM0370 START
    'strXML = strXML & "<OTHERS>DXID<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DXID<TYPE>dbdtString</TYPE></OTHERS>"
    'GD BM0370 END
    

    strXML = strXML & "<OTHERS>DXLOCATION<TYPE>dbdtString</TYPE></OTHERS>"
    'BM0244 MDC 10/01/2003 - End
    
    strXML = strXML & "</TABLENAME>"

    Set IAppTPClassDef_LoadLegalRepData = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    
    Set objXmlAssist = Nothing

End Function


