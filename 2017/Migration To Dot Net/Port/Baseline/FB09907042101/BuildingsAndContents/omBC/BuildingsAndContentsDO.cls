VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "BuildingsAndContentsDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      BuildingsAndContentsDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Code template for BuildingsAndContents Data Object
'
'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'JLD    21/12/1999  Created
'JLD    17/01/2000  Fixed sql query - bcstartdate should be <= today
'LD     07/11/2000  Explicity close recordsets
'LD     25/06/2001  SQL Server Port - IBuildingsAndContentsDO_FindProductList modified
'AD     04/10/2001  Removed MoveFirst as it was a ForwardOnly cursor. SYS2718
'------------------------------------------------------------------------------------------

Option Explicit

Implements IBuildingsAndContentsDO
 
 '=============================================
 'Variable Declaration Section
 '=============================================
 
 '=============================================
 'Constant Declaration Section
 '=============================================

Private Sub IBuildingsAndContentsDO_AddDerivedData(ByVal vxmlData As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   XML elements must be created for any derived values as specified.
'   Add any derived values to XML. E.g. data type 'double' fields will
'   need to be formatted as strings to required precision & rounding.
' pass:
'   vxmlData          base XML node
'                       as:
'                           <tablename>
'                               <element1>element1 value</element1>
'                               <elementn>elementn value</elementn>
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo AddDerivedDataVBErr

    Const strFunctionName As String = "IBuildingsAndContentsDO_AddDerivedData"

    Dim objErrAssist As New ErrAssist

    Set objErrAssist = Nothing
    
    Exit Sub
            
    Set objErrAssist = Nothing
    
    Exit Sub
    
AddDerivedDataVBErr:
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


Private Function IBuildingsAndContentsDO_FindProductList(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindProductListVbErr

    Const strFunctionName As String = "IBuildingsAndContentsDO_FindProductList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist As IADOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomBCClassDef
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlListElem As IXMLDOMElement
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
     
    Dim strSQL As String
    Dim strSystemDate As String

    Dim objSQLAssist As SQLAssist
    Set objSQLAssist = New SQLAssist
    strSystemDate = objSQLAssist.GetSystemDate()
    Set objSQLAssist = Nothing
    strSQL = "select * from BUILDINGANDCONTENTSPRODUCT where BCSTARTDATE <= " + strSystemDate + " and ENDDATE > " + strSystemDate
    
    Dim cmd As New ADODB.Command
        
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    
    cmd.CommandText = strSQL
    
    Dim rstThisRecSet As New ADODB.Recordset

    Set rstThisRecSet = cmd.Execute(, , adCmdText)

    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
        
    Set xmlListElem = xmlOut.createElement("BUILDINGANDCONTENTSPRODUCTLIST")
    xmlOut.appendChild xmlListElem
    
    ' AD 04/10/01 Removed the MoveFirst as it was a default ForwardOnly cursor. SYS2718
    
    Set objIClassDef = New omBCClassDef
    Set xmlClassDefDoc = objIClassDef.LoadBCProductData
    
    ' loop through the record set
    While Not rstThisRecSet.EOF
        objIDOAssist.GetXMLFromRecordSet rstThisRecSet, xmlClassDefDoc, xmlListElem
        rstThisRecSet.MoveNext
    Wend
    rstThisRecSet.Close

    Set rstThisRecSet = Nothing
    Set cmd = Nothing

    Set IBuildingsAndContentsDO_FindProductList = xmlListElem.cloneNode(True)


    ' TODO ************************************************************************
    '   Add Any derived data here
    ' *****************************************************************************

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set xmlOut = Nothing
    Set xmlListElem = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindProductListVbErr:
    
    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set xmlOut = Nothing
    Set xmlListElem = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


