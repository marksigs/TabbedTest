Attribute VB_Name = "globalAssist"
' DM    SYS3185 7/01/02 Added method to get a boolean value from a global parameter
' PB    EP2_1627 09/03/2007 Initialise gblnSection2

Option Explicit

' Section2 flag - indicates whether section 2 is used
Global gblnSection2 As Boolean

Private Enum GLOBALPARAMTYPE
    oegptString
    oegptAmount
    oegptMaximumAmount
    oegptPercentage
    oegptBoolean
End Enum

Private Const gstrMODULEPREFIX As String = "globalAssist."

' DM    SYS3185 7/01/02 Added method to get a boolean value from a global parameter
Public Function GetGlobalParamBoolean(ByVal vstrParamName As String) As Boolean

    Dim varParamValue As Boolean
    If GetGlobalParam(vstrParamName, oegptBoolean, varParamValue) Then
        GetGlobalParamBoolean = varParamValue
    End If

End Function

' DM    SYS3185 7/01/02 Added method to get a boolean value from a global parameter
Public Function GetMandatoryGlobalParamBoolean(ByVal vstrParamName As String) As Boolean
    
    Const cstrFunctionName As String = "GetMandatoryGlobalParamBoolean"
    
    Dim varParamValue As String
    
    If GetGlobalParam(vstrParamName, oegptBoolean, varParamValue) Then
        
        GetMandatoryGlobalParamBoolean = varParamValue
    
    Else
        
        errThrowError _
            gstrMODULEPREFIX & cstrFunctionName, _
            oeMissingParameter, _
            vstrParamName
            
    End If

End Function

Public Function GetMandatoryGlobalParamString(ByVal vstrParamName As String) As String
    
    Const cstrFunctionName As String = "GetMandatoryGlobalParamString"
    
    Dim varParamValue As String
    
    If GetGlobalParam(vstrParamName, oegptString, varParamValue) Then
        
        GetMandatoryGlobalParamString = varParamValue
    
    Else
        
        errThrowError _
            gstrMODULEPREFIX & cstrFunctionName, _
            oeMissingParameter, _
            vstrParamName
            
    End If

End Function

' APS SYS1920 Added new method
Public Function GetMandatoryGlobalParamAmount(ByVal vstrParamName As String) As Long
    
    Const cstrFunctionName As String = "GetMandatoryGlobalParamAmount"
    
    Dim varParamValue As String
    
    If GetGlobalParam(vstrParamName, oegptAmount, varParamValue) Then
        
        GetMandatoryGlobalParamAmount = varParamValue
    
    Else
        
        errThrowError _
            gstrMODULEPREFIX & cstrFunctionName, _
            oeMissingParameter, _
            vstrParamName
            
    End If

End Function

Public Function GetGlobalParamString(ByVal vstrParamName As String) As String
    Dim varParamValue As String
    If GetGlobalParam(vstrParamName, oegptString, varParamValue) Then
        GetGlobalParamString = varParamValue
    End If
End Function
' APS SYS1920 Added new method
Public Function GetGlobalParamAmount(ByVal vstrParamName As String) As Long
    Dim varParamValue As String
    If GetGlobalParam(vstrParamName, oegptAmount, varParamValue) Then
        GetGlobalParamAmount = varParamValue
    End If
End Function

' PK MAR1648 Add new procedure for global parameters with percentages
Public Function GetGlobalParamPercentage(ByVal vstrParamName As String) As Double
    Dim varParamValue As String
    If GetGlobalParam(vstrParamName, oegptPercentage, varParamValue) Then
        GetGlobalParamPercentage = varParamValue
    End If
End Function

Private Function GetGlobalParam( _
    ByVal vstrParamName As String, _
    ByVal enumParamType As GLOBALPARAMTYPE, _
    ByRef vvarParamValue As Variant) _
    As Boolean
    
    Const cstrFunctionName As String = "GetGlobalParam"
    
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlRootNode As IXMLDOMNode
    Dim xmlSchemaNode As IXMLDOMNode
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlResponseNode As IXMLDOMNode
    
    ' create one DOMDocument to contain schema, request & response
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    ' create BOGUS root node
    Set xmlElem = xmlDoc.createElement("BOGUS")
    Set xmlRootNode = xmlDoc.appendChild(xmlElem)
    
    ' create schema for OM_GLOBALPARAM view
    Set xmlElem = xmlDoc.createElement("GLOBALPARAM")
    xmlElem.setAttribute "ENTITYTYPE", "LOGICAL"
    xmlElem.setAttribute "DATASRCE", "OM_GLOBALPARAM"
    Set xmlSchemaNode = xmlRootNode.appendChild(xmlElem)
    Set xmlElem = xmlDoc.createElement("NAME")
    xmlElem.setAttribute "DATATYPE", "STRING"
    xmlElem.setAttribute "LENGTH", "30"
    xmlSchemaNode.appendChild xmlElem
    Set xmlElem = xmlDoc.createElement("GLOBALPARAMETERSTARTDATE")
    xmlElem.setAttribute "DATATYPE", "DATE"
    xmlSchemaNode.appendChild xmlElem
    Set xmlElem = xmlDoc.createElement("STRING")
    xmlElem.setAttribute "DATATYPE", "STRING"
    xmlElem.setAttribute "LENGTH", "30"
    xmlSchemaNode.appendChild xmlElem
    Set xmlElem = xmlDoc.createElement("DESCRIPTION")
    xmlElem.setAttribute "DATATYPE", "STRING"
    xmlElem.setAttribute "LENGTH", "255"
    xmlSchemaNode.appendChild xmlElem
    Set xmlElem = xmlDoc.createElement("AMOUNT")
    xmlElem.setAttribute "DATATYPE", "DOUBLE"
    xmlSchemaNode.appendChild xmlElem
    Set xmlElem = xmlDoc.createElement("AMOUNT")
    xmlElem.setAttribute "DATATYPE", "DOUBLE"
    xmlSchemaNode.appendChild xmlElem
    Set xmlElem = xmlDoc.createElement("MAXIMUMAMOUNT")
    xmlElem.setAttribute "DATATYPE", "DOUBLE"
    xmlSchemaNode.appendChild xmlElem
    Set xmlElem = xmlDoc.createElement("PERCENTAGE")
    xmlElem.setAttribute "DATATYPE", "DOUBLE"
    xmlSchemaNode.appendChild xmlElem
    Set xmlElem = xmlDoc.createElement("BOOLEAN")
    xmlElem.setAttribute "DATATYPE", "BOOLEAN"
    xmlSchemaNode.appendChild xmlElem
    
    ' create REQUEST details
    Set xmlElem = xmlDoc.createElement("GLOBALPARAM")
    xmlElem.setAttribute "NAME", vstrParamName
    Set xmlRequestNode = xmlRootNode.appendChild(xmlElem)
    
    ' create RESPONSE node
    Set xmlElem = xmlDoc.createElement("RESPONSE")
    Set xmlResponseNode = xmlRootNode.appendChild(xmlElem)
    
    adoGetRecordSetAsXML xmlRequestNode, xmlSchemaNode, xmlResponseNode
    
    Set xmlResponseNode = xmlGetNode(xmlResponseNode, "GLOBALPARAM")
    
    If Not xmlResponseNode Is Nothing Then
    
        Select Case enumParamType
    
            Case oegptString
                If xmlAttributeValueExists(xmlResponseNode, "STRING") Then
                    vvarParamValue = xmlGetAttributeText(xmlResponseNode, "STRING")
                    GetGlobalParam = True
                End If
            
            Case oegptAmount
                If xmlAttributeValueExists(xmlResponseNode, "AMOUNT") Then
                    vvarParamValue = xmlGetAttributeAsDouble(xmlResponseNode, "AMOUNT")
                    GetGlobalParam = True
                End If
            
            Case oegptMaximumAmount
                If xmlAttributeValueExists(xmlResponseNode, "MAXIMUMAMOUNT") Then
                    vvarParamValue = xmlGetAttributeAsDouble(xmlResponseNode, "MAXIMUMAMOUNT")
                    GetGlobalParam = True
                End If
            
            Case oegptPercentage
                If xmlAttributeValueExists(xmlResponseNode, "PERCENTAGE") Then
                    vvarParamValue = xmlGetAttributeAsDouble(xmlResponseNode, "PERCENTAGE")
                    GetGlobalParam = True
                End If
            ' DM    SYS3185 7/01/02 Added method to get a boolean value from a global parameter
            Case oegptBoolean
                If xmlAttributeValueExists(xmlResponseNode, "BOOLEAN") Then
                    ' SYS3185 Changed this to call the xmlGetAttributeAsBoolean method
                    vvarParamValue = xmlGetAttributeAsBoolean(xmlResponseNode, "BOOLEAN")
                    GetGlobalParam = True
                End If
        
        End Select
        
    End If
    
    Set xmlElem = Nothing
    Set xmlRootNode = Nothing
    Set xmlSchemaNode = Nothing
    Set xmlRequestNode = Nothing
    Set xmlResponseNode = Nothing
    Set xmlDoc = Nothing
    
End Function
