Attribute VB_Name = "om4COGGlobals"
Option Explicit
Private gxmlComponentsDoc As FreeThreadedDOMDocument40
Private Sub LoadComponentDetails()
    
    Dim strFileSpec As String
    ' pick up XML map from "...\Omiga 4\XML" directory
    ' Only do the subsitution once to change DLL -> XML
    strFileSpec = App.Path & "\" & App.Title & ".xml"
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    Set gxmlComponentsDoc = New FreeThreadedDOMDocument40
    gxmlComponentsDoc.async = False
    gxmlComponentsDoc.setProperty "NewParser", True
    gxmlComponentsDoc.validateOnParse = False
    gxmlComponentsDoc.Load strFileSpec
End Sub
Public Sub Main()
    ' adoAssist
    adoBuildDbConnectionString
    LoadComponentDetails
End Sub
Public Function RequestConversionRequired(ByVal vstrComponentName As String) As Boolean
    
    If gxmlComponentsDoc Is Nothing Then
        App.LogEvent "omCOG no component details [xml] file"
        Exit Function
    End If
    Dim strFilter As String
    strFilter = "COMPONENTS/" & UCase(vstrComponentName) & "/@CTYPE"
    If Not gxmlComponentsDoc.selectSingleNode(strFilter) Is Nothing Then
        If gxmlComponentsDoc.selectSingleNode(strFilter).Text = "1" Then
            RequestConversionRequired = True
        End If
    End If
End Function
Public Function GetConversionType(ByVal vstrComponentName As String) As Integer
    
    If gxmlComponentsDoc Is Nothing Then
        App.LogEvent "omCOG no component details [xml] file"
        Exit Function
    End If
    Dim strFilter As String
    Dim intConvType As Integer
    ' default 'conversion type' is 2
    ' i.e. phase 2 component
    intConvType = 2
    strFilter = "COMPONENTS/" & UCase(vstrComponentName) & "/@CTYPE"
    If Not gxmlComponentsDoc.selectSingleNode(strFilter) Is Nothing Then
        If IsNumeric(gxmlComponentsDoc.selectSingleNode(strFilter).Text) Then
            intConvType = CInt(gxmlComponentsDoc.selectSingleNode(strFilter).Text)
        End If
    End If
    GetConversionType = intConvType
End Function
Public Function ResponseConversionRequired(ByVal vstrComponentName As String) As Boolean
    ' currently always return TRUE
    ' this will translate element based ERROR responses from phase2 components
    ResponseConversionRequired = True
End Function
