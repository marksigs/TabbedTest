#Region "Header Comments"
'Workfile:      AlphaPlusSoapDoc.asmx.vb
'Copyright:     Copyright © 2004 Marlborough Stirling

'Description:   Provides Alpha Plus web service in SOAP Document format
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date     Description
' RF    25/02/04 Created
'------------------------------------------------------------------------------------------
#End Region

Option Strict On

Imports System.Web.Services
Imports System.Web.Services.Protocols

Namespace MAS.AlphaPlus.WebServices.AlphaPlusWebService

    <WebService(Namespace:="http://marlborough-stirling.com/WebServices/AlphaPlus")> _
        Public Class AlphaPlusSoapDoc
        Inherits System.Web.Services.WebService

#Region " Web Services Designer Generated Code "

        Public Sub New()
            MyBase.New()

            'This call is required by the Web Services Designer.
            InitializeComponent()

            'Add your own initialization code after the InitializeComponent() call

        End Sub

        'Required by the Web Services Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Web Services Designer
        'It can be modified using the Web Services Designer.  
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            components = New System.ComponentModel.Container()
        End Sub

        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            'CODEGEN: This procedure is required by the Web Services Designer
            'Do not modify it using the code editor.
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

#End Region

#Region "Web Methods"

        <WebMethod( _
            Description:="Service which allows a request to be submitted to the Marlborough Stirling Alpha Plus Calculations Engine")> _
        Public Function ACERequest(ByVal vstrRequest As String) As String
            Dim strRetVal As String
            Dim strErrorDescription As String
            Try
                Dim objCalc As ALPHACOMPLUSLib.IAlpha = New ALPHACOMPLUSLib.Alpha()
                strRetVal = objCalc.aceRequest(vstrRequest)
            Catch ex As System.Exception
                strErrorDescription = ex.Message()
                strRetVal = "<Response><Errors><Error TypeId=""1"" ErrorNumber=""1"" " & _
                   "Source=""AlphaPlusSoapDoc.ACERequest"" Severity=""1"" CodeTextId=""1"" " & _
                  "Description=""" & strErrorDescription & """></Error></Errors></Response>"
            Catch
                strErrorDescription = "Unknown exception"
                strRetVal = "<Response><Errors><Error TypeId=""1"" ErrorNumber=""1"" " & _
                   "Source=""AlphaPlusSoapDoc.ACERequest"" Severity=""1"" CodeTextId=""1"" " & _
                  "Description=""" & strErrorDescription & """></Error></Errors></Response>"
                Exit Try
            End Try
            Return strRetVal
        End Function

#End Region

    End Class

End Namespace



