Option Explicit On 

Imports System.Xml

Public Class FormMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents btnCalc As System.Windows.Forms.Button
    Friend WithEvents txtInputFile As System.Windows.Forms.TextBox
    Friend WithEvents txtOutputFile As System.Windows.Forms.TextBox
    Friend WithEvents btnSelectInput As System.Windows.Forms.Button
    Friend WithEvents btnSelectOutput As System.Windows.Forms.Button
    Friend WithEvents optDoc As System.Windows.Forms.RadioButton
    Friend WithEvents optRpc As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnCalc = New System.Windows.Forms.Button()
        Me.btnSelectInput = New System.Windows.Forms.Button()
        Me.btnSelectOutput = New System.Windows.Forms.Button()
        Me.txtInputFile = New System.Windows.Forms.TextBox()
        Me.txtOutputFile = New System.Windows.Forms.TextBox()
        Me.optDoc = New System.Windows.Forms.RadioButton()
        Me.optRpc = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnCalc
        '
        Me.btnCalc.Location = New System.Drawing.Point(304, 160)
        Me.btnCalc.Name = "btnCalc"
        Me.btnCalc.Size = New System.Drawing.Size(112, 23)
        Me.btnCalc.TabIndex = 1
        Me.btnCalc.Text = "Calculate"
        '
        'btnSelectInput
        '
        Me.btnSelectInput.Location = New System.Drawing.Point(360, 56)
        Me.btnSelectInput.Name = "btnSelectInput"
        Me.btnSelectInput.Size = New System.Drawing.Size(56, 23)
        Me.btnSelectInput.TabIndex = 0
        Me.btnSelectInput.Text = "Select"
        '
        'btnSelectOutput
        '
        Me.btnSelectOutput.Location = New System.Drawing.Point(360, 120)
        Me.btnSelectOutput.Name = "btnSelectOutput"
        Me.btnSelectOutput.Size = New System.Drawing.Size(56, 23)
        Me.btnSelectOutput.TabIndex = 2
        Me.btnSelectOutput.Text = "Select"
        '
        'txtInputFile
        '
        Me.txtInputFile.Location = New System.Drawing.Point(64, 24)
        Me.txtInputFile.Name = "txtInputFile"
        Me.txtInputFile.Size = New System.Drawing.Size(352, 20)
        Me.txtInputFile.TabIndex = 3
        Me.txtInputFile.Text = ""
        '
        'txtOutputFile
        '
        Me.txtOutputFile.Location = New System.Drawing.Point(64, 96)
        Me.txtOutputFile.Name = "txtOutputFile"
        Me.txtOutputFile.Size = New System.Drawing.Size(352, 20)
        Me.txtOutputFile.TabIndex = 4
        Me.txtOutputFile.Text = ""
        '
        'optDoc
        '
        Me.optDoc.Checked = True
        Me.optDoc.Location = New System.Drawing.Point(40, 160)
        Me.optDoc.Name = "optDoc"
        Me.optDoc.TabIndex = 5
        Me.optDoc.TabStop = True
        Me.optDoc.Text = "Soap Doc"
        '
        'optRpc
        '
        Me.optRpc.Location = New System.Drawing.Point(152, 160)
        Me.optRpc.Name = "optRpc"
        Me.optRpc.TabIndex = 6
        Me.optRpc.Text = "Soap Rpc"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 23)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Input file"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 23)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Output file"
        '
        'FormMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(432, 205)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label2, Me.Label1, Me.optRpc, Me.optDoc, Me.txtOutputFile, Me.txtInputFile, Me.btnSelectOutput, Me.btnSelectInput, Me.btnCalc})
        Me.Name = "FormMain"
        Me.Text = "Web Service Test Harness"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim m_docRequest As XmlDocument

    Private Sub btnCalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalc.Click
        Try
            If m_docRequest Is Nothing Then
                Throw New Exception("No request specified")
            End If
            If m_docRequest.DocumentElement.IsEmpty() Then
                Throw New Exception("No request specified")
            End If

            If Me.txtOutputFile.TextLength = 0 Then
                Throw New Exception("No output file specified")
            End If

            Dim strResponse As String

            If Me.optDoc.Checked Then
                Dim objCalc As webreference_doc.AlphaPlusSoapDoc = _
                    New webreference_doc.AlphaPlusSoapDoc()

                If objCalc Is Nothing Then
                    Throw New Exception("Cannot instantiate web reference")
                End If
                strResponse = objCalc.ACERequest(m_docRequest.DocumentElement.OuterXml())
            Else
                Dim objCalc As webreference_rpc.AlphaPlusSoapRpc = _
                    New webreference_rpc.AlphaPlusSoapRpc()
                If objCalc Is Nothing Then
                    Throw New Exception("Cannot instantiate web reference")
                End If
                strResponse = objCalc.ACERequest(m_docRequest.DocumentElement.OuterXml())
            End If

            MsgBox("Finished")

            Dim docResponse As XmlDocument = New XmlDocument()
            docResponse.LoadXml(strResponse)
            docResponse.Save(Me.txtOutputFile.Text)
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try

    End Sub

    Private Sub btnSelectInput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectInput.Click
        Try
            Dim objDlg As OpenFileDialog = New OpenFileDialog()
            objDlg.FileName = Me.txtInputFile.Text
            objDlg.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
            objDlg.FilterIndex = 1

            If objDlg.ShowDialog() = DialogResult.OK Then
                Me.txtInputFile.Text = objDlg.FileName()
                m_docRequest.Load(objDlg.FileName())
            End If
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub

    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        m_docRequest = New XmlDocument()
    End Sub

    Private Sub btnSelectOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectOutput.Click
        Try
            Dim objDlg As SaveFileDialog = New SaveFileDialog()
            objDlg.FileName = Me.txtOutputFile.Text
            objDlg.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*"
            objDlg.FilterIndex = 1
            If objDlg.ShowDialog() = DialogResult.OK Then
                Me.txtOutputFile.Text = objDlg.FileName()
            End If
        Catch ex As Exception
            MsgBox(ex.Message())
        End Try
    End Sub

End Class
