Option Explicit On 
Option Strict On

Public Class frmFileNew
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
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents grpImageParams As System.Windows.Forms.GroupBox
    Friend WithEvents cboHeight As System.Windows.Forms.ComboBox
    Friend WithEvents cboWidth As System.Windows.Forms.ComboBox
    Friend WithEvents txtHeight As System.Windows.Forms.TextBox
    Friend WithEvents lblHeight As System.Windows.Forms.Label
    Friend WithEvents txtWidth As System.Windows.Forms.TextBox
    Friend WithEvents lblWidth As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.lblName = New System.Windows.Forms.Label()
        Me.grpImageParams = New System.Windows.Forms.GroupBox()
        Me.cboHeight = New System.Windows.Forms.ComboBox()
        Me.cboWidth = New System.Windows.Forms.ComboBox()
        Me.txtHeight = New System.Windows.Forms.TextBox()
        Me.lblHeight = New System.Windows.Forms.Label()
        Me.txtWidth = New System.Windows.Forms.TextBox()
        Me.lblWidth = New System.Windows.Forms.Label()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.grpImageParams.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtName
        '
        Me.txtName.Enabled = False
        Me.txtName.Location = New System.Drawing.Point(48, 8)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(176, 20)
        Me.txtName.TabIndex = 1
        Me.txtName.Text = "Untitled"
        '
        'lblName
        '
        Me.lblName.Enabled = False
        Me.lblName.Location = New System.Drawing.Point(8, 8)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(40, 20)
        Me.lblName.TabIndex = 0
        Me.lblName.Text = "Name:"
        Me.lblName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'grpImageParams
        '
        Me.grpImageParams.Controls.AddRange(New System.Windows.Forms.Control() {Me.cboHeight, Me.cboWidth, Me.txtHeight, Me.lblHeight, Me.txtWidth, Me.lblWidth})
        Me.grpImageParams.Location = New System.Drawing.Point(8, 40)
        Me.grpImageParams.Name = "grpImageParams"
        Me.grpImageParams.Size = New System.Drawing.Size(216, 80)
        Me.grpImageParams.TabIndex = 2
        Me.grpImageParams.TabStop = False
        Me.grpImageParams.Text = "Image Parameters"
        '
        'cboHeight
        '
        Me.cboHeight.Items.AddRange(New Object() {"tile rows"})
        Me.cboHeight.Location = New System.Drawing.Point(120, 48)
        Me.cboHeight.Name = "cboHeight"
        Me.cboHeight.Size = New System.Drawing.Size(88, 21)
        Me.cboHeight.TabIndex = 5
        Me.cboHeight.Text = "tile rows"
        '
        'cboWidth
        '
        Me.cboWidth.Items.AddRange(New Object() {"tile columns"})
        Me.cboWidth.Location = New System.Drawing.Point(120, 24)
        Me.cboWidth.Name = "cboWidth"
        Me.cboWidth.Size = New System.Drawing.Size(88, 21)
        Me.cboWidth.TabIndex = 2
        Me.cboWidth.Text = "tile columns"
        '
        'txtHeight
        '
        Me.txtHeight.Location = New System.Drawing.Point(56, 48)
        Me.txtHeight.Name = "txtHeight"
        Me.txtHeight.Size = New System.Drawing.Size(56, 20)
        Me.txtHeight.TabIndex = 4
        Me.txtHeight.Text = ""
        '
        'lblHeight
        '
        Me.lblHeight.Location = New System.Drawing.Point(8, 48)
        Me.lblHeight.Name = "lblHeight"
        Me.lblHeight.Size = New System.Drawing.Size(48, 20)
        Me.lblHeight.TabIndex = 3
        Me.lblHeight.Text = "&Height:"
        Me.lblHeight.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtWidth
        '
        Me.txtWidth.Location = New System.Drawing.Point(56, 24)
        Me.txtWidth.Name = "txtWidth"
        Me.txtWidth.Size = New System.Drawing.Size(56, 20)
        Me.txtWidth.TabIndex = 1
        Me.txtWidth.Text = ""
        '
        'lblWidth
        '
        Me.lblWidth.Location = New System.Drawing.Point(3, 24)
        Me.lblWidth.Name = "lblWidth"
        Me.lblWidth.Size = New System.Drawing.Size(53, 20)
        Me.lblWidth.TabIndex = 0
        Me.lblWidth.Text = "&Width:"
        Me.lblWidth.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(240, 8)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.TabIndex = 3
        Me.btnOK.Text = "OK"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(240, 40)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        '
        'frmFileNew
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(328, 133)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnOK, Me.grpImageParams, Me.lblName, Me.txtName, Me.btnCancel})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmFileNew"
        Me.Text = "New"
        Me.grpImageParams.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private mblnOK As Boolean = False
    Private mstrFilename As String
    Private mintWidth, mintHeight As Integer
    Private mintWidthCboIdx, mintHeightCboIdx As Integer

    Public ReadOnly Property OK() As Boolean
        Get
            Return Me.mblnOK
        End Get
    End Property
    Public Property ImgWidth() As Integer
        Get
            Return mintWidth
        End Get
        Set(ByVal Value As Integer)
            mintWidth = Value
        End Set
    End Property
    Public Property ImgHeight() As Integer
        Get
            Return mintHeight
        End Get
        Set(ByVal Value As Integer)
            mintHeight = Value
        End Set
    End Property

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If IsIntegerGreaterThan(txtWidth.Text, 0) And IsIntegerGreaterThan(txtHeight.Text, 0) Then
            Me.ImgWidth = CInt(txtWidth.Text)
            Me.ImgHeight = CInt(txtHeight.Text)
            mblnOK = True
            Me.Hide()
        Else
            MsgBox("Image width and height must be integers greater than zero.", MsgBoxStyle.Information And MsgBoxStyle.OKOnly, "Error: Illegal Value")
        End If
    End Sub

End Class
