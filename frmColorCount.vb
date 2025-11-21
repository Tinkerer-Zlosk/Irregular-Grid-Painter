' ColorCount.vb
'
' Created by Dunedon
' Modified by Zlosk so it would be less memory intensive
' Added to IGP Distribution on 12-31-2003

Option Strict On
Option Explicit On 

Public Class frmColorCount
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
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lblColor As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents lblCount As System.Windows.Forms.Label
    Friend WithEvents btnSaveAsText As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents sfd As System.Windows.Forms.SaveFileDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.btnSaveAsText = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.lblColor = New System.Windows.Forms.Label
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.lblCount = New System.Windows.Forms.Label
        Me.sfd = New System.Windows.Forms.SaveFileDialog
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.PictureBox1.Location = New System.Drawing.Point(200, 24)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(120, 112)
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'btnSaveAsText
        '
        Me.btnSaveAsText.Location = New System.Drawing.Point(200, 144)
        Me.btnSaveAsText.Name = "btnSaveAsText"
        Me.btnSaveAsText.Size = New System.Drawing.Size(120, 24)
        Me.btnSaveAsText.TabIndex = 3
        Me.btnSaveAsText.Text = "&Save as text..."
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(200, 176)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(120, 24)
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        '
        'lblColor
        '
        Me.lblColor.Location = New System.Drawing.Point(8, 8)
        Me.lblColor.Name = "lblColor"
        Me.lblColor.Size = New System.Drawing.Size(32, 16)
        Me.lblColor.TabIndex = 5
        Me.lblColor.Text = "Color"
        '
        'ListBox1
        '
        Me.ListBox1.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.ItemHeight = 14
        Me.ListBox1.Location = New System.Drawing.Point(8, 24)
        Me.ListBox1.MultiColumn = True
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(184, 172)
        Me.ListBox1.TabIndex = 6
        '
        'lblCount
        '
        Me.lblCount.Location = New System.Drawing.Point(144, 8)
        Me.lblCount.Name = "lblCount"
        Me.lblCount.Size = New System.Drawing.Size(40, 16)
        Me.lblCount.TabIndex = 7
        Me.lblCount.Text = "Count"
        '
        'sfd
        '
        '
        'frmColorCount
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(328, 205)
        Me.Controls.Add(Me.lblCount)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.lblColor)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSaveAsText)
        Me.Controls.Add(Me.PictureBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmColorCount"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "ColorCount"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Public Cols() As Color, Cts() As Integer
    Public bmpColorPattern As System.Drawing.Bitmap

    Private Sub ColorCount_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rgb As Color
        Dim ColCt As Integer
        Dim i, j, c As Integer
        Dim blnColorExists As Boolean

        ColCt = 0
        For j = 0 To bmpColorPattern.Height - 1
            For i = 0 To bmpColorPattern.Width - 1
                rgb = bmpColorPattern.GetPixel(i, j)
                blnColorExists = False
                For c = 1 To ColCt
                    If (rgb.ToArgb = Cols(c).ToArgb) Then
                        blnColorExists = True
                        Cts(c) = Cts(c) + 1
                        Exit For
                    End If
                Next c
                If Not blnColorExists Then
                    ColCt = ColCt + 1
                    ReDim Preserve Cols(ColCt)
                    ReDim Preserve Cts(ColCt)
                    Cols(ColCt) = rgb
                    Cts(ColCt) = 1
                End If
            Next i
        Next j
        Dim iTotal As Integer = 0
        For c = 1 To ColCt
            Me.ListBox1.Items.Add(Cols(c).R.ToString.PadLeft(3) & Cols(c).G.ToString.PadLeft(6) & Cols(c).B.ToString.PadLeft(7) & Cts(c).ToString.PadLeft(9))
            iTotal += Cts(c)
            'Me.ListBox1.Items.Add("RGB(" & Cols(c).R.ToString & "," & Cols(c).G.ToString & "," & Cols(c).B.ToString & ")")
        Next c
        Me.ListBox1.Items.Add("")
        Me.ListBox1.Items.Add("Total".PadRight(16) & iTotal.ToString.PadLeft(9))
        Me.ListBox1.SelectedIndex = 0
    End Sub

    Private Sub ListBox1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.DoubleClick
        ' 2009.09.01a Addition to prevent crash when selecting items beyond color array. 
        Dim i As Integer
        i = ListBox1.SelectedIndex + 1
        If i > Cols.GetUpperBound(0) Then Exit Sub
        ' 2009.09.01a End code addition. PLW
        Dim str As String
        str = InputBox("Enter the new description (16 characters or less):", "Change Description", Me.ListBox1.Items(Me.ListBox1.SelectedIndex).ToString.Substring(0, 16))

        If str.Length > 0 Then
            If str.Length > 16 Then
                str = str.Substring(0, 16)
            End If

            ' 2009.09.01b Moved code below into IF block so Cancel button actually works.
            Me.ListBox1.Items(Me.ListBox1.SelectedIndex) = str.PadRight(16) & Cts(Me.ListBox1.SelectedIndex + 1).ToString.PadLeft(9)
        End If
        ' 2009.09.01b Original location of code. PLW

        'PictureBox1.BackColor = Cols(ListBox1.SelectedIndex + 1)
    End Sub

    Private Sub ListBox1_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        ' 2009.09.01c Addition to prevent crash when selecting items beyond color array. 
        Dim i As Integer
        i = ListBox1.SelectedIndex + 1
        If i > Cols.GetUpperBound(0) Then Exit Sub
        ' 2009.09.01c End code addition. PLW
        PictureBox1.BackColor = Cols(ListBox1.SelectedIndex + 1)
    End Sub

    Private Sub btnSaveAsText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveAsText.Click
        sfd.Filter = "Text (*.txt)|*.txt|All (*.*)|*.*"
        sfd.ShowDialog()
    End Sub
    Private Sub sfd_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles sfd.FileOk
        ' Create an instance of StreamWriter to write text to a file.
        Dim sw As System.IO.StreamWriter = New System.IO.StreamWriter(sfd.FileName)
        ' Add some text to the file.
        Dim i, j, argb, ct As Integer
        Dim txt As String
        Try
            For i = 0 To Me.ListBox1.Items.Count - 1
                sw.WriteLine(Me.ListBox1.Items(i))
            Next
            For j = 0 To bmpColorPattern.Height - 1
                ct = 0
                argb = -1
                sw.WriteLine()
                sw.Write((j + 1).ToString & ": ")
                For i = 0 To bmpColorPattern.Width - 1
                    If bmpColorPattern.GetPixel(i, j).ToArgb = argb Then
                        ct += 1
                    Else
                        If ct = 0 Then
                            argb = bmpColorPattern.GetPixel(i, j).ToArgb
                            ct = 1
                        Else
                            txt = Me.ListBox1.Items(FindColorInArray(argb) - 1).ToString.Substring(0, 16).Trim(" "c)
                            sw.Write(txt & "-" & ct.ToString & " ")
                            ct = 1
                            argb = bmpColorPattern.GetPixel(i, j).ToArgb
                        End If
                    End If
                Next
                txt = Me.ListBox1.Items(FindColorInArray(argb) - 1).ToString.Substring(0, 16).Trim(" "c)
                sw.Write(txt & "-" & ct.ToString)
            Next
            sw.WriteLine()
        Catch
            sw.Close()
        End Try
        sw.Close()
    End Sub

    Private Function FindColorInArray(ByVal argb As Integer) As Integer
        Dim c As Integer
        For c = 1 To Cols.GetUpperBound(0)
            If (argb = Cols(c).ToArgb) Then Return c
        Next c
        Return -1
    End Function

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click

    End Sub
End Class
