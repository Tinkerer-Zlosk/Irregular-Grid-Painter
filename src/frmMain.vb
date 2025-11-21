Option Explicit On 
Option Strict On

Public Class frmMain
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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents ofd As System.Windows.Forms.OpenFileDialog
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileOpen As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileNew As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPatterns As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelpAbout As System.Windows.Forms.MenuItem
    Friend WithEvents sbMain As System.Windows.Forms.StatusBar
    Friend WithEvents mnuColors As System.Windows.Forms.MenuItem
    Friend WithEvents mnuColorsOutlines As System.Windows.Forms.MenuItem
    Friend WithEvents mnuColorsBackground As System.Windows.Forms.MenuItem
    Friend WithEvents mnuColorsLeft As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSeparator2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSave As System.Windows.Forms.MenuItem
    Friend WithEvents cd As System.Windows.Forms.ColorDialog
    Friend WithEvents sfd As System.Windows.Forms.SaveFileDialog
    Friend WithEvents mnuFileSavePicture As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSeparator1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSaveAs As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSeparator4 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileImport As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileImportMaterialLibrary As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileExport As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileExportMaterialLibrary As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileExportPovRayScene As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileExportPattern As System.Windows.Forms.MenuItem
    Friend WithEvents pnlColors As System.Windows.Forms.Panel
    Friend WithEvents btnLeft As System.Windows.Forms.Button
    Friend WithEvents btnRight As System.Windows.Forms.Button
    Friend WithEvents btnOutlines As System.Windows.Forms.Button
    Friend WithEvents btnBackground As System.Windows.Forms.Button
    Friend WithEvents mnuWindow As System.Windows.Forms.MenuItem
    Friend WithEvents mnuWindowColorPanel As System.Windows.Forms.MenuItem
    Friend WithEvents lblLeft As System.Windows.Forms.Label
    Friend WithEvents lblRight As System.Windows.Forms.Label
    Friend WithEvents lblOutlines As System.Windows.Forms.Label
    Friend WithEvents lblBackground As System.Windows.Forms.Label
    Friend WithEvents mnuFileClose As System.Windows.Forms.MenuItem
    Friend WithEvents mnuColorsRight As System.Windows.Forms.MenuItem
    Friend WithEvents pnlDisplay As System.Windows.Forms.Panel
    Friend WithEvents mnuFilePrint As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents tip As System.Windows.Forms.ToolTip
    Friend WithEvents pnlOuter As System.Windows.Forms.Panel
    Friend WithEvents btnColorCount As System.Windows.Forms.Button
    Friend WithEvents mnuColorsCountColors As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnuFileNew = New System.Windows.Forms.MenuItem
        Me.mnuFileOpen = New System.Windows.Forms.MenuItem
        Me.mnuSeparator1 = New System.Windows.Forms.MenuItem
        Me.mnuFileClose = New System.Windows.Forms.MenuItem
        Me.mnuFileSave = New System.Windows.Forms.MenuItem
        Me.mnuFileSaveAs = New System.Windows.Forms.MenuItem
        Me.mnuFileSavePicture = New System.Windows.Forms.MenuItem
        Me.mnuSeparator4 = New System.Windows.Forms.MenuItem
        Me.mnuFileImport = New System.Windows.Forms.MenuItem
        Me.mnuFileImportMaterialLibrary = New System.Windows.Forms.MenuItem
        Me.mnuFileExport = New System.Windows.Forms.MenuItem
        Me.mnuFileExportMaterialLibrary = New System.Windows.Forms.MenuItem
        Me.mnuFileExportPattern = New System.Windows.Forms.MenuItem
        Me.mnuFileExportPovRayScene = New System.Windows.Forms.MenuItem
        Me.mnuSeparator2 = New System.Windows.Forms.MenuItem
        Me.mnuFilePrint = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.mnuFileExit = New System.Windows.Forms.MenuItem
        Me.mnuPatterns = New System.Windows.Forms.MenuItem
        Me.mnuColors = New System.Windows.Forms.MenuItem
        Me.mnuColorsLeft = New System.Windows.Forms.MenuItem
        Me.mnuColorsRight = New System.Windows.Forms.MenuItem
        Me.mnuColorsOutlines = New System.Windows.Forms.MenuItem
        Me.mnuColorsBackground = New System.Windows.Forms.MenuItem
        Me.mnuColorsCountColors = New System.Windows.Forms.MenuItem
        Me.mnuWindow = New System.Windows.Forms.MenuItem
        Me.mnuWindowColorPanel = New System.Windows.Forms.MenuItem
        Me.mnuHelp = New System.Windows.Forms.MenuItem
        Me.mnuHelpAbout = New System.Windows.Forms.MenuItem
        Me.ofd = New System.Windows.Forms.OpenFileDialog
        Me.sbMain = New System.Windows.Forms.StatusBar
        Me.cd = New System.Windows.Forms.ColorDialog
        Me.sfd = New System.Windows.Forms.SaveFileDialog
        Me.pnlColors = New System.Windows.Forms.Panel
        Me.btnColorCount = New System.Windows.Forms.Button
        Me.lblLeft = New System.Windows.Forms.Label
        Me.btnLeft = New System.Windows.Forms.Button
        Me.btnRight = New System.Windows.Forms.Button
        Me.lblRight = New System.Windows.Forms.Label
        Me.lblOutlines = New System.Windows.Forms.Label
        Me.btnOutlines = New System.Windows.Forms.Button
        Me.lblBackground = New System.Windows.Forms.Label
        Me.btnBackground = New System.Windows.Forms.Button
        Me.pnlDisplay = New System.Windows.Forms.Panel
        Me.tip = New System.Windows.Forms.ToolTip(Me.components)
        Me.pnlOuter = New System.Windows.Forms.Panel
        Me.pnlColors.SuspendLayout()
        Me.pnlOuter.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuPatterns, Me.mnuColors, Me.mnuWindow, Me.mnuHelp})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFileNew, Me.mnuFileOpen, Me.mnuSeparator1, Me.mnuFileClose, Me.mnuFileSave, Me.mnuFileSaveAs, Me.mnuFileSavePicture, Me.mnuSeparator4, Me.mnuFileImport, Me.mnuFileExport, Me.mnuSeparator2, Me.mnuFilePrint, Me.MenuItem2, Me.mnuFileExit})
        Me.mnuFile.Shortcut = System.Windows.Forms.Shortcut.CtrlN
        Me.mnuFile.Text = "&File"
        '
        'mnuFileNew
        '
        Me.mnuFileNew.Index = 0
        Me.mnuFileNew.Shortcut = System.Windows.Forms.Shortcut.CtrlN
        Me.mnuFileNew.Text = "&New..."
        '
        'mnuFileOpen
        '
        Me.mnuFileOpen.Index = 1
        Me.mnuFileOpen.Shortcut = System.Windows.Forms.Shortcut.CtrlO
        Me.mnuFileOpen.Text = "&Open..."
        '
        'mnuSeparator1
        '
        Me.mnuSeparator1.Index = 2
        Me.mnuSeparator1.Text = "-"
        '
        'mnuFileClose
        '
        Me.mnuFileClose.Enabled = False
        Me.mnuFileClose.Index = 3
        Me.mnuFileClose.Shortcut = System.Windows.Forms.Shortcut.CtrlW
        Me.mnuFileClose.Text = "&Close"
        '
        'mnuFileSave
        '
        Me.mnuFileSave.Enabled = False
        Me.mnuFileSave.Index = 4
        Me.mnuFileSave.Shortcut = System.Windows.Forms.Shortcut.CtrlS
        Me.mnuFileSave.Text = "&Save..."
        '
        'mnuFileSaveAs
        '
        Me.mnuFileSaveAs.Enabled = False
        Me.mnuFileSaveAs.Index = 5
        Me.mnuFileSaveAs.Shortcut = System.Windows.Forms.Shortcut.CtrlShiftS
        Me.mnuFileSaveAs.Text = "Save &As..."
        '
        'mnuFileSavePicture
        '
        Me.mnuFileSavePicture.Enabled = False
        Me.mnuFileSavePicture.Index = 6
        Me.mnuFileSavePicture.Text = "Save Patterned Picture..."
        '
        'mnuSeparator4
        '
        Me.mnuSeparator4.Index = 7
        Me.mnuSeparator4.Text = "-"
        '
        'mnuFileImport
        '
        Me.mnuFileImport.Enabled = False
        Me.mnuFileImport.Index = 8
        Me.mnuFileImport.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFileImportMaterialLibrary})
        Me.mnuFileImport.Text = "&Import"
        '
        'mnuFileImportMaterialLibrary
        '
        Me.mnuFileImportMaterialLibrary.Enabled = False
        Me.mnuFileImportMaterialLibrary.Index = 0
        Me.mnuFileImportMaterialLibrary.Text = "&Material Library"
        '
        'mnuFileExport
        '
        Me.mnuFileExport.Enabled = False
        Me.mnuFileExport.Index = 9
        Me.mnuFileExport.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFileExportMaterialLibrary, Me.mnuFileExportPattern, Me.mnuFileExportPovRayScene})
        Me.mnuFileExport.Text = "&Export"
        '
        'mnuFileExportMaterialLibrary
        '
        Me.mnuFileExportMaterialLibrary.Index = 0
        Me.mnuFileExportMaterialLibrary.Text = "&Material Library"
        '
        'mnuFileExportPattern
        '
        Me.mnuFileExportPattern.Index = 1
        Me.mnuFileExportPattern.Text = "&Pattern"
        '
        'mnuFileExportPovRayScene
        '
        Me.mnuFileExportPovRayScene.Index = 2
        Me.mnuFileExportPovRayScene.Text = "POV-Ray &Scene"
        '
        'mnuSeparator2
        '
        Me.mnuSeparator2.Index = 10
        Me.mnuSeparator2.Text = "-"
        '
        'mnuFilePrint
        '
        Me.mnuFilePrint.Enabled = False
        Me.mnuFilePrint.Index = 11
        Me.mnuFilePrint.Shortcut = System.Windows.Forms.Shortcut.CtrlP
        Me.mnuFilePrint.Text = "&Print..."
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 12
        Me.MenuItem2.Text = "-"
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Index = 13
        Me.mnuFileExit.Text = "E&xit" & Microsoft.VisualBasic.ChrW(9) & "Alt+F4"
        '
        'mnuPatterns
        '
        Me.mnuPatterns.Index = 1
        Me.mnuPatterns.Text = "&Patterns"
        '
        'mnuColors
        '
        Me.mnuColors.Index = 2
        Me.mnuColors.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuColorsLeft, Me.mnuColorsRight, Me.mnuColorsOutlines, Me.mnuColorsBackground, Me.mnuColorsCountColors})
        Me.mnuColors.Text = "&Colors"
        '
        'mnuColorsLeft
        '
        Me.mnuColorsLeft.Index = 0
        Me.mnuColorsLeft.Text = "Set &Left Button Color..."
        '
        'mnuColorsRight
        '
        Me.mnuColorsRight.Index = 1
        Me.mnuColorsRight.Text = "Set &Right Button Color..."
        '
        'mnuColorsOutlines
        '
        Me.mnuColorsOutlines.Index = 2
        Me.mnuColorsOutlines.Text = "Set &Outline Color..."
        '
        'mnuColorsBackground
        '
        Me.mnuColorsBackground.Index = 3
        Me.mnuColorsBackground.Text = "Set &Background Color..."
        '
        'mnuColorsCountColors
        '
        Me.mnuColorsCountColors.Index = 4
        Me.mnuColorsCountColors.Text = "&Count Colors..."
        '
        'mnuWindow
        '
        Me.mnuWindow.Index = 3
        Me.mnuWindow.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuWindowColorPanel})
        Me.mnuWindow.Text = "&Window"
        '
        'mnuWindowColorPanel
        '
        Me.mnuWindowColorPanel.Index = 0
        Me.mnuWindowColorPanel.Text = "Hide &Color Panel"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 4
        Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuHelpAbout})
        Me.mnuHelp.Text = "&Help"
        '
        'mnuHelpAbout
        '
        Me.mnuHelpAbout.Index = 0
        Me.mnuHelpAbout.Text = "&About Irregular Grid Painter..."
        '
        'ofd
        '
        '
        'sbMain
        '
        Me.sbMain.Location = New System.Drawing.Point(0, 413)
        Me.sbMain.Name = "sbMain"
        Me.sbMain.Size = New System.Drawing.Size(632, 20)
        Me.sbMain.TabIndex = 0
        '
        'sfd
        '
        Me.sfd.FileName = "doc1"
        '
        'pnlColors
        '
        Me.pnlColors.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.pnlColors.BackColor = System.Drawing.SystemColors.Control
        Me.pnlColors.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlColors.Controls.Add(Me.btnColorCount)
        Me.pnlColors.Controls.Add(Me.lblLeft)
        Me.pnlColors.Controls.Add(Me.btnLeft)
        Me.pnlColors.Controls.Add(Me.btnRight)
        Me.pnlColors.Controls.Add(Me.lblRight)
        Me.pnlColors.Controls.Add(Me.lblOutlines)
        Me.pnlColors.Controls.Add(Me.btnOutlines)
        Me.pnlColors.Controls.Add(Me.lblBackground)
        Me.pnlColors.Controls.Add(Me.btnBackground)
        Me.pnlColors.Location = New System.Drawing.Point(512, 0)
        Me.pnlColors.Name = "pnlColors"
        Me.pnlColors.Size = New System.Drawing.Size(120, 136)
        Me.pnlColors.TabIndex = 2
        '
        'btnColorCount
        '
        Me.btnColorCount.Location = New System.Drawing.Point(8, 104)
        Me.btnColorCount.Name = "btnColorCount"
        Me.btnColorCount.Size = New System.Drawing.Size(104, 24)
        Me.btnColorCount.TabIndex = 2
        Me.btnColorCount.Text = "Count Colors..."
        '
        'lblLeft
        '
        Me.lblLeft.Location = New System.Drawing.Point(32, 8)
        Me.lblLeft.Name = "lblLeft"
        Me.lblLeft.Size = New System.Drawing.Size(80, 16)
        Me.lblLeft.TabIndex = 1
        Me.lblLeft.Text = "Left Button"
        Me.tip.SetToolTip(Me.lblLeft, "Set the drawing color of the left mouse button")
        '
        'btnLeft
        '
        Me.btnLeft.Location = New System.Drawing.Point(8, 8)
        Me.btnLeft.Name = "btnLeft"
        Me.btnLeft.Size = New System.Drawing.Size(16, 16)
        Me.btnLeft.TabIndex = 0
        Me.tip.SetToolTip(Me.btnLeft, "Set the drawing color of the left mouse button")
        '
        'btnRight
        '
        Me.btnRight.Location = New System.Drawing.Point(8, 32)
        Me.btnRight.Name = "btnRight"
        Me.btnRight.Size = New System.Drawing.Size(16, 16)
        Me.btnRight.TabIndex = 0
        Me.tip.SetToolTip(Me.btnRight, "Set the drawing color of the right mouse button")
        '
        'lblRight
        '
        Me.lblRight.Location = New System.Drawing.Point(32, 32)
        Me.lblRight.Name = "lblRight"
        Me.lblRight.Size = New System.Drawing.Size(80, 16)
        Me.lblRight.TabIndex = 1
        Me.lblRight.Text = "Right Button"
        Me.tip.SetToolTip(Me.lblRight, "Set the drawing color of the right mouse button")
        '
        'lblOutlines
        '
        Me.lblOutlines.Location = New System.Drawing.Point(32, 56)
        Me.lblOutlines.Name = "lblOutlines"
        Me.lblOutlines.Size = New System.Drawing.Size(80, 16)
        Me.lblOutlines.TabIndex = 1
        Me.lblOutlines.Text = "Outlines"
        Me.tip.SetToolTip(Me.lblOutlines, "Set the color of the outlines on the display")
        '
        'btnOutlines
        '
        Me.btnOutlines.Location = New System.Drawing.Point(8, 56)
        Me.btnOutlines.Name = "btnOutlines"
        Me.btnOutlines.Size = New System.Drawing.Size(16, 16)
        Me.btnOutlines.TabIndex = 0
        Me.tip.SetToolTip(Me.btnOutlines, "Set the color of the outlines on the display")
        '
        'lblBackground
        '
        Me.lblBackground.Location = New System.Drawing.Point(32, 80)
        Me.lblBackground.Name = "lblBackground"
        Me.lblBackground.Size = New System.Drawing.Size(80, 16)
        Me.lblBackground.TabIndex = 1
        Me.lblBackground.Text = "Background"
        Me.tip.SetToolTip(Me.lblBackground, "Set the color of the background on the display")
        '
        'btnBackground
        '
        Me.btnBackground.Location = New System.Drawing.Point(8, 80)
        Me.btnBackground.Name = "btnBackground"
        Me.btnBackground.Size = New System.Drawing.Size(16, 16)
        Me.btnBackground.TabIndex = 0
        Me.tip.SetToolTip(Me.btnBackground, "Set the color of the background on the display")
        '
        'pnlDisplay
        '
        Me.pnlDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlDisplay.Name = "pnlDisplay"
        Me.pnlDisplay.TabIndex = 3
        Me.pnlDisplay.Visible = False
        '
        'pnlOuter
        '
        Me.pnlOuter.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlOuter.AutoScroll = True
        Me.pnlOuter.Controls.Add(Me.pnlDisplay)
        Me.pnlOuter.Location = New System.Drawing.Point(0, 0)
        Me.pnlOuter.Name = "pnlOuter"
        Me.pnlOuter.Size = New System.Drawing.Size(632, 408)
        Me.pnlOuter.TabIndex = 4
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(632, 433)
        Me.Controls.Add(Me.pnlColors)
        Me.Controls.Add(Me.sbMain)
        Me.Controls.Add(Me.pnlOuter)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmMain"
        Me.Text = "Irregular Grid Painter"
        Me.pnlColors.ResumeLayout(False)
        Me.pnlOuter.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    WithEvents myPattern As Zlosk.Patterns.Pattern
    Private mcolCurrent As Color
    Private mstrSave, mstrFileName As String
    Private mblnMouseDown As Boolean = False, mblnDirty As Boolean
    WithEvents d2m As New Dir2Mnu()

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        myPattern = New Zlosk.Patterns.Pattern()
        Me.myPattern.mctrlClient = Me
        Dim strIniFile As String = Application.StartupPath & "\igp.ini"
        Dim ini As New INIReader(strIniFile)
        Me.pnlOuter.Width = Me.ClientSize.Width
        Me.pnlOuter.Height = Me.ClientSize.Height - Me.sbMain.Height

        ' Determine location of Patterns directory
        Dim strPatternPath As String
        strPatternPath = ini.ReadString("General", "PatternsDirectory", "")
        ' Per http://www.codeproject.com/csharp/effectivecspart1.asp, 
        ' don't compare strings; compare length instead.
        If strPatternPath.Length = 0 Then
            MsgBox("A directory of Patterns has not been specified. In the IGP.INI file, add the key 'PatternsDirectory=' to the [General] section.", MsgBoxStyle.Exclamation, "Error: No Patterns")
            Me.Close()
            Exit Sub
        Else
            strPatternPath = System.IO.Path.GetDirectoryName(strPatternPath)
            If System.IO.Directory.Exists(strPatternPath) Then
                ' Add patterns to Patterns menu
                d2m.MenuBase = mnuPatterns
                d2m.DirName = strPatternPath
                d2m.CreateMenus()
            Else
                MsgBox("The 'Patterns' directory does not exist. When extracting IGP from the compressed zip file, the directory structure must be extracted as well. For example, if using WinZip 8.0, the 'Use folder names' checkbox should be checked before extracting. ", MsgBoxStyle.Exclamation, "Error: Missing Subdirectory")
                Me.Close()
                Exit Sub
            End If
        End If

        ' Load Current Pattern
        Dim strIniName As String = ini.ReadString("General", "CurrentPattern", "")
        If strPatternPath.Length = 0 Then
            MsgBox("A current Pattern has not been specified. In the IGP.INI file, add the key 'CurrentPattern=' to the [General] section.", MsgBoxStyle.Exclamation, "Error: No Patterns")
            Exit Sub
        End If
        Dim FullPath As String
        FullPath = strPatternPath & "\" & strIniName & "\" & System.IO.Path.GetFileNameWithoutExtension(strIniName) & ".ini"
        ' Per http://www.codeproject.com/csharp/effectivecspart1.asp, 
        ' don't compare strings; compare length instead.
        If System.IO.Path.GetPathRoot(FullPath).Length = 0 Then
            FullPath = Application.StartupPath & "\" & FullPath
        End If
        Me.myPattern.LoadIni(FullPath)

        ' Set Open/Save starting directory
        Dim strTemp As String = ini.ReadString("General", "OpenSaveDirectory", "img")
        FullPath = System.IO.Path.GetDirectoryName(strTemp)
        ' Per http://www.codeproject.com/csharp/effectivecspart1.asp, 
        ' don't compare strings; compare length instead.
        If System.IO.Path.GetPathRoot(FullPath).Length = 0 Then
            FullPath = Application.StartupPath & "\" & FullPath
        End If
        ofd.InitialDirectory = FullPath
        sfd.InitialDirectory = FullPath

        Me.mcolCurrent = StringToColor(ini.ReadString("General", "PaintColor", "255,0,0"))
        Me.myPattern.ForeColor = StringToColor(ini.ReadString("General", "ForeColor", "32,32,32"))
        Me.myPattern.BackColor = StringToColor(ini.ReadString("General", "BackColor", "96,96,96"))
        Me.myPattern.BaseColor = StringToColor(ini.ReadString("General", "BaseColor", "192,192,192"))
        Me.btnLeft.BackColor = Me.mcolCurrent
        Me.btnRight.BackColor = Me.myPattern.BaseColor
        Me.btnOutlines.BackColor = Me.myPattern.ForeColor
        Me.btnBackground.BackColor = Me.myPattern.BackColor
    End Sub
    Private Sub Form1_Invalidated(ByVal sender As Object, ByVal e As System.Windows.Forms.InvalidateEventArgs) Handles myPattern.Invalidated
        If Me.myPattern Is Nothing Then Exit Sub
        If Me.myPattern.RenderedImage Is Nothing Then Exit Sub
        Me.CreateGraphics.DrawImage(Me.myPattern.RenderedImage, e.InvalidRect, e.InvalidRect, System.Drawing.GraphicsUnit.Pixel)
    End Sub
    Private Sub Form1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint, pnlDisplay.Paint
        '    If Me.myPattern Is Nothing Then Exit Sub
        '    If Me.myPattern.RenderedImage Is Nothing Then Exit Sub
        '    e.Graphics.DrawImage(Me.myPattern.RenderedImage, 0, 0)
        Me.RefreshPanel()
        Me.SetPnlColorsLocation()
    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        ' pnlColors placement needs to be handled outside of the "Anchor" property
        ' to bypass the ScrollBars
        SetPnlColorsLocation()
        'Me.Form1_Paint(Me, New PaintEventArgs(Me.CreateGraphics, New Rectangle(0, 0, 1, 1)))
    End Sub
    Private Sub SetPnlColorsLocation()
        Dim intLeft, intTop As Integer
        intLeft = Me.pnlOuter.ClientSize.Width - Me.pnlColors.Width - 3
        If intLeft < 3 Then intLeft = 3
        intTop = 3
        Me.pnlColors.Left = intLeft
        Me.pnlColors.Top = intTop
    End Sub

    Private Sub Form1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Click, pnlDisplay.Click
        ' Translate click location from screen coords to form coords
        Dim ptLocalMousePosition As Point
        Dim pc As Zlosk.Patterns.Pattern.PatternCoordinate

        ptLocalMousePosition = Me.PointToClient(Cursor.Position)
        ptLocalMousePosition.X -= Me.pnlDisplay.Left
        ptLocalMousePosition.Y -= Me.pnlDisplay.Top

        ' Translate click location from form coords to pattern coords
        pc = Me.myPattern.PointToPatternCoordinate(ptLocalMousePosition)
        If pc.x = -1 Then Exit Sub
        If Me.myPattern.GetPixel(pc).ToArgb <> Me.mcolCurrent.ToArgb Then
            Me.myPattern.SetPixel(pc, Me.mcolCurrent)
            Me.mblnDirty = True
            Me.RefreshPanel()
        End If
    End Sub
    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown, pnlDisplay.MouseDown
        mblnMouseDown = True
        If e.Button = MouseButtons.Left Then
            Me.mcolCurrent = Me.btnLeft.BackColor
        ElseIf e.Button = MouseButtons.Right Then
            Me.mcolCurrent = Me.btnRight.BackColor
        End If
    End Sub
    Private Sub Form1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp, pnlDisplay.MouseUp
        mblnMouseDown = False
    End Sub
    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove, pnlDisplay.MouseMove
        ' Translate click location from screen coords to form coords
        Dim ptLocalMousePosition As Point
        Dim pc As Zlosk.Patterns.Pattern.PatternCoordinate
        ptLocalMousePosition = Me.pnlDisplay.PointToClient(Cursor.Position)
        ' Translate click location from form coords to pattern coords
        pc = Me.myPattern.PointToPatternCoordinate(ptLocalMousePosition)
        If pc.x = -1 Then
            Me.sbMain.Text = ""
        Else
            Me.sbMain.Text = "(" & pc.x.ToString & "," & pc.y.ToString & "," & pc.Index.ToString & ")"
        End If
        If Me.mblnMouseDown Then Form1_Click(sender, e)
    End Sub

#Region " File menu "
    Private Sub mnuFileNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileNew.Click
        If mblnDirty Then
            Dim Response As MsgBoxResult = MsgBox("The document has changed. Do you want to save the changes?", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Exclamation, Me.Text)
            Select Case Response
                Case MsgBoxResult.Yes
                    mnuFileSave_Click(sender, e)
                Case MsgBoxResult.Cancel
                    Exit Sub
            End Select
        End If
        Dim frm As New frmFileNew()
        frm.ShowDialog()
        If frm.OK Then
            Me.myPattern.NewBaseImage(frm.ImgWidth, frm.ImgHeight)
            Me.EnableMenuSaveGroup(True)
            Me.RefreshPanel()
            mblnDirty = True
        End If
    End Sub
    Private Sub mnuFileOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileOpen.Click
        If mblnDirty Then
            Dim Response As MsgBoxResult = MsgBox("The document has changed. Do you want to save the changes?", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Exclamation, Me.Text)
            Select Case Response
                Case MsgBoxResult.Yes
                    mnuFileSave_Click(sender, e)
                Case MsgBoxResult.Cancel
                    Exit Sub
            End Select
        End If
        ofd.Filter = "All files (*.*)|*.*|Bitmap (*.bmp)|*.bmp|PC Stitch Pattern (*.pat)|*.pat|Portable Network Graphic (*.png)|*.png|TIFF (*.tif)|*.tif"
        ofd.ShowDialog()
    End Sub
    Private Sub ofd_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ofd.FileOk
        Me.mstrFileName = ofd.FileName
        ' Load new bitmap
        Select Case UCase(System.IO.Path.GetExtension(ofd.FileName))
            Case Is = ".BMP", ".GIF", ".PNG", ".TIF"
                Me.myPattern.BaseImage = New Bitmap(Image.FromFile(ofd.FileName))
            Case Is = ".PAT"
                Dim pcs As New Zlosk.PCStitch()
                pcs.FromFile(ofd.FileName)
                If pcs.Bitmap Is Nothing Then Exit Sub
                pcs.Bitmap.Save("~temp.bmp", System.Drawing.Imaging.ImageFormat.Bmp)
                'Dim img As Image = pcs.Bitmap
                'Me.myPattern.BaseImage = img
                Me.myPattern.BaseImage = New Bitmap(Image.FromFile("~temp.bmp"))
        End Select
#If DEBUG Then
        Dim sw As New Org.Mentalis.Utilities.StopWatch
        sw.Reset()
        Dim i As Int16
        For i = 1 To 5
#End If
            Me.myPattern.BuildImage()
#If DEBUG Then
        Next i
        MsgBox((sw.Peek / 5).ToString)
#End If
        Me.EnableMenuSaveGroup(True)
        Me.mblnDirty = False
        Me.RefreshPanel()
    End Sub
    Private Sub EnableMenuSaveGroup(ByVal bln As Boolean)
        Me.mnuFileClose.Enabled = bln
        Me.mnuFileSave.Enabled = bln
        Me.mnuFileSaveAs.Enabled = bln
        Me.mnuFileSavePicture.Enabled = bln
        Me.pnlDisplay.Visible = bln
        'Me.pnlOuter.Visible = bln
    End Sub
    Private Sub mnuFileClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileClose.Click
        If mblnDirty Then
            Dim Response As MsgBoxResult = MsgBox("The document has changed. Do you want to save the changes?", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Exclamation, Me.Text)
            Select Case Response
                Case MsgBoxResult.Yes
                    mnuFileSave_Click(sender, e)
                Case MsgBoxResult.Cancel
                    Exit Sub
            End Select
        End If
        Me.myPattern.BaseImage = Nothing
        EnableMenuSaveGroup(False)
        mblnDirty = False
        'Me.RefreshPanel()
    End Sub
    Private Sub mnuFileSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSave.Click
        mstrSave = "Save"
        ' Per http://www.codeproject.com/csharp/effectivecspart1.asp, 
        ' don't compare strings; compare length instead.
        If mstrFileName.Length = 0 Then
            mnuFileSaveAs_Click(sender, e)
            Exit Sub
        End If
        sfd.FileName = Me.mstrFileName
        sfd_FileOk(sender, New System.ComponentModel.CancelEventArgs)
    End Sub
    Private Sub mnuFileSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSaveAs.Click
        mstrSave = "Save"
        'Jpeg is not allowed because it is stupid for pics that this program will output
        sfd.Filter = "Bitmap (*.bmp)|*.bmp|Portable Network Graphic (*.png)|*.png|TIFF (*.tif)|*.tif"
        sfd.ShowDialog()
    End Sub
    Private Sub mnuFileSavePicture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSavePicture.Click
        mstrSave = "SavePicture"
        sfd.Filter = "Bitmap (*.bmp)|*.bmp|Portable Network Graphic (*.png)|*.png|TIFF (*.tif)|*.tif"
        sfd.ShowDialog()
    End Sub
    Private Sub sfd_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles sfd.FileOk
        Dim imgFormat As System.Drawing.Imaging.ImageFormat

        Select Case sfd.FilterIndex
            Case 1
                imgFormat = System.Drawing.Imaging.ImageFormat.Bmp
            Case 2
                imgFormat = System.Drawing.Imaging.ImageFormat.Png
            Case 3
                imgFormat = System.Drawing.Imaging.ImageFormat.Tiff
            Case Else
                MsgBox("IGP is not able to use that image format. Please choose an image in a PNG, TIF, or BMP format.")
                Exit Sub
        End Select
        Select Case mstrSave
            Case "SavePicture"
                Me.myPattern.RenderedImage.Save(sfd.FileName, imgFormat)
            Case "Save"
                Me.myPattern.BaseImage.Save(sfd.FileName, imgFormat)
        End Select
        Me.mblnDirty = False
    End Sub
    Private Sub mnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
        Me.Close()
    End Sub
#End Region

#Region " Patterns menu "
    Private Sub d2m_Click(ByVal sender As MenuItem, ByVal Tag As Dir2Mnu.d2mData) Handles d2m.Click
        Me.myPattern.LoadIni(Tag.FullPath)
        Me.myPattern.BuildImage()
        Refresh()
    End Sub
#End Region

#Region " Colors menu "
    Private Sub mnuColorsLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuColorsLeft.Click, lblLeft.Click
        btnColorsLeftRight_Click(Me.btnLeft, e)
    End Sub
    Private Sub mnuColorsRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuColorsRight.Click, lblRight.Click
        btnColorsLeftRight_Click(Me.btnRight, e)
    End Sub
    Private Sub mnuColorsBackground_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuColorsBackground.Click, lblBackground.Click
        btnColorsBackground_Click(Me.btnBackground, e)
    End Sub
    Private Sub mnuColorsOutlines_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuColorsOutlines.Click, lblOutlines.Click
        btnColorsOutlines_Click(Me.btnOutlines, e)
    End Sub
    Private Sub mnuColorsCountColors_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuColorsCountColors.Click
        btnColorCount_Click(Me.btnColorCount, e)
    End Sub
    Private Sub btnColorsLeftRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLeft.Click, btnRight.Click
        Dim colTmp As Color = CType(sender, Button).BackColor
        cd.Color = colTmp
        cd.ShowDialog()
        If cd.Color.ToArgb <> colTmp.ToArgb Then
            If True Then
                Me.mcolCurrent = cd.Color
                Me.Refresh()
                CType(sender, Button).BackColor = cd.Color
            End If
        End If
    End Sub
    Private Sub btnColorsBackground_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBackground.Click
        Dim colTmp As Color = CType(sender, Button).BackColor
        cd.Color = colTmp
        cd.ShowDialog()
        If cd.Color.ToArgb <> colTmp.ToArgb Then
            CType(sender, Button).BackColor = cd.Color
            Me.myPattern.BackColor = cd.Color
            Me.Refresh()
        End If
    End Sub
    Private Sub btnColorsOutlines_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOutlines.Click
        Dim colTmp As Color = CType(sender, Button).BackColor
        cd.Color = colTmp
        cd.ShowDialog()
        If cd.Color.ToArgb <> colTmp.ToArgb Then
            CType(sender, Button).BackColor = cd.Color
            Me.myPattern.ForeColor = cd.Color
            Me.Refresh()
        End If
    End Sub
    Private Sub btnColorCount_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnColorCount.Click
        If Not myPattern.BaseImage Is Nothing Then
            Dim frm As New frmColorCount
            frm.bmpColorPattern = New Bitmap(myPattern.BaseImage)
            frm.ShowDialog()
        End If
    End Sub
#End Region

    Function StringToColor(ByVal str As String) As Color
        Dim r, g, b As Integer
        Dim strArray() As String
        strArray = str.Split(System.Convert.ToChar(","))
        r = System.Convert.ToInt16(strArray(0))
        g = System.Convert.ToInt16(strArray(1))
        b = System.Convert.ToInt16(strArray(2))
        Return Color.FromArgb(r, g, b)
    End Function

    Private Sub ToggleColorPanel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuWindowColorPanel.Click
        Me.pnlColors.Visible = Not Me.pnlColors.Visible
        If Me.pnlColors.Visible Then
            Me.mnuWindowColorPanel.Text = "Hide &Color Panel"
        Else
            Me.mnuWindowColorPanel.Text = "Show &Color Panel"
        End If
    End Sub

    Private Sub mnuHelpAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpAbout.Click
        ' Per http://www.codeproject.com/csharp/effectivecspart1.asp - Item 2,
        ' use StringBuilder when assembling a long string instead of 
        ' strTemp += "blah"
        Dim sbAbout As New System.Text.StringBuilder("Irregular Grid Painter")
        sbAbout.Append(vbCrLf)
        sbAbout.Append("Version 0.6.0.5")
        sbAbout.Append(vbCrLf)
        sbAbout.Append("Copyright ©2009 zlosk.com")
        sbAbout.Append(vbCrLf)
        sbAbout.Append(vbCrLf)
        sbAbout.Append("Send comments, criticisms and requests")
        sbAbout.Append(vbCrLf)
        sbAbout.Append("for new features to programmer@zlosk.com")
        MsgBox(sbAbout.ToString, MsgBoxStyle.OKOnly, "Irregular Grid Painter")
    End Sub

    Private Sub RefreshPanel()
        If Me.myPattern Is Nothing Then Exit Sub
        If Me.myPattern.RenderedImage Is Nothing Then Exit Sub
        Me.pnlDisplay.Size = Me.myPattern.RenderedImage.Size
        Me.pnlDisplay.CreateGraphics.DrawImage(Me.myPattern.RenderedImage, 0, 0)
    End Sub

    Private Sub frmMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If mblnDirty Then
            Dim Response As MsgBoxResult = MsgBox("The document has changed. Do you want to save the changes?", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Exclamation, Me.Text)
            Select Case Response
                Case MsgBoxResult.Yes
                    mnuFileSave_Click(sender, e)
                Case MsgBoxResult.Cancel
                    e.Cancel = True
            End Select
        End If
    End Sub

    Private Sub pnlOuter_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles pnlOuter.Paint

    End Sub
End Class
