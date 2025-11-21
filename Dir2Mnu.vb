'Class: Dir2Mnu
'Purpose: To create a submenu structure from a directory structure

'Programmer: Paul Werstler
'Web site: http://www.zlosk.com
'Date: 4-4-03

#Region " A code example showing the implemetation "
'Public Class frmTest
'    Inherits System.Windows.Forms.Form
'    Friend WithEvents md2m As Dir2Mnu

'    Private Sub frmTest_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'        md2m = New Dir2Mnu(mnuMainMenu, "c:\temp")
'        Me.SuspendLayout()
'        md2m.CreateMenus()
'        Me.ResumeLayout(False)
'    End Sub

'    Private Sub Dir2Mnu_Click(ByVal mnu As System.Windows.Forms.MenuItem, ByVal str As String) Handles md2m.Click
'        MsgBox(str)
'    End Sub
'End Class
#End Region

Public Class Dir2Mnu

    Structure d2mData
        Dim FullPath As String
        Dim MenuPath As String
    End Structure


#Region " Member variables "
    Dim mmnuBase As System.Windows.Forms.MenuItem
    Dim mstrDirName, mstrDirResult As String
    Dim mintCt As Integer = 0
    Dim mTags As New Collection()
#End Region

#Region " Constructors "
    Public Sub New()

    End Sub
    Public Sub New(ByVal mnuBase As System.Windows.Forms.MenuItem, ByVal strDirName As String)
        MenuBase = mnuBase
        DirName = strDirName
    End Sub
#End Region

#Region " Events "
    Public Event Click(ByVal mnuSender As System.Windows.Forms.MenuItem, ByVal Tag As d2mData)
#End Region

#Region " Event Handlers "
    Private Sub Dir2Mnu_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim mnu As System.Windows.Forms.MenuItem
        mnu = CType(sender, System.Windows.Forms.MenuItem)
        RaiseEvent Click(mnu, GetTag(mnu))
    End Sub
#End Region

#Region " Methods "
    Public Sub CreateMenus()
        CreateMenus(Me.mmnuBase, Me.mstrDirName)
    End Sub
    Private Sub CreateMenus(ByVal mnuBase As System.Windows.Forms.MenuItem, ByVal strDirName As String)
        Dim oFS As New System.IO.DirectoryInfo(strDirName & "\")
        Dim oDir As System.IO.DirectoryInfo
        'Dim oFile As System.IO.FileInfo
        If Not oFS.Exists Then Exit Sub
        Dim ini As New INIReader("")
        Dim Tag As d2mData
        For Each oDir In oFS.GetDirectories()
            Dim mnu As New System.Windows.Forms.MenuItem()
            ini.Filename = oDir.FullName & "\" & oDir.Name & ".ini"
            mnu.Text = ini.ReadString("General", "Description", oDir.Name)
            mnuBase.MenuItems.Add(mnu)
            Tag.FullPath = ini.Filename
            Tag.MenuPath = GetMenuPath(mnu)
            mTags.Add(Tag)
            AddHandler mnu.Click, AddressOf Dir2Mnu_Click
            CreateMenus(mnu, strDirName & "\" & oDir.Name)
        Next
    End Sub
    Public Sub RemoveMenus()
        RemoveMenu(mmnuBase)
    End Sub
    Private Sub RemoveMenu(ByVal mnu As System.Windows.Forms.MenuItem)
        If mnu.IsParent Then
            RemoveMenu(mnu.MenuItems(0))
            RemoveMenu(mnu)
        Else
            If mnu.ToString <> mmnuBase.ToString Then
                mnu.Dispose()
            End If
        End If
    End Sub
#End Region

#Region " Properties "
    Public Property MenuBase() As System.Windows.Forms.MenuItem
        Get
            ' The Get property procedure is called when the value
            ' of a property is retrieved. 
            Return mmnuBase
        End Get
        Set(ByVal Value As System.Windows.Forms.MenuItem)
            ' The Set property procedure is called when the value 
            ' of a property is modified. 
            ' The value to be assigned is passed in the argument to Set. 
            mmnuBase = Value
        End Set
    End Property
    Public Property DirName() As String
        Get
            ' The Get property procedure is called when the value
            ' of a property is retrieved. 
            Return mstrDirName
        End Get
        Set(ByVal Value As String)
            ' The Set property procedure is called when the value 
            ' of a property is modified. 
            ' The value to be assigned is passed in the argument to Set. 
            Do Until Value.Substring(Value.Length - 1) <> "\"
                Value = Value.Substring(0, Value.Length - 1)
            Loop
            mstrDirName = Value
        End Set
    End Property
#End Region

#Region " Miscellaneous functions / subroutines "
    Private Function GetMenuPath(ByVal mnu As System.Windows.Forms.MenuItem) As String
        Dim str As String = ""
        Do Until mnu.Text = mmnuBase.Text
            str = GetAmpersandlessText(mnu) & "\" & str
            mnu = mnu.Parent
        Loop
        Return str
    End Function
    Private Function GetAmpersandlessText(ByVal mnu As System.Windows.Forms.MenuItem) As String
        Dim c As Char
        Dim str As String
        c = mnu.Mnemonic
        str = mnu.Text
        If AscW(c) > 0 Then
            Dim i As Integer
            i = InStr(UCase(str), "&" & UCase(CStr(c)))
            str = str.Remove(i - 1, 1)
        End If
        Return str
    End Function
    Private Function GetTag(ByVal mnu As System.Windows.Forms.MenuItem) As d2mData
        Dim mnuPath As String
        Dim Tag As d2mData
        mnuPath = GetMenuPath(mnu)
        For Each Tag In mTags
            If mnuPath = Tag.MenuPath Then
                Return Tag
                Exit For
            End If
        Next
        Return Nothing
    End Function
#End Region

End Class
