'
'    DirectoryDialog class for VB.NET
'		Version: 1.0		Date: 2002/04/24
'
'
'    Copyright © 2002, The KPD-Team
'    All rights reserved.
'    http://www.mentalis.org/
'
'  Redistribution and use in source and binary forms, with or without
'  modification, are permitted provided that the following conditions
'  are met:
'
'    - Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer. 
'
'    - Neither the name of the KPD-Team, nor the names of its contributors
'       may be used to endorse or promote products derived from this
'       software without specific prior written permission. 
'
'  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'  "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'  LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
'  FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL
'  THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT,
'  INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
'  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
'  SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
'  HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
'  STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
'  ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED
'  OF THE POSSIBILITY OF SUCH DAMAGE.
'

Imports System
Imports System.Text
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

'<summary>The <strong>Org.Mentalis.Utilities.DirectoryDialog</strong> namespace provides a managed implementation of the SHBrowseForFolder function that creates a directory dialog.</summary>
Namespace Org.Mentalis.Utilities.DirectoryDialog
    '///<summary>Contains parameters for the SHBrowseForFolder function and receives information about the folder selected by the user.</summary>
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> _
    Friend Structure BROWSEINFO
        '///<summary>Handle to the owner window for the dialog box.</summary>
        Public hWndOwner As IntPtr
        '///<summary>Address of an ITEMIDLIST structure specifying the location of the root folder from which to browse. Only the specified folder and its subfolders appear in the dialog box. This member can be IntPtr.Zero; in that case, the namespace root (the desktop folder) is used.</summary>
        Public pIDLRoot As IntPtr
        '///<summary>Address of a buffer to receive the display name of the folder selected by the user. The size of this buffer is assumed to be MAX_PATH bytes.</summary>
        Public pszDisplayName As String
        '///<summary>Address of a null-terminated string that is displayed above the tree view control in the dialog box. This string can be used to specify instructions to the user.</summary>
        Public lpszTitle As String
        '///<summary>Flags specifying the options for the dialog box.</summary>
        Public ulFlags As Integer
        '///<summary>Address of an application-defined function that the dialog box calls when an event occurs. This member can be IntPtr.Zero.</summary>
        Public lpfnCallback As IntPtr
        '///<summary>Application-defined value that the dialog box passes to the callback function, if one is specified.</summary>
        Public lParam As Integer
        '///<summary>Variable to receive the image associated with the selected folder. The image is specified as an index to the system image list.</summary>
        Public iImage As Integer
    End Structure

    '///<summary>Specifies the type of items to browse for.</summary>
    Public Enum BrowseForTypes
        '///<summary>Browse for computers.</summary>
        Computers = &H1000
        '///<summary>Browse for directories.</summary>
        Directories = &H1
        '///<summary>Browse for files and directories.</summary>
        '///<remarks>Only works on shell version 4.71 and higher!</remarks>
        FilesAndDirectories = &H4000 ' Version 4.71
        '///<summary>Browse for file system ancestors (an ancestor is a subfolder that is beneath the root folder in the namespace hierarchy.).</summary>
        FileSystemAncestors = &H8
    End Enum

    '///<summary>Implements the SHBrowseForFolder function to show a common folder dialog.</summary>
    Public Class DirectoryDialog
        '///<summary>Initializes a new DirectoryDialog instance.</summary>
        Public Sub New()
        End Sub
        '///<summary>Shows the common folder dialog.</summary>
        '///<param name="hWndOwner">The owner of the folder dialog.</param>
        '///<returns>True when successful, false otherwise.</returns>
        Protected Function RunDialog(ByVal hWndOwner As IntPtr) As Boolean
            Dim udtBI As New BROWSEINFO()
            Dim lpIDList As IntPtr
            udtBI.pIDLRoot = IntPtr.Zero
            udtBI.lpfnCallback = IntPtr.Zero
            ' set the owner of the window
            udtBI.hWndOwner = hWndOwner
            ' set the title of the window
            udtBI.lpszTitle = Title
            ' set the flags of the dialog
            udtBI.ulFlags = CType(BrowseFor, Integer)
            ' create string buffer for display name
            Dim buffer As New StringBuilder(MAX_PATH)
            buffer.Length = MAX_PATH
            udtBI.pszDisplayName = buffer.ToString()
            ' show the 'Browse for folder' dialog
            lpIDList = SHBrowseForFolder(udtBI)
            ' examine the result
            If Not lpIDList.Equals(IntPtr.Zero) Then
                If BrowseFor = BrowseForTypes.Computers Then
                    m_Selected = udtBI.pszDisplayName.Trim()
                Else
                    Dim path As New StringBuilder(MAX_PATH)
                    ' get the path from the IDList
                    SHGetPathFromIDList(lpIDList, path)
                    m_Selected = path.ToString()
                End If
                ' free the block of memory
                CoTaskMemFree(lpIDList)
                Return True
            Else
                ' user pressed cancel
                Return False
            End If
        End Function
        '///<summary>Shows the folder dialog.</summary>
        '///<returns>DialogResult.OK when successful, DialogResult.Cancel otherwise.</returns>
        Public Function ShowDialog() As DialogResult
            Return ShowDialog(Nothing)
        End Function
        '///<summary>Shows the folder dialog.</summary>
        '///<param name="owner">The owner of the folder dialog.</param>
        '///<returns>DialogResult.OK when successful, DialogResult.Cancel otherwise.</returns>
        Public Function ShowDialog(ByVal owner As IWin32Window) As DialogResult
            Dim handle As IntPtr
            If Not owner Is Nothing Then
                handle = owner.Handle
            Else
                handle = IntPtr.Zero
            End If
            If RunDialog(handle) Then
                Return DialogResult.OK
            Else
                Return DialogResult.Cancel
            End If
        End Function
        '///<summary>Gets or sets the title of the dialog.</summary>
        '///<value>A String representing the title of the dialog.</value>
        '///<exceptions cref="ArgumentNullException">The specified value is null (VB.NET: Nothing)</exceptions>
        Public Property Title() As String
            Get
                Return m_Title
            End Get
            Set(ByVal Value As String)
                If Value Is Nothing Then Throw New ArgumentNullException()
                m_Title = Value
            End Set
        End Property
        '///<summary>Gets the selected item.</summary>
        '///<value>A String representing the selected item.</value>
        Public ReadOnly Property Selected() As String
            Get
                Return m_Selected
            End Get
        End Property
        '///<summary>Gets or sets the type of items to browse for.</summary>
        '///<value>One of the BrowseForTypes values.</value>
        Public Property BrowseFor() As BrowseForTypes
            Get
                Return m_BrowseFor
            End Get
            Set(ByVal Value As BrowseForTypes)
                m_BrowseFor = Value
            End Set
        End Property
        ' private declares
        '///<summary>Frees a block of task memory previously allocated through a call to the CoTaskMemAlloc or CoTaskMemRealloc function.</summary>
        '///<param name="hMem">Pointer to the memory block to be freed.</param>
        Private Declare Auto Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As IntPtr)
        '///<summary>The lstrcat function appends one string to another.</summary>
        '///<param name="lpString1">Pointer to a null-terminated string. The buffer must be large enough to contain both strings.</param>
        '///<param name="lpString2">Pointer to the null-terminated string to be appended to the string specified in the lpString1 parameter.</param>
        '///<returns>If the function succeeds, the return value is a pointer to the buffer.<br>If the function fails, the return value is IntPtr.Zero.</br></returns>
        Private Declare Ansi Function lstrcat Lib "kernel32.dll" (ByVal lpString1 As String, ByVal lpString2 As String) As IntPtr
        '///<summary>Displays a dialog box that enables the user to select a shell folder.</summary>
        '///<param name="lpbi">Address of a BROWSEINFO structure that contains information used to display the dialog box.</param>
        '///<returns>Returns the address of an item identifier list that specifies the location of the selected folder relative to the root of the namespace. If the user chooses the Cancel button in the dialog box, the return value is IntPtr.Zero.</returns>
        Private Declare Ansi Function SHBrowseForFolder Lib "shell32.dll" (ByRef lpbi As BROWSEINFO) As IntPtr
        '///<summary>Converts an item identifier list to a file system path.</summary>
        '///<param name="pidList">Address of an item identifier list that specifies a file or directory location relative to the root of the namespace (the desktop).</param>
        '///<param name="lpBuffer">Address of a buffer to receive the file system path. This buffer must be at least MAX_PATH characters in size.</param>
        '///<returns>Returns a nonzero value if successful, or zero otherwise.</returns>
        Private Declare Ansi Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As IntPtr, ByVal lpBuffer As StringBuilder) As Integer
        ' private variables
        '///<summary>Specifies the maximum number of characters in a pathname.</summary>
        '///<remarks>The value of this integer is 260.</remarks>
        Private Const MAX_PATH As Integer = 260
        '///<summary>Holds a BrowseForTypes value that indicates what type of items to browse for.</summary>
        Private m_BrowseFor As BrowseForTypes = BrowseForTypes.Directories
        '///<summary>Holds the string with the title of the directory dialog.</summary>
        Private m_Title As String = ""
        '///<summary>The a string with the selected item.</summary>
        Private m_Selected As String = ""
    End Class
End Namespace