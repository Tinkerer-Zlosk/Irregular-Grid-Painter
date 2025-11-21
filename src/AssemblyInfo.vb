'Copyright (C) 2002 Microsoft Corporation
'All rights reserved.
'THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
'EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
'MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.

'Requires the Trial or Release version of Visual Studio .NET Professional (or greater).

' Version History
'
' 0.6.0.5 2009.09.01 Added validation code to Color Counter so choosing IGP won't
'                    crash when choosing lines that don't have colors. Also fixed
'                    Change Description InputBox so Cancel button works properly.
' 0.6.0.4 2005.05.02 Changed a couple of objects from String to StringBuilder
'                    Added a new error check to see if folder structure was
'                    extracted when unzipping the zip file (and inform the
'                    user if it was not).
' 0.6.0.3 2005.04.13 Fixed About dialog to show correct version
'                    Fixed bug in BuildImage that caused IGP to crash on
'                    some images.
' 0.6.0.2 2005.04.12 Did some optimization in igpPattern.vb - 
'                    Patterned images build about 4 times faster
' 0.6.0.1 2003.01.11 Fixed error in PC Stitch 7 import
' 0.6.0.0 2003.01.09 Added "Save as text..." to Color Counter (request by SirGuido)
'                    Added PC Stitch 7 import (request by Charlie AZ)
' 0.5.0.0 2003.12.31 Added Color Counter (created by Dunedon, modified by Zlosk)

Option Strict On

Imports System.Reflection

<Assembly: AssemblyTitle("Irregular Grid Painter")> 
<Assembly: AssemblyDescription("The Irregular Grid Painter allows painting on a non-uniform grid of tiles.")> 
<Assembly: AssemblyCompany("zlosk.com")> 
<Assembly: AssemblyProduct("zlosk.com: Irregular Grid Painter")> 
<Assembly: AssemblyCopyright("Copyright © 2009 zlosk.com.  All rights reserved.")> 
<Assembly: CLSCompliant(True)> 

' Version Definition
'
' Version numbers consist of two to four components: major, minor, build, and revision. 
' Components major and minor are required. Components build and revision are optional, but 
' the revision component is only optional if the build component is not defined. All 
' defined components must be decimal integers greater than or equal to 0. Metadata 
' restricts the major, minor, build, and revision components to a maximum value of 
' MaxValue - 1.
'
' The format of the version number is as follows. Optional components are shown in square 
' brackets ('[' and ']'):
' major.minor[.build[.revision]]

' The components are used by convention as follows:
' * Assemblies with the same name but different major versions are not interchangeable. 
'   This would be appropriate, for example, for a major rewrite of a product where backward 
'   compatibility cannot be assumed. 
' * If the name and major number on two assemblies are the same, but the minor number is 
'   different, this indicates significant enhancement with the intention of backward 
'   compatibility. This would be appropriate, for example, on a point release of a product
'   or a fully backward compatible new version of a product. 
' * A difference in build number represents a recompilation of the same source. This would
'   be appropriate because of processor, platform, or compiler changes. 
' * Assemblies with the same name, major, and minor version numbers but different revisions 
'   are intended to be fully interchangeable. This would be appropriate to fix a security 
'   hole in a previously released assembly. 
'
' Subsequent versions of an assembly that differ only by build or revision numbers are 
' considered to be Quick Fix Engineering (QFE) updates of the prior version. If necessary, 
' the build and revision numbers can be honored by changing the version policy in the
' configuration.

<Assembly: AssemblyVersion("0.6.0.5")> 

#Region " Helper Class to Get Information for the About form. "
' This class uses the System.Reflection.Assembly class to
' access assembly meta-data
' This class is not a normal feature of AssemblyInfo.vb
Public Class AssemblyInfo
    ' Used by Helper Functions to access information from Assembly Attributes
    Private myType As Type

    Public Sub New()
        myType = GetType(frmMain)
    End Sub

    Public ReadOnly Property AsmName() As String
        Get
            Return myType.Assembly.GetName.Name.ToString()
        End Get
    End Property
    Public ReadOnly Property AsmFQName() As String
        Get
            Return myType.Assembly.GetName.FullName.ToString()
        End Get
    End Property
    Public ReadOnly Property CodeBase() As String
        Get
            Return myType.Assembly.CodeBase
        End Get
    End Property
    Public ReadOnly Property Copyright() As String
        Get
            Dim at As Type = GetType(AssemblyCopyrightAttribute)
            Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
            Dim ct As AssemblyCopyrightAttribute = CType(r(0), AssemblyCopyrightAttribute)
            Return ct.Copyright
        End Get
    End Property
    Public ReadOnly Property Company() As String
        Get
            Dim at As Type = GetType(AssemblyCompanyAttribute)
            Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
            Dim ct As AssemblyCompanyAttribute = CType(r(0), AssemblyCompanyAttribute)
            Return ct.Company
        End Get
    End Property
    Public ReadOnly Property Description() As String
        Get
            Dim at As Type = GetType(AssemblyDescriptionAttribute)
            Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
            Dim da As AssemblyDescriptionAttribute = CType(r(0), AssemblyDescriptionAttribute)
            Return da.Description
        End Get
    End Property
    Public ReadOnly Property Product() As String
        Get
            Dim at As Type = GetType(AssemblyProductAttribute)
            Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
            Dim pt As AssemblyProductAttribute = CType(r(0), AssemblyProductAttribute)
            Return pt.Product
        End Get
    End Property
    Public ReadOnly Property Title() As String
        Get
            Dim at As Type = GetType(AssemblyTitleAttribute)
            Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
            Dim ta As AssemblyTitleAttribute = CType(r(0), AssemblyTitleAttribute)
            Return ta.Title
        End Get
    End Property
    Public ReadOnly Property Version() As String
        Get
            Return myType.Assembly.GetName.Version.ToString()
        End Get
    End Property
End Class

#End Region
'Imports System.Reflection
'Imports System.Runtime.InteropServices

'' General Information about an assembly is controlled through the following 
'' set of attributes. Change these attribute values to modify the information
'' associated with an assembly.

'' Review the values of the assembly attributes

'<Assembly: AssemblyTitle("")> 
'<Assembly: AssemblyDescription("")> 
'<Assembly: AssemblyCompany("")> 
'<Assembly: AssemblyProduct("")> 
'<Assembly: AssemblyCopyright("")> 
'<Assembly: AssemblyTrademark("")> 
'<Assembly: CLSCompliant(True)> 

''The following GUID is for the ID of the typelib if this project is exposed to COM
'<Assembly: Guid("348A4E7F-9FA7-4523-852D-E1FB2C7EF5E4")> 

'' Version information for an assembly consists of the following four values:
''
''      Major Version
''      Minor Version 
''      Build Number
''      Revision
''
'' You can specify all the values or you can default the Build and Revision Numbers 
'' by using the '*' as shown below:

'<Assembly: AssemblyVersion("1.0.*")> 
