'Copyright (C) 2013 by Lee Millward

'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be included in
'all copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
'THE SOFTWARE.

''' <summary>
''' Compiler settings based on the any command line options.
''' </summary>
''' <remarks></remarks>
Public Class CompilerSettings

    ''' <summary>
    ''' Sets the setting values to the default
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        OutputTarget = OutputTarget.WinExe
        XMLDocumentationEnabled = False
        RemoveIntegerOverflowChecks = False
        DebugInfoGenerationEnabled = False
        OptionCompare = OptionCompareValues.Binary
        OptionExplicit = False
        OptionInfer = False
        OptionStrict = OptionStrictViolationValue.Ignore

        InteropAssemblies = New HashSet(Of String)
        ReferencedAssemblies = New HashSet(Of String)
        InputFiles = New List(Of SourceFile)
        Resources = New List(Of AssemblyResource)
        Defines = New List(Of KeyValuePair(Of String, String))
        GlobalImports = New HashSet(Of String)
        LibraryPaths = New List(Of String)

    End Sub

    ''' <summary>
    ''' The output filename.
    ''' </summary>
    ''' <remarks></remarks>
    Public OutputFile As String

    ''' <summary>
    ''' Type of output assembly to create.
    ''' </summary>
    ''' <remarks></remarks>
    Public OutputTarget As OutputTarget

    ''' <summary>
    ''' Controls whether XML documentation file is produced.
    ''' </summary>
    ''' <remarks></remarks>
    Public XMLDocumentationEnabled As Boolean

    ''' <summary>
    ''' Outputs XML documentation to the specified file.
    ''' </summary>
    ''' <remarks></remarks>
    Public XMLDocumentationOutputFile As String

    ''' <summary>
    ''' Whether to show help text or not.
    ''' </summary>
    ''' <remarks></remarks>
    Public ShowHelp As Boolean

    ''' <summary>
    ''' Interop assemblies to embed metadata from.
    ''' </summary>
    ''' <remarks></remarks>
    Public InteropAssemblies As HashSet(Of String)

    ''' <summary>
    ''' Set of files to reference metadata form.
    ''' </summary>
    ''' <remarks></remarks>
    Public ReferencedAssemblies As HashSet(Of String)

    ''' <summary>
    ''' List of source files to compile.
    ''' </summary>
    ''' <remarks></remarks>
    Public InputFiles As List(Of SourceFile)

    ''' <summary>
    ''' Set of resources to include.
    ''' </summary>
    ''' <remarks></remarks>
    Public Resources As List(Of AssemblyResource)

    ''' <summary>
    ''' Indicates whether the default manifest should not be embedded in the manifest section of the output PE file.
    ''' </summary>
    ''' <remarks></remarks>
    Public NoWin32Manifest As Boolean

    ''' <summary>
    ''' Win32 icon file for the default Win32 resources.
    ''' </summary>
    ''' <remarks></remarks>
    Public Win32IconFileName As String

    ''' <summary>
    ''' The file to embed in the manifest section of the output PE file.
    ''' </summary>
    ''' <remarks></remarks>
    Public Win32ManifestFileName As String

    ''' <summary>
    ''' Win32 resource file name.
    ''' </summary>
    ''' <remarks></remarks>
    Public Win32ResourceFileName As String

    ''' <summary>
    ''' Enables optimisations
    ''' </summary>
    ''' <remarks></remarks>
    Public Optimise As Boolean

    ''' <summary>
    ''' Removes integer overflow checks.
    ''' </summary>
    ''' <remarks></remarks>
    Public RemoveIntegerOverflowChecks As Boolean

    ''' <summary>
    ''' Whether debug info generation is enabled or not.
    ''' </summary>
    ''' <remarks></remarks>
    Public DebugInfoGenerationEnabled As Boolean

    ''' <summary>
    ''' Conditional compilation symbols.
    ''' </summary>
    ''' <remarks></remarks>
    Public Defines As List(Of KeyValuePair(Of String, String))

    ''' <summary>
    ''' Global imports list.
    ''' </summary>
    ''' <remarks></remarks>
    Public GlobalImports As HashSet(Of String)

    ''' <summary>
    ''' Whether explicit declaration of variables is required.
    ''' </summary>
    ''' <remarks></remarks>
    Public OptionExplicit As Boolean

    ''' <summary>
    ''' Allow type inference of variables.
    ''' </summary>
    ''' <remarks></remarks>
    Public OptionInfer As Boolean

    ''' <summary>
    ''' The root Namespace for all type declarations.
    ''' </summary>
    ''' <remarks></remarks>
    Public RootNamespace As String

    ''' <summary>
    ''' Enforce strict language semantics.
    ''' </summary>
    ''' <remarks></remarks>
    Public OptionStrict As OptionStrictViolationValue

    ''' <summary>
    ''' The method used for string comparisons.
    ''' </summary>
    ''' <remarks></remarks>
    Public OptionCompare As OptionCompareValues

    ''' <summary>
    ''' Whether to auto-include VBC.RSP file or not.
    ''' </summary>
    ''' <remarks></remarks>
    Public NoConfig As Boolean

    ''' <summary>
    ''' Do not display compiler copyright banner.
    ''' </summary>
    ''' <remarks></remarks>
    Public NoLogo As Boolean

    ''' <summary>
    ''' Quiet output mode.
    ''' </summary>
    ''' <remarks></remarks>
    Public QuietMode As Boolean

    ''' <summary>
    ''' Display verbose messages.
    ''' </summary>
    ''' <remarks></remarks>
    Public Verbose As Boolean

    ''' <summary>
    ''' Name of the bug report file to create.
    ''' </summary>
    ''' <remarks></remarks>
    Public BugReportFile As String

    ''' <summary>
    ''' The codepage to use when opening source files.
    ''' </summary>
    ''' <remarks></remarks>
    Public CodePage As Integer?

    ''' <summary>
    ''' Delay-sign the assembly using only the public portion of the strong name key.
    ''' </summary>
    ''' <remarks></remarks>
    Public DelaySign As Boolean

    ''' <summary>
    ''' The strong key name container to use when signing the assembly.
    ''' </summary>
    ''' <remarks></remarks>
    Public KeyContainer As String

    ''' <summary>
    ''' The strong name key file.
    ''' </summary>
    ''' <remarks></remarks>
    Public KeyFile As String

    ''' <summary>
    ''' List of directories to search for metadata references.
    ''' </summary>
    ''' <remarks></remarks>
    Public LibraryPaths As IList(Of String)

    ''' <summary>
    ''' The Class or Module that contains Sub Main
    ''' </summary>
    ''' <remarks></remarks>
    Public MainClassName As String

    ''' <summary>
    ''' Whether to target the Compact Framework or not.
    ''' </summary>
    ''' <remarks></remarks>
    Public TargetCompactFramework As Boolean

    ''' <summary>
    ''' Whether or not to reference standard libraries (system.dll and VBC.RSP file).
    ''' </summary>
    ''' <remarks></remarks>
    Public NoStandardLibraries As Boolean

    ''' <summary>
    ''' Location of the .NET Framework SDK directory
    ''' </summary>
    ''' <remarks></remarks>
    Public SDKPath As String

    ''' <summary>
    ''' Compile with/without the default Visual Basic runtime.
    ''' </summary>
    ''' <remarks></remarks>
    Public VBRuntime As VBRuntimeValues

    ''' <summary>
    ''' Name of the file containing the VB runtime to use.
    ''' </summary>
    ''' <remarks></remarks>
    Public VBRuntimeFileName As String

End Class
