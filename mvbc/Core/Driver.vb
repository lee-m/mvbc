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

#Region "Imports"

Imports System.Text

#End Region

''' <summary>
''' Compiler driver which oversees the compilation process.
''' </summary>
''' <remarks></remarks>
Public Class Driver

    ''' <summary>
    ''' Diagnostics manager used for error reporting.
    ''' </summary>
    ''' <remarks></remarks>
    Private mDiagnosticsManager As DiagnosticsManager

    ''' <summary>
    ''' Compiler settings based on the parsed command line options.
    ''' </summary>
    ''' <remarks></remarks>
    Private mCompilerSettings As CompilerSettings

#Region "Public Methods"

    ''' <summary>
    ''' Defauult constructor.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        mDiagnosticsManager = New DiagnosticsManager()
    End Sub

    ''' <summary>
    ''' Main entry point for the compilation process.
    ''' </summary>
    ''' <param name="commandLineArgs">Command line arguments.</param>
    ''' <returns>Application exit code.</returns>
    ''' <remarks></remarks>
    Public Function Compile(commandLineArgs As String()) As Integer

        'Suppress any command line errors until after the options have been parsed
        mDiagnosticsManager.EnableMessageQueuing()

        Dim cmdLine As New CommandLineParser(mDiagnosticsManager)
        Dim settings = cmdLine.Parse(commandLineArgs)

        If settings Is Nothing Then
            Return 1
        End If

        If settings.ShowHelp Then
            ShowUsage()
            Return 0
        End If

        'If we're not showing the help text, emit any errors/warnings found whilst parsing
        'the command line options
        mDiagnosticsManager.EmitQueuedMessages()

        If mDiagnosticsManager.NumberOfErrors > 0 Then
            Return 1
        End If

        Return 0

    End Function

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Outputs details of the various command line options which can be specified.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ShowUsage()

        Dim usageText As New StringBuilder
        usageText.AppendLine("                  Visual Basic Compiler Options")
        usageText.AppendLine("")
        usageText.AppendLine("                                  - OUTPUT FILE -")
        usageText.AppendLine("/out:<file>                       Specifies the output file name.")
        usageText.AppendLine("/target:exe                       Create a console application (default). (Short form: /t)")
        usageText.AppendLine("/target:winexe                    Create a Windows application.")
        usageText.AppendLine("/target:library                   Create a library assembly.")
        usageText.AppendLine("/doc[+|-]                         Generates XML documentation file.")
        usageText.AppendLine("/doc:<file>                       Generates XML documentation file to <file>.")
        usageText.AppendLine("")
        usageText.AppendLine("                                  - INPUT FILES -")
        usageText.AppendLine("/addmodule:<file_list>            Reference metadata from the specified modules.")
        usageText.AppendLine("/link:<file_list>                 Embed metadata from the specified interop assembly. (Short form: /l)")
        usageText.AppendLine("/recurse:<wildcard>               Include all files in the current directory and subdirectories according to the wildcard specifications.")
        usageText.AppendLine("/reference:<file_list>            Reference metadata from the specified assembly. (Short form: /r)")
        usageText.AppendLine("")
        usageText.AppendLine("                                  - RESOURCES -")
        usageText.AppendLine("/linkresource:<resinfo>           Links the specified file as an external assembly resource. resinfo:<file>[,<name>[,public|private]] (Short form: /linkres)")
        usageText.AppendLine("/nowin32manifest                  The default manifest should not be embedded in the manifest section of the output PE.")
        usageText.AppendLine("/resource:<resinfo>               Adds the specified file as an embedded assembly resource. resinfo:<file>[,<name>[,public|private]] (Short form: /res)")
        usageText.AppendLine("/win32icon:<file>                 Specifies a Win32 icon file (.ico) for the default Win32 resources.")
        usageText.AppendLine("/win32manifest:<file>             The provided file is embedded in the manifest section of the output PE.")
        usageText.AppendLine("/win32resource:<file>             Specifies a Win32 resource file (.res).")
        usageText.AppendLine("")
        usageText.AppendLine("                                  - CODE GENERATION -")
        usageText.AppendLine("/optimize[+|-]                    Enable optimizations.")
        usageText.AppendLine("/removeintchecks[+|-]             Remove integer checks. Default off.")
        usageText.AppendLine("/debug[+|-]                       Emit debugging information.")
        usageText.AppendLine("/debug:full                       Emit full debugging information (default).")
        usageText.AppendLine("/debug:pdbonly                    Emit PDB file only.")
        usageText.AppendLine("")
        usageText.AppendLine("                                  - ERRORS AND WARNINGS -")
        usageText.AppendLine("/nowarn                           Disable all warnings.")
        usageText.AppendLine("/nowarn:<number_list>             Disable a list of individual warnings.")
        usageText.AppendLine("/warnaserror[+|-]                 Treat all warnings as errors.")
        usageText.AppendLine("/warnaserror[+|-]:<number_list>   Treat a list of warnings as errors.")
        usageText.AppendLine("")
        usageText.AppendLine("                                  - LANGUAGE -")
        usageText.AppendLine("/define:<symbol_list>             Declare global conditional compilation symbol(s). symbol_list:name=value,... (Short form: /d)")
        usageText.AppendLine("/imports:<import_list>            Declare global Imports for namespaces in referenced metadata files. import_list:namespace,...")
        usageText.AppendLine("/optionexplicit[+|-]              Require explicit declaration of variables.")
        usageText.AppendLine("/optioninfer[+|-]                 Allow type inference of variables.")
        usageText.AppendLine("/rootnamespace:<string>           Specifies the root Namespace for all type declarations.")
        usageText.AppendLine("/optionstrict[+|-]                Enforce strict language semantics.")
        usageText.AppendLine("/optionstrict:custom              Warn when strict language semantics are not respected.")
        usageText.AppendLine("/optioncompare:binary             Specifies binary-style string comparisons. This is the default.")
        usageText.AppendLine("/optioncompare:text               Specifies text-style string comparisons.")
        usageText.AppendLine("")
        usageText.AppendLine("                                  - MISCELLANEOUS -")
        usageText.AppendLine("/help                             Display this usage message. (Short form: /?)")
        usageText.AppendLine("/noconfig                         Do not auto-include VBC.RSP file.")
        usageText.AppendLine("/nologo                           Do not display compiler copyright banner.")
        usageText.AppendLine("/quiet                            Quiet output mode.")
        usageText.AppendLine("/verbose                          Display verbose messages.")
        usageText.AppendLine("")
        usageText.AppendLine("                                  - ADVANCED -")
        usageText.AppendLine("/baseaddress:<number>             The base address for a library or module (hex).")
        usageText.AppendLine("/bugreport:<file>                 Create bug report file.")
        usageText.AppendLine("/codepage:<number>                Specifies the codepage to use when opening source files.")
        usageText.AppendLine("/delaysign[+|-]                   Delay-sign the assembly using only the public portion of the strong name key.")
        usageText.AppendLine("/errorreport:<string>             Specifies how to handle internal compiler errors; must be prompt, send, none, or queue (default).")
        usageText.AppendLine("/filealign:<number>               Specify the alignment used for output file sections.")
        usageText.AppendLine("/keycontainer:<string>            Specifies a strong name key container.")
        usageText.AppendLine("/keyfile:<file>                   Specifies a strong name key file.")
        usageText.AppendLine("/libpath:<path_list>              List of directories to search for metadata references. (Semi-colon delimited.)")
        usageText.AppendLine("/main:<class>                     Specifies the Class or Module that contains Sub Main. It can also be a Class that inherits from System.Windows.Forms.Form. (Short form: /m)")
        usageText.AppendLine("/netcf                            Target the .NET Compact Framework.")
        usageText.AppendLine("/nostdlib                         Do not reference standard libraries (system.dll and VBC.RSP file).")
        usageText.AppendLine("/platform:<string>                Limit which platforms this code can run on; must be x86, x64, Itanium, arm, AnyCPU32BitPreferred or anycpu (default).")
        usageText.AppendLine("/sdkpath:<path>                   Location of the .NET Framework SDK directory (mscorlib.dll).")
        usageText.AppendLine("/utf8output[+|-]                  Emit compiler output in UTF8 character encoding.")
        usageText.AppendLine("@<file>                           Insert command-line settings from a text file.")
        usageText.AppendLine("/vbruntime[+|-|*]                 Compile with/without the default Visual Basic runtime.")
        usageText.AppendLine("/vbruntime:<file>                 Compile with the alternate Visual Basic runtime in <file>.")
        Console.Write(usageText.ToString())

    End Sub

#End Region

End Class
