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

'                  Visual Basic Compiler Options
'
'                                  - OUTPUT FILE -
'/out:<file>                       Specifies the output file name.
'/target:exe                       Create a console application (default). (Short form: /t)
'/target:winexe                    Create a Windows application.
'/target:library                   Create a library assembly.
'/target:module                    Create a module that can be added to an assembly.
'/target:appcontainerexe           Create a Windows application that runs in AppContainer.
'/target:winmdobj                  Create a Windows Metadata intermediate file
'/doc[+|-]                         Generates XML documentation file.
'/doc:<file>                       Generates XML documentation file to <file>.

'                                  - INPUT FILES -
'/addmodule:<file_list>            Reference metadata from the specified modules.
'/link:<file_list>                 Embed metadata from the specified interop assembly. (Short form: /l)
'/recurse:<wildcard>               Include all files in the current directory and subdirectories according to the wildcard specifications.
'/reference:<file_list>            Reference metadata from the specified assembly. (Short form: /r)

'                                  - RESOURCES -
'/linkresource:<resinfo>           Links the specified file as an external assembly resource. resinfo:<file>[,<name>[,public|private]] (Short form: /linkres)
'/nowin32manifest                  The default manifest should not be embedded in the manifest section of the output PE.
'/resource:<resinfo>               Adds the specified file as an embedded assembly resource. resinfo:<file>[,<name>[,public|private]] (Short form: /res)
'/win32icon:<file>                 Specifies a Win32 icon file (.ico) for the default Win32 resources.
'/win32manifest:<file>             The provided file is embedded in the manifest section of the output PE.
'/win32resource:<file>             Specifies a Win32 resource file (.res).

'                                  - CODE GENERATION -
'/optimize[+|-]                    Enable optimizations.
'/removeintchecks[+|-]             Remove integer checks. Default off.
'/debug[+|-]                       Emit debugging information.
'/debug:full                       Emit full debugging information (default).
'/debug:pdbonly                    Emit PDB file only.

'                                  - ERRORS AND WARNINGS -
'/nowarn                           Disable all warnings.
'/nowarn:<number_list>             Disable a list of individual warnings.
'/warnaserror[+|-]                 Treat all warnings as errors.
'/warnaserror[+|-]:<number_list>   Treat a list of warnings as errors.

'                                  - LANGUAGE -
'/define:<symbol_list>             Declare global conditional compilation symbol(s). symbol_list:name=value,... (Short form: /d)
'/imports:<import_list>            Declare global Imports for namespaces in referenced metadata files. import_list:namespace,...
'/langversion:<number>             Specify language version: 9|10|11.
'/optionexplicit[+|-]              Require explicit declaration of variables.
'/optioninfer[+|-]                 Allow type inference of variables.
'/rootnamespace:<string>           Specifies the root Namespace for all type declarations.
'/optionstrict[+|-]                Enforce strict language semantics.
'/optionstrict:custom              Warn when strict language semantics are not respected.
'/optioncompare:binary             Specifies binary-style string comparisons. This is the default.
'/optioncompare:text               Specifies text-style string comparisons.

'                                  - MISCELLANEOUS -
'/help                             Display this usage message. (Short form: /?)
'/noconfig                         Do not auto-include VBC.RSP file.
'/nologo                           Do not display compiler copyright banner.
'/quiet                            Quiet output mode.
'/verbose                          Display verbose messages.

'                                  - ADVANCED -
'/baseaddress:<number>             The base address for a library or module (hex).
'/bugreport:<file>                 Create bug report file.
'/codepage:<number>                Specifies the codepage to use when opening source files.
'/delaysign[+|-]                   Delay-sign the assembly using only the public portion of the strong name key.
'/errorreport:<string>             Specifies how to handle internal compiler errors; must be prompt, send, none, or queue (default).
'/filealign:<number>               Specify the alignment used for output file sections.
'/highentropyva[+|-]               Enable high-entropy ASLR.
'/keycontainer:<string>            Specifies a strong name key container.
'/keyfile:<file>                   Specifies a strong name key file.
'/libpath:<path_list>              List of directories to search for metadata references. (Semi-colon delimited.)
'/main:<class>                     Specifies the Class or Module that contains Sub Main. It can also be a Class that inherits from System.Windows.Forms.Form. (Short form: /m)
'/moduleassemblyname:<string>      Name of the assembly which this module will be a part of.
'/netcf                            Target the .NET Compact Framework.
'/nostdlib                         Do not reference standard libraries (system.dll and VBC.RSP file).
'/platform:<string>                Limit which platforms this code can run on; must be x86, x64, Itanium, arm, AnyCPU32BitPreferred or anycpu (default).
'/sdkpath:<path>                   Location of the .NET Framework SDK directory (mscorlib.dll).
'/subsystemversion:<version>       Specify subsystem version of the output PE. version:<number>[.<number>]
'/utf8output[+|-]                  Emit compiler output in UTF8 character encoding.
'@<file>                           Insert command-line settings from a text file.
'/vbruntime[+|-|*]                 Compile with/without the default Visual Basic runtime.
'/vbruntime:<file>                 Compile with the alternate Visual Basic runtime in <file>.

''' <summary>
''' Parser for the command line options.
''' </summary>
''' <remarks></remarks>
Public Class CommandLineParser

    ''' <summary>
    ''' Diagnostics manager used to output any errors with the options.
    ''' </summary>
    ''' <remarks></remarks>
    Private mDiagnosticsMngr As DiagnosticsManager

    ''' <summary>
    ''' Used as the return value of ParseOption.
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum ParseOptionResult

        ''' <summary>
        ''' Parsing of remaining options should continue.
        ''' </summary>
        ''' <remarks></remarks>
        [Continue]

        ''' <summary>
        ''' Parsing of remaining options should be stopped.
        ''' </summary>
        ''' <remarks></remarks>
        [Stop]

    End Enum

#Region "Public Methods"

    ''' <summary>
    ''' Constructor to pass in the diagnostics manager instance.
    ''' </summary>
    ''' <param name="diagnosticMngr">Diagnostics manager instance to use for error reporting.</param>
    ''' <remarks></remarks>
    Public Sub New(diagnosticMngr As DiagnosticsManager)
        mDiagnosticsMngr = diagnosticMngr
    End Sub

    ''' <summary>
    ''' Creates this instance with the specified command line arguments.
    ''' </summary>
    ''' <param name="args">Command line arguments to parse.</param>
    ''' <remarks></remarks>
    Public Function Parse(args() As String) As CompilerSettings

        Dim settings As New CompilerSettings

        For Each argument In args

            Dim optParseResult As ParseOptionResult = ParseOptionResult.Continue

            'Switches may start with - or /
            If argument.StartsWith("/") _
               OrElse argument.StartsWith("-") Then
                optParseResult = ParseOption(argument.TrimStart(New Char() {"/"c, "-"c}), settings)
            Else
                settings.InputFiles.Add(New SourceFile(argument))
            End If

            If optParseResult = ParseOptionResult.Stop Then
                Exit For
            End If

        Next

        Return settings

    End Function

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Parses a single command line option and updates the compiler settings to reflect the value of that option.
    ''' </summary>
    ''' <param name="argument">The argument to parse.</param>
    ''' <param name="settings">Compiler settings to update.</param>
    ''' <returns>A parse result indicating whether the remaining options should be parsed or not.</returns>
    ''' <remarks></remarks>
    Private Function ParseOption(argument As String, settings As CompilerSettings) As ParseOptionResult

        Dim colonPosition As Integer = argument.IndexOf(":")
        Dim argName As String
        Dim argValue As String

        If colonPosition > 0 Then
            argName = argument.Substring(0, colonPosition)
            argValue = argument.Substring(colonPosition + 1)
        Else
            argName = argument
            argValue = String.Empty
        End If

        Select Case argName

            Case "debug+"
                settings.DebugInfoGenerationEnabled = True

            Case "debug-"
                settings.DebugInfoGenerationEnabled = False

            Case "debug"

                If String.IsNullOrEmpty(argValue) Then

                    If colonPosition > 0 Then
                        mDiagnosticsMngr.CommandLineError(2006, "debug", ":pdbonly|full")
                    End If

                ElseIf argValue = "pdb" OrElse argValue = "full" Then
                    settings.DebugInfoGenerationEnabled = True
                Else
                    mDiagnosticsMngr.CommandLineError(2014, argValue, "debug")
                End If

            Case "delaysign", "delaysign+"
                settings.DelaySign = True

            Case "delaysign-"
                settings.DelaySign = False

            Case "doc-"
                settings.XMLDocumentationEnabled = False

            Case "doc", "doc+"
                settings.XMLDocumentationEnabled = True

            Case "help", "?"
                settings.ShowHelp = True
                Return ParseOptionResult.Stop

            Case "noconfig"
                settings.NoConfig = True

            Case "nologo"
                settings.NoLogo = True

            Case "nostdlib"
                settings.NoStandardLibraries = True

            Case "nowarn"
                settings.DisableWarnings = True

            Case "nowin32manifest"
                settings.NoWin32Manifest = True

            Case "optimize", "optimize+"
                settings.Optimise = True

            Case "optimize-"
                settings.Optimise = False

            Case "optioncompare"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "optioncompare", ":binary|text")
                Else

                    If argValue = "binary" Then
                        settings.OptionCompare = OptionCompareValues.Binary
                    ElseIf argValue = "text" Then
                        settings.OptionCompare = OptionCompareValues.Text
                    Else
                        mDiagnosticsMngr.CommandLineError(2014, argValue, "optioncompare")
                    End If

                End If

            Case "optionexplicit", "optionexplicit+"
                settings.OptionExplicit = True

            Case "optionexplicit-"
                settings.OptionExplicit = False

            Case "optioninfer", "optioninfer+"
                settings.OptionInfer = True

            Case "optioninfer-"
                settings.OptionInfer = False

            Case "optionstrict"

                If String.IsNullOrEmpty(argValue) Then

                    If colonPosition > 0 Then
                        mDiagnosticsMngr.CommandLineError(2006, "optionstrict", ":custom")
                    Else
                        settings.OptionStrict = OptionStrictViolationValue.IssueError
                    End If

                ElseIf argValue = "custom" Then
                    settings.OptionStrict = OptionStrictViolationValue.IssueWarning
                Else
                    mDiagnosticsMngr.CommandLineError(2014, argValue, "optionstrict")
                End If

            Case "optionstrict+"
                settings.OptionStrict = OptionStrictViolationValue.IssueError

            Case "optionstrict-"
                settings.OptionStrict = OptionStrictViolationValue.Ignore

            Case "out"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "out", ":<file>")
                Else
                    settings.OutputFile = argValue
                End If

            Case "quiet"
                settings.QuietMode = True

            Case "removeintchecks", "removeintchecks+"
                settings.RemoveIntegerOverflowChecks = True

            Case "removeintchecks-"
                settings.RemoveIntegerOverflowChecks = False

            Case "rootnamespace"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "rootnamespace", ":<string>")
                Else
                    settings.RootNamespace = argValue
                End If

            Case "target", "t"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "target", ":exe|winexe|library")
                Else

                    'Check that the value matches one of the allowed target types
                    If argValue = "exe" Then
                        settings.OutputTarget = OutputTarget.Exe
                    ElseIf argValue = "winexe" Then
                        settings.OutputTarget = OutputTarget.WinExe
                    ElseIf argValue = "library" Then
                        settings.OutputTarget = OutputTarget.Library
                    Else
                        mDiagnosticsMngr.CommandLineError(2014, argValue, "target")
                    End If

                End If

            Case "verbose"
                settings.Verbose = True

            Case Else
                mDiagnosticsMngr.CommandLineWarning(2007, argName)

        End Select

        Return ParseOptionResult.Continue

    End Function

#End Region

End Class
