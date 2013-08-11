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
        Dim processedResponseFiles As New HashSet(Of String)

        For Each argument In args

            Dim optParseResult As ParseOptionResult = ParseOptionResult.Continue

            'Switches may start with - or /
            If argument.StartsWith("/") _
               OrElse argument.StartsWith("-") Then
                optParseResult = ParseOption(argument.TrimStart(New Char() {"/"c, "-"c}), settings)
            ElseIf argument.StartsWith("@") Then

                Dim absPath = New FileInfo(argument.TrimStart("@"c)).FullName

                If Not File.Exists(absPath) Then
                    mDiagnosticsMngr.FatalError(2011, absPath)
                ElseIf processedResponseFiles.Contains(absPath) Then
                    mDiagnosticsMngr.CommandLineError(2003, absPath)
                Else

                    processedResponseFiles.Add(absPath)

                    Using reader As New StreamReader(absPath)
                        ParseResponseFile(reader, settings)
                    End Using

                End If

            Else
                settings.InputFiles.Add(New SourceFile(argument))
            End If

            If optParseResult = ParseOptionResult.Stop Then
                Exit For
            End If

        Next

        'Process the standard VBC.rsp file unless /noconfig was specified
        If Not settings.NoConfig Then

            Using reader As New StreamReader("mvbc.rsp")
                ParseResponseFile(reader, settings)
            End Using

        End If

        'Make sure that any referenced file can be found
        ValidateReferencedFiles(settings.ReferencedAssemblies, settings.LibraryPaths)

        Return settings

    End Function

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Ensures that any referenced file can be found.
    ''' </summary>
    ''' <param name="referencedFiles">The set of referenced files specified on the command line.</param>
    ''' <param name="libraryPaths">Library search paths to process when trying to find a referenced file.</param>
    ''' <remarks></remarks>
    Private Sub ValidateReferencedFiles(referencedFiles As HashSet(Of String),
                                        libraryPaths As IEnumerable(Of String))

        Dim searchPaths As New List(Of String)
        searchPaths.Add(Environment.CurrentDirectory)
        searchPaths.AddRange(libraryPaths)

        'Keep track of the name of each referenced assembly so we can detect duplicates
        Dim referencedAssemblyNames As New HashSet(Of String)

        For Each referencedFile In referencedFiles

            Dim foundFilePath As String = Nothing

            'If we have an absolute path, use that. Otherwise make a pass through each library path and try and find
            'the file in there
            If Path.IsPathRooted(referencedFile) Then

                If File.Exists(referencedFile) Then
                    foundFilePath = referencedFile
                End If

            Else

                For Each searchPath In searchPaths

                    Dim tempPath As String = Path.Combine(searchPath, referencedFile)

                    If File.Exists(tempPath) Then
                        foundFilePath = tempPath
                        Exit For
                    End If

                Next

            End If

            If String.IsNullOrEmpty(foundFilePath) Then

                mDiagnosticsMngr.CommandLineError(2017, Path.GetFileName(referencedFile))
                Exit Sub

            Else
                CheckReferencedFileIsValidAssembly(foundFilePath, referencedAssemblyNames)
            End If

        Next

    End Sub

    ''' <summary>
    ''' Checks that a referenced file specified via /r: or /reference: (and has already been checked that it does exist) is 
    ''' a valid assembly and hasn't already been loaded via another file path
    ''' </summary>
    ''' <param name="assemblyPath"></param>
    ''' <param name="prevLoadedAssemblies"></param>
    ''' <remarks></remarks>
    Private Sub CheckReferencedFileIsValidAssembly(assemblyPath As String, prevLoadedAssemblies As HashSet(Of String))

        'Make sure that it is an assembly
        Dim validAssembly As Boolean = False
        Dim referencedAssemblyName As String = Nothing

        Try

            'GetAssemblyName will throw an exception if the file is not a valid assembly
            referencedAssemblyName = AssemblyName.GetAssemblyName(assemblyPath).Name
            validAssembly = True

        Catch badFormatEx As BadImageFormatException
            validAssembly = False

        End Try

        If Not validAssembly Then
            mDiagnosticsMngr.FatalError(2000, String.Format("'{0}' cannot be referenced because it is not an assembly.", assemblyPath))
            Exit Sub
        End If

        'Reject attempts to load the same assembly via two different files
        If Not prevLoadedAssemblies.Add(referencedAssemblyName) Then

            mDiagnosticsMngr.FatalError(2000,
                                        String.Format("Project already has a reference to assembly {0}. A second reference to '{1}' cannot be added.",
                                                      referencedAssemblyName, assemblyPath))
            Exit Sub

        End If

    End Sub

    ''' <summary>
    ''' Parses a set of options from the specified response file
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <remarks></remarks>
    Private Sub ParseResponseFile(responseFileReader As TextReader, settings As CompilerSettings)

        Dim optionLine As String

        While True

            optionLine = responseFileReader.ReadLine()

            If String.IsNullOrEmpty(optionLine) Then
                Exit Sub
            End If

            'Lines starting with a # are comments
            If optionLine.StartsWith("#") Then
                Continue While
            End If

            Dim result = ParseOption(optionLine.TrimStart(New Char() {"/"c, "-"c}), settings)

            If result = ParseOptionResult.Stop Then
                Exit Sub
            End If

        End While

    End Sub

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

            Case "bugreport"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "bugreport", ":<file>")
                Else
                    settings.BugReportFile = argValue
                End If

            Case "codepage"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "codepage", ":<number>")
                Else

                    Try

                        Dim codePageID As Integer = Integer.Parse(argValue)
                        Dim encoding As Encoding = encoding.GetEncoding(codePageID)

                        'Code page ID is valid
                        settings.CodePage = codePageID

                    Catch ex As Exception
                        mDiagnosticsMngr.CommandLineError(2016, argValue)
                    End Try

                End If

            Case "debug+"
                settings.DebugInfoGenerationEnabled = True

                If Not String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2009, "debug")
                End If

            Case "debug-"
                settings.DebugInfoGenerationEnabled = False

                If Not String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2009, "debug")
                End If

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

            Case "define", "d"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, argName, ":<symbol_list>")
                Else

                    Dim values As String() = argValue.Split(","c)

                    For Each defineValue In values

                        Dim symbolName As String = Nothing
                        Dim symbolValue As String = Nothing

                        If defineValue.Contains("="c) Then

                            Dim eqPos As Integer = defineValue.IndexOf("=")
                            symbolName = defineValue.Substring(0, eqPos)

                            If Not defineValue.EndsWith("=") Then
                                symbolValue = defineValue.Substring(eqPos + 1, defineValue.Length - eqPos - 1)
                            End If

                        Else
                            symbolName = defineValue
                        End If

                        settings.Defines.Add(New KeyValuePair(Of String, String)(symbolName, symbolValue))

                    Next

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

            Case "imports"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "imports", ":<import_list>")
                Else
                    settings.GlobalImports.UnionWith(argValue.Split(","c))
                End If

            Case "keycontainer"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "keycontainer", ":<string>")
                Else
                    settings.KeyContainer = argValue
                End If

            Case "keyfile"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "keyfile", ":<file>")
                Else
                    settings.KeyFile = argValue
                End If

            Case "libpath"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "libpath", ":<path_list>")
                Else

                    Dim paths As String() = argValue.Split(";"c)

                    For Each libPath In paths

                        Try

                            'VBC doesn't warn/error for invalid paths but check that the path exists before adding it to the list
                            If Directory.Exists(New DirectoryInfo(libPath).FullName) Then
                                settings.LibraryPaths.Add(libPath)
                            End If

                        Catch argEx As ArgumentException
                        Catch pathEx As PathTooLongException
                            'Ignore - we may get here if the user does something silly like /libpath:? 
                            'or /libpath:<path longer than 255 chars>
                        End Try

                    Next

                End If

            Case "main"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "main", ":<class>")
                Else
                    settings.MainClassName = argValue
                End If

            Case "netcf"
                settings.TargetCompactFramework = True

            Case "noconfig"
                settings.NoConfig = True

            Case "nologo"
                settings.NoLogo = True

            Case "nostdlib"
                settings.NoStandardLibraries = True

            Case "nowarn"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.SuppressAllWarnings()
                Else

                    Dim parsedWarnings As HashSet(Of Integer) = ParseWarningList(argValue, argName)
                    mDiagnosticsMngr.DisabledWarnings.UnionWith(parsedWarnings)

                End If

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

                If Not String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2009, "optionstrict")
                End If

            Case "optionstrict-"
                settings.OptionStrict = OptionStrictViolationValue.Ignore

                If Not String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2009, "optionstrict")
                End If

            Case "out"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "out", ":<file>")
                Else
                    settings.OutputFile = argValue
                End If

            Case "reference", "r"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, argName, ":<file_list>")
                Else

                    'Any referenced file's validity is checked after all the options are parsed so that we can take into 
                    'account any /libpath options
                    Dim referencedFiles As String() = argValue.Split(","c)
                    settings.ReferencedAssemblies.UnionWith(referencedFiles)

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

            Case "sdkpath"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "sdkpath", ":<path>")
                Else
                    settings.SDKPath = argValue
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

            Case "vbruntime-", "vbruntime+"

                If colonPosition > 0 Then
                    mDiagnosticsMngr.CommandLineError(2009, "vbruntime")
                ElseIf argName.EndsWith("+") Then
                    settings.VBRuntime = VBRuntimeValues.WithVBRuntime
                ElseIf argName.EndsWith("-") Then
                    settings.VBRuntime = VBRuntimeValues.WithoutVBRuntime
                End If

            Case "vbruntime"
                settings.VBRuntime = VBRuntimeValues.WithVBRuntime
                settings.VBRuntimeFileName = argValue

            Case "verbose"
                settings.Verbose = True

            Case "warnaserror", "warnaserror+", "warnaserror-"

                '/warnaserror+: or /warnaserror-: doesn't do anything
                If colonPosition > 0 Then

                    If Not String.IsNullOrEmpty(argValue) Then

                        Dim parsedWarnings As HashSet(Of Integer) = ParseWarningList(argValue, "warnaserror")

                        If argName.EndsWith("-") Then
                            mDiagnosticsMngr.WarningsAsErrors.ExceptWith(parsedWarnings)
                        Else
                            mDiagnosticsMngr.WarningsAsErrors.UnionWith(parsedWarnings)
                        End If

                    End If

                Else

                    If argName.EndsWith("-") Then
                        mDiagnosticsMngr.WarningsAsErrors.Clear()
                    Else
                        mDiagnosticsMngr.TreatAllWarningsAsErrors()
                    End If

                End If

            Case "win32icon"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "win32icon", ":<file>")
                ElseIf Not String.IsNullOrEmpty(settings.Win32ResourceFileName) Then
                    mDiagnosticsMngr.CommandLineError(2023)
                Else
                    settings.Win32IconFileName = argValue
                End If

            Case "win32manifest"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "win32manifest", ":<file>")
                Else
                    settings.Win32ManifestFileName = argValue
                End If

            Case "win32resource"

                If String.IsNullOrEmpty(argValue) Then
                    mDiagnosticsMngr.CommandLineError(2006, "win32resource", ":<file>")
                ElseIf Not String.IsNullOrEmpty(settings.Win32IconFileName) Then
                    mDiagnosticsMngr.CommandLineError(2023)
                Else
                    settings.Win32ResourceFileName = argValue
                End If

            Case Else
                mDiagnosticsMngr.CommandLineWarning(2007, argName)

        End Select

        Return ParseOptionResult.Continue

    End Function

    ''' <summary>
    ''' Parses a set of warning numbers from a command line option's value.
    ''' </summary>
    ''' <param name="argValue">Option value containing the comma separated list of warning numbers.</param>
    ''' <param name="optionName">The name of the option the warning IDs belong to.</param>
    ''' <returns>A set of parsed warning IDs.</returns>
    ''' <remarks></remarks>
    Private Function ParseWarningList(argValue As String, optionName As String) As HashSet(Of Integer)

        Dim warningIDs = argValue.Split(New String() {","}, StringSplitOptions.RemoveEmptyEntries)
        Dim parsedWarnings As New HashSet(Of Integer)

        For Each warningID In warningIDs

            Dim parsedWarningID As Integer = 0

            If Not Integer.TryParse(warningID, parsedWarningID) Then
                mDiagnosticsMngr.CommandLineError(2014, warningID, optionName)
            ElseIf Not mDiagnosticsMngr.IsValidWarningID(parsedWarningID) Then
                mDiagnosticsMngr.CommandLineWarning(2026, warningID, optionName)
            Else
                parsedWarnings.Add(parsedWarningID)
            End If

        Next

        Return parsedWarnings

    End Function

#End Region

End Class
