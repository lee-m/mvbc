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
''' Possible values for the /target switch
''' </summary>
''' <remarks></remarks>
Public Enum OutputTarget
    Exe
    WinExe
    Library
End Enum

Public Enum OptionStrictViolationValue
    IssueError
    Ignore
    IssueWarning
End Enum

Public Enum OptionCompareValues
    Text
    Binary
End Enum

Public Enum InternalCompilerErrorReportValues
    Prompt
    Send
    None
    Queue
End Enum

Public Enum PlatformValues
    x86
    x64
    Itanium
    ARM
    AnyCPU32BitPreferred
    AnyCPU
End Enum

Public Enum VBRuntimeValues
    WithVBRuntime
    WithoutVBRuntime
    EmbedVBRuntime
End Enum