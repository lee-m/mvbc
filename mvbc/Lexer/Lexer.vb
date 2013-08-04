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
''' Lexes a source file into a series of tokens for parsing.
''' </summary>
''' <remarks></remarks>
Public Class Lexer

    ''' <summary>
    ''' List of files to process.
    ''' </summary>
    ''' <remarks></remarks>
    Private mFiles As Queue(Of SourceFile)

    ''' <summary>
    ''' Current position in the current file.
    ''' </summary>
    ''' <remarks></remarks>
    Private mCurrPosition As CharEnumerator

    ''' <summary>
    ''' Current line number of the current position.
    ''' </summary>
    ''' <remarks></remarks>
    Private mLineNumber As Integer

    ''' <summary>
    ''' Column number of the urrent position in the current file.
    ''' </summary>
    ''' <remarks></remarks>
    Private mColumnNumber As Integer

    ''' <summary>
    ''' List of keywords.
    ''' </summary>
    ''' <remarks></remarks>
    Private mKeywords As HashSet(Of String)

    ''' <summary>
    ''' List of characters which indicate the start of a comment
    ''' </summary>
    ''' <remarks></remarks>
    Private mStartOfCommentChars As HashSet(Of Char)

    ''' <summary>
    ''' List of characters which indicate a new line.
    ''' </summary>
    ''' <remarks></remarks>
    Private mLineTerminatorChars As HashSet(Of Char)

#Region "Public Methods"

    ''' <summary>
    ''' Initialise the lexer with a set of files to process.
    ''' </summary>
    ''' <param name="files">Set of input files.</param>
    ''' <remarks></remarks>
    Public Sub New(files As IEnumerable(Of SourceFile))

        mFiles = New Queue(Of SourceFile)(files)
        mKeywords = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From
            {"AddHandler", "AddressOf", "Alias", "And", "AndAlso", "As", "Boolean", "ByRef", "Byte", "ByVal", "Call",
             "Case", "Catch", "CBool", "CByte", "CChar", "CDate", "CDbl", "CDec", "Char", "CInt", "Class", "CLng",
             "CObj", "Const", "Continue", "CSByte", "CShort", "CSng", "CStr", "CType", "CUInt", "CULng", "CUShort", "Date",
             "Decimal", "Declare", "Default", "Delegate", "Dim", "DirectCast", "Do", "Double", "Each", "Else", "ElseIf",
             "End", "EndIf", "Enum", "Erase", "Error", "Event", "Exit", "False", "Finally", "For", "Friend", "Function",
             "Get", "GetType", "GetXmlNamespace", "Global", "GoSub", "GoTo", "Handles", "If", "Implements", "Imports",
             "In", "Inherits", "Integer", "Interface", "Is", "IsNot", "Let", "Lib", "Like", "Long", "Loop", "Me", "Mod",
             "Module", "MustInherit", "MustOverride", "MyBase", "MyClass", "Namespace", "Narrowing", "New", "Next", "Not",
             "Nothing", "NotInheritable", "NotOverridable", "Object", "Of", "On", "Operator", "Option", "Optional", "Or",
             "OrElse", "Overloads", "Overridable", "Overrides", "ParamArray", "Partial", "Private", "Property", "Protected",
             "Public", "RaiseEvent", "ReadOnly", "ReDim", "REM", "RemoveHandler", "Resume", "Return", "SByte", "Select",
             "Set", "Shadows", "Shared", "Short", "Single", "Static", "Step", "Stop", "String", "Structure", "Sub", "SyncLock",
             "Then", "Throw", "To", "True", "Try", "TryCast", "TypeOf", "UInteger", "ULong", "UShort", "Using", "Variant",
             "Wend", "When", "While", "Widening", "With", "WithEvents", "WriteOnly", "Xor"}
        mStartOfCommentChars = New HashSet(Of Char) From {"'"c, ChrW(&H2018), ChrW(&H2019)}
        NextFile()

    End Sub

    ''' <summary>
    ''' Lexes and returns the next token
    ''' </summary>
    ''' <returns>The next token that was lexed.</returns>
    ''' <remarks></remarks>
    Public Function NextToken() As Token

        SkipWhitespaceAndComments()
        Return New Token

    End Function

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Skips over and whitespace and comment characters.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SkipWhitespaceAndComments()

        While Not EndOfInput()

            Dim ch = NextCharacter()

            If Char.IsWhiteSpace(ch) Then
                Continue While
            ElseIf mStartOfCommentChars.Contains(ch) Then

                ch = NextCharacter()

                'Eat characters until we hit the start of a new line sequence
                While Not EndOfInput() _
                      AndAlso Not mLineTerminatorChars.Contains(ch)
                    ch = NextCharacter()
                End While

                'Skip over the new line sequence
                SkipNewLineCharacters(ch)

            End If

        End While

    End Sub

    ''' <summary>
    ''' Skips over a new line character sequence.
    ''' </summary>
    ''' <param name="ch">The first character of any potential multi-character end of line sequence.</param>
    ''' <remarks></remarks>
    Private Sub SkipNewLineCharacters(ch As Char)

    End Sub

    ''' <summary>
    ''' Indicates if the end of all input files has been reached.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function EndOfInput() As Boolean
        Return False
    End Function

    ''' <summary>
    ''' Advances the current position in the current file and returns the character at that position.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function NextCharacter() As Char

        If Not mCurrPosition.MoveNext() Then

            If Not NextFile() Then
                Return Chr(0)
            End If

        End If

        mColumnNumber += 1
        Return mCurrPosition.Current

    End Function

    ''' <summary>
    ''' Reads the contents of the next file in for lexing.
    ''' </summary>
    ''' <remarks></remarks>
    Private Function NextFile() As Boolean

        If mFiles.Any() Then

            Using reader As New IO.StreamReader(mFiles.Dequeue().FileName)
                mCurrPosition = reader.ReadToEnd().GetEnumerator()
            End Using

            mLineNumber = 1
            mColumnNumber = 0
            Return True

        Else
            Return False
        End If

    End Function

#End Region

End Class
