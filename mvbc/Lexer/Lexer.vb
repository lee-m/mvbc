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
    Implements IDisposable

    ''' <summary>
    ''' Encoding to open the source files in.
    ''' </summary>
    ''' <remarks></remarks>
    Private mEncoding As Encoding

    ''' <summary>
    ''' Input buffer to lex to the tokens from.
    ''' </summary>
    ''' <remarks></remarks>
    Private mInputBuffer As LexerInputBuffer

    ''' <summary>
    ''' The current source file being lexed.
    ''' </summary>
    ''' <remarks></remarks>
    Private mCurrentFile As SourceFile

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

#Region "Constants"

    ''' <summary>
    ''' Carriage return character constant.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const CarriageReturnChar As Char = ChrW(13)

    ''' <summary>
    ''' Line feed character constant.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const LineFeedChar As Char = ChrW(10)

    ''' <summary>
    ''' Unicode line separator character constant.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const LineSeparatorChar As Char = ChrW(&H2028)

    ''' <summary>
    ''' Unicode paragraph separator character constant.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const ParagraphSeparatorChar As Char = ChrW(&H2029)

#End Region

#Region "Public Methods"

    ''' <summary>
    ''' Initialise the lexer with a set of files to process.
    ''' </summary>
    ''' <param name="inputFile">Input file to lex.</param>
    ''' <param name="encoding">Encoding to open the input file in.</param>
    ''' <remarks></remarks>
    Public Sub New(inputFile As SourceFile, encoding As Encoding)

        mLineNumber = 1
        mColumnNumber = 0
        mCurrentFile = inputFile
        mEncoding = encoding
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
        mLineTerminatorChars = New HashSet(Of Char) From {CarriageReturnChar, LineFeedChar, LineSeparatorChar, ParagraphSeparatorChar}

        If mEncoding IsNot Nothing Then
            mInputBuffer = New LexerInputBuffer(New StreamReader(mCurrentFile.FileName, mEncoding))
        Else
            mInputBuffer = New LexerInputBuffer(New StreamReader(mCurrentFile.FileName))
        End If

    End Sub

    ''' <summary>
    ''' Releases the input file handle.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Dispose() Implements IDisposable.Dispose
        mInputBuffer.Dispose()
    End Sub

    ''' <summary>
    ''' Lexes and returns the next token
    ''' </summary>
    ''' <returns>The next token that was lexed.</returns>
    ''' <remarks></remarks>
    Public Function NextToken() As Token

        SkipWhitespaceAndComments()

        Select Case CurrentCharacter

            Case CarriageReturnChar, LineFeedChar, LineSeparatorChar, ParagraphSeparatorChar
                SkipNewLineCharacters(mInputBuffer.CurrentCharacter)
                Return CreateToken(TokenType.TT_END_OF_STATEMENT)

            Case "+"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(TokenType.OP_PLUS_EQ)
                Else
                    Return CreateToken(TokenType.OP_PLUS)
                End If

            Case "-"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(TokenType.OP_SUBTRACT_EQ)
                Else
                    Return CreateToken(TokenType.OP_SUBTRACT)
                End If

            Case "*"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(TokenType.OP_MULTIPLY_EQ)
                Else
                    Return CreateToken(TokenType.OP_MULTIPLY)
                End If

            Case "/"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(TokenType.OP_DIVISION_EQ)
                Else
                    Return CreateToken(TokenType.OP_DIVISION)
                End If

            Case "\"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(TokenType.OP_INT_DIVISION_EQ)
                Else
                    Return CreateToken(TokenType.OP_INT_DIVISION)
                End If

            Case "&"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(TokenType.OP_CONCAT_EQ)
                Else
                    Return CreateToken(TokenType.OP_CONCAT)
                End If

            Case "^"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(TokenType.OP_EXPONENTIAL_EQ)
                Else
                    Return CreateToken(TokenType.OP_EXPONENTIAL)
                End If

            Case "<"c

                'Consume the <
                NextCharacter()

                'May be one of <, <<, <=, <<=, <>
                Dim currChar As Char = CurrentCharacter

                If currChar = ">"c Then
                    NextCharacter()
                    Return CreateToken(TokenType.OP_NOT_EQUAL)
                ElseIf currChar = "<" Then

                    'Consume the second < and peek the following character
                    NextCharacter()
                    currChar = CurrentCharacter

                    If currChar = "=" Then

                        'Consume the =
                        NextCharacter()
                        Return CreateToken(TokenType.OP_LSHIFT_EQ)

                    Else
                        Return CreateToken(TokenType.OP_LSHIFT)
                    End If

                ElseIf currChar = "=" Then
                    NextCharacter()
                    Return CreateToken(TokenType.OP_LESS_THAN_EQ)
                Else
                    Return CreateToken(TokenType.OP_LESS_THAN)
                End If

            Case ">"c

                'Consume the >
                NextCharacter()

                If CurrentCharacter = ">" Then

                    'Consume the second >
                    NextCharacter()

                    If CurrentCharacter = "=" Then
                        NextCharacter()
                        Return CreateToken(TokenType.OP_RSHIFT_EQ)
                    Else
                        Return CreateToken(TokenType.OP_RSHIFT)
                    End If

                ElseIf CurrentCharacter = "=" Then

                    NextCharacter()
                    Return CreateToken(TokenType.OP_GREATER_THAN_EQ)

                Else
                    Return CreateToken(TokenType.OP_GREATER_THAN)
                End If

            Case "="c
                NextCharacter()
                Return CreateToken(TokenType.OP_EQUALS, Nothing)

        End Select

        Return Nothing

    End Function

    ''' <summary>
    ''' Indicates if the end of all input files has been reached.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function EndOfInput() As Boolean
        Return mInputBuffer.EndOfBuffer
    End Function

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Helper method to create a token given its type and value.
    ''' </summary>
    ''' <param name="tokType">Type of the token.</param>
    ''' <param name="value">Value of the token.</param>
    ''' <returns>A new Token instance initialised with the specified type and value.</returns>
    ''' <remarks></remarks>
    Private Function CreateToken(tokType As TokenType, Optional value As Object = Nothing) As Token
        Return New Token(mCurrentFile, mLineNumber, mColumnNumber, tokType, value)
    End Function

    ''' <summary>
    ''' Skips over and whitespace and comment characters.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SkipWhitespaceAndComments()

        While Not EndOfInput()

            Dim ch As Char = CurrentCharacter

            'Don't want to skip newlines, those are treated as their own token
            If Char.IsWhiteSpace(ch) _
               AndAlso Not mLineTerminatorChars.Contains(ch) Then
                NextCharacter()
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

            Else
                Return
            End If

        End While

    End Sub

    ''' <summary>
    ''' Skips over a new line character sequence.
    ''' </summary>
    ''' <param name="ch">The first character of any potential multi-character end of line sequence.</param>
    ''' <remarks></remarks>
    Private Sub SkipNewLineCharacters(ch As Char)

        If ch = CarriageReturnChar Then

            If NextCharacter() = LineFeedChar Then
                NextCharacter()
            End If

        ElseIf ch = LineFeedChar _
               OrElse ch = LineSeparatorChar _
               OrElse ch = ParagraphSeparatorChar Then
            NextCharacter()
        End If

        mLineNumber += 1
        mColumnNumber = 0

    End Sub

    ''' <summary>
    ''' Advances the current position in the current file and returns the character at that position.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function NextCharacter() As Char

        mColumnNumber += 1
        Return mInputBuffer.NextCharacter()

    End Function

    ''' <summary>
    ''' Peeks a single character from the input buffer.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function PeekCharacter() As Char
        Return mInputBuffer.PeekCharacter()
    End Function

#End Region

#Region "Private Properties"

    ''' <summary>
    ''' Accessor for the character at the current position in the input file.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private ReadOnly Property CurrentCharacter() As Char
        Get
            Return mInputBuffer.CurrentCharacter
        End Get
    End Property

#End Region

End Class