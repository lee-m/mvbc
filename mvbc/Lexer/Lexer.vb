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
Public NotInheritable Class Lexer
    Implements IDisposable

    ''' <summary>
    ''' Diagnostics manager for reporting errors.
    ''' </summary>
    ''' <remarks></remarks>
    Private mDiagnosticsMngr As DiagnosticsManager

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
    Private mKeywords As Dictionary(Of String, TokenType)

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

    ''' <summary>
    ''' List of valid double-quote characters.
    ''' </summary>
    ''' <remarks></remarks>
    Private mDoubleQuoteChars As HashSet(Of Char)

    ''' <summary>
    ''' Set of Unicode character categories which are valid for an alpha character.
    ''' </summary>
    ''' <remarks></remarks>
    Private mAlphaCharUnicodeCategories As HashSet(Of UnicodeCategory)

    ''' <summary>
    ''' Set of Unicode character categories which are valid for inclusion in an identifier.
    ''' </summary>
    ''' <remarks></remarks>
    Private mIdentifierCharUnicodeCategories As HashSet(Of UnicodeCategory)

    ''' <summary>
    ''' Set of characters which can appear in an octal literal.
    ''' </summary>
    ''' <remarks></remarks>
    Private mOctalDigits As HashSet(Of Char)

    ''' <summary>
    ''' Set of characters which are valid in a hex literal
    ''' </summary>
    ''' <remarks></remarks>
    Private mHexDigits As HashSet(Of Char)

    ''' <summary>
    ''' Enumeration of potential type characters which may appear after an integral literal
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum IntegralLiteralTypeCharacter
        None
        ShortTypeChar
        UnsignedShortTypeChar
        IntTypeChar
        UnsignedIntTypeChar
        LongTypeChar
        UnsignedLongTypeChar
        DecimalTypeChar
    End Enum

    ''' <summary>
    ''' Enumeration of potential type characters which may appear after a floating point literal.
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum FloatingPointLiteralTypeCharacter
        None
        SingleTypeChar
        DoubleTypeChar
        DecimalTypeChar
    End Enum

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
    Public Sub New(inputFile As SourceFile, encoding As Encoding, diagnosticsMngr As DiagnosticsManager)

        mLineNumber = 1
        mColumnNumber = 1
        mCurrentFile = inputFile
        mEncoding = encoding
        mDiagnosticsMngr = diagnosticsMngr

        mKeywords = New Dictionary(Of String, TokenType)(StringComparer.OrdinalIgnoreCase) From
            {{"AddHandler", TokenType.KW_ADDHANDLER},
             {"AddressOf", TokenType.KW_ADDRESSOF},
             {"Alias", TokenType.KW_ALIAS},
             {"And", TokenType.KW_AND},
             {"AndAlso", TokenType.KW_ANDALSO},
             {"As", TokenType.KW_AS},
             {"Boolean", TokenType.KW_BOOLEAN},
             {"ByRef", TokenType.KW_BYREF},
             {"Byte", TokenType.KW_BYTE},
             {"ByVal", TokenType.KW_BYVAL},
             {"Call", TokenType.KW_CALL},
             {"Case", TokenType.KW_CASE},
             {"Catch", TokenType.KW_CATCH},
             {"CBool", TokenType.KW_CBOOL},
             {"CByte", TokenType.KW_CBYTE},
             {"CChar", TokenType.KW_CCHAR},
             {"CDate", TokenType.KW_CDATE},
             {"CDbl", TokenType.KW_CDBL},
             {"CDec", TokenType.KW_CDEC},
             {"Char", TokenType.KW_CHAR},
             {"CInt", TokenType.KW_CINT},
             {"Class", TokenType.KW_CLASS},
             {"CLng", TokenType.KW_CLNG},
             {"CObj", TokenType.KW_COBJ},
             {"Const", TokenType.KW_CONST},
             {"Continue", TokenType.KW_CONTINUE},
             {"CSByte", TokenType.KW_CSBYTE},
             {"CShort", TokenType.KW_CSHORT},
             {"CSng", TokenType.KW_CSNG},
             {"CStr", TokenType.KW_CSTR},
             {"CType", TokenType.KW_CTYPE},
             {"CUInt", TokenType.KW_CUINT},
             {"CULng", TokenType.KW_CULNG},
             {"CUShort", TokenType.KW_CUSHORT},
             {"Date", TokenType.KW_DATE},
             {"Decimal", TokenType.KW_DECIMAL},
             {"Declare", TokenType.KW_DECLARE},
             {"Default", TokenType.KW_DEFAULT},
             {"Delegate", TokenType.KW_DELEGATE},
             {"Dim", TokenType.KW_DIM},
             {"DirectCast", TokenType.KW_DIRECTCAST},
             {"Do", TokenType.KW_DO},
             {"Double", TokenType.KW_DOUBLE},
             {"Each", TokenType.KW_EACH},
             {"Else", TokenType.KW_ELSE},
             {"ElseIf", TokenType.KW_ELSEIF},
             {"End", TokenType.KW_END},
             {"EndIf", TokenType.KW_ENDIF},
             {"Enum", TokenType.KW_ENUM},
             {"Erase", TokenType.KW_ERASE},
             {"Error", TokenType.KW_ERROR},
             {"Event", TokenType.KW_EVENT},
             {"Exit", TokenType.KW_EXIT},
             {"False", TokenType.KW_FALSE},
             {"Finally", TokenType.KW_FINALLY},
             {"For", TokenType.KW_FOR},
             {"Friend", TokenType.KW_FRIEND},
             {"Function", TokenType.KW_FUNCTION},
             {"Get", TokenType.KW_GET},
             {"GetType", TokenType.KW_GETTYPE},
             {"GetXmlNamespace", TokenType.KW_GETXMLNAMESPACE},
             {"Global", TokenType.KW_GLOBAL},
             {"GoSub", TokenType.KW_GOSUB},
             {"GoTo", TokenType.KW_GOTO},
             {"Handles", TokenType.KW_HANDLES},
             {"If", TokenType.KW_IF},
             {"Implements", TokenType.KW_IMPLEMENTS},
             {"Imports", TokenType.KW_IMPORTS},
             {"In", TokenType.KW_IN},
             {"Inherits", TokenType.KW_INHERITS},
             {"Integer", TokenType.KW_INTEGER},
             {"Interface", TokenType.KW_INTERFACE},
             {"Is", TokenType.KW_IS},
             {"IsNot", TokenType.KW_ISNOT},
             {"Let", TokenType.KW_LET},
             {"Lib", TokenType.KW_LIB},
             {"Like", TokenType.KW_LIKE},
             {"Long", TokenType.KW_LONG},
             {"Loop", TokenType.KW_LOOP},
             {"Me", TokenType.KW_ME},
             {"Mod", TokenType.KW_MOD},
             {"Module", TokenType.KW_MODULE},
             {"MustInherit", TokenType.KW_MUSTINHERIT},
             {"MustOverride", TokenType.KW_MUSTOVERRIDE},
             {"MyBase", TokenType.KW_MYBASE},
             {"MyClass", TokenType.KW_MYCLASS},
             {"Namespace", TokenType.KW_NAMESPACE},
             {"Narrowing", TokenType.KW_NARROWING},
             {"New", TokenType.KW_NEW},
             {"Next", TokenType.KW_NEXT},
             {"Not", TokenType.KW_NOT},
             {"Nothing", TokenType.KW_NOTHING},
             {"NotInheritable", TokenType.KW_NOTINHERITABLE},
             {"NotOverridable", TokenType.KW_NOTOVERRIDABLE},
             {"Object", TokenType.KW_OBJECT},
             {"Of", TokenType.KW_OF},
             {"On", TokenType.KW_ON},
             {"Operator", TokenType.KW_OPERATOR},
             {"Option", TokenType.KW_OPTION},
             {"Optional", TokenType.KW_OPTIONAL},
             {"Or", TokenType.OP_OR},
             {"OrElse", TokenType.KW_ORELSE},
             {"Overloads", TokenType.KW_OVERLOADS},
             {"Overridable", TokenType.KW_OVERRIDABLE},
             {"Overrides", TokenType.KW_OVERRIDES},
             {"ParamArray", TokenType.KW_PARAMARRAY},
             {"Partial", TokenType.KW_PARTIAL},
             {"Private", TokenType.KW_PRIVATE},
             {"Property", TokenType.KW_PROPERTY},
             {"Protected", TokenType.KW_PROTECTED},
             {"Public", TokenType.KW_PUBLIC},
             {"RaiseEvent", TokenType.KW_RAISEEVENT},
             {"ReadOnly", TokenType.KW_READONLY},
             {"ReDim", TokenType.KW_REDIM},
             {"RemoveHandler", TokenType.KW_REMOVEHANDLER},
             {"Resume", TokenType.KW_RESUME},
             {"Return", TokenType.KW_RETURN},
             {"SByte", TokenType.KW_SBYTE},
             {"Select", TokenType.KW_SELECT},
             {"Set", TokenType.KW_SET},
             {"Shadows", TokenType.KW_SHADOWS},
             {"Shared", TokenType.KW_SHARED},
             {"Short", TokenType.KW_SHORT},
             {"Single", TokenType.KW_SINGLE},
             {"Static", TokenType.KW_STATIC},
             {"Step", TokenType.KW_STEP},
             {"Stop", TokenType.KW_STOP},
             {"String", TokenType.KW_STRING},
             {"Structure", TokenType.KW_STRUCTURE},
             {"Sub", TokenType.KW_SUB},
             {"SyncLock", TokenType.KW_SYNCLOCK},
             {"Then", TokenType.KW_THEN},
             {"Throw", TokenType.KW_THROW},
             {"To", TokenType.KW_TO},
             {"True", TokenType.KW_TRUE},
             {"Try", TokenType.KW_TRY},
             {"TryCast", TokenType.KW_TRYCAST},
             {"TypeOf", TokenType.KW_TYPEOF},
             {"UInteger", TokenType.KW_UINTEGER},
             {"ULong", TokenType.KW_CULNG},
             {"UShort", TokenType.KW_USHORT},
             {"Using", TokenType.KW_USING},
             {"Variant", TokenType.KW_VARIANT},
             {"Wend", TokenType.KW_WEND},
             {"When", TokenType.KW_WHEN},
             {"While", TokenType.KW_WHILE},
             {"Widening", TokenType.KW_WIDENING},
             {"With", TokenType.KW_WITH},
             {"WithEvents", TokenType.KW_WITHEVENTS},
             {"WriteOnly", TokenType.KW_WRITEONLY},
             {"Xor", TokenType.KW_XOR}}

        'Comment/line/quote terminators
        mStartOfCommentChars = New HashSet(Of Char) From {"'"c, ChrW(&H2018), ChrW(&H2019)}
        mLineTerminatorChars = New HashSet(Of Char) From {CarriageReturnChar, LineFeedChar, LineSeparatorChar, ParagraphSeparatorChar}
        mDoubleQuoteChars = New HashSet(Of Char) From {""""c, ChrW(&H201C), ChrW(&H201D)}

        'Hex/octal digits
        mOctalDigits = New HashSet(Of Char) From {"0"c, "1"c, "2"c, "3"c, "4"c, "5"c, "6"c, "7"c, "8"c, "9"c}
        mHexDigits = New HashSet(Of Char)(New CaseInsensitiveCharComparer()) From {"A"c, "B"c, "C"c, "D"c, "E"c, "F"c}
        mHexDigits.UnionWith(mOctalDigits)

        'Identifier character categories
        mAlphaCharUnicodeCategories = New HashSet(Of UnicodeCategory) From {UnicodeCategory.UppercaseLetter,
                                                                            UnicodeCategory.LowercaseLetter,
                                                                            UnicodeCategory.TitlecaseLetter,
                                                                            UnicodeCategory.ModifierLetter,
                                                                            UnicodeCategory.OtherLetter,
                                                                            UnicodeCategory.LetterNumber}
        mIdentifierCharUnicodeCategories = New HashSet(Of UnicodeCategory) From {UnicodeCategory.NonSpacingMark,
                                                                                 UnicodeCategory.SpacingCombiningMark,
                                                                                 UnicodeCategory.Format,
                                                                                 UnicodeCategory.ConnectorPunctuation,
                                                                                 UnicodeCategory.DecimalDigitNumber}
        mIdentifierCharUnicodeCategories.UnionWith(mAlphaCharUnicodeCategories)

        'Create the input buffer
        If mEncoding IsNot Nothing Then
            mInputBuffer = New LexerInputBuffer(New StreamReader(mCurrentFile.FileName, mEncoding))
        Else
            mInputBuffer = New LexerInputBuffer(New StreamReader(mCurrentFile.FileName))
        End If

        'Move to the start of the first token
        SkipWhitespaceAndComments(True)

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

        SkipWhitespaceAndComments(False)

        If mInputBuffer.EndOfBuffer Then
            Return CreateToken(Nothing, TokenType.TT_END_OF_FILE)
        End If

        Dim tokenStartLocation As Location = CurrentLocation

        Select Case CurrentCharacter

            Case CarriageReturnChar, LineFeedChar, LineSeparatorChar, ParagraphSeparatorChar
                SkipWhitespaceAndComments(True)
                Return CreateToken(tokenStartLocation, TokenType.TT_END_OF_STATEMENT)

            Case "+"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(tokenStartLocation, TokenType.OP_PLUS_EQ)
                Else
                    Return CreateToken(tokenStartLocation, TokenType.OP_PLUS)
                End If

            Case "-"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(tokenStartLocation, TokenType.OP_SUBTRACT_EQ)
                Else
                    Return CreateToken(tokenStartLocation, TokenType.OP_SUBTRACT)
                End If

            Case "*"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(tokenStartLocation, TokenType.OP_MULTIPLY_EQ)
                Else
                    Return CreateToken(tokenStartLocation, TokenType.OP_MULTIPLY)
                End If

            Case "/"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(tokenStartLocation, TokenType.OP_DIVISION_EQ)
                Else
                    Return CreateToken(tokenStartLocation, TokenType.OP_DIVISION)
                End If

            Case "\"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(tokenStartLocation, TokenType.OP_INT_DIVISION_EQ)
                Else
                    Return CreateToken(tokenStartLocation, TokenType.OP_INT_DIVISION)
                End If

            Case "&"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(tokenStartLocation, TokenType.OP_CONCAT_EQ)
                ElseIf Char.ToUpperInvariant(CurrentCharacter) = "H"c Then

                    NextCharacter()
                    Return LexHexIntegralLiteral(tokenStartLocation)

                ElseIf Char.ToUpperInvariant(CurrentCharacter) = "O"c Then

                    NextCharacter()
                    Return LexOctalIntegralToken(tokenStartLocation)

                Else
                    Return CreateToken(tokenStartLocation, TokenType.OP_CONCAT)
                End If

            Case "^"c

                NextCharacter()

                If CurrentCharacter = "=" Then
                    NextCharacter()
                    Return CreateToken(tokenStartLocation, TokenType.OP_EXPONENTIAL_EQ)
                Else
                    Return CreateToken(tokenStartLocation, TokenType.OP_EXPONENTIAL)
                End If

            Case "<"c

                'Consume the <
                NextCharacter()

                'May be one of <, <<, <=, <<=, <>
                Dim currChar As Char = CurrentCharacter

                If currChar = ">"c Then
                    NextCharacter()
                    Return CreateToken(tokenStartLocation, TokenType.OP_NOT_EQUAL)
                ElseIf currChar = "<" Then

                    'Consume the second < and peek the following character
                    NextCharacter()
                    currChar = CurrentCharacter

                    If currChar = "=" Then

                        'Consume the =
                        NextCharacter()
                        Return CreateToken(tokenStartLocation, TokenType.OP_LSHIFT_EQ)

                    Else
                        Return CreateToken(tokenStartLocation, TokenType.OP_LSHIFT)
                    End If

                ElseIf currChar = "=" Then
                    NextCharacter()
                    Return CreateToken(tokenStartLocation, TokenType.OP_LESS_THAN_EQ)
                Else
                    Return CreateToken(tokenStartLocation, TokenType.OP_LESS_THAN)
                End If

            Case ">"c

                'Consume the >
                NextCharacter()

                If CurrentCharacter = ">" Then

                    'Consume the second >
                    NextCharacter()

                    If CurrentCharacter = "=" Then
                        NextCharacter()
                        Return CreateToken(tokenStartLocation, TokenType.OP_RSHIFT_EQ)
                    Else
                        Return CreateToken(tokenStartLocation, TokenType.OP_RSHIFT)
                    End If

                ElseIf CurrentCharacter = "=" Then

                    NextCharacter()
                    Return CreateToken(tokenStartLocation, TokenType.OP_GREATER_THAN_EQ)

                Else
                    Return CreateToken(tokenStartLocation, TokenType.OP_GREATER_THAN)
                End If

            Case "="c
                NextCharacter()
                Return CreateToken(tokenStartLocation, TokenType.OP_EQUALS, Nothing)

            Case "("c
                NextCharacter()
                Return CreateToken(tokenStartLocation, TokenType.SEP_LPAREN)

            Case ")"c
                NextCharacter()
                Return CreateToken(tokenStartLocation, TokenType.SEP_RPAREN)

            Case "{"c
                NextCharacter()
                Return CreateToken(tokenStartLocation, TokenType.SEP_LBRACE)

            Case "}"c
                NextCharacter()
                Return CreateToken(tokenStartLocation, TokenType.SEP_RBRACE)

            Case "!"c
                NextCharacter()
                Return CreateToken(tokenStartLocation, TokenType.SEP_EXPMARK)

            Case "#"c

                'TODO: Date literals, directives
                NextCharacter()
                Return CreateToken(tokenStartLocation, TokenType.SEP_HASH)

            Case ","c
                NextCharacter()
                Return CreateToken(tokenStartLocation, TokenType.SEP_COMMA)

            Case "."c

                'If followed by a digit this is a number, otherwise it's a token by itself
                If Char.IsDigit(PeekCharacter()) Then

                    Dim floatingPointLiteral As New StringBuilder
                    floatingPointLiteral.Append(CurrentCharacter)
                    NextCharacter()

                    LexNumericDigitsSequence(floatingPointLiteral)

                    'See if we have an exponent
                    If Char.ToUpperInvariant(CurrentCharacter) = "E"c Then
                        MaybeLexExponent(floatingPointLiteral, tokenStartLocation)
                    End If

                    'Handle any type character
                    Dim typeChar = MaybeLexFloatingPointTypeCharacter()
                    Return CreateFloatingPointLiteralToken(tokenStartLocation, floatingPointLiteral.ToString(), typeChar)

                Else
                    NextCharacter()
                    Return CreateToken(tokenStartLocation, TokenType.SEP_DOT)
                End If

            Case ":"c
                NextCharacter()
                Return CreateToken(tokenStartLocation, TokenType.SEP_COLON)

            Case "?"c
                NextCharacter()
                Return CreateToken(tokenStartLocation, TokenType.SEP_QMARK)

            Case "0"c To "9"c
                Return LexIntegralOrFloatingPointNumber(tokenStartLocation)

            Case """"c, ChrW(&H201C), ChrW(&H201D)

                'Advance past the opening quote
                NextCharacter()

                Dim literal As New StringBuilder

                While True

                    'A double quote immediately followed by a double quote is an escaped double quote
                    If mDoubleQuoteChars.Contains(CurrentCharacter) Then

                        If mDoubleQuoteChars.Contains(PeekCharacter()) Then
                            literal.Append(CurrentCharacter)
                            SkipCharacters(2)
                        Else
                            NextCharacter()
                            Exit While
                        End If

                    ElseIf mLineTerminatorChars.Contains(CurrentCharacter) Then

                        'Unterminated string literal
                        mDiagnosticsMngr.EmitError(30648, tokenStartLocation)
                        Exit While

                    Else
                        literal.Append(CurrentCharacter)
                        NextCharacter()
                    End If

                End While

                Return CreateToken(tokenStartLocation, TokenType.LIT_STRING, literal.ToString())

            Case "["c

                NextCharacter()

                If Not IsValidAlphaCharacter(CurrentCharacter) Then
                    mDiagnosticsMngr.EmitError(30203, CurrentLocation)
                Else

                    Dim escapedIdentifier As New StringBuilder()

                    While IsValidAlphaCharacter(CurrentCharacter)
                        escapedIdentifier.Append(CurrentCharacter)
                        NextCharacter()
                    End While

                    If CurrentCharacter <> "]"c Then
                        mDiagnosticsMngr.EmitError(30034, CurrentLocation)
                    Else
                        NextCharacter()
                    End If

                    Return CreateToken(tokenStartLocation, TokenType.TT_IDENTIFIER, escapedIdentifier.ToString())

                End If

            Case Else

                'See if this character can start an identifier
                If IsValidAlphaCharacter(CurrentCharacter) _
                    OrElse (CurrentCharacter = "_" _
                            AndAlso IsValidAlphaCharacter(PeekCharacter())) Then

                    Dim identifier As New StringBuilder()
                    identifier.Append(CurrentCharacter)

                    'Skip over the initial character which has been added
                    Dim ch = NextCharacter()

                    'Keep going whilst we have a valid identifier character
                    While IsValidIdentifierCharacter(ch)
                        identifier.Append(ch)
                        ch = NextCharacter()
                    End While

                    'Check if it's a keyword
                    Dim completeIdentifier As String = identifier.ToString()

                    If mKeywords.ContainsKey(completeIdentifier) Then
                        Return CreateToken(tokenStartLocation, mKeywords(completeIdentifier))
                    Else
                        Return CreateToken(tokenStartLocation, TokenType.TT_IDENTIFIER, identifier.ToString())
                    End If

                Else

                    mDiagnosticsMngr.EmitError(30037, CurrentLocation)

                    NextCharacter()
                    Return NextToken()

                End If


        End Select

        Return Nothing

    End Function

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Lexes and creates an integral or floating-point number.
    ''' </summary>
    ''' <param name="tokenLocation">Location of the token.</param>
    ''' <returns>An integral or floating-point literal token.</returns>
    ''' <remarks></remarks>
    Private Function LexIntegralOrFloatingPointNumber(tokenLocation As Location) As Token

        Dim tokenValue As New StringBuilder()
        LexNumericDigitsSequence(tokenValue)

        If Char.ToUpperInvariant(CurrentCharacter) = "E" Then

            MaybeLexExponent(tokenValue, tokenLocation)

            Dim typeChar = MaybeLexFloatingPointTypeCharacter()
            Return CreateFloatingPointLiteralToken(tokenLocation, tokenValue.ToString(), typeChar)

        ElseIf CurrentCharacter = "."c _
               AndAlso Char.IsDigit(PeekCharacter()) Then

            'Consume the . and the digit sequence
            tokenValue.Append(CurrentCharacter)
            NextCharacter()
            LexNumericDigitsSequence(tokenValue)

            'Handle any exponent
            If Char.ToUpperInvariant(CurrentCharacter) = "E" Then
                MaybeLexExponent(tokenValue, tokenLocation)
            End If

            Dim typeChar = MaybeLexFloatingPointTypeCharacter()
            Return CreateFloatingPointLiteralToken(tokenLocation, tokenValue.ToString(), typeChar)

        Else

            'Plain ol' integer
            Dim typeCharacter = MaybeLexIntegralTypeCharacter()
            Dim literalValue As BigInteger = BigInteger.Parse(tokenValue.ToString(), CultureInfo.InvariantCulture)

            Return CreateIntegralLiteralToken(tokenLocation, literalValue, typeCharacter)

        End If

    End Function

    ''' <summary>
    ''' Lexes a sequence of numeric digits
    ''' </summary>
    ''' <param name="literal">The literal value to append each character to.</param>
    ''' <remarks></remarks>
    Private Sub LexNumericDigitsSequence(tokenValue As StringBuilder)

        While Char.IsDigit(CurrentCharacter)
            tokenValue.Append(CurrentCharacter)
            NextCharacter()
        End While

    End Sub

    ''' <summary>
    ''' Handles the lexing and validation of an exponent character sequence for a floating point literal.
    ''' </summary>
    ''' <param name="literal">Current literal being lexed the exponent belongs to.</param>
    ''' <param name="location">Location of the literal the exponent belongs to.</param>
    ''' <remarks>
    ''' Exponent  ::=  E  [  Sign  ]  IntLiteral
    ''' Sign  ::=  +  |  -
    ''' </remarks>
    Private Sub MaybeLexExponent(literal As StringBuilder, location As Location)

        'Append the E character
        Debug.Assert(Char.ToUpperInvariant(CurrentCharacter) = "E")

        literal.Append(CurrentCharacter)
        NextCharacter()

        'If the sign is missing and is followed by a digit, the sign defaults to +
        If CurrentCharacter <> "-"c _
           AndAlso CurrentCharacter <> "+" Then

            If Not Char.IsDigit(PeekCharacter()) Then
                mDiagnosticsMngr.EmitError(30495, location)
            End If

            literal.Append("+"c)

        End If

        'Append and consume the sign character
        literal.Append(CurrentCharacter)
        NextCharacter()

        'Should be at least one digit following the sign. If not then append a 1 as an error recovery measure
        If Not Char.IsDigit(CurrentCharacter) Then
            mDiagnosticsMngr.EmitError(30495, location)
            literal.Append("1"c)
        End If

        LexNumericDigitsSequence(literal)

    End Sub

    ''' <summary>
    ''' Lexes any floating point literal type character at the current position.
    ''' </summary>
    ''' <returns>The FP literal type character (if any) at the current position.</returns>
    ''' <remarks></remarks>
    Private Function MaybeLexFloatingPointTypeCharacter() As FloatingPointLiteralTypeCharacter

        Select Case Char.ToUpperInvariant(CurrentCharacter)

            Case "!"c, "f"c, "F"c

                If Char.IsWhiteSpace(PeekCharacter()) Then
                    NextCharacter()
                    Return FloatingPointLiteralTypeCharacter.SingleTypeChar
                Else
                    Return FloatingPointLiteralTypeCharacter.None
                End If

            Case "#"c, "r"c, "R"c

                If Char.IsWhiteSpace(PeekCharacter()) Then
                    NextCharacter()
                    Return FloatingPointLiteralTypeCharacter.DoubleTypeChar
                Else
                    Return FloatingPointLiteralTypeCharacter.None
                End If

            Case "@"c, "d"c, "D"c

                If Char.IsWhiteSpace(PeekCharacter()) Then
                    NextCharacter()
                    Return FloatingPointLiteralTypeCharacter.DecimalTypeChar
                Else
                    Return FloatingPointLiteralTypeCharacter.None
                End If

        End Select

        Return FloatingPointLiteralTypeCharacter.None

    End Function

    ''' <summary>
    ''' Lexes any integral literal type character at the current position.
    ''' </summary>
    ''' <returns>The integral literal type character (if any) at the current position.</returns>
    ''' <remarks></remarks>
    Private Function MaybeLexIntegralTypeCharacter() As IntegralLiteralTypeCharacter

        Select Case Char.ToUpperInvariant(CurrentCharacter)

            Case "S"c

                If Char.IsWhiteSpace(PeekCharacter()) Then
                    NextCharacter()
                    Return IntegralLiteralTypeCharacter.ShortTypeChar
                Else
                    Return IntegralLiteralTypeCharacter.None
                End If

            Case "U"c

                'US, UI, UL
                Dim peekedChar As Char = Char.ToUpperInvariant(PeekCharacter(1))
                Dim peekedSecondChar As Char = PeekCharacter(2)

                If Char.IsWhiteSpace(peekedSecondChar) Then

                    If peekedChar = "S"c Then
                        SkipCharacters(2)
                        Return IntegralLiteralTypeCharacter.UnsignedShortTypeChar
                    ElseIf peekedChar = "I"c Then
                        SkipCharacters(2)
                        Return IntegralLiteralTypeCharacter.UnsignedIntTypeChar
                    ElseIf peekedChar = "L" Then
                        SkipCharacters(2)
                        Return IntegralLiteralTypeCharacter.UnsignedLongTypeChar
                    End If

                Else
                    Return IntegralLiteralTypeCharacter.None
                End If

            Case "%"c, "I"c

                If Char.IsWhiteSpace(PeekCharacter()) Then
                    NextCharacter()
                    Return IntegralLiteralTypeCharacter.IntTypeChar
                Else
                    Return IntegralLiteralTypeCharacter.None
                End If

            Case "L"c, "&"c

                If Char.IsWhiteSpace(PeekCharacter()) Then
                    NextCharacter()
                    Return IntegralLiteralTypeCharacter.LongTypeChar
                Else
                    Return IntegralLiteralTypeCharacter.None
                End If

            Case "@"c, "D"c, "d"c

                If Char.IsWhiteSpace(PeekCharacter()) Then
                    NextCharacter()
                    Return IntegralLiteralTypeCharacter.DecimalTypeChar
                End If

        End Select

        Return IntegralLiteralTypeCharacter.None

    End Function

    ''' <summary>
    ''' Indicates if the end of all input files has been reached.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function EndOfInput() As Boolean
        Return mInputBuffer.EndOfBuffer
    End Function

    ''' <summary>
    ''' Lexes an integral literal represented in octal form.
    ''' </summary>
    ''' <param name="literalLocation">Start location of the literal.</param>
    ''' <returns>An integral literal token.</returns>
    ''' <remarks></remarks>
    Private Function LexOctalIntegralToken(literalLocation As Location) As Token

        Dim octalLiteral As New StringBuilder

        While mOctalDigits.Contains(CurrentCharacter)
            octalLiteral.Append(CurrentCharacter)
            NextCharacter()
        End While

        Dim typeCharacter As IntegralLiteralTypeCharacter = MaybeLexIntegralTypeCharacter()

        'Convert the octal value to decimal
        Dim zeroCharASCIIVal As Integer = Asc("0"c)
        Dim literalValue As New BigInteger(Asc(octalLiteral(0)) - zeroCharASCIIVal)

        For digitPos As Integer = 1 To octalLiteral.Length - 1
            literalValue = (literalValue * 8) + Asc(octalLiteral(digitPos)) - zeroCharASCIIVal
        Next

        Return CreateIntegralLiteralToken(literalLocation, literalValue, typeCharacter)

    End Function

    ''' <summary>
    ''' Lexes a hex integral literal token.
    ''' </summary>
    ''' <param name="literalLocation">Start location of the token.</param>
    ''' <returns>An integral literal token.</returns>
    ''' <remarks></remarks>
    Private Function LexHexIntegralLiteral(literalLocation As Location) As Token

        Dim hexliteral As New StringBuilder

        While mHexDigits.Contains(CurrentCharacter)
            hexliteral.Append(CurrentCharacter)
            NextCharacter()
        End While

        'Handle any type character
        Dim typeCharacter As IntegralLiteralTypeCharacter = MaybeLexIntegralTypeCharacter()

        'Parse the literal into a numeric value
        Dim literalValue As BigInteger = BigInteger.Parse(hexliteral.ToString(),
                                                          NumberStyles.HexNumber,
                                                          CultureInfo.InvariantCulture)
        Return CreateIntegralLiteralToken(literalLocation, literalValue, typeCharacter)

    End Function

    ''' <summary>
    ''' Creates an floating-point literal token, the type of which depends on the value of the literal or 
    ''' any type specifier applied.
    ''' </summary>
    ''' <param name="location"><Location of the token./param>
    ''' <param name="value">Lexed value of the token.</param>
    ''' <param name="typeCharacter">Any type specified for the value to dictate the literal type.</param>
    ''' <returns>A new floating-point literal token.</returns>
    ''' <remarks></remarks>
    Private Function CreateFloatingPointLiteralToken(location As Location,
                                                     value As String,
                                                     typeCharacter As FloatingPointLiteralTypeCharacter) As Token

        Dim literalValue As Double

        Try
            literalValue = Double.Parse(value, NumberStyles.Float, CultureInfo.InvariantCulture)
        Catch ex As OverflowException

            mDiagnosticsMngr.EmitError(30036, location)
            literalValue = Double.MaxValue

        End Try

        '2.4.3 - By default, a floating-point literal is of type Double. If the Single, Double, or Decimal type 
        'character is specified, the literal is of that type. If a floating-point literal's type is of insufficient 
        'size to hold the floating-point literal, a compile-time error results.
        Dim literalTokenType As TokenType = TokenType.LIT_DOUBLE
        Dim maxTokenValue As Double = Double.MaxValue

        If typeCharacter = FloatingPointLiteralTypeCharacter.SingleTypeChar Then

            literalTokenType = TokenType.LIT_SINGLE
            maxTokenValue = Single.MaxValue

        ElseIf typeCharacter = FloatingPointLiteralTypeCharacter.DecimalTypeChar Then

            literalTokenType = TokenType.LIT_DECIMAL
            maxTokenValue = Decimal.MaxValue

        ElseIf typeCharacter <> FloatingPointLiteralTypeCharacter.None Then
            Debug.Assert(typeCharacter = FloatingPointLiteralTypeCharacter.DoubleTypeChar)
        End If

        If literalValue > maxTokenValue Then
            mDiagnosticsMngr.EmitError(30036, location)
            literalValue = maxTokenValue
        End If

        Return New Token(location, literalTokenType, literalValue)

    End Function

    ''' <summary>
    ''' Creates an integral literal token, the type of which depends on the value of the literal or 
    ''' any type specifier applied.
    ''' </summary>
    ''' <param name="location"><Location of the token./param>
    ''' <param name="value">Lexed value of the token.</param>
    ''' <param name="typeCharacter">Any type specified for the value to dictate the literal type.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CreateIntegralLiteralToken(location As Location,
                                                value As BigInteger,
                                                typeCharacter As IntegralLiteralTypeCharacter) As Token

        Dim literalTokenType As TokenType

        If typeCharacter = IntegralLiteralTypeCharacter.None Then

            '2.4.2 - If no type character is specified, values in the range of the Integer type are typed as Integer; values
            'outside the range for Integer are typed as Long. If an integer literal's type is of insufficient size to
            'hold the integer literal, a compile-time error results.
            If value <= Integer.MaxValue Then
                typeCharacter = IntegralLiteralTypeCharacter.IntTypeChar
            Else
                typeCharacter = IntegralLiteralTypeCharacter.LongTypeChar
            End If

        End If

        'Determine the max allowed value of the literal based on the type.
        Dim maxValue As BigInteger

        Select Case typeCharacter

            Case IntegralLiteralTypeCharacter.ShortTypeChar
                maxValue = New BigInteger(Short.MaxValue)
                literalTokenType = TokenType.LIT_SHORT

            Case IntegralLiteralTypeCharacter.UnsignedShortTypeChar
                maxValue = New BigInteger(UShort.MaxValue)
                literalTokenType = TokenType.LIT_USHORT

            Case IntegralLiteralTypeCharacter.IntTypeChar
                maxValue = New BigInteger(Integer.MaxValue)
                literalTokenType = TokenType.LIT_INTEGER

            Case IntegralLiteralTypeCharacter.UnsignedIntTypeChar
                maxValue = New BigInteger(UInt32.MaxValue)
                literalTokenType = TokenType.LIT_UINTEGER

            Case IntegralLiteralTypeCharacter.LongTypeChar
                maxValue = New BigInteger(Long.MaxValue)
                literalTokenType = TokenType.LIT_LONG

            Case IntegralLiteralTypeCharacter.UnsignedLongTypeChar
                maxValue = New BigInteger(ULong.MaxValue)
                literalTokenType = TokenType.KW_ULONG

            Case IntegralLiteralTypeCharacter.DecimalTypeChar
                maxValue = New BigInteger(Decimal.MaxValue)
                literalTokenType = TokenType.LIT_DECIMAL

            Case Else
                Throw New InvalidOperationException("Unknown literal type character")

        End Select

        'If the value overflows the max allowed, clamp the literal to the max allowed for error recovery.
        If value > maxValue Then
            mDiagnosticsMngr.EmitError(30036, location)
            value = maxValue
        End If

        Return New Token(location, literalTokenType, value)

    End Function

    ''' <summary>
    ''' Helper method to create a token given its type and value.
    ''' </summary>
    ''' <param name="location">Location of the token.</param>
    ''' <param name="tokType">Type of the token.</param>
    ''' <param name="value">Value of the token.</param>
    ''' <returns>A new Token instance initialised with the specified type and value.</returns>
    ''' <remarks></remarks>
    Private Function CreateToken(location As Location, tokType As TokenType, Optional value As Object = Nothing) As Token
        Return New Token(location, tokType, value)
    End Function

    ''' <summary>
    ''' Skips over and whitespace and comment characters.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SkipWhitespaceAndComments(skipNewLines As Boolean)

        While Not EndOfInput()

            Dim ch As Char = CurrentCharacter

            If Char.IsWhiteSpace(ch) Then

                If mLineTerminatorChars.Contains(ch) Then

                    If Not skipNewLines Then
                        Exit Sub
                    End If

                    SkipNewLineCharacters(ch)

                Else
                    NextCharacter()
                End If

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
        mColumnNumber = 1

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
    ''' Consumes the required number of characters from the current position.
    ''' </summary>
    ''' <param name="count">Number of characters to consume.</param>
    ''' <remarks></remarks>
    Private Sub SkipCharacters(count As Integer)

        For idx As Integer = 1 To count
            NextCharacter()
        Next

    End Sub

    ''' <summary>
    ''' Peeks a single character from the input buffer.
    ''' </summary>
    ''' <param name="lookAhead">Number of characters to peek ahead.</param>
    ''' <returns>The peeked character.</returns>
    ''' <remarks></remarks>
    Private Function PeekCharacter(Optional lookAhead As Integer = 1) As Char
        Return mInputBuffer.PeekCharacter(lookAhead)
    End Function

    ''' <summary>
    ''' Determines if the specified character is a valid identifier character.
    ''' </summary>
    ''' <param name="ch"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsValidIdentifierCharacter(ch As Char) As Boolean

        Dim category As UnicodeCategory = Char.GetUnicodeCategory(ch)
        Return mIdentifierCharUnicodeCategories.Contains(category)

    End Function

    ''' <summary>
    ''' Determines if the specified character is a valid alpha character.
    ''' </summary>
    ''' <param name="ch"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsValidAlphaCharacter(ch As Char) As Boolean

        Dim category As UnicodeCategory = Char.GetUnicodeCategory(ch)
        Return mAlphaCharUnicodeCategories.Contains(category)

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

    ''' <summary>
    ''' Accessor for the current location.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private ReadOnly Property CurrentLocation As Location
        Get
            Return New Location(mCurrentFile, mLineNumber, mColumnNumber)
        End Get
    End Property

#End Region

End Class