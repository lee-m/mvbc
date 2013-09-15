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
''' Represents a single token lexed from an input file.
''' </summary>
''' <remarks></remarks>
Public Class Token

    ''' <summary>
    ''' Location of the token.
    ''' </summary>
    ''' <remarks></remarks>
    Private mLocation As Location

    ''' <summary>
    ''' The type of the token.
    ''' </summary>
    ''' <remarks></remarks>
    Private mTokenType As TokenType

    ''' <summary>
    ''' The actual value of this token.
    ''' </summary>
    ''' <remarks></remarks>
    Private mValue As Object

#Region "Public Methods"

    ''' <summary>
    ''' Creates a new Token instance initialised with the specified values.
    ''' </summary>
    ''' <param name="location">Location of the token.</param>
    ''' <param name="tokType">The type of token.</param>
    ''' <param name="value">The actual textual value of the token as lexed from the source file.</param>
    ''' <remarks></remarks>
    Public Sub New(location As Location,
                   tokType As TokenType,
                   value As Object)
        mLocation = location
        mTokenType = TokType
        mValue = Value
    End Sub

#End Region

#Region "Properties"

    ''' <summary>
    ''' Accessor for the location of the token.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Location As Location
        Get
            Return mLocation
        End Get
    End Property

    ''' <summary>
    ''' Accessor for the type of the token.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property TokType As TokenType
        Get
            Return mTokenType
        End Get
    End Property

    ''' <summary>
    ''' Accessor for the token's value.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Value As Object
        Get
            Return mValue
        End Get
    End Property

#End Region

End Class

#Region "TokenType Enumeration"

''' <summary>
''' Enumeration of potential token types.
''' </summary>
''' <remarks></remarks>
Public Enum TokenType

    'Keywords
    KW_ADDHANDLER
    KW_ADDRESSOF
    KW_ALIAS
    KW_AND
    KW_ANDALSO
    KW_AS
    KW_BOOLEAN
    KW_BYREF
    KW_BYTE
    KW_BYVAL
    KW_CALL
    KW_CASE
    KW_CATCH
    KW_CBOOL
    KW_CBYTE
    KW_CCHAR
    KW_CDATE
    KW_CDBL
    KW_CDEC
    KW_CHAR
    KW_CINT
    KW_CLASS
    KW_CLNG
    KW_COBJ
    KW_CONST
    KW_CONTINUE
    KW_CSBYTE
    KW_CSHORT
    KW_CSNG
    KW_CSTR
    KW_CTYPE
    KW_CUINT
    KW_CULNG
    KW_CUSHORT
    KW_DATE
    KW_DECIMAL
    KW_DECLARE
    KW_DEFAULT
    KW_DELEGATE
    KW_DIM
    KW_DIRECTCAST
    KW_DO
    KW_DOUBLE
    KW_EACH
    KW_ELSE
    KW_ELSEIF
    KW_END
    KW_ENDIF
    KW_ENUM
    KW_ERASE
    KW_ERROR
    KW_EVENT
    KW_EXIT
    KW_FALSE
    KW_FINALLY
    KW_FOR
    KW_FRIEND
    KW_FUNCTION
    KW_GET
    KW_GETTYPE
    KW_GETXMLNAMESPACE
    KW_GLOBAL
    KW_GOSUB
    KW_GOTO
    KW_HANDLES
    KW_IF
    KW_IMPLEMENTS
    KW_IMPORTS
    KW_IN
    KW_INHERITS
    KW_INTEGER
    KW_INTERFACE
    KW_IS
    KW_ISNOT
    KW_LET
    KW_LIB
    KW_LIKE
    KW_LONG
    KW_LOOP
    KW_ME
    KW_MOD
    KW_MODULE
    KW_MUSTINHERIT
    KW_MUSTOVERRIDE
    KW_MYBASE
    KW_MYCLASS
    KW_NAMESPACE
    KW_NARROWING
    KW_NEW
    KW_NEXT
    KW_NOT
    KW_NOTHING
    KW_NOTINHERITABLE
    KW_NOTOVERRIDABLE
    KW_OBJECT
    KW_OF
    KW_ON
    KW_OPERATOR
    KW_OPTION
    KW_OPTIONAL
    KW_OR
    KW_ORELSE
    KW_OVERLOADS
    KW_OVERRIDABLE
    KW_OVERRIDES
    KW_PARAMARRAY
    KW_PARTIAL
    KW_PRIVATE
    KW_PROPERTY
    KW_PROTECTED
    KW_PUBLIC
    KW_RAISEEVENT
    KW_READONLY
    KW_REDIM
    KW_REMOVEHANDLER
    KW_RESUME
    KW_RETURN
    KW_SBYTE
    KW_SELECT
    KW_SET
    KW_SHADOWS
    KW_SHARED
    KW_SHORT
    KW_SINGLE
    KW_STATIC
    KW_STEP
    KW_STOP
    KW_STRING
    KW_STRUCTURE
    KW_SUB
    KW_SYNCLOCK
    KW_THEN
    KW_THROW
    KW_TO
    KW_TRUE
    KW_TRY
    KW_TRYCAST
    KW_TYPEOF
    KW_UINTEGER
    KW_ULONG
    KW_USHORT
    KW_USING
    KW_VARIANT
    KW_WEND
    KW_WHEN
    KW_WHILE
    KW_WIDENING
    KW_WITH
    KW_WITHEVENTS
    KW_WRITEONLY
    KW_XOR

    'Operators
    OP_PLUS
    OP_PLUS_EQ
    OP_SUBTRACT
    OP_SUBTRACT_EQ
    OP_MULTIPLY
    OP_MULTIPLY_EQ
    OP_DIVISION
    OP_DIVISION_EQ
    OP_INT_DIVISION
    OP_INT_DIVISION_EQ
    OP_CONCAT
    OP_CONCAT_EQ
    OP_LIKE
    OP_MOD
    OP_AND
    OP_OR
    OP_XOR
    OP_EXPONENTIAL
    OP_EXPONENTIAL_EQ
    OP_LSHIFT
    OP_LSHIFT_EQ
    OP_RSHIFT
    OP_RSHIFT_EQ
    OP_EQUALS
    OP_NOT_EQUAL
    OP_LESS_THAN
    OP_GREATER_THAN
    OP_LESS_THAN_EQ
    OP_GREATER_THAN_EQ
    OP_NOT
    OP_ISTRUE
    OP_ISFALSE
    OP_CTYPE

    'Separators
    SEP_LPAREN
    SEP_RPAREN
    SEP_LBRACE
    SEP_RBRACE
    SEP_EXPMARK
    SEP_QMARK
    SEP_HASH
    SEP_COMMA
    SEP_DOT
    SEP_COLON

    'Literals
    LIT_SHORT
    LIT_USHORT
    LIT_INTEGER
    LIT_UINTEGER
    LIT_LONG
    LIT_ULONG
    LIT_DECIMAL
    LIT_SINGLE
    LIT_DOUBLE
    LIT_STRING
    LIT_CHAR

    'Miscellaneous
    TT_IDENTIFIER
    TT_END_OF_STATEMENT
    TT_END_OF_FILE

End Enum

#End Region
