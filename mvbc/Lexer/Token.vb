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
    ''' Line number in the source file this token was from.
    ''' </summary>
    ''' <remarks></remarks>
    Private mLineNumber As Integer

    ''' <summary>
    ''' Column number the first character of the token appeared on.
    ''' </summary>
    ''' <remarks></remarks>
    Private mColumnNumber As Integer

    ''' <summary>
    ''' The type of the token.
    ''' </summary>
    ''' <remarks></remarks>
    Private mTokenType As TokenType

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
        KW_REM
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
    End Enum

#End Region

End Class
