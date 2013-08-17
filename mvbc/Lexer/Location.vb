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
''' Represents the location of a single token.
''' </summary>
''' <remarks></remarks>
Public Class Location

    ''' <summary>
    ''' Containing source file.
    ''' </summary>
    ''' <remarks></remarks>
    Private ReadOnly mSourceFile As SourceFile

    ''' <summary>
    ''' Line number of the token.
    ''' </summary>
    ''' <remarks></remarks>
    Private ReadOnly mLine As Integer

    ''' <summary>
    ''' Column number of the token.
    ''' </summary>
    ''' <remarks></remarks>
    Private ReadOnly mColumn As Integer

    ''' <summary>
    ''' Creates a new Location instance initialised with the specified values.
    ''' </summary>
    ''' <param name="sourceFile">Containing source file.</param>
    ''' <param name="line">Line number of the token.</param>
    ''' <param name="column">Column number of the token.</param>
    ''' <remarks></remarks>
    Public Sub New(sourceFile As SourceFile, line As Integer, column As Integer)
        mSourceFile = sourceFile
        mLine = line
        mColumn = column
    End Sub

    ''' <summary>
    ''' Accessor for the source file.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property SourceFile As SourceFile
        Get
            Return mSourceFile
        End Get
    End Property

    ''' <summary>
    ''' Accessor for the line number.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property LineNumber As Integer
        Get
            Return mLine
        End Get
    End Property

    ''' <summary>
    ''' Accessor for the column number.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property ColumnNumber As Integer
        Get
            Return mColumn
        End Get
    End Property

End Class
