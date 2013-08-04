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
''' Represents a single source file being compiled.
''' </summary>
''' <remarks></remarks>
Public Class SourceFile

    ''' <summary>
    ''' Name of the file this instance represents.
    ''' </summary>
    ''' <remarks></remarks>
    Private mFileName As String

    ''' <summary>
    ''' Iniitialise this instance with the specified file.
    ''' </summary>
    ''' <param name="fileName">Name of the file this instance represents.</param>
    ''' <remarks></remarks>
    Public Sub New(fileName As String)
        mFileName = fileName
    End Sub

#Region "Properties"

    ''' <summary>
    ''' Accessor for t
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property FileName As String
        Get
            Return mFileName
        End Get
    End Property

    ''' <summary>
    ''' Accessor for the Option Strict value for this file.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OptionStrict As Boolean

    ''' <summary>
    ''' Accessor for the Option Explicit value for this file.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OptionExplicit As Boolean

    ''' <summary>
    ''' Accessor for the Option Compare value for this file.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property OptionCompare As OptionCompareValues

#End Region

End Class
