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
''' Wrapper around an input stream used by the Lexer.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class LexerInputBuffer
    Implements IDisposable

    ''' <summary>
    ''' Handle to the source file to read the input from.
    ''' </summary>
    ''' <remarks></remarks>
    Private mReader As StreamReader

    ''' <summary>
    ''' Current position within the input file.
    ''' </summary>
    ''' <remarks></remarks>
    Private mCurrentStreamPosition As Integer

    ''' <summary>
    ''' Buffered text from the input file.
    ''' </summary>
    ''' <remarks></remarks>
    Private mBuffer(BufferSize) As Char

    ''' <summary>
    ''' Current position within the input buffer.
    ''' </summary>
    ''' <remarks></remarks>
    Private mBufferPosition As Integer

    ''' <summary>
    ''' Actual number of bytes in the buffer.
    ''' </summary>
    ''' <remarks></remarks>
    Private mActualBufferSize As Integer

    ''' <summary>
    ''' Maximum size of the buffered file contents.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const BufferSize As Integer = 64 * 1024

#Region "Public Methods"

    ''' <summary>
    ''' Create a new input buffer wrapped around the input stream.
    ''' </summary>
    ''' <param name="input"></param>
    ''' <remarks></remarks>
    Public Sub New(input As StreamReader)
        mReader = input
        RefreshBuffer()
    End Sub

    ''' <summary>
    ''' Peeks a character from the buffer without advancing the current position.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PeekCharacter() As Char

        If mBufferPosition = mBuffer.Length - 1 Then
            Throw New InvalidOperationException("Attempt to peek past end of buffer")
        End If

        Return mBuffer(mBufferPosition + 1)

    End Function

    ''' <summary>
    ''' Returns the next character and advances the current position.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function NextCharacter() As Char

        If mBufferPosition = mActualBufferSize - 1 Then

            If mReader.EndOfStream Then
                Return Chr(0)
            End If

            RefreshBuffer()
            mBufferPosition = 0

        Else
            mBufferPosition += 1
        End If

        Return mBuffer(mBufferPosition)

    End Function

    ''' <summary>
    ''' Releases the file handle.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Dispose() Implements IDisposable.Dispose
        mReader.Dispose()
        mBuffer = Nothing
    End Sub

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Refreshes the input buffer with the next chunk of data from the input file.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RefreshBuffer()
        Array.Clear(mBuffer, 0, BufferSize)
        mActualBufferSize = mReader.ReadBlock(mBuffer, mCurrentStreamPosition, BufferSize)
        mCurrentStreamPosition += mActualBufferSize
        mBufferPosition = 0
    End Sub

#End Region

#Region "Properties"

    ''' <summary>
    ''' Indicates whether the end of the buffer has been reached.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property EndOfBuffer As Boolean
        Get
            Return mBufferPosition = mActualBufferSize - 1 _
                   AndAlso mReader.EndOfStream
        End Get
    End Property

    ''' <summary>
    ''' Accessor for the character at the current position.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property CurrentCharacter As Char
        Get
            Return mBuffer(mBufferPosition)
        End Get
    End Property

#End Region

End Class
