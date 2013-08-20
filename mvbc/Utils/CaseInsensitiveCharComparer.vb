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
''' Case insensitive equality comparer.
''' </summary>
''' <remarks></remarks>
Public Class CaseInsensitiveCharComparer
    Implements IEqualityComparer(Of Char)

    ''' <summary>
    ''' Checks two characters for equality.
    ''' </summary>
    ''' <param name="x">First character to compare.</param>
    ''' <param name="y">Second character to compare.</param>
    ''' <returns>True if the two characters are equal.</returns>
    ''' <remarks></remarks>
    Public Function Equals1(x As Char, y As Char) As Boolean Implements IEqualityComparer(Of Char).Equals
        Return Char.ToUpperInvariant(x) = Char.ToUpperInvariant(y)
    End Function

    ''' <summary>
    ''' Calculates the hash code of the character in a case insensitive manner.
    ''' </summary>
    ''' <param name="obj">Character to hash.</param>
    ''' <returns>Character hash code.</returns>
    ''' <remarks></remarks>
    Public Function GetHashCode1(obj As Char) As Integer Implements IEqualityComparer(Of Char).GetHashCode
        Return Char.ToUpperInvariant(obj).GetHashCode()
    End Function

End Class
