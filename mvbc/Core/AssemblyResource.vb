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
''' Represents a linked or embedded resource to include in the output assembly.
''' </summary>
''' <remarks></remarks>
Public Class AssemblyResource

    ''' <summary>
    ''' Name of the resource
    ''' </summary>
    ''' <remarks></remarks>
    Private mName As String

    ''' <summary>
    ''' Name of the file containing the resource.
    ''' </summary>
    ''' <remarks></remarks>
    Private mFileName As String

    ''' <summary>
    ''' Public/private attribute of the resource.
    ''' </summary>
    ''' <remarks></remarks>
    Private mAttributes As ResourceAttributes

    ''' <summary>
    ''' Whether the resource is embedded or not.
    ''' </summary>
    ''' <remarks></remarks>
    Private mEmbedded As Boolean

    ''' <summary>
    ''' Initialises this instance with the specified field values.
    ''' </summary>
    ''' <param name="name">Name of the resource.</param>
    ''' <param name="fileName">Name of the file containing the resource.</param>
    ''' <param name="attributes">Public/private attributes.</param>
    ''' <param name="embedded">Whether the resource is embedded or not.</param>
    ''' <remarks></remarks>
    Public Sub New(name As String, fileName As String, attributes As ResourceAttributes, embedded As Boolean)
        mName = name
        mFileName = fileName
        mAttributes = attributes
        mEmbedded = embedded
    End Sub

    ''' <summary>
    ''' Accessor for the name of the file.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Name As String
        Get
            Return mName
        End Get
    End Property

    ''' <summary>
    ''' Accessor for the filename of the resource.
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
    ''' Accessor for the public/private attribute of the resource.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Attributes As ResourceAttributes
        Get
            Return mAttributes
        End Get
    End Property

    ''' <summary>
    ''' Whether the resource is embedded or not.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property Embedded As Boolean
        Get
            Return mEmbedded
        End Get
    End Property

End Class
