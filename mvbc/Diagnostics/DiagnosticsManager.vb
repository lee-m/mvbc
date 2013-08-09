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
''' Class to handle the reporting of errors or warnings.
''' </summary>
''' <remarks></remarks>
Public Class DiagnosticsManager

    ''' <summary>
    ''' Number of errors that have been issued.
    ''' </summary>
    ''' <remarks></remarks>
    Private mNumErrorsIssued As Integer

    ''' <summary>
    ''' Number of warnings that have been issued,
    ''' </summary>
    ''' <remarks></remarks>
    Private mNumWarningsIssued As Integer

    ''' <summary>
    ''' Resource manager used to get the message text from.
    ''' </summary>
    ''' <remarks></remarks>
    Private mResourceManager As ResourceManager

    ''' <summary>
    ''' A "stack" of whether messages should be queued or not.
    ''' </summary>
    ''' <remarks></remarks>
    Private mQueueMessageOutput As Integer

    ''' <summary>
    ''' Any queued messages.
    ''' </summary>
    ''' <remarks></remarks>
    Private mQueuedMessages As Queue(Of String)

    ''' <summary>
    ''' Set of all warning numbers.
    ''' </summary>
    ''' <remarks></remarks>
    Private mAllWarnings As HashSet(Of Integer)

    ''' <summary>
    ''' Maximum number of errors that may be reported ina a single run.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const MaximumNumberOfErrors As Integer = 100

#Region "Public Methods"

    ''' <summary>
    ''' Default constructor.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        mResourceManager = New ResourceManager("mvbc.Messages", Me.GetType().Assembly)
        mQueueMessageOutput = 0
        mQueuedMessages = New Queue(Of String)
        mAllWarnings = New HashSet(Of Integer)(New Integer() {40008, 40025, 40027, 40028, 40029, 40031, 40032, 40033, 40035, 40039,
                                                              40042, 40059, 41999, 42004, 42015, 42016, 42017, 42018, 42019, 42020,
                                                              42021, 42022, 42023, 42024, 42025, 42026, 42028, 42029, 42030, 42031,
                                                              42032, 42036, 42104, 42105, 42107, 42110, 42324, 42326, 42358})
    End Sub

    ''' <summary>
    ''' Indicates that any errors/warnings emitted after this point should be queued for output.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub EnableMessageQueuing()
        mQueueMessageOutput += 1
    End Sub

    ''' <summary>
    ''' Indicates that any errors/warnings emitted after this point should be output immediately.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub DisableMessageQueuing()
        mQueueMessageOutput -= 1
    End Sub

    ''' <summary>
    ''' Emits any queued messages.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub EmitQueuedMessages()

        mQueueMessageOutput = 0

        For Each queuedMessage As String In mQueuedMessages
            EmitMessage(queuedMessage)
        Next

    End Sub

    ''' <summary>
    ''' Dicards any queued messages.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub DiscardQueuedMessages()

        mQueueMessageOutput = 0
        mQueuedMessages.Clear()

    End Sub

    ''' <summary>
    ''' Treats all warnings as errors.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub TreatAllWarningsAsErrors()

        For Each warningID In mAllWarnings
            WarningsAsErrors.Add(warningID)
        Next

    End Sub

    ''' <summary>
    ''' Suppresses emission of all warnings
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SuppressAllWarnings()

        For Each warningID In mAllWarnings
            DisabledWarnings.Add(warningID)
        Next

    End Sub

    ''' <summary>
    ''' Determines whether a particular warning ID is valid or not.
    ''' </summary>
    ''' <param name="warningID">Warning ID to check.</param>
    ''' <returns>True if the warning ID is valid, otherwise false.</returns>
    ''' <remarks></remarks>
    Public Function IsValidWarningID(warningID As Integer) As Boolean
        Return mAllWarnings.Contains(warningID)
    End Function

    ''' <summary>
    ''' Outputs an internal compiler error.
    ''' </summary>
    ''' <param name="ex">Unhandled exception.</param>
    ''' <remarks></remarks>
    Public Sub InternalCompilerError(ex As Exception)

        'We want ICEs to be reported right away
        DisableMessageQueuing()

        mNumErrorsIssued += 1
        EmitMessage(String.Format("Internal compiler error : {0}", ex.Message))

    End Sub

    ''' <summary>
    ''' Outputs a fatal error.
    ''' </summary>
    ''' <param name="number">Error number.</param>
    ''' <param name="args">Any string format arguments to include in the message.</param>
    ''' <remarks></remarks>
    Public Sub FatalError(number As Integer, ParamArray args() As String)

        mNumErrorsIssued += 1
        EmitMessage(String.Format("Fatal error BC{0} : {1}", number.ToString(), GetMessageText(number, args)))

    End Sub

    ''' <summary>
    ''' Outputs a command line error.
    ''' </summary>
    ''' <param name="number">Number of the error.</param>
    ''' <param name="args">Any string format arguments to include in the message.</param>
    ''' <remarks></remarks>
    Public Sub CommandLineError(number As Integer, ParamArray args() As String)

        mNumErrorsIssued += 1
        EmitMessage(String.Format("Command line error BC{0} : {1}", number.ToString(), GetMessageText(number, args)))

    End Sub

    ''' <summary>
    ''' Outputs a command line warning.
    ''' </summary>
    ''' <param name="number">Number of the error.</param>
    ''' <param name="args">Any string format arguments to include in the message.</param>
    ''' <remarks></remarks>
    Public Sub CommandLineWarning(number As Integer, ParamArray args() As String)

        mNumWarningsIssued += 1
        EmitMessage(String.Format("Command line warning BC{0} : {1}", number.ToString(), GetMessageText(number, args)))

    End Sub

#End Region

#Region "Private Methods"

    ''' <summary>
    ''' Emits or queues the message for output.
    ''' </summary>
    ''' <param name="message">The message to output or to queue.</param>
    ''' <remarks></remarks>
    Private Sub EmitMessage(message As String)

        If mQueueMessageOutput = 0 Then
            Console.WriteLine("mvbc : " & message)
        Else
            mQueuedMessages.Enqueue(message)
        End If

    End Sub

    ''' <summary>
    ''' Gets the text for a particular message from the resource file.
    ''' </summary>
    ''' <param name="messageNumber">Number of the message to get.</param>
    ''' <param name="args">Any string format arguments to embed in the message.</param>
    ''' <returns>The formatted message text.</returns>
    ''' <remarks></remarks>
    Private Function GetMessageText(messageNumber As Integer, ParamArray args() As String) As String

        Dim text = mResourceManager.GetString(messageNumber.ToString())
        Return String.Format(text, args)

    End Function

#End Region

#Region "Properties"

    ''' <summary>
    ''' Accessor for the number of errors that were reported.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property NumberOfErrors As Integer
        Get
            Return mNumErrorsIssued
        End Get
    End Property

    ''' <summary>
    ''' Accessor for the number of warnings that were issued.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property NumberOfWarnings As Integer
        Get
            Return mNumWarningsIssued
        End Get
    End Property

    ''' <summary>
    ''' Set of warnings which should be output as errors
    ''' </summary>
    ''' <remarks></remarks>
    Public WarningsAsErrors As New HashSet(Of Integer)

    ''' <summary>
    ''' Those warnings which are disabled.
    ''' </summary>
    ''' <remarks></remarks>
    Public DisabledWarnings As New HashSet(Of Integer)

#End Region

End Class
