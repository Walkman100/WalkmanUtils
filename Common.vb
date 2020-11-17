Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System

Public Module Common
    ''' <summary>Runs <paramref name="func"/> and returns <see langword="True"/>.</summary>
    Public Function DoAndReturn(func As Action) As Boolean
        func()
        Return True
    End Function

    ''' <summary>
    ''' Converts <paramref name="input"/> to a boolean and returns it, exits the program with an error if convert failed.
    ''' <br />If <paramref name="input"/> is <see langword="Null"/> or <see cref="String"/>.<see langword="Empty"/>, returns <paramref name="default"/>.
    ''' </summary>
    Public Function GetBoolOrExit(input As String, Optional [default] As Boolean = False) As Boolean
        If String.IsNullOrEmpty(input) Then Return [default]

        Dim rtn As Boolean
        If Boolean.TryParse(input, rtn) Then
            Return rtn
        Else
            ExitE("""{0}"" is not True or False!", input)
            End ' fix warning
        End If
    End Function

    Public Function GetProgramFileName() As String
        Dim programPath As String = Reflection.Assembly.GetExecutingAssembly().CodeBase
        Dim programFile As String = IO.Path.GetFileName(programPath)

        If My.Computer.Info.OSPlatform = "Unix" Then
            programFile = "mono " & programFile
        End If
        Return programFile
    End Function

    ''' <summary>Writes <paramref name="msg"/> to STDOUT optionally formatted with <paramref name="formatItem"/>, and exits the program with <paramref name="errorCode"/> or 1.</summary>
    Public Sub ExitE(msg As String, Optional formatItem As String = Nothing, Optional errorCode As Integer = 1)
        If formatItem IsNot Nothing Then
            msg = String.Format(msg, formatItem)
        End If
        Console.Error.WriteLine(msg)
        Environment.Exit(errorCode)
    End Sub
End Module
