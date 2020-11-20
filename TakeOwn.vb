Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports System.Windows.Forms

<Assembly: AssemblyTitle("TakeOwn")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyConfiguration("")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("TakeOwn")>

Class TakeOwn
    Inherits Form

    Function GetUsage(Optional input As String = Nothing) As String
        Dim usageText As String = "Usage: " & GetProgramFileName() & " <FILE...>" & Environment.NewLine
        usageText &= "Take Ownership of a file or recursively Take Ownership of a folder" & WalkmanUtilsText & Environment.NewLine & Environment.NewLine

        Using sw As New IO.StringWriter
            WalkmanLib.EchoHelp(flagDict, input, sw)
            usageText &= sw.ToString()
        End Using

        Return usageText
    End Function

    Function ShowUsage(Optional input As String = Nothing) As Boolean
        MessageBox.Show(GetUsage(input), "Program Usage", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Environment.Exit(0)
        Return True
    End Function

    Private flagDict As New Dictionary(Of String, WalkmanLib.FlagInfo) From {
        {"help", New WalkmanLib.FlagInfo With {
            .shortFlag = "h"c,
            .description = "Show Help",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = "[flag]",
            .action = AddressOf ShowUsage
        }}
    }

    Private Sub DoTakeOwn(paths As List(Of String))
        For Each path As String In paths
            If WalkmanLib.IsFileOrDirectory(path).HasFlag(PathEnum.Exists) Then
                path = IO.Path.GetFullPath(path)
                WalkmanLib.TakeOwnership(path)
            Else
                MessageBox.Show("Path """ & path & """ not found!", "Path not found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Next
    End Sub

    Public Sub New()
        Dim args As String() = Environment.GetCommandLineArgs().Skip(1).ToArray()
        Dim res As WalkmanLib.ResultInfo = WalkmanLib.ProcessArgs(args, flagDict, True)

        If res.gotError Then
            MessageBox.Show(res.errorInfo, "Error processing arguments", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf res.extraParams.Count < 1 Then
            ShowUsage()
        Else
            DoTakeOwn(res.extraParams)
        End If

        Do Until 0 <> 0
            Application.Exit()
            End
        Loop
    End Sub
End Class

Namespace My
    Partial Class MyApplication
        Protected Overrides Sub OnCreateMainForm()
            Me.MainForm = My.Forms.TakeOwn
        End Sub
    End Class
End Namespace
