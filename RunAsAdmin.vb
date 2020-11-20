Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Windows.Forms

<Assembly: AssemblyTitle("RunAsAdmin")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyConfiguration("")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("RunAsAdmin")>

Class RunAsAdmin
    Inherits Form

    Function GetUsage(Optional input As String = Nothing) As String
        Dim usageText As String = "Usage: " & GetProgramFileName() & " <FILE> [ARGS...]" & Environment.NewLine
        usageText &= "Starts a program with a set of command-line arguments as an administrator" & WalkmanUtilsText & Environment.NewLine & Environment.NewLine

        Using sw As New StringWriter
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

    Private Sub DoRunAsAdmin(args As List(Of String))
        Dim f As String = args(0)
        If File.Exists(f) Then
            f = Path.GetFullPath(f)

            If args.Count = 1 Then
                WalkmanLib.RunAsAdmin(f)
            Else
                Dim arguments As String = ""
                For i As Integer = 1 To args.Count - 1
                    arguments &= args(i) & " "
                Next
                arguments = arguments.Remove(arguments.Length - 1) ' to get rid of the extra space at the end

                WalkmanLib.RunAsAdmin(f, arguments)
            End If
        Else
            MessageBox.Show("Executable """ & f & """ not found!", "Executable not found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Public Sub New()
        Dim args As String() = Environment.GetCommandLineArgs().Skip(1).ToArray()
        Dim res As WalkmanLib.ResultInfo = WalkmanLib.ProcessArgs(args, flagDict, True)

        If res.gotError Then
            MessageBox.Show(res.errorInfo, "Error processing arguments", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf res.extraParams.Count < 1 Then
            ShowUsage()
        Else
            DoRunAsAdmin(res.extraParams)
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
            Me.MainForm = My.Forms.RunAsAdmin
        End Sub
    End Class
End Namespace
