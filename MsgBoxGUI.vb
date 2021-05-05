Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Linq
Imports System.Reflection
Imports System.Windows.Forms

<Assembly: AssemblyTitle("MsgBoxGUI")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyConfiguration("")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("MsgBoxGUI")>

Module MsgBoxCommonCompat
    Function WriteUsage(Optional input As String = Nothing) As Boolean
        Return MsgBoxGUI.ShowUsage(input)
    End Function
End Module

Class MsgBoxGUI
    Inherits Form

    Shared Function ShowUsage(Optional input As String = Nothing) As Boolean
        Dim formToShow As New CustomMsgBoxForm With {
            .Prompt = GetUsage(input),
            .Title = "Program Usage",
            .Buttons = MessageBoxButtons.OK,
            .FormLevel = MessageBoxIcon.Information,
            .ShowInTaskbar = True,
            .Width = 1100
        }
        formToShow.MainText.Font = New Drawing.Font("Consolas", 9)
        formToShow.MainText.MaximumSize = New Drawing.Size(1000, 0)
        formToShow.MainText.Size = New Drawing.Size(1000, 0)
        formToShow.ShowDialog()

        Environment.Exit(0)
        Return True
    End Function

    Public Sub New()
        Dim args As String() = Environment.GetCommandLineArgs().Skip(1).ToArray()
        Dim res As WalkmanLib.ResultInfo = WalkmanLib.ProcessArgs(args, flagDict, True)

        If res.gotError Then
            MessageBox.Show(res.errorInfo, "Error processing arguments", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf res.extraParams.Count < 1 Then
            ShowUsage()
        Else
            DoMsgBox(res.extraParams)
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
            Me.MainForm = My.Forms.MsgBoxGUI
        End Sub
    End Class
End Namespace
