Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Reflection
Imports System.Windows.Forms
Imports Microsoft.VisualBasic

<assembly: AssemblyTitle("OpenWith")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("OpenWith")>

Class OpenWith
    Inherits System.Windows.Forms.Form
    Public Sub New()
        Dim args As String() = System.Environment.GetCommandLineArgs()
        If args.Length <> 2 Then
            Application.EnableVisualStyles()
            MsgBox(GetUsage(), MsgBoxStyle.Information, "Command line arguments incorrect")
        ElseIf IO.File.Exists(args(1)) Or IO.Directory.Exists(args(1)) Then
            WalkmanLib.OpenWith(args(1))
        Else
            Application.EnableVisualStyles()
            MsgBox("File or directory """ & args(1) & """ not found!" & Environment.NewLine & Environment.NewLine & GetUsage(), MsgBoxStyle.Exclamation, "Path not found")
        End If
        
        Do Until 0 <> 0
            Application.Exit
            End
        Loop
    End Sub
    
    Function GetUsage() As String
        Dim flags As String = " <path>" & Environment.NewLine & "Open the Open With dialog box for a path"
        flags &= Environment.NewLine & "WalkmanUtils - https://github.com/Walkman100/WalkmanUtils"
        Dim programPath As String = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
        Dim programFile As String = programPath.Substring(programPath.LastIndexOf("/") +1)
        If My.Computer.Info.OSPlatform = "Unix" Then
            programFile = "mono " & programFile
        End If
        Return "Usage: " & programFile & flags
    End Function
End Class

Namespace My
    Partial Class MyApplication
        Protected Overrides Sub OnCreateMainForm()
            Me.MainForm = My.Forms.OpenWith
        End Sub
    End Class
End Namespace
