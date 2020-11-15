Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Reflection
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.ApplicationServices

<assembly: AssemblyTitle("TakeOwn")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("TakeOwn")>

Class TakeOwn
    Inherits System.Windows.Forms.Form
    Public Sub New()
        Dim args As String() = System.Environment.GetCommandLineArgs()
        If args.Length <> 2 Then
            Application.EnableVisualStyles()
            MsgBox(GetUsage(), MsgBoxStyle.Information, "Command line arguments incorrect")
        ElseIf IO.File.Exists(args(1)) Or IO.Directory.Exists(args(1)) Then
            WalkmanLib.TakeOwnership(args(1))
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
        Dim flags As String = " <path>" & Environment.NewLine & "Take Ownership of a file or recursively Take Ownership of a folder"
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
            Me.MainForm = My.Forms.TakeOwn
        End Sub
    End Class
End Namespace
