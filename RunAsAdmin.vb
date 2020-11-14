Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Reflection
Imports System.Windows.Forms
Imports Microsoft.VisualBasic

<assembly: AssemblyTitle("RunAsAdmin")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("RunAsAdmin")>

Class RunAsAdmin
    Inherits System.Windows.Forms.Form
    Public Sub New()
        Dim args As String() = System.Environment.GetCommandLineArgs()
        If args.Length < 2 Then
            Application.EnableVisualStyles()
            MsgBox(GetUsage(), MsgBoxStyle.Information, "Command line arguments incorrect")
        ElseIf IO.File.Exists(args(1)) Then
            If args.Length = 2 Then
                WalkmanLib.RunAsAdmin(args(1))
            Else
                Dim arguments As String = ""
                For i As Integer = 2 To args.Length - 1
                    arguments &= args(i) & " "
                Next
                arguments = arguments.Remove(arguments.Length - 1) ' to get rid of the extra space at the end
                
                WalkmanLib.RunAsAdmin(args(1), arguments)
            End If
        Else
            Application.EnableVisualStyles()
            MsgBox("Executable """ & args(1) & """ not found!" & Environment.NewLine & Environment.NewLine & GetUsage(), MsgBoxStyle.Exclamation, "Executable not found")
        End If
        
        Do Until 0 <> 0
            Application.Exit
            End
        Loop
    End Sub
    
    Function GetUsage() As String
        Dim flags As String = " <path>" & Environment.NewLine & "Starts a program with a set of command-line arguments as an administrator"
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
            Me.MainForm = My.Forms.RunAsAdmin
        End Sub
    End Class
End Namespace
