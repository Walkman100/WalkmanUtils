Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Reflection
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.ApplicationServices

<assembly: AssemblyTitle("WinProperties")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("WinProperties")>

Class WinProperties
    Inherits System.Windows.Forms.Form
    Public Sub New()
        Dim args As String() = System.Environment.GetCommandLineArgs()
        If args.Length <> 2 Then
            Application.EnableVisualStyles()
            MsgBox(GetUsage(), MsgBoxStyle.Information, "Command line arguments incorrect")
        ElseIf IO.File.Exists(args(1)) Or IO.Directory.Exists(args(1)) Then
            Application.EnableVisualStyles()
            If WalkmanLib.ShowProperties(args(1)) Then
                System.Threading.Thread.Sleep(2000)
                'FIXME: properties window only shows for a second
            Else
                MsgBox("There was an error opening the properties window for """ & args(1) & """!")
            End If
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
        Dim flags As String = " <path>" & Environment.NewLine & "Opens the Windows properties window for a path"
        flags &= Environment.NewLine & "WalkmanUtils - https://github.com/Walkman100/WalkmanUtils"
        Dim programPath As String = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
        Dim programFile As String = programPath.Substring(programPath.LastIndexOf("/") +1)
        If My.Computer.Info.OSPlatform = "Unix" Then
            programFile = "mono " & programFile
        End If
        Return "Usage: " & programFile & flags
        
        ''' <summary>Opens the Windows properties window for a path.</summary>
        ''' <param name="path">The path to show the window for.</param>
        ''' <param name="tab">Optional tab to open to. Beware, this name is culture-specific!</param>
        ''' <returns>Whether the properties window was shown successfully or not.</returns>
    End Function
End Class

Namespace My
    Partial Class MyApplication
        Protected Overrides Sub OnCreateMainForm()
            Me.MainForm = My.Forms.WinProperties
        End Sub
    End Class
End Namespace
