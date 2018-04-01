Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.ApplicationServices
' Information about this assembly is defined by the following attributes.
<assembly: AssemblyTitle("OpenWith")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("OpenWith")>
<assembly: AssemblyCopyright("Copyright 2018")>
<assembly: AssemblyTrademark("")>
<assembly: AssemblyCulture("")>
' This sets the default COM visibility of types in the assembly to invisible.
' If you need to expose a type to COM, use <ComVisible(true)> on that type.
<assembly: ComVisible(False)>
' The assembly version has following format :
' Major.Minor.Build.Revision
' You can specify all values by your own or you can build default build and revision
' numbers with the '*' character (the default):
<assembly: AssemblyVersion("1.0.*")>

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
            MsgBox("File or directory """ & args(1) & """ not found!" & vbNewLine & vbNewLine & GetUsage(), MsgBoxStyle.Exclamation, "Path not found")
        End If
        
        Do Until 0 <> 0
            Application.Exit
            End
        Loop
    End Sub
    
    Function GetUsage() As String
        Dim flags As String = " <path>" & vbNewLine & "Open the Open With dialog box for a path"
        flags &= vbNewLine & "WalkmanUtils - https://github.com/Walkman100/WalkmanUtils"
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
