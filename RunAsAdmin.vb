Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.ApplicationServices
' Information about this assembly is defined by the following attributes.
<assembly: AssemblyTitle("RunAsAdmin")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("RunAsAdmin")>
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
                For i = 2 To args.Length - 1
                    arguments &= args(i) & " "
                Next
                arguments = arguments.Remove(arguments.Length - 1) ' to get rid of the extra space at the end
                
                WalkmanLib.RunAsAdmin(args(1), arguments)
            End If
        Else
            Application.EnableVisualStyles()
            MsgBox("Executable """ & args(1) & """ not found!" & vbNewLine & vbNewLine & GetUsage(), MsgBoxStyle.Exclamation, "Executable not found")
        End If
        
        Do Until 0 <> 0
            Application.Exit
            End
        Loop
    End Sub
    
    Function GetUsage() As String
        Dim flags As String = " <path>" & vbNewLine & "Starts a program with a set of command-line arguments as an administrator"
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
            Me.MainForm = My.Forms.RunAsAdmin
        End Sub
    End Class
End Namespace
