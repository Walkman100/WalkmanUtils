Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Reflection

<assembly: AssemblyTitle("GetFolderIcon")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("GetFolderIcon")>

Module Program
    Sub Main(args() As String)
        If args.Length <> 1 Then
            WriteUsage
        ElseIf IO.Directory.Exists(args(0)) Then
            Dim folderPath As String = args(0)
            If folderPath.EndsWith(IO.Path.VolumeSeparatorChar) Then folderPath &= IO.Path.DirectorySeparatorChar
            
            Console.WriteLine(WalkmanLib.GetFolderIconPath(folderPath))
        Else
            Console.WriteLine("Folder """ & args(0) & """ not found!")
            WriteUsage
        End If
    End Sub
    
    Sub WriteUsage()
        Dim flags As String = " <folder path>" & Environment.NewLine & "Get the path to a folder icon, or ""no icon found"" if none is set (WalkmanUtils - https://github.com/Walkman100/WalkmanUtils)"
        Dim programPath As String = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
        Dim programFile As String = programPath.Substring(programPath.LastIndexOf("/") +1)
        If My.Computer.Info.OSPlatform = "Unix" Then
            programFile = "mono " & programFile
        End If
        Console.WriteLine("Usage: " & programFile & flags)
    End Sub
End Module
