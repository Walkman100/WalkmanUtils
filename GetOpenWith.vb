Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Reflection

<assembly: AssemblyTitle("GetOpenWith")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("GetOpenWith")>

Module Program
    Sub Main(args() As String)
        If args.Length <> 1 Then
            WriteUsage
        ElseIf IO.File.Exists(args(0)) Then
            Console.WriteLine(WalkmanLib.GetOpenWith(args(0)))
        Else
            Console.WriteLine("File """ & args(0) & """ not found!")
            WriteUsage
        End If
    End Sub
    
    Sub WriteUsage()
        Dim flags As String = " <path>" & Environment.NewLine & "Get the path to the program specified to open a file, or ""Filetype not associated!"" if none is set (WalkmanUtils - https://github.com/Walkman100/WalkmanUtils)"
        Dim programPath As String = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
        Dim programFile As String = programPath.Substring(programPath.LastIndexOf("/") +1)
        If My.Computer.Info.OSPlatform = "Unix" Then
            programFile = "mono " & programFile
        End If
        Console.WriteLine("Usage: " & programFile & flags)
    End Sub
End Module
