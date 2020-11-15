Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Reflection

<assembly: AssemblyTitle("GetCompressedSize")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("GetCompressedSize")>

Module Program
    Sub Main(args() As String)
        If args.Length <> 1 Then
            WriteUsage
        ElseIf IO.File.Exists(args(0)) Then
            Try
                Dim compressedSize As Double = WalkmanLib.GetCompressedSize(args(0))
                Console.WriteLine(compressedSize)
            Catch ex As IO.IOException
                Console.WriteLine("Exception getting compressed size: " & ex.Message)
            End Try
        Else
            Console.WriteLine("File """ & args(0) & """ not found!")
            WriteUsage
        End If
    End Sub
    
    Sub WriteUsage()
        Dim flags As String = " <path>" & Environment.NewLine & "Get the compressed size of a file (WalkmanUtils - https://github.com/Walkman100/WalkmanUtils)"
        Dim programPath As String = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
        Dim programFile As String = programPath.Substring(programPath.LastIndexOf("/") +1)
        If My.Computer.Info.OSPlatform = "Unix" Then
            programFile = "mono " & programFile
        End If
        Console.WriteLine("Usage: " & programFile & flags)
    End Sub
End Module
