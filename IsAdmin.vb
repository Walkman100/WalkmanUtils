Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Reflection

<assembly: AssemblyTitle("IsAdmin")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("IsAdmin")>

Module Program
    Sub Main(args() As String)
        If args.Length <> 0 Then
            WriteUsage
        Else
            Console.WriteLine(WalkmanLib.IsAdmin)
        End If
    End Sub
    
    Sub WriteUsage()
        Dim flags As String = Environment.NewLine & "Checks whether the current process is elevated (running with administrator permissions). Outputs True if running with administrator permissions, False if not (WalkmanUtils - https://github.com/Walkman100/WalkmanUtils)"
        Dim programPath As String = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
        Dim programFile As String = programPath.Substring(programPath.LastIndexOf("/") +1)
        If My.Computer.Info.OSPlatform = "Unix" Then
            programFile = "mono " & programFile
        End If
        Console.WriteLine("Usage: " & programFile & flags)
    End Sub
End Module
