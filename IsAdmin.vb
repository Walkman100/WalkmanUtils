Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
' Information about this assembly is defined by the following attributes.
<assembly: AssemblyTitle("IsAdmin")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("IsAdmin")>
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

Module Program
    Sub Main(args() As String)
        If args.Length <> 0 Then
            WriteUsage
        Else
            Console.WriteLine(WalkmanLib.IsAdmin)
        End If
    End Sub
    
    Sub WriteUsage()
        Dim flags As String = vbNewLine & "Checks whether the current process is elevated (running with administrator permissions). Outputs True if running with administrator permissions, False if not (WalkmanUtils - https://github.com/Walkman100/WalkmanUtils)"
        Dim programPath As String = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
        Dim programFile As String = programPath.Substring(programPath.LastIndexOf("/") +1)
        If My.Computer.Info.OSPlatform = "Unix" Then
            programFile = "mono " & programFile
        End If
        Console.WriteLine("Usage: " & programFile & flags)
    End Sub
End Module
