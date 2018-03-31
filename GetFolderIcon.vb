Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
' Information about this assembly is defined by the following attributes.
<assembly: AssemblyTitle("GetFolderIcon")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("GetFolderIcon")>
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
        If args.Length <> 1 Then
            WriteUsage
        ElseIf IO.Directory.Exists(args(0)) Then
            Dim folderPath = args(0)
            If folderPath.EndsWith(IO.Path.VolumeSeparatorChar) Then folderPath &= IO.Path.DirectorySeparatorChar
            
            Console.WriteLine(WalkmanLib.GetFolderIconPath(folderPath))
        Else
            Console.WriteLine("Folder """ & args(0) & """ not found!")
            WriteUsage
        End If
    End Sub
    
    Sub WriteUsage()
        Dim flags As String = " <folder path>" & vbNewLine & "Get the path to a folder icon, or ""no icon found"" if none is set (WalkmanUtils - https://github.com/Walkman100/WalkmanUtils)"
        Dim programPath As String = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
        Dim programFile As String = programPath.Substring(programPath.LastIndexOf("/") +1)
        If My.Computer.Info.OSPlatform = "Unix" Then
            programFile = "mono " & programFile
        End If
        Console.WriteLine("Usage: " & programFile & flags)
    End Sub
End Module
