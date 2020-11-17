Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.Reflection

<Assembly: AssemblyTitle("GetFolderIcon")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyConfiguration("")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("GetFolderIcon")>

Module Program
    Function WriteUsage(Optional input As String = Nothing) As Boolean
        Console.Error.WriteLine("Usage: " & GetProgramFileName() & " [OPTION] <FOLDER...>")
        Console.Error.WriteLine("Get the path to a folder icon, or ""no icon found"" if none is set" &
                                WalkmanUtilsText & Environment.NewLine)
        WalkmanLib.EchoHelp(flagDict, input)
        Environment.Exit(0)
        Return True
    End Function

    Private flagDict As New Dictionary(Of String, WalkmanLib.FlagInfo) From {
        {"help", New WalkmanLib.FlagInfo With {
            .shortFlag = "h"c,
            .description = "Show Help",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = "[flag]",
            .action = AddressOf WriteUsage
        }},
        {"folders", New WalkmanLib.FlagInfo With {
            .shortFlag = "f"c,
            .description = "Print folder name before icon path",
            .hasArgs = False,
            .action = Function() DoAndReturn(Sub() printFolderPaths = True)
        }}
    }

    Private printFolderPaths As Boolean = False

    Private Sub WriteFolderIconPaths(folders As List(Of String))
        For Each folder As String In folders
            If IO.Directory.Exists(folder) Then
                If folder.EndsWith(IO.Path.VolumeSeparatorChar) Then folder &= IO.Path.DirectorySeparatorChar
                folder = IO.Path.GetFullPath(folder)

                If printFolderPaths Then Console.Write(folder & ": ")
                Console.WriteLine(WalkmanLib.GetFolderIconPath(folder))
            Else
                If Console.IsOutputRedirected Then
                    If printFolderPaths Then Console.Write(folder & ": ")
                    Console.WriteLine("Path not found")
                End If

                Console.Error.WriteLine("Folder """ & folder & """ not found!")
            End If
        Next
    End Sub

    Sub Main(args() As String)
        Dim res As WalkmanLib.ResultInfo = WalkmanLib.ProcessArgs(args, flagDict, True)

        If res.gotError Then
            ExitE(res.errorInfo)
        ElseIf res.extraParams.Count < 1 Then
            WriteUsage()
        Else
            WriteFolderIconPaths(res.extraParams)
        End If
    End Sub
End Module
