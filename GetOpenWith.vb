Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.Reflection

<Assembly: AssemblyTitle("GetOpenWith")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyConfiguration("")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("GetOpenWith")>

Module Program
    Function WriteUsage(Optional input As String = Nothing) As Boolean
        Console.Error.WriteLine("Usage: " & GetProgramFileName() & " [OPTION] <FILE...>")
        Console.Error.WriteLine("Get the path to the program specified to open a file, or ""Filetype not associated!"" if none is set" &
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
        {"files", New WalkmanLib.FlagInfo With {
            .shortFlag = "f"c,
            .description = "Print file path before associated program path",
            .hasArgs = False,
            .action = Function() DoAndReturn(Sub() printFilePaths = True)
        }}
    }

    Private printFilePaths As Boolean = False

    Private Sub WriteOpenWiths(files As List(Of String))
        For Each file As String In files
            If IO.File.Exists(file) Then
                file = IO.Path.GetFullPath(file)
                If printFilePaths Then Console.Write(file & ": ")

                Try
                    Dim openWith As String = WalkmanLib.GetOpenWith(file)
                    Console.WriteLine(openWith)
                Catch ex As Exception
                    If Console.IsOutputRedirected Then Console.WriteLine("Error")
                    Console.Error.WriteLine("Exception getting associated program: " & ex.Message)
                End Try
            Else
                If Console.IsOutputRedirected Then
                    If printFilePaths Then Console.Write(file & ": ")
                    Console.WriteLine("File not found")
                End If

                Console.Error.WriteLine("File """ & file & """ not found!")
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
            WriteOpenWiths(res.extraParams)
        End If
    End Sub
End Module
