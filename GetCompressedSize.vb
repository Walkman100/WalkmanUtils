Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.Reflection

<Assembly: AssemblyTitle("GetCompressedSize")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyConfiguration("")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("GetCompressedSize")>

Module Program
    Function WriteUsage(Optional input As String = Nothing) As Boolean
        Console.Error.WriteLine("Usage: " & GetProgramFileName() & " [OPTION] <FILE>")
        Console.Error.WriteLine("Gets the compressed size of a file (WalkmanUtils - https://github.com/Walkman100/WalkmanUtils)" & Environment.NewLine)
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
        }}
    }

    Sub Main(args() As String)
        Dim res As WalkmanLib.ResultInfo = WalkmanLib.ProcessArgs(args, flagDict, True)

        If res.gotError Then
            ExitE(res.errorInfo)
        ElseIf res.extraParams.Count < 1 Then
            WriteUsage()

        Else
            For Each file As String In res.extraParams
                If IO.File.Exists(file) Then
                    Try
                        Dim compressedSize As Double = WalkmanLib.GetCompressedSize(file)
                        Console.WriteLine(compressedSize)
                    Catch ex As IO.IOException
                        If Console.IsOutputRedirected Then Console.WriteLine("?")
                        Console.Error.WriteLine("Exception getting compressed size: " & ex.Message)
                    End Try
                Else
                    If Console.IsOutputRedirected Then Console.WriteLine("?")
                    Console.Error.WriteLine("File """ & file & """ not found!")
                End If
            Next
        End If
    End Sub
End Module
