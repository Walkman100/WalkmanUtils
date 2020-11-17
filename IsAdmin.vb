Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.Reflection

<Assembly: AssemblyTitle("IsAdmin")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyConfiguration("")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("IsAdmin")>

Module Program
    Function WriteUsage(Optional input As String = Nothing) As Boolean
        Console.Error.WriteLine("Usage: " & GetProgramFileName() & " [OPTION]")
        Console.Error.WriteLine("Checks whether the current process is elevated (running with administrator permissions). Outputs True if running with administrator permissions, False if not" &
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
        }}
    }

    Sub Main(args() As String)
        Dim res As WalkmanLib.ResultInfo = WalkmanLib.ProcessArgs(args, flagDict)

        If res.gotError Then
            ExitE(res.errorInfo)
        Else
            Console.WriteLine(WalkmanLib.IsAdmin())
        End If
    End Sub
End Module
