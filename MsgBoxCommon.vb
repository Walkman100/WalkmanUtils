Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Collections.Generic

Public Module MsgBoxCommon
    Function GetUsage(Optional input As String = Nothing) As String
        Dim usageText As String = "Usage: " & GetProgramFileName() & " [OPTIONS] <TEXT>" & Environment.NewLine
        usageText &= "Show a MessageBox" & WalkmanUtilsText & Environment.NewLine & Environment.NewLine

        Using sw As New IO.StringWriter
            WalkmanLib.EchoHelp(flagDict, input, sw)
            usageText &= sw.ToString()
        End Using

        Return usageText
    End Function

    Public ReadOnly flagDict As New Dictionary(Of String, WalkmanLib.FlagInfo) From {
        {"help", New WalkmanLib.FlagInfo With {
            .shortFlag = "h"c,
            .description = "Show Help",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = "[flag]",
            .action = AddressOf WriteUsage
        }},
        {"caption", New WalkmanLib.FlagInfo With {
            .shortFlag = "c"c,
            .description = "Set the Message Box window title",
            .hasArgs = True,
            .argsInfo = "<text>"
        }},
        {"buttons", New WalkmanLib.FlagInfo With {
            .shortFlag = "b"c,
            .description = "Set the Message Box buttons",
            .hasArgs = True,
            .argsInfo = "<int|MessageBoxButtons value>"
        }},
        {"icon", New WalkmanLib.FlagInfo With {
            .shortFlag = "i"c,
            .description = "Set the Message Box icon",
            .hasArgs = True,
            .argsInfo = "<int|MessageBoxIcon value>"
        }},
        {"defaultButton", New WalkmanLib.FlagInfo With {
            .shortFlag = "d"c,
            .description = "Set the Message Box default button",
            .hasArgs = True,
            .argsInfo = "<int|MessageBoxDefaultButton value>"
        }},
        {"options", New WalkmanLib.FlagInfo With {
            .shortFlag = "o"c,
            .description = "Set the Message Box extra options",
            .hasArgs = True,
            .argsInfo = "<int|MessageBoxOptions value>"
        }},
        {"helpFile", New WalkmanLib.FlagInfo With {
            .shortFlag = "f"c,
            .description = "Set the Message Box Help File path",
            .hasArgs = True,
            .argsInfo = "<path>"
        }},
        {"number", New WalkmanLib.FlagInfo With {
            .shortFlag = "n"c,
            .description = "Change selected button output (MsgBoxCMD only) to the associated number instead of string"
        }}
    }

    Sub DoMsgBox(args As List(Of String))
        
    End Sub
End Module
