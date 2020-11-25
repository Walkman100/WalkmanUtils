Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Collections.Generic

Public Module MsgBoxCommon
    Function GetUsage(Optional input As String = Nothing) As String
        Dim usageText As String = "Usage: " & GetProgramFileName() & " " & Environment.NewLine
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
        }}
    }

    Sub DoMsgBox(args As List(Of String))
        
    End Sub
End Module
