Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms

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
            .argsInfo = "<text>",
            .action = Function(value As String) DoAndReturn(Sub() caption = value)
        }},
        {"buttons", New WalkmanLib.FlagInfo With {
            .shortFlag = "b"c,
            .description = "Set the Message Box buttons",
            .hasArgs = True,
            .argsInfo = "<int|MessageBoxButtons value>",
            .action = Function(value As String) TryParseEnumOrInt(value, buttons)
        }},
        {"icon", New WalkmanLib.FlagInfo With {
            .shortFlag = "i"c,
            .description = "Set the Message Box icon",
            .hasArgs = True,
            .argsInfo = "<int|MessageBoxIcon value>",
            .action = Function(value As String) TryParseEnumOrInt(value, icon)
        }},
        {"defaultButton", New WalkmanLib.FlagInfo With {
            .shortFlag = "d"c,
            .description = "Set the Message Box default button",
            .hasArgs = True,
            .argsInfo = "<int|MessageBoxDefaultButton value>",
            .action = Function(value As String) TryParseEnumOrInt(value, defaultButton)
        }},
        {"options", New WalkmanLib.FlagInfo With {
            .shortFlag = "o"c,
            .description = "Set the Message Box extra options",
            .hasArgs = True,
            .argsInfo = "<int|MessageBoxOptions value>",
            .action = Function(value As String) TryParseEnumOrInt(value, options)
        }},
        {"helpFile", New WalkmanLib.FlagInfo With {
            .shortFlag = "f"c,
            .description = "Set the Message Box Help File path",
            .hasArgs = True,
            .argsInfo = "<path>",
            .action = Function(value As String) DoAndReturn(Sub() helpFile = value)
        }},
        {"number", New WalkmanLib.FlagInfo With {
            .shortFlag = "n"c,
            .description = "Change selected button output (MsgBoxCMD only) to the associated number instead of string",
            .action = Function() DoAndReturn(Sub() outputNumber = True)
        }}
    }

    Private caption As String = Nothing
    Private buttons As MessageBoxButtons
    Private icon As MessageBoxIcon
    Private defaultButton As MessageBoxDefaultButton
    Private options As MessageBoxOptions
    Private helpFile As String = Nothing
    Private outputNumber As Boolean = False

    Private Function TryParseEnumOrInt(Of TEnum As Structure)(value As String, ByRef result As TEnum) As Boolean
        result = Nothing

        Dim number As Integer
        If Integer.TryParse(value, number) Then
            If Not [Enum].IsDefined(GetType(TEnum), number) Then
                Return False
            End If
        End If

        Return [Enum].TryParse(value, True, result)
    End Function

    Sub DoMsgBox(args As List(Of String))
        Dim text As String = String.Join(" ", args.ToArray())
        Dim result As DialogResult

        If helpFile Is Nothing Then
            result = MessageBox.Show(text, caption, buttons, icon, defaultButton, options)
        Else
            result = MessageBox.Show(text, caption, buttons, icon, defaultButton, options, helpFile)
        End If

        If outputNumber Then
            Console.WriteLine(DirectCast(result, Integer))
        Else
            Console.WriteLine(result.ToString())
        End If
    End Sub
End Module
