Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Reflection

<Assembly: AssemblyTitle("SetAttribute")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyConfiguration("")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("SetAttribute")>

Module Program
    Function WriteUsage(Optional input As String = Nothing) As Boolean
        Console.Error.WriteLine("Usage: " & GetProgramFileName() & " [OPTION] <FILE...>")
        Console.Error.WriteLine("Sets the provided attributes to all items specified" & WalkmanUtilsText & Environment.NewLine)
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
        {"quiet", New WalkmanLib.FlagInfo With {
            .shortFlag = "q"c,
            .description = "Don't output whether setting attributes succeeded or failed",
            .action = Function() DoAndReturn(Sub() QuietOutput = True)
        }},
        {"files", New WalkmanLib.FlagInfo With {
            .shortFlag = "f"c,
            .description = "Print file path before success value",
            .action = Function() DoAndReturn(Sub() PrintFilePaths = True)
        }},
        {"ReadOnly", New WalkmanLib.FlagInfo With {
            .shortFlag = "r"c,
            .description = "Set or Unset the Read Only flag",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = AllowedValues,
            .action = Function(input As String) DoAndReturn(Sub() [ReadOnly] = GetSetMode(input))
        }},
        {"Hidden", New WalkmanLib.FlagInfo With {
            .shortFlag = "H"c,
            .description = "Set or Unset the Hidden flag",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = AllowedValues,
            .action = Function(input As String) DoAndReturn(Sub() Hidden = GetSetMode(input))
        }},
        {"System", New WalkmanLib.FlagInfo With {
            .shortFlag = "s"c,
            .description = "Set or Unset the System flag",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = AllowedValues,
            .action = Function(input As String) DoAndReturn(Sub() System = GetSetMode(input))
        }},
        {"Archive", New WalkmanLib.FlagInfo With {
            .shortFlag = "a"c,
            .description = "Set or Unset the Archive flag",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = AllowedValues,
            .action = Function(input As String) DoAndReturn(Sub() Archive = GetSetMode(input))
        }},
        {"Normal", New WalkmanLib.FlagInfo With {
            .shortFlag = "n"c,
            .description = "Set or Unset the Normal flag - Setting clears most other flags",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = AllowedValues,
            .action = Function(input As String) DoAndReturn(Sub() Normal = GetSetMode(input))
        }},
        {"Temporary", New WalkmanLib.FlagInfo With {
            .shortFlag = "t"c,
            .description = "Set or Unset the Temporary flag",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = AllowedValues,
            .action = Function(input As String) DoAndReturn(Sub() Temporary = GetSetMode(input))
        }},
        {"Offline", New WalkmanLib.FlagInfo With {
            .shortFlag = "o"c,
            .description = "Set or Unset the Offline flag",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = AllowedValues,
            .action = Function(input As String) DoAndReturn(Sub() Offline = GetSetMode(input))
        }},
        {"NotContentIndexed", New WalkmanLib.FlagInfo With {
            .shortFlag = "c"c,
            .description = "Set or Unset the NotContentIndexed flag",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = AllowedValues,
            .action = Function(input As String) DoAndReturn(Sub() NotContentIndexed = GetSetMode(input))
        }},
        {"Encrypted", New WalkmanLib.FlagInfo With {
            .shortFlag = "e"c,
            .description = "Set or Unset the Encrypted flag",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = AllowedValues,
            .action = Function(input As String) DoAndReturn(Sub() Encrypted = GetSetMode(input))
        }},
        {"IntegrityStream", New WalkmanLib.FlagInfo With {
            .shortFlag = "i"c,
            .description = "Set or Unset the IntegrityStream flag",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = AllowedValues,
            .action = Function(input As String) DoAndReturn(Sub() IntegrityStream = GetSetMode(input))
        }},
        {"NoScrubData", New WalkmanLib.FlagInfo With {
            .shortFlag = "d"c,
            .description = "Set or Unset the NoScrubData flag",
            .hasArgs = True,
            .optionalArgs = True,
            .argsInfo = AllowedValues,
            .action = Function(input As String) DoAndReturn(Sub() NoScrubData = GetSetMode(input))
        }}
    }

    Private QuietOutput As Boolean = False
    Private PrintFilePaths As Boolean = False

    Private Enum SetMode
        None
        [Set]
        Unset
    End Enum
    Private [ReadOnly] As SetMode
    Private Hidden As SetMode
    Private System As SetMode
    Private Archive As SetMode
    Private Normal As SetMode
    Private Temporary As SetMode
    Private Offline As SetMode
    Private NotContentIndexed As SetMode
    Private Encrypted As SetMode
    Private IntegrityStream As SetMode
    Private NoScrubData As SetMode

    Private Function StrEquCaseInsensitive(str As String, ParamArray values As String()) As Boolean
        Return values.Any(Function(value As String) str.Equals(value, StringComparison.OrdinalIgnoreCase))
    End Function

    Private Const AllowedValues As String = "[true|false|yes|no|1|0]"
    Private Function GetSetMode(input As String, Optional [default] As SetMode = SetMode.Set) As SetMode
        If String.IsNullOrEmpty(input) Then Return [default]

        If StrEquCaseInsensitive(input, "true", "t", "yes", "y", "1") Then
            Return SetMode.Set
        ElseIf StrEquCaseInsensitive(input, "false", "f", "no", "n", "0") Then
            Return SetMode.Unset
        Else
            ExitE("""{0}"" is not an accepted value!", input)
            End ' fix warning
        End If
    End Function

    Private Sub SetAttributes(files As List(Of String))
        For Each file As String In files
            If IO.File.Exists(file) Then
                file = Path.GetFullPath(file)
                If PrintFilePaths Then Console.Write(file & ": ")

                Dim rtn As Boolean

                If [ReadOnly] <> SetMode.None Then rtn = WalkmanLib.ChangeAttribute(file, FileAttributes.ReadOnly, [ReadOnly] = SetMode.Set)
                If Hidden <> SetMode.None Then rtn = WalkmanLib.ChangeAttribute(file, FileAttributes.Hidden, Hidden = SetMode.Set)
                If System <> SetMode.None Then rtn = WalkmanLib.ChangeAttribute(file, FileAttributes.System, System = SetMode.Set)
                If Archive <> SetMode.None Then rtn = WalkmanLib.ChangeAttribute(file, FileAttributes.Archive, Archive = SetMode.Set)
                If Normal <> SetMode.None Then rtn = WalkmanLib.ChangeAttribute(file, FileAttributes.Normal, Normal = SetMode.Set)
                If Temporary <> SetMode.None Then rtn = WalkmanLib.ChangeAttribute(file, FileAttributes.Temporary, Temporary = SetMode.Set)
                If Offline <> SetMode.None Then rtn = WalkmanLib.ChangeAttribute(file, FileAttributes.Offline, Offline = SetMode.Set)
                If NotContentIndexed <> SetMode.None Then rtn = WalkmanLib.ChangeAttribute(file, FileAttributes.NotContentIndexed, NotContentIndexed = SetMode.Set)
                If Encrypted <> SetMode.None Then rtn = WalkmanLib.ChangeAttribute(file, FileAttributes.Encrypted, Encrypted = SetMode.Set)
                If IntegrityStream <> SetMode.None Then rtn = WalkmanLib.ChangeAttribute(file, FileAttributes.IntegrityStream, IntegrityStream = SetMode.Set)
                If NoScrubData <> SetMode.None Then rtn = WalkmanLib.ChangeAttribute(file, FileAttributes.NoScrubData, NoScrubData = SetMode.Set)

                If Not QuietOutput Then Console.Write(rtn)

                Console.WriteLine()
            Else
                If Console.IsOutputRedirected Then
                    If PrintFilePaths Then Console.Write(file & ": ")
                    Console.WriteLine("?")
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
            SetAttributes(res.extraParams)
        End If
    End Sub
End Module
