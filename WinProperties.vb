Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms

<Assembly: AssemblyTitle("WinProperties")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyConfiguration("")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("WinProperties")>

Class WinProperties
    Inherits Form

    Function GetUsage(Optional input As String = Nothing) As String
        Dim usageText As String = "Usage: " & GetProgramFileName() & " <FILE...>" & Environment.NewLine
        usageText &= "Opens the Windows properties window for a path" & WalkmanUtilsText & Environment.NewLine & Environment.NewLine

        Using sw As New IO.StringWriter
            WalkmanLib.EchoHelp(flagDict, input, sw)
            usageText &= sw.ToString()
        End Using

        Return usageText
    End Function

    Function ShowUsage(Optional input As String = Nothing) As Boolean
        MessageBox.Show(GetUsage(input), "Program Usage", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
            .action = AddressOf ShowUsage
        }},
        {"tab", New WalkmanLib.FlagInfo With {
            .shortFlag = "t"c,
            .description = "Open the properties window on a specified tab. Tab name is Language-specific.",
            .hasArgs = True,
            .argsInfo = "<tab name>",
            .action = Function(value As String) DoAndReturn(Sub() tabToShow = value)
        }}
    }

    Private tabToShow As String = Nothing

    Private Sub ShowProperties(paths As List(Of String))
        Dim waitThreads As New List(Of Task)

        For Each path As String In paths
            If WalkmanLib.IsFileOrDirectory(path).HasFlag(PathEnum.Exists) Then
                If WalkmanLib.ShowProperties(path, tabToShow) Then
                    Dim itemName As String = Nothing

                    ' in case . or .. is used        GetFileName returns empty string if path ends in directorySeparatorChar
                    path = IO.Path.GetFullPath(path).TrimEnd(IO.Path.DirectorySeparatorChar)

                    If WalkmanLib.IsFileOrDirectory(path & IO.Path.DirectorySeparatorChar).HasFlag(PathEnum.IsDrive) Then
                        ' handle drives separately.      ^ IsFileOrDirectory needs DirectorySeparatorChar to create fileinfo in the correct dir
                        Dim driveLabel As String = Nothing
                        Try
                            driveLabel = New IO.DriveInfo(path).VolumeLabel
                        Catch
                            ' HACK: attempt to get label from type if getting label failed
                            driveLabel = New IO.DriveInfo(path).DriveType.ToString()
                        End Try
                        itemName = "(" & path & ") " & driveLabel
                    Else
                        itemName = IO.Path.GetFileName(path)
                    End If

                    waitThreads.Add(Task.Run(Sub()
                                                 Thread.Sleep(600)
                                                 WalkmanLib.WaitForWindow(itemName & " Properties", "#32770")
                                             End Sub))
                Else
                    MessageBox.Show("There was an error opening the properties window for """ & path & """!",
                                    "WinProperties", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                End If
            Else
                MessageBox.Show("File or directory """ & path & """ not found!" & Environment.NewLine & Environment.NewLine & GetUsage(),
                                "Path not found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Next

        Try
            Task.WaitAll(waitThreads.ToArray())
            Thread.Sleep(30000) ' wait for shell thread to exit completely
        Catch ex As AggregateException
            MessageBox.Show(ex.InnerException.ToString(), "Error waiting for one or more properties windows!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error waiting for one or more properties windows!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub New()
        Dim args As String() = Environment.GetCommandLineArgs().Skip(1).ToArray()
        Dim res As WalkmanLib.ResultInfo = WalkmanLib.ProcessArgs(args, flagDict, True)

        If res.gotError Then
            MessageBox.Show(res.errorInfo, "Error processing arguments", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf res.extraParams.Count < 1 Then
            ShowUsage()
        Else
            ShowProperties(res.extraParams)
        End If

        Do Until 0 <> 0
            Application.Exit()
            End
        Loop
    End Sub
End Class

Namespace My
    Partial Class MyApplication
        Protected Overrides Sub OnCreateMainForm()
            Me.MainForm = My.Forms.WinProperties
        End Sub
    End Class
End Namespace
