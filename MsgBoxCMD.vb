Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Reflection

<Assembly: AssemblyTitle("MsgBoxCMD")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyConfiguration("")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("MsgBoxCMD")>

Module Program
    Function WriteUsage(Optional input As String = Nothing) As Boolean
        Console.Error.WriteLine(GetUsage(input))
        Environment.Exit(0)
        Return True
    End Function

    Sub Main(args() As String)
        Dim res As WalkmanLib.ResultInfo = WalkmanLib.ProcessArgs(args, flagDict, True)

        If res.gotError Then
            ExitE(res.errorInfo)
        ElseIf res.extraParams.Count < 1 Then
            WriteUsage()
        Else
            Windows.Forms.Application.EnableVisualStyles()
            DoMsgBox(res.extraParams)
        End If
    End Sub
End Module
