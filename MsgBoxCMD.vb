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

    Sub WriteError(format As String, arg0 As String)
        Console.Error.WriteLine(format, arg0)
    End Sub

    Sub Main(args() As String)
        Dim res As WalkmanLib.ResultInfo = WalkmanLib.ProcessArgs(args, flagDict, True)

        ' if res.gotError is true and res.errorInfo is Nothing, then there was an error but it was shown in ProcessArgs
        If res.gotError AndAlso res.errorInfo IsNot Nothing Then
            ExitE(res.errorInfo)
        ElseIf res.extraParams.Count < 1 Then
            WriteUsage()
        Else
            Windows.Forms.Application.EnableVisualStyles()
            DoMsgBox(res.extraParams)
        End If
    End Sub
End Module
