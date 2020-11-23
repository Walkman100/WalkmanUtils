Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Linq
Imports System.Reflection
Imports System.Windows.Forms

<Assembly: AssemblyTitle("UncompressFile")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyConfiguration("")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("UncompressFile")>

Class UncompressFile
    Inherits Form

    Function GetUsage(Optional input As String = Nothing) As String
        Dim usageText As String = "Usage: " & GetProgramFileName() & " <FILE...>" & Environment.NewLine
        usageText &= "Decompresses the provided files/folders. If given a folder, only marks it as decompressed" & WalkmanUtilsText & Environment.NewLine & Environment.NewLine

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
        {"hidden", New WalkmanLib.FlagInfo With {
            .shortFlag = "n"c,
            .description = "Start window hidden",
            .action = Function() DoAndReturn(Sub() hideWindow = True)
        }},
        {"quiet", New WalkmanLib.FlagInfo With {
            .shortFlag = "q"c,
            .description = "Don't show error messages",
            .action = Function() DoAndReturn(Sub() showErrors = False)
        }}
    }

    Private hideWindow As Boolean = False
    Private showErrors As Boolean = True

    Private Sub DoUncompress(sender As Object, e As DoWorkEventArgs) Handles bwUncompress.DoWork
        Dim paths As List(Of String) = DirectCast(e.Argument, List(Of String))

        SetStatus("Starting...")

        Dim allSucceeded As Boolean = True

        For Each path As String In paths
            If WalkmanLib.IsFileOrDirectory(path).HasFlag(PathEnum.Exists) Then
                path = IO.Path.GetFullPath(path)

                SetStatus("Decompressing " & path & "...")
                Dim rtn As Boolean = WalkmanLib.UncompressFile(path)
                SetStatus("Decompressing " & path & " succeeded: " & rtn)

                If rtn = False Then allSucceeded = False
            Else
                If showErrors Then MessageBox.Show("Path """ & path & """ not found!", "Path not found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                allSucceeded = False
            End If
        Next

        If Not allSucceeded AndAlso showErrors Then MessageBox.Show("Some items failed to decompress!", "Decompressing items", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub
    Private Sub SetStatus(text As String)
        If lblStatus IsNot Nothing Then
            lblStatus.Text = text
        End If
    End Sub
    Private Sub bwUncompress_WorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles bwUncompress.RunWorkerCompleted
        If e.Error IsNot Nothing Then
            If showErrors Then WalkmanLib.ErrorDialog(e.Error)
            SetStatus("Failed. Close window to exit")
        Else
            Me.Close()
            Application.Exit()
        End If
    End Sub

    Public Sub New()
        Dim args As String() = Environment.GetCommandLineArgs().Skip(1).ToArray()
        Dim res As WalkmanLib.ResultInfo = WalkmanLib.ProcessArgs(args, flagDict, True)

        If res.gotError Then
            MessageBox.Show(res.errorInfo, "Error processing arguments", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0) ' ShowUsage() below exits as well
        ElseIf res.extraParams.Count < 1 Then
            ShowUsage()
        ElseIf hideWindow Then
            DoUncompress(Nothing, New DoWorkEventArgs(res.extraParams))
            Environment.Exit(0)
        Else
            Me.InitializeComponent()

            Using msI As New IO.MemoryStream(Convert.FromBase64String(uncompressIconBase64))
                Me.Icon = New Drawing.Icon(msI)
            End Using

            ' MemoryStream used to create image must NOT be closed - https://stackoverflow.com/q/336387/2999220
            Dim msL As New IO.MemoryStream(Convert.FromBase64String(loadingImageBase64))
            pbxLoading.Image = Drawing.Image.FromStream(msL)

            bwUncompress.RunWorkerAsync(res.extraParams)
        End If
    End Sub

    Private Sub btnHide_Click() Handles btnHide.Click
        Me.Hide()
    End Sub

    Private Const loadingImageBase64 As String =
        "R0lGODlhEAAQALMPAHp6evf394qKiry8vJOTk83NzYKCgubm5t7e3qysrMXFxe7u7pubm7S0tKOjo///" &
        "/yH/C05FVFNDQVBFMi4wAwEAAAAh+QQJCAAPACwAAAAAEAAQAAAETPDJSau9NRDAgWxDYGmdZADCkQnl" &
        "U7CCOA3oNgXsQG2FRhUAAoWDIU6MGeSDR0m4ghRa7JjIUXCogqQzpRxYhi2HILsOGuJxGcNuTyIAIfkE" &
        "CQgADwAsAAAAABAAEAAABGLwSXmMmjhLAQjSWDAYQHmAz8GVQPIESxZwggIYS0AIATYAvAdh8OIQJwRA" &
        "QbJkdjAlUCA6KfU0VEmyGWgWnpNfcEAoAo6SmWtBUtCuk9gjwQKeQAeWYQAHIZICKBoKBncTEQAh+QQJ" &
        "CAAPACwAAAAAEAAQAAAEWvDJORejGCtQsgwDAQAGGWSHMK7jgAWq0CGj0VEDIJxPnvAU0a13eAQKrsnI" &
        "81gqAZ6AUzIonA7JRwFAyAQSgCQsjCmUAIhjDEhlrQTFV+lMGLApWwUzw1jsIwAh+QQJCAAPACwAAAAA" &
        "EAAQAAAETvDJSau9L4QaBgEAMWgEQh0CqALCZ0pBKhRSkYLvM7Ab/OGThoE2+QExyAdiuexhVglKwdCg" &
        "qKKTGGBgBc00Np7VcVsJDpVo5ydyJt/wCAAh+QQJCAAPACwAAAAAEAAQAAAEWvDJSau9OAwCABnBtQhd" &
        "CQjHlQhFWJBCOKWPLAXk8KQIkCwWBcAgMDw4Q5CkgOwohCVCYTIwdAgPolVhWSQAiN1jcLLVQrQbrBV4" &
        "EcySA8l0Alo0yA8cw+9TIgAh+QQFCAAPACwAAAAAEAAQAAAEWvDJSau9WA4AyAhWMChPwXHCQRUGYARg" &
        "KQBCzJxAQgXzIC2KFkc1MREoHMTAhwQ0Y5oBgkMhAAqUw8mgWGho0EcCx5DwaAUQrGXATg6zE7bwCQ2s" &
        "AGZmz7dEAAA7"

    Private Const uncompressIconBase64 As String =
        "AAABAAEAEBAAAAEAIABoBAAAFgAAACgAAAAQAAAAIAAAAAEAIAAAAAAAAAQAAMIOAADCDgAAAAAAAAAA" &
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP///wD///8A////AP///wD///8A////AP//" &
        "/wD///8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKZoHgCmaB4Qpmkgz7JzKP/HhDL/x4U0/8aF" &
        "Nv+3fDX/////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAApmkfAKZpHxCmaSHv3JI3/+ic" &
        "Pf/nnD3/x4c4/////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKZnGwCmZxufv3so/+ib" &
        "Ov/omzv/55s8/8eGNf////8AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKZkFQCmZBWfv3kj/+iZ" &
        "Nv/omjj/55o5/9uRNv/HhTP/////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACmYhBwt3Aa/+iY" &
        "M//omTX/55g1/756KP6maiTvsnUr/////wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAApmEPEKZi" &
        "EM/PhCb/55cy/754I/6maB6epmojEKZrJc////8AAAAAAAAAAAAAAAAAAAAAAKZwMACmcDBvpnAyD6Zi" &
        "EACmYhAQpmMRz7ZwG/6mZRieAAAAAKZrJACmayQQ////AP///wCmaB4PAAAAAKZtKwCmbSuftnoy/qZw" &
        "Mc4AAAAApmIRAKZiERCmYxNvAAAAAAAAAAAAAAAAAAAAAAAAAAD///8ApmgdzqZoHxCmayWfv34w/+ed" &
        "P//OjDr+pnEyzqZxMxAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////ALFvIP6maB/vv3wr/+ic" &
        "Pf/onT7/550//7Z6NP6mcTJwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP///wDHfyf/3JEz/+ib" &
        "Of/omzv/55s8/75+Mf6mby2eAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD///8Ax38l/+iZ" &
        "Nv/omjj/55o4/758LP6mbCieAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////AMd+" &
        "I//omDT/55g1/9uQMv6maiLupmojDwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP//" &
        "/wC3cBn/x34k/8d/Jv/Gfyj/sXEj/qZqIs6mayQPAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" &
        "AAD///8A////AP///wD///8A////AP///wD///8A////AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" &
        "AAAAAAAA//8AAP8BAAD/gQAA/4EAAP8BAAD+AQAA/gEAAPkNAACxnwAAgH8AAIB/AACA/wAAgf8AAIH/" &
        "AACA/wAA//8AAA=="

#Region "Form Design"
    ''' <summary>
    ''' Designer variable used to keep track of non-visual components.
    ''' </summary>
    Private components As System.ComponentModel.IContainer

    ''' <summary>
    ''' Disposes resources used by the form.
    ''' </summary>
    ''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If components IsNot Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ''' <summary>
    ''' This method is required for Windows Forms designer support.
    ''' Do not change the method contents inside the source code editor. The Forms designer might
    ''' not be able to load this method if it was changed manually.
    ''' </summary>
    Private Sub InitializeComponent()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.btnHide = New System.Windows.Forms.Button()
        Me.pbxLoading = New System.Windows.Forms.PictureBox()
        Me.bwUncompress = New System.ComponentModel.BackgroundWorker()
        CType(Me.pbxLoading, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(34, 14)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(25, 13)
        Me.lblStatus.TabIndex = 0
        Me.lblStatus.Text = "???"
        '
        'btnHide
        '
        Me.btnHide.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.btnHide.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnHide.Location = New System.Drawing.Point(638, 9)
        Me.btnHide.Name = "btnHide"
        Me.btnHide.Size = New System.Drawing.Size(75, 23)
        Me.btnHide.TabIndex = 1
        Me.btnHide.Text = "Hide"
        Me.btnHide.UseVisualStyleBackColor = True
        '
        'pbxLoading
        '
        Me.pbxLoading.Location = New System.Drawing.Point(12, 12)
        Me.pbxLoading.Name = "pbxLoading"
        Me.pbxLoading.Size = New System.Drawing.Size(16, 16)
        Me.pbxLoading.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.pbxLoading.TabIndex = 2
        Me.pbxLoading.TabStop = False
        '
        'bwUncompress
        '
        Me.bwUncompress.WorkerReportsProgress = True
        Me.bwUncompress.WorkerSupportsCancellation = True
        '
        'UncompressFile
        '
        Me.AcceptButton = Me.btnHide
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnHide
        Me.ClientSize = New System.Drawing.Size(725, 40)
        Me.Controls.Add(Me.btnHide)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.pbxLoading)
        Me.Name = "UncompressFile"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Decompressing..."
        CType(Me.pbxLoading, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents lblStatus As Label
    Friend WithEvents btnHide As Button
    Friend WithEvents pbxLoading As PictureBox
    Friend WithEvents bwUncompress As BackgroundWorker
#End Region
End Class

Namespace My
    Partial Class MyApplication
        Protected Overrides Sub OnCreateMainForm()
            Me.MainForm = My.Forms.UncompressFile
        End Sub
    End Class
End Namespace
