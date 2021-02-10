Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Reflection
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports WalkmanLib

<Assembly: AssemblyTitle("HandleManager")>
<Assembly: AssemblyDescription("")>
<Assembly: AssemblyConfiguration("")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("HandleManager")>

Class HandleManager
    Inherits Form

    Function GetUsage(Optional input As String = Nothing) As String
        Dim usageText As String = "Usage: " & GetProgramFileName() & " <FILE>" & Environment.NewLine
        usageText &= "Shows open handles on the specified file/folder." & WalkmanUtilsText & Environment.NewLine & Environment.NewLine

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
        {"autoClose", New WalkmanLib.FlagInfo With {
            .shortFlag = "a"c,
            .description = "Automatically close all handles found",
            .action = Function() DoAndReturn(Sub() autoClose = True)
        }},
        {"autoKill", New WalkmanLib.FlagInfo With {
            .shortFlag = "k"c,
            .description = "Automatically KILL ALL PROGRAMS FOUND",
            .action = Function() DoAndReturn(Sub() autoKill = True)
        }}
    }

    Private autoClose As Boolean = False
    Private autoKill As Boolean = False
    Private currentPath As String = Nothing

    Public Sub New()
        Dim args As String() = Environment.GetCommandLineArgs().Skip(1).ToArray()
        Dim res As WalkmanLib.ResultInfo = WalkmanLib.ProcessArgs(args, flagDict, True)

        If res.gotError Then
            MessageBox.Show(res.errorInfo, "Error processing arguments", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Environment.Exit(0) ' ShowUsage() below exits as well
        ElseIf res.extraParams.Count <> 1 Then
            ShowUsage()
        Else
            Me.InitializeComponent()
            currentPath = res.extraParams(0)

            Me.Text = "Processes using: " & currentPath
        End If
    End Sub

    Private Sub HandleManager_Shown() Handles Me.Shown
        If Not bwHandleScan.IsBusy Then
            btnScan_Click()
        End If
    End Sub

    Private Sub lstHandles_ItemSelectionChanged() Handles lstHandles.ItemSelectionChanged
        If lstHandles.SelectedItems.Count = 0 Then
            btnKillProcess.Enabled = False
            btnCloseHandle.Enabled = False
        Else
            btnKillProcess.Enabled = True
            btnCloseHandle.Enabled = False
            For Each item As ListViewItem In lstHandles.SelectedItems
                If Not String.IsNullOrWhiteSpace(item.SubItems(0).Text) AndAlso Not String.IsNullOrWhiteSpace(item.SubItems(2).Text) Then
                    btnCloseHandle.Enabled = True
                End If
            Next
        End If
    End Sub

    Private Sub lstHandles_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles lstHandles.ColumnClick
        If e.Column = 0 Then
            lstHandles.Sorting = If(lstHandles.Sorting = SortOrder.Ascending, SortOrder.Descending, SortOrder.Ascending)
        Else
            'lstHandles.Sort(e.Column)
        End If
    End Sub

    Sub btnScan_Click() Handles btnScan.Click
        btnScan.Enabled = False
        If bwHandleScan.IsBusy Then
            bwHandleScan.CancelAsync()
        ElseIf currentPath IsNot Nothing Then
            lstHandles.Items.Clear()
            bwHandleScan.RunWorkerAsync(currentPath)
        End If
    End Sub

    Sub btnKillProcess_Click() Handles btnKillProcess.Click
        If Not MessageBox.Show("Are you sure you want to kill the selected program(s)?", "Kill Process(es)",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Exit Sub
        End If
        For Each item As ListViewItem In lstHandles.SelectedItems
            Try
                Diagnostics.Process.GetProcessById(Integer.Parse(item.SubItems(0).Text)).Kill()
                item.Selected = False
                item.ForeColor = Drawing.SystemColors.GrayText
            Catch ex As Exception
                WalkmanLib.ErrorDialog(ex)
            End Try
        Next
    End Sub

    Sub btnCloseHandle_Click() Handles btnCloseHandle.Click
        For Each item As ListViewItem In lstHandles.SelectedItems
            If Not String.IsNullOrWhiteSpace(item.SubItems(0).Text) AndAlso Not String.IsNullOrWhiteSpace(item.SubItems(2).Text) Then
                Try
                    Dim handleID As UShort = UShort.Parse(item.SubItems(2).Text.Substring(2), Globalization.NumberStyles.AllowHexSpecifier)

                    SystemHandles.CloseSystemHandle(UInteger.Parse(item.SubItems(0).Text), handleID)

                    item.Selected = False
                    item.ForeColor = Drawing.SystemColors.GrayText
                Catch ex As Exception
                    WalkmanLib.ErrorDialog(ex)
                End Try
            End If
        Next
    End Sub

    Sub btnClose_Click() Handles btnClose.Click
        Me.Close()
    End Sub

    Sub bwHandleScan_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles bwHandleScan.ProgressChanged
        lblStatus.Text = DirectCast(e.UserState, String)
    End Sub

    Sub bwHandleScan_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles bwHandleScan.RunWorkerCompleted
        btnScan.Enabled = True
        btnScan.Text = "Start New Scan"

        lblStatus.Text = "Status: Idle."
        If e.Error IsNot Nothing Then
            WalkmanLib.ErrorDialog(e.Error)
        ElseIf e.Cancelled Then
            MessageBox.Show("The scan waiting was cancelled! Running handle checks are still running and can only be stopped by restarting HandleManager...",
                            "Canceled Scan", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Sub bwHandleScan_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles bwHandleScan.DoWork
        btnScan.Enabled = True
        btnScan.Text = "Cancel Waiting"

        Dim filePath As String = DirectCast(e.Argument, String)
        filePath = New IO.FileInfo(filePath).FullName

        bwHandleScan.ReportProgress(0, "Status: Getting Processes from Restart Manager...")

        Try
            For Each process As Diagnostics.Process In RestartManager.GetLockingProcesses(filePath)
                Dim tmpListViewItem As New ListViewItem({process.Id.ToString(), process.MainModule.FileName, "", ""})
                lstHandles.Items.Add(tmpListViewItem)
            Next

            If autoKill Then
                lstHandles.Items.Cast(Of ListViewItem).ToList().ForEach(Sub(item) item.Selected = True)
                btnKillProcess.PerformClick()
            End If
        Catch ex As Exception When TypeOf ex.InnerException Is ComponentModel.Win32Exception
            ' ignore exceptions on folders - restart manager doesn't allow getting folder locks
        Catch ex As Exception
            ' allow systemHandles code to run even if restartManager fails
            WalkmanLib.ErrorDialog(ex)
        End Try

        lstHandles.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
        lstHandles_ItemSelectionChanged()

        bwHandleScan.ReportProgress(0, "Status: Getting System Handles...")
        Dim taskList As New List(Of Task)

        For Each systemHandle As SystemHandles.SYSTEM_HANDLE In SystemHandles.GetSystemHandles()
            taskList.Add(Task.Run(Sub()
                                      Dim handleInfo As SystemHandles.HandleInfo = SystemHandles.GetHandleInfo(systemHandle, True, SystemHandles.SYSTEM_HANDLE_TYPE.FILE)
                                      If handleInfo.Type = SystemHandles.SYSTEM_HANDLE_TYPE.FILE AndAlso handleInfo.Name IsNot Nothing Then
                                          handleInfo.Name = SystemHandles.ConvertDevicePathToDosPath(handleInfo.Name)
                                          If handleInfo.Name.Contains(filePath) Then
                                              addHandle(handleInfo)
                                          End If
                                      End If
                                  End Sub
            ))

            If bwHandleScan.CancellationPending Then e.Cancel = True : Return
        Next

        ' while any tasks haven't finished
        While taskList.Any(Function(task As Task)
                               Return task.IsCompleted = False
                           End Function)

            ' report unfinished tasks count
            Dim taskCount As Integer = taskList.Where(Function(task As Task)
                                                          Return task.IsCompleted = False
                                                      End Function).Count()
            bwHandleScan.ReportProgress(0, "Status: Waiting for tasks: " & taskCount & " remaining...")

            If bwHandleScan.CancellationPending Then e.Cancel = True : Return
            ' wait for any tasks to complete
            Task.WaitAny(taskList.ToArray(), 200)
            If bwHandleScan.CancellationPending Then e.Cancel = True : Return

            ' remove complete tasks so they aren't waited
            taskList.RemoveAll(Function(task As Task)
                                   Return task.IsCompleted
                               End Function)
        End While
    End Sub

    Private Sub addHandle(handleInfo As SystemHandles.HandleInfo)
        Dim processPath As String
        Try : processPath = Diagnostics.Process.GetProcessById(CType(handleInfo.ProcessID, Integer)).MainModule.FileName
        Catch ex As InvalidOperationException : processPath = ""
        End Try

        Dim tmpListviewItem As New ListViewItem({handleInfo.ProcessID.ToString(), processPath, "0x" & handleInfo.HandleID.ToString("x"), handleInfo.Name})
        If autoKill OrElse autoClose Then tmpListviewItem.Selected = True

        lstHandles.Items.Add(tmpListviewItem)

        lstHandles.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
        lstHandles.Refresh()
        lstHandles_ItemSelectionChanged()

        If autoKill Then
            btnKillProcess.PerformClick()
        ElseIf autoClose Then
            btnCloseHandle.PerformClick()
        End If
    End Sub

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
        Me.btnScan = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnKillProcess = New System.Windows.Forms.Button()
        Me.btnCloseHandle = New System.Windows.Forms.Button()
        Me.lstHandles = New System.Windows.Forms.ListView()
        Me.colHeadProcessID = New System.Windows.Forms.ColumnHeader()
        Me.colHeadProcessPath = New System.Windows.Forms.ColumnHeader()
        Me.colHeadHandleID = New System.Windows.Forms.ColumnHeader()
        Me.colHeadHandleName = New System.Windows.Forms.ColumnHeader()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.bwHandleScan = New System.ComponentModel.BackgroundWorker()
        Me.SuspendLayout()
        '
        'btnScan
        '
        Me.btnScan.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.btnScan.Location = New System.Drawing.Point(822, 75)
        Me.btnScan.Name = "btnScan"
        Me.btnScan.Size = New System.Drawing.Size(100, 23)
        Me.btnScan.TabIndex = 11
        Me.btnScan.Text = "Start Scanning"
        Me.btnScan.UseVisualStyleBackColor = True
        '
        'btnClose
        '
        Me.btnClose.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(822, 162)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 23)
        Me.btnClose.TabIndex = 15
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnKillProcess
        '
        Me.btnKillProcess.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.btnKillProcess.Enabled = False
        Me.btnKillProcess.Location = New System.Drawing.Point(822, 104)
        Me.btnKillProcess.Name = "btnKillProcess"
        Me.btnKillProcess.Size = New System.Drawing.Size(100, 23)
        Me.btnKillProcess.TabIndex = 12
        Me.btnKillProcess.Text = "Kill Process"
        Me.btnKillProcess.UseVisualStyleBackColor = True
        '
        'btnCloseHandle
        '
        Me.btnCloseHandle.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.btnCloseHandle.Enabled = False
        Me.btnCloseHandle.Location = New System.Drawing.Point(822, 133)
        Me.btnCloseHandle.Name = "btnCloseHandle"
        Me.btnCloseHandle.Size = New System.Drawing.Size(100, 23)
        Me.btnCloseHandle.TabIndex = 13
        Me.btnCloseHandle.Text = "Close Handle"
        Me.btnCloseHandle.UseVisualStyleBackColor = True
        '
        'lstHandles
        '
        Me.lstHandles.AllowColumnReorder = True
        Me.lstHandles.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstHandles.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colHeadProcessID, Me.colHeadProcessPath, Me.colHeadHandleID, Me.colHeadHandleName})
        Me.lstHandles.FullRowSelect = True
        Me.lstHandles.GridLines = True
        Me.lstHandles.HideSelection = False
        Me.lstHandles.LabelWrap = False
        Me.lstHandles.Location = New System.Drawing.Point(12, 12)
        Me.lstHandles.Name = "lstHandles"
        Me.lstHandles.Size = New System.Drawing.Size(804, 223)
        Me.lstHandles.TabIndex = 8
        Me.lstHandles.UseCompatibleStateImageBehavior = False
        Me.lstHandles.View = System.Windows.Forms.View.Details
        '
        'colHeadProcessID
        '
        Me.colHeadProcessID.Text = "Process ID"
        Me.colHeadProcessID.Width = 200
        '
        'colHeadProcessPath
        '
        Me.colHeadProcessPath.Text = "Process Path"
        Me.colHeadProcessPath.Width = 200
        '
        'colHeadHandleID
        '
        Me.colHeadHandleID.Text = "Handle ID"
        Me.colHeadHandleID.Width = 200
        '
        'colHeadHandleName
        '
        Me.colHeadHandleName.Text = "Handle Name"
        Me.colHeadHandleName.Width = 200
        '
        'lblStatus
        '
        Me.lblStatus.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(12, 238)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(88, 13)
        Me.lblStatus.TabIndex = 16
        Me.lblStatus.Text = "Status: Starting..."
        '
        'bwHandleScan
        '
        Me.bwHandleScan.WorkerReportsProgress = True
        Me.bwHandleScan.WorkerSupportsCancellation = True
        '
        'HandleManager
        '
        Me.AcceptButton = Me.btnScan
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(934, 260)
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.btnScan)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnKillProcess)
        Me.Controls.Add(Me.btnCloseHandle)
        Me.Controls.Add(Me.lstHandles)
        Me.Name = "HandleManager"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Processes using file: Checking..."
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents bwHandleScan As System.ComponentModel.BackgroundWorker
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents colHeadHandleName As System.Windows.Forms.ColumnHeader
    Friend WithEvents colHeadHandleID As System.Windows.Forms.ColumnHeader
    Friend WithEvents colHeadProcessPath As System.Windows.Forms.ColumnHeader
    Friend WithEvents colHeadProcessID As System.Windows.Forms.ColumnHeader
    Friend WithEvents lstHandles As System.Windows.Forms.ListView
    Friend WithEvents btnCloseHandle As System.Windows.Forms.Button
    Friend WithEvents btnKillProcess As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnScan As System.Windows.Forms.Button
#End Region
End Class

Namespace My
    Partial Class MyApplication
        Protected Overrides Sub OnCreateMainForm()
            Me.MainForm = My.Forms.HandleManager
        End Sub
    End Class
End Namespace