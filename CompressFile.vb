Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports System
Imports System.Reflection

<assembly: AssemblyTitle("CompressFile")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("CompressFile")>

Class CompressFile
    Inherits System.Windows.Forms.Form
    
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
        Me.SuspendLayout
        '
        'CompressFile
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(546, 125)
        Me.Name = "CompressFile"
        Me.Text = "CompressFile"
        Me.ResumeLayout(false)
    End Sub
    
    Public Sub New()
        ' The Me.InitializeComponent call is required for Windows Forms designer support.
        Me.InitializeComponent()
        
        '
        ' TODO : Add constructor code after InitializeComponents
        '
    End Sub
End Class

Namespace My
    Partial Class MyApplication
        Protected Overrides Sub OnCreateMainForm()
            Me.MainForm = My.Forms.CompressFile
        End Sub
    End Class
End Namespace
