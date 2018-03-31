Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.ApplicationServices
' Information about this assembly is defined by the following attributes.
<assembly: AssemblyTitle("RunAsAdmin")>
<assembly: AssemblyDescription("")>
<assembly: AssemblyConfiguration("")>
<assembly: AssemblyCompany("")>
<assembly: AssemblyProduct("RunAsAdmin")>
<assembly: AssemblyCopyright("Copyright 2018")>
<assembly: AssemblyTrademark("")>
<assembly: AssemblyCulture("")>
' This sets the default COM visibility of types in the assembly to invisible.
' If you need to expose a type to COM, use <ComVisible(true)> on that type.
<assembly: ComVisible(False)>
' The assembly version has following format :
' Major.Minor.Build.Revision
' You can specify all values by your own or you can build default build and revision
' numbers with the '*' character (the default):
<assembly: AssemblyVersion("1.0.*")>

Class RunAsAdmin
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
        'RunAsAdmin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6!, 13!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(546, 125)
        Me.Name = "RunAsAdmin"
        Me.Text = "RunAsAdmin"
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
    ' This file controls the behaviour of the application.
    Partial Class MyApplication
        Public Sub New()
            MyBase.New(AuthenticationMode.Windows)
            Me.IsSingleInstance = False
            Me.EnableVisualStyles = True
            Me.SaveMySettingsOnExit = True
            Me.ShutDownStyle = ShutdownMode.AfterMainFormCloses
        End Sub
        
        Protected Overrides Sub OnCreateMainForm()
            Me.MainForm = My.Forms.RunAsAdmin
        End Sub
    End Class
End Namespace
