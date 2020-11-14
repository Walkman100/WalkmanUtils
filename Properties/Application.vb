Option Explicit On
Option Strict On
Option Compare Binary
Option Infer Off

Imports Microsoft.VisualBasic.ApplicationServices

Namespace My
    ' This file controls the behaviour of the application.
    Partial Class MyApplication
        Public Sub New()
            MyBase.New(AuthenticationMode.Windows)
            Me.IsSingleInstance = False
            Me.EnableVisualStyles = True
            Me.SaveMySettingsOnExit = True
            Me.ShutdownStyle = ShutdownMode.AfterMainFormCloses
        End Sub
    End Class
End Namespace
