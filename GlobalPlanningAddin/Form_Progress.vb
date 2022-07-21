Option Explicit On
Option Strict On

Imports System.Threading


Public Class Form_Progress
    Sub New(StatusText As String)
        InitializeComponent() ' Cet appel est requis par le concepteur.

        Me.Label_Status.Text = StatusText
        Me.ProgressBar1.Value = 0
        Me.ProgressBar1.Maximum = 100
    End Sub

    Public Sub SetProgress(Progress As Integer, Optional StatusText As String = "[EMPTY]")
        If Progress <= 0 Then
            Me.ProgressBar1.Value = 1
            Me.ProgressBar1.Value = 0
        Else
            Me.ProgressBar1.Value = Progress
            Me.ProgressBar1.Value = Progress - 1 'the ProgressBar control animates itself to expand to the value. this creates problems. If we move the progress backwards, the animation is not shown
        End If
        If StatusText <> "[EMPTY]" Then Me.Label_Status.Text = StatusText
        Me.Invalidate(True)
        Me.Update()
    End Sub


    Private Sub Form_Progress_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Globals.CenterForm(Me) 'center the form on the Excel Window
    End Sub
End Class

