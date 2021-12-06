Public Class Form_ErrorsDisplay
    Sub New(Errors As List(Of String))


        InitializeComponent() ' Cet appel est requis par le concepteur.

        For Each ErrorStr As String In Errors
            ListBox_Errors.Items.Add(ErrorStr)
        Next

    End Sub

    Private Sub Form_ErrorsDisplay_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Globals.CenterForm(Me) 'center the form on the Excel Window
    End Sub
End Class