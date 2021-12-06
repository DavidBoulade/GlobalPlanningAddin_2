Public Class Form_Changelog
    Public Sub SetText(NewText As Byte())
        Using ms As New IO.MemoryStream(NewText)
            TextBox_Changelog.LoadFile(ms, RichTextBoxStreamType.PlainText)
        End Using
    End Sub

    Private Sub Form_Changelog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Globals.CenterForm(Me)
    End Sub
End Class