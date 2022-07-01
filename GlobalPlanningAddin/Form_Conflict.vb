Public Class Form_Conflict

    Public Property UserDecision As String
    Public Property RememberChoice As Boolean


    Sub New(KeyValues() As String, FieldName As String, UserChanges As String, OtherChanges_DataSet As DataSet)
        InitializeComponent() ' Cet appel est requis par le concepteur.

        Dim KeyValuesConcatStr As String = ""
        For i As Integer = 0 To UBound(KeyValues) 'Concat the Key fields values
            KeyValuesConcatStr &= KeyValues(i)
            If i < UBound(KeyValues) Then KeyValuesConcatStr &= "/"
        Next

        TextBox_SKU.Text = KeyValuesConcatStr
        TextBox_Field.Text = FieldName
        RichTextBox_UserModification.Text = UserChanges

        DataGridView_OtherChanges.Columns.Clear()
        DataGridView_OtherChanges.DataSource = OtherChanges_DataSet.Tables(0)
        DataGridView_OtherChanges.ColumnHeadersVisible = True

        _RememberChoice = False
    End Sub
    Private Sub Button_Overwrite_Click(sender As Object, e As EventArgs) Handles Button_Overwrite.Click
        _UserDecision = "OVERWRITE"
        Me.Close()
    End Sub
    Private Sub Button_Abandon_Click(sender As Object, e As EventArgs) Handles Button_Abandon.Click
        _UserDecision = "ABANDON"
        Me.Close()
    End Sub
    Private Sub CheckBox_RememberChoice_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_RememberChoice.CheckedChanged
        _RememberChoice = CheckBox_RememberChoice.Checked
    End Sub

    Private Sub Form_Conflict_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Globals.CenterForm(Me) 'center the form on the Excel Window
    End Sub
End Class