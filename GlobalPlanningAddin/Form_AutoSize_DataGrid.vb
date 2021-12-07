Public Class Form_AutoSize_DataGrid
    Private Const MarginPx As Integer = 2
    Private Sub Form_KeyValues_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim TotalWidth As Integer = 0
        For Each Col As DataGridViewColumn In DataGridView1.Columns
            Col.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            TotalWidth += Col.Width
        Next
        DataGridView1.Width = TotalWidth

        Dim TotalHeight As Integer = 0
        For Each Rowx As DataGridViewRow In DataGridView1.Rows
            TotalHeight += Rowx.Height
        Next
        DataGridView1.Height = TotalHeight + IIf(DataGridView1.ColumnHeadersVisible = True, DataGridView1.ColumnHeadersHeight, 0)

        DataGridView1.Left = MarginPx
        DataGridView1.Top = MarginPx

        Me.Width = DataGridView1.Width + MarginPx * 2
        Me.Height = DataGridView1.Height + MarginPx * 2

        Me.Location = New System.Drawing.Point(Cursor.Position.X - Me.Width / 3, Cursor.Position.Y + 10)

    End Sub
    Private Sub Form_KeyValues_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        DataGridView1.ClearSelection()
        DataGridView1.CurrentCell = Nothing
    End Sub

    Sub New(DataRows As List(Of Object())) 'This New is used for the More Info Drop Down

        ' Cet appel est requis par le concepteur.
        InitializeComponent()
        For Each DataRow As Object() In DataRows
            DataGridView1.Rows.Add(DataRow)
        Next

    End Sub

    Sub New(Data_DataSet As DataSet) 'This new is used for the Change Log form

        ' Cet appel est requis par le concepteur.
        InitializeComponent()
        DataGridView1.Columns.Clear()
        DataGridView1.DataSource = Data_DataSet.Tables(0)
        DataGridView1.ColumnHeadersVisible = True

    End Sub

    'Private Sub Form_KeyValues_MouseLeave(sender As Object, e As EventArgs) Handles Me.MouseLeave
    '    Dim MouseX As Integer = MousePosition.X
    '    Dim MouseY As Integer = MousePosition.Y

    '    If MouseX > Me.Left + 5 And MouseX < Me.Right - 5 And
    '        MouseY > Me.Top + 5 And MouseY < Me.Bottom - 5 Then
    '        Exit Sub
    '    End If

    '    Me.Close()
    'End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_MouseLeave(sender As Object, e As EventArgs) Handles DataGridView1.MouseLeave
        Me.Close()
    End Sub

    Private Sub Form_KeyValues_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        'The color and the width of the border.
        Dim borderColor As Drawing.Color = Drawing.SystemColors.ScrollBar
        Dim borderWidth As Integer = 1
        Dim formRegion As Drawing.Rectangle
        formRegion = New Drawing.Rectangle(0, 0, Me.Width, Me.Height)

        'Draws the border.
        ControlPaint.DrawBorder(e.Graphics, formRegion,
                                borderColor, borderWidth, ButtonBorderStyle.Solid,
                                borderColor, borderWidth, ButtonBorderStyle.Solid,
                                borderColor, borderWidth, ButtonBorderStyle.Solid,
                                borderColor, borderWidth, ButtonBorderStyle.Solid)
    End Sub

    Private Sub DataGridView1_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseEnter

        If e.RowIndex = -1 Then
            'We are in the header
        Else
            For Each TheCell As DataGridViewCell In DataGridView1.Rows(e.RowIndex).Cells
                TheCell.Selected = True
            Next
        End If

    End Sub

    Private Sub DataGridView1_CellMouseLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseLeave

        If e.RowIndex = -1 Then
            'We left the header
        Else
            For Each TheCell As DataGridViewCell In DataGridView1.Rows(e.RowIndex).Cells
                TheCell.Selected = False
            Next
        End If

    End Sub

End Class