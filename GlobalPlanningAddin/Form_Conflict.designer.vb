Imports System.Windows.Forms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form_Conflict
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox_SKU = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox_Field = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.RichTextBox_UserModification = New System.Windows.Forms.RichTextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button_Overwrite = New System.Windows.Forms.Button()
        Me.Button_Abandon = New System.Windows.Forms.Button()
        Me.CheckBox_RememberChoice = New System.Windows.Forms.CheckBox()
        Me.DataGridView_OtherChanges = New System.Windows.Forms.DataGridView()
        Me.Field = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Value = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.DataGridView_OtherChanges, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(149, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "A conflict has been found for :"
        '
        'TextBox_SKU
        '
        Me.TextBox_SKU.Enabled = False
        Me.TextBox_SKU.Location = New System.Drawing.Point(167, 6)
        Me.TextBox_SKU.Name = "TextBox_SKU"
        Me.TextBox_SKU.Size = New System.Drawing.Size(259, 20)
        Me.TextBox_SKU.TabIndex = 1
        Me.TextBox_SKU.Text = "123456789123456789@1234"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(126, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Field :"
        '
        'TextBox_Field
        '
        Me.TextBox_Field.Enabled = False
        Me.TextBox_Field.Location = New System.Drawing.Point(167, 33)
        Me.TextBox_Field.Name = "TextBox_Field"
        Me.TextBox_Field.Size = New System.Drawing.Size(259, 20)
        Me.TextBox_Field.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 79)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(159, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "You made the following change:"
        '
        'RichTextBox_UserModification
        '
        Me.RichTextBox_UserModification.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox_UserModification.Location = New System.Drawing.Point(15, 96)
        Me.RichTextBox_UserModification.Name = "RichTextBox_UserModification"
        Me.RichTextBox_UserModification.Size = New System.Drawing.Size(557, 50)
        Me.RichTextBox_UserModification.TabIndex = 5
        Me.RichTextBox_UserModification.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(15, 190)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(289, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "but in the meantime the following change have been made :"
        '
        'Button_Overwrite
        '
        Me.Button_Overwrite.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button_Overwrite.Location = New System.Drawing.Point(401, 152)
        Me.Button_Overwrite.Name = "Button_Overwrite"
        Me.Button_Overwrite.Size = New System.Drawing.Size(171, 23)
        Me.Button_Overwrite.TabIndex = 8
        Me.Button_Overwrite.Text = "Overwrite (keep my changes)"
        Me.Button_Overwrite.UseVisualStyleBackColor = True
        '
        'Button_Abandon
        '
        Me.Button_Abandon.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button_Abandon.Location = New System.Drawing.Point(401, 322)
        Me.Button_Abandon.Name = "Button_Abandon"
        Me.Button_Abandon.Size = New System.Drawing.Size(171, 23)
        Me.Button_Abandon.TabIndex = 9
        Me.Button_Abandon.Text = "Abandon (don't change)"
        Me.Button_Abandon.UseVisualStyleBackColor = True
        '
        'CheckBox_RememberChoice
        '
        Me.CheckBox_RememberChoice.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.CheckBox_RememberChoice.AutoSize = True
        Me.CheckBox_RememberChoice.Location = New System.Drawing.Point(15, 352)
        Me.CheckBox_RememberChoice.Name = "CheckBox_RememberChoice"
        Me.CheckBox_RememberChoice.Size = New System.Drawing.Size(218, 17)
        Me.CheckBox_RememberChoice.TabIndex = 10
        Me.CheckBox_RememberChoice.Text = "Remember my choice for further conflicts"
        Me.CheckBox_RememberChoice.UseVisualStyleBackColor = True
        '
        'DataGridView_OtherChanges
        '
        Me.DataGridView_OtherChanges.AllowUserToAddRows = False
        Me.DataGridView_OtherChanges.AllowUserToDeleteRows = False
        Me.DataGridView_OtherChanges.AllowUserToResizeRows = False
        Me.DataGridView_OtherChanges.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView_OtherChanges.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DataGridView_OtherChanges.BackgroundColor = System.Drawing.SystemColors.Window
        Me.DataGridView_OtherChanges.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView_OtherChanges.ColumnHeadersVisible = False
        Me.DataGridView_OtherChanges.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Field, Me.Value})
        Me.DataGridView_OtherChanges.Location = New System.Drawing.Point(12, 206)
        Me.DataGridView_OtherChanges.Name = "DataGridView_OtherChanges"
        Me.DataGridView_OtherChanges.RowHeadersVisible = False
        Me.DataGridView_OtherChanges.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView_OtherChanges.ShowEditingIcon = False
        Me.DataGridView_OtherChanges.Size = New System.Drawing.Size(560, 110)
        Me.DataGridView_OtherChanges.TabIndex = 11
        '
        'Field
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle9.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        DataGridViewCellStyle9.ForeColor = System.Drawing.SystemColors.GrayText
        DataGridViewCellStyle9.Padding = New System.Windows.Forms.Padding(2, 0, 0, 0)
        DataGridViewCellStyle9.SelectionBackColor = System.Drawing.SystemColors.ScrollBar
        DataGridViewCellStyle9.SelectionForeColor = System.Drawing.Color.Black
        Me.Field.DefaultCellStyle = DataGridViewCellStyle9
        Me.Field.HeaderText = "Field"
        Me.Field.Name = "Field"
        Me.Field.Width = 5
        '
        'Value
        '
        DataGridViewCellStyle10.ForeColor = System.Drawing.SystemColors.GrayText
        DataGridViewCellStyle10.Padding = New System.Windows.Forms.Padding(0, 0, 2, 0)
        DataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.ScrollBar
        DataGridViewCellStyle10.SelectionForeColor = System.Drawing.Color.Black
        Me.Value.DefaultCellStyle = DataGridViewCellStyle10
        Me.Value.HeaderText = "Value"
        Me.Value.Name = "Value"
        Me.Value.Width = 5
        '
        'Form_Conflict
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(584, 381)
        Me.ControlBox = False
        Me.Controls.Add(Me.DataGridView_OtherChanges)
        Me.Controls.Add(Me.CheckBox_RememberChoice)
        Me.Controls.Add(Me.Button_Abandon)
        Me.Controls.Add(Me.Button_Overwrite)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.RichTextBox_UserModification)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox_Field)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox_SKU)
        Me.Controls.Add(Me.Label1)
        Me.MinimumSize = New System.Drawing.Size(450, 360)
        Me.Name = "Form_Conflict"
        Me.Text = "Conflict found"
        Me.TopMost = True
        CType(Me.DataGridView_OtherChanges, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox_SKU As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents TextBox_Field As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents RichTextBox_UserModification As RichTextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Button_Overwrite As Button
    Friend WithEvents Button_Abandon As Button
    Friend WithEvents CheckBox_RememberChoice As CheckBox
    Friend WithEvents DataGridView_OtherChanges As DataGridView
    Friend WithEvents Field As DataGridViewTextBoxColumn
    Friend WithEvents Value As DataGridViewTextBoxColumn
End Class
