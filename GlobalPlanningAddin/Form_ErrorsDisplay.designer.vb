Imports System.Windows.Forms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form_ErrorsDisplay
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
        Me.ListBox_Errors = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'ListBox_Errors
        '
        Me.ListBox_Errors.FormattingEnabled = True
        Me.ListBox_Errors.Location = New System.Drawing.Point(12, 12)
        Me.ListBox_Errors.Name = "ListBox_Errors"
        Me.ListBox_Errors.Size = New System.Drawing.Size(658, 303)
        Me.ListBox_Errors.TabIndex = 0
        '
        'Form_ErrorsDisplay
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(683, 331)
        Me.Controls.Add(Me.ListBox_Errors)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form_ErrorsDisplay"
        Me.ShowIcon = False
        Me.Text = "Errors while saving changes to the database"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ListBox_Errors As ListBox
End Class
