<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form_PluginInfos
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label_Version = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox_PluginPath = New System.Windows.Forms.TextBox()
        Me.Btn_Install = New System.Windows.Forms.Button()
        Me.Button_CheckUpdates = New System.Windows.Forms.Button()
        Me.Btn_Close = New System.Windows.Forms.Button()
        Me.Label_Install_Status = New System.Windows.Forms.Label()
        Me.Label_bitness = New System.Windows.Forms.Label()
        Me.Button_Changelog = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(150, 15)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Global Planning Addin"
        '
        'Label_Version
        '
        Me.Label_Version.AutoSize = True
        Me.Label_Version.Location = New System.Drawing.Point(13, 38)
        Me.Label_Version.Name = "Label_Version"
        Me.Label_Version.Size = New System.Drawing.Size(69, 13)
        Me.Label_Version.TabIndex = 1
        Me.Label_Version.Text = "Version 1.0.0"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(17, 111)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Plugin path"
        '
        'TextBox_PluginPath
        '
        Me.TextBox_PluginPath.Location = New System.Drawing.Point(93, 108)
        Me.TextBox_PluginPath.Name = "TextBox_PluginPath"
        Me.TextBox_PluginPath.ReadOnly = True
        Me.TextBox_PluginPath.Size = New System.Drawing.Size(169, 20)
        Me.TextBox_PluginPath.TabIndex = 3
        '
        'Btn_Install
        '
        Me.Btn_Install.Location = New System.Drawing.Point(164, 80)
        Me.Btn_Install.Name = "Btn_Install"
        Me.Btn_Install.Size = New System.Drawing.Size(98, 25)
        Me.Btn_Install.TabIndex = 4
        Me.Btn_Install.Text = "Install"
        Me.Btn_Install.UseVisualStyleBackColor = True
        '
        'Button_CheckUpdates
        '
        Me.Button_CheckUpdates.Location = New System.Drawing.Point(20, 144)
        Me.Button_CheckUpdates.Name = "Button_CheckUpdates"
        Me.Button_CheckUpdates.Size = New System.Drawing.Size(117, 24)
        Me.Button_CheckUpdates.TabIndex = 5
        Me.Button_CheckUpdates.Text = "Check updates..."
        Me.Button_CheckUpdates.UseVisualStyleBackColor = True
        '
        'Btn_Close
        '
        Me.Btn_Close.Location = New System.Drawing.Point(145, 144)
        Me.Btn_Close.Name = "Btn_Close"
        Me.Btn_Close.Size = New System.Drawing.Size(117, 24)
        Me.Btn_Close.TabIndex = 6
        Me.Btn_Close.Text = "Close"
        Me.Btn_Close.UseVisualStyleBackColor = True
        '
        'Label_Install_Status
        '
        Me.Label_Install_Status.AutoSize = True
        Me.Label_Install_Status.Location = New System.Drawing.Point(17, 86)
        Me.Label_Install_Status.Name = "Label_Install_Status"
        Me.Label_Install_Status.Size = New System.Drawing.Size(96, 13)
        Me.Label_Install_Status.TabIndex = 7
        Me.Label_Install_Status.Text = "Plugin install status"
        '
        'Label_bitness
        '
        Me.Label_bitness.AutoSize = True
        Me.Label_bitness.Location = New System.Drawing.Point(97, 38)
        Me.Label_bitness.Name = "Label_bitness"
        Me.Label_bitness.Size = New System.Drawing.Size(27, 13)
        Me.Label_bitness.TabIndex = 8
        Me.Label_bitness.Text = "0 bit"
        '
        'Button_Changelog
        '
        Me.Button_Changelog.Location = New System.Drawing.Point(164, 32)
        Me.Button_Changelog.Name = "Button_Changelog"
        Me.Button_Changelog.Size = New System.Drawing.Size(98, 24)
        Me.Button_Changelog.TabIndex = 9
        Me.Button_Changelog.Text = "Changelog"
        Me.Button_Changelog.UseVisualStyleBackColor = True
        '
        'Form_PluginInfos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(274, 182)
        Me.Controls.Add(Me.Button_Changelog)
        Me.Controls.Add(Me.Label_bitness)
        Me.Controls.Add(Me.Label_Install_Status)
        Me.Controls.Add(Me.Btn_Close)
        Me.Controls.Add(Me.Button_CheckUpdates)
        Me.Controls.Add(Me.Btn_Install)
        Me.Controls.Add(Me.TextBox_PluginPath)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label_Version)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "Form_PluginInfos"
        Me.Text = "Plugin Infos"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label_Version As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents TextBox_PluginPath As TextBox
    Friend WithEvents Btn_Install As Button
    Friend WithEvents Button_CheckUpdates As Button
    Friend WithEvents Btn_Close As Button
    Friend WithEvents Label_Install_Status As Label
    Friend WithEvents Label_bitness As Label
    Friend WithEvents Button_Changelog As Button
End Class
