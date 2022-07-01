Imports ExcelDna.Integration
Imports System.Net
Imports System.Threading


Public Class Form_PluginInfos

    Private _NbBits As Integer
    Private _FormLoading As Boolean

    Private Sub Form_PluginInfos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        _FormLoading = True

        Globals.CenterForm(Me) 'center the form on the Excel Window

        If My.Application.Info.Version.Revision = Nothing Then 'Display current plugin version
            Label_Version.Text = "Version " & My.Application.Info.Version.ToString(3)
        Else
            Label_Version.Text = "Version " & My.Application.Info.Version.ToString()
        End If

        TextBox_PluginPath.Text = Globals.PluginInstallMgr.GetCurrentPluginLocation 'Display current plugin file path
        If IntPtr.Size = 4 Then 'Check how many bits is the Excel version we are running this plugin on
            _NbBits = 32
        ElseIf IntPtr.Size = 8 Then
            _NbBits = 64
        Else
            _NbBits = 0 'unknown!!!
        End If
        Label_bitness.Text = _NbBits.ToString & " bits"

        Btn_Install.Enabled = False 'for now, disable the install button until we know if the plugin is installed or not
        Button_CheckUpdates.Enabled = False 'we also disable the upgrade if the plugin is not yet installed

        'Fill the systems drop down list
        For Each PluginSystem In ListSystems
            ComboBox_System.Items.Add(PluginSystem.Name)
        Next
        ComboBox_System.SelectedItem = Current_Plugin_System.Name

        Btn_Close.Select() 'focus on the close button, to avoid that the text in the textbox is selected by default

        Globals.PluginInstallMgr.CheckPluginInstallStatus(AddressOf InstallStatusKnown)

        _FormLoading = False
    End Sub

    Friend Sub InstallStatusKnown()
        If Globals.PluginInstallMgr.PluginIsInstalled = False Then
            Label_Install_Status.Text = "Plugin is not installed"
            Btn_Install.Enabled = True
        Else
            Label_Install_Status.Text = "Plugin is installed correctly"
            Btn_Install.Enabled = False
            Button_CheckUpdates.Enabled = True
        End If

    End Sub

    Private Sub Btn_Install_Click(sender As Object, e As EventArgs) Handles Btn_Install.Click
        If Globals.PluginInstallMgr.PluginIsInstalled = False Then
            Globals.PluginInstallMgr.InstallPluginNow(AddressOf InstallFinished)
        End If
    End Sub

    Friend Sub InstallFinished()
        InstallStatusKnown()
        MsgBox(Globals.PluginInstallMgr.LastMessage)
        If Globals.PluginInstallMgr.PluginIsInstalled = True Then
            Globals.PluginEnabled = False
            Globals.ThisRibbon.Invalidate()
        End If
    End Sub

    Private Sub Button_CheckUpdates_Click(sender As Object, e As EventArgs) Handles Button_CheckUpdates.Click
        Globals.PluginInstallMgr.CheckUpdates()
        If Globals.PluginInstallMgr.LatestPluginVersion <> Globals.PluginInstallMgr.CurrentPluginVersion Then
            If MsgBox("A new version of the plugin is availalbe (v." & Globals.PluginInstallMgr.LatestPluginVersion.ToString & ")." & vbCrLf & "Do you want to upgrade now?", MsgBoxStyle.YesNo, "New version available") = MsgBoxResult.Yes Then
                Globals.PluginInstallMgr.UpdateToNewVersion(_NbBits)
                MsgBox(Globals.PluginInstallMgr.LastMessage, MsgBoxStyle.Information, "Plugin update")
            End If
        Else
            MsgBox("Your plugin is up to date.", MsgBoxStyle.Information, "Plugin update")
        End If

    End Sub

    Private Sub Btn_Close_Click(sender As Object, e As EventArgs) Handles Btn_Close.Click
        Me.Close()
    End Sub

    Private Sub Button_Changelog_Click(sender As Object, e As EventArgs) Handles Button_Changelog.Click
        Try
            Dim webClient = New WebClient
            Dim FileBuffer As Byte()

            FileBuffer = webClient.DownloadData(Globals.VersionCheckFolder & ChangelogFileName)
            Dim Changelogform As New Form_Changelog()
            Changelogform.SetText(FileBuffer)
            Changelogform.Show()
        Catch ex As Exception
            MsgBox("Unable to load the changelog: " & ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

    End Sub

    Private Sub ComboBox_System_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_System.SelectedIndexChanged

        If _FormLoading = True Then Exit Sub

        If NbGlobalPlanningAddinCompatibleWorkbooksOpen > 0 Then
            MsgBox("Please close all Global Planning addin compatible workbooks before changing the system.", MsgBoxStyle.Critical, "Global planning Addin")
            Exit Sub
        End If

        Current_Plugin_System = ListSystems.Find(Function(x) x.Name = ComboBox_System.SelectedItem.ToString)
        MsgBox("Successfully switched to " & Current_Plugin_System.Name & " system.", MsgBoxStyle.Information, "Global planning Addin")

        If Current_Plugin_System.ID = 0 Then
            ExcelAsyncUtil.QueueAsMacro(AddressOf CurRibbonActions.CheckUpdatesDone) 'Since we need COM introp, make sure this is run in the main ui thread
        Else
            ExcelAsyncUtil.QueueAsMacro(AddressOf CurRibbonActions.Setup_TestSystemInfoButton) 'Since we need COM introp, make sure this is run in the main ui thread
        End If

    End Sub



    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    If Globals.PluginInstallMgr.PluginIsInstalled = True Then
    '        Globals.PluginInstallMgr.UnInstallPluginNow(AddressOf UnInstallFinished)
    '    End If
    'End Sub

    'Friend Sub UnInstallFinished()
    '    InstallStatusKnown()
    '    MsgBox(Globals.PluginInstallMgr.LastMessage)
    '    If Globals.PluginInstallMgr.PluginIsInstalled = False Then
    '        'Globals.PluginEnabled = False
    '        Globals.ThisRibbon.Invalidate()
    '    End If
    'End Sub

End Class
