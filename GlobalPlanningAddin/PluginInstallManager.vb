Option Explicit On
Option Strict On
Imports ExcelDna.Integration
Imports System.IO
Imports System.Net
Imports System.Threading
Public Class PluginInstallManager

    ''' <summary>Full path to this plugin instance xll file excluding the file name</summary>
    Private ReadOnly _xllPath As String
    ''' <summary>Path to the Excel custom plugins folder + install directory</summary>
    Private ReadOnly _InstallPath As String
    ''' <summary>Current xll file name for this instance of the plugin</summary>
    Private ReadOnly _CurrentXllFileName As String




    Sub New()
        _InstallPath = Globals.ExcelApplication.UserLibraryPath & Globals.PluginXllInstallSubFolder
        _xllPath = ExcelDnaUtil.XllPath
        _CurrentXllFileName = System.IO.Path.GetFileName(_xllPath)
        _xllPath = Strings.Left(_xllPath, Len(_xllPath) - Len(_CurrentXllFileName)) 'keep the path only (don't use IO.path.GetDirectory as a \ may be missing if using c:\)
    End Sub

    Friend Function GetCurrentPluginLocation() As String
        Return _xllPath '& _CurrentXllFileName
    End Function


    ''' <summary>True if the Plugin is installed</summary>
    Friend Property PluginIsInstalled As Nullable(Of Boolean) = Nothing 'Nullable(Of Boolean) can also be written "Boolean?"
    Friend Sub CheckPluginInstallStatus(Callback As Action)
        ExcelAsyncUtil.QueueAsMacro(AddressOf CheckIfPluginIsInstalled) 'Since we need COM introp, make sure this is run in the main ui thread
        _CallbackCheckPluginInstallStatus = Callback
    End Sub
    Private _CallbackCheckPluginInstallStatus As Action
    Private Sub CheckIfPluginIsInstalled()
        Dim IsPluginInstalled As Boolean = False
        For Each theaddin As Microsoft.Office.Interop.Excel.AddIn In Globals.ExcelApplication.AddIns2
            Try
                If theaddin.Name = Globals.PluginXllInstalledFileName Then ' theaddin.FullName = _xllPath & _CurrentXllFileName
                    If theaddin.Installed = True Then IsPluginInstalled = True
                End If 'continue looping as it is possible to have multiple instances of this plugin
            Catch
                'sometimes we can't read the plugin name??? don't know why but simply ignore it
            End Try

        Next
        _PluginIsInstalled = IsPluginInstalled
        _CallbackCheckPluginInstallStatus()
    End Sub


    ''' <summary>Output message of the last action</summary>
    Friend Property LastMessage As String = ""
    Friend Sub InstallPluginNow(Callback As Action)

        ExcelAsyncUtil.QueueAsMacro(AddressOf InstallNow) 'Since we need COM introp, make sure this is run in the main ui thread
        _CallbackInstallPluginNow = Callback
    End Sub
    Private _CallbackInstallPluginNow As Action
    Private Sub InstallNow()

        _LastMessage = ""
        Dim TempWb As Microsoft.Office.Interop.Excel.Workbook = Nothing

        If _PluginIsInstalled Is Nothing Then
            'we need to check first if the plugin is alreday installed
            _LastMessage = "Internal error: The install status was not checked before attempting to install the plugin"
        Else
            If _PluginIsInstalled = False Then
                'Let's install it!

                'Before copying the xll, verify the destination file doesn't exist already
                If Globals.DoesFileExists(_InstallPath, Globals.PluginXllInstalledFileName) Then

                    'if an old copy also exists, we need to delete it before
                    If Globals.DoesFileExists(_InstallPath, Globals.PluginXllInstalledFileName & ".old") Then System.IO.File.Delete(_InstallPath & Globals.PluginXllInstalledFileName & ".old")

                    'Save to existing plugin
                    Globals.RenameFile(_InstallPath, Globals.PluginXllInstalledFileName, Globals.PluginXllInstalledFileName & ".old")
                End If

                Try
                    If System.IO.Directory.Exists(_InstallPath) = False Then System.IO.Directory.CreateDirectory(_InstallPath) 'Create the folder if doesn't exist
                    System.IO.File.Copy(_xllPath & _CurrentXllFileName, _InstallPath & Globals.PluginXllInstalledFileName) 'Copy the xll to the install folder
                Catch ex As Exception
                    _LastMessage = "Error: " & ex.Message
                    _CallbackInstallPluginNow()
                    Exit Sub
                End Try

                'and add the plugin
                Try
                    'Plugin install fails if no workbook is open!!?? so open a new one if none is open
                    If Globals.ExcelApplication.Workbooks.Count = 0 Then TempWb = Globals.ExcelApplication.Workbooks.Add()

                    Dim NewInstalledAddin As Microsoft.Office.Interop.Excel.AddIn
                    NewInstalledAddin = Globals.ExcelApplication.AddIns.Add(_InstallPath & Globals.PluginXllInstalledFileName, False)
                    NewInstalledAddin.Installed = True 'this actually activates the plugin
                    _PluginIsInstalled = True
                    _LastMessage = "Plugin installed ok. Please restart excel."
                    If Not (TempWb Is Nothing) Then TempWb.Close(False) 'Close the temp workbook if it was open
                    _CallbackInstallPluginNow()
                    Exit Sub
                Catch ex As Exception
                    _LastMessage = "Error: " & ex.Message
                    If Not (TempWb Is Nothing) Then TempWb.Close(False)
                    _CallbackInstallPluginNow()
                    Exit Sub
                End Try

            Else
                _LastMessage = "Internal error: The plugin is already installed"
            End If

        End If



        _CallbackInstallPluginNow()
    End Sub


    'The uninstal is disregarded for now as Excel crashes when calling Addin.installed=False
    'Friend Sub UnInstallPluginNow(Callback As Action)
    '    ExcelAsyncUtil.QueueAsMacro(AddressOf UnInstallNow) 'Since we need COM introp, make sure this is run in the main ui thread
    '    _CallbackUNInstallPluginNow = Callback
    'End Sub
    'Private _CallbackUNInstallPluginNow As Action
    'Private Sub UnInstallNow()

    '    _LastMessage = ""

    '    If _PluginIsInstalled Is Nothing Then
    '        'we need to check first if the plugin is alreday installed
    '        _LastMessage = "Internal error: The install status was not checked before attempting to install the plugin"
    '    Else
    '        If _PluginIsInstalled = True Then
    '            'Let's uninstall it

    '            'Shearch for the installed plugin
    '            For Each theaddin As Microsoft.Office.Interop.Excel.AddIn In Globals.ExcelApplication.AddIns2
    '                If UCase(theaddin.Name) = UCase(Globals.PluginXllInstalledFileName) Then
    '                    'we found an instance of this plugin, now check if it is the open installed

    '                    If theaddin.Installed = True Then
    '                        'We found it, now uninstall
    '                        theaddin.Installed = False 'this deactivates the plugin, should we also delete from the list? not easy...
    '                        _LastMessage = "Plugin uninstalled correctly"
    '                        _PluginIsInstalled = False
    '                        Exit For
    '                    End If


    '                End If
    '            Next

    '        Else
    '            _LastMessage = "Internal error: The plugin is not installed"
    '        End If

    '    End If

    '    _CallbackUNInstallPluginNow()
    'End Sub




    'Private Sub IsThereAnotherRunningInstance()  ->>>> this won't work as we can't RELIABLY identify if another addin is from the same type of this addin

    '    _AnotherInstanceIsOpen = False
    '    For Each theaddin As Microsoft.Office.Interop.Excel.AddIn In Globals.ExcelApplication.AddIns2
    '        If UCase(theaddin.Name) = UCase(Globals.PluginXllTitle) Then
    '            'we found an instance of this plugin, now check if it is this instance
    '            If theaddin.FullName <> _xllPath & _CurrentXllFileName Then
    '                'We found another instance
    '                If theaddin.IsOpen = True Then
    '                    _AnotherInstanceIsOpen = True
    '                    Exit For
    '                End If
    '            End If
    '        End If
    '    Next
    '    _CallbackIsThereAnotherRunningInstance()

    'End Sub

    Friend Property CurrentPluginVersion As Version = My.Application.Info.Version
    Friend Property LatestPluginVersion As Version = My.Application.Info.Version
    Private _CheckUpdateCallback As Action
    Friend Sub CheckUpdates(Callback As Action)
        _CheckUpdateCallback = Callback
        'Thread.Sleep(5000)
        CheckUpdates()
        ExcelAsyncUtil.QueueAsMacro(AddressOf CheckUpdatesDone)
    End Sub
    Private Sub CheckUpdatesDone() 'this is back to the main UI thread, safe to call the given callback
        _CheckUpdateCallback()
    End Sub

    Friend Sub CheckUpdates()
        Dim webClient = New WebClient

        Try 'find the latest version
            _LatestPluginVersion = New Version(webClient.DownloadString(Globals.VersionCheckFolder & Globals.LatestVersionInfoFileName)) '.Split('\n')[0];

        Catch ex As Exception
            _LastMessage = "Could not check for updates: " & ex.Message
            _LatestPluginVersion = My.Application.Info.Version
            Exit Sub
        End Try

        'If (_LatestPluginVersion.Equals(My.Application.Info.Version.ToString)) Then
        '    _LastMessage = "No updates available."
        'End If

    End Sub

    Friend Sub UpdateToNewVersion(NbBit As Integer)
        Dim webClient = New WebClient
        Dim NewXllFileBuffer As Byte()

        If NbBit <> 32 And NbBit <> 64 Then
            _LastMessage = "Error: only 32 and 64 bit are supported!"
            Exit Sub
        End If
        If (_LatestPluginVersion.Equals(My.Application.Info.Version.ToString)) Then
            _LastMessage = "Error: No new version availalbe"
        Else

            Try 'Get latest xll
                If NbBit = 32 Then
                    NewXllFileBuffer = webClient.DownloadData(Globals.VersionCheckFolder & Globals.PluginXll32InstallerFileName)
                Else '64
                    NewXllFileBuffer = webClient.DownloadData(Globals.VersionCheckFolder & Globals.PluginXll64InstallerFileName)
                End If

            Catch ex As Exception
                _LastMessage = "Could not download latest version:" + ex.Message
                Return
            End Try

            'Before copying the new xll, verify the destination file doesn't exist already, and normally it should!
            If Globals.DoesFileExists(_InstallPath, Globals.PluginXllInstalledFileName) Then

                'if an old copy also exists, we need to delete it before
                If Globals.DoesFileExists(_InstallPath, Globals.PluginXllInstalledFileName & ".old") Then System.IO.File.Delete(_InstallPath & Globals.PluginXllInstalledFileName & ".old")

                'Save to existing plugin
                Globals.RenameFile(_InstallPath, Globals.PluginXllInstalledFileName, Globals.PluginXllInstalledFileName & ".old")
            End If


            Try 'write new xll version
                Dim fs As FileStream = New FileStream(_InstallPath & Globals.PluginXllInstalledFileName, FileMode.Create)
                fs.Write(NewXllFileBuffer, 0, NewXllFileBuffer.Length)
                fs.Close()
                _LastMessage = "Restart Excel to upgrade to " + _LatestPluginVersion.ToString + "."
            Catch ex As Exception
                'restore old xll
                If Globals.DoesFileExists(_InstallPath, Globals.PluginXllInstalledFileName & ".old") Then Globals.RenameFile(_InstallPath, Globals.PluginXllInstalledFileName & ".old", Globals.PluginXllInstalledFileName)

                _LastMessage = "Could not write updated plug-in: " + ex.Message
                Return
            End Try

            _LastMessage = "The plugin has been upgraded to the latest version (" + _LatestPluginVersion.ToString + "). The changes will take effect once you restart Excel."

            _CurrentPluginVersion = _LatestPluginVersion
        End If

    End Sub

End Class
