Imports System.Runtime.InteropServices
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop.Excel
'Imports System.Threading

''' <summary>Global variables and functions for the Addin</summary>
Public Module Globals
    ' general Global objects/variables
    Friend Const PluginXllInstallSubFolder As String = "Global Planning Addin\"
    Friend Const PluginXll32InstallerFileName As String = "GlobalPlanningAddIn-packed.xll"
    Friend Const PluginXll64InstallerFileName As String = "GlobalPlanningAddIn64-packed.xll"
    Friend Const PluginXllInstalledFileName As String = "GlobalPlanningAddIn.xll"
    Friend Const VersionCheckFolder As String = "\\USSANTFILE01\shared\Global Planning\Global Planning AddIn\"
    Friend Const LatestVersionInfoFileName As String = "version.txt"
    Friend Const ChangelogFileName As String = "changelog.txt"
    Friend Const TemplatesMenuXmlFileName As String = "templates.xml"
    Friend Const TemplatesSubFolder As String = "Templates\"

    Friend Property PluginInstallMgr As PluginInstallManager

    Friend ReadOnly Property PluginVersion As Version = My.Application.Info.Version
    Friend Property PluginEnabled As Boolean = True

    ''' <summary>reference object for the Addins ribbon</summary>
    Friend Property ThisRibbon As CustomUI.IRibbonUI
    Friend Property CurRibbonActions As RibbonActions

    Friend Property ExcelApplication As Microsoft.Office.Interop.Excel.Application
    Friend Property ExcelApplication_UserLibraryPath As String

    Friend Property WindowsHandle As IntPtr

    Private ReadOnly _WorkbooksData As New List(Of OpenWorkbookData)
    Private _ThisWorkbookData As OpenWorkbookData
    Friend ReadOnly Property ThisWorkbook As Microsoft.Office.Interop.Excel.Workbook
        Get
            If Not (_ThisWorkbookData Is Nothing) Then
                Return _ThisWorkbookData.Workbook
            Else
                Return Nothing
            End If
        End Get
    End Property

    Friend Property ReportDate As Date
        Get
            If Not (_ThisWorkbookData Is Nothing) Then
                Return _ThisWorkbookData.ReportDate
            Else
                Return Nothing
            End If
        End Get
        Set(value As Date)
            If Not (_ThisWorkbookData Is Nothing) Then _ThisWorkbookData.ReportDate = value
        End Set
    End Property

    Friend Property Reader As DatabaseReader
        Get
            If Not (_ThisWorkbookData Is Nothing) Then
                Return _ThisWorkbookData.Reader
            Else
                Return Nothing
            End If
        End Get
        Set(value As DatabaseReader)
            If Not (_ThisWorkbookData Is Nothing) Then _ThisWorkbookData.Reader = value
        End Set
    End Property

    Friend ReadOnly Property DatabaseReaderType As String
        Get
            If Not (_ThisWorkbookData Is Nothing) Then
                Return _ThisWorkbookData.DatabaseReaderType
            Else
                Return Nothing
            End If
        End Get
    End Property

    Friend ReadOnly Property ConfigSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Friend ReadOnly Property ReportSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
    Friend ReadOnly Property DetailsSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing


    Public Sub WorkbookActivated(Wb As Microsoft.Office.Interop.Excel.Workbook)
        Dim ThisIsANewWorkbook As Boolean = False


        'Check if this workbook is a workbook this plugin can handle
        If GetCustomDocumentProperty(Wb, "CustomDocType") = "SKUAlertsUI" Or GetCustomDocumentProperty(Wb, "CustomDocType") = "GRUT_UI" Then

            'check if it is the first time we see this workbook
            If _WorkbooksData.Exists(Function(x) x.Workbook Is Wb) = False Then
                _WorkbooksData.Add(New OpenWorkbookData(Wb)) 'if yes, add in the list
                ThisIsANewWorkbook = True

                'Check also the version of the required Plugin
            End If

            _ThisWorkbookData = _WorkbooksData.Find(Function(x) x.Workbook Is Wb)
            _ThisWorkbookData.DatabaseReaderType = GetCustomDocumentProperty(Wb, "CustomDocType")

            'Create a reference to the SKUAlerts worksheets
            For Each wrksheet As Microsoft.Office.Interop.Excel.Worksheet In _ThisWorkbookData.Workbook.Sheets
                Select Case Globals.GetCustomWorksheetProperty(wrksheet, "CustomSheetType")
                    Case "SKUAlertsConfig", "GRUTConfig"
                        _ConfigSheet = wrksheet
                    Case "SKUAlertsReport", "GRUTReport"
                        _ReportSheet = wrksheet
                    Case "SKUAlertsDetails", "GRUTDetails"
                        _DetailsSheet = wrksheet
                End Select
            Next

            'Verify that the references have been found
            If _ConfigSheet Is Nothing Then Throw New System.Exception("Unable to get a reference to the config worksheet")
            If _ReportSheet Is Nothing Then Throw New System.Exception("Unable to get a reference to the report worksheet")
            If _DetailsSheet Is Nothing Then Throw New System.Exception("Unable to get a reference to the detailed view worksheet")


            _CurRibbonActions.TemplateLoaded(GetCustomDocumentProperty(Wb, "TemplateID"), GetCustomDocumentProperty(Wb, "TemplateVersion"))

        Else
            _ThisWorkbookData = Nothing
            _CurRibbonActions.NonCompatibleFileLoaded()
        End If

    End Sub

    Friend Sub WorkbookClosed(Wb As Microsoft.Office.Interop.Excel.Workbook)
        'check if that workbook is in the list of workbooks this plugin handles
        If _WorkbooksData.Exists(Function(x) x.Workbook Is Wb) = True Then
            _ThisWorkbookData = _WorkbooksData.Find(Function(x) x.Workbook Is Wb) 'normally this should not be needed... but just to be sure
            'Reader = _ThisWorkbookData.Reader 'normally this should not be needed... but just to be sure
            If Not (_ThisWorkbookData.Reader Is Nothing) Then 'Dispose the reader
                _ThisWorkbookData.Reader.Dispose()
            End If
            _WorkbooksData.Remove(_ThisWorkbookData) 'remove from the list
        End If
    End Sub

    Public Function GetCustomDocumentProperty(Wb As Microsoft.Office.Interop.Excel.Workbook, PropertyName As String) As String
        For Each CustomDocProperty As Microsoft.Office.Core.DocumentProperty In DirectCast(Wb.CustomDocumentProperties, Microsoft.Office.Core.DocumentProperties)
            If CustomDocProperty.Name = PropertyName Then
                Return CustomDocProperty.Value
            End If
        Next
        Return ""
    End Function

    Public Function GetCustomWorksheetProperty(Ws As Microsoft.Office.Interop.Excel.Worksheet, PropertyName As String) As String
        For Each CustomNamedRange As Microsoft.Office.Interop.Excel.Name In Ws.Names
            If Strings.InStr(UCase(CustomNamedRange.Name), UCase(PropertyName)) <> 0 Then Return CustomNamedRange.Comment
        Next
        Return ""
    End Function

    Public Sub SetWorksheetNamedRangeValue(Ws As Microsoft.Office.Interop.Excel.Worksheet, NamedRangeName As String, NewValue As String)
        Dim FoundSw As Boolean = False
        For Each CustomNamedRange As Microsoft.Office.Interop.Excel.Name In Ws.Names
            If Strings.InStr(UCase(CustomNamedRange.Name), UCase(NamedRangeName)) <> 0 Then
                'The Name already exists, just update it
                CustomNamedRange.RefersTo = NewValue
                FoundSw = True
                Exit For
            End If
        Next
        If FoundSw = False Then
            'The name doesn't exist, create it now
            Dim NewName As Name = Ws.Names.Add(NamedRangeName, NewValue, True)
        End If
    End Sub


    'This function is needed as the standard .fileExists() function doesn't always work properly over the network
    Public Function DoesFileExists(FolderPath As String, FileName As String) As Boolean
        Dim FileNamesInFolder() As String
        Dim i As Integer

        Try
            FileNamesInFolder = System.IO.Directory.GetFiles(FolderPath)
            For i = 0 To UBound(FileNamesInFolder)
                If FileNamesInFolder(i) = FolderPath & FileName Then Return True
            Next i
            Return False
        Catch
            Return False
        End Try

    End Function

    Public Sub RenameFile(FolderPath As String, OldFile As String, NewFile As String)
        FileSystem.Rename(FolderPath & OldFile, FolderPath & NewFile)
        If DoesFileExists(FolderPath, OldFile) = True Then
            System.IO.File.Delete(FolderPath & OldFile)
        End If
    End Sub

    '********************************************************************************************************************
    '*** Code to center a form on the Excel Window
    '********************************************************************************************************************
    <StructLayout(LayoutKind.Sequential)> Public Structure RECT
        Dim Left As Integer
        Dim Top As Integer
        Dim Right As Integer
        Dim Bottom As Integer
    End Structure

    Public Declare Function GetWindowRect Lib "user32" (ByVal HWND As Integer, ByRef lpRect As RECT) As Integer
    Public Sub CenterForm(FormToCenter As System.Windows.Forms.Form)
        Try
            Dim xlWindowRect As RECT = New RECT()
            GetWindowRect(ExcelDnaUtil.WindowHandle, xlWindowRect)
            Dim X As Integer = (xlWindowRect.Right - xlWindowRect.Left - FormToCenter.Width) / 2 + xlWindowRect.Left
            Dim Y As Integer = (xlWindowRect.Bottom - xlWindowRect.Top - FormToCenter.Height) / 2 + xlWindowRect.Top

            FormToCenter.StartPosition = FormStartPosition.Manual
            FormToCenter.Location = New System.Drawing.Point(X, Y)
        Catch ex As Exception
            'No error should happen
        End Try

    End Sub

    Function IsEditing() As Boolean

        If _ExcelApplication.Interactive = False Then Return False
        Try
            _ExcelApplication.Interactive = False
            _ExcelApplication.Interactive = True
        Catch
            Return True
        End Try
        Return False
    End Function

End Module

Friend Class OpenWorkbookData 'All workbooks that will be opened will get an instance of this class
    Friend Workbook As Microsoft.Office.Interop.Excel.Workbook
    Friend Reader As DatabaseReader
    Friend ReportDate As Date = Nothing
    Friend DatabaseReaderType As String

    Sub New(Wb As Microsoft.Office.Interop.Excel.Workbook)
        Workbook = Wb
    End Sub

End Class