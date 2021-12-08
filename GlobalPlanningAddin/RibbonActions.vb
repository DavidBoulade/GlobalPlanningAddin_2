Option Explicit On
Option Strict On

Imports ExcelDna.Integration
Imports ExcelDna.Integration.CustomUI
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Net
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Xml

<Assembly: System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0060:Supprimer le paramètre inutilisé", Justification:="<En attente>", Scope:="module")>

'**************************************************************************************************************************************************
'*** Use this in case you need custom worksheet functions
'**************************************************************************************************************************************************
'Public Class MyFunctions
'    Implements IExcelAddIn
'    Public Sub AutoOpen() Implements IExcelAddIn.AutoOpen

'    End Sub

'    Public Sub AutoClose() Implements IExcelAddIn.AutoClose

'    End Sub

'    <ExcelFunction(Description:="Test .NET function", Category:="Useful functions")>
'    Public Shared Function HelloDNA(name As String) As String
'        Return "Hello " & name
'    End Function

'    '<ExcelCommand(MenuName:="Test", MenuText:="Set with C API")>
'    'Sub SetValueAPI() 'Application.Run "SetValueAPI"
'    '    Dim target As ExcelReference = New ExcelReference(2, 2)
'    '    target.SetValue("The quick brown fox ...")
'    'End Sub

'End Class



<ComVisible(True)> Public Class RibbonActions
    Inherits ExcelRibbon

    '**************************************************************************************************************************************************
    '*** Entry point of this program - Callback after Excel loaded the Ribbon
    '**************************************************************************************************************************************************
    Dim PluginUpdateThread As Thread 'This thread is used to check for plugin updates in the background
    ''' <summary>callback after Excel loaded the Ribbon</summary>
    Public Sub RibbonLoaded(theRibbon As CustomUI.IRibbonUI)
        Globals.ThisRibbon = theRibbon
        Globals.CurRibbonActions = Me
        Globals.WindowsHandle = ExcelDnaUtil.WindowHandle
        Globals.ExcelApplication = CType(ExcelDnaUtil.Application, Microsoft.Office.Interop.Excel.Application)
        Globals.ExcelApplication_UserLibraryPath = Globals.ExcelApplication.UserLibraryPath 'Save the local user librairy path where to store all files for this plugin
        System.Windows.Forms.Application.EnableVisualStyles() 'enables visual styles (colors, fonts, and other visual elements from the operating system theme) 

        Globals.PluginInstallMgr = New PluginInstallManager()
        PluginUpdateThread = New Thread(Sub() Globals.PluginInstallMgr.CheckUpdates(AddressOf CheckUpdatesDone))
        PluginUpdateThread.Start()

        AddHandler Globals.ExcelApplication.WorkbookActivate, AddressOf MyWbChange
        AddHandler Globals.ExcelApplication.SheetActivate, AddressOf MyWsChange
        AddHandler Globals.ExcelApplication.WorkbookBeforeClose, AddressOf MyWbClose

        LoadListOfTemplates()
    End Sub

    'Public Overrides Sub OnBeginShutdown(ByRef custom As Array)
    '    MsgBox("Excel is shutting down")
    '    MyBase.OnBeginShutdown(custom)
    'End Sub

    '**************************************************************************************************************************************************
    '*** This is called when the check for updates is finished
    '**************************************************************************************************************************************************
    Friend Sub CheckUpdatesDone()
        If Globals.PluginInstallMgr.CurrentPluginVersion.CompareTo(Globals.PluginInstallMgr.LatestPluginVersion) < 0 Then
            Setup_NewVersionAvailableButton()
        ElseIf Globals.PluginInstallMgr.CurrentPluginVersion.CompareTo(Globals.PluginInstallMgr.LatestPluginVersion) > 0 Then
            Setup_BetaVersionButton()
        End If
        PluginUpdateThread.Join()
    End Sub


    '**************************************************************************************************************************************************
    '*** This is called when Excel closes a workbook
    '**************************************************************************************************************************************************
    ''' <summary>callback before Excel closes a workbook</summary>
    Private Sub MyWbClose(Wb As Workbook, ByRef Cancel As Boolean)
        Globals.WorkbookClosed(Wb)
    End Sub


    '**************************************************************************************************************************************************
    '*** Callback when user activated a new workbook
    '**************************************************************************************************************************************************
    ''' <summary>callback when workbook is activated</summary>
    Private Sub MyWbChange(Wb As Microsoft.Office.Interop.Excel.Workbook)

        Globals.WorkbookActivated(Wb)

        SetReportDateBtnLabel(Globals.ReportDate)
        If Not (Globals.ThisWorkbook Is Nothing) Then
            SetReportCreationGroupVisibleState(True)
            Globals.ThisRibbon.ActivateTab("GlobalPlanning") 'or use ActivateTabMso("TabAddIns")
        Else
            SetReportCreationGroupVisibleState(False)
            SetReportActionsGroupVisiblility(False)
            SetProjectionDetailsGroupVisiblility(False)
        End If

        MyWsChange(Wb.ActiveSheet)

    End Sub

    '**************************************************************************************************************************************************
    '*** Callback when user activated a new worksheet
    '**************************************************************************************************************************************************
    ''' <summary>callback when new worksheet is activated</summary>
    Private Sub MyWsChange(Ws As Object) 'Microsoft.Office.Interop.Excel.Worksheet
        Dim ActivetedWorksheet As Microsoft.Office.Interop.Excel.Worksheet = CType(Ws, Microsoft.Office.Interop.Excel.Worksheet)

        Select Case Globals.GetCustomWorksheetProperty(ActivetedWorksheet, "CustomSheetType")
            Case "SKUAlertsConfig", "GRUTConfig"
                SetReportActionsGroupVisiblility(False)
                SetProjectionDetailsGroupVisiblility(False)
            Case "SKUAlertsReport", "GRUTReport"
                SetReportActionsGroupVisiblility(True)
                SetProjectionDetailsGroupVisiblility(False)
            Case "SKUAlertsDetails", "GRUTDetails"
                SetReportActionsGroupVisiblility(False)
                SetProjectionDetailsGroupVisiblility(True)
            Case Else
                SetReportActionsGroupVisiblility(False)
                SetProjectionDetailsGroupVisiblility(False)
        End Select
    End Sub

    '**************************************************************************************************************************************************
    '*** Report Date Button
    '**************************************************************************************************************************************************
    Private _ReportDateBtnLabel As String = "Report Date"
    Private Sub SetReportDateBtnLabel(ReportDate As Date)

        If ReportDate = Nothing Then
            _ReportDateBtnLabel = "Report Date"
        Else
            _ReportDateBtnLabel = "Report Date : " & Format(ReportDate, "yyyy-MM-dd")
        End If
        Globals.ThisRibbon.InvalidateControl("Btn_ReportDate")
    End Sub
    Public Function GetReportDateBtnLabel(control As CustomUI.IRibbonControl) As String
        Return _ReportDateBtnLabel
    End Function


    '**************************************************************************************************************************************************
    '*** Plugin Infos Button
    '**************************************************************************************************************************************************
    Private _PluginInfosBtnLabel As String = "Version"


    Public Function GetPluginInfosBtnLabel(control As CustomUI.IRibbonControl) As String
        If Globals.PluginEnabled = False Then
            Return "Please restart Excel"
        Else
            Return _PluginInfosBtnLabel
        End If

    End Function
    Private Sub Setup_NewVersionAvailableButton()
        _PluginInfosBtnLabel = "New version available"
        _PluginInfosBtnImage = My.Resources.Icon32_Info_red
        Globals.ThisRibbon.InvalidateControl("Btn_CheckVersion")
    End Sub

    Private Sub Setup_BetaVersionButton()
        _PluginInfosBtnLabel = "This is a beta version"
        _PluginInfosBtnImage = My.Resources.Icon32_Info_red
        Globals.ThisRibbon.InvalidateControl("Btn_CheckVersion")
    End Sub

    Public Sub Btn_CheckVersion_Click(control As CustomUI.IRibbonControl)

        Dim Form_Info As New Form_PluginInfos
        Form_Info.Show()

    End Sub


    '**************************************************************************************************************************************************
    '*** Report Creation Group visibility
    '**************************************************************************************************************************************************
    Private _ReportCreationGroupVisible As Boolean = False
    Private Sub SetReportCreationGroupVisibleState(NewVisibleSw As Boolean)
        If _ReportCreationGroupVisible <> NewVisibleSw Then
            _ReportCreationGroupVisible = NewVisibleSw
            Globals.ThisRibbon.InvalidateControl("Btn_ReportDate")
            Globals.ThisRibbon.InvalidateControl("Btn_CreateReport")
        End If
    End Sub

    '**************************************************************************************************************************************************
    '*** Report actions group visibility
    '**************************************************************************************************************************************************
    Private _ReportActionsGroupVisible As Boolean = False
    Private Sub SetReportActionsGroupVisiblility(Visible As Boolean)
        If _ReportActionsGroupVisible <> Visible Then
            _ReportActionsGroupVisible = Visible
            Globals.ThisRibbon.InvalidateControl("ReportActions_RibbonGroup")
        End If
    End Sub

    '**************************************************************************************************************************************************
    '*** Detailled view group visibility
    '**************************************************************************************************************************************************
    Private _ProjectionDetailsGroupVisible As Boolean = False
    Private Sub SetProjectionDetailsGroupVisiblility(Visible As Boolean)
        If _ProjectionDetailsGroupVisible <> Visible Then
            _ProjectionDetailsGroupVisible = Visible
            Globals.ThisRibbon.InvalidateControl("ProjectionDetails_RibbonGroup")
        End If
    End Sub

    '**************************************************************************************************************************************************
    '*** Ribbon GetImage callback
    '**************************************************************************************************************************************************
    Private _ReportTemplatesBtnImage As System.Drawing.Bitmap = My.Resources.Icon32_DownloadTemplate
    Private _PluginInfosBtnImage As System.Drawing.Bitmap = My.Resources.Icon32_Info
    Private _TemplateImage As System.Drawing.Bitmap = Nothing

    Public Function GetBtnImage(IRibbonControl As CustomUI.IRibbonControl) As System.Drawing.Bitmap
        Select Case IRibbonControl.Id
            Case "Btn_CheckVersion" : Return _PluginInfosBtnImage
            Case "ReportTemplates_DynamicMenu" : Return _ReportTemplatesBtnImage
            Case "Btn_ReportDate", "Btn_DetailedView_Date" : Return My.Resources.Icon32_Calendar
            Case "Btn_CreateReport" : Return My.Resources.Icon32_NewReport
            Case "Btn_ReportActions_Save" : Return My.Resources.Icon32_Save
            Case "Btn_ReportActions_Details", "Gallery_DetailedView_Infos", "Btn_ContextMenuCell_Details", "Btn_ContextMenuCellLayout_Details" : Return My.Resources.Icon32_MagnifyingGlass
            Case "TemplateImage" : Return _TemplateImage
            Case Else
                Return Nothing
        End Select
    End Function

    '**************************************************************************************************************************************************
    '*** Global callback setting visibility of the different ribbon elements
    '**************************************************************************************************************************************************
    Public Function GetVisible(control As CustomUI.IRibbonControl) As Boolean

        If Globals.PluginEnabled = False Then

            Select Case control.Id
                Case "AddinVersion_RibbonGroup" : Return True
                Case Else : Return False
            End Select

        Else

            Select Case control.Id
                Case "ReportCreation_RibbonGroup" : Return _ReportCreationGroupVisible
                Case "ReportActions_RibbonGroup" : Return _ReportActionsGroupVisible
                Case "ProjectionDetails_RibbonGroup" : Return _ProjectionDetailsGroupVisible
                Case "Btn_ContextMenuCell_Details", "Btn_ContextMenuCellLayout_Details" : Return _ReportActionsGroupVisible
                Case Else : Return True
            End Select

        End If

    End Function

    ''**************************************************************************************************************************************************
    ''*** Global callback setting Enabled state of the different ribbon elements NOT USED ANYMORE FOR NOW
    ''**************************************************************************************************************************************************
    'Public Function GetEnabled(control As CustomUI.IRibbonControl) As Boolean
    '    Select Case control.Id
    '        Case "Btn_ReportDate", "Btn_CreateReport"
    '            Return _ReportCreationButtonsEnabled
    '        Case Else
    '            Return True
    '    End Select
    'End Function

    '**************************************************************************************************************************************************
    '*** Create report
    '**************************************************************************************************************************************************
    Public Sub Btn_CreateReport_Click(ByVal control As IRibbonControl)
        Dim NewReportDate As Date = Globals.ReportDate
        If NewReportDate = Nothing Then NewReportDate = Today
        CreateReport(NewReportDate)
    End Sub

    Private Sub CreateReport(NewReportDate As Date)

        'Check if there is already a report in place
        If Not (Globals.Reader Is Nothing) Then
            If MsgBox("Warning: This will erase the existing report and any unsaved change you may have done. Continue?", MsgBoxStyle.OkCancel, "Global planning Addin") = MsgBoxResult.Cancel Then Exit Sub
            Globals.Reader.Dispose()
        End If

        Globals.ReportDate = NewReportDate
        SetReportDateBtnLabel(Globals.ReportDate) 'Make sure the report date displayed is correct

        Globals.ReportSheet.Visible = XlSheetVisibility.xlSheetVisible

        'Create a new DatabaseReader object
        Globals.Reader = New DatabaseReader(Globals.ReportDate, Globals.DatabaseReaderType)
        If Globals.Reader.CreateReport() = False Then
            Globals.ReportSheet.Visible = XlSheetVisibility.xlSheetHidden
            Globals.DetailsSheet.Visible = XlSheetVisibility.xlSheetHidden
            MsgBox(Globals.Reader.LastMessage, MsgBoxStyle.Critical, "Global planning Addin")
            Globals.Reader.Dispose() 'if it didn't work (no data returned), dispose the object
        End If
    End Sub

    '**************************************************************************************************************************************************
    '*** Save Changes
    '**************************************************************************************************************************************************
    Public Sub Btn_Save_Click(ByVal control As IRibbonControl)

        If Not (Globals.Reader Is Nothing) Then
            If Globals.Reader.SendUserModificationsToDB() = False Then
                MsgBox(Globals.Reader.LastMessage, MsgBoxStyle.Critical, "Global planning Addin")
            End If
        Else
            MsgBox("Error: The report has not yet been generated", MsgBoxStyle.Critical, "Global planning Addin")
        End If

    End Sub

    '**************************************************************************************************************************************************
    '*** Summary report - report date selection
    '**************************************************************************************************************************************************
    Public Sub Btn_ReportDate_Click(ByVal control As IRibbonControl)
        Dim DefaultSelectedDate As Date = Globals.ReportDate
        If DefaultSelectedDate = Nothing Then DefaultSelectedDate = Today
        Dim ReportDateWindow As New Form_ReportDate(DefaultSelectedDate) 'Create a new date selection form
        ReportDateWindow.ShowDialog() 'show it in modal mode
        If ReportDateWindow.WasCancelled = False Then
            CreateReport(ReportDateWindow.SelectedDate)
        End If
        ReportDateWindow.Dispose()
    End Sub

    '**************************************************************************************************************************************************
    '*** Detailled view
    '**************************************************************************************************************************************************
    Dim _KeyValuesOfCurrentRow() As String
    Public Sub Btn_Details_Click(ByVal control As IRibbonControl)

        If Globals.Reader Is Nothing Then Exit Sub
        If Globals.ThisWorkbook.Application.ActiveCell.Row <= DatabaseReader.REPORT_FIRSTROW - 1 Then Exit Sub

        _KeyValuesOfCurrentRow = Globals.Reader.GetKeyValues_For_SummaryReportRow(Globals.ThisWorkbook.Application.ActiveCell.Row, CType(Globals.ThisWorkbook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet))

        _AvailableMD04Dates = Nothing 'Reset the list of dates on which we have available MD04 snapshots
        _CurrentMD04Date = Globals.ReportDate
        Globals.DetailsSheet.Visible = XlSheetVisibility.xlSheetVisible
        DisplayDetails(_KeyValuesOfCurrentRow, Globals.ReportDate)

    End Sub

    Private Sub DisplayDetails(KeyValues() As String, SnapshotDate As Date)

        If Globals.Reader Is Nothing Then Exit Sub
        For Each KeyValue As String In KeyValues
            If KeyValue = "" Then Exit Sub
        Next
        If Globals.Reader.ReadDetailedProjectionData(SnapshotDate, KeyValues) = True Then
            SetDetailedView_Date(SnapshotDate)
            Globals.DetailsSheet.Activate()
            If Globals.Reader.ReadSingleSummaryTableRow(SnapshotDate, KeyValues) = False Then 'TODO move this to DBReader
                MsgBox("Error: Unable to read detailled data", MsgBoxStyle.Critical, "Global planning Addin")
            Else
                Globals.ThisRibbon.InvalidateControl("Lbl_DetailedView_Header")
                Globals.ThisRibbon.InvalidateControl("Btn_DetailedView_Info")
                Globals.ThisRibbon.InvalidateControl("Btn_DetailedView_Date")
            End If
        Else
            MsgBox("Error: Unable to read detailled data", MsgBoxStyle.Critical, "Global planning Addin")
        End If
    End Sub

    Private Sub SetDetailedView_Date(NewDate As Date)
        _DetailedView_Date = NewDate
        Globals.ThisRibbon.InvalidateControl("Btn_DetailedView_Date")
    End Sub
    Private _DetailedView_Date As Date = Today
    Public Function Get_DetailedView_DateBtn_Label(control As CustomUI.IRibbonControl) As String
        If Globals.Reader Is Nothing Then Return "(Not defined)"
        Return Format(_DetailedView_Date, "yyyy-MM-dd")
    End Function


    '**************************************************************************************************************************************************
    '*** Detailled view - Data Date
    '**************************************************************************************************************************************************
    Private _AvailableMD04Dates() As Date = Nothing
    Private _CurrentMD04Date As Date
    Public Sub Btn_DetailedView_Date_Click(control As CustomUI.IRibbonControl)

        If Globals.Reader Is Nothing Then Exit Sub

        If _KeyValuesOfCurrentRow Is Nothing Then Exit Sub

        If _AvailableMD04Dates Is Nothing Then
            If Globals.Reader.Read_DetailsTable_Available_Dates(_KeyValuesOfCurrentRow) = False Then
                MsgBox("Error: Unable to read Availalble dates", MsgBoxStyle.Critical, "Global planning Addin")
                Exit Sub
            Else
                ReDim _AvailableMD04Dates(0 To Globals.Reader.DetailsTable_AvailableDates_DataSet.Tables(0).Rows.Count - 1)
                For i As Integer = 0 To Globals.Reader.DetailsTable_AvailableDates_DataSet.Tables(0).Rows.Count - 1
                    _AvailableMD04Dates(i) = CDate(Globals.Reader.DetailsTable_AvailableDates_DataSet.Tables(0).Rows(i).Item("ReportDate"))
                Next
            End If
        End If

        Dim ReportDateWindow As New Form_ReportDate(_CurrentMD04Date, _AvailableMD04Dates) 'Create a new date selection form
        ReportDateWindow.ShowDialog() 'show it in modal mode
        If ReportDateWindow.WasCancelled = False Then

            If _AvailableMD04Dates.Contains(ReportDateWindow.SelectedDate) = False Then
                MsgBox("There is no data for the date you selected" & vbCrLf & "Please select a date in bold.", MsgBoxStyle.Exclamation, "Global planning Addin")
            Else
                _CurrentMD04Date = ReportDateWindow.SelectedDate
                DisplayDetails(_KeyValuesOfCurrentRow, ReportDateWindow.SelectedDate)
            End If


        End If
        ReportDateWindow.Dispose()

    End Sub

    '**************************************************************************************************************************************************
    '*** Detailled view - Info Drop Down
    '**************************************************************************************************************************************************
    Dim DetailledView_InfosDropDownForm As Form_AutoSize_DataGrid
    Public Sub Btn_DetailedView_Info_Click(control As CustomUI.IRibbonControl)

        If Globals.Reader Is Nothing Then Exit Sub
        If _KeyValuesOfCurrentRow Is Nothing Then Exit Sub

        If Not (DetailledView_InfosDropDownForm Is Nothing) Then DetailledView_InfosDropDownForm.Close()

        Dim DataRows As New List(Of Object())

        DataRows.AddRange(Globals.Reader.Get_ListOf_DetailedView_Info_Columns)

        DetailledView_InfosDropDownForm = New Form_AutoSize_DataGrid(DataRows)

        DetailledView_InfosDropDownForm.Show()

    End Sub

    Public Function Get_DetailedView_Info_Image(IRibbonControl As CustomUI.IRibbonControl) As System.Drawing.Bitmap
        Return My.Resources.Icon32_MagnifyingGlass
    End Function

    Public Function Get_DetailedView_Header_Label(control As CustomUI.IRibbonControl) As String
        If Globals.Reader Is Nothing Then Return "(Not Defined)"
        Return Globals.Reader.Get_DetailledView_HeaderText
    End Function

    '**************************************************************************************************************************************************
    '*** Change log form
    '**************************************************************************************************************************************************
    Private ChangeLogForm As Form_AutoSize_DataGrid
    Public Sub Btn_ChangeLog_Click(ByVal control As IRibbonControl)
        If Globals.Reader Is Nothing Then Exit Sub

        If Not (ChangeLogForm Is Nothing) Then ChangeLogForm.Close()


        Dim ChangeLogDataset As DataSet = Globals.Reader.Get_ChangeLog(Globals.ThisWorkbook.Application.ActiveCell.Row, Globals.ThisWorkbook.Application.ActiveCell.Column, CType(Globals.ThisWorkbook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet))

        If Not (ChangeLogDataset Is Nothing) AndAlso ChangeLogDataset.Tables(0).Rows.Count > 0 Then
            ChangeLogForm = New Form_AutoSize_DataGrid(ChangeLogDataset)
            ChangeLogForm.Show()
        Else
            MsgBox("No change made to this cell", MsgBoxStyle.Information, "Global planning Addin")
        End If

    End Sub
    Public Function Get_ChangeLogBtn_Visible(control As CustomUI.IRibbonControl) As Boolean

        If Globals.ThisWorkbook.Application.ActiveSheet IsNot Globals.ReportSheet Then Return False 'if we are not on the right sheet, don't display the button
        If Globals.Reader.Is_SummaryWorksheetColumn_Modifiable(Globals.ThisWorkbook.Application.ActiveCell.Column) Then 'check if the column is modifiable
            Return True
        Else
            Return False
        End If

    End Function

    '**************************************************************************************************************************************************
    '*** Management of the file templates
    '**************************************************************************************************************************************************

    Private _DownloadTemplatesConfigFile_Thread As Thread
    Private ReadOnly _templates As New List(Of FileTemplate)
    Private ReadOnly _templates_Lock As New Object 'This lock is used to protect against concurrent accesses of _template
    Const DEFAULT_MENU_XML As String = "<menu xmlns='http://schemas.microsoft.com/office/2006/01/customui' itemSize='normal'><button id='BtnLoading' label='Loading...' imageMso='RefreshAll'/></menu >"
    Private _templatesMenuXml As String = DEFAULT_MENU_XML

    Public Function GetReportTemplatesMenuContent(ByVal control As CustomUI.IRibbonControl) As String
        'When Excel asks, returns the current Menu XML
        Return _templatesMenuXml
    End Function

    Public Sub RefreshListOfTemplates_MenuBtnClick(ByVal control As CustomUI.IRibbonControl) 'This is called when the user clicks the Refresh button on the templates Menu
        TriggerRefreshListOfTemplates()
    End Sub

    Public Sub LoadListOfTemplates() 'Load the list of templates (first from the local copy if it exists, and then from the central repository)

        'Start with the default "Loading..." menu
        _templatesMenuXml = DEFAULT_MENU_XML

        'If the local copy of the templates config file exists, loads it
        Dim InstallPath As String = Globals.ExcelApplication_UserLibraryPath & Globals.PluginXllInstallSubFolder
        If System.IO.File.Exists(InstallPath & Globals.TemplatesMenuXmlFileName) = True Then
            LoadTemplatesConfigFile(InstallPath & Globals.TemplatesMenuXmlFileName)
        End If

        'and trigger the refresh of the config file
        TriggerRefreshListOfTemplates()
    End Sub
    Public Sub TriggerRefreshListOfTemplates() 'Refresh List of Templates from the central repository

        If _DownloadTemplatesConfigFile_Thread IsNot Nothing AndAlso
            _DownloadTemplatesConfigFile_Thread.ThreadState = ThreadState.Running Then Exit Sub 'if the thread is already running, exit

        _DownloadTemplatesConfigFile_Thread = New Thread(Sub() DownloadAndLoad_Latest_TemplatesConfigFile()) '.NET Garbage collector shoud free up any previous thread object when we lose the reference
        _DownloadTemplatesConfigFile_Thread.Start() 'Start the download in background

    End Sub

    Public Sub DownloadAndLoad_Latest_TemplatesConfigFile() 'Download the Templates config file from the central repository, and trigger its load
        'As this is run in a background thread, we must make sure no COM call is made to Excel
        Dim webClient = New WebClient
        Dim TemplatesConfigFileBuffer As Byte()

        Dim InstallPath As String = Globals.ExcelApplication_UserLibraryPath & Globals.PluginXllInstallSubFolder

        Try
            If System.IO.Directory.Exists(InstallPath) = False Then System.IO.Directory.CreateDirectory(InstallPath) 'Create the folder if doesn't exist
            TemplatesConfigFileBuffer = webClient.DownloadData(Globals.VersionCheckFolder & Globals.TemplatesMenuXmlFileName)
            Dim fs As FileStream = New FileStream(InstallPath & Globals.TemplatesMenuXmlFileName, FileMode.Create)
            fs.Write(TemplatesConfigFileBuffer, 0, TemplatesConfigFileBuffer.Length)
            fs.Close()
            fs.Dispose()
            'Templates config file downloaded correctly, now loads it
            LoadTemplatesConfigFile(InstallPath & Globals.TemplatesMenuXmlFileName)
        Catch ex As Exception
            'Error while downloading Templates config file
        End Try

    End Sub

    Public Sub LoadTemplatesConfigFile(FileFullPath As String) 'Load a given templates config file

        Try

            Dim XmlDoc As XmlDocument = New XmlDocument()
            XmlDoc.Load(FileFullPath)

            Dim FoldersCounter As Integer = 0
            Dim NewMenuXml As String = "<menu xmlns='http://schemas.microsoft.com/office/2006/01/customui' itemSize='normal'>"

            SyncLock _templates_Lock 'Make sure no one tries to read the templates while we are making modifications to the list here
                _templates.Clear() 'Clear the list of templates
                For Each NodeElt As XmlNode In XmlDoc.DocumentElement.ChildNodes
                    Select Case NodeElt.Name

                        Case "template"
                            Dim CurTemplate As New FileTemplate With {.ID = NodeElt.Attributes("id").Value,
                                                              .TemplateDescription = NodeElt.Attributes("description").Value,
                                                              .TemplateFileName = NodeElt.Attributes("filename").Value,
                                                              .TemplateVersion = New Version(NodeElt.Attributes("version").Value),
                                                              .MinPluginVersion = New Version(NodeElt.Attributes("minpluginversion").Value)}
                            _templates.Add(CurTemplate)
                            NewMenuXml &= "<button id='Btn_" & CurTemplate.ID & "' tag='" & CurTemplate.ID & "' label='" & CurTemplate.TemplateDescription & " (v" & CurTemplate.GetVersionText & ")" & "' imageMso='FileSaveAsExcelXlsx' onAction='ReportTemplate_Open_Click' />"

                        Case "folder"

                            FoldersCounter += 1
                            NewMenuXml &= "<menu id='Menu" & FoldersCounter & "' label='" & NodeElt.Attributes("description").Value & "' imageMso='Folder' >"

                            For Each SubNodeElt As XmlNode In NodeElt.ChildNodes

                                If SubNodeElt.Name = "template" Then 'this is not an elegant programming solution, ideally it should be recursive and code not repeated
                                    Dim CurTemplate As New FileTemplate With {.ID = SubNodeElt.Attributes("id").Value,
                                                          .TemplateDescription = SubNodeElt.Attributes("description").Value,
                                                          .TemplateFileName = SubNodeElt.Attributes("filename").Value,
                                                          .TemplateVersion = New Version(SubNodeElt.Attributes("version").Value),
                                                          .MinPluginVersion = New Version(SubNodeElt.Attributes("minpluginversion").Value)}
                                    _templates.Add(CurTemplate)
                                    NewMenuXml &= "<button id='Btn_" & CurTemplate.ID & "' tag='" & CurTemplate.ID & "' label='" & CurTemplate.TemplateDescription & " (v" & CurTemplate.GetVersionText & ")" & "' imageMso='FileSaveAsExcelXlsx' onAction='ReportTemplate_Open_Click' />"
                                End If

                            Next

                            NewMenuXml &= "</menu >"
                    End Select

                Next
            End SyncLock

            NewMenuXml &= "<button id='ButtonRefresh' label='Refresh list' imageMso='RefreshAll' onAction='RefreshListOfTemplates_MenuBtnClick' />"
            NewMenuXml &= "</menu>"
            _templatesMenuXml = NewMenuXml

        Catch ex As Exception
            Dim NewMenuXml = "<menu xmlns='http://schemas.microsoft.com/office/2006/01/customui' itemSize='normal'>"
            NewMenuXml &= "<button id='ButtonError' label='Unable to read Templates list' imageMso='HighImportance' onAction='RefreshListOfTemplates_MenuBtnClick' />"
            NewMenuXml &= "<button id='ButtonRefresh' label='Refresh list' imageMso='RefreshAll' onAction='RefreshListOfTemplates_MenuBtnClick' />"
            NewMenuXml &= "</menu>"
            _templatesMenuXml = NewMenuXml
        End Try

    End Sub


    Public Sub ReportTemplate_Open_Click(ByVal control As CustomUI.IRibbonControl)

        Dim selectedTemplateID As String = control.Tag
        Dim selectedTemplate As FileTemplate = _templates.Find(Function(c) c.ID = selectedTemplateID)

        If selectedTemplate.MinPluginVersion.CompareTo(Globals.PluginVersion) < 0 Then
            'Removing this message box as it is painful to click ok each time
            ' If MsgBox("Do you want to download a new template (" & selectedTemplate.TemplateDescription & ") ?", MsgBoxStyle.OkCancel, "Global planning Addin") = MsgBoxResult.Cancel Then Exit Sub
        Else
            If MsgBox("This template requires a plugin version " & selectedTemplate.MinPluginVersion.ToString & vbCrLf &
                      "Your current version is " & Globals.PluginVersion.ToString & vbCrLf &
                      "It is highly recommended to upgrade your addin before using this template." & vbCrLf &
                      "Do you want to continue anyway?", MsgBoxStyle.YesNo, "Global planning Addin") = MsgBoxResult.No Then Exit Sub
        End If

        Dim ProgressWindow As Form_Progress
        ProgressWindow = New Form_Progress("Downloading Template")
        ProgressWindow.Show()

        'Globals.ThisWorkbook.Application.ScreenUpdating = False

        ProgressWindow.SetProgress(10, "Downloading Template")


        Dim webClient = New WebClient
        Dim TemplateFileBuffer As Byte()

        Dim InstallPath As String = Globals.ExcelApplication_UserLibraryPath & Globals.PluginXllInstallSubFolder

        Try
            If System.IO.Directory.Exists(InstallPath) = False Then System.IO.Directory.CreateDirectory(InstallPath) 'Create the folder if doesn't exist
            TemplateFileBuffer = webClient.DownloadData(Globals.VersionCheckFolder & TemplatesSubFolder & selectedTemplate.TemplateFileName)
            ProgressWindow.SetProgress(50, "Saving Template")
            Dim fs As FileStream = New FileStream(InstallPath & selectedTemplate.TemplateFileName, FileMode.Create)
            fs.Write(TemplateFileBuffer, 0, TemplateFileBuffer.Length)
            fs.Close()
            ProgressWindow.SetProgress(90, "Opening Template")
            Globals.ExcelApplication.Workbooks.Open(InstallPath & selectedTemplate.TemplateFileName)
            ProgressWindow.Close()
            'MsgBox("Template downloaded correctly. You can now save this file with your own selection criteria to your preferred location on your computer.", MsgBoxStyle.Information, "Done")
        Catch ex As Exception
            ProgressWindow.Close()
            MsgBox("Unable to load the template : " & ex.Message, MsgBoxStyle.Critical, "Global planning Addin")
        End Try



    End Sub


    Private _ReportTemplatesBtnLabel As String = "Download template"
    Public Function GetReportTemplatesBtnLabel(control As CustomUI.IRibbonControl) As String
        Return _ReportTemplatesBtnLabel
    End Function

    Private _UsedWarned As Boolean = False
    Public Sub TemplateLoaded(TemplateID As String, TemplateVersion As String) 'This is called after a compatible Workbook is activated

        SyncLock _templates_Lock 'Make sure this is not executed while the Template config file is being loaded

            Dim CurTemplate As FileTemplate

            If _templates.Exists(Function(c) c.ID = TemplateID) Then 'check if the template is in the list of defined templated
                CurTemplate = _templates.Find(Function(c) c.ID = TemplateID)

                If TemplateVersion = "" Then
                    'This must be an old template with undefined version
                    NewReportTemplateVersionAvailable()
                    _CurTemplateVersion = "(Not defined)"
                Else
                    'Version is known
                    Dim CurrentFileVersion As New Version(TemplateVersion)
                    If CurrentFileVersion.CompareTo(CurTemplate.TemplateVersion) < 0 Then
                        'If the version is lower than the latest template, warn user
                        NewReportTemplateVersionAvailable()
                    Else
                        ReportTemplateIsUpToDate()
                    End If
                    _CurTemplateVersion = TemplateVersion
                End If
                _CurTemplateDescr = CurTemplate.TemplateDescription
                If _CurTemplateDescr = "" Then _CurTemplateDescr = "(Not defined)" 'This shound NOT happen!!

            Else 'the template that the user opened is not in the list!?

                _CurTemplateDescr = "(Not defined)"
                _CurTemplateVersion = TemplateVersion
                If _CurTemplateVersion = "" Then _CurTemplateVersion = "(Not defined)"
                If _templates.Count = 0 Then
                    'We have probably not yet loaded a templates config file!
                    'Do nothing... all will be ok when the user will reactivate this workbook after the templates config file is loaded
                Else
                    'If we ended up here, it means that a template is missing in the template.xml, or the TemplateID property is not set for this workbook
                    NewReportTemplateVersionAvailable()
                    If _UsedWarned = False Then
                        _UsedWarned = True 'We will only raise this message box one time
                        MsgBox("You opened an old template version." & vbCrLf & "Please download the latest version from the templates download menu.", MsgBoxStyle.Exclamation, "Global planning Addin")
                    End If
                End If

            End If

        End SyncLock

        Select Case TemplateID
            Case "MROB"
                _TemplateImage = My.Resources.Icon32_MROB
            Case "RTT"
                _TemplateImage = My.Resources.Icon32_RTT
            Case "GRUT_MPS"
                _TemplateImage = My.Resources.Icon32_GRUT
            Case Else
                _TemplateImage = Nothing
        End Select

    End Sub

    Public Sub NonCompatibleFileLoaded()
        ReportTemplateIsUpToDate()
    End Sub
    Private Sub ReportTemplateIsUpToDate()
        _ReportTemplatesBtnImage = My.Resources.Icon32_DownloadTemplate
        _ReportTemplatesBtnLabel = "Download template"
        Globals.ThisRibbon.InvalidateControl("ReportTemplates_DynamicMenu")
    End Sub
    Private Sub NewReportTemplateVersionAvailable()
        _ReportTemplatesBtnImage = My.Resources.Icon32_DownloadTemplate_red
        _ReportTemplatesBtnLabel = "New version available"
        Globals.ThisRibbon.InvalidateControl("ReportTemplates_DynamicMenu")
    End Sub

    Dim _CurTemplateDescr As String = ""
    Dim _CurTemplateVersion As String = ""

    Public Function GetTemplateName(control As CustomUI.IRibbonControl) As String
        Return _CurTemplateDescr
    End Function

    Public Function GetTemplateVersion(control As CustomUI.IRibbonControl) As String
        Return "Version " & _CurTemplateVersion
    End Function

    Public Sub Btn_CheckTemplateVersion_Click(control As CustomUI.IRibbonControl)
        MsgBox("Current template:" & vbCrLf & _CurTemplateDescr & vbCrLf & "Version " & _CurTemplateVersion, MsgBoxStyle.OkOnly, "Global planning Addin")
    End Sub


End Class


'**************************************************************************************************************************************************
'*** This class is used to store properties of available file templates
'**************************************************************************************************************************************************
Friend Class FileTemplate
    Public Property ID As String
    Public Property TemplateDescription As String
    Public Property TemplateFileName As String
    Public Property TemplateVersion As Version
    Public Property MinPluginVersion As Version

    Public Function GetVersionText() As String
        If _TemplateVersion.Revision = Nothing Then 'Display current plugin version
            Return _TemplateVersion.ToString(3)
        Else
            Return _TemplateVersion.ToString()
        End If
    End Function
End Class