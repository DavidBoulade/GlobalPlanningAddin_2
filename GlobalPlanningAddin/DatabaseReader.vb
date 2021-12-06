Option Explicit On
Option Strict On

Imports Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient
Imports System.Threading
Imports System.Threading.Tasks
Friend Class DatabaseReader : Implements IDisposable

    Public Const PARAMS_FIRSTROW As Integer = 2
    Public Const PARAMS_COL_HEADERTEXT As Integer = 1
    Public Const PARAMS_COL_COLUMNNAME As Integer = 2
    Public Const PARAMS_COL_WIDTH As Integer = 3
    Public Const PARAMS_COL_FORMAT As Integer = 4
    Public Const PARAMS_COL_HIDDEN As Integer = 5
    Public Const PARAMS_COL_GROUPPED As Integer = 6
    Public Const PARAMS_COL_COLFILTER As Integer = 7
    Public Const REPORT_FIRSTROW As Integer = 2
    Public Const DETAILS_FIRSTROW As Integer = 2

    Private ReadOnly _DBAdapter As DatabaseAdapterBase

    Friend Property LastMessage As String

    Private _Param_ColHeader_Rng As ExcelRangeArray
    Private _Param_ColName_Rng As ExcelRangeArray
    Private _Param_ColWidth_Rng As ExcelRangeArray
    'Private _Param_ColFormat_Rng As ExcelRangeArray
    Private _Param_ColHidden_Rng As ExcelRangeArray
    Private _Param_ColGroupped_Rng As ExcelRangeArray
    Private _Param_ColFilter_Rng As ExcelRangeArray

    Private ReadOnly _Report_Param_Row() As Integer 'Row number of each column in the params worksheet
    Private ReadOnly _Report_ColNumber() As Integer 'Column number in the summary report
    Private ReadOnly _ReportDate As Date
    Private _ReportAlreadyCreated As Boolean = False


    Private _LastDBUpdateDateTime As DateTime = Date.MinValue

    Private _ReportNbColumns As Integer

    Private _ProgressWindow As Form_Progress


    Sub New(ReportDate As Date, DatabaseAdapterType As String)
        _ReportDate = ReportDate

        Select Case DatabaseAdapterType
            Case "SKUAlertsUI"
                _DBAdapter = New SKUAlertsDatabaseAdapter()
            Case "GRUT_UI"
                _DBAdapter = New GRUTDatabaseAdapter()
        End Select


        ReDim _Report_Param_Row(0 To _DBAdapter.Get_SummaryTable_Columns.Count - 1) 'Row number of each column in the params worksheet
        ReDim _Report_ColNumber(0 To _DBAdapter.Get_SummaryTable_Columns.Count - 1)  'Column number in the summary report

    End Sub

    Private Function ReadParams() As Boolean

        Dim ParamRow As Integer
        Dim ParamsNbRows As Integer

        ParamsNbRows = ConfigSheet.UsedRange.Rows.End(XlDirection.xlDown).Row

        _Param_ColHeader_Rng = New ExcelRangeArray("Column Header",
                                                   ConfigSheet.Range(
                                                        ConfigSheet.Cells(PARAMS_FIRSTROW, PARAMS_COL_HEADERTEXT),
                                                        ConfigSheet.Cells(ParamsNbRows, PARAMS_COL_HEADERTEXT)))
        _Param_ColName_Rng = New ExcelRangeArray("Column Name",
                                                   ConfigSheet.Range(
                                                        ConfigSheet.Cells(PARAMS_FIRSTROW, PARAMS_COL_COLUMNNAME),
                                                        ConfigSheet.Cells(ParamsNbRows, PARAMS_COL_COLUMNNAME)))
        _Param_ColWidth_Rng = New ExcelRangeArray("Width",
                                                   ConfigSheet.Range(
                                                        ConfigSheet.Cells(PARAMS_FIRSTROW, PARAMS_COL_WIDTH),
                                                        ConfigSheet.Cells(ParamsNbRows, PARAMS_COL_WIDTH)))
        '_Param_ColFormat_Rng = New ExcelRangeArray("Format",
        '                                            ConfigSheet.Range(
        '                                             ConfigSheet.Cells(PARAMS_FIRSTROW, PARAMS_COL_FORMAT),
        '                                             ConfigSheet.Cells(ParamsNbRows, PARAMS_COL_FORMAT)
        '                                                 )) 'copy also the format attributes
        _Param_ColHidden_Rng = New ExcelRangeArray("Hidden",
                                                   ConfigSheet.Range(
                                                        ConfigSheet.Cells(PARAMS_FIRSTROW, PARAMS_COL_HIDDEN),
                                                        ConfigSheet.Cells(ParamsNbRows, PARAMS_COL_HIDDEN)))
        _Param_ColGroupped_Rng = New ExcelRangeArray("Groupped",
                                                   ConfigSheet.Range(
                                                        ConfigSheet.Cells(PARAMS_FIRSTROW, PARAMS_COL_GROUPPED),
                                                        ConfigSheet.Cells(ParamsNbRows, PARAMS_COL_GROUPPED)))
        _Param_ColFilter_Rng = New ExcelRangeArray("Filter",
                                                   ConfigSheet.Range(
                                                        ConfigSheet.Cells(PARAMS_FIRSTROW, PARAMS_COL_COLFILTER),
                                                        ConfigSheet.Cells(ParamsNbRows, PARAMS_COL_COLFILTER)))



        For ColIndex = 0 To _DBAdapter.Get_SummaryTable_Columns.Count - 1
            _Report_Param_Row(ColIndex) = 0 '0 by default
            _Report_ColNumber(ColIndex) = 0
        Next ColIndex

        Dim FoundSw As Boolean
        ParamRow = 1
        _ReportNbColumns = 0
        Do While _Param_ColName_Rng.CellValue_Str(ParamRow) <> ""

            FoundSw = False
            For ColIndex = 0 To _DBAdapter.Get_SummaryTable_Columns.Count - 1
                If _Param_ColName_Rng.CellValue_Str(ParamRow) = _DBAdapter.GetColName(ColIndex) Then
                    FoundSw = True
                    _ReportNbColumns += 1
                    _Report_Param_Row(ColIndex) = ParamRow + PARAMS_FIRSTROW - 1 'Row number in the params worksheet
                    _Report_ColNumber(ColIndex) = ParamRow 'Column number in the summary report
                    Exit For
                End If
            Next ColIndex

            If FoundSw = False Then
                _LastMessage = "'" & _Param_ColName_Rng.CellValue_Str(ParamRow) & "' is not a valid column name."
                Return False
            End If

            ParamRow += 1
            If ParamRow > _Param_ColName_Rng.NbRows Then Exit Do
        Loop

        Return True

    End Function


    Private Sub SendDisplayFieldsToDBAdapter()
        Dim ParamRow As Integer
        Dim ColumnName As String

        'List all the fields we want to display on the report
        _DBAdapter.ClearDisplayFields()
        For ParamRow = 1 To _ReportNbColumns
            ColumnName = _Param_ColName_Rng.CellValue_Str(ParamRow)
            _DBAdapter.AddDisplayField(ColumnName)
        Next ParamRow

    End Sub

    Private Sub SetupQueryFiltersInDBAdapter()
        Dim ParamRow As Integer
        Dim FilterValues() As String
        Dim FilterNbValuesIncl As Integer
        Dim FilterNbValuesExcl As Integer
        Dim FilterValueStrIncl As String
        Dim FilterValueStrExcl As String
        Dim ColumnName As String

        _DBAdapter.ClearQueryFilters()
        ParamRow = 1 'PARAMS_FIRSTROW
        Do While _Param_ColName_Rng.CellValue_Str(ParamRow) <> ""

            If Trim(_Param_ColFilter_Rng.CellValue_Str(ParamRow)) <> "" Then
                'We found a new filter

                FilterValues = Split(Trim(_Param_ColFilter_Rng.CellValue_Str(ParamRow)), ",")

                FilterNbValuesIncl = 0
                FilterNbValuesExcl = 0
                FilterValueStrIncl = ""
                FilterValueStrExcl = ""
                For i = LBound(FilterValues) To UBound(FilterValues)

                    If Left(FilterValues(i), 1) = "!" Then
                        'This value needs to be EXcluded
                        FilterNbValuesExcl += 1
                        If FilterNbValuesExcl > 1 Then FilterValueStrExcl &= ","
                        FilterValueStrExcl = FilterValueStrExcl & "'" & Right(FilterValues(i), Len(FilterValues(i)) - 1) & "'"
                    Else
                        'This value needs to be INcluded
                        FilterNbValuesIncl += 1
                        If FilterNbValuesIncl > 1 Then FilterValueStrIncl &= ","
                        FilterValueStrIncl = FilterValueStrIncl & "'" & FilterValues(i) & "'"
                    End If

                Next i

                Dim FilterText As String = ""
                ColumnName = _Param_ColName_Rng.CellValue_Str(ParamRow)
                If FilterNbValuesIncl > 0 Then FilterText = SKUAlertsDatabaseAdapter.GetColDatabaseNameFromColName(ColumnName) & " IN (" & FilterValueStrIncl & ")"
                If FilterNbValuesIncl > 0 And FilterNbValuesExcl > 0 Then FilterText &= " AND "
                If FilterNbValuesExcl > 0 Then FilterText = FilterText & SKUAlertsDatabaseAdapter.GetColDatabaseNameFromColName(ColumnName) & " NOT IN (" & FilterValueStrExcl & ")"

                _DBAdapter.AddQueryFilter(FilterText)

            End If
            ParamRow += 1
            If ParamRow > _Param_ColName_Rng.NbRows Then Exit Do
        Loop
    End Sub

    Private Sub CleanReportWorksheet()
        'Delete everything from the report worksheet
        If ReportSheet.ProtectContents = True Then ReportSheet.Unprotect()
        ReportSheet.Range(ReportSheet.Columns(1), ReportSheet.Columns(500)).Delete(XlDeleteShiftDirection.xlShiftToLeft)
    End Sub

    Private Sub ApplyReportColumnsSetup(ReportNbRow As Integer)

        ''Header text
        'ConfigSheet.Range(
        '        ConfigSheet.Cells(PARAMS_FIRSTROW, PARAMS_COL_HEADERTEXT),
        '        ConfigSheet.Cells(PARAMS_FIRSTROW + _Param_ColName_Rng.NbRows - 1, PARAMS_COL_FORMAT)).Copy()
        'ReportSheet.Range(
        '        ReportSheet.Cells(REPORT_FIRSTROW - 1, 1),
        '        ReportSheet.Cells(REPORT_FIRSTROW - 1, _Param_ColName_Rng.NbRows)).PasteSpecial(XlPasteType.xlPasteValues,,, True)

        ''Header Format
        'CType(ConfigSheet.Cells(PARAMS_FIRSTROW - 1, PARAMS_COL_HEADERTEXT), Range).Copy()
        'ReportSheet.Range(
        '        ReportSheet.Cells(REPORT_FIRSTROW - 1, 1),
        '        ReportSheet.Cells(REPORT_FIRSTROW - 1, _Param_ColName_Rng.NbRows)).PasteSpecial(XlPasteType.xlPasteFormats,,, True)

        ''Header row Height
        'CType(ReportSheet.Rows(REPORT_FIRSTROW - 1), Range).RowHeight = CType(ConfigSheet.Rows(PARAMS_FIRSTROW - 1), Range).RowHeight

        ''Rows Format
        'ConfigSheet.Range(
        '        ConfigSheet.Cells(PARAMS_FIRSTROW, PARAMS_COL_FORMAT),
        '        ConfigSheet.Cells(PARAMS_FIRSTROW + _Param_ColName_Rng.NbRows - 1, PARAMS_COL_FORMAT)).Copy()
        'ReportSheet.Range(
        '        ReportSheet.Cells(REPORT_FIRSTROW, 1),
        '        ReportSheet.Cells(REPORT_FIRSTROW + ReportNbRow - 1, _Param_ColName_Rng.NbRows)).PasteSpecial(XlPasteType.xlPasteFormats,,, True)
        'ReportSheet.Range(
        '        ReportSheet.Cells(REPORT_FIRSTROW, 1),
        '        ReportSheet.Cells(REPORT_FIRSTROW + ReportNbRow - 1, _Param_ColName_Rng.NbRows)).PasteSpecial(XlPasteType.xlPasteValidation,,, True)


        Dim ColIndex As Integer

        'Setup the columns *****************************************

        CType(ReportSheet.Rows(REPORT_FIRSTROW - 1), Range).RowHeight = 45
        CType(ReportSheet.Rows(REPORT_FIRSTROW - 1), Range).HorizontalAlignment = XlHAlign.xlHAlignCenter
        CType(ReportSheet.Rows(REPORT_FIRSTROW - 1), Range).VerticalAlignment = XlVAlign.xlVAlignTop
        CType(ReportSheet.Rows(REPORT_FIRSTROW - 1), Range).WrapText = True

        For ColIndex = 0 To _DBAdapter.Get_SummaryTable_Columns.Count - 1
            'Column
            If _Report_Param_Row(ColIndex) > 0 Then
                With ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW, _Report_ColNumber(ColIndex)), ReportSheet.Cells(REPORT_FIRSTROW + ReportNbRow - 1, _Report_ColNumber(ColIndex)))
                    .NumberFormat = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).NumberFormat
                    .Interior.Color = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Interior.Color
                    '.ColumnWidth = CellValue_Int(ConfigSheet, _Report_Param_Row(ColIndex), PARAMS_COL_WIDTH) 'we will do that after loading the data as it changes columns width
                    .Locked = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Locked
                    .Font.Name = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Font.Name
                    '.Font.Background = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Font.Background
                    .Font.Bold = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Font.Bold
                    .Font.Color = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Font.Color
                    '.Font.ColorIndex = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Font.ColorIndex
                    .Font.FontStyle = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Font.FontStyle
                    .Font.Italic = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Font.Italic
                    .Font.Size = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Font.Size
                    '.Font.ThemeColor = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Font.ThemeColor
                    '.Font.ThemeFont = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Font.ThemeFont
                    '.Font.TintAndShade = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Font.TintAndShade
                    .HorizontalAlignment = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).HorizontalAlignment
                    .VerticalAlignment = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).VerticalAlignment

                    'Validation
                    If HasValidation(CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range)) Then
                        Try
                            With .Validation
                                .Delete()
                                .Add(CType(CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Validation.Type, XlDVType),
                                    CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Validation.AlertStyle,
                                    CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Validation.Operator,
                                    CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Validation.Formula1,
                                    CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Validation.Formula2)

                                .IgnoreBlank = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Validation.IgnoreBlank
                                .InCellDropdown = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Validation.InCellDropdown
                                .InputTitle = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Validation.InputTitle
                                .ErrorTitle = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Validation.ErrorTitle
                                .InputMessage = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Validation.InputMessage
                                .ErrorMessage = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Validation.ErrorMessage
                                .ShowInput = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Validation.ShowInput
                                .ShowError = CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Validation.ShowError
                            End With
                        Catch ex As Exception
                            MsgBox("Incorrect data validation definition for column '" & _Param_ColHeader_Rng.CellValue_Str(_Report_ColNumber(ColIndex)) & "'", MsgBoxStyle.Critical, "Global planning Addin")
                        End Try
                    End If

                    'Conditional format
                    If CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).FormatConditions.Count > 0 Then
                        CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Copy()
                        .PasteSpecial(XlPasteType.xlPasteAllMergingConditionalFormats)
                        Globals.ThisWorkbook.Application.CutCopyMode = CType(0, XlCutCopyMode)
                    End If

                End With

                'CType(ConfigSheet.Cells(_Report_Param_Row(ColIndex), PARAMS_COL_FORMAT), Range).Copy()
                'ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW, _Report_ColNumber(ColIndex)), ReportSheet.Cells(REPORT_FIRSTROW + ReportNbRow - 1, _Report_ColNumber(ColIndex))).PasteSpecial(XlPasteType.xlPasteFormats)
                'ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW, _Report_ColNumber(ColIndex)), ReportSheet.Cells(REPORT_FIRSTROW + ReportNbRow - 1, _Report_ColNumber(ColIndex))).PasteSpecial(XlPasteType.xlPasteValidation)

                ''Header
                CType(ReportSheet.Cells(REPORT_FIRSTROW - 1, _Report_ColNumber(ColIndex)), Range).Value = _Param_ColHeader_Rng.CellValue_Str(_Report_ColNumber(ColIndex))
                CType(ReportSheet.Cells(REPORT_FIRSTROW - 1, _Report_ColNumber(ColIndex)), Range).Interior.Color = CType(ConfigSheet.Cells(PARAMS_FIRSTROW - 1, PARAMS_COL_HEADERTEXT), Range).Interior.Color
                CType(ReportSheet.Cells(REPORT_FIRSTROW - 1, _Report_ColNumber(ColIndex)), Range).Font.Color = CType(ConfigSheet.Cells(PARAMS_FIRSTROW - 1, PARAMS_COL_HEADERTEXT), Range).Font.Color
                CType(ReportSheet.Cells(REPORT_FIRSTROW - 1, _Report_ColNumber(ColIndex)), Range).Font.Bold = CType(ConfigSheet.Cells(PARAMS_FIRSTROW - 1, PARAMS_COL_HEADERTEXT), Range).Font.Bold
                CType(ReportSheet.Cells(REPORT_FIRSTROW - 1, _Report_ColNumber(ColIndex)), Range).Font.Name = CType(ConfigSheet.Cells(PARAMS_FIRSTROW - 1, PARAMS_COL_HEADERTEXT), Range).Font.Name

            End If
        Next ColIndex

    End Sub

    Private Sub CopyDataToReport(ReportNbRow As Integer)

        ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW, 1), ReportSheet.Cells(REPORT_FIRSTROW + ReportNbRow - 1, _ReportNbColumns)).Value = ConvertDataTableToArray()

        'Make sure the text in cells is not wrapped
        ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW, 1), ReportSheet.Cells(REPORT_FIRSTROW + ReportNbRow - 1, _ReportNbColumns)).WrapText = False


    End Sub

    Public Function ConvertDataTableToArray() As Object(,)
        Dim dtTable As System.Data.DataTable = _DBAdapter.SummaryTable_Dataset.Tables(0)
        Dim myArray(0 To dtTable.Rows.Count - 1, 0 To dtTable.Columns.Count - 1) As Object

        For i As Integer = 0 To dtTable.Rows.Count - 1
            For j As Integer = 0 To dtTable.Columns.Count - 1
                myArray(i, j) = dtTable.Rows(i)(j)
            Next
        Next

        Return myArray
    End Function

    Private Sub AdjustReportColumnsWidth(ReportNbRow As Integer)
        'Columns Width
        For ColIndex = 0 To _DBAdapter.Get_SummaryTable_Columns.Count - 1
            'Column
            If _Report_ColNumber(ColIndex) > 0 Then
                With ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW, _Report_ColNumber(ColIndex)), ReportSheet.Cells(REPORT_FIRSTROW + ReportNbRow - 1, _Report_ColNumber(ColIndex)))
                    .ColumnWidth = _Param_ColWidth_Rng.CellValue_Int(_Report_ColNumber(ColIndex))
                End With

                'Hidden columns -> Fix 30/09/2021 : Move this piece of code here as it was done before in the process, so setting the size of the column was preventing hiding columns
                If _Param_ColHidden_Rng.CellValue_Bool(_Report_ColNumber(ColIndex)) = True Then
                    CType(ReportSheet.Columns(_Report_ColNumber(ColIndex)), Range).Hidden = True
                End If
            End If
        Next
    End Sub

    Private Sub ApplyReportColumnsGrouping()
        Dim ParamRow As Integer
        Dim Groupping As Boolean
        Dim FirstGroupCol As Integer
        Dim LastGroupCol As Integer

        'Columns Groupping *************************
        Try
            ReportSheet.Columns.Ungroup()
        Catch
        End Try

        FirstGroupCol = 0
        'LastGroupCol = 0
        Groupping = False
        ParamRow = 1 'PARAMS_FIRSTROW
        Do While _Param_ColName_Rng.CellValue_Str(ParamRow) <> ""

            If _Param_ColGroupped_Rng.CellValue_Str(ParamRow) <> "" Then
                If Groupping = True Then
                    'Inside a group
                Else
                    'Start of group detected
                    Groupping = True
                    FirstGroupCol = ParamRow
                End If
            Else
                If Groupping = True Then
                    'End of group detected
                    Groupping = False
                    LastGroupCol = ParamRow - 1
                    ReportSheet.Range(ReportSheet.Columns(FirstGroupCol), ReportSheet.Columns(LastGroupCol)).Columns.Group()
                Else
                    'Not in a group
                End If
            End If

            ParamRow += 1
            If ParamRow > _Param_ColName_Rng.NbRows Then Exit Do
        Loop
        If Groupping = True Then 'If last column is in a group
            'End of group detected
            LastGroupCol = ParamRow - 1
            ReportSheet.Range(ReportSheet.Columns(FirstGroupCol), ReportSheet.Columns(LastGroupCol)).Columns.Group()
        End If
        ReportSheet.Outline.ShowLevels(ColumnLevels:=1)
    End Sub

    Private Sub AddReportFilterAndSort(ReportNbRow As Integer)

        'Add autofilter *******************************************
        ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW - 1, 1), ReportSheet.Cells(REPORT_FIRSTROW + ReportNbRow - 1, _ReportNbColumns)).AutoFilter(Field:=1) 'Field:=Report_Col_xxx, Criteria1:="=X"

        'Sort the report **********************************
        For i As Integer = 0 To _DBAdapter.Get_SummaryTable_DefaultSortColumns.Count - 1
            Dim ColIndex As Integer = _DBAdapter.GetColIndexFromDatabaseName(_DBAdapter.Get_SummaryTable_DefaultSortColumns(i))
            If _Report_ColNumber(ColIndex) > 0 Then 'is this column mapped in the report? if yes sort by it
                ReportSheet.AutoFilter.Sort.SortFields.Clear()
                ReportSheet.AutoFilter.Sort.SortFields.Add(ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW, _Report_ColNumber(ColIndex)), ReportSheet.Cells(REPORT_FIRSTROW + ReportNbRow - 1, _Report_ColNumber(ColIndex))), XlSortOn.xlSortOnValues, XlSortOrder.xlDescending, XlSortDataOption.xlSortNormal)
                With ReportSheet.AutoFilter.Sort
                    .Header = XlYesNoGuess.xlYes
                    .MatchCase = False
                    .Orientation = XlSortOrientation.xlSortColumns
                    .SortMethod = XlSortMethod.xlPinYin
                    .Apply()
                End With
                Exit For
            End If
        Next i

    End Sub

    Private Sub ReportFreezePanes()
        'Freeze panes ***************************
        ReportSheet.Activate()
        CType(ReportSheet.Cells(2, 3), Range).Select()
        Try
            Globals.ThisWorkbook.Windows(1).FreezePanes = False
        Catch
        End Try
        Globals.ThisWorkbook.Windows(1).FreezePanes = True

    End Sub

    Public Function CreateReport() As Boolean
        Dim ReportNbRow As Integer

        Try

            _ProgressWindow = New Form_Progress("Starting")
            _ProgressWindow.Show()

            Globals.ThisWorkbook.Application.ScreenUpdating = False

            _ProgressWindow.SetProgress(5, "Reading params")

            'Check that we have at least one column mapped -> this is not robust
            If ConfigSheet.UsedRange.Rows.End(XlDirection.xlDown).Row - PARAMS_FIRSTROW + 1 < 1 Then
                Globals.ThisWorkbook.Application.ScreenUpdating = True
                _LastMessage = "Error: Report columns must be defined in the config worksheet."
                CloseProgressWindow()
                Return False
            End If

            If ReadParams() = False Then
                Globals.ThisWorkbook.Application.ScreenUpdating = True
                CloseProgressWindow()
                Return False
            End If

            'check if the key columns are mapped
            Dim AllKeysAreMapped As Boolean = True
            Dim MissingKeyColumnName As String = ""
            For i As Integer = 0 To _DBAdapter.Get_SummaryTable_KeyColumns.Count - 1
                Dim ColIndex As Integer = _DBAdapter.GetColIndexFromDatabaseName(_DBAdapter.Get_SummaryTable_KeyColumns(i))
                If _Report_ColNumber(ColIndex) = 0 Then
                    AllKeysAreMapped = False
                    MissingKeyColumnName = _DBAdapter.GetColName(ColIndex)
                    Exit For
                End If
            Next
            If AllKeysAreMapped = False Then
                Globals.ThisWorkbook.Application.ScreenUpdating = True
                _LastMessage = "Error: " & MissingKeyColumnName & " must be mapped in the report."
                CloseProgressWindow()
                Return False
            End If

            SendDisplayFieldsToDBAdapter() 'List all the fields we want to display on the report
            SetupQueryFiltersInDBAdapter() 'Look for all the filters setup in the config sheet

            _ProgressWindow.SetProgress(10, "Reading data")

            ReportNbRow = _DBAdapter.Read_SummaryTable_Data(_ReportDate) 'Read data from database

            'stop if no data returned
            If ReportNbRow = 0 Then
                Globals.ThisWorkbook.Application.ScreenUpdating = True
                _LastMessage = "No data returned"
                CloseProgressWindow()
                Return False
            End If

            _ProgressWindow.SetProgress(65, "Cleaning old worksheet")

            CleanReportWorksheet() 'Delete everything from the report worksheet

            _ProgressWindow.SetProgress(70, "Formatting columns")
            ApplyReportColumnsSetup(ReportNbRow) 'Setup the columns

            _ProgressWindow.SetProgress(75, "Copying new data")
            CopyDataToReport(ReportNbRow) 'Copy the data to the report worksheet

            _ProgressWindow.SetProgress(80, "Adjusting columns width")
            AdjustReportColumnsWidth(ReportNbRow) 'Columns Width

            _ProgressWindow.SetProgress(85, "Adding column groups")
            ApplyReportColumnsGrouping() 'Columns Groupping

            _ProgressWindow.SetProgress(90, "Drawing borders")
            'Draw borders
            DrawAllBorders(ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW - 1, 1), ReportSheet.Cells(REPORT_FIRSTROW + ReportNbRow - 1, _ReportNbColumns)))

            _ProgressWindow.SetProgress(95, "Adding filter and sort")
            AddReportFilterAndSort(ReportNbRow) 'Add autofilter and sort the report
            ReportFreezePanes() 'Freeze panes (this also activates the report sheet)

            _ProgressWindow.SetProgress(100, "OK")

            Globals.ThisWorkbook.Application.ScreenUpdating = True

            _ReportAlreadyCreated = True

            CloseProgressWindow()
            Return True

        Catch ex As Exception
            Globals.ThisWorkbook.Application.ScreenUpdating = True
            _LastMessage = ex.Message
            CloseProgressWindow()
            Return False
        End Try


    End Function

    Private Sub CloseProgressWindow()
        If Not (_ProgressWindow Is Nothing) Then
            _ProgressWindow.Close()
            _ProgressWindow = Nothing
        End If
    End Sub

    Public Function SendUserModificationsToDB() As Boolean
        Dim ColIndex As Integer
        Dim ParamRow As Integer
        Dim AtLeastOneModifiableCol As Boolean = False
        Dim CurColNum As Integer
        Dim CurRowNum As Integer
        Dim ErrorsFound As New List(Of String)
        Dim OldDateValue As Nullable(Of Date)
        Dim NewDateValue As Nullable(Of Date)
        Dim OldNumericValue As Nullable(Of Double)
        Dim NewNumericValue As Nullable(Of Double)
        Dim OldStringValue As String
        Dim NewStringValue As String
        Dim KeyValues(_DBAdapter.Get_SummaryTable_KeyColumns.Count - 1) As String
        Dim KeyValuesConcatStr As String = ""

        'Check that the report was created first...
        If _ReportAlreadyCreated = False Then
            _LastMessage = "Error: The report has not yet been generated"
            Return False
        End If

        'check that the key columns are still mapped correctly
        For Each KeyColumn As String In _DBAdapter.Get_SummaryTable_KeyColumns

            Dim KeyColumnIndex As Integer = _DBAdapter.GetColIndexFromDatabaseName(KeyColumn)
            ParamRow = _Report_Param_Row(KeyColumnIndex)
            If CellValue_Str(ReportSheet, REPORT_FIRSTROW - 1, _Report_ColNumber(KeyColumnIndex)) <> CellValue_Str(ConfigSheet, ParamRow, PARAMS_COL_HEADERTEXT) Then
                MsgBox("Error: Columns have been modified. Column " & _Report_ColNumber(KeyColumnIndex) & " should be '" & CellValue_Str(ConfigSheet, ParamRow, PARAMS_COL_HEADERTEXT) & "'.", MsgBoxStyle.Critical, "Global planning Addin")
                Return False
            End If

        Next

        'check that the modifiable columns are still mapped correctly
        For Each ModifiableColumnName As String In _DBAdapter.Get_SummaryTable_ListOfModifiableColumns()
            ColIndex = _DBAdapter.GetColIndexFromDatabaseName(ModifiableColumnName)
            ParamRow = _Report_Param_Row(ColIndex)
            If ParamRow <> 0 Then 'if the row is listed for display
                AtLeastOneModifiableCol = True
                If CellValue_Str(ReportSheet, REPORT_FIRSTROW - 1, _Report_ColNumber(ColIndex)) <> CellValue_Str(ConfigSheet, ParamRow, PARAMS_COL_HEADERTEXT) Then
                    MsgBox("Error: Columns have been modified. Column " & _Report_ColNumber(ColIndex) & " should be '" & CellValue_Str(ConfigSheet, ParamRow, PARAMS_COL_HEADERTEXT) & "'.", MsgBoxStyle.Critical, "Global planning Addin")
                    Return False
                End If

            End If
        Next

        'If there is no modifiable column, no need to continue
        If AtLeastOneModifiableCol = False Then
            _LastMessage = "Error: None of the displayed columns are modifiable. Nothing to update."
            Return False
        End If

        _ProgressWindow = New Form_Progress("Checking modifications")
        _ProgressWindow.Show()


        _DBAdapter.ResetModifications() 'Reset the list of modifications. We will create a new one


        'Make a copy of the data in the Key & modifiable columns
        Dim ColumnsSnapshot As New List(Of ExcelRangeArray)

        For Each KeyColumn As String In _DBAdapter.Get_SummaryTable_KeyColumns

            Dim KeyColumnIndex As Integer = _DBAdapter.GetColIndexFromDatabaseName(KeyColumn)
            ColumnsSnapshot.Add(New ExcelRangeArray(KeyColumn,
                                                     ReportSheet.Range(
                                                            ReportSheet.Cells(REPORT_FIRSTROW, _Report_ColNumber(KeyColumnIndex)),
                                                            ReportSheet.Cells(ReportSheet.UsedRange.Rows.End(XlDirection.xlDown).Row, _Report_ColNumber(KeyColumnIndex))
                                                                )))
        Next

        For Each ColumnName As String In _DBAdapter.Get_SummaryTable_ListOfModifiableColumns
            ColIndex = _DBAdapter.GetColIndexFromDatabaseName(ColumnName)
            ParamRow = _Report_Param_Row(ColIndex)
            If ParamRow <> 0 Then 'if the row is listed for display

                ColumnsSnapshot.Add(New ExcelRangeArray(ColumnName, ReportSheet.Range(
                                                        ReportSheet.Cells(REPORT_FIRSTROW, _Report_ColNumber(ColIndex)),
                                                        ReportSheet.Cells(ReportSheet.UsedRange.Rows.End(XlDirection.xlDown).Row, _Report_ColNumber(ColIndex))
                                                            )))
            End If
        Next

        'Create a table to keep a reference to the key columns snapshots
        Dim KeyCol_Snapshot(_DBAdapter.Get_SummaryTable_KeyColumns.Count - 1) As ExcelRangeArray
        For i As Integer = 0 To _DBAdapter.Get_SummaryTable_KeyColumns.Count - 1
            Dim ColName = _DBAdapter.Get_SummaryTable_KeyColumns(i)
            KeyCol_Snapshot(i) = ColumnsSnapshot.Find(Function(c) c.ColumnName = ColName)
        Next


        CurRowNum = 1 'Start from row 1

        'Copy the values of the keys for the first row
        For i As Integer = 0 To _DBAdapter.Get_SummaryTable_KeyColumns.Count - 1
            KeyValues(i) = KeyCol_Snapshot(i).CellValue_Str(CurRowNum)
            KeyValuesConcatStr &= KeyValues(i)
            If i < _DBAdapter.Get_SummaryTable_KeyColumns.Count - 1 Then KeyValuesConcatStr &= "/"
        Next i

        Do While KeyValues(0) <> "" 'Loop while the first key has some value; Should we check also other keys?
            'process that row

            'Retrieve the record in the original dataset
            Dim OriginalDataSetRecord As DataRow = _DBAdapter.Get_SummaryTable_DatasetRecord(KeyValues)
            If Not (OriginalDataSetRecord Is Nothing) Then 'if we find it

                For Each ModifiableColumnName As String In _DBAdapter.Get_SummaryTable_ListOfModifiableColumns

                    ColIndex = _DBAdapter.GetColIndexFromDatabaseName(ModifiableColumnName)
                    CurColNum = _Report_ColNumber(ColIndex)

                    If CurColNum <> 0 Then 'if the row is listed for display

                        Dim Cur_Col As ExcelRangeArray = ColumnsSnapshot.Find(Function(c) c.ColumnName = ModifiableColumnName)

                        'Check data type for this column
                        Select Case _DBAdapter.GetColumnDataType(ModifiableColumnName)

                            Case "NUMERIC"

                                If IsDBNull(OriginalDataSetRecord.Item(ModifiableColumnName)) Then
                                    OldNumericValue = Nothing
                                Else
                                    OldNumericValue = CDbl(OriginalDataSetRecord.Item(ModifiableColumnName)) 'this cast should never fail!
                                End If

                                If Cur_Col.CellValue_Dbl(CurRowNum).HasValue = False Then
                                    NewNumericValue = Nothing
                                Else
                                    NewNumericValue = Math.Round(CDbl(Cur_Col.CellValue_Dbl(CurRowNum)), 3) 'Only 3 decimals can be stored in the DB
                                End If

                                If NewNumericValue = Double.MinValue Then
                                    ErrorsFound.Add(KeyValuesConcatStr & ": cannot convert the field '" & ModifiableColumnName & "' in numeric format")
                                Else
                                    If OldNumericValue <> NewNumericValue Or OldNumericValue.HasValue <> NewNumericValue.HasValue Then
                                        'We found a modification! Add it to the list
                                        _DBAdapter.AddUserValueModification(KeyValues, ModifiableColumnName, OldNumericValue, NewNumericValue, "NUMERIC", CurRowNum)
                                    End If
                                End If


                            Case "DATE"

                                If IsDBNull(OriginalDataSetRecord.Item(ModifiableColumnName)) OrElse CStr(OriginalDataSetRecord.Item(ModifiableColumnName)) = "" Then
                                    OldDateValue = Nothing 'DateSerial(1900, 1, 1)
                                Else
                                    OldDateValue = CDate(OriginalDataSetRecord.Item(ModifiableColumnName))
                                End If

                                NewDateValue = Cur_Col.CellValue_Date(CurRowNum)
                                If NewDateValue = Date.MinValue Then
                                    ErrorsFound.Add(KeyValuesConcatStr & ": cannot convert the field '" & ModifiableColumnName & "' in date format")
                                Else
                                    If OldDateValue <> NewDateValue Or OldDateValue.HasValue <> NewDateValue.HasValue Then
                                        'We found a modification! Add it to the list
                                        _DBAdapter.AddUserValueModification(KeyValues, ModifiableColumnName, OldDateValue, NewDateValue, "DATE", CurRowNum)
                                    End If
                                End If


                            Case "STRING"

                                If IsDBNull(OriginalDataSetRecord.Item(ModifiableColumnName)) Then
                                    OldStringValue = Nothing
                                Else
                                    OldStringValue = CStr(OriginalDataSetRecord.Item(ModifiableColumnName))
                                End If

                                NewStringValue = Cur_Col.CellValue_NullableStr(CurRowNum) 'string is nullable by default

                                If OldStringValue <> NewStringValue Then 'Or OldStringValue.HasValue <> NewDateValue.HasValue
                                    'We found a modification! Add it to the list
                                    _DBAdapter.AddUserValueModification(KeyValues, ModifiableColumnName, OldStringValue, NewStringValue, "STRING", CurRowNum)
                                End If

                            Case Else
                                _LastMessage = "Internal Error: Unknown data type."
                                Return False

                        End Select

                    End If
                Next
            Else
                'error, we can't find this SKU
                ErrorsFound.Add(KeyValuesConcatStr & ": not found in the original data snapshot")
            End If


            If CurRowNum Mod (Math.Ceiling(KeyCol_Snapshot(0).NbRows / 50)) = 0 Then 'split the progress bar updates in 50 steps
                _ProgressWindow.SetProgress(CInt(33 * CurRowNum / KeyCol_Snapshot(0).NbRows), "Checking modifications")
            End If

            'move to next row
            CurRowNum += 1
            If CurRowNum > KeyCol_Snapshot(0).NbRows Then Exit Do
            'Copy the values of the keys for the new row
            KeyValuesConcatStr = ""
            For i As Integer = 0 To _DBAdapter.Get_SummaryTable_KeyColumns.Count - 1
                KeyValues(i) = KeyCol_Snapshot(i).CellValue_Str(CurRowNum)
                KeyValuesConcatStr &= KeyValues(i)
                If i < _DBAdapter.Get_SummaryTable_KeyColumns.Count - 1 Then KeyValuesConcatStr &= "/"
            Next i
        Loop


        If ErrorsFound.Count <> 0 Then
            Dim ErrorMsg As String = ""
            For i As Integer = 0 To ErrorsFound.Count - 1
                ErrorMsg &= ErrorsFound(i) & vbCrLf
                If i = 20 Then
                    ErrorMsg &= "(...)"
                    Exit For
                End If
            Next i
            _ProgressWindow.Close()
            _ProgressWindow = Nothing
            'MsgBox("Errors to fix before data can be sent to database:" & vbCrLf & ErrorMsg)
            _LastMessage = "Errors to fix before data can be sent to database:" & vbCrLf & ErrorMsg
            Return False
        Else
            If _DBAdapter.ModificationsCount = 0 Then
                _ProgressWindow.Close()
                _ProgressWindow = Nothing
                MsgBox("Nothing to update.", MsgBoxStyle.Information, "Global planning Addin")
                Return True
            Else
                'All ok, now send the data to the database
                _ProgressWindow.SetProgress(33, "Sending modifications to database")

                _LastDBUpdateDateTime = Now
                _LastDBUpdateDateTime = New DateTime(_LastDBUpdateDateTime.Year, _LastDBUpdateDateTime.Month, _LastDBUpdateDateTime.Day, _LastDBUpdateDateTime.Hour, _LastDBUpdateDateTime.Minute, _LastDBUpdateDateTime.Second, _LastDBUpdateDateTime.Kind)

                Dim NbChangesProcessed As Integer = _DBAdapter.SendChangesToDB(_LastDBUpdateDateTime, _ReportDate, Environment.UserName, AddressOf UpdateProgress)
                If NbChangesProcessed > 0 Then MsgBox(NbChangesProcessed.ToString & " change" & CStr(IIf(NbChangesProcessed > 1, "s", "")) & " " & CStr(IIf(NbChangesProcessed > 1, "have", "has")) & " been sent successfully to the database", MsgBoxStyle.Information, "Global planning Addin")

                For Each AbandonnedConflict As SummaryTable_ValueModification In _DBAdapter.AbandonnedConflictsModifications
                    CType(ReportSheet.Cells(
                            AbandonnedConflict.ExcelReportRow + REPORT_FIRSTROW - 1,
                            _Report_ColNumber(_DBAdapter.GetColIndexFromDatabaseName(AbandonnedConflict.FieldName))), Range
                            ).Value = AbandonnedConflict.NewValue
                Next
            End If
        End If

        _ProgressWindow.Close()
        _ProgressWindow = Nothing
        Return True

    End Function

    Public Sub UpdateProgress(Progress As Integer, Status As String)
        _ProgressWindow.SetProgress(Progress, Status)
    End Sub

    Public Function GetKeyValuesForReportRow(ReportRow As Integer, ReportSheet As Worksheet) As String()

        Dim KeyValues(_DBAdapter.Get_SummaryTable_KeyColumns.Count - 1) As String
        For i As Integer = 0 To _DBAdapter.Get_SummaryTable_KeyColumns.Count - 1
            Dim ColIndex As Integer = _DBAdapter.GetColIndexFromDatabaseName(_DBAdapter.Get_SummaryTable_KeyColumns(i))
            KeyValues(i) = CStr(CType(ReportSheet.Cells(ReportRow, _Report_ColNumber(ColIndex)), Microsoft.Office.Interop.Excel.Range).Value)
        Next
        Return KeyValues

    End Function

    Public Function Get_ChangeLog(ReportRow As Integer, ReportColumn As Integer, ReportSheet As Worksheet) As DataSet


        If _DBAdapter.Read_ChangeLog(_ReportDate, GetKeyValuesForReportRow(ReportRow, ReportSheet), _DBAdapter.Get_SummaryTable_Columns(ReportColumn)) = True Then
            Return _DBAdapter.ChangeLog_Dataset
        Else
            Return Nothing
        End If


    End Function

    Public Function ReadDetailedProjectionData(ReportDate As Date, KeyValues() As String) As Boolean

        _ProgressWindow = New Form_Progress("Starting")
        _ProgressWindow.Show()

        Dim NbRows As Integer = _DBAdapter.ReadDetailedProjectionData(ReportDate, KeyValues)

        CopyDetailedProjectionDataToReport(NbRows)

        _ProgressWindow.SetProgress(100, "Done")

        _ProgressWindow.Close()
        _ProgressWindow = Nothing

        Return True
    End Function

    Private Sub CopyDetailedProjectionDataToReport(NbRows As Integer)

        '30/09/2021: Fix display issue. If there is a filter in place, the resulting data is messed up
        If Not (IsNothing(DetailsSheet.AutoFilter)) AndAlso DetailsSheet.AutoFilter.FilterMode = True Then DetailsSheet.AutoFilter.ShowAllData()

        DetailsSheet.Range(DetailsSheet.Cells(DETAILS_FIRSTROW, 1), DetailsSheet.Cells(DETAILS_FIRSTROW + 100000, _DBAdapter.Get_DetailedView_Columns.Count)).ClearContents()
        DetailsSheet.Range(DetailsSheet.Cells(DETAILS_FIRSTROW, 1), DetailsSheet.Cells(DETAILS_FIRSTROW + NbRows - 1, _DBAdapter.Get_DetailedView_Columns.Count)).Value = ConvertSummaryTableToArray(_DBAdapter.Get_DetailedView_Columns)

    End Sub


    Public Function ConvertSummaryTableToArray(MappedColumns() As String) As Object(,)
        Dim dtTable As System.Data.DataTable = _DBAdapter.DetailsTable_Dataset.Tables(0)
        Dim myArray(0 To dtTable.Rows.Count - 1, 0 To MappedColumns.Count - 1) As Object

        For i As Integer = 0 To dtTable.Rows.Count - 1
            For j As Integer = 0 To MappedColumns.Count - 1
                If dtTable.Columns(MappedColumns(j)).DataType = System.Type.GetType("System.DateTime") Then
                    If Not (dtTable.Rows(i)(MappedColumns(j)) Is DBNull.Value) AndAlso CDate(dtTable.Rows(i)(MappedColumns(j))).ToOADate = 0 Then
                        myArray(i, j) = Nothing
                    Else
                        myArray(i, j) = dtTable.Rows(i)(MappedColumns(j))
                    End If

                Else
                    myArray(i, j) = dtTable.Rows(i)(MappedColumns(j))
                End If

            Next
        Next

        Return myArray
    End Function

    Public Function ReadSingleSummaryTableRow(ReportDate As Date, KeyValues() As String) As Boolean
        If _DBAdapter.ReadSingleSummaryTableRowData(ReportDate, KeyValues) = False Then
            Return False
        Else
            Return True
        End If
    End Function


    Public Function Read_DetailsTable_Available_Dates(KeyValues() As String) As Boolean

        If _DBAdapter.Read_DetailsTable_Availabe_Dates(KeyValues) = False Then
            Return False
        Else
            Return True
        End If
    End Function

    Public ReadOnly Property DetailsTable_AvailableDates_DataSet As DataSet
        Get
            Return _DBAdapter.DetailsTable_AvailableDates_Dataset
        End Get
    End Property

    Public Function Get_DetailledView_HeaderText() As String

        If _DBAdapter.CurrentSummaryTableRow_Dataset Is Nothing Then Return "(Not Defined)"

        Get_DetailledView_HeaderText = ""
        For i As Integer = 0 To _DBAdapter.Get_DetailedView_CurItem_HeaderText.Count - 1

            Dim ParamStr As String = _DBAdapter.Get_DetailedView_CurItem_HeaderText(i)
            If Left(ParamStr, 1) = "'" And Right(ParamStr, 1) = "'" Then
                Get_DetailledView_HeaderText &= RemoveXMLReservedChars(Mid(ParamStr, 2, Len(ParamStr) - 2))
            Else
                Get_DetailledView_HeaderText &= RemoveXMLReservedChars(CStr(_DBAdapter.CurrentSummaryTableRow_Dataset.Tables(0).Rows(0).Item(ParamStr)))
            End If

        Next

    End Function

    Private Function RemoveXMLReservedChars(ByVal InputStr As String) As String
        'InputStr.Replace("""", "ˮ") '	For " similar looking chars: “ ” ″ ʺ ˮ 
        'InputStr.Replace("'", "′") '	For ' similar looking chars: ΄  ҆ ‘ ’ ′
        'InputStr.Replace("<", "˂") '	For < similar looking chars: ‹ ᐸ ˂
        'InputStr.Replace(">", "˃") '	For > similar looking chars: › ᐳ ˃
        InputStr = InputStr.Replace("&", "ꝸ") '    For & similar looking chars: ₰ ꝸ (only that one seems to be needed)
        Return InputStr
    End Function



    Public Function Get_ListOf_DetailedView_Info_Columns() As List(Of String())

        Dim TempList As New List(Of String())

        For i As Integer = 0 To _DBAdapter.Get_DetailedView_InfoDropDown_Items.Count - 1

            Dim ValueStr As String = ""

            For j As Integer = 1 To _DBAdapter.Get_DetailedView_InfoDropDown_Items(i).Count - 1
                Dim ParamStr As String = _DBAdapter.Get_DetailedView_InfoDropDown_Items(i)(j)
                If Left(ParamStr, 1) = "'" And Right(ParamStr, 1) = "'" Then
                    ValueStr &= Mid(ParamStr, 2, Len(ParamStr) - 2)
                Else
                    If Not (IsDBNull(_DBAdapter.CurrentSummaryTableRow_Dataset.Tables(0).Rows(0).Item(ParamStr))) Then
                        ValueStr &= CStr(_DBAdapter.CurrentSummaryTableRow_Dataset.Tables(0).Rows(0).Item(ParamStr))
                    End If
                End If
            Next j
            If Len(ValueStr) > 150 Then ValueStr = Left(ValueStr, 150) & "[...]"
            TempList.Add({_DBAdapter.Get_DetailedView_InfoDropDown_Items(i)(0).ToString, ValueStr})

        Next i
        Return TempList

    End Function


#Region "IDisposable Support"
    Private disposedValue As Boolean ' Pour détecter les appels redondants

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: supprimer l'état managé (objets managés).
            End If

            If Not (_DBAdapter Is Nothing) Then
                Try
                    _DBAdapter.Dispose()
                Catch
                End Try
            End If
        End If
        disposedValue = True
    End Sub

    ' TODO: remplacer Finalize() seulement si la fonction Dispose(disposing As Boolean) ci-dessus a du code pour libérer les ressources non managées.
    Protected Overrides Sub Finalize()
        ' Ne modifiez pas ce code. Placez le code de nettoyage dans Dispose(disposing As Boolean) ci-dessus.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' Ce code est ajouté par Visual Basic pour implémenter correctement le modèle supprimable.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Ne modifiez pas ce code. Placez le code de nettoyage dans Dispose(disposing As Boolean) ci-dessus.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class


Module DatabaseReaderUtils
    Public Function HasValidation(RangeToTest As Range) As Boolean
        Dim TempStr As String

        Try
            TempStr = RangeToTest.Validation.Formula1
            Return True
        Catch
            Return False
        End Try

    End Function

    'Draw all borders in the given range
    Public Sub DrawAllBorders(TheRange As Range)

        Try
            TheRange.Borders.LineStyle = XlLineStyle.xlContinuous
            TheRange.Borders.Weight = XlBorderWeight.xlThin
            TheRange.Borders.ColorIndex = XlColorIndex.xlColorIndexAutomatic
            'TheRange.Borders(XlBordersIndex.xlDiagonalDown).LineStyle = XlLineStyle.xlLineStyleNone
            'TheRange.Borders(XlBordersIndex.xlDiagonalUp).LineStyle = XlLineStyle.xlLineStyleNone
        Catch
        End Try

    End Sub

    Public Function CellValue_Str(Worksheet As Worksheet, Row As Integer, Col As Integer) As String
        If IsNothing(CType(Worksheet.Cells(Row, Col), Range).Value) Then
            Return ""
        Else
            Return CStr(CType(Worksheet.Cells(Row, Col), Range).Value)
        End If
    End Function

    'Public Function RangeValue_Str(Worksheet As Worksheet, RangeAdress As String) As String

    '    If IsNothing(Worksheet.Range(RangeAdress).Value) Then
    '        Return ""
    '    Else
    '        Return CStr(Worksheet.Range(RangeAdress).Value)
    '    End If
    'End Function

    'Public Function CellValue_Int(Worksheet As Worksheet, Row As Integer, Col As Integer) As Integer
    '    If IsNothing(CType(Worksheet.Cells(Row, Col), Range).Value) Then
    '        Return 0
    '    Else
    '        Try
    '            Return CInt(CType(Worksheet.Cells(Row, Col), Range).Value)
    '        Catch
    '            Return Nothing
    '            '   'Throw New System.Exception("Worksheet " & Worksheet.Name & " Row " & Row.ToString & " Col " & Col.ToString & " cannot be interpreted as an integer")
    '        End Try
    '    End If

    'End Function

    'Public Function CellValue_Dbl(Worksheet As Worksheet, Row As Integer, Col As Integer) As Double
    '    If IsNothing(CType(Worksheet.Cells(Row, Col), Range).Value) Then
    '        Return 0
    '    Else
    '        Try
    '            Return CDbl(CType(Worksheet.Cells(Row, Col), Range).Value)
    '        Catch
    '            Return Double.MinValue
    '            'Throw New System.Exception("Worksheet " & Worksheet.Name & " Row " & Row.ToString & " Col " & Col.ToString & " cannot be interpreted as a double")
    '        End Try
    '    End If
    'End Function

    'Public Function CellValue_Date(Worksheet As Worksheet, Row As Integer, Col As Integer) As Date
    '    If IsNothing(CType(Worksheet.Cells(Row, Col), Range).Value) Then
    '        Return DateSerial(1900, 1, 1)
    '    Else
    '        Try
    '            Return CDate(CType(Worksheet.Cells(Row, Col), Range).Value)
    '        Catch
    '            Return Date.MinValue
    '            'Throw New System.Exception("Worksheet " & Worksheet.Name & " Row " & Row.ToString & " Col " & Col.ToString & " cannot be interpreted as a date")
    '        End Try
    '    End If
    'End Function


End Module


Public Class ExcelRangeArray

    Public Property DataArray As Object(,)
    Public Property ColumnName As String
    Public Property NbRows As Integer

    Sub New(ColumnName As String, ExcelRange As Range)
        _ColumnName = ColumnName
        _DataArray = CType(ExcelRange.Value, Object(,)) 'DataArray
        _NbRows = _DataArray.GetUpperBound(0)
    End Sub

    Public Function CellValue_Str(Row As Integer) As String
        If IsNothing(DataArray(Row, 1)) Then
            Return ""
        Else
            Return CStr(DataArray(Row, 1))
        End If
    End Function

    Public Function CellValue_NullableStr(Row As Integer) As String
        If IsNothing(DataArray(Row, 1)) Then
            Return Nothing
        Else
            Return CStr(DataArray(Row, 1))
        End If
    End Function

    Public Function CellValue_Int(Row As Integer) As Integer
        If IsNothing(DataArray(Row, 1)) Then
            Return 0
        Else
            Try
                Return CInt(DataArray(Row, 1))
            Catch
                Return Nothing
                '   'Throw New System.Exception("Worksheet " & Worksheet.Name & " Row " & Row.ToString & " Col " & Col.ToString & " cannot be interpreted as an integer")
            End Try
        End If
    End Function

    Public Function CellValue_Dbl(Row As Integer) As Nullable(Of Double)
        If IsNothing(DataArray(Row, 1)) Then
            Return Nothing
        Else
            Try
                Return CDbl(DataArray(Row, 1))
            Catch
                Return Double.MinValue
                'Throw New System.Exception("Worksheet " & Worksheet.Name & " Row " & Row.ToString & " Col " & Col.ToString & " cannot be interpreted as a double")
            End Try
        End If
    End Function

    Public Function CellValue_Date(Row As Integer) As Nullable(Of Date)
        If IsNothing(DataArray(Row, 1)) Then
            Return Nothing 'DateSerial(1900, 1, 1)
        Else
            Try
                Return CDate(DataArray(Row, 1))
            Catch
                Return Date.MinValue
                'Throw New System.Exception("Worksheet " & Worksheet.Name & " Row " & Row.ToString & " Col " & Col.ToString & " cannot be interpreted as a date")
            End Try
        End If
    End Function

    Public Function CellValue_Bool(Row As Integer) As Boolean
        If IsNothing(DataArray(Row, 1)) Then
            Return False
        Else
            Try
                Return CBool(DataArray(Row, 1))
            Catch
                Return Nothing
                '   'Throw New System.Exception("Worksheet " & Worksheet.Name & " Row " & Row.ToString & " Col " & Col.ToString & " cannot be interpreted as an integer")
            End Try
        End If
    End Function


End Class