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

    'Private _LastDBUpdateDateTime As DateTime = Date.MinValue 'Date of the latest changes updated -> but useless as in fact we should keep a lastUpdatedDateTime per field

    Private _ReportNbColumns As Integer

    Private _ProgressWindow As Form_Progress

    Private _ReportNbRow As Integer

    Public Function HasModifiableColumns() As Boolean
        Return _DBAdapter.HasModifiableColumns()
    End Function

    Sub New(ReportDate As Date, DatabaseAdapterType As String, TemplateID As String)
        _ReportDate = ReportDate

        Select Case DatabaseAdapterType
            Case "SKUAlertsUI"
                _DBAdapter = New SKUAlertsDatabaseAdapter(TemplateID)
            Case "GRUT_UI"
                _DBAdapter = New GRUTDatabaseAdapter(TemplateID)
            Case "GRUT_MARKET_UI"
                _DBAdapter = New GRUTMarketDatabaseAdapter(TemplateID)
        End Select

        ReDim _Report_Param_Row(0 To _DBAdapter.SummaryTableColumns.Count - 1) 'Row number of each column in the params worksheet
        ReDim _Report_ColNumber(0 To _DBAdapter.SummaryTableColumns.Count - 1)  'Column number in the summary report

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



        For ColIndex = 0 To _DBAdapter.SummaryTableColumns.Count - 1
            _Report_Param_Row(ColIndex) = 0 '0 by default
            _Report_ColNumber(ColIndex) = 0
        Next ColIndex

        Dim FoundSw As Boolean
        ParamRow = 1
        _ReportNbColumns = 0
        Do While _Param_ColName_Rng.CellValue_Str(ParamRow) <> ""

            FoundSw = False
            For ColIndex = 0 To _DBAdapter.SummaryTableColumns.Count - 1
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

        Dim NbValues_Std_Incl As Integer
        Dim NbValues_Std_Excl As Integer
        Dim Values_Std_Text_Incl As String
        Dim Values_Std_Text_Excl As String

        Dim NbValues_Like_Incl As Integer
        Dim NbValues_Like_Excl As Integer
        Dim Values_Like_Text_Incl As String
        Dim Values_Like_Text_Excl As String


        Dim ColumnName As String

        _DBAdapter.ClearQueryFilters()
        ParamRow = 1 'PARAMS_FIRSTROW
        Do While _Param_ColName_Rng.CellValue_Str(ParamRow) <> ""



            If Trim(_Param_ColFilter_Rng.CellValue_Str(ParamRow)) <> "" Then
                'We found a new filter

                ColumnName = _Param_ColName_Rng.CellValue_Str(ParamRow)
                FilterValues = Split(Trim(Replace(_Param_ColFilter_Rng.CellValue_Str(ParamRow), "'", "''")), ",")

                NbValues_Std_Incl = 0
                NbValues_Std_Excl = 0
                Values_Std_Text_Incl = ""
                Values_Std_Text_Excl = ""

                NbValues_Like_Incl = 0
                NbValues_Like_Excl = 0
                Values_Like_Text_Incl = ""
                Values_Like_Text_Excl = ""

                For i = LBound(FilterValues) To UBound(FilterValues)

                    If FilterValues(i) = "{EMPTY}" Then FilterValues(i) = ""
                    If FilterValues(i) = "!{EMPTY}" Then FilterValues(i) = "!"

                    If Left(FilterValues(i), 1) = "!" Then
                        'This value needs to be EXcluded
                        If FilterValues(i).Contains("%") Then
                            'we found a wildcard, need to use the LIKE operator
                            NbValues_Like_Excl += 1
                            If NbValues_Like_Excl > 1 Then Values_Like_Text_Excl &= " AND "
                            Values_Like_Text_Excl &= SKUAlertsDatabaseAdapter.GetColDatabaseNameFromColName(ColumnName) & " NOT LIKE '" & Right(FilterValues(i), Len(FilterValues(i)) - 1) & "'"
                        Else
                            'no wildcard found, use the IN operator
                            NbValues_Std_Excl += 1
                            If NbValues_Std_Excl > 1 Then Values_Std_Text_Excl &= ","
                            Values_Std_Text_Excl = Values_Std_Text_Excl & "'" & Right(FilterValues(i), Len(FilterValues(i)) - 1) & "'"
                        End If
                    Else
                        'This value needs to be INcluded
                        If FilterValues(i).Contains("%") Then
                            'we found a wildcard, need to use the LIKE operator
                            NbValues_Like_Incl += 1
                            If NbValues_Like_Incl > 1 Then Values_Like_Text_Incl &= " OR "
                            Values_Like_Text_Incl &= SKUAlertsDatabaseAdapter.GetColDatabaseNameFromColName(ColumnName) & " LIKE '" & FilterValues(i) & "'"
                        Else
                            'no wildcard found, use the IN operator
                            NbValues_Std_Incl += 1
                            If NbValues_Std_Incl > 1 Then Values_Std_Text_Incl &= ","
                            Values_Std_Text_Incl = Values_Std_Text_Incl & "'" & FilterValues(i) & "'"
                        End If

                    End If

                Next i

                Dim FilterText As String = ""

                If NbValues_Std_Incl > 0 And NbValues_Std_Excl > 0 Then
                    FilterText &= "("
                End If

                If NbValues_Std_Incl > 0 Then
                    FilterText &= "ISNULL(" & SKUAlertsDatabaseAdapter.GetColDatabaseNameFromColName(ColumnName) & ",'') IN (" & Values_Std_Text_Incl & ")"
                End If

                If NbValues_Std_Excl > 0 Then
                    If NbValues_Std_Incl > 0 Then FilterText &= " AND "
                    FilterText &= "ISNULL(" & SKUAlertsDatabaseAdapter.GetColDatabaseNameFromColName(ColumnName) & ",'') NOT IN (" & Values_Std_Text_Excl & ")"
                End If

                If NbValues_Std_Incl > 0 And NbValues_Std_Excl > 0 Then
                    FilterText &= ")"
                End If

                If NbValues_Std_Incl + NbValues_Std_Excl > 0 And NbValues_Like_Incl + NbValues_Like_Excl > 0 Then
                    FilterText &= " OR "
                End If

                If NbValues_Like_Incl > 0 And NbValues_Like_Excl > 0 Then
                    FilterText &= "("
                End If

                If NbValues_Like_Incl > 0 Then
                    'If FilterText <> "" Then FilterText &= " AND "
                    FilterText &= " (" & Values_Like_Text_Incl & ")"
                End If

                If NbValues_Like_Excl > 0 Then
                    If NbValues_Like_Incl > 0 Then FilterText &= " AND "
                    FilterText &= " (" & Values_Like_Text_Excl & ")"
                End If

                If NbValues_Like_Incl > 0 And NbValues_Like_Excl > 0 Then
                    FilterText &= ")"
                End If


                _DBAdapter.AddQueryFilter("(" & FilterText & ")")

            End If

            ParamRow += 1
            If ParamRow > _Param_ColName_Rng.NbRows Then Exit Do
        Loop
    End Sub

    Private Sub CleanReportWorksheet()
        'Delete everything from the report worksheet
        If ReportSheet.ProtectContents = True Then ReportSheet.Unprotect()
        ReportSheet.Range(ReportSheet.Columns(1), ReportSheet.Columns(500)).Delete(XlDeleteShiftDirection.xlShiftToLeft)
        ReportSheet.Rows.ClearOutline()
    End Sub

    Private Sub ApplyReportColumnsSetup()

        'Header text
        ConfigSheet.Range(
                ConfigSheet.Cells(PARAMS_FIRSTROW, PARAMS_COL_HEADERTEXT),
                ConfigSheet.Cells(PARAMS_FIRSTROW + _Param_ColName_Rng.NbRows - 1, PARAMS_COL_HEADERTEXT)).Copy()
        ReportSheet.Range(
                ReportSheet.Cells(REPORT_FIRSTROW - 1, 1),
                ReportSheet.Cells(REPORT_FIRSTROW - 1, _Param_ColName_Rng.NbRows)).PasteSpecial(XlPasteType.xlPasteValues,,, True)

        'Header Format
        ConfigSheet.Range(
                ConfigSheet.Cells(PARAMS_FIRSTROW, PARAMS_COL_COLUMNNAME),
                ConfigSheet.Cells(PARAMS_FIRSTROW + _Param_ColName_Rng.NbRows - 1, PARAMS_COL_COLUMNNAME)).Copy()
        ReportSheet.Range(
                ReportSheet.Cells(REPORT_FIRSTROW - 1, 1),
                ReportSheet.Cells(REPORT_FIRSTROW - 1, _Param_ColName_Rng.NbRows)).PasteSpecial(XlPasteType.xlPasteFormats,,, True)

        'Copy Header comments
        ConfigSheet.Range(
                ConfigSheet.Cells(PARAMS_FIRSTROW, PARAMS_COL_HEADERTEXT),
                ConfigSheet.Cells(PARAMS_FIRSTROW + _Param_ColName_Rng.NbRows - 1, PARAMS_COL_HEADERTEXT)).Copy()
        ReportSheet.Range(
                ReportSheet.Cells(REPORT_FIRSTROW - 1, 1),
                ReportSheet.Cells(REPORT_FIRSTROW - 1, _Param_ColName_Rng.NbRows)).PasteSpecial(XlPasteType.xlPasteComments,,, True)

        'Header row Height
        CType(ReportSheet.Rows(REPORT_FIRSTROW - 1), Range).RowHeight = CType(ConfigSheet.Rows(PARAMS_FIRSTROW - 1), Range).RowHeight

        'Rows Format
        ConfigSheet.Range(
                ConfigSheet.Cells(PARAMS_FIRSTROW, PARAMS_COL_FORMAT),
                ConfigSheet.Cells(PARAMS_FIRSTROW + _Param_ColName_Rng.NbRows - 1, PARAMS_COL_FORMAT)).Copy()

        If LCase(Left(Globals.ThisWorkbook.Path, 4)) = "http" Then
            'If the template is opened from OneDrive, the copy paste is extremely slow. This is a dirty workarround
            Dim StepSize As Integer = CInt(200000 / _Param_ColName_Rng.NbRows)
            For FromLine As Integer = 0 To _ReportNbRow Step StepSize
                ReportSheet.Range(
                    ReportSheet.Cells(REPORT_FIRSTROW + FromLine, 1),
                    ReportSheet.Cells(REPORT_FIRSTROW + FromLine + StepSize - 1, _Param_ColName_Rng.NbRows)).PasteSpecial(XlPasteType.xlPasteFormats,,, True)
                ReportSheet.Range(
                    ReportSheet.Cells(REPORT_FIRSTROW + FromLine, 1),
                    ReportSheet.Cells(REPORT_FIRSTROW + FromLine + StepSize - 1, _Param_ColName_Rng.NbRows)).PasteSpecial(XlPasteType.xlPasteValidation,,, True)
            Next
        Else
            'The template is not synchrnoized with OneDrive, paste noramlly in one go
            ReportSheet.Range(
                    ReportSheet.Cells(REPORT_FIRSTROW, 1),
                    ReportSheet.Cells(REPORT_FIRSTROW + _ReportNbRow - 1, _Param_ColName_Rng.NbRows)).PasteSpecial(XlPasteType.xlPasteFormats,,, True)
            ReportSheet.Range(
                    ReportSheet.Cells(REPORT_FIRSTROW, 1),
                    ReportSheet.Cells(REPORT_FIRSTROW + _ReportNbRow - 1, _Param_ColName_Rng.NbRows)).PasteSpecial(XlPasteType.xlPasteValidation,,, True)
        End If


    End Sub

    Private Sub CopyDataToReport()

        ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW, 1), ReportSheet.Cells(REPORT_FIRSTROW + _ReportNbRow - 1, _ReportNbColumns)).Value = ConvertDataTableToArray()

        'Make sure the text in cells is not wrapped
        ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW, 1), ReportSheet.Cells(REPORT_FIRSTROW + _ReportNbRow - 1, _ReportNbColumns)).WrapText = False


    End Sub

    Public ReadOnly Property Report_Range_Incl_Header As Range
        Get
            Return ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW - 1, 1), ReportSheet.Cells(REPORT_FIRSTROW + _ReportNbRow - 1, _ReportNbColumns))
        End Get
    End Property


    Private Function ConvertDataTableToArray() As Object(,)
        Dim dtTable As System.Data.DataTable = _DBAdapter.SummaryTable_Dataset.Tables(0)
        Dim myArray(0 To dtTable.Rows.Count - 1, 0 To dtTable.Columns.Count - 1) As Object

        For i As Integer = 0 To dtTable.Rows.Count - 1
            For j As Integer = 0 To dtTable.Columns.Count - 1
                myArray(i, j) = dtTable.Rows(i)(j)
            Next
        Next

        Return myArray
    End Function

    Private Sub AdjustReportColumnsWidth()
        'Columns Width
        For ColIndex = 0 To _DBAdapter.SummaryTableColumns.Count - 1
            'Column
            If _Report_ColNumber(ColIndex) > 0 Then
                With ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW, _Report_ColNumber(ColIndex)), ReportSheet.Cells(REPORT_FIRSTROW + _ReportNbRow - 1, _Report_ColNumber(ColIndex)))
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

    Private Function CountWsFrozenRows(ws As Worksheet) As Integer
        Dim CurSheet As Worksheet
        CurSheet = DirectCast(ThisWorkbook.ActiveSheet, Worksheet)

        ws.Activate()
        If ws.Application.ActiveWindow.FreezePanes = True Then
            Return Globals.ThisWorkbook.Windows(1).SplitRow
        Else
            Return 0
        End If
        CurSheet.Activate() 'return to active sheet   
    End Function

    Private Function CountWsFrozenCols(ws As Worksheet) As Integer
        Dim CurSheet As Worksheet
        CurSheet = DirectCast(ThisWorkbook.ActiveSheet, Worksheet)

        ws.Activate()
        If ws.Application.ActiveWindow.FreezePanes = True Then
            Return Globals.ThisWorkbook.Windows(1).SplitColumn
        Else
            Return 0
        End If
        CurSheet.Activate() 'return to active sheet   
    End Function

    Private Sub WorksheetRemoveFreezePanes(ws As Worksheet)
        Dim CurSheet As Worksheet
        CurSheet = DirectCast(ThisWorkbook.ActiveSheet, Worksheet)
        ws.Activate()
        If ws.Application.ActiveWindow.FreezePanes = True Then
            ws.Application.ActiveWindow.FreezePanes = False
        End If
        CurSheet.Activate() 'return to active sheet   
    End Sub

    Private Sub WorksheetFreezePanes(ws As Worksheet, FrozenColumnsCount As Integer, FrozenRowsCount As Integer)
        ''Freeze panes ***************************
        ws.Activate()
        CType(ReportSheet.Cells(FrozenRowsCount + 1, FrozenColumnsCount + 1), Range).Select()
        If FrozenRowsCount > 0 Or FrozenColumnsCount > 0 Then
            ws.Application.ActiveWindow.FreezePanes = True
        End If

    End Sub

    Public Function CreateReport() As Boolean
        Dim ReportFrozenColumnsCount As Integer
        Dim ReportFrozenRowsCount As Integer

        Try

            _ProgressWindow = New Form_Progress("Starting")
            _ProgressWindow.Show()

            Globals.ThisWorkbook.Application.ScreenUpdating = False

            If ConfigSheet.AutoFilterMode = True Then ConfigSheet.AutoFilter.ShowAllData()

            _ProgressWindow.SetProgress(1, "Checking if data is ready")
            If Run_Preliminary_Check_Query() = False Then
                _LastMessage = "Process cancelled"
                CloseProgressWindow()
                Globals.ThisWorkbook.Application.ScreenUpdating = True
                Return False
            End If

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
            For i As Integer = 0 To _DBAdapter.SummaryTable_KeyColumns.Count - 1
                Dim ColIndex As Integer = _DBAdapter.GetColIndexFromDatabaseName(_DBAdapter.SummaryTable_KeyColumns(i).ColumnName)
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

            _ReportNbRow = _DBAdapter.Read_SummaryTable_Data(_ReportDate) 'Read data from database

            'stop if no data returned
            If _ReportNbRow = 0 Then
                Globals.ThisWorkbook.Application.ScreenUpdating = True
                _LastMessage = "No data returned"
                CloseProgressWindow()
                Return False
            End If

            _ProgressWindow.SetProgress(65, "Cleaning old worksheet")
            ReportFrozenColumnsCount = CountWsFrozenCols(ReportSheet)
            ReportFrozenRowsCount = CountWsFrozenRows(ReportSheet)
            WorksheetRemoveFreezePanes(ReportSheet)
            CleanReportWorksheet() 'Delete everything from the report worksheet

            _ProgressWindow.SetProgress(70, "Formatting columns")
            ApplyReportColumnsSetup() 'Setup the columns

            _ProgressWindow.SetProgress(75, "Copying new data")
            CopyDataToReport() 'Copy the data to the report worksheet

            _ProgressWindow.SetProgress(80, "Adjusting columns width")
            AdjustReportColumnsWidth() 'Columns Width

            _ProgressWindow.SetProgress(85, "Adding column groups")
            ApplyReportColumnsGrouping() 'Columns Groupping

            _ProgressWindow.SetProgress(90, "Drawing borders")
            'Draw borders
            DrawAllBorders(ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW - 1, 1), ReportSheet.Cells(REPORT_FIRSTROW + _ReportNbRow - 1, _ReportNbColumns)))

            _ProgressWindow.SetProgress(95, "Adding filter and sort")
            AutoFilterAndSortReport(_DBAdapter.SummaryTable_DefaultSortColumns)
            WorksheetFreezePanes(ReportSheet, ReportFrozenColumnsCount, ReportFrozenRowsCount) 'Freeze panes (this also activates the report sheet)

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

    Private Function Run_Preliminary_Check_Query() As Boolean
        Dim ResultStr As String = _DBAdapter.Run_Preliminary_Check_Query()
        If ResultStr <> "" Then
            If MsgBox("Warning! " & ResultStr & vbCrLf & "Do you want to continue anyway?", MsgBoxStyle.YesNo, "Global planning Addin") = MsgBoxResult.Yes Then
                Return True
            Else
                Return False
            End If
        Else
            Return True
        End If
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
        Dim KeyValues(_DBAdapter.SummaryTable_KeyColumns.Count - 1) As String
        Dim KeyValuesConcatStr As String = ""

        'Check that the report was created first...
        If _ReportAlreadyCreated = False Then
            _LastMessage = "Error: The report has not yet been generated"
            Return False
        End If

        'check that the key columns are still mapped correctly
        For Each KeyColumn As DatabaseAdapterColumn In _DBAdapter.SummaryTable_KeyColumns

            Dim KeyColumnIndex As Integer = _DBAdapter.GetColIndexFromDatabaseName(KeyColumn.ColumnName)
            ParamRow = _Report_Param_Row(KeyColumnIndex)
            If CellValue_Str(ReportSheet, REPORT_FIRSTROW - 1, _Report_ColNumber(KeyColumnIndex)) <> CellValue_Str(ConfigSheet, ParamRow, PARAMS_COL_HEADERTEXT) Then
                MsgBox("Error: Columns have been modified. Column " & _Report_ColNumber(KeyColumnIndex) & " should be '" & CellValue_Str(ConfigSheet, ParamRow, PARAMS_COL_HEADERTEXT) & "'.", MsgBoxStyle.Critical, "Global planning Addin")
                Return False
            End If

        Next

        'check that the modifiable columns are still mapped correctly
        For Each ModifiableColumn As DatabaseAdapterColumn In _DBAdapter.SummaryTable_ModifiableColumns
            ColIndex = _DBAdapter.GetColIndexFromDatabaseName(ModifiableColumn.ColumnName)
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

        For Each KeyColumn As DatabaseAdapterColumn In _DBAdapter.SummaryTable_KeyColumns

            Dim KeyColumnIndex As Integer = _DBAdapter.GetColIndexFromDatabaseName(KeyColumn.ColumnName)
            ColumnsSnapshot.Add(New ExcelRangeArray(KeyColumn.ColumnName,
                                                     ReportSheet.Range(
                                                            ReportSheet.Cells(REPORT_FIRSTROW, _Report_ColNumber(KeyColumnIndex)),
                                                            ReportSheet.Cells(ReportSheet.UsedRange.Rows.End(XlDirection.xlDown).Row, _Report_ColNumber(KeyColumnIndex))
                                                                )))
        Next

        For Each ModifiableColumn As DatabaseAdapterColumn In _DBAdapter.SummaryTable_ModifiableColumns
            ColIndex = _DBAdapter.GetColIndexFromDatabaseName(ModifiableColumn.ColumnName)
            ParamRow = _Report_Param_Row(ColIndex)
            If ParamRow <> 0 Then 'if the row is listed for display

                ColumnsSnapshot.Add(New ExcelRangeArray(ModifiableColumn.ColumnName, ReportSheet.Range(
                                                        ReportSheet.Cells(REPORT_FIRSTROW, _Report_ColNumber(ColIndex)),
                                                        ReportSheet.Cells(ReportSheet.UsedRange.Rows.End(XlDirection.xlDown).Row, _Report_ColNumber(ColIndex))
                                                            )))
            End If
        Next

        'Create a table to keep a reference to the key columns snapshots
        Dim KeyCol_Snapshot(_DBAdapter.SummaryTable_KeyColumns.Count - 1) As ExcelRangeArray
        For i As Integer = 0 To _DBAdapter.SummaryTable_KeyColumns.Count - 1
            Dim ColName = _DBAdapter.SummaryTable_KeyColumns(i).ColumnName
            KeyCol_Snapshot(i) = ColumnsSnapshot.Find(Function(c) c.ColumnName = ColName)
        Next


        CurRowNum = 1 'Start from row 1

        'Copy the values of the keys for the first row
        For i As Integer = 0 To _DBAdapter.SummaryTable_KeyColumns.Count - 1
            KeyValues(i) = KeyCol_Snapshot(i).CellValue_Str(CurRowNum)
            KeyValuesConcatStr &= KeyValues(i)
            If i < _DBAdapter.SummaryTable_KeyColumns.Count - 1 Then KeyValuesConcatStr &= "/"
        Next i

        'Create a table that says for each modifiable column if it is displayed, so we don't have to check for each row
        Dim NbModifiableColumns As Integer = _DBAdapter.SummaryTable_ModifiableColumns.Count
        Dim IsModifiableColumnMapped(0 To NbModifiableColumns - 1) As Boolean
        For i = 0 To NbModifiableColumns - 1
            ColIndex = _DBAdapter.GetColIndexFromDatabaseName(_DBAdapter.SummaryTable_ModifiableColumns(i).ColumnName)
            CurColNum = _Report_ColNumber(ColIndex)

            If CurColNum <> 0 Then
                IsModifiableColumnMapped(i) = True
            Else
                IsModifiableColumnMapped(i) = False
            End If
        Next

        Do While KeyValues(0) <> "" 'Loop while the first key has some value; Should we check also other keys?
            'process that row

            'Retrieve the record in the original dataset
            Dim OriginalDataSetRecord As DataRow = _DBAdapter.Get_SummaryTable_DatasetRecord(KeyValues)
            If Not (OriginalDataSetRecord Is Nothing) Then 'if we find it

                For i = 0 To NbModifiableColumns - 1

                    Dim ModifiableColumnName As String = _DBAdapter.SummaryTable_ModifiableColumns(i).ColumnName

                    If IsModifiableColumnMapped(i) Then

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
            For i As Integer = 0 To _DBAdapter.SummaryTable_KeyColumns.Count - 1
                KeyValues(i) = KeyCol_Snapshot(i).CellValue_Str(CurRowNum)
                KeyValuesConcatStr &= KeyValues(i)
                If i < _DBAdapter.SummaryTable_KeyColumns.Count - 1 Then KeyValuesConcatStr &= "/"
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

                Dim UpdateDateTime As DateTime = Now
                UpdateDateTime = New DateTime(UpdateDateTime.Year, UpdateDateTime.Month, UpdateDateTime.Day, UpdateDateTime.Hour, UpdateDateTime.Minute, UpdateDateTime.Second, UpdateDateTime.Kind)

                Dim NbChangesProcessed As Integer = _DBAdapter.SendChangesToDB(UpdateDateTime, _ReportDate, Environment.UserName, AddressOf UpdateProgress)
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

    Public Function Is_SummaryWorksheetColumn_Modifiable(ReportColumn As Integer) As Boolean 'Check if the given Worksheet column is modifiable
        If ReportColumn > _ReportNbColumns Or ReportColumn < 1 Then Return False
        Dim ColIndex As Integer = GetReportColIndex(ReportColumn)
        If ColIndex = -1 Then Return False
        Dim ColumnName As String = _DBAdapter.SummaryTableColumns(ColIndex).ColumnName
        Return _DBAdapter.SummaryTable_ModifiableColumns.Find(Function(x) x.ColumnName = ColumnName) IsNot Nothing '(ReportColumn).Contains(ColumnName)
    End Function
    Public Function GetKeyValues_For_SummaryReportRow(ReportRow As Integer, ReportSheet As Worksheet) As String()

        Dim KeyValues(_DBAdapter.SummaryTable_KeyColumns.Count - 1) As String
        For i As Integer = 0 To _DBAdapter.SummaryTable_KeyColumns.Count - 1
            Dim ColIndex As Integer = _DBAdapter.GetColIndexFromDatabaseName(_DBAdapter.SummaryTable_KeyColumns(i).ColumnName)
            KeyValues(i) = CStr(CType(ReportSheet.Cells(ReportRow, _Report_ColNumber(ColIndex)), Microsoft.Office.Interop.Excel.Range).Value)
        Next
        Return KeyValues

    End Function

    Public Function Get_ChangeLog(ReportRow As Integer, ReportColumn As Integer, ReportSheet As Worksheet) As DataSet

        If _DBAdapter.Read_ChangeLog(GetKeyValues_For_SummaryReportRow(ReportRow, ReportSheet), _DBAdapter.SummaryTableColumns(GetReportColIndex(ReportColumn)).ColumnName) = True Then
            Return _DBAdapter.ChangeLog_Dataset
        Else
            Return Nothing
        End If

    End Function

    Private Function GetReportColIndex(ColumnNumber As Integer) As Integer

        For ColIndex As Integer = 0 To _Report_ColNumber.Count - 1
            If _Report_ColNumber(ColIndex) = ColumnNumber Then
                Return ColIndex
            End If
        Next
        Return -1

    End Function

    Public Function ReadDetailedProjectionData(ReportDate As Date, KeyValues() As String) As Boolean

        _ProgressWindow = New Form_Progress("Reading Data")
        _ProgressWindow.Show()

        Dim NbRows As Integer = _DBAdapter.ReadDetailedProjectionData(ReportDate, KeyValues)

        _ProgressWindow.SetProgress(50, "Copying new data")

        CopyDetailedProjectionDataToReport(NbRows)

        _ProgressWindow.SetProgress(80, "Applying filters")

        Add_DetailedProjectionReport_AutoFilter(KeyValues)

        _ProgressWindow.SetProgress(100, "Done")

        _ProgressWindow.Close()
        _ProgressWindow = Nothing

        Return True
    End Function

    Private Sub Add_DetailedProjectionReport_AutoFilter(KeyValues() As String)
        Dim ColumnFilters As List(Of DatabaseAdapterBase.ColumnFilter)
        ColumnFilters = _DBAdapter.Get_DetailledView_ColumnFilter(KeyValues)

        'Add autofilter *******************************************
        Select Case ColumnFilters.Count
            Case 0
                DetailsSheet.Range(
                    DetailsSheet.Cells(DETAILS_FIRSTROW - 1, 1),
                    DetailsSheet.Cells(DETAILS_FIRSTROW + _ReportNbRow - 1, _DBAdapter.Get_DetailedView_Columns.Count)
                    ).AutoFilter(Field:=1)
            Case Else
                For i As Integer = 0 To ColumnFilters.Count - 1
                    DetailsSheet.Range(
                        DetailsSheet.Cells(DETAILS_FIRSTROW - 1, 1),
                        DetailsSheet.Cells(DETAILS_FIRSTROW + _ReportNbRow - 1, _DBAdapter.Get_DetailedView_Columns.Count)
                        ).AutoFilter(Field:=ColumnFilters(i).ColumnNumber, Criteria1:=ColumnFilters(i).FilterValue)
                Next


        End Select


    End Sub

    Private Sub CopyDetailedProjectionDataToReport(NbRows As Integer)

        '30/09/2021: Fix display issue. If there is a filter in place, the resulting data is messed up
        If Not (IsNothing(DetailsSheet.AutoFilter)) AndAlso DetailsSheet.AutoFilter.FilterMode = True Then DetailsSheet.AutoFilter.ShowAllData()

        DetailsSheet.Range(DetailsSheet.Cells(DETAILS_FIRSTROW, 1), DetailsSheet.Cells(DETAILS_FIRSTROW + 100000, _DBAdapter.Get_DetailedView_Columns.Count)).ClearContents()
        DetailsSheet.Range(DetailsSheet.Cells(DETAILS_FIRSTROW, 1), DetailsSheet.Cells(DETAILS_FIRSTROW + NbRows - 1, _DBAdapter.Get_DetailedView_Columns.Count)).Value = ConvertDetailsTableToArray(_DBAdapter.Get_DetailedView_Columns)

    End Sub


    Public Function ConvertDetailsTableToArray(MappedColumns() As String) As Object(,)
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

    Private Sub AutoFilterAndSortReport(SortFields As List(Of SortField))

        'Make sure the autofilter is setup
        If ReportSheet.AutoFilterMode = False Then
            ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW - 1, 1), ReportSheet.Cells(REPORT_FIRSTROW + _ReportNbRow - 1, _ReportNbColumns)).AutoFilter(Field:=1)
        End If

        ReportSheet.AutoFilter.Sort.SortFields.Clear()

        If Not (SortFields Is Nothing) AndAlso SortFields.Count > 0 Then

            For i As Integer = 0 To SortFields.Count - 1
                Dim ColNum As Integer = _Report_ColNumber(_DBAdapter.GetColIndexFromDatabaseName(SortFields(i).ColumnDatabaseName))
                If ColNum <> 0 Then
                    ReportSheet.AutoFilter.Sort.SortFields.Add(
                                ReportSheet.Range(ReportSheet.Cells(REPORT_FIRSTROW, ColNum), ReportSheet.Cells(REPORT_FIRSTROW + _ReportNbRow - 1, ColNum)),
                                XlSortOn.xlSortOnValues,
                                IIf(SortFields(i).SortOrder = SortField.SortOrders.Ascending, XlSortOrder.xlAscending, XlSortOrder.xlDescending),
                                XlSortDataOption.xlSortNormal)
                End If
            Next

            With ReportSheet.AutoFilter.Sort
                .Header = XlYesNoGuess.xlYes
                .MatchCase = False
                .Orientation = XlSortOrientation.xlSortColumns
                .SortMethod = XlSortMethod.xlPinYin
                .Apply()
            End With
        End If

    End Sub

    Public Sub GRUTReport_SortBySKURisk()

        If _DBAdapter.GetColIndexFromDatabaseName("ServiceRiskFactor") = -1 Then
            MsgBox("The column 'ServiceRiskFactor' must be present in the columns definition table", MsgBoxStyle.Critical, "Global planning Addin")
            Return
        End If
        If _Report_ColNumber(_DBAdapter.GetColIndexFromDatabaseName("ServiceRiskFactor")) = 0 Then
            MsgBox("The column 'ServiceRiskFactor' is mandatory in order to use the sorting", MsgBoxStyle.Critical, "Global planning Addin")
            Return
        End If

        'Clear filters if we have some
        If ReportSheet.AutoFilterMode = True Then
            If ReportSheet.AutoFilter.FilterMode = True Then
                If MsgBox("This will reset your filters. Do you want to continue?", MsgBoxStyle.YesNo, "Global planning Addin") = MsgBoxResult.No Then
                    Return
                Else
                    ReportSheet.AutoFilter.ShowAllData()
                End If
            End If
        End If

        Globals.ThisWorkbook.Application.ScreenUpdating = False

        'Clear the groups, sorting won't work properly otherwise.
        ReportSheet.Outline.ShowLevels(8) 'Display all levels before removing the outline, otherwise the lines remain hidden
        For i As Integer = 1 To 8 'there are 8 levels max in Excel
            Try
                ReportSheet.Rows.Ungroup()
            Catch ex As Exception 'Didn't find a way to count existing outline levels... so we need to try and catch the error
                Exit For
            End Try
        Next

        AutoFilterAndSortReport(New List(Of SortField)(
                {New SortField("ServiceRiskFactor", SortField.SortOrders.Descending)}
                )
           )

        Globals.ThisWorkbook.Application.ScreenUpdating = True
    End Sub

    Public Sub GRUTReport_SortByItemRisk_GroupBySKULevel()

        Dim RangeRowNo As Integer
        Dim CurrentLevel As Integer
        Dim NbColumnsToColor As Integer

        If _DBAdapter.GetColIndexFromDatabaseName("Item_ServiceRiskFactor") = -1 Or _DBAdapter.GetColIndexFromDatabaseName("Sourcing_Path") = -1 Then
            MsgBox("The columns 'Item_ServiceRiskFactor' and 'Sourcing_Path' must be present in the columns definition table", MsgBoxStyle.Critical, "Global planning Addin")
            Return
        End If
        If _Report_ColNumber(_DBAdapter.GetColIndexFromDatabaseName("Item_ServiceRiskFactor")) = 0 Then
            MsgBox("The column 'Item_ServiceRiskFactor' is mandatory in order to use the hierarchical sorting", MsgBoxStyle.Critical, "Global planning Addin")
            Return
        End If
        If _Report_ColNumber(_DBAdapter.GetColIndexFromDatabaseName("Sourcing_Path")) = 0 Then
            MsgBox("The column 'Sourcing_Path' is mandatory in order to use the hierarchical sorting", MsgBoxStyle.Critical, "Global planning Addin")
            Return
        End If

        'Clear filters if we have some
        If ReportSheet.AutoFilterMode = True Then
            If ReportSheet.AutoFilter.FilterMode = True Then
                If MsgBox("This will reset your filters. Do you want to continue?", MsgBoxStyle.YesNo, "Global planning Addin") = MsgBoxResult.No Then
                    Return
                Else
                    ReportSheet.AutoFilter.ShowAllData()
                End If
            End If
        End If

        NbColumnsToColor = CountWsFrozenCols(ReportSheet)

        Globals.ThisWorkbook.Application.ScreenUpdating = False

        _ProgressWindow = New Form_Progress("Formatting the report")
        _ProgressWindow.Show()



        'Clear the groups, sorting won't work properly otherwise. Didn't find a way to count existing outline levels... so we need to catch the error
        For i As Integer = 1 To 8 'there are 8 levels max in Excel
            Try
                ReportSheet.Rows.Ungroup()
            Catch ex As Exception
                Exit For
            End Try
        Next

        ReportSheet.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove 'shows the group + on the top line

        AutoFilterAndSortReport(New List(Of SortField)(
                        {New SortField("Item_ServiceRiskFactor", SortField.SortOrders.Descending),
                        New SortField("Item", SortField.SortOrders.Ascending),
                        New SortField("Sourcing_Path", SortField.SortOrders.Ascending)}
                        )
                   )

        Dim SKULevelColNum As Integer = _Report_ColNumber(_DBAdapter.GetColIndexFromDatabaseName("SKU_Level"))

        Dim SKULevel_Rng = New ExcelRangeArray("SKU_Level",
                                                   ReportSheet.Range(
                                                        ReportSheet.Cells(REPORT_FIRSTROW, SKULevelColNum),
                                                        ReportSheet.Cells(REPORT_FIRSTROW + _ReportNbRow - 1, SKULevelColNum)))

        'Make grouping
        For CurrentLevel = 3 To 1 Step -1

            RangeRowNo = 1
            Dim FirstGroupRow As Integer
            Dim InAGroup As Boolean = False
            Do
                If CInt(((SKULevel_Rng.NbRows * 3) / 100)) = 0 OrElse (RangeRowNo + (3 - CurrentLevel) * SKULevel_Rng.NbRows) Mod CInt(((SKULevel_Rng.NbRows * 3) / 100)) = 0 Then
                    _ProgressWindow.SetProgress(CInt(80 * ((RangeRowNo + (3 - CurrentLevel) * SKULevel_Rng.NbRows) / (SKULevel_Rng.NbRows * 3))), "Formatting the report")
                End If

                Dim CurrentRowLevel As Integer
                If SKULevel_Rng.CellValue_Str(RangeRowNo) = "" Then
                    CurrentRowLevel = 1
                Else
                    CurrentRowLevel = SKULevel_Rng.CellValue_Int(RangeRowNo)
                End If

                If CurrentRowLevel = CurrentLevel And Not InAGroup Then

                    FirstGroupRow = RangeRowNo
                    InAGroup = True

                ElseIf CurrentRowLevel < CurrentLevel And InAGroup Then

                    Dim FirstReportRowNo = FirstGroupRow + REPORT_FIRSTROW - 1
                    Dim LastReportRowNo = RangeRowNo - 1 + REPORT_FIRSTROW - 1
                    ReportSheet.Range(ReportSheet.Rows(FirstReportRowNo), ReportSheet.Rows(LastReportRowNo)).Rows.Group()
                    InAGroup = False

                End If

                RangeRowNo += 1
                If RangeRowNo > SKULevel_Rng.NbRows Then Exit Do
            Loop
            If InAGroup Then
                Dim FirstReportRowNo = FirstGroupRow + REPORT_FIRSTROW - 1
                Dim LastReportRowNo = RangeRowNo - 1 + REPORT_FIRSTROW - 1
                ReportSheet.Range(ReportSheet.Rows(FirstReportRowNo), ReportSheet.Rows(LastReportRowNo)).Rows.Group()
            End If
        Next

        'Coloring of the cells
        If NbColumnsToColor > 0 Then
            RangeRowNo = 1
            Dim ReportRowNo As Integer
            Do

                CurrentLevel = SKULevel_Rng.CellValue_Int(RangeRowNo)
                ReportRowNo = RangeRowNo + REPORT_FIRSTROW - 1

                If CInt((SKULevel_Rng.NbRows / 20)) = 0 OrElse RangeRowNo Mod CInt((SKULevel_Rng.NbRows / 20)) = 0 Then
                    _ProgressWindow.SetProgress(80 + CInt(20 * (RangeRowNo / SKULevel_Rng.NbRows)), "Coloring the report")
                End If

                If SKULevel_Rng.CellValue_Str(RangeRowNo) = "" Then
                    'Gray
                    ReportSheet.Range(ReportSheet.Cells(ReportRowNo, 1), ReportSheet.Cells(ReportRowNo, NbColumnsToColor)).Interior.Color = XlRgbColor.rgbLightGray 'HslToRgba(148 / 255, 150 / 255, (114 + CurrentLevel * 40) / 255)

                Else
                    'Blue
                    ReportSheet.Range(ReportSheet.Cells(ReportRowNo, 1), ReportSheet.Cells(ReportRowNo, NbColumnsToColor)).Interior.Color = HslToRgba(148 / 255, 150 / 255, (114 + CurrentLevel * 40) / 255)
                    If CurrentLevel = 0 Then ReportSheet.Range(ReportSheet.Cells(ReportRowNo, 1), ReportSheet.Cells(ReportRowNo, NbColumnsToColor)).Font.Color = System.Drawing.Color.White
                End If

                RangeRowNo += 1
                If RangeRowNo > SKULevel_Rng.NbRows Then Exit Do
            Loop
        End If

        'Close the groups one by one
        ReportSheet.Outline.ShowLevels(4)
        ReportSheet.Outline.ShowLevels(3)
        ReportSheet.Outline.ShowLevels(2)
        ReportSheet.Outline.ShowLevels(1)

        CloseProgressWindow()
        Globals.ThisWorkbook.Application.ScreenUpdating = True

    End Sub


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

    Public Function HslToRgba(h As Double, s As Double, l As Double) As System.Drawing.Color
        Dim r, g, b As Double

        If h > 1 Then h = 1
        If s > 1 Then s = 1
        If l > 1 Then l = 1

        If s = 0 Then
            r = l
            g = l
            b = l
        Else
            Dim q As Double = If(l < 0.5, l * (1 + s), l + s - l * s)
            Dim p As Double = 2 * l - q
            r = HueToRgb(p, q, h + 1 / 3)
            g = HueToRgb(p, q, h)
            b = HueToRgb(p, q, h - 1 / 3)
        End If

        Return System.Drawing.Color.FromArgb(CInt((r * 255)), CInt((g * 255)), CInt((b * 255)))
    End Function

    Private Function HueToRgb(ByVal p As Double, ByVal q As Double, ByVal t As Double) As Double
        If t < 0 Then t += 1
        If t > 1 Then t -= 1
        If t < 1 / 6 Then Return p + (q - p) * 6 * t
        If t < 1 / 2 Then Return q
        If t < 2 / 3 Then Return p + (q - p) * (2 / 3 - t) * 6
        Return p
    End Function

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