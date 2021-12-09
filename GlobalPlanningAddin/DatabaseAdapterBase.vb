Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Threading
Public MustInherit Class DatabaseAdapterBase : Implements IDisposable

    Protected MustOverride Function Get_ConnectionString() As String
    Protected MustOverride Function Get_DatabaseSchema() As String
    Protected MustOverride Function Get_SummaryTable_Name() As String
    Protected MustOverride Function Get_SummaryTableUpdates_TableName() As String
    Protected MustOverride Function Get_DetailsTable_Name() As String
    Public MustOverride Function Get_SummaryTable_ListOfModifiableColumns() As String()
    Public MustOverride Function Get_SummaryTable_ListOfNumericColumns() As String()
    Public MustOverride Function Get_SummaryTable_ListOfDateColumns() As String()
    Public MustOverride Function Get_SummaryTable_KeyColumns() As String()
    Public MustOverride Function Get_SummaryTable_Columns() As String()
    Public MustOverride Function Get_SummaryTable_SQLQueryForField(DatabaseColName As String) As String
    Public MustOverride Function Get_SummaryTable_DefaultSortColumns() As String()
    Protected Overridable Function Get_DetailledView_Optional_OrderBySQLClause() As String 'Optional ORDER BY clause used when querying the detailled data
        Return ""
    End Function
    Public MustOverride Function Get_DetailedView_Columns() As String()
    Public MustOverride Function Get_DetailedView_CurItem_HeaderText() As String()
    Public MustOverride Function Get_DetailedView_InfoDropDown_Items() As List(Of String())
    Public Class ColumnFilter
        Public Property ColumnNumber As Integer
        Public Property FilterValue As String
    End Class

    Public Overridable Function Get_DetailledView_ColumnFilter(KeyValues() As String) As List(Of ColumnFilter)
        Static Columnsfilters As New List(Of ColumnFilter) 'no need to create a new list each time this is called
        Return Columnsfilters 'By default, we return an empty list
    End Function

    Protected ReadOnly _Connection As SqlConnection
    Protected _Command As SqlCommand
    Protected _Adapter As SqlDataAdapter
    Protected _ResultDataSet As DataSet
    Public Property SummaryTable_Dataset As DataSet
    Public Property DetailsTable_Dataset As DataSet
    Public Property CurrentSummaryTableRow_Dataset As DataSet
    Public Property DetailsTable_AvailableDates_Dataset As DataSet
    Public Property ChangeLog_Dataset As DataSet

    Protected ReadOnly _DisplayedFields As New List(Of String)
    Protected ReadOnly _QueryFilters As New List(Of String)

    Protected ReadOnly _ValueModifications As New List(Of SummaryTable_ValueModification)
    Public Property AbandonnedConflictsModifications As New List(Of SummaryTable_ValueModification)
    Public Property CopyToDBPercentCompleted As Single '??????
    Protected _ProgressCallback As System.Action(Of Integer, String)


    Public Function GetColName(ColIndex As Integer) As String
        Return "[" & Get_SummaryTable_Columns(ColIndex) & "]"
    End Function

    Public Shared Function GetColDatabaseNameFromColName(ColName As String) As String
        Dim StrLen As Integer

        StrLen = Len(ColName)
        Return Right(Left(ColName, StrLen - 1), StrLen - 2)

    End Function

    Public Function GetColIndexFromDatabaseName(ColDatabaseName As String) As Integer

        For i As Integer = 0 To Get_SummaryTable_Columns.Count - 1
            If Get_SummaryTable_Columns(i) = ColDatabaseName Then
                Return i
            End If
        Next
        Return -1 'we didn't find it

    End Function

    Public Function GetColumnDataType(DatabaseColumnName As String) As String

        'Check if numeric format
        For Each NumericColumnName As String In Get_SummaryTable_ListOfNumericColumns()
            If DatabaseColumnName = NumericColumnName Then
                Return "NUMERIC"
            End If
        Next

        'Check if date format
        For Each DateColumnName As String In Get_SummaryTable_ListOfDateColumns()
            If DatabaseColumnName = DateColumnName Then
                Return "DATE"
            End If
        Next
        Return "STRING" 'string by default

    End Function

    Private ReadOnly _CultureInf As Globalization.CultureInfo

    Sub New()

        'Connect to SQL server
        _Connection = GetSQLConnection(Get_ConnectionString())
        If CheckSQLConnectionAndReconnect(_Connection, 5) = False Then
            Throw New System.Exception("Unable to connect to the database server")
        End If

        'Create our own CultureInfo to match SQL standards
        _CultureInf = DirectCast(Globalization.CultureInfo.InvariantCulture.Clone(), Globalization.CultureInfo)
        _CultureInf.NumberFormat.NumberDecimalSeparator = "."
        _CultureInf.NumberFormat.NumberGroupSeparator = ""
        _CultureInf.DateTimeFormat.DateSeparator = "-"
        _CultureInf.DateTimeFormat.ShortDatePattern = "yyyy-MM-dd"

    End Sub



    Public Sub ClearDisplayFields()
        _DisplayedFields.Clear()
    End Sub
    Public Sub AddDisplayField(FieldName As String)
        _DisplayedFields.Add(FieldName)
    End Sub

    Public Sub ClearQueryFilters()
        _QueryFilters.Clear()
    End Sub
    Public Sub AddQueryFilter(FilterText As String)
        _QueryFilters.Add(FilterText)
    End Sub

    Public Sub ResetModifications()
        _ValueModifications.Clear()
        _AbandonnedConflictsModifications.Clear()
    End Sub

    Public Function ModificationsCount() As Integer
        Return _ValueModifications.Count
    End Function

    Public Function Read_SummaryTable_Data(ReportDate As Date) As Integer
        Dim ex As Exception
        Dim SQLQuery As String
        Dim i As Integer
        Dim NbFailedQueries As Integer
        Dim ReportNbRow As Integer

        If Not (_ResultDataSet Is Nothing) Then _ResultDataSet.Dispose()
        If Not (_Adapter Is Nothing) Then _Adapter.Dispose()
        If Not (_Command Is Nothing) Then _Command.Dispose()

        If _DisplayedFields.Count = 0 Then Throw New System.Exception("Please setup the list of fields to be displayed")

        ' Create the SQL query to read data from database
        SQLQuery = "SELECT "

        For i = 0 To _DisplayedFields.Count - 1
            Dim ColumnName As String = _DisplayedFields(i)
            SQLQuery &= Get_SummaryTable_SQLQueryForField(GetColDatabaseNameFromColName(ColumnName))
            If i < _DisplayedFields.Count - 1 Then SQLQuery &= ","
        Next

        SQLQuery = SQLQuery & " FROM [" & Get_DatabaseSchema() & "].[" & Get_SummaryTable_Name() & "] WHERE ReportDate = '" & Format(ReportDate, "yyyy-MM-dd") & "' "

        For Each FilterText As String In _QueryFilters
            SQLQuery = SQLQuery & "AND " & FilterText & " "
        Next

        SQLQuery &= ";"

        'Now trigger it
        NbFailedQueries = 0
        Do

            _Command = New SqlCommand(SQLQuery, _Connection)
            _Adapter = New SqlDataAdapter(_Command)
            _ResultDataSet = New DataSet

            ex = Nothing
            Try
                ReportNbRow = _Adapter.Fill(_ResultDataSet, "ResultTable") 'Try to run the query, and update the number of rows
            Catch ex
                'Dispose these object, we will recreate new instances in the next loop
                _ResultDataSet.Dispose()
                _Adapter.Dispose()
                _Command.Dispose()

                NbFailedQueries += 1
                'AddMessageToStack(New LogMessage("Warning: Unable to xxxx (attempt " & NbFailedQueries.ToString(Globalization.CultureInfo.InvariantCulture) & ") : " & ex.Message, LogLevel.LogWarning))

                If NbFailedQueries = 5 Then 'after 5 failed trials, we give up for now
                    'AddMessageToStack(New LogMessage("Error: Unable to xxxx (will retry in 5 min) : " & ex.Message, LogLevel.LogError))
                    Throw New System.Exception("Error while running the data read query from the database")
                Else
                    Thread.Sleep(NbFailedQueries * 1000) 'if we tried less than 5 times, pause the thread for an increasing amount of time until next trial
                End If

                'Re-check the connection to the database server for next loop
                If CheckSQLConnectionAndReconnect(_Connection, 5) = False Then
                    Throw New System.Exception("Error: Unable to connect to the database")
                End If

            End Try

        Loop While Not (ex Is Nothing)

        _SummaryTable_Dataset = _ResultDataSet.Copy()
        _ResultDataSet.Dispose()
        _Adapter.Dispose()
        _Command.Dispose()

        Dim PrimaryKeyColumns(Get_SummaryTable_KeyColumns.Count - 1) As DataColumn
        For i = 0 To Get_SummaryTable_KeyColumns.Count - 1
            PrimaryKeyColumns(i) = _SummaryTable_Dataset.Tables(0).Columns(Get_SummaryTable_KeyColumns(i))
        Next
        _SummaryTable_Dataset.Tables(0).PrimaryKey = PrimaryKeyColumns

        Return ReportNbRow

    End Function

    Public Sub AddUserValueModification(KeyValues As String(), FieldName As String, OldValue As Object, NewValue As Object, DataType As String, ExcelReportRow As Integer)

        _ValueModifications.Add(New SummaryTable_ValueModification(KeyValues, FieldName, OldValue, NewValue, DataType, ExcelReportRow))

    End Sub



    Protected Sub ModifyValue(KeyValues() As String, FieldName As String, NewValue As Object)
        If _SummaryTable_Dataset Is Nothing Then
            Throw New System.Exception("Error: There is no dataset available")
        End If

        Dim dataRow As DataRow = Get_SummaryTable_DatasetRecord(KeyValues)

        If dataRow Is Nothing Then
            Throw New System.Exception("Error: Row not found in the dataset")
        Else

            Select Case GetColumnDataType(FieldName)

                Case "NUMERIC"
                    If IsDBNull(NewValue) Then
                        dataRow(FieldName) = DBNull.Value
                    Else
                        dataRow(FieldName) = Double.Parse(DirectCast(NewValue, String), _CultureInf)
                    End If

                Case "DATE"
                    If IsDBNull(NewValue) Then
                        dataRow(FieldName) = DBNull.Value
                    Else
                        dataRow(FieldName) = Date.Parse(DirectCast(NewValue, String), _CultureInf)
                    End If

                Case "STRING"
                    If IsDBNull(NewValue) Then
                        dataRow(FieldName) = DBNull.Value
                    Else
                        dataRow(FieldName) = NewValue
                    End If


                Case Else
                    dataRow(FieldName) = NewValue

            End Select

        End If

    End Sub

    Public Function Get_SummaryTable_DatasetRecord(KeyValues() As String) As DataRow

        If _SummaryTable_Dataset Is Nothing Then
            Throw New System.Exception("Error: There is no dataset available")
        End If

        Dim foundRow As DataRow = _SummaryTable_Dataset.Tables("ResultTable").Rows.Find(KeyValues)

        Return foundRow

    End Function

    Public Function SendChangesToDB(ChangeDateTime As DateTime, ReportDate As Date, UserName As String, ProgressCallback As System.Action(Of Integer, String)) As Integer

        Dim ex As Exception
        Dim SQLQuery As String
        Dim NbFailedQueries As Integer
        Dim NewDataTable As DataTable
        Dim ErrorsFound As New List(Of String)
        Dim ConflictForm As Form_Conflict
        Dim UserChoiceToAllConflicts As String = ""
        Dim UserChoice As String = ""
        Dim NeedReprocess As Boolean = False
        Dim NbChangesProcessed As Integer = 0
        Dim KeyValues As String()
        Dim ColumnName As String
        Dim ReportRowNumber As Integer

        _ProgressCallback = ProgressCallback

        CopyToDBPercentCompleted = 0 'This percentage will inform about the progress of data copy

        If _ValueModifications.Count = 0 Then 'if there is nothing to do, return immediately
            Return 0
        End If


        ' Create a new DataTable object
        NewDataTable = New DataTable(Get_SummaryTableUpdates_TableName())

        ' Add columns to the table
        Dim keys(4 + Get_SummaryTable_KeyColumns.Count) As DataColumn
        Dim TableColumn As DataColumn

        TableColumn = New DataColumn("ChangeDateTime", System.Type.GetType("System.DateTime")) : keys(0) = TableColumn : NewDataTable.Columns.Add(TableColumn)
        TableColumn = New DataColumn("ReportDate", System.Type.GetType("System.DateTime")) : keys(1) = TableColumn : NewDataTable.Columns.Add(TableColumn)
        For i As Integer = 0 To Get_SummaryTable_KeyColumns.Count - 1 'Create one column per Key
            TableColumn = New DataColumn(Get_SummaryTable_KeyColumns(i), System.Type.GetType("System.String"))
            keys(i + 2) = TableColumn
            NewDataTable.Columns.Add(TableColumn)
        Next
        TableColumn = New DataColumn("ChangedBy", System.Type.GetType("System.String")) : keys(Get_SummaryTable_KeyColumns.Count + 2) = TableColumn : NewDataTable.Columns.Add(TableColumn)
        TableColumn = New DataColumn("ColumName", System.Type.GetType("System.String")) : keys(Get_SummaryTable_KeyColumns.Count + 3) = TableColumn : NewDataTable.Columns.Add(TableColumn)
        TableColumn = New DataColumn("OldValue", System.Type.GetType("System.String")) : NewDataTable.Columns.Add(TableColumn)
        TableColumn = New DataColumn("NewValue", System.Type.GetType("System.String")) : NewDataTable.Columns.Add(TableColumn)
        TableColumn = New DataColumn("Status", System.Type.GetType("System.String")) : NewDataTable.Columns.Add(TableColumn)
        TableColumn = New DataColumn("Comment", System.Type.GetType("System.String")) : NewDataTable.Columns.Add(TableColumn)

        ' Set the primary keys 
        NewDataTable.PrimaryKey = keys

        'Add the rows to the Data table object
        Dim NewDataRow As DataRow
        For Each ValueModification As SummaryTable_ValueModification In _ValueModifications

            If ValueModification.Processed = False Then

                ValueModification.Processed = True
                NewDataRow = NewDataTable.NewRow() 'Create a new row object with the right format
                For i As Integer = 0 To Get_SummaryTable_KeyColumns.Count - 1 'Create one column per Key
                    NewDataRow.Item(Get_SummaryTable_KeyColumns(i)) = ValueModification.KeyValues(i)
                Next
                NewDataRow.Item("ChangeDateTime") = ChangeDateTime
                NewDataRow.Item("ReportDate") = ReportDate
                NewDataRow.Item("ChangedBy") = UserName
                NewDataRow.Item("ColumName") = ValueModification.FieldName
                Select Case ValueModification.DataType
                    Case "STRING"
                        If IsNothing(DirectCast(ValueModification.OldValue, String)) = True Then
                            NewDataRow.Item("OldValue") = Nothing
                        Else
                            NewDataRow.Item("OldValue") = ValueModification.OldValue.ToString
                        End If

                        If IsNothing(DirectCast(ValueModification.NewValue, String)) = True Then
                            NewDataRow.Item("NewValue") = Nothing
                        Else
                            NewDataRow.Item("NewValue") = ValueModification.NewValue.ToString
                        End If

                    Case "NUMERIC"
                        If DirectCast(ValueModification.OldValue, Nullable(Of Double)).HasValue = False Then
                            NewDataRow.Item("OldValue") = Nothing
                        Else
                            NewDataRow.Item("OldValue") = CDbl(ValueModification.OldValue).ToString(Globalization.CultureInfo.InvariantCulture)
                        End If

                        If DirectCast(ValueModification.NewValue, Nullable(Of Double)).HasValue = False Then
                            NewDataRow.Item("NewValue") = Nothing
                        Else
                            NewDataRow.Item("NewValue") = CDbl(ValueModification.NewValue).ToString(Globalization.CultureInfo.InvariantCulture)
                        End If

                    Case "DATE"
                        If DirectCast(ValueModification.OldValue, Nullable(Of Date)).HasValue = False Then
                            NewDataRow.Item("OldValue") = Nothing
                        Else
                            NewDataRow.Item("OldValue") = CDate(ValueModification.OldValue).ToString("yyyy'-'MM'-'dd")
                        End If

                        If DirectCast(ValueModification.NewValue, Nullable(Of Date)).HasValue = False Then
                            NewDataRow.Item("NewValue") = Nothing
                        Else
                            NewDataRow.Item("NewValue") = CDate(ValueModification.NewValue).ToString("yyyy'-'MM'-'dd")
                        End If

                End Select
                NewDataRow.Item("Status") = "WAITING"
                NewDataRow.Item("Comment") = ""

                NewDataTable.Rows.Add(NewDataRow) 'add this row to the table
            End If

        Next

        NewDataTable.AcceptChanges() 'validate the updates

        'Now the Data table object contains the data in memory
        Using Transaction As SqlTransaction = _Connection.BeginTransaction() 'IsolationLevel.ReadCommitted)
            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(_Connection, SqlBulkCopyOptions.FireTriggers, Transaction) 'Create a bulk copy object

                bulkCopy.DestinationTableName = "[" & Get_DatabaseSchema() & "].[" & Get_SummaryTableUpdates_TableName() & "]" 'set the destination of the copy to the right table

                AddHandler bulkCopy.SqlRowsCopied, AddressOf OnSqlRowsCopied 'add a callback function to follow the progress of the copy
                bulkCopy.NotifyAfter = CInt(Math.Ceiling(_ValueModifications.Count / 50)) 'callback every 100 records processed
                bulkCopy.BulkCopyTimeout = 15 * 60 'Timeout after 15 min
                Try
                    ' Write unchanged rows from the source to the destination
                    bulkCopy.WriteToServer(NewDataTable) 'DataRowState.Unchanged)
                    Transaction.Commit()

                Catch ex ' As Exception
                    Throw New System.Exception("Error filling the table: " & ex.Message)
                End Try
            End Using
        End Using

        'The table has been successfully updated in the database!
        NewDataTable.Dispose() 'Dispose the memory data oject


        _ProgressCallback(66, "Checking results")

        'Check the result
        SQLQuery = "SELECT "
        For i As Integer = 0 To Get_SummaryTable_KeyColumns.Count - 1 'Create one column per Key
            SQLQuery &= Get_SummaryTable_KeyColumns(i) & ","
        Next
        SQLQuery &= "COLUMNAME,OLDVALUE,NEWVALUE,STATUS,COMMENT FROM [" & Get_DatabaseSchema() & "].[" & Get_SummaryTableUpdates_TableName() & "]"
        SQLQuery &= "WHERE [ChangedBy]='" & UserName & "' AND "
        SQLQuery &= "[ChangeDateTime]='" & ChangeDateTime.ToString("yyyy'-'MM'-'dd' 'HH':'mm':'ss'.000'") & "'"

        'Now trigger it
        NbFailedQueries = 0
        Do

            _Command = New SqlCommand(SQLQuery, _Connection)
            _Adapter = New SqlDataAdapter(_Command)
            _ResultDataSet = New DataSet

            ex = Nothing
            Try
                _Adapter.Fill(_ResultDataSet, "ResultTable") 'Try to run the query, and update the number of rows
            Catch ex
                'Dispose these object, we will recreate new instances in the next loop
                _ResultDataSet.Dispose()
                _Adapter.Dispose()
                _Command.Dispose()

                NbFailedQueries += 1
                'AddMessageToStack(New LogMessage("Warning: Unable to xxxx (attempt " & NbFailedQueries.ToString(Globalization.CultureInfo.InvariantCulture) & ") : " & ex.Message, LogLevel.LogWarning))

                If NbFailedQueries = 5 Then 'after 5 failed trials, we give up for now
                    'AddMessageToStack(New LogMessage("Error: Unable to xxxx (will retry in 5 min) : " & ex.Message, LogLevel.LogError))
                    Throw New System.Exception("Error while running the check result query")
                Else
                    Thread.Sleep(NbFailedQueries * 1000) 'if we tried less than 5 times, pause the thread for an increasing amount of time until next trial
                End If

                'Re-check the connection to the database server for next loop
                If CheckSQLConnectionAndReconnect(_Connection, 5) = False Then
                    Throw New System.Exception("Error: Unable to connect to the database")
                End If

            End Try

        Loop While Not (ex Is Nothing)

        ReDim KeyValues(Get_SummaryTable_KeyColumns.Count - 1)

        If _ResultDataSet.Tables("ResultTable").Rows.Count = 0 Then
            Throw New System.Exception("Error: The check result query didn't return anything")
        Else
            Dim CurRowNum As Integer = 0
            For Each ResultRow As DataRow In _ResultDataSet.Tables("ResultTable").Rows

                CurRowNum += 1
                If CurRowNum Mod (Math.Ceiling(_ValueModifications.Count / 50)) = 0 Then 'split the progress bar updates in 50 steps
                    _ProgressCallback(66 + CInt(33 * (CurRowNum / _ValueModifications.Count)), "Checking results")
                End If

                Dim KeyValuesConcatStr As String = ""
                For i As Integer = 0 To Get_SummaryTable_KeyColumns.Count - 1 'Read the values of the Key fields
                    KeyValues(i) = CStr(ResultRow.Item(Get_SummaryTable_KeyColumns(i)))
                    KeyValuesConcatStr &= KeyValues(i)
                    If i < Get_SummaryTable_KeyColumns.Count - 1 Then KeyValuesConcatStr &= "/"
                Next

                ColumnName = CStr(ResultRow.Item("COLUMNAME"))

                Select Case ResultRow.Item("STATUS").ToString
                    Case "PROCESSED OK"
                        'Recorsd processed ok-> update the record in memory
                        ModifyValue(KeyValues, ColumnName, ResultRow.Item("NEWVALUE"))
                        NbChangesProcessed += 1

                    Case "WAITING"
                        'These records are still waiting! (should not happen!)
                        ErrorsFound.Add(KeyValuesConcatStr & "/" & ColumnName & ": Record still in WAITING status!! Please contact your BPX")

                    Case "INVALID TARGET COLUMN"
                        'The requested column is not modifiable or doesn't exist (should not happen!)
                        ErrorsFound.Add(KeyValuesConcatStr & "/" & ColumnName & ": The requested column is not modifiable or doesn't exist. Please contact your BPX")

                    Case "RECORD NOT FOUND"
                        'The record was not found for the given Keys and Report date (should not happen!)
                        ErrorsFound.Add(KeyValuesConcatStr & "/" & ColumnName & ": The record was not found. Please contact your BPX")

                    Case "NEW VALUE HAS INVALID FORMAT"
                        'The New value doesn't have the right numeric or date format
                        ErrorsFound.Add(KeyValuesConcatStr & "/" & ColumnName & ": The new value is not in the correct format")

                    Case "OLD VALUE HAS INVALID FORMAT"
                        'The old value doesn't have the right numeric or date format
                        ErrorsFound.Add(KeyValuesConcatStr & "/" & ColumnName & ": The old value is not in the correct format. Please contact your BPX")

                    Case "NEW VALUE LENGHT EXCEEDS MAX SIZE"
                        'The lenght of the new value exceeds max size
                        ErrorsFound.Add(KeyValuesConcatStr & "/" & ColumnName & ": The lenght of the new value exceeds max size. Please contact your BPX")

                    Case "OLD VALUE LENGHT EXCEEDS MAX SIZE"
                        'The lenght of the old value exceeds max size
                        ErrorsFound.Add(KeyValuesConcatStr & "/" & ColumnName & ": The lenght of the old value exceeds max size. Please contact your BPX")

                    Case "ERROR"
                        'Records processed with errors (the value after the update is different than what we expect!! (should not happen!)
                        ErrorsFound.Add(KeyValuesConcatStr & "/" & ColumnName & ": Value not updated properly. Please contact your BPX")

                    Case "INVALID ROOT CAUSE"
                        'The provided root cause is not in the list of valid root causes
                        ErrorsFound.Add(KeyValuesConcatStr & "/" & ColumnName & ": Invalid Root cause")

                    Case "CONFLICT FOUND"
                        'The current database value is not the one that we expect. Data has been modified by another user in the meantime.


                        If UserChoiceToAllConflicts = "" Then
                            'the user didn't make a decision yet

                            'create a new conflic form with the details of the problem
                            ConflictForm = New Form_Conflict(
                                                        KeyValues,
                                                        ColumnName,
                                                        "From '" & CStr(ResultRow.Item("OLDVALUE")) & "' to '" & CStr(ResultRow.Item("NEWVALUE")) & "'",
                                                        "Changed to '" & CStr(ResultRow.Item("COMMENT")) & "'"
                                                        )
                            ConflictForm.ShowDialog() 'display the form

                            'now that it is closed, take note of the user decision
                            UserChoice = ConflictForm.UserDecision
                            If ConflictForm.RememberChoice Then UserChoiceToAllConflicts = UserChoice

                            ConflictForm.Dispose() 'dispose the form
                        Else
                            'the user already made a decision for all remaining conflicts
                            UserChoice = UserChoiceToAllConflicts
                        End If

                        ReportRowNumber = _ValueModifications.Find(Function(x) x.EqualsKeys(New SummaryTable_ValueModification(KeyValues))).ExcelReportRow

                        If UserChoice = "OVERWRITE" Then
                            NeedReprocess = True 'we have at least one conflict to reprocess
                            _ValueModifications.Add(New SummaryTable_ValueModification(KeyValues,
                                                                                  ColumnName,
                                                                                  CStr(ResultRow.Item("COMMENT")),
                                                                                  CStr(ResultRow.Item("NEWVALUE")),
                                                                                  GetColumnDataType(ColumnName),
                                                                                  ReportRowNumber))

                        Else '"ABANDON
                            'Update to the database value
                            ModifyValue(KeyValues, ColumnName, ResultRow.Item("COMMENT"))

                            'and update the list of modifications to make to the local displayed report
                            Dim ModificationItem As SummaryTable_ValueModification
                            'check if there is already a modification record for that Keys combination & column
                            ModificationItem = AbandonnedConflictsModifications.Find(Function(x) x.EqualsKeysAndFieldName(New SummaryTable_ValueModification(KeyValues, ColumnName)))
                            If ModificationItem Is Nothing Then
                                'if not, create a new item
                                AbandonnedConflictsModifications.Add(New SummaryTable_ValueModification(
                                                                                      KeyValues,
                                                                                      ColumnName,
                                                                                      "",
                                                                                      CStr(ResultRow.Item("COMMENT")),
                                                                                      GetColumnDataType(ColumnName),
                                                                                      ReportRowNumber))
                            Else
                                'otherwise, update the existing one. This is very unlikely to happen! this can happen only if we had several conflicts in series
                                ModificationItem.NewValue = CStr(ResultRow.Item("COMMENT"))
                            End If

                        End If

                    Case Else
                        ErrorsFound.Add(KeyValuesConcatStr & "/" & ColumnName & ": Unknown error. Please contact your BPX")
                End Select

            Next
        End If

        _ResultDataSet.Dispose()
        _Adapter.Dispose()
        _Command.Dispose()

        If ErrorsFound.Count > 0 Then
            Dim FormErrors As New Form_ErrorsDisplay(ErrorsFound)
            FormErrors.ShowDialog()
            FormErrors.Dispose()
        End If

        If NeedReprocess = True Then
            ChangeDateTime = Now
            ChangeDateTime = New DateTime(ChangeDateTime.Year, ChangeDateTime.Month, ChangeDateTime.Day, ChangeDateTime.Hour, ChangeDateTime.Minute, ChangeDateTime.Second, ChangeDateTime.Kind)
            SendChangesToDB(ChangeDateTime, ReportDate, UserName, ProgressCallback)
        End If

        Return NbChangesProcessed

    End Function

    Protected Sub OnSqlRowsCopied(ByVal sender As Object, ByVal args As SqlRowsCopiedEventArgs)

        If _ValueModifications.Count <> 0 Then
            _CopyToDBPercentCompleted = CSng(args.RowsCopied / _ValueModifications.Count)
            _ProgressCallback(33 + CInt(_CopyToDBPercentCompleted * 100 / 3), "Sending modifications to database")
        Else
            _CopyToDBPercentCompleted = 0
        End If
    End Sub

    Protected Overridable Function Get_ReadDetailedProjectionData_QueryString(ReportDate As Date, KeyValues As String()) As String
        Dim SQLQuery As String

        ' Create the standard SQL query to read data from database. Each DatabaseAdapter can override this function if needed
        SQLQuery = "SELECT * FROM [" & Get_DatabaseSchema() & "].[" & Get_DetailsTable_Name() & "] WHERE "
        For i As Integer = 0 To Get_SummaryTable_KeyColumns.Count - 1
            SQLQuery &= Get_SummaryTable_KeyColumns(i) & " = '" & KeyValues(i) & "' "
            If i < Get_SummaryTable_KeyColumns.Count - 1 Then SQLQuery &= "AND "
        Next
        SQLQuery &= " AND ReportDate = '" & ReportDate.ToString("yyyy'-'MM'-'dd") & "' "
        SQLQuery &= Get_DetailledView_Optional_OrderBySQLClause()
        SQLQuery &= ";"
        Return SQLQuery

    End Function

    Public Function ReadDetailedProjectionData(ReportDate As Date, KeyValues As String()) As Integer
        Dim ex As Exception

        Dim NbFailedQueries As Integer
        Dim ReportNbRow As Integer

        If Not (_ResultDataSet Is Nothing) Then _ResultDataSet.Dispose()
        If Not (_Adapter Is Nothing) Then _Adapter.Dispose()
        If Not (_Command Is Nothing) Then _Command.Dispose()

        Dim SQLQuery As String = Get_ReadDetailedProjectionData_QueryString(ReportDate, KeyValues)

        'Now trigger it
        NbFailedQueries = 0
        Do

            _Command = New SqlCommand(SQLQuery, _Connection)
            _Adapter = New SqlDataAdapter(_Command)
            _ResultDataSet = New DataSet

            ex = Nothing
            Try
                ReportNbRow = _Adapter.Fill(_ResultDataSet, "ResultTable") 'Try to run the query, and update the number of rows
            Catch ex
                'Dispose these object, we will recreate new instances in the next loop
                _ResultDataSet.Dispose()
                _Adapter.Dispose()
                _Command.Dispose()

                NbFailedQueries += 1
                'AddMessageToStack(New LogMessage("Warning: Unable to xxxx (attempt " & NbFailedQueries.ToString(Globalization.CultureInfo.InvariantCulture) & ") : " & ex.Message, LogLevel.LogWarning))

                If NbFailedQueries = 5 Then 'after 5 failed trials, we give up for now
                    'AddMessageToStack(New LogMessage("Error: Unable to xxxx (will retry in 5 min) : " & ex.Message, LogLevel.LogError))
                    Throw New System.Exception("Error while running the data read query from the database")
                Else
                    Thread.Sleep(NbFailedQueries * 1000) 'if we tried less than 5 times, pause the thread for an increasing amount of time until next trial
                End If

                'Re-check the connection to the database server for next loop
                If CheckSQLConnectionAndReconnect(_Connection, 5) = False Then
                    Throw New System.Exception("Error: Unable to connect to the database")
                End If

            End Try

        Loop While Not (ex Is Nothing)

        DetailsTable_Dataset = _ResultDataSet.Copy()
        _ResultDataSet.Dispose()
        _Adapter.Dispose()
        _Command.Dispose()

        Return ReportNbRow

    End Function

    Public Function ReadSingleSummaryTableRowData(ReportDate As Date, KeyValues As String()) As Boolean
        Dim ex As Exception
        Dim SQLQuery As String
        Dim NbFailedQueries As Integer
        Dim ReportNbRow As Integer

        If Not (_ResultDataSet Is Nothing) Then _ResultDataSet.Dispose()
        If Not (_Adapter Is Nothing) Then _Adapter.Dispose()
        If Not (_Command Is Nothing) Then _Command.Dispose()

        SQLQuery = "SELECT * FROM [" & Get_DatabaseSchema() & "].[" & Get_SummaryTable_Name() & "] WHERE "
        For i As Integer = 0 To Get_SummaryTable_KeyColumns.Count - 1
            SQLQuery &= Get_SummaryTable_KeyColumns(i) & " = '" & KeyValues(i) & "' "
            If i < Get_SummaryTable_KeyColumns.Count - 1 Then SQLQuery &= "AND "
        Next
        SQLQuery &= " AND ReportDate = '" & ReportDate.ToString("yyyy'-'MM'-'dd") & "';"

        'Now trigger it
        NbFailedQueries = 0
        Do

            _Command = New SqlCommand(SQLQuery, _Connection)
            _Adapter = New SqlDataAdapter(_Command)
            _ResultDataSet = New DataSet

            ex = Nothing
            Try
                ReportNbRow = _Adapter.Fill(_ResultDataSet, "ResultTable") 'Try to run the query, and update the number of rows
            Catch ex
                'Dispose these object, we will recreate new instances in the next loop
                _ResultDataSet.Dispose()
                _Adapter.Dispose()
                _Command.Dispose()

                NbFailedQueries += 1
                'AddMessageToStack(New LogMessage("Warning: Unable to xxxx (attempt " & NbFailedQueries.ToString(Globalization.CultureInfo.InvariantCulture) & ") : " & ex.Message, LogLevel.LogWarning))

                If NbFailedQueries = 5 Then 'after 5 failed trials, we give up for now
                    'AddMessageToStack(New LogMessage("Error: Unable to xxxx (will retry in 5 min) : " & ex.Message, LogLevel.LogError))
                    Throw New System.Exception("Error while running the data read query from the database")
                Else
                    Thread.Sleep(NbFailedQueries * 1000) 'if we tried less than 5 times, pause the thread for an increasing amount of time until next trial
                End If

                'Re-check the connection to the database server for next loop
                If CheckSQLConnectionAndReconnect(_Connection, 5) = False Then
                    Throw New System.Exception("Error: Unable to connect to the database")
                End If

            End Try

        Loop While Not (ex Is Nothing)

        _CurrentSummaryTableRow_Dataset = _ResultDataSet.Copy()
        _ResultDataSet.Dispose()
        _Adapter.Dispose()
        _Command.Dispose()

        If ReportNbRow = 1 Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Function Read_DetailsTable_Availabe_Dates(KeyValues As String()) As Boolean
        Dim ex As Exception
        Dim SQLQuery As String
        Dim NbFailedQueries As Integer
        Dim ReportNbRow As Integer

        If Not (_ResultDataSet Is Nothing) Then _ResultDataSet.Dispose()
        If Not (_Adapter Is Nothing) Then _Adapter.Dispose()
        If Not (_Command Is Nothing) Then _Command.Dispose()

        SQLQuery = "SELECT DISTINCT(ReportDate) FROM [" & Get_DatabaseSchema() & "].[" & Get_DetailsTable_Name() & "] WHERE "
        For i As Integer = 0 To Get_SummaryTable_KeyColumns.Count - 1
            SQLQuery &= Get_SummaryTable_KeyColumns(i) & " = '" & KeyValues(i) & "' "
            If i < Get_SummaryTable_KeyColumns.Count - 1 Then SQLQuery &= "AND "
        Next
        SQLQuery &= ";"

        'Now trigger it
        NbFailedQueries = 0
        Do

            _Command = New SqlCommand(SQLQuery, _Connection)
            _Adapter = New SqlDataAdapter(_Command)
            _ResultDataSet = New DataSet

            ex = Nothing
            Try
                ReportNbRow = _Adapter.Fill(_ResultDataSet, "ResultTable") 'Try to run the query, and update the number of rows
            Catch ex
                'Dispose these object, we will recreate new instances in the next loop
                _ResultDataSet.Dispose()
                _Adapter.Dispose()
                _Command.Dispose()

                NbFailedQueries += 1
                'AddMessageToStack(New LogMessage("Warning: Unable to xxxx (attempt " & NbFailedQueries.ToString(Globalization.CultureInfo.InvariantCulture) & ") : " & ex.Message, LogLevel.LogWarning))

                If NbFailedQueries = 5 Then 'after 5 failed trials, we give up for now
                    'AddMessageToStack(New LogMessage("Error: Unable to xxxx (will retry in 5 min) : " & ex.Message, LogLevel.LogError))
                    Throw New System.Exception("Error while checking the avaialbe StockReqLists")
                Else
                    Thread.Sleep(NbFailedQueries * 1000) 'if we tried less than 5 times, pause the thread for an increasing amount of time until next trial
                End If

                'Re-check the connection to the database server for next loop
                If CheckSQLConnectionAndReconnect(_Connection, 5) = False Then
                    Throw New System.Exception("Error: Unable to connect to the database")
                End If

            End Try

        Loop While Not (ex Is Nothing)

        DetailsTable_AvailableDates_Dataset = _ResultDataSet.Copy()
        _ResultDataSet.Dispose()
        _Adapter.Dispose()
        _Command.Dispose()

        Return True


    End Function

    Public Function Read_ChangeLog(ReportDate As Date, KeyValues As String(), ColumnName As String) As Boolean
        Dim ex As Exception
        Dim SQLQuery As String
        Dim NbFailedQueries As Integer
        Dim ReportNbRow As Integer

        If Not (_ResultDataSet Is Nothing) Then _ResultDataSet.Dispose()
        If Not (_Adapter Is Nothing) Then _Adapter.Dispose()
        If Not (_Command Is Nothing) Then _Command.Dispose()

        SQLQuery = "SELECT TOP(50) ChangeDateTime, ReportDate, ChangedBy, OldValue, NewValue, Status FROM [" & Get_DatabaseSchema() & "].[" & Get_SummaryTableUpdates_TableName() & "] WHERE "
        For i As Integer = 0 To Get_SummaryTable_KeyColumns.Count - 1
            SQLQuery &= Get_SummaryTable_KeyColumns(i) & " = '" & KeyValues(i) & "' "
            SQLQuery &= "AND "
        Next
        'SQLQuery &= " ReportDate = '" & ReportDate.ToString("yyyy'-'MM'-'dd") & "'" 'Should not filter on ReportDate as modifications are copied from day to day...
        SQLQuery &= " ColumName = '" & ColumnName & "'"
        SQLQuery &= " Order by ChangeDateTime DESC;"

        'Now trigger it
        NbFailedQueries = 0
        Do

            _Command = New SqlCommand(SQLQuery, _Connection)
            _Adapter = New SqlDataAdapter(_Command)
            _ResultDataSet = New DataSet

            ex = Nothing
            Try
                ReportNbRow = _Adapter.Fill(_ResultDataSet, "ResultTable") 'Try to run the query, and update the number of rows
            Catch ex
                'Dispose these object, we will recreate new instances in the next loop
                _ResultDataSet.Dispose()
                _Adapter.Dispose()
                _Command.Dispose()

                NbFailedQueries += 1
                'AddMessageToStack(New LogMessage("Warning: Unable to xxxx (attempt " & NbFailedQueries.ToString(Globalization.CultureInfo.InvariantCulture) & ") : " & ex.Message, LogLevel.LogWarning))

                If NbFailedQueries = 5 Then 'after 5 failed trials, we give up for now
                    'AddMessageToStack(New LogMessage("Error: Unable to xxxx (will retry in 5 min) : " & ex.Message, LogLevel.LogError))
                    Throw New System.Exception("Error while checking the avaialbe StockReqLists")
                Else
                    Thread.Sleep(NbFailedQueries * 1000) 'if we tried less than 5 times, pause the thread for an increasing amount of time until next trial
                End If

                'Re-check the connection to the database server for next loop
                If CheckSQLConnectionAndReconnect(_Connection, 5) = False Then
                    Throw New System.Exception("Error: Unable to connect to the database")
                End If

            End Try

        Loop While Not (ex Is Nothing)

        _ChangeLog_Dataset = _ResultDataSet.Copy()
        _ResultDataSet.Dispose()
        _Adapter.Dispose()
        _Command.Dispose()

        Return True


    End Function


#Region "IDisposable Support"
    Private disposedValue As Boolean ' Pour détecter les appels redondants

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: supprimer l'état managé (objets managés).
            End If

            If Not (_ResultDataSet Is Nothing) Then
                Try
                    _ResultDataSet.Dispose()
                Catch
                End Try
            End If
            If Not (_Adapter Is Nothing) Then
                Try
                    _Adapter.Dispose()
                Catch
                End Try
            End If
            If Not (_Connection Is Nothing) Then
                Try
                    _Connection.Close()
                Catch
                End Try
                Try
                    _Connection.Dispose()
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



Public Class SummaryTable_ValueModification

    Public Property KeyValues As String()
    Public Property FieldName As String
    Public Property NewValue As Object
    Public Property OldValue As Object
    Public Property DataType As String
    Public Property ExcelReportRow As Integer

    Public Property Processed As Boolean = False

    Sub New(KeyValues As String(), FieldName As String, OldValue As Object, NewValue As Object, DataType As String, ExcelReportRow As Integer)

        _KeyValues = DirectCast(KeyValues.Clone(), String())
        _FieldName = FieldName
        _NewValue = NewValue
        _OldValue = OldValue
        _DataType = DataType
        _ExcelReportRow = ExcelReportRow

    End Sub

    Sub New(KeyValues As String(), FieldName As String) 'This conctructor is to be used only for the .Find function, searching for a record that has the same Keys and FieldName
        _KeyValues = KeyValues
        _FieldName = FieldName

    End Sub

    Sub New(KeyValues As String()) 'This conctructor is to be used only for the .Find function, searching for a record that has the same Keys (whatever FieldName)
        _KeyValues = KeyValues

    End Sub

    Public Function EqualsKeys(otherMyObject As SummaryTable_ValueModification) As Boolean
        Dim EqualFlag As Boolean = True

        For i As Integer = 0 To UBound(otherMyObject.KeyValues)
            If _KeyValues(i) <> otherMyObject.KeyValues(i) Then
                EqualFlag = False
                Exit For
            End If
        Next
        Return EqualFlag

    End Function

    Public Function EqualsKeysAndFieldName(otherMyObject As SummaryTable_ValueModification) As Boolean
        Dim EqualFlag As Boolean = True

        If _FieldName <> otherMyObject.FieldName Then
            EqualFlag = False
        Else
            For i As Integer = 0 To UBound(otherMyObject.KeyValues)
                If _KeyValues(i) <> otherMyObject.KeyValues(i) Then
                    EqualFlag = False
                    Exit For
                End If
            Next
        End If

        Return EqualFlag

    End Function

End Class