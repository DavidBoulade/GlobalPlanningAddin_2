Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Data

Public Class GRUTMarketDatabaseAdapter : Inherits DatabaseAdapterBase

    Public Sub New(TemplateID As String)0
        MyBase.New(TemplateID)
    End Sub

    Protected Overrides Function Get_ConnectionString() As String
        Return "Server=USSANTDB02P\NA_SUPPLY_CHAIN;Database=GlobalPlanning;UID=GlobalPlanningAddinUser;PWD=iojrgRGRE**$8421;"
    End Function

    Protected Overrides Function Get_DatabaseSchema() As String
        Return "risk"
    End Function

    Protected Overrides Function Get_SummaryTable_Name() As String
        Return "GRUT_MARKET_VIEW"
    End Function

    Protected Overrides Function Get_SummaryTableUpdates_TableName() As String 'Used to send changes to database
        Return "GRUT_MARKET_UPDATES"
    End Function

    Protected Overrides Function Get_Preliminary_Check_Query(ReportDate As Date) As String
        Return "EXEC [Risk].[List_Pending_GRUT_Factories] @ReportDate ='" & Format(ReportDate, "yyyy-MM-dd") & "'"
    End Function

    Public Overrides Function SummaryTable_DefaultSortColumns() As List(Of SortField)
        Return _SortFields
    End Function

    Private ReadOnly _SortFields As New List(Of SortField)(
        {New SortField("Market_Local_Indep_Risk_6M", SortField.SortOrders.Descending)}
        )


End Class

