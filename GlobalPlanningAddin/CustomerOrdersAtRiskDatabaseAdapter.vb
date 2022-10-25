Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Data

Public Class CustomerOrdersAtRiskDatabaseAdapter : Inherits DatabaseAdapterBase

    Public Sub New(TemplateID As String)
        MyBase.New(TemplateID)
    End Sub

    Protected Overrides Function Get_ConnectionString() As String
        Dim ConnectionString_Production As String = "Server=USSANTDB02P\NA_SUPPLY_CHAIN;Database=SKUAlerts;UID=GlobalPlanningAddinUser;PWD=iojrgRGRE**$8421;"
        Dim ConnectionString_Test As String = "Server=USSANTDB01T\NA_SUPPLY_CHAIN;Database=SKUAlerts;UID=GlobalPlanningAddinUser;PWD=iojrgRGRE**$8421;"

        Select Case Globals.Current_Plugin_System.ID
            Case 0
                Return ConnectionString_Production
            Case 1
                Return ConnectionString_Test
            Case Else
                Return ConnectionString_Production
        End Select
    End Function

    Protected Overrides Function Get_DatabaseSchema() As String
        Return "dbo"
    End Function

    Protected Overrides Function Get_SummaryTable_Name() As String
        Return "CustOrdersAtRisk_VIEW"
    End Function

    Public Overrides Function SummaryTable_DefaultSortColumns() As List(Of SortField)
        Return _SortFields
    End Function

    Private ReadOnly _SortFields As New List(Of SortField)(
        {New SortField("Material", SortField.SortOrders.Ascending),
        New SortField("sortkeyAVST", SortField.SortOrders.Ascending)}
        )


End Class
