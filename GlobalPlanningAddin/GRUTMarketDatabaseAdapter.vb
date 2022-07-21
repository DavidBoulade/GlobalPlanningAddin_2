Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Data

Public Class GRUTMarketDatabaseAdapter : Inherits DatabaseAdapterBase

    Public Sub New(TemplateID As String)
        MyBase.New(TemplateID)
    End Sub

    'Public Overrides ReadOnly Property SummaryTable_KeyColumns As List(Of DatabaseAdapterColumn)
    'Public Overrides ReadOnly Property SummaryTable_ModifiableColumns As List(Of DatabaseAdapterColumn)

    Protected Overrides Function Get_ConnectionString() As String
        Return "Server=USSANTDB02P\NA_SUPPLY_CHAIN;Database=GlobalPlanning;UID=GlobalPlanningAddinUser;PWD=iojrgRGRE**$8421;"
    End Function

    Protected Overrides Function Get_DatabaseSchema() As String
        Return "risk"
    End Function

    Protected Overrides Function Get_SummaryTable_Name() As String
        Return "GRUT_MARKET_VIEW"
    End Function

    'Protected Overrides Function Get_SummaryTable_Alias() As String
    '    Return "GMV"
    'End Function

    Protected Overrides Function Get_SummaryTableUpdates_TableName() As String 'Used to send changes to database
        Return "" 'GRUT_MARKET_UPDATES"
    End Function

    Protected Overrides Function Get_SummaryTableUpdates_ViewName() As String 'Used to read change history
        Return "" '"GRUT_MARKET_UPDATES_VIEW"
    End Function

    Protected Overrides Function Get_DetailsTable_Name() As String
        Return "" '"GRUT_PROJECTION"
    End Function

    Public Overrides Function SummaryTable_DefaultSortColumns() As List(Of SortField)
        Return _SortFields
    End Function

    Private ReadOnly _SortFields As New List(Of SortField)(
        {New SortField("Market_Local_Indep_Risk_6M", SortField.SortOrders.Descending)}
        )

    Public Overrides Function Get_DetailedView_Columns() As String()
        Return Nothing

    End Function

    Public Overrides Function Get_DetailedView_CurItem_HeaderText() As String()

        Return {"Loc", "' | '", "Item_Description"}

    End Function

    Private ReadOnly _DetailedView_InfoDropDown_Items As New List(Of String()) '({
    '        ({"Extraction Date/time", "EXTRACT_DATETIME"}),
    '        ({"Loc", "Loc"}),
    '        ({"Item", "Item"}),
    '        ({"Description", "Item_Description"}),
    '        ({"Source", "Cur_Source"}),
    '        ({"Factory", "Factory_ID", "' ('", "Factory", "')'"}),
    '        ({"Leadtime from Source", "Cur_FromSource_Transportation_LeadTime"}),
    '        ({"Leadtime from Factory", "Cur_FromFactory_Transportation_LeadTime"}),
    '        ({"Internal Factory reaction time", "Cur_Factory_Internal_Reaction_Time"}),
    '        ({"Total Replenishment Leadtime", "Cur_FromFactory_Total_Replanishment_LeadTime"}),
    '        ({"DRP Cov Dur", "Cur_DRPCovDur"}),
    '        ({"Article SubType", "Article_SubType"}),
    '        ({"Division", "Division"}),
    '        ({"House", "House"}),
    '        ({"Brand", "Brand"}),
    '        ({"Line", "Line"}),
    '        ({"Product Segment", "Product_Segment"}),
    '        ({"Planner", "Planner_Code", "' ('", "Planner_Name", "')'"}),
    '        ({"SKU Planner", "SKU_Planner_Code", "' ('", "SKU_Planner_Name", "')'"}),
    '        ({"ABC Item", "ABC_Item"}),
    '        ({"ABC SKU", "ABC_SKU"})
    '        }
    '        )

    Public Overrides Function Get_DetailedView_InfoDropDown_Items() As List(Of String())
        Return _DetailedView_InfoDropDown_Items
    End Function


End Class

