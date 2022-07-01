Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Data

Public Class GRUTMarketDatabaseAdapter : Inherits DatabaseAdapterBase

    Public Sub New()
        'Populate lists of specific columns
        SummaryTable_KeyColumns = _SummaryTableColumns.FindAll(Function(x) x.IsKey = True)
        SummaryTable_ModifiableColumns = _SummaryTableColumns.FindAll(Function(x) x.IsModifiable = True)
    End Sub

    Public Overrides ReadOnly Property SummaryTable_KeyColumns As List(Of DatabaseAdapterColumn)
    Public Overrides ReadOnly Property SummaryTable_ModifiableColumns As List(Of DatabaseAdapterColumn)

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


    Private ReadOnly _SummaryTableColumns As New List(Of DatabaseAdapterColumn)( 'the order of the key columns should match the order of the columns in the XXX_UPDATE table in SQL server as the BulkCopy fails if columns are not in the right order
        {
        New DatabaseAdapterColumn("ReportDate", False, False, "", ""),
        New DatabaseAdapterColumn("Item", True, False, "", ""),
        New DatabaseAdapterColumn("Loc", True, False, "", ""),
        New DatabaseAdapterColumn("DFULoc", True, False, "", "ISNULL(DFULOC,'') AS DFULOC"),
        New DatabaseAdapterColumn("DMDGroup", True, False, "", "ISNULL(DMDGroup,'') AS DMDGroup"),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_M1", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_M2", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_M3", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_M4", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_M5", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_M6", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W01", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W02", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W03", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W04", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W05", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W06", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W07", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W08", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W09", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W10", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W11", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W12", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W13", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W14", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W15", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W16", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W17", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W18", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W19", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W20", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W21", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W22", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W23", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W24", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W25", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_W26", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_1M", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_2M", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_3M", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_4M", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_5M", False, False, "", ""),
        New DatabaseAdapterColumn("Market_Local_Indep_Risk_6M", False, False, "", "")
        })

    Public Overrides ReadOnly Property SummaryTableColumns As List(Of DatabaseAdapterColumn)
        Get
            Return _SummaryTableColumns
        End Get
    End Property


    Public Overrides Function Get_SummaryTable_DefaultSortColumns() As String()
        '"ServiceRiskFactor"  'if user mapped this column, the report will be sorted by this column by default
        Return {
            "Market_Local_Indep_Risk_6M"
            }
    End Function

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

