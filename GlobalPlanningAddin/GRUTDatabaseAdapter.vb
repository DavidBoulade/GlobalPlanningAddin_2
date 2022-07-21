Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Data

Public Class GRUTDatabaseAdapter : Inherits DatabaseAdapterBase

    Public Sub New(TemplateID As String)
        MyBase.New(TemplateID)
    End Sub

    Protected Overrides Function Get_ConnectionString() As String

        Dim ConnectionString_Production As String = "Server=USSANTDB02P\NA_SUPPLY_CHAIN;Database=GlobalPlanning;UID=GlobalPlanningAddinUser;PWD=iojrgRGRE**$8421;"
        Dim ConnectionString_Test As String = "Server=USSANTDB01T\NA_SUPPLY_CHAIN;Database=GlobalPlanning;UID=GlobalPlanningAddinUser;PWD=iojrgRGRE**$8421;"

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
        Return "risk"
    End Function

    Protected Overrides Function Get_SummaryTable_Name() As String
        Return "GRUT_VIEW"
    End Function

    Protected Overrides Function Get_SummaryTable_Alias() As String
        Return "GV" 'We can define an alias here, could be useful for custom columns alternate SQL definition.
    End Function

    Protected Overrides Function Get_SummaryTableUpdates_TableName() As String 'Used to send changes to database
        Return "GRUT_UPDATES"
    End Function

    Protected Overrides Function Get_SummaryTableUpdates_ViewName() As String 'Used to read change history
        Return "GRUT_UPDATES_VIEW"
    End Function

    Protected Overrides Function Get_Preliminary_Check_Query() As String
        Return "EXEC [Risk].[List_Pending_GRUT_Factories]"
    End Function

    Protected Overrides Function Get_DetailsTable_Name() As String
        Return "GRUT_PROJECTION"
    End Function

    Public Overrides Function SummaryTable_DefaultSortColumns() As List(Of SortField)
        Return _SortFields
    End Function

    Private ReadOnly _SortFields As New List(Of SortField)(
        {New SortField("ServiceRiskFactor", SortField.SortOrders.Descending)}
        )

    Public Overrides Function Get_DetailedView_Columns() As String()
        Return {
                "Loc",
                "WEEK+",
                "STARTDATE",
                "SOURCE",
                "STARTDATE_SOURCE",
                "ALTSUPSDCONSTRPROJOH",
                "SUPSDCONSTRPROJOH", 'Added 14/03/2022
                "SS",
                "SS_VS_OH",
                "CUT_COMMITINTRANSOUT", 'Updated 21/02/2022
                "CUT_ADJFCSTCUSTORDERS", 'Updated 21/02/2022
                "CUT_NONFCSTCUSTORDERS", 'Updated 21/02/2022
                "CUT_ADJALLOCTOTFCST", 'Updated 21/02/2022
                "CUT_DEPDMD", 'Updated 21/02/2022
                "CUT_RECSHIP", 'Updated 21/02/2022
                "CUT_PROXYRECSHIP", 'Updated 21/02/2022
                "CUT_CONSTRPROXYDEMAND", 'Updated 21/02/2022
                "CUT_IND_DMD", 'Updated 21/02/2022
                "CUT_OTHER_DMD", 'Updated 21/02/2022
                "CUT_VS_INDEPDMD_DOWNSTREAM", 'Updated 21/02/2022
                "CUT_VS_INDEPDMD_NML", 'Updated 21/02/2022
                "ACTUALINTRANSIN",
                "SCHEDRCPTS",
                "COMMITINTRANSIN",
                "RECARRIV",
                "PLANARRIV", 'Added 14/03/2022
                "FIRMPLANARRIV", 'Added 14/03/2022
                "CONSTRPROXYSUPPLY",
                "PROXYRECARRIV",
                "ADJALLOCTOTFCST",
                "ADJFCSTCUSTORDERS",
                "NONFCSTCUSTORDERS",
                "COMMITINTRANSOUT",
                "DEPDMD",
                "RECSHIP",
                "PLANSHIP", 'Added 14/03/2022
                "FIRMPLANSHIP", 'Added 14/03/2022
                "CONSTRPROXYDEMAND",
                "PROXYRECSHIP",
                "CUT_COMMITINTRANSOUT_ORDERS_ONLY", 'Updated 21/02/2022
                "CUT_ADJFCSTCUSTORDERS_ORDERS_ONLY", 'Updated 21/02/2022
                "CUT_NONFCSTCUSTORDERS_ORDERS_ONLY", 'Updated 21/02/2022
                "CUT_DEPDMD_ORDERS_ONLY", 'Updated 21/02/2022
                "CUT_PROXYRECSHIP_ORDERS_ONLY", 'Updated 21/02/2022
                "CUT_CONSTRPROXYDEMAND_ORDERS_ONLY", 'Updated 21/02/2022
                "CUT_IND_DMD_ORDERS_ONLY", 'Updated 21/02/2022
                "CUT_OTHER_DMD_ORDERS_ONLY" 'Updated 21/02/2022
                }

    End Function

    Public Overrides Function Get_DetailedView_CurItem_HeaderText() As String()

        Return {
                "Loc", "' | '", "Item_Description"}
    End Function

    Private ReadOnly _DetailedView_InfoDropDown_Items As New List(Of String())({
                ({"Extraction Date/time", "EXTRACT_DATETIME"}),
                ({"Loc", "Loc"}),
                ({"Item", "Item"}),
                ({"Description", "Item_Description"}),
                ({"Source", "Cur_Source"}),
                ({"Factory", "Factory_ID", "' ('", "Factory", "')'"}),
                ({"Leadtime from Source", "Cur_FromSource_Transportation_LeadTime"}),
                ({"Leadtime from Factory", "Cur_FromFactory_Transportation_LeadTime"}),
                ({"Internal Factory reaction time", "Cur_Factory_Internal_Reaction_Time"}),
                ({"Total Replenishment Leadtime", "Cur_FromFactory_Total_Replanishment_LeadTime"}),
                ({"DRP Cov Dur", "Cur_DRPCovDur"}),
                ({"Article SubType", "Article_SubType"}),
                ({"Division", "Division"}),
                ({"House", "House"}),
                ({"Brand", "Brand"}),
                ({"Line", "Line"}),
                ({"Product Segment", "Product_Segment"}),
                ({"Planner", "Planner_Code", "' ('", "Planner_Name", "')'"}),
                ({"SKU Planner", "SKU_Planner_Code", "' ('", "SKU_Planner_Name", "')'"}),
                ({"ABC Item", "ABC_Item"}),
                ({"ABC SKU", "ABC_SKU"})
                })

    Public Overrides Function Get_DetailedView_InfoDropDown_Items() As List(Of String())
        Return _DetailedView_InfoDropDown_Items
    End Function

    Protected Overrides Function Get_ReadDetailedProjectionData_QueryString(ReportDate As Date, KeyValues As String()) As String
        Dim SQLQuery As String

        ' Create the SQL query to read the detailled data from database. This overrides the default one from the DatabaseAdapterBase
        'The KeyValues tables contains the Values of the Key fields for the current Summary table row
        'In the same order as defined in the Get_SummaryTable_KeyColumns function
        'KeyValues(0) -> Item
        'KeyValues(1) -> Loc
        SQLQuery = "SELECT * FROM [" & Get_DatabaseSchema() & "].[" & Get_DetailsTable_Name() & "] WHERE"
        SQLQuery &= " Item" & " = '" & KeyValues(0) & "'"
        SQLQuery &= " AND ReportDate = '" & ReportDate.ToString("yyyy'-'MM'-'dd") & "'"
        SQLQuery &= " ORDER BY (CASE Loc WHEN '" & KeyValues(1) & "' THEN '0' ELSE Loc END) ASC, STARTDATE ASC" 'Show the current Loc first
        SQLQuery &= ";"
        Return SQLQuery

    End Function

    Public Overrides Function Get_DetailledView_ColumnFilter(KeyValues() As String) As List(Of ColumnFilter)
        'The KeyValues tables contains the Values of the Key fields for the current Summary table row
        'In the same order as defined in the Get_SummaryTable_KeyColumns function
        'KeyValues(0) -> Item
        'KeyValues(1) -> Loc

        Dim Columnsfilters As New List(Of ColumnFilter) From {
            New ColumnFilter With {.ColumnNumber = 1, .FilterValue = KeyValues(1)} '1 is "Loc"
            }
        'New ColumnFilter With {.ColumnNumber = 3, .FilterValue = "<=8"} '3 is "Week+"
        Return Columnsfilters
    End Function

End Class

