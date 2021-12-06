Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Data

Public Class GRUTDatabaseAdapter : Inherits DatabaseAdapterBase

    Protected Overrides Function Get_ConnectionString() As String
        Return "Server=USSANTDB02P\NA_SUPPLY_CHAIN;Database=GlobalPlanning;UID=GlobalPlanningAddinUser;PWD=iojrgRGRE**$8421;"
    End Function

    Protected Overrides Function Get_DatabaseSchema() As String
        Return "risk"
    End Function

    Protected Overrides Function Get_SummaryTable_Name() As String
        Return "GRUT"
    End Function

    Protected Overrides Function Get_SummaryTableUpdates_TableName() As String
        Return "GRUT_UPDATES"
    End Function

    Protected Overrides Function Get_DetailsTable_Name() As String
        Return "GRUT_PROJECTION"
    End Function

    Protected Overrides Function Get_DetailledView_Optional_OrderBySQLClause() As String
        Return "ORDER BY STARTDATE ASC"
    End Function

    Public Overrides Function Get_SummaryTable_ListOfModifiableColumns() As String()
        Return {
                "Next_Dely_Date",
                "Next_Dely_Qty",
                "Root_Cause",
                "RCA_Comment",
                "U_Comment",
                "Fa_Comment",
                "GDO_Comment"
                }
    End Function

    'List all the modifiable columns with numeric type
    Public Overrides Function Get_SummaryTable_ListOfNumericColumns() As String()
        Return {
                "Next_Dely_Qty" 'double
                }
    End Function

    'List all modifiable columns with Date type
    Public Overrides Function Get_SummaryTable_ListOfDateColumns() As String()
        Return {
                "Next_Dely_Date" 'date
                }
    End Function

    Public Overrides Function Get_SummaryTable_KeyColumns() As String() 'Key columns (excluding the ReportDate that must be key as well). These columns should also be Keys in the detailed table
        Return { 'the order should match the order of the columns in the XXX_UPDATE table in SQL server as the BulkCopy fails if columns are not in the right order
            "Item",
            "Loc"}
    End Function

    Public Overrides Function Get_SummaryTable_Columns() As String()
        Return {
                "EXTRACT_DATETIME",
                "ReportDate",
                "Loc",
                "Item",
                "Item_Description",
                "SKU",
                "Cur_Source",
                "Cur_FromSource_Transportation_LeadTime",
                "Cur_FromFactory_Transportation_LeadTime",
                "Cur_Factory_Internal_Reaction_Time",
                "Cur_FromFactory_Total_Replanishment_LeadTime",
                "Cur_DRPCovDur",
                "Factory_ID",
                "Article_SubType",
                "Factory",
                "Division",
                "House",
                "Brand",
                "Line",
                "Product_Segment",
                "Planner_Code",
                "Planner_Name",
                "SKU_Planner_Code",
                "SKU_Planner_Name",
                "Free_Stock",
                "OnHand_QC_Qty",
                "OnHand_Qty",
                "Todays_orders",
                "ABC_Item",
                "ABC_SKU",
                "DOC",
                "DOC_Change",
                "SS_VS_OH_today",
                "PlanSwitch",
                "Quota_From",
                "Quota_To",
                "Quota_Type",
                "New_Item",
                "Old_Item",
                "ICW",
                "ABC_SKU_Rank",
                "UnderForecast_switch",
                "UnderForecast_rules_results",
                "UnderForecast_Cond_1",
                "UnderForecast_Cond_2",
                "UnderForecast_Cond_3",
                "UnderForecast_Cond_4",
                "Current_Month_Orders",
                "Next_Month_Orders",
                "Total_Orders",
                "Current_Month_forecast",
                "Current_Month_forecast_Var",
                "Next_Month_forecast",
                "End_of_Demand_Date",
                "Ex_Factory_Date",
                "Factory_Disc_Date",
                "Stock_Out_Date",
                "Oracle_Code",
                "SSCov",
                "MinSS",
                "BISS",
                "Factory_OH",
                "Factory_QC",
                "FCST_PERF_M1",
                "FCST_PERF_M2",
                "FCST_PERF_M3",
                "SALES_VS_CM_FCST",
                "ServiceRiskFactor",
                "MinSSOH_Within_Leadtime",
                "Total_IntransitIn_Qty",
                "First_IntransitIn_Qty",
                "First_IntransitIn_Arrival_Date",
                "First_Committed_Qty",
                "First_Committed_Date",
                "First_Recship_Qty",
                "First_Recship_Date",
                "Next_Dely_Qty",
                "Next_Dely_Date",
                "Root_Cause",
                "RCA_Comment",
                "U_Comment",
                "Fa_Comment",
                "GDO_Comment",
                "GDO_Reviewed",
                "REVIEWED_DATE",
                "Risk_on_order_qty_D",
                "Risk_on_order_qty_D1",
                "Risk_on_order_qty_D2",
                "Risk_on_order_qty_D3",
                "Risk_on_order_qty_D4",
                "Risk_on_order_qty_D5",
                "Risk_on_order_qty_D6",
                "Risk_on_order_qty_7days",
                "Total_Indep_Risk_Within_Leadtime",
                "Local_Indep_Risk_Within_Leadtime",
                "Local_Indep_Risk_Late_For_Replanishment",
                "Local_Indep_Risk_Within_Replanishment_period",
                "Local_Indep_Risk_After_Leadtime",
                "Total_Risk_Within_Leadtime",
                "Total_Risk_Late_For_Replanishment",
                "Total_Risk_Within_Replanishment_period",
                "Total_Risk_After_Leadtime",
                "Total_Indep_Risk_Late_For_Replanishment",
                "Total_Indep_Risk_Within_Replanishment_period",
                "Total_Indep_Risk_After_Leadtime",
                "Total_Indep_Risk_PAST",
                "Total_Indep_Risk_W01",
                "Total_Indep_Risk_W02",
                "Total_Indep_Risk_W03",
                "Total_Indep_Risk_W04",
                "Total_Indep_Risk_W05",
                "Total_Indep_Risk_W06",
                "Total_Indep_Risk_W07",
                "Total_Indep_Risk_W08",
                "Total_Indep_Risk_W09",
                "Total_Indep_Risk_W10",
                "Total_Indep_Risk_W11",
                "Total_Indep_Risk_W12",
                "Total_Indep_Risk_W13",
                "Total_Indep_Risk_W14",
                "Total_Indep_Risk_W15",
                "Total_Indep_Risk_W16",
                "Total_Indep_Risk_W17",
                "Total_Indep_Risk_W18",
                "Total_Indep_Risk_W19",
                "Total_Indep_Risk_W20",
                "Total_Indep_Risk_W21",
                "Total_Indep_Risk_W22",
                "Total_Indep_Risk_W23",
                "Total_Indep_Risk_W24",
                "Total_Indep_Risk_W25",
                "Total_Indep_Risk_W26",
                "Local_Indep_Risk_Past",
                "Local_Indep_Risk_W01",
                "Local_Indep_Risk_W02",
                "Local_Indep_Risk_W03",
                "Local_Indep_Risk_W04",
                "Local_Indep_Risk_W05",
                "Local_Indep_Risk_W06",
                "Local_Indep_Risk_W07",
                "Local_Indep_Risk_W08",
                "Local_Indep_Risk_W09",
                "Local_Indep_Risk_W10",
                "Local_Indep_Risk_W11",
                "Local_Indep_Risk_W12",
                "Local_Indep_Risk_W13",
                "Local_Indep_Risk_W14",
                "Local_Indep_Risk_W15",
                "Local_Indep_Risk_W16",
                "Local_Indep_Risk_W17",
                "Local_Indep_Risk_W18",
                "Local_Indep_Risk_W19",
                "Local_Indep_Risk_W20",
                "Local_Indep_Risk_W21",
                "Local_Indep_Risk_W22",
                "Local_Indep_Risk_W23",
                "Local_Indep_Risk_W24",
                "Local_Indep_Risk_W25",
                "Local_Indep_Risk_W26",
                "Total_Risk_Past",
                "Total_Risk_W01",
                "Total_Risk_W02",
                "Total_Risk_W03",
                "Total_Risk_W04",
                "Total_Risk_W05",
                "Total_Risk_W06",
                "Total_Risk_W07",
                "Total_Risk_W08",
                "Total_Risk_W09",
                "Total_Risk_W10",
                "Total_Risk_W11",
                "Total_Risk_W12",
                "Total_Risk_W13",
                "Total_Risk_W14",
                "Total_Risk_W15",
                "Total_Risk_W16",
                "Total_Risk_W17",
                "Total_Risk_W18",
                "Total_Risk_W19",
                "Total_Risk_W20",
                "Total_Risk_W21",
                "Total_Risk_W22",
                "Total_Risk_W23",
                "Total_Risk_W24",
                "Total_Risk_W25",
                "Total_Risk_W26",
                "MinSSOH_W01",
                "MinSSOH_W02",
                "MinSSOH_W03",
                "MinSSOH_W04",
                "MinSSOH_W05",
                "MinSSOH_W06",
                "MinSSOH_W07",
                "MinSSOH_W08",
                "MinSSOH_W09",
                "MinSSOH_W10",
                "MinSSOH_W11",
                "MinSSOH_W12",
                "MinSSOH_W13",
                "MinSSOH_W14",
                "MinSSOH_W15",
                "MinSSOH_W16",
                "MinSSOH_W17",
                "MinSSOH_W18",
                "MinSSOH_W19",
                "MinSSOH_W20",
                "MinSSOH_W21",
                "MinSSOH_W22",
                "MinSSOH_W23",
                "MinSSOH_W24",
                "MinSSOH_W25",
                "MinSSOH_W26"
                }
    End Function

    Public Overrides Function Get_SummaryTable_SQLQueryForField(DatabaseColName As String) As String

        Select Case DatabaseColName
            Case "First_IntransitIn_Arrival_Date" : Return "CASE WHEN First_IntransitIn_Arrival_Date='1970-01-01' THEN NULL ELSE FORMAT(First_IntransitIn_Arrival_Date,'yyyy-MM-dd') END as First_IntransitIn_Arrival_Date" '"First_IntransitIn_Arrival_Date" '**
            Case "First_Committed_Date" : Return "CASE WHEN First_Committed_Date='1970-01-01' THEN NULL ELSE FORMAT(First_Committed_Date,'yyyy-MM-dd') END as First_Committed_Date" '"First_Committed_Date" '**
            Case "First_Recship_Date" : Return "CASE WHEN First_Recship_Date='1970-01-01' THEN NULL ELSE FORMAT(First_Recship_Date,'yyyy-MM-dd') END as First_Recship_Date" '"First_Recship_Date" '**
            Case "GDO_Reviewed" : Return "CASE WHEN GDO_Reviewed='1970-01-01' THEN NULL ELSE FORMAT(GDO_Reviewed,'yyyy-MM-dd') END as GDO_Reviewed" '"GDO_Reviewed" '**
            Case "REVIEWED_DATE" : Return "CASE WHEN REVIEWED_DATE='1970-01-01' THEN NULL ELSE FORMAT(REVIEWED_DATE,'yyyy-MM-dd') END as REVIEWED_DATE" '"REVIEWED_DATE" '**
            Case "End_of_Demand_Date" : Return "CASE WHEN End_of_Demand_Date='1970-01-01' THEN NULL ELSE FORMAT(End_of_Demand_Date,'yyyy-MM-dd') END as End_of_Demand_Date" '"End_of_Demand_Date" '**
            Case "Ex_Factory_Date" : Return "CASE WHEN Ex_Factory_Date='1970-01-01' THEN NULL ELSE FORMAT(Ex_Factory_Date,'yyyy-MM-dd') END as Ex_Factory_Date" '"Ex_Factory_Date" '**
            Case "Factory_Disc_Date" : Return "CASE WHEN Factory_Disc_Date='1970-01-01' THEN NULL ELSE FORMAT(Factory_Disc_Date,'yyyy-MM-dd') END as Factory_Disc_Date" '"Factory_Disc_Date" '**
            Case "Stock_Out_Date" : Return "CASE WHEN Stock_Out_Date='1970-01-01' THEN NULL ELSE FORMAT(Stock_Out_Date,'yyyy-MM-dd') END as Stock_Out_Date" '"Stock_Out_Date" '**
            Case "BISS" : Return "CASE WHEN BISS='1970-01-01' THEN NULL ELSE FORMAT(BISS,'yyyy-MM-dd') END as BISS" '"BISS" '**
            Case "ICW" : Return "CASE WHEN ICW='1970-01-01' THEN NULL ELSE FORMAT(ICW,'yyyy-MM-dd') END as ICW" '"ICW" '**

            Case Else : Return """" & DatabaseColName & """"
        End Select
    End Function

    Public Overrides Function Get_SummaryTable_DefaultSortColumns() As String()
        Return {
            "Total_Indep_Risk_W01"  ' if user mapped this column, the report will be sorted by this column by default
        }
    End Function

    Public Overrides Function Get_DetailedView_Columns() As String()
        Return {
                "DAY+",
                "WEEK+",
                "STARTDATE",
                "STARTDATE_SOURCE",
                "ALTSUPSDCONSTRPROJOH",
                "SS",
                "SS_VS_OH",
                "STEP0_StockBegDay",
                "STEP1_ConsOtherDmd",
                "STEP2_Cut_Ind_Dmd",
                "STEP2_Cut_Other_Dmd",
                "STEP3_TotalSoldOfDay",
                "STEP4_StockEndDay",
                "CUT_VS_INDEPDMD_DOWNSTREAM_1",
                "CUT_VS_INDEPDMD_DOWNSTREAM_2",
                "CUT_VS_INDEPDMD_DOWNSTREAM_3",
                "CUT_VS_INDEPDMD_DOWNSTREAM_4",
                "CUT_VS_INDEPDMD_DOWNSTREAM_5",
                "CUT_VS_INDEPDMD_DOWNSTREAM",
                "ACTUALINTRANSIN",
                "SCHEDRCPTS",
                "COMMITINTRANSIN",
                "RECARRIV",
                "CONSTRPROXYSUPPLY",
                "PROXYRECARRIV",
                "TOTAL_SUPPLY",
                "TOTAL_SUPPLY_ORDERS_ONLY",
                "ADJALLOCTOTFCST",
                "ADJFCSTCUSTORDERS",
                "NONFCSTCUSTORDERS",
                "INDEP_DMD",
                "INDEP_DMD_ORDERS_ONLY",
                "COMMITINTRANSOUT",
                "DEPDMD",
                "RECSHIP",
                "CONSTRPROXYDEMAND",
                "PROXYRECSHIP",
                "OTHER_DMD",
                "OTHER_DMD_ORDERS_ONLY",
                "STEP0_StockBegDay_ORDERS_ONLY",
                "STEP1_ConsOtherDmd_ORDERS_ONLY",
                "STEP2_Cut_Ind_Dmd_ORDERS_ONLY",
                "STEP2_Cut_Other_Dmd_ORDERS_ONLY",
                "STEP3_TotalSoldOfDay_ORDERS_ONLY",
                "STEP4_StockEndDay_ORDERS_ONLY"
                }
    End Function

    Public Overrides Function Get_DetailedView_CurItem_HeaderText() As String()

        Return {
                "Loc", "' | '", "Item_Description"}
    End Function

    Private ReadOnly _DetailedView_InfoDropDown_Items As New List(Of String())({
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


End Class

