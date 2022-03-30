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

    Public Overrides Function Get_SummaryTable_ListOfModifiableColumns() As String()
        Return {
                "Next_Dely_Date",
                "Next_Dely_Qty",
                "Root_Cause",
                "RCA_Comment",
                "U_Comment",
                "Fa_Comment",
                "GDO_Comment",
                "UserDefined_1",
                "UserDefined_2",
                "UserDefined_3",
                "UserDefined_4",
                "UserDefined_5",
                "UserDefined_6",
                "UserDefined_7",
                "UserDefined_8",
                "UserDefined_9",
                "UserDefined_10",
                "RiskToReview_Flag"
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
                "Article_Type",
                "Article_SubType",
                "Factory",
                "Division",
                "House",
                "Brand",
                "Line",
                "Product_Segment",
                "Core_category",
                "Market_category",
                "Sub_category",
                "Sp_category",
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
                "RiskToReview_Flag",
                "RiskToReview_Date",
                "First_Supply_Date",
                "First_Supply_Qty",
                "First_Local_OOS_Date",
                "First_Local_Recovery_Date",
                "First_Local_Back_in_SS_Date",
                "Final_Local_Recovery_Date",
                "Final_Local_Back_in_SS_Date",
                "First_OOS_Date",
                "First_Total_Recovery_Date",
                "Final_Total_Recovery_Date",
                "First_Total_Indep_Risk_Date",
                "Final_Total_Indep_Recovery_Date",
                "Item_Total_Indep_Risk_6M",
                "MinSSOH_Within_Leadtime",
                "Total_IntransitIn_Qty",
                "OnHand_Blocked",
                "PRODQTY",
                "PRODDATE",
                "PRODORDERID",
                "DMDTODATE",
                "E_O",
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
                "Total_Indep_Risk_Before_Leadtime_From_Factory",
                "Total_Indep_Risk_After_Leadtime_From_Factory",
                "Total_Indep_Risk_Before_Leadtime_From_Source",
                "Total_Indep_Risk_Within_Replanishment_period_From_Source",
                "Local_Indep_Risk_Before_Leadtime_From_Factory",
                "Local_Indep_Risk_After_Leadtime_From_Factory",
                "Local_Indep_Risk_Before_Leadtime_From_Source",
                "Local_Indep_Risk_Within_Replanishment_period_From_Source",
                "Total_Risk_Before_Leadtime_From_Factory",
                "Total_Risk_After_Leadtime_From_Factory",
                "Total_Risk_Before_Leadtime_From_Source",
                "Total_Risk_Within_Replanishment_period_From_Source",
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
                "Local_Dep_Risk_W01",
                "Local_Dep_Risk_W02",
                "Local_Dep_Risk_W03",
                "Local_Dep_Risk_W04",
                "Local_Dep_Risk_W05",
                "Local_Dep_Risk_W06",
                "Local_Dep_Risk_W07",
                "Local_Dep_Risk_W08",
                "Local_Dep_Risk_W09",
                "Local_Dep_Risk_W10",
                "Local_Dep_Risk_W11",
                "Local_Dep_Risk_W12",
                "Local_Dep_Risk_W13",
                "Local_Dep_Risk_W14",
                "Local_Dep_Risk_W15",
                "Local_Dep_Risk_W16",
                "Local_Dep_Risk_W17",
                "Local_Dep_Risk_W18",
                "Local_Dep_Risk_W19",
                "Local_Dep_Risk_W20",
                "Local_Dep_Risk_W21",
                "Local_Dep_Risk_W22",
                "Local_Dep_Risk_W23",
                "Local_Dep_Risk_W24",
                "Local_Dep_Risk_W25",
                "Local_Dep_Risk_W26",
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
                "MinSSOH_W26",
                "Min_AltSupsdConstrProjOH_W01",
                "Min_AltSupsdConstrProjOH_W02",
                "Min_AltSupsdConstrProjOH_W03",
                "Min_AltSupsdConstrProjOH_W04",
                "Min_AltSupsdConstrProjOH_W05",
                "Min_AltSupsdConstrProjOH_W06",
                "Min_AltSupsdConstrProjOH_W07",
                "Min_AltSupsdConstrProjOH_W08",
                "Min_AltSupsdConstrProjOH_W09",
                "Min_AltSupsdConstrProjOH_W10",
                "Min_AltSupsdConstrProjOH_W11",
                "Min_AltSupsdConstrProjOH_W12",
                "Min_AltSupsdConstrProjOH_W13",
                "Min_AltSupsdConstrProjOH_W14",
                "Min_AltSupsdConstrProjOH_W15",
                "Min_AltSupsdConstrProjOH_W16",
                "Min_AltSupsdConstrProjOH_W17",
                "Min_AltSupsdConstrProjOH_W18",
                "Min_AltSupsdConstrProjOH_W19",
                "Min_AltSupsdConstrProjOH_W20",
                "Min_AltSupsdConstrProjOH_W21",
                "Min_AltSupsdConstrProjOH_W22",
                "Min_AltSupsdConstrProjOH_W23",
                "Min_AltSupsdConstrProjOH_W24",
                "Min_AltSupsdConstrProjOH_W25",
                "Min_AltSupsdConstrProjOH_W26",
                "UserDefined_1",
                "UserDefined_2",
                "UserDefined_3",
                "UserDefined_4",
                "UserDefined_5",
                "UserDefined_6",
                "UserDefined_7",
                "UserDefined_8",
                "UserDefined_9",
                "UserDefined_10",
                "Local_Indep_Risk_M1",
                "Local_Indep_Risk_M2",
                "Local_Indep_Risk_M3",
                "Local_Indep_Risk_M4",
                "Local_Indep_Risk_M5",
                "Local_Indep_Risk_M6",
                "Local_Dep_Risk_M1",
                "Local_Dep_Risk_M2",
                "Local_Dep_Risk_M3",
                "Local_Dep_Risk_M4",
                "Local_Dep_Risk_M5",
                "Local_Dep_Risk_M6",
                "Total_Indep_Risk_M1",
                "Total_Indep_Risk_M2",
                "Total_Indep_Risk_M3",
                "Total_Indep_Risk_M4",
                "Total_Indep_Risk_M5",
                "Total_Indep_Risk_M6",
                "Total_Risk_M1",
                "Total_Risk_M2",
                "Total_Risk_M3",
                "Total_Risk_M4",
                "Total_Risk_M5",
                "Total_Risk_M6",
                "Indep_Dmd_M1",
                "Indep_Dmd_M2",
                "Indep_Dmd_M3",
                "Indep_Dmd_M4",
                "Indep_Dmd_M5",
                "Indep_Dmd_M6",
                "Dep_Dmd_M1",
                "Dep_Dmd_M2",
                "Dep_Dmd_M3",
                "Dep_Dmd_M4",
                "Dep_Dmd_M5",
                "Dep_Dmd_M6",
                "Other_Dmd_M1",
                "Other_Dmd_M2",
                "Other_Dmd_M3",
                "Other_Dmd_M4",
                "Other_Dmd_M5",
                "Other_Dmd_M6",
                "Total_Indep_Risk_1M",
                "Total_Indep_Risk_2M",
                "Total_Indep_Risk_3M",
                "Total_Indep_Risk_4M",
                "Total_Indep_Risk_5M",
                "Total_Indep_Risk_6M",
                "Local_Indep_Risk_1M",
                "Local_Indep_Risk_2M",
                "Local_Indep_Risk_3M",
                "Local_Indep_Risk_4M",
                "Local_Indep_Risk_5M",
                "Local_Indep_Risk_6M",
                "Total_Risk_1M",
                "Total_Risk_2M",
                "Total_Risk_3M",
                "Total_Risk_4M",
                "Total_Risk_5M",
                "Total_Risk_6M"
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
            Case "Next_Dely_Date" : Return "CASE WHEN Next_Dely_Date='1970-01-01' THEN NULL ELSE FORMAT(Next_Dely_Date,'yyyy-MM-dd') END as Next_Dely_Date" '"Next_Dely_Date" '***

            Case "Total_Indep_Risk_1M" : Return "Total_Indep_Risk_PAST + Total_Indep_Risk_M1"
            Case "Total_Indep_Risk_2M" : Return "Total_Indep_Risk_PAST + Total_Indep_Risk_M1 + Total_Indep_Risk_M2"
            Case "Total_Indep_Risk_3M" : Return "Total_Indep_Risk_PAST + Total_Indep_Risk_M1 + Total_Indep_Risk_M2 + Total_Indep_Risk_M3"
            Case "Total_Indep_Risk_4M" : Return "Total_Indep_Risk_PAST + Total_Indep_Risk_M1 + Total_Indep_Risk_M2 + Total_Indep_Risk_M3 + Total_Indep_Risk_M4"
            Case "Total_Indep_Risk_5M" : Return "Total_Indep_Risk_PAST + Total_Indep_Risk_M1 + Total_Indep_Risk_M2 + Total_Indep_Risk_M3 + Total_Indep_Risk_M4 + Total_Indep_Risk_M5"
            Case "Total_Indep_Risk_6M" : Return "Total_Indep_Risk_PAST + Total_Indep_Risk_M1 + Total_Indep_Risk_M2 + Total_Indep_Risk_M3 + Total_Indep_Risk_M4 + Total_Indep_Risk_M5 + Total_Indep_Risk_M6"

            Case "Local_Indep_Risk_1M" : Return "Local_Indep_Risk_M1"
            Case "Local_Indep_Risk_2M" : Return "Local_Indep_Risk_M1 + Local_Indep_Risk_M2"
            Case "Local_Indep_Risk_3M" : Return "Local_Indep_Risk_M1 + Local_Indep_Risk_M2 + Local_Indep_Risk_M3"
            Case "Local_Indep_Risk_4M" : Return "Local_Indep_Risk_M1 + Local_Indep_Risk_M2 + Local_Indep_Risk_M3 + Local_Indep_Risk_M4"
            Case "Local_Indep_Risk_5M" : Return "Local_Indep_Risk_M1 + Local_Indep_Risk_M2 + Local_Indep_Risk_M3 + Local_Indep_Risk_M4 + Local_Indep_Risk_M5"
            Case "Local_Indep_Risk_6M" : Return "Local_Indep_Risk_M1 + Local_Indep_Risk_M2 + Local_Indep_Risk_M3 + Local_Indep_Risk_M4 + Local_Indep_Risk_M5 + Local_Indep_Risk_M6"

            Case "Total_Risk_1M" : Return "Total_Risk_M1"
            Case "Total_Risk_2M" : Return "Total_Risk_M1 + Total_Risk_M2"
            Case "Total_Risk_3M" : Return "Total_Risk_M1 + Total_Risk_M2 + Total_Risk_M3"
            Case "Total_Risk_4M" : Return "Total_Risk_M1 + Total_Risk_M2 + Total_Risk_M3 + Total_Risk_M4"
            Case "Total_Risk_5M" : Return "Total_Risk_M1 + Total_Risk_M2 + Total_Risk_M3 + Total_Risk_M4 + Total_Risk_M5"
            Case "Total_Risk_6M" : Return "Total_Risk_M1 + Total_Risk_M2 + Total_Risk_M3 + Total_Risk_M4 + Total_Risk_M5 + Total_Risk_M6"

            Case "PRODDATE" : Return "CASE WHEN PRODDATE='1970-01-01' THEN NULL ELSE FORMAT(PRODDATE,'yyyy-MM-dd') END as PRODDATE" '"PRODDATE" '***
            Case "RiskToReview_Flag" : Return "CASE WHEN RiskToReview_Flag=1 THEN 1 WHEN RiskToReview_Flag IS NULL THEN NULL ELSE 0 END as RiskToReview_Flag"

            Case Else : Return """" & DatabaseColName & """"
        End Select
    End Function

    Public Overrides Function Get_SummaryTable_DefaultSortColumns() As String()
        Return {
            "ServiceRiskFactor"  ' if user mapped this column, the report will be sorted by this column by default
        }
    End Function

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

