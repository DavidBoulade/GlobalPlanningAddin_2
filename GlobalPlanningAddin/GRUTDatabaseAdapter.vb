Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Data

Public Class GRUTDatabaseAdapter : Inherits DatabaseAdapterBase

    Public Sub New()
        'Populate lists of specific columns
        SummaryTable_KeyColumns = _SummaryTableColumns.FindAll(Function(x) x.IsKey = True)
        SummaryTable_ModifiableColumns = _SummaryTableColumns.FindAll(Function(x) x.IsModifiable = True)
    End Sub

    Public Overrides ReadOnly Property SummaryTable_KeyColumns As List(Of DatabaseAdapterColumn)
    Public Overrides ReadOnly Property SummaryTable_ModifiableColumns As List(Of DatabaseAdapterColumn)

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

    Protected Overrides Function Get_DetailsTable_Name() As String
        Return "GRUT_PROJECTION"
    End Function


    Private ReadOnly _SummaryTableColumns As New List(Of DatabaseAdapterColumn)( 'the order of the key columns should match the order of the columns in the XXX_UPDATE table in SQL server as the BulkCopy fails if columns are not in the right order
        { 'The Data type is mandatory if the column is modifiable, not used otherwise
        New DatabaseAdapterColumn("EXTRACT_DATETIME", False, False, "", ""),
        New DatabaseAdapterColumn("ReportDate", False, False, "", ""), 'ReportDate should NOT be set as key column here, it is a mandatory key for all reports
        New DatabaseAdapterColumn("Item", True, False, "", ""),
        New DatabaseAdapterColumn("Loc", True, False, "", ""),
        New DatabaseAdapterColumn("Item_Description", False, False, "", ""),
        New DatabaseAdapterColumn("SKU", False, False, "", ""),
        New DatabaseAdapterColumn("Cur_Source", False, False, "", ""),
        New DatabaseAdapterColumn("Cur_FromSource_Transportation_LeadTime", False, False, "", ""),
        New DatabaseAdapterColumn("Cur_FromFactory_Transportation_LeadTime", False, False, "", ""),
        New DatabaseAdapterColumn("Cur_Factory_Internal_Reaction_Time", False, False, "", ""),
        New DatabaseAdapterColumn("Cur_FromFactory_Total_Replanishment_LeadTime", False, False, "", ""),
        New DatabaseAdapterColumn("Cur_DRPCovDur", False, False, "", ""),
        New DatabaseAdapterColumn("Factory_ID", False, False, "", ""),
        New DatabaseAdapterColumn("Article_Type", False, False, "", ""),
        New DatabaseAdapterColumn("Article_SubType", False, False, "", ""),
        New DatabaseAdapterColumn("Factory", False, False, "", ""),
        New DatabaseAdapterColumn("Division", False, False, "", ""),
        New DatabaseAdapterColumn("House", False, False, "", ""),
        New DatabaseAdapterColumn("Brand", False, False, "", ""),
        New DatabaseAdapterColumn("Line", False, False, "", ""),
        New DatabaseAdapterColumn("Product_Segment", False, False, "", ""),
        New DatabaseAdapterColumn("Core_category", False, False, "", ""),
        New DatabaseAdapterColumn("Market_category", False, False, "", ""),
        New DatabaseAdapterColumn("Sub_category", False, False, "", ""),
        New DatabaseAdapterColumn("Sp_category", False, False, "", ""),
        New DatabaseAdapterColumn("Planner_Code", False, False, "", ""),
        New DatabaseAdapterColumn("Planner_Name", False, False, "", ""),
        New DatabaseAdapterColumn("SKU_Planner_Code", False, False, "", ""),
        New DatabaseAdapterColumn("SKU_Planner_Name", False, False, "", ""),
        New DatabaseAdapterColumn("Free_Stock", False, False, "", ""),
        New DatabaseAdapterColumn("OnHand_QC_Qty", False, False, "", ""),
        New DatabaseAdapterColumn("OnHand_Qty", False, False, "", ""),
        New DatabaseAdapterColumn("Todays_orders", False, False, "", ""),
        New DatabaseAdapterColumn("ABC_Item", False, False, "", ""),
        New DatabaseAdapterColumn("ABC_SKU", False, False, "", ""),
        New DatabaseAdapterColumn("DOC", False, False, "", ""),
        New DatabaseAdapterColumn("DOC_Change", False, False, "", ""),
        New DatabaseAdapterColumn("SS_VS_OH_today", False, False, "", ""),
        New DatabaseAdapterColumn("PlanSwitch", False, False, "", ""),
        New DatabaseAdapterColumn("Quota_From", False, False, "", ""),
        New DatabaseAdapterColumn("Quota_To", False, False, "", ""),
        New DatabaseAdapterColumn("Quota_Type", False, False, "", ""),
        New DatabaseAdapterColumn("New_Item", False, False, "", ""),
        New DatabaseAdapterColumn("Old_Item", False, False, "", ""),
        New DatabaseAdapterColumn("ICW", False, False, "", "CASE WHEN ICW='1970-01-01' THEN NULL ELSE FORMAT(ICW,'yyyy-MM-dd') END as ICW"),
        New DatabaseAdapterColumn("ABC_SKU_Rank", False, False, "", ""),
        New DatabaseAdapterColumn("UnderForecast_switch", False, False, "", ""),
        New DatabaseAdapterColumn("UnderForecast_rules_results", False, False, "", ""),
        New DatabaseAdapterColumn("UnderForecast_Cond_1", False, False, "", ""),
        New DatabaseAdapterColumn("UnderForecast_Cond_2", False, False, "", ""),
        New DatabaseAdapterColumn("UnderForecast_Cond_3", False, False, "", ""),
        New DatabaseAdapterColumn("UnderForecast_Cond_4", False, False, "", ""),
        New DatabaseAdapterColumn("Current_Month_Orders", False, False, "", ""),
        New DatabaseAdapterColumn("Next_Month_Orders", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Orders", False, False, "", ""),
        New DatabaseAdapterColumn("Current_Month_forecast", False, False, "", ""),
        New DatabaseAdapterColumn("Current_Month_forecast_Var", False, False, "", ""),
        New DatabaseAdapterColumn("Next_Month_forecast", False, False, "", ""),
        New DatabaseAdapterColumn("End_of_Demand_Date", False, False, "", "CASE WHEN End_of_Demand_Date='1970-01-01' THEN NULL ELSE FORMAT(End_of_Demand_Date,'yyyy-MM-dd') END as End_of_Demand_Date"),
        New DatabaseAdapterColumn("Ex_Factory_Date", False, False, "", "CASE WHEN Ex_Factory_Date='1970-01-01' THEN NULL ELSE FORMAT(Ex_Factory_Date,'yyyy-MM-dd') END as Ex_Factory_Date"),
        New DatabaseAdapterColumn("Factory_Disc_Date", False, False, "", "CASE WHEN Factory_Disc_Date='1970-01-01' THEN NULL ELSE FORMAT(Factory_Disc_Date,'yyyy-MM-dd') END as Factory_Disc_Date"),
        New DatabaseAdapterColumn("Stock_Out_Date", False, False, "", "CASE WHEN Stock_Out_Date='1970-01-01' THEN NULL ELSE FORMAT(Stock_Out_Date,'yyyy-MM-dd') END as Stock_Out_Date"),
        New DatabaseAdapterColumn("Oracle_Code", False, False, "", ""),
        New DatabaseAdapterColumn("SSCov", False, False, "", ""),
        New DatabaseAdapterColumn("MinSS", False, False, "", ""),
        New DatabaseAdapterColumn("BISS", False, False, "", "CASE WHEN BISS='1970-01-01' THEN NULL ELSE FORMAT(BISS,'yyyy-MM-dd') END as BISS"),
        New DatabaseAdapterColumn("Factory_OH", False, False, "", ""),
        New DatabaseAdapterColumn("Factory_QC", False, False, "", ""),
        New DatabaseAdapterColumn("FCST_PERF_M1", False, False, "", ""),
        New DatabaseAdapterColumn("FCST_PERF_M2", False, False, "", ""),
        New DatabaseAdapterColumn("FCST_PERF_M3", False, False, "", ""),
        New DatabaseAdapterColumn("SALES_VS_CM_FCST", False, False, "", ""),
        New DatabaseAdapterColumn("ServiceRiskFactor", False, False, "", ""),
        New DatabaseAdapterColumn("RiskToReview_Flag", False, True, "STRING", "CASE WHEN RiskToReview_Flag=1 THEN 1 WHEN RiskToReview_Flag IS NULL THEN NULL ELSE 0 END as RiskToReview_Flag"),
        New DatabaseAdapterColumn("RiskToReview_Date", False, False, "", ""),
        New DatabaseAdapterColumn("First_Supply_Date", False, False, "", ""),
        New DatabaseAdapterColumn("First_Supply_Qty", False, False, "", ""),
        New DatabaseAdapterColumn("First_Local_OOS_Date", False, False, "", ""),
        New DatabaseAdapterColumn("First_Local_Recovery_Date", False, False, "", ""),
        New DatabaseAdapterColumn("First_Local_Back_in_SS_Date", False, False, "", ""),
        New DatabaseAdapterColumn("Final_Local_Recovery_Date", False, False, "", ""),
        New DatabaseAdapterColumn("Final_Local_Back_in_SS_Date", False, False, "", ""),
        New DatabaseAdapterColumn("First_OOS_Date", False, False, "", ""),
        New DatabaseAdapterColumn("First_Total_Recovery_Date", False, False, "", ""),
        New DatabaseAdapterColumn("Final_Total_Recovery_Date", False, False, "", ""),
        New DatabaseAdapterColumn("First_Total_Indep_Risk_Date", False, False, "", ""),
        New DatabaseAdapterColumn("Final_Total_Indep_Recovery_Date", False, False, "", ""),
        New DatabaseAdapterColumn("Item_Total_Indep_Risk_6M", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_Within_Leadtime", False, False, "", ""),
        New DatabaseAdapterColumn("Total_IntransitIn_Qty", False, False, "", ""),
        New DatabaseAdapterColumn("OnHand_Blocked", False, False, "", ""),
        New DatabaseAdapterColumn("PRODQTY", False, False, "", ""),
        New DatabaseAdapterColumn("PRODDATE", False, False, "", "CASE WHEN PRODDATE='1970-01-01' THEN NULL ELSE FORMAT(PRODDATE,'yyyy-MM-dd') END as PRODDATE"),
        New DatabaseAdapterColumn("PRODORDERID", False, False, "", ""),
        New DatabaseAdapterColumn("DMDTODATE", False, False, "", ""),
        New DatabaseAdapterColumn("E_O", False, False, "", ""),
        New DatabaseAdapterColumn("First_IntransitIn_Qty", False, False, "", ""),
        New DatabaseAdapterColumn("First_IntransitIn_Arrival_Date", False, False, "", "CASE WHEN First_IntransitIn_Arrival_Date='1970-01-01' THEN NULL ELSE FORMAT(First_IntransitIn_Arrival_Date,'yyyy-MM-dd') END as First_IntransitIn_Arrival_Date"),
        New DatabaseAdapterColumn("First_Committed_Qty", False, False, "", ""),
        New DatabaseAdapterColumn("First_Committed_Date", False, False, "", "CASE WHEN First_Committed_Date='1970-01-01' THEN NULL ELSE FORMAT(First_Committed_Date,'yyyy-MM-dd') END as First_Committed_Date"),
        New DatabaseAdapterColumn("First_Recship_Qty", False, False, "", ""),
        New DatabaseAdapterColumn("First_Recship_Date", False, False, "", "CASE WHEN First_Recship_Date='1970-01-01' THEN NULL ELSE FORMAT(First_Recship_Date,'yyyy-MM-dd') END as First_Recship_Date"),
        New DatabaseAdapterColumn("Next_Dely_Qty", False, False, "", ""),
        New DatabaseAdapterColumn("Next_Dely_Date", False, False, "", ""),
        New DatabaseAdapterColumn("Next_Dely_Qty_Override", False, True, "NUMERIC", ""),
        New DatabaseAdapterColumn("Next_Dely_Date_Override", False, True, "DATE", ""),
        New DatabaseAdapterColumn("Service_RootCause_ID", False, False, "", ""),
        New DatabaseAdapterColumn("Service_RootCause_Level_3", False, True, "STRING", ""),
        New DatabaseAdapterColumn("Service_RootCause_Level_2", False, False, "", ""),
        New DatabaseAdapterColumn("Service_RootCause_Level_1", False, False, "", ""),
        New DatabaseAdapterColumn("Service_RootCause_Accountable", False, False, "", ""),
        New DatabaseAdapterColumn("RCA_Comment", False, True, "STRING", ""),
        New DatabaseAdapterColumn("U_Comment", False, True, "STRING", ""),
        New DatabaseAdapterColumn("Fa_Comment", False, True, "STRING", ""),
        New DatabaseAdapterColumn("GDO_Comment", False, True, "STRING", ""),
        New DatabaseAdapterColumn("GDO_Reviewed", False, False, "", "CASE WHEN GDO_Reviewed='1970-01-01' THEN NULL ELSE FORMAT(GDO_Reviewed,'yyyy-MM-dd') END as GDO_Reviewed"),
        New DatabaseAdapterColumn("VAR_COMMENT", False, True, "STRING", ""),
        New DatabaseAdapterColumn("Service_Risk_Action", False, True, "STRING", ""),
        New DatabaseAdapterColumn("Service_Risk_Action_Owner", False, True, "STRING", ""),
        New DatabaseAdapterColumn("Service_Risk_Action_Date", False, True, "DATE", ""),
        New DatabaseAdapterColumn("Service_Risk_Constraint", False, True, "STRING", ""),
        New DatabaseAdapterColumn("REVIEWED_DATE", False, False, "", "CASE WHEN REVIEWED_DATE='1970-01-01' THEN NULL ELSE FORMAT(REVIEWED_DATE,'yyyy-MM-dd') END as REVIEWED_DATE"),
        New DatabaseAdapterColumn("Risk_on_order_qty_D", False, False, "", ""),
        New DatabaseAdapterColumn("Risk_on_order_qty_D1", False, False, "", ""),
        New DatabaseAdapterColumn("Risk_on_order_qty_D2", False, False, "", ""),
        New DatabaseAdapterColumn("Risk_on_order_qty_D3", False, False, "", ""),
        New DatabaseAdapterColumn("Risk_on_order_qty_D4", False, False, "", ""),
        New DatabaseAdapterColumn("Risk_on_order_qty_D5", False, False, "", ""),
        New DatabaseAdapterColumn("Risk_on_order_qty_D6", False, False, "", ""),
        New DatabaseAdapterColumn("Risk_on_order_qty_7days", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_Before_Leadtime_From_Factory", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_After_Leadtime_From_Factory", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_Before_Leadtime_From_Source", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_Within_Replanishment_period_From_Source", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_Before_Leadtime_From_Factory", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_After_Leadtime_From_Factory", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_Before_Leadtime_From_Source", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_Within_Replanishment_period_From_Source", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_Before_Leadtime_From_Factory", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_After_Leadtime_From_Factory", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_Before_Leadtime_From_Source", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_Within_Replanishment_period_From_Source", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_PAST", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W01", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W02", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W03", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W04", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W05", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W06", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W07", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W08", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W09", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W10", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W11", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W12", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W13", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W14", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W15", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W16", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W17", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W18", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W19", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W20", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W21", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W22", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W23", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W24", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W25", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_W26", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W01", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W02", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W03", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W04", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W05", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W06", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W07", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W08", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W09", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W10", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W11", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W12", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W13", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W14", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W15", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W16", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W17", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W18", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W19", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W20", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W21", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W22", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W23", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W24", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W25", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_W26", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W01", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W02", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W03", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W04", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W05", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W06", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W07", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W08", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W09", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W10", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W11", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W12", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W13", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W14", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W15", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W16", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W17", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W18", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W19", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W20", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W21", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W22", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W23", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W24", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W25", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_W26", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W01", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W02", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W03", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W04", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W05", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W06", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W07", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W08", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W09", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W10", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W11", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W12", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W13", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W14", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W15", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W16", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W17", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W18", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W19", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W20", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W21", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W22", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W23", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W24", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W25", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_W26", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W01", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W02", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W03", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W04", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W05", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W06", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W07", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W08", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W09", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W10", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W11", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W12", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W13", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W14", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W15", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W16", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W17", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W18", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W19", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W20", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W21", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W22", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W23", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W24", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W25", False, False, "", ""),
        New DatabaseAdapterColumn("MinSSOH_W26", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W01", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W02", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W03", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W04", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W05", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W06", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W07", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W08", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W09", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W10", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W11", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W12", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W13", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W14", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W15", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W16", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W17", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W18", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W19", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W20", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W21", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W22", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W23", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W24", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W25", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_W26", False, False, "", ""),
        New DatabaseAdapterColumn("UserDefined_1", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_2", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_3", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_4", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_5", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_6", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_7", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_8", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_9", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_10", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_Reset_1", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_Reset_2", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_Reset_3", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_Reset_4", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_Reset_5", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_Reset_6", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_Reset_7", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_Reset_8", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_Reset_9", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined_Reset_10", False, True, "STRING", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_M1", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_M2", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_M3", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_M4", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_M5", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_M6", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_M1", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_M2", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_M3", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_M4", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_M5", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Dep_Risk_M6", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_M1", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_M2", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_M3", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_M4", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_M5", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_M6", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_M1", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_M2", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_M3", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_M4", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_M5", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_M6", False, False, "", ""),
        New DatabaseAdapterColumn("Indep_Dmd_M1", False, False, "", ""),
        New DatabaseAdapterColumn("Indep_Dmd_M2", False, False, "", ""),
        New DatabaseAdapterColumn("Indep_Dmd_M3", False, False, "", ""),
        New DatabaseAdapterColumn("Indep_Dmd_M4", False, False, "", ""),
        New DatabaseAdapterColumn("Indep_Dmd_M5", False, False, "", ""),
        New DatabaseAdapterColumn("Indep_Dmd_M6", False, False, "", ""),
        New DatabaseAdapterColumn("Dep_Dmd_M1", False, False, "", ""),
        New DatabaseAdapterColumn("Dep_Dmd_M2", False, False, "", ""),
        New DatabaseAdapterColumn("Dep_Dmd_M3", False, False, "", ""),
        New DatabaseAdapterColumn("Dep_Dmd_M4", False, False, "", ""),
        New DatabaseAdapterColumn("Dep_Dmd_M5", False, False, "", ""),
        New DatabaseAdapterColumn("Dep_Dmd_M6", False, False, "", ""),
        New DatabaseAdapterColumn("Other_Dmd_M1", False, False, "", ""),
        New DatabaseAdapterColumn("Other_Dmd_M2", False, False, "", ""),
        New DatabaseAdapterColumn("Other_Dmd_M3", False, False, "", ""),
        New DatabaseAdapterColumn("Other_Dmd_M4", False, False, "", ""),
        New DatabaseAdapterColumn("Other_Dmd_M5", False, False, "", ""),
        New DatabaseAdapterColumn("Other_Dmd_M6", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_1M", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_2M", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_3M", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_4M", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_5M", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Indep_Risk_6M", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_1M", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_2M", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_3M", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_4M", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_5M", False, False, "", ""),
        New DatabaseAdapterColumn("Local_Indep_Risk_6M", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_1M", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_2M", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_3M", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_4M", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_5M", False, False, "", ""),
        New DatabaseAdapterColumn("Total_Risk_6M", False, False, "", ""),
        New DatabaseAdapterColumn("U_EAN", False, False, "", ""),
        New DatabaseAdapterColumn("U_ICP", False, False, "", ""),
        New DatabaseAdapterColumn("QT_Quota_M1", False, False, "", ""),
        New DatabaseAdapterColumn("QT_Quota_M2", False, False, "", ""),
        New DatabaseAdapterColumn("QT_Quota_M3", False, False, "", ""),
        New DatabaseAdapterColumn("QT_Quota_M4", False, False, "", ""),
        New DatabaseAdapterColumn("QT_Quota_M5", False, False, "", ""),
        New DatabaseAdapterColumn("QT_Quota_M6", False, False, "", ""),
        New DatabaseAdapterColumn("QT_Quota_Type", False, False, "", ""),
        New DatabaseAdapterColumn("Phase_In_Item", False, False, "", ""),
        New DatabaseAdapterColumn("DFUs_Strategies", False, False, "", ""),
        New DatabaseAdapterColumn("SKU_Level", False, False, "", ""),
        New DatabaseAdapterColumn("Sourcing_Path", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_M1", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_M2", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_M3", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_M4", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_M5", False, False, "", ""),
        New DatabaseAdapterColumn("Min_AltSupsdConstrProjOH_M6", False, False, "", "")
        })

    Public Overrides ReadOnly Property SummaryTableColumns As List(Of DatabaseAdapterColumn)
        Get
            Return _SummaryTableColumns
        End Get
    End Property





    Public Overrides Function Get_SummaryTable_DefaultSortColumns() As String()
        '"ServiceRiskFactor"  'if user mapped this column, the report will be sorted by this column by default
        Return {
            "Item_Total_Indep_Risk_6M"
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

