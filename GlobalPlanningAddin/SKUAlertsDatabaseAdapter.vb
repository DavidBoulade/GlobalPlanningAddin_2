Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Data

Public Class SKUAlertsDatabaseAdapter : Inherits DatabaseAdapterBase
    Public Sub New()
        'Populate lists of specific columns
        SummaryTable_KeyColumns = _SummaryTableColumns.FindAll(Function(x) x.IsKey = True)
        SummaryTable_ModifiableColumns = _SummaryTableColumns.FindAll(Function(x) x.IsModifiable = True)
    End Sub

    Public Overrides ReadOnly Property SummaryTable_KeyColumns As List(Of DatabaseAdapterColumn)
    Public Overrides ReadOnly Property SummaryTable_ModifiableColumns As List(Of DatabaseAdapterColumn)

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
        Return "SKUALERTS"
    End Function

    Protected Overrides Function Get_SummaryTableUpdates_TableName() As String
        Return "SKUALERTS_UPDATES"
    End Function

    Protected Overrides Function Get_DetailsTable_Name() As String
        Return "STOCKREQLIST"
    End Function

    Protected Overrides Function Get_DetailledView_Optional_OrderBySQLClause() As String
        Return "ORDER BY SortIdAVST ASC"
    End Function

    Private ReadOnly _SummaryTableColumns As New List(Of DatabaseAdapterColumn)( 'the order of the key columns should match the order of the columns in the XXX_UPDATE table in SQL server as the BulkCopy fails if columns are not in the right order
        {
        New DatabaseAdapterColumn("Material", True, False, "", ""),
        New DatabaseAdapterColumn("Plant", True, False, "", ""),
        New DatabaseAdapterColumn("MaterialDescr", False, False, "", ""),
        New DatabaseAdapterColumn("MRPC", False, False, "", ""),
        New DatabaseAdapterColumn("MatType", False, False, "", ""),
        New DatabaseAdapterColumn("ABC", False, False, "", ""),
        New DatabaseAdapterColumn("LCS", False, False, "", ""),
        New DatabaseAdapterColumn("StdPrice", False, False, "", ""),
        New DatabaseAdapterColumn("PriceBase", False, False, "", ""),
        New DatabaseAdapterColumn("PurchGroup", False, False, "", ""),
        New DatabaseAdapterColumn("PTF", False, False, "", ""),
        New DatabaseAdapterColumn("LotSize", False, False, "", ""),
        New DatabaseAdapterColumn("MinLotSizeMM", False, False, "", ""),
        New DatabaseAdapterColumn("RoundingValue", False, False, "", ""),
        New DatabaseAdapterColumn("UOM", False, False, "", ""),
        New DatabaseAdapterColumn("PDT", False, False, "", ""),
        New DatabaseAdapterColumn("GRPT", False, False, "", ""),
        New DatabaseAdapterColumn("FirstOOSDate", False, False, "", "CASE WHEN FirstOOSDate='1900-01-01' THEN '' ELSE FORMAT(FirstOOSDate,'yyyy-MM-dd') END as FirstOOSDate"),
        New DatabaseAdapterColumn("MissedQtyPDT", False, False, "", ""),
        New DatabaseAdapterColumn("MissedQtyW", False, False, "", ""),
        New DatabaseAdapterColumn("MissedQtyW1", False, False, "", ""),
        New DatabaseAdapterColumn("MissedQtyW2", False, False, "", ""),
        New DatabaseAdapterColumn("MissedQtyW3", False, False, "", ""),
        New DatabaseAdapterColumn("MissedQtyW4", False, False, "", ""),
        New DatabaseAdapterColumn("MissedQtyW5", False, False, "", ""),
        New DatabaseAdapterColumn("MissedQtyW6+", False, False, "", ""),
        New DatabaseAdapterColumn("MissingSafetyPDT", False, False, "", ""),
        New DatabaseAdapterColumn("MissingSafetyQtyPDT", False, False, "", ""),
        New DatabaseAdapterColumn("MissingSafetyTimePDT", False, False, "", ""),
        New DatabaseAdapterColumn("MissingGRPTPDT", False, False, "", ""),
        New DatabaseAdapterColumn("FirstDeliveryDate", False, True, "DATE", "CASE WHEN FirstDeliveryDate='1900-01-01' THEN '' ELSE FORMAT(FirstDeliveryDate,'yyyy-MM-dd') END as FirstDeliveryDate"),
        New DatabaseAdapterColumn("FirstDeliveryQty", False, True, "NUMERIC", ""),
        New DatabaseAdapterColumn("RecoveryDate", False, False, "", "CASE WHEN RecoveryDate='1900-01-01' THEN '' ELSE FORMAT(RecoveryDate,'yyyy-MM-dd') END as RecoveryDate"),
        New DatabaseAdapterColumn("ServiceRootCause1", False, True, "STRING", ""),
        New DatabaseAdapterColumn("ServiceRootCause2", False, True, "STRING", ""),
        New DatabaseAdapterColumn("ServiceComment", False, True, "STRING", ""),
        New DatabaseAdapterColumn("SKU", False, False, "", "Material+'@'+Plant AS SKU"),
        New DatabaseAdapterColumn("OOSParents", False, False, "", ""),
        New DatabaseAdapterColumn("OOSParentsMRPC", False, False, "", ""),
        New DatabaseAdapterColumn("ActiveVendorCodes", False, False, "", ""),
        New DatabaseAdapterColumn("ActiveVendorNames", False, False, "", ""),
        New DatabaseAdapterColumn("MRPType", False, False, "", ""),
        New DatabaseAdapterColumn("MissedQtyPDT+", False, False, "", ""),
        New DatabaseAdapterColumn("MissedValPDT", False, False, "", ""),
        New DatabaseAdapterColumn("MissedValPDT+", False, False, "", ""),
        New DatabaseAdapterColumn("ServiceRisk", False, False, "", ""),
        New DatabaseAdapterColumn("UltimateOOSParents", False, False, "", ""),
        New DatabaseAdapterColumn("UltimateOOSParentsMRPC", False, False, "", ""),
        New DatabaseAdapterColumn("MissedQtyTotal", False, False, "", ""),
        New DatabaseAdapterColumn("MissedValTotal", False, False, "", ""),
        New DatabaseAdapterColumn("ServiceAlert", False, False, "", ""),
        New DatabaseAdapterColumn("SourceListFixVendor", False, False, "", ""),
        New DatabaseAdapterColumn("UserDefined1", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined2", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined3", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined4", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined5", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined6", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined7", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined8", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined9", False, True, "STRING", ""),
        New DatabaseAdapterColumn("UserDefined10", False, True, "STRING", ""),
        New DatabaseAdapterColumn("MRPCDescription", False, False, "", ""),
        New DatabaseAdapterColumn("House", False, False, "", ""),
        New DatabaseAdapterColumn("Brand", False, False, "", ""),
        New DatabaseAdapterColumn("Line", False, False, "", ""),
        New DatabaseAdapterColumn("Currency", False, False, "", ""),
        New DatabaseAdapterColumn("USDollarExchangeRate", False, False, "", ""),
        New DatabaseAdapterColumn("PlanningCal", False, False, "", ""),
        New DatabaseAdapterColumn("MinlotsizeVal", False, False, "", "(MinLotSizeMM * StdPrice * USDollarExchangeRate / NULLIF(PriceBase,0)) AS MinlotsizeVal"),
        New DatabaseAdapterColumn("SafetyTimeInd", False, False, "", ""),
        New DatabaseAdapterColumn("SafetyTime", False, False, "", ""),
        New DatabaseAdapterColumn("CompanyCode", False, False, "", ""),
        New DatabaseAdapterColumn("ProductHierarchy", False, False, "", ""),
        New DatabaseAdapterColumn("EAN", False, False, "", ""),
        New DatabaseAdapterColumn("SKUProjTooEarlyAreaVal", False, False, "", ""),
        New DatabaseAdapterColumn("SKUProjTooMuchAreaVal", False, False, "", ""),
        New DatabaseAdapterColumn("AverageTooEarlyNbDays", False, False, "", ""),
        New DatabaseAdapterColumn("SumTooMuchValues", False, False, "", ""),
        New DatabaseAdapterColumn("SKUProjExcessAreaVal", False, False, "", ""),
        New DatabaseAdapterColumn("SKUInvTooEarlyAreaVal", False, False, "", ""),
        New DatabaseAdapterColumn("SKUInvTooMuchAreaVal", False, False, "", ""),
        New DatabaseAdapterColumn("SKUInvExcessAreaVal", False, False, "", ""),
        New DatabaseAdapterColumn("SKUInventoryAreaVal", False, False, "", ""),
        New DatabaseAdapterColumn("ExcessAlert", False, False, "", ""),
        New DatabaseAdapterColumn("ExcessAlertCreationDate", False, False, "", "CASE WHEN ExcessAlertCreationDate='1900-01-01' THEN '' ELSE FORMAT(ExcessAlertCreationDate,'yyyy-MM-dd') END as ExcessAlertCreationDate"),
        New DatabaseAdapterColumn("NbDaysCurrentExcessAlert", False, False, "", ""),
        New DatabaseAdapterColumn("ExcessActionned", False, True, "STRING", ""),
        New DatabaseAdapterColumn("ExcessRootCause1", False, True, "STRING", ""),
        New DatabaseAdapterColumn("ExcessRootCause2", False, True, "STRING", ""),
        New DatabaseAdapterColumn("ExcessComment", False, True, "STRING", ""),
        New DatabaseAdapterColumn("SKUProjExcessAreaValM1", False, False, "", ""),
        New DatabaseAdapterColumn("SKUProjExcessAreaValM2", False, False, "", ""),
        New DatabaseAdapterColumn("SKUProjExcessAreaValM3", False, False, "", ""),
        New DatabaseAdapterColumn("SKUProjExcessAreaValM4", False, False, "", ""),
        New DatabaseAdapterColumn("SKUProjExcessAreaValM5", False, False, "", ""),
        New DatabaseAdapterColumn("SKUProjExcessAreaValM6", False, False, "", "")
        })

    Public Overrides ReadOnly Property SummaryTableColumns As List(Of DatabaseAdapterColumn)
        Get
            Return _SummaryTableColumns
        End Get
    End Property


    Public Overrides Function Get_SummaryTable_DefaultSortColumns() As String()
        Return {
            "ServiceRisk", 'if user mapped this column, the report will be sorted by this column by default
            "SKUProjExcessAreaVal" 'if the first is not mapped and this one is, the report will be sorted by this column
        }
    End Function

    Public Overrides Function Get_DetailedView_Columns() As String()
        Return {
                "ReceiptDateGR",
                "ReceiptDateAv",
                "ReceiptDateAvST",
                "StartDate",
                "FinishDate",
                "GR_Week",
                "GR_Month",
                "GR_FY",
                "OpeningDate",
                "MRPElement",
                "MRPElementDescr",
                "MRPElementData",
                "ReschedulingDaysGR",
                "ReschDateGR",
                "ReschDateAv",
                "ExceptionKey",
                "ExceptionNo",
                "RequirementSign",
                "RequirementQty",
                "AvailableQtyGR",
                "AvailableQtyAvST",
                "FreeTextEntry",
                "ProdVersion",
                "ProdLine",
                "StorageLoc",
                "VendorNo",
                "VendorName",
                "CustNo",
                "CustName",
                "DOC",
                "DOCVariation",
                "TooEarlyNbDays",
                "InitialTooMuchQty",
                "InitialTooMuchVal",
                "AddedTooEarlyAreaVal",
                "AddedTooMuchAreaVal",
                "AddedExcessAreaVal",
                "AddedInventoryAreaVal",
                "Excess6MQty",
                "Excess12MQty",
                "Excess6MVal",
                "Excess12MVal"}
    End Function

    Public Overrides Function Get_DetailedView_CurItem_HeaderText() As String()
        Return {
                "Plant", "' | '", "MaterialDescr"}
    End Function


    Private ReadOnly _DetailedView_InfoDropDown_Items As New List(Of String())({
                ({"Plant", "Plant"}),
                ({"Material", "Material"}),
                ({"Description", "MaterialDescr"}),
                ({"Updated Date Time", "UpdateDateTime"}),
                ({"MRP Controller", "MRPC", "' ('", "MRPCDescription", "')'"}),
                ({"House", "House"}),
                ({"Brand", "Brand"}),
                ({"Line", "Line"}),
                ({"Source List default Vendor", "SourceListFixVendor"}),
                ({"Active vendor Codes", "ActiveVendorCodes", "' ('", "ActiveVendorNames", "')'"}),
                ({"ABC", "ABC"}),
                ({"LCS", "LCS"}),
                ({"PTF", "PTF"}),
                ({"Std Price", "StdPrice", "' '", "Currency", "' / '", "PriceBase", "' '", "UOM"}),
                ({"US Dollar Exchange Rate", "USDollarExchangeRate", "' USD/'", "Currency"}),
                ({"Currency", "Currency"}),
                ({"MRP Type", "MRPType"}),
                ({"Lot Size", "LotSize"}),
                ({"Planning Calendar", "PlanningCal"}),
                ({"Min Lot Size (MM)", "MinLotSizeMM"}),
                ({"Rounding Value", "RoundingValue"}),
                ({"Unit of measure", "UOM"}),
                ({"Material Type", "MatType"}),
                ({"Purchasing Group", "PurchGroup"}),
                ({"PDT", "PDT"}),
                ({"GRPT", "GRPT"}),
                ({"SafetyTimeInd", "SafetyTimeInd"}),
                ({"SafetyTime", "SafetyTime"}),
                ({"Company Code", "CompanyCode"}),
                ({"Product Hierarchy", "ProductHierarchy"}),
                ({"EAN", "EAN"})
                })

    Public Overrides Function Get_DetailedView_InfoDropDown_Items() As List(Of String())
        Return _DetailedView_InfoDropDown_Items
    End Function
End Class