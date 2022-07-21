Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Data

Public Class SKUAlertsDatabaseAdapter : Inherits DatabaseAdapterBase
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
        Return "SKUALERTS_VIEW"
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

    Public Overrides Function SummaryTable_DefaultSortColumns() As List(Of SortField)
        Return _SortFields
    End Function

    Private ReadOnly _SortFields As New List(Of SortField)(
        {New SortField("ServiceRisk", SortField.SortOrders.Descending),
        New SortField("SKUProjExcessAreaVal", SortField.SortOrders.Descending)}
        )

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