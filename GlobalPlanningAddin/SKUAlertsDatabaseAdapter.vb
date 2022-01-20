Option Explicit On
Option Strict On

Imports System.Data.SqlClient
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Data

Public Class SKUAlertsDatabaseAdapter : Inherits DatabaseAdapterBase

    Protected Overrides Function Get_ConnectionString() As String
        Return "Server=USSANTDB02P\NA_SUPPLY_CHAIN;Database=SKUAlerts;UID=GlobalPlanningAddinUser;PWD=iojrgRGRE**$8421;"
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

    Public Overrides Function Get_SummaryTable_ListOfModifiableColumns() As String()
        Return {
                "ExcessActionned", 'char(1)
                "ExcessRootCause1", 'dropdown list of string
                "ExcessRootCause2", 'dropdownlist of string
                "ExcessComment", 'free string
                "UserDefined1", 'free string
                "UserDefined2", 'free string
                "UserDefined3", 'free string
                "UserDefined4", 'free string
                "UserDefined5", 'free string
                "UserDefined6", 'free string
                "UserDefined7", 'free string
                "UserDefined8", 'free string
                "UserDefined9", 'free string
                "UserDefined10", 'free string
                "FirstDeliveryDate", 'date
                "FirstDeliveryQty", 'double
                "ServiceRootCause1", 'dropdown list of string
                "ServiceRootCause2", 'dropdown list of string
                "ServiceComment" 'free string
                }
    End Function

    'List all the modifiable columns with numeric type
    Public Overrides Function Get_SummaryTable_ListOfNumericColumns() As String()
        Return {
                "FirstDeliveryQty" 'double
                }
    End Function

    'List all modifiable columns with Date type
    Public Overrides Function Get_SummaryTable_ListOfDateColumns() As String()
        Return {
                "FirstDeliveryDate" 'date
                }
    End Function

    Public Overrides Function Get_SummaryTable_KeyColumns() As String() 'Key columns (excluding the ReportDate that must be key as well). These columns should also be Keys in the detailed table
        Return { 'the order should match the order of the columns in the XXX_UPDATE table in SQL server as the BulkCopy fails if columns are not in the right order
            "Material",
            "Plant"}
    End Function

    Public Overrides Function Get_SummaryTable_Columns() As String()
        Return {
                "Plant",
                "Material",
                "MaterialDescr",
                "MRPC",
                "MatType",
                "ABC",
                "LCS",
                "StdPrice",
                "PriceBase",
                "PurchGroup",
                "PTF",
                "LotSize",
                "MinLotSizeMM",
                "RoundingValue",
                "UOM",
                "PDT",
                "GRPT",
                "FirstOOSDate",
                "MissedQtyPDT",
                "MissedQtyW",
                "MissedQtyW1",
                "MissedQtyW2",
                "MissedQtyW3",
                "MissedQtyW4",
                "MissedQtyW5",
                "MissedQtyW6+",
                "MissingSafetyPDT",
                "MissingSafetyQtyPDT",
                "MissingSafetyTimePDT",
                "MissingGRPTPDT",
                "FirstDeliveryDate",
                "FirstDeliveryQty",
                "RecoveryDate",
                "ServiceRootCause1",
                "ServiceRootCause2",
                "ServiceComment",
                "SKU",
                "OOSParents",
                "OOSParentsMRPC",
                "ActiveVendorCodes",
                "ActiveVendorNames",
                "MRPType",
                "MissedQtyPDT+",
                "MissedValPDT",
                "MissedValPDT+",
                "ServiceRisk",
                "UltimateOOSParents",
                "UltimateOOSParentsMRPC",
                "MissedQtyTotal",
                "MissedValTotal",
                "ServiceAlert",
                "SourceListFixVendor",
                "UserDefined1",
                "UserDefined2",
                "UserDefined3",
                "UserDefined4",
                "UserDefined5",
                "UserDefined6",
                "UserDefined7",
                "UserDefined8",
                "UserDefined9",
                "UserDefined10",
                "MRPCDescription",
                "House",
                "Brand",
                "Line",
                "Currency",
                "USDollarExchangeRate",
                "PlanningCal",
                "MinlotsizeVal",
                "SafetyTimeInd",
                "SafetyTime",
                "CompanyCode",
                "ProductHierarchy",
                "EAN",
                "SKUProjTooEarlyAreaVal",
                "SKUProjTooMuchAreaVal",
                "AverageTooEarlyNbDays",
                "SumTooMuchValues",
                "SKUProjExcessAreaVal",
                "SKUInvTooEarlyAreaVal",
                "SKUInvTooMuchAreaVal",
                "SKUInvExcessAreaVal",
                "SKUInventoryAreaVal",
                "ExcessAlert",
                "ExcessAlertCreationDate",
                "NbDaysCurrentExcessAlert",
                "ExcessActionned",
                "ExcessRootCause1",
                "ExcessRootCause2",
                "ExcessComment",
                "SKUProjExcessAreaValM1",
                "SKUProjExcessAreaValM2",
                "SKUProjExcessAreaValM3",
                "SKUProjExcessAreaValM4",
                "SKUProjExcessAreaValM5",
                "SKUProjExcessAreaValM6"
                }
    End Function
    Public Overrides Function Get_SummaryTable_SQLQueryForField(DatabaseColName As String) As String

        Select Case DatabaseColName
            Case "FirstOOSDate" : Return "CASE WHEN FirstOOSDate='1900-01-01' THEN '' ELSE FORMAT(FirstOOSDate,'yyyy-MM-dd') END as FirstOOSDate"'"FirstOOSDate"
            Case "FirstDeliveryDate" : Return "CASE WHEN FirstDeliveryDate='1900-01-01' THEN '' ELSE FORMAT(FirstDeliveryDate,'yyyy-MM-dd') END as FirstDeliveryDate"'"FirstDeliveryDate"
            Case "RecoveryDate" : Return "CASE WHEN RecoveryDate='1900-01-01' THEN '' ELSE FORMAT(RecoveryDate,'yyyy-MM-dd') END as RecoveryDate"'"RecoveryDate"
            Case "SKU" : Return "Material+'@'+Plant AS SKU" 'SKU
            Case "MinlotsizeVal" : Return "(MinLotSizeMM * StdPrice * USDollarExchangeRate / NULLIF(PriceBase,0)) AS MinlotsizeVal" 'MinlotsizeVal
            Case "ExcessAlertCreationDate" : Return "CASE WHEN ExcessAlertCreationDate='1900-01-01' THEN '' ELSE FORMAT(ExcessAlertCreationDate,'yyyy-MM-dd') END as ExcessAlertCreationDate" '"ExcessAlertCreationDate"
            Case Else : Return """" & DatabaseColName & """"
        End Select
    End Function

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