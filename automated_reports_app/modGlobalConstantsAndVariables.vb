Module modGlobalConstantsAndVariables

    Public SalesDailyFlag As Boolean = False
    Public SalesWeeklyFlag As Boolean = False
    Public SalesMonthlyFlag As Boolean = False
    Public SalesYearlyFlag As Boolean = False

    Public InventoryDailyFlag As Boolean = False
    Public InventoryWeeklyFlag As Boolean = False
    Public InventoryMonthlyFlag As Boolean = False
    Public InventoryYearlyFlag As Boolean = False

    Public TaxDailyFlag As Boolean = False
    Public TaxWeeklyFlag As Boolean = False
    Public TaxMonthlyFlag As Boolean = False
    Public TaxYearlyFlag As Boolean = False

    Public DepositDailyFlag As Boolean = False
    Public DepositWeeklyFlag As Boolean = False
    Public DepositMonthlyFlag As Boolean = False
    Public DepositYearlyFlag As Boolean = False

    Public strConnectionUsername As String
    Public strConnectionPassword As String
    'Public btmButtonDefault As Bitmap = My.Resources.Button
    'Public btmButtonDefaultGray As Bitmap = My.Resources.Button
    'Public btmButtonShort As Bitmap = My.Resources.ButtonShort
    'Public btmButtonShortGray As Bitmap = My.Resources.ButtonShort
    ' PLEASE NOTE THAT IN VB.NET THE AXES OR SWAPPED. (Y, X)
    Public aastrCSVFile(,) As String = New String(15, 1) {
        {"Sales Report Daily", "false"},
        {"Sales Report Weekly", "false"},
        {"Sales Report Monthly", "false"},
        {"Sales Report Yearly", "false"},
        {"Sales Tax Report Daily", "false"},
        {"Sales Tax Report Weekly", "false"},
        {"Sales Tax Report Monthly", "false"},
        {"Sales Tax Report Yearly", "false"},
        {"Sales Inventory Daily", "false"},
        {"Sales Inventory Weekly", "false"},
        {"Sales Inventory Monthly", "false"},
        {"Sales Inventory Yearly", "false"},
        {"Sales Deposit Daily", "false"},
        {"Sales Deposit Weekly", "false"},
        {"Sales Deposit Monthly", "false"},
        {"Sales Deposit Yearly", "false"}
    }


    Public strEmailSalesReport As String
    Public strEmailSalesTaxReport As String
    Public strEmailInventoryReport As String
    Public strEmailDespositReport As String

    Public blnSalesReportDaily As Boolean
    Public blnSalesReportMonthly As Boolean
    Public blnSalesReportWeekly As Boolean
    Public blnSalesReportYearly As Boolean

    Public blnInventoryReportDaily As Boolean
    Public blnInventoryReportMonthly As Boolean
    Public blnInventoryReportWeekly As Boolean
    Public blnInventoryReportYearly As Boolean

    Public blnSalesTaxReportDaily As Boolean
    Public blnSalesTaxReportMonthly As Boolean
    Public blnSalesTaxReportWeekly As Boolean
    Public blnSalesTaxReportYearly As Boolean

    Public blnDepositReportDaily As Boolean
    Public blnDepositReportMonthly As Boolean
    Public blnDepositReportWeekly As Boolean
    Public blnDepositReportYearly As Boolean

    Public dtmDailySalesReport As DateTime
    Public dtmDailySalesTaxReport As DateTime
    Public dtmDailyInventoryReport As DateTime
    Public dtmDailyDepositReport As DateTime

    Public dtmWeeklySalesReport As DateTime
    Public dtmWeeklySalesTaxReport As DateTime
    Public dtmWeeklyInventoryReport As DateTime
    Public dtmWeeklyDepositReport As DateTime

    Public dtmMonthlySalesReport As DateTime
    Public dtmMonthlySalesTaxReport As DateTime
    Public dtmMonthlyInventoryReport As DateTime
    Public dtmMonthlyDepositReport As DateTime

    Public dtmYearlySalesReport As DateTime
    Public dtmYearlySalesTaxReport As DateTime
    Public dtmYearlyInventoryReport As DateTime
    Public dtmYearlyDepositReport As DateTime

    Public btmButtonDefault As Bitmap = My.Resources.Button
    Public btmButtonDefaultGray As Bitmap = My.Resources.Button
    Public btmButtonLittleLong As Bitmap = My.Resources.ButtonLittleLong
    Public btmButtonLittleLongGray As Bitmap = My.Resources.ButtonLittleLong
    Public btmButtonShort As Bitmap = My.Resources.ButtonShort
    Public btmButtonShortGray As Bitmap = My.Resources.ButtonShort
    Public btmSkinnyButton As Bitmap = My.Resources.SkinnyButton
    Public btmSkinnyButtonGray As Bitmap = My.Resources.SkinnyButton

End Module
