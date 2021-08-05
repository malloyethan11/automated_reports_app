

Public Class frmMain

    Dim intElapsedTimeUpdateConfig As Integer = 1000000
    Dim intElapsedTimeRunReport As Integer = 0
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        tmrCheckIfReportRun.Enabled = True
        tmrUpdateLocalConfig.Enabled = True

        If (My.Computer.FileSystem.FileExists("ReportFlags.csv") = False) Then

            ' create the record file if it doesn't exist
            WriteCSVFile()

        Else

            ' read in the CSV file
            ReadCSVFile()

        End If

    End Sub

    Private Sub frmMain_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        Me.Visible = False

    End Sub

    ' ┍━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┑
    ' │ From here down is the functionality to automate the reports │
    ' ┕━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┙

    Private Sub tmrUpdateLocalConfig_Tick(sender As Object, e As EventArgs) Handles tmrUpdateLocalConfig.Tick

        Try

            ' Elapse time
            intElapsedTimeUpdateConfig += 1
            ' Label1.Text = intElapsedTimeUpdateConfig

            ' Has X minutes elapsed?
            If (intElapsedTimeUpdateConfig >= 45) Then ' 60 * 5 * 2
                ' Get db update
                If OpenDatabaseConnectionSQLServer(True) = False Then

                    '' The database is not open
                    'MessageBox.Show(Me, "Database connection error." & vbNewLine &
                    '        "The form will now close.",
                    '        Me.Text + " Error",
                    '        MessageBoxButtons.OK, MessageBoxIcon.Error)

                    ' Close the form/application
                    ' Me.Close()

                    ' Reset
                    intElapsedTimeUpdateConfig = 0
                    Exit Sub

                End If

                ' Init select statement string
                Dim strSelect As String = ""
                ' Init select statement Db command
                Dim cmdSelect As OleDb.OleDbCommand
                ' Init data reader
                Dim drSourceTable As OleDb.OleDbDataReader
                ' Init data table
                Dim dt As Data.DataTable = New Data.DataTable

                ' Build the select statement
                strSelect = "SELECT * FROM TReports"

                ' Retrieve all the records 
                cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
                drSourceTable = cmdSelect.ExecuteReader
                ' Get data
                dt.Load(drSourceTable)

                ' Load from data table
                strEmailSalesReport = dt(0)("strTargetEmail").ToString
                strEmailSalesTaxReport = dt(2)("strTargetEmail").ToString
                strEmailInventoryReport = dt(1)("strTargetEmail").ToString
                strEmailDespositReport = dt(3)("strTargetEmail").ToString

                blnSalesReportDaily = dt(0)("blnDaily").ToString
                blnSalesReportMonthly = dt(0)("blnMonthly").ToString
                blnSalesReportWeekly = dt(0)("blnWeekly").ToString
                blnSalesReportYearly = dt(0)("blnYearly").ToString

                blnInventoryReportDaily = dt(1)("blnDaily").ToString
                blnInventoryReportMonthly = dt(1)("blnMonthly").ToString
                blnInventoryReportWeekly = dt(1)("blnWeekly").ToString
                blnInventoryReportYearly = dt(1)("blnYearly").ToString

                blnSalesTaxReportDaily = dt(2)("blnDaily").ToString
                blnSalesTaxReportMonthly = dt(2)("blnMonthly").ToString
                blnSalesTaxReportWeekly = dt(2)("blnWeekly").ToString
                blnSalesTaxReportYearly = dt(2)("blnYearly").ToString

                blnDepositReportDaily = dt(3)("blnDaily").ToString
                blnDepositReportMonthly = dt(3)("blnMonthly").ToString
                blnDepositReportWeekly = dt(3)("blnWeekly").ToString
                blnDepositReportYearly = dt(3)("blnYearly").ToString

                dtmDailySalesReport = Convert.ToDateTime(dt(0)("dtDailyReportDate").ToString())
                dtmDailySalesTaxReport = Convert.ToDateTime(dt(2)("dtDailyReportDate").ToString())
                dtmDailyInventoryReport = Convert.ToDateTime(dt(1)("dtDailyReportDate").ToString())
                dtmDailyDepositReport = Convert.ToDateTime(dt(3)("dtDailyReportDate").ToString())

                dtmWeeklySalesReport = Convert.ToDateTime(dt(0)("dtWeeklyReportDate").ToString())
                dtmWeeklySalesTaxReport = Convert.ToDateTime(dt(2)("dtWeeklyReportDate").ToString())
                dtmWeeklyInventoryReport = Convert.ToDateTime(dt(1)("dtWeeklyReportDate").ToString())
                dtmWeeklyDepositReport = Convert.ToDateTime(dt(3)("dtWeeklyReportDate").ToString())

                dtmMonthlySalesReport = Convert.ToDateTime(dt(0)("dtMonthlyReportDate").ToString())
                dtmMonthlySalesTaxReport = Convert.ToDateTime(dt(2)("dtMonthlyReportDate").ToString())
                dtmMonthlyInventoryReport = Convert.ToDateTime(dt(1)("dtMonthlyReportDate").ToString())
                dtmMonthlyDepositReport = Convert.ToDateTime(dt(3)("dtMonthlyReportDate").ToString())

                dtmYearlySalesReport = Convert.ToDateTime(dt(0)("dtYearlyReportDate").ToString())
                dtmYearlySalesTaxReport = Convert.ToDateTime(dt(2)("dtYearlyReportDate").ToString())
                dtmYearlyInventoryReport = Convert.ToDateTime(dt(1)("dtYearlyReportDate").ToString())
                dtmYearlyDepositReport = Convert.ToDateTime(dt(3)("dtYearlyReportDate").ToString())

                ' Reset
                intElapsedTimeUpdateConfig = 0
            End If
        Catch excError As Exception

            Console.WriteLine("Something went wrong in updating the configuration")

        End Try

    End Sub

    Private Sub tmrCheckIfReportRun_Tick(sender As Object, e As EventArgs) Handles tmrCheckIfReportRun.Tick

        Try

            ' Elapse time
            intElapsedTimeRunReport += 1

            ' Has 5 minutes elapsed?
            If (intElapsedTimeRunReport >= 1) Then

                ' Run sales report
                AutomateSalesReport()

                ' Reset
                intElapsedTimeRunReport = 0

            End If

        Catch excError As Exception

            Console.WriteLine("Something went wrong in generating the reports")

        End Try

    End Sub

    Private Sub AutomateSalesReport()

        ' Sales report
        SalesDaily()
        SalesWeekly()
        SalesMonthly()
        SalesYearly()

        ' Inventory report
        InventoryDaily()
        InventoryWeekly()
        InventoryMonthly()
        InventoryYearly()

        ' Sales tax report
        TaxDaily()
        TaxMonthly()
        TaxWeekly()
        TaxYearly()

        ' Sales tax report
        DepositDaily()
        DepositMonthly()
        DepositWeekly()
        DepositYearly()

        ' Generate sales
        GenerateSalesDaily()
        GenerateSalesWeekly()
        GenerateSalesMonthly()
        GenerateSalesYearly()

        ' Generate inventory
        GenerateInventoryDaily()
        GenerateInventoryWeekly()
        GenerateInventoryMonthly()
        GenerateInventoryYearly()

        ' Generate tax
        GenerateTaxDaily()
        GenerateTaxWeekly()
        GenerateTaxMonthly()
        GenerateTaxYearly()

        ' Generate deposit
        GenerateDepositDaily()
        GenerateDepositWeekly()
        GenerateDepositMonthly()
        GenerateDepositYearly()

    End Sub

#Region "Sales Automation"

    Private Sub SalesDaily()

        ' Run daily?
        If (blnSalesReportDaily = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("HH:mm")
            Dim strTarget As String = dtmDailySalesReport.ToString("HH:mm")

            ' Is the flag false?
            If (aastrCSVFile(0, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Turn on
                    SalesDailyFlag = True

                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(0, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(0, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateSalesDaily()

        If (SalesDailyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = CreateSalesReport(Me, "last day", True)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email sales report
                SendMail(strEmailSalesReport, "TeamBeesCapstone@gmail.com", "Daily Sales Report", "Automated message: See attached sales report.", "TeamBeesCapstone@gmail.com", "cincystate123", "SalesReport.xlsx", True)
            End If

            ' Turn off
            SalesDailyFlag = False

        End If

    End Sub

    Private Sub SalesWeekly()

        ' Run weekly?
        If (blnSalesReportWeekly = True) Then

            ' Time period
            Dim strNow As String = WeekdayName(Weekday(DateTime.Now)) & " " & DateTime.Now.ToString("HH:mm")
            Dim strTarget As String = WeekdayName(Weekday(dtmWeeklySalesReport)) & " " & dtmWeeklySalesReport.ToString("HH:mm")

            ' Is the flag false?
            If (aastrCSVFile(1, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Turn on
                    SalesWeeklyFlag = True

                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(1, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(1, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateSalesWeekly()

        If (SalesWeeklyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = CreateSalesReport(Me, "last week", True)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email sales report
                SendMail(strEmailSalesReport, "TeamBeesCapstone@gmail.com", "Weekly Sales Report", "Automated message: See attached sales report.", "TeamBeesCapstone@gmail.com", "cincystate123", "SalesReport.xlsx", True)
            End If

            ' Turn off
            SalesWeeklyFlag = False

        End If

    End Sub

    Private Sub SalesMonthly()

        ' Run weekly?
        If (blnSalesReportMonthly = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("dd HH:mm")
            Dim strTarget As String
            Dim intDaysInMonth As Integer = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month)

            ' Clamp the days within this month
            If intDaysInMonth < dtmMonthlySalesReport.Day Then
                strTarget = dtmMonthlySalesReport.ToString(intDaysInMonth & " HH:mm")
            Else
                strTarget = dtmMonthlySalesReport.ToString("dd HH:mm")
            End If

            ' Is the flag false?
            If (aastrCSVFile(2, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Turn on
                    SalesMonthlyFlag = True

                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(2, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(2, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateSalesMonthly()

        If (SalesMonthlyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = CreateSalesReport(Me, "last month (30 days)", True)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email sales report
                SendMail(strEmailSalesReport, "TeamBeesCapstone@gmail.com", "Monthly Sales Report", "Automated message: See attached sales report.", "TeamBeesCapstone@gmail.com", "cincystate123", "SalesReport.xlsx", True)
            End If

            ' Turn off
            SalesMonthlyFlag = False

        End If

    End Sub

    Private Sub SalesYearly()

        ' Run weekly?
        If (blnSalesReportYearly = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("MM/dd HH:mm")
            Dim strTarget As String
            Dim intDaysInMonth As Integer = DateTime.DaysInMonth(DateTime.Now.Year, dtmYearlySalesReport.Month)

            ' Clamp the days within this month
            If intDaysInMonth < dtmYearlySalesReport.Day Then
                strTarget = dtmYearlySalesReport.ToString("MM/" & intDaysInMonth & " HH:mm")
            Else
                strTarget = dtmYearlySalesReport.ToString("MM/dd HH:mm")
            End If

            ' Is the flag false?
            If (aastrCSVFile(3, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Turn on
                    SalesYearlyFlag = True

                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(3, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(3, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateSalesYearly()

        If (SalesYearlyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = CreateSalesReport(Me, "last year (365 days)", True)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email sales report
                SendMail(strEmailSalesReport, "TeamBeesCapstone@gmail.com", "Yearly Sales Report", "Automated message: See attached sales report.", "TeamBeesCapstone@gmail.com", "cincystate123", "SalesReport.xlsx", True)
            End If

            ' Turn off
            SalesYearlyFlag = False

        End If

    End Sub

#End Region

#Region "Inventory Automation"

    Private Sub InventoryDaily()

        ' Run daily?
        If (blnInventoryReportDaily = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("HH:mm")
            Dim strTarget As String = dtmDailyInventoryReport.ToString("HH:mm")

            ' Is the flag false?
            If (aastrCSVFile(4, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' On
                    InventoryDailyFlag = True
                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(4, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(4, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateInventoryDaily()

        If (InventoryDailyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = RunInventoryReport(Me, True)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email Inventory report
                SendMail(strEmailInventoryReport, "TeamBeesCapstone@gmail.com", "Daily Inventory Report", "Automated message: See attached Inventory report.", "TeamBeesCapstone@gmail.com", "cincystate123", "InventoryReport.xlsx", True)
            End If

            ' Turn off
            InventoryDailyFlag = False

        End If

    End Sub

    Private Sub InventoryWeekly()

        ' Run weekly?
        If (blnInventoryReportWeekly = True) Then

            ' Time period
            Dim strNow As String = WeekdayName(Weekday(DateTime.Now)) & " " & DateTime.Now.ToString("HH:mm")
            Dim strTarget As String = WeekdayName(Weekday(dtmWeeklyInventoryReport)) & " " & dtmWeeklyInventoryReport.ToString("HH:mm")

            ' Is the flag false?
            If (aastrCSVFile(5, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Turn on
                    InventoryWeeklyFlag = True
                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(5, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(5, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateInventoryWeekly()

        If (InventoryWeeklyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = RunInventoryReport(Me, True)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email Inventory report
                SendMail(strEmailInventoryReport, "TeamBeesCapstone@gmail.com", "Weekly Inventory Report", "Automated message: See attached Inventory report.", "TeamBeesCapstone@gmail.com", "cincystate123", "InventoryReport.xlsx", True)
            End If

            ' Turn off
            InventoryWeeklyFlag = False

        End If

    End Sub

    Private Sub InventoryMonthly()

        ' Run weekly?
        If (blnInventoryReportMonthly = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("dd HH:mm")
            Dim strTarget As String
            Dim intDaysInMonth As Integer = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month)

            ' Clamp the days within this month
            If intDaysInMonth < dtmMonthlyInventoryReport.Day Then
                strTarget = dtmMonthlyInventoryReport.ToString(intDaysInMonth & " HH:mm")
            Else
                strTarget = dtmMonthlyInventoryReport.ToString("dd HH:mm")
            End If

            ' Is the flag false?
            If (aastrCSVFile(6, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Turn on
                    InventoryMonthlyFlag = True
                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(6, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(6, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateInventoryMonthly()

        If (InventoryMonthlyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = RunInventoryReport(Me, True)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email Inventory report
                SendMail(strEmailInventoryReport, "TeamBeesCapstone@gmail.com", "Monthly Inventory Report", "Automated message: See attached Inventory report.", "TeamBeesCapstone@gmail.com", "cincystate123", "InventoryReport.xlsx", True)
            End If

            ' Turn off
            InventoryMonthlyFlag = False

        End If

    End Sub

    Private Sub InventoryYearly()

        ' Run weekly?
        If (blnInventoryReportYearly = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("MM/dd HH:mm")
            Dim strTarget As String
            Dim intDaysInMonth As Integer = DateTime.DaysInMonth(DateTime.Now.Year, dtmYearlyInventoryReport.Month)

            ' Clamp the days within this month
            If intDaysInMonth < dtmYearlyInventoryReport.Day Then
                strTarget = dtmYearlyInventoryReport.ToString("MM/" & intDaysInMonth & " HH:mm")
            Else
                strTarget = dtmYearlyInventoryReport.ToString("MM/dd HH:mm")
            End If

            ' Is the flag false?
            If (aastrCSVFile(7, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Turn on
                    InventoryYearlyFlag = True
                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(7, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(7, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateInventoryYearly()

        If (InventoryYearlyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = RunInventoryReport(Me, True)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email Inventory report
                SendMail(strEmailInventoryReport, "TeamBeesCapstone@gmail.com", "Yearly Inventory Report", "Automated message: See attached Inventory report.", "TeamBeesCapstone@gmail.com", "cincystate123", "InventoryReport.xlsx", True)
            End If

            ' Turn off
            InventoryYearlyFlag = False

        End If

    End Sub

#End Region

#Region "Sales Tax Automation"

    Private Sub TaxDaily()

        ' Run daily?
        If (blnSalesTaxReportDaily = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("HH:mm")
            Dim strTarget As String = dtmDailySalesTaxReport.ToString("HH:mm")

            ' Is the flag false?
            If (aastrCSVFile(8, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    TaxDailyFlag = True
                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(8, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(8, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateTaxDaily()

        If (TaxDailyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = RunTaxReport(Me, True, DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Year)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email Tax report
                SendMail(strEmailSalesTaxReport, "TeamBeesCapstone@gmail.com", "Daily Tax Report", "Automated message: See attached Tax report.", "TeamBeesCapstone@gmail.com", "cincystate123", "TaxReport.xlsx", True)
            End If

            TaxDailyFlag = False

        End If

    End Sub

    Private Sub TaxWeekly()

        ' Run weekly?
        If (blnSalesTaxReportWeekly = True) Then

            ' Time period
            Dim strNow As String = WeekdayName(Weekday(DateTime.Now)) & " " & DateTime.Now.ToString("HH:mm")
            Dim strTarget As String = WeekdayName(Weekday(dtmWeeklySalesTaxReport)) & " " & dtmWeeklySalesTaxReport.ToString("HH:mm")

            ' Is the flag false?
            If (aastrCSVFile(9, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    TaxWeeklyFlag = True
                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(9, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(9, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateTaxWeekly()

        If (TaxWeeklyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = RunTaxReport(Me, True, DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Year)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email Tax report
                SendMail(strEmailSalesTaxReport, "TeamBeesCapstone@gmail.com", "Weekly Tax Report", "Automated message: See attached Tax report.", "TeamBeesCapstone@gmail.com", "cincystate123", "TaxReport.xlsx", True)
            End If

            TaxWeeklyFlag = False

        End If

    End Sub

    Private Sub TaxMonthly()

        ' Run weekly?
        If (blnSalesTaxReportMonthly = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("dd HH:mm")
            Dim strTarget As String
            Dim intDaysInMonth As Integer = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month)

            ' Clamp the days within this month
            If intDaysInMonth < dtmMonthlySalesTaxReport.Day Then
                strTarget = dtmMonthlySalesTaxReport.ToString(intDaysInMonth & " HH:mm")
            Else
                strTarget = dtmMonthlySalesTaxReport.ToString("dd HH:mm")
            End If

            ' Is the flag false?
            If (aastrCSVFile(10, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    TaxMonthlyFlag = True
                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(10, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(10, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateTaxMonthly()

        If (TaxMonthlyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = RunTaxReport(Me, True, DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Year)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email Tax report
                SendMail(strEmailSalesTaxReport, "TeamBeesCapstone@gmail.com", "Monthly Tax Report", "Automated message: See attached Tax report.", "TeamBeesCapstone@gmail.com", "cincystate123", "TaxReport.xlsx", True)
            End If

            TaxMonthlyFlag = False

        End If

    End Sub

    Private Sub TaxYearly()

        ' Run weekly?
        If (blnSalesTaxReportYearly = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("MM/dd HH:mm")
            Dim strTarget As String
            Dim intDaysInMonth As Integer = DateTime.DaysInMonth(DateTime.Now.Year, dtmYearlySalesTaxReport.Month)

            ' Clamp the days within this month
            If intDaysInMonth < dtmYearlySalesTaxReport.Day Then
                strTarget = dtmYearlySalesTaxReport.ToString("MM/" & intDaysInMonth & " HH:mm")
            Else
                strTarget = dtmYearlySalesTaxReport.ToString("MM/dd HH:mm")
            End If

            ' Is the flag false?
            If (aastrCSVFile(11, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    TaxYearlyFlag = True
                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(11, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(11, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateTaxYearly()

        If (TaxYearlyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = RunTaxReport(Me, True, DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Year)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email Tax report
                SendMail(strEmailSalesTaxReport, "TeamBeesCapstone@gmail.com", "Yearly Tax Report", "Automated message: See attached Tax report.", "TeamBeesCapstone@gmail.com", "cincystate123", "TaxReport.xlsx", True)
            End If

            TaxYearlyFlag = False

        End If

    End Sub

#End Region

#Region "Deposit Automation"

    Private Sub DepositDaily()

        ' Run daily?
        If (blnDepositReportDaily = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("HH:mm")
            Dim strTarget As String = dtmDailyDepositReport.ToString("HH:mm")

            ' Is the flag false?
            If (aastrCSVFile(12, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    DepositDailyFlag = True
                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(12, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(12, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateDepositDaily()

        If (DepositDailyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = RunCashCreditReport(Me, True, DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Year)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email Deposit report
                SendMail(strEmailDespositReport, "TeamBeesCapstone@gmail.com", "Daily Deposit Report", "Automated message: See attached Deposit report.", "TeamBeesCapstone@gmail.com", "cincystate123", "CashCreditDepositReport.xlsx", True)
            End If

            DepositDailyFlag = False

        End If

    End Sub

    Private Sub DepositWeekly()

        ' Run weekly?
        If (blnDepositReportWeekly = True) Then

            ' Time period
            Dim strNow As String = WeekdayName(Weekday(DateTime.Now)) & " " & DateTime.Now.ToString("HH:mm")
            Dim strTarget As String = WeekdayName(Weekday(dtmWeeklyDepositReport)) & " " & dtmWeeklyDepositReport.ToString("HH:mm")

            ' Is the flag false?
            If (aastrCSVFile(13, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    DepositWeeklyFlag = True
                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(13, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(13, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateDepositWeekly()

        If (DepositWeeklyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = RunCashCreditReport(Me, True, DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Year)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email Deposit report
                SendMail(strEmailDespositReport, "TeamBeesCapstone@gmail.com", "Weekly Deposit Report", "Automated message: See attached Deposit report.", "TeamBeesCapstone@gmail.com", "cincystate123", "CashCreditDepositReport.xlsx", True)
            End If

            DepositWeeklyFlag = False

        End If

    End Sub

    Private Sub DepositMonthly()

        ' Run weekly?
        If (blnDepositReportMonthly = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("dd HH:mm")
            Dim strTarget As String
            Dim intDaysInMonth As Integer = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month)

            ' Clamp the days within this month
            If intDaysInMonth < dtmMonthlyDepositReport.Day Then
                strTarget = dtmMonthlyDepositReport.ToString(intDaysInMonth & " HH:mm")
            Else
                strTarget = dtmMonthlyDepositReport.ToString("dd HH:mm")
            End If

            ' Is the flag false?
            If (aastrCSVFile(14, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    DepositMonthlyFlag = True
                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(14, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(14, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateDepositMonthly()

        If (DepositMonthlyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = RunCashCreditReport(Me, True, DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Year)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email Deposit report
                SendMail(strEmailDespositReport, "TeamBeesCapstone@gmail.com", "Monthly Deposit Report", "Automated message: See attached Deposit report.", "TeamBeesCapstone@gmail.com", "cincystate123", "CashCreditDepositReport.xlsx", True)
            End If

            DepositMonthlyFlag = False

        End If

    End Sub

    Private Sub DepositYearly()

        ' Run weekly?
        If (blnDepositReportYearly = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("MM/dd HH:mm")
            Dim strTarget As String
            Dim intDaysInMonth As Integer = DateTime.DaysInMonth(DateTime.Now.Year, dtmYearlyDepositReport.Month)

            ' Clamp the days within this month
            If intDaysInMonth < dtmYearlyDepositReport.Day Then
                strTarget = dtmYearlyDepositReport.ToString("MM/" & intDaysInMonth & " HH:mm")
            Else
                strTarget = dtmYearlyDepositReport.ToString("MM/dd HH:mm")
            End If

            ' Is the flag false?
            If (aastrCSVFile(15, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    DepositYearlyFlag = True
                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(15, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(15, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub GenerateDepositYearly()

        If (DepositYearlyFlag = True) Then

            ' Okay, we're in. Run the report in quiet mode
            Dim blnResult As Boolean = RunCashCreditReport(Me, True, DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Year)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            If (blnResult = True) Then
                ' email Deposit report
                SendMail(strEmailDespositReport, "TeamBeesCapstone@gmail.com", "Yearly Deposit Report", "Automated message: See attached Deposit report.", "TeamBeesCapstone@gmail.com", "cincystate123", "CashCreditDepositReport.xlsx", True)
            End If

            DepositYearlyFlag = False

        End If

    End Sub

#End Region

End Class
