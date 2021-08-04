Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Net.Mail


Module modCommonUtilities

    ' Open a form and hide the current form
    Public Function OpenFormMaintainParent(ByRef frmSelf As Form, ByVal frmToOpen As Form) As DialogResult

        ' Init variables
        Dim dlgResult As DialogResult

        ' Make self invisible
        frmSelf.Visible = False

        ' Make new form
        dlgResult = frmToOpen.ShowDialog()

        ' Make self visible
        If (frmSelf.IsDisposed = False) Then
            frmSelf.Visible = True
        End If

        ' Return result
        Return dlgResult

    End Function

    ' Open a form and close the current form
    Public Function OpenFormKillParent(ByRef frmSelf As Form, ByVal frmToOpen As Form) As DialogResult

        ' Init variables
        Dim dlgResult As DialogResult

        ' Make new form
        frmToOpen.Show()

        ' Kill self
        frmSelf.Close()

        ' Return result
        Return dlgResult

    End Function

    Public Sub ButtonColor(ByVal pntPosition As Point, ByRef btnItemLookup As Button, ByVal frmMe As Form, ByVal btmButtonGray As Bitmap, ByVal btmButton As Bitmap, Optional ByVal intUpperOffset As Integer = -9, Optional ByVal intLowerOffset As Integer = -8)

        If (MouseIsHovering(pntPosition, btnItemLookup, frmMe, intUpperOffset, intLowerOffset) And frmMe.ContainsFocus = True) Then
            btnItemLookup.Image = btmButtonGray
        Else
            btnItemLookup.Image = btmButton
        End If

    End Sub

    Public Function MouseIsHovering(ByVal MousePosition As Point, ByRef ctlControl As Control, ByRef frmForm As Form, Optional ByVal intUpperOffset As Integer = -9, Optional ByVal intLowerOffset As Integer = -8)

        ' This if statement adapted from the Waveslash game launch I made for my latest game: https://gravityhamster.itch.io/waveslash
        If (MousePosition.X > ctlControl.Left + frmForm.Left And MousePosition.X < ctlControl.Right + frmForm.Left) And (MousePosition.Y > ctlControl.Top + frmForm.Top + ctlControl.Height + intUpperOffset And MousePosition.Y < ctlControl.Bottom + frmForm.Top + ctlControl.Height + intLowerOffset) Then
            Return True
        Else
            Return False
        End If

    End Function

    ' Clamp a value between two values
    Public Function Clamp(ByVal intValue As Integer, ByVal intMin As Integer, ByVal intMax As Integer) As Integer

        ' Clamp into range
        If (intValue > intMax) Then
            intValue = intMax
        ElseIf (intValue < intMin) Then
            intValue = intMin
        End If

        ' Return
        Return intValue

    End Function

    Public Function CreateSalesReport(ByRef frmMe As Form, ByVal strTimePeriod As String, ByVal blnQuiet As Boolean) As Boolean

        ' instantiate excel objects and declare variables
        ' THE EXCEL CODE IS BASED ON THIS TUTORIAL: https://www.tutorialspoint.com/vb.net/vb.net_excel_sheet.htm

        Dim ExcelApp As Excel.Application
        Dim ExcelWkBk As Excel.Workbook
        Dim ExcelWkSht As Excel.Worksheet
        Dim ExcelRange As Excel.Range
        Dim blnSuccess As Boolean = True

        ' start excel and get application object
        ExcelApp = CreateObject("Excel.Application")
        ExcelApp.Visible = False ' for testing only, set to false when go to prod

        ' Add a new workbook
        ExcelWkBk = ExcelApp.Workbooks.Add
        ExcelWkSht = ExcelWkBk.ActiveSheet

        Try

            ' add table headers going cell by cell
            ExcelWkSht.Cells(1, 1) = "Sales for the " & strTimePeriod
            ExcelWkSht.Cells(2, 1) = "Medals"
            ExcelWkSht.Cells(2, 2) = "Statues"
            ExcelWkSht.Cells(2, 3) = "Books"
            ExcelWkSht.Cells(2, 4) = "Church Goods"
            ExcelWkSht.Cells(2, 5) = "Tokens"
            ExcelWkSht.Cells(2, 6) = "Baptism"
            ExcelWkSht.Cells(2, 7) = "Rosary"
            ExcelWkSht.Cells(2, 8) = "Wedding"
            ExcelWkSht.Cells(2, 9) = "Anniversary"
            ExcelWkSht.Cells(2, 10) = "Cards"
            ExcelWkSht.Cells(2, 11) = "Holy Cards"
            ExcelWkSht.Cells(2, 12) = "Pictures/Artwork"
            ExcelWkSht.Cells(2, 13) = "Confirmation"
            ExcelWkSht.Cells(2, 14) = "First Communion"

            ' add table data
            ExcelWkSht.Cells(3, 1) = GetSales(frmMe, 1, strTimePeriod, blnQuiet)
            ExcelWkSht.Cells(3, 2) = GetSales(frmMe, 2, strTimePeriod, blnQuiet)
            ExcelWkSht.Cells(3, 3) = GetSales(frmMe, 3, strTimePeriod, blnQuiet)
            ExcelWkSht.Cells(3, 4) = GetSales(frmMe, 4, strTimePeriod, blnQuiet)
            ExcelWkSht.Cells(3, 5) = GetSales(frmMe, 5, strTimePeriod, blnQuiet)
            ExcelWkSht.Cells(3, 6) = GetSales(frmMe, 6, strTimePeriod, blnQuiet)
            ExcelWkSht.Cells(3, 7) = GetSales(frmMe, 7, strTimePeriod, blnQuiet)
            ExcelWkSht.Cells(3, 8) = GetSales(frmMe, 8, strTimePeriod, blnQuiet)
            ExcelWkSht.Cells(3, 9) = GetSales(frmMe, 9, strTimePeriod, blnQuiet)
            ExcelWkSht.Cells(3, 10) = GetSales(frmMe, 10, strTimePeriod, blnQuiet)
            ExcelWkSht.Cells(3, 11) = GetSales(frmMe, 11, strTimePeriod, blnQuiet)
            ExcelWkSht.Cells(3, 12) = GetSales(frmMe, 12, strTimePeriod, blnQuiet)
            ExcelWkSht.Cells(3, 13) = GetSales(frmMe, 13, strTimePeriod, blnQuiet)
            ExcelWkSht.Cells(3, 14) = GetSales(frmMe, 14, strTimePeriod, blnQuiet)

            Dim strFile As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + "\SalesReport.xlsx"

            ' Save
            If (My.Computer.FileSystem.FileExists("SalesReport.xlsx") = True) Then
                My.Computer.FileSystem.DeleteFile("SalesReport.xlsx")
            End If
            ExcelWkSht.SaveAs(strFile)

        Catch excError As Exception

            blnSuccess = False
            If (blnQuiet = False) Then
                ' Log and display error message
                MessageBox.Show(excError.Message)
            Else
                ' Log and display error message
                Console.WriteLine(excError.Message)
            End If

        End Try

        ' Release object references.
        ExcelWkBk.Saved = True
        ExcelApp.Workbooks.Close()
        ExcelApp.Quit()
        ExcelApp = Nothing
        ExcelRange = Nothing
        ExcelWkSht = Nothing
        ExcelWkBk = Nothing

        Return blnSuccess

    End Function

    Public Function GetSales(ByRef frmMe As Form, ByVal intCategory As Integer, ByVal strTimePeriod As String, ByVal blnQuiet As Boolean)

        Dim dblTotalSales As Double

        'Try

        Dim strSelect As String
        Dim cmdSelect As OleDb.OleDbCommand
        Dim dt As DataTable = New DataTable

        ' Open the DB
        OpenDatabaseConnectionSQLServer(blnQuiet)
        'If OpenDatabaseConnectionSQLServer(blnQuiet) = False Then

        '    '' The database is not open
        '    'If blnQuiet = False Then
        '    '    MessageBox.Show(frmMe, "Database connection error." & vbNewLine &
        '    '                "The form will now close.",
        '    '                frmMe.Text + " Error",
        '    '                MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    '    ' Close the form/application
        '    '    ' frmMe.Close()
        '    'Else
        '    '    Console.WriteLine("Database connection error." & vbNewLine & "Report not generated.")
        '    'End If

        ' Build the select statement based on user-selected time period
        If strTimePeriod = "last day" Then
            strSelect = "SELECT SUM(decCurrentItemPrice) from vItems_Sold_By_Category_With_Date_And_Price where intCategoryID = " & intCategory & " AND dtTransactionDate > (DATEADD(DAY, -1, GETDATE()))"

        ElseIf strTimePeriod = "last week" Then
            strSelect = "Select SUM(decCurrentItemPrice) from vItems_Sold_By_Category_With_Date_And_Price where intCategoryID = " & intCategory & " and dtTransactionDate > (DATEADD(DAY, -7, GETDATE()))"

        ElseIf strTimePeriod = "last month (30 days)" Then
            strSelect = "SELECT SUM(decCurrentItemPrice) from vItems_Sold_By_Category_With_Date_And_Price where intCategoryID = " & intCategory & " and dtTransactionDate > (DATEADD(DAY, -30, GETDATE()))"

        ElseIf strTimePeriod = "last year (365 days)" Then
            strSelect = "SELECT SUM(decCurrentItemPrice) from vItems_Sold_By_Category_With_Date_And_Price where intCategoryID = " & intCategory & " and dtTransactionDate > (DATEADD(DAY, -365, GETDATE()))"

        End If

        ' Retrieve all the records 
        cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
        Dim objTotalSales As Object = cmdSelect.ExecuteScalar

        ' check for null entries (zeroes), set to zero
        If IsDBNull(objTotalSales) Then
            dblTotalSales = 0
        Else
            dblTotalSales = CDbl(objTotalSales)
        End If

        ' close the database connection
        CloseDatabaseConnection()

        'Catch excError As Exception

        '    If blnQuiet = False Then
        '        ' Log and display error message
        '        MessageBox.Show(excError.Message)
        '    Else
        '        ' Log message in console
        '        Console.WriteLine(excError.Message)
        '    End If

        'End Try

        Return dblTotalSales

    End Function

    Public Function RunTaxReport(ByRef frmMe As Form, ByVal blnQuiet As Boolean, ByVal strYear As String, ByVal strMonth As String, ByVal strDay As String) As Boolean

        Dim ExcelApp As Excel.Application
        Dim ExcelWkBk As Excel.Workbook
        Dim ExcelWkSht As Excel.Worksheet
        Dim ExcelRange As Excel.Range
        Dim blnSuccess As Boolean = True

        ' start excel and get application object
        ExcelApp = CreateObject("Excel.Application")
        ExcelApp.Visible = False ' for testing only, set to false when go to prod

        ' Add a new workbook
        ExcelWkBk = ExcelApp.Workbooks.Add
        ExcelWkSht = ExcelWkBk.ActiveSheet

        Try

            ' declare variables
            Dim strSelect As String
            Dim cmdSelect As OleDb.OleDbCommand
            'Dim drSourceTable As OleDb.OleDbDataReader
            Dim dt As DataTable = New DataTable
            Dim objResults As Object
            Dim dblOhioTaxRate As Double = 7.8
            Dim intGrossSalesQty As Integer
            Dim dblGrossSalesTotal As Double
            'Dim dblGrossSalesAvg As Double
            Dim intReturnsQty As Integer
            Dim dblReturnsTotal As Double
            'Dim dblReturnsAvg As Double
            Dim intNetSalesQty As Integer
            Dim dblNetSalesTotal As Double
            'Dim dblNetSalesAvg As Double
            Dim dblTaxesTotal As Double
            Dim intTicketTotalQty As Integer
            Dim dblTicketTotalTotal As Double
            'Dim dblTicketTotalAvg As Double
            Dim intPaymentsWithCash As Integer
            Dim intPaymentsWithCredit As Integer
            Dim dblCashAmount As Double
            Dim dblCreditAmount As Double
            Dim dblTotalCash As Double
            Dim dblTotalCredit As Double
            Dim intPayinQty As Integer
            Dim intPayoutQty As Integer
            Dim dblPayinAmount As Double
            Dim dblPayoutAmount As Double
            Dim dblTaxableSubtotal As Double

            ' Open the DB
            If OpenDatabaseConnectionSQLServer(blnQuiet) = False Then

                '' The database is not open
                'If blnQuiet = False Then
                '    MessageBox.Show(frmMe, "Database connection error." & vbNewLine &
                '                "The form will now close.",
                '                frmMe.Text + " Error",
                '                MessageBoxButtons.OK, MessageBoxIcon.Error)
                '    ' Close the form/application
                '    ' frmMe.Close()
                'Else
                '    Console.WriteLine("Database connection error." & vbNewLine & "Report not generated.")
                'End If

                Exit Try

            End If

            ' BUILD THE SELECT STATEMENTS

            ' GET GROSS SALES QUANTITY 
            strSelect = "SELECT COUNT(intTransactionID) FROM TTransactions WHERE intTransactionTypeID = 1 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                intGrossSalesQty = 0
            Else
                intGrossSalesQty = CDbl(objResults)
            End If



            ' GET GROSS SALES TOTAL 
            strSelect = "SELECT SUM(decTotalPrice + decSalesTax) FROM TTransactions WHERE intTransactionTypeID = 1 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                dblGrossSalesTotal = 0
            Else
                dblGrossSalesTotal = CDbl(objResults)
            End If


            ' GET RETURNS QUANTITY
            strSelect = "SELECT COUNT(intTransactionTypeID) FROM TTransactions WHERE intTransactionTypeID = 2 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                intReturnsQty = 0
            Else
                intReturnsQty = CDbl(objResults)
            End If


            ' GET RETURNS TOTAL 
            strSelect = "SELECT SUM(decTotalPrice + decSalesTax) FROM TTransactions WHERE intTransactionTypeID = 2 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                dblReturnsTotal = 0
            Else
                dblReturnsTotal = CDbl(objResults)
            End If


            ' GET NET SALES QUANTITY
            strSelect = "SELECT COUNT(intTransactionID) FROM TTransactions WHERE intTransactionTypeID = 1 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                intNetSalesQty = 0
            Else
                intNetSalesQty = CDbl(objResults)
            End If


            ' GET NET SALES TOTAL
            strSelect = "SELECT SUM(decTotalPrice) FROM TTransactions WHERE intTransactionTypeID = 1 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                dblNetSalesTotal = 0
            Else
                dblNetSalesTotal = CDbl(objResults)
            End If


            ' GET TOTAL TAXES
            strSelect = "SELECT SUM(decSalesTax) FROM TTransactions WHERE intTransactionTypeID = 1 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                dblTaxesTotal = 0
            Else
                dblTaxesTotal = CDbl(objResults)
            End If


            ' GET TICKET TOTAL QUANTITY
            strSelect = "SELECT COUNT(intTransactionID) FROM TTransactions WHERE intTransactionTypeID = 1 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                intTicketTotalQty = 0
            Else
                intTicketTotalQty = CDbl(objResults)
            End If


            ' GET NUM PAYMENTS WITH CASH
            strSelect = "SELECT COUNT(intPaymentTypeID) FROM TTransactions WHERE intPaymentTypeID = 1 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                intPaymentsWithCash = 0
            Else
                intPaymentsWithCash = CDbl(objResults)
            End If


            ' GET NUM PAYMENTS WITH CREDIT
            strSelect = "SELECT COUNT(intPaymentTypeID) FROM TTransactions WHERE intPaymentTypeID = 2 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                intPaymentsWithCredit = 0
            Else
                intPaymentsWithCredit = CDbl(objResults)
            End If


            ' GET CASH AMOUNT PAID
            strSelect = "SELECT SUM(decTotalPrice + decSalesTax) FROM TTransactions WHERE intPaymentTypeID = 1 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                dblCashAmount = 0
            Else
                dblCashAmount = CDbl(objResults)
            End If


            ' GET CREDIT AMOUNT PAID
            strSelect = "SELECT SUM(decTotalPrice + decSalesTax) FROM TTransactions WHERE intPaymentTypeID = 2 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                dblCreditAmount = 0
            Else
                dblCreditAmount = CDbl(objResults)
            End If


            ' GET NUM PAY-INS
            strSelect = "SELECT COUNT(intTransactionID) FROM TTransactions WHERE intTransactionTypeID = 7 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                intPayinQty = 0
            Else
                intPayinQty = CDbl(objResults)
            End If


            ' GET NUM PAY-OUTS
            strSelect = "SELECT COUNT(intTransactionID) FROM TTransactions WHERE intTransactionTypeID = 8 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                intPayoutQty = 0
            Else
                intPayoutQty = CDbl(objResults)
            End If


            ' GET PAY-IN TOTAL
            strSelect = "SELECT SUM(decTotalPrice) FROM TTransactions WHERE intTransactionTypeID = 7 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                dblPayinAmount = 0
            Else
                dblPayinAmount = CDbl(objResults)
            End If


            ' GET PAY-OUT TOTAL
            strSelect = "SELECT SUM(decTotalPrice) FROM TTransactions WHERE intTransactionTypeID = 8 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                dblPayoutAmount = 0
            Else
                dblPayoutAmount = CDbl(objResults)
            End If


            ' add table headers going cell by cell
            ExcelWkSht.Cells(1, 1) = "TRINITY CHURCH SUPPLY"
            ExcelWkSht.Cells(2, 1) = "5479 NORTH BEND RD."
            ExcelWkSht.Cells(3, 1) = "CINCINNATI, OH 45247"
            ExcelWkSht.Cells(1, 10) = "Store Summary"
            ExcelWkSht.Cells(2, 10) = System.DateTime.Now.ToString

            ExcelWkSht.Cells(5, 2) = "Quantity"
            ExcelWkSht.Cells(5, 3) = "Total"
            ExcelWkSht.Cells(5, 4) = "Average"
            ExcelWkSht.Cells(6, 1) = "GROSS SALES"
            ExcelWkSht.Cells(6, 2) = intGrossSalesQty
            ExcelWkSht.Cells(6, 3) = dblGrossSalesTotal
            ExcelWkSht.Cells(6, 4) = (dblGrossSalesTotal / intGrossSalesQty)
            ExcelWkSht.Cells(7, 1) = "Gross Returns"
            ExcelWkSht.Cells(7, 2) = intReturnsQty
            ExcelWkSht.Cells(7, 3) = dblReturnsTotal
            ExcelWkSht.Cells(7, 4) = (dblReturnsTotal / intReturnsQty)
            ExcelWkSht.Cells(8, 1) = "NET SALES"
            ExcelWkSht.Cells(8, 2) = intNetSalesQty
            ExcelWkSht.Cells(8, 3) = dblNetSalesTotal
            ExcelWkSht.Cells(8, 4) = (dblNetSalesTotal / intNetSalesQty)
            ExcelWkSht.Cells(9, 1) = "Taxes"
            ExcelWkSht.Cells(9, 3) = dblTaxesTotal
            ExcelWkSht.Cells(10, 1) = "TICKET TOTAL"
            ExcelWkSht.Cells(10, 2) = intTicketTotalQty
            dblTicketTotalTotal = dblNetSalesTotal + dblTaxesTotal
            ExcelWkSht.Cells(10, 3) = dblTicketTotalTotal
            ExcelWkSht.Cells(10, 4) = (dblTicketTotalTotal / intTicketTotalQty)

            ExcelWkSht.Cells(12, 1) = "PAYMENT TYPES"
            ExcelWkSht.Cells(12, 2) = "Quantity"
            ExcelWkSht.Cells(12, 3) = "Amount"
            ExcelWkSht.Cells(12, 4) = "Total Amount"
            ExcelWkSht.Cells(13, 1) = "Cash"
            ExcelWkSht.Cells(13, 2) = intPaymentsWithCash
            ExcelWkSht.Cells(13, 3) = dblCashAmount
            ExcelWkSht.Cells(13, 4) = dblTotalCash
            ExcelWkSht.Cells(14, 1) = "Credit Card"
            ExcelWkSht.Cells(14, 2) = intPaymentsWithCredit
            ExcelWkSht.Cells(14, 3) = dblCreditAmount
            ExcelWkSht.Cells(14, 4) = dblTotalCredit

            ExcelWkSht.Cells(16, 2) = "Quantity"
            ExcelWkSht.Cells(16, 3) = "Amount"
            ExcelWkSht.Cells(17, 1) = "Payins"
            ExcelWkSht.Cells(17, 2) = intPayinQty
            ExcelWkSht.Cells(17, 3) = dblPayinAmount
            ExcelWkSht.Cells(18, 1) = "Payouts"
            ExcelWkSht.Cells(18, 2) = intPayoutQty
            ExcelWkSht.Cells(18, 3) = dblPayoutAmount

            ExcelWkSht.Cells(20, 1) = "TAX CATEGORIES"
            ExcelWkSht.Cells(20, 2) = "Rate %"
            ExcelWkSht.Cells(20, 3) = "Taxable Subtotal"
            ExcelWkSht.Cells(21, 1) = "Ohio Tax"
            ExcelWkSht.Cells(21, 2) = dblOhioTaxRate
            ExcelWkSht.Cells(21, 3) = dblNetSalesTotal


            ' close the database connection
            CloseDatabaseConnection()

            Dim strFile As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + "\TaxReport.xlsx"

            ' Save
            If (My.Computer.FileSystem.FileExists("TaxReport.xlsx") = True) Then
                My.Computer.FileSystem.DeleteFile("TaxReport.xlsx")
            End If
            ExcelWkSht.SaveAs(strFile)

        Catch excError As Exception

            blnSuccess = False

            If blnQuiet = False Then
                ' Log and display error message
                MessageBox.Show(excError.Message)
            Else
                ' Log message in console
                Console.WriteLine(excError.Message)
            End If

        End Try

        ' Release object references.
        ExcelWkBk.Saved = True
        ExcelApp.Workbooks.Close()
        ExcelApp.Quit()
        ExcelApp = Nothing
        ExcelRange = Nothing
        ExcelWkSht = Nothing
        ExcelWkBk = Nothing

        Return blnSuccess

    End Function

    Public Function RunInventoryReport(ByRef frmMe As Form, ByVal blnQuiet As Boolean) As Boolean

        ' add table data

        ' instantiate excel objects and declare variables
        ' THE EXCEL CODE IS BASED ON THIS TUTORIAL: https://www.tutorialspoint.com/vb.net/vb.net_excel_sheet.htm
        Dim ExcelApp As Excel.Application
        Dim ExcelWkBk As Excel.Workbook
        Dim ExcelWkSht As Excel.Worksheet
        Dim ExcelRange As Excel.Range
        Dim intNumRecords As Integer
        Dim intIndex As Integer = 2 ' starts at 2 to account for header row, Excel rows are also 1-based
        Dim intRecordIndex As Integer = 0
        Dim blnSuccess As Boolean = True

        ' start excel and get application object
        ExcelApp = CreateObject("Excel.Application")
        ExcelApp.Visible = False ' for testing only, set to false when go to prod

        ' Add a new workbook
        ExcelWkBk = ExcelApp.Workbooks.Add
        ExcelWkSht = ExcelWkBk.ActiveSheet

        Try

            ' add table headers
            ExcelWkSht.Cells(1, 1) = "SKU"
            ExcelWkSht.Cells(1, 2) = "Item Name"
            ExcelWkSht.Cells(1, 3) = "Item Description"
            ExcelWkSht.Cells(1, 4) = "Vendor Name"
            ExcelWkSht.Cells(1, 5) = "Retail Price"
            ExcelWkSht.Cells(1, 6) = "Current Inventory"
            ExcelWkSht.Cells(1, 7) = "Safety Stock"
            ExcelWkSht.Cells(1, 8) = "UPC"

            ' Init select statement string
            Dim strSelect As String = ""
            ' Init select statement Db command
            Dim cmdSelect As OleDb.OleDbCommand
            ' Init data reader
            Dim drSourceTable As OleDb.OleDbDataReader
            ' Init data table
            Dim dt As DataTable = New DataTable

            ' Open the DB
            If OpenDatabaseConnectionSQLServer(blnQuiet) = False Then

                'If (blnQuiet = False) Then
                '    ' The database is not open
                '    MessageBox.Show(frmMe, "Database connection error." & vbNewLine &
                '                "The form will now close.",
                '                frmMe.Text + " Error",
                '                MessageBoxButtons.OK, MessageBoxIcon.Error)

                '    ' Close the form/application
                '    ' frmMe.Close()
                'Else
                '    Console.WriteLine("Database connection error." & vbNewLine & "Report not generated.")
                'End If

                Exit Try

            End If

            ' Build the select statement based on user-selected time period
            strSelect = "SELECT strSKU, strItemName, strItemDesc, strVendorName, decItemPrice, intInventoryAmt, intSafetyStockAmt, strUPC FROM TItems, TVendors WHERE intInventoryAmt <= intSafetyStockAmt and TItems.intVendorID = TVendors.intVendorID"

            ' Retrieve all the records 
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader

            ' load table from data reader
            dt.Load(drSourceTable)

            ' add data to excel spreadsheet
            intNumRecords = dt.Rows.Count

            While intIndex <= (intNumRecords + 1)

                ExcelWkSht.Cells(intIndex, 1).Value = dt.Rows.Item(intRecordIndex).ItemArray(0)
                ExcelWkSht.Cells(intIndex, 2).Value = dt.Rows.Item(intRecordIndex).ItemArray(1)
                ExcelWkSht.Cells(intIndex, 3).Value = dt.Rows.Item(intRecordIndex).ItemArray(2)
                ExcelWkSht.Cells(intIndex, 4).Value = dt.Rows.Item(intRecordIndex).ItemArray(3)
                ExcelWkSht.Cells(intIndex, 5).Value = dt.Rows.Item(intRecordIndex).ItemArray(4)
                ExcelWkSht.Cells(intIndex, 6).Value = dt.Rows.Item(intRecordIndex).ItemArray(5)
                ExcelWkSht.Cells(intIndex, 7).Value = dt.Rows.Item(intRecordIndex).ItemArray(6)
                ExcelWkSht.Cells(intIndex, 8).Value = dt.Rows.Item(intRecordIndex).ItemArray(7).ToString() & "​" ' <- Whitespace character to trick excel formatting. DO NOT DELETE THIS
                intIndex += 1
                intRecordIndex += 1

            End While

            ' close the database connection
            CloseDatabaseConnection()

            Dim strFile As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + "\InventoryReport.xlsx"

            ' Save
            If (My.Computer.FileSystem.FileExists("InventoryReport.xlsx") = True) Then
                My.Computer.FileSystem.DeleteFile("InventoryReport.xlsx")
            End If
            ExcelWkSht.SaveAs(strFile)

        Catch excError As Exception

            blnSuccess = False

            If (blnQuiet = False) Then
                ' Log and display error message
                MessageBox.Show(excError.Message)
            Else
                Console.WriteLine(excError.Message)
            End If

        End Try

        ' Release object references.
        ExcelWkBk.Saved = True
        ExcelApp.Workbooks.Close()
        ExcelApp.Quit()
        ExcelApp = Nothing
        ExcelRange = Nothing
        ExcelWkSht = Nothing
        ExcelWkBk = Nothing

        Return blnSuccess

    End Function

    Public Function RunCashCreditReport(ByRef frmMe As Form, ByVal blnQuiet As Boolean, ByVal strYear As String, ByVal strMonth As String, ByVal strDay As String) As Boolean

        Dim ExcelApp As Excel.Application
        Dim ExcelWkBk As Excel.Workbook
        Dim ExcelWkSht As Excel.Worksheet
        Dim ExcelRange As Excel.Range
        Dim blnSuccess As Boolean = True

        ' start excel and get application object
        ExcelApp = CreateObject("Excel.Application")
        ExcelApp.Visible = False

        ' instantiate excel objects and declare variables
        ' THE EXCEL CODE IS BASED ON THIS TUTORIAL: https://www.tutorialspoint.com/vb.net/vb.net_excel_sheet.htm
        Dim objResults As Object
        Dim dblCashDeposit As Double
        Dim dblCreditDeposit As Double
        Dim strSelect As String
        Dim cmdSelect As OleDb.OleDbCommand

        ' Add a new workbook
        ExcelWkBk = ExcelApp.Workbooks.Add
        ExcelWkSht = ExcelWkBk.ActiveSheet

        ' add table data
        Try

            ' Open the DB
            If OpenDatabaseConnectionSQLServer(blnQuiet) = False Then

                'If (blnQuiet = False) Then
                '    ' The database is not open
                '    MessageBox.Show(frmMe, "Database connection error." & vbNewLine &
                '                "The form will now close.",
                '                frmMe.Text + " Error",
                '                MessageBoxButtons.OK, MessageBoxIcon.Error)

                '    ' Close the form/application
                '    ' frmMe.Close()
                'Else
                '    Console.WriteLine("Database connection error." & vbNewLine & "Report not generated.")
                'End If

                Exit Try

            End If

            ' build the select statements
            ' get cash deposits
            strSelect = "SELECT SUM(decTotalPrice + decSalesTax) from TTransactions WHERE intTransactionTypeID = 1 AND intPaymentTypeID = 1 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                dblCashDeposit = 0
            Else
                dblCashDeposit = CDbl(objResults)
            End If


            ' get credit deposits
            strSelect = "SELECT SUM(decTotalPrice + decSalesTax) from TTransactions WHERE intTransactionTypeID = 1 AND intPaymentTypeID = 2 AND (DATEPART(yyyy, dtTransactionDate) = ? AND DATEPART(MM, dtTransactionDate) = ? AND DATEPART(DD, dtTransactionDate) = ?)"
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            cmdSelect.Parameters.AddWithValue("dtTransactionYear", strYear)
            cmdSelect.Parameters.AddWithValue("dtTransactionMonth", strMonth)
            cmdSelect.Parameters.AddWithValue("dtTransactionDay", strDay)
            objResults = cmdSelect.ExecuteScalar
            If IsDBNull(objResults) Then
                dblCreditDeposit = 0
            Else
                dblCreditDeposit = CDbl(objResults)
            End If


            ' add table headers and data
            ExcelWkSht.Cells(1, 1) = "Cash Deposit for " & strYear & "-" & strMonth & "-" & strDay
            ExcelWkSht.Cells(1, 2) = "Credit Deposit for " & strYear & "-" & strMonth & "-" & strDay
            ExcelWkSht.Cells(2, 1) = dblCashDeposit
            ExcelWkSht.Cells(2, 2) = dblCreditDeposit

            Dim strFile As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + "\CashCreditDepositReport.xlsx"

            ' Save
            If (My.Computer.FileSystem.FileExists("CashCreditDepositReport.xlsx") = True) Then
                My.Computer.FileSystem.DeleteFile("CashCreditDepositReport.xlsx")
            End If

            ExcelWkSht.SaveAs(strFile)

        Catch excError As Exception

            blnSuccess = False

            If (blnQuiet = False) Then
                ' Log and display error message
                MessageBox.Show(excError.Message)
            Else
                Console.WriteLine(excError.Message)
            End If

        End Try

        ' Release object references.
        ExcelWkBk.Saved = True
        ExcelApp.Workbooks.Close()
        ExcelApp.Quit()
        ExcelApp = Nothing
        ExcelRange = Nothing
        ExcelWkSht = Nothing
        ExcelWkBk = Nothing

        Return blnSuccess

    End Function

    Public Sub ReadCSVFile()

        Try
            Dim intY As Integer
            Dim intReader As Integer = 0
            Dim strFile As String = My.Computer.FileSystem.ReadAllText("ReportFlags.csv")

            ' Reset CSV array
            aastrCSVFile = {
                {"", ""},
                {"", ""},
                {"", ""},
                {"", ""},
                {"", ""},
                {"", ""},
                {"", ""},
                {"", ""},
                {"", ""},
                {"", ""},
                {"", ""},
                {"", ""},
                {"", ""},
                {"", ""},
                {"", ""},
                {"", ""}
            }

            For intY = 0 To 15

                ' Read until comma
                While (strFile.Chars(intReader) <> ",")
                    aastrCSVFile(intY, 0) += strFile.Chars(intReader)
                    intReader += 1
                End While
                intReader += 1

                ' Read until NewLine
                While (strFile.Chars(intReader) <> vbCr)
                    aastrCSVFile(intY, 1) += strFile.Chars(intReader)
                    intReader += 1
                End While
                intReader += 2
            Next

        Catch excError As Exception

            Console.WriteLine("ERROR: Could not read CSV file. " & excError.Message)

            ' Attempt to delete the CSV file
            Try

                My.Computer.FileSystem.DeleteFile("ReportFlags.csv")

                ' Re-write
                If (My.Computer.FileSystem.FileExists("ReportFlags.csv") = False) Then

                    ' Reset the local file
                    aastrCSVFile = {
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

                    WriteCSVFile()

                    ' Re-read
                    If (My.Computer.FileSystem.FileExists("ReportFlags.csv") = True) Then

                        ReadCSVFile()

                    End If

                End If

            Catch excNestError As Exception

                Console.WriteLine("ERROR: Could not delete CSV file. " & excError.Message)

            End Try

        End Try



    End Sub

    Public Sub WriteCSVFile()

        Try

            Dim strWriteLine As String =
            aastrCSVFile(0, 0) & "," & aastrCSVFile(0, 1) & vbNewLine &
            aastrCSVFile(1, 0) & "," & aastrCSVFile(1, 1) & vbNewLine &
            aastrCSVFile(2, 0) & "," & aastrCSVFile(2, 1) & vbNewLine &
            aastrCSVFile(3, 0) & "," & aastrCSVFile(3, 1) & vbNewLine &
            aastrCSVFile(4, 0) & "," & aastrCSVFile(4, 1) & vbNewLine &
            aastrCSVFile(5, 0) & "," & aastrCSVFile(5, 1) & vbNewLine &
            aastrCSVFile(6, 0) & "," & aastrCSVFile(6, 1) & vbNewLine &
            aastrCSVFile(7, 0) & "," & aastrCSVFile(7, 1) & vbNewLine &
            aastrCSVFile(8, 0) & "," & aastrCSVFile(8, 1) & vbNewLine &
            aastrCSVFile(9, 0) & "," & aastrCSVFile(9, 1) & vbNewLine &
            aastrCSVFile(10, 0) & "," & aastrCSVFile(10, 1) & vbNewLine &
            aastrCSVFile(11, 0) & "," & aastrCSVFile(11, 1) & vbNewLine &
            aastrCSVFile(12, 0) & "," & aastrCSVFile(12, 1) & vbNewLine &
            aastrCSVFile(13, 0) & "," & aastrCSVFile(13, 1) & vbNewLine &
            aastrCSVFile(14, 0) & "," & aastrCSVFile(14, 1) & vbNewLine &
            aastrCSVFile(15, 0) & "," & aastrCSVFile(15, 1) & vbNewLine

            My.Computer.FileSystem.WriteAllText("ReportFlags.csv", strWriteLine, False)

        Catch excError As Exception
            Console.WriteLine("ERROR: Could not write CSV file. " & excError.Message)
        End Try



    End Sub

    ' Send Mail Function copied from: http://vb.net-informations.com/communications/vb.net_smtp_mail.htm
    Public Function SendMail(strTO As String, strFrom As String, strSubject As String, strBody As String, strUsername As String, strPassword As String, strAttachmentPath As String, blnQuiet As Boolean)
        Try
            Dim SmtpServer As New SmtpClient()
            Dim mail As New MailMessage()

            ' Add Email ID and Password of gmail account.
            ' This will be used to send the email from
            ' In this GMail account one need to  turn on from setting -> Allow less Secure App

            SmtpServer.Port = 587
            SmtpServer.Host = "smtp.gmail.com"
            ' Citation start: https://stackoverflow.com/questions/13424096/contact-form-is-not-sending-email-to-my-gmail-acount
            SmtpServer.EnableSsl = True
            SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network
            SmtpServer.UseDefaultCredentials = False
            SmtpServer.Credentials = New Net.NetworkCredential(strUsername, strPassword)
            ' Citation end.
            mail = New MailMessage()
            mail.From = New MailAddress(strFrom)
            mail.To.Add(strTO)
            mail.Subject = strSubject
            mail.Body = strBody
            ' Add Attachment
            ' Code from: http://vb.net-informations.com/communications/vb-email-attachment.htm
            ' Reference for excel: https://www.tutorialspoint.com/vb.net/vb.net_excel_sheet.htm
            ' Reference for PDF And Easy Report RDLC: https://www.youtube.com/watch?v=HX8hG29s3r8
            Dim attachment As System.Net.Mail.Attachment
            attachment = New System.Net.Mail.Attachment(strAttachmentPath)
            mail.Attachments.Add(attachment)

            SmtpServer.Send(mail)

            SmtpServer.Dispose()
            mail.Dispose()

            If (blnQuiet = False) Then
                MsgBox("Message sent")
            End If
            Return 0
        Catch ex As Exception
            If (blnQuiet = False) Then
                MsgBox(ex.ToString)
            End If
            Return ex.Message.Length
        End Try
    End Function

End Module
