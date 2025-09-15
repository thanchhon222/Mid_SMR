Imports System.Data.OleDb
Imports System.Reflection.Emit
Imports System.Text
Imports ACCPAC.Advantage
Imports AccpacCOMAPI
Imports Mid_SMR.Read
Imports Mid_SMR.SQLGetData
Public Class Internal_Usage

    Public Shared SageSess As New ACCPAC.Advantage.Session
    Public Property Username = frmLogin.TextBox1.Text
    Public Property Password = frmLogin.TextBox2.Text

    Private Sub btnPostSage_Click(sender As Object, e As EventArgs) Handles btnPostSage.Click

        If dgDetail.RowCount <= 0 Then
            MsgBox("No transactions found.", MsgBoxStyle.Information, "Alert")
            Exit Sub
        End If
        Try
            ' Validate data before posting
            If Not ValidateSageData() Then
                Exit Sub
            End If
            ' Check stock availability right before posting
            If Not CheckStockAvailability() Then
                MessageBox.Show("Cannot post to Sage due to insufficient stock. Please adjust quantities or locations and try again.",
                          "Insufficient Stock",
                          MessageBoxButtons.OK,
                          MessageBoxIcon.Error)
                Exit Sub
            End If

            SageSess.Init("", "XY", "XY1000", "72A")
            SageSess.Open(Username, Password, ReadConnectionNotedPage.DBName_St, Today, 0)

            '   FFDsageSess.Open(UCase(frmLogin.txt_UserName.Text.Trim()), UCase(frmLogin.txt_frmLogin.Text.Trim()), (Form1.TextBox2.Text.Trim()), Today, 0)
            Dim mDBLinkCmpRW As DBLink
            mDBLinkCmpRW = SageSess.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)

            Dim ICICE1header As View = mDBLinkCmpRW.OpenView("IC0288")
            Dim ICICE1detail1 As View = mDBLinkCmpRW.OpenView("IC0286")
            Dim ICICE1detail2 As View = mDBLinkCmpRW.OpenView("IC0289")
            Dim ICICE1detail3 As View = mDBLinkCmpRW.OpenView("IC0287")
            Dim ICICE1detail4 As View = mDBLinkCmpRW.OpenView("IC0282")
            Dim ICICE1detail5 As View = mDBLinkCmpRW.OpenView("IC0284")

            ICICE1header.Compose(New View() {ICICE1detail1, ICICE1detail2})
            ICICE1detail1.Compose(New View() {ICICE1header, ICICE1detail3, ICICE1detail5, Nothing, Nothing, Nothing, Nothing, Nothing, ICICE1detail4})
            ICICE1detail2.Compose(New View() {ICICE1header})
            ICICE1detail3.Compose(New View() {ICICE1detail1})
            ICICE1detail4.Compose(New View() {ICICE1detail1})
            ICICE1detail5.Compose(New View() {ICICE1detail1})
            ICICE1header.Init()


            With ICICE1header.Fields
                .FieldByName("TRANSDATE").SetValue(dt1.Value.ToString("dd-MMM-yy"), True)
                .FieldByName("HDRDESC").SetValue(TextBox3.Text.Trim(), False)
                .FieldByName("REFERENCE").SetValue(TextBox2.Text.Trim(), False)
                .FieldByName("EMPLOYEENO").SetValue(TextBox4.Text.Trim(), False)
            End With
            With ICICE1detail1.Fields
                Dim i As Integer

                For i = 0 To dgDetail.Rows.Count - 1

                    ICICE1detail1.RecordCreate(0)

                    'Dim itemCode As String = dgDetail.Rows(i).Cells(1).Value?.ToString()?.Trim()
                    'Dim glAccount As String = dgDetail.Rows(i).Cells(4).Value?.ToString()?.Trim()

                    ' itemCode = itemCode.Replace("-", "")
                    ' glAccount = glAccount.Replace("-", "")

                    '.FieldByName("ITEMNO").SetValue(itemCode, False)
                    .FieldByName("ITEMNO").SetValue(dgDetail.Rows(i).Cells(1).Value.ToString, False)
                    .FieldByName("LOCATION").SetValue(dgDetail.Rows(i).Cells(3).Value.ToString, False)     ' Location
                    .FieldByName("GLACCT").SetValue(dgDetail.Rows(i).Cells(4).Value.ToString, False)
                    '.FieldByName("GLACCT").SetValue(glAccount, False)
                    .FieldByName("QUANTITY").SetValue(dgDetail.Rows(i).Cells(5).Value.ToString, False)
                    .FieldByName("COMMENTS").SetValue(dgDetail.Rows(i).Cells(6).Value.ToString, False)

                    ICICE1detail1.Process()
                    ICICE1detail1.Insert()

                Next
                'ICICE1detail1.Insert()
            End With
            ICICE1header.Fields.FieldByName("STATUS").SetValue("1", False)             ' Record Status
            ICICE1header.Insert()

            MsgBox("Import internal usage completed.", MsgBoxStyle.Information, "Alert")

            dgDetail.DataSource = Nothing
            dgDetail.Columns.Clear()
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""

        Catch ex As Exception
            HandleACCPACError(ex)
        Finally
            'If SageSess IsNot Nothing AndAlso SageSess.IsActive Then
            '    SageSess.Close()
            'End If
        End Try
    End Sub
    Private Sub HandleACCPACError(ex As Exception)
        If SageSess IsNot Nothing AndAlso SageSess.Errors IsNot Nothing Then
            Dim errorMessage As New StringBuilder()
            errorMessage.AppendLine("ACCPAC Error Details:")

            For i As Integer = 0 To SageSess.Errors.Count - 1
                errorMessage.AppendLine($"Error {i + 1}: {SageSess.Errors.Item(i).Message}")
            Next

            MessageBox.Show(errorMessage.ToString(), "ACCPAC Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            MessageBox.Show($"Error: {ex.Message}", "General Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Function ValidateSageData() As Boolean
        Dim missingItems As New List(Of String)
        Dim missingAccounts As New List(Of String)
        Dim isValid As Boolean = True

        ' Clear previous highlighting
        For Each row As DataGridViewRow In dgDetail.Rows
            If Not row.IsNewRow Then
                row.DefaultCellStyle.BackColor = Color.White
            End If
        Next

        Try
            ' Check each row in the DataGridView
            For i As Integer = 0 To dgDetail.Rows.Count - 1
                Dim row As DataGridViewRow = dgDetail.Rows(i)

                ' Skip empty rows
                If row.IsNewRow Then Continue For

                Dim rowNumber As Integer = i + 1 ' Convert to 1-based indexing for user display
                Dim hasError As Boolean = False

                ' Check ItemCode (assuming column index 1 for ItemCode)
                Dim itemCode As String = If(row.Cells(1).Value IsNot Nothing, row.Cells(1).Value.ToString().Trim(), "")
                If Not String.IsNullOrEmpty(itemCode) Then
                    ' Remove hyphens from ItemCode for Sage comparison
                    Dim sageItemCode As String = itemCode.Replace("-", "")
                    Dim itemCheckDt As DataTable = SQL_GetTable_St("SELECT ITEMNO FROM ICITEM WHERE ITEMNO = '" & sageItemCode.Replace("'", "''") & "'")
                    If itemCheckDt.Rows.Count = 0 Then
                        missingItems.Add($"Row {rowNumber}: {itemCode} (Sage expects: {sageItemCode})")
                        hasError = True
                        isValid = False
                    End If
                Else
                    missingItems.Add($"Row {rowNumber}: (Empty ItemCode)")
                    hasError = True
                    isValid = False
                End If

                ' Check Account Number (assuming column index 4 for account)
                Dim accountNumber As String = If(row.Cells(4).Value IsNot Nothing, row.Cells(4).Value.ToString().Trim(), "")
                If Not String.IsNullOrEmpty(accountNumber) Then
                    ' Remove hyphens from Account Number for Sage comparison
                    Dim sageAccountNumber As String = accountNumber.Replace("-", "")
                    Dim glCheckDt As DataTable = SQL_GetTable_St("SELECT ACCTID FROM GLAMF WHERE ACCTID = '" & sageAccountNumber.Replace("'", "''") & "'")
                    If glCheckDt.Rows.Count = 0 Then
                        missingAccounts.Add($"Row {rowNumber}: {accountNumber} (Sage expects: {sageAccountNumber})")
                        hasError = True
                        isValid = False
                    End If
                Else
                    missingAccounts.Add($"Row {rowNumber}: (Empty Account)")
                    hasError = True
                    isValid = False
                End If

                ' Highlight row if it has errors
                If hasError Then
                    row.DefaultCellStyle.BackColor = Color.LightPink
                End If
            Next

            ' Show error messages if any missing items/accounts found
            If missingItems.Count > 0 OrElse missingAccounts.Count > 0 Then
                Dim errorMessage As New StringBuilder()
                errorMessage.AppendLine("Validation failed! Please correct the following issues:")

                If missingItems.Count > 0 Then
                    errorMessage.AppendLine()
                    errorMessage.AppendLine("Missing Item Codes in Sage:")
                    For Each item In missingItems
                        errorMessage.AppendLine($"- {item}")
                    Next
                End If

                If missingAccounts.Count > 0 Then
                    errorMessage.AppendLine()
                    errorMessage.AppendLine("Missing Chart of Accounts in Sage:")
                    For Each account In missingAccounts
                        errorMessage.AppendLine($"- {account}")
                    Next
                End If

                errorMessage.AppendLine()
                errorMessage.AppendLine("Note: Rows with errors are highlighted in pink.")
                errorMessage.AppendLine("Sage stores codes without hyphens. Your data contains hyphens that will be automatically removed during posting.")

                MessageBox.Show(errorMessage.ToString(), "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            MessageBox.Show($"Error during validation: {ex.Message}", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            isValid = False
        End Try

        Return isValid
    End Function
    Private Function CheckStockAvailability() As Boolean
        Dim outOfStockItems As New List(Of String)
        Dim isValid As Boolean = True

        Try
            ' Clear previous stock highlighting
            For Each row As DataGridViewRow In dgDetail.Rows
                If Not row.IsNewRow Then
                    ' Only remove orange highlighting (keep pink for validation errors)
                    If row.DefaultCellStyle.BackColor = Color.Orange Then
                        row.DefaultCellStyle.BackColor = Color.White
                    End If
                End If
            Next

            ' Check each row in the DataGridView
            For i As Integer = 0 To dgDetail.Rows.Count - 1
                Dim row As DataGridViewRow = dgDetail.Rows(i)

                ' Skip empty rows
                If row.IsNewRow Then Continue For

                Dim rowNumber As Integer = i + 1 ' Convert to 1-based indexing for user display

                ' Get ItemCode and Location
                Dim itemCode As String = If(row.Cells(1).Value IsNot Nothing, row.Cells(1).Value.ToString().Trim(), "")
                Dim location As String = If(row.Cells(3).Value IsNot Nothing, row.Cells(3).Value.ToString().Trim(), "")
                Dim quantity As Decimal = 0
                Decimal.TryParse(If(row.Cells(5).Value IsNot Nothing, row.Cells(5).Value.ToString().Trim(), "0"), quantity)

                If Not String.IsNullOrEmpty(itemCode) AndAlso Not String.IsNullOrEmpty(location) AndAlso quantity > 0 Then
                    ' Remove hyphens from ItemCode for Sage comparison
                    Dim sageItemCode As String = itemCode.Replace("-", "")

                    ' Check stock availability in Sage
                    Dim stockQuery As String = "SELECT QTYONHAND FROM ICILOC " &
                                          "WHERE ITEMNO = '" & sageItemCode.Replace("'", "''") & "' " &
                                          "AND LOCATION = '" & location.Replace("'", "''") & "'"

                    Dim stockCheckDt As DataTable = SQL_GetTable_St(stockQuery)

                    If stockCheckDt.Rows.Count > 0 Then
                        Dim qtyOnHand As Decimal = 0
                        Decimal.TryParse(stockCheckDt.Rows(0)("QTYONHAND").ToString(), qtyOnHand)

                        If qtyOnHand < quantity Then
                            outOfStockItems.Add($"Row {rowNumber}: {itemCode} at {location} (Available: {qtyOnHand}, Required: {quantity})")
                            isValid = False

                            ' Highlight the row in orange for stock issues
                            row.DefaultCellStyle.BackColor = Color.Orange
                        End If
                    Else
                        ' Item-location combination not found
                        outOfStockItems.Add($"Row {rowNumber}: {itemCode} not found at location {location}")
                        isValid = False
                        row.DefaultCellStyle.BackColor = Color.Orange
                    End If
                End If
            Next

            ' Show error messages if any out-of-stock items found
            If outOfStockItems.Count > 0 Then
                Dim errorMessage As New StringBuilder()
                errorMessage.AppendLine("Insufficient stock! The following items cannot be posted:")

                For Each item In outOfStockItems
                    errorMessage.AppendLine($"- {item}")
                Next

                errorMessage.AppendLine()
                errorMessage.AppendLine("Rows with stock issues are highlighted in orange.")
                errorMessage.AppendLine("Please adjust quantities or locations before posting.")

                MessageBox.Show(errorMessage.ToString(), "Stock Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            MessageBox.Show($"Error during stock check: {ex.Message}", "Stock Check Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            isValid = False
        End Try

        Return isValid
    End Function

    Private Sub Internal_Usage_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For f As Integer = 0 To dgDetail.Rows.Count - 1
            Dim num As Integer = Val(dgDetail.Rows(f).Cells(11).Value)
            If num <= 0 Then
                dgDetail.Rows(f).DefaultCellStyle.BackColor = Color.Red
            End If
        Next

        dt1.Format = DateTimePickerFormat.Custom
        dt1.CustomFormat = "dd-MMM-yy"

        If dgDetail.Columns.Count > 0 Then
            SetupDataGridViewEditableColumns()
        End If
    End Sub

    ' Add these event handlers to your form

    ' Validate Quantity to ensure it's a positive number
    Private Sub dgDetail_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgDetail.CellValidating
        ' Check if it's the Quantity column (adjust index as needed)
        If e.ColumnIndex = 5 OrElse (dgDetail.Columns.Count > e.ColumnIndex AndAlso dgDetail.Columns(e.ColumnIndex).Name = "Quantity") Then
            Dim newValue As String = e.FormattedValue.ToString()
            If Not String.IsNullOrEmpty(newValue) Then
                Dim quantity As Decimal
                If Not Decimal.TryParse(newValue, quantity) OrElse quantity < 0 Then
                    MessageBox.Show("Quantity must be a positive number.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    e.Cancel = True
                End If
            End If
        End If
    End Sub

    ' Provide visual feedback when editing editable cells
    Private Sub dgDetail_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgDetail.CellBeginEdit
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim column As DataGridViewColumn = dgDetail.Columns(e.ColumnIndex)
            If Not column.ReadOnly Then
                ' Change background color when editing editable cells
                dgDetail.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.LightYellow
            Else
                ' Prevent editing of read-only cells
                e.Cancel = True
            End If
        End If
    End Sub

    ' Restore normal background color after editing
    Private Sub dgDetail_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgDetail.CellEndEdit
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            dgDetail.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.White
        End If
    End Sub

    ' Optional: Double-click to edit for better usability
    Private Sub dgDetail_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgDetail.CellDoubleClick
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim column As DataGridViewColumn = dgDetail.Columns(e.ColumnIndex)
            If Not column.ReadOnly Then
                dgDetail.BeginEdit(True)
            End If
        End If
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub


    Private Sub B_browse_Click(sender As Object, e As EventArgs) Handles B_browse.Click
        OpenFileDialog1.Filter = "Excel Files|*.xls;*.xlsx"

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName

            Try
                Dim connStr As String = ""
                Dim filePath As String = TextBox1.Text.Trim()

                ' Choose correct provider based on extension
                If filePath.EndsWith(".xls") Then
                    connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filePath & ";Extended Properties='Excel 8.0;HDR=YES;'"
                ElseIf filePath.EndsWith(".xlsx") Then
                    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES;'"
                Else
                    MessageBox.Show("Unsupported file type.")
                    Exit Sub
                End If

                ' Get sheet name
                Dim sheetName As String = ""
                Using conn As New OleDbConnection(connStr)
                    conn.Open()
                    Dim dtSheet As DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                    sheetName = dtSheet.Rows(0)("TABLE_NAME").ToString()
                    conn.Close()
                End Using

                ' Read data
                Dim da As New OleDbDataAdapter("SELECT * FROM [" & sheetName & "] where ItemCode <> ''", connStr)
                Dim ds As New DataSet()
                da.Fill(ds)

                If ds.Tables.Count > 0 Then
                    dgDetail.DataSource = ds.Tables(0)

                    ' Apply Grid View styling and set editable columns
                    SetupDataGridViewEditableColumns()
                Else
                    MessageBox.Show("No data found.")
                End If

            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub SetupDataGridViewEditableColumns()
        With dgDetail
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            .ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            .ColumnHeadersDefaultCellStyle.BackColor = Color.Navy
            .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            .EnableHeadersVisualStyles = False

            .DefaultCellStyle.BackColor = Color.White
            .AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue

            .DefaultCellStyle.Font = New Font("Segoe UI", 9)
            .DefaultCellStyle.SelectionBackColor = Color.DodgerBlue
            .DefaultCellStyle.SelectionForeColor = Color.White

            .RowHeadersVisible = False
            .AllowUserToAddRows = False

            ' Set ReadOnly to False to allow editing
            .ReadOnly = False

            ' Make all columns read-only first
            For Each column As DataGridViewColumn In .Columns
                column.ReadOnly = True
            Next

            ' Make specific columns editable - adjust column names/indices based on your actual data
            SetEditableColumn("ItemCode", 1)     ' ItemCode column
            SetEditableColumn("AccountNumber", 4) ' AccountNumber column
            SetEditableColumn("Quantity", 5)      ' Quantity column

            ' Optional: Set better header names if needed
            If .Columns.Contains("ItemCode") Then
                .Columns("ItemCode").HeaderText = "Item Code (Editable)"
            End If
            If .Columns.Contains("AccountNumber") Then
                .Columns("AccountNumber").HeaderText = "Account Code (Editable)"
            End If
            If .Columns.Contains("Quantity") Then
                .Columns("Quantity").HeaderText = "Quantity (Editable)"
            End If
        End With
    End Sub

    Private Sub SetEditableColumn(columnName As String, columnIndex As Integer)
        If dgDetail.Columns.Contains(columnName) Then
            dgDetail.Columns(columnName).ReadOnly = False
        ElseIf dgDetail.Columns.Count > columnIndex Then
            dgDetail.Columns(columnIndex).ReadOnly = False
        End If
    End Sub


    Private Sub btn_export_template_excel_Click(sender As Object, e As EventArgs) Handles btn_export_template_excel.Click
        Try
            ' Create a new SaveFileDialog to let the user choose where to save the Excel file
            Dim sfd As New SaveFileDialog
            sfd.Filter = "Excel Files (*.xlsx)|*.xlsx"
            sfd.FileName = "Internal_Usage_Template.xlsx"
            sfd.Title = "Save Excel Template"

            If sfd.ShowDialog = DialogResult.OK Then
                ' Use ClosedXML to create an Excel file
                Using wb As New ClosedXML.Excel.XLWorkbook()
                    Dim ws As ClosedXML.Excel.IXLWorksheet = wb.Worksheets.Add("Invoice Template")

                    ' Define the headers based on the provided structure
                    Dim headers As String() = {"NO", "ItemCode", "ItemName", "Location", "AccountNumber", "Quantity", "Comments"}

                    ' Write headers to the first row
                    For colIndex As Integer = 0 To headers.Length - 1
                        ws.Cell(1, colIndex + 1).Value = headers(colIndex)
                    Next

                    ' Optional: Format the header row
                    Dim headerRange = ws.Range(1, 1, 1, headers.Length)
                    headerRange.Style.Font.Bold = True
                    headerRange.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.LightBlue
                    headerRange.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center

                    ' Auto-adjust column widths based on content
                    ws.Columns().AdjustToContents()

                    ' Save the workbook to the selected file path
                    wb.SaveAs(sfd.FileName)

                    MessageBox.Show("Excel template exported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Using
            End If
        Catch ex As Exception
            MessageBox.Show("Error exporting template: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class