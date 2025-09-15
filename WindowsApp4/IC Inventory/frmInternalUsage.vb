Imports System.Data.OleDb
Imports System.Text
Imports ACCPAC.Advantage
Imports ClosedXML.Excel
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
                    .FieldByName("ITEMNO").SetValue(dgDetail.Rows(i).Cells(1).Value.ToString, False)
                    .FieldByName("LOCATION").SetValue(dgDetail.Rows(i).Cells(3).Value.ToString, False)
                    .FieldByName("GLACCT").SetValue(dgDetail.Rows(i).Cells(4).Value.ToString, False)
                    .FieldByName("QUANTITY").SetValue(dgDetail.Rows(i).Cells(5).Value.ToString, False)
                    .FieldByName("COMMENTS").SetValue(dgDetail.Rows(i).Cells(6).Value.ToString, False)

                    ICICE1detail1.Process()
                    ICICE1detail1.Insert()
                Next
            End With
            ICICE1header.Fields.FieldByName("STATUS").SetValue("1", False)
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
        Dim missingQuantities As New List(Of String)
        Dim isValid As Boolean = True

        ' Clear previous highlighting and error messages
        For Each row As DataGridViewRow In dgDetail.Rows
            If Not row.IsNewRow Then
                row.DefaultCellStyle.BackColor = Color.White
                row.Cells(1).Style.BackColor = Color.White ' ItemCode
                row.Cells(4).Style.BackColor = Color.White ' AccountNumber
                row.Cells(5).Style.BackColor = Color.White ' Quantity
                If dgDetail.Columns.Contains("Error") Then
                    row.Cells("Error").Value = ""
                End If
            End If
        Next

        Try
            For i As Integer = 0 To dgDetail.Rows.Count - 1
                Dim row As DataGridViewRow = dgDetail.Rows(i)
                If row.IsNewRow Then Continue For

                Dim rowNumber As Integer = i + 1
                Dim hasError As Boolean = False
                Dim errorMessage As New StringBuilder()

                Dim itemCode As String = If(row.Cells(1).Value IsNot Nothing, row.Cells(1).Value.ToString().Trim(), "")
                If String.IsNullOrEmpty(itemCode) Then
                    errorMessage.AppendLine("Missing ItemCode")
                    row.Cells(1).Style.BackColor = Color.LightPink
                    hasError = True
                    isValid = False
                    missingItems.Add($"Row {rowNumber}: ItemCode is missing")
                Else
                    Dim sageItemCode As String = itemCode.Replace("-", "")
                    Dim itemCheckDt As DataTable = SQL_GetTable_St("SELECT ITEMNO FROM ICITEM WHERE ITEMNO = '" & sageItemCode.Replace("'", "''") & "'")
                    If itemCheckDt.Rows.Count = 0 Then
                        errorMessage.AppendLine($"ItemCode '{itemCode}' not found in Sage")
                        row.Cells(1).Style.BackColor = Color.LightPink
                        hasError = True
                        isValid = False
                        missingItems.Add($"Row {rowNumber}: ItemCode '{itemCode}' not found")
                    End If
                End If

                Dim accountNumber As String = If(row.Cells(4).Value IsNot Nothing, row.Cells(4).Value.ToString().Trim(), "")
                If String.IsNullOrEmpty(accountNumber) Then
                    errorMessage.AppendLine("Missing AccountCode")
                    row.Cells(4).Style.BackColor = Color.LightPink
                    hasError = True
                    isValid = False
                    missingAccounts.Add($"Row {rowNumber}: AccountCode is missing")
                Else
                    Dim sageAccountNumber As String = accountNumber.Replace("-", "")
                    Dim glCheckDt As DataTable = SQL_GetTable_St("SELECT ACCTID FROM GLAMF WHERE ACCTID = '" & sageAccountNumber.Replace("'", "''") & "'")
                    If glCheckDt.Rows.Count = 0 Then
                        errorMessage.AppendLine($"AccountCode '{accountNumber}' not found in Sage")
                        row.Cells(4).Style.BackColor = Color.LightPink
                        hasError = True
                        isValid = False
                        missingAccounts.Add($"Row {rowNumber}: AccountCode '{accountNumber}' not found")
                    End If
                End If

                Dim quantity As String = If(row.Cells(5).Value IsNot Nothing, row.Cells(5).Value.ToString().Trim(), "")
                If String.IsNullOrEmpty(quantity) Then
                    errorMessage.AppendLine("Missing Quantity")
                    row.Cells(5).Style.BackColor = Color.LightPink
                    hasError = True
                    isValid = False
                    missingQuantities.Add($"Row {rowNumber}: Quantity is missing")
                Else
                    Dim qtyValue As Decimal
                    If Not Decimal.TryParse(quantity, qtyValue) OrElse qtyValue <= 0 Then
                        errorMessage.AppendLine("Invalid Quantity (must be > 0)")
                        row.Cells(5).Style.BackColor = Color.LightPink
                        hasError = True
                        isValid = False
                        missingQuantities.Add($"Row {rowNumber}: Invalid Quantity '{quantity}'")
                    End If
                End If

                If hasError AndAlso dgDetail.Columns.Contains("Error") Then
                    row.Cells("Error").Value = errorMessage.ToString().Trim()
                End If
            Next

            If missingItems.Count > 0 OrElse missingAccounts.Count > 0 OrElse missingQuantities.Count > 0 Then
                Dim errorMessage As New StringBuilder()
                errorMessage.AppendLine("Validation failed! Please correct the following issues:")

                If missingItems.Count > 0 Then
                    errorMessage.AppendLine()
                    errorMessage.AppendLine("Missing/Invalid Item Codes:")
                    For Each item In missingItems
                        errorMessage.AppendLine($"- {item}")
                    Next
                End If

                If missingAccounts.Count > 0 Then
                    errorMessage.AppendLine()
                    errorMessage.AppendLine("Missing/Invalid Account Codes:")
                    For Each account In missingAccounts
                        errorMessage.AppendLine($"- {account}")
                    Next
                End If

                If missingQuantities.Count > 0 Then
                    errorMessage.AppendLine()
                    errorMessage.AppendLine("Missing/Invalid Quantities:")
                    For Each qty In missingQuantities
                        errorMessage.AppendLine($"- {qty}")
                    Next
                End If

                errorMessage.AppendLine()
                errorMessage.AppendLine("Note: Cells with errors are highlighted in pink.")
                errorMessage.AppendLine("Sage stores codes without hyphens. Your data contains hyphens that will be automatically removed during posting.")

                MessageBox.Show(errorMessage.ToString(), "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

                For Each row As DataGridViewRow In dgDetail.Rows
                    If Not row.IsNewRow Then
                        For Each cell As DataGridViewCell In row.Cells
                            If cell.Style.BackColor = Color.LightPink Then
                                dgDetail.CurrentCell = cell
                                dgDetail.BeginEdit(True)
                                Exit For
                            End If
                        Next
                        If dgDetail.IsCurrentCellInEditMode Then Exit For
                    End If
                Next
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
            For Each row As DataGridViewRow In dgDetail.Rows
                If Not row.IsNewRow Then
                    If row.DefaultCellStyle.BackColor = Color.Orange Then
                        row.DefaultCellStyle.BackColor = Color.White
                    End If
                    If row.Cells(5).Style.BackColor = Color.LightCoral Then
                        row.Cells(5).Style.BackColor = Color.White
                    End If
                End If
            Next

            For i As Integer = 0 To dgDetail.Rows.Count - 1
                Dim row As DataGridViewRow = dgDetail.Rows(i)
                If row.IsNewRow Then Continue For

                Dim rowNumber As Integer = i + 1
                Dim itemCode As String = If(row.Cells(1).Value IsNot Nothing, row.Cells(1).Value.ToString().Trim(), "")
                Dim location As String = If(row.Cells(3).Value IsNot Nothing, row.Cells(3).Value.ToString().Trim(), "")
                Dim quantity As Decimal = 0
                Decimal.TryParse(If(row.Cells(5).Value IsNot Nothing, row.Cells(5).Value.ToString().Trim(), "0"), quantity)

                If Not String.IsNullOrEmpty(itemCode) AndAlso Not String.IsNullOrEmpty(location) AndAlso quantity > 0 Then
                    Dim sageItemCode As String = itemCode.Replace("-", "")
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
                            row.Cells(5).Style.BackColor = Color.LightCoral
                            If dgDetail.Columns.Contains("Error") Then
                                Dim currentError As String = If(row.Cells("Error").Value IsNot Nothing, row.Cells("Error").Value.ToString(), "")
                                If Not String.IsNullOrEmpty(currentError) Then
                                    currentError += Environment.NewLine
                                End If
                                row.Cells("Error").Value = currentError + $"Insufficient stock in Quantity cell: Available {qtyOnHand}, Required {quantity}"
                            End If
                        End If
                    Else
                        outOfStockItems.Add($"Row {rowNumber}: {itemCode} not found at location {location}")
                        isValid = False
                        row.Cells(5).Style.BackColor = Color.LightCoral
                        If dgDetail.Columns.Contains("Error") Then
                            Dim currentError As String = If(row.Cells("Error").Value IsNot Nothing, row.Cells("Error").Value.ToString(), "")
                            If Not String.IsNullOrEmpty(currentError) Then
                                currentError += Environment.NewLine
                            End If
                            row.Cells("Error").Value = currentError + "Item not found at specified location in Quantity cell"
                        End If
                    End If
                End If
            Next

            If outOfStockItems.Count > 0 Then
                Dim errorMessage As New StringBuilder()
                errorMessage.AppendLine("Insufficient stock! The following items cannot be posted:")
                For Each item In outOfStockItems
                    errorMessage.AppendLine($"- {item}")
                Next
                errorMessage.AppendLine()
                errorMessage.AppendLine("Cells with stock issues are highlighted in coral.")
                errorMessage.AppendLine("Please adjust the Quantity cell (highlighted) before posting.")

                MessageBox.Show(errorMessage.ToString(), "Stock Check Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                For Each row As DataGridViewRow In dgDetail.Rows
                    If Not row.IsNewRow Then
                        If row.Cells(5).Style.BackColor = Color.LightCoral Then
                            dgDetail.CurrentCell = row.Cells(5)
                            dgDetail.BeginEdit(True)
                            Exit For
                        End If
                    End If
                Next
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

    Private Sub B_browse_Click(sender As Object, e As EventArgs) Handles B_browse.Click
        OpenFileDialog1.Filter = "Excel Files|*.xls;*.xlsx"

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName

            Try
                Dim connStr As String = ""
                Dim filePath As String = TextBox1.Text.Trim()

                If filePath.EndsWith(".xls") Then
                    connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filePath & ";Extended Properties='Excel 8.0;HDR=YES;'"
                ElseIf filePath.EndsWith(".xlsx") Then
                    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES;'"
                Else
                    MessageBox.Show("Unsupported file type. Please select an Excel file (.xls or .xlsx).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If

                Dim sheetName As String = ""
                Using conn As New OleDbConnection(connStr)
                    conn.Open()
                    Dim dtSheet As DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                    sheetName = dtSheet.Rows(0)("TABLE_NAME").ToString()
                    conn.Close()
                End Using

                Dim da As New OleDbDataAdapter("SELECT * FROM [" & sheetName & "] WHERE ItemCode <> ''", connStr)
                Dim ds As New DataSet()
                da.Fill(ds)

                If ds.Tables.Count > 0 Then
                    dgDetail.DataSource = Nothing
                    dgDetail.Columns.Clear()
                    dgDetail.DataSource = ds.Tables(0)
                    SetupDataGridViewEditableColumns()
                    If Not dgDetail.Columns.Contains("Error") Then
                        Dim errorColumn As New DataGridViewTextBoxColumn()
                        errorColumn.Name = "Error"
                        errorColumn.HeaderText = "Error Messages"
                        errorColumn.ReadOnly = True
                        errorColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                        dgDetail.Columns.Add(errorColumn)
                    End If
                    For i As Integer = 0 To dgDetail.Rows.Count - 1
                        If Not dgDetail.Rows(i).IsNewRow Then
                            ValidateRow(i)
                        End If
                    Next
                    MessageBox.Show("Data loaded successfully from " & filePath, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("No data found in the Excel file.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If

            Catch ex As Exception
                MessageBox.Show("Error loading data: " & ex.Message & Environment.NewLine & "Ensure the file is not open in another program and try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub dgDetail_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgDetail.CellValidating
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

    Private Sub dgDetail_CellBeginEdit(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgDetail.CellBeginEdit
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim column As DataGridViewColumn = dgDetail.Columns(e.ColumnIndex)
            If Not column.ReadOnly Then
                dgDetail.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.LightYellow
            Else
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub dgDetail_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgDetail.CellEndEdit
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            dgDetail.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.White
            ValidateRow(e.RowIndex)
        End If
    End Sub

    Private Sub ValidateRow(rowIndex As Integer)
        If rowIndex < 0 OrElse rowIndex >= dgDetail.Rows.Count OrElse dgDetail.Rows(rowIndex).IsNewRow Then
            Return
        End If

        Dim row As DataGridViewRow = dgDetail.Rows(rowIndex)
        Dim errorMessage As New StringBuilder()
        Dim hasError As Boolean = False

        If dgDetail.Columns.Contains("Error") Then
            row.Cells("Error").Value = ""
        End If
        row.Cells(1).Style.BackColor = Color.White
        row.Cells(4).Style.BackColor = Color.White
        row.Cells(5).Style.BackColor = Color.White

        Dim itemCode As String = If(row.Cells(1).Value IsNot Nothing, row.Cells(1).Value.ToString().Trim(), "")
        If String.IsNullOrEmpty(itemCode) Then
            errorMessage.AppendLine("Missing ItemCode")
            row.Cells(1).Style.BackColor = Color.LightPink
            hasError = True
        End If

        Dim accountNumber As String = If(row.Cells(4).Value IsNot Nothing, row.Cells(4).Value.ToString().Trim(), "")
        If String.IsNullOrEmpty(accountNumber) Then
            errorMessage.AppendLine("Missing AccountCode")
            row.Cells(4).Style.BackColor = Color.LightPink
            hasError = True
        End If

        Dim quantity As String = If(row.Cells(5).Value IsNot Nothing, row.Cells(5).Value.ToString().Trim(), "")
        If String.IsNullOrEmpty(quantity) Then
            errorMessage.AppendLine("Missing Quantity")
            row.Cells(5).Style.BackColor = Color.LightPink
            hasError = True
        Else
            Dim qtyValue As Decimal
            If Not Decimal.TryParse(quantity, qtyValue) OrElse qtyValue <= 0 Then
                errorMessage.AppendLine("Invalid Quantity (must be > 0)")
                row.Cells(5).Style.BackColor = Color.LightPink
                hasError = True
            End If
        End If

        If hasError AndAlso dgDetail.Columns.Contains("Error") Then
            row.Cells("Error").Value = errorMessage.ToString().Trim()
        End If
    End Sub

    Private Sub dgDetail_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgDetail.CellDoubleClick
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim column As DataGridViewColumn = dgDetail.Columns(e.ColumnIndex)
            If Not column.ReadOnly Then
                dgDetail.CurrentCell = dgDetail.Rows(e.RowIndex).Cells(e.ColumnIndex)
                dgDetail.BeginEdit(True)
            End If
        End If
    End Sub

    Private Sub dgDetail_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgDetail.CellClick
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim column As DataGridViewColumn = dgDetail.Columns(e.ColumnIndex)
            If Not column.ReadOnly AndAlso (dgDetail.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.LightPink OrElse
               dgDetail.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = Color.LightCoral) Then
                dgDetail.CurrentCell = dgDetail.Rows(e.RowIndex).Cells(e.ColumnIndex)
                dgDetail.BeginEdit(True)
            End If
        End If
    End Sub

    Private Sub dgDetail_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgDetail.CellMouseEnter
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim cell As DataGridViewCell = dgDetail.Rows(e.RowIndex).Cells(e.ColumnIndex)
            If cell.Style.BackColor = Color.LightPink OrElse cell.Style.BackColor = Color.LightCoral Then
                Dim errorText As String = If(dgDetail.Rows(e.RowIndex).Cells("Error").Value IsNot Nothing, dgDetail.Rows(e.RowIndex).Cells("Error").Value.ToString(), "")
                cell.ToolTipText = errorText
            Else
                cell.ToolTipText = ""
            End If
        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
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
            .ReadOnly = False

            For Each column As DataGridViewColumn In .Columns
                column.ReadOnly = True
            Next

            SetEditableColumn("ItemCode", 1)
            SetEditableColumn("AccountNumber", 4)
            SetEditableColumn("Quantity", 5)

            If .Columns.Contains("ItemCode") Then
                .Columns("ItemCode").HeaderText = "Item Code (Editable)"
            End If
            If .Columns.Contains("AccountNumber") Then
                .Columns("AccountNumber").HeaderText = "Account Code (Editable)"
            End If
            If .Columns.Contains("Quantity") Then
                .Columns("Quantity").HeaderText = "Quantity (Editable)"
            End If

            If Not .Columns.Contains("Error") Then
                Dim errorColumn As New DataGridViewTextBoxColumn()
                errorColumn.Name = "Error"
                errorColumn.HeaderText = "Error Messages"
                errorColumn.ReadOnly = True
                errorColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Columns.Add(errorColumn)
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
            Dim sfd As New SaveFileDialog
            sfd.Filter = "Excel Files (*.xlsx)|*.xlsx"
            sfd.FileName = "Internal_Usage_Template.xlsx"
            sfd.Title = "Save Excel Template"

            If sfd.ShowDialog = DialogResult.OK Then
                Using wb As New XLWorkbook()
                    Dim ws As IXLWorksheet = wb.Worksheets.Add("Invoice Template")
                    Dim headers As String() = {"NO", "ItemCode", "ItemName", "Location", "AccountNumber", "Quantity", "Comments"}
                    For colIndex As Integer = 0 To headers.Length - 1
                        ws.Cell(1, colIndex + 1).Value = headers(colIndex)
                    Next
                    Dim headerRange = ws.Range(1, 1, 1, headers.Length)
                    headerRange.Style.Font.Bold = True
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center
                    ws.Columns().AdjustToContents()
                    wb.SaveAs(sfd.FileName)
                    MessageBox.Show("Excel template exported successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Using
            End If
        Catch ex As Exception
            MessageBox.Show("Error exporting template: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class