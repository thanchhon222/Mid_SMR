Imports System.IO
Imports System.Data
Imports ACCPAC.Advantage
Imports System.Text
Imports System.Data.SqlClient
Imports Microsoft.Graph.Drives.Item.Items.Item.Workbook.Functions
Imports Mid_SMR.SQLGetData

Public Class frm_InventoryUpload
    Dim TBL As DataTable = Nothing
    Public Shared SageSess As New ACCPAC.Advantage.Session
    Dim Errors As Object
    Public Property Username = frmLogin.TextBox1.Text
    Public Property Password = frmLogin.TextBox2.Text
    Private Sub btn_GetData_Click(sender As Object, e As EventArgs) Handles btn_GetData.Click
        If txt_Location.Text = "" Then
            MessageBox.Show("No data to locatuion – Select locaiton first.",
                V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        OpenFileDialog1.Filter = "CSV Files|*.csv|Excel Files|*.xls;*.xlsx"

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            TextBox1.Text = OpenFileDialog1.FileName
            Try
                Dim filePath As String = TextBox1.Text.Trim()

                ' Check if it's a CSV file
                If filePath.EndsWith(".csv") Then
                    ' Read CSV file directly
                    Dim dt As New DataTable()

                    ' Read all lines from the CSV file
                    Dim lines As String() = File.ReadAllLines(filePath)

                    ' Check if there are any lines
                    If lines.Length = 0 Then
                        MessageBox.Show("The CSV file is empty.")
                        Exit Sub
                    End If

                    ' Get headers (first line)
                    Dim headers As String() = lines(0).Split(","c)
                    For Each header As String In headers
                        dt.Columns.Add(header.Trim())
                    Next

                    ' Add data rows (skip first line if it's headers)
                    For i As Integer = 1 To lines.Length - 1
                        Dim fields As String() = lines(i).Split(","c)
                        ' Ensure we have the right number of fields
                        If fields.Length = headers.Length Then
                            dt.Rows.Add(fields)
                        End If
                    Next

                    ' Filter out empty ItemCode rows
                    Dim dv As DataView = dt.DefaultView
                    dv.RowFilter = "ItemCode <> ''"

                    ' Bind to DataGridView
                    dgDetail.DataSource = dv.ToTable()

                ElseIf filePath.EndsWith(".xlsx") OrElse filePath.EndsWith(".xls") Then
                    ' Original Excel reading code (for backward compatibility)
                    Dim connStr As String = If(filePath.EndsWith(".xlsx"),
                        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES;'",
                        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filePath & ";Extended Properties='Excel 8.0;HDR=YES;'")

                    ' Get sheet name
                    Dim sheetName As String = ""
                    Using conn As New OleDb.OleDbConnection(connStr)
                        conn.Open()
                        Dim dtSheet As DataTable = conn.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, Nothing)
                        sheetName = dtSheet.Rows(0)("TABLE_NAME").ToString()
                        conn.Close()
                    End Using

                    ' Read data
                    Dim da As New OleDb.OleDbDataAdapter("SELECT * FROM [" & sheetName & "] where ItemCode <> '' And Location = ' " & txt_Location.Text & " '", connStr)
                    Dim ds As New DataSet()
                    da.Fill(ds)

                    If ds.Tables.Count > 0 Then
                        dgDetail.DataSource = ds.Tables(0)
                    Else
                        MessageBox.Show("No data found.")
                        Exit Sub
                    End If
                Else
                    MessageBox.Show("Unsupported file type.")
                    Exit Sub
                End If

                ' Apply Grid View styling
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
                    .ReadOnly = True
                End With

            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub btn_SavetoSage_Click(sender As Object, e As EventArgs) Handles btn_SavetoSage.Click

        If dgDetail.Rows.Count <= 0 Then
            MessageBox.Show("No data to export – Get Data first.",
                V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        ' 1️⃣ Adjust SQL
        Const SqlUpdate As String = "UPDATE ICWKUH SET QTYCOUNTED = @ColA, COSTVAR = @ColB, QTYVAR = @ColC WHERE ITEMNO = @Pk "
        Using conn As New SqlConnection(V_CNNstringStock)
            conn.Open()
            Using tran As SqlTransaction = conn.BeginTransaction()
                Try
                    Using cmd As New SqlCommand(SqlUpdate, conn, tran)
                        ' 🛠 Declare both parameters here
                        cmd.Parameters.Add("@ColA", SqlDbType.Decimal)
                        cmd.Parameters.Add("@ColB", SqlDbType.Decimal)
                        cmd.Parameters.Add("@ColC", SqlDbType.Decimal)
                        cmd.Parameters.Add("@Pk", SqlDbType.NVarChar, 50)

                        For Each gridRow As DataGridViewRow In dgDetail.Rows
                            If gridRow.IsNewRow Then Continue For   ' skip new blank row

                            ' 🧾 Get values from the grid
                            Dim pkValue As String = CStr(gridRow.Cells("ItemCode").Value).Trim()
                            Dim colAValue As Decimal = CDec(gridRow.Cells("ScannedCount").Value)
                            Dim colACost As Decimal = CDec(gridRow.Cells("COSTVAR").Value)
                            Dim colQtyCost As Decimal = CDec(gridRow.Cells("QtyVar").Value)

                            ' 🗝 Assign parameter values
                            cmd.Parameters("@Pk").Value = pkValue
                            cmd.Parameters("@ColA").Value = colAValue
                            cmd.Parameters("@ColB").Value = colACost
                            cmd.Parameters("@ColC").Value = colQtyCost

                            ' 🔄 Execute update
                            cmd.ExecuteNonQuery()
                        Next
                    End Using


                    tran.Commit()

                    MessageBox.Show("Save completed for header.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Save failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        End Using


        ' 1️⃣ Adjust SQL
        Const SqlUpdate_detail As String = "UPDATE ICWKUD SET QTYCOUNTED = @ColA WHERE ITEMNO = @Pk "
        Using conn_detail As New SqlConnection(V_CNNstringStock)
            conn_detail.Open()
            Using tran_detail As SqlTransaction = conn_detail.BeginTransaction()
                Try
                    Using cmd As New SqlCommand(SqlUpdate_detail, conn_detail, tran_detail)
                        ' 🛠 Declare both parameters here
                        cmd.Parameters.Add("@ColA", SqlDbType.Decimal)
                        cmd.Parameters.Add("@Pk", SqlDbType.NVarChar, 50)

                        For Each gridRow As DataGridViewRow In dgDetail.Rows
                            If gridRow.IsNewRow Then Continue For   ' skip new blank row

                            ' 🧾 Get values from the grid
                            Dim pkValue As String = CStr(gridRow.Cells("ItemCode").Value).Trim()
                            Dim colAValue As Decimal = CDec(gridRow.Cells("ScannedCount").Value)

                            ' 🗝 Assign parameter values
                            cmd.Parameters("@Pk").Value = pkValue
                            cmd.Parameters("@ColA").Value = colAValue

                            ' 🔄 Execute update
                            cmd.ExecuteNonQuery()
                        Next
                    End Using
                    tran_detail.Commit()

                    MessageBox.Show("Save completed for detail.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    dgDetail.Columns.Clear()
                    txt_Location.Text = ""
                Catch ex As Exception
                    MessageBox.Show("Save failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        End Using

    End Sub

    Private Sub btn_Find_Click(sender As Object, e As EventArgs) Handles btn_Find.Click
        Try
            F_Barcode = ""
            Dim F As New frm_Find_Location
            F.StartPosition = FormStartPosition.CenterScreen
            F.ShowDialog()
            If F_Barcode <> "" Then
                txt_Location.Text = F_Barcode
            End If
            txt_Location.Focus()
        Catch ex As Exception
            WriteError(ex.Message)
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class