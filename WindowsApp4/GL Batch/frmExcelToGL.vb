Imports System.Text
Imports ACCPAC.Advantage
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports AccpacCOMAPI
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ListView
Public Class frmExcelToGL
    Dim con As OleDbConnection
    Dim ds As DataSet
    Dim da As OleDbDataAdapter
    Dim TMPOE As New DataTable
    Dim temp As Boolean
    Dim ExcelFile As Object
    Dim Errors As Object
    Private _array As Object
    Private _array1 As Object
    Private FILENAME As String
    Public Shared SageSess As New ACCPAC.Advantage.Session
    Public Property Username = frmLogin.TextBox1.Text
    Public Property Password = frmLogin.TextBox2.Text
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim op As New OpenFileDialog
        Try
            op.Filter = ("MS Excel (*.xls)|*.xlsx")

            If op.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                TxtBrowse.Text = op.FileName

                Dim connStr As String = ""
                Dim filePath As String = TxtBrowse.Text.Trim()

                ' Choose correct provider based on extension
                If filePath.EndsWith(".xls") Then
                    connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filePath & ";Extended Properties='Excel 8.0;HDR=YES;'"
                ElseIf filePath.EndsWith(".xlsx") Then
                    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES;'"
                Else
                    MessageBox.Show("Unsupported file type.")
                    Exit Sub
                End If

                'Get sheet name
                Dim sheetName As String = ""
                Using conn As New OleDbConnection(connStr)
                    conn.Open()
                    Dim dtSheet As DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                    sheetName = dtSheet.Rows(0)("TABLE_NAME").ToString()
                    conn.Close()
                End Using

                da = New OleDbDataAdapter("SELECT * FROM [" & sheetName & "] where D_Account <> 0 ", connStr)

                ds = New DataSet
                da.Fill(ds)
                dgDetail.DataSource = ds.Tables(0)

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

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub BtnImport_Click(sender As Object, e As EventArgs) Handles BtnImport.Click

        If String.IsNullOrWhiteSpace(TxtDescription.Text) Then
            MessageBox.Show("Please input Batch Description", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            ' Initialize Sage session
            SageSess.Init("", "XY", "XY1000", "72A")
            SageSess.Open(Username, Password, ReadConnectionNotedPage.DBName_St, Today, 0)

            Using mDBLinkCmpRW As ACCPAC.Advantage.DBLink = SageSess.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)
                ' Open views for batch processing
                Dim GLBATCH1batch As View = mDBLinkCmpRW.OpenView("GL0008")
                Dim GLBATCH1header As View = mDBLinkCmpRW.OpenView("GL0006")
                Dim GLBATCH1detail1 As View = mDBLinkCmpRW.OpenView("GL0010")
                Dim GLBATCH1detail2 As View = mDBLinkCmpRW.OpenView("GL0402")

                ' Compose views
                GLBATCH1batch.Compose(New View() {GLBATCH1header})
                GLBATCH1header.Compose(New View() {GLBATCH1batch, GLBATCH1detail1})
                GLBATCH1detail1.Compose(New View() {GLBATCH1header, GLBATCH1detail2})
                GLBATCH1detail2.Compose(New View() {GLBATCH1detail1})

                ' Initialize batch
                GLBATCH1batch.RecordCreate(1)
                GLBATCH1batch.Fields.FieldByName("BTCHDESC").SetValue(TxtDescription.Text.Trim(), False)
                GLBATCH1batch.Fields.FieldByName("PROCESSCMD").SetValue("1", False) ' Lock Batch
                GLBATCH1batch.Update()

                ' Group data by H_Date and H_JournalDes
                Dim OEHEAD = ds.Tables(0).Select().CopyToDataTable()
                Dim groupedData = OEHEAD.AsEnumerable() _
                .GroupBy(Function(row) New With {
                    Key .Date = row.Field(Of Date)("H_Date"),
                    Key .JournalDes = row.Field(Of String)("H_JournalDes").Trim()
                })

                For Each group In groupedData
                    Dim docDate As Date = group.Key.Date
                    Dim journalDesc As String = group.Key.JournalDes

                    ' Create new header for this group
                    GLBATCH1header.RecordCreate(2)
                    GLBATCH1header.Fields.FieldByName("BTCHENTRY").SetValue("00000", False)
                    GLBATCH1header.Fields.FieldByName("DOCDATE").SetValue(docDate, False)
                    GLBATCH1header.Fields.FieldByName("SRCETYPE").SetValue("JE", False)
                    GLBATCH1header.Fields.FieldByName("DATEENTRY").SetValue(docDate, False)
                    GLBATCH1header.Fields.FieldByName("JRNLDESC").SetValue(journalDesc, False)

                    Dim docDateTime As DateTime = Convert.ToDateTime(docDate)
                    GLBATCH1header.Fields.FieldByName("FSCSPERD").SetValue(docDateTime.ToString("MM"), False)
                    GLBATCH1header.Fields.FieldByName("FSCSYR").SetValue(docDateTime.ToString("yyyy"), False)

                    ' Validate balance
                    Dim totalDebit As Decimal = group.Sum(Function(d) If(IsDBNull(d("D_Debit")), 0, Convert.ToDecimal(d("D_Debit"))))
                    Dim totalCredit As Decimal = group.Sum(Function(d) If(IsDBNull(d("D_Credit")), 0, Convert.ToDecimal(d("D_Credit"))))
                    If totalDebit <> totalCredit Then
                        Throw New Exception($"Journal entry for {journalDesc} on {docDate} is not balanced. Debit: {totalDebit}, Credit: {totalCredit}")
                    End If

                    ' Process details for this header
                    For Each d In group
                        GLBATCH1detail1.RecordCreate(0)
                        GLBATCH1detail1.Fields.FieldByName("TRANSREF").SetValue(Convert.ToString(d("D_Ref")).Trim(), False)
                        GLBATCH1detail1.Fields.FieldByName("TRANSDESC").SetValue(Convert.ToString(d("D_Desc")).Trim(), False)
                        GLBATCH1detail1.Fields.FieldByName("ACCTID").SetValue(Convert.ToString(d("D_Account")).Trim(), False)

                        Dim debit As Decimal = If(IsDBNull(d("D_Debit")), 0, Convert.ToDecimal(d("D_Debit")))
                        Dim credit As Decimal = If(IsDBNull(d("D_Credit")), 0, Convert.ToDecimal(d("D_Credit")))
                        GLBATCH1detail1.Fields.FieldByName("SCURNAMT").SetValue(debit - credit, False)

                        GLBATCH1detail1.Insert()
                    Next

                    ' Insert header after all details are added
                    GLBATCH1header.Insert()
                Next

                ' Finalize batch
                GLBATCH1batch.Process()
            End Using



            ' Handle success
            If SageSess.Errors Is Nothing OrElse SageSess.Errors.Count = 0 Then
                MessageBox.Show("Your Import is Completed!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                dgDetail.DataSource = Nothing
                TxtBrowse.Text = ""
                TxtDescription.Text = ""
            Else
                For lIndex As Integer = 0 To SageSess.Errors.Count - 1
                    MessageBox.Show(SageSess.Errors.Item(lIndex).Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Next
                SageSess.Errors.Clear()
            End If

        Catch ex As Exception
            MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If SageSess IsNot Nothing Then

            End If
        End Try

    End Sub

End Class