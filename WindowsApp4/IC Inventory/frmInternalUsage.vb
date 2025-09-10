Imports System.Reflection.Emit
Imports Mid_SMR.Read
Imports System.Data.OleDb
Imports ACCPAC.Advantage
Imports Mid_SMR.SQLGetData
Imports AccpacCOMAPI
Public Class Internal_Usage

    Public Shared SageSess As New ACCPAC.Advantage.Session
    Public Property Username = frmLogin.TextBox1.Text
    Public Property Password = frmLogin.TextBox2.Text

    Private Sub btnPostSage_Click(sender As Object, e As EventArgs) Handles btnPostSage.Click

        If dgDetail.RowCount > 0 Then

            On Error GoTo ACCPACErrorHandler


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
                .FieldByName("TRANSDATE").SetValue(dt1.Value.ToString("dd-MM-yyyy"), True)
                .FieldByName("HDRDESC").SetValue(TextBox3.Text.Trim(), False)
                .FieldByName("REFERENCE").SetValue(TextBox2.Text.Trim(), False)
                .FieldByName("EMPLOYEENO").SetValue(TextBox4.Text.Trim(), False)
            End With
            With ICICE1detail1.Fields
                Dim i As Integer
                For i = 0 To dgDetail.Rows.Count - 1

                    ICICE1detail1.RecordCreate(0)
                    .FieldByName("ITEMNO").SetValue(dgDetail.Rows(i).Cells(1).Value.ToString, False)
                    .FieldByName("LOCATION").SetValue(dgDetail.Rows(i).Cells(3).Value.ToString, False)                  ' Location
                    .FieldByName("GLACCT").SetValue(dgDetail.Rows(i).Cells(4).Value.ToString, False)                    ' Location
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

ACCPACErrorHandler:
            Dim lCount As Long
            Dim lIndex As Long
            If SageSess.Errors Is Nothing Then
                ' MessageBox.Show(Err.Description, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
                '  MessageBox.Show("Please check Username and Password!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                lCount = SageSess.Errors.Count
                If lCount = 0 Then
                    'do nothing
                Else
                    MessageBox.Show(Err.Description, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If

        Else
            MsgBox("No transactions found.", MsgBoxStyle.Information, "Alert")
        End If

    End Sub

    Private Sub Internal_Usage_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For f As Integer = 0 To dgDetail.Rows.Count - 1
            Dim num As Integer = Val(dgDetail.Rows(f).Cells(11).Value)
            If num <= 0 Then
                dgDetail.Rows(f).DefaultCellStyle.BackColor = Color.Red
            End If
        Next

        dt1.Format = DateTimePickerFormat.Custom
        dt1.CustomFormat = "yyyy-MM-dd"
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        Me.Close()
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
                Else
                    MessageBox.Show("No data found.")
                End If

            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End If

    End Sub

End Class