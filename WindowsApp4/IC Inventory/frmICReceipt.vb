Imports System.Data.OleDb
Imports ACCPAC.Advantage
Imports Mid_SMR.Read

Public Class frmICReceipt
    Public Shared SageSess As New ACCPAC.Advantage.Session
    Public Property Username = frmLogin.TextBox1.Text
    Public Property Password = frmLogin.TextBox2.Text
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim SageSess As Session = New Session()

        SageSess.Init("", "XY", "XY1000", "72A")
        SageSess.Open(Username, Password, ReadConnectionNotedPage.DBName_St, Today, 0)

        Dim myDBLink As DBLink = SageSess.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)

        Dim ICREE1header As View = myDBLink.OpenView("IC0590")
        Dim ICREE1detail1 As View = myDBLink.OpenView("IC0580")
        Dim ICREE1detail2 As View = myDBLink.OpenView("IC0595")
        Dim ICREE1detail3 As View = myDBLink.OpenView("IC0585")
        Dim viewArray() As View = {ICREE1detail1, ICREE1detail2}
        ICREE1header.Compose(viewArray)
        viewArray = New View() {ICREE1header, Nothing, Nothing, Nothing, Nothing, Nothing, ICREE1detail3}
        ICREE1detail1.Compose(viewArray)
        viewArray = New View() {ICREE1header}
        ICREE1detail2.Compose(viewArray)
        viewArray = New View() {ICREE1detail1}
        ICREE1detail3.Compose(viewArray)
        ICREE1header.Cancel()
        ICREE1header.Init()
        Dim fields As ViewFields = ICREE1header.Fields
        fields.FieldByName("RECPDATE").SetValue(dt1.Value.ToString("dd-MM-yyyy"), False)
        fields.FieldByName("REFERENCE").SetValue(Me.txt_Ref.Text, False)
        fields.FieldByName("RECPDESC").SetValue(Me.Txt_Desc.Text, False)
        fields = Nothing
        Dim i As Integer = 0
        Dim count As Integer = Me.dgDetail.Rows.Count - 1
        i = 0
        Do
            ICREE1detail1.RecordClear()
            ICREE1detail1.RecordCreate(DirectCast(CLng(0), ViewRecordCreate))
            Dim viewField As ViewFields = ICREE1detail1.Fields
            viewField.FieldByName("ITEMNO").SetValue(Me.dgDetail.Rows(i).Cells(1).Value.ToString().Trim(), True)
            viewField.FieldByName("PROCESSCMD").SetValue(1, False)
            ICREE1detail1.Process()
            viewField.FieldByName("LOCATION").SetValue(Me.dgDetail.Rows(i).Cells(3).Value.ToString().Trim(), True)
            viewField.FieldByName("RECPQTY").SetValue(Conversion.Val(Me.dgDetail.Rows(i).Cells(4).Value.ToString()), False)
            viewField.FieldByName("UNITCOST").SetValue(Conversion.Val(Me.dgDetail.Rows(i).Cells(5).Value.ToString()), False)
            ICREE1detail1.Insert()
            viewField = Nothing
            i = i + 1
        Loop While i <= count
        ICREE1header.Fields.FieldByName("STATUS").SetValue(1, False)
        ICREE1header.Insert()
        SageSess = Nothing
        SageSess = New Session()
        MessageBox.Show("Saving Complete !")

        dgDetail.DataSource = Nothing
        dgDetail.Columns.Clear()
        TextBox1.Text = ""
        txt_Ref.Text = ""
        Txt_Desc.Text = ""


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

    Private Sub B_close_Click(sender As Object, e As EventArgs) Handles B_close.Click
        Me.Close()
    End Sub

    Private Sub frmICReceipt_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dt1.Format = DateTimePickerFormat.Custom
        dt1.CustomFormat = "yyyy-MM-dd"
    End Sub
End Class