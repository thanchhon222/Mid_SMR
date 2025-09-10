Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports ACCPAC.Advantage
Imports Mid_SMR.SQLGetData
Imports AccpacCOMAPI
Public Class frmRevenueImport
    Public Shared SageSess As New Session
    Dim con As OleDbConnection
    Dim ds As DataSet
    Dim da As OleDbDataAdapter
    Dim TMPOE As New DataTable
    Dim temp As Boolean
    Dim ExcelFile As Object
    Dim Errors As Object
    Private _array As Object
    Private _array1 As Object
    Dim objConn As Object

    Dim BaustellenpreisdateiDataGridView As Object
    Dim dtExportItems As Object

    Dim Tbl As DataTable = Nothing
    Dim _Tbl1 As DataTable = Nothing
    Public Property Username As String
    Public Property Password As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Button2.Enabled = False
            DataGridView1.Visible = False
            DataGridView1.DataSource = ""
            Dim openFileDialog As OpenFileDialog = New OpenFileDialog
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx"
            openFileDialog.ShowDialog()

            '   openFileDialog.ShowDialog()
            Dim trandate As Date = DateTimePicker1.Value.AddDays(0)
            Dim formattedDate As String = trandate.ToString("yyyyMMdd")
            Dim FileName As String = "IDB_GL_" & formattedDate & ".xlsx"

            Dim CurrentDay As Date = DateTimePicker1.Value.AddDays(0)
            Dim fmCurrentDay As String = trandate.ToString("yyyyMMdd")
            Dim sheet As String = "IDB_GL_" & fmCurrentDay & ""

            TextBox1.Text = openFileDialog.FileName


            If Not String.IsNullOrEmpty(TextBox1.Text) Then
                con = New OleDbConnection(("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" _
                                + (TextBox1.Text + ";Extended Properties=Excel 12.0;")))
                con.Open()
                Dim dt As DataTable = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
                con.Close()
                Dim i As Integer = 0
                Do While (i < dt.Rows.Count)
                    Dim sheetName As String = dt.Rows(i)("TABLE_NAME").ToString
                    sheetName = sheetName.Substring(0, (sheetName.Length))
                    i = (i + 1)
                Loop

            End If

            da = New OleDbDataAdapter("SELECT '1' As [Status],
                                       FORMAT(DateSerial(Mid(BatchDate, 7, 4), Mid(BatchDate, 4, 2), Mid(BatchDate, 1, 2)), 'yyyy-mm-dd') As BatchDate,
                                       BatchDesc,EntryDesc,
                                       FORMAT(DateSerial(Mid(TranDate, 7, 4), Mid(TranDate, 4, 2), Mid(TranDate, 1, 2)), 'yyyy-mm-dd')  As TranDate,
                                       LineRefer,
                                       LineDescription,Comments,AccountCode,AccountName,[TypeD(D/C)] As [TypeDC], REPLACE(REPLACE(REPLACE([Total Amount(D/C)], '(', '-'), ')', ''), ',', '') As TotalAmount
                                       From [" & sheet & "$] where [BatchDate] is not null", con)
            ds = New DataSet
            da.Fill(ds)
            ''  DataGridView1.DataSource = ds.Tables(0)
            con.Close()

            TextBox5.Text = SQL_GetField_SC("select [SheetName] from [z_tbl_revenue] where [SheetName] = '" & sheet & "' group by [SheetName]")

            If TextBox5.Text = sheet Then
                MessageBox.Show("Can't Load agian. This data already exist in database temporary.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""
                DataGridView1.Columns.Clear()
            Else

                For Each i As DataRow In ds.Tables(0).Rows

                    ' For Each row As DataGridViewRow In DataGridView1.Rows
                    Dim sqlComm As New SqlCommand()
                    sqlComm.Connection = SQLCNN
                    sqlComm.CommandText = "[sp_Insert_Revenue]"
                    sqlComm.CommandType = CommandType.StoredProcedure
                    sqlComm.Parameters.AddWithValue("@SheetName", sheet)
                    sqlComm.Parameters.AddWithValue("@Status", (Convert.ToString(i("Status")).Trim))
                    sqlComm.Parameters.AddWithValue("@BatchDate", (Convert.ToString(i("BatchDate")).Trim))
                    sqlComm.Parameters.AddWithValue("@BatchDesc", (Convert.ToString(i("BatchDesc")).Trim))
                    sqlComm.Parameters.AddWithValue("@EntryDesc", (Convert.ToString(i("EntryDesc")).Trim))
                    sqlComm.Parameters.AddWithValue("@TranDate", (Convert.ToString(i("TranDate")).Trim))
                    sqlComm.Parameters.AddWithValue("@LineRefer", (Convert.ToString(i("LineRefer")).Trim))
                    sqlComm.Parameters.AddWithValue("@LineDescription", (Convert.ToString(i("LineDescription")).Trim))
                    sqlComm.Parameters.AddWithValue("@Comments", (Convert.ToString(i("Comments")).Trim))
                    sqlComm.Parameters.AddWithValue("@AccountCode", (Convert.ToString(i("AccountCode")).Trim))
                    sqlComm.Parameters.AddWithValue("@AccountName", (Convert.ToString(i("AccountName")).Trim))
                    sqlComm.Parameters.AddWithValue("@TypeDC", (Convert.ToString(i("TypeDC")).Trim))
                    sqlComm.Parameters.AddWithValue("@TotalAmount", (Convert.ToString(i("TotalAmount")).Trim))

                    If SQLCNN.State <> ConnectionState.Closed Then SQLCNN.Close()
                    SQLCNN.ConnectionString = Read_CNN_St()
                    SQLCNN.Open()
                    sqlComm.ExecuteNonQuery()
                    '   Next

                Next
                MessageBox.Show("Uploading data completed")
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Me.Cursor = Cursors.WaitCursor

        Try
            Dim SageSess As New AccpacSession
            SageSess.Init("", "XY", "XY1000", "71A")
            SageSess.Open(Username, Password, ReadConnectionNotedPage.DBName_St, Today, 0, "")
            ' Dim mDBLinkCmpRW As ACCPAC.Advantage.DBLink
            ' mDBLinkCmpRW = SageSess.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)

            Dim mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkCmpRW = SageSess.OpenDBLink(tagDBLinkTypeEnum.DBLINK_COMPANY, tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)

            Dim GLBATCH1batch As AccpacView = Nothing
            mDBLinkCmpRW.OpenView("GL0008", GLBATCH1batch)
            Dim GLBATCH1header As AccpacView = Nothing
            mDBLinkCmpRW.OpenView("GL0006", GLBATCH1header)
            Dim GLBATCH1detail1 As AccpacView = Nothing
            mDBLinkCmpRW.OpenView("GL0010", GLBATCH1detail1)
            Dim GLBATCH1detail2 As AccpacView = Nothing
            mDBLinkCmpRW.OpenView("GL0402", GLBATCH1detail2)
            Dim GLPOST2 As AccpacView = Nothing
            mDBLinkCmpRW.OpenView("GL0030", GLPOST2)
            Dim GLBATCH3batch As AccpacView = Nothing
            mDBLinkCmpRW.OpenView("GL0008", GLBATCH3batch)
            Dim GLBATCH3header As AccpacView = Nothing
            mDBLinkCmpRW.OpenView("GL0006", GLBATCH3header)
            Dim GLBATCH3detail1 As AccpacView = Nothing
            mDBLinkCmpRW.OpenView("GL0010", GLBATCH3detail1)
            Dim GLBATCH3detail2 As AccpacView = Nothing
            mDBLinkCmpRW.OpenView("GL0402", GLBATCH3detail2)
            Dim GLPOST4 As AccpacView = Nothing
            mDBLinkCmpRW.OpenView("GL0030", GLPOST4)

            GLBATCH1batch.Compose(New AccpacView() {GLBATCH1header})
            GLBATCH1header.Compose(New AccpacView() {GLBATCH1batch, GLBATCH1detail1})
            GLBATCH1detail1.Compose(New AccpacView() {GLBATCH1header, GLBATCH1detail2})
            GLBATCH1detail2.Compose(New AccpacView() {GLBATCH1detail1})

            GLBATCH3batch.Compose(New AccpacView() {GLBATCH3header})
            GLBATCH3header.Compose(New AccpacView() {GLBATCH3batch, GLBATCH3detail1})
            GLBATCH3detail1.Compose(New AccpacView() {GLBATCH3header, GLBATCH3detail2})
            GLBATCH3detail2.Compose(New AccpacView() {GLBATCH3detail1})

            Dim batDate As Date = DateTimePicker1.Value.AddDays(0)
            Dim formatbatDate As String = batDate.ToString("dd MMM-yy")

            GLBATCH1batch.RecordCreate(1)

            GLBATCH1batch.Fields.FieldByName("PROCESSCMD").Value = ("1")        ' Lock Batch Switch
            GLBATCH1batch.Process()
            GLBATCH1header.Fields.FieldByName("BTCHENTRY").Value = ("")         ' Entry Number
            ' GLBATCH1header.Fetch(0)
            GLBATCH1header.Fields.FieldByName("BTCHENTRY").Value = ("00000")                    ' Entry Number
            GLBATCH1header.RecordCreate(2)
            GLBATCH1batch.Fields.FieldByName("BTCHDESC").Value = ("Revenue import from IDB " & formatbatDate & "")   ' Description
            GLBATCH1batch.Update()

            Dim trandate As Date = DateTimePicker1.Value.AddDays(0)
            Dim formattedDate As String = trandate.ToString("yyyy-MM-dd")
            Dim months As String = trandate.ToString("MM")

            Dim batchSize As Integer = 8000
            _Tbl1 = SQL_GetTable_St("select * from z_tbl_revenue where TranDate = '" & formattedDate & "'")
            Dim totalRecords As Integer = _Tbl1.Rows.Count
            If _Tbl1 IsNot Nothing Then

                For Each i As DataRow In _Tbl1.Rows

                    GLBATCH1header.Fields.FieldByName("DOCDATE").Value = ((Convert.ToString(i("TranDate")).Trim))      ' Document Date
                    GLBATCH1header.Fields.FieldByName("FSCSPERD").Value = (months)
                    GLBATCH1header.Fields.FieldByName("SRCETYPE").Value = ("JE")                       ' Source Type
                    GLBATCH1header.Fields.FieldByName("DATEENTRY").Value = ((Convert.ToString(i("TranDate")).Trim))     ' Posting Date

                    GLBATCH1header.Fields.FieldByName("JRNLDESC").Value = ("Revenue On " & (Convert.ToString(i("TranDate")).Trim))     ' Description
                    GLBATCH1header.Fields.FieldByName("BTCHENTRY").Value = ("00000")                    ' Entry Number
                    GLBATCH1detail1.RecordCreate(0)

                    GLBATCH1detail1.Fields.FieldByName("ACCTID").Value = ((Convert.ToString(i("AccountCode")).Trim))                    ' Account Number
                    GLBATCH1detail1.Fields.FieldByName("SCURNAMT").Value = ((Convert.ToString(i("TotalAmount")).Trim))                    ' Source Currency Amount
                    GLBATCH1detail1.Fields.FieldByName("TRANSDESC").Value = ((Convert.ToString(i("LineDescription")).Trim))                      ' Description
                    GLBATCH1detail1.Fields.FieldByName("TRANSREF").Value = ((Convert.ToString(i("LineRefer")).Trim))                    ' Referrence
                    GLBATCH1detail1.Fields.FieldByName("COMMENT").Value = ((Convert.ToString(i("Comments")).Trim))                    ' Referrence

                    GLBATCH1detail1.Insert()

                Next

                GLBATCH1header.Insert()
                GLBATCH1header.RecordCreate(2)

                MessageBox.Show("Data Loading to Sage Completed.")
                DataGridView1.Columns.Clear()
                DataGridView1.DataSource = Nothing
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox5.Text = ""
                '  SageSess.Close()
            Else
                MessageBox.Show("No data loading.")
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub frmRevenueImport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button2.Enabled = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            DataGridView1.Visible = True
            DataGridView1.Columns.Clear()
            Dim trandate As Date = DateTimePicker1.Value.AddDays(0)
            Dim formattedDate As String = trandate.ToString("yyyy-MM-dd")

            Using connection As New SqlConnection(V_CNNstring_Central)
                Dim selectCommand As String = "SELECT * FROM [z_tbl_revenue] where [Status] = '1' and Trandate = '" & formattedDate & "'"
                Dim dataAdapter As New SqlDataAdapter(selectCommand, connection)
                Dim dataSet As New DataSet()
                dataAdapter.Fill(dataSet, "[z_tbl_revenue]")

                ' Bind the DataGridView to the dataset
                DataGridView1.DataSource = dataSet.Tables("[z_tbl_revenue]")
            End Using
            ' Sum Total COH
            Dim Dr As Decimal = 0
            Dim Cr As Decimal = 0
            Dim Va As Decimal = 0
            For i As Integer = 0 To DataGridView1.Rows.Count() - 1 Step +1
                If DataGridView1.Rows(i).Cells(12).Value > 0 Then
                    Dr = Dr + DataGridView1.Rows(i).Cells(12).Value
                ElseIf DataGridView1.Rows(i).Cells(12).Value < 0 Then
                    Cr = Cr + DataGridView1.Rows(i).Cells(12).Value
                Else
                End If
            Next
            TextBox2.Text = Format(Val(Dr.ToString()), v_NoFormart)
            TextBox3.Text = Format(Val(-Cr.ToString()), v_NoFormart)
            TextBox4.Text = Val(Dr.ToString()) + Val(Cr.ToString())

            Dim RowCount As String = ""
            RowCount = DataGridView1.Rows.Count

            If RowCount = 0 Then
                Button2.Enabled = False
            Else
                Button2.Enabled = True
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class