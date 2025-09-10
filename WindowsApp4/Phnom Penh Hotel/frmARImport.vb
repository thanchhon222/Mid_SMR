Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports ACCPAC.Advantage
Imports ACCPAC.Advantage.Session
Imports Mid_SMR.SQLGetData

Public Class frmARImport
    Public Shared SageSess As New Session
    Dim con As OleDbConnection
    Dim ds As DataSet
    Dim da As OleDbDataAdapter
    Dim temp As Boolean
    Dim ExcelFile As Object
    Dim Errors As Object
    Private _array As Object
    Private _array1 As Object
    Dim objConn As Object

    Dim consql As New SqlConnection(Read_CNN_St())
    Dim cmd As New SqlCommand
    Private dttts As DataTable = New DataTable()
    Dim dssql As DataSet

    Dim BaustellenpreisdateiDataGridView As Object
    Dim dtExportItems As Object
    Public Property Username As String
    Public Property Password As String
    Private Sub frmARImport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton1.Checked = True
        RadioButton2.Checked = False
        Button1.Enabled = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            DataGridView1.DataSource = ""
            Dim openFileDialog As OpenFileDialog = New OpenFileDialog
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx"
            openFileDialog.ShowDialog()
            TextBox1.Text = openFileDialog.FileName


            If Not String.IsNullOrEmpty(openFileDialog.FileName) Then
                con = New OleDbConnection(("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" _
                                + (openFileDialog.FileName + ";Extended Properties=Excel 12.0;")))
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

            da = New OleDbDataAdapter("SELECT * From [revenue$] where [Customer] is not null", con)
            ds = New DataSet
            da.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
            'Add decimal on column
            DataGridView1.Columns(9).DefaultCellStyle.Format = "#.00"
            con.Close()

            ' Sum Total COH
            Dim Dr As Decimal = 0
            Dim Cr As Double = 0
            Dim Va As Decimal = 0
            For i As Integer = 0 To DataGridView1.Rows.Count() - 1 Step +1
                If DataGridView1.Rows(i).Cells(9).Value > 0 Then
                    Cr = Cr + DataGridView1.Rows(i).Cells(9).Value
                ElseIf DataGridView1.Rows(i).Cells(9).Value < 0 Then
                    Cr = Cr + DataGridView1.Rows(i).Cells(9).Value
                Else
                End If
            Next
            TextBox3.Text = Format(Val(Cr.ToString()), v_NoFormart)

            MessageBox.Show("Completed")

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        DataGridView1.Columns.Clear()
        DataGridView1.DataSource = ""


        Dim com As String = "SELECT * FROM _ARREVENUE where DocDate = '" & DateTimePicker1.Text.Trim & "'"
        Dim Adpt As New SqlDataAdapter(com, consql)
        Adpt.Fill(dttts)
        DataGridView1.DataSource = dttts

        ' Sum Total COH
        Dim Dr As Decimal = 0
        Dim Cr As Double = 0
        Dim Va As Decimal = 0
        For i As Integer = 0 To DataGridView1.Rows.Count() - 1 Step +1
            If DataGridView1.Rows(i).Cells(9).Value > 0 Then
                Cr = Cr + DataGridView1.Rows(i).Cells(9).Value
            ElseIf DataGridView1.Rows(i).Cells(9).Value < 0 Then
                Cr = Cr + DataGridView1.Rows(i).Cells(9).Value
            Else
            End If
        Next
        TextBox3.Text = Format(Val(Cr.ToString()), v_NoFormart)

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            SageSess.Init("", "XY", "XY1000", "71A")
            SageSess.Open(UCase(Username.Trim()), (Password.Trim()), (ReadConnectionNotedPage.DBName_St.Trim()), Today, 0)
            'SageSess.Open("IMPORT", "SCA2021", "REPD21", Today, 0)
            Dim mDBLinkCmpRW As ACCPAC.Advantage.DBLink
            mDBLinkCmpRW = SageSess.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)
            '***** Start***
            ' Dim mDBLinkCmpRW As ACCPAC.Advantage.DBLink
            mDBLinkCmpRW = SageSess.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)

            'Auto AR Invoice Batch

            Dim ARINVCPOST1 As View = mDBLinkCmpRW.OpenView("AR0048")
            Dim ARPAYMPOST2 As View = mDBLinkCmpRW.OpenView("AR0049")
            Dim ARRFPOST3 As View = mDBLinkCmpRW.OpenView("AR0150")

            ' Dim ICICE1header As View = mDBLinkCmpRW.OpenView("IC0288")

            Dim ARINVOICE1batch As View = mDBLinkCmpRW.OpenView("AR0031")
            Dim ARINVOICE1header As View = mDBLinkCmpRW.OpenView("AR0032")
            Dim ARINVOICE1detail1 As View = mDBLinkCmpRW.OpenView("AR0033")
            Dim ARINVOICE1detail2 As View = mDBLinkCmpRW.OpenView("AR0034")
            Dim ARINVOICE1detail3 As View = mDBLinkCmpRW.OpenView("AR0402")
            Dim ARINVOICE1detail4 As View = mDBLinkCmpRW.OpenView("AR0401")
            Dim ARINVOICE3header As View = mDBLinkCmpRW.OpenView("AR0032")

            Dim ARINVCPOST2 As View = mDBLinkCmpRW.OpenView("AR0048")
            Dim ARINVOICE3batch As View = mDBLinkCmpRW.OpenView("AR0031")
            ' Dim ARINVOICE3header As View = mDBLinkCmpRW.OpenView("AR0032")
            Dim ARINVOICE3detail1 As View = mDBLinkCmpRW.OpenView("AR0033")
            Dim ARINVOICE3detail2 As View = mDBLinkCmpRW.OpenView("AR0034")
            Dim ARINVOICE3detail3 As View = mDBLinkCmpRW.OpenView("AR0402")
            Dim ARINVOICE3detail4 As View = mDBLinkCmpRW.OpenView("AR0401")

            ARINVOICE3batch.Compose(New View() {ARINVOICE3header})
            ARINVOICE3header.Compose(New View() {ARINVOICE3batch, ARINVOICE3detail1, ARINVOICE3detail2,
                                 ARINVOICE3detail3, Nothing})
            ARINVOICE3detail1.Compose(New View() {ARINVOICE3header, ARINVOICE3batch, ARINVOICE3detail4})
            ARINVOICE3detail2.Compose(New View() {ARINVOICE3header})
            ARINVOICE3detail3.Compose(New View() {ARINVOICE3header})
            ARINVOICE3detail4.Compose(New View() {ARINVOICE3detail1})

            ARINVOICE1batch.Compose(New View() {ARINVOICE1header})
            ARINVOICE1header.Compose(New View() {ARINVOICE1batch, ARINVOICE1detail1, ARINVOICE1detail2, ARINVOICE1detail3, Nothing})
            ARINVOICE1detail1.Compose(New View() {ARINVOICE1header, ARINVOICE1batch, ARINVOICE1detail4})
            ARINVOICE1detail2.Compose(New View() {ARINVOICE1header})
            ARINVOICE1detail3.Compose(New View() {ARINVOICE1header})
            ARINVOICE1detail4.Compose(New View() {ARINVOICE1detail1})

            If RadioButton2.Checked = True Then

                ARINVOICE1batch.RecordCreate(1)

                ARINVOICE1batch.Fields.FieldByName("PROCESSCMD").SetValue("1", False)      ' Process Command
                ARINVOICE1batch.Process()
                ARINVOICE1batch.Read(0)
                ARINVOICE1header.RecordCreate(2)
                ARINVOICE1detail1.Cancel()

                ARINVOICE1batch.Fields.FieldByName("BTCHDESC").SetValue("Imported from Excel " & DateTime.Now.ToString("dd/MMM/yy HH:mm:ss"), False)   ' Description

                ARINVOICE1batch.Update()
                ARINVOICE1batch.Fields.FieldByName("DATEBTCH").SetValue(Today, False)   ' Batch Date
                ARINVOICE1batch.Update()
                ARINVOICE1batch.Read(0)
                ARINVOICE1header.RecordCreate(2)
                ARINVOICE1detail1.Cancel()

                Dim ARHEAD = ds.Tables(0).[Select].CopyToDataTable()
                Dim ARDETAILS =
                  New DataView(ARHEAD).ToTable(True,
                                                Convert.ToString("Customer").Trim,
                                                Convert.ToString("Document").Trim,
                                                Convert.ToString("DocDate").Trim,
                                                Convert.ToString("EntryDescription").Trim,
                                                Convert.ToString("TYPE").Trim)

                For Each h As DataRow In ARDETAILS.Rows

                    ARINVOICE1header.Fields.FieldByName("IDCUST").SetValue(Convert.ToString(h("Customer")).Trim, False)   ' Customer Number
                    ARINVOICE1header.Fields.FieldByName("IDINVC").SetValue(Convert.ToString(h("Document")).Trim, False)              ' Document Number
                    ARINVOICE1header.Fields.FieldByName("DATEINVC").SetValue(Convert.ToString(h("DocDate")).Trim, False)   ' Document Date
                    ARINVOICE1header.Fields.FieldByName("INVCDESC").SetValue(Convert.ToString(h("EntryDescription")).Trim, False)   ' Invoice Description
                    ARINVOICE1header.Fields.FieldByName("INVCTYPE").SetValue("2", False)     ' Invoice Type
                    ARINVOICE1header.Fields.FieldByName("TEXTTRX").SetValue(Convert.ToString(h("TYPE")).Trim, False) 'Document Type
                    ARINVOICE1header.Fields.FieldByName("TAXSTTS1").SetValue("2", False)                   ' Tax Class 1

                    For Each d In ARHEAD.[Select](
                    String.Format("Customer='{0}' AND Document='{1}'AND DocDate='{2}'AND EntryDescription = '{3}'AND TYPE = '{4}'", Convert.ToString(h("Customer")).Trim, Convert.ToString(h("Document")).Trim, Convert.ToString(h("DocDate")).Trim, Convert.ToString(h("EntryDescription")).Trim, Convert.ToString(h("TYPE")).Trim))

                        ARINVOICE1detail1.RecordClear()
                        ARINVOICE1detail1.RecordCreate(0)
                        ARINVOICE1detail1.Fields.FieldByName("TEXTDESC").SetValue(Convert.ToString(d("DetailDescription")).Trim, False)                     ' Description
                        ARINVOICE1detail1.Fields.FieldByName("IDACCTREV").SetValue(Convert.ToString(d("GLAccount")).Trim, False)                ' Revenue Account
                        ARINVOICE1detail1.Insert()
                        ARINVOICE1detail1.Read(0)
                        ARINVOICE1detail1.Fields.FieldByName("AMTEXTN").SetValue(Convert.ToString(d("GLAmount")).Trim, False)              ' Extended Amount w/ TIP
                        ARINVOICE1detail1.Update()
                        ARINVOICE1detail1.Read(0)

                    Next

                    ARINVOICE1header.Insert()
                    ARINVOICE1detail1.Read(0)
                    ARINVOICE1detail1.Read(0)
                    ARINVOICE1batch.Read(0)
                    ARINVOICE1header.RecordCreate(2)
                    ARINVOICE1detail1.Cancel()

                Next

                MessageBox.Show("Imported Completed", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)

                TextBox1.Text = ""
                DataGridView1.DataSource = ""

            End If

            If RadioButton1.Checked = True Then

                ARINVOICE1batch.RecordCreate(1)

                ARINVOICE1batch.Fields.FieldByName("PROCESSCMD").SetValue("1", False)      ' Process Command
                ARINVOICE1batch.Process()
                ARINVOICE1batch.Read(0)
                ARINVOICE1header.RecordCreate(2)
                ARINVOICE1detail1.Cancel()

                ARINVOICE1batch.Fields.FieldByName("BTCHDESC").SetValue("Imported from SQL " & DateTime.Now.ToString("dd/MMM/yy HH:mm:ss"), False)   ' Description

                ARINVOICE1batch.Update()
                ARINVOICE1batch.Fields.FieldByName("DATEBTCH").SetValue(Today, False)   ' Batch Date
                ARINVOICE1batch.Update()
                ARINVOICE1batch.Read(0)
                ARINVOICE1header.RecordCreate(2)
                ARINVOICE1detail1.Cancel()


                Dim ARHEAD1 = dttts.[Select]("TYPE='1'").CopyToDataTable()

                Dim ARDETAILS1 =
                  New DataView(ARHEAD1).ToTable(True,
                                                Convert.ToString("Customer").Trim,
                                                Convert.ToString("Document").Trim,
                                                Convert.ToString("DocDate").Trim,
                                                Convert.ToString("EntryDescription").Trim,
                                                Convert.ToString("TYPE").Trim)

                For Each h As DataRow In ARDETAILS1.Rows

                    ARINVOICE1header.Fields.FieldByName("IDCUST").SetValue(Convert.ToString(h("Customer")).Trim, False)   ' Customer Number
                    ARINVOICE1header.Fields.FieldByName("IDINVC").SetValue(Convert.ToString(h("Document")).Trim, False)              ' Document Number
                    ARINVOICE1header.Fields.FieldByName("DATEINVC").SetValue(Convert.ToString(h("DocDate")).Trim, False)   ' Document Date
                    ARINVOICE1header.Fields.FieldByName("INVCDESC").SetValue(Convert.ToString(h("EntryDescription")).Trim, False)   ' Invoice Description
                    ARINVOICE1header.Fields.FieldByName("INVCTYPE").SetValue("2", False)     ' Invoice Type
                    ARINVOICE1header.Fields.FieldByName("TEXTTRX").SetValue(Convert.ToString(h("TYPE")).Trim, False) 'Document Type
                    ARINVOICE1header.Fields.FieldByName("TAXSTTS1").SetValue("2", False)                   ' Tax Class 1

                    For Each d In ARHEAD1.[Select](
                        String.Format("Customer='{0}' AND Document='{1}'AND DocDate='{2}'AND EntryDescription = '{3}'AND TYPE = '{4}'", Convert.ToString(h("Customer")).Trim, Convert.ToString(h("Document")).Trim, Convert.ToString(h("DocDate")).Trim, Convert.ToString(h("EntryDescription")).Trim, Convert.ToString(h("TYPE")).Trim))

                        ARINVOICE1detail1.RecordClear()
                        ARINVOICE1detail1.RecordCreate(0)
                        ARINVOICE1detail1.Fields.FieldByName("TEXTDESC").SetValue(Convert.ToString(d("DetailDescription")).Trim, False)                     ' Description
                        ARINVOICE1detail1.Fields.FieldByName("IDACCTREV").SetValue(Convert.ToString(d("GLAccount")).Trim, False)                ' Revenue Account
                        ARINVOICE1detail1.Insert()
                        ARINVOICE1detail1.Read(0)
                        ARINVOICE1detail1.Fields.FieldByName("AMTEXTN").SetValue(Convert.ToString(d("GLAmount")).Trim, False)              ' Extended Amount w/ TIP
                        ARINVOICE1detail1.Update()
                        ARINVOICE1detail1.Read(0)

                    Next

                    ARINVOICE1header.Insert()
                    ARINVOICE1detail1.Read(0)
                    ARINVOICE1detail1.Read(0)
                    ARINVOICE1batch.Read(0)
                    ARINVOICE1header.RecordCreate(2)
                    ARINVOICE1detail1.Cancel()

                Next

                MessageBox.Show("Imported Completed", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)

                TextBox1.Text = ""
                DataGridView1.DataSource = ""

            End If

        Catch ex As Exception
            For i As Integer = 0 To SageSess.Errors.Count - 1
                MessageBox.Show("Sage Error Message: " & SageSess.Errors(i).Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Next
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            Button4.Enabled = True
        Else
            Button4.Enabled = False
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            Button1.Enabled = True
        Else
            Button1.Enabled = False
        End If
    End Sub
End Class