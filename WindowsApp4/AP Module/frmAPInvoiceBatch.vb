Imports System.Text
Imports ACCPAC.Advantage
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ListView
Public Class AP_Invoice_Batch
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
    Private dt As DataTable = New DataTable()
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

                da = New OleDbDataAdapter("SELECT * FROM [" & sheetName & "] where Invoice <> '' ", connStr)
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

        On Error GoTo ACCPACErrorHandlerPOST

        SageSess.Init("", "XY", "XY1000", "72A")
        SageSess.Open(Username, Password, ReadConnectionNotedPage.DBName_St, Today, 0)

        Dim mDBLinkCmpRW As ACCPAC.Advantage.DBLink
        mDBLinkCmpRW = SageSess.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)


        Dim APINVOICE1batch As View = mDBLinkCmpRW.OpenView("AP0020")
        Dim APINVOICE1header As View = mDBLinkCmpRW.OpenView("AP0021")
        Dim APINVOICE1detail1 As View = mDBLinkCmpRW.OpenView("AP0022")
        Dim APINVOICE1detail2 As View = mDBLinkCmpRW.OpenView("AP0023")
        Dim APINVOICE1detail3 As View = mDBLinkCmpRW.OpenView("AP0402")
        Dim APINVOICE1detail4 As View = mDBLinkCmpRW.OpenView("AP0401")
        Dim APINVCPOST2 As View = mDBLinkCmpRW.OpenView("AP0039")


        APINVOICE1batch.Compose(New View() {APINVOICE1header})
        APINVOICE1header.Compose(New View() {APINVOICE1batch, APINVOICE1detail1, APINVOICE1detail2, APINVOICE1detail3})
        APINVOICE1detail1.Compose(New View() {APINVOICE1header, APINVOICE1batch, APINVOICE1detail4})
        APINVOICE1detail2.Compose(New View() {APINVOICE1header})
        APINVOICE1detail3.Compose(New View() {APINVOICE1header})
        APINVOICE1detail4.Compose(New View() {APINVOICE1detail1})


        Dim ARHEAD = ds.Tables(0).[Select].CopyToDataTable()
        Dim ARDETAILS =
                New DataView(ARHEAD).ToTable(True,
                                                         Convert.ToString("DocDate").Trim,
                                                         Convert.ToString("Invoice").Trim,
                                                         Convert.ToString("DocAmt").Trim,
                                                         Convert.ToString("VendorCode").Trim,
                                                         Convert.ToString("PONumber").Trim,
                                                         Convert.ToString("DocType").Trim)


        APINVOICE1batch.RecordCreate(1)
        APINVOICE1batch.Fields.FieldByName("PROCESSCMD").SetValue("1", False) ' Process Command Code
        APINVOICE1batch.Process()
        APINVOICE1batch.Read(0)
        APINVOICE1header.RecordCreate(2)
        'APINVOICE1batch.Fields.FieldByName("BTCHDESC").SetValue(DateTimePicker1.Value.ToString("yyyyMMdd"), False)   ' Description
        APINVOICE1batch.Fields.FieldByName("BTCHDESC").SetValue(TxtDescription.Text.ToString, False)   ' Description
        APINVOICE1batch.Update()
        'APINVOICE1batch.Fields.FieldByName("DATEBTCH").SetValue(DateTimePicker1.Value.ToString, False)  ' Batch Date

        ' Import Header Transaction 
        For Each h As DataRow In ARDETAILS.Rows

            APINVOICE1batch.Update()
            APINVOICE1batch.Read(0)
            APINVOICE1header.RecordCreate(2)
            APINVOICE1detail1.Cancel()
            APINVOICE1header.Fields.FieldByName("IDVEND").SetValue(Convert.ToString(h("VendorCode")).Trim, False)                  ' Vendor Number
            APINVOICE1header.Fields.FieldByName("PROCESSCMD").SetValue("7", False)     ' Process Command Code
            APINVOICE1header.Process()
            APINVOICE1header.Fields.FieldByName("PROCESSCMD").SetValue("7", False)     ' Process Command Code
            APINVOICE1header.Fields.FieldByName("DATEINVC").SetValue(Convert.ToString(h("DocDate")).Trim, False)  ' Document Date
            APINVOICE1header.Process()
            APINVOICE1header.Fields.FieldByName("TEXTTRX").SetValue(Convert.ToString(h("DocType")).Trim, False)                    ' Document Type
            APINVOICE1header.Fields.FieldByName("PROCESSCMD").SetValue("7", False)     ' Process Command Code
            APINVOICE1header.Fields.FieldByName("INVCDESC").SetValue(TxtDescription.Text.ToString, False)   ' Invoice Description
            APINVOICE1header.Fields.FieldByName("TAXCLASS1").SetValue("2", False)                     ' Tax Class 1

            APINVOICE1header.Process()
            APINVOICE1header.Fields.FieldByName("IDINVC").SetValue(Convert.ToString(h("Invoice")).Trim, False)                ' Document Number
            APINVOICE1header.Fields.FieldByName("AMTGROSTOT").SetValue(Convert.ToString(h("DocAmt")).Trim, False)                 ' Document Total Including Tax
            APINVOICE1header.Fields.FieldByName("PONBR").SetValue(Convert.ToString(h("PONumber")).Trim, False)     ' PO Number
            APINVOICE1detail1.RecordClear()
            APINVOICE1detail1.RecordCreate(0)

            ' Import Detail Transaction 
            For Each d In ARHEAD.[Select](
                           String.Format("DocDate='{0}' AND Invoice='{1}'AND DocAmt='{2}'AND VendorCode='{3}'AND PONumber='{4}'AND DocType='{5}'", Convert.ToString(h("DocDate")).Trim, Convert.ToString(h("Invoice")).Trim, Convert.ToString(h("DocAmt")).Trim, Convert.ToString(h("VendorCode")).Trim, Convert.ToString(h("PONumber")).Trim, Convert.ToString(h("DocType")).Trim))

                APINVOICE1detail1.Fields.FieldByName("IDGLACCT").SetValue(Convert.ToString(d("GLAccount")).Trim, False)                ' G/L Account
                APINVOICE1detail1.Fields.FieldByName("AMTDIST").SetValue(Convert.ToString(d("Amount")).Trim, False)                  ' Distributed Amount

                APINVOICE1detail1.Insert()
                APINVOICE1detail1.Fields.FieldByName("CNTLINE").SetValue("-1", False)      ' Line Number
                APINVOICE1detail1.Read(0)
                APINVOICE1detail1.RecordCreate(0)
                APINVOICE1detail1.Process()
                'APINVOICE1header.Fields.FieldByName("ORDRNBR").SetValue(DataGridView1.Rows(i).Cells(10).Value.ToString, False)   ' Order Number
            Next

            APINVOICE1header.Insert()
            APINVOICE1header.Fields.FieldByName("PROCESSCMD").SetValue("7", False)     ' Process Command Code
            APINVOICE1header.Process()
            APINVOICE1batch.Read(0)
            APINVOICE1header.RecordCreate(2)
            APINVOICE1detail1.Cancel()
        Next
        MessageBox.Show("Imported Completed", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
        dgDetail.Columns.Clear()
        TxtBrowse.Text = ""
        TxtDescription.Text = ""

ACCPACErrorHandlerPOST:

        Dim lCount As Long
        Dim lIndex As Long

        If SageSess.Errors Is Nothing Then
            MsgBox(Err.Description)
        Else
            lCount = SageSess.Errors.Count
            If lCount = 0 Then
                'MsgBox("Added", MsgBoxStyle.Information, "Done")
                ' MsgBox(Err.Description)
                'MsgBox(SageSess.Errors.Item(lIndex).Message)
            Else
                For lIndex = 0 To lCount - 1
                    MsgBox(SageSess.Errors.Item(lIndex).Message, MsgBoxStyle.Information, "Error")
                Next
                SageSess.Errors.Clear()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim s As String = "000123456781963500C"
        Dim C As String = Strings.Right(s, 1)

        'Dim v As Decimal = Strings.Left(s, 15) & "." & Strings.Right(CDec(Val(s)), 3)
        'MessageBox.Show(C & v)
        Dim t As String = Replace((Strings.Right(s, 1)), "C", "-") & Strings.Left(s, 15) & "." & Strings.Right(CDec(Val(s)), 3)
        MessageBox.Show(t.ToString)

    End Sub

    Private Sub AP_Invoice_Batch_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub
End Class