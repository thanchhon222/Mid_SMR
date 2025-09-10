Imports System.Text
Imports ACCPAC.Advantage
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.CompilerServices
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Public Class Market_Segment
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
        Try

            Try
                Me.OpenFileDialog1.Title = "Choose a text file"
                Me.OpenFileDialog1.DefaultExt = "txt"
                Me.OpenFileDialog1.Filter = "*.txt|*.txt"
                Me.OpenFileDialog1.Multiselect = False

                If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                    FILENAME = OpenFileDialog1.FileName
                    Me.TxtBrowse.Text = FILENAME
                End If
                LoadData()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub LoadData()

        Dim numCol As Integer = 8
        Dim col1 As Integer = 15
        Dim col2 As Integer = 17
        Dim col3 As Integer = 15
        Dim col4 As Integer = 20
        Dim col5 As Integer = 10
        Dim col6 As Integer = 15
        Dim col7 As Integer = 174
        Dim col8 As Integer = 15
        Dim col9 As Integer = 15
        'Dim col11 As Integer


        Try
            Dim colCount As Integer = 10
            Dim cols(colCount) As Integer
            cols(0) = 15
            cols(1) = 17
            cols(2) = 15
            cols(3) = 20
            cols(4) = 10
            cols(5) = 15
            cols(6) = 174
            cols(7) = 15
            cols(8) = 15
            cols(9) = 15

            Dim dt As DataTable = New DataTable()
            Dim row As DataRow

            'Ádd Columns to dataTable
            For j As Integer = 1 To colCount
                dt.Columns.Add(j.ToString())
            Next

            Dim lines = File.ReadAllLines(FILENAME)
            ''Dim rowCount As Integer = 0
            Dim lineNum As Integer = 1
            Dim rowStart As Integer = 1
            For Each line In lines

                'Read data from rowStart
                If lineNum >= rowStart Then
                    Dim txt = line.ToString() 'Get string each line
                    Dim strStart As Integer = 0
                    Dim strLength As Integer = 0

                    row = dt.NewRow()
                    For i As Integer = 0 To colCount - 1

                        'First col
                        If i = 0 Then
                            strStart = 0 'First column, substring start from 0
                        Else
                            strStart = cols(i - 1)

                            If strStart > txt.Length Then
                                strStart = txt.Length
                            End If
                            If txt.Length = 0 Then
                                strStart = 0
                            End If
                        End If
                        'Return string after removing previus col
                        txt = txt.Substring(strStart)

                        strLength = cols(i) 'strLength=charactor length of each column
                        If strLength > txt.Length Then
                            strLength = txt.Length
                        End If
                        If txt.Length = 0 Then
                            strLength = 0
                        End If

                        Dim value As String = txt.Substring(0, strLength)
                        If i = 3 Then

                            Dim v As String = value.Trim()
                            Dim amountStr As String = v.Substring(0, v.Length - 1)
                            Dim c As String = v.Substring(v.Length - 1, 1)
                            Dim n As Decimal = If((c = "C"), -1, 1)
                            Dim amount As Decimal = Decimal.Parse(amountStr)
                            amount = amount / 1000
                            amount = amount * n
                            value = amount.ToString("#0.000")
                        End If
                        row(i) = value.Trim()

                    Next

                    dt.Rows.Add(row)
                End If


                lineNum += 1
            Next

            DataGridView1.DataSource = dt
            ' Apply Grid View styling
            With DataGridView1
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
            MessageBox.Show(ex.Message)
        End Try
        Dim sum As Decimal
        For i As Decimal = 0 To DataGridView1.Rows.Count() - 1 Step +1
            sum = sum + DataGridView1.Rows(i).Cells(3).Value
        Next
        txtvariance.Text = sum.ToString
    End Sub

    Private Sub BtnImport_Click(sender As Object, e As EventArgs) Handles BtnImport.Click
        If DataGridView1.RowCount > 1 Then
            On Error GoTo ACCPACErrorHandlerPOST

            SageSess.Init("", "XY", "XY1000", "72A")
            SageSess.Open(Username, Password, ReadConnectionNotedPage.DBName_St, Today, 0)

            Dim mDBLinkCmpRW As ACCPAC.Advantage.DBLink
            mDBLinkCmpRW = SageSess.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)


            Dim GLbatch As ACCPAC.Advantage.View = mDBLinkCmpRW.OpenView("GL0008")      ' The view of the batch list in Accpac
            Dim GLheader As ACCPAC.Advantage.View = mDBLinkCmpRW.OpenView("GL0006")     ' Is this the header for the batch or for once specific invoice in the batch 
            Dim GLdetail1 As ACCPAC.Advantage.View = mDBLinkCmpRW.OpenView("GL0010")
            Dim GLdetail2 As ACCPAC.Advantage.View = mDBLinkCmpRW.OpenView("GL0402")
            ' This isCreating the data structure or some sort in memory of some sort of how th batch will be structured
            GLbatch.Compose(New ACCPAC.Advantage.View() {GLheader})
            GLheader.Compose(New ACCPAC.Advantage.View() {GLbatch, GLdetail1})
            GLdetail1.Compose(New ACCPAC.Advantage.View() {GLheader, GLdetail2})
            GLdetail2.Compose(New ACCPAC.Advantage.View() {GLdetail1})
            ' GLbatch.Browse("((BATCHSTAT = ""1"" OR BATCHSTAT = ""6"" OR BATCHSTAT = ""9""))", 1)''
            GLbatch.RecordCreate(1) ' Create the actual batch (ViewRecordCreate.Insert )

            Dim rowNum As Integer = 1
            Dim i As Integer = 0
            GLbatch.Fields.FieldByName("PROCESSCMD").SetValue("1", False)        ' Lock Batch Switch
            GLbatch.Process()
            GLheader.RecordCreate(2)
            GLbatch.Fields.FieldByName("BTCHDESC").SetValue("Daily Market Segment for " & DateTimePicker1.Value.ToString("dd-MMM-yy "), False)   ' Description
            'GLbatch.Fields.FieldByName("BTCHDESC").SetValue(TxtDescription.Text.ToString, False)   ' Description
            GLbatch.Update()
            GLheader.Fields.FieldByName("DOCDATE").SetValue(DateTimePicker1.Value.ToString, False)    ' Document Date
            GLheader.Fields.FieldByName("DATEENTRY").SetValue(DateTimePicker1.Value.ToString, False)    ' Posting Date
            Dim FSCSPERD As Integer = DateTimePicker1.Value.Month
            Dim FSCSYEAR As Integer = DateTimePicker1.Value.Year
            GLheader.Fields.FieldByName("FSCSPERD").SetValue(FSCSPERD, False)                        ' Fiscal Period
            GLheader.Fields.FieldByName("FSCSYR").SetValue(FSCSYEAR, False)                        ' Fiscal Period
            GLheader.Fields.FieldByName("SRCETYPE").SetValue("JE", False)                     ' Source Type
            'GLBATCH3header.Fields("SRCETYPE").Value = "JE"                         ' Source Type
            For i = 0 To DataGridView1.Rows.Count - 1
                GLdetail1.RecordClear()
                GLdetail1.RecordCreate(0)
                GLdetail1.Fields.FieldByName("ACCTID").SetValue(DataGridView1.Rows(i).Cells(0).Value.ToString, False)                  ' Account Number
                GLdetail1.Fields.FieldByName("PROCESSCMD").SetValue("0", False)      ' Process switches
                GLdetail1.Fields.FieldByName("TRANSDESC").SetValue(DataGridView1.Rows(i).Cells(6).Value.ToString, False)   ' Description
                'GLdetail1.Fields.FieldByName("TRANSREF").SetValue("Daily Revenue " & DateTimePicker1.Value.ToString("dd-MMM-yy "), False)   ' Reference
                GLdetail1.Fields.FieldByName("TRANSREF").SetValue(DataGridView1.Rows(i).Cells(5).Value.ToString, False)   ' Reference
                GLdetail1.Process()
                GLdetail1.Fields.FieldByName("SCURNAMT").SetValue((DataGridView1.Rows(i).Cells(3).Value.ToString), False)                 ' Source Currency Amount
                'GLdetail1.Fields.FieldByName("COMMENT").SetValue(DataGridView1.Rows(i).Cells(0).Value.ToString, False)   ' Comment
                GLdetail1.Insert()
            Next
            ''********** Need to be enable if use direct with SQL
            ''GLheader.Fields.FieldByName("JRNLDESC").SetValue("Entry Description", False)   ' Description
            If txtvariance.Text.Trim <> "" Then
                ' GLdetail1.RecordClear()
                GLdetail1.RecordCreate(0)
                GLdetail1.Fields.FieldByName("ACCTID").SetValue(txtGlCodeoffset.Text.Trim, False)                  ' Account Number
                GLdetail1.Fields.FieldByName("PROCESSCMD").SetValue("0", False)      ' Process switches
                GLdetail1.Fields.FieldByName("TRANSDESC").SetValue("Daily Market Segment for " & DateTimePicker1.Value.ToString("dd-MMM-yy "), False)   ' Description
                'GLdetail1.Fields.FieldByName("TRANSREF").SetValue("Reference", False)   ' Reference
                GLdetail1.Process()
                GLdetail1.Fields.FieldByName("SCURNAMT").SetValue(-txtvariance.Text.Trim, False)                 ' Source Currency Amount
                'GLdetail1.Fields.FieldByName("COMMENT").SetValue("comment", False)   ' Comment
                GLdetail1.Insert()
            End If

            GLheader.Fields.FieldByName("JRNLDESC").SetValue("" & TxtDescription.Text & " " & DateTimePicker1.Value.ToString("dd-MMM-yy "), False)   ' Description
            GLheader.Insert()
            SageSess = Nothing
            SageSess = New ACCPAC.Advantage.Session
            DataGridView1.DataSource = Nothing
            txtvariance.Text = ""
            TxtBrowse.Text = ""
            TxtDescription.Text = ""

ACCPACErrorHandlerPOST:
            Dim lCount As Long
            Dim lIndex As Long
            If SageSess.Errors Is Nothing Then
                MessageBox.Show("Completed")
            Else
                lCount = SageSess.Errors.Count
                If lCount = 0 Then
                    MsgBox(Err.Description)
                    MsgBox(SageSess.Errors.Item(lIndex).Message)
                Else
                    For lIndex = 0 To lCount - 1
                        MsgBox(SageSess.Errors.Item(lIndex).Message)
                    Next
                    SageSess.Errors.Clear()
                End If
                Resume Next
            End If
        Else
            MessageBox.Show("No data!!!")
        End If
    End Sub


End Class