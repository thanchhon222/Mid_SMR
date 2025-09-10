Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.Data.SqlClient
Imports System.IO


Public Class SQLGetData
    'Public Shared SQLCNN As New OleDb.OleDbConnection
    'Public Shared stSQLCNN As String = "Provider = SQLOledb.1; Data Source=203.196.89.102; User ID=sa; Password=Pa$$w0rd; Initial Catalog='LMGDAT'"
    'Dim ST_SQLCNN As String = "Data Source=203.196.89.100; User ID=sa; Password=Pa$$w0rd; Initial Catalog='TLMG'"

    'Public Shared V_CNN_FPC As String = "Data Source=10.25.1.53; User ID=sa; Password=DFI@dm1n; Initial Catalog='FPCPOS'"


    Public Shared Function SQL_GetDataSet(ByVal SqlString As String) As DataSet
        Try
            If SQLCNN.State = ConnectionState.Open Then SQLCNN.Close()
            SQLCNN.ConnectionString = Read_CNN()
            SQLCNN.Open()
            Dim CMD As New SqlCommand
            With CMD
                .Connection = SQLCNN
                .CommandTimeout = 0
                .CommandType = CommandType.Text
                .CommandText = SqlString
            End With
            Dim DAP As New SqlDataAdapter
            DAP.SelectCommand = CMD
            Dim ds As New DataSet
            DAP.Fill(ds)
            SQLCNN.Close()
            Return ds
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Shared Function SQL_GetTable(ByVal SqlString As String) As DataTable

        Try
            If SQLCNN.State = ConnectionState.Open Then SQLCNN.Close()
            SQLCNN.ConnectionString = Read_CNN()
            SQLCNN.Open()

            Dim CMD As New SqlCommand
            With CMD
                .Connection = SQLCNN
                .CommandTimeout = 0
                .CommandType = CommandType.Text
                .CommandText = SqlString
            End With
            Dim DAP As New SqlDataAdapter
            DAP.SelectCommand = CMD
            Dim ds As New DataSet
            DAP.Fill(ds)
            Dim Table As New DataTable
            Table = ds.Tables(0)
            SQLCNN.Close()
            Return Table
        Catch ex As Exception
            Return Nothing
        End Try

    End Function


    Public Shared Function SQL_GetTable_St(ByVal SqlString As String) As DataTable

        Try
            If SQLCNN.State = ConnectionState.Open Then SQLCNN.Close()
            SQLCNN.ConnectionString = Read_CNN_St()
            SQLCNN.Open()

            Dim CMD As New SqlCommand
            With CMD
                .Connection = SQLCNN
                .CommandTimeout = 0
                .CommandType = CommandType.Text
                .CommandText = SqlString
            End With
            Dim DAP As New SqlDataAdapter
            DAP.SelectCommand = CMD
            Dim ds As New DataSet
            DAP.Fill(ds)
            Dim Table As New DataTable
            Table = ds.Tables(0)
            SQLCNN.Close()
            Return Table
        Catch ex As Exception
            Return Nothing
        End Try

    End Function

    Public Shared Function SQL_GetTable_Ct(ByVal SqlString As String) As DataTable

        Try
            If SQLCNN.State = ConnectionState.Open Then SQLCNN.Close()
            SQLCNN.ConnectionString = Read_CNN_CT()
            SQLCNN.Open()

            Dim CMD As New SqlCommand
            With CMD
                .Connection = SQLCNN
                .CommandTimeout = 0
                .CommandType = CommandType.Text
                .CommandText = SqlString
            End With
            Dim DAP As New SqlDataAdapter
            DAP.SelectCommand = CMD
            Dim ds As New DataSet
            DAP.Fill(ds)
            Dim Table As New DataTable
            Table = ds.Tables(0)
            SQLCNN.Close()
            Return Table
        Catch ex As Exception
            Return Nothing
        End Try

    End Function


    Public Shared Function SQL_GetField(ByVal ST_SQL As String) As String
        Try
            If SQLCNN.State = ConnectionState.Open Then SQLCNN.Close()
            SQLCNN.ConnectionString = Read_CNN()
            SQLCNN.Open()
            Dim CMD As SqlCommand
            Dim DTReader As SqlDataReader
            CMD = New SqlCommand(ST_SQL, SQLCNN)
            DTReader = CMD.ExecuteReader()
            DTReader.Read()

            If DTReader.HasRows = True Then
                Return DTReader(0)
                DTReader.Close()
                SQLCNN.Close()
                Exit Function
            End If
            DTReader.Close()
            SQLCNN.Close()
            Return ""
        Catch ex As Exception
            WriteError(Err.Description)
            Return ""
        End Try
    End Function

    Public Shared Function SQL_GetField_SC(ByVal ST_SQL As String) As String
        Try
            If SQLCNN.State = ConnectionState.Open Then SQLCNN.Close()
            SQLCNN.ConnectionString = Read_CNN_St()
            Dim CMD As SqlCommand
            Dim DTReader As SqlDataReader

            SQLCNN.Open()
            CMD = New SqlCommand(ST_SQL, SQLCNN)
            DTReader = CMD.ExecuteReader()
            DTReader.Read()

            If DTReader.HasRows = True Then
                Return DTReader(0)
                DTReader.Close()
                SQLCNN.Close()
                Exit Function
            End If
            DTReader.Close()
            SQLCNN.Close()
            Return "False"
        Catch ex As Exception
            MessageBox.Show(ex.ToString & vbNewLine & ex.StackTrace, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return "False"
        End Try
    End Function

    Public Shared Function SQL_ExecuteNonQuery(ByVal ST_SQL As String) As Boolean
        Try
            If SQLCNN.State = ConnectionState.Open Then SQLCNN.Close()
            SQLCNN.ConnectionString = Read_CNN()
            SQLCNN.Open()
            Dim CMD As New SqlCommand(ST_SQL, SQLCNN)
            CMD.CommandTimeout = 0
            CMD.ExecuteNonQuery()
            SQLCNN.Close()
            Return True
        Catch ex As Exception
            WriteError(Err.Description)
            Return False
        End Try
    End Function

    Public Shared Sub SQL_ExecQuery(ByVal ST_SQL As String)
        Try
            If SQLCNN.State = ConnectionState.Open Then SQLCNN.Close()
            SQLCNN.Open()
            Dim CMD As New SqlCommand(ST_SQL, SQLCNN)
            CMD.CommandTimeout = 0
            CMD.ExecuteNonQuery()
            SQLCNN.Close()
        Catch ex As Exception
            WriteError(Err.Description)
        End Try
    End Sub

    Shared Sub WriteError(ByVal ST_Error As String)
        Try
            Dim ws As New StreamWriter(Application.StartupPath + "\LOG.txt", True, System.Text.Encoding.Default)
            ws.WriteLine("")
            ws.WriteLine("")
            ws.WriteLine("----------------------------------------------")
            ws.WriteLine(Format(Now(), "dd-MM-yyyy HH:mm:ss"))
            ws.WriteLine(ST_Error)
            ws.Close()
        Catch ex As Exception
            MessageBox.Show(Err.Description, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Shared Function NumToDate(ByVal fNum As String) As Date
        Dim ndate As New Windows.Forms.DateTimePicker
        Dim yr As Integer = CInt(Val(fNum.Substring(0, 4)))
        Dim mt As Integer = CInt(Val(fNum.Substring(4, 2)))
        Dim dy As Integer = CInt(Val(fNum.Substring(6, 2)))
        ndate.Value = DateSerial(yr, mt, dy)
        Return ndate.Value
    End Function

    Public Shared Function ReadMailSetting() As Boolean
        Try
            If SQLCNN.State = ConnectionState.Open Then SQLCNN.Close()
            SQLCNN.ConnectionString = Read_CNN()
            SQLCNN.Open()

            Dim CMD As New SqlCommand
            With CMD
                .Connection = SQLCNN
                .CommandTimeout = 0
                .CommandType = CommandType.Text
                .CommandText = "SELECT SmtpMailServer, IncommingPort, OutgoingPort, DispayName, MailAddress, UserID, Passw, SSL_Mode FROM sy_mail_setting"
            End With
            Dim DAP As New SqlDataAdapter
            DAP.SelectCommand = CMD
            Dim ds As New DataSet
            DAP.Fill(ds)
            Dim TBL As DataTable = ds.Tables(0)
            vSmtpMailServer = TBL.Rows(0)("SmtpMailServer").ToString.Trim
            vIncommingPort = TBL.Rows(0)("IncommingPort").ToString.Trim
            vOutgoingPort = TBL.Rows(0)("OutgoingPort").ToString.Trim
            vDispayName = TBL.Rows(0)("DispayName").ToString.Trim
            vMailAddress = TBL.Rows(0)("MailAddress").ToString.Trim
            vUserID = TBL.Rows(0)("UserID").ToString.Trim
            vPassword = TBL.Rows(0)("Passw").ToString.Trim
            vSSL = TBL.Rows(0)("SSL_Mode")
            SQLCNN.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function


End Class
