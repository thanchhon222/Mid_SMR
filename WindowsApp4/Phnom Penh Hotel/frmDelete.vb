Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports Mid_SMR.SQLGetData
Public Class frmDelete
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try

            Dim record As Int32
            record = SQL_GetField_SC("select count(Status) As TranDate from z_tbl_revenue where TranDate = '" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' ")
            If record > 0 Then

                ' For Each row As DataGridViewRow In DataGridView1.Rows
                Dim sqlComm As New SqlCommand()
                sqlComm.Connection = SQLCNN
                sqlComm.CommandText = "[sp_Delete_Revenue]"
                sqlComm.CommandType = CommandType.StoredProcedure
                sqlComm.Parameters.AddWithValue("@TranDate", DateTimePicker1.Value.ToString("yyyy-MM-dd"))

                If SQLCNN.State <> ConnectionState.Closed Then SQLCNN.Close()
                SQLCNN.ConnectionString = Read_CNN_St()
                SQLCNN.Open()
                sqlComm.ExecuteNonQuery()
                MessageBox.Show("Data already deleted." & " Transaction Date : " & DateTimePicker1.Value.ToString("yyyy-MM-dd"))

            Else
                MessageBox.Show("No data for" & " Transaction Date : " & DateTimePicker1.Value.ToString("yyyy-MM-dd"))
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
    End Sub
End Class