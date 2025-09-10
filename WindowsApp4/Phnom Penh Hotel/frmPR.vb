Imports Mid_SMR.SQLGetData
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices
Public Class frmPR
    Dim TBL As DataTable = Nothing

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        getVendor()
    End Sub
    Public Sub getVendor()
        Try
            TBL = Nothing
            DataGridView1.Rows.Clear()

            TBL = SQL_GetTable_St("select VENDORID,VENDNAME, rtrim(VendorName) As VendorName  from APVEN where VENDORID = '" & TextBox1.Text & "'")

            For Each DR As DataRow In TBL.Rows
                DataGridView1.Rows.Add(DR("VENDORID"),
                                       DR("VENDNAME"),
                                       DR("VendorName"))
            Next
            TBL = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            WriteError(ex.Message)
            Exit Sub
        End Try
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        getVendor()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim item_no As String = DataGridView1.CurrentRow.Cells("Column3").Value.ToString.Trim
            Dim VendorName As String = DataGridView1.CurrentRow.Cells("Column2").Value.ToString.Trim

            Using (SQLCNN)
                Dim sqlComm As New SqlCommand()
                sqlComm.Connection = SQLCNN
                sqlComm.CommandText = "PST_UP_Vendor"
                sqlComm.CommandType = CommandType.StoredProcedure
                sqlComm.Parameters.AddWithValue("@VendorID", item_no)
                sqlComm.Parameters.AddWithValue("@VendorName", VendorName)

                If SQLCNN.State <> ConnectionState.Closed Then SQLCNN.Close()
                SQLCNN.ConnectionString = Read_CNN_St()
                SQLCNN.Open()
                sqlComm.ExecuteNonQuery()
            End Using

            MessageBox.Show("Data Updated.")
            TextBox1.Text = ""

        Catch ex As Exception
            WriteError(ex.Message)
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub
End Class