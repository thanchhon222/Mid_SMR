Imports System.Data.SqlClient
Imports Mid_SMR.ReadConnectionNotedPage
Imports System.Diagnostics
Public Class frmConnectionLoading
    Private Sub texting_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '  ReadConnectionNotedPage.LoadConfig()

    End Sub

    Public Class ConnectionTester
        Public Shared Function TestConnection(serverName As String, dbName As String, username As String, password As String) As Boolean

            Dim connectionString As String = $"Data Source={serverName};Initial Catalog={dbName};User ID={username};Password={password};"
            Using connection As New SqlConnection(connectionString)
                Try
                    connection.Open()
                    Return True
                Catch ex As SqlException
                    Console.WriteLine($"Error: {ex.Message}")
                    Return False
                End Try
            End Using
        End Function

    End Class

    Public Sub connetesting()
        ReadConnectionNotedPage.LoadConfig()
        If ConnectionTester.TestConnection(ReadConnectionNotedPage.ServerName_St, ReadConnectionNotedPage.DBName_St, ReadConnectionNotedPage.DBUser_St, ReadConnectionNotedPage.DBPassword_St) Then

            frmHome.Show()
            Me.Hide()

        Else
            ' MessageBox.Show("Connection failed. Please check your connection settings.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Button1.Text = "Pleaase check your connection by click on the Setup Connection."

        End If

    End Sub
    Private Sub texting_LoadShown(sender As Object, e As EventArgs) Handles Me.Shown
        connetesting()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        connetesting()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim filePath As String = Application.StartupPath + "\config.txt"

        If IO.File.Exists(filePath) Then
            Process.Start(filePath)
        Else
            MessageBox.Show("The file does not exist.", "File Not  Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

End Class