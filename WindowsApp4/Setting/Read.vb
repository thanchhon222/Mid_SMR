Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Public Class Read
    Public Function GetDataFromExcel(ByVal a_sFilepath As String) As DataSet
        Dim ds As New DataSet()
        Dim cn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & a_sFilepath & ";Extended Properties= Excel 8.0")
        Try
            cn.Open()
        Catch ex As OleDbException
            Console.WriteLine(ex.Message)
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        Dim dt As New System.Data.DataTable()
        dt = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, Nothing)
        If dt IsNot Nothing OrElse dt.Rows.Count > 0 Then
            For sheet_count As Integer = 0 To dt.Rows.Count - 1
                Try
                    Dim sheetname As String = dt.Rows(sheet_count)("table_name").ToString()
                    Dim da As New OleDbDataAdapter("SELECT * FROM [" & sheetname & "]", cn)
                    da.Fill(ds, sheetname)
                Catch ex As DataException
                    Console.WriteLine(ex.Message)
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                End Try
            Next
        End If
        cn.Close()
        Return ds
    End Function
    Private Function GetData(ByVal a_dtData As System.Data.DataTable) As Object(,)
        Dim obj As Object(,) = New Object((a_dtData.Rows.Count + 1) - 1, a_dtData.Columns.Count - 1) {}
        Try
            For j As Integer = 0 To a_dtData.Columns.Count - 1
                obj(0, j) = a_dtData.Columns(j).Caption
            Next
            Dim dt As New DateTime()
            Dim sTmpStr As String = String.Empty
            For i As Integer = 1 To a_dtData.Rows.Count
                For j As Integer = 0 To a_dtData.Columns.Count - 1
                    If a_dtData.Columns(j).DataType Is dt.[GetType]() Then
                        If a_dtData.Rows(i - 1)(j) IsNot DBNull.Value Then
                            DateTime.TryParse(a_dtData.Rows(i - 1)(j).ToString(), dt)
                            obj(i, j) = dt.ToString("MM/dd/yy hh:mm tt")
                        Else
                            obj(i, j) = a_dtData.Rows(i - 1)(j)
                        End If
                    ElseIf a_dtData.Columns(j).DataType Is sTmpStr.[GetType]() Then
                        If a_dtData.Rows(i - 1)(j) IsNot DBNull.Value Then
                            sTmpStr = a_dtData.Rows(i - 1)(j).ToString().Replace(vbCr, "")
                            obj(i, j) = sTmpStr
                        Else
                            obj(i, j) = a_dtData.Rows(i - 1)(j)
                        End If
                    Else
                        obj(i, j) = a_dtData.Rows(i - 1)(j)
                    End If
                Next
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        Return obj
    End Function
End Class
