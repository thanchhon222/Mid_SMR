Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Mid_SMR.RGBColors
Imports Mid_SMR.SQLGetData
Public Class frm_Find_Location
    Dim TBL As DataTable = Nothing

    Private Sub frm_Find_Location_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            TBL = Nothing
            dgDetail.Rows.Clear()

            TBL = SQL_GetTable_St(" SELECT 
                                DISTINCT
                                RTRIM(B.[Location]) As 'Code',
                                RTRIM(B.[DESC]) As 'Location'
                                from ICWKUH A
                                INNER JOIN ICLOC B ON A.LOCATION = B.LOCATION")

            If TBL IsNot Nothing Then
                For Each DR As DataRow In TBL.Rows
                    dgDetail.Rows.Add(DR("Code"),
                                      DR("Location"))
                Next

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btn_select.Click
        Try
            If dgDetail.RowCount = 0 Then
                Exit Sub
            End If
            F_Barcode = dgDetail.CurrentRow.Cells("Column1").Value.ToString.Trim
            Me.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class