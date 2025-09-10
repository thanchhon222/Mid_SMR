Public Class frm_InventoryCount
    Private Sub btn_download_Click(sender As Object, e As EventArgs) Handles btn_download.Click

        Dim childForm As New frm_InventoryDownload()
        ' Me.MdiParent = frmHome
        childForm.MdiParent = frmHome
        childForm.WindowState = FormWindowState.Normal
        childForm.Show()
    End Sub

    Private Sub btn_Upload_Click(sender As Object, e As EventArgs) Handles btn_Upload.Click
        Dim childForm As New frm_InventoryUpload()
        ' Me.MdiParent = frmHome
        childForm.MdiParent = frmHome
        childForm.WindowState = FormWindowState.Normal
        childForm.Show()
    End Sub
End Class