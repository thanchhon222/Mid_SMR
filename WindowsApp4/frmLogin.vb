Imports ACCPAC.Advantage.Session
Imports ACCPAC.Advantage
Imports Mid_SMR.SQLGetData
Imports AccpacCOMAPI
Public Class frmLogin
    Public Shared SageSess As New Session

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        '  On Error GoTo ACCPACErrorHandler

        Try

            Dim SageSess As New AccpacSession
            SageSess.Init("", "XY", "XY1000", "70A")
            SageSess.Open(TextBox1.Text.Trim(), (TextBox2.Text.Trim()), (ReadConnectionNotedPage.DBName_St.Trim()), Today, 0, "")
            ' SageSess.Open(UCase(txtUsername.Text.Trim()), UCase(txtPwd.Text.Trim.Trim()), (txtCompanyName.Text.Trim()), Today, 0)
            '   Dim mDBLinkCmpRW As DBLink
            '  mDBLinkCmpRW = SageSess.OpenDBLink(DBLinkType.Company, DBLinkFlags.ReadWrite)

            Dim mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkCmpRW = SageSess.OpenDBLink(tagDBLinkTypeEnum.DBLINK_COMPANY, tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)


            Me.Hide()

            SageSess.Close()

            ' Create and show a child form
            Dim childForm As New frmModule()
            childForm.MdiParent = frmHome
            childForm.Username = TextBox1.Text
            childForm.Password = TextBox2.Text

            childForm.Show()
            frmHome.MenuStrip1.Visible = True


        Catch ex As Exception
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        frmHome.Close()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        TextBox2.PasswordChar = ""
        PictureBox2.Hide()
        PictureBox3.Show()
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        TextBox2.PasswordChar = "*"
        PictureBox2.Show()
        PictureBox3.Hide()
    End Sub

    Private Sub frmLogin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PictureBox2.Show()
        PictureBox3.Hide()
    End Sub
End Class