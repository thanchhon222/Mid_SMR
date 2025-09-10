Public Class frmHome
    Inherits Form

    Public Sub New()
        InitializeComponent()
        Me.IsMdiContainer = True

        Dim mdiclient As MdiClient = Me.Controls.OfType(Of MdiClient)().FirstOrDefault()
        If mdiclient IsNot Nothing Then
            mdiclient.BackColor = Color.DarkCyan
        End If
    End Sub


    Private Sub frmHome_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set default color in Load event handler

        frmConnectionLoading.Show()
        frmConnectionLoading.Hide()

        MenuStrip1.Visible = False

        ' Create and show a child form
        Dim childForm As New frmLogin()
        childForm.MdiParent = Me
        childForm.Show()


    End Sub

    Private Sub ShowMainMenuToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowMainMenuToolStripMenuItem.Click
        ' Create and show a child form
        Dim childForm As New frmModule()
        childForm.MdiParent = Me
        childForm.Show()

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click
        MessageBox.Show("Developed by Positron Multiverse. Please contact us Hotline: +855 17 222 009")
    End Sub

    Private Sub VersionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VersionToolStripMenuItem.Click
        MessageBox.Show("v1.01")
    End Sub

End Class