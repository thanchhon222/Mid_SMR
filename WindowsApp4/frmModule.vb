Public Class frmModule

    Public Property Username As String
    Public Property Password As String
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        ' Add nodes at design time
        Dim AdminNode As TreeNode = TreeView1.Nodes.Add("Administrator")
        AdminNode.Nodes.Add("Connection")

        Dim GLBatch As TreeNode = TreeView1.Nodes.Add("GL Batch Import")
        GLBatch.Nodes.Add("GL Batch To Sage 300")

        Dim AP As TreeNode = TreeView1.Nodes.Add("AP Module")
        AP.Nodes.Add("AP Inovice To Sage 300")


        Dim Revenue As TreeNode = TreeView1.Nodes.Add("Hotel Revenue")
        Revenue.Nodes.Add("Revenue To Sage 300")
        Revenue.Nodes.Add("Statistic To Sage 300")
        Revenue.Nodes.Add("Market Segment To Sage 300")

        Dim purchaseNode As TreeNode = TreeView1.Nodes.Add("IC Inventory")

        purchaseNode.Nodes.Add("IC - Receipt")
        purchaseNode.Nodes.Add("IC - Internal Usage")
        purchaseNode.Nodes.Add("IC - Inventory Count")

        ' Revenue.Nodes.Add("Import Revenue To AR")

        ' Set the font for main nodes to make the text bold
        AdminNode.NodeFont = New Font(TreeView1.Font, FontStyle.Bold)
        purchaseNode.NodeFont = New Font(TreeView1.Font, FontStyle.Bold)
        Revenue.NodeFont = New Font(TreeView1.Font, FontStyle.Bold)
        GLBatch.NodeFont = New Font(TreeView1.Font, FontStyle.Bold)
        AP.NodeFont = New Font(TreeView1.Font, FontStyle.Bold)
        ' GLBatch.ForeColor = Color.LightCyan ' Or any other color


        ' Calculate the maximum height needed for the nodes based on the font
        Dim maxNodeHeight As Integer = CInt(Math.Ceiling(Math.Max(AdminNode.NodeFont.Height, purchaseNode.NodeFont.Height)))

        ' Set the ItemHeight property to accommodate the text of all nodes
        TreeView1.Font = New Font("Arial", 10)
        TreeView1.ItemHeight = maxNodeHeight + 12 ' You can adjust the value as needed

    End Sub

    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick

        For Each node As TreeNode In TreeView1.Nodes
            If node IsNot e.Node Then
                node.Collapse()
            Else
                e.Node.Expand()
            End If
        Next

        If e.Node.Text = "IC - Internal Usage" Then
            Dim childForm As New Internal_Usage()
            ' Me.MdiParent = frmHome
            childForm.MdiParent = frmHome
            childForm.Username = Username
            childForm.Password = Password
            childForm.WindowState = FormWindowState.Normal
            childForm.Show()
        End If
        If e.Node.Text = "IC - Receipt" Then
            Dim childForm As New frmICReceipt()
            ' Me.MdiParent = frmHome
            childForm.MdiParent = frmHome
            childForm.Username = Username
            childForm.Password = Password
            childForm.WindowState = FormWindowState.Normal
            childForm.Show()
        End If
        If e.Node.Text = "IC - Inventory Count" Then
            Dim childForm As New frm_InventoryCount()
            ' Me.MdiParent = frmHome
            childForm.MdiParent = frmHome
            childForm.WindowState = FormWindowState.Normal
            childForm.Show()
        End If

        If e.Node.Text = "GL Batch To Sage 300" Then
            Dim childForm As New frmExcelToGL()
            ' Me.MdiParent = frmHome
            childForm.MdiParent = frmHome
            childForm.Username = Username
            childForm.Password = Password
            childForm.WindowState = FormWindowState.Normal
            childForm.Show()
        End If

        If e.Node.Text = "Revenue To Sage 300" Then
            Dim childForm As New frmRevenue()
            ' Me.MdiParent = frmHome
            childForm.MdiParent = frmHome
            childForm.Username = Username
            childForm.Password = Password
            childForm.WindowState = FormWindowState.Normal
            childForm.Show()
        End If

        If e.Node.Text = "Statistic To Sage 300" Then
            Dim childForm As New frmStatistic()
            ' Me.MdiParent = frmHome
            childForm.MdiParent = frmHome
            childForm.Username = Username
            childForm.Password = Password
            childForm.WindowState = FormWindowState.Normal
            childForm.Show()
        End If

        If e.Node.Text = "Market Segment To Sage 300" Then
            Dim childForm As New Market_Segment()
            ' Me.MdiParent = frmHome
            childForm.MdiParent = frmHome
            childForm.Username = Username
            childForm.Password = Password
            childForm.WindowState = FormWindowState.Normal
            childForm.Show()
        End If
        If e.Node.Text = "AP Inovice To Sage 300" Then
            Dim childForm As New AP_Invoice_Batch()
            ' Me.MdiParent = frmHome
            childForm.MdiParent = frmHome
            childForm.Username = Username
            childForm.Password = Password
            childForm.WindowState = FormWindowState.Normal
            childForm.Show()
        End If

        If e.Node.Text = "Connection" Then
            Dim childForm As New frmConnectionLoading()
            childForm.MdiParent = frmHome
            childForm.WindowState = FormWindowState.Normal
            childForm.Show()

        End If

    End Sub

End Class