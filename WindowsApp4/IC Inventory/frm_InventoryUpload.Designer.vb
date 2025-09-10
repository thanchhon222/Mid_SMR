<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_InventoryUpload
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.dgDetail = New System.Windows.Forms.DataGridView()
        Me.btn_SavetoSage = New System.Windows.Forms.Button()
        Me.btn_GetData = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btn_Find = New System.Windows.Forms.Button()
        Me.txt_Location = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.dgDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgDetail
        '
        Me.dgDetail.AllowUserToAddRows = False
        Me.dgDetail.AllowUserToDeleteRows = False
        Me.dgDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgDetail.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight
        Me.dgDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgDetail.Location = New System.Drawing.Point(12, 111)
        Me.dgDetail.Margin = New System.Windows.Forms.Padding(2)
        Me.dgDetail.Name = "dgDetail"
        Me.dgDetail.ReadOnly = True
        Me.dgDetail.RowHeadersWidth = 51
        Me.dgDetail.RowTemplate.Height = 24
        Me.dgDetail.Size = New System.Drawing.Size(777, 328)
        Me.dgDetail.TabIndex = 7
        '
        'btn_SavetoSage
        '
        Me.btn_SavetoSage.Location = New System.Drawing.Point(464, 57)
        Me.btn_SavetoSage.Name = "btn_SavetoSage"
        Me.btn_SavetoSage.Size = New System.Drawing.Size(107, 36)
        Me.btn_SavetoSage.TabIndex = 6
        Me.btn_SavetoSage.Text = "Save to Sage"
        Me.btn_SavetoSage.UseVisualStyleBackColor = True
        '
        'btn_GetData
        '
        Me.btn_GetData.Location = New System.Drawing.Point(464, 21)
        Me.btn_GetData.Name = "btn_GetData"
        Me.btn_GetData.Size = New System.Drawing.Size(107, 21)
        Me.btn_GetData.TabIndex = 5
        Me.btn_GetData.Text = "Get Data"
        Me.btn_GetData.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(15, 28)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(44, 13)
        Me.Label4.TabIndex = 134
        Me.Label4.Text = "CSV file"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(90, 22)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(369, 20)
        Me.TextBox1.TabIndex = 133
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'btn_Find
        '
        Me.btn_Find.Location = New System.Drawing.Point(264, 57)
        Me.btn_Find.Name = "btn_Find"
        Me.btn_Find.Size = New System.Drawing.Size(62, 20)
        Me.btn_Find.TabIndex = 137
        Me.btn_Find.Text = "Find"
        Me.btn_Find.UseVisualStyleBackColor = True
        '
        'txt_Location
        '
        Me.txt_Location.Location = New System.Drawing.Point(88, 57)
        Me.txt_Location.Name = "txt_Location"
        Me.txt_Location.Size = New System.Drawing.Size(169, 20)
        Me.txt_Location.TabIndex = 136
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 62)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 13)
        Me.Label1.TabIndex = 135
        Me.Label1.Text = "Location"
        '
        'frm_InventoryUpload
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.btn_Find)
        Me.Controls.Add(Me.txt_Location)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.dgDetail)
        Me.Controls.Add(Me.btn_SavetoSage)
        Me.Controls.Add(Me.btn_GetData)
        Me.Name = "frm_InventoryUpload"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Inventory Count Upload"
        CType(Me.dgDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dgDetail As DataGridView
    Friend WithEvents btn_SavetoSage As Button
    Friend WithEvents btn_GetData As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents btn_Find As Button
    Friend WithEvents txt_Location As TextBox
    Friend WithEvents Label1 As Label
End Class
