<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_InventoryDownload
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_InventoryDownload))
        Me.btn_GetData = New System.Windows.Forms.Button()
        Me.btn_SavetoXml = New System.Windows.Forms.Button()
        Me.dgDetail = New System.Windows.Forms.DataGridView()
        Me.Column1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txt_Location = New System.Windows.Forms.TextBox()
        Me.btn_Find = New System.Windows.Forms.Button()
        CType(Me.dgDetail, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btn_GetData
        '
        Me.btn_GetData.Location = New System.Drawing.Point(12, 12)
        Me.btn_GetData.Name = "btn_GetData"
        Me.btn_GetData.Size = New System.Drawing.Size(138, 36)
        Me.btn_GetData.TabIndex = 0
        Me.btn_GetData.Text = "Get Data"
        Me.btn_GetData.UseVisualStyleBackColor = True
        '
        'btn_SavetoXml
        '
        Me.btn_SavetoXml.Location = New System.Drawing.Point(166, 12)
        Me.btn_SavetoXml.Name = "btn_SavetoXml"
        Me.btn_SavetoXml.Size = New System.Drawing.Size(138, 36)
        Me.btn_SavetoXml.TabIndex = 1
        Me.btn_SavetoXml.Text = "Save to XML"
        Me.btn_SavetoXml.UseVisualStyleBackColor = True
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
        Me.dgDetail.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column1, Me.Column2, Me.Column3, Me.Column4, Me.Column5, Me.Column6, Me.Column7})
        Me.dgDetail.Location = New System.Drawing.Point(12, 111)
        Me.dgDetail.Margin = New System.Windows.Forms.Padding(2)
        Me.dgDetail.Name = "dgDetail"
        Me.dgDetail.ReadOnly = True
        Me.dgDetail.RowHeadersWidth = 51
        Me.dgDetail.RowTemplate.Height = 24
        Me.dgDetail.Size = New System.Drawing.Size(777, 328)
        Me.dgDetail.TabIndex = 2
        '
        'Column1
        '
        Me.Column1.HeaderText = "Item Code"
        Me.Column1.Name = "Column1"
        Me.Column1.ReadOnly = True
        Me.Column1.Width = 150
        '
        'Column2
        '
        Me.Column2.HeaderText = "BarCode"
        Me.Column2.Name = "Column2"
        Me.Column2.ReadOnly = True
        Me.Column2.Width = 150
        '
        'Column3
        '
        Me.Column3.HeaderText = "ItemDesc"
        Me.Column3.Name = "Column3"
        Me.Column3.ReadOnly = True
        '
        'Column4
        '
        Me.Column4.HeaderText = "StockUoM"
        Me.Column4.Name = "Column4"
        Me.Column4.ReadOnly = True
        Me.Column4.Width = 150
        '
        'Column5
        '
        Me.Column5.HeaderText = "QtyOnHand"
        Me.Column5.Name = "Column5"
        Me.Column5.ReadOnly = True
        '
        'Column6
        '
        Me.Column6.HeaderText = "CalCCost"
        Me.Column6.Name = "Column6"
        Me.Column6.ReadOnly = True
        '
        'Column7
        '
        Me.Column7.HeaderText = "Location"
        Me.Column7.Name = "Column7"
        Me.Column7.ReadOnly = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 73)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 13)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Location"
        '
        'txt_Location
        '
        Me.txt_Location.Location = New System.Drawing.Point(66, 69)
        Me.txt_Location.Name = "txt_Location"
        Me.txt_Location.Size = New System.Drawing.Size(169, 20)
        Me.txt_Location.TabIndex = 5
        '
        'btn_Find
        '
        Me.btn_Find.Location = New System.Drawing.Point(242, 69)
        Me.btn_Find.Name = "btn_Find"
        Me.btn_Find.Size = New System.Drawing.Size(62, 20)
        Me.btn_Find.TabIndex = 6
        Me.btn_Find.Text = "Find"
        Me.btn_Find.UseVisualStyleBackColor = True
        '
        'frm_InventoryDownload
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.btn_Find)
        Me.Controls.Add(Me.txt_Location)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgDetail)
        Me.Controls.Add(Me.btn_SavetoXml)
        Me.Controls.Add(Me.btn_GetData)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frm_InventoryDownload"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Inventory Count Download"
        CType(Me.dgDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btn_GetData As Button
    Friend WithEvents btn_SavetoXml As Button
    Friend WithEvents dgDetail As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents Column1 As DataGridViewTextBoxColumn
    Friend WithEvents Column2 As DataGridViewTextBoxColumn
    Friend WithEvents Column3 As DataGridViewTextBoxColumn
    Friend WithEvents Column4 As DataGridViewTextBoxColumn
    Friend WithEvents Column5 As DataGridViewTextBoxColumn
    Friend WithEvents Column6 As DataGridViewTextBoxColumn
    Friend WithEvents Column7 As DataGridViewTextBoxColumn
    Friend WithEvents txt_Location As TextBox
    Friend WithEvents btn_Find As Button
End Class
