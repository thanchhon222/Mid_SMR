<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmICReceipt
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmICReceipt))
        Me.dgDetail = New System.Windows.Forms.DataGridView()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.B_close = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.dt1 = New System.Windows.Forms.DateTimePicker()
        Me.Txt_Desc = New System.Windows.Forms.TextBox()
        Me.txt_Ref = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.B_browse = New System.Windows.Forms.Button()
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
        Me.dgDetail.Location = New System.Drawing.Point(11, 115)
        Me.dgDetail.Margin = New System.Windows.Forms.Padding(2)
        Me.dgDetail.Name = "dgDetail"
        Me.dgDetail.ReadOnly = True
        Me.dgDetail.RowHeadersWidth = 51
        Me.dgDetail.RowTemplate.Height = 24
        Me.dgDetail.Size = New System.Drawing.Size(833, 328)
        Me.dgDetail.TabIndex = 1
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(102, 83)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(407, 20)
        Me.TextBox1.TabIndex = 3
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(733, 69)
        Me.Button2.Margin = New System.Windows.Forms.Padding(2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(104, 32)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Post to Sage"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'B_close
        '
        Me.B_close.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.B_close.Location = New System.Drawing.Point(733, 449)
        Me.B_close.Margin = New System.Windows.Forms.Padding(2)
        Me.B_close.Name = "B_close"
        Me.B_close.Size = New System.Drawing.Size(111, 34)
        Me.B_close.TabIndex = 6
        Me.B_close.Text = "Close"
        Me.B_close.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'dt1
        '
        Me.dt1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dt1.Location = New System.Drawing.Point(520, 31)
        Me.dt1.Margin = New System.Windows.Forms.Padding(2)
        Me.dt1.Name = "dt1"
        Me.dt1.Size = New System.Drawing.Size(99, 20)
        Me.dt1.TabIndex = 7
        '
        'Txt_Desc
        '
        Me.Txt_Desc.Location = New System.Drawing.Point(102, 31)
        Me.Txt_Desc.Margin = New System.Windows.Forms.Padding(2)
        Me.Txt_Desc.Name = "Txt_Desc"
        Me.Txt_Desc.Size = New System.Drawing.Size(382, 20)
        Me.Txt_Desc.TabIndex = 8
        '
        'txt_Ref
        '
        Me.txt_Ref.Location = New System.Drawing.Point(102, 54)
        Me.txt_Ref.Margin = New System.Windows.Forms.Padding(2)
        Me.txt_Ref.Name = "txt_Ref"
        Me.txt_Ref.Size = New System.Drawing.Size(382, 20)
        Me.txt_Ref.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(693, 18)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 13)
        Me.Label1.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(10, 31)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Description "
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(10, 54)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(57, 13)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Reference"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 87)
        Me.Label4.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(49, 13)
        Me.Label4.TabIndex = 132
        Me.Label4.Text = "Excel file"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(517, 9)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(89, 13)
        Me.Label5.TabIndex = 133
        Me.Label5.Text = "Transaction Date"
        '
        'B_browse
        '
        Me.B_browse.Location = New System.Drawing.Point(520, 83)
        Me.B_browse.Margin = New System.Windows.Forms.Padding(2)
        Me.B_browse.Name = "B_browse"
        Me.B_browse.Size = New System.Drawing.Size(78, 20)
        Me.B_browse.TabIndex = 4
        Me.B_browse.Text = "Browse..."
        Me.B_browse.UseVisualStyleBackColor = True
        '
        'frmICReceipt
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(855, 494)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txt_Ref)
        Me.Controls.Add(Me.Txt_Desc)
        Me.Controls.Add(Me.dt1)
        Me.Controls.Add(Me.B_close)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.B_browse)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.dgDetail)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "frmICReceipt"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "IC Receipt"
        CType(Me.dgDetail, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dgDetail As DataGridView
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button2 As Button
    Friend WithEvents B_close As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents dt1 As DateTimePicker
    Friend WithEvents Txt_Desc As TextBox
    Friend WithEvents txt_Ref As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents B_browse As Button
End Class
