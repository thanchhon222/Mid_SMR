<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_InventoryCount
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_InventoryCount))
        Me.btn_download = New System.Windows.Forms.Button()
        Me.btn_Upload = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'btn_download
        '
        Me.btn_download.Location = New System.Drawing.Point(26, 40)
        Me.btn_download.Name = "btn_download"
        Me.btn_download.Size = New System.Drawing.Size(134, 48)
        Me.btn_download.TabIndex = 0
        Me.btn_download.Text = "Download"
        Me.btn_download.UseVisualStyleBackColor = True
        '
        'btn_Upload
        '
        Me.btn_Upload.Location = New System.Drawing.Point(195, 40)
        Me.btn_Upload.Name = "btn_Upload"
        Me.btn_Upload.Size = New System.Drawing.Size(134, 48)
        Me.btn_Upload.TabIndex = 1
        Me.btn_Upload.Text = "Upload"
        Me.btn_Upload.UseVisualStyleBackColor = True
        '
        'frm_InventoryCount
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(371, 147)
        Me.Controls.Add(Me.btn_Upload)
        Me.Controls.Add(Me.btn_download)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frm_InventoryCount"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Inventory Count"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btn_download As Button
    Friend WithEvents btn_Upload As Button
End Class
