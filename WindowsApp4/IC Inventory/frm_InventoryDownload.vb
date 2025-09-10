Imports System.Reflection.Emit
Imports System.IO
Imports System.Drawing
Imports System.Data.SqlClient
Imports Mid_SMR.RGBColors
Imports Mid_SMR.SQLGetData
Imports System.Diagnostics

Public Class frm_InventoryDownload
    Dim TBL As DataTable = Nothing

    Private Sub btn_Find_Click(sender As Object, e As EventArgs) Handles btn_Find.Click

        Try
            F_Barcode = ""
            Dim F As New frm_Find_Location
            F.StartPosition = FormStartPosition.CenterScreen
            F.ShowDialog()
            If F_Barcode <> "" Then
                txt_Location.Text = F_Barcode

            End If
            txt_Location.Focus()
        Catch ex As Exception
            WriteError(ex.Message)
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub btn_GetData_Click(sender As Object, e As EventArgs) Handles btn_GetData.Click

        Try

            TBL = Nothing
            dgDetail.Rows.Clear()

            TBL = SQL_GetTable_St(" SELECT 
                                    RTRIM(ITEMNO) As 'ItemCode',
                                    RTRIM(FMTITEMNO) As 'BarCode',
                                    RTRIM(a.[DESC]) As 'ItemDesc',
                                    RTRIM(A.STOCKUNIT) As 'StockUoM',
                                    CAST(QTYONHAND AS DECIMAL(18,2)) As 'QtyOnHand',
                                    CAST(CALCCOST AS DECIMAL(18,2)) As CalCCost,
                                    RTRIM(B.[DESC]) As 'Location'
                                    from ICWKUH A
                                    INNER JOIN ICLOC B ON A.LOCATION = B.LOCATION Where B.[Location] = '" & txt_Location.Text & "'")

            If TBL IsNot Nothing Then
                For Each DR As DataRow In TBL.Rows
                    dgDetail.Rows.Add(DR("ItemCode"),
                                      DR("Barcode"),
                                      DR("ItemDesc"),
                                      DR("StockUoM"),
                                      DR("QtyOnHand"),
                                      DR("CalCCost"),
                                      DR("Location"))
                Next

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try

    End Sub

    Private Sub btn_SavetoXml_Click(sender As Object, e As EventArgs) Handles btn_SavetoXml.Click
        ' 0. Guard clauses ----------------------------------------------------
        If TBL Is Nothing OrElse TBL.Rows.Count = 0 Then
            MessageBox.Show("No data to export – click Get Data first.",
                        V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim Location As String = txt_Location.Text

        ' 1. Let the user choose where to save -------------------------------
        Using sfd As New SaveFileDialog With {
        .Title = "Save inventory as XML",
        .Filter = "XML files (*.xml)|*.xml",
        .FileName = "" & Location & ".xml"}

            If sfd.ShowDialog(Me) <> DialogResult.OK Then Return
            Dim filePath As String = sfd.FileName

            ' 2. Build the XML ------------------------------------------------
            Dim xml As New XElement("items",
            From r In TBL.AsEnumerable()
            Select New XElement("item",
                New XAttribute("id", r.Field(Of String)("ItemCode").Trim()),
                New XElement("name", r.Field(Of String)("ItemDesc").Trim()),
                New XElement("Barcode", r.Field(Of String)("Barcode").Trim()),
                New XElement("StockUoM", r.Field(Of String)("StockUoM").Trim()),
                New XElement("QtyOnHand", r.Field(Of Decimal)("QtyOnHand")),
                New XElement("CalCCost", r.Field(Of Decimal)("CalCCost")),
                New XElement("Location", r.Field(Of String)("Location").Trim())
            )
        )

            ' 3. Save ---------------------------------------------------------
            xml.Save(filePath)

            ' 4. ADB Push to device ------------------------------------------
            Try

                Dim platformToolsPath As String = "C:\Android_Tools_platform_tools"
                Dim adbPath As String = Path.Combine(platformToolsPath, "adb.exe")
                Dim devicePath As String = "/sdcard/DCIM/StockCount/" & Location & ".xml"

                ' Check if ADB exists
                If Not File.Exists(adbPath) Then
                    MessageBox.Show($"ADB not found at:{Environment.NewLine}{adbPath}",
                                V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                ' Create process to run ADB commands
                Dim process As New Process()
                process.StartInfo.FileName = adbPath
                process.StartInfo.Arguments = $"devices"
                process.StartInfo.UseShellExecute = False
                process.StartInfo.RedirectStandardOutput = True
                process.StartInfo.CreateNoWindow = True
                process.Start()

                ' Check if any devices are connected
                Dim output As String = process.StandardOutput.ReadToEnd()
                process.WaitForExit()

                If Not output.Contains("device") Then
                    MessageBox.Show("No Android devices connected via ADB",
                                V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                ' Push the file to the device
                process = New Process()
                process.StartInfo.FileName = adbPath
                process.StartInfo.Arguments = $"push ""{filePath}"" ""{devicePath}"""
                process.StartInfo.UseShellExecute = False
                process.StartInfo.RedirectStandardOutput = True
                process.StartInfo.CreateNoWindow = True
                process.Start()

                output = process.StandardOutput.ReadToEnd()
                process.WaitForExit()

                MessageBox.Show($"XML saved to:{Environment.NewLine}{filePath}" &
                           $"{Environment.NewLine}{Environment.NewLine}" &
                           $"Successfully pushed to device at:{Environment.NewLine}{devicePath}",
                           V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                dgDetail.Rows.Clear()

            Catch ex As Exception
                MessageBox.Show($"Error pushing to device:{Environment.NewLine}{ex.Message}",
                           V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using

        txt_Location.Text = ""

    End Sub


End Class