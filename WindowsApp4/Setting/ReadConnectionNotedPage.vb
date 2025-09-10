Imports System.IO

Module ReadConnectionNotedPage

    ' Connection to Stock DB
    Public ServerName_St As String = ""
    Public DBUser_St As String = ""
    Public DBPassword_St As String = ""
    Public DBName_St As String = ""

    ' Connection to SAP DB
    Public ServerName_CT As String = ""
    Public DBUser_CT As String = ""
    Public DBPassword_CT As String = ""
    Public DBName_CT As String = ""

    ' Connection to SAP DB
    Public ServerName_Sap As String = ""
    Public DBUser_Sap As String = ""
    Public DBPassword_Sap As String = ""
    Public DBName_Sap As String = ""
    Public SAPUser_Sap As String = ""
    Public SAPPassword_Sap As String = ""
    Public server As String = ""

    ' Load configuration from file
    Public Sub LoadConfig()
        Try
            ' Specify the path to your configuration file
            Dim configFilePath As String = Application.StartupPath + "\config.txt"

            ' Read the configuration file
            Dim lines As String() = File.ReadAllLines(configFilePath)

            ' Parse the lines to set connection variables
            For Each line As String In lines
                Dim parts As String() = line.Split("="c)
                If parts.Length = 2 Then
                    Dim key As String = parts(0).Trim()
                    Dim value As String = parts(1).Trim()

                    Select Case key
                        Case "St_Server"
                            ServerName_St = value
                        Case "St_UserId"
                            DBUser_St = value
                        Case "St_Password"
                            DBPassword_St = value
                        Case "St_Database"
                            DBName_St = value

                            'Case "CT_Server"
                            '    ServerName_CT = value
                            'Case "CT_UserId"
                            '    DBUser_CT = value
                            'Case "CT_Password"
                            '    DBPassword_CT = value
                            'Case "CT_Database"
                            '    DBName_CT = value

                            'Case "SAP_Server"
                            '    ServerName_Sap = value
                            'Case "SAP_UserId"
                            '    DBUser_Sap = value
                            'Case "SAP_Password"
                            '    DBPassword_Sap = value
                            'Case "SAP_Database"
                            '    DBName_Sap = value
                            'Case "SAP_SAPUser"
                            '    SAPUser_Sap = value
                            'Case "SAP_SAPPassword"
                            '    SAPPassword_Sap = value
                            'Case "server"
                            '    server = value
                    End Select
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Error loading configuration: " & ex.Message)
        End Try

    End Sub


End Module
