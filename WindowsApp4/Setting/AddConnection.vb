Imports Microsoft.Win32
Imports System.Data.SqlClient
Imports System.IO

Module AddConnection



    Public v_Date As String = Nothing

    Public V_ProjectName As String = "School_Mgt"
    Public V_LoginStatus As Boolean = False
    Public V_UserID As String = Nothing
    Public FileINI = Application.StartupPath + "\School_Mgt.ini"
    Public v_StudID As String = Nothing
    Public v_ItemNO As String = Nothing
    Public v_NoFormart As String = "###,###.00"


    Private V_LastUserLogin = "LastUserLogin"
    Public V_FindUser As String = Nothing
    Public SQLCNN As New SqlConnection
    Private Section = "Settings"

    ''***********************SAP Connection Bong Mara***************************
    ' Connection SAP for SDK Import Data
    'Public SAPServer = "192.168.1.9"
    'Public SAPDB = "LSH_PRD_US"
    'Public SAPUser = "manager"
    'Public SAPPwd = "Lsh@12345"
    'Public SQLUser = "sa"
    'Public SQLPwd = "LSH@123"
    'Public DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019

    ' Connection SAP DB for Download Data
    'Private ServerName = "192.168.1.9"
    'Private DBUser = "sa"
    'Private DBPassword = "LSH@123"
    'Private DBName = "LSH_PRD_US"


    ''Connection to Stock Count DB
    'Private ServerName_St = "DESKTOP-FLICUTA"
    'Private DBUser_St = "sa"
    'Private DBPassword_St = "123456"
    'Private DBName_St = "StockCount"

    '***********************SAP Connection Than***************************
    '' Connection SAP for SDK Import Data
    'Public SAPServer = ReadConnectionNotedPage.ServerName_Sap
    'Public SAPDB = ReadConnectionNotedPage.DBName_Sap
    'Public SAPUser = ReadConnectionNotedPage.SAPUser_Sap
    'Public SAPPwd = ReadConnectionNotedPage.SAPPassword_Sap
    'Public SQLUser = ReadConnectionNotedPage.DBUser_Sap
    'Public SQLPwd = ReadConnectionNotedPage.DBPassword_Sap
    'Public DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019

    ' Connection SAP DB for Download Data
    Private ServerName = ReadConnectionNotedPage.ServerName_St
    Private DBName = ReadConnectionNotedPage.DBName_St
    Private DBUser = ReadConnectionNotedPage.DBUser_St
    Private DBPassword = ReadConnectionNotedPage.DBPassword_St


    'Connection to Stock Count DB
    Private ServerName_St = ReadConnectionNotedPage.ServerName_St '  "chhonthan"
    Private DBName_St = ReadConnectionNotedPage.DBName_St ' "StockCount"
    Private DBUser_St = ReadConnectionNotedPage.DBUser_St ' "sa"
    Private DBPassword_St = ReadConnectionNotedPage.DBPassword_St '  "123456"


    '***************************************************************************

    'Public V_CNNstring As String = "Data Source=10.25.1.35;Initial Catalog=eLabel;Persist Security Info=True;User ID=eLabel;Password=##eLabel!@#8899QQ"

    '' SAP Database
    'Public V_CNNstring As String = "Data Source= " & SAPServer & ";Initial Catalog=" & SAPDB & ";Persist Security Info=True;User ID=" & SQLUser & ";Password=" & SQLPwd & ""

    ' Stock Database
    Public V_CNNstringStock As String = "Data Source= " & ServerName_St & ";Initial Catalog=" & DBName_St & ";Persist Security Info=True;User ID=" & DBUser_St & ";Password=" & DBPassword_St & ""

    ' Central Database
    Public V_CNNstring_Central As String = "Data Source= " & ServerName & ";Initial Catalog=" & DBName & ";Persist Security Info=True;User ID=" & DBUser & ";Password=" & DBPassword & ""


    '***********************Email Setting***************************
    Public vSmtpMailServer As String = Nothing
    Public vIncommingPort As String = Nothing
    Public vOutgoingPort As String = Nothing
    Public vDispayName As String = Nothing
    Public vMailAddress As String = Nothing
    Public vUserID As String = Nothing
    Public vPassword As String = Nothing
    Public vSSL As Boolean = False

    '***********************SQLCNN***************************
    Public Function Write_CNN(ByVal pServerName As String, ByVal pDBUser As String, ByVal pDBPassword As String, ByVal pDBName As String) As Boolean
        Try
            WriteIni(FileINI, Section, ServerName, Encryption(pServerName))
            WriteIni(FileINI, Section, DBUser, Encryption(pDBUser))
            WriteIni(FileINI, Section, DBPassword, Encryption(pDBPassword))
            WriteIni(FileINI, Section, DBName, Encryption(pDBName))
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function


    Public Function Read_CNN_CT() As String
        Try
            Return V_CNNstring_Central

        Catch ex As Exception
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Public Function Read_CNN_St() As String
        Try
            Return V_CNNstringStock

        Catch ex As Exception
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Public Function Read_CNN() As String
        Try

            '   Return V_CNNstring

        Catch ex As Exception
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function


    Public Function WriteLastLogin(ByVal LastUserLogin As String) As Boolean
        Try
            WriteIni(FileINI, Section, V_LastUserLogin, Encryption(LastUserLogin))
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message, V_ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    'Public Function ReadLastLogin() As String
    '    Try
    '        rLastUserLogin = Decryption(ReadIni(FileINI, Section, V_LastUserLogin, ""))
    '        Return rLastUserLogin
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message, ProjectName, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Return Nothing
    '    End Try
    'End Function




    Private Declare Unicode Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String,
        ByVal lpKeyName As String, ByVal lpString As String,
        ByVal lpFileName As String) As Int32

    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String,
        ByVal lpKeyName As String, ByVal lpDefault As String,
        ByVal lpReturnedString As String, ByVal nSize As Int32,
        ByVal lpFileName As String) As Int32

    Public Sub WriteIni(ByVal iniFileName As String, ByVal Section As String, ByVal ParamName As String, ByVal ParamVal As String)
        Dim Result As Integer = WritePrivateProfileString(Section, ParamName, ParamVal, iniFileName)
    End Sub

    Public Function ReadIni(ByVal IniFileName As String, ByVal Section As String, ByVal ParamName As String, ByVal ParamDefault As String) As String
        Dim ParamVal As String = Space$(1024)
        Dim LenParamVal As Long = GetPrivateProfileString(Section, ParamName, ParamDefault, ParamVal, Len(ParamVal), IniFileName)
        ReadIni = Left$(ParamVal, LenParamVal)
    End Function


    Public Function Encryption(ByVal StrConnection As String) As String
        Dim i As Integer
        Dim st As String = ""
        For i = 1 To Len(StrConnection)
            st = st & Chr(Asc(Mid(StrConnection, i, 1)) + 130)
        Next
        Encryption = st
    End Function

    Public Function Decryption(ByVal StrEncriptConnection As String) As String
        Dim i As Integer
        Dim st As String = ""
        For i = 1 To Len(StrEncriptConnection)
            st = st & Chr(Asc(Mid(StrEncriptConnection, i, 1)) - 130)
        Next
        Decryption = st
    End Function

End Module
