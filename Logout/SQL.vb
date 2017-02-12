Imports System.Reflection
Imports System.IO
Imports MySql.Data.MySqlClient

Imports System.Net

Public Class SQL
    Public ReadOnly SQLUSername As String = "itprefect"
    Public ReadOnly SQLPassword As String = "lycmc2013"
    Public ReadOnly SQLIP As String = "itdc"
    Public ReadOnly SQLDB As String = "pc_room_manger_sys"

    Dim SQL_Conn As String = "server=" + SQLIP + ";Port=3306;userid=" + SQLUSername + ";password=" + SQLPassword + ";database=" + SQLDB + ";Character Set=utf8"
    Public ReadOnly SQLString As String = SQL_Conn

    Dim mysqlconn As MySql.Data.MySqlClient.MySqlConnection
    Dim reader As MySql.Data.MySqlClient.MySqlDataReader
    Dim command As New MySql.Data.MySqlClient.MySqlCommand
    Public Function SQL_Return(ByVal Cmd As String, ByVal rt As String)
        Try
            mysqlconn = New MySql.Data.MySqlClient.MySqlConnection
            mysqlconn.ConnectionString = SQL_Conn
            mysqlconn.Open()
            Dim Query As String

            Query = Cmd
            command = New MySql.Data.MySqlClient.MySqlCommand(Query, mysqlconn)
            reader = command.ExecuteReader
            Dim str As String = ""

            While reader.Read
                str = reader.GetString(rt)
            End While
            mysqlconn.Close()
            mysqlconn.Dispose()
            Return str
        Catch ex As Exception
            MsgBox(ex.Message + ",系統即將關閉")
            Application.Exit()
            Return ""
        End Try
    End Function

    Public Function telegram(ByVal info As String)
        Dim key As String = SQL_Return("SELECT * FROM `config` WHERE `name`='telegram_api_key'", "data")
        Dim request As WebRequest
        Dim response1 As WebResponse
        Try
            mysqlconn = New MySql.Data.MySqlClient.MySqlConnection
            mysqlconn.ConnectionString = SQL_Conn
            mysqlconn.Open()
            Dim Query As String

            Query = "SELECT * FROM `config` WHERE `name`='telegram_chat_id'"
            command = New MySql.Data.MySqlClient.MySqlCommand(Query, mysqlconn)
            reader = command.ExecuteReader
            Dim resp = ""
            While reader.Read
                resp = reader.GetString("data")
                request = WebRequest.Create("https://api.telegram.org/bot" + key + "/sendmessage?text=" + info + "&chat_id=" + resp)
                Console.WriteLine("https://api.telegram.org/bot" + key + "/sendmessage?text=" + info + "&chat_id=" + resp)
                'MsgBox("https://api.telegram.org/bot" + key + "/sendmessage?text=" + info + "&chat_id=" + resp)
                response1 = request.GetResponse()
                response1.Close()
                Threading.Thread.Sleep(500)

            End While
            mysqlconn.Close()
            mysqlconn.Dispose()
            Return 0
        Catch ex As Exception
            Return 0
        End Try



        Return 0
    End Function

End Class
