Imports System.Reflection
Imports System.IO
Imports MySql.Data.MySqlClient
Imports MetroFramework
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


    Public Function SQLTEST()
        Try
            mysqlconn = New MySql.Data.MySqlClient.MySqlConnection
            mysqlconn.ConnectionString = SQL_Conn
            mysqlconn.Open()

            mysqlconn.Close()
            mysqlconn.Dispose()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

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
            Return ""
        End Try
    End Function

    Public Function SQLListbox(ByVal Cmd As String, ByVal rt As String, ByVal LB As ListBox)
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
                LB.Items.Add(reader.GetString(rt))
            End While
            mysqlconn.Close()
            mysqlconn.Dispose()
            Return 0
        Catch ex As Exception
            Return ""
            LB.Items.Add(ex.Message)
        End Try
    End Function
    Public Function SQLComboBox(ByVal Cmd As String, ByVal rt As String, ByVal CB As ComboBox)
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
                CB.Items.Add(reader.GetString(rt))
            End While
            mysqlconn.Close()
            mysqlconn.Dispose()
            Return 0
        Catch ex As Exception
            Return ex.Message
            CB.Items.Add(ex.Message)
            Return ""
        End Try
    End Function


    Public Function SQLRichTextbox(ByVal Cmd As String, ByVal rt As String, ByVal RTB As RichTextBox, Optional ByVal PROCESSBAR As ProgressBar = Nothing)
        Dim mysqlconn As MySql.Data.MySqlClient.MySqlConnection
        Dim reader As MySql.Data.MySqlClient.MySqlDataReader
        Dim command As New MySql.Data.MySqlClient.MySqlCommand

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
                PROCESSBAR.Value = PROCESSBAR.Value + 1
                RTB.Text = RTB.Text + reader.GetString(rt) + vbCrLf

            End While
            mysqlconn.Close()
            mysqlconn.Dispose()
            Return 0
        Catch ex As Exception
            RTB.Text = ex.Message
            Return ex.Message


        End Try
    End Function

    Public Function SQLCOUNT(ByVal Cmd As String)

        mysqlconn = New MySql.Data.MySqlClient.MySqlConnection
        mysqlconn.ConnectionString = SQL_Conn
        mysqlconn.Open()
        Dim Query As String

        Query = Cmd
        command = New MySql.Data.MySqlClient.MySqlCommand(Query, mysqlconn)
        reader = command.ExecuteReader
        Dim int As Integer = 0

        While reader.Read
            int = int + 1
        End While
        mysqlconn.Close()
        mysqlconn.Dispose()
        Return int




    End Function

    Public Function SQLALLINTEXTBOX(ByVal Cmd As String, ByVal rt As String)

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
                str = str + reader.GetString(rt)
            End While
            mysqlconn.Close()
            mysqlconn.Dispose()
            Return str
        Catch ex As Exception
            Return ex.Message
        End Try

    End Function
    Public Function SQLBACKUP(ByVal Location As String)
        Try

            Using conn As New MySqlConnection(SQL_Conn)
                Using cmd As New MySqlCommand()
                    Using mb As New MySqlBackup(cmd)
                        cmd.Connection = conn
                        conn.Open()
                        mb.ExportToFile(Location + "/db.sql")
                        conn.Close()
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Return ex.Message.ToString

        Finally

        End Try
        Return "OK"
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

    Public Function DataGrid(ByVal CMD As String, ByVal DG As DataGridView)
        Dim myconn As New MySqlConnection(SQL_Conn)
        Dim cmd1 As New MySqlCommand
        Dim ada As New MySqlDataAdapter
        Dim table As DataTable

        Try
            myconn.Open()
            cmd1.Connection = myconn
            ada = New MySqlDataAdapter(CMD, myconn)
            table = New DataTable
            ada.Fill(table)
            DG.DataSource = table
            myconn.Close()
            myconn.Dispose()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return 0
    End Function
    Public Function SQLRestore(ByVal file As String)
        Try
            Using conn As New MySqlConnection(SQL_Conn)
                Using cmd As New MySqlCommand()
                    Using mb As New MySqlBackup(cmd)
                        cmd.Connection = conn
                        conn.Open()
                        mb.ImportFromFile(file)
                        conn.Close()
                    End Using
                End Using
            End Using
            Return 0
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function SQLInjection(ByVal Text As String)
        If Text.ToUpper.Contains("'") = True Or Text.ToUpper.Contains("--") = True Or Text.ToUpper.Contains("#") = True Or Text.ToUpper.Contains("=") = True Then
            Return False
        Else
            Return True
        End If

    End Function


End Class
