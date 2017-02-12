Imports System.Drawing.Text
Imports MetroFramework
Imports MySql.Data.MySqlClient
Imports System.DirectoryServices
Imports System.Text
Imports System.Globalization

Public Class MainForm
    Public Username As String
    Public Password As String
    Public Displayname As String
    Dim s As New SQL
    Private Sub TabControl1_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles TabControl1.DrawItem
        Dim g As Graphics
        Dim sText As String
        Dim iX As Integer
        Dim iY As Integer
        Dim sizeText As SizeF
        Dim ctlTab As TabControl

        ctlTab = CType(sender, TabControl)

        g = e.Graphics

        sText = ctlTab.TabPages(e.Index).Text
        sizeText = g.MeasureString(sText, ctlTab.Font)
        iX = e.Bounds.Left + 6
        iY = e.Bounds.Top + (e.Bounds.Height - sizeText.Height) / 2
        g.DrawString(sText, ctlTab.Font, Brushes.Black, iX, iY)


    End Sub

    Private Sub MainForm_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        s.SQL_Return("UPDATE `user` SET `last_login`=now() , `logined` = 'false' WHERE `user`='" + Username + "';", "last_login")
        log("管理員登出", "管理員: " + Username + " 登出了系統", Username)
        Application.Exit()
    End Sub



    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TabControl1.ShowTab(HomePage_Tab)
        s.DataGrid("SELECT * FROM `seat`", DataGridView1)
        s.DataGrid("SELECT * FROM `banned_user`", DataGridView2)
        s.SQLComboBox("SELECT * FROM `config` WHERE `name`='ban_reason'", "data", MetroComboBox2)
        s.SQLComboBox("SELECT * FROM `config` WHERE `name`='ban_reason'", "data", MetroComboBox3)
        s.SQL_Return("UPDATE `user` SET `logined` = 'true' WHERE `user`='" + Username + "';", "last_login")
        s.SQLComboBox("SELECT * FROM `seat` WHERE 1", "computer_name", MetroComboBox1)
        From_Textbox.Text = Today.Year & "/" & Today.Month & "/" & Today.Day
        To_Textbox.Text = Today.Date.AddDays(1)

        If s.SQLTEST() = False Then
            MetroMessageBox.Show(Me, "連接失敗" _
                            + vbNewLine + "目前無法和MYSQL連接，系統現在無法使用。" _
                            + vbNewLine + "請和系統管理員取得協助",
                            "MYSQL 系統故障", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        refresh_homepage()
        Username_Label.Text = Displayname + "歡迎回來"
        LastLogin_Label.Text = "你上一次登入系統時間為 : " + s.SQL_Return("SELECT * FROM `user` WHERE `user`='" + Username + "';", "last_login")
        Username_2_Label.Text = "使用者名稱 : " + Displayname
        Group_Label.Text = "用戶組 : " + s.SQL_Return("SELECT * FROM `user` WHERE `user`='" + Username + "';", "level")
        LastLogin_1_Label.Text = "上一次登入 : " + s.SQL_Return("SELECT * FROM `user` WHERE `user`='" + Username + "';", "last_login")

        'Init SQL Data into log box
        If showalllog_toggle.Checked = True Then
            Logrefresh("all")
        Else
            Logrefresh("today")
        End If

        If s.SQL_Return("SELECT * FROM `user` WHERE `user`='" + Username + "'", "level") = "admin" Then
            EnableSQL_Btn.Enabled = True
        End If
    End Sub

    Private Sub OnOffSystem_Btn_Click(sender As Object, e As EventArgs) Handles OnOffSystem_Btn.Click
        If s.SQL_Return("SELECT * FROM `config` WHERE `name`='system_enable'", "data") = "false" Then
            s.SQL_Return("UPDATE `config` SET `data`='true' WHERE `name`='system_enable'", "data")
            log("啟動系統", "啟動系統成功", "GLOBAL")
        Else
            s.SQL_Return("UPDATE `config` SET `data`='false' WHERE `name`='system_enable'", "data")
            log("關閉系統", "關閉系統成功", "GLOBAL")
        End If
    End Sub



    Private Sub RefreshTile_Btn_Click(sender As Object, e As EventArgs) Handles RefreshTile_Btn.Click
        refresh_homepage()
    End Sub


    Private Sub ColorText(ByRef RTBToSearch As RichTextBox, ByVal TextToColor As String, ByVal Color As System.Drawing.Color)

        Dim StrPos As Integer
        Dim StrLen As Integer
        Dim FindStr As String

        FindStr = TextToColor

        StrLen = Len(FindStr)

        StrPos = RTBToSearch.Find(FindStr)

        Do While StrPos >= 0
            RTBToSearch.SelectionStart = StrPos
            RTBToSearch.SelectionLength = StrLen
            RTBToSearch.SelectionColor = Color
            StrPos = RTBToSearch.Find(FindStr, StrPos + StrLen, RichTextBoxFinds.None)
        Loop

    End Sub
    Private Sub Logout_Btn_Click(sender As Object, e As EventArgs) Handles Logout_Btn.Click
        Me.Close()

    End Sub
    Public Function log(ByVal type As String, ByVal info As String, ByVal ppl As String)
        s.SQL_Return("INSERT INTO `logs` (`uid`, `time`, `level`, `type`, `details`, `operator`, `affected_people`) VALUES (NULL, CURRENT_TIMESTAMP, '資訊', '" + type + "', '[資訊][" + Now().ToString + "]" + info + "', '" + Username + "', '" + ppl + "');", "details")
        Return 0
    End Function

    Public Function log_warning(ByVal type As String, ByVal info As String, ByVal ppl As String)
        s.SQL_Return("INSERT INTO `logs` (`uid`, `time`, `level`, `type`, `details`, `operator`, `affected_people`) VALUES (NULL, CURRENT_TIMESTAMP, '警告', '" + type + "', '[警告][" + Now().ToString + "]" + info + "', '" + Username + "', '" + ppl + "');", "details")
        Return 0
    End Function


    Public Function refresh_homepage()
        Dim initsqlerr As Boolean = False
        'init test 
        If s.SQLTEST() = False Then
            SystemInfo_Tile.Style = MetroFramework.MetroColorStyle.Red
            SystemInfo_Tile.Text = "目前系統狀態：故障"
            initsqlerr = True


        Else
            SystemInfo_Tile.Style = MetroFramework.MetroColorStyle.Green
            SystemInfo_Tile.Text = "目前系統狀態：正常"
        End If
        'end

        'init version tile
        Dim CKNew As New CheckNew("http://api.alanyeung.co/checknew_crms.txt", "crms") '建立CheckNew類別的CKNew物件
        CKNew.CheckNew()
        Select Case CKNew.GetCheckConsequence '取得更新結果
            Case 0 '沒有更新
                ProgramVersion_Tile.Style = MetroFramework.MetroColorStyle.Green
                ProgramVersion_Tile.Text = "程式版本 : 最新"
            Case 1 '有更新
                ProgramVersion_Tile.Style = MetroFramework.MetroColorStyle.Yellow
                ProgramVersion_Tile.Text = "程式版本 : 有更新"
            Case 2 '更新失敗
                ProgramVersion_Tile.Style = MetroFramework.MetroColorStyle.Red
                ProgramVersion_Tile.Text = "程式版本 : 無法取得"
        End Select
        'end

        'Init controller info
        If initsqlerr = False Then
            If s.SQL_Return("SELECT * FROM `config` WHERE `name`='system_enable'", "data") = "false" Then
                ControllerInfo_Tile.Style = MetroFramework.MetroColorStyle.Yellow
                ControllerInfo_Tile.Text = "目前電腦控制系統狀態：未啟動"
            Else
                ControllerInfo_Tile.Style = MetroFramework.MetroColorStyle.Green
                ControllerInfo_Tile.Text = "目前電腦控制系統狀態：啟動"
            End If

        End If

        'end

        'Init attend ppl info


        Dim date_d As Integer = Now.DayOfWeek()
        If initsqlerr = False Then
            If (s.SQLCOUNT("SELECT * FROM `user` WHERE `attend` LIKE '%" + date_d.ToString + "%'") - s.SQLCOUNT("SELECT * FROM `user` WHERE `attend` LIKE '%" + date_d.ToString + "%' AND `logined` LIKE 'true'").ToString) = 0 Then
                RequireAttendInt_Tile.Text = "本日當值中人數: " + s.SQLCOUNT("SELECT * FROM `user` WHERE `attend` LIKE '%" + date_d.ToString + "%' AND `logined` LIKE 'true'").ToString
                RequireAttendInt_Tile.Style = MetroFramework.MetroColorStyle.Green
            Else
                RequireAttendInt_Tile.Text = MetroFramework.MetroColorStyle.Red
                RequireAttendInt_Tile.Text = "本日當值中人數: " + s.SQLCOUNT("SELECT * FROM `user` WHERE `attend` LIKE '%" + date_d.ToString + "%' AND `logined` LIKE 'true'").ToString
            End If
        End If
        'end 

        'init borrow ppl info
        If initsqlerr = False Then
            UsingComputer_Tile.Text = "目前使用中電腦 : " + s.SQLCOUNT("SELECT * FROM `seat` WHERE `borrowed`='1'").ToString
            UsingComputer_Tile.Style = MetroFramework.MetroColorStyle.Green
        End If
        'end

        'init attend ppl name
        If initsqlerr = False Then
            If s.SQLCOUNT("SELECT * FROM `user` WHERE `attend` LIKE '%" + date_d.ToString + "%' AND `logined` LIKE 'false'") > 0 Then
                RequireAttendText_Tile.Text = "本日當值但未到人士 : " + s.SQLALLINTEXTBOX("SELECT * FROM `user` WHERE `attend` LIKE '%" + date_d.ToString + "%' AND `logined` LIKE 'false'", "nickname").ToString
                RequireAttendText_Tile.Style = MetroFramework.MetroColorStyle.Red
            Else
                RequireAttendText_Tile.Text = "所有人已在當值 " + s.SQLALLINTEXTBOX("SELECT * FROM `user` WHERE `attend` LIKE '%" + date_d.ToString + "%' AND `logined` LIKE 'false'", "nickname").ToString
                RequireAttendText_Tile.Style = MetroFramework.MetroColorStyle.Green
            End If

        End If
        'end




        'end
        Return 0
    End Function

    Public Function Logrefresh(ByVal log As String)
        Dim dt As Date = Date.Today
        Dim a As String = dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        If log = "today" Then
            Log_ProgressBar.Value = 0
            Log_ProgressBar.Maximum = s.SQL_Return("SELECT COUNT(*) FROM `logs` WHERE `time` BETWEEN '" + a + " 00:00:00.000000' AND '" + a + " 23:59:59.999999'", "COUNT(*)")
            Log_Textbox.Visible = False
            Log_Textbox.Text = ""
            s.SQLRichTextbox("SELECT * FROM `logs` WHERE `time` BETWEEN '" + a + " 00:00:00.000000' AND '" + a + " 23:59:59.999999'", "details", Log_Textbox, Log_ProgressBar)
            ColorText(Log_Textbox, "[警告]", Color.Red)
            ColorText(Log_Textbox, "[資訊]", Color.Blue)
            Log_Textbox.ScrollToCaret()

            Log_Textbox.Visible = True
        End If
        If log = "all" Then
            Log_ProgressBar.Value = 0
            Log_ProgressBar.Maximum = s.SQL_Return("SELECT COUNT(*) FROM `logs`", "COUNT(*)")
            Log_Textbox.Visible = False
            Log_Textbox.Text = ""
            s.SQLRichTextbox("SELECT * FROM `logs`", "details", Log_Textbox, Log_ProgressBar)
            ColorText(Log_Textbox, "[警告]", Color.Red)
            ColorText(Log_Textbox, "[資訊]", Color.Blue)
            Log_Textbox.ScrollToCaret()

            Log_Textbox.Visible = True
        End If

        If log = "warning" Then
            Log_ProgressBar.Value = 0
            Log_ProgressBar.Maximum = s.SQL_Return("SELECT COUNT(*) FROM `logs`  WHERE `level` = '警告'", "COUNT(*)")
            Log_Textbox.Visible = False
            Log_Textbox.Text = ""
            s.SQLRichTextbox("SELECT * FROM `logs` WHERE `level` = '警告'", "details", Log_Textbox, Log_ProgressBar)
            ColorText(Log_Textbox, "[警告]", Color.Red)
            ColorText(Log_Textbox, "[資訊]", Color.Blue)
            Log_Textbox.ScrollToCaret()
            Log_Textbox.Visible = True
        End If
        Return 0
    End Function


    Private Sub LogRefresh_Btn_Click(sender As Object, e As EventArgs) Handles LogRefresh_Btn.Click
        If showalllog_toggle.Checked = True Then
            Logrefresh("all")
        Else
            Logrefresh("today")
        End If
    End Sub


    Private Sub MetroToggle1_CheckedChanged(sender As Object, e As EventArgs) Handles showalllog_toggle.CheckedChanged
        If showalllog_toggle.Checked = True Then
            MetroMessageBox.Show(Me, "系統目前偵測到大量SQL正準備執行" _
                        + vbNewLine + "在執行時，請減少電腦執行中程式!",
                        "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Logrefresh("all")
        Else
            Logrefresh("today")
        End If
    End Sub


    Private Sub ShowWarningLogOnly_Btn_Click(sender As Object, e As EventArgs) Handles ShowWarningLogOnly_Btn.Click
        Logrefresh("warning")
    End Sub

    Private Sub Output_Btn_Click(sender As Object, e As EventArgs) Handles Output_Btn.Click
        BackupSQLBrowser.ShowDialog()
        Dim result As String
        result = s.SQLBACKUP(BackupSQLBrowser.SelectedPath)
        If result = "OK" Then
            MetroMessageBox.Show(Me, "完成備份SQL DB" _
                    ,
                     "完成", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
        Else
            MetroMessageBox.Show(Me, "系統目前無法備份SQL DB" _
                    + vbNewLine + "請稍後再試" + vbNewLine + "SQL錯誤資料:" + result,
                    "失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        End If


    End Sub

    Private Sub ResetSeat_Btn_Click(sender As Object, e As EventArgs) Handles ResetSeat_Btn.Click
        s.SQL_Return("UPDATE `seat` SET `borrowed_by`='',`borrowed`='0',`logout_time`= '2000-01-01 00:00:00' WHERE 1", "computer_name")
        s.DataGrid("SELECT * FROM `seat`", DataGridView1)
    End Sub

    Private Sub Refresh_Timer_Tick(sender As Object, e As EventArgs) Handles Refresh_Timer.Tick
        refresh_homepage()
    End Sub

    Private Sub RefreshSeat_Btn_Click(sender As Object, e As EventArgs) Handles RefreshSeat_Btn.Click
        s.DataGrid("SELECT * FROM `seat`", DataGridView1)
    End Sub

    Private Sub QueryUsername_Btn_Click(sender As Object, e As EventArgs) Handles QueryUsername_Btn.Click
        QueryNameResult_Label.Text = "學生名稱 : " + chkuser("LDAP://LYCMC2013.EDU.HK", Username, Password, QueryUsername_Textbox.Text)
        Try
            Query_Picturebox.Load("http://eclass.lycmc.edu.hk/file/user_photo/" + QueryUsername_Textbox.Text + ".jpg")
        Catch ex As Exception
            'Query_Picturebox.Image 
        End Try
        MetroLabel3.Text = "警告次數 : " + s.SQL_Return("SELECT * FROM `warning_page` WHERE `affected_people` ='" + QueryUsername_Textbox.Text + "' AND `user_type` = 'normal';", "total_warning")
        Ban_Btn.Text = "停權此用戶(" + chkuser("LDAP://LYCMC2013.EDU.HK", Username, Password, QueryUsername_Textbox.Text) + ")"
    End Sub
    Private Function chkuser(LDAPstr As String, LDAPuser As String, LDAPpass As String, UserName As String)
        Dim entry As DirectoryEntry = New DirectoryEntry(LDAPstr, LDAPuser, LDAPpass)
        Dim FullName As String = ""

        Dim obj As Object = entry.NativeObject
        Dim search As DirectorySearcher = New DirectorySearcher(entry)
        Try
            search.Filter = "(&(objectClass=user)(SAMAccountName=" & UserName & "))"
            search.PropertiesToLoad.Add("displayName")

            Dim result As SearchResult = search.FindOne()
            If (result Is Nothing) Then
                FullName = "未找到用戶"
            Else
                FullName = result.Properties("displayName")(0).ToString()
            End If

        Catch ex As System.DirectoryServices.DirectoryServicesCOMException

            Return "使用者名稱或密碼錯誤"
        Catch ex As Exception
            Return ex.Message

        End Try
        Return FullName
    End Function

    Private Sub Submit_Btn_Click(sender As Object, e As EventArgs) Handles Submit_Btn.Click
        LoginRecord_Listbox.Items.Clear()
        s.SQLListbox("SELECT * FROM `logs` WHERE `time` BETWEEN '" + From_Textbox.Text + " 00:00:00.000000' AND '" + To_Textbox.Text + " 00:00:00.000000' AND `type` = '管理員登入成功' AND `operator` = '" + Username + "' AND `affected_people` = '" + Username + "';", "time", LoginRecord_Listbox)
        LoginRecord_Listbox.Items.Add("========記錄完========")
        LoginRecord_Listbox.Items.Add("由" + From_Textbox.Text + "至" + To_Textbox.Text + "共" + s.SQL_Return("SELECT COUNT(*) FROM `logs` WHERE `time` BETWEEN '" + From_Textbox.Text + " 00:00:00.000000' AND '" + To_Textbox.Text + " 23:59:59.000000' AND `type` = '管理員登入成功' AND `operator` = '" + Username + "' AND `affected_people` = '" + Username + "';", "COUNT(*)") + "筆記錄")
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Process.Start("https://telegram.me/Ivybridge_i3")
    End Sub

    Private Sub EnableSQL_Btn_Click(sender As Object, e As EventArgs) Handles EnableSQL_Btn.Click
        Dim s As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopgrstuvwxyz0123456789"
        Dim r As New Random
        Dim sb As New StringBuilder
        For i As Integer = 1 To 5
            Dim idx As Integer = r.Next(0, 35)
            sb.Append(s.Substring(idx, 1))
        Next
        MetroLabel2.Text = "驗證碼: " + sb.ToString()

        EnableSQL_Btn.Enabled = False
    End Sub

    Private Sub MetroButton2_Click(sender As Object, e As EventArgs) Handles MetroButton2.Click
        Dim api As String
        api = s.SQL_Return("SELECT * FROM `config` WHERE `name`='telegram_api_key'", "data")
        If (api = "") Then
            MetroMessageBox.Show(Me, "API Key問題" _
                       + vbNewLine + "API Key錯誤，Telegram拒絕閣下要求。" _
                       + vbNewLine + "請重新輸入",
                       "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim objWebClient As New System.Net.WebClient
            Dim Url As String = "https://api.telegram.org/bot" + api + "/getupdates"
            Dim JsonStr As String = ""
            JsonStr = Encoding.UTF8.GetString(objWebClient.DownloadData(New Uri(Url.Trim())))

            Dim Obj As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.JsonConvert.DeserializeObject(JsonStr)
            '取得物件明細的方法 依據階層 Obj.item("階層1")("階層2") 以此類推
            Dim int As Integer = 0
            RichTextBox1.Text = ""
            RichTextBox1.Text = "====以下是Telegram UID資訊====" + Environment.NewLine
            For Each Items In Obj.Item("result")

                Dim ChatId As String = Obj.Item("result")(int)("message")("chat")("id").ToString
                Dim ChatUID As String = Obj.Item("result")(int)("message")("message_id").ToString
                Dim ChatText As String = Obj.Item("result")(int)("message")("text").ToString
                int = int + 1
                RichTextBox1.Text += "ChatID - " + ChatId + " , UID - " + ChatUID + " , Text - " + ChatText + Environment.NewLine
            Next
            int = 0
            RichTextBox1.Text += "====以下是儲存在SQL 之Telegram UID資訊====" + Environment.NewLine
            RichTextBox1.Text += s.SQLALLINTEXTBOX("SELECT * FROM `config` WHERE `name`='telegram_chat_id'", "data")
        End If


    End Sub

    Private Sub MetroComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MetroComboBox1.SelectedIndexChanged
        If s.SQL_Return("SELECT * FROM `seat` WHERE `computer_name` = '" + MetroComboBox1.SelectedItem + "';", "enabled") = "0" Then
            MetroToggle2.Checked = False
        Else
            MetroToggle2.Checked = True
        End If
        MetroTextBox4.Text = s.SQL_Return("SELECT * FROM `seat` WHERE `computer_name` = '" + MetroComboBox1.SelectedItem + "';", "borrowed_by")
    End Sub

    Private Sub MetroToggle2_CheckedChanged(sender As Object, e As EventArgs) Handles MetroToggle2.CheckedChanged
        If MetroToggle2.Checked = True Then
            s.SQL_Return("UPDATE `seat` SET `enabled`='1' WHERE `computer_name`='" + MetroComboBox1.SelectedItem + "';", "enabled")
            log("座位設定", "座位: " + MetroComboBox1.SelectedItem + " 已設定為啟用", Username)
        Else
            s.SQL_Return("UPDATE `seat` SET `enabled`='0' WHERE `computer_name`='" + MetroComboBox1.SelectedItem + "';", "enabled")
            log("座位設定", "座位: " + MetroComboBox1.SelectedItem + " 已設定為停用", Username)
        End If
        s.DataGrid("SELECT * FROM `seat`", DataGridView1)
    End Sub

    Private Sub MetroButton6_Click(sender As Object, e As EventArgs) Handles MetroButton6.Click
        s.SQL_Return("UPDATE `seat` SET `borrowed_by`='" + MetroTextBox4.Text + "' WHERE `computer_name`='" + MetroComboBox1.SelectedItem + "';", "enabled")
        log("座位設定", "座位: " + MetroComboBox1.SelectedItem + " 已設定使用者為 : " + MetroTextBox4.Text, Username)
        s.DataGrid("SELECT * FROM `seat`", DataGridView1)
    End Sub


    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        If ("驗證碼: " + MetroTextBox1.Text = MetroLabel2.Text) Then
            SQL_Groupbox.Enabled = True
            GroupBox3.Enabled = True
        Else
            MetroMessageBox.Show(Me, "驗證碼問題" _
                         + vbNewLine + "驗證碼錯誤，系統現在拒絕閣下要求。" _
                         + vbNewLine + "請重新輸入",
                         "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        MetroButton1.Enabled = False
    End Sub

    Private Sub ClearLog_Btn_Click(sender As Object, e As EventArgs) Handles ClearLog_Btn.Click
        If (MetroMessageBox.Show(Me, "提示" _
                          + vbNewLine + "您正要清空(TRUNCATE)整張資料表! 您確定要執行?" _
                          + vbNewLine + "SQL指令: TRUNCATE logs",
                          "", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) = Windows.Forms.DialogResult.Yes) Then
            s.SQL_Return("TRUNCATE logs", "")
            MetroMessageBox.Show(Me, "提示" _
          + vbNewLine + "系統成功執行要求。" _
          + vbNewLine + "",
          "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MetroMessageBox.Show(Me, "提示" _
             + vbNewLine + "系統成功取消要求。" _
             + vbNewLine + "",
             "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If


    End Sub

    Private Sub MetroButton5_Click(sender As Object, e As EventArgs) Handles MetroButton5.Click
        If (MetroMessageBox.Show(Me, "提示" _
                       + vbNewLine + "您正要清空(TRUNCATE)整個資料庫! 您確定要執行?" _
                       + vbNewLine + "SQL指令: (多重指令)",
                       "", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) = Windows.Forms.DialogResult.Yes) Then
            'Clear All Data First
            s.SQL_Return("TRUNCATE logs", "")
            s.SQL_Return("TRUNCATE banned_user", "")
            s.SQL_Return("TRUNCATE config", "")
            s.SQL_Return("TRUNCATE seat", "")
            s.SQL_Return("TRUNCATE user", "")
            s.SQL_Return("TRUNCATE warning_page", "")

            'Insert Back Config
            s.SQL_Return("INSERT INTO `config` (`name`, `data`) VALUES('system_enable', 'false'),('telegram_api_key', ''),('telegram_chat_id', ''),('telegram_chat_id', '');", "")
            s.SQL_Return("INSERT INTO `user` (`user`, `nickname`, `level`, `enabled`, `pwd_attempt`, `attend`, `logined`, `last_login`) VALUES('administrator', '', 'admin', 1, 0, '1,2,3,4,5,6,7,8,9,0', 'false', '2016-08-17 13:10:12')", "")

            MetroMessageBox.Show(Me, "提示" _
          + vbNewLine + "系統成功執行要求。" _
          + vbNewLine + "請重新新增使用者，預設使用者為: Administrator",
          "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MetroMessageBox.Show(Me, "提示" _
             + vbNewLine + "系統成功取消要求。" _
             + vbNewLine + "",
             "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub Browse_Btn_Click(sender As Object, e As EventArgs) Handles Browse_Btn.Click
        RestoreSQLBrowser.ShowDialog()
        LogPath_Textbox.Text = RestoreSQLBrowser.FileName
    End Sub

    Private Sub Restore_Btn_Click(sender As Object, e As EventArgs) Handles Restore_Btn.Click
        If (MetroMessageBox.Show(Me, "提示" _
                 + vbNewLine + "您正要匯入(IMPORT)整個資料庫! 您確定要執行？ 您確定要執行?" _
                 + vbNewLine + "SQL指令: (多重指令)",
                 "", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk) = Windows.Forms.DialogResult.Yes) Then
            MetroMessageBox.Show(Me, "提示" _
            + vbNewLine + "開始匯入資料表！" _
            + vbNewLine + "",
            "", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Dim pass As String
            pass = s.SQLRestore(LogPath_Textbox.Text).ToString
            If (pass = "0") Then
                MetroMessageBox.Show(Me, "提示" _
                        + vbNewLine + "成功匯入資料表！" _
                        + vbNewLine + "",
                        "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MetroMessageBox.Show(Me, "錯誤" _
           + vbNewLine + "系統無法完成要求。" _
           + vbNewLine + "SQL錯誤 : " + pass,
           "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If


        Else
            MetroMessageBox.Show(Me, "提示" _
             + vbNewLine + "系統成功取消要求。" _
             + vbNewLine + "",
             "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        s.SQL_Return("UPDATE `config` SET `data`='" + MetroTextBox3.Text + "' WHERE `name`='telegram_api_key'", "")
        MetroMessageBox.Show(Me, "提示" _
         + vbNewLine + "設定完成！" _
         + vbNewLine + "API KEY :" + MetroTextBox3.Text,
         "", MessageBoxButtons.OK, MessageBoxIcon.Information)
        log("設定", "Telegram API KEY已設定", Username)
    End Sub

    Private Sub MetroButton4_Click(sender As Object, e As EventArgs) Handles MetroButton4.Click
        If (s.SQL_Return("SELECT * FROM `config` WHERE `name`='telegram_chat_id' AND data = '" + MetroTextBox2.Text + "';", "data") = MetroTextBox2.Text) Then

            s.SQL_Return("DELETE FROM `config` WHERE  `name`='telegram_chat_id' AND data = '" + MetroTextBox2.Text + "';", "data")
            MetroMessageBox.Show(Me, "提示" _
     + vbNewLine + "已刪除使用者！" _
     + vbNewLine + "",
     "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            log("設定", "Telegram 已刪除使用者", Username)
        Else
            s.SQL_Return("INSERT INTO `config`(`name`, `data`) VALUES ('telegram_chat_id','" + MetroTextBox2.Text + "')", "")
            MetroMessageBox.Show(Me, "提示" _
    + vbNewLine + "已增加使用者！" _
    + vbNewLine + "",
    "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            log("設定", "Telegram 已增加使用者", Username)
        End If

    End Sub

    Private Sub Ban_Btn_Click(sender As Object, e As EventArgs) Handles Ban_Btn.Click
        If (s.SQL_Return("SELECT * FROM `banned_user` WHERE `user_id` ='" + QueryUsername_Textbox.Text + "';", "user_id") = QueryUsername_Textbox.Text) Then
            MetroMessageBox.Show(Me, "提示" _
                     + vbNewLine + QueryUsername_Textbox.Text + "於早前已被停權！" _
                     + vbNewLine + "",
                     "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            s.SQL_Return("INSERT INTO `banned_user` (`uid`, `user_id`, `reason`, `operator`, `banned_until`) VALUES (NULL, '" + QueryUsername_Textbox.Text + "', '" + MetroComboBox2.SelectedItem.ToString + "', '" + Username + "', DATE_ADD(now(),INTERVAL 2 WEEK));", "")
            log("設定", QueryUsername_Textbox.Text + "被停權! 原因:" + MetroComboBox2.SelectedItem.ToString, Username)
            MetroMessageBox.Show(Me, "提示" _
                       + vbNewLine + MetroTextBox5.Text + "停權成功！" _
                       + vbNewLine + "",
                       "", MessageBoxButtons.OK, MessageBoxIcon.Information)

        End If

    End Sub
    Private Sub MetroButton7_Click(sender As Object, e As EventArgs) Handles MetroButton7.Click
        If (s.SQL_Return("SELECT * FROM `banned_user` WHERE user_id='" + MetroTextBox5.Text + "'", "user_id") = MetroTextBox5.Text) Then
            s.SQL_Return("DELETE FROM `banned_user` WHERE `user_id`='" + MetroTextBox5.Text + "'", "")
            MetroMessageBox.Show(Me, "提示" _
                      + vbNewLine + MetroTextBox5.Text + "取消停權！" _
                      + vbNewLine + "",
                      "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            log("使用者取消停權", MetroTextBox5.Text + "取消停權!", Username)
        Else
            s.SQL_Return("INSERT INTO `banned_user` (`uid`, `user_id`, `reason`, `operator`, `banned_until`) VALUES (NULL, '" + MetroTextBox5.Text + "', '" + MetroComboBox3.SelectedItem.ToString + "', '" + Username + "', DATE_ADD(now(),INTERVAL 2 WEEK));", "")
            MetroMessageBox.Show(Me, "提示" _
                         + vbNewLine + MetroTextBox5.Text + "停權成功！" _
                         + vbNewLine + "",
                         "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            log("使用者停權", MetroTextBox5.Text + "被停權! 原因:" + MetroComboBox3.SelectedItem.ToString, Username)
        End If
    End Sub

    Private Sub MetroButton22_Click(sender As Object, e As EventArgs) Handles MetroButton22.Click
        s.DataGrid("SELECT * FROM `banned_user`", DataGridView2)
    End Sub



    Private Sub MetroButton8_Click(sender As Object, e As EventArgs) Handles MetroButton8.Click

    End Sub

    Private Sub RequireAttendText_Tile_Click(sender As Object, e As EventArgs) Handles RequireAttendText_Tile.Click
        Dim date_d As Integer = Now.DayOfWeek()
        MetroMessageBox.Show(Me, "提示" _
                         + vbNewLine + s.SQLALLINTEXTBOX("SELECT * FROM `user` WHERE `attend` LIKE '%" + date_d.ToString + "%' AND `logined` LIKE 'false'", "nickname").ToString() _
                         + vbNewLine + "",
                         "", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub 請勿在本格輸入文字_Click(sender As Object, e As EventArgs) Handles 請勿在本格輸入文字.Click

    End Sub
End Class
