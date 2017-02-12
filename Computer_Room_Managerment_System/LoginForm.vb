'定義模組
Imports System.DirectoryServices
Imports MetroFramework

Public Class LoginForm
    '定義變數
    Dim int As Integer = 0
    Dim s As New SQL

    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click




        '查詢AD和SQL系統用戶名稱
        If (AuthenticateUser("LDAP://LYCMC2013.EDU.HK", MetroTextBox2.Text, MetroTextBox1.Text) = True And s.SQL_Return("SELECT * FROM `user` WHERE `user`='" + MetroTextBox2.Text + "'", "user") = MetroTextBox2.Text) Then
            '判定用戶是否被封鎖
            If (s.SQL_Return("SELECT * FROM `user` WHERE `user`='" + MetroTextBox2.Text + "';", "pwd_attempt") = "15") Then
                '顯示警告
                MetroMessageBox.Show(Me,
                      vbNewLine + "帳戶自動停權" _
                      + vbNewLine + "帳戶封鎖，請和系統管理員聯絡以取得協助。" + vbNewLine +
                      "本次錯誤系統已記錄至日誌！",
                      "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                '
                s.SQL_Return("UPDATE `user` SET `pwd_attempt`=0 WHERE `user`='" + MetroTextBox2.Text + "'", "")
                '傳送變數至MainForm
                '傳送AD使用者名(AD = DisplayName)
                MainForm.Displayname = chkuser("LDAP://LYCMC2013.EDU.HK", MetroTextBox2.Text, MetroTextBox1.Text, MetroTextBox2.Text)
                '傳送AD使用者名稱
                MainForm.Username = MetroTextBox2.Text
                '傳送AD密碼
                MainForm.Password = MetroTextBox1.Text
                '顯示MainForm
                MainForm.Show()
                '記錄資料
                s.telegram("資訊:" + vbNewLine + "帳戶" + MetroTextBox2.Text + "登入成功" + vbNewLine + "--系統管理員啟")
                log("管理員登入成功", "使用者:" + MetroTextBox2.Text + "登入成功", MetroTextBox2.Text)
                Me.Hide()
            End If
        Else
            '登入失敗
            '系統登入錯誤變數+1
            int = int + 1
            '定義SQL登入次數
            Dim attempt As String
            '判定用戶是否被封鎖
            If (s.SQL_Return("SELECT * FROM `user` WHERE `user`='" + MetroTextBox2.Text + "';", "pwd_attempt") = "15") Then
                '顯示警告
                MetroMessageBox.Show(Me,
                      vbNewLine + "帳戶自動停權" _
                      + vbNewLine + "帳戶封鎖，請和系統管理員聯絡以取得協助。" + vbNewLine +
                      "本次錯誤系統已記錄至日誌！",
                      "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                '記錄警告
                s.telegram("警告:" + vbNewLine + "帳戶" + MetroTextBox2.Text + "嘗試登入，但正在停權" + vbNewLine + "--系統管理員啟")
                log("管理員登入失敗", "使用者:" + MetroTextBox2.Text + "登入失敗，使用之電腦名稱:" + My.Computer.Name, MetroTextBox2.Text)
            Else
                'SQL系統登入錯誤變數+1
                s.SQL_Return("UPDATE `user` SET `pwd_attempt`= `pwd_attempt`+1 WHERE `user`='" + MetroTextBox2.Text + "'", "pwd_attempt")
            End If
            '儲存登入錯誤變數
            attempt = s.SQL_Return("SELECT * FROM `user` WHERE `user`='" + MetroTextBox2.Text + "';", "pwd_attempt")
            '如果SQL查無此人,則使用本機緩存
            If (attempt = "") Then
                attempt = int.ToString
            End If
            '判定是否錯誤15次
            If (attempt = "15") Then
                '顯示封鎖信息
                MetroMessageBox.Show(Me,
                      vbNewLine + "帳戶自動停權" _
                      + vbNewLine + "帳戶封鎖，請和系統管理員聯絡以取得協助。" + vbNewLine +
                      "本次錯誤系統已記錄至日誌！",
                      "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error)
                '記錄警告
                s.telegram("警告:" + vbNewLine + "帳戶" + MetroTextBox2.Text + "因被停權而登入失敗" + vbNewLine + "--系統管理員啟")
                log_warning("管理員登入失敗", "使用者:" + MetroTextBox2.Text + "因被停權而登入失敗", MetroTextBox2.Text)
            Else
                '顯示登入錯誤
                MetroMessageBox.Show(Me,
                                        vbNewLine + "使用者名稱或密碼錯誤。" _
                                        + vbNewLine + "請確定是使用登入電腦之登入資訊" + vbNewLine +
                                        "本次錯誤系統已記錄至日誌" + vbNewLine + "目前已錯誤次數 : " + attempt + vbNewLine + " 當錯誤次數達到15，帳戶便會自動停權！",
                                        "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                '記錄警告
                s.telegram("警告:" + vbNewLine + "帳戶" + MetroTextBox2.Text + "密碼錯誤" + vbNewLine + "--系統管理員啟")
                log_warning("管理員登入失敗", "使用者:" + MetroTextBox2.Text + "因被密碼錯誤而登入失敗", MetroTextBox2.Text)
            End If
        End If
    End Sub


    '函數AD登入
    Function AuthenticateUser(path As String, user As String, pass As String) As Boolean
        '定義變數並登入AD
        Dim de As New DirectoryEntry(path, user, pass, AuthenticationTypes.Secure)
        '在AD中尋找自己
        Try
            Dim ds As DirectorySearcher = New DirectorySearcher(de)
            ds.FindOne()
            '成功則回傳True
            Return True
        Catch
            '失敗則回傳False
            Return False
        End Try
    End Function

    '函數AD查詢
    Private Function chkuser(LDAPstr As String, LDAPuser As String, LDAPpass As String, UserName As String)
        '定義變數並登入AD
        Dim entry As DirectoryEntry = New DirectoryEntry(LDAPstr, LDAPuser, LDAPpass)
        '定義全名
        Dim FullName As String = ""
        '定義變數
        Dim obj As Object = entry.NativeObject
        Dim search As DirectorySearcher = New DirectorySearcher(entry)
        '在AD中尋找使用者
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
            '錯誤登入資料則回傳使用者名稱或密碼錯誤
            Return "使用者名稱或密碼錯誤"
        Catch ex As Exception
            'Bug則回傳錯誤資料
            Return ex.Message
        End Try
        '成功則回傳全名
        Return FullName
    End Function

    Public Function log(ByVal type As String, ByVal info As String, ByVal ppl As String)
        s.SQL_Return("INSERT INTO `logs` (`uid`, `time`, `level`, `type`, `details`, `operator`, `affected_people`) VALUES (NULL, CURRENT_TIMESTAMP, '資訊', '" + type + "', '[資訊][" + Now().ToString + "]" + info + "', '" + MetroTextBox2.Text + "', '" + ppl + "');", "details")
        Return 0
    End Function

    Public Function log_warning(ByVal type As String, ByVal info As String, ByVal ppl As String)
        s.SQL_Return("INSERT INTO `logs` (`uid`, `time`, `level`, `type`, `details`, `operator`, `affected_people`) VALUES (NULL, CURRENT_TIMESTAMP, '警告', '" + type + "', '[警告][" + Now().ToString + "]" + info + "', '" + MetroTextBox2.Text + "', '" + ppl + "');", "details")
        Return 0
    End Function

    Private Sub LoginForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class