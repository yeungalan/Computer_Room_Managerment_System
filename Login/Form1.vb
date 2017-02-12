Public Class Form1
    Dim s As New SQL
    Dim user As String
    Dim second As Integer = 0
    Dim min As Integer = 1
    Dim accept_close As Boolean = False
    Dim computer As String = My.Computer.Name.Substring(0, My.Computer.Name.Length - 1)
    Public Function ex_accept()
        accept_close = True
        Return 0
    End Function
    Public Function log(ByVal type As String, ByVal info As String, ByVal ppl As String)
        Dim UserName As String = My.User.Name
        Dim intBackSlash As Integer = UserName.LastIndexOf("\")
        Dim DisplayName As String = UserName.Substring(intBackSlash + 1)
        s.SQL_Return("INSERT INTO `logs` (`uid`, `time`, `level`, `type`, `details`, `operator`, `affected_people`) VALUES (NULL, CURRENT_TIMESTAMP, '資訊', '" + type + "', '[資訊][" + Now().ToString + "]" + info + "', '" + DisplayName + "', '" + ppl + "');", "details")
        Return 0
    End Function

    Public Function log_warning(ByVal type As String, ByVal info As String, ByVal ppl As String)
        Dim UserName As String = My.User.Name
        Dim intBackSlash As Integer = UserName.LastIndexOf("\")
        Dim DisplayName As String = UserName.Substring(intBackSlash + 1)
        s.SQL_Return("INSERT INTO `logs` (`uid`, `time`, `level`, `type`, `details`, `operator`, `affected_people`) VALUES (NULL, CURRENT_TIMESTAMP, '警告', '" + type + "', '[警告][" + Now().ToString + "]" + info + "', '" + DisplayName + "', '" + ppl + "');", "details")
        Return 0
    End Function
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If accept_close = True Or s.expection = True Then
            e.Cancel = False
        Else
            e.Cancel = True
        End If

    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim UserName As String = My.User.Name
        Dim intBackSlash As Integer = UserName.LastIndexOf("\")
        Dim DisplayName As String = UserName.Substring(intBackSlash + 1)
        Me.Visible = False
        If (s.SQL_Return("SELECT * FROM `config` WHERE name = 'system_enable'", "data") = "false") Then
            log("電腦登入成功", "系統偵測到功能停用，" + computer + "將不檢查使用者名稱", DisplayName)
            accept_close = True
            Application.Exit()
        Else
            Me.Visible = True


            user = s.SQL_Return("SELECT * FROM `seat` WHERE `computer_name`='" + computer + "';", "borrowed_by")
            If (user = DisplayName) Then
                log("電腦登入成功", "" + computer + "登入成功", DisplayName)
                s.telegram("資訊:" + computer + "電腦登入成功" + vbNewLine + "--系統管理員啟")
                accept_close = True
                NotifyIcon1.ShowBalloonTip(2000)
                Application.Exit()
            Else
                log_warning("電腦登入失敗", "" + computer + "登入失敗", DisplayName)
                s.telegram("資訊:" + computer + "電腦登入失敗" + vbNewLine + "--系統管理員啟")
                Timer1.Enabled = True
            End If
        End If
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Visible = True
        TextBox2.Visible = True
        TextBox3.Visible = True
        TextBox4.Visible = True
        TextBox5.Visible = True
        TextBox6.Visible = True
        Button2.Visible = True
        TextBox1.Focus()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If (second = 0) Then
            min = min - 1
            second = 59
        End If
        second = second - 1
        Label2.Text = "0 : " + min.ToString + " : " + second.ToString


        Dim Proc As New System.Diagnostics.Process
        Proc.StartInfo = New ProcessStartInfo("C:\Windows\System32\cmd.exe")
        Proc.StartInfo.FileName = "shutdown.exe"
        Proc.StartInfo.Arguments = "-s -t " & second
        Proc.StartInfo.UseShellExecute = False
        Proc.StartInfo.CreateNoWindow = True
        Proc.Start()
        ' Allows script to execute sequentially instead of simultaneously
        Proc.WaitForExit()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        TextBox2.Focus()

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        TextBox3.Focus()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        TextBox4.Focus()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        TextBox5.Focus()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        TextBox6.Focus()
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Button2.Focus()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim pwd As String
        Dim UserName As String = My.User.Name
        Dim intBackSlash As Integer = UserName.LastIndexOf("\")
        Dim DisplayName As String = UserName.Substring(intBackSlash + 1)
        pwd = s.SQL_Return("SELECT * FROM `config` WHERE `name`='force_login';", "data")
        If (pwd = TextBox1.Text + TextBox2.Text + TextBox3.Text + TextBox4.Text + TextBox5.Text + TextBox6.Text) Then
            log("電腦登入成功", "" + computer + "被管理員使用密碼登入，系統取消登出", DisplayName)
            s.telegram("資訊:" + computer + "使用密碼登入成功" + vbNewLine + "--系統管理員啟")
            s.SQL_Return("UPDATE `seat` SET `borrowed_by`='" + DisplayName + "',`borrowed`='1' WHERE `computer_name`='" + computer + "'", "")
            accept_close = True
            Timer1.Enabled = False
            Dim Proc As New System.Diagnostics.Process
            Proc.StartInfo = New ProcessStartInfo("C:\Windows\System32\cmd.exe")
            Proc.StartInfo.FileName = "shutdown.exe"
            Proc.StartInfo.Arguments = "-a"
            Proc.StartInfo.UseShellExecute = False
            Proc.StartInfo.CreateNoWindow = True
            Proc.Start()
            ' Allows script to execute sequentially instead of simultaneously
            Proc.WaitForExit()
            MsgBox("系統取消登出!")
            Me.Close()

        Else
            second = second - 10
        End If
    End Sub


End Class
