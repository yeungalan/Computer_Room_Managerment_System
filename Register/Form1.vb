Public Class Form1
    Dim s As New SQL
    Private GIFAnim As Image = My.Resources.info
    Private frames As Integer
    Dim MetroProgressBar1 As New ProgressBar
    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        Timer1.Enabled = False
        Label8.Text = TextBox1.Text
        Dim UserName As String = My.User.Name
        Dim intBackSlash As Integer = UserName.LastIndexOf("\")
        Dim DisplayName As String = UserName.Substring(intBackSlash + 1)
        Dim ban As Boolean = False
        If s.SQL_Return("SELECT * FROM `banned_user` WHERE `user_id`='" + TextBox1.Text + "'", "user_id") = TextBox1.Text Then
            GIFAnim = My.Resources.warn
            frames = GIFAnim.GetFrameCount(Imaging.FrameDimension.Time)
            ImageAnimator.Animate(GIFAnim, AddressOf paintFrame)
            Label7.Text = "未分配座位"
            Timer1.Enabled = True
            Label6.Text = "閣下目前被停權，原因：" + s.SQL_Return("SELECT * FROM `banned_user` WHERE `user_id`='" + TextBox1.Text + "'", "reason") + "      "
            Label5.Text = "未知"
            ban = True


            log_warning("註冊失敗", TextBox1.Text + "被停權而注冊失敗", UserName)
            s.telegram("警告:" + vbNewLine + "帳戶" + TextBox1.Text + "被停權而注冊失敗" + vbNewLine + "--系統管理員啟")
        End If

        If ban = False Then
            Dim seat As String = s.SQL_Return("SELECT * FROM `seat` WHERE `borrowed`='0' AND `enabled`='1' ORDER BY `seat`.`computer_name` ASC LIMIT 1", "computer_name")
            If seat = "" Then
                GIFAnim = My.Resources.warn
                frames = GIFAnim.GetFrameCount(Imaging.FrameDimension.Time)
                ImageAnimator.Animate(GIFAnim, AddressOf paintFrame)

                Label7.Text = "無"
                Label6.Text = "無可用座位"
                Label5.Text = s.SQL_Return("SELECT COUNT(*) FROM `seat` WHERE `borrowed`='0' AND `enabled`='1'", "COUNT(*)")
                log_warning("註冊失敗", TextBox1.Text + "因沒有座位而注冊失敗", UserName)
            Else
                GIFAnim = My.Resources.success
                frames = GIFAnim.GetFrameCount(Imaging.FrameDimension.Time)
                ImageAnimator.Animate(GIFAnim, AddressOf paintFrame)
                s.SQL_Return("UPDATE `seat` SET `borrowed_by`='" + TextBox1.Text + "',`borrowed`='1' WHERE `computer_name`='" + seat + "'", "")
                Label7.Text = seat
                Label6.Text = "成功"
                Label5.Text = s.SQL_Return("SELECT COUNT(*) FROM `seat` WHERE `borrowed`='0' AND `enabled`='1'", "COUNT(*)")
         
                log("註冊座位成功", TextBox1.Text + "注冊成功，座位 :" + seat, UserName)
                s.telegram("資訊:" + vbNewLine + TextBox1.Text + "注冊成功，座位 :" + seat + vbNewLine + "--系統管理員啟")
            End If

        End If
        MetroButton1.Enabled = False
        Timer3.Enabled = True
    End Sub
    Private Sub paintFrame(ByVal sender As Object, ByVal e As EventArgs)
        If frames > 0 Then PictureBox1.Invalidate() Else ImageAnimator.StopAnimate(GIFAnim, AddressOf StopAnim)
    End Sub

    Private Sub StopAnim(ByVal sender As Object, ByVal e As EventArgs)
        PictureBox1.Invalidate()
    End Sub

    Private Sub PictureBox1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox1.Paint
        If frames > 0 Then
            ImageAnimator.UpdateFrames()
            e.Graphics.DrawImage(GIFAnim, Point.Empty)
            frames -= 1
        End If
    End Sub
    Function 文字旋轉(ByVal s As String, ByVal v As Integer) As String
        '--- S=目標字串， V=捲動次數（ +v 右捲, -v 左捲）---
        文字旋轉 = Mid(s & s & s, Len(s) - (v Mod Len(s)) + 1, Len(s))
    End Function
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Label6.Text = 文字旋轉(Label6.Text, -1)
    End Sub
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
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.FormBorderStyle = FormBorderStyle.None
        Me.Location = New Point(0, 0)
        Me.Size = SystemInformation.PrimaryMonitorSize
        GroupBox1.Location = New Point(Me.Width \ 2 - GroupBox1.Width \ 2, GroupBox1.Top)
        GroupBox2.Location = New Point(Me.Width \ 2 - GroupBox2.Width \ 2, GroupBox2.Top)
        PictureBox1.Location = New Point(Me.Width \ 2 - PictureBox1.Width \ 2, PictureBox1.Top)
        Try
            If My.Computer.Network.Ping("itdc", 1000) Then
                GIFAnim = My.Resources.info
                frames = GIFAnim.GetFrameCount(Imaging.FrameDimension.Time)
                ImageAnimator.Animate(GIFAnim, AddressOf paintFrame)
                Label8.Text = "N/A"
                Label7.Text = "N/A"
                Label6.Text = "伺服器響應Ping成功"
                Label5.Text = "N/A"
            Else
                GIFAnim = My.Resources.warn
                frames = GIFAnim.GetFrameCount(Imaging.FrameDimension.Time)
                ImageAnimator.Animate(GIFAnim, AddressOf paintFrame)
                Label8.Text = "N/A"
                Label7.Text = "N/A"
                Label6.Text = "伺服器未響應Ping"
                Label5.Text = "N/A"
            End If
        Catch ex As InvalidOperationException
            GIFAnim = My.Resources.warn
            frames = GIFAnim.GetFrameCount(Imaging.FrameDimension.Time)
            ImageAnimator.Animate(GIFAnim, AddressOf paintFrame)
            Label8.Text = "N/A"
            Label7.Text = "N/A"
            Label6.Text = "沒有可用的網路連線/URL 無效"
            Label5.Text = "N/A"
        End Try
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs)

    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        MetroButton1.Enabled = True
        Timer3.Enabled = False
    End Sub
End Class
