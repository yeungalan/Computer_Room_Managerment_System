Public Class Form1
    Dim s As New SQL
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim UserName As String = My.User.Name
        Dim intBackSlash As Integer = UserName.LastIndexOf("\")
        Dim DisplayName As String = UserName.Substring(intBackSlash + 1)
        s.telegram("資訊:" + My.Computer.Name + "電腦登出" + vbNewLine + "--系統管理員啟")
        s.SQL_Return("UPDATE `seat` SET `borrowed_by`='',`borrowed`='0',`logout_time`= CURRENT_TIMESTAMP WHERE `computer_name`='" + My.Computer.Name + "'", "")
        log("電腦登出成功", "" + My.Computer.Name + "登出成功", DisplayName)
        s.telegram("資訊:" + My.Computer.Name + "電腦登出成功" + vbNewLine + "--系統管理員啟")
        Application.Exit()

    End Sub

    Public Function log(ByVal type As String, ByVal info As String, ByVal ppl As String)
        Dim UserName As String = My.User.Name
        Dim intBackSlash As Integer = UserName.LastIndexOf("\")
        Dim DisplayName As String = UserName.Substring(intBackSlash + 1)
        s.SQL_Return("INSERT INTO `logs` (`uid`, `time`, `level`, `type`, `details`, `operator`, `affected_people`) VALUES (NULL, CURRENT_TIMESTAMP, '資訊', '" + type + "', '[資訊][" + Now().ToString + "]" + info + "', '" + DisplayName + "', '" + ppl + "');", "details")
        Return 0
    End Function
End Class
