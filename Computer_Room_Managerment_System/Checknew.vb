Public Class CheckNew '檢查更新
    '字串常數
    Private Const DefaultInfoCode As String = "UTF-8" '儲存預設編碼為「UTF-8」
    Private Const DefaultInfoFeature As String = "crms" '儲存預設更新版本資訊的特徵為「電腦不難」
    Private Const DefaultInfoSeparator As String = "|" '儲存預設更新版本資訊的分隔字元符號為「|」
    Private Const DefaultNotCheckReturn As String = "Not yet." '儲存還沒檢查更新時所傳回的字串訊息
    Private Const DefaultCheckErrorReturn As String = "Error." '儲存檢查更新失敗時所傳回的字串訊息


    '更新資訊組態
    Private sInfoURL As String '儲存更新版本資訊的網址
    Private sInfoCode As String = DefaultInfoCode '編碼
    Private sInfoFeature As String = DefaultInfoFeature '儲存更新版本資訊的特徵，預設為「電腦不難」

    '檢查結果
    Private shInfoConsequence As Short = -1 '0為沒有新版本；1為有新版本；2為檢查更新失敗；-1為尚未檢查
    Private sInfoNumber As String '儲存取得的版本號碼

    Public Sub New(ByVal SetInfoURL As String) 'CheckNew建構子(多載1)。引數為SetInfoURL，載入更新版本資訊的網址。
        sInfoURL = SetInfoURL '設定更新版本資訊的網址
    End Sub

    Public Sub New(ByVal SetInfoURL As String, ByVal SetInfoFeature As String) 'CheckNew建構子(多載2)。引數為SetInfoURL和SetInfoFeature，載入更新版本資訊的網址和判斷特徵。
        sInfoURL = SetInfoURL '設定更新版本資訊的網址
        sInfoFeature = SetInfoFeature '設定更新版本資訊的特徵
    End Sub

    Public WriteOnly Property SetInfoCode() As String '設定更新版本資訊的編碼
        Set(ByVal Code As String)
            sInfoCode = Code
        End Set
    End Property

    Public WriteOnly Property SetInfoURL() As String '設定網址
        Set(ByVal URL As String)
            sInfoURL = URL
        End Set
    End Property

    Public WriteOnly Property SetInfoFeature() As String '設定特徵
        Set(ByVal Feature As String)
            sInfoFeature = Feature
        End Set
    End Property

    Public ReadOnly Property GetMyVersion() As String '取得程式的版本號碼(格式：X.X.X)
        Get
            Return My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Build '取得版本號碼
        End Get
    End Property

    Public ReadOnly Property GetCheckConsequence() As String '取得檢查更新的結果
        Get
            Return shInfoConsequence
        End Get
    End Property

    Public ReadOnly Property GetCheckConsequenceNumber() As String '取得最新版本號碼
        Get
            If shInfoConsequence = -1 Then '如果還沒檢查
                Return DefaultNotCheckReturn
            ElseIf shInfoConsequence = 2 Then '如果檢查失敗
                Return DefaultCheckErrorReturn
            Else '如果檢查成功
                Return sInfoNumber
            End If
        End Get
    End Property

    Public Sub CheckNew() '檢查更新
        On Error Resume Next
        Dim CheckNewClient As New Net.WebClient() '宣告CheckNewClient為WebClient物件
        Dim Encode As System.Text.Encoding = System.Text.Encoding.GetEncoding(sInfoCode) '設定編碼
        Dim Info As IO.StreamReader = New IO.StreamReader(CheckNewClient.OpenRead(sInfoURL), Encode) '取得更新資訊
        Dim sInfoData As String = Info.ReadLine '儲存更新資訊

        Debug.Print("InfoData = """ & sInfoData & """") '(除錯用)顯示出從網路上下載的字串

        If Strings.InStr(1, sInfoData, sInfoFeature) > 0 Then '若能找到識別關鍵字(InfoFeature)，則更新資訊取得成功
            Dim InfoArr() As String = Split(sInfoData, DefaultInfoSeparator) '分解字串為版本號碼(0)和識別關鍵字(1)
            sInfoNumber = InfoArr(0) '儲存最新版本號碼
            If InfoArr(0) <> GetMyVersion() Then '比對版本號碼，若不相同，代表有新版本
                shInfoConsequence = 1
            Else '若版本號碼相同，則代表沒有新版本
                shInfoConsequence = 0
            End If
        Else '若未能找到識別關鍵字(專案名稱)，則更新資訊取得失敗
            shInfoConsequence = 2
        End If
    End Sub
End Class