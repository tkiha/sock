Imports System
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading
Imports Microsoft.VisualBasic




''' <summary>
''' 送受信管理クラス
''' </summary>
''' <remarks>送受信管理クラス</remarks>
Public Class StateObject
    ''' <summary>
    ''' 送受信バッファ
    ''' </summary>
    ''' <remarks></remarks>
    Private _Buffer As Byte()

    ''' <summary>
    ''' ワークソケット
    ''' </summary>
    ''' <remarks>ワークソケット</remarks>
    Public Property WorkSocket() As Socket = Nothing

    ''' <summary>
    ''' 受信サイズ最大
    ''' </summary>
    ''' <remarks>受信サイズ最大</remarks>
    Public Const BufferSize As Integer = 1024

    ''' <summary>
    ''' 受信バッファ
    ''' </summary>
    ''' <remarks>受信バッファ</remarks>
    Public ReadOnly Property Buffer As Byte()
        Get
            Return _Buffer
        End Get
    End Property


    ''' <summary>
    ''' 受信データを文字にしたもの
    ''' </summary>
    ''' <remarks>受信データを文字にしたもの</remarks>
    Public Property Sb() As New StringBuilder

    ''' <summary>
    ''' 受信データ
    ''' </summary>
    ''' <remarks>受信データ</remarks>
    Public Property bRecv() As Byte()

    ''' <summary>
    ''' 送信データ
    ''' </summary>
    ''' <remarks>送信データ</remarks>
    Public Property bSend() As Byte()

    ''' <summary>
    ''' 相手IP
    ''' </summary>
    ''' <remarks>相手IP</remarks>
    Public Property RemoteIP() As String

    ''' <summary>
    ''' 相手PORT
    ''' </summary>
    ''' <remarks>相手PORT</remarks>
    Public Property RemotePort() As String

    ''' <summary>
    ''' 自分IP
    ''' </summary>
    ''' <remarks>自分IP</remarks>
    Public Property LocalIP() As String

    ''' <summary>
    ''' 自分PORT
    ''' </summary>
    ''' <remarks>自分PORT</remarks>
    Public Property LocalPort() As String


    Public Sub New()
        _Buffer = New Byte(BufferSize) {}
    End Sub
End Class 'StateObject


''' <summary>
''' ソケットメッセージイベント引数
''' </summary>
''' <remarks>ソケットメッセージイベント引数</remarks>
Public Class SockMessageEventArgs
    Inherits EventArgs

    ''' <summary>
    ''' メッセージ
    ''' </summary>
    ''' <remarks>メッセージ</remarks>
    Public Property Msg() As String
End Class

''' <summary>
''' ソケットエラーイベント引数
''' </summary>
''' <remarks>ソケットエラーイベント引数</remarks>
Public Class SockErrorEventArgs
    Inherits EventArgs

    ''' <summary>
    ''' メッセージ
    ''' </summary>
    ''' <remarks>メッセージ</remarks>
    Public Property Msg() As String

	''' <summary>
	''' エラー種別
	''' </summary>
	''' <remarks>エラー種別</remarks>
    Public Property ErrorType() As SockErrors

	''' <summary>
	''' 相手IP (エラー種別によっては入ってこない）
	''' </summary>
	''' <remarks>相手IP (エラー種別によっては入ってこない）</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' 相手PORT(エラー種別によっては入ってこない）
	''' </summary>
	''' <remarks>相手PORT(エラー種別によっては入ってこない）</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' 送信時に設定したKey
	''' </summary>
	''' <remarks>送信時に設定したKey</remarks>
    Public Property SendKey() As String

End Class

''' <summary>
''' ソケット受信開始イベント引数
''' </summary>
''' <remarks>ソケット受信開始イベント引数</remarks>
Public Class SockReceiveStartEventArgs
    Inherits EventArgs

	''' <summary>
	''' 相手IP
	''' </summary>
	''' <remarks>相手IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' 相手PORT
	''' </summary>
	''' <remarks>相手PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' 自分IP
	''' </summary>
	''' <remarks>自分IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' 自分PORT
	''' </summary>
	''' <remarks>自分PORT</remarks>
    Public Property LocalPort() As String

End Class

''' <summary>
''' ソケット受信完了イベント引数
''' </summary>
''' <remarks>ソケット受信完了イベント引数</remarks>
Public Class SockReceiveEndEventArgs
    Inherits EventArgs

	''' <summary>
	''' 相手IP
	''' </summary>
	''' <remarks>相手IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' 相手PORT
	''' </summary>
	''' <remarks>相手PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' 自分IP
	''' </summary>
	''' <remarks>自分IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' 自分PORT
	''' </summary>
	''' <remarks>自分PORT</remarks>
    Public Property LocalPort() As String

	''' <summary>
	''' 受信内容
	''' </summary>
	''' <remarks>受信内容</remarks>
    Public Property RecvChar() As String

	''' <summary>
	''' 送信時に設定したKey
	''' </summary>
	''' <remarks>送信時に設定したKey</remarks>
    Public Property SendKey() As String

End Class

''' <summary>
''' ソケット送信完了イベント引数
''' </summary>
''' <remarks>ソケット送信完了イベント引数</remarks>
Public Class SockSendEndEventArgs
    Inherits EventArgs

	''' <summary>
	''' 相手IP
	''' </summary>
	''' <remarks>相手IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' 相手PORT
	''' </summary>
	''' <remarks>相手PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' 自分IP
	''' </summary>
	''' <remarks>自分IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' 自分PORT
	''' </summary>
	''' <remarks>自分PORT</remarks>
    Public Property LocalPort() As String

	''' <summary>
	''' 送信内容
	''' </summary>
	''' <remarks>送信内容</remarks>
    Public Property SendChar() As String

	''' <summary>
	''' 送信時に設定したKey
	''' </summary>
	''' <remarks>送信時に設定したKey</remarks>
    Public Property SendKey() As String


End Class

''' <summary>
''' ソケットListen完了イベント引数
''' </summary>
''' <remarks>ソケットListen完了イベント引数</remarks>
Public Class SockListenEndEventArgs
    Inherits EventArgs

	''' <summary>
	''' 自分IP
	''' </summary>
	''' <remarks>自分IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' 自分PORT
	''' </summary>
	''' <remarks>自分PORT</remarks>
    Public Property LocalPort() As String

End Class

''' <summary>
''' ソケット接続完了イベント引数
''' </summary>
''' <remarks>ソケット接続完了イベント引数</remarks>
Public Class SockConnectEndEventArgs
    Inherits EventArgs

	''' <summary>
	''' 相手IP
	''' </summary>
	''' <remarks>相手IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' 相手PORT
	''' </summary>
	''' <remarks>相手PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' 自分IP
	''' </summary>
	''' <remarks>自分IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' 自分PORT
	''' </summary>
	''' <remarks>自分PORT</remarks>
    Public Property LocalPort() As String

End Class

''' <summary>
''' ソケット接続リトライイベント引数
''' </summary>
''' <remarks>ソケット接続リトライイベント引数</remarks>
Public Class SockConnectRetryEventArgs
    Inherits EventArgs

	''' <summary>
	''' 相手IP
	''' </summary>
	''' <remarks>相手IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' 相手PORT
	''' </summary>
	''' <remarks>相手PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' リトライ回数
	''' </summary>
	''' <remarks>リトライ回数</remarks>
    Public Property RetryDo() As Long

End Class

''' <summary>
''' ソケットAccept完了イベント引数
''' </summary>
''' <remarks>ソケットAccept完了イベント引数</remarks>
Public Class SockAcceptEndEventArgs
    Inherits EventArgs

	''' <summary>
	''' 相手IP
	''' </summary>
	''' <remarks>相手IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' 相手PORT
	''' </summary>
	''' <remarks>相手PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' 自分IP
	''' </summary>
	''' <remarks>自分IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' 自分PORT
	''' </summary>
	''' <remarks>自分PORT</remarks>
    Public Property LocalPort() As String

End Class

''' <summary>
''' ソケット送信リトライイベント引数
''' </summary>
''' <remarks>ソケット送信リトライイベント引数</remarks>
Public Class SockSendRetryEventArgs
    Inherits EventArgs

	''' <summary>
	''' 相手IP
	''' </summary>
	''' <remarks>相手IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' 相手PORT
	''' </summary>
	''' <remarks>相手PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' 自分IP
	''' </summary>
	''' <remarks>自分IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' 自分PORT
	''' </summary>
	''' <remarks>自分PORT</remarks>
    Public Property LocalPort() As String

	''' <summary>
	''' リトライ回数
	''' </summary>
	''' <remarks>リトライ回数</remarks>
    Public Property RetryDo() As Long

	''' <summary>
	''' 送信時に設定したKey
	''' </summary>
	''' <remarks>送信時に設定したKey</remarks>
    Public Property SendKey() As String

End Class







''' <summary>
''' ソケット結果
''' </summary>
''' <remarks>ソケット結果</remarks>
Public Enum SockErrors As Integer

	''' <summary>
	''' BINDエラー
	''' </summary>
	''' <remarks>BINDエラー</remarks>
	BindErr = 1

	''' <summary>
	''' Listenエラー
	''' </summary>
	''' <remarks>Listenエラー</remarks>
	ListenErr = 2

	''' <summary>
	''' Acceptエラー
	''' </summary>
	''' <remarks>Acceptエラー</remarks>
	AcceptErr = 3

	''' <summary>
	''' Connectエラー
	''' </summary>
	''' <remarks>Connectエラー</remarks>
	ConnectErr = 4

	''' <summary>
	''' Sendエラー
	''' </summary>
	''' <remarks>Sendエラー</remarks>
	SendErr = 5

	''' <summary>
	''' Sendリトライオーバ
	''' </summary>
	''' <remarks>Sendリトライオーバ</remarks>
	SendRetryOver = 6

    ''' <summary>
    ''' Receiveエラー
    ''' </summary>
    ''' <remarks>Receiveエラー</remarks>
	RecvErr = 8

	''' <summary>
	''' BeginReceiveエラー
	''' </summary>
	''' <remarks>BeginReceiveエラー</remarks>
	BeginRecvErr = 9

	''' <summary>
	''' Connectリトライオーバ
	''' </summary>
	''' <remarks>Connectリトライオーバ</remarks>
	ConnectRetryOver = 10

	''' <summary>
	''' その他エラー
	''' </summary>
	''' <remarks>その他エラー</remarks>
	OtherErr = 999

End Enum

''' <summary>
''' ソケット状態
''' </summary>
''' <remarks>ソケット状態</remarks>
Public Enum SockStatuses As Integer

	''' <summary>
	''' 接続中
	''' </summary>
	''' <remarks>接続中</remarks>
	Connecting = 1

	''' <summary>
	''' 接続した
	''' </summary>
	''' <remarks>接続した</remarks>
	Connected = 2

	''' <summary>
	''' 送信中
	''' </summary>
	''' <remarks>送信中</remarks>
	Sending = 3

	''' <summary>
	''' 送信した
	''' </summary>
	''' <remarks>送信した</remarks>
	Sended = 4

	''' <summary>
	''' 受信中
	''' </summary>
	''' <remarks>受信中</remarks>
	Receiving = 5

	''' <summary>
	''' 受信した
	''' </summary>
	''' <remarks>受信した</remarks>
	Received = 6

	''' <summary>
	''' BIND中
	''' </summary>
	''' <remarks>BIND中</remarks>
	Binding = 7

	''' <summary>
	''' BINDした
	''' </summary>
	''' <remarks>BINDした</remarks>
	Binded = 8

	''' <summary>
	''' Listen中
	''' </summary>
	''' <remarks>Listen中</remarks>
	Listening = 9

	''' <summary>
	''' Listenした
	''' </summary>
	''' <remarks>Listenした</remarks>
	Listned = 10

	''' <summary>
	''' 送信待ち中
	''' </summary>
	''' <remarks>送信待ち中</remarks>
    SendWaiting = 11

End Enum




''' <summary>
''' 受信ソケットクラス
''' </summary>
''' <remarks>受信ソケットクラス</remarks>
Public Class RecvSockClass

    ''' <summary>
    ''' 受信終了判定文字列。この文字が来たら電文の終了とする。この文字を指定した場合、この文字が来るまで電文を連結して受信を続ける。指定していない場合は１回の通信単位＝電文単位となる
    ''' </summary>
    ''' <remarks>
    ''' 受信終了判定文字列。この文字が来たら電文の終了とする。この文字を指定した場合、この文字が来るまで電文を連結して受信を続ける。指定していない場合は１回の通信単位＝電文単位となる
    ''' </remarks>
    Public Property EndCheckChar() As Byte()

    ''' <summary>
    ''' 待ち受けIP
    ''' </summary>
    ''' <remarks>待ち受けIP</remarks>
    Public Property LocalIP() As String

    ''' <summary>
    ''' 待ち受けPort
    ''' </summary>
    ''' <remarks>待ち受けPort</remarks>
    Public Property LocalPort() As String

    ''' <summary>
    ''' タグ（呼び出し側で自由に使ってよい）
    ''' </summary>
    ''' <remarks>タグ（呼び出し側で自由に使ってよい）</remarks>
    Public TAG As String


    ''' <summary>
    ''' 受信ステータス
    ''' </summary>
    ''' <remarks>受信ステータス</remarks>
    Private Property RecvStatus() As SockStatuses

    ''' <summary>
    ''' Bind処理のリミット
    ''' </summary>
    ''' <remarks>Bind処理のリミット</remarks>
    Private Property BindTimeLimit() As DateTime

    ''' <summary>
    ''' Bind処理タイムアウトミリ秒（とりあえず固定、必要であればPublicにしても良い）
    ''' </summary>
    ''' <remarks>Bind処理タイムアウトミリ秒（とりあえず固定、必要であればPublicにしても良い）</remarks>
    Private BindTimeOut As Long = 3000

    ''' <summary>
    ''' Listen処理のリミット
    ''' </summary>
    ''' <remarks>Listen処理のリミット</remarks>
    Private ListenTimeLimit As DateTime

    ''' <summary>
    ''' Listen処理タイムアウトミリ秒（とりあえず固定、必要であればPublicにしても良い）
    ''' </summary>
    ''' <remarks>Listen処理タイムアウトミリ秒（とりあえず固定、必要であればPublicにしても良い）</remarks>
    Private ListenTimeOut As Long = 3000

    ''' <summary>
    ''' ソケット受付監視タイマー
    ''' </summary>
    ''' <remarks>ソケット受付監視タイマー</remarks>
    Private WithEvents ListenCheckTimer As New System.Timers.Timer

    ''' <summary>
    ''' ソケット受付監視タイマー処理中フラグ
    ''' </summary>
    ''' <remarks>ソケット受付監視タイマー処理中フラグ</remarks>
    Private ListenCheckTimerExecuting As Boolean = False

    ''' <summary>
    ''' 送信状態監視タイマー
    ''' </summary>
    ''' <remarks>送信状態監視タイマー</remarks>
    Private WithEvents SendCheckTimer As New System.Timers.Timer

    ''' <summary>
    ''' 送信状態監視タイマー処理中フラグ
    ''' </summary>
    ''' <remarks>送信状態監視タイマー処理中フラグ</remarks>
    Private SendCheckTimerExecuting As Boolean = False '


    ''' <summary>
    ''' 送信用パラメータ
    ''' </summary>
    ''' <remarks>送信メソッドのパラメータ</remarks>
    Public Class SendParmClass
        '呼び出し側でセットするパラメータ

        ''' <summary>
        ''' ソケット識別キー（呼び出し側で自由に使ってよい）
        ''' </summary>
        ''' <remarks>ソケット識別キー（呼び出し側で自由に使ってよい）</remarks>
        Public SendKEY As String

        ''' <summary>
        ''' 送信電文
        ''' </summary>
        ''' <remarks>送信電文</remarks>
        Public Property sndCmd As Byte()

        ''' <summary>
        ''' 送信タイムアウトミリ秒。既定値1000
        ''' </summary>
        ''' <remarks>送信タイムアウトミリ秒</remarks>
        Public Property SendTimeOut() As Long = 1000

        ''' <summary>
        ''' 送信リトライを行う回数。既定値0
        ''' </summary>
        ''' <remarks>送信リトライを行う回数</remarks>
        Public Property SendRetry() As Long = 0

        ''' <summary>
        ''' 送信リトライ間隔（ミリ秒）。既定値1000
        ''' </summary>
        ''' <remarks>送信リトライ間隔（ミリ秒）</remarks>
        Public Property SendRetryWait() As Long = 1000

        ''' <summary>
        ''' 送信相手IP
        ''' </summary>
        ''' <remarks>送信相手IP</remarks>
        Public Property RemoteIP() As String

        ''' <summary>
        ''' 送信相手ポート
        ''' </summary>
        ''' <remarks>送信相手ポート</remarks>
        Public Property RemotePort() As String

        '処理内で使うパラメータ

        ''' <summary>
        ''' 送信リトライを行った回数
        ''' </summary>
        ''' <remarks>送信リトライを行った回数</remarks>
        Public Property SendRetryDo() As Long = 0

        ''' <summary>
        ''' 送信処理のリミット（送信開始日時＋送信タイムアウト）
        ''' </summary>
        ''' <remarks>送信処理のリミット（送信開始日時＋送信タイムアウト）</remarks>
        Public Property SendTimeLimit() As Date

        ''' <summary>
        ''' 送信ステータス
        ''' </summary>
        ''' <remarks>送信ステータス</remarks>
        Public Property SendStatus() As SockStatuses

    End Class

    ''' <summary>
    ''' 送信電文の待ち行列
    ''' </summary>
    ''' <remarks>送信電文の待ち行列</remarks>
    Private SendCmdList As New System.Collections.Generic.List(Of SendParmClass)


    ''' <summary>
    ''' メッセージ出力イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>メッセージ出力イベント</remarks>
    Public Event OnMessage(ByVal sender As Object, ByVal e As SockMessageEventArgs)

    ''' <summary>
    ''' エラー発生イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>エラー発生イベント</remarks>
    Public Event OnError(ByVal sender As Object, ByVal e As SockErrorEventArgs)

    ''' <summary>
    ''' LISTEN完了イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>LISTEN完了イベント</remarks>
    Public Event OnListenEnd(ByVal sender As Object, ByVal e As SockListenEndEventArgs)

    ''' <summary>
    ''' ACCEPT完了イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>ACCEPT完了イベント</remarks>
    Public Event OnAcceptEnd(ByVal sender As Object, ByVal e As SockAcceptEndEventArgs)

    ''' <summary>
    ''' 受信開始イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>受信開始イベント</remarks>
    Public Event OnRecvStart(ByVal sender As Object, ByVal e As SockReceiveStartEventArgs)

    ''' <summary>
    ''' 受信完了イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>受信完了イベント</remarks>
    Public Event OnRecvEnd(ByVal sender As Object, ByVal e As SockReceiveEndEventArgs)

    ''' <summary>
    ''' 送信リトライ開始イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>送信リトライ開始イベント</remarks>
    Public Event OnSendRetry(ByVal sender As Object, ByVal e As SockSendRetryEventArgs)

    ''' <summary>
    ''' 送信完了イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>送信完了イベント</remarks>
    Public Event OnSendEnd(ByVal sender As Object, ByVal e As SockSendEndEventArgs)


    ''' <summary>
    ''' 接続毎に持つ
    ''' </summary>
    ''' <remarks>接続毎に持つ</remarks>
    Private tState(0) As StateObject

    ''' <summary>
    ''' ソケット
    ''' </summary>
    ''' <remarks>ソケット</remarks>
    Private listener As Socket

    ''' <summary>
    ''' ソケット生存確認
    ''' </summary>
    ''' <returns>True:生きている False:死んでいる</returns>
    ''' <remarks>ソケット生存確認</remarks>
    Public Function IsAlive()
        Dim result As Boolean = False

        Try
            If listener.Poll(1000, Sockets.SelectMode.SelectRead) = True Then
                result = False
            Else
                result = True
            End If
        Catch

        End Try

        Return result

    End Function

    ''' <summary>
    ''' サーバソケット開始
    ''' </summary>
    ''' <remarks>サーバソケット開始</remarks>
    Public Sub RcvSockStart()

        Try
            RecvStatus = SockStatuses.Binding
            BindTimeLimit = DateTime.Now.AddMilliseconds(BindTimeOut)

            ' Data buffer for incoming data.
            Dim bytes() As Byte = New [Byte](1023) {}

            Dim cipAddress As IPAddress = IPAddress.Parse(LocalIP)
            Dim localEndPoint As New IPEndPoint(cipAddress, CInt(LocalPort))

            Call CloseSockAll()

            '監視タイマーを起動
            ListenCheckTimer.Interval = 10
            ListenCheckTimer.Start()


            ' Create a TCP/IP socket.
            listener = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)

            ' Bind the socket to the local endpoint and listen for incoming connections.
            Try
                listener.Bind(localEndPoint)
            Catch ex As Exception
                Dim msg As String = LocalIP & ":" & LocalPort & " Bindエラー ソケット受信できません:" & ex.Message
                Call RaiseMessage(msg)
                Return
            End Try
            RecvStatus = SockStatuses.Binded

            Try
                ListenTimeLimit = DateTime.Now.AddMilliseconds(BindTimeOut)
                RecvStatus = SockStatuses.Listening
                listener.Listen(100)
            Catch ex As Exception
                Dim msg As String = LocalIP & ":" & LocalPort & " Listenエラー ソケット受信できません:" & ex.Message
                Call RaiseMessage(msg)
                Return
            End Try
            RecvStatus = SockStatuses.Listned

            Call RaiseMessage("Listen開始:" & LocalIP & "(" & LocalPort & ")")
            SendCheckTimer.Interval = 10
            SendCheckTimer.Start()

        Catch ex As Exception
            Dim msg As String = LocalIP & ":" & LocalPort & " ソケット開始できません:" & ex.Message
            Call RaiseMessage(msg)

        End Try

    End Sub 'Main


    ''' <summary>
    ''' Listen状態監視タイマー
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>Listen状態監視タイマー</remarks>
    Private Sub ListenCheckTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles ListenCheckTimer.Elapsed
        If ListenCheckTimerExecuting Then
            'タイマー処理実行中であればreturn
            Return
        End If


        Try
            ListenCheckTimerExecuting = True


            Select Case RecvStatus
                Case SockStatuses.Binding
                    'BINDタイムアウトチェック
                    If DateTime.Now >= BindTimeLimit Then
                        'タイムアウト

                        Dim msg As String = LocalIP & ":" & LocalPort & " BINDタイムアウト"

                        Call RaiseError(msg, SockErrors.BindErr)

                        ListenCheckTimer.Stop()

                    End If
                Case SockStatuses.Listening
                    'Listenタイムアウトチェック
                    If DateTime.Now >= ListenTimeLimit Then
                        'タイムアウト

                        Dim msg As String = LocalIP & ":" & LocalPort & " Listenタイムアウト"

                        Call RaiseError(msg, SockErrors.ListenErr)

                        ListenCheckTimer.Stop()

                    End If
                Case SockStatuses.Listned
                    'Listen状態になった
                    Call RaiseListenEnd(LocalIP, LocalPort)
                    ListenCheckTimer.Stop()
                    Call BeginAccept()

            End Select


        Catch ex As Exception
            Dim msg As String = LocalIP & ":" & LocalPort & " ListenCheckTimer_Elapsedエラー:" & ex.Message
            Call RaiseMessage(msg)

        End Try
        ListenCheckTimerExecuting = False

    End Sub

    ''' <summary>
    ''' すべてのソケットを閉じる
    ''' </summary>
    ''' <remarks>すべてのソケットを閉じる</remarks>
    Public Sub CloseSockAll()
        Try
            ListenCheckTimer.Stop()
            SendCheckTimer.Stop()

            For l = LBound(tState) To UBound(tState)
                If tState(l) Is Nothing Then

                Else
                    Call CloseSock(tState(l).RemoteIP, tState(l).RemotePort)
                End If
            Next
            Try
                listener.Shutdown(SocketShutdown.Both)
            Catch ex As Exception

            End Try
            listener.Close()
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' 指定されたソケットを閉じる
    ''' </summary>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <remarks>指定されたソケットを閉じる</remarks>
    Public Sub CloseSock(ByVal remoteIP As String, ByVal remotePort As String)
        Try
            Dim l As Long = FindStateIndex(remoteIP, remotePort)
            If l >= 0 Then
                With tState(l)
                    Try
                        .WorkSocket.Shutdown(SocketShutdown.Both)
                    Catch ex As Exception

                    End Try
                    Try
                        .WorkSocket.Close()
                    Catch ex As Exception

                    End Try

                    tState(l) = Nothing
                End With
            End If
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' メッセージイベントを発生させる
    ''' </summary>
    ''' <param name="msg">メッセージ</param>
    ''' <remarks>メッセージイベントを発生させる</remarks>
    Private Sub RaiseMessage(ByVal msg As String)
        Dim e As New SockMessageEventArgs
        e.Msg = msg
        RaiseEvent OnMessage(Me, e)
    End Sub

    ''' <summary>
    ''' エラーイベントを発生させる
    ''' </summary>
    ''' <param name="msg">メッセージ</param>
    ''' <param name="sockError">ソケット結果</param>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <param name="sendKey">送信時に設定したKey</param>
    ''' <remarks>
    ''' エラーイベントを発生させる
    ''' </remarks>
    Private Sub RaiseError(ByVal msg As String, ByVal sockError As SockErrors, _
            Optional ByVal remoteIP As String = "", _
            Optional ByVal remotePort As String = "", _
            Optional ByVal sendKey As String = "")
        Dim e As New SockErrorEventArgs
        e.Msg = msg
        e.ErrorType = sockError
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort
        e.SendKey = sendKey

        If Me.IsAlive And e.ErrorType = SockErrors.AcceptErr Then
            Call BeginAccept() '再度ソケット受信受付する
        End If

        RaiseEvent OnError(Me, e)
    End Sub

    ''' <summary>
    ''' Listen完了イベントを発生させる
    ''' </summary>
    ''' <param name="localIP">自IP</param>
    ''' <param name="localPort">自PORT</param>
    ''' <remarks>Listen完了イベントを発生させる</remarks>
    Private Sub RaiseListenEnd(ByVal localIP As String, ByVal localPort As String)
        Dim e As New SockListenEndEventArgs
        e.LocalIP = localIP
        e.LocalPort = localPort

        RaiseEvent OnListenEnd(Me, e)
    End Sub

    ''' <summary>
    ''' 接続受付完了イベントを発生させる
    ''' </summary>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <param name="localIP">自IP</param>
    ''' <param name="localPort">自PORT</param>
    ''' <remarks>接続受付完了イベントを発生させる</remarks>
    Private Sub RaiseAcceptEnd(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String)
        Dim e As New SockAcceptEndEventArgs
        e.LocalIP = localIP
        e.LocalPort = localPort
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort

        Call BeginAccept() '再度ソケット受信受付する



        RaiseEvent OnAcceptEnd(Me, e)
    End Sub

    ''' <summary>
    ''' 受信開始イベントを発生させる
    ''' </summary>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <param name="localIP">自IP</param>
    ''' <param name="localPort">自PORT</param>
    ''' <remarks>受信開始イベントを発生させる</remarks>
    Private Sub RaiseRecvStart(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String)
        Dim e As New SockReceiveStartEventArgs
        e.LocalIP = localIP
        e.LocalPort = localPort
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort

        RaiseEvent OnRecvStart(Me, e)
    End Sub
    ''' <summary>
    ''' 受信完了イベントを発生させる
    ''' </summary>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <param name="localIP">自IP</param>
    ''' <param name="localPort">自PORT</param>
    ''' <param name="recvChar">受信内容</param>
    ''' <remarks>受信完了イベントを発生させる</remarks>
    Private Sub RaiseRecvEnd(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String, ByVal recvChar As String)
        Dim e As New SockReceiveEndEventArgs
        e.LocalIP = localIP
        e.LocalPort = localPort
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort
        e.RecvChar = recvChar

        RaiseEvent OnRecvEnd(Me, e)
    End Sub

    ''' <summary>
    ''' 送信リトライイベントを発生させる
    ''' </summary>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <param name="localIP">自IP</param>
    ''' <param name="localPort">自PORT</param>
    ''' <param name="RetryDoCnt">リトライ回数</param>
    ''' <param name="sendKey">送信KEY</param>
    ''' <remarks>送信リトライイベントを発生させる</remarks>
    Private Sub RaiseSendRetry(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String, ByVal RetryDoCnt As Long, ByVal sendKey As String)
        Dim e As New SockSendRetryEventArgs
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort
        e.LocalIP = localIP
        e.LocalPort = localPort
        e.RetryDo = RetryDoCnt
        e.SendKey = sendKey

        RaiseEvent OnSendRetry(Me, e)
    End Sub

    ''' <summary>
    ''' 送信完了イベントを発生させる
    ''' </summary>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <param name="localIP">自IP</param>
    ''' <param name="localPort">自PORT</param>
    ''' <param name="sendChar">送信内容</param>
    ''' <param name="sendkey">送信KEY</param>
    ''' <remarks>送信完了イベントを発生させる</remarks>
    Private Sub RaiseSendEnd(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String, ByVal sendChar As String, ByVal sendkey As String)
        Dim e As New SockSendEndEventArgs
        e.LocalIP = localIP
        e.LocalPort = localPort
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort
        e.SendChar = sendChar
        e.SendKey = sendkey

        RaiseEvent OnSendEnd(Me, e)

    End Sub


    ''' <summary>
    ''' Accept開始処理
    ''' </summary>
    ''' <remarks>Accept開始処理</remarks>
    Private Sub BeginAccept()
        Try
            listener.BeginAccept(New AsyncCallback(AddressOf AcceptCallback), listener)
        Catch ex As Exception
            Dim msg As String = LocalIP & ":" & LocalPort & " BeginAcceptエラー:" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.AcceptErr)

        End Try

    End Sub

    ''' <summary>
    ''' Accept開始コールバック
    ''' </summary>
    ''' <param name="ar">ソケット</param>
    ''' <remarks>Accept開始コールバック</remarks>
    Private Sub AcceptCallback(ByVal ar As IAsyncResult)

        Dim listener As Socket = CType(ar.AsyncState, Socket)
        Dim handler As Socket
        Try
            handler = listener.EndAccept(ar)
        Catch ex As ObjectDisposedException
            Dim msg As String = LocalIP & ":" & LocalPort & " Closeされました"
            Call RaiseMessage(msg)
            Return
        Catch ex As Exception
            'listen中にクローズされた場合の処理
            Dim msg As String = LocalIP & ":" & LocalPort & " Acceptエラー:" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.AcceptErr)

            Return
        End Try



        Dim remoteIP As String = CType(handler.RemoteEndPoint, IPEndPoint).Address.ToString()
        Dim remotePORT As String = CType(handler.RemoteEndPoint, IPEndPoint).Port.ToString()

        Try
            Dim l = GetStateIndex(remoteIP, remotePORT, LocalIP, LocalPort)

            Call RaiseMessage(remoteIP & ":" & remotePORT & " から接続あり StateNo:" & l)

            Call RaiseAcceptEnd(remoteIP, remotePORT, LocalIP, LocalPort)




            With tState(l)
                .WorkSocket = handler

            End With

            Try
                handler.BeginReceive(tState(l).Buffer, 0, StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReadCallback), tState(l))
            Catch ex As Exception
                Dim msg As String = LocalIP & ":" & LocalPort & " BeginReceiveエラー:" & ex.Message
                Call RaiseMessage(msg)
                Call RaiseError(msg, SockErrors.BeginRecvErr, remoteIP, remotePORT)
            End Try

        Catch ex As Exception
            Dim msg As String = LocalIP & ":" & LocalPort & " AcceptCallbackエラー:" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.AcceptErr, remoteIP, remotePORT)
        End Try

    End Sub 'AcceptCallback

    ''' <summary>
    ''' 指定されたソケットのインデックスを返す。なければ新たに作成する
    ''' </summary>
    ''' <param name="remoteIP">指定IP</param>
    ''' <param name="remotePort">指定PORT</param>
    ''' <param name="localIP">自IP</param>
    ''' <param name="localPort">自PORT</param>
    ''' <returns>インデックス</returns>
    ''' <remarks>指定されたソケットのインデックスを返す</remarks>
    Private Function GetStateIndex(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String) As Long
        Dim lRet As Long
        Dim l As Long
        lRet = -999
        Try


            '指定されたIP,Portのステートが存在しているか調べる
            For l = LBound(tState) To UBound(tState)
                If tState(l) Is Nothing Then
                Else
                    If tState(l).RemoteIP = remoteIP And tState(l).RemotePort = remotePort Then
                        '割り当て済みなのでこのステートを使う
                        lRet = l
                        Exit For
                    End If
                End If
            Next

            If lRet = -999 Then
                '指定されたIP,Portのステートは存在していなかったので空きステートを探す
                For l = LBound(tState) To UBound(tState)
                    If tState(l) Is Nothing Then
                        lRet = l
                        tState(l) = New StateObject
                        Exit For
                    End If
                Next
                If lRet = -999 Then
                    ReDim Preserve tState(0 To UBound(tState) + 1)
                    lRet = UBound(tState)
                    tState(lRet) = New StateObject
                End If

                '初期値
                ReDim tState(lRet).bRecv(0)
                tState(lRet).RemoteIP = remoteIP
                tState(lRet).RemotePort = remotePort
                tState(lRet).LocalIP = localIP
                tState(lRet).LocalPort = localPort
                tState(lRet).Sb = New StringBuilder

            End If



        Catch ex As Exception
            Dim msg As String = "GetStateIndexエラー:" & remoteIP & ":" & remotePort & " " & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.OtherErr)

        End Try
        Return lRet
    End Function

    ''' <summary>
    ''' 指定されたソケットのインデックスを返す。なければ-999を返す
    ''' </summary>
    ''' <param name="remoteIP">指定IP</param>
    ''' <param name="remotePort">指定PORT</param>
    ''' <returns>インデックス</returns>
    ''' <remarks>指定されたソケットのインデックスを返す。</remarks>
    Private Function FindStateIndex(ByVal remoteIP As String, ByVal remotePort As String) As Long
        Dim lRet As Long
        Dim l As Long
        lRet = -999
        Try


            '指定されたIP,Portのステートが存在しているか調べる
            For l = LBound(tState) To UBound(tState)
                If tState(l) Is Nothing Then
                Else
                    If tState(l).RemoteIP = remoteIP And tState(l).RemotePort = remotePort Then
                        '割り当て済みなのでこのステートを使う
                        lRet = l
                        Exit For
                    End If
                End If
            Next


        Catch ex As Exception
            RecvStatus = "ERROR"
            Dim msg As String = "FindStateIndexエラー:" & remoteIP & ":" & remotePort & " " & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.OtherErr)

        End Try
        Return lRet
    End Function


    ''' <summary>
    ''' Readコールバック
    ''' </summary>
    ''' <param name="ar">ソケット</param>
    ''' <remarks>Readコールバック</remarks>
    Private Sub ReadCallback(ByVal ar As IAsyncResult)
        Dim state As StateObject = CType(ar.AsyncState, StateObject)
        Try

            Dim content As String = String.Empty
            Dim l As Long
            Dim lix As Long
            Dim lSt As Long

            Dim handler As Socket = state.WorkSocket
            Dim IsReadEnd As Boolean = False

            Dim bytesRead As Integer
            Try
                bytesRead = handler.EndReceive(ar)
            Catch ex As ObjectDisposedException
                Dim msg As String = "[ReadCallback]" & state.RemoteIP & ":" & state.RemotePort & " はClose済みです"
                Call RaiseMessage(msg)
                Return
            Catch ex As Exception
                'RECEIVE中にクローズされた場合の処理
                Dim msg As String = LocalIP & ":" & LocalPort & " EndReceiveエラー:" & ex.Message
                Call RaiseMessage(msg)
                Call RaiseError(msg, SockErrors.RecvErr, state.RemoteIP, state.RemotePort)

                Return
            End Try


            If bytesRead > 0 Then
                Call RaiseRecvStart(state.RemoteIP, state.RemotePort, LocalIP, LocalPort)
                If UBound(state.bRecv) = 0 Then
                    lSt = 0
                    ReDim state.bRecv(0 To bytesRead - 1)
                Else
                    lSt = UBound(state.bRecv) + 1
                    ReDim Preserve state.bRecv(0 To UBound(state.bRecv) + bytesRead)
                End If

                l = 0
                For lix = lSt To UBound(state.bRecv)
                    state.bRecv(lix) = state.Buffer(l)
                    l += 1
                Next

                state.Sb.Append(Encoding.ASCII.GetString(state.Buffer, 0, bytesRead))

                content = state.Sb.ToString()
                Call RaiseMessage(state.RemoteIP & ":" & state.RemotePort & "->" & state.LocalIP & ":" & state.LocalPort & " " & content.Length.ToString & "バイト受信-> " & F.Byte2String(state.bRecv) & "(" & F.Byte2HexChar(state.bRecv) & ")")

                If content.Length > StateObject.BufferSize Then
                    'メッセージ表示
                    Call RaiseMessage(state.RemoteIP & ":" & state.RemotePort & "からバッファサイズを超えるデータを受け取りました。受信データを破棄します : " + F.Byte2HexChar(state.bRecv))
                    ReDim state.bRecv(0) '受信データの初期化
                    state.Sb = New StringBuilder

                Else
                    If EndCheckChar.Length = 0 Then
                        IsReadEnd = True
                    Else
                        '終了判定文字列を受信していれば完了
                        If content.Length >= EndCheckChar.Length Then
                            If F.Byte2HexChar(state.bRecv).EndsWith(F.Byte2HexChar(EndCheckChar)) Then
                                IsReadEnd = True
                            End If
                        End If

                    End If
                End If


            End If


            If IsReadEnd Then
                '受信完了イベント
                Call RaiseRecvEnd(state.RemoteIP, state.RemotePort, LocalIP, LocalPort, F.Byte2HexChar(state.bRecv))

                ReDim state.bRecv(0) '受信データの初期化
                state.Sb = New StringBuilder

            End If

            Try
                '受信待機
                handler.BeginReceive(state.Buffer, 0, StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReadCallback), state)
                ''Call BeginAccept() '再度ソケット受信受付する
            Catch ex As Exception
                Dim msg As String = LocalIP & ":" & LocalPort & " BeginReceiveエラー:" & ex.Message
                Call RaiseMessage(msg)
                Call RaiseError(msg, SockErrors.BeginRecvErr, state.RemoteIP, state.RemotePort)
            End Try


        Catch ex As Exception
            Dim msg As String = LocalIP & ":" & LocalPort & " ReadCallbackエラー:" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.RecvErr, state.RemoteIP, state.RemotePort)
        End Try
    End Sub 'ReadCallback


    ''' <summary>
    ''' ソケット送信登録
    ''' </summary>
    ''' <param name="SendParm">送信パラメータ</param>
    ''' <remarks>ソケット送信登録</remarks>
    Public Sub SockSend(ByVal SendParm As SendParmClass)
        SyncLock SendCmdList
            SendParm.SendRetryDo = 0
            SendParm.SendStatus = SockStatuses.SendWaiting
            SendCmdList.Add(SendParm)
        End SyncLock
    End Sub

    ''' <summary>
    ''' ソケット送信監視タイマー
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>ソケット送信監視タイマー</remarks>
    Private Sub SendCheckTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles SendCheckTimer.Elapsed
        If SendCheckTimerExecuting Then
            'タイマー処理実行中であればreturn
            Return
        End If

        Try
            SendCheckTimerExecuting = True

            Dim IsIdle As Boolean = True
            For sendParmIndex As Long = 0 To SendCmdList.Count - 1
                Dim sendparm As SendParmClass = SendCmdList(sendParmIndex)
                If sendparm.SendStatus = SockStatuses.Sended Then

                Else
                    IsIdle = False
                End If
            Next
            If IsIdle Then
                '送信待ち行列をチェックして送信済みがあれば削除する
                SendCmdList.RemoveAll(AddressOf IsRemoveSendParm)
            End If

            For sendParmIndex As Long = 0 To SendCmdList.Count - 1
                Dim sendparm As SendParmClass = SendCmdList(sendParmIndex)
                Select Case sendparm.SendStatus
                    Case SockStatuses.SendWaiting

                        '同一IP、Portで送信中のものがあるかチェックする
                        Dim IsSending As Boolean = False
                        For l As Long = 0 To SendCmdList.Count - 1
                            If SendCmdList(l).RemoteIP = sendparm.RemoteIP And _
                               SendCmdList(l).RemotePort = sendparm.RemotePort And _
                               SendCmdList(l).SendStatus = SockStatuses.Sending Then
                                IsSending = True '送信中のソケットあり
                                Exit For

                            End If
                        Next
                        If IsSending = False Then
                            '送信開始
                            sendparm.SendTimeLimit = DateTime.Now.AddMilliseconds(sendparm.SendTimeOut)
                            sendparm.SendStatus = SockStatuses.Sending
                            Call Send(sendparm.RemoteIP, sendparm.RemotePort, sendparm.sndCmd)
                        End If
                    Case SockStatuses.Sending
                        If LocalIP = "0.0.0.0" Then
                            Dim msg As String = String.Empty
                            msg = LocalIP & ":" & LocalPort & "->" & sendparm.RemoteIP & ":" & sendparm.RemotePort & " セッションが切れているのでリトライオーバー扱いとします "
                            Call RaiseMessage(msg)

                            sendparm.SendTimeLimit = DateTime.Now
                            sendparm.SendRetryDo = sendparm.SendRetry
                        End If


                        '送信タイムアウトチェック
                        If DateTime.Now >= sendparm.SendTimeLimit Then
                            If sendparm.SendRetryDo = sendparm.SendRetry Then
                                'リトライオーバー

                                Dim msg As String = LocalIP & ":" & LocalPort & "->" & sendparm.RemoteIP & ":" & sendparm.RemotePort & " 送信リトライオーバ"

                                Call RaiseError(msg, SockErrors.SendRetryOver, sendparm.RemoteIP, sendparm.RemotePort, sendparm.SendKEY)

                                sendparm.SendStatus = SockStatuses.Sended

                            Else
                                'リトライ
                                sendparm.SendRetryDo += 1
                                Dim msg As String = String.Empty
                                msg = LocalIP & ":" & LocalPort & "->" & sendparm.RemoteIP & ":" & sendparm.RemotePort & " 送信リトライ開始待ち..."
                                Call RaiseMessage(msg)

                                Call F.Wait(sendparm.SendRetryWait)

                                msg = LocalIP & ":" & LocalPort & "->" & sendparm.RemoteIP & ":" & sendparm.RemotePort & " へ送信開始 リトライ" & sendparm.SendRetryDo & "回目"
                                Call RaiseMessage(msg)
                                Call RaiseSendRetry(sendparm.RemoteIP, sendparm.RemotePort, LocalIP, LocalPort, sendparm.SendRetryDo, sendparm.SendKEY)

                                Call Send(sendparm.RemoteIP, sendparm.RemotePort, sendparm.sndCmd)
                            End If

                        End If

                End Select

            Next




        Catch ex As Exception
            Dim msg As String = "SendCheckTimer_Elapsedエラー:" & ex.Message
            Call RaiseMessage(msg)

        End Try
        SendCheckTimerExecuting = False


    End Sub

    ''' <summary>
    ''' 送信パラメータが削除可能か判定する
    ''' </summary>
    ''' <param name="sendParm">送信パラメータ</param>
    ''' <returns>True:削除可能 False:削除不可</returns>
    ''' <remarks>送信パラメータが削除可能か判定する</remarks>
    Private Function IsRemoveSendParm(ByVal sendParm As SendParmClass) As Boolean
        If sendParm.SendStatus = SockStatuses.Sended Then
            Return True
        Else
            Return False
        End If
    End Function


    ''' <summary>
    ''' ソケット送信
    ''' </summary>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <param name="sndData">送信内容</param>
    ''' <remarks>ソケット送信</remarks>
    Public Sub Send(ByVal remoteIP As String, ByVal remotePort As String, ByVal sndData As Byte())
        Try
            Dim l As Long = FindStateIndex(remoteIP, remotePort)
            Dim msg As String = String.Empty

            If l < 0 Then
                '指定されたIP、Portの接続が無ければ
                msg = remoteIP & ":" & remotePort & " のソケットが見つからないので送信できません"
                Call RaiseMessage(msg)
                Call RaiseError(msg, SockErrors.SendErr, remoteIP, remotePort)
                Return

            End If

            Dim sendingIndex As Long = GetSendingIndex(remoteIP, remotePort)
            If sendingIndex >= 0 Then
                msg = LocalIP & ":" & LocalPort & "->" & remoteIP & ":" & remotePort & " 送信開始 タイムアウト->" & Format(SendCmdList(sendingIndex).SendTimeLimit, "yyyy/MM/dd HH:mm:ss.ff") & " index:" & sendingIndex
            Else
                msg = LocalIP & ":" & LocalPort & "->" & remoteIP & ":" & remotePort & " 送信開始" & " index:" & sendingIndex
            End If
            Call RaiseMessage(msg)

            tState(l).WorkSocket.BeginSend(sndData, 0, sndData.Length, 0, New AsyncCallback(AddressOf SendCallback), tState(l).WorkSocket)
        Catch ex As Exception

            Dim msg As String = remoteIP & ":" & remotePort & " への送信エラー：" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.SendErr, remoteIP, remotePort)
        End Try
    End Sub 'Send


    ''' <summary>
    ''' ソケット送信コールバック
    ''' </summary>
    ''' <param name="ar">ソケット</param>
    ''' <remarks>ソケット送信コールバック</remarks>
    Private Sub SendCallback(ByVal ar As IAsyncResult)
        Dim handler As Socket
        Dim remoteIP As String
        Dim remotePORT As String

        Try
            handler = CType(ar.AsyncState, Socket)
            remoteIP = CType(handler.RemoteEndPoint, IPEndPoint).Address.ToString()
            remotePORT = CType(handler.RemoteEndPoint, IPEndPoint).Port.ToString()
        Catch ex As Exception
            Dim msg As String = "送信相手IP取得エラー(SendCallback)：" & ex.Message
            Call RaiseMessage(msg)
            Return
            'ここではremoteIP,remotePortは分からない可能性がある。Call RaiseError( msg, SockErrors.HealthCheckErr, remoteIP, remotePORT)
        End Try



        Try

            Dim bytesSent As Integer = handler.EndSend(ar)

            Dim localIP As String = CType(handler.LocalEndPoint, IPEndPoint).Address.ToString()
            Dim localPort As String = CType(handler.LocalEndPoint, IPEndPoint).Port.ToString()

            If localIP.Equals("0.0.0.0") Then
                'セッションが切れている
                Dim msg As String = localIP & ":" & localPort & " - " & remoteIP & ":" & remotePORT & " 間の接続が切れています"
                Call RaiseMessage(msg)
                Call RaiseError(msg, SockErrors.SendErr, remoteIP, remotePORT)
                Return

            End If



            '送信中の電文を探す
            Dim l As Long = GetSendingIndex(remoteIP, remotePORT)

            If l >= 0 Then


                Call RaiseSendEnd(SendCmdList(l).RemoteIP, SendCmdList(l).RemotePort, localIP, localPort, F.Byte2HexChar(SendCmdList(l).sndCmd), SendCmdList(l).SendKEY)

                'メッセージ表示
                Call RaiseMessage(localIP & ":" & localPort & "->" & remoteIP & ":" & remotePORT & " " & _
                      bytesSent & "バイト送信-> " & F.Byte2String(SendCmdList(l).sndCmd) & "(" & F.Byte2HexChar(SendCmdList(l).sndCmd) & ")")


                SendCmdList(l).SendStatus = SockStatuses.Sended
            End If


            Call RaiseMessage(remoteIP & ":" & remotePORT & " へ送信しました")

        Catch ex As Exception
            Dim msg As String = remoteIP & ":" & remotePORT & " への送信エラー(SendCallback)：" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.SendErr, remoteIP, remotePORT)
        End Try
    End Sub 'SendCallback


    ''' <summary>
    ''' 指定されたソケットの送信中状態のインデックスを取得する
    ''' </summary>
    ''' <param name="remoteIP">指定IP</param>
    ''' <param name="remotePort">指定PORT</param>
    ''' <returns>インデックス</returns>
    ''' <remarks>指定されたソケットの送信中状態のインデックスを取得する</remarks>
    Private Function GetSendingIndex(ByVal remoteIP, ByVal remotePort) As Long
        Dim result As Long = -999
        Dim retryCnt As Integer = 0

        Try
RETRY_POINT:


            For l As Long = 0 To SendCmdList.Count - 1
                If SendCmdList(l).RemoteIP = remoteIP And _
                   SendCmdList(l).RemotePort = remotePort And _
                   SendCmdList(l).SendStatus = SockStatuses.Sending Then

                    result = l

                    Exit For

                End If
            Next

        Catch ex As Exception
            If retryCnt < 3 Then
                '同時実行されると、処理中に配列が削除されることがあるためリトライする
                retryCnt += 1
                GoTo RETRY_POINT
            End If
            Throw ex
        End Try

        'Dim msg As String = remoteIP & ":" & remotePort & " index:" & result
        'Call RaiseMessage( msg)

        Return result

    End Function

    Public Sub New()
        EndCheckChar = New Byte() {}
    End Sub
End Class



''' <summary>
''' 送信ソケットクラス
''' </summary>
''' <remarks>送信ソケットクラス</remarks>
Public Class SendSockClass

    ''' <summary>
    ''' 受信終了判定文字列。この文字が来たら電文の終了とする。この文字を指定した場合、この文字が来るまで電文を連結して受信を続ける。指定していない場合は１回の通信単位＝電文単位となる
    ''' </summary>
    ''' <remarks>受信終了判定文字列。この文字が来たら電文の終了とする。この文字を指定した場合、この文字が来るまで電文を連結して受信を続ける。指定していない場合は１回の通信単位＝電文単位となる </remarks>
    Public Property EndCheckChar() As Byte()

    ''' <summary>
    ''' 相手IP
    ''' </summary>
    ''' <remarks>相手IP</remarks>
    Public Property RemoteIP() As String

    ''' <summary>
    ''' 相手Port
    ''' </summary>
    ''' <remarks>相手Port</remarks>
    Public Property RemotePort() As String

    ''' <summary>
    ''' 接続タイムアウトミリ秒。既定値1000
    ''' </summary>
    ''' <remarks>接続タイムアウトミリ秒</remarks>
    Public Property ConnectTimeOut() As Long = 1000

    ''' <summary>
    ''' 接続リトライを行う回数。既定値0
    ''' </summary>
    ''' <remarks>接続リトライを行う回数</remarks>
    Public Property ConnectRetry() As Long

    ''' <summary>
    ''' 接続リトライ間隔（ミリ秒）。既定値1000
    ''' </summary>
    ''' <remarks>接続リトライ間隔（ミリ秒）</remarks>
    Public Property ConnectRetryWait() As Long = 1000


    ''' <summary>
    ''' 接続処理のリミット（接続開始日時＋接続タイムアウト）
    ''' </summary>
    ''' <remarks>接続処理のリミット（接続開始日時＋接続タイムアウト）</remarks>
    Private ConnectTimeLimit As DateTime

    ''' <summary>
    ''' 接続リトライを行った回数
    ''' </summary>
    ''' <remarks>接続リトライを行った回数</remarks>
    Private ConnectRetryDo As Long = 0

    ''' <summary>
    ''' 接続状態監視タイマー
    ''' </summary>
    ''' <remarks>接続状態監視タイマー</remarks>
    Private WithEvents ConnectCheckTimer As New System.Timers.Timer

    ''' <summary>
    ''' 接続状態監視タイマー処理中フラグ
    ''' </summary>
    ''' <remarks>接続状態監視タイマー処理中フラグ</remarks>
    Private ConnectCheckTimerExecuting As Boolean = False

    ''' <summary>
    ''' 送信ステータス
    ''' </summary>
    ''' <remarks>送信ステータス</remarks>
    Private SendStatus As SockStatuses

    ''' <summary>
    ''' 受信ステータス
    ''' </summary>
    ''' <remarks>受信ステータス</remarks>
    Private RecvStatus As SockStatuses


    ''' <summary>
    ''' 送信用パラメータ
    ''' </summary>
    ''' <remarks>送信メソッドのパラメータ</remarks>
    Public Class SendParmClass
        '呼び出し側でセットするパラメータ

        ''' <summary>
        ''' ソケット識別キー（呼び出し側で自由に使ってよい）
        ''' </summary>
        ''' <remarks>ソケット識別キー（呼び出し側で自由に使ってよい）</remarks>
        Public Property SendKEY() As String

        ''' <summary>
        ''' 送信電文
        ''' </summary>
        ''' <remarks>送信電文</remarks>
        Public Property sndCmd() As Byte()

        ''' <summary>
        ''' 送信タイムアウトミリ秒。既定値1000
        ''' </summary>
        ''' <remarks>送信タイムアウトミリ秒</remarks>
        Public Property SendTimeOut() As Long = 1000

        ''' <summary>
        ''' 送信リトライを行う回数。既定値0
        ''' </summary>
        ''' <remarks>送信リトライを行う回数</remarks>
        Public Property SendRetry() As Long = 0

        ''' <summary>
        ''' 送信リトライ間隔（ミリ秒）
        ''' </summary>
        ''' <remarks>送信リトライ間隔（ミリ秒）</remarks>
        Public Property SendRetryWait() As Long

        ''' <summary>
        ''' 送信完了後にサーバーソケットからの返信待ちを行うか
        ''' </summary>
        ''' <remarks>送信完了後にサーバーソケットからの返信待ちを行うか</remarks>
        Public Property IsRecv() As Boolean

        ''' <summary>
        ''' サーバーソケットからの返信待ちタイムアウトミリ秒
        ''' </summary>
        ''' <remarks>サーバーソケットからの返信待ちタイムアウトミリ秒</remarks>
        Public Property RecvTimeOut() As Long

        '処理内で使うパラメータ

        ''' <summary>
        ''' 送信リトライを行った回数
        ''' </summary>
        ''' <remarks>送信リトライを行った回数</remarks>
        Public Property SendRetryDo() As Long = 0

        ''' <summary>
        ''' 送信処理のリミット（送信開始日時＋送信タイムアウト）
        ''' </summary>
        ''' <remarks>送信処理のリミット（送信開始日時＋送信タイムアウト）</remarks>
        Public Property SendTimeLimit() As DateTime

        ''' <summary>
        ''' 送信処理済みであればTrue
        ''' </summary>
        ''' <remarks>送信処理済みであればTrue</remarks>
        Public Property IsSendEnd() As Boolean

        ''' <summary>
        ''' 受信処理のリミット（受信開始日時＋送信タイムアウト）
        ''' </summary>
        ''' <remarks>受信処理のリミット（受信開始日時＋送信タイムアウト）</remarks>
        Public Property RecvTimeLimit() As DateTime

        ''' <summary>
        ''' 受信処理済みであればTrue
        ''' </summary>
        ''' <remarks>受信処理済みであればTrue</remarks>
        Public Property IsRecvEnd() As Boolean

    End Class

    ''' <summary>
    ''' 送信電文の待ち行列
    ''' </summary>
    ''' <remarks>送信電文の待ち行列</remarks>
    Private SendCmdList As New System.Collections.Generic.List(Of SendParmClass)

    ''' <summary>
    ''' 現在送信処理中の待ち行列
    ''' </summary>
    ''' <remarks>現在送信処理中の待ち行列</remarks>
    Private SendingIndex As Long

    ''' <summary>
    ''' 送信状態監視タイマー
    ''' </summary>
    ''' <remarks>送信状態監視タイマー</remarks>
    Private WithEvents SendCheckTimer As New System.Timers.Timer

    ''' <summary>
    ''' 送信状態監視タイマー処理中フラグ
    ''' </summary>
    ''' <remarks>送信状態監視タイマー処理中フラグ</remarks>
    Private SendCheckTimerExecuting As Boolean = False

    ''' <summary>
    ''' 受信状態監視タイマー
    ''' </summary>
    ''' <remarks>受信状態監視タイマー</remarks>
    Private WithEvents RecvCheckTimer As New System.Timers.Timer

    ''' <summary>
    ''' 受信状態監視タイマー処理中フラグ
    ''' </summary>
    ''' <remarks>受信状態監視タイマー処理中フラグ</remarks>
    Private RecvCheckTimerExecuting As Boolean = False

    ''' <summary>
    ''' タグ（呼び出し側で自由に使ってよい）
    ''' </summary>
    ''' <remarks>タグ（呼び出し側で自由に使ってよい）</remarks>
    Public TAG As String



    ''' <summary>
    ''' メッセージ出力イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>メッセージ出力イベント</remarks>
    Public Event OnMessage(ByVal sender As Object, ByVal e As SockMessageEventArgs)


    ''' <summary>
    ''' 接続完了イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>接続完了イベント</remarks>
    Public Event OnConnectEnd(ByVal sender As Object, ByVal e As SockConnectEndEventArgs)

    ''' <summary>
    ''' 接続リトライ開始イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>接続リトライ開始イベント</remarks>
    Public Event OnConnectRetry(ByVal sender As Object, ByVal e As SockConnectRetryEventArgs)

    ''' <summary>
    ''' 送信リトライ開始イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>送信リトライ開始イベント</remarks>
    Public Event OnSendRetry(ByVal sender As Object, ByVal e As SockSendRetryEventArgs)

    ''' <summary>
    ''' 送信完了イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>送信完了イベント</remarks>
    Public Event OnSendEnd(ByVal sender As Object, ByVal e As SockSendEndEventArgs)

    ''' <summary>
    ''' 受信完了イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>受信完了イベント</remarks>
    Public Event OnRecvEnd(ByVal sender As Object, ByVal e As SockReceiveEndEventArgs)

    ''' <summary>
    ''' エラー発生イベント
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>エラー発生イベント</remarks>
    Public Event OnError(ByVal sender As Object, ByVal e As SockErrorEventArgs)


    ''' <summary>
    ''' IPアドレス毎に持つ
    ''' </summary>
    ''' <remarks>IPアドレス毎に持つ</remarks>
    Private tState(0) As StateObject

    ''' <summary>
    ''' クライアントソケット
    ''' </summary>
    ''' <remarks>クライアントソケット</remarks>
    Private client As Socket

    Public Sub New()
        EndCheckChar = New Byte() {}
    End Sub

    ''' <summary>
    ''' ソケット生存確認
    ''' </summary>
    ''' <returns>True:生きている False:死んでいる</returns>
    ''' <remarks>ソケット生存確認</remarks>
    <Obsolete("このメソッドは相手（サーバーソケット）から切断された場合は機能しません（Trueのまま変わらない）ので、使い方に注意して下さい")> _
    Public Function IsAlive()
        Dim result As Boolean = False

        Try
            If client.Poll(1000, Sockets.SelectMode.SelectRead) = True Then
                result = False
            Else
                result = True
            End If
        Catch

        End Try

        Return result

        'Dim result As Boolean = True

        'Try
        '    Dim localIP As String = CType(client.LocalEndPoint, IPEndPoint).Address.ToString()
        '    Dim localPort As String = CType(client.LocalEndPoint, IPEndPoint).Port.ToString()

        '    If localIP.Equals("0.0.0.0") Then
        '        result = False
        '    End If

        'Catch ex As Exception
        '    result = False

        'End Try

        'Return result

    End Function


    ''' <summary>
    ''' メッセージイベントを発生させる
    ''' </summary>
    ''' <param name="msg">メッセージ</param>
    ''' <remarks>メッセージイベントを発生させる</remarks>
    Private Sub RaiseMessage(ByVal msg As String)
        Dim e As New SockMessageEventArgs
        e.Msg = msg
        RaiseEvent OnMessage(Me, e)
    End Sub

    ''' <summary>
    ''' エラーイベントを発生させる
    ''' </summary>
    ''' <param name="msg">メッセージ</param>
    ''' <param name="sockError">エラータイプ</param>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <param name="sendKey">送信KEY</param>
    ''' <remarks>エラーイベントを発生させる</remarks>
    Private Sub RaiseError(ByVal msg As String, ByVal sockError As SockErrors, _
                           Optional ByVal remoteIP As String = "", _
                           Optional ByVal remotePort As String = "", _
                           Optional ByVal sendKey As String = "")
        Dim e As New SockErrorEventArgs
        e.Msg = msg
        e.ErrorType = sockError
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort
        e.SendKey = sendKey
        RaiseEvent OnError(Me, e)
    End Sub

    ''' <summary>
    ''' 接続完了イベントを発生させる
    ''' </summary>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <param name="localIP">自IP</param>
    ''' <param name="localPort">自PORT</param>
    ''' <remarks>接続完了イベントを発生させる</remarks>
    Private Sub RaiseConnectEnd(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String)
        Dim e As New SockConnectEndEventArgs
        e.LocalIP = localIP
        e.LocalPort = localPort
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort

        RaiseEvent OnConnectEnd(Me, e)
    End Sub

    ''' <summary>
    ''' 接続リトライイベントを発生させる
    ''' </summary>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <param name="RetryDoCnt">リトライ回数</param>
    ''' <remarks>接続リトライイベントを発生させる</remarks>
    Private Sub RaiseConnectRetry(ByVal remoteIP As String, ByVal remotePort As String, ByVal RetryDoCnt As Long)
        Dim e As New SockConnectRetryEventArgs
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort
        e.RetryDo = RetryDoCnt

        RaiseEvent OnConnectRetry(Me, e)
    End Sub


    ''' <summary>
    ''' 受信完了イベントを発生させる
    ''' </summary>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <param name="localIP">自IP</param>
    ''' <param name="localPort">自PORT</param>
    ''' <param name="recvChar">受信内容</param>
    ''' <remarks>受信完了イベントを発生させる</remarks>
    Private Sub RaiseRecvEnd(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String, ByVal recvChar As String, ByVal sendKey As String)
        Dim e As New SockReceiveEndEventArgs
        e.LocalIP = localIP
        e.LocalPort = localPort
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort
        e.RecvChar = recvChar
        e.SendKey = sendKey

        RaiseEvent OnRecvEnd(Me, e)
    End Sub

    ''' <summary>
    ''' 送信リトライイベントを発生させる
    ''' </summary>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <param name="localIP">自IP</param>
    ''' <param name="localPort">自PORT</param>
    ''' <param name="RetryDoCnt">リトライ回数</param>
    ''' <param name="sendKey">送信KEY</param>
    ''' <remarks>送信リトライイベントを発生させる</remarks>
    Private Sub RaiseSendRetry(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String, ByVal RetryDoCnt As Long, ByVal sendKey As String)
        Dim e As New SockSendRetryEventArgs
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort
        e.LocalIP = localIP
        e.LocalPort = localPort
        e.RetryDo = RetryDoCnt
        e.SendKey = sendKey

        RaiseEvent OnSendRetry(Me, e)
    End Sub

    ''' <summary>
    ''' 送信完了イベントを発生させる
    ''' </summary>
    ''' <param name="remoteIP">相手IP</param>
    ''' <param name="remotePort">相手PORT</param>
    ''' <param name="localIP">自IP</param>
    ''' <param name="localPort">自PORT</param>
    ''' <param name="sendChar">送信内容</param>
    ''' <param name="sendkey">送信KEY</param>
    ''' <remarks>送信完了イベントを発生させる</remarks>
    Private Sub RaiseSendEnd(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String, ByVal sendChar As String, ByVal sendkey As String)
        Dim e As New SockSendEndEventArgs
        e.LocalIP = localIP
        e.LocalPort = localPort
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort
        e.SendChar = sendChar
        e.SendKey = sendkey

        RaiseEvent OnSendEnd(Me, e)



    End Sub


    ''' <summary>
    ''' ソケット接続開始処理
    ''' </summary>
    ''' <remarks>ソケット接続開始処理</remarks>
    Public Sub SockConnect()

        Call CloseSock()
        ConnectRetryDo = 0
        Call Connect()
        ConnectCheckTimer.Interval = 10
        ConnectCheckTimer.Start()

    End Sub

    ''' <summary>
    ''' 接続チェックタイマー
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>接続チェックタイマー</remarks>
    Private Sub ConnectCheckTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles ConnectCheckTimer.Elapsed
        If ConnectCheckTimerExecuting Then
            'タイマー処理実行中であればreturn
            Return
        End If

        ConnectCheckTimerExecuting = True
        Select Case SendStatus
            Case SockStatuses.Connecting
                '接続タイムアウトチェック
                If DateTime.Now >= ConnectTimeLimit Then
                    If ConnectRetryDo = ConnectRetry Then
                        'リトライオーバー
                        Dim msg As String = RemoteIP & ":" & RemotePort & " へ接続リトライオーバー"
                        Call RaiseError(msg, SockErrors.ConnectRetryOver)
                        ConnectCheckTimer.Stop()
                    Else
                        'リトライ
                        ConnectRetryDo += 1
                        Dim msg As String = String.Empty
                        msg = RemoteIP & ":" & RemotePort & " へ接続リトライ開始待ち..."
                        Call RaiseMessage(msg)

                        Call F.Wait(ConnectRetryWait)

                        msg = RemoteIP & ":" & RemotePort & " へ接続開始 リトライ" & ConnectRetryDo & "回目"
                        Call RaiseMessage(msg)
                        Call RaiseConnectRetry(RemoteIP, RemotePort, ConnectRetryDo)

                        Call Connect() '接続処理を呼ぶ
                    End If

                End If
            Case SockStatuses.Connected
                ConnectCheckTimer.Stop()
                '送信監視タイマーを起動
                SendCheckTimer.Interval = 10
                SendCheckTimer.Start()

        End Select

        ConnectCheckTimerExecuting = False



    End Sub



    ''' <summary>
    ''' ソケット接続
    ''' </summary>
    ''' <remarks>ソケット接続</remarks>
    Private Sub Connect()
        Try
            SendStatus = SockStatuses.Connecting
            ConnectTimeLimit = DateTime.Now.AddMilliseconds(ConnectTimeOut)


            Dim cipAddress As IPAddress = IPAddress.Parse(RemoteIP)
            Dim remoteEP As New IPEndPoint(cipAddress, CInt(RemotePort))

            ' Create a TCP/IP socket.
            client = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)


            Try
                Dim msg As String = RemoteIP & ":" & RemotePort & " へ接続開始 タイムアウト->" & Format(ConnectTimeLimit, "yyyy/MM/dd HH:mm:ss.ff")
                Call RaiseMessage(msg)
                client.BeginConnect(remoteEP, New AsyncCallback(AddressOf ConnectCallback), client)

            Catch e As Exception
                Dim msg As String = RemoteIP & ":" & RemotePort & " Connectエラー:" & e.Message
                Call RaiseMessage(msg)
                If ConnectCheckTimer.Enabled Then '接続監視タイマーが動いている間はエラーイベントを起こす
                    Call RaiseError(msg, SockErrors.ConnectErr)
                End If

                Exit Sub
            End Try

        Catch ex As Exception
            Dim msg As String = RemoteIP & ":" & RemotePort & " ソケット接続できません:" & ex.Message
            Call RaiseMessage(msg)
            If ConnectCheckTimer.Enabled Then '接続監視タイマーが動いている間はエラーイベントを起こす
                Call RaiseError(msg, SockErrors.ConnectErr)
            End If

        End Try

    End Sub

    ''' <summary>
    ''' ソケット接続コールバック
    ''' </summary>
    ''' <param name="ar">ソケット</param>
    ''' <remarks>ソケット接続コールバック</remarks>
    Private Sub ConnectCallback(ByVal ar As IAsyncResult)
        Try
            ' Retrieve the socket from the state object.
            Dim client As Socket = CType(ar.AsyncState, Socket)

            Try
                ' Complete the connection.
                client.EndConnect(ar)
            Catch ex As Exception
                'Connect中にクローズされた場合の処理
                Dim msg As String = RemoteIP & ":" & RemotePort & " ConnectCallbackエラー:" & ex.Message
                Call RaiseMessage(msg)
                If ConnectCheckTimer.Enabled Then '接続監視タイマーが動いている間はエラーイベントを起こす
                    Call RaiseError(msg, SockErrors.ConnectErr)
                End If
                Return
            End Try

            SendStatus = SockStatuses.Connected

            Dim localIP As String = CType(client.LocalEndPoint, IPEndPoint).Address.ToString()
            Dim localPort As String = CType(client.LocalEndPoint, IPEndPoint).Port.ToString()

            Call RaiseConnectEnd(RemoteIP, RemotePort, localIP, localPort)

            '受信開始
            Call RaiseMessage("RECEIVE開始:" & localIP & "(" & localPort & ")")
            Dim l = GetStateIndex(RemoteIP, RemotePort, localIP, localPort)
            With tState(l)
                .WorkSocket = client
            End With
            client.BeginReceive(tState(l).Buffer, 0, StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReadCallback), tState(l))


        Catch ex As Exception
            Dim msg As String = RemoteIP & ":" & RemotePort & " ConnectCallbackエラー:" & ex.Message
            Call RaiseMessage(msg)
            If ConnectCheckTimer.Enabled Then '接続監視タイマーが動いている間はエラーイベントを起こす
                Call RaiseError(msg, SockErrors.ConnectErr)
            End If

        End Try

    End Sub 'ConnectCallback


    ''' <summary>
    ''' ソケット切断処理
    ''' </summary>
    ''' <remarks>ソケット切断処理</remarks>
    Public Sub CloseSock()
        Try
            Try
                client.Shutdown(SocketShutdown.Both)
            Catch ex As Exception

            End Try

            ConnectCheckTimer.Stop()
            SendCheckTimer.Stop()
            RecvCheckTimer.Stop()
            SendCmdList.Clear()
            RecvStatus = SockStatuses.Received
            SendStatus = SockStatuses.Sended

            Dim localIP As String = CType(client.LocalEndPoint, IPEndPoint).Address.ToString()
            Dim localPort As String = CType(client.LocalEndPoint, IPEndPoint).Port.ToString()

            client.Close()

            Dim msg As String = localIP & ":" & localPort & " Close"
            Call RaiseMessage(msg)

        Catch ex As Exception

        End Try
    End Sub


    ''' <summary>
    ''' 送信開始処理
    ''' </summary>
    ''' <param name="SendParm">送信パラメータ</param>
    ''' <remarks>送信開始処理</remarks>
    Public Sub SockSend(ByVal SendParm As SendParmClass)

        SyncLock SendCmdList

            SendParm.SendRetryDo = 0
            SendParm.IsSendEnd = False
            SendCmdList.Add(SendParm)

            Dim msg As String = String.Empty
            Dim localIP As String = "0.0.0.0"
            Dim localPort As String = "????"
            Try
                localIP = CType(client.LocalEndPoint, IPEndPoint).Address.ToString()
                localPort = CType(client.LocalEndPoint, IPEndPoint).Port.ToString()
            Catch ex As Exception
                msg = "(SockSend)LocalIP取得エラー:" & ex.Message
                Call RaiseMessage(msg)
            End Try

            msg = "送信待ちに追加しました:" & localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & _
             " Key:" & SendParm.SendKEY & " 送信内容:" & F.Byte2String(SendParm.sndCmd)
            Call RaiseMessage(msg)

        End SyncLock
    End Sub

    ''' <summary>
    ''' 送信監視タイマー処理
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>送信監視タイマー処理</remarks>
    Private Sub SendCheckTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles SendCheckTimer.Elapsed
        If SendCheckTimerExecuting Then
            'タイマー処理実行中であればreturn
            Return
        End If

        If RecvStatus = SockStatuses.Receiving Then
            '受信処理中であればreturn
            Return
        End If


        Try
            SendCheckTimerExecuting = True


            Dim localIP As String = "0.0.0.0"
            Dim localPort As String = "????"
            Try
                localIP = CType(client.LocalEndPoint, IPEndPoint).Address.ToString()
                localPort = CType(client.LocalEndPoint, IPEndPoint).Port.ToString()
            Catch ex As Exception
                Dim msg As String = "(SendCheckTimer_Elapsed)LocalIP取得エラー:" & ex.Message
                Call RaiseMessage(msg)
            End Try

            Select Case SendStatus
                Case SockStatuses.Sended, SockStatuses.Connected
                    '送信待ち行列をチェックして送信済みがあれば削除する
                    SendCmdList.RemoveAll(AddressOf IsRemoveSendParm)

                    '送信待ち行列をチェックして待ちがあれば送信開始する
                    For l As Integer = 0 To SendCmdList.Count - 1
                        If SendCmdList(l).IsSendEnd = False Then

                            SendingIndex = l

                            Dim msg As String = "送信待ちあり " & localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " SendingIndex:" & SendingIndex
                            Call RaiseMessage(msg)

                            Call Send()
                            Exit For
                        End If
                    Next

                Case SockStatuses.Sending
                    If localIP = "0.0.0.0" Then
                        Dim msg As String = String.Empty
                        msg = localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " セッションが切れているのでリトライオーバー扱いとします " & " SendingIndex:" & SendingIndex
                        Call RaiseMessage(msg)

                        SendCmdList(SendingIndex).SendTimeLimit = DateTime.Now
                        SendCmdList(SendingIndex).SendRetryDo = SendCmdList(SendingIndex).SendRetry

                    End If


                    '送信タイムアウトチェック
                    If DateTime.Now >= SendCmdList(SendingIndex).SendTimeLimit Then
                        If SendCmdList(SendingIndex).SendRetryDo = SendCmdList(SendingIndex).SendRetry Then
                            'リトライオーバー

                            Dim msg As String = localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " 送信リトライオーバ" & " SendingIndex:" & SendingIndex

                            Call RaiseError(msg, SockErrors.SendRetryOver, RemoteIP, RemotePort, SendCmdList(SendingIndex).SendKEY)

                            SendCheckTimer.Stop()

                            Call SockConnect() '接続処理を呼ぶ

                        Else
                            'リトライ
                            SendCmdList(SendingIndex).SendRetryDo += 1
                            Dim msg As String = String.Empty
                            msg = localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " 送信リトライ開始待ち..." & " SendingIndex:" & SendingIndex
                            Call RaiseMessage(msg)

                            Call F.Wait(SendCmdList(SendingIndex).SendRetryWait)

                            msg = localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " へ送信開始 リトライ" & SendCmdList(SendingIndex).SendRetryDo & "回目" & " SendingIndex:" & SendingIndex
                            Call RaiseMessage(msg)
                            Call RaiseSendRetry(RemoteIP, RemotePort, localIP, localPort, SendCmdList(SendingIndex).SendRetryDo, SendCmdList(SendingIndex).SendKEY)

                            Call Send() '送信処理を呼ぶ
                        End If

                    End If

            End Select


        Catch ex As Exception
            Dim msg As String = RemoteIP & ":" & RemotePort & " SendCheckTimer_Elapsedエラー:" & ex.Message
            Call RaiseMessage(msg)

        End Try
        SendCheckTimerExecuting = False


    End Sub

    ''' <summary>
    ''' 送信パラメータを削除して良いかチェックする
    ''' </summary>
    ''' <param name="sendParm">送信パラメータ</param>
    ''' <returns>True:削除可 False:削除不可</returns>
    ''' <remarks>送信パラメータを削除して良いかチェックする</remarks>
    Private Function IsRemoveSendParm(ByVal sendParm As SendParmClass) As Boolean
        If sendParm.IsRecvEnd Then
            '受信処理まで完了していれば消して良い
            Return True
        Else
            If sendParm.IsRecv = False And sendParm.IsSendEnd Then
                '受信処理不要で送信完了していれば消して良い
                Return True
            End If

            If sendParm.IsSendEnd And sendParm.IsRecv And DateTime.Now >= sendParm.RecvTimeLimit Then
                '受信待ちタイムアウトになっていれば消して良い
                Return True
            End If


            Return False
        End If
    End Function

    ''' <summary>
    ''' ソケット送信
    ''' </summary>
    ''' <remarks>ソケット送信</remarks>
    Private Sub Send()
        Try


            SendStatus = SockStatuses.Sending
            SendCmdList(SendingIndex).SendTimeLimit = DateTime.Now.AddMilliseconds(SendCmdList(SendingIndex).SendTimeOut)


            Dim localIP As String = CType(client.LocalEndPoint, IPEndPoint).Address.ToString()
            Dim localPort As String = CType(client.LocalEndPoint, IPEndPoint).Port.ToString()

            Dim msg As String = localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " 送信開始 タイムアウト->" & Format(SendCmdList(SendingIndex).SendTimeLimit, "yyyy/MM/dd HH:mm:ss.ff") & " SendingIndex:" & SendingIndex
            Call RaiseMessage(msg)

            client.BeginSend(SendCmdList(SendingIndex).sndCmd, 0, SendCmdList(SendingIndex).sndCmd.Length, 0, New AsyncCallback(AddressOf SendCallback), client)

        Catch ex As Exception
            Dim msg As String = RemoteIP & ":" & RemotePort & " BeginSendエラー:" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.SendErr, RemoteIP, RemotePort)

        End Try
    End Sub 'Send

    ''' <summary>
    ''' ソケット送信コールバック
    ''' </summary>
    ''' <param name="ar">ソケット</param>
    ''' <remarks>ソケット送信コールバック</remarks>
    Private Sub SendCallback(ByVal ar As IAsyncResult)
        Dim pos As Integer = 0

        Try
            pos = 100


            pos = 110
            '受信待ちの必要があればステータス変更を行う（このタイミングで行わないとステータス変更前にReceiveされることがある）
            If SendCmdList(SendingIndex).IsRecv Then
                pos = 120
                'ステータス変更
                RecvStatus = SockStatuses.Receiving
            End If

            Dim client As Socket = CType(ar.AsyncState, Socket)
            Dim bytesSent As Integer = client.EndSend(ar)

            pos = 200

            Dim localIP As String = CType(client.LocalEndPoint, IPEndPoint).Address.ToString()
            Dim localPort As String = CType(client.LocalEndPoint, IPEndPoint).Port.ToString()

            pos = 300

            If localIP.Equals("0.0.0.0") Then
                'セッションが切れている
                Dim msg As String = localIP & ":" & localPort & " - " & RemoteIP & ":" & RemotePort & " 間の接続が切れています"

                RecvStatus = SockStatuses.Received '上で受信中にしているので、受信完了に戻しておく

                Call RaiseMessage(msg)
                If SendingIndex <= SendCmdList.Count - 1 Then
                    Call RaiseError(msg, SockErrors.SendErr, RemoteIP, RemotePort, SendCmdList(SendingIndex).SendKEY)
                End If
                Return

            End If

            pos = 400
            Call RaiseSendEnd(RemoteIP, RemotePort, localIP, localPort, F.Byte2HexChar(SendCmdList(SendingIndex).sndCmd), SendCmdList(SendingIndex).SendKEY)

            'メッセージ表示
            Call RaiseMessage(localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " " & _
                              bytesSent & "バイト送信-> " & F.Byte2String(SendCmdList(SendingIndex).sndCmd) & "(" & F.Byte2HexChar(SendCmdList(SendingIndex).sndCmd) & ")" & " SendingIndex:" & SendingIndex)


            pos = 500
            '受信待ちの必要があれば受信待ちを行う
            If SendCmdList(SendingIndex).IsRecv Then
                '受信監視タイマーを起動
                SendCmdList(SendingIndex).RecvTimeLimit = DateTime.Now.AddMilliseconds(SendCmdList(SendingIndex).RecvTimeOut)
                Dim msg As String = RemoteIP & ":" & RemotePort & "->" & localIP & ":" & localPort & " 戻り待ち開始 タイムアウト->" & Format(SendCmdList(SendingIndex).RecvTimeLimit, "yyyy/MM/dd HH:mm:ss.ff") & " SendingIndex:" & SendingIndex
                Call RaiseMessage(msg)

                pos = 600

                'このタイミングでステータス変更はだめ。RecvStatus = SockStatuses.Receiving
                RecvCheckTimer.Interval = 10
                RecvCheckTimer.Start()

            End If

            pos = 700


            '送信完了状態にする
            SendCmdList(SendingIndex).IsSendEnd = True
            SendStatus = SockStatuses.Sended

            pos = 800
        Catch ex As Exception
            Dim msg As String = RemoteIP & ":" & RemotePort & " SendCallbackエラー:" & ex.Message & " pos=" & pos
            Call RaiseMessage(msg)

            If SendingIndex <= SendCmdList.Count - 1 Then
                Call RaiseError(msg, SockErrors.SendErr, RemoteIP, RemotePort, SendCmdList(SendingIndex).SendKEY)
            Else
                'Call RaiseError( msg, SockErrors.SendErr, RemoteIP, RemotePort, "Key不明")
                msg = "SendCallbackエラー すでにソケットがクローズされています"
                Call RaiseMessage(msg)
                Return
            End If

        End Try

    End Sub 'SendCallback

    ''' <summary>
    ''' 指定されたIPのステートのインデックスを返す。なければ新たに作成する
    ''' </summary>
    ''' <param name="remoteIP">指定IP</param>
    ''' <param name="remotePort">指定PORT</param>
    ''' <param name="localIP">自IP</param>
    ''' <param name="localPort">自PORT</param>
    ''' <returns>インデックス</returns>
    ''' <remarks>指定されたソケットのインデックスを返す</remarks>
    Private Function GetStateIndex(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String) As Long
        Dim lRet As Long
        Dim l As Long
        lRet = -999


        Try
            '指定されたIPのステートが存在しているか調べる
            For l = LBound(tState) To UBound(tState)
                If tState(l) Is Nothing Then
                Else
                    If tState(l).RemoteIP = remoteIP Then
                        '割り当て済みなのでこのステートを使う
                        tState(l).RemotePort = remotePort
                        tState(l).LocalIP = localIP
                        tState(l).LocalPort = localPort

                        lRet = l
                        Exit For
                    End If
                End If
            Next

            If lRet = -999 Then
                '指定されたIPのステートは存在していなかったので空きステートを探す
                For l = LBound(tState) To UBound(tState)
                    If tState(l) Is Nothing Then
                        lRet = l
                        tState(l) = New StateObject
                        Exit For
                    End If
                Next
                If lRet = -999 Then
                    ReDim Preserve tState(0 To UBound(tState) + 1)
                    lRet = UBound(tState)
                    tState(lRet) = New StateObject
                End If

                '初期値
                ReDim tState(lRet).bRecv(0)
                tState(lRet).RemoteIP = remoteIP
                tState(lRet).RemotePort = remotePort
                tState(lRet).LocalIP = localIP
                tState(lRet).LocalPort = localPort
                tState(lRet).Sb = New StringBuilder

            End If

        Catch ex As Exception
            Dim msg As String = "GetStateIndexエラー:" & remoteIP & ":" & remotePort & " " & ex.Message
            Call RaiseMessage(msg)

        End Try


        Return lRet
    End Function


    ''' <summary>
    ''' 受信コールバック
    ''' </summary>
    ''' <param name="ar">ソケット</param>
    ''' <remarks>受信コールバック</remarks>
    Private Sub ReadCallback(ByVal ar As IAsyncResult)
        Dim state As StateObject = CType(ar.AsyncState, StateObject)
        Dim pos As Integer = 0

        Try
            Dim content As String = String.Empty
            Dim l As Long
            Dim lix As Long
            Dim lSt As Long

            pos = 100

            Dim handler As Socket = state.WorkSocket

            Dim IsReadEnd As Boolean = False

            Dim bytesRead As Integer
            Try
                bytesRead = handler.EndReceive(ar)
            Catch ex As ObjectDisposedException
                Dim msg As String = "[ReadCallback]" & state.LocalIP & ":" & state.LocalPort & " はClose済みです"
                Call RaiseMessage(msg)
                Return
            Catch ex As Exception
                'RECEIVE中にクローズされた場合の処理
                Dim msg As String = state.LocalIP & ":" & state.LocalPort & " EndReceiveエラー:" & ex.Message
                Call RaiseMessage(msg)
                Return
            End Try


            pos = 200

            If bytesRead = 0 Then
                'サーバー側から切断された
                Call CloseSock()
                Return
            Else
                If UBound(state.bRecv) = 0 Then
                    lSt = 0
                    ReDim state.bRecv(0 To bytesRead - 1)
                Else
                    lSt = UBound(state.bRecv) + 1
                    ReDim Preserve state.bRecv(0 To UBound(state.bRecv) + bytesRead)
                End If

                l = 0
                For lix = lSt To UBound(state.bRecv)
                    state.bRecv(lix) = state.Buffer(l)
                    l += 1
                Next


                state.Sb.Append(Encoding.ASCII.GetString(state.Buffer, 0, bytesRead))
                content = state.Sb.ToString()
                'メッセージ表示
                Call RaiseMessage(state.RemoteIP & ":" & state.RemotePort & "->" & state.LocalIP & ":" & state.LocalPort & " " & content.Length.ToString & "バイト受信-> " & F.Byte2String(state.bRecv) & "(" & F.Byte2HexChar(state.bRecv) & ")" & " SendingIndex:" & SendingIndex)

            End If

            pos = 300

            If content.Length > StateObject.BufferSize Then
                'メッセージ表示
                Call RaiseMessage(state.RemoteIP & "からバッファサイズを超えるデータを受け取りました。受信データを破棄します : " + F.Byte2HexChar(state.bRecv) & " SendingIndex:" & SendingIndex)
                ReDim state.bRecv(0) '受信データの初期化
                state.Sb = New StringBuilder

            Else
                If EndCheckChar.Length = 0 Then
                    IsReadEnd = True
                Else
                    '終了判定文字列を受信していれば完了
                    If content.Length >= EndCheckChar.Length Then
                        If F.Byte2HexChar(state.bRecv).EndsWith(F.Byte2HexChar(EndCheckChar)) Then
                            IsReadEnd = True
                        End If
                    End If

                End If


            End If

            pos = 400


            If IsReadEnd Then
                Dim recvCmds = New String() {F.Byte2String(state.bRecv)}
                Dim kugiri As String = String.Empty

                If EndCheckChar.Length <> 0 Then
                    '複数のメッセージがくっついてくることがあるのでスプリットする
                    kugiri = F.Byte2String(EndCheckChar)

                    recvCmds = Split(F.Byte2String(state.bRecv), kugiri)

                End If

                Dim isRecvEnd As Boolean = False

                pos = 410

                For cmdix As Long = LBound(recvCmds) To UBound(recvCmds)
                    If recvCmds(cmdix).ToString <> String.Empty Then
                        If RecvStatus = SockStatuses.Receiving Then
                            pos = 430
                            SendCmdList(SendingIndex).IsRecvEnd = True
                            pos = 431
                            RecvStatus = SockStatuses.Received

                            pos = 432
                            '受信完了イベント
                            Call RaiseRecvEnd(state.RemoteIP, state.RemotePort, state.LocalIP, state.LocalPort, F.Byte2HexChar(F.String2Byte(recvCmds(cmdix).ToString & kugiri)), SendCmdList(SendingIndex).SendKEY)
                            pos = 433

                            isRecvEnd = True
                            Exit For
                        Else
                            pos = 440
                            '受信完了イベント
                            Call RaiseRecvEnd(state.RemoteIP, state.RemotePort, state.LocalIP, state.LocalPort, F.Byte2HexChar(F.String2Byte(recvCmds(cmdix).ToString & kugiri)), "Key不明" & cmdix)

                        End If

                    End If
                Next

                'If isRecvEnd Then
                '    '受信完了
                '    SendCmdList(SendingIndex).IsRecvEnd = True
                '    RecvStatus = SockStatuses.Received
                'End If


                pos = 450
                ReDim state.bRecv(0) '受信データの初期化
                state.Sb = New StringBuilder

            End If

            pos = 500


            Try
                handler.BeginReceive(state.Buffer, 0, StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReadCallback), state)
            Catch ex As Exception
                Dim msg As String = RemoteIP & ":" & RemotePort & " (再)BeginReceiveエラー:" & ex.Message
                Call RaiseMessage(msg)
            End Try
        Catch ex As Exception
            Dim msg As String = state.LocalIP & ":" & state.LocalPort & " ReadCallbackエラー:" & ex.Message & " pos:" & pos
            Call RaiseMessage(msg)

        End Try
    End Sub 'ReadCallback


    ''' <summary>
    ''' 受信監視タイマー処理
    ''' </summary>
    ''' <param name="sender">イベント発生元</param>
    ''' <param name="e">イベント引数</param>
    ''' <remarks>受信監視タイマー処理</remarks>
    Private Sub RecvCheckTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles RecvCheckTimer.Elapsed
        If RecvCheckTimerExecuting Then
            'タイマー処理実行中であればreturn
            Return
        End If


        Try
            RecvCheckTimerExecuting = True


            Dim localIP As String = "0.0.0.0"
            Dim localPort As String = "????"
            Try
                localIP = CType(client.LocalEndPoint, IPEndPoint).Address.ToString()
                localPort = CType(client.LocalEndPoint, IPEndPoint).Port.ToString()
            Catch ex As Exception
                Dim msg As String = RemoteIP & ":" & RemotePort & " LocalIP取得エラー:" & ex.Message
                Call RaiseMessage(msg)
            End Try

            Select Case RecvStatus
                Case SockStatuses.Received
                    Dim msg As String = localIP & ":" & localPort & " -> " & RemoteIP & ":" & RemotePort & " 結果受信完了 SendingIndex:" & SendingIndex
                    Call RaiseMessage(msg)

                    RecvCheckTimer.Stop() '受信完了したのでタイマーを止める

                Case SockStatuses.Receiving
                    '受信タイムアウトチェック
                    If DateTime.Now >= SendCmdList(SendingIndex).RecvTimeLimit Then
                        'タイムアウト

                        Dim msg As String = RemoteIP & ":" & RemotePort & "->" & localIP & ":" & localPort & " 受信待ちタイムアウト SendingIndex:" & SendingIndex

                        Call RaiseError(msg, SockErrors.RecvErr)

                        RecvStatus = SockStatuses.Received
                        RecvCheckTimer.Stop()

                    End If

            End Select


        Catch ex As Exception
            Dim msg As String = RemoteIP & ":" & RemotePort & " RecvCheckTimer_Elapsedエラー:" & ex.Message
            Call RaiseMessage(msg)

        End Try
        RecvCheckTimerExecuting = False

    End Sub
End Class 'clsSndSock

Public Class F
    ''' <summary>
    ''' バイト配列をhex文字列に 例)000112FF（F.Byte2HexChar）
    ''' </summary>
    ''' <param name="bytearr">バイト配列</param>
    ''' <returns>inputをhex文字列にした結果</returns>
    ''' <remarks>バイト配列をhex文字列に 例)000112FF（F.Byte2HexChar）</remarks>
    Public Shared Function Byte2HexChar(ByVal bytearr() As Byte) As String
        Dim l As Long           'ループ添え字
        Dim result As String    '戻り値

        result = ""
        If bytearr Is Nothing Then

        Else
            For l = LBound(bytearr) To UBound(bytearr)
                result &= Right("0" & Hex(bytearr(l)), 2)
            Next
        End If
        Return result
    End Function

    ''' <summary>
    ''' hex文字列をバイト配列に
    ''' </summary>
    ''' <param name="hexString">hex文字列</param>
    ''' <returns>バイト配列</returns>
    ''' <remarks>hex文字列をバイト配列に</remarks>
    Public Shared Function HexChar2Byte(ByVal hexString As String) As Byte()
        Dim l As Long           'ループ添え字
        Dim result(0) As Byte   '戻り値

        For l = 0 To hexString.Length - 1 Step 2
            If UBound(result) < l / 2 Then
                ReDim Preserve result(0 To l / 2)
            End If
            result(l / 2) = Convert.ToInt16(hexString.Substring(l, 2), 16)
        Next

        Return result
    End Function


    ''' <summary>
    ''' バイト配列をString文字列に
    ''' </summary>
    ''' <param name="bytearr">バイト配列</param>
    ''' <returns>String文字列</returns>
    ''' <remarks>バイト配列をString文字列に</remarks>
    Public Shared Function Byte2String(ByVal bytearr() As Byte) As String
        Dim l As Long           'ループ添え字
        Dim result As String    '戻り値

        result = ""
        If bytearr Is Nothing Then

        Else
            For l = LBound(bytearr) To UBound(bytearr)
                result &= Chr(bytearr(l))
            Next
        End If
        Return result
    End Function
    ''' <summary>
    ''' String文字列をバイト配列に
    ''' </summary>
    ''' <param name="Value">String文字列</param>
    ''' <returns>バイト配列</returns>
    ''' <remarks>String文字列をバイト配列に</remarks>
    Public Shared Function String2Byte(ByVal Value As String) As Byte()
        Dim byteArr() As Byte

        If Value.Length = 0 Then
            Return Nothing
        Else
            ReDim byteArr(0)
            For i As Integer = 1 To Value.Length
                ReDim Preserve byteArr(i - 1)
                byteArr(i - 1) = Asc(Value.Substring(i - 1, 1))
            Next

            Return byteArr
        End If
    End Function

    ''' <summary>
    ''' 待ち処理
    ''' </summary>
    ''' <param name="msec">待つ時間（ミリ秒）</param>
    ''' <remarks>待ち処理</remarks>
    Public Shared Sub Wait(ByVal msec As Long)
        Dim loopEnd As DateTime = DateTime.Now.AddMilliseconds(msec)
        Do
            System.Windows.Forms.Application.DoEvents()
            If DateTime.Now > loopEnd Then
                Exit Do
            End If
        Loop

    End Sub


End Class
