Imports System
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading
Imports Microsoft.VisualBasic




''' <summary>
''' ����M�Ǘ��N���X
''' </summary>
''' <remarks>����M�Ǘ��N���X</remarks>
Public Class StateObject
    ''' <summary>
    ''' ����M�o�b�t�@
    ''' </summary>
    ''' <remarks></remarks>
    Private _Buffer As Byte()

    ''' <summary>
    ''' ���[�N�\�P�b�g
    ''' </summary>
    ''' <remarks>���[�N�\�P�b�g</remarks>
    Public Property WorkSocket() As Socket = Nothing

    ''' <summary>
    ''' ��M�T�C�Y�ő�
    ''' </summary>
    ''' <remarks>��M�T�C�Y�ő�</remarks>
    Public Const BufferSize As Integer = 1024

    ''' <summary>
    ''' ��M�o�b�t�@
    ''' </summary>
    ''' <remarks>��M�o�b�t�@</remarks>
    Public ReadOnly Property Buffer As Byte()
        Get
            Return _Buffer
        End Get
    End Property


    ''' <summary>
    ''' ��M�f�[�^�𕶎��ɂ�������
    ''' </summary>
    ''' <remarks>��M�f�[�^�𕶎��ɂ�������</remarks>
    Public Property Sb() As New StringBuilder

    ''' <summary>
    ''' ��M�f�[�^
    ''' </summary>
    ''' <remarks>��M�f�[�^</remarks>
    Public Property bRecv() As Byte()

    ''' <summary>
    ''' ���M�f�[�^
    ''' </summary>
    ''' <remarks>���M�f�[�^</remarks>
    Public Property bSend() As Byte()

    ''' <summary>
    ''' ����IP
    ''' </summary>
    ''' <remarks>����IP</remarks>
    Public Property RemoteIP() As String

    ''' <summary>
    ''' ����PORT
    ''' </summary>
    ''' <remarks>����PORT</remarks>
    Public Property RemotePort() As String

    ''' <summary>
    ''' ����IP
    ''' </summary>
    ''' <remarks>����IP</remarks>
    Public Property LocalIP() As String

    ''' <summary>
    ''' ����PORT
    ''' </summary>
    ''' <remarks>����PORT</remarks>
    Public Property LocalPort() As String


    Public Sub New()
        _Buffer = New Byte(BufferSize) {}
    End Sub
End Class 'StateObject


''' <summary>
''' �\�P�b�g���b�Z�[�W�C�x���g����
''' </summary>
''' <remarks>�\�P�b�g���b�Z�[�W�C�x���g����</remarks>
Public Class SockMessageEventArgs
    Inherits EventArgs

    ''' <summary>
    ''' ���b�Z�[�W
    ''' </summary>
    ''' <remarks>���b�Z�[�W</remarks>
    Public Property Msg() As String
End Class

''' <summary>
''' �\�P�b�g�G���[�C�x���g����
''' </summary>
''' <remarks>�\�P�b�g�G���[�C�x���g����</remarks>
Public Class SockErrorEventArgs
    Inherits EventArgs

    ''' <summary>
    ''' ���b�Z�[�W
    ''' </summary>
    ''' <remarks>���b�Z�[�W</remarks>
    Public Property Msg() As String

	''' <summary>
	''' �G���[���
	''' </summary>
	''' <remarks>�G���[���</remarks>
    Public Property ErrorType() As SockErrors

	''' <summary>
	''' ����IP (�G���[��ʂɂ���Ă͓����Ă��Ȃ��j
	''' </summary>
	''' <remarks>����IP (�G���[��ʂɂ���Ă͓����Ă��Ȃ��j</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' ����PORT(�G���[��ʂɂ���Ă͓����Ă��Ȃ��j
	''' </summary>
	''' <remarks>����PORT(�G���[��ʂɂ���Ă͓����Ă��Ȃ��j</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' ���M���ɐݒ肵��Key
	''' </summary>
	''' <remarks>���M���ɐݒ肵��Key</remarks>
    Public Property SendKey() As String

End Class

''' <summary>
''' �\�P�b�g��M�J�n�C�x���g����
''' </summary>
''' <remarks>�\�P�b�g��M�J�n�C�x���g����</remarks>
Public Class SockReceiveStartEventArgs
    Inherits EventArgs

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property LocalPort() As String

End Class

''' <summary>
''' �\�P�b�g��M�����C�x���g����
''' </summary>
''' <remarks>�\�P�b�g��M�����C�x���g����</remarks>
Public Class SockReceiveEndEventArgs
    Inherits EventArgs

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property LocalPort() As String

	''' <summary>
	''' ��M���e
	''' </summary>
	''' <remarks>��M���e</remarks>
    Public Property RecvChar() As String

	''' <summary>
	''' ���M���ɐݒ肵��Key
	''' </summary>
	''' <remarks>���M���ɐݒ肵��Key</remarks>
    Public Property SendKey() As String

End Class

''' <summary>
''' �\�P�b�g���M�����C�x���g����
''' </summary>
''' <remarks>�\�P�b�g���M�����C�x���g����</remarks>
Public Class SockSendEndEventArgs
    Inherits EventArgs

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property LocalPort() As String

	''' <summary>
	''' ���M���e
	''' </summary>
	''' <remarks>���M���e</remarks>
    Public Property SendChar() As String

	''' <summary>
	''' ���M���ɐݒ肵��Key
	''' </summary>
	''' <remarks>���M���ɐݒ肵��Key</remarks>
    Public Property SendKey() As String


End Class

''' <summary>
''' �\�P�b�gListen�����C�x���g����
''' </summary>
''' <remarks>�\�P�b�gListen�����C�x���g����</remarks>
Public Class SockListenEndEventArgs
    Inherits EventArgs

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property LocalPort() As String

End Class

''' <summary>
''' �\�P�b�g�ڑ������C�x���g����
''' </summary>
''' <remarks>�\�P�b�g�ڑ������C�x���g����</remarks>
Public Class SockConnectEndEventArgs
    Inherits EventArgs

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property LocalPort() As String

End Class

''' <summary>
''' �\�P�b�g�ڑ����g���C�C�x���g����
''' </summary>
''' <remarks>�\�P�b�g�ڑ����g���C�C�x���g����</remarks>
Public Class SockConnectRetryEventArgs
    Inherits EventArgs

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' ���g���C��
	''' </summary>
	''' <remarks>���g���C��</remarks>
    Public Property RetryDo() As Long

End Class

''' <summary>
''' �\�P�b�gAccept�����C�x���g����
''' </summary>
''' <remarks>�\�P�b�gAccept�����C�x���g����</remarks>
Public Class SockAcceptEndEventArgs
    Inherits EventArgs

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property LocalPort() As String

End Class

''' <summary>
''' �\�P�b�g���M���g���C�C�x���g����
''' </summary>
''' <remarks>�\�P�b�g���M���g���C�C�x���g����</remarks>
Public Class SockSendRetryEventArgs
    Inherits EventArgs

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property RemoteIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property RemotePort() As String

	''' <summary>
	''' ����IP
	''' </summary>
	''' <remarks>����IP</remarks>
    Public Property LocalIP() As String

	''' <summary>
	''' ����PORT
	''' </summary>
	''' <remarks>����PORT</remarks>
    Public Property LocalPort() As String

	''' <summary>
	''' ���g���C��
	''' </summary>
	''' <remarks>���g���C��</remarks>
    Public Property RetryDo() As Long

	''' <summary>
	''' ���M���ɐݒ肵��Key
	''' </summary>
	''' <remarks>���M���ɐݒ肵��Key</remarks>
    Public Property SendKey() As String

End Class







''' <summary>
''' �\�P�b�g����
''' </summary>
''' <remarks>�\�P�b�g����</remarks>
Public Enum SockErrors As Integer

	''' <summary>
	''' BIND�G���[
	''' </summary>
	''' <remarks>BIND�G���[</remarks>
	BindErr = 1

	''' <summary>
	''' Listen�G���[
	''' </summary>
	''' <remarks>Listen�G���[</remarks>
	ListenErr = 2

	''' <summary>
	''' Accept�G���[
	''' </summary>
	''' <remarks>Accept�G���[</remarks>
	AcceptErr = 3

	''' <summary>
	''' Connect�G���[
	''' </summary>
	''' <remarks>Connect�G���[</remarks>
	ConnectErr = 4

	''' <summary>
	''' Send�G���[
	''' </summary>
	''' <remarks>Send�G���[</remarks>
	SendErr = 5

	''' <summary>
	''' Send���g���C�I�[�o
	''' </summary>
	''' <remarks>Send���g���C�I�[�o</remarks>
	SendRetryOver = 6

    ''' <summary>
    ''' Receive�G���[
    ''' </summary>
    ''' <remarks>Receive�G���[</remarks>
	RecvErr = 8

	''' <summary>
	''' BeginReceive�G���[
	''' </summary>
	''' <remarks>BeginReceive�G���[</remarks>
	BeginRecvErr = 9

	''' <summary>
	''' Connect���g���C�I�[�o
	''' </summary>
	''' <remarks>Connect���g���C�I�[�o</remarks>
	ConnectRetryOver = 10

	''' <summary>
	''' ���̑��G���[
	''' </summary>
	''' <remarks>���̑��G���[</remarks>
	OtherErr = 999

End Enum

''' <summary>
''' �\�P�b�g���
''' </summary>
''' <remarks>�\�P�b�g���</remarks>
Public Enum SockStatuses As Integer

	''' <summary>
	''' �ڑ���
	''' </summary>
	''' <remarks>�ڑ���</remarks>
	Connecting = 1

	''' <summary>
	''' �ڑ�����
	''' </summary>
	''' <remarks>�ڑ�����</remarks>
	Connected = 2

	''' <summary>
	''' ���M��
	''' </summary>
	''' <remarks>���M��</remarks>
	Sending = 3

	''' <summary>
	''' ���M����
	''' </summary>
	''' <remarks>���M����</remarks>
	Sended = 4

	''' <summary>
	''' ��M��
	''' </summary>
	''' <remarks>��M��</remarks>
	Receiving = 5

	''' <summary>
	''' ��M����
	''' </summary>
	''' <remarks>��M����</remarks>
	Received = 6

	''' <summary>
	''' BIND��
	''' </summary>
	''' <remarks>BIND��</remarks>
	Binding = 7

	''' <summary>
	''' BIND����
	''' </summary>
	''' <remarks>BIND����</remarks>
	Binded = 8

	''' <summary>
	''' Listen��
	''' </summary>
	''' <remarks>Listen��</remarks>
	Listening = 9

	''' <summary>
	''' Listen����
	''' </summary>
	''' <remarks>Listen����</remarks>
	Listned = 10

	''' <summary>
	''' ���M�҂���
	''' </summary>
	''' <remarks>���M�҂���</remarks>
    SendWaiting = 11

End Enum




''' <summary>
''' ��M�\�P�b�g�N���X
''' </summary>
''' <remarks>��M�\�P�b�g�N���X</remarks>
Public Class RecvSockClass

    ''' <summary>
    ''' ��M�I�����蕶����B���̕�����������d���̏I���Ƃ���B���̕������w�肵���ꍇ�A���̕���������܂œd����A�����Ď�M�𑱂���B�w�肵�Ă��Ȃ��ꍇ�͂P��̒ʐM�P�ʁ��d���P�ʂƂȂ�
    ''' </summary>
    ''' <remarks>
    ''' ��M�I�����蕶����B���̕�����������d���̏I���Ƃ���B���̕������w�肵���ꍇ�A���̕���������܂œd����A�����Ď�M�𑱂���B�w�肵�Ă��Ȃ��ꍇ�͂P��̒ʐM�P�ʁ��d���P�ʂƂȂ�
    ''' </remarks>
    Public Property EndCheckChar() As Byte()

    ''' <summary>
    ''' �҂���IP
    ''' </summary>
    ''' <remarks>�҂���IP</remarks>
    Public Property LocalIP() As String

    ''' <summary>
    ''' �҂���Port
    ''' </summary>
    ''' <remarks>�҂���Port</remarks>
    Public Property LocalPort() As String

    ''' <summary>
    ''' �^�O�i�Ăяo�����Ŏ��R�Ɏg���Ă悢�j
    ''' </summary>
    ''' <remarks>�^�O�i�Ăяo�����Ŏ��R�Ɏg���Ă悢�j</remarks>
    Public TAG As String


    ''' <summary>
    ''' ��M�X�e�[�^�X
    ''' </summary>
    ''' <remarks>��M�X�e�[�^�X</remarks>
    Private Property RecvStatus() As SockStatuses

    ''' <summary>
    ''' Bind�����̃��~�b�g
    ''' </summary>
    ''' <remarks>Bind�����̃��~�b�g</remarks>
    Private Property BindTimeLimit() As DateTime

    ''' <summary>
    ''' Bind�����^�C���A�E�g�~���b�i�Ƃ肠�����Œ�A�K�v�ł����Public�ɂ��Ă��ǂ��j
    ''' </summary>
    ''' <remarks>Bind�����^�C���A�E�g�~���b�i�Ƃ肠�����Œ�A�K�v�ł����Public�ɂ��Ă��ǂ��j</remarks>
    Private BindTimeOut As Long = 3000

    ''' <summary>
    ''' Listen�����̃��~�b�g
    ''' </summary>
    ''' <remarks>Listen�����̃��~�b�g</remarks>
    Private ListenTimeLimit As DateTime

    ''' <summary>
    ''' Listen�����^�C���A�E�g�~���b�i�Ƃ肠�����Œ�A�K�v�ł����Public�ɂ��Ă��ǂ��j
    ''' </summary>
    ''' <remarks>Listen�����^�C���A�E�g�~���b�i�Ƃ肠�����Œ�A�K�v�ł����Public�ɂ��Ă��ǂ��j</remarks>
    Private ListenTimeOut As Long = 3000

    ''' <summary>
    ''' �\�P�b�g��t�Ď��^�C�}�[
    ''' </summary>
    ''' <remarks>�\�P�b�g��t�Ď��^�C�}�[</remarks>
    Private WithEvents ListenCheckTimer As New System.Timers.Timer

    ''' <summary>
    ''' �\�P�b�g��t�Ď��^�C�}�[�������t���O
    ''' </summary>
    ''' <remarks>�\�P�b�g��t�Ď��^�C�}�[�������t���O</remarks>
    Private ListenCheckTimerExecuting As Boolean = False

    ''' <summary>
    ''' ���M��ԊĎ��^�C�}�[
    ''' </summary>
    ''' <remarks>���M��ԊĎ��^�C�}�[</remarks>
    Private WithEvents SendCheckTimer As New System.Timers.Timer

    ''' <summary>
    ''' ���M��ԊĎ��^�C�}�[�������t���O
    ''' </summary>
    ''' <remarks>���M��ԊĎ��^�C�}�[�������t���O</remarks>
    Private SendCheckTimerExecuting As Boolean = False '


    ''' <summary>
    ''' ���M�p�p�����[�^
    ''' </summary>
    ''' <remarks>���M���\�b�h�̃p�����[�^</remarks>
    Public Class SendParmClass
        '�Ăяo�����ŃZ�b�g����p�����[�^

        ''' <summary>
        ''' �\�P�b�g���ʃL�[�i�Ăяo�����Ŏ��R�Ɏg���Ă悢�j
        ''' </summary>
        ''' <remarks>�\�P�b�g���ʃL�[�i�Ăяo�����Ŏ��R�Ɏg���Ă悢�j</remarks>
        Public SendKEY As String

        ''' <summary>
        ''' ���M�d��
        ''' </summary>
        ''' <remarks>���M�d��</remarks>
        Public Property sndCmd As Byte()

        ''' <summary>
        ''' ���M�^�C���A�E�g�~���b�B����l1000
        ''' </summary>
        ''' <remarks>���M�^�C���A�E�g�~���b</remarks>
        Public Property SendTimeOut() As Long = 1000

        ''' <summary>
        ''' ���M���g���C���s���񐔁B����l0
        ''' </summary>
        ''' <remarks>���M���g���C���s����</remarks>
        Public Property SendRetry() As Long = 0

        ''' <summary>
        ''' ���M���g���C�Ԋu�i�~���b�j�B����l1000
        ''' </summary>
        ''' <remarks>���M���g���C�Ԋu�i�~���b�j</remarks>
        Public Property SendRetryWait() As Long = 1000

        ''' <summary>
        ''' ���M����IP
        ''' </summary>
        ''' <remarks>���M����IP</remarks>
        Public Property RemoteIP() As String

        ''' <summary>
        ''' ���M����|�[�g
        ''' </summary>
        ''' <remarks>���M����|�[�g</remarks>
        Public Property RemotePort() As String

        '�������Ŏg���p�����[�^

        ''' <summary>
        ''' ���M���g���C���s������
        ''' </summary>
        ''' <remarks>���M���g���C���s������</remarks>
        Public Property SendRetryDo() As Long = 0

        ''' <summary>
        ''' ���M�����̃��~�b�g�i���M�J�n�����{���M�^�C���A�E�g�j
        ''' </summary>
        ''' <remarks>���M�����̃��~�b�g�i���M�J�n�����{���M�^�C���A�E�g�j</remarks>
        Public Property SendTimeLimit() As Date

        ''' <summary>
        ''' ���M�X�e�[�^�X
        ''' </summary>
        ''' <remarks>���M�X�e�[�^�X</remarks>
        Public Property SendStatus() As SockStatuses

    End Class

    ''' <summary>
    ''' ���M�d���̑҂��s��
    ''' </summary>
    ''' <remarks>���M�d���̑҂��s��</remarks>
    Private SendCmdList As New System.Collections.Generic.List(Of SendParmClass)


    ''' <summary>
    ''' ���b�Z�[�W�o�̓C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>���b�Z�[�W�o�̓C�x���g</remarks>
    Public Event OnMessage(ByVal sender As Object, ByVal e As SockMessageEventArgs)

    ''' <summary>
    ''' �G���[�����C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>�G���[�����C�x���g</remarks>
    Public Event OnError(ByVal sender As Object, ByVal e As SockErrorEventArgs)

    ''' <summary>
    ''' LISTEN�����C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>LISTEN�����C�x���g</remarks>
    Public Event OnListenEnd(ByVal sender As Object, ByVal e As SockListenEndEventArgs)

    ''' <summary>
    ''' ACCEPT�����C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>ACCEPT�����C�x���g</remarks>
    Public Event OnAcceptEnd(ByVal sender As Object, ByVal e As SockAcceptEndEventArgs)

    ''' <summary>
    ''' ��M�J�n�C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>��M�J�n�C�x���g</remarks>
    Public Event OnRecvStart(ByVal sender As Object, ByVal e As SockReceiveStartEventArgs)

    ''' <summary>
    ''' ��M�����C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>��M�����C�x���g</remarks>
    Public Event OnRecvEnd(ByVal sender As Object, ByVal e As SockReceiveEndEventArgs)

    ''' <summary>
    ''' ���M���g���C�J�n�C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>���M���g���C�J�n�C�x���g</remarks>
    Public Event OnSendRetry(ByVal sender As Object, ByVal e As SockSendRetryEventArgs)

    ''' <summary>
    ''' ���M�����C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>���M�����C�x���g</remarks>
    Public Event OnSendEnd(ByVal sender As Object, ByVal e As SockSendEndEventArgs)


    ''' <summary>
    ''' �ڑ����Ɏ���
    ''' </summary>
    ''' <remarks>�ڑ����Ɏ���</remarks>
    Private tState(0) As StateObject

    ''' <summary>
    ''' �\�P�b�g
    ''' </summary>
    ''' <remarks>�\�P�b�g</remarks>
    Private listener As Socket

    ''' <summary>
    ''' �\�P�b�g�����m�F
    ''' </summary>
    ''' <returns>True:�����Ă��� False:����ł���</returns>
    ''' <remarks>�\�P�b�g�����m�F</remarks>
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
    ''' �T�[�o�\�P�b�g�J�n
    ''' </summary>
    ''' <remarks>�T�[�o�\�P�b�g�J�n</remarks>
    Public Sub RcvSockStart()

        Try
            RecvStatus = SockStatuses.Binding
            BindTimeLimit = DateTime.Now.AddMilliseconds(BindTimeOut)

            ' Data buffer for incoming data.
            Dim bytes() As Byte = New [Byte](1023) {}

            Dim cipAddress As IPAddress = IPAddress.Parse(LocalIP)
            Dim localEndPoint As New IPEndPoint(cipAddress, CInt(LocalPort))

            Call CloseSockAll()

            '�Ď��^�C�}�[���N��
            ListenCheckTimer.Interval = 10
            ListenCheckTimer.Start()


            ' Create a TCP/IP socket.
            listener = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)

            ' Bind the socket to the local endpoint and listen for incoming connections.
            Try
                listener.Bind(localEndPoint)
            Catch ex As Exception
                Dim msg As String = LocalIP & ":" & LocalPort & " Bind�G���[ �\�P�b�g��M�ł��܂���:" & ex.Message
                Call RaiseMessage(msg)
                Return
            End Try
            RecvStatus = SockStatuses.Binded

            Try
                ListenTimeLimit = DateTime.Now.AddMilliseconds(BindTimeOut)
                RecvStatus = SockStatuses.Listening
                listener.Listen(100)
            Catch ex As Exception
                Dim msg As String = LocalIP & ":" & LocalPort & " Listen�G���[ �\�P�b�g��M�ł��܂���:" & ex.Message
                Call RaiseMessage(msg)
                Return
            End Try
            RecvStatus = SockStatuses.Listned

            Call RaiseMessage("Listen�J�n:" & LocalIP & "(" & LocalPort & ")")
            SendCheckTimer.Interval = 10
            SendCheckTimer.Start()

        Catch ex As Exception
            Dim msg As String = LocalIP & ":" & LocalPort & " �\�P�b�g�J�n�ł��܂���:" & ex.Message
            Call RaiseMessage(msg)

        End Try

    End Sub 'Main


    ''' <summary>
    ''' Listen��ԊĎ��^�C�}�[
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>Listen��ԊĎ��^�C�}�[</remarks>
    Private Sub ListenCheckTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles ListenCheckTimer.Elapsed
        If ListenCheckTimerExecuting Then
            '�^�C�}�[�������s���ł����return
            Return
        End If


        Try
            ListenCheckTimerExecuting = True


            Select Case RecvStatus
                Case SockStatuses.Binding
                    'BIND�^�C���A�E�g�`�F�b�N
                    If DateTime.Now >= BindTimeLimit Then
                        '�^�C���A�E�g

                        Dim msg As String = LocalIP & ":" & LocalPort & " BIND�^�C���A�E�g"

                        Call RaiseError(msg, SockErrors.BindErr)

                        ListenCheckTimer.Stop()

                    End If
                Case SockStatuses.Listening
                    'Listen�^�C���A�E�g�`�F�b�N
                    If DateTime.Now >= ListenTimeLimit Then
                        '�^�C���A�E�g

                        Dim msg As String = LocalIP & ":" & LocalPort & " Listen�^�C���A�E�g"

                        Call RaiseError(msg, SockErrors.ListenErr)

                        ListenCheckTimer.Stop()

                    End If
                Case SockStatuses.Listned
                    'Listen��ԂɂȂ���
                    Call RaiseListenEnd(LocalIP, LocalPort)
                    ListenCheckTimer.Stop()
                    Call BeginAccept()

            End Select


        Catch ex As Exception
            Dim msg As String = LocalIP & ":" & LocalPort & " ListenCheckTimer_Elapsed�G���[:" & ex.Message
            Call RaiseMessage(msg)

        End Try
        ListenCheckTimerExecuting = False

    End Sub

    ''' <summary>
    ''' ���ׂẴ\�P�b�g�����
    ''' </summary>
    ''' <remarks>���ׂẴ\�P�b�g�����</remarks>
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
    ''' �w�肳�ꂽ�\�P�b�g�����
    ''' </summary>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <remarks>�w�肳�ꂽ�\�P�b�g�����</remarks>
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
    ''' ���b�Z�[�W�C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="msg">���b�Z�[�W</param>
    ''' <remarks>���b�Z�[�W�C�x���g�𔭐�������</remarks>
    Private Sub RaiseMessage(ByVal msg As String)
        Dim e As New SockMessageEventArgs
        e.Msg = msg
        RaiseEvent OnMessage(Me, e)
    End Sub

    ''' <summary>
    ''' �G���[�C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="msg">���b�Z�[�W</param>
    ''' <param name="sockError">�\�P�b�g����</param>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <param name="sendKey">���M���ɐݒ肵��Key</param>
    ''' <remarks>
    ''' �G���[�C�x���g�𔭐�������
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
            Call BeginAccept() '�ēx�\�P�b�g��M��t����
        End If

        RaiseEvent OnError(Me, e)
    End Sub

    ''' <summary>
    ''' Listen�����C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="localIP">��IP</param>
    ''' <param name="localPort">��PORT</param>
    ''' <remarks>Listen�����C�x���g�𔭐�������</remarks>
    Private Sub RaiseListenEnd(ByVal localIP As String, ByVal localPort As String)
        Dim e As New SockListenEndEventArgs
        e.LocalIP = localIP
        e.LocalPort = localPort

        RaiseEvent OnListenEnd(Me, e)
    End Sub

    ''' <summary>
    ''' �ڑ���t�����C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <param name="localIP">��IP</param>
    ''' <param name="localPort">��PORT</param>
    ''' <remarks>�ڑ���t�����C�x���g�𔭐�������</remarks>
    Private Sub RaiseAcceptEnd(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String)
        Dim e As New SockAcceptEndEventArgs
        e.LocalIP = localIP
        e.LocalPort = localPort
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort

        Call BeginAccept() '�ēx�\�P�b�g��M��t����



        RaiseEvent OnAcceptEnd(Me, e)
    End Sub

    ''' <summary>
    ''' ��M�J�n�C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <param name="localIP">��IP</param>
    ''' <param name="localPort">��PORT</param>
    ''' <remarks>��M�J�n�C�x���g�𔭐�������</remarks>
    Private Sub RaiseRecvStart(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String)
        Dim e As New SockReceiveStartEventArgs
        e.LocalIP = localIP
        e.LocalPort = localPort
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort

        RaiseEvent OnRecvStart(Me, e)
    End Sub
    ''' <summary>
    ''' ��M�����C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <param name="localIP">��IP</param>
    ''' <param name="localPort">��PORT</param>
    ''' <param name="recvChar">��M���e</param>
    ''' <remarks>��M�����C�x���g�𔭐�������</remarks>
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
    ''' ���M���g���C�C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <param name="localIP">��IP</param>
    ''' <param name="localPort">��PORT</param>
    ''' <param name="RetryDoCnt">���g���C��</param>
    ''' <param name="sendKey">���MKEY</param>
    ''' <remarks>���M���g���C�C�x���g�𔭐�������</remarks>
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
    ''' ���M�����C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <param name="localIP">��IP</param>
    ''' <param name="localPort">��PORT</param>
    ''' <param name="sendChar">���M���e</param>
    ''' <param name="sendkey">���MKEY</param>
    ''' <remarks>���M�����C�x���g�𔭐�������</remarks>
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
    ''' Accept�J�n����
    ''' </summary>
    ''' <remarks>Accept�J�n����</remarks>
    Private Sub BeginAccept()
        Try
            listener.BeginAccept(New AsyncCallback(AddressOf AcceptCallback), listener)
        Catch ex As Exception
            Dim msg As String = LocalIP & ":" & LocalPort & " BeginAccept�G���[:" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.AcceptErr)

        End Try

    End Sub

    ''' <summary>
    ''' Accept�J�n�R�[���o�b�N
    ''' </summary>
    ''' <param name="ar">�\�P�b�g</param>
    ''' <remarks>Accept�J�n�R�[���o�b�N</remarks>
    Private Sub AcceptCallback(ByVal ar As IAsyncResult)

        Dim listener As Socket = CType(ar.AsyncState, Socket)
        Dim handler As Socket
        Try
            handler = listener.EndAccept(ar)
        Catch ex As ObjectDisposedException
            Dim msg As String = LocalIP & ":" & LocalPort & " Close����܂���"
            Call RaiseMessage(msg)
            Return
        Catch ex As Exception
            'listen���ɃN���[�Y���ꂽ�ꍇ�̏���
            Dim msg As String = LocalIP & ":" & LocalPort & " Accept�G���[:" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.AcceptErr)

            Return
        End Try



        Dim remoteIP As String = CType(handler.RemoteEndPoint, IPEndPoint).Address.ToString()
        Dim remotePORT As String = CType(handler.RemoteEndPoint, IPEndPoint).Port.ToString()

        Try
            Dim l = GetStateIndex(remoteIP, remotePORT, LocalIP, LocalPort)

            Call RaiseMessage(remoteIP & ":" & remotePORT & " ����ڑ����� StateNo:" & l)

            Call RaiseAcceptEnd(remoteIP, remotePORT, LocalIP, LocalPort)




            With tState(l)
                .WorkSocket = handler

            End With

            Try
                handler.BeginReceive(tState(l).Buffer, 0, StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReadCallback), tState(l))
            Catch ex As Exception
                Dim msg As String = LocalIP & ":" & LocalPort & " BeginReceive�G���[:" & ex.Message
                Call RaiseMessage(msg)
                Call RaiseError(msg, SockErrors.BeginRecvErr, remoteIP, remotePORT)
            End Try

        Catch ex As Exception
            Dim msg As String = LocalIP & ":" & LocalPort & " AcceptCallback�G���[:" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.AcceptErr, remoteIP, remotePORT)
        End Try

    End Sub 'AcceptCallback

    ''' <summary>
    ''' �w�肳�ꂽ�\�P�b�g�̃C���f�b�N�X��Ԃ��B�Ȃ���ΐV���ɍ쐬����
    ''' </summary>
    ''' <param name="remoteIP">�w��IP</param>
    ''' <param name="remotePort">�w��PORT</param>
    ''' <param name="localIP">��IP</param>
    ''' <param name="localPort">��PORT</param>
    ''' <returns>�C���f�b�N�X</returns>
    ''' <remarks>�w�肳�ꂽ�\�P�b�g�̃C���f�b�N�X��Ԃ�</remarks>
    Private Function GetStateIndex(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String) As Long
        Dim lRet As Long
        Dim l As Long
        lRet = -999
        Try


            '�w�肳�ꂽIP,Port�̃X�e�[�g�����݂��Ă��邩���ׂ�
            For l = LBound(tState) To UBound(tState)
                If tState(l) Is Nothing Then
                Else
                    If tState(l).RemoteIP = remoteIP And tState(l).RemotePort = remotePort Then
                        '���蓖�čς݂Ȃ̂ł��̃X�e�[�g���g��
                        lRet = l
                        Exit For
                    End If
                End If
            Next

            If lRet = -999 Then
                '�w�肳�ꂽIP,Port�̃X�e�[�g�͑��݂��Ă��Ȃ������̂ŋ󂫃X�e�[�g��T��
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

                '�����l
                ReDim tState(lRet).bRecv(0)
                tState(lRet).RemoteIP = remoteIP
                tState(lRet).RemotePort = remotePort
                tState(lRet).LocalIP = localIP
                tState(lRet).LocalPort = localPort
                tState(lRet).Sb = New StringBuilder

            End If



        Catch ex As Exception
            Dim msg As String = "GetStateIndex�G���[:" & remoteIP & ":" & remotePort & " " & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.OtherErr)

        End Try
        Return lRet
    End Function

    ''' <summary>
    ''' �w�肳�ꂽ�\�P�b�g�̃C���f�b�N�X��Ԃ��B�Ȃ����-999��Ԃ�
    ''' </summary>
    ''' <param name="remoteIP">�w��IP</param>
    ''' <param name="remotePort">�w��PORT</param>
    ''' <returns>�C���f�b�N�X</returns>
    ''' <remarks>�w�肳�ꂽ�\�P�b�g�̃C���f�b�N�X��Ԃ��B</remarks>
    Private Function FindStateIndex(ByVal remoteIP As String, ByVal remotePort As String) As Long
        Dim lRet As Long
        Dim l As Long
        lRet = -999
        Try


            '�w�肳�ꂽIP,Port�̃X�e�[�g�����݂��Ă��邩���ׂ�
            For l = LBound(tState) To UBound(tState)
                If tState(l) Is Nothing Then
                Else
                    If tState(l).RemoteIP = remoteIP And tState(l).RemotePort = remotePort Then
                        '���蓖�čς݂Ȃ̂ł��̃X�e�[�g���g��
                        lRet = l
                        Exit For
                    End If
                End If
            Next


        Catch ex As Exception
            RecvStatus = "ERROR"
            Dim msg As String = "FindStateIndex�G���[:" & remoteIP & ":" & remotePort & " " & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.OtherErr)

        End Try
        Return lRet
    End Function


    ''' <summary>
    ''' Read�R�[���o�b�N
    ''' </summary>
    ''' <param name="ar">�\�P�b�g</param>
    ''' <remarks>Read�R�[���o�b�N</remarks>
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
                Dim msg As String = "[ReadCallback]" & state.RemoteIP & ":" & state.RemotePort & " ��Close�ς݂ł�"
                Call RaiseMessage(msg)
                Return
            Catch ex As Exception
                'RECEIVE���ɃN���[�Y���ꂽ�ꍇ�̏���
                Dim msg As String = LocalIP & ":" & LocalPort & " EndReceive�G���[:" & ex.Message
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
                Call RaiseMessage(state.RemoteIP & ":" & state.RemotePort & "->" & state.LocalIP & ":" & state.LocalPort & " " & content.Length.ToString & "�o�C�g��M-> " & F.Byte2String(state.bRecv) & "(" & F.Byte2HexChar(state.bRecv) & ")")

                If content.Length > StateObject.BufferSize Then
                    '���b�Z�[�W�\��
                    Call RaiseMessage(state.RemoteIP & ":" & state.RemotePort & "����o�b�t�@�T�C�Y�𒴂���f�[�^���󂯎��܂����B��M�f�[�^��j�����܂� : " + F.Byte2HexChar(state.bRecv))
                    ReDim state.bRecv(0) '��M�f�[�^�̏�����
                    state.Sb = New StringBuilder

                Else
                    If EndCheckChar.Length = 0 Then
                        IsReadEnd = True
                    Else
                        '�I�����蕶�������M���Ă���Ί���
                        If content.Length >= EndCheckChar.Length Then
                            If F.Byte2HexChar(state.bRecv).EndsWith(F.Byte2HexChar(EndCheckChar)) Then
                                IsReadEnd = True
                            End If
                        End If

                    End If
                End If


            End If


            If IsReadEnd Then
                '��M�����C�x���g
                Call RaiseRecvEnd(state.RemoteIP, state.RemotePort, LocalIP, LocalPort, F.Byte2HexChar(state.bRecv))

                ReDim state.bRecv(0) '��M�f�[�^�̏�����
                state.Sb = New StringBuilder

            End If

            Try
                '��M�ҋ@
                handler.BeginReceive(state.Buffer, 0, StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReadCallback), state)
                ''Call BeginAccept() '�ēx�\�P�b�g��M��t����
            Catch ex As Exception
                Dim msg As String = LocalIP & ":" & LocalPort & " BeginReceive�G���[:" & ex.Message
                Call RaiseMessage(msg)
                Call RaiseError(msg, SockErrors.BeginRecvErr, state.RemoteIP, state.RemotePort)
            End Try


        Catch ex As Exception
            Dim msg As String = LocalIP & ":" & LocalPort & " ReadCallback�G���[:" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.RecvErr, state.RemoteIP, state.RemotePort)
        End Try
    End Sub 'ReadCallback


    ''' <summary>
    ''' �\�P�b�g���M�o�^
    ''' </summary>
    ''' <param name="SendParm">���M�p�����[�^</param>
    ''' <remarks>�\�P�b�g���M�o�^</remarks>
    Public Sub SockSend(ByVal SendParm As SendParmClass)
        SyncLock SendCmdList
            SendParm.SendRetryDo = 0
            SendParm.SendStatus = SockStatuses.SendWaiting
            SendCmdList.Add(SendParm)
        End SyncLock
    End Sub

    ''' <summary>
    ''' �\�P�b�g���M�Ď��^�C�}�[
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>�\�P�b�g���M�Ď��^�C�}�[</remarks>
    Private Sub SendCheckTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles SendCheckTimer.Elapsed
        If SendCheckTimerExecuting Then
            '�^�C�}�[�������s���ł����return
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
                '���M�҂��s����`�F�b�N���đ��M�ς݂�����΍폜����
                SendCmdList.RemoveAll(AddressOf IsRemoveSendParm)
            End If

            For sendParmIndex As Long = 0 To SendCmdList.Count - 1
                Dim sendparm As SendParmClass = SendCmdList(sendParmIndex)
                Select Case sendparm.SendStatus
                    Case SockStatuses.SendWaiting

                        '����IP�APort�ő��M���̂��̂����邩�`�F�b�N����
                        Dim IsSending As Boolean = False
                        For l As Long = 0 To SendCmdList.Count - 1
                            If SendCmdList(l).RemoteIP = sendparm.RemoteIP And _
                               SendCmdList(l).RemotePort = sendparm.RemotePort And _
                               SendCmdList(l).SendStatus = SockStatuses.Sending Then
                                IsSending = True '���M���̃\�P�b�g����
                                Exit For

                            End If
                        Next
                        If IsSending = False Then
                            '���M�J�n
                            sendparm.SendTimeLimit = DateTime.Now.AddMilliseconds(sendparm.SendTimeOut)
                            sendparm.SendStatus = SockStatuses.Sending
                            Call Send(sendparm.RemoteIP, sendparm.RemotePort, sendparm.sndCmd)
                        End If
                    Case SockStatuses.Sending
                        If LocalIP = "0.0.0.0" Then
                            Dim msg As String = String.Empty
                            msg = LocalIP & ":" & LocalPort & "->" & sendparm.RemoteIP & ":" & sendparm.RemotePort & " �Z�b�V�������؂�Ă���̂Ń��g���C�I�[�o�[�����Ƃ��܂� "
                            Call RaiseMessage(msg)

                            sendparm.SendTimeLimit = DateTime.Now
                            sendparm.SendRetryDo = sendparm.SendRetry
                        End If


                        '���M�^�C���A�E�g�`�F�b�N
                        If DateTime.Now >= sendparm.SendTimeLimit Then
                            If sendparm.SendRetryDo = sendparm.SendRetry Then
                                '���g���C�I�[�o�[

                                Dim msg As String = LocalIP & ":" & LocalPort & "->" & sendparm.RemoteIP & ":" & sendparm.RemotePort & " ���M���g���C�I�[�o"

                                Call RaiseError(msg, SockErrors.SendRetryOver, sendparm.RemoteIP, sendparm.RemotePort, sendparm.SendKEY)

                                sendparm.SendStatus = SockStatuses.Sended

                            Else
                                '���g���C
                                sendparm.SendRetryDo += 1
                                Dim msg As String = String.Empty
                                msg = LocalIP & ":" & LocalPort & "->" & sendparm.RemoteIP & ":" & sendparm.RemotePort & " ���M���g���C�J�n�҂�..."
                                Call RaiseMessage(msg)

                                Call F.Wait(sendparm.SendRetryWait)

                                msg = LocalIP & ":" & LocalPort & "->" & sendparm.RemoteIP & ":" & sendparm.RemotePort & " �֑��M�J�n ���g���C" & sendparm.SendRetryDo & "���"
                                Call RaiseMessage(msg)
                                Call RaiseSendRetry(sendparm.RemoteIP, sendparm.RemotePort, LocalIP, LocalPort, sendparm.SendRetryDo, sendparm.SendKEY)

                                Call Send(sendparm.RemoteIP, sendparm.RemotePort, sendparm.sndCmd)
                            End If

                        End If

                End Select

            Next




        Catch ex As Exception
            Dim msg As String = "SendCheckTimer_Elapsed�G���[:" & ex.Message
            Call RaiseMessage(msg)

        End Try
        SendCheckTimerExecuting = False


    End Sub

    ''' <summary>
    ''' ���M�p�����[�^���폜�\�����肷��
    ''' </summary>
    ''' <param name="sendParm">���M�p�����[�^</param>
    ''' <returns>True:�폜�\ False:�폜�s��</returns>
    ''' <remarks>���M�p�����[�^���폜�\�����肷��</remarks>
    Private Function IsRemoveSendParm(ByVal sendParm As SendParmClass) As Boolean
        If sendParm.SendStatus = SockStatuses.Sended Then
            Return True
        Else
            Return False
        End If
    End Function


    ''' <summary>
    ''' �\�P�b�g���M
    ''' </summary>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <param name="sndData">���M���e</param>
    ''' <remarks>�\�P�b�g���M</remarks>
    Public Sub Send(ByVal remoteIP As String, ByVal remotePort As String, ByVal sndData As Byte())
        Try
            Dim l As Long = FindStateIndex(remoteIP, remotePort)
            Dim msg As String = String.Empty

            If l < 0 Then
                '�w�肳�ꂽIP�APort�̐ڑ����������
                msg = remoteIP & ":" & remotePort & " �̃\�P�b�g��������Ȃ��̂ő��M�ł��܂���"
                Call RaiseMessage(msg)
                Call RaiseError(msg, SockErrors.SendErr, remoteIP, remotePort)
                Return

            End If

            Dim sendingIndex As Long = GetSendingIndex(remoteIP, remotePort)
            If sendingIndex >= 0 Then
                msg = LocalIP & ":" & LocalPort & "->" & remoteIP & ":" & remotePort & " ���M�J�n �^�C���A�E�g->" & Format(SendCmdList(sendingIndex).SendTimeLimit, "yyyy/MM/dd HH:mm:ss.ff") & " index:" & sendingIndex
            Else
                msg = LocalIP & ":" & LocalPort & "->" & remoteIP & ":" & remotePort & " ���M�J�n" & " index:" & sendingIndex
            End If
            Call RaiseMessage(msg)

            tState(l).WorkSocket.BeginSend(sndData, 0, sndData.Length, 0, New AsyncCallback(AddressOf SendCallback), tState(l).WorkSocket)
        Catch ex As Exception

            Dim msg As String = remoteIP & ":" & remotePort & " �ւ̑��M�G���[�F" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.SendErr, remoteIP, remotePort)
        End Try
    End Sub 'Send


    ''' <summary>
    ''' �\�P�b�g���M�R�[���o�b�N
    ''' </summary>
    ''' <param name="ar">�\�P�b�g</param>
    ''' <remarks>�\�P�b�g���M�R�[���o�b�N</remarks>
    Private Sub SendCallback(ByVal ar As IAsyncResult)
        Dim handler As Socket
        Dim remoteIP As String
        Dim remotePORT As String

        Try
            handler = CType(ar.AsyncState, Socket)
            remoteIP = CType(handler.RemoteEndPoint, IPEndPoint).Address.ToString()
            remotePORT = CType(handler.RemoteEndPoint, IPEndPoint).Port.ToString()
        Catch ex As Exception
            Dim msg As String = "���M����IP�擾�G���[(SendCallback)�F" & ex.Message
            Call RaiseMessage(msg)
            Return
            '�����ł�remoteIP,remotePort�͕�����Ȃ��\��������BCall RaiseError( msg, SockErrors.HealthCheckErr, remoteIP, remotePORT)
        End Try



        Try

            Dim bytesSent As Integer = handler.EndSend(ar)

            Dim localIP As String = CType(handler.LocalEndPoint, IPEndPoint).Address.ToString()
            Dim localPort As String = CType(handler.LocalEndPoint, IPEndPoint).Port.ToString()

            If localIP.Equals("0.0.0.0") Then
                '�Z�b�V�������؂�Ă���
                Dim msg As String = localIP & ":" & localPort & " - " & remoteIP & ":" & remotePORT & " �Ԃ̐ڑ����؂�Ă��܂�"
                Call RaiseMessage(msg)
                Call RaiseError(msg, SockErrors.SendErr, remoteIP, remotePORT)
                Return

            End If



            '���M���̓d����T��
            Dim l As Long = GetSendingIndex(remoteIP, remotePORT)

            If l >= 0 Then


                Call RaiseSendEnd(SendCmdList(l).RemoteIP, SendCmdList(l).RemotePort, localIP, localPort, F.Byte2HexChar(SendCmdList(l).sndCmd), SendCmdList(l).SendKEY)

                '���b�Z�[�W�\��
                Call RaiseMessage(localIP & ":" & localPort & "->" & remoteIP & ":" & remotePORT & " " & _
                      bytesSent & "�o�C�g���M-> " & F.Byte2String(SendCmdList(l).sndCmd) & "(" & F.Byte2HexChar(SendCmdList(l).sndCmd) & ")")


                SendCmdList(l).SendStatus = SockStatuses.Sended
            End If


            Call RaiseMessage(remoteIP & ":" & remotePORT & " �֑��M���܂���")

        Catch ex As Exception
            Dim msg As String = remoteIP & ":" & remotePORT & " �ւ̑��M�G���[(SendCallback)�F" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.SendErr, remoteIP, remotePORT)
        End Try
    End Sub 'SendCallback


    ''' <summary>
    ''' �w�肳�ꂽ�\�P�b�g�̑��M����Ԃ̃C���f�b�N�X���擾����
    ''' </summary>
    ''' <param name="remoteIP">�w��IP</param>
    ''' <param name="remotePort">�w��PORT</param>
    ''' <returns>�C���f�b�N�X</returns>
    ''' <remarks>�w�肳�ꂽ�\�P�b�g�̑��M����Ԃ̃C���f�b�N�X���擾����</remarks>
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
                '�������s�����ƁA�������ɔz�񂪍폜����邱�Ƃ����邽�߃��g���C����
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
''' ���M�\�P�b�g�N���X
''' </summary>
''' <remarks>���M�\�P�b�g�N���X</remarks>
Public Class SendSockClass

    ''' <summary>
    ''' ��M�I�����蕶����B���̕�����������d���̏I���Ƃ���B���̕������w�肵���ꍇ�A���̕���������܂œd����A�����Ď�M�𑱂���B�w�肵�Ă��Ȃ��ꍇ�͂P��̒ʐM�P�ʁ��d���P�ʂƂȂ�
    ''' </summary>
    ''' <remarks>��M�I�����蕶����B���̕�����������d���̏I���Ƃ���B���̕������w�肵���ꍇ�A���̕���������܂œd����A�����Ď�M�𑱂���B�w�肵�Ă��Ȃ��ꍇ�͂P��̒ʐM�P�ʁ��d���P�ʂƂȂ� </remarks>
    Public Property EndCheckChar() As Byte()

    ''' <summary>
    ''' ����IP
    ''' </summary>
    ''' <remarks>����IP</remarks>
    Public Property RemoteIP() As String

    ''' <summary>
    ''' ����Port
    ''' </summary>
    ''' <remarks>����Port</remarks>
    Public Property RemotePort() As String

    ''' <summary>
    ''' �ڑ��^�C���A�E�g�~���b�B����l1000
    ''' </summary>
    ''' <remarks>�ڑ��^�C���A�E�g�~���b</remarks>
    Public Property ConnectTimeOut() As Long = 1000

    ''' <summary>
    ''' �ڑ����g���C���s���񐔁B����l0
    ''' </summary>
    ''' <remarks>�ڑ����g���C���s����</remarks>
    Public Property ConnectRetry() As Long

    ''' <summary>
    ''' �ڑ����g���C�Ԋu�i�~���b�j�B����l1000
    ''' </summary>
    ''' <remarks>�ڑ����g���C�Ԋu�i�~���b�j</remarks>
    Public Property ConnectRetryWait() As Long = 1000


    ''' <summary>
    ''' �ڑ������̃��~�b�g�i�ڑ��J�n�����{�ڑ��^�C���A�E�g�j
    ''' </summary>
    ''' <remarks>�ڑ������̃��~�b�g�i�ڑ��J�n�����{�ڑ��^�C���A�E�g�j</remarks>
    Private ConnectTimeLimit As DateTime

    ''' <summary>
    ''' �ڑ����g���C���s������
    ''' </summary>
    ''' <remarks>�ڑ����g���C���s������</remarks>
    Private ConnectRetryDo As Long = 0

    ''' <summary>
    ''' �ڑ���ԊĎ��^�C�}�[
    ''' </summary>
    ''' <remarks>�ڑ���ԊĎ��^�C�}�[</remarks>
    Private WithEvents ConnectCheckTimer As New System.Timers.Timer

    ''' <summary>
    ''' �ڑ���ԊĎ��^�C�}�[�������t���O
    ''' </summary>
    ''' <remarks>�ڑ���ԊĎ��^�C�}�[�������t���O</remarks>
    Private ConnectCheckTimerExecuting As Boolean = False

    ''' <summary>
    ''' ���M�X�e�[�^�X
    ''' </summary>
    ''' <remarks>���M�X�e�[�^�X</remarks>
    Private SendStatus As SockStatuses

    ''' <summary>
    ''' ��M�X�e�[�^�X
    ''' </summary>
    ''' <remarks>��M�X�e�[�^�X</remarks>
    Private RecvStatus As SockStatuses


    ''' <summary>
    ''' ���M�p�p�����[�^
    ''' </summary>
    ''' <remarks>���M���\�b�h�̃p�����[�^</remarks>
    Public Class SendParmClass
        '�Ăяo�����ŃZ�b�g����p�����[�^

        ''' <summary>
        ''' �\�P�b�g���ʃL�[�i�Ăяo�����Ŏ��R�Ɏg���Ă悢�j
        ''' </summary>
        ''' <remarks>�\�P�b�g���ʃL�[�i�Ăяo�����Ŏ��R�Ɏg���Ă悢�j</remarks>
        Public Property SendKEY() As String

        ''' <summary>
        ''' ���M�d��
        ''' </summary>
        ''' <remarks>���M�d��</remarks>
        Public Property sndCmd() As Byte()

        ''' <summary>
        ''' ���M�^�C���A�E�g�~���b�B����l1000
        ''' </summary>
        ''' <remarks>���M�^�C���A�E�g�~���b</remarks>
        Public Property SendTimeOut() As Long = 1000

        ''' <summary>
        ''' ���M���g���C���s���񐔁B����l0
        ''' </summary>
        ''' <remarks>���M���g���C���s����</remarks>
        Public Property SendRetry() As Long = 0

        ''' <summary>
        ''' ���M���g���C�Ԋu�i�~���b�j
        ''' </summary>
        ''' <remarks>���M���g���C�Ԋu�i�~���b�j</remarks>
        Public Property SendRetryWait() As Long

        ''' <summary>
        ''' ���M������ɃT�[�o�[�\�P�b�g����̕ԐM�҂����s����
        ''' </summary>
        ''' <remarks>���M������ɃT�[�o�[�\�P�b�g����̕ԐM�҂����s����</remarks>
        Public Property IsRecv() As Boolean

        ''' <summary>
        ''' �T�[�o�[�\�P�b�g����̕ԐM�҂��^�C���A�E�g�~���b
        ''' </summary>
        ''' <remarks>�T�[�o�[�\�P�b�g����̕ԐM�҂��^�C���A�E�g�~���b</remarks>
        Public Property RecvTimeOut() As Long

        '�������Ŏg���p�����[�^

        ''' <summary>
        ''' ���M���g���C���s������
        ''' </summary>
        ''' <remarks>���M���g���C���s������</remarks>
        Public Property SendRetryDo() As Long = 0

        ''' <summary>
        ''' ���M�����̃��~�b�g�i���M�J�n�����{���M�^�C���A�E�g�j
        ''' </summary>
        ''' <remarks>���M�����̃��~�b�g�i���M�J�n�����{���M�^�C���A�E�g�j</remarks>
        Public Property SendTimeLimit() As DateTime

        ''' <summary>
        ''' ���M�����ς݂ł����True
        ''' </summary>
        ''' <remarks>���M�����ς݂ł����True</remarks>
        Public Property IsSendEnd() As Boolean

        ''' <summary>
        ''' ��M�����̃��~�b�g�i��M�J�n�����{���M�^�C���A�E�g�j
        ''' </summary>
        ''' <remarks>��M�����̃��~�b�g�i��M�J�n�����{���M�^�C���A�E�g�j</remarks>
        Public Property RecvTimeLimit() As DateTime

        ''' <summary>
        ''' ��M�����ς݂ł����True
        ''' </summary>
        ''' <remarks>��M�����ς݂ł����True</remarks>
        Public Property IsRecvEnd() As Boolean

    End Class

    ''' <summary>
    ''' ���M�d���̑҂��s��
    ''' </summary>
    ''' <remarks>���M�d���̑҂��s��</remarks>
    Private SendCmdList As New System.Collections.Generic.List(Of SendParmClass)

    ''' <summary>
    ''' ���ݑ��M�������̑҂��s��
    ''' </summary>
    ''' <remarks>���ݑ��M�������̑҂��s��</remarks>
    Private SendingIndex As Long

    ''' <summary>
    ''' ���M��ԊĎ��^�C�}�[
    ''' </summary>
    ''' <remarks>���M��ԊĎ��^�C�}�[</remarks>
    Private WithEvents SendCheckTimer As New System.Timers.Timer

    ''' <summary>
    ''' ���M��ԊĎ��^�C�}�[�������t���O
    ''' </summary>
    ''' <remarks>���M��ԊĎ��^�C�}�[�������t���O</remarks>
    Private SendCheckTimerExecuting As Boolean = False

    ''' <summary>
    ''' ��M��ԊĎ��^�C�}�[
    ''' </summary>
    ''' <remarks>��M��ԊĎ��^�C�}�[</remarks>
    Private WithEvents RecvCheckTimer As New System.Timers.Timer

    ''' <summary>
    ''' ��M��ԊĎ��^�C�}�[�������t���O
    ''' </summary>
    ''' <remarks>��M��ԊĎ��^�C�}�[�������t���O</remarks>
    Private RecvCheckTimerExecuting As Boolean = False

    ''' <summary>
    ''' �^�O�i�Ăяo�����Ŏ��R�Ɏg���Ă悢�j
    ''' </summary>
    ''' <remarks>�^�O�i�Ăяo�����Ŏ��R�Ɏg���Ă悢�j</remarks>
    Public TAG As String



    ''' <summary>
    ''' ���b�Z�[�W�o�̓C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>���b�Z�[�W�o�̓C�x���g</remarks>
    Public Event OnMessage(ByVal sender As Object, ByVal e As SockMessageEventArgs)


    ''' <summary>
    ''' �ڑ������C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>�ڑ������C�x���g</remarks>
    Public Event OnConnectEnd(ByVal sender As Object, ByVal e As SockConnectEndEventArgs)

    ''' <summary>
    ''' �ڑ����g���C�J�n�C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>�ڑ����g���C�J�n�C�x���g</remarks>
    Public Event OnConnectRetry(ByVal sender As Object, ByVal e As SockConnectRetryEventArgs)

    ''' <summary>
    ''' ���M���g���C�J�n�C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>���M���g���C�J�n�C�x���g</remarks>
    Public Event OnSendRetry(ByVal sender As Object, ByVal e As SockSendRetryEventArgs)

    ''' <summary>
    ''' ���M�����C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>���M�����C�x���g</remarks>
    Public Event OnSendEnd(ByVal sender As Object, ByVal e As SockSendEndEventArgs)

    ''' <summary>
    ''' ��M�����C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>��M�����C�x���g</remarks>
    Public Event OnRecvEnd(ByVal sender As Object, ByVal e As SockReceiveEndEventArgs)

    ''' <summary>
    ''' �G���[�����C�x���g
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>�G���[�����C�x���g</remarks>
    Public Event OnError(ByVal sender As Object, ByVal e As SockErrorEventArgs)


    ''' <summary>
    ''' IP�A�h���X���Ɏ���
    ''' </summary>
    ''' <remarks>IP�A�h���X���Ɏ���</remarks>
    Private tState(0) As StateObject

    ''' <summary>
    ''' �N���C�A���g�\�P�b�g
    ''' </summary>
    ''' <remarks>�N���C�A���g�\�P�b�g</remarks>
    Private client As Socket

    Public Sub New()
        EndCheckChar = New Byte() {}
    End Sub

    ''' <summary>
    ''' �\�P�b�g�����m�F
    ''' </summary>
    ''' <returns>True:�����Ă��� False:����ł���</returns>
    ''' <remarks>�\�P�b�g�����m�F</remarks>
    <Obsolete("���̃��\�b�h�͑���i�T�[�o�[�\�P�b�g�j����ؒf���ꂽ�ꍇ�͋@�\���܂���iTrue�̂܂ܕς��Ȃ��j�̂ŁA�g�����ɒ��ӂ��ĉ�����")> _
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
    ''' ���b�Z�[�W�C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="msg">���b�Z�[�W</param>
    ''' <remarks>���b�Z�[�W�C�x���g�𔭐�������</remarks>
    Private Sub RaiseMessage(ByVal msg As String)
        Dim e As New SockMessageEventArgs
        e.Msg = msg
        RaiseEvent OnMessage(Me, e)
    End Sub

    ''' <summary>
    ''' �G���[�C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="msg">���b�Z�[�W</param>
    ''' <param name="sockError">�G���[�^�C�v</param>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <param name="sendKey">���MKEY</param>
    ''' <remarks>�G���[�C�x���g�𔭐�������</remarks>
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
    ''' �ڑ������C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <param name="localIP">��IP</param>
    ''' <param name="localPort">��PORT</param>
    ''' <remarks>�ڑ������C�x���g�𔭐�������</remarks>
    Private Sub RaiseConnectEnd(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String)
        Dim e As New SockConnectEndEventArgs
        e.LocalIP = localIP
        e.LocalPort = localPort
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort

        RaiseEvent OnConnectEnd(Me, e)
    End Sub

    ''' <summary>
    ''' �ڑ����g���C�C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <param name="RetryDoCnt">���g���C��</param>
    ''' <remarks>�ڑ����g���C�C�x���g�𔭐�������</remarks>
    Private Sub RaiseConnectRetry(ByVal remoteIP As String, ByVal remotePort As String, ByVal RetryDoCnt As Long)
        Dim e As New SockConnectRetryEventArgs
        e.RemoteIP = remoteIP
        e.RemotePort = remotePort
        e.RetryDo = RetryDoCnt

        RaiseEvent OnConnectRetry(Me, e)
    End Sub


    ''' <summary>
    ''' ��M�����C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <param name="localIP">��IP</param>
    ''' <param name="localPort">��PORT</param>
    ''' <param name="recvChar">��M���e</param>
    ''' <remarks>��M�����C�x���g�𔭐�������</remarks>
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
    ''' ���M���g���C�C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <param name="localIP">��IP</param>
    ''' <param name="localPort">��PORT</param>
    ''' <param name="RetryDoCnt">���g���C��</param>
    ''' <param name="sendKey">���MKEY</param>
    ''' <remarks>���M���g���C�C�x���g�𔭐�������</remarks>
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
    ''' ���M�����C�x���g�𔭐�������
    ''' </summary>
    ''' <param name="remoteIP">����IP</param>
    ''' <param name="remotePort">����PORT</param>
    ''' <param name="localIP">��IP</param>
    ''' <param name="localPort">��PORT</param>
    ''' <param name="sendChar">���M���e</param>
    ''' <param name="sendkey">���MKEY</param>
    ''' <remarks>���M�����C�x���g�𔭐�������</remarks>
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
    ''' �\�P�b�g�ڑ��J�n����
    ''' </summary>
    ''' <remarks>�\�P�b�g�ڑ��J�n����</remarks>
    Public Sub SockConnect()

        Call CloseSock()
        ConnectRetryDo = 0
        Call Connect()
        ConnectCheckTimer.Interval = 10
        ConnectCheckTimer.Start()

    End Sub

    ''' <summary>
    ''' �ڑ��`�F�b�N�^�C�}�[
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>�ڑ��`�F�b�N�^�C�}�[</remarks>
    Private Sub ConnectCheckTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles ConnectCheckTimer.Elapsed
        If ConnectCheckTimerExecuting Then
            '�^�C�}�[�������s���ł����return
            Return
        End If

        ConnectCheckTimerExecuting = True
        Select Case SendStatus
            Case SockStatuses.Connecting
                '�ڑ��^�C���A�E�g�`�F�b�N
                If DateTime.Now >= ConnectTimeLimit Then
                    If ConnectRetryDo = ConnectRetry Then
                        '���g���C�I�[�o�[
                        Dim msg As String = RemoteIP & ":" & RemotePort & " �֐ڑ����g���C�I�[�o�["
                        Call RaiseError(msg, SockErrors.ConnectRetryOver)
                        ConnectCheckTimer.Stop()
                    Else
                        '���g���C
                        ConnectRetryDo += 1
                        Dim msg As String = String.Empty
                        msg = RemoteIP & ":" & RemotePort & " �֐ڑ����g���C�J�n�҂�..."
                        Call RaiseMessage(msg)

                        Call F.Wait(ConnectRetryWait)

                        msg = RemoteIP & ":" & RemotePort & " �֐ڑ��J�n ���g���C" & ConnectRetryDo & "���"
                        Call RaiseMessage(msg)
                        Call RaiseConnectRetry(RemoteIP, RemotePort, ConnectRetryDo)

                        Call Connect() '�ڑ��������Ă�
                    End If

                End If
            Case SockStatuses.Connected
                ConnectCheckTimer.Stop()
                '���M�Ď��^�C�}�[���N��
                SendCheckTimer.Interval = 10
                SendCheckTimer.Start()

        End Select

        ConnectCheckTimerExecuting = False



    End Sub



    ''' <summary>
    ''' �\�P�b�g�ڑ�
    ''' </summary>
    ''' <remarks>�\�P�b�g�ڑ�</remarks>
    Private Sub Connect()
        Try
            SendStatus = SockStatuses.Connecting
            ConnectTimeLimit = DateTime.Now.AddMilliseconds(ConnectTimeOut)


            Dim cipAddress As IPAddress = IPAddress.Parse(RemoteIP)
            Dim remoteEP As New IPEndPoint(cipAddress, CInt(RemotePort))

            ' Create a TCP/IP socket.
            client = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)


            Try
                Dim msg As String = RemoteIP & ":" & RemotePort & " �֐ڑ��J�n �^�C���A�E�g->" & Format(ConnectTimeLimit, "yyyy/MM/dd HH:mm:ss.ff")
                Call RaiseMessage(msg)
                client.BeginConnect(remoteEP, New AsyncCallback(AddressOf ConnectCallback), client)

            Catch e As Exception
                Dim msg As String = RemoteIP & ":" & RemotePort & " Connect�G���[:" & e.Message
                Call RaiseMessage(msg)
                If ConnectCheckTimer.Enabled Then '�ڑ��Ď��^�C�}�[�������Ă���Ԃ̓G���[�C�x���g���N����
                    Call RaiseError(msg, SockErrors.ConnectErr)
                End If

                Exit Sub
            End Try

        Catch ex As Exception
            Dim msg As String = RemoteIP & ":" & RemotePort & " �\�P�b�g�ڑ��ł��܂���:" & ex.Message
            Call RaiseMessage(msg)
            If ConnectCheckTimer.Enabled Then '�ڑ��Ď��^�C�}�[�������Ă���Ԃ̓G���[�C�x���g���N����
                Call RaiseError(msg, SockErrors.ConnectErr)
            End If

        End Try

    End Sub

    ''' <summary>
    ''' �\�P�b�g�ڑ��R�[���o�b�N
    ''' </summary>
    ''' <param name="ar">�\�P�b�g</param>
    ''' <remarks>�\�P�b�g�ڑ��R�[���o�b�N</remarks>
    Private Sub ConnectCallback(ByVal ar As IAsyncResult)
        Try
            ' Retrieve the socket from the state object.
            Dim client As Socket = CType(ar.AsyncState, Socket)

            Try
                ' Complete the connection.
                client.EndConnect(ar)
            Catch ex As Exception
                'Connect���ɃN���[�Y���ꂽ�ꍇ�̏���
                Dim msg As String = RemoteIP & ":" & RemotePort & " ConnectCallback�G���[:" & ex.Message
                Call RaiseMessage(msg)
                If ConnectCheckTimer.Enabled Then '�ڑ��Ď��^�C�}�[�������Ă���Ԃ̓G���[�C�x���g���N����
                    Call RaiseError(msg, SockErrors.ConnectErr)
                End If
                Return
            End Try

            SendStatus = SockStatuses.Connected

            Dim localIP As String = CType(client.LocalEndPoint, IPEndPoint).Address.ToString()
            Dim localPort As String = CType(client.LocalEndPoint, IPEndPoint).Port.ToString()

            Call RaiseConnectEnd(RemoteIP, RemotePort, localIP, localPort)

            '��M�J�n
            Call RaiseMessage("RECEIVE�J�n:" & localIP & "(" & localPort & ")")
            Dim l = GetStateIndex(RemoteIP, RemotePort, localIP, localPort)
            With tState(l)
                .WorkSocket = client
            End With
            client.BeginReceive(tState(l).Buffer, 0, StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReadCallback), tState(l))


        Catch ex As Exception
            Dim msg As String = RemoteIP & ":" & RemotePort & " ConnectCallback�G���[:" & ex.Message
            Call RaiseMessage(msg)
            If ConnectCheckTimer.Enabled Then '�ڑ��Ď��^�C�}�[�������Ă���Ԃ̓G���[�C�x���g���N����
                Call RaiseError(msg, SockErrors.ConnectErr)
            End If

        End Try

    End Sub 'ConnectCallback


    ''' <summary>
    ''' �\�P�b�g�ؒf����
    ''' </summary>
    ''' <remarks>�\�P�b�g�ؒf����</remarks>
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
    ''' ���M�J�n����
    ''' </summary>
    ''' <param name="SendParm">���M�p�����[�^</param>
    ''' <remarks>���M�J�n����</remarks>
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
                msg = "(SockSend)LocalIP�擾�G���[:" & ex.Message
                Call RaiseMessage(msg)
            End Try

            msg = "���M�҂��ɒǉ����܂���:" & localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & _
             " Key:" & SendParm.SendKEY & " ���M���e:" & F.Byte2String(SendParm.sndCmd)
            Call RaiseMessage(msg)

        End SyncLock
    End Sub

    ''' <summary>
    ''' ���M�Ď��^�C�}�[����
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>���M�Ď��^�C�}�[����</remarks>
    Private Sub SendCheckTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles SendCheckTimer.Elapsed
        If SendCheckTimerExecuting Then
            '�^�C�}�[�������s���ł����return
            Return
        End If

        If RecvStatus = SockStatuses.Receiving Then
            '��M�������ł����return
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
                Dim msg As String = "(SendCheckTimer_Elapsed)LocalIP�擾�G���[:" & ex.Message
                Call RaiseMessage(msg)
            End Try

            Select Case SendStatus
                Case SockStatuses.Sended, SockStatuses.Connected
                    '���M�҂��s����`�F�b�N���đ��M�ς݂�����΍폜����
                    SendCmdList.RemoveAll(AddressOf IsRemoveSendParm)

                    '���M�҂��s����`�F�b�N���đ҂�������Α��M�J�n����
                    For l As Integer = 0 To SendCmdList.Count - 1
                        If SendCmdList(l).IsSendEnd = False Then

                            SendingIndex = l

                            Dim msg As String = "���M�҂����� " & localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " SendingIndex:" & SendingIndex
                            Call RaiseMessage(msg)

                            Call Send()
                            Exit For
                        End If
                    Next

                Case SockStatuses.Sending
                    If localIP = "0.0.0.0" Then
                        Dim msg As String = String.Empty
                        msg = localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " �Z�b�V�������؂�Ă���̂Ń��g���C�I�[�o�[�����Ƃ��܂� " & " SendingIndex:" & SendingIndex
                        Call RaiseMessage(msg)

                        SendCmdList(SendingIndex).SendTimeLimit = DateTime.Now
                        SendCmdList(SendingIndex).SendRetryDo = SendCmdList(SendingIndex).SendRetry

                    End If


                    '���M�^�C���A�E�g�`�F�b�N
                    If DateTime.Now >= SendCmdList(SendingIndex).SendTimeLimit Then
                        If SendCmdList(SendingIndex).SendRetryDo = SendCmdList(SendingIndex).SendRetry Then
                            '���g���C�I�[�o�[

                            Dim msg As String = localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " ���M���g���C�I�[�o" & " SendingIndex:" & SendingIndex

                            Call RaiseError(msg, SockErrors.SendRetryOver, RemoteIP, RemotePort, SendCmdList(SendingIndex).SendKEY)

                            SendCheckTimer.Stop()

                            Call SockConnect() '�ڑ��������Ă�

                        Else
                            '���g���C
                            SendCmdList(SendingIndex).SendRetryDo += 1
                            Dim msg As String = String.Empty
                            msg = localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " ���M���g���C�J�n�҂�..." & " SendingIndex:" & SendingIndex
                            Call RaiseMessage(msg)

                            Call F.Wait(SendCmdList(SendingIndex).SendRetryWait)

                            msg = localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " �֑��M�J�n ���g���C" & SendCmdList(SendingIndex).SendRetryDo & "���" & " SendingIndex:" & SendingIndex
                            Call RaiseMessage(msg)
                            Call RaiseSendRetry(RemoteIP, RemotePort, localIP, localPort, SendCmdList(SendingIndex).SendRetryDo, SendCmdList(SendingIndex).SendKEY)

                            Call Send() '���M�������Ă�
                        End If

                    End If

            End Select


        Catch ex As Exception
            Dim msg As String = RemoteIP & ":" & RemotePort & " SendCheckTimer_Elapsed�G���[:" & ex.Message
            Call RaiseMessage(msg)

        End Try
        SendCheckTimerExecuting = False


    End Sub

    ''' <summary>
    ''' ���M�p�����[�^���폜���ėǂ����`�F�b�N����
    ''' </summary>
    ''' <param name="sendParm">���M�p�����[�^</param>
    ''' <returns>True:�폜�� False:�폜�s��</returns>
    ''' <remarks>���M�p�����[�^���폜���ėǂ����`�F�b�N����</remarks>
    Private Function IsRemoveSendParm(ByVal sendParm As SendParmClass) As Boolean
        If sendParm.IsRecvEnd Then
            '��M�����܂Ŋ������Ă���Ώ����ėǂ�
            Return True
        Else
            If sendParm.IsRecv = False And sendParm.IsSendEnd Then
                '��M�����s�v�ő��M�������Ă���Ώ����ėǂ�
                Return True
            End If

            If sendParm.IsSendEnd And sendParm.IsRecv And DateTime.Now >= sendParm.RecvTimeLimit Then
                '��M�҂��^�C���A�E�g�ɂȂ��Ă���Ώ����ėǂ�
                Return True
            End If


            Return False
        End If
    End Function

    ''' <summary>
    ''' �\�P�b�g���M
    ''' </summary>
    ''' <remarks>�\�P�b�g���M</remarks>
    Private Sub Send()
        Try


            SendStatus = SockStatuses.Sending
            SendCmdList(SendingIndex).SendTimeLimit = DateTime.Now.AddMilliseconds(SendCmdList(SendingIndex).SendTimeOut)


            Dim localIP As String = CType(client.LocalEndPoint, IPEndPoint).Address.ToString()
            Dim localPort As String = CType(client.LocalEndPoint, IPEndPoint).Port.ToString()

            Dim msg As String = localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " ���M�J�n �^�C���A�E�g->" & Format(SendCmdList(SendingIndex).SendTimeLimit, "yyyy/MM/dd HH:mm:ss.ff") & " SendingIndex:" & SendingIndex
            Call RaiseMessage(msg)

            client.BeginSend(SendCmdList(SendingIndex).sndCmd, 0, SendCmdList(SendingIndex).sndCmd.Length, 0, New AsyncCallback(AddressOf SendCallback), client)

        Catch ex As Exception
            Dim msg As String = RemoteIP & ":" & RemotePort & " BeginSend�G���[:" & ex.Message
            Call RaiseMessage(msg)
            Call RaiseError(msg, SockErrors.SendErr, RemoteIP, RemotePort)

        End Try
    End Sub 'Send

    ''' <summary>
    ''' �\�P�b�g���M�R�[���o�b�N
    ''' </summary>
    ''' <param name="ar">�\�P�b�g</param>
    ''' <remarks>�\�P�b�g���M�R�[���o�b�N</remarks>
    Private Sub SendCallback(ByVal ar As IAsyncResult)
        Dim pos As Integer = 0

        Try
            pos = 100


            pos = 110
            '��M�҂��̕K�v������΃X�e�[�^�X�ύX���s���i���̃^�C�~���O�ōs��Ȃ��ƃX�e�[�^�X�ύX�O��Receive����邱�Ƃ�����j
            If SendCmdList(SendingIndex).IsRecv Then
                pos = 120
                '�X�e�[�^�X�ύX
                RecvStatus = SockStatuses.Receiving
            End If

            Dim client As Socket = CType(ar.AsyncState, Socket)
            Dim bytesSent As Integer = client.EndSend(ar)

            pos = 200

            Dim localIP As String = CType(client.LocalEndPoint, IPEndPoint).Address.ToString()
            Dim localPort As String = CType(client.LocalEndPoint, IPEndPoint).Port.ToString()

            pos = 300

            If localIP.Equals("0.0.0.0") Then
                '�Z�b�V�������؂�Ă���
                Dim msg As String = localIP & ":" & localPort & " - " & RemoteIP & ":" & RemotePort & " �Ԃ̐ڑ����؂�Ă��܂�"

                RecvStatus = SockStatuses.Received '��Ŏ�M���ɂ��Ă���̂ŁA��M�����ɖ߂��Ă���

                Call RaiseMessage(msg)
                If SendingIndex <= SendCmdList.Count - 1 Then
                    Call RaiseError(msg, SockErrors.SendErr, RemoteIP, RemotePort, SendCmdList(SendingIndex).SendKEY)
                End If
                Return

            End If

            pos = 400
            Call RaiseSendEnd(RemoteIP, RemotePort, localIP, localPort, F.Byte2HexChar(SendCmdList(SendingIndex).sndCmd), SendCmdList(SendingIndex).SendKEY)

            '���b�Z�[�W�\��
            Call RaiseMessage(localIP & ":" & localPort & "->" & RemoteIP & ":" & RemotePort & " " & _
                              bytesSent & "�o�C�g���M-> " & F.Byte2String(SendCmdList(SendingIndex).sndCmd) & "(" & F.Byte2HexChar(SendCmdList(SendingIndex).sndCmd) & ")" & " SendingIndex:" & SendingIndex)


            pos = 500
            '��M�҂��̕K�v������Ύ�M�҂����s��
            If SendCmdList(SendingIndex).IsRecv Then
                '��M�Ď��^�C�}�[���N��
                SendCmdList(SendingIndex).RecvTimeLimit = DateTime.Now.AddMilliseconds(SendCmdList(SendingIndex).RecvTimeOut)
                Dim msg As String = RemoteIP & ":" & RemotePort & "->" & localIP & ":" & localPort & " �߂�҂��J�n �^�C���A�E�g->" & Format(SendCmdList(SendingIndex).RecvTimeLimit, "yyyy/MM/dd HH:mm:ss.ff") & " SendingIndex:" & SendingIndex
                Call RaiseMessage(msg)

                pos = 600

                '���̃^�C�~���O�ŃX�e�[�^�X�ύX�͂��߁BRecvStatus = SockStatuses.Receiving
                RecvCheckTimer.Interval = 10
                RecvCheckTimer.Start()

            End If

            pos = 700


            '���M������Ԃɂ���
            SendCmdList(SendingIndex).IsSendEnd = True
            SendStatus = SockStatuses.Sended

            pos = 800
        Catch ex As Exception
            Dim msg As String = RemoteIP & ":" & RemotePort & " SendCallback�G���[:" & ex.Message & " pos=" & pos
            Call RaiseMessage(msg)

            If SendingIndex <= SendCmdList.Count - 1 Then
                Call RaiseError(msg, SockErrors.SendErr, RemoteIP, RemotePort, SendCmdList(SendingIndex).SendKEY)
            Else
                'Call RaiseError( msg, SockErrors.SendErr, RemoteIP, RemotePort, "Key�s��")
                msg = "SendCallback�G���[ ���łɃ\�P�b�g���N���[�Y����Ă��܂�"
                Call RaiseMessage(msg)
                Return
            End If

        End Try

    End Sub 'SendCallback

    ''' <summary>
    ''' �w�肳�ꂽIP�̃X�e�[�g�̃C���f�b�N�X��Ԃ��B�Ȃ���ΐV���ɍ쐬����
    ''' </summary>
    ''' <param name="remoteIP">�w��IP</param>
    ''' <param name="remotePort">�w��PORT</param>
    ''' <param name="localIP">��IP</param>
    ''' <param name="localPort">��PORT</param>
    ''' <returns>�C���f�b�N�X</returns>
    ''' <remarks>�w�肳�ꂽ�\�P�b�g�̃C���f�b�N�X��Ԃ�</remarks>
    Private Function GetStateIndex(ByVal remoteIP As String, ByVal remotePort As String, ByVal localIP As String, ByVal localPort As String) As Long
        Dim lRet As Long
        Dim l As Long
        lRet = -999


        Try
            '�w�肳�ꂽIP�̃X�e�[�g�����݂��Ă��邩���ׂ�
            For l = LBound(tState) To UBound(tState)
                If tState(l) Is Nothing Then
                Else
                    If tState(l).RemoteIP = remoteIP Then
                        '���蓖�čς݂Ȃ̂ł��̃X�e�[�g���g��
                        tState(l).RemotePort = remotePort
                        tState(l).LocalIP = localIP
                        tState(l).LocalPort = localPort

                        lRet = l
                        Exit For
                    End If
                End If
            Next

            If lRet = -999 Then
                '�w�肳�ꂽIP�̃X�e�[�g�͑��݂��Ă��Ȃ������̂ŋ󂫃X�e�[�g��T��
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

                '�����l
                ReDim tState(lRet).bRecv(0)
                tState(lRet).RemoteIP = remoteIP
                tState(lRet).RemotePort = remotePort
                tState(lRet).LocalIP = localIP
                tState(lRet).LocalPort = localPort
                tState(lRet).Sb = New StringBuilder

            End If

        Catch ex As Exception
            Dim msg As String = "GetStateIndex�G���[:" & remoteIP & ":" & remotePort & " " & ex.Message
            Call RaiseMessage(msg)

        End Try


        Return lRet
    End Function


    ''' <summary>
    ''' ��M�R�[���o�b�N
    ''' </summary>
    ''' <param name="ar">�\�P�b�g</param>
    ''' <remarks>��M�R�[���o�b�N</remarks>
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
                Dim msg As String = "[ReadCallback]" & state.LocalIP & ":" & state.LocalPort & " ��Close�ς݂ł�"
                Call RaiseMessage(msg)
                Return
            Catch ex As Exception
                'RECEIVE���ɃN���[�Y���ꂽ�ꍇ�̏���
                Dim msg As String = state.LocalIP & ":" & state.LocalPort & " EndReceive�G���[:" & ex.Message
                Call RaiseMessage(msg)
                Return
            End Try


            pos = 200

            If bytesRead = 0 Then
                '�T�[�o�[������ؒf���ꂽ
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
                '���b�Z�[�W�\��
                Call RaiseMessage(state.RemoteIP & ":" & state.RemotePort & "->" & state.LocalIP & ":" & state.LocalPort & " " & content.Length.ToString & "�o�C�g��M-> " & F.Byte2String(state.bRecv) & "(" & F.Byte2HexChar(state.bRecv) & ")" & " SendingIndex:" & SendingIndex)

            End If

            pos = 300

            If content.Length > StateObject.BufferSize Then
                '���b�Z�[�W�\��
                Call RaiseMessage(state.RemoteIP & "����o�b�t�@�T�C�Y�𒴂���f�[�^���󂯎��܂����B��M�f�[�^��j�����܂� : " + F.Byte2HexChar(state.bRecv) & " SendingIndex:" & SendingIndex)
                ReDim state.bRecv(0) '��M�f�[�^�̏�����
                state.Sb = New StringBuilder

            Else
                If EndCheckChar.Length = 0 Then
                    IsReadEnd = True
                Else
                    '�I�����蕶�������M���Ă���Ί���
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
                    '�����̃��b�Z�[�W���������Ă��邱�Ƃ�����̂ŃX�v���b�g����
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
                            '��M�����C�x���g
                            Call RaiseRecvEnd(state.RemoteIP, state.RemotePort, state.LocalIP, state.LocalPort, F.Byte2HexChar(F.String2Byte(recvCmds(cmdix).ToString & kugiri)), SendCmdList(SendingIndex).SendKEY)
                            pos = 433

                            isRecvEnd = True
                            Exit For
                        Else
                            pos = 440
                            '��M�����C�x���g
                            Call RaiseRecvEnd(state.RemoteIP, state.RemotePort, state.LocalIP, state.LocalPort, F.Byte2HexChar(F.String2Byte(recvCmds(cmdix).ToString & kugiri)), "Key�s��" & cmdix)

                        End If

                    End If
                Next

                'If isRecvEnd Then
                '    '��M����
                '    SendCmdList(SendingIndex).IsRecvEnd = True
                '    RecvStatus = SockStatuses.Received
                'End If


                pos = 450
                ReDim state.bRecv(0) '��M�f�[�^�̏�����
                state.Sb = New StringBuilder

            End If

            pos = 500


            Try
                handler.BeginReceive(state.Buffer, 0, StateObject.BufferSize, 0, New AsyncCallback(AddressOf ReadCallback), state)
            Catch ex As Exception
                Dim msg As String = RemoteIP & ":" & RemotePort & " (��)BeginReceive�G���[:" & ex.Message
                Call RaiseMessage(msg)
            End Try
        Catch ex As Exception
            Dim msg As String = state.LocalIP & ":" & state.LocalPort & " ReadCallback�G���[:" & ex.Message & " pos:" & pos
            Call RaiseMessage(msg)

        End Try
    End Sub 'ReadCallback


    ''' <summary>
    ''' ��M�Ď��^�C�}�[����
    ''' </summary>
    ''' <param name="sender">�C�x���g������</param>
    ''' <param name="e">�C�x���g����</param>
    ''' <remarks>��M�Ď��^�C�}�[����</remarks>
    Private Sub RecvCheckTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles RecvCheckTimer.Elapsed
        If RecvCheckTimerExecuting Then
            '�^�C�}�[�������s���ł����return
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
                Dim msg As String = RemoteIP & ":" & RemotePort & " LocalIP�擾�G���[:" & ex.Message
                Call RaiseMessage(msg)
            End Try

            Select Case RecvStatus
                Case SockStatuses.Received
                    Dim msg As String = localIP & ":" & localPort & " -> " & RemoteIP & ":" & RemotePort & " ���ʎ�M���� SendingIndex:" & SendingIndex
                    Call RaiseMessage(msg)

                    RecvCheckTimer.Stop() '��M���������̂Ń^�C�}�[���~�߂�

                Case SockStatuses.Receiving
                    '��M�^�C���A�E�g�`�F�b�N
                    If DateTime.Now >= SendCmdList(SendingIndex).RecvTimeLimit Then
                        '�^�C���A�E�g

                        Dim msg As String = RemoteIP & ":" & RemotePort & "->" & localIP & ":" & localPort & " ��M�҂��^�C���A�E�g SendingIndex:" & SendingIndex

                        Call RaiseError(msg, SockErrors.RecvErr)

                        RecvStatus = SockStatuses.Received
                        RecvCheckTimer.Stop()

                    End If

            End Select


        Catch ex As Exception
            Dim msg As String = RemoteIP & ":" & RemotePort & " RecvCheckTimer_Elapsed�G���[:" & ex.Message
            Call RaiseMessage(msg)

        End Try
        RecvCheckTimerExecuting = False

    End Sub
End Class 'clsSndSock

Public Class F
    ''' <summary>
    ''' �o�C�g�z���hex������� ��)000112FF�iF.Byte2HexChar�j
    ''' </summary>
    ''' <param name="bytearr">�o�C�g�z��</param>
    ''' <returns>input��hex������ɂ�������</returns>
    ''' <remarks>�o�C�g�z���hex������� ��)000112FF�iF.Byte2HexChar�j</remarks>
    Public Shared Function Byte2HexChar(ByVal bytearr() As Byte) As String
        Dim l As Long           '���[�v�Y����
        Dim result As String    '�߂�l

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
    ''' hex��������o�C�g�z���
    ''' </summary>
    ''' <param name="hexString">hex������</param>
    ''' <returns>�o�C�g�z��</returns>
    ''' <remarks>hex��������o�C�g�z���</remarks>
    Public Shared Function HexChar2Byte(ByVal hexString As String) As Byte()
        Dim l As Long           '���[�v�Y����
        Dim result(0) As Byte   '�߂�l

        For l = 0 To hexString.Length - 1 Step 2
            If UBound(result) < l / 2 Then
                ReDim Preserve result(0 To l / 2)
            End If
            result(l / 2) = Convert.ToInt16(hexString.Substring(l, 2), 16)
        Next

        Return result
    End Function


    ''' <summary>
    ''' �o�C�g�z���String�������
    ''' </summary>
    ''' <param name="bytearr">�o�C�g�z��</param>
    ''' <returns>String������</returns>
    ''' <remarks>�o�C�g�z���String�������</remarks>
    Public Shared Function Byte2String(ByVal bytearr() As Byte) As String
        Dim l As Long           '���[�v�Y����
        Dim result As String    '�߂�l

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
    ''' String��������o�C�g�z���
    ''' </summary>
    ''' <param name="Value">String������</param>
    ''' <returns>�o�C�g�z��</returns>
    ''' <remarks>String��������o�C�g�z���</remarks>
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
    ''' �҂�����
    ''' </summary>
    ''' <param name="msec">�҂��ԁi�~���b�j</param>
    ''' <remarks>�҂�����</remarks>
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
