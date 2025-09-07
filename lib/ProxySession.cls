VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProxySession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Front As cTlsSocket  ' accepted client (TLS or plain)
Attribute Front.VB_VarHelpID = -1
Private WithEvents Back  As cTlsSocket  ' outbound to backend (plain or TLS)
Attribute Back.VB_VarHelpID = -1
Private mParent As ProxyMain            ' the form (see below)
Private mBackReady As Boolean
Private mQueueToBack As String          ' temp buffer until Back connects

Friend Sub Init(ByVal accepted As cTlsSocket, ByVal parent As ProxyMain, _
                ByVal backendHost As String, ByVal backendPort As Long, _
                Optional ByVal backendUseTls As Boolean = False, _
                Optional ByVal clientPfxPath As String = "", _
                Optional ByVal clientPfxPassword As String = "")
    Set Front = accepted
    Set mParent = parent

    Set Back = New cTlsSocket
    mBackReady = False
    mQueueToBack = vbNullString

        If backendUseTls And Len(clientPfxPath) > 0 Then
        Dim certs As New Collection
        certs.Add clientPfxPath
        Set Back.LocalCertificates = certs
        Dim keys As New Collection
        If Len(clientPfxPassword) > 0 Then keys.Add clientPfxPassword
        Set Back.LocalPrivateKey = keys
    End If

    If Not Back.Connect(backendHost, backendPort, UseTls:=backendUseTls) Then
        mParent.Trace "backend connect failed"
        CloseBoth
    End If
End Sub


Friend Sub ShutDown()
    CloseBoth
End Sub

Private Sub Back_OnConnect()
    mBackReady = True
    If LenB(mQueueToBack) > 0 Then
        Dim a() As Byte
        a = StrConv(mQueueToBack, vbFromUnicode)
        Back.SendArray a
        mQueueToBack = vbNullString
    End If
End Sub

Private Sub Front_OnReceive()
    Dim a() As Byte
    If Not Front.ReceiveArray(a) Then Exit Sub
    If mBackReady Then
        Back.SendArray a
    Else
        ' buffer until back-end connects
        mQueueToBack = mQueueToBack & StrConv(a, vbUnicode)
    End If
End Sub

Private Sub Back_OnReceive()
    Dim a() As Byte
    If Not Back.ReceiveArray(a) Then Exit Sub
    Front.SendArray a
End Sub

Private Sub Front_OnClose()
    CloseBoth
End Sub

Private Sub Back_OnClose()
    CloseBoth
End Sub

Private Sub Front_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    mParent.Trace "front error " & CStr(ErrorCode)
    CloseBoth
End Sub

Private Sub Back_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    mParent.Trace "back error " & CStr(ErrorCode)
    CloseBoth
End Sub

Private Sub CloseBoth()
    On Error Resume Next
    If Not Front Is Nothing Then Front.Close_
    If Not Back Is Nothing Then Back.Close_
    mParent.SessionClosed Me
    Set Front = Nothing
    Set Back = Nothing
    Set mParent = Nothing
End Sub

