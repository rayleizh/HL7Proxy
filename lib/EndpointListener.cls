VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndpointListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Raise when a client connects. The sink must implement EP_Accepted with this exact signature.
Public Event Accepted(ByVal cli As cTlsSocket, ByVal nameTag As String)

Private WithEvents Sock As cTlsSocket
Attribute Sock.VB_VarHelpID = -1

Private mPort As Long
Private mUseTls As Boolean
Private mServerPfxPath As String
Private mServerPfxPassword As String
Private mNameTag As String

Public Property Get Port() As Long: Port = mPort: End Property
Public Property Let Port(ByVal v As Long): mPort = v: End Property

Public Property Get UseTls() As Boolean: UseTls = mUseTls: End Property
Public Property Let UseTls(ByVal v As Boolean): mUseTls = v: End Property

Public Property Get ServerPfxPath() As String: ServerPfxPath = mServerPfxPath: End Property
Public Property Let ServerPfxPath(ByVal v As String): mServerPfxPath = v: End Property

Public Property Get ServerPfxPassword() As String: ServerPfxPassword = mServerPfxPassword: End Property
Public Property Let ServerPfxPassword(ByVal v As String): mServerPfxPassword = v: End Property

Public Property Get nameTag() As String: nameTag = mNameTag: End Property
Public Property Let nameTag(ByVal v As String): mNameTag = v: End Property

Public Function Start() As Boolean
    On Error GoTo oops
    Set Sock = New cTlsSocket
    
    If Not Sock.Create(mPort) Then GoTo oops
    If Not Sock.Listen() Then GoTo oops
    Start = True
    Exit Function
oops:
    Start = False
End Function

Public Sub StopListen()
    On Error Resume Next
    If Not Sock Is Nothing Then Sock.Close_
End Sub

Private Sub Sock_OnAccept()
   Dim cli As cTlsSocket
   
   ' 1) Always accept as plain first.
   If Not Sock.Accept(cli, UseTls:=False) Then Exit Sub
   
   
   ' 2) If TLS is desired, initialize it *now* from the PFX.
   If mUseTls Then
    If Not cli.InitServerTls(CertFile:=mServerPfxPath, Password:=mServerPfxPassword) Then
        cli.Close_
        Exit Sub
    End If
   End If
   
   ' 3) Hand off the connected (possibly TLS) socket.
   RaiseEvent Accepted(cli, mNameTag)
End Sub

