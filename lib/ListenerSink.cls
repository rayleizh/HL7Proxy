VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListenerSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Parent form (ProxyMain) must implement:
' Friend Sub EndpointAccepted(cli As cTlsSocket, backendHost As String, backendPort As Long, _
'                             backendUseTls As Boolean, clientPfxPath As String, _
'                             clientPfxPassword As String, nameTag As String)

Public Parent As ProxyMain

Private WithEvents EP As EndpointListener
Attribute EP.VB_VarHelpID = -1

' Backend routing for this endpoint
Public BackendHost As String
Public BackendPort As Long
Public BackendUseTls As Boolean
Public ClientPfxPath As String
Public ClientPfxPassword As String
Public nameTag As String

Public Sub Bind(ByVal endpoint As EndpointListener)
    Set EP = endpoint
End Sub

' Must exactly match the EndpointListener.Accepted event signature.
Private Sub EP_Accepted(ByVal cli As cTlsSocket, ByVal nameTag As String)
    If Parent Is Nothing Then Exit Sub
    Parent.EndpointAccepted cli, BackendHost, BackendPort, BackendUseTls, ClientPfxPath, ClientPfxPassword, nameTag
End Sub


