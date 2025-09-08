VERSION 5.00
Begin VB.Form ProxyMain 
   Caption         =   "HL7Proxy"
   ClientHeight    =   420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   420
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
End
Attribute VB_Name = "ProxyMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' ========= Runtime state =========
Private Sessions As Collection          ' active ProxySession objects
Private InboundSinks As Collection      ' ListenerSink for inbound TLS listeners
Private OutboundSinks As Collection     ' ListenerSink for outbound plain listeners

' ========= Logging =========
Private LogPath As String

' ========= Form lifecycle =========
Private Sub Form_Load()
    On Error Resume Next

    Me.Caption = "HL7Proxy"

    Set Sessions = New Collection
    Set InboundSinks = New Collection
    Set OutboundSinks = New Collection

    ' Default log file (edit if you prefer)
    LogPath = LocalLogPath(App.EXEName & ".log")  ' e.g., HL7Proxy.exe -> HL7Proxy.log
    
    WriteLog "==== HL7Proxy starting ===="

    ' Spin up listeners for every HL7Server* service that is TLSProxyEnabled
    StartAllFromRegistry

    WriteLog "==== HL7Proxy ready ===="
End Sub

Private Function LocalLogPath(ByVal FileName As String) As String
    Dim base As String
    base = App.Path & "\Log\HL7Proxy"
    If Len(base) = 0 Then base = CurDir$      ' fallback, just in case
    If Right$(base, 1) <> "\" Then base = base & "\"
    LocalLogPath = base & FileName
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim i As Long

    ' Stop inbound listeners
    If Not InboundSinks Is Nothing Then
        For i = InboundSinks.Count To 1 Step -1
            InboundSinks(i).Parent = Nothing
            ' Each sink holds an EndpointListener; StopListen inside class
            ' (We don't need to call here because sockets close when form ends,
            ' but safe to have StopListen in class when destroyed.)
            InboundSinks.Remove i
        Next i
    End If

    ' Stop outbound listeners
    If Not OutboundSinks Is Nothing Then
        For i = OutboundSinks.Count To 1 Step -1
            OutboundSinks(i).Parent = Nothing
            OutboundSinks.Remove i
        Next i
    End If

    ' Close sessions
    If Not Sessions Is Nothing Then
        For i = Sessions.Count To 1 Step -1
            Sessions(i).ShutDown
            Sessions.Remove i
        Next i
    End If

    WriteLog "==== HL7Proxy stopped ===="
End Sub

' ========= Registry-driven startup =========
Private Sub StartAllFromRegistry()
    On Error Resume Next

    Dim services As Collection
    Set services = EnumMeridianServiceKeys()   ' from modRegEnum.bas

    If services Is Nothing Or services.Count = 0 Then
        WriteLog "*** No HL7Server* services found under HKLM\SOFTWARE\Meridian"
        Exit Sub
    End If

    Dim i As Long
    For i = 1 To services.Count
        Dim svc As String: svc = services(i)
        Dim base As String
        ' RegistryKey = "HKLM\Software\Meridian\" & RegistryKey & "\"
        base = "HKLM\SOFTWARE\Meridian\" & svc & "\"

        Dim enabled As Boolean
        enabled = (ReadRegistryRaw(base & "TLSProxyEnabled", "1") <> "0")
        If Not enabled Then
            WriteLog "TLSProxyEnabled=0 for " & svc & " — skipping"
            GoTo NextService
        End If

        ' Legacy ports your existing services already have:
        Dim inboundPort As Long
        Dim outboundPort As Long
        inboundPort = CLng(Val(ReadRegistryRaw(base & "InboundPort", "0")))
        outboundPort = CLng(Val(ReadRegistryRaw(base & "OutboundPort", "0")))

        ' New TLS proxy keys you specified:
        Dim tlsInboundPort As Long
        Dim tlsOutHost As String
        Dim tlsOutPort As Long
        Dim pfxPath As String
        Dim pfxPass As String

        tlsInboundPort = CLng(Val(ReadRegistryRaw(base & "TLSProxyInboundPort", "0")))
        tlsOutHost = ReadRegistryRaw(base & "TLSProxyOutboundHost", "")
        tlsOutPort = CLng(Val(ReadRegistryRaw(base & "TLSProxyOutboundPort", "0")))
        pfxPath = ReadRegistryRaw(base & "TLSCertificatePath", "")
        pfxPass = ReadRegistryRaw(base & "TLSCertificatePassword", "")

        ' Inbound TLS terminator — partner -> TLS :tlsInboundPort -> plain -> 127.0.0.1:InboundPort
        If tlsInboundPort > 0 And inboundPort > 0 Then
            StartInboundSvc svc, tlsInboundPort, pfxPath, pfxPass, inboundPort
        End If

        ' Outbound TLS wrapper — local HL7Outbound -> plain :OutboundPort -> TLS -> host:tlsOutPort
        If Len(tlsOutHost) > 0 And tlsOutPort > 0 And outboundPort > 0 Then
            StartOutboundSvc svc, outboundPort, tlsOutHost, tlsOutPort, pfxPath, pfxPass
        End If

NextService:
    Next i
End Sub

' ========= Listener provisioning =========
Private Sub StartInboundSvc(ByVal nameTag As String, ByVal tlsListenPort As Long, _
                            ByVal ServerPfxPath As String, ByVal ServerPfxPassword As String, _
                            ByVal backendLocalInboundPort As Long)
    On Error Resume Next

    Dim EP As New EndpointListener
    EP.nameTag = nameTag & "_IN"
    EP.Port = tlsListenPort
    EP.UseTls = True
    EP.ServerPfxPath = ServerPfxPath
    EP.ServerPfxPassword = ServerPfxPassword
    
    Dim sink As New ListenerSink
    Set sink.Parent = Me
    sink.nameTag = EP.nameTag
    sink.BackendHost = "127.0.0.1"
    sink.BackendPort = backendLocalInboundPort
    sink.BackendUseTls = False
    sink.ClientPfxPath = ""
    sink.ClientPfxPassword = ""

    sink.Bind EP

    If EP.UseTls And Len(ServerPfxPath) = 0 Then
        WriteLog "*** WARNING: TLS listener started without a certificate (" & nameTag & ")"
    End If
    
    If EP.Start Then
        InboundSinks.Add sink, EP.nameTag
        WriteLog "Inbound TLS listening : " & CStr(tlsListenPort) & _
                 "  ->  127.0.0.1:" & CStr(backendLocalInboundPort) & _
                 "  (" & nameTag & ")"
    Else
        WriteLog "*** Failed to start inbound TLS :" & CStr(tlsListenPort) & " (" & nameTag & ")"
    End If
End Sub

Private Sub StartOutboundSvc(ByVal nameTag As String, ByVal localPlainPort As Long, _
                             ByVal tlsHost As String, ByVal tlsPort As Long, _
                             ByVal ClientPfxPath As String, ByVal ClientPfxPassword As String)
    On Error Resume Next

    Dim EP As New EndpointListener
    EP.nameTag = nameTag & "_OUT"
    EP.Port = localPlainPort
    EP.UseTls = False   ' plain listener on localhost

    Dim sink As New ListenerSink
    Set sink.Parent = Me
    sink.nameTag = EP.nameTag
    sink.BackendHost = tlsHost
    sink.BackendPort = tlsPort
    sink.BackendUseTls = True       ' connect outward using TLS
    sink.ClientPfxPath = ClientPfxPath
    sink.ClientPfxPassword = ClientPfxPassword

    sink.Bind EP

    If EP.Start Then
        OutboundSinks.Add sink, EP.nameTag
        WriteLog "Outbound PLAINTEXT listening : 127.0.0.1:" & CStr(localPlainPort) & _
                 "  ->  TLS " & tlsHost & ":" & CStr(tlsPort) & _
                 IIf(Len(ClientPfxPath) > 0, " (mTLS cert provided)", "") & _
                 "  (" & nameTag & ")"
    Else
        WriteLog "*** Failed to start outbound plain :" & CStr(localPlainPort) & " (" & nameTag & ")"
    End If
End Sub

' ========= Session creation (called by ListenerSink on Accept) =========
Friend Sub EndpointAccepted(ByVal cli As cTlsSocket, _
                            ByVal BackendHost As String, ByVal BackendPort As Long, _
                            ByVal BackendUseTls As Boolean, _
                            ByVal ClientPfxPath As String, ByVal ClientPfxPassword As String, _
                            ByVal nameTag As String)
    On Error Resume Next

    Dim s As New ProxySession
    If BackendUseTls Then
        s.Init cli, Me, BackendHost, BackendPort, nameTag, True, ClientPfxPath, ClientPfxPassword
    Else
        s.Init cli, Me, BackendHost, BackendPort, nameTag
    End If
    Sessions.Add s
    WriteLog "Accepted client (" & nameTag & ")"
End Sub

' ========= Utilities called by ProxySession =========
Friend Sub Trace(ByVal s As String)
    WriteLog s
End Sub

Friend Sub SessionClosed(ByVal sess As ProxySession)
    On Error Resume Next
    Dim i As Long
    For i = Sessions.Count To 1 Step -1
        If Sessions(i) Is sess Then
            Sessions.Remove i
            Exit For
        End If
    Next i
End Sub

' ========= Registry + Logging helpers =========
Private Function ReadRegistryRaw(ByVal fullKeyAndValue As String, ByVal defaultValue As String) As String
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    ReadRegistryRaw = wsh.RegRead(fullKeyAndValue)
    If Err.Number <> 0 Then
        ReadRegistryRaw = defaultValue
        Err.Clear
    End If
End Function

Private Sub WriteLog(ByVal s As String)
    On Error Resume Next
    Dim line As String
    line = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & s
    Debug.Print line
    Dim ff As Integer
    ff = FreeFile
    Open LogPath For Append As #ff
    Print #ff, line
    Close #ff
End Sub

