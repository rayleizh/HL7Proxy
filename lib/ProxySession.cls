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

Private WithEvents Front As cTlsSocket
Attribute Front.VB_VarHelpID = -1
Private WithEvents Back As cTlsSocket
Attribute Back.VB_VarHelpID = -1
Private mParent As ProxyMain
Private mBackReady As Boolean
Private mQueueToBack As String
Private mNameTag As String

Friend Sub Init(ByVal accepted As cTlsSocket, ByVal parent As ProxyMain, _
                ByVal backendHost As String, ByVal backendPort As Long, _
                ByVal nameTag As String, _
                Optional ByVal backendUseTls As Boolean = False, _
                Optional ByVal clientPfxPath As String = "", _
                Optional ByVal clientPfxPassword As String = "")
    Set Front = accepted
    Set mParent = parent
    Set Back = New cTlsSocket
    mBackReady = False
    mNameTag = nameTag
    mQueueToBack = vbNullString

    mParent.Trace "Initiating connection to " & backendHost & ":" & backendPort & " (TLS=" & backendUseTls & ") for " & mNameTag
    If backendUseTls Then
        mParent.Trace "TLS enabled for backend, no client certificate used, LocalFeatures=ucsTlsSupportTls12 for " & mNameTag
    End If
    On Error Resume Next
    If Not Back.Connect(backendHost, backendPort, UseTls:=backendUseTls, LocalFeatures:=ucsTlsSupportTls12) Then
        mParent.Trace "ERROR: Backend connect failed for " & mNameTag & " (" & backendHost & ":" & backendPort & "), Error: " & Err.Description
        CloseBoth
    End If
    On Error GoTo 0
End Sub

Friend Sub ShutDown()
    CloseBoth
End Sub

Private Sub Back_OnConnect()
    mBackReady = True
    mParent.Trace "Backend connected for " & mNameTag
    If LenB(mQueueToBack) > 0 Then
        Dim a() As Byte
        a = StrConv(mQueueToBack, vbFromUnicode)
        On Error Resume Next
        Back.SendArray a
        If Err.Number = 0 Then
            mParent.Trace "Sent " & UBound(a) + 1 & " bytes to backend for " & mNameTag & ": " & Left(mQueueToBack, 100) & "..."
            mParent.Trace "Raw bytes sent: " & ByteArrayToHex(a)
            mQueueToBack = vbNullString
        Else
            mParent.Trace "ERROR: Failed to send buffered data to backend for " & mNameTag & ": " & Err.Description
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub Front_OnReceive()
    Dim a() As Byte
    If Not Front.ReceiveArray(a) Then
        mParent.Trace "ERROR: Front.ReceiveArray failed for " & mNameTag
        Exit Sub
    End If
    mParent.Trace "Front received " & UBound(a) + 1 & " bytes for " & mNameTag & ": " & Left(StrConv(a, vbUnicode), 100) & "..."
    mParent.Trace "Raw bytes received: " & ByteArrayToHex(a)
    If EndsWithFS_CR_Bytes(a) Then
        mParent.Trace "Valid MLLP frame detected for " & mNameTag
        If mBackReady Then
            On Error Resume Next
            Back.SendArray a
            If Err.Number = 0 Then
                mParent.Trace "Sent " & UBound(a) + 1 & " bytes to backend for " & mNameTag & ": " & Left(StrConv(a, vbUnicode), 100) & "..."
                mParent.Trace "Raw bytes sent: " & ByteArrayToHex(a)
            Else
                mParent.Trace "ERROR: Failed to send to backend for " & mNameTag & ": " & Err.Description
                mQueueToBack = mQueueToBack & StrConv(a, vbUnicode)
                mParent.Trace "Buffered " & Len(mQueueToBack) & " characters for " & mNameTag & ": " & Left(mQueueToBack, 100) & "..."
            End If
            On Error GoTo 0
        Else
            mQueueToBack = mQueueToBack & StrConv(a, vbUnicode)
            mParent.Trace "Buffered " & Len(mQueueToBack) & " characters for " & mNameTag & ": " & Left(mQueueToBack, 100) & "..."
        End If
    Else
        mQueueToBack = mQueueToBack & StrConv(a, vbUnicode)
        mParent.Trace "Incomplete MLLP frame, buffering " & Len(mQueueToBack) & " characters for " & mNameTag & ": " & Left(mQueueToBack, 100) & "..."
    End If
End Sub

Private Sub Back_OnReceive()
    Dim a() As Byte
    If Not Back.ReceiveArray(a) Then
        mParent.Trace "ERROR: Back.ReceiveArray failed for " & mNameTag
        Exit Sub
    End If
    mParent.Trace "Back received " & UBound(a) + 1 & " bytes for " & mNameTag & ": " & Left(StrConv(a, vbUnicode), 100) & "..."
    mParent.Trace "Raw bytes received: " & ByteArrayToHex(a)
    If EndsWithFS_CR_Bytes(a) Then
        mParent.Trace "Valid MLLP frame detected for " & mNameTag
        On Error Resume Next
        Front.SendArray a
        If Err.Number = 0 Then
            mParent.Trace "Sent " & UBound(a) + 1 & " bytes to front for " & mNameTag & ": " & Left(StrConv(a, vbUnicode), 100) & "..."
            mParent.Trace "Raw bytes sent: " & ByteArrayToHex(a)
        Else
            mParent.Trace "ERROR: Failed to send to front for " & mNameTag & ": " & Err.Description
        End If
        On Error GoTo 0
    Else
        mParent.Trace "Incomplete MLLP frame from backend, buffering " & Len(mQueueToBack) & " characters for " & mNameTag & ": " & Left(mQueueToBack, 100) & "..."
    End If
End Sub

Private Sub Front_OnClose()
    mParent.Trace "Front closed for " & mNameTag
    CloseBoth
End Sub

Private Sub Back_OnClose()
    mParent.Trace "Back closed for " & mNameTag
    CloseBoth
End Sub

Private Sub Front_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    mParent.Trace "Front error " & ErrorCode & " (EventMask=" & EventMask & ") for " & mNameTag
    CloseBoth
End Sub

Private Sub Back_OnError(ByVal ErrorCode As Long, ByVal EventMask As UcsAsyncSocketEventMaskEnum)
    mParent.Trace "Back error " & ErrorCode & " (EventMask=" & EventMask & ") for " & mNameTag
    CloseBoth
End Sub

Private Sub CloseBoth()
    On Error Resume Next
    If Not Front Is Nothing Then Front.Close_
    If Not Back Is Nothing Then Back.Close_
    mParent.Trace "Closed both sockets for " & mNameTag
    mParent.SessionClosed Me
    Set Front = Nothing
    Set Back = Nothing
    Set mParent = Nothing
End Sub

Private Function EndsWithFS_CR_Bytes(b() As Byte) As Boolean
    On Error Resume Next
    If UBound(b) < 0 Then
        mParent.Trace "MLLP check failed: Byte array empty for " & mNameTag
        EndsWithFS_CR_Bytes = False
        Exit Function
    End If
    Dim lastByte As Byte, secondLastByte As Byte
    If UBound(b) >= 1 Then
        secondLastByte = b(UBound(b) - 1)
        lastByte = b(UBound(b))
    ElseIf UBound(b) = 0 Then
        secondLastByte = 0
        lastByte = b(0)
    End If
    mParent.Trace "MLLP check: Last byte=" & Hex(lastByte) & ", Second last byte=" & Hex(secondLastByte) & ", Array length=" & UBound(b) + 1 & " for " & mNameTag
    EndsWithFS_CR_Bytes = (secondLastByte = &H1C And lastByte = &HD)
    On Error GoTo 0
End Function

Private Function ByteArrayToHex(b() As Byte) As String
    Dim i As Long
    Dim s As String
    If UBound(b) < 0 Then
        ByteArrayToHex = "(empty)"
        Exit Function
    End If
    For i = 0 To UBound(b)
        s = s & Right("0" & Hex(b(i)), 2) & " "
    Next i
    ByteArrayToHex = Trim(s)
End Function

