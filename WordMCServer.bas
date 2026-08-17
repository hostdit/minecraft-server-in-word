Attribute VB_Name = "WordMCServer"
Option Explicit

Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, ByRef lpWSAData As Any) As Long
Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare PtrSafe Function api_socket Lib "ws2_32.dll" Alias "socket" (ByVal af As Long, ByVal stype As Long, ByVal protocol As Long) As LongPtr
Private Declare PtrSafe Function api_bind Lib "ws2_32.dll" Alias "bind" (ByVal s As LongPtr, ByRef nm As SOCKADDR_IN, ByVal namelen As Long) As Long
Private Declare PtrSafe Function api_listen Lib "ws2_32.dll" Alias "listen" (ByVal s As LongPtr, ByVal backlog As Long) As Long
Private Declare PtrSafe Function api_accept Lib "ws2_32.dll" Alias "accept" (ByVal s As LongPtr, ByVal addr As LongPtr, ByVal addrlen As LongPtr) As LongPtr
Private Declare PtrSafe Function api_recv Lib "ws2_32.dll" Alias "recv" (ByVal s As LongPtr, ByRef buf As Any, ByVal ln As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function api_send Lib "ws2_32.dll" Alias "send" (ByVal s As LongPtr, ByRef buf As Any, ByVal ln As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function api_close Lib "ws2_32.dll" Alias "closesocket" (ByVal s As LongPtr) As Long
Private Declare PtrSafe Function api_ioctl Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal s As LongPtr, ByVal cmd As Long, ByRef argp As Long) As Long
Private Declare PtrSafe Function api_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal s As LongPtr, ByVal level As Long, ByVal optname As Long, ByRef optval As Any, ByVal optlen As Long) As Long
Private Declare PtrSafe Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer

Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByRef src As Any, ByVal ln As LongPtr)

Private Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Private Type DrawItem
    key As Long
    bx As Long
    by As Long
    bz As Long
    col As Long
End Type

Private Const AF_INET As Long = 2
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const FIONBIO As Long = &H8004667E
Private Const SOL_SOCKET As Long = &HFFFF&
Private Const SO_REUSEADDR As Long = &H4
Private Const INVALID_SOCKET As LongPtr = -1
Private Const WOULDBLOCK As Long = 10035

Private Const PORT As Long = 25565
Private Const MOTD As String = "Hosted in Microsoft Word"
Private Const VIEW_DIST As Long = 2
Private Const TICK_MS As Long = 20
Private Const READS_PER_TICK As Long = 8
Private Const REPAINT_TICKS As Long = 15

Private Const RADIUS As Long = 4
Private Const TILE_W As Single = 40
Private Const TILE_H As Single = 22
Private Const BLOCK_H As Single = 20
Private Const POOL As Long = 140
Private Const GROUND_Y As Long = 3
Private Const MAX_LOG As Long = 30

Private mListen As LongPtr
Private mSock As LongPtr
Private mState As Long
Private mRunning As Boolean
Private mWsaUp As Boolean

Private mTimerId As LongPtr
Private mInTick As Boolean
Private mTick As Long

Private mBuf() As Byte
Private mLen As Long
Private mOutBuf() As Byte
Private mOutLen As Long
Private mOutPos As Long

Private mKeepId As Long
Private mLastKeep As Double
Private mStart As Double
Private mUser As String
Private mPX As Double, mPY As Double, mPZ As Double
Private mYaw As Single, mPitch As Single
Private mPktIn As Long, mPktOut As Long

Private mChunk() As Byte
Private mChunkReady As Boolean
Private mFavicon As String

Private mBlocks As Object
Private mDirty As Boolean
Private mLastBX As Long, mLastBZ As Long
Private mLastChunkX As Long, mLastChunkZ As Long
Private mChunkKnown As Object

Private mDoc As Object
Private mCube(1 To POOL) As Object
Private mPoolReady As Boolean
Private mCX As Single, mCY As Single
Private mLogCount As Long

Public Sub Setup()
    On Error GoTo fail
    BuildDoc
    MsgBox "Document ready." & vbCrLf & vbCrLf & _
           "Save as .docm, then run StartServer." & vbCrLf & _
           "Connect to localhost:25565 in Minecraft 1.8.9." & vbCrLf & vbCrLf & _
           "Set View to Print Layout, Review > All Markup, and Show Markup > Balloons > Show Revisions in Balloons."
    Exit Sub
fail:
    MsgBox "Setup failed: " & Err.Number & " " & Err.Description
End Sub

Public Function ServerRunning() As Boolean
    ServerRunning = mRunning
End Function

Public Sub StartServer()
    On Error GoTo fail
    Dim wsa(0 To 511) As Byte, nb As Long, sa As SOCKADDR_IN

    If mRunning Then Exit Sub

    HardReset
    If Not mPoolReady Then BuildDoc

    If WSAStartup(&H202, wsa(0)) <> 0 Then
        MsgBox "WSAStartup failed"
        Exit Sub
    End If
    mWsaUp = True

    mListen = api_socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If mListen = INVALID_SOCKET Then
        MsgBox "socket failed err=" & WSAGetLastError()
        HardReset
        Exit Sub
    End If

    nb = 1
    api_setsockopt mListen, SOL_SOCKET, SO_REUSEADDR, nb, 4

    sa.sin_family = AF_INET
    sa.sin_port = htons(CInt(PORT))
    sa.sin_addr = 0

    If api_bind(mListen, sa, LenB(sa)) <> 0 Then
        MsgBox "bind failed err=" & WSAGetLastError()
        HardReset
        Exit Sub
    End If

    If api_listen(mListen, 5) <> 0 Then
        MsgBox "listen failed err=" & WSAGetLastError()
        HardReset
        Exit Sub
    End If

    nb = 1
    api_ioctl mListen, FIONBIO, nb

    ReDim mBuf(0 To 65535)
    ReDim mOutBuf(0 To 1048575)
    mLen = 0
    mOutLen = 0
    mOutPos = 0
    mSock = INVALID_SOCKET
    mPktIn = 0
    mPktOut = 0
    mKeepId = 0
    mUser = ""
    mState = 0
    mStart = Timer
    mTick = 0
    Set mBlocks = CreateObject("Scripting.Dictionary")
    Set mChunkKnown = CreateObject("Scripting.Dictionary")
    mDirty = True

    LoadFavicon

    mDoc.TrackRevisions = True

    mRunning = True
    mTimerId = SetTimer(0, 0, TICK_MS, AddressOf TimerProc)

    If mTimerId = 0 Then
        MsgBox "SetTimer failed"
        HardReset
        Exit Sub
    End If

    SetStatus "listening on port " & PORT
    Exit Sub

fail:
    MsgBox "start error " & Err.Number & " " & Err.Description
    Err.Clear
    HardReset
End Sub

Public Sub StopServer()
    On Error Resume Next
    HardReset
    If Not mDoc Is Nothing Then mDoc.TrackRevisions = False
    SetStatus "stopped"
    Err.Clear
End Sub

Private Sub HardReset()
    On Error Resume Next
    mRunning = False
    If mTimerId <> 0 Then
        KillTimer 0, mTimerId
        mTimerId = 0
    End If
    If mSock <> 0 And mSock <> INVALID_SOCKET Then api_close mSock
    If mListen <> 0 And mListen <> INVALID_SOCKET Then api_close mListen
    mSock = INVALID_SOCKET
    mListen = INVALID_SOCKET
    mState = 0
    mLen = 0
    mOutLen = 0
    mOutPos = 0
    mInTick = False
    If mWsaUp Then
        WSACleanup
        mWsaUp = False
    End If
    Err.Clear
End Sub

Public Sub TimerProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
    On Error GoTo bail
    If mInTick Then Exit Sub
    If Not mRunning Then Exit Sub
    mInTick = True

    PollAccept
    PollRead
    FlushOut
    PollKeepAlive

    mTick = mTick + 1
    If mTick Mod REPAINT_TICKS = 0 Then Repaint

    mInTick = False
    Exit Sub

bail:
    mInTick = False
    Err.Clear
End Sub

Private Sub PollAccept()
    Dim s As LongPtr, nb As Long
    If mListen = 0 Or mListen = INVALID_SOCKET Then Exit Sub
    If mSock <> INVALID_SOCKET Then Exit Sub
    s = api_accept(mListen, 0, 0)
    If s = INVALID_SOCKET Then Exit Sub
    nb = 1
    api_ioctl s, FIONBIO, nb
    mSock = s
    mState = 0
    mLen = 0
    mOutLen = 0
    mOutPos = 0
End Sub

Private Sub PollRead()
    Dim tmp(0 To 8191) As Byte, n As Long, e As Long, i As Long, pass As Long

    For pass = 1 To READS_PER_TICK
        If mSock = INVALID_SOCKET Then Exit Sub

        n = api_recv(mSock, tmp(0), 8192, 0)
        e = WSAGetLastError()

        If n = 0 Then
            Disconnect "closed by client"
            Exit Sub
        ElseIf n < 0 Then
            If e <> WOULDBLOCK Then Disconnect "recv err=" & e
            Exit Sub
        End If

        If mLen + n > UBound(mBuf) Then ReDim Preserve mBuf(0 To (mLen + n) * 2)
        For i = 0 To n - 1
            mBuf(mLen + i) = tmp(i)
        Next
        mLen = mLen + n

        DrainPackets
        If mSock = INVALID_SOCKET Then Exit Sub
        If n < 8192 Then Exit Sub
    Next
End Sub

Private Sub DrainPackets()
    Dim p As Long, plen As Long, hdr As Long, i As Long
    Dim body() As Byte

    Do
        If mLen = 0 Then Exit Do
        p = 0
        plen = PeekVarInt(p)
        If plen < 0 Then Exit Do
        If plen = 0 Or plen > 2097152 Then
            Disconnect "bad packet length " & plen
            Exit Do
        End If
        hdr = p
        If mLen < hdr + plen Then Exit Do

        ReDim body(0 To plen - 1)
        For i = 0 To plen - 1
            body(i) = mBuf(hdr + i)
        Next

        For i = hdr + plen To mLen - 1
            mBuf(i - hdr - plen) = mBuf(i)
        Next
        mLen = mLen - hdr - plen
        mPktIn = mPktIn + 1

        Handle body, plen
        If mSock = INVALID_SOCKET Then Exit Do
    Loop
End Sub

Private Sub Handle(ByRef b() As Byte, ByVal n As Long)
    Dim p As Long, pid As Long
    Dim proto As Long, addr As String, nxt As Long
    Dim echo() As Byte, i As Long
    Dim bx As Long, by As Long, bz As Long
    Dim face As Long, item As Long, st As Long

    p = 0
    pid = ReadVarInt(b, p)

    Select Case mState
        Case 0
            If pid = 0 Then
                proto = ReadVarInt(b, p)
                addr = ReadStr(b, p)
                p = p + 2
                nxt = ReadVarInt(b, p)
                mState = IIf(nxt = 1, 1, 2)
            End If

        Case 1
            If pid = 0 Then
                SendPacket &H0, StrBytes(StatusJson())
            ElseIf pid = 1 Then
                ReDim echo(0 To 7)
                For i = 0 To 7
                    If p + i > UBound(b) Then Exit For
                    echo(i) = b(p + i)
                Next
                SendPacket &H1, echo
            End If

        Case 2
            If pid = 0 Then
                mUser = ReadStr(b, p)
                SendPacket &H2, Cat(StrBytes("069a79f4-44e9-4726-a5be-fca90e38aaf5"), StrBytes(mUser))
                mState = 3
                SendJoinSequence
                mDirty = True
                LogHeading "Session: " & mUser
                SetStatus mUser & " connected"
            End If

        Case 3
            Select Case pid
                Case &H4
                    If p + 23 > UBound(b) Then Exit Sub
                    mPX = ReadDouble(b, p)
                    mPY = ReadDouble(b, p)
                    mPZ = ReadDouble(b, p)
                Case &H5
                    If p + 7 > UBound(b) Then Exit Sub
                    mYaw = ReadFloat(b, p)
                    mPitch = ReadFloat(b, p)
                Case &H6
                    If p + 31 > UBound(b) Then Exit Sub
                    mPX = ReadDouble(b, p)
                    mPY = ReadDouble(b, p)
                    mPZ = ReadDouble(b, p)
                    mYaw = ReadFloat(b, p)
                    mPitch = ReadFloat(b, p)
                Case &H7
                    If p + 9 > UBound(b) Then Exit Sub
                    st = b(p)
                    p = p + 1
                    DecodePosition b, p, bx, by, bz
                    If st = 0 Or st = 2 Then RemoveBlock bx, by, bz
                Case &H8
                    If p + 10 > UBound(b) Then Exit Sub
                    DecodePosition b, p, bx, by, bz
                    face = b(p)
                    p = p + 1
                    If face > 5 Then Exit Sub
                    If p + 1 > UBound(b) Then Exit Sub
                    item = b(p) * 256& + b(p + 1)
                    If item = 65535 Then Exit Sub
                    If item < 1 Or item > 255 Then Exit Sub
                    OffsetByFace face, bx, by, bz
                    AddBlock bx, by, bz, item
            End Select
    End Select
End Sub

Private Sub DecodePosition(ByRef b() As Byte, ByRef p As Long, ByRef bx As Long, ByRef by As Long, ByRef bz As Long)
    Dim b0 As Long, b1 As Long, b2 As Long, b3 As Long
    Dim b4 As Long, b5 As Long, b6 As Long, b7 As Long

    b0 = b(p): b1 = b(p + 1): b2 = b(p + 2): b3 = b(p + 3)
    b4 = b(p + 4): b5 = b(p + 5): b6 = b(p + 6): b7 = b(p + 7)
    p = p + 8

    bx = b0 * &H40000 + b1 * &H400 + b2 * &H4 + (b3 \ &H40)
    by = (b3 And &H3F) * &H40 + (b4 \ &H4)
    bz = (b4 And &H3) * &H1000000 + b5 * &H10000 + b6 * &H100 + b7

    If bx >= &H2000000 Then bx = bx - &H4000000
    If by >= &H800 Then by = by - &H1000
    If bz >= &H2000000 Then bz = bz - &H4000000
End Sub

Private Sub OffsetByFace(ByVal face As Long, ByRef bx As Long, ByRef by As Long, ByRef bz As Long)
    Select Case face
        Case 0: by = by - 1
        Case 1: by = by + 1
        Case 2: bz = bz - 1
        Case 3: bz = bz + 1
        Case 4: bx = bx - 1
        Case 5: bx = bx + 1
    End Select
End Sub

Private Function BlockKey(ByVal bx As Long, ByVal by As Long, ByVal bz As Long) As String
    BlockKey = bx & ":" & by & ":" & bz
End Function

Private Sub AddBlock(ByVal bx As Long, ByVal by As Long, ByVal bz As Long, ByVal id As Long)
    On Error Resume Next
    If mBlocks Is Nothing Then Set mBlocks = CreateObject("Scripting.Dictionary")
    If by < 0 Or by > 40 Then Exit Sub
    If mBlocks.Exists(BlockKey(bx, by, bz)) Then Exit Sub
    mBlocks(BlockKey(bx, by, bz)) = id
    mDirty = True
    LogLine "placed " & BlockName(id) & " at " & bx & ", " & by & ", " & bz
    Err.Clear
End Sub

Private Sub RemoveBlock(ByVal bx As Long, ByVal by As Long, ByVal bz As Long)
    On Error Resume Next
    Dim id As Long
    If mBlocks Is Nothing Then Exit Sub
    If mBlocks.Exists(BlockKey(bx, by, bz)) Then
        id = CLng(mBlocks(BlockKey(bx, by, bz)))
        mBlocks.Remove BlockKey(bx, by, bz)
        mDirty = True
        LogLine "broke " & BlockName(id) & " at " & bx & ", " & by & ", " & bz
    End If
    Err.Clear
End Sub

Private Sub SendJoinSequence()
    Dim d() As Byte, cx As Long, cz As Long

    d = LongBE(1)
    d = Cat(d, Bt(1))
    d = Cat(d, Bt(0))
    d = Cat(d, Bt(0))
    d = Cat(d, Bt(1))
    d = Cat(d, StrBytes("flat"))
    d = Cat(d, Bt(0))
    SendPacket &H1, d

    SendPacket &H5, PositionBytes(0, 5, 0)

    For cx = -VIEW_DIST To VIEW_DIST
        For cz = -VIEW_DIST To VIEW_DIST
            SendPacket &H21, FlatChunk(cx, cz)
        Next
    Next

    mPX = 0.5: mPY = 5#: mPZ = 0.5
    mYaw = 0!: mPitch = 0!
    d = DoubleBE(mPX)
    d = Cat(d, DoubleBE(mPY))
    d = Cat(d, DoubleBE(mPZ))
    d = Cat(d, FloatBE(0!))
    d = Cat(d, FloatBE(0!))
    d = Cat(d, Bt(0))
    SendPacket &H8, d

    SendChat "Served from Word. Place a block and watch the margin."
    mLastKeep = Timer
End Sub

Private Sub PollKeepAlive()
    If mState <> 3 Then Exit Sub
    If Timer - mLastKeep < 10 Then Exit Sub
    mKeepId = mKeepId + 1
    SendPacket &H0, WriteVarInt(mKeepId)
    mLastKeep = Timer
End Sub

Public Sub SendChat(ByVal msg As String)
    If mState <> 3 Then Exit Sub
    SendPacket &H2, Cat(StrBytes("{""text"":""" & msg & """}"), Bt(0))
End Sub

Private Sub SendPacket(ByVal pid As Long, ByRef payload() As Byte)
    Dim body() As Byte, full() As Byte
    body = Cat(WriteVarInt(pid), payload)
    full = Cat(WriteVarInt(UBound(body) + 1), body)
    QueueOut full, UBound(full) + 1
    mPktOut = mPktOut + 1
End Sub

Private Sub QueueOut(ByRef b() As Byte, ByVal n As Long)
    Dim i As Long, need As Long
    If n <= 0 Then Exit Sub

    If mOutPos > 0 Then
        For i = mOutPos To mOutLen - 1
            mOutBuf(i - mOutPos) = mOutBuf(i)
        Next
        mOutLen = mOutLen - mOutPos
        mOutPos = 0
    End If

    need = mOutLen + n
    If need > UBound(mOutBuf) + 1 Then ReDim Preserve mOutBuf(0 To need * 2)

    For i = 0 To n - 1
        mOutBuf(mOutLen + i) = b(i)
    Next
    mOutLen = mOutLen + n
End Sub

Private Sub FlushOut()
    Dim sent As Long, e As Long
    If mSock = INVALID_SOCKET Then Exit Sub

    Do While mOutPos < mOutLen
        sent = api_send(mSock, mOutBuf(mOutPos), mOutLen - mOutPos, 0)
        e = WSAGetLastError()
        If sent > 0 Then
            mOutPos = mOutPos + sent
        ElseIf e = WOULDBLOCK Then
            Exit Do
        Else
            Disconnect "send err=" & e
            Exit Sub
        End If
    Loop

    If mOutPos >= mOutLen Then
        mOutPos = 0
        mOutLen = 0
    End If
End Sub

Private Sub Disconnect(ByVal why As String)
    On Error Resume Next
    If mSock <> 0 And mSock <> INVALID_SOCKET Then api_close mSock
    mSock = INVALID_SOCKET
    mState = 0
    mLen = 0
    mOutLen = 0
    mOutPos = 0
    mUser = ""
    mDirty = True
    SetStatus "disconnected: " & why
    HideFrom 1
    Err.Clear
End Sub

Private Function StatusJson() As String
    Dim fav As String
    If Len(mFavicon) > 0 Then fav = ",""favicon"":""data:image/png;base64," & mFavicon & """"
    StatusJson = "{""version"":{""name"":""1.8.9"",""protocol"":47}," & _
                 """players"":{""max"":1,""online"":" & IIf(mState = 3, 1, 0) & "}," & _
                 """description"":{""text"":""" & MOTD & """}" & fav & "}"
End Function

Private Sub LoadFavicon()
    Dim p As String, st As Object, dom As Object, nd As Object
    mFavicon = ""
    On Error GoTo fail
    If mDoc Is Nothing Then Exit Sub
    If Len(mDoc.Path) = 0 Then Exit Sub
    p = mDoc.Path & "\favicon.png"
    If Len(Dir$(p)) = 0 Then Exit Sub
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.LoadFromFile p
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    Set nd = dom.createElement("b64")
    nd.DataType = "bin.base64"
    nd.nodeTypedValue = st.Read
    st.Close
    mFavicon = Replace(Replace(nd.Text, vbCr, ""), vbLf, "")
    Exit Sub
fail:
    mFavicon = ""
    Err.Clear
End Sub

Private Function FindShape(ByVal nm As String) As Object
    On Error Resume Next
    Set FindShape = mDoc.Shapes(nm)
    Err.Clear
End Function

Private Function BindDoc() As Boolean
    Dim i As Long, sh As Object
    BindDoc = False
    On Error GoTo bail
    If Not mDoc.Bookmarks.Exists("MC_STATUS") Then Exit Function
    If Not mDoc.Bookmarks.Exists("MC_LOGEND") Then Exit Function
    For i = 1 To POOL
        Set sh = FindShape("MC_CUBE_" & i)
        If sh Is Nothing Then Exit Function
        Set mCube(i) = sh
    Next
    mCX = mDoc.PageSetup.PageWidth / 2
    mCY = 320
    mPoolReady = True
    BindDoc = True
    Exit Function
bail:
    Err.Clear
End Function

Private Sub BuildDoc()
    Dim i As Long, sh As Object, r As Object

    Set mDoc = Application.ActiveDocument
    mDoc.TrackRevisions = False

    If BindDoc() Then Exit Sub

    mDoc.Content.Delete
    For i = mDoc.Shapes.Count To 1 Step -1
        mDoc.Shapes(i).Delete
    Next

    mCX = mDoc.PageSetup.PageWidth / 2
    mCY = 320

    Set r = mDoc.Content
    r.InsertAfter "MINECRAFT SERVER" & vbCr
    r.InsertAfter "hosted in microsoft word" & vbCr
    r.InsertAfter "idle" & vbCr
    For i = 1 To 16
        r.InsertAfter vbCr
    Next
    r.InsertAfter "BUILD LOG" & vbCr

    mDoc.Paragraphs(1).Range.Style = mDoc.Styles(wdStyleHeading1)
    mDoc.Paragraphs(2).Range.Font.Name = "Consolas"
    mDoc.Paragraphs(3).Range.Font.Name = "Consolas"
    mDoc.Paragraphs(mDoc.Paragraphs.Count - 1).Range.Style = mDoc.Styles(wdStyleHeading1)

    Set r = mDoc.Paragraphs(3).Range
    r.MoveEnd wdCharacter, -1
    mDoc.Bookmarks.Add "MC_STATUS", r

    Set r = mDoc.Content
    r.Collapse wdCollapseEnd
    mDoc.Bookmarks.Add "MC_LOGEND", r
    mLogCount = 0

    For i = 1 To POOL
        Set sh = mDoc.Shapes.AddShape(msoShapeCube, -800, -800, TILE_W, TILE_H + BLOCK_H)
        sh.Name = "MC_CUBE_" & i
        sh.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        sh.RelativeVerticalPosition = wdRelativeVerticalPositionPage
        sh.WrapFormat.Type = wdWrapFront
        sh.LockAnchor = True
        sh.Line.Visible = msoFalse
        sh.Fill.ForeColor.RGB = RGB(120, 170, 120)
        sh.Visible = msoFalse
        Set mCube(i) = sh
    Next
    mPoolReady = True
End Sub

Private Sub SetStatus(ByVal s As String)
    On Error Resume Next
    Dim was As Boolean, r As Object
    If mDoc Is Nothing Then Exit Sub
    If Not mDoc.Bookmarks.Exists("MC_STATUS") Then Exit Sub
    was = mDoc.TrackRevisions
    mDoc.TrackRevisions = False
    Set r = mDoc.Bookmarks("MC_STATUS").Range
    r.Text = s
    r.Font.Name = "Consolas"
    mDoc.Bookmarks.Add "MC_STATUS", r
    mDoc.TrackRevisions = was
    Err.Clear
End Sub

Private Sub LogLine(ByVal s As String)
    On Error Resume Next
    Dim r As Object
    If mDoc Is Nothing Then Exit Sub
    If Not mDoc.Bookmarks.Exists("MC_LOGEND") Then Exit Sub
    Set r = mDoc.Bookmarks("MC_LOGEND").Range
    r.Collapse wdCollapseEnd
    r.InsertAfter s & vbCr
    r.Font.Name = "Consolas"
    r.Style = mDoc.Styles(wdStyleNormal)
    Set r = mDoc.Content
    r.Collapse wdCollapseEnd
    mDoc.Bookmarks.Add "MC_LOGEND", r
    mLogCount = mLogCount + 1
    TrimLog
    Err.Clear
End Sub

Private Sub LogHeading(ByVal s As String)
    On Error Resume Next
    Dim r As Object
    If mDoc Is Nothing Then Exit Sub
    If Not mDoc.Bookmarks.Exists("MC_LOGEND") Then Exit Sub
    Set r = mDoc.Bookmarks("MC_LOGEND").Range
    r.Collapse wdCollapseEnd
    r.InsertAfter s & vbCr
    r.Style = mDoc.Styles(wdStyleHeading2)
    Set r = mDoc.Content
    r.Collapse wdCollapseEnd
    mDoc.Bookmarks.Add "MC_LOGEND", r
    mLogCount = mLogCount + 1
    TrimLog
    Err.Clear
End Sub

Private Sub TrimLog()
    On Error Resume Next
    Dim was As Boolean, idx As Long
    If mLogCount <= MAX_LOG Then Exit Sub
    was = mDoc.TrackRevisions
    mDoc.TrackRevisions = False
    idx = mDoc.Paragraphs.Count - mLogCount + 1
    If idx >= 1 And idx <= mDoc.Paragraphs.Count Then
        mDoc.Paragraphs(idx).Range.Delete
        mLogCount = mLogCount - 1
    End If
    mDoc.TrackRevisions = was
    Err.Clear
End Sub

Private Function BlockName(ByVal id As Long) As String
    Select Case id
        Case 1: BlockName = "stone"
        Case 2: BlockName = "grass"
        Case 3: BlockName = "dirt"
        Case 4: BlockName = "cobblestone"
        Case 5: BlockName = "planks"
        Case 7: BlockName = "bedrock"
        Case 12: BlockName = "sand"
        Case 17: BlockName = "log"
        Case 20: BlockName = "glass"
        Case 24: BlockName = "sandstone"
        Case 35: BlockName = "wool"
        Case 41: BlockName = "gold block"
        Case 42: BlockName = "iron block"
        Case 45: BlockName = "bricks"
        Case 49: BlockName = "obsidian"
        Case 57: BlockName = "diamond block"
        Case 89: BlockName = "glowstone"
        Case 152: BlockName = "redstone block"
        Case Else: BlockName = "block " & id
    End Select
End Function

Private Function BlockColour(ByVal id As Long) As Long
    Select Case id
        Case 1: BlockColour = RGB(122, 122, 125)
        Case 2: BlockColour = RGB(106, 170, 80)
        Case 3: BlockColour = RGB(134, 96, 67)
        Case 4: BlockColour = RGB(128, 128, 128)
        Case 5: BlockColour = RGB(162, 130, 78)
        Case 7: BlockColour = RGB(45, 45, 48)
        Case 12: BlockColour = RGB(219, 207, 163)
        Case 17: BlockColour = RGB(102, 81, 50)
        Case 20: BlockColour = RGB(210, 235, 245)
        Case 24: BlockColour = RGB(216, 203, 155)
        Case 35: BlockColour = RGB(233, 233, 233)
        Case 41: BlockColour = RGB(243, 200, 50)
        Case 42: BlockColour = RGB(219, 219, 219)
        Case 45: BlockColour = RGB(150, 90, 78)
        Case 49: BlockColour = RGB(20, 18, 30)
        Case 57: BlockColour = RGB(98, 219, 214)
        Case 89: BlockColour = RGB(216, 190, 120)
        Case 152: BlockColour = RGB(210, 60, 45)
        Case Else: BlockColour = RGB(232, 120, 40)
    End Select
End Function

Private Sub Repaint()
    On Error GoTo bail
    Dim items() As DrawItem, cnt As Long
    Dim baseX As Long, baseZ As Long
    Dim dx As Long, dz As Long, i As Long, j As Long
    Dim k As Variant, parts() As String
    Dim bx As Long, by As Long, bz As Long
    Dim fx As Single, fz As Single
    Dim tmp As DrawItem
    Dim sh As Object
    Dim isoX As Single, isoY As Single
    Dim cxk As String

    If Not mPoolReady Then Exit Sub
    If mState <> 3 Then Exit Sub

    baseX = Int(mPX)
    baseZ = Int(mPZ)

    cxk = (baseX \ 16) & "," & (baseZ \ 16)
    If Not mChunkKnown Is Nothing Then
        If Not mChunkKnown.Exists(cxk) Then
            mChunkKnown(cxk) = 1
            LogHeading "Chunk " & cxk
        End If
    End If

    If Not mDirty And baseX = mLastBX And baseZ = mLastBZ Then
        UpdateStatusLine
        Exit Sub
    End If
    mLastBX = baseX
    mLastBZ = baseZ
    mDirty = False

    ReDim items(1 To POOL)
    cnt = 0

    For dx = -RADIUS To RADIUS
        For dz = -RADIUS To RADIUS
            If cnt < POOL Then
                cnt = cnt + 1
                items(cnt).bx = baseX + dx
                items(cnt).by = GROUND_Y
                items(cnt).bz = baseZ + dz
                If ((baseX + dx) Mod 16 = 0) Or ((baseZ + dz) Mod 16 = 0) Then
                    items(cnt).col = RGB(70, 120, 60)
                ElseIf ((baseX + dx + baseZ + dz) And 1) = 0 Then
                    items(cnt).col = RGB(106, 170, 80)
                Else
                    items(cnt).col = RGB(96, 158, 72)
                End If
            End If
        Next
    Next

    If Not mBlocks Is Nothing Then
        For Each k In mBlocks.Keys
            parts = Split(CStr(k), ":")
            bx = CLng(parts(0))
            by = CLng(parts(1))
            bz = CLng(parts(2))
            If Abs(bx - baseX) <= RADIUS And Abs(bz - baseZ) <= RADIUS Then
                If cnt < POOL Then
                    cnt = cnt + 1
                    items(cnt).bx = bx
                    items(cnt).by = by
                    items(cnt).bz = bz
                    items(cnt).col = BlockColour(CLng(mBlocks(k)))
                End If
            End If
        Next
    End If

    For i = 1 To cnt
        items(i).key = (items(i).bx + items(i).bz) * 128& + items(i).by
    Next

    For i = 2 To cnt
        tmp = items(i)
        j = i - 1
        Do While j >= 1
            If items(j).key <= tmp.key Then Exit Do
            items(j + 1) = items(j)
            j = j - 1
        Loop
        items(j + 1) = tmp
    Next

    fx = CSng(mPX - baseX)
    fz = CSng(mPZ - baseZ)

    For i = 1 To cnt
        dx = items(i).bx - baseX
        dz = items(i).bz - baseZ
        isoX = (CSng(dx) - fx - (CSng(dz) - fz)) * (TILE_W / 2)
        isoY = (CSng(dx) - fx + CSng(dz) - fz) * (TILE_H / 2) - (items(i).by - GROUND_Y) * BLOCK_H

        Set sh = mCube(i)
        sh.Left = mCX + isoX - TILE_W / 2
        sh.Top = mCY + isoY - TILE_H / 2
        sh.Fill.ForeColor.RGB = items(i).col
        If sh.Visible = msoFalse Then sh.Visible = msoTrue
    Next

    HideFrom cnt + 1
    UpdateStatusLine
    Exit Sub

bail:
    Err.Clear
End Sub

Private Sub HideFrom(ByVal startIdx As Long)
    On Error Resume Next
    Dim i As Long
    For i = startIdx To POOL
        If Not mCube(i) Is Nothing Then
            If mCube(i).Visible = msoTrue Then mCube(i).Visible = msoFalse
        End If
    Next
    Err.Clear
End Sub

Private Sub UpdateStatusLine()
    On Error Resume Next
    Dim n As Long
    If Not mBlocks Is Nothing Then n = mBlocks.Count
    SetStatus IIf(Len(mUser) = 0, "-", mUser) & _
              "   X " & Format$(mPX, "0.0") & "  Y " & Format$(mPY, "0.0") & "  Z " & Format$(mPZ, "0.0") & _
              "   yaw " & Format$(mYaw, "0") & _
              "   blocks " & n & _
              "   in " & mPktIn & "  out " & mPktOut
    Err.Clear
End Sub

Private Sub BuildChunkBody()
    Dim y As Long, z As Long, x As Long, i As Long, idx As Long, v As Long

    ReDim mChunk(0 To 12543)

    For y = 0 To 3
        Select Case y
            Case 0: v = 7 * 16
            Case 1, 2: v = 3 * 16
            Case 3: v = 2 * 16
        End Select
        For z = 0 To 15
            For x = 0 To 15
                idx = ((y * 256) + (z * 16) + x) * 2
                mChunk(idx) = v And &HFF
                mChunk(idx + 1) = (v \ &H100) And &HFF
            Next
        Next
    Next

    For i = 10240 To 12287
        mChunk(i) = &HFF
    Next
    For i = 12288 To 12543
        mChunk(i) = 1
    Next

    mChunkReady = True
End Sub

Private Function FlatChunk(ByVal cx As Long, ByVal cz As Long) As Byte()
    Dim r() As Byte, hx() As Byte, hz() As Byte, sz() As Byte
    Dim n As Long, i As Long

    If Not mChunkReady Then BuildChunkBody

    hx = LongBE(cx)
    hz = LongBE(cz)
    sz = WriteVarInt(12544)

    ReDim r(0 To 11 + UBound(sz) + 12544)

    For i = 0 To 3
        r(i) = hx(i)
        r(4 + i) = hz(i)
    Next
    r(8) = 1
    r(9) = 0
    r(10) = 1

    n = 11
    For i = 0 To UBound(sz)
        r(n) = sz(i)
        n = n + 1
    Next
    For i = 0 To 12543
        r(n + i) = mChunk(i)
    Next

    FlatChunk = r
End Function

Private Function WriteVarInt(ByVal v As Long) As Byte()
    Dim out(0 To 4) As Byte, n As Long, t As Long, r() As Byte, i As Long
    Do
        t = v And &H7F
        v = URShift7(v)
        If v <> 0 Then t = t Or &H80
        out(n) = CByte(t)
        n = n + 1
    Loop While v <> 0
    ReDim r(0 To n - 1)
    For i = 0 To n - 1
        r(i) = out(i)
    Next
    WriteVarInt = r
End Function

Private Function URShift7(ByVal v As Long) As Long
    If v < 0 Then
        URShift7 = ((v And &H7FFFFFFF) \ &H80) Or &H1000000
    Else
        URShift7 = v \ &H80
    End If
End Function

Private Function ReadVarInt(ByRef b() As Byte, ByRef p As Long) As Long
    Dim r As Long, sh As Long, cur As Long
    Do
        If p < 0 Or p > UBound(b) Or sh > 28 Then
            ReadVarInt = r
            Exit Function
        End If
        cur = b(p)
        p = p + 1
        r = r Or ((cur And &H7F) * (2 ^ sh))
        sh = sh + 7
    Loop While (cur And &H80) <> 0
    ReadVarInt = r
End Function

Private Function PeekVarInt(ByRef p As Long) As Long
    Dim r As Long, sh As Long, cur As Long
    p = 0
    Do
        If p >= mLen Or p > 4 Then
            PeekVarInt = -1
            Exit Function
        End If
        cur = mBuf(p)
        p = p + 1
        r = r Or ((cur And &H7F) * (2 ^ sh))
        sh = sh + 7
    Loop While (cur And &H80) <> 0
    PeekVarInt = r
End Function

Private Function ReadStr(ByRef b() As Byte, ByRef p As Long) As String
    Dim n As Long, i As Long, s As String
    n = ReadVarInt(b, p)
    If n < 0 Then n = 0
    If n > 32767 Then n = 0
    For i = 1 To n
        If p > UBound(b) Then Exit For
        s = s & Chr$(b(p))
        p = p + 1
    Next
    ReadStr = s
End Function

Private Function ReadDouble(ByRef b() As Byte, ByRef p As Long) As Double
    Dim t(0 To 7) As Byte, i As Long, d As Double
    If p < 0 Or p + 7 > UBound(b) Then
        p = UBound(b) + 1
        ReadDouble = 0
        Exit Function
    End If
    For i = 0 To 7
        t(7 - i) = b(p + i)
    Next
    p = p + 8
    CopyMemory d, t(0), 8
    ReadDouble = d
End Function

Private Function ReadFloat(ByRef b() As Byte, ByRef p As Long) As Single
    Dim t(0 To 3) As Byte, i As Long, f As Single
    If p < 0 Or p + 3 > UBound(b) Then
        p = UBound(b) + 1
        ReadFloat = 0
        Exit Function
    End If
    For i = 0 To 3
        t(3 - i) = b(p + i)
    Next
    p = p + 4
    CopyMemory f, t(0), 4
    ReadFloat = f
End Function

Private Function StrBytes(ByVal s As String) As Byte()
    Dim n As Long, i As Long, body() As Byte
    n = Len(s)
    If n = 0 Then
        StrBytes = WriteVarInt(0)
        Exit Function
    End If
    ReDim body(0 To n - 1)
    For i = 1 To n
        body(i - 1) = CByte(Asc(Mid$(s, i, 1)) And &HFF)
    Next
    StrBytes = Cat(WriteVarInt(n), body)
End Function

Private Function LongBE(ByVal v As Long) As Byte()
    Dim t(0 To 3) As Byte, r(0 To 3) As Byte, i As Long
    CopyMemory t(0), v, 4
    For i = 0 To 3
        r(i) = t(3 - i)
    Next
    LongBE = r
End Function

Private Function DoubleBE(ByVal v As Double) As Byte()
    Dim t(0 To 7) As Byte, r(0 To 7) As Byte, i As Long
    CopyMemory t(0), v, 8
    For i = 0 To 7
        r(i) = t(7 - i)
    Next
    DoubleBE = r
End Function

Private Function FloatBE(ByVal v As Single) As Byte()
    Dim t(0 To 3) As Byte, r(0 To 3) As Byte, i As Long
    CopyMemory t(0), v, 4
    For i = 0 To 3
        r(i) = t(3 - i)
    Next
    FloatBE = r
End Function

Private Function PositionBytes(ByVal x As Long, ByVal y As Long, ByVal z As Long) As Byte()
    Dim r(0 To 7) As Byte, hi As Long, lo As Long
    hi = ((x And &H3FFFFFF) * &H40) Or ((y \ &H40) And &H3F)
    lo = ((y And &H3F) * &H4000000) Or (z And &H3FFFFFF)
    r(0) = (hi \ &H1000000) And &HFF
    r(1) = (hi \ &H10000) And &HFF
    r(2) = (hi \ &H100) And &HFF
    r(3) = hi And &HFF
    r(4) = (lo \ &H1000000) And &HFF
    r(5) = (lo \ &H10000) And &HFF
    r(6) = (lo \ &H100) And &HFF
    r(7) = lo And &HFF
    PositionBytes = r
End Function

Private Function Bt(ByVal v As Long) As Byte()
    Dim r(0 To 0) As Byte
    r(0) = CByte(v And &HFF)
    Bt = r
End Function

Private Function Cat(ByRef a() As Byte, ByRef b() As Byte) As Byte()
    Dim la As Long, lb As Long, r() As Byte, i As Long
    la = UBound(a) + 1
    lb = UBound(b) + 1
    ReDim r(0 To la + lb - 1)
    For i = 0 To la - 1
        r(i) = a(i)
    Next
    For i = 0 To lb - 1
        r(la + i) = b(i)
    Next
    Cat = r
End Function
