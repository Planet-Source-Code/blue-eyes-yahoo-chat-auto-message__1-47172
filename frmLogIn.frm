VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmLogIn 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Log in"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3675
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogIn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   3675
   StartUpPosition =   3  'Windows Default
   Begin Yahoo.MyButton cmdLogin 
      Default         =   -1  'True
      Height          =   375
      Left            =   2220
      TabIndex        =   7
      Top             =   1140
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&Log in"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   14933984
      BCOLO           =   14933984
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLogIn.frx":0442
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer tmrAnimate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   1050
      Top             =   1320
   End
   Begin VB.Timer tmrLoadfrmMain 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   450
      Top             =   1350
   End
   Begin VB.Timer tmrBuddyList 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   -150
      Top             =   1350
   End
   Begin MSWinsockLib.Winsock sckYahoo 
      Left            =   3570
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   5050
   End
   Begin VB.TextBox txtID 
      Height          =   330
      Left            =   1230
      TabIndex        =   1
      Top             =   210
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1230
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   2325
   End
   Begin VB.Label lblStatus 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   90
      TabIndex        =   6
      Top             =   1500
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   ":"
      Height          =   285
      Index           =   1
      Left            =   990
      TabIndex        =   5
      Top             =   720
      Width           =   165
   End
   Begin VB.Label Label3 
      Caption         =   ":"
      Height          =   285
      Index           =   0
      Left            =   990
      TabIndex        =   4
      Top             =   240
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password  "
      Height          =   270
      Left            =   60
      TabIndex        =   3
      Top             =   720
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Yahoo ID  "
      Height          =   270
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   885
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function YMSG5_GetLoginStrings Lib "YMSG5LOG.DLL" (LPYMSG5LOGIN As YMSG5_LOGIN) As Long

Private Type YMSG5_LOGIN
    ymsg5_id As String * 34
    ymsg5_pwd As String * 34
    encryption_string As String * 26
    encrypted_string1 As String * 26
    encrypted_string2 As String * 26
End Type

Private ymsg5_log As YMSG5_LOGIN

Private LogInBuffer As String
Private IsBuddyListComplete As Boolean
Private LastBufferAdded
Private sckErrorFound As Boolean
'

Private Sub cmdLogin_Click()



    Me.Visible = False
    sckErrorFound = False
    StartChatLogin = False
    Dim StartAt, IsConnected As Boolean
    tmrAnimate.Enabled = True
    lblStatus = "Connecting to Yahoo server ..."
    SysTrayIcon.szTip = "Connecting to Yahoo server ..." & Chr$(0)
    Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
    UserName = txtID
    Password = txtPassword
    
'   Load AutoMessage Setting :

    LoadAutoSetting
    
    IsLoggedIn = False
    IsBuddyListComplete = False
    LogInBuffer = ""
    IsfrmMainLoaded = False
    If sckYahoo.State <> sckClosed Then
        sckYahoo.Close
    End If
    
    StartAt = Time
    IsConnected = True
    sckYahoo.Connect "cs38.msg.sc5.yahoo.com", 5050
    
    Do While sckYahoo.State <> sckConnected
        
        DoEvents
        If Mid(Format(Time - StartAt, "hh:mm"), 4, 2) > 1 Or sckErrorFound Then
            sckYahoo.Close
            lblStatus = "Unable to connect"
            tmrAnimate.Enabled = False
            SysTrayIcon.szTip = "Yahoo! Messenger: Disconnected" & Chr$(0)
            SysTrayIcon.hIcon = frmMenu.iml32.ListImages.Item(2).Picture
            Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
            IsConnected = False
            Exit Do
        End If
        
    Loop
    
    If IsConnected Then
        lblStatus = "Sending login request"
        SysTrayIcon.szTip = "Sending login request" & Chr$(0)
        Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
        Step = 1
        sckYahoo.SendData "YMSG" & Chr(9) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(Len(UserName) + 5) & Chr(0) & Chr(&H57) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(&HC2) & Chr(&H23) & Chr(&H12) & Chr(&H58) & Chr(&H31) & Chr(&HC0) & Chr(&H80) & UserName & Chr(&HC0) & Chr(&H80)
    End If
    IsLoggedIn = False
    If Not sckErrorFound Then
        tmrCheck.Enabled = True
    Else
        Unload Me
    End If
End Sub

Private Sub sckYahoo_Close()
        StartChatLogin = False
        If sckYahoo.State <> sckClosed Then sckYahoo.Close
        lblStatus = "Logged Out"
        IsLoggedIn = False
        Me.Visible = True
        If IsfrmMainLoaded Then
            Me.Left = frmMain.Left
            Me.Top = frmMain.Top + (frmMain.Width - Me.Width) / 2
        End If
        tmrAnimate.Enabled = False
        SysTrayIcon.szTip = "Yahoo! Messanger: Disconnected" & Chr$(0)
        SysTrayIcon.hIcon = frmMenu.iml32.ListImages.Item(2).Picture
        Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
End Sub

Private Sub sckYahoo_DataArrival(ByVal bytesTotal As Long)
    sckYahoo.GetData buffer
    If Len(buffer) = 20 And Mid(buffer, 12, 6) = "ÿÿÿÿo" Then
        If sckYahoo.State <> sckClosed Then sckYahoo.Close
        lblStatus = "Logged Out"
        IsLoggedIn = False
        Me.Visible = True
        If IsfrmMainLoaded Then
            Me.Left = frmMain.Left
            Me.Top = frmMain.Top + (frmMain.Width - Me.Width) / 2
        End If
        tmrAnimate.Enabled = False
        SysTrayIcon.szTip = "Yahoo! Messanger: Disconnected" & Chr$(0)
        SysTrayIcon.hIcon = frmMenu.iml32.ListImages.Item(2).Picture
        Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
        Me.Visible = True
        StartChatLogin = False
        Exit Sub
    End If
    
    If IsBuddyListComplete = False And Step > 1 Then
        LogInBuffer = LogInBuffer & buffer
        tmrCheck.Enabled = False
        tmrCheck.Enabled = True
    End If
    If InStr(1, LogInBuffer, "À€89À€", vbBinaryCompare) > 0 Then
        If Not IsBuddyListComplete And InStr(InStr(1, LogInBuffer, "À€89À€", vbBinaryCompare), LogInBuffer, "À€8À€", vbBinaryCompare) <> 0 Then
            tmrBuddyList.Enabled = True
            tmrCheck.Enabled = False
            IsBuddyListComplete = True
        End If
    End If
    If IsLoggedIn Then
    
    Else
        Select Case Step
        
        Case 1
            ' Log in request accept. Now send the login information
            lblStatus = "Sending login information"
            ymsg5_log.ymsg5_id = UserName & Chr(0)
            ymsg5_log.ymsg5_pwd = Password & Chr(0)
            ymsg5_log.encryption_string = Mid(buffer, 30 + Len(UserName), 24) & Chr(0)
            YMSG5_GetLoginStrings ymsg5_log
            buffer = Chr(&H30) & Chr(&HC0) & Chr(&H80) & LCase(UserName) & Chr(&HC0) & Chr(&H80) & "6" & Chr(&HC0) & Chr(&H80) & Mid(ymsg5_log.encrypted_string1, 1, 24) & Chr(&HC0) & Chr(&H80) & "96" & Chr(&HC0) & Chr(&H80) & Mid(ymsg5_log.encrypted_string2, 1, 24) & Chr(&HC0) & Chr(&H80) & "2" & Chr(&HC0) & Chr(&H80) & "1" & Chr(&HC0) & Chr(&H80) & "1" & Chr(&HC0) & Chr(&H80) & LCase(UserName) & Chr(&HC0) & Chr(&H80)
            packet = "YMSG" & Chr(9) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr((Len(UserName) * 2) + 75) & Chr(0) & Chr(&H54) & Chr(0) & Chr(0) & Chr(0) & Chr(0) & Chr(&HC2) & Chr(&H23) & Chr(&H12) & Chr(&H58) & buffer
            sckYahoo.SendData packet
            Step = 2
            Exit Sub
        Case 2
            ' Find out if login or not
            If Mid(buffer, 13, 4) = Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF) Then
                sckYahoo.Close
                Step = 1
                IsLoggedIn = False
                lblStatus = "Invalid Username / Password"
                tmrCheck.Enabled = False
                tmrAnimate.Enabled = False
                SysTrayIcon.szTip = "Yahoo! Messanger: Disconnected" & Chr$(0)
                SysTrayIcon.hIcon = frmMenu.iml32.ListImages.Item(2).Picture
                Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
                NM = ""
                Me.Visible = True
                StartChatLogin = False
                Exit Sub
            Else
                Step = 3
                Exit Sub
            End If
        Case 3
            ' User is successfully logged in
            IsLoggedIn = True
            If NM = "" Then NM = Mid(buffer, 17, 4)
            
            lblStatus = "Status: logged onto id!"
            tmrAnimate.Enabled = False
            SysTrayIcon.szTip = "Yahoo! Messenger: Logged In " & Chr$(0)
            SysTrayIcon.hIcon = frmMenu.iml32.ListImages.Item(1).Picture
            Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
            
        End Select
    End If

    If InStr(1, buffer, NM & "8À€1À€7À€", vbBinaryCompare) > 0 Or InStr(1, buffer, NM & "7À€", vbBinaryCompare) > 0 Or InStr(1, buffer, NM & "0À€" & UserName & "À€7À€", vbBinaryCompare) > 0 Then  '   Online friend found
        Call StatusManagement(buffer)
    End If
    If InStr(1, buffer, NM & "3À€", vbBinaryCompare) > 0 Then
        ' Someone deny my add invite
        Call DenyManagement(buffer)
    End If
    If InStr(1, buffer, "À€49À€TYPINGÀ€", vbBinaryCompare) > 0 Then
        Typing
    End If
'   4À€ami_1951À€1À€ami_1951À€5À€rubel_176À€14À€helloÀ€97À€1À€63À€;0À€64À€0À€
    If InStr(1, buffer, "À€97À€1À€", vbBinaryCompare) > 0 Then
        GetPM
    End If

'   Add invite
    If InStr(1, buffer, NM & "1À€" & UserName & "À€3À€", vbBinaryCompare) Then
'        AddInvite TO DO
    End If
    If StartChatLogin Then
        Call HandleProto(buffer, sckYahoo, frmChat.lvwChatPerson)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If sckYahoo.State <> sckClosed Then sckYahoo.Close
    tmrAnimate.Enabled = False
    SysTrayIcon.hIcon = frmMenu.iml32.ListImages.Item(2).Picture
    Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
    IsLoggedIn = False
    StartChatLogin = False
End Sub

Private Sub BuddyManagement()
    Dim OnLineFriend As String, BuddyList As String
    Dim StartPosition As Long, EndPosition As Long
    StartPosition = InStr(1, LogInBuffer, "87À€", vbBinaryCompare) + Len("87À€")
    EndPosition = InStr(1, LogInBuffer, "88À€", vbBinaryCompare)
    BuddyList = Mid(LogInBuffer, StartPosition, EndPosition - StartPosition)
    EndPosition = InStr(1, BuddyList, Chr(10) & "À€")
    BuddyList = Left(BuddyList, EndPosition - 1)
    BuddyList = Replace(BuddyList, Chr(10), Chr(255))
    If StrComp(Right(BuddyList, 1), Chr(255), vbBinaryCompare) = 0 Then
        BuddyList = Left(BuddyList, Len(BuddyList) - 1)
    End If
    Dim LoopCounter As Long, GroupList As String
    If Len(BuddyList) = 0 Then
        GroupError = True
        Exit Sub
    End If
    ReDim GroupName(ListLen(BuddyList, Chr(255)) - 1)
    ReDim lstFriendID(ListLen(BuddyList, Chr(255)) - 1)
    
    For LoopCounter = 1 To ListLen(BuddyList, Chr(255))
        GroupList = ListGetAt(BuddyList, LoopCounter, Chr(255))
        GroupName(LoopCounter - 1) = ListGetAt(GroupList, 1, ":")
        lstFriendID(LoopCounter - 1) = ListGetAt(GroupList, 2, ":")
    Next LoopCounter
    If InStr(1, LogInBuffer, "À€8À€", vbBinaryCompare) <> 0 Then
        StartPosition = InStr(InStr(1, LogInBuffer, "À€8À€", vbBinaryCompare), LogInBuffer, "7À€", vbBinaryCompare)
        EndPosition = InStr(StartPosition, LogInBuffer, "YMSG", vbTextCompare)
        If EndPosition > 0 Then
            OnLineFriend = Mid(LogInBuffer, StartPosition, EndPosition - StartPosition)
        Else
            OnLineFriend = Mid(LogInBuffer, StartPosition)
        End If
    
        StartPosition = InStr(1, OnLineFriend, "7À€", vbBinaryCompare) + Len("7À€")
        Dim tmpOnLineFriend As String
        Do While StartPosition > 0
            EndPosition = InStr(StartPosition, OnLineFriend, "À€7À€", vbBinaryCompare)
    
            If EndPosition = 0 Then Exit Do
    
            tmpOnLineFriend = Mid(OnLineFriend, StartPosition, EndPosition - StartPosition) & "À€"
            
            lstOnLineFriend = ListAppend(lstOnLineFriend, Left(tmpOnLineFriend, InStr(1, tmpOnLineFriend, "À€10À€", vbBinaryCompare) - 1), Chr(255))
            lstStatus = ListAppend(lstStatus, GetStatus(tmpOnLineFriend), Chr(255))
            StartPosition = EndPosition + Len("À€7À€")
        Loop
        tmpOnLineFriend = Mid(OnLineFriend, StartPosition)
        lstOnLineFriend = ListAppend(lstOnLineFriend, Left(tmpOnLineFriend, InStr(1, tmpOnLineFriend, "À€10À€", vbBinaryCompare) - 1), Chr(255))
        lstStatus = ListAppend(lstStatus, GetStatus(tmpOnLineFriend), Chr(255))
    Else
        lstOnLineFriend = ""
        lstStatus = ""
    End If
End Sub

Private Sub sckYahoo_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
        sckErrorFound = True
        If sckYahoo.State <> sckClosed Then sckYahoo.Close
        lblStatus = Description
        IsLoggedIn = False
        
        If IsfrmMainLoaded Then
            Me.Visible = True
            Me.Left = frmMain.Left
            Me.Top = frmMain.Top + (frmMain.Width - Me.Width) / 2
        End If
        tmrAnimate.Enabled = False
        SysTrayIcon.hIcon = frmMenu.iml32.ListImages.Item(2).Picture
        SysTrayIcon.szTip = "Yahoo! Messenger: Disconnected " & Chr$(0)
        Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
End Sub

Private Sub tmrAnimate_Timer()
    Randomize
    Dim icoIndex As Integer
    icoIndex = Int((Rnd * 2) + 1)
    SysTrayIcon.hIcon = frmMenu.iml32.ListImages.Item(icoIndex).Picture
    Shell_NotifyIcon NIM_MODIFY, SysTrayIcon
End Sub

Private Sub tmrBuddyList_Timer()
    BuddyManagement
    tmrBuddyList.Enabled = False
    tmrLoadfrmMain.Enabled = True
End Sub

Private Function GetStatus(EncodedStatus As String) As String
    Dim StartPosition As Long, EndPosition As Long
    Dim tmpStatus As String
    StartPosition = InStr(1, EncodedStatus, "À€10À€", vbBinaryCompare) + Len("À€10À€")
    EndPosition = InStr(1, EncodedStatus, "À€11À€", vbBinaryCompare)
    tmpStatus = Mid(EncodedStatus, StartPosition, EndPosition - StartPosition)
    EndPosition = InStr(1, tmpStatus, "À€47À€", vbBinaryCompare)
    If EndPosition > 0 And Left(tmpStatus, 2) <> "99" Then
        tmpStatus = Mid(tmpStatus, 1, EndPosition - 1)
    End If
    Select Case tmpStatus
    
    Case "0"
        If InStr(1, EncodedStatus, "À€13À€0À€", vbBinaryCompare) = 0 Then
            GetStatus = "I'm Available"
        Else
            GetStatus = "Log off"
        End If
    
    Case "1"
        GetStatus = "Be Right Back"
        
    Case "2"
        GetStatus = "Busy"
        
    Case "3"
        GetStatus = "Not At Home"
        
    Case "4"
        GetStatus = "Not At My Desk"
        
    Case "5"
        GetStatus = "Not In The Office"
        
    Case "6"
        GetStatus = "On The Phone"
        
    Case "7"
        GetStatus = "On Vacation"
        
    Case "8"
        GetStatus = "Out To Lunch"
        
    Case "9"
        GetStatus = "Stepped Out"
    
    Case "999"
        Debug.Print "Unknown"
    Case Else
        StartPosition = InStr(1, tmpStatus, "À€19À€", vbBinaryCompare) + Len("À€19À€")
        EndPosition = InStr(StartPosition, tmpStatus, "À€47À€", vbBinaryCompare)
        GetStatus = Mid(tmpStatus, StartPosition, EndPosition - StartPosition) & Chr(254) & Right(tmpStatus, 1)
    End Select
End Function


Private Sub tmrCheck_Timer()
    tmrCheck.Enabled = False
    IsBuddyListComplete = True
    If InStr(1, LogInBuffer, "87À€", vbBinaryCompare) <> 0 Then
        BuddyManagement
        tmrLoadfrmMain.Enabled = True
    Else
        frmLogIn.Visible = False
        GroupError = True
        Load frmMain
        frmMain.Visible = True
    End If
End Sub

Private Sub tmrLoadfrmMain_Timer()
    frmLogIn.Visible = False
    Load frmMain
    frmMain.Visible = True
    tmrLoadfrmMain.Enabled = False
End Sub

Private Sub StatusManagement(StatusInfo As String)
    Dim StartPosition As Long, EndPosition As Long
    Dim FriendID As String, HisStatus As String, HavePreStatus As Long, IsBusy As Boolean
    Dim i As Long
    StartPosition = InStr(1, StatusInfo, NM & "8À€1À€7À€", vbBinaryCompare) + Len(NM & "8À€1À€7À€")
    If StartPosition = Len(NM & "8À€1À€7À€") Then
        StartPosition = InStr(1, StatusInfo, NM & "7À€", vbBinaryCompare) + Len(NM & "7À€")
        If StartPosition = Len(NM & "7À€") Then
            StartPosition = InStr(1, StatusInfo, NM & "0À€" & UserName & "À€7À€", vbBinaryCompare) + Len(NM & "0À€" & UserName & "À€7À€")
        End If
    End If
    EndPosition = InStr(StartPosition, StatusInfo, "À€10À€", vbBinaryCompare)
    FriendID = Mid(StatusInfo, StartPosition, EndPosition - StartPosition)
    
    HisStatus = GetStatus(StatusInfo)
    IsBusy = True
    If ListLen(HisStatus, Chr(254)) > 1 Then
        If ListGetAt(HisStatus, 2, Chr(254)) = 0 Then
            IsBusy = False
        Else
            IsBusy = True
        End If
        HisStatus = ListGetAt(HisStatus, 1, Chr(254))
    End If
    HavePreStatus = ListFindNoCase(lstOnLineFriend, FriendID, Chr(255))
    
    If HavePreStatus > 0 Then
        If StrComp(HisStatus, "Log off", vbTextCompare) = 0 Then
            lstStatus = ListDeleteAt(lstStatus, HavePreStatus, Chr(255))
            lstOnLineFriend = ListDeleteAt(lstOnLineFriend, HavePreStatus, Chr(255))
        Else
            lstStatus = ListReplaceAt(lstStatus, HisStatus, HavePreStatus, Chr(255))
        End If
    Else
        lstOnLineFriend = ListAppend(lstOnLineFriend, FriendID, Chr(255))
        lstStatus = ListAppend(lstStatus, HisStatus, Chr(255))
        If ChkPMValue = 1 Then
            PM_Send FriendID, AutoPMMessage, frmLogIn.sckYahoo
        End If
    End If
    If IsfrmMainLoaded Then
        For i = 1 To frmMain.tvwBuddy.Nodes.Count
            If StrComp(ListGetAt(frmMain.tvwBuddy.Nodes.Item(i).Text, 1, " "), FriendID, vbTextCompare) = 0 Then
                If StrComp(HisStatus, "I'm Available", vbTextCompare) = 0 Then
                    frmMain.tvwBuddy.Nodes.Item(i).Text = FriendID
                    frmMain.tvwBuddy.Nodes.Item(i).Image = 1
                    frmMain.tvwBuddy.Nodes.Item(i).Bold = True
                ElseIf StrComp(HisStatus, "Log off", vbTextCompare) = 0 Then
                    frmMain.tvwBuddy.Nodes.Item(i).Text = FriendID
                    frmMain.tvwBuddy.Nodes.Item(i).Image = 2
                    frmMain.tvwBuddy.Nodes.Item(i).Bold = False
                Else
                    frmMain.tvwBuddy.Nodes.Item(i).Text = FriendID & " (" & HisStatus & ")"
                    If IsBusy Then
                        frmMain.tvwBuddy.Nodes.Item(i).Image = 3
                        frmMain.tvwBuddy.Nodes.Item(i).Bold = True
                    Else
                        frmMain.tvwBuddy.Nodes.Item(i).Image = 1
                        frmMain.tvwBuddy.Nodes.Item(i).Bold = True
                    End If
                End If
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub DenyManagement(DenyInfo As String)
    Dim StartPosition As Long
    Dim EndPosition As Long
    Dim PersonID As String
    Dim Message As String
    Dim i As Long
    
    StartPosition = InStr(1, DenyInfo, NM & "3À€", vbBinaryCompare) + Len(NM & "3À€")
    EndPosition = InStr(StartPosition, DenyInfo, "À€", vbBinaryCompare)
    PersonID = Mid(DenyInfo, StartPosition, EndPosition - StartPosition)
    
    StartPosition = InStr(EndPosition, DenyInfo, "À€14À€", vbBinaryCompare) + Len("À€14À€")
    EndPosition = InStr(StartPosition, DenyInfo, "À€", vbBinaryCompare)
    Message = Mid(DenyInfo, StartPosition, EndPosition - StartPosition)
    If IsfrmMainLoaded Then
        For i = 1 To frmMain.tvwBuddy.Nodes.Count
            If StrComp(ListGetAt(frmMain.tvwBuddy.Nodes.Item(i).Text, 1, " "), PersonID, vbTextCompare) = 0 Then
                
                If frmMain.tvwBuddy.Nodes.Item(i).Parent.Children > 1 Then
                    frmMain.tvwBuddy.Nodes.Remove (i)
                Else
                    frmMain.tvwBuddy.Nodes.Remove (i)
                    frmMain.tvwBuddy.Nodes.Remove (i - 1)
                End If
                If frmMain.tvwBuddy.Nodes.Count = 1 Then
                    frmMain.tvwBuddy.Nodes.Item(1).Text = "No friend for - " & UserName
                End If
                Exit For
            End If
        Next i
    End If
    For i = 0 To UBound(lstFriendID) - 1
        If ListFindNoCase(lstFriendID(i), PersonID) > 0 Then
            lstFriendID(i) = ListDeleteAt(lstFriendID(i), ListFindNoCase(lstFriendID(i), PersonID))
            Exit For
        End If
    Next i

    lstStatus = ListDeleteAt(lstStatus, ListFindNoCase(lstOnLineFriend, PersonID, Chr(255)), Chr(255))
    lstOnLineFriend = ListDeleteAt(lstOnLineFriend, ListFindNoCase(lstOnLineFriend, PersonID, Chr(255)), Chr(255))
    MsgBox PersonID & " don't accept your add message and send the following message to you: " & vbCrLf & Message, , "Sorry"
End Sub

Private Sub Typing()
    Dim i As Long, StartPosition As Long, EndPosition As Long
    Dim PersonID As String
    On Error Resume Next
    StartPosition = InStr(1, buffer, NM & "4À€", vbBinaryCompare) + Len(NM & "4À€")
    If StartPosition = Len(NM & "4À€") Then
        StartPosition = InStr(1, buffer, "À€4À€", vbBinaryCompare) + Len("À€4À€")
    End If
    EndPosition = InStr(StartPosition, buffer, "À€", vbBinaryCompare)
    PersonID = Mid(buffer, StartPosition, EndPosition - StartPosition)
    For i = 0 To UBound(frmNewPager)
        If StrComp(PersonID, frmNewPager(i).Tag, vbTextCompare) = 0 Then
            frmNewPager(i).sBarMain.SimpleText = "Typing..."
            Exit Sub
        End If
    Next i
End Sub

Private Sub GetPM()
    Dim i As Long, StartPosition As Long, EndPosition As Long
    Dim PersonID As String, Message As String, IsOffLine As Boolean
    Dim IsBold As Boolean, IsItalic As Boolean, IsUnderline As Boolean, FontName As String, FontSize As String, FontColor As String
    Dim CurrentLeft As Long, j As Long
    On Error Resume Next
    If InStr(1, buffer, "À€63À€;", vbBinaryCompare) = 0 Then
        IsOffLine = True
    End If
    
    StartPosition = InStr(1, buffer, NM & "4À€", vbBinaryCompare) + Len(NM & "4À€")
    If StartPosition = Len(NM & "4À€") Then
        StartPosition = InStr(1, buffer, "À€4À€", vbBinaryCompare) + Len("À€4À€")
    End If
    EndPosition = InStr(StartPosition, buffer, "À€", vbBinaryCompare)
    PersonID = Mid(buffer, StartPosition, EndPosition - StartPosition)
    StartPosition = InStr(1, buffer, "À€14À€", vbBinaryCompare) + Len("À€14À€")
    EndPosition = InStr(StartPosition, buffer, "À€", vbBinaryCompare)
    
    Message = Mid(buffer, StartPosition, EndPosition - StartPosition)
    
    If InStr(1, Message, "[1m", vbBinaryCompare) > 0 Then
        Message = Replace(Message, "[1m", "")
        IsBold = True
    Else
        IsBold = False
    End If
    
    If InStr(1, Message, "[2m", vbBinaryCompare) > 0 Then
        Message = Replace(Message, "[2m", "")
        IsItalic = True
    Else
        IsItalic = False
    End If
    
    If InStr(1, Message, "[4m", vbBinaryCompare) > 0 Then
        Message = Replace(Message, "[4m", "")
        IsUnderline = True
    Else
        IsUnderline = False
    End If
    Message = Replace(Message, "[1m", "")
    
    
    
    Message = Replace(Message, "[2m", "")
    Message = Replace(Message, "[38m", "")
    Message = Replace(Message, "[36m", "")
    Message = Replace(Message, "[33m", "")
    
    
    
    Dim FadeR As String
    FadeR = Split(Message, "<FADE")(1): FadeR = Split(FadeR, ">")(0)
        Message = Replace(Message, "<FADE" & FadeR & ">", "")
        Message = Replace(Message, "</FADE>", "")
        
        
        Dim FadeR2 As String
    FadeR2 = Split(Message, "<fade")(1): FadeR2 = Split(FadeR2, ">")(0)
        Message = Replace(Message, "<fade" & FadeR2 & ">", "")
        Message = Replace(Message, "</fade>", "")
        
        
        Dim FontR As String
    FontR = Split(Message, "<font")(1): FontR = Split(FontR, ">")(0)
    FontName = Split(FontR, " face=" & Chr(34))(1): FontName = Split(FontName, Chr(34))(0)
    FontSize = Split(FontR, " size=" & Chr(34))(1): FontSize = Split(FontSize, Chr(34))(0)
    Message = Replace(Message, "<font" & FontR & ">", "")
    Message = Replace(Message, "</font>", "")
            
    Dim TextC As String
    TextC = Split(Message, "#")(1): TextC = Split(TextC, "m")(0)
    Message = Replace(Message, "#" & TextC & "m", "")
    
    TextC = Split(Message, "")(1): TextC = Split(TextC, "m")(0)
    Message = Replace(Message, "" & TextC & "m", "")
    Message = Replace(Message, "[", "")
    Dim IsBuzz As Boolean
    If StrComp(Message, "<ding>", vbTextCompare) = 0 Then
        Message = "BUZZ!!!"
        IsBuzz = True
    End If
    Err.Clear
    For i = 0 To UBound(frmNewPager)
        If Err > 0 Then Exit For
        If StrComp(PersonID, frmNewPager(i).Tag, vbTextCompare) = 0 Then
            If IsOffLine Then
                Load frmMessageSent
                frmMessageSent.Left = frmNewPager(i).Left
                frmMessageSent.Top = frmNewPager(i).Top
                frmMessageSent.Visible = True
                Exit Sub
            End If
            frmNewPager(i).SetFocus
            frmNewPager(i).rtbData.SelStart = Len(frmNewPager(i).rtbData.Text)
            frmNewPager(i).rtbData.SelColor = vbBlack
            frmNewPager(i).rtbData.SelFontName = "Trebuchet MS"
            frmNewPager(i).rtbData.SelFontSize = 10
            frmNewPager(i).rtbData.SelBold = True


            If frmNewPager(i).rtbData.Text = "" Then
                frmNewPager(i).rtbData.Text = PersonID & ": "
            Else
                frmNewPager(i).rtbData.SelText = vbCrLf & PersonID & ": "
            End If
            
            frmNewPager(i).rtbData.SelStart = Len(frmNewPager(i).rtbData.Text) - Len(PersonID & ": ")
            frmNewPager(i).rtbData.SelLength = Len(PersonID & ": ")
            frmNewPager(i).rtbData.SelBold = True
            
            frmNewPager(i).rtbData.SelStart = Len(frmNewPager(i).rtbData.Text)
            frmNewPager(i).rtbData.SelFontName = FontName
            frmNewPager(i).rtbData.SelFontSize = CInt(FontSize)
            
            frmNewPager(i).rtbData.SelStart = Len(frmNewPager(i).rtbData.Text)
            frmNewPager(i).rtbData.SelText = Message
            
            frmNewPager(i).rtbData.SelStart = Len(frmNewPager(i).rtbData.Text) - Len(Message)
            frmNewPager(i).rtbData.SelLength = Len(Message)
            If IsBold Then
                frmNewPager(i).rtbData.SelBold = True
            Else
                frmNewPager(i).rtbData.SelBold = False
            End If
            frmNewPager(i).rtbData.SelStart = Len(frmNewPager(i).rtbData.Text) - Len(Message)
            frmNewPager(i).rtbData.SelLength = Len(Message)
            If IsItalic Then
                frmNewPager(i).rtbData.SelItalic = True
            Else
                frmNewPager(i).rtbData.SelItalic = False
            End If
            frmNewPager(i).rtbData.SelStart = Len(frmNewPager(i).rtbData.Text) - Len(Message)
            frmNewPager(i).rtbData.SelLength = Len(Message)
            If IsUnderline Then
                frmNewPager(i).rtbData.SelUnderline = True
            Else
                frmNewPager(i).rtbData.SelUnderline = False
            End If
            frmNewPager(i).rtbData.SelStart = Len(frmNewPager(i).rtbData.Text) - Len(Message)
            frmNewPager(i).rtbData.SelLength = Len(Message)
            If IsBuzz Then
                frmNewPager(i).rtbData.SelBold = True
                frmNewPager(i).rtbData.SelStart = Len(frmNewPager(i).rtbData.Text) - Len(Message)
                frmNewPager(i).rtbData.SelLength = Len(Message)
                frmNewPager(i).rtbData.SelColor = vbRed
                CurrentLeft = frmNewPager(i).Left
                For j = 1 To 50
                    If j Mod 3 = 1 Then
                        frmNewPager(i).Left = CurrentLeft + 50
                    ElseIf j Mod 3 = 2 Then
                        frmNewPager(i).Left = CurrentLeft
                    Else
                        frmNewPager(i).Left = CurrentLeft - 50
                    End If
                Next j
                frmNewPager(i).Left = CurrentLeft
            End If
            frmNewPager(i).sBarMain.SimpleText = "Last message received on " & Now
            Exit Sub
        End If
    Next i
    If UBound(frmNewPager) >= 0 Then
        ReDim Preserve frmNewPager(UBound(frmNewPager) + 1) As New frmPager
    Else
        ReDim frmNewPager(0)
    End If
    If Err.Number > 0 Then
        ReDim frmNewPager(0)
    End If
    Load frmNewPager(UBound(frmNewPager))
    lstNowChatting = ListAppend(lstNowChatting, PersonID, Chr(255))
    frmNewPager(UBound(frmNewPager)).Tag = PersonID
    frmNewPager(UBound(frmNewPager)).Visible = True
    frmNewPager(UBound(frmNewPager)).lblTo = "To: " & PersonID
    frmNewPager(UBound(frmNewPager)).txtTo.Text = PersonID
    frmNewPager(UBound(frmNewPager)).txtTo.Visible = False
    If IsOffLine Then
        Load frmMessageSent
        frmMessageSent.Left = frmNewPager(UBound(frmNewPager)).Left
        frmMessageSent.Top = frmNewPager(UBound(frmNewPager)).Top
        frmMessageSent.Visible = True
        Exit Sub
    End If
    frmNewPager(UBound(frmNewPager)).rtbData.SelStart = Len(frmNewPager(UBound(frmNewPager)).rtbData.Text)
    frmNewPager(UBound(frmNewPager)).rtbData.SelColor = vbBlack
    frmNewPager(UBound(frmNewPager)).rtbData.SelFontName = "Trebuchet MS"
    frmNewPager(UBound(frmNewPager)).rtbData.SelFontSize = 10
    frmNewPager(UBound(frmNewPager)).rtbData.SelBold = True
    
    If frmNewPager(UBound(frmNewPager)).rtbData.Text = "" Then
        frmNewPager(UBound(frmNewPager)).rtbData.Text = PersonID & ": "
    Else
        frmNewPager(UBound(frmNewPager)).rtbData.SelText = vbCrLf & PersonID & ": "
    End If
    
    frmNewPager(UBound(frmNewPager)).rtbData.SelStart = Len(frmNewPager(UBound(frmNewPager)).rtbData.Text) - Len(PersonID & ": ")
    frmNewPager(UBound(frmNewPager)).rtbData.SelLength = Len(PersonID & ": ")
    frmNewPager(UBound(frmNewPager)).rtbData.SelBold = True
    
    frmNewPager(UBound(frmNewPager)).rtbData.SelStart = Len(frmNewPager(UBound(frmNewPager)).rtbData.Text)
    frmNewPager(UBound(frmNewPager)).rtbData.SelText = Message
    
    frmNewPager(UBound(frmNewPager)).rtbData.SelStart = Len(frmNewPager(UBound(frmNewPager)).rtbData.Text) - Len(Message)
    frmNewPager(UBound(frmNewPager)).rtbData.SelLength = Len(Message)
    If IsBold Then
        frmNewPager(UBound(frmNewPager)).rtbData.SelBold = True
    Else
        frmNewPager(UBound(frmNewPager)).rtbData.SelBold = False
    End If
    frmNewPager(UBound(frmNewPager)).rtbData.SelStart = Len(frmNewPager(UBound(frmNewPager)).rtbData.Text) - Len(Message)
    frmNewPager(UBound(frmNewPager)).rtbData.SelLength = Len(Message)
    If IsItalic Then
        frmNewPager(UBound(frmNewPager)).rtbData.SelItalic = True
    Else
        frmNewPager(UBound(frmNewPager)).rtbData.SelItalic = False
    End If
    frmNewPager(UBound(frmNewPager)).rtbData.SelStart = Len(frmNewPager(UBound(frmNewPager)).rtbData.Text) - Len(Message)
    frmNewPager(UBound(frmNewPager)).rtbData.SelLength = Len(Message)
    If IsUnderline Then
        frmNewPager(UBound(frmNewPager)).rtbData.SelUnderline = True
    Else
        frmNewPager(UBound(frmNewPager)).rtbData.SelUnderline = False
    End If
    frmNewPager(UBound(frmNewPager)).rtbData.SelStart = Len(frmNewPager(UBound(frmNewPager)).rtbData.Text) - Len(Message)
    frmNewPager(UBound(frmNewPager)).rtbData.SelLength = Len(Message)
    If IsBuzz Then
        frmNewPager(UBound(frmNewPager)).rtbData.SelBold = True
        frmNewPager(UBound(frmNewPager)).rtbData.SelStart = Len(frmNewPager(UBound(frmNewPager)).rtbData.Text) - Len(Message)
        frmNewPager(UBound(frmNewPager)).rtbData.SelLength = Len(Message)
        frmNewPager(UBound(frmNewPager)).rtbData.SelColor = vbRed
        CurrentLeft = frmNewPager(UBound(frmNewPager)).Left
    End If

    frmNewPager(UBound(frmNewPager)).sBarMain.SimpleText = "Last message received on " & Now
    
End Sub
