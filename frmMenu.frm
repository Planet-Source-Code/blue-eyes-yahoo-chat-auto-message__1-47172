VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMenu 
   BackColor       =   &H00FF0000&
   Caption         =   "Menu"
   ClientHeight    =   795
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   795
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList iml32 
      Left            =   180
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0164
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIconMenu 
      Left            =   870
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":02C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":03E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":04FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":061C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0C84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0FCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":10EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1200
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1314
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1428
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":187C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2558
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":4234
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgMenuBackground 
      Height          =   165
      Left            =   0
      Picture         =   "frmMenu.frx":5F10
      Stretch         =   -1  'True
      Top             =   60
      Width           =   5370
   End
   Begin VB.Menu mnuLoggedOut 
      Caption         =   "Logged Out"
      Visible         =   0   'False
      Begin VB.Menu mnuLogIn 
         Caption         =   "Log In"
      End
      Begin VB.Menu mnuSp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnload1 
         Caption         =   "Close && Exit"
      End
   End
   Begin VB.Menu mnuLoggedIn 
      Caption         =   "Logged In"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuShow 
         Caption         =   "Show Yahoo! Messenger"
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeUser 
         Caption         =   "Change User"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu sp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendMessage 
         Caption         =   "Send Message"
      End
      Begin VB.Menu mnuAutoMessages 
         Caption         =   "Auto Messages"
      End
      Begin VB.Menu mnuChatRoom 
         Caption         =   "Go to Chat Room"
      End
      Begin VB.Menu mnuAddFriend 
         Caption         =   "Add a Friend"
      End
      Begin VB.Menu mnuBoot 
         Caption         =   "Boot"
      End
      Begin VB.Menu sp4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Change My Status"
         WindowList      =   -1  'True
         Begin VB.Menu mnuMyStatus 
            Caption         =   "Available"
            Index           =   0
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "Be Right Back"
            Index           =   2
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "Busy"
            Index           =   3
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "Not At Home"
            Index           =   4
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "Not At My Desk"
            Index           =   5
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "Not In My Office"
            Index           =   6
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "On The Phone"
            Index           =   7
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "On Vacation"
            Index           =   8
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "Out To Lunch"
            Index           =   9
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "Stepped Out"
            Index           =   10
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "-"
            Index           =   11
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "New Status Message"
            Index           =   12
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "-"
            Index           =   13
         End
         Begin VB.Menu mnuMyStatus 
            Caption         =   "Invisible"
            Index           =   14
         End
      End
      Begin VB.Menu sp5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnload2 
         Caption         =   "Close && Exit"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objIconMenu As New cIconMenu
'


Private Sub Form_Load()
'    IsLoggedIn = True
'    IsfrmMainLoaded = True
    
    objIconMenu.ActiveMenuForeColor = &H40&
    objIconMenu.HighlightStyle = ECPHighlightStyleButton
    objIconMenu.MenuBackgroundColor = &HF3F3F3
    objIconMenu.InActiveMenuForeColor = &H6F4744

    objIconMenu.BackgroundPicture = frmMenu.imgMenuBackground
    Dim MyFont As New StdFont
    Set MyFont = Me.Font
    MyFont.Size = 9
    
    Set objIconMenu.Font = MyFont
 
    objIconMenu.Attach hwnd
'    objIconMenu.ImageList = imlIconMenu
'
'
'    objIconMenu.IconIndex(mnuChangeUser.Caption) = 1
'    objIconMenu.IconIndex(mnuStatus.Caption) = 6
'    objIconMenu.IconIndex(mnuDisconnect.Caption) = 5
'    objIconMenu.IconIndex(mnuUnload1.Caption) = 9
'    objIconMenu.IconIndex(mnuUnload2.Caption) = 9
'    objIconMenu.IconIndex(mnuMyStatus(0).Caption) = 4
'    Dim mnuIndex As Integer
'
'    For mnuIndex = 2 To 10
'        objIconMenu.IconIndex(mnuMyStatus(mnuIndex).Caption) = 2
'    Next
'    objIconMenu.IconIndex(mnuMyStatus(7).Caption) = 8
'    objIconMenu.IconIndex(mnuMyStatus(14).Caption) = 3
'    objIconMenu.IconIndex(mnuLogIn.Caption) = 11
'    objIconMenu.IconIndex(mnuShow.Caption) = 10
'    objIconMenu.IconIndex(mnuChatRoom.Caption) = 14
'    objIconMenu.IconIndex(mnuSendMessage.Caption) = 13
'    objIconMenu.IconIndex(mnuAddFriend.Caption) = 15
    Dim rtValue As Long
    If IsLoggedIn Then
        objIconMenu.ActiveMenuForeColor = &HFFC0FF
        objIconMenu.HighlightStyle = ECPHighlightStyleGradient
        Icon = iml32.ListImages.Item(1).Picture
        SysTrayIcon.szTip = "Yahoo! Messenger: Logged In " & Chr$(0)
    Else
        objIconMenu.ActiveMenuForeColor = &H40&
        objIconMenu.HighlightStyle = ECPHighlightStyleButton
        Icon = iml32.ListImages.Item(2).Picture
        SysTrayIcon.szTip = "Yahoo! Messenger: Disconnected " & Chr$(0)
    End If
    Me.Visible = False
    SysTrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    SysTrayIcon.hwnd = hwnd
    SysTrayIcon.uID = vbNull
    SysTrayIcon.uCallbackMessage = WM_MOUSEMOVE
    SysTrayIcon.hIcon = Me.Icon
    SysTrayIcon.cbSize = Len(SysTrayIcon)
    
    rtValue = Shell_NotifyIcon(NIM_ADD, SysTrayIcon)
    

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If IsLoggedIn Then
        objIconMenu.ActiveMenuForeColor = &HFFC0FF
        objIconMenu.HighlightStyle = ECPHighlightStyleGradient
    Else
        objIconMenu.ActiveMenuForeColor = &H40&
        objIconMenu.HighlightStyle = ECPHighlightStyleButton
    End If
    Dim Result As Long
    Dim msg As Long
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    Select Case msg
    
    Case WM_LBUTTONUP        '514 restore form window

    Case WM_LBUTTONDBLCLK    '515 restore form window
        mnuShow_Click

    Case WM_RBUTTONUP        '517 display popup menu
        If IsfrmMainLoaded And IsLoggedIn Then
            Me.PopupMenu Me.mnuLoggedIn
        Else
'            Icon = iml32.ListImages.Item(1).Picture
            Me.PopupMenu Me.mnuLoggedOut
        End If
       
    End Select
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim rtValue As Long
    rtValue = Shell_NotifyIcon(NIM_DELETE, SysTrayIcon)
'    On Error Resume Next
    Unload frmAddMe
    Unload frmBoot
End Sub

Private Sub mnuAddFriend_Click()
    frmAddMe.Show
End Sub

Private Sub mnuAutoMessages_Click()
    frmAutoMessages.Show
End Sub

Private Sub mnuBoot_Click()
    Load frmBoot
    frmBoot.Visible = True
    frmBoot.SetFocus
End Sub

Private Sub mnuChangeUser_Click()
    frmMain.mnuChangeUser_Click
End Sub

Private Sub mnuChatRoom_Click()
        RoomName = InputBox("Please enter the Chat-Room name:", "Chat Room", , Me.Left, Me.Top)
        If Trim(RoomName) <> "" And Not InRoom Then
            StartChatLogin = False
            frmLogIn.sckYahoo.SendData Chat_RoomLogin
        
        ElseIf Trim(RoomName) <> "" And InRoom Then
            Chat_Exit frmLogIn.sckYahoo
            Pause 2
            StartChatLogin = False
            InRoom = False
            frmLogIn.sckYahoo.SendData Chat_RoomLogin

        End If
End Sub

Private Sub mnuDisconnect_Click()
    frmMain.mnuDisconnect_Click
End Sub

Private Sub mnuLogIn_Click()
    Load frmLogIn
    frmLogIn.Visible = True
    frmLogIn.SetFocus
End Sub

Private Sub mnuMyStatus_Click(Index As Integer)
    frmMain.mnuMyStatus_Click Index
End Sub

Private Sub mnuSendMessage_Click()
    On Error Resume Next
    If UBound(frmNewPager) >= 0 Then
        ReDim Preserve frmNewPager(UBound(frmNewPager) + 1)
    Else
        ReDim frmNewPager(0)
    End If
    If Err.Number > 0 Then
        ReDim frmNewPager(0)
    End If
    Load frmNewPager(UBound(frmNewPager))
    frmNewPager(UBound(frmNewPager)).lblTo = "To: "
    frmNewPager(UBound(frmNewPager)).txtTo.Text = ""
    frmNewPager(UBound(frmNewPager)).txtTo.Visible = True
    frmNewPager(UBound(frmNewPager)).Visible = True
    frmNewPager(UBound(frmNewPager)).SetFocus
End Sub

Private Sub mnuShow_Click()
    If IsfrmMainLoaded Then
        frmMain.Visible = True
        frmMain.SetFocus
    Else
        Load frmLogIn
        frmLogIn.Visible = True
        frmLogIn.SetFocus
    End If
End Sub

Private Sub mnuUnload1_Click()
    Dim rtValue As Long
    rtValue = Shell_NotifyIcon(NIM_DELETE, SysTrayIcon)
    frmMain.mnuUnload_Click
    Unload Me
End Sub

Private Sub mnuUnload2_Click()
    Dim rtValue As Long
    rtValue = Shell_NotifyIcon(NIM_DELETE, SysTrayIcon)
    frmMain.mnuUnload_Click
    On Error Resume Next
        Unload frmChat
    Unload Me
End Sub


