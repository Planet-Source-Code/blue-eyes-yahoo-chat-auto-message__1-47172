VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "My Client"
   ClientHeight    =   4935
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3660
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   4935
   ScaleWidth      =   3660
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlIconMenu 
      Left            =   2430
      Top             =   4380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2450
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2564
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2684
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2798
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcon 
      Left            =   3000
      Top             =   4350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3144
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3258
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":336A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":347C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":358E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlCustom 
      Left            =   3060
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   34
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3FB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":489E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   1111
      ButtonWidth     =   1455
      ButtonHeight    =   1058
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlCustom"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwBuddy 
      Height          =   4275
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7541
      _Version        =   393217
      LabelEdit       =   1
      Style           =   1
      ImageList       =   "imlIcon"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgMenuBackground 
      Height          =   495
      Left            =   0
      Picture         =   "frmMain.frx":6F86
      Stretch         =   -1  'True
      Top             =   4410
      Visible         =   0   'False
      Width           =   5370
   End
   Begin VB.Menu mnuLogin 
      Caption         =   "Login"
      Begin VB.Menu mnuChangeUser 
         Caption         =   "Change User"
         Shortcut        =   ^O
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
            Caption         =   "Not At My Desh"
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
            Caption         =   "Invisible Mode"
            Index           =   14
         End
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuUnload 
         Caption         =   "Close && Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objIconMenu As New cIconMenu
Private EndOfEXE As Boolean
'

Private Sub Form_Load()
    frmLogIn.tmrAnimate.Enabled = False
    objIconMenu.ActiveMenuForeColor = &HFF0000
    objIconMenu.HighlightStyle = ECPHighlightStyleButton
    objIconMenu.MenuBackgroundColor = &HF3F3F3
    objIconMenu.InActiveMenuForeColor = &H6F4744
    objIconMenu.BackgroundPicture = imgMenuBackground
    Dim MyFont As New StdFont
    Set MyFont = Me.Font
    
    MyFont.Size = 9
    MyFont.Name = "Tahoma"
    Set objIconMenu.Font = MyFont
    objIconMenu.Attach hwnd
'    objIconMenu.ImageList = imlIconMenu
'
'    objIconMenu.IconIndex(mnuChangeUser.Caption) = 1
'    objIconMenu.IconIndex(mnuStatus.Caption) = 6
'    objIconMenu.IconIndex(mnuDisconnect.Caption) = 5
'    objIconMenu.IconIndex(mnuUnload.Caption) = 9
'    objIconMenu.IconIndex(mnuMyStatus(0).Caption) = 4
'    Dim mnuIndex As Integer
'
'    For mnuIndex = 2 To 10
'        objIconMenu.IconIndex(mnuMyStatus(mnuIndex).Caption) = 2
'    Next
'    objIconMenu.IconIndex(mnuMyStatus(7).Caption) = 8
'    objIconMenu.IconIndex(mnuMyStatus(14).Caption) = 3
    IsfrmMainLoaded = True
    Dim i As Long, j As Long, imgIndex As Integer, tmpPosition As Long, tmpStatusID As String, tmpFriendName As String
    
    Dim nodx As Node
    
    If Not GroupError Then
        Set nodx = tvwBuddy.Nodes.Add(, , "Main", "friends for - " & UserName, 4)
        nodx.Bold = True
        For i = 0 To UBound(GroupName)
            Set nodx = tvwBuddy.Nodes.Add(, , "GroupName(" & i & ")", GroupName(i), 6)
            nodx.Expanded = True
            For j = 1 To ListLen(lstFriendID(i))
                tmpFriendName = ListGetAt(lstFriendID(i), j)
                tmpPosition = ListFindNoCase(lstOnLineFriend, tmpFriendName, Chr(255))
                If tmpPosition > 0 Then
                    ' online, now finding the status
                    If StrComp(ListGetAt(lstStatus, tmpPosition, Chr(255)), "I'm Available", vbTextCompare) = 0 Then
                        ' the status is online
                        imgIndex = 1
                    Else
                        ' Is any custom status message
                        If InStr(1, ListGetAt(lstStatus, tmpPosition, Chr(255)), Chr(254), vbBinaryCompare) > 0 Then
                            If Right(ListGetAt(lstStatus, tmpPosition, Chr(255)), 1) = "0" Then
                                imgIndex = 1
                            Else
                                imgIndex = 3
                            End If
                        Else
                            ' No its not an custom status message
                            imgIndex = 3
                        End If
                        tmpFriendName = tmpFriendName & " (" & ListGetAt(ListGetAt(lstStatus, tmpPosition, Chr(255)), 1, Chr(254)) & ")"
                    End If
                Else
                    ' offline
                    imgIndex = 2
                End If
                Set nodx = tvwBuddy.Nodes.Add("GroupName(" & i & ")", tvwChild, , tmpFriendName, imgIndex)
                If imgIndex <> 2 Then nodx.Bold = True
                
            Next j
        Next i
    Else
        Set nodx = tvwBuddy.Nodes.Add(, , "Main", "No friend for - " & UserName, 4)
        nodx.Bold = True
    End If
    Set nodx = Nothing
End Sub

Private Sub Form_Paint()
    If Me.Height < 5625 Then
        Me.Height = 5625
    End If
    If Me.Width < 3780 Then
        Me.Width = 3780
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Height < 5625 Then
        Me.Height = 5625
'        Exit Sub
    End If
    If Me.Width < 3780 Then
        Me.Width = 3780
'        Exit Sub
    End If
    If Me.WindowState = vbMinimized Then Exit Sub
    tvwBuddy.Width = Me.ScaleWidth
    tvwBuddy.Height = Me.ScaleHeight - tlbMain.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not EndOfEXE Then
        Cancel = 1
        frmMain.Visible = False
    Else
        
    End If
End Sub

Public Sub mnuChangeUser_Click()
    tvwBuddy.Nodes.Clear
    
    Erase GroupName
    Erase lstFriendID
    lstOnLineFriend = ""
    lstStatus = ""
    
    If frmLogIn.sckYahoo.State <> sckClosed Then frmLogIn.sckYahoo.Close
    frmLogIn.Visible = True
    frmLogIn.txtID = ""
    frmLogIn.txtPassword = ""
    frmLogIn.lblStatus = "Provide YahooID and Password"
    frmLogIn.Left = frmMain.Left
    frmLogIn.Top = frmMain.Top + (frmMain.Width - Me.Width) / 2
    EndOfEXE = True
    Unload Me
End Sub

Private Sub mnuClose_Click()
    frmMain.Visible = False
End Sub

Public Sub mnuDisconnect_Click()
    
    If frmLogIn.sckYahoo.State <> sckClosed Then frmLogIn.sckYahoo.Close
'    frmLogIn.Visible = True
'    frmLogIn.txtID = ""
'    frmLogIn.txtPassword = ""
'    frmLogIn.lblStatus = "Disconnected"
'    frmLogIn.Left = frmMain.Left
'    frmLogIn.Top = frmMain.Top + (frmMain.Width - Me.Width) / 2
    Unload frmLogIn
    IsLoggedIn = False
    EndOfEXE = True
    Unload Me
End Sub

Public Sub mnuMyStatus_Click(Index As Integer)
    Select Case Index
    
    Case 0
        ChangeStatus Index, frmLogIn.sckYahoo, , False
    Case 2 To 10
        ChangeStatus Index - 1, frmLogIn.sckYahoo
    Case 14
        ChangeStatus Index - 2, frmLogIn.sckYahoo
    Case 12
        Load frmNewStatusMessage
        frmNewStatusMessage.Visible = True
    End Select
End Sub

Public Sub mnuUnload_Click()
    On Error Resume Next
    Dim i As Long
    For i = 0 To UBound(frmNewPager)
        Unload frmNewPager(i)
    Next
    Erase frmNewPager
    frmLogIn.sckYahoo.Close
    Unload frmLogIn
    EndOfEXE = True
    Unload frmChat
    Unload frmMenu
    Unload Me
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
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
    ElseIf Button.Index = 2 Then
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
    ElseIf Button.Index = 3 Then
        frmAddMe.Show
    ElseIf Button.Index = 4 Then
        frmBoot.Show
    End If
End Sub

Private Sub tvwBuddy_DblClick()
    If tvwBuddy.SelectedItem.Children > 0 Or tvwBuddy.SelectedItem.Index = 1 Then Exit Sub
    If ListFindNoCase(lstNowChatting, ListGetAt(tvwBuddy.SelectedItem.Text, 1, " "), Chr(255)) > 0 Then
        Exit Sub
    End If
    On Error Resume Next
    If UBound(frmNewPager) >= 0 Then
        ReDim Preserve frmNewPager(UBound(frmNewPager) + 1) As New frmPager
    Else
        ReDim frmNewPager(0)
    End If
    If Err.Number > 0 Then
        ReDim frmNewPager(0)
    End If
    Load frmNewPager(UBound(frmNewPager))
    lstNowChatting = ListAppend(lstNowChatting, ListGetAt(tvwBuddy.SelectedItem.Text, 1, " "), Chr(255))
    frmNewPager(UBound(frmNewPager)).Tag = ListGetAt(tvwBuddy.SelectedItem.Text, 1, " ")
    frmNewPager(UBound(frmNewPager)).Visible = True
    frmNewPager(UBound(frmNewPager)).lblTo = "To: " & ListGetAt(tvwBuddy.SelectedItem.Text, 1, " ")
    frmNewPager(UBound(frmNewPager)).txtTo.Text = ListGetAt(tvwBuddy.SelectedItem.Text, 1, " ")
    frmNewPager(UBound(frmNewPager)).txtTo.Visible = False
End Sub
