VERSION 5.00
Begin VB.Form frmAddMe 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add a friend"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddMe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Yahoo.MyButton cmdFinish 
      Height          =   375
      Left            =   2730
      TabIndex        =   8
      Top             =   3990
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&Finish"
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
      MICON           =   "frmAddMe.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtMessage 
      Height          =   1065
      Left            =   2490
      TabIndex        =   6
      Top             =   2790
      Width           =   3555
   End
   Begin VB.ComboBox cboGroup 
      Height          =   390
      Left            =   2490
      TabIndex        =   3
      Top             =   1350
      Width           =   3555
   End
   Begin VB.TextBox txtFriendsID 
      Height          =   390
      Left            =   2490
      TabIndex        =   2
      Top             =   810
      Width           =   3555
   End
   Begin Yahoo.MyButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      Top             =   3990
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&Cancel"
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
      MICON           =   "frmAddMe.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Message to your Friend:"
      Height          =   270
      Left            =   270
      TabIndex        =   7
      Top             =   2730
      Width           =   2115
   End
   Begin VB.Label Label4 
      Caption         =   "A message will be sent to notify your friend that you added him or her as a friend."
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   120
      TabIndex        =   5
      Top             =   2100
      Width           =   6285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Select or enter a group:"
      Height          =   270
      Left            =   300
      TabIndex        =   4
      Top             =   1380
      Width           =   2085
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter Friend's ID:"
      Height          =   270
      Left            =   300
      TabIndex        =   1
      Top             =   840
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "If you know your friend's Yahoo! ID, you may add him or her as a friend by entering their Yahoo! ID."
      ForeColor       =   &H00800000&
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   6255
   End
End
Attribute VB_Name = "frmAddMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lstGroupName As String
'

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFinish_Click()
    Dim IsNewGroupAdded As Boolean, Position As Long, i As Long
    Dim IsExists As Boolean
    Dim nodx As Node
    If Not GroupError Then
        For i = 0 To UBound(lstFriendID)
            If ListFindNoCase(lstFriendID(i), txtFriendsID) > 0 Then
                IsExists = True
                MsgBox txtFriendsID & " already exists in your friend list", , "Not Added"
                Exit Sub
            End If
        Next i
        Position = ListFindNoCase(lstGroupName, cboGroup.Text, Chr(255))
        If Position = 0 Then
            IsNewGroupAdded = True
            Set nodx = frmMain.tvwBuddy.Nodes.Add(, , "GroupName(" & UBound(GroupName) + 1 & ")", cboGroup.Text, 6)
            nodx.Expanded = True
            Set nodx = frmMain.tvwBuddy.Nodes.Add("GroupName(" & UBound(GroupName) + 1 & ")", tvwChild, , LCase(txtFriendsID), 1)
            ReDim Preserve GroupName(UBound(GroupName) + 1)
            GroupName(UBound(GroupName)) = cboGroup.Text
            ReDim Preserve lstFriendID(UBound(lstFriendID) + 1)
            lstFriendID(UBound(lstFriendID)) = txtFriendsID
        Else
            IsNewGroupAdded = False
            Set nodx = frmMain.tvwBuddy.Nodes.Add("GroupName(" & Position - 1 & ")", tvwChild, , LCase(txtFriendsID), 1)
            lstFriendID(Position - 1) = ListAppend(lstFriendID(Position - 1), txtFriendsID)
        End If
    
    Else
        ReDim GroupName(0)
        ReDim lstFriendID(0)
        GroupName(UBound(GroupName)) = cboGroup.Text
        lstFriendID(UBound(lstFriendID)) = txtFriendsID
        Set nodx = frmMain.tvwBuddy.Nodes.Add(, , "GroupName(" & UBound(GroupName) & ")", cboGroup.Text, 6)
        nodx.Expanded = True
        Set nodx = frmMain.tvwBuddy.Nodes.Add("GroupName(" & UBound(GroupName) & ")", tvwChild, , txtFriendsID, 1)
        frmMain.tvwBuddy.Nodes.Item(1).Text = "friends for - " & UserName
    End If
    
    Set nodx = Nothing
    Add_Me txtFriendsID, txtMessage, cboGroup.Text, frmLogIn.sckYahoo
    Unload Me
End Sub

Private Sub Form_Load()
    Dim GroupIndex As Long
    On Error Resume Next
    If GroupError Then Exit Sub
    For GroupIndex = 0 To UBound(GroupName)
        cboGroup.AddItem GroupName(GroupIndex)
        lstGroupName = ListAppend(lstGroupName, GroupName(GroupIndex), Chr(255))
    Next GroupIndex
    cboGroup.ListIndex = 0
End Sub
