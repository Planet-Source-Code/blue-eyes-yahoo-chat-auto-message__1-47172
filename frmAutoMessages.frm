VERSION 5.00
Begin VB.Form frmAutoMessages 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Auto Message"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "[Chat Room]"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1785
      Left            =   90
      TabIndex        =   6
      Top             =   2820
      Width           =   5235
      Begin VB.CheckBox chkPM 
         Caption         =   "Send Welcome Message (Buddy List)"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   3855
      End
      Begin VB.TextBox txtPMMessage 
         Height          =   765
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   960
         Width           =   5115
      End
      Begin VB.Label Label2 
         Caption         =   "Message:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   9
         Top             =   630
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "[Chat Room]"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2505
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   5235
      Begin VB.OptionButton optByChat 
         Caption         =   "By Chat Room"
         Height          =   345
         Left            =   810
         TabIndex        =   4
         Top             =   960
         Width           =   1995
      End
      Begin VB.OptionButton optByPM 
         Caption         =   "By PM"
         Height          =   345
         Left            =   810
         TabIndex        =   3
         Top             =   660
         Width           =   1215
      End
      Begin VB.TextBox txtChatMessage 
         Height          =   795
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1620
         Width           =   5055
      End
      Begin VB.CheckBox chkChat 
         Caption         =   "Send Welcome Message (Chat Room)"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Message:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   5
         Top             =   1290
         Width           =   1425
      End
   End
   Begin Yahoo.MyButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2610
      TabIndex        =   10
      Top             =   4620
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "&OK"
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
      MICON           =   "frmAutoMessages.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Yahoo.MyButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4020
      TabIndex        =   11
      Top             =   4620
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "frmAutoMessages.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmAutoMessages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    ChkChatValue = chkChat.Value
    ChkPMValue = chkPM.Value
    optByPMValue = optByPM.Value
    optByChatValue = optByChat.Value
    AutoChatMessage = Left(Trim(txtChatMessage), 255)
    AutoPMMessage = Left(Trim(txtPMMessage), 255)
    
    SaveSetting "MyClient", "AutoMessage_" & UserName, "ChkChatValue", ChkChatValue
    SaveSetting "MyClient", "AutoMessage_" & UserName, "ChkPMValue", ChkPMValue
    SaveSetting "MyClient", "AutoMessage_" & UserName, "optByPMValue", optByPMValue
    SaveSetting "MyClient", "AutoMessage_" & UserName, "optByChatValue", optByChatValue
    SaveSetting "MyClient", "AutoMessage_" & UserName, "AutoChatMessage", AutoChatMessage
    SaveSetting "MyClient", "AutoMessage_" & UserName, "AutoPMMessage", AutoPMMessage
    
    Unload Me
End Sub

Private Sub Form_Load()
    chkChat.Value = ChkChatValue
    chkPM.Value = ChkPMValue
    optByPM.Value = optByPMValue
    optByChat.Value = optByChatValue
    txtChatMessage = AutoChatMessage
    txtPMMessage = AutoPMMessage
End Sub


