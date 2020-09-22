VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmChat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Room"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7470
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
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlIconMenu 
      Left            =   3720
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":0114
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwChatPerson 
      Height          =   4965
      Left            =   5340
      TabIndex        =   10
      Top             =   0
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   8758
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "imlIconMenu"
      SmallIcons      =   "imlIconMenu"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Chat ID"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.TextBox txtSend 
      Height          =   585
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   5430
      Width           =   6495
   End
   Begin VB.ComboBox cboFontSize 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6630
      TabIndex        =   1
      Text            =   "cboFontSize"
      Top             =   5010
      Width           =   765
   End
   Begin VB.ComboBox cboFontName 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3300
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5040
      Width           =   2175
   End
   Begin Yahoo.MyButton cmdBold 
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   5070
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      BTYPE           =   8
      TX              =   "B"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      MICON           =   "frmChat.frx":0568
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Yahoo.MyButton cmdItalic 
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   5070
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      BTYPE           =   8
      TX              =   "I"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      MICON           =   "frmChat.frx":0584
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Yahoo.MyButton cmdUnderline 
      Height          =   315
      Left            =   660
      TabIndex        =   5
      Top             =   5070
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      BTYPE           =   8
      TX              =   "U"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
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
      MICON           =   "frmChat.frx":05A0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Yahoo.MyButton cmdSend 
      Default         =   -1  'True
      Height          =   555
      Left            =   6630
      TabIndex        =   6
      Top             =   5460
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   "&Send"
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
      MICON           =   "frmChat.frx":05BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RichTextLib.RichTextBox rtbChatData 
      Height          =   4965
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   8758
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmChat.frx":05D8
   End
   Begin VB.Label lblFontSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6180
      TabIndex        =   8
      Top             =   5070
      Width           =   375
   End
   Begin VB.Label lblFontName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Font:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2820
      TabIndex        =   7
      Top             =   5100
      Width           =   435
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsBoldPressed As Boolean
Dim IsItalicPressed As Boolean
Dim IsUnderlinePressed As Boolean
'

Private Sub cboFontName_Click()
    txtSend.FontName = Trim(cboFontName.Text)
End Sub

Private Sub cboFontSize_Change()
    If IsNumeric((Trim(cboFontSize.Text))) Then
        If CInt(Trim(cboFontSize.Text)) > 6 Then
            txtSend.FontSize = CInt(Trim(cboFontSize.Text))
        End If
    End If
End Sub

Private Sub cmdBold_Click()
    If IsBoldPressed Then
        cmdBold.ButtonType = [Flat Highlight]
        IsBoldPressed = False
        cmdBold.ForeColor = vbBlack
        cmdBold.FontItalic = False
    Else
        cmdBold.ButtonType = [Java metal]
        IsBoldPressed = True
        cmdBold.ForeColor = &H800000
        cmdBold.FontItalic = True
    End If
    txtSend.FontBold = IsBoldPressed
End Sub

Private Sub cmdItalic_Click()
    If IsItalicPressed Then
        cmdItalic.ButtonType = [Flat Highlight]
        IsItalicPressed = False
        cmdItalic.ForeColor = vbBlack
        cmdItalic.FontItalic = False
    Else
        cmdItalic.ButtonType = [Java metal]
        IsItalicPressed = True
        cmdItalic.ForeColor = &H800000
        cmdItalic.FontItalic = True
    End If
    txtSend.FontItalic = IsItalicPressed
End Sub

Private Sub cmdUnderline_Click()
    If IsUnderlinePressed Then
        cmdUnderline.ButtonType = [Flat Highlight]
        IsUnderlinePressed = False
        cmdUnderline.ForeColor = vbBlack
        cmdUnderline.FontItalic = False
    Else
        cmdUnderline.ButtonType = [Java metal]
        IsUnderlinePressed = True
        cmdUnderline.ForeColor = &H800000
        cmdUnderline.FontItalic = True
    End If
    txtSend.FontUnderline = IsUnderlinePressed
End Sub
Public Sub cmdSend_Click()
    rtbChatData.SelStart = Len(rtbChatData.Text)
    rtbChatData.SelColor = vbBlack
    rtbChatData.SelFontName = "Trebuchet MS"
    rtbChatData.SelFontSize = 10
    rtbChatData.SelBold = True
    
    If rtbChatData.Text = "" Then
        rtbChatData.Text = UserName & ": "
    Else
        rtbChatData.SelText = vbCrLf & UserName & ": "
    End If
    
    rtbChatData.SelStart = Len(rtbChatData.Text) - Len(UserName & ": ")
    rtbChatData.SelLength = Len(UserName & ": ")
    rtbChatData.SelBold = True
    
    rtbChatData.SelStart = Len(rtbChatData.Text)
    rtbChatData.SelBold = IsBoldPressed
    rtbChatData.SelItalic = IsItalicPressed
    rtbChatData.SelUnderline = IsUnderlinePressed
    rtbChatData.SelFontName = Trim(cboFontName.Text)
    rtbChatData.SelFontSize = txtSend.FontSize
    
    rtbChatData.SelStart = Len(rtbChatData.Text)
    rtbChatData.SelText = txtSend.Text
    
    Dim SendDataBuffer As String, ChangedInFontName As Boolean
    If IsBoldPressed Then
        SendDataBuffer = "[1m"
    End If
    If IsItalicPressed Then
        SendDataBuffer = SendDataBuffer & "[2m"
    End If
    If IsUnderlinePressed Then
        SendDataBuffer = SendDataBuffer & "[4m"
    End If
    If StrComp("Arial", Trim(cboFontName.Text), vbTextCompare) <> 0 Then
        SendDataBuffer = SendDataBuffer & "<font face=" & Chr(34) & txtSend.FontName & Chr(34) & " size=" & Chr(34) & txtSend.FontSize & Chr(34) & ">"
        SendDataBuffer = SendDataBuffer & txtSend.Text & "</font>"
    Else
        SendDataBuffer = SendDataBuffer & txtSend.Text
    End If
    
    Chat_ChatSend SendDataBuffer, frmLogIn.sckYahoo
    txtSend.Text = ""
    rtbChatData.SelStart = Len(rtbChatData.Text)
    rtbChatData.SelFontName = "Trebuchet MS"
    rtbChatData.SelFontSize = 10
    

End Sub

Private Sub Form_Load()
    Me.Icon = frmLogIn.Icon
    Me.Caption = " Chat Room : " & RoomName
    
    cmdSend.Enabled = False
    
    FillComboWithFonts cboFontName
    cboFontName.ListIndex = 0
    Dim i As Long
    i = 6
    Do While i <= 72
        cboFontSize.AddItem CStr(i)
        If i >= 14 And i < 30 Then
            i = i + 2
        ElseIf i >= 30 And i < 50 Then
            i = i + 5
        ElseIf i >= 50 Then
            i = i + 10
        Else
            i = i + 1
        End If
    Loop
    cboFontSize.ListIndex = 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    InRoom = False
    StartChatLogin = False
    Chat_Exit frmLogIn.sckYahoo
End Sub

Private Sub lvwChatPerson_DblClick()

    StartPM lvwChatPerson.SelectedItem.Text
    
End Sub

Private Sub txtSend_Change()
    If Len(Trim(txtSend.Text)) > 0 Then
        cmdSend.Enabled = True
    Else
        cmdSend.Enabled = False
    End If
End Sub
