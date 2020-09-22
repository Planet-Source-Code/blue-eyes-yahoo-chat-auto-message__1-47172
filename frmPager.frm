VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmPager 
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPager.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   StartUpPosition =   3  'Windows Default
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
      Left            =   2880
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3900
      Width           =   2175
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
      Left            =   5550
      TabIndex        =   11
      Text            =   "cboFontSize"
      Top             =   3900
      Width           =   765
   End
   Begin RichTextLib.RichTextBox rtbData 
      Height          =   2745
      Left            =   30
      TabIndex        =   5
      Top             =   1110
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   4842
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmPager.frx":0442
   End
   Begin VB.TextBox rtbSend 
      Height          =   585
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4290
      Width           =   5415
   End
   Begin MSComctlLib.StatusBar sBarMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   4920
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTo 
      Height          =   315
      Left            =   690
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   750
      Width           =   2505
   End
   Begin MSComctlLib.ImageList imlIcon 
      Left            =   3840
      Top             =   660
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
            Picture         =   "frmPager.frx":04C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPager.frx":05D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPager.frx":0B27
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPager.frx":0C3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPager.frx":0D4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPager.frx":0E5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPager.frx":0F71
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlCustom 
      Left            =   3270
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   34
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPager.frx":1083
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPager.frx":23F7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Yahoo.MyButton cmdBold 
      Height          =   315
      Left            =   30
      TabIndex        =   6
      Top             =   3930
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
      MICON           =   "frmPager.frx":376B
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
      Left            =   330
      TabIndex        =   7
      Top             =   3930
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
      MICON           =   "frmPager.frx":3787
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
      Left            =   630
      TabIndex        =   8
      Top             =   3930
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
      MICON           =   "frmPager.frx":37A3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   1111
      ButtonWidth     =   1455
      ButtonHeight    =   1058
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlCustom"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin Yahoo.MyButton cmdSend 
      Default         =   -1  'True
      Height          =   555
      Left            =   5520
      TabIndex        =   14
      Top             =   4320
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
      MICON           =   "frmPager.frx":37BF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
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
      Left            =   2400
      TabIndex        =   12
      Top             =   3960
      Width           =   435
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
      Left            =   5100
      TabIndex        =   10
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblSendAs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   270
      Left            =   4560
      TabIndex        =   2
      Top             =   780
      Width           =   570
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   270
      Left            =   90
      TabIndex        =   0
      Top             =   780
      Width           =   570
   End
End
Attribute VB_Name = "frmPager"
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
    rtbSend.FontName = Trim(cboFontName.Text)
End Sub

Private Sub cboFontSize_Change()
    If IsNumeric((Trim(cboFontSize.Text))) Then
        If CInt(Trim(cboFontSize.Text)) > 6 Then
            rtbSend.FontSize = CInt(Trim(cboFontSize.Text))
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
    rtbSend.FontBold = IsBoldPressed
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
    rtbSend.FontItalic = IsItalicPressed
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
    rtbSend.FontUnderline = IsUnderlinePressed
End Sub
Private Sub cmdSend_Click()
    
    If txtTo = "" Then
        Exit Sub
    End If
    If ListFindNoCase(lstNowChatting, txtTo, Chr(255)) = 0 Then
        lstNowChatting = ListAppend(lstNowChatting, txtTo, Chr(255))
        Me.Tag = txtTo
    End If
    
    lblTo = "To: " & txtTo.Text
    txtTo.Visible = False
    
    rtbData.SelStart = Len(rtbData.Text)
    rtbData.SelColor = vbBlack
    rtbData.SelFontName = "Trebuchet MS"
    rtbData.SelFontSize = 10
    rtbData.SelBold = True
    
    If rtbData.Text = "" Then
        rtbData.Text = UserName & ": "
    Else
        rtbData.SelText = vbCrLf & UserName & ": "
    End If
    
    rtbData.SelStart = Len(rtbData.Text) - Len(UserName & ": ")
    rtbData.SelLength = Len(UserName & ": ")
    rtbData.SelBold = True
    
    rtbData.SelStart = Len(rtbData.Text)
    rtbData.SelBold = IsBoldPressed
    rtbData.SelItalic = IsItalicPressed
    rtbData.SelUnderline = IsUnderlinePressed
    rtbData.SelFontName = Trim(cboFontName.Text)
    rtbData.SelFontSize = rtbSend.FontSize
    
    rtbData.SelStart = Len(rtbData.Text)
    rtbData.SelText = rtbSend.Text
    
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
        SendDataBuffer = SendDataBuffer & "<font face=" & Chr(34) & rtbSend.FontName & Chr(34) & " size=" & Chr(34) & rtbSend.FontSize & Chr(34) & ">"
        SendDataBuffer = SendDataBuffer & rtbSend.Text & "</font>"
    Else
        SendDataBuffer = SendDataBuffer & rtbSend.Text
    End If
    
    PM_Send txtTo, SendDataBuffer, frmLogIn.sckYahoo
    rtbSend.Text = ""
    rtbData.SelStart = Len(rtbData.Text)
    rtbData.SelFontName = "Trebuchet MS"
    rtbData.SelFontSize = 10
End Sub



Private Sub Form_Load()
'    Me.Icon = imlIcon.ListImages.Item(7).ExtractIcon

    Me.Caption = " Instant Message"
    lblSendAs = "Send As: " & UserName
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

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    If Me.Width <= 6435 Then
        Me.Width = 6435
    End If
    If Me.Height <= 5600 Then
        Me.Height = 5600
    End If

    rtbData.Width = Me.ScaleWidth - 2
    rtbSend.Width = Me.ScaleWidth - cmdSend.Width - 8
    cmdSend.Left = rtbSend.Width + rtbSend.Left + 4
    
    rtbData.Height = Me.ScaleHeight - (rtbSend.Height + txtTo.Height + tlbMain.Height + sBarMain.Height + cmdBold.Height + 24)
    cmdBold.Top = rtbData.Top + rtbData.Height + 4
    cmdItalic.Top = rtbData.Top + rtbData.Height + 4
    cmdUnderline.Top = rtbData.Top + rtbData.Height + 4
    
    lblFontName.Top = rtbData.Top + rtbData.Height + 7
    cboFontName.Top = rtbData.Top + rtbData.Height + 4
    
    lblFontSize.Top = rtbData.Top + rtbData.Height + 7
    lblFontSize.Left = Me.ScaleWidth - lblFontSize.Width - cboFontSize.Width - 5
    
    cboFontSize.Top = rtbData.Top + rtbData.Height + 4
    cboFontSize.Left = Me.ScaleWidth - cboFontSize.Width - 1
    
    cboFontName.Left = lblFontSize.Left - cboFontName.Width - 10
    lblFontName.Left = cboFontName.Left - lblFontName.Width - 5
    
    rtbSend.Top = cboFontSize.Top + cboFontSize.Height + 4
    cmdSend.Top = cboFontSize.Top + cboFontSize.Height + 4
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim pos As Long
    pos = ListFindNoCase(lstNowChatting, Me.Tag, Chr(255))
    If pos > 0 Then
        lstNowChatting = ListDeleteAt(lstNowChatting, pos, Chr(255))
    End If
    Me.Tag = ""
End Sub

Private Sub rtbData_Change()
    rtbData.SelStart = Len(rtbData.Text)
End Sub

Private Sub rtbSend_KeyPress(KeyAscii As Integer)
    If rtbSend.Text <> "" Then
        cmdSend.Enabled = True
    Else
        cmdSend.Enabled = False
    End If
End Sub

