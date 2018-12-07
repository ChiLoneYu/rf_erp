VERSION 5.00
Begin VB.Form aboutmms 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关於同盛软体 -- 幼童玩具ERP管理系统(V6.0)"
   ClientHeight    =   5100
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Aboutmms.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox AresButton1 
      Height          =   4815
      Left            =   210
      ScaleHeight     =   4755
      ScaleWidth      =   2040
      TabIndex        =   14
      Top             =   120
      Width           =   2100
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   2580
      TabIndex        =   13
      Top             =   3150
      Width           =   5085
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   " 确定(&Y)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6570
      TabIndex        =   11
      Top             =   4470
      Width           =   1000
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   2580
      ScaleHeight     =   1065
      ScaleWidth      =   4845
      TabIndex        =   2
      Top             =   1950
      Width           =   4905
      Begin VB.Label serial 
         BackStyle       =   0  'Transparent
         Caption         =   "序列号:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   7
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label comp_cname 
         BackStyle       =   0  'Transparent
         Caption         =   "公司"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   6
         Top             =   390
         Width           =   4365
      End
      Begin VB.Label resp_name 
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   60
         TabIndex        =   5
         Top             =   60
         Width           =   4245
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "任何个人及团体不得私自复制,如经发现将"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2580
      TabIndex        =   12
      Top             =   3810
      Width           =   4875
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "追究法律责任,其造成一切後果由违法者自行承担."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2580
      TabIndex        =   10
      Top             =   4080
      Width           =   4875
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "本产品使用权受法律保护"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2580
      TabIndex        =   9
      Top             =   3570
      Width           =   4875
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "警告:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2580
      TabIndex        =   8
      Top             =   3270
      Width           =   1005
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "本产品使用权授予:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2610
      TabIndex        =   4
      Top             =   1620
      Width           =   1995
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright  2000-2003 FuXing Corp."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2670
      TabIndex        =   3
      Top             =   1080
      Width           =   4665
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "TeamSoft ERP Ver 6.0 For Win98/2000/XP/Win NT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   750
      Width           =   5505
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "东莞同盛软体有限公司"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2610
      TabIndex        =   0
      Top             =   330
      Width           =   3675
   End
End
Attribute VB_Name = "aboutmms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim St_000 As New ADODB.Recordset
Dim Counter As Long

Private Sub Cmd_Ok_Click()
Unload Me
End Sub

Sub init_form() '初始化表单
Dim W_rs As New ADODB.Recordset

W_rs.Open "Select * FROM mmst000", G_Con, adOpenKeyset, adLockOptimistic
If W_rs.EOF = False Then
   Resp_Name.Caption = Resp_Name.Caption & Space(2) & W_rs!Resp_Name
   Comp_Cname.Caption = Comp_Cname.Caption & Space(2) & W_rs!cmp_cname
End If
Set W_rs = Nothing
serial.Caption = serial.Caption & Space(2) & "412922-731202-209"
'AresButton1.Picture = LoadPicture(App.Path & "\bmp\sysbmp\about.gif")
'If g_Language = "C" Then
'        Label7.Top = 3990
'    Else
'        Label7.Top = 4140
'    End If
End Sub
Private Sub Form_Load()
'    Set Me.Picture = GetMdiForm.Picture
'    Set Me.Picture2 = GetMdiForm.Picture
    Me.AresButton1 = LoadPicture(App.Path + "\Picture\about.jpg")
    
    
    '窗口置中
    With Me
        .ScaleMode = vbPixels
        .Left = Int(Screen.Width \ Screen.TwipsPerPixelX - Me.ScaleWidth) * 15 / 2
        .Top = Int(Screen.Height \ Screen.TwipsPerPixelY - Me.ScaleHeight - 120) * 15 / 2 + 450
        .ScaleMode = vbTwips
        .BackColor = &H80000005
    End With

    Call init_form
End Sub

