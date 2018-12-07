VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form mmss810 
   BorderStyle     =   1  '单线固定
   Caption         =   "公司名称维护档(810)"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "mmss810.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "mmss810.frx":08CA
   ScaleHeight     =   5055
   ScaleWidth      =   8895
   StartUpPosition =   2  '萤幕中央
   Tag             =   "Set company's data(810)"
   Begin TabDlg.SSTab SSTab1 
      Height          =   4215
      Left            =   30
      TabIndex        =   2
      Top             =   840
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   -2147483634
      TabCaption(0)   =   "公司基本信息"
      TabPicture(0)   =   "mmss810.frx":86D0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "comp_email"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "comp_htlm"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "post_code"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "resp_name"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmp_fax"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmp_tel"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmp_eaddr"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmp_caddr"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmp_ename"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmp_cname"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "loc_name"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Cmd_quit"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmd_ok"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "公司怠行帐户信息"
      TabPicture(1)   =   "mmss810.frx":86D28
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label14"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmd_quit1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmd_ok1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Acc_No2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "acc_no1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox acc_no1 
         Appearance      =   0  '平面
         BeginProperty Font 
            Name            =   "新细明体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   -73890
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直卷轴
         TabIndex        =   31
         Top             =   540
         Width           =   7665
      End
      Begin VB.TextBox Acc_No2 
         Appearance      =   0  '平面
         BeginProperty Font 
            Name            =   "新细明体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   -73890
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直卷轴
         TabIndex        =   30
         Top             =   2070
         Width           =   7665
      End
      Begin VB.CommandButton cmd_ok1 
         Caption         =   "确认(&Y)"
         Height          =   330
         Left            =   -70260
         TabIndex        =   28
         Tag             =   "&OK"
         Top             =   3690
         Width           =   975
      End
      Begin VB.CommandButton cmd_quit1 
         Cancel          =   -1  'True
         Caption         =   "退出(&Q)"
         Height          =   330
         Left            =   -72120
         TabIndex        =   27
         Tag             =   "&Cancel"
         Top             =   3690
         Width           =   975
      End
      Begin VB.CommandButton cmd_ok 
         Caption         =   "确认(&Y)"
         Height          =   330
         Left            =   4740
         TabIndex        =   26
         Tag             =   "&OK"
         Top             =   3690
         Width           =   975
      End
      Begin VB.CommandButton Cmd_quit 
         Caption         =   "退出(&Q)"
         Height          =   330
         Left            =   2880
         TabIndex        =   25
         Tag             =   "&Cancel"
         Top             =   3690
         Width           =   975
      End
      Begin VB.TextBox loc_name 
         Appearance      =   0  '平面
         Height          =   300
         Left            =   1425
         TabIndex        =   13
         Top             =   540
         Width           =   3045
      End
      Begin VB.TextBox cmp_cname 
         Appearance      =   0  '平面
         Height          =   300
         Left            =   1425
         TabIndex        =   12
         Top             =   870
         Width           =   7215
      End
      Begin VB.TextBox cmp_ename 
         Appearance      =   0  '平面
         Height          =   300
         Left            =   1425
         TabIndex        =   11
         Top             =   1200
         Width           =   7215
      End
      Begin VB.TextBox cmp_caddr 
         Appearance      =   0  '平面
         Height          =   300
         Left            =   1425
         TabIndex        =   10
         Top             =   1530
         Width           =   7215
      End
      Begin VB.TextBox cmp_eaddr 
         Appearance      =   0  '平面
         Height          =   300
         Left            =   1425
         TabIndex        =   9
         Top             =   1860
         Width           =   7215
      End
      Begin VB.TextBox cmp_tel 
         Appearance      =   0  '平面
         Height          =   300
         Left            =   1425
         TabIndex        =   8
         Top             =   2190
         Width           =   3045
      End
      Begin VB.TextBox cmp_fax 
         Appearance      =   0  '平面
         Height          =   300
         Left            =   5595
         TabIndex        =   7
         Top             =   2190
         Width           =   3045
      End
      Begin VB.TextBox resp_name 
         Appearance      =   0  '平面
         Height          =   300
         Left            =   1425
         TabIndex        =   6
         Top             =   2520
         Width           =   3045
      End
      Begin VB.TextBox post_code 
         Appearance      =   0  '平面
         Height          =   300
         Left            =   5595
         MaxLength       =   6
         TabIndex        =   5
         Top             =   2520
         Width           =   3045
      End
      Begin VB.TextBox comp_htlm 
         Appearance      =   0  '平面
         Height          =   300
         Left            =   1425
         TabIndex        =   4
         Top             =   3180
         Width           =   7215
      End
      Begin VB.TextBox comp_email 
         Appearance      =   0  '平面
         Height          =   300
         Left            =   1425
         TabIndex        =   3
         Top             =   2850
         Width           =   7215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  '透明
         Caption         =   "香港账户:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   32
         Tag             =   "Short for Corp:"
         Top             =   2160
         Width           =   825
      End
      Begin VB.Label Label15 
         BackStyle       =   0  '透明
         Caption         =   "台湾账户:"
         Height          =   270
         Left            =   -74790
         TabIndex        =   29
         Tag             =   "Short for Corp:"
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '透明
         Caption         =   "公司简称:"
         Height          =   270
         Left            =   300
         TabIndex        =   24
         Tag             =   "Short for Corp:"
         Top             =   600
         Width           =   2145
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '透明
         Caption         =   "中文名称:"
         Height          =   270
         Left            =   300
         TabIndex        =   23
         Tag             =   "Chinese Name:"
         Top             =   930
         Width           =   1905
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '透明
         Caption         =   "英文名称:"
         Height          =   270
         Left            =   300
         TabIndex        =   22
         Tag             =   "English Name:"
         Top             =   1260
         Width           =   1830
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '透明
         Caption         =   "公司电话:"
         Height          =   270
         Left            =   300
         TabIndex        =   21
         Tag             =   "Tel:"
         Top             =   2250
         Width           =   1845
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '透明
         Caption         =   "中文地址:"
         Height          =   270
         Left            =   300
         TabIndex        =   20
         Tag             =   "Addr(Chinese):"
         Top             =   1590
         Width           =   1740
      End
      Begin VB.Label Label7 
         BackStyle       =   0  '透明
         Caption         =   "英文地址:"
         Height          =   270
         Left            =   300
         TabIndex        =   19
         Tag             =   "Addr(English ):"
         Top             =   1920
         Width           =   1800
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '透明
         Caption         =   "传真:"
         Height          =   330
         Left            =   4935
         TabIndex        =   18
         Tag             =   "Fax:"
         Top             =   2250
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  '透明
         Caption         =   "负  责  人:"
         Height          =   270
         Left            =   300
         TabIndex        =   17
         Tag             =   "Principal:"
         Top             =   2580
         Width           =   1590
      End
      Begin VB.Label Label11 
         BackStyle       =   0  '透明
         Caption         =   "邮编:"
         Height          =   330
         Left            =   4935
         TabIndex        =   16
         Tag             =   "Post code:"
         Top             =   2580
         Width           =   1065
      End
      Begin VB.Label Label12 
         BackStyle       =   0  '透明
         Caption         =   "网页地址:"
         Height          =   270
         Left            =   300
         TabIndex        =   15
         Tag             =   "Web site:"
         Top             =   3210
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackStyle       =   0  '透明
         Caption         =   "邮件地址:"
         Height          =   270
         Left            =   300
         TabIndex        =   14
         Tag             =   "E-mail:"
         Top             =   2910
         Width           =   2100
      End
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   255
      Picture         =   "mmss810.frx":86D44
      Top             =   90
      Width           =   705
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "注意:请认真填写以上内容"
      Height          =   255
      Left            =   300
      TabIndex        =   1
      Tag             =   "Please earnestly input about data"
      Top             =   4140
      Width           =   3285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "欢迎使用同盛科技ERP管理系统(SQL Server 版)"
      BeginProperty Font 
         Name            =   "新细明体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   1050
      TabIndex        =   0
      Tag             =   "Welcome to use NCST ERP System(SQL SERVER V7.0)"
      Top             =   330
      Width           =   5505
   End
End
Attribute VB_Name = "mmss810"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_000 As DAO.Recordset
Dim title_db As Database
Dim st_000 As New ADODB.Recordset

Private Sub cmd_ok_Click()
On Error Resume Next

st_000!loc_name = loc_name.Text
st_000!cmp_cname = cmp_cname.Text
st_000!cmp_ename = cmp_ename.Text
st_000!cmp_caddr = cmp_caddr.Text
st_000!cmp_eaddr = cmp_eaddr.Text
st_000!cmp_tel = cmp_tel.Text
st_000!cmp_fax = cmp_fax.Text

st_000!resp_name = resp_name.Text

st_000!post_code = post_code.Text
st_000!comp_email = comp_email.Text
st_000!comp_htlm = comp_htlm.Text
st_000!cmp_caddr = cmp_caddr.Text
st_000.Update

Set w_000 = title_db.OpenRecordset("mmst000")
w_000.Edit
w_000!loc_name = loc_name.Text
w_000!cmp_cname = cmp_cname.Text
w_000!cmp_ename = cmp_ename.Text
w_000!cmp_caddr = cmp_caddr.Text
w_000!cmp_eaddr = cmp_eaddr.Text
w_000!cmp_tel = cmp_tel.Text
w_000!cmp_fax = cmp_fax.Text
w_000!resp_name = resp_name.Text

w_000!post_code = post_code.Text
w_000!comp_email = comp_email.Text
w_000!comp_htlm = comp_htlm.Text

w_000.Update
Unload Me
End Sub

Private Sub cmd_ok1_Click()

On Error Resume Next

st_000!bank_no1 = Trim(acc_no1.Text)
st_000!bank_No2 = Trim(Acc_No2.Text)

st_000.Update
End Sub

Sub cmd_quit_Click()
Unload Me
End Sub


Private Sub cmd_quit1_Click()
Unload Me
End Sub

Private Sub Form_Load()
st_000.CursorLocation = adUseClient
st_000.Open "select * from mmst000", G_Con, adOpenDynamic, adLockPessimistic

Set title_db = OpenDatabase(G_Path & "\data\T_title.mdb")

st_000.MoveFirst
loc_name.Text = NullSetValue(st_000!loc_name, "")
cmp_cname.Text = NullSetValue(st_000!cmp_cname, "")
cmp_ename.Text = NullSetValue(st_000!cmp_ename, "")
cmp_caddr.Text = NullSetValue(st_000!cmp_caddr, "")
cmp_eaddr.Text = NullSetValue(st_000!cmp_eaddr, "")
cmp_tel.Text = NullSetValue(st_000!cmp_tel, "")
cmp_fax.Text = NullSetValue(st_000!cmp_fax, "")

resp_name.Text = NullSetValue(st_000!resp_name, "")

post_code.Text = NullSetValue(st_000!post_code, "")
comp_email.Text = NullSetValue(st_000!comp_email, "")
comp_htlm.Text = NullSetValue(st_000!comp_htlm, "")


'acc_name1.Text = NullSetValue(st_000!acc_name1, "")
acc_no1.Text = NullSetValue(st_000!bank_no1, "")
'acc_name2.Text = NullSetValue(st_000!acc_name2, "")
Acc_No2.Text = NullSetValue(st_000!bank_No2, "")



End Sub


Private Sub Form_Unload(Cancel As Integer)
If st_000.State = adStateOpen Then
    Set st_000 = Nothing
End If
title_db.Close
End Sub


