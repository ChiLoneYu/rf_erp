VERSION 5.00
Object = "{F4732CE3-9A6C-11D2-8018-0080AD70A386}#5.7#0"; "ARESBUTTONPRO.OCX"
Begin VB.Form Form_add_note 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   4395
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   6615
      Begin ARESBUTTONLib.AresButton AresButton1 
         Height          =   420
         Left            =   5160
         TabIndex        =   17
         Top             =   1230
         Width           =   495
         _Version        =   327687
         PrevPointer     =   56078148
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin VB.CommandButton Command2 
         Caption         =   "我不明白(&N)"
         Height          =   405
         Left            =   4860
         TabIndex        =   15
         Top             =   3270
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "我明白了(&X)"
         Height          =   405
         Left            =   4830
         TabIndex        =   1
         Top             =   3810
         Width           =   1365
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "請致電:0769-2495162"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4740
         TabIndex        =   16
         Top             =   2970
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3: 完成一張表單的輸入"
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Top             =   3990
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "2.6 完成全部的明細資料輸入后點選取消<按鈕>"
         Height          =   285
         Left            =   810
         TabIndex        =   13
         Top             =   3690
         Width           =   3855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "2.5 增加另外一筆明細記錄重復2.3項"
         Height          =   285
         Left            =   810
         TabIndex        =   12
         Top             =   3360
         Width           =   2985
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "2.4 點選确認鈕,完成一筆明細記錄輸入"
         Height          =   225
         Left            =   810
         TabIndex        =   11
         Top             =   3030
         Width           =   3315
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "2.3 在彈出的畫面依次輸入表身的明細資料"
         Height          =   225
         Left            =   810
         TabIndex        =   10
         Top             =   2730
         Width           =   3825
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "2.2 在彈出的菜單項選擇<新增>項"
         Height          =   225
         Left            =   810
         TabIndex        =   9
         Top             =   2430
         Width           =   3285
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "2.1 在表身范圍內點滑鼠右鍵"
         Height          =   225
         Left            =   840
         TabIndex        =   8
         Top             =   2130
         Width           =   2445
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2: 再來增加表身資料"
         Height          =   285
         Left            =   270
         TabIndex        =   7
         Top             =   1830
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "1.4 完成表頭輸入"
         Height          =   225
         Left            =   870
         TabIndex        =   6
         Top             =   1530
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "1.3 點選<确定>按鈕"
         Height          =   225
         Left            =   870
         TabIndex        =   5
         Top             =   1230
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "1.2 按順序輸入表頭資料,如編號,日期等"
         Height          =   225
         Left            =   870
         TabIndex        =   4
         Top             =   930
         Width           =   3285
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "1.1 點選<新增>按鈕"
         Height          =   225
         Left            =   870
         TabIndex        =   3
         Top             =   660
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1: 首先增加表頭資料"
         Height          =   285
         Left            =   300
         TabIndex        =   2
         Top             =   330
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form_add_note"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Label14.Visible = True
End Sub
