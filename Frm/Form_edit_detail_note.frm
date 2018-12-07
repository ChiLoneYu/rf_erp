VERSION 5.00
Begin VB.Form Form_edit_detail_note 
   BorderStyle     =   0  '沒有框線
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   4395
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   6615
      Begin VB.CommandButton Command2 
         Caption         =   "我不明白(&N)"
         Height          =   405
         Left            =   4860
         TabIndex        =   2
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
      Begin VB.Label Label6 
         BackStyle       =   0  '透明
         Caption         =   "1.1 在表身范圍內點滑鼠左鍵選取欲修改的記錄"
         Height          =   225
         Left            =   645
         TabIndex        =   9
         Top             =   810
         Width           =   3825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "1.4 依次修改該記錄內容"
         Height          =   225
         Left            =   645
         TabIndex        =   8
         Top             =   1875
         Width           =   2055
      End
      Begin VB.Label Label14 
         BackStyle       =   0  '透明
         Caption         =   "請致電:0769-2495162"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4740
         TabIndex        =   7
         Top             =   2970
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '透明
         Caption         =   "1.6 完成表身該條記錄修改"
         Height          =   225
         Left            =   645
         TabIndex        =   6
         Top             =   2580
         Width           =   2565
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '透明
         Caption         =   "1.5 點選<确定>按鈕"
         Height          =   225
         Left            =   645
         TabIndex        =   5
         Top             =   2220
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '透明
         Caption         =   "1.3 在彈出的菜單中選取修改菜單項"
         Height          =   225
         Left            =   645
         TabIndex        =   4
         Top             =   1515
         Width           =   3285
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '透明
         Caption         =   "1.2 在表身范圍內點滑鼠右鍵"
         Height          =   225
         Left            =   645
         TabIndex        =   3
         Top             =   1170
         Width           =   2805
      End
   End
End
Attribute VB_Name = "Form_edit_detail_note"
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

