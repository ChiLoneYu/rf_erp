VERSION 5.00
Begin VB.Form Form_delete_note 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   4395
      Left            =   0
      TabIndex        =   0
      Top             =   0
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "1.1 選出欲刪除的單据"
         Height          =   225
         Left            =   900
         TabIndex        =   6
         Top             =   870
         Width           =   2055
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "請致電:0769-2495162"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4740
         TabIndex        =   5
         Top             =   2970
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "1.3 在彈出的提示信息窗口點選<是>按鈕"
         Height          =   225
         Left            =   870
         TabIndex        =   4
         Top             =   1890
         Width           =   3285
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "1.2 點選<刪除>按鈕"
         Height          =   225
         Left            =   870
         TabIndex        =   3
         Top             =   1380
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form_delete_note"
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
