VERSION 5.00
Begin VB.Form Form_delete_detail_note 
   BorderStyle     =   0  '�S���ؽu
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   4395
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "�ک��դF(&X)"
         Height          =   405
         Left            =   4830
         TabIndex        =   2
         Top             =   3810
         Width           =   1365
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�ڤ�����(&N)"
         Height          =   405
         Left            =   4860
         TabIndex        =   1
         Top             =   3270
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  '�z��
         Caption         =   "1.2 �b���S���I�ƹ��k��"
         Height          =   225
         Left            =   645
         TabIndex        =   8
         Top             =   1170
         Width           =   2805
      End
      Begin VB.Label Label3 
         BackStyle       =   0  '�z��
         Caption         =   "1.3 �b�u�X����椤����R����涵"
         Height          =   225
         Left            =   645
         TabIndex        =   7
         Top             =   1515
         Width           =   3285
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '�z��
         Caption         =   "1.4�b�u�X���H�����f�I��<�O>���s"
         Height          =   225
         Left            =   645
         TabIndex        =   6
         Top             =   1890
         Width           =   2865
      End
      Begin VB.Label Label5 
         BackStyle       =   0  '�z��
         Caption         =   "1.5 �������ӱ��O���R��"
         Height          =   225
         Left            =   645
         TabIndex        =   5
         Top             =   2280
         Width           =   2565
      End
      Begin VB.Label Label14 
         BackStyle       =   0  '�z��
         Caption         =   "�ЭP�q:0769-2495162"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4740
         TabIndex        =   4
         Top             =   2970
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label6 
         BackStyle       =   0  '�z��
         Caption         =   "1.1 �b���S���I�ƹ����������R�����O��"
         Height          =   225
         Left            =   645
         TabIndex        =   3
         Top             =   810
         Width           =   3825
      End
   End
End
Attribute VB_Name = "Form_delete_detail_note"
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


