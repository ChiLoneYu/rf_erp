VERSION 5.00
Begin VB.Form FrmA022Mx 
   Appearance      =   0  'ƽ��
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '˫�߹̶��Ի�����
   Caption         =   "�����ʼ���Ŀ"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "��ϸ����"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   Icon            =   "Frma022Mx.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'өĻ����
   Begin VB.CommandButton Cmd_Brow 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2520
      TabIndex        =   8
      Top             =   390
      Width           =   300
   End
   Begin VB.TextBox Ill_Name 
      Appearance      =   0  'ƽ��
      Height          =   300
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "100 chars"
      Top             =   1635
      Width           =   4425
   End
   Begin VB.TextBox Ill_No 
      Appearance      =   0  'ƽ��
      Height          =   300
      Left            =   960
      MaxLength       =   10
      TabIndex        =   3
      ToolTipText     =   "21 Chars"
      Top             =   360
      Width           =   1875
   End
   Begin VB.TextBox Ill_Type 
      Appearance      =   0  'ƽ��
      Height          =   300
      Left            =   960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "100 chars"
      Top             =   1005
      Width           =   1905
   End
   Begin VB.CommandButton Cmd_OK 
      Height          =   330
      Left            =   1260
      Picture         =   "Frma022Mx.frx":000C
      Style           =   1  'ͼƬ���
      TabIndex        =   0
      Top             =   2280
      Width           =   1125
   End
   Begin VB.CommandButton Cmd_Cancel 
      Height          =   345
      Left            =   3450
      Picture         =   "Frma022Mx.frx":15AE
      Style           =   1  'ͼƬ���
      TabIndex        =   1
      Top             =   2250
      Width           =   1110
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  '͸��
      Caption         =   "��Ŀ���:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   420
      Width           =   765
   End
   Begin VB.Label lbl4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  '͸��
      Caption         =   "��Ŀ����:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1065
      Width           =   765
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BackStyle       =   0  '͸��
      Caption         =   "��������:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1695
      Width           =   765
   End
End
Attribute VB_Name = "FrmA022Mx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim w_mtr_no As String '���ϱ���


Public Property Let Get_Mtr_No(Mtr_No As String)
w_mtr_no = Mtr_No
End Property


Private Sub Cmd_OK_MouseClick()

If check_ok() Then
    Call Upd_Data
    
End If

End Sub




Private Sub cmd_brow_Click()
With FrmIllList
    .G_Ill_Filter = " Ill_no like '" & Trim(Ill_No.Text) & "%'"
    .Show vbModal
    If .Ill_No <> "" Then
        Ill_No.Text = .Ill_No
        Ill_Name.Text = .Ill_Name
        Ill_Type.Text = .Ill_Type
    End If
End With


End Sub


Private Sub cmd_cancel_Click()
Unload Me
End Sub

Private Sub Cmd_OK_Click()
If check_ok() Then
    Call Upd_Data
    Unload Me
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If ActiveControl.MultiLine = False Then
        SendKeys "{TAB}"
    End If
    Exit Sub
End If

If Shift = 0 Then
    Select Case KeyCode
    
    Case vbKeyF5               'ȷ��
        Call Cmd_OK_MouseClick
    Case vbKeyEscape           'ȡ��
        Call cmd_cancel_MouseClick
    End Select
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set FrmA05 = Nothing
End Sub

'�������ݿ�
Private Sub Upd_Data()
Dim w_tmp As New ADODB.Recordset


w_tmp.CursorLocation = adUseClient
w_tmp.Open "Select * from mmsta02_2 where Mtr_No='" & w_mtr_no & "'", G_Con, adOpenDynamic, adLockPessimistic, adCmdText
With w_tmp
    .AddNew
    !Mtr_No = w_mtr_no
    !Ill_No = Trim(Ill_No.Text)
    !upd_name = Trim(G_User_Name)
    !upd_date = Get_SQLDATE
   .Update
End With
Set w_tmp = Nothing
Call ClearFields


End Sub



Public Sub ClearFields()
'׼������
Ill_No.Text = ""
Ill_Name.Text = ""
Ill_Type.Text = ""

End Sub

Private Function check_ok() As Boolean
Dim w_tmp As New ADODB.Recordset


If Ill_No.Text = "" Then
    MsgBox "�����ʼ���Ŀ!", 64, g_CON_CTitle
    Ill_No.SetFocus
    check_ok = False
    Exit Function
Else
    '��֤�Ƿ��ظ�
    w_tmp.CursorLocation = adUseClient
    w_tmp.Open " Select ill_no from mmsta02_2 where Mtr_No='" & w_mtr_no & "'" & _
               " and ill_no = '" & Ill_No.Text & "'", G_Con, adOpenForwardOnly, adLockReadOnly, adCmdText
    If w_tmp.EOF = False Then
        MsgBox "�ʼ���Ŀ�ظ�!", 64, g_CON_CTitle
        Ill_No.Text = ""
        Ill_Type.Text = ""
        Ill_Name.Text = ""
        Ill_No.SetFocus
        check_ok = False
        Set w_tmp = Nothing
        Exit Function
    End If
    Set w_tmp = Nothing
End If

check_ok = True
End Function

Private Sub cmd_cancel_MouseClick()
Unload Me
End Sub

