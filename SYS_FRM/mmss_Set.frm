VERSION 5.00
Begin VB.Form mmss_set 
   Appearance      =   0  'ƽ��
   BackColor       =   &H80000004&
   BorderStyle     =   1  '���߹̶�
   Caption         =   "����趨"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "mmss_Set.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5130
   Begin VB.Frame Frame3 
      Height          =   1485
      Left            =   2850
      TabIndex        =   12
      Top             =   3120
      Width           =   2265
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "ȡ��(&N)"
         Height          =   375
         Left            =   540
         TabIndex        =   8
         Top             =   900
         Width           =   1155
      End
      Begin VB.CommandButton cmd_ok 
         Caption         =   "ȷ��(&Y)"
         Height          =   375
         Left            =   540
         TabIndex        =   7
         Top             =   270
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3075
      Left            =   2850
      TabIndex        =   9
      Top             =   60
      Width           =   2265
      Begin VB.CommandButton Cmd_ALL1 
         Caption         =   "ȫ��"
         Height          =   300
         Left            =   1500
         TabIndex        =   3
         Top             =   810
         Width           =   600
      End
      Begin VB.CommandButton Cmd_ALL2 
         Caption         =   "ȫ��"
         Height          =   300
         Left            =   1500
         TabIndex        =   5
         Top             =   1650
         Width           =   600
      End
      Begin VB.ComboBox head_Aligment 
         Height          =   300
         Left            =   150
         TabIndex        =   2
         Top             =   810
         Width           =   1335
      End
      Begin VB.TextBox col_width 
         Height          =   300
         Left            =   150
         TabIndex        =   6
         Top             =   2490
         Width           =   1935
      End
      Begin VB.ComboBox col_Aligment 
         Height          =   300
         Left            =   150
         TabIndex        =   4
         Top             =   1650
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "����ԡ���ʽ:"
         Height          =   285
         Left            =   150
         TabIndex        =   13
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "��λ���:"
         Height          =   375
         Left            =   150
         TabIndex        =   11
         Top             =   2130
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "��λ�ԡ���ʽ:"
         Height          =   285
         Left            =   150
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4545
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   2835
      Begin VB.ListBox List1 
         Height          =   4260
         Left            =   60
         Style           =   1  '��Ŀ������ȡ����
         TabIndex        =   1
         Top             =   210
         Width           =   2715
      End
   End
   Begin VB.Menu menu_modify 
      Caption         =   "modify"
      Visible         =   0   'False
      Begin VB.Menu menu_edit 
         Caption         =   "�޸�"
      End
   End
End
Attribute VB_Name = "mmss_set"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tmp_Gridc(127) As Grid_Data
Dim W_C_File As String '���� INI �ļ���
Dim W_Form As Form     '��
Dim W_FormName As String
Dim W_GridName As String

'��������ȡ�� INI �ļ���
'*******************************************************************
'���� INI �ļ�
Public Property Get Gridc_File() As String
    Gridc_File = W_C_File
End Property

Public Property Let Gridc_File(New_cFile As String)
    W_C_File = New_cFile
End Property
'FormName
Public Property Get Get_FormName() As String
    Get_FormName = W_FormName
End Property

Public Property Let Get_FormName(New_FormName As String)
    W_FormName = New_FormName
End Property

'GridName
Public Property Get Get_GridName() As String
    Get_GridName = W_GridName
End Property

Public Property Let Get_GridName(New_GridName As String)
    W_GridName = New_GridName
End Property

'��
Public Property Get Parent_form() As Form
    Set Parent_form = W_Form
End Property

Public Property Set Parent_form(New_Form As Form)
    Set W_Form = New_Form
End Property

'***************************************************************************

Private Sub Form_Load()
    '��������
    With Me
        .ScaleMode = vbPixels
        .Left = Int(Screen.Width \ Screen.TwipsPerPixelX - Me.ScaleWidth) * 15 / 2
        .Top = Int(Screen.Height \ Screen.TwipsPerPixelY - Me.ScaleHeight - 120) * 15 / 2
        .ScaleMode = vbTwips
    End With
'************************************************************
'***�趨Ĭ��INI�ļ�,���ͱ��ؼ� ************
'************************************************************
    If W_C_File = "" Then
        W_C_File = "sys_gridc.ini"
    End If
    If W_GridName = "" Then
        W_GridName = G_Grid.Name
    End If
    If W_FormName = "" Then
        W_FormName = G_Form.Name
    End If
    If W_Form Is Nothing Then
        Set W_Form = G_Form
    End If
'************************************************************
    Call init_form
End Sub

Sub init_form() '��ʼ����
'ˢ�� ComboBox �ؼ�
col_Aligment.Clear
head_Aligment.Clear
col_Aligment.AddItem "���϶ԡ�"
col_Aligment.AddItem "���жԡ�"
col_Aligment.AddItem "���¶ԡ�"
col_Aligment.AddItem "���϶ԡ�"
col_Aligment.AddItem "���жԡ�"
col_Aligment.AddItem "���¶ԡ�"
col_Aligment.AddItem "���϶ԡ�"
col_Aligment.AddItem "���жԡ�"
col_Aligment.AddItem "ͨ    ��"

head_Aligment.AddItem "���϶ԡ�"
head_Aligment.AddItem "���жԡ�"
head_Aligment.AddItem "���¶ԡ�"
head_Aligment.AddItem "���϶ԡ�"
head_Aligment.AddItem "���жԡ�"
head_Aligment.AddItem "���¶ԡ�"
head_Aligment.AddItem "���϶ԡ�"
head_Aligment.AddItem "���жԡ�"
head_Aligment.AddItem "ͨ    ��"

Cmd_ALL1.Caption = "ȫ��"
Cmd_ALL2.Caption = "ȫ��"
        
'�����ʱ����
For i = 0 To 127
    Tmp_Gridc(i).Grid_DataField = ""
Next i
    
'�� INI �ļ�ȡ������
Call GetGridSetting(W_FormName, W_GridName, Tmp_Gridc, W_C_File)
    
'ˢ�� List ����
List1.Clear
For i = 0 To 127
    If Tmp_Gridc(i).Grid_DataField <> "" Then
        List1.AddItem Tmp_Gridc(i).Grid_Caption
        If Mid(Tmp_Gridc(i).Grid_Visible, 1, 1) = "T" Then List1.Selected(i) = True
    Else
        Exit For
    End If
Next i
    
'�� List �ó�ֵ
If List1.ListCount > 0 Then
    List1.ListIndex = 0
    Cmd_ALL1.Enabled = True
    Cmd_ALL2.Enabled = True
Else
    Cmd_ALL1.Enabled = False
    Cmd_ALL2.Enabled = False
End If
End Sub

'*****************************************************************
'��������¼�
'�������жԡ���ʽ(����ͷ)
Private Sub Cmd_ALL1_Click()
If List1.ListCount > 0 And head_Aligment.ListIndex >= 0 Then
    For i = 0 To List1.ListCount - 1
        Tmp_Gridc(i).Grid_HeadAligment = head_Aligment.ListIndex
    Next i
End If
End Sub
'(������)
Private Sub Cmd_ALL2_Click()
If List1.ListCount > 0 And col_Aligment.ListIndex >= 0 Then
    For i = 0 To List1.ListCount - 1
        Tmp_Gridc(i).Grid_ColAligment = col_Aligment.ListIndex
    Next i
End If
End Sub
'ȷ����ȡ��
Private Sub cmd_ok_Click()
'�洢�޸�����
Call SaveGridSetting(W_FormName, W_GridName, Tmp_Gridc, W_C_File)
On Error Resume Next
Call W_Form.Form_Activate
Unload Me
End Sub

Private Sub cmd_cancel_Click()
    Unload Me
End Sub

'***************************************************************************
'���ؼ��¼�
'���ԡ��ı�ʱ
Private Sub col_Aligment_Click()
If List1.ListCount > 0 And col_Aligment.ListIndex >= 0 Then
    Tmp_Gridc(List1.ListIndex).Grid_ColAligment = col_Aligment.ListIndex
End If
End Sub
'�п�ȸı�ʱ
Private Sub col_width_Change()
If List1.ListCount > 0 Then
    Tmp_Gridc(List1.ListIndex).Grid_Width = col_width.Text
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set W_Form = Nothing
Set W_Grid = Nothing
Set mmss_set = Nothing
End Sub

'�жԡ��ı�ʱ
Private Sub head_Aligment_Click()
If List1.ListCount > 0 And head_Aligment.ListIndex >= 0 Then
    Tmp_Gridc(List1.ListIndex).Grid_HeadAligment = head_Aligment.ListIndex
End If
End Sub

'�������λʱ
Private Sub List1_Click()
Dim W_False As Boolean
    
W_False = False
If List1.ListCount > 0 Then
    col_width.Text = Round(Val(Tmp_Gridc(List1.ListIndex).Grid_Width), 2)
    head_Aligment.ListIndex = Tmp_Gridc(List1.ListIndex).Grid_HeadAligment
    col_Aligment.ListIndex = Tmp_Gridc(List1.ListIndex).Grid_ColAligment
    If List1.Selected(List1.ListIndex) = True Then
        Tmp_Gridc(List1.ListIndex).Grid_Visible = "True"
    Else
        Tmp_Gridc(List1.ListIndex).Grid_Visible = "False"
    End If
End If
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) = True Then
        W_False = False
        Exit For
    Else
        W_False = True
    End If
Next i
If W_False Then
    MsgBox "�벻Ҫ��ȫ����λ����", vbOKOnly, g_CON_CTitle
    List1.Selected(0) = True
End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If List1.ListCount > 0 And Button = 2 Then
    If List1.ListIndex >= 0 Then
        PopupMenu Me.menu_modify
    End If
End If
End Sub

Private Sub menu_edit_Click()
Dim W_Default As String
Dim W_Modify As String

W_Default = Trim(List1.List(List1.ListIndex))
W_Modify = InputBox("�޸ı���ʶ����", "�޸ı�ʶ", W_Default)
If W_Modify <> "" Then '�޸ĺ�
    List1.List(List1.ListIndex) = W_Modify
    Tmp_Gridc(List1.ListIndex).Grid_Caption = W_Modify
End If
End Sub
