VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmpoInvSh 
   BackColor       =   &H80000014&
   BorderStyle     =   3  '˫�߹̶��Ի�����
   Caption         =   "��Ʒ���ݲ�ѯ(FrmpoInvSh)"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "��ϸ����"
      Size            =   9.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmpoInvSh.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'өĻ����
   Tag             =   "Order Search"
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'ƽ��
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3690
      Left            =   0
      ScaleHeight     =   3660
      ScaleWidth      =   6015
      TabIndex        =   4
      Top             =   0
      Width           =   6045
      Begin VB.ComboBox cb_check 
         Appearance      =   0  'ƽ��
         Height          =   315
         ItemData        =   "FrmpoInvSh.frx":000C
         Left            =   1410
         List            =   "FrmpoInvSh.frx":000E
         TabIndex        =   7
         Top             =   945
         Width           =   1725
      End
      Begin VB.TextBox inv_no 
         Appearance      =   0  'ƽ��
         Height          =   345
         Left            =   1410
         TabIndex        =   6
         Top             =   330
         Width           =   1725
      End
      Begin VB.CommandButton CmdCancel 
         BeginProperty Font 
            Name            =   "��ϸ����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         Picture         =   "FrmpoInvSh.frx":0010
         Style           =   1  'ͼƬ���
         TabIndex        =   3
         Tag             =   "&Cancel"
         Top             =   2565
         Width           =   1185
      End
      Begin VB.CommandButton CmdOK 
         BeginProperty Font 
            Name            =   "��ϸ����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1515
         Picture         =   "FrmpoInvSh.frx":15B2
         Style           =   1  'ͼƬ���
         TabIndex        =   2
         Tag             =   "&OK"
         Top             =   2565
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker date1 
         Height          =   345
         Left            =   1410
         TabIndex        =   0
         Top             =   1515
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "��ϸ����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   61931523
         UpDown          =   -1  'True
         CurrentDate     =   37217
         MaxDate         =   65745
         MinDate         =   32874
      End
      Begin MSComCtl2.DTPicker date2 
         Height          =   345
         Left            =   3540
         TabIndex        =   1
         Top             =   1515
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "��ϸ����"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   61931523
         UpDown          =   -1  'True
         CurrentDate     =   37217
         MaxDate         =   65745
         MinDate         =   32874
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '͸��
         Caption         =   "����״̬:"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   10
         Tag             =   "Order Date:"
         Top             =   990
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '͸��
         Caption         =   "���ݵ���:"
         Height          =   195
         Index           =   4
         Left            =   330
         TabIndex        =   9
         Tag             =   "Order No.:"
         Top             =   390
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '͸��
         Caption         =   "��������:"
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   8
         Tag             =   "Order Date:"
         Top             =   1575
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '͸��
         Caption         =   "__"
         Height          =   300
         Index           =   3
         Left            =   3240
         TabIndex        =   5
         Top             =   1485
         Width           =   210
      End
   End
End
Attribute VB_Name = "FrmpoInvSh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_CaNcel As Boolean 'ָʾ�Ƿ���"ȡ��"��ť
Dim w_Rs As New ADODB.Recordset
Dim W_CallCoNtrol As Object
Dim W_DefTable As String         '�����ѯ�ı�
Dim W_DefField As String         '�����ѯ���ֶ�
Dim W_DefInvDate As String       '����Ӹ����崫���������ֶ�
Dim W_DefInvType As String       '���嵥������
Dim W_suplOrSupl As String       '�ж��ǿͻ����� 'S':����,'C':�ͻ�
Public G_File As String

'���ݿؼ�
Property Get CallCoNtrol() As Object
Set CallCoNtrol = W_CallCoNtrol
End Property

Property Set CallCoNtrol(p_CallCoNtrol As Object)
Set W_CallCoNtrol = p_CallCoNtrol
End Property
'��ѯ�ı�
Public Property Get DefTable() As String
DefTable = W_DefTable
End Property

Public Property Let DefTable(NewTable As String)
W_DefTable = NewTable
End Property
'��ѯ���ֶ�
Public Property Get DefField() As String
DefField = W_DefField
End Property

Public Property Let DefField(NewField As String)
W_DefField = NewField
End Property
'��ѯ�������ֶ�
Public Property Get DefInvDate() As String
DefInvDate = W_DefInvDate
End Property

Public Property Let DefInvDate(NewInvDate As String)
W_DefInvDate = NewInvDate
End Property
'�жϳ��̻�ͻ�
Public Property Get suplOrSupl() As String
suplOrSupl = W_suplOrSupl
End Property

Public Property Let suplOrSupl(NewsuplOrSupl As String)
W_suplOrSupl = NewsuplOrSupl
If NewsuplOrSupl = "C" Then
    Label1(1).Caption = "�ͻ����:"
    p_line_no.Visible = False
ElseIf NewsuplOrSupl = "S" Then
    Label1(1).Caption = "���̱��:"
    p_line_no.Visible = False
ElseIf NewsuplOrSupl = "P" Then
'    Label1(1).Caption = "��  ��  ��:"
'    cmd_brow.Visible = False
'    supl_no.Visible = False
'    Call AddRsToList(p_line_no, "SELECT p_line_name FROM mmst811 order by p_line_name")
End If
End Property

Public Property Get DefInvType() As String
DefInvType = W_DefInvType
End Property

Public Property Let DefInvType(NewInvType As String)
W_DefInvType = NewInvType
End Property

'���������ļ�
Property Get ClickCaNcel() As Boolean 'Ϊ��ʱ��ʾ����"ȡ��"
ClickCaNcel = W_CaNcel
End Property

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 34
End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And ActiveControl.Name <> "note" Then
    SendKeys "{TAB}"
End If
If Shift = 0 Then
    Select Case KeyCode
    Case vbKeyF5               'ȷ��
         If CmdOK.Enabled = True Then
             Call CmdOK_Click
         End If
    Case vbKeyEscape           'ȡ��
         If CmdCancel.Enabled = True Then
             Call CmdCaNcel_Click
         End If
    End Select
End If
End Sub

Private Sub Form_Load()
date2.Value = Get_SQLDATE
date1.Value = Get_SQLDATE - 30
W_CaNcel = False
Me.KeyPreview = True
End Sub

Private Sub CmdOK_Click()
Dim W_inv_no As String
Dim W_Supl_No As String
Dim W_SQL As String
'On Error Resume Next
W_inv_no = Trim(Inv_No.Text)

'�ж϶�������
If Not IsNull(date1.Value) And Not IsNull(date2.Value) Then
    W_SQL = " AND " & W_DefInvDate & " BETWEEN '" & _
            Format(date1.Value, "m/d/yyyy") & "' AND '" & _
            Format(date2.Value, "m/d/yyyy") & "' "
ElseIf Not IsNull(date1.Value) And IsNull(date2.Value) Then
    W_SQL = " AND " & W_DefInvDate & " >='" & Format(date1.Value, "m/d/yyyy") & "' "
ElseIf IsNull(date1.Value) And Not IsNull(date2.Value) Then
    W_SQL = " AND " & W_DefInvDate & " <='" & Format(date2.Value, "m/d/yyyy") & "' "
End If

'��������
If W_inv_no <> "" Then
   W_SQL = " AND " & W_DefField & " LIKE '" & W_inv_no & "%'" & W_SQL
' Else
'   W_SQL = " AND " & W_DefField & " LIKE '" & W_inv_No & "%'" & W_SQL
' End If
End If
'��������
If W_DefInvType <> "" Then
    W_SQL = " AND " & W_DefInvType & "  " & W_SQL
End If

'�жϵ���״̬
If G_File = "1" Then
   If cb_check.ListIndex = 0 Then
        W_SQL = " status IN ('0') " & W_SQL
    ElseIf cb_check.ListIndex = 1 Then
        W_SQL = " status In ('2') " & W_SQL
    Else
        W_SQL = " status IN ('0','2') " & W_SQL
    End If
Else
    If cb_check.ListIndex = 0 Then
        W_SQL = " status IN ('0') " & W_SQL
    ElseIf cb_check.ListIndex = 1 Then
        W_SQL = " status In ('1') " & W_SQL
    ElseIf cb_check.ListIndex = 2 Then
        W_SQL = " status In ('2') " & W_SQL
    Else
        W_SQL = " status IN ('0','1','2') " & W_SQL
    End If
End If
W_SQL = "SELECT " & W_DefField & " FROM " & W_DefTable & " WHERE " & W_SQL & " ORDER BY " & W_DefField

w_Rs.CursorLocation = adUseClient
w_Rs.Open W_SQL, G_Con, adOpenDynamic

'���ص�ComboBox��
W_CallCoNtrol.Clear
Do Until w_Rs.EOF
    W_CallCoNtrol.AddItem w_Rs.Fields(W_DefField)
    w_Rs.MoveNext
Loop
w_Rs.Close
W_CaNcel = False
Unload Me
End Sub

Private Sub cmd_brow_Click()

    With FrmSuplList
        
        .G_Supl_Filter = Trim(Supl_No.Text)
        
        .Show vbModal
        If .Supl_No <> "" Then
            Supl_No.Text = .Supl_No
        End If
    End With

End Sub

Private Sub CmdCaNcel_Click()
W_CaNcel = True
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set w_Rs = Nothing
G_File = ""
Set FrmpoInvSh = Nothing
End Sub


