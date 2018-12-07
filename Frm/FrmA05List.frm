VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form FrmA05List 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ձ��б�(FrmA05List)"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   12420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Material List"
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   4815
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   1290
      Picture         =   "FrmA05List.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "&Cancel"
      Top             =   90
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   30
      Picture         =   "FrmA05List.frx":15A2
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "&OK"
      Top             =   90
      Width           =   1155
   End
   Begin VSFlex7Ctl.VSFlexGrid Grid1 
      Height          =   5250
      Left            =   0
      TabIndex        =   2
      Top             =   510
      Width           =   12360
      _cx             =   21802
      _cy             =   9260
      _ConvInfo       =   -1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16777215
      ForeColorFixed  =   16711680
      BackColorSel    =   49152
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483643
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16711680
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   13
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   1
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "FrmA05List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�����ľֲ�����
Private W_Inv_No As String
Private W_Dup_Name As String
Private W_Remark As String

Public G_Inv_Filter As String '����

Dim W_CallForm As Form

'ֻ������
Public Property Get Get_InvNo() As String '��ⵥ��
    Get_InvNo = W_Inv_No
End Property

Public Property Get CallForm() As Form
Set CallForm = W_CallForm
End Property

Public Property Set CallForm(f As Form)
Set W_CallForm = f
End Property

Public Property Get Get_DupName() As String 'Ʒ��
    Get_DupName = W_Dup_Name
End Property

Public Property Get Get_Remark() As String
    Get_Remark = W_Remark
End Property

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
    Dim W_Row As Long
    W_Row = Grid1.Row
    If W_Row > 0 Then
        With Grid1
            W_Inv_No = .TextMatrix(W_Row, 0)
        End With
    End If
End If
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
    
    Case vbKeyF5               'ȷ��
        Call Command1_Click(0)
    Case vbKeyEscape           'ȡ��
        Call Command1_Click(1)
    End Select
End If
End Sub

Private Sub Form_Load()
Set Me.Picture = GetMdiForm.Picture
'�����б�
'On Error Resume Next
Dim W_RS As New ADODB.Recordset
W_RS.CursorLocation = adUseClient
W_RS.Open "select  a.inv_no as ��ⵥ��,a.inv_date �������,d.p_line_name �����߱�,c.order_no ������,c.cust_name �ͻ�����," & _
" c.cust_order_no �ͻ�����,c.cust_mtr_no �ͻ��Ϻ�,c.prod_name ��Ʒ����,c.prod_dim ��Ʒ���,c.color_script ��ɫ����,b.mtr_amt ������� from mmst531 a inner join mmst532 b on a.inv_No=b.inv_No " & _
" inner join mmsp011 c on c.order_no=b.order_no and b.mtr_no=c.mtr_no " & _
" inner join mmst811 d on d.p_LIne_no=a.p_Line_No " & _
              " and " & G_Inv_Filter & " order by a.inv_No desc ", G_Con, , , adCmdText

Set Grid1.DataSource = W_RS

'�趨�����
With Grid1
    .AutoResize = True
    For i = 1 To .Cols - 1
       .AutoSize (i)
    Next
    For i = 0 To .Rows - 1
        .RowHeight(i) = 350
    Next
End With

If Grid1.Rows = 2 Then
   Timer1.Enabled = True
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'���������ı�����ֵ,��Ϊ�еĿͻ�����ʱ��δ�������е�����
Set FrmA05List = Nothing
End Sub

Private Sub Grid1_DblClick()
Call Command1_Click(0)
End Sub

Private Sub Timer1_Timer()
If Grid1.Rows = 2 Then
    Call Command1_Click(0)
End If
End Sub
