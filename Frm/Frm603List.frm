VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form Frm603List 
   BorderStyle     =   4  '单线固定工具视窗
   Caption         =   "通知单资料列表"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '萤幕中央
   Tag             =   "Customers List"
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5820
      Left            =   0
      ScaleHeight     =   5790
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   0
      Width           =   8265
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   4815
         Top             =   45
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Height          =   345
         Index           =   1
         Left            =   1350
         Picture         =   "Frm603List.frx":0000
         Style           =   1  '图片外观
         TabIndex        =   2
         Tag             =   "&Cancel"
         Top             =   150
         Width           =   1155
      End
      Begin VB.CommandButton Command1 
         Default         =   -1  'True
         Height          =   345
         Index           =   0
         Left            =   90
         Picture         =   "Frm603List.frx":15A2
         Style           =   1  '图片外观
         TabIndex        =   1
         Tag             =   "&OK"
         Top             =   150
         Width           =   1155
      End
      Begin VSFlex7Ctl.VSFlexGrid Grid1 
         Height          =   5205
         Left            =   0
         TabIndex        =   3
         Top             =   600
         Width           =   8220
         _cx             =   14499
         _cy             =   9181
         _ConvInfo       =   -1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新细明体"
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
End
Attribute VB_Name = "Frm603List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private W_Mo_No As String       '制单单号
Private W_Mo_Date As String     '制单日期
Private W_Order_No As String    '订单单号
Private W_Cust_Order_No As String   '客户订单号
Private W_Cust_Name As String       '客户名称
Private W_Cust_No As String         '客户编号
Private W_Order_Amt As Double       '订单数量
Private W_Mtr_Amt As Double         '制单数量
Private W_Cust_Mtr_No As String     '客户料号
Private W_Dele_Amt As Double        '已出货数量
Private W_Diff_Amt As Double        '差量
Private W_Close_Date As String      '结关日期
Private w_mtr_no As String          '成品编号
Private w_mtr_name As String        '品名
Private w_mtr_dim As String         '规格
Private W_Unit_Name As String       '单位
Private W_Color_Name As String      '颜色

Private W_Calcel_Status As Boolean

Public G_Filter As String

Dim W_CallForm As Form

Public Property Get CallForm() As Form
Set CallForm = W_CallForm
End Property
Public Property Set CallForm(f As Form)
Set W_CallForm = f
End Property
Public Property Get Mo_No() As String
    Mo_No = W_Mo_No
End Property
Public Property Get Mo_Date() As String
    Mo_Date = W_Mo_Date
End Property

Public Property Get Order_No() As String
    Order_No = W_Order_No
End Property
Public Property Get Cust_Order_No() As String
    Cust_Order_No = W_Cust_Order_No
End Property
Public Property Get Cust_No() As String
    Cust_No = W_Cust_No
End Property
Public Property Get Cust_Name() As String
    Cust_Name = W_Cust_Name
End Property

Public Property Get mtr_no() As String
    mtr_no = w_mtr_no
End Property
Public Property Get Mtr_Dim() As String
    Mtr_Dim = w_mtr_dim
End Property

Public Property Get Unit_Name() As String
    Unit_Name = W_Unit_Name
End Property
Public Property Get Color_Name() As String
    Color_Name = W_Color_Name
End Property
Public Property Get Order_Amt() As Double   '订单数量
    Order_Amt = W_Order_Amt
End Property
Public Property Get Mtr_Amt() As Double     '制单数量
    Mtr_Amt = W_Mtr_Amt
End Property

Public Property Get Close_Date() As String
    Close_Date = W_Close_Date
End Property
Public Property Get mtr_name() As String
    mtr_name = w_mtr_name
End Property
Public Property Get Cust_Mtr_No() As String
    Cust_Mtr_No = W_Cust_Mtr_No
End Property
Public Property Get Cancel_Status() As Boolean
    Cancel_Status = W_Calcel_Status
End Property
Public Property Get Deliv_Amt() As Double
    Deliv_Amt = W_Deliv_Amt
End Property
Public Property Get Diff_Amt() As Double
    Diff_Amt = W_Diff_Amt
End Property

Private Sub Command1_Click(Index As Integer)
Dim Temp As New ADODB.Recordset
Dim W_Row As Long
On Error Resume Next

    If Index = 0 Then
        W_Row = Grid1.Row
        If W_Row > 0 Then
            With Grid1
                W_Mo_No = .TextMatrix(W_Row, 0)
                W_Mo_Date = .TextMatrix(W_Row, 1)
                W_Close_Date = .TextMatrix(W_Row, 2)
                W_Cust_Order_No = .TextMatrix(W_Row, 3)
                W_Order_No = .TextMatrix(W_Row, 4)
                W_Cust_No = .TextMatrix(W_Row, 5)
                W_Cust_Name = .TextMatrix(W_Row, 6)

                w_mtr_no = .TextMatrix(W_Row, 7)
                W_Cust_Mtr_No = .TextMatrix(W_Row, 8)
                w_mtr_name = .TextMatrix(W_Row, 9)
                w_mtr_dim = .TextMatrix(W_Row, 10)
                W_Color_Name = .TextMatrix(W_Row, 11)
                W_Unit_Name = .TextMatrix(W_Row, 12)
                W_Order_Amt = .TextMatrix(W_Row, 13)
                W_Mtr_Amt = .TextMatrix(W_Row, 14)
                W_Deliv_Amt = .TextMatrix(W_Row, 15)
                W_Diff_Amt = .TextMatrix(W_Row, 16)
                
                Set Temp = Nothing
            End With
            W_Calcel_Status = False
        Else
           W_Calcel_Status = False
           W_Order_No = ""
           W_Cust_Order_No = ""
        End If
    Else
        W_Calcel_Status = False
        W_Order_No = ""
        W_Cust_Order_No = ""
    End If
    
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 34
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 0 Then
    Select Case KeyCode
    Case vbKeyF5               '
       Call Command1_Click(0)
    Case vbKeyEscape           '
       Call Command1_Click(1)
    End Select
End If
End Sub

Private Sub Form_Load()
    '加载列量
    Dim W_Rs As New ADODB.Recordset
    Dim W_Str As String
    W_Calcel_Status = False
    Set Me.Picture1.Picture = G_MDIForm.Picture
   
    
       W_Rs.Open " Select a.Mo_No as 批号 ,a.Mo_Date as 单据日期 ,a.close_date as 结关日期 ,a.Cust_Order_No as 客户订单,a.Order_No as 订单单号, " & _
                        " b.Cust_No as 客户编号 ,d.Cust_Name as 客户名称 ,a.Mtr_No as 产品料号, " & _
                        " b.Cust_Mtr_No as 客户料号  ,c.Mtr_Name as 品名 ,c.Mtr_Dim as 规格 ,b.Color_Name  as 颜色描述," & _
                        " c.unit_name as 单位  , b.Mtr_Amt as  订单数量 ,a.Mtr_Amt as 制单数量 ,isnull(b.Deliv_Amt,0)  as 出货数量," & _
                        " (b.Mtr_Amt -isnull(b.Deliv_Amt,0)) as 差量 , a.Remark  as 备注" & _
                 " From mmst401_m a Inner Join mmst012 b On a.Order_No=b.Order_No And a.Cust_Order_No=b.Cust_Order_No And a.Mtr_No=b.Mtr_No " & _
                                  " Inner Join mmsp611 c On a.Mtr_No=c.Mtr_No " & _
                                  " Inner Join mmst021 d On b.Cust_No=d.Cust_No " & _
                 " Where a.Status='2'  " & G_Filter & "", G_Con

    Set Grid1.DataSource = W_Rs
    
   '设定表格宽度
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
    
    Me.KeyPreview = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
G_Order_Filter = ""
G_Filter = ""
W_S_Order_No = ""
'W_Mtr_No = ""


End Sub

Private Sub Grid1_DblClick()
    Call Command1_Click(0)
End Sub

Private Sub Timer1_Timer()
If Grid1.Rows = 2 Then
    Call Command1_Click(0)
End If
End Sub
