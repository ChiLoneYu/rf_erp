VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form Frm50AList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "加工单号列表"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   13440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Order List"
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   4815
      Top             =   75
   End
   Begin VB.CommandButton Command1 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   30
      Picture         =   "Frm50AList.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "&OK"
      Top             =   120
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   1290
      Picture         =   "Frm50AList.frx":15A2
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "&Cancel"
      Top             =   120
      Width           =   1155
   End
   Begin VSFlex7Ctl.VSFlexGrid Grid1 
      Height          =   7710
      Left            =   0
      TabIndex        =   2
      Top             =   585
      Width           =   13380
      _cx             =   23601
      _cy             =   13600
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
Attribute VB_Name = "Frm50AList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private W_Inv_No As String
Private W_Inv_Need_No As String
Private W_Cancel_Click As Boolean
Private W_Date As Date

Private W_Po As Boolean

Public g_Inv_Filter As String  '条件
Public G_Filter As String  '　　条件

Private W_Mo_No As String
Private W_P_Mtr As String
Private W_Cust_Mtr As String
Private w_mtr_name As String
Private w_mtr_dim As String
Private w_mtr_no As String
Private W_Mtr_Name1 As String
Private W_Mtr_Dim1 As String
Private W_Plan_No As String
Private W_order_no As String
Private W_Cust_Order_No As String
Private W_Po_Name As String
               
                
Public Property Get inv_no() As String
inv_no = W_Inv_No
End Property

Public Property Get Inv_Need_No() As String
Inv_Need_No = W_Inv_Need_No
End Property



Public Property Get inv_date() As Date
inv_date = W_Date
End Property

Public Property Get Supl_Name() As String
Supl_Name = W_Supl_Name
End Property


'************************************
Public Property Get mo_no() As String
mo_no = W_Mo_No
End Property

Public Property Get P_Mtr() As String
P_Mtr = W_P_Mtr
End Property

Public Property Get cust_mtr() As String
cust_mtr = W_Cust_Mtr
End Property

Public Property Get mtr_name() As String
mtr_name = w_mtr_name
End Property

Public Property Get Mtr_Dim() As String
Mtr_Dim = w_mtr_dim
End Property

Public Property Get mtr_no() As String
mtr_no = w_mtr_no
End Property

Public Property Get Mtr_Name1() As String
Mtr_Name1 = W_Mtr_Name1
End Property

Public Property Get Mtr_Dim1() As String
Mtr_Dim1 = W_Mtr_Dim1
End Property

Public Property Get Plan_No() As String
Plan_No = W_Plan_No
End Property

Public Property Get Order_No() As String
Order_No = W_order_no
End Property

Public Property Get Po_Name() As String
Po_Name = W_Po_Name
End Property

Public Property Get Cust_Order_No() As String
Cust_Order_No = W_Cust_Order_No
End Property

Public Property Get Cancel_Click() As Boolean
Cancel_Click = W_Cancel_Click
End Property

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
    Dim W_Row As Long
    W_Row = Grid1.Row
    If G_Filter = "50A" Then
        If W_Row > 0 Then
            With Grid1
                W_Cancel_Click = False
                W_Inv_No = .TextMatrix(W_Row, 0)
                W_inv_date = .TextMatrix(W_Row, 1)
     
            End With
        End If

                
    ElseIf G_Filter = "50AP" Then
        If W_Row > 0 Then
            With Grid1
                W_Cancel_Click = False
                W_Mo_No = .TextMatrix(W_Row, 0)
              
                W_Plan_No = .TextMatrix(W_Row, 1)
                W_order_no = .TextMatrix(W_Row, 2)
                W_Cust_Order_No = .TextMatrix(W_Row, 3)
                
            End With
        End If
    ElseIf G_Filter = "50AP1" Then
        If W_Row > 0 Then
            With Grid1
                W_Cancel_Click = False
                W_Mo_No = .TextMatrix(W_Row, 0)
              
                W_Plan_No = .TextMatrix(W_Row, 1)
                W_order_no = .TextMatrix(W_Row, 2)
                W_Cust_Order_No = .TextMatrix(W_Row, 3)
                
                W_P_Mtr = .TextMatrix(W_Row, 4)
                W_Cust_Mtr = .TextMatrix(W_Row, 5)
                w_mtr_name = .TextMatrix(W_Row, 6)
                w_mtr_dim = .TextMatrix(W_Row, 7)
         
                
            End With
        End If
     ElseIf G_Filter = "50AMTR" Then
        If W_Row > 0 Then
            With Grid1
                W_Cancel_Click = False
                W_Mo_No = .TextMatrix(W_Row, 0)
                W_Po_Name = .TextMatrix(W_Row, 1)
                W_P_Mtr = .TextMatrix(W_Row, 2)
                W_Cust_Mtr = .TextMatrix(W_Row, 3)
                w_mtr_name = .TextMatrix(W_Row, 4)
                w_mtr_dim = .TextMatrix(W_Row, 5)
                w_mtr_no = .TextMatrix(W_Row, 6)
                W_Mtr_Name1 = .TextMatrix(W_Row, 7)
                W_Mtr_Dim1 = .TextMatrix(W_Row, 8)
                W_Plan_No = .TextMatrix(W_Row, 9)
                W_order_no = .TextMatrix(W_Row, 10)
                W_Cust_Order_No = .TextMatrix(W_Row, 11)
                
            End With
        End If
    Else
        If W_Row > 0 Then
            With Grid1
                W_Cancel_Click = False
                W_Inv_No = .TextMatrix(W_Row, 0)
                W_inv_date = .TextMatrix(W_Row, 1)
                W_Supl_Name = .TextMatrix(W_Row, 2)
            End With
        End If
    End If
End If
If Index = 1 Then
    W_Cancel_Click = True
    W_Inv_No = ""
    W_Supl_Name = ""
End If
Unload Me
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
Dim W_Str As String
Dim w_rs As New ADODB.Recordset

Set Me.Picture = G_MDIForm.Picture
If G_Filter = "50A" Then '加工单

   W_Str = "SELECT DISTINCT Inv_No As 加工单号, " & _
              "Inv_Date As 下单日期 " & _
        "FROM mmst413   " & _
        "WHERE  status='2' " & _
          " AND " & g_Inv_Filter & _
           "ORDER BY Inv_No DESC "
           

           
ElseIf G_Filter = "50AP" Then
  W_Str = " SELECT distinct mmst414.Mo_No as 通知单号,mmst414.Plan_No as 排程单号,  " & _
                " mmst414.Order_No as 订单单号,mmst414.Cust_Order_No as 客户订单  " & _
          " FROM  mmst414 INNER JOIN " & _
                " mmst611 ON mmst414.Mtr_No = mmst611.Mtr_No  " & _
                " AND " & g_Inv_Filter & _
           " "
ElseIf G_Filter = "50AP1" Then
  W_Str = " SELECT distinct mmst414.Mo_No as 通知单号,mmst414.Plan_No as 排程单号,  " & _
                " mmst414.Order_No as 订单单号,mmst414.Cust_Order_No as 客户订单,  " & _
                " mmst414.P_Mtr as 成品料号, mmst414.Cust_Mtr_No as 客户料号,mmst611.Mtr_Name as 成品名称, mmst611.Mtr_Dim as 成品规格 " & _
          " FROM  mmst414 INNER JOIN " & _
                " mmst611 ON mmst414.Mtr_No = mmst611.Mtr_No  " & _
                " AND " & g_Inv_Filter & _
           " "
ElseIf G_Filter = "50AMTR" Then
  W_Str = " SELECT mmst414.Mo_No as 通知单号,  " & _
                " mmst414.Po_Name as 加工工序,mmst414.P_Mtr as 成品料号, mmst414.Cust_Mtr_No as 客户料号,mmst611.Mtr_Name as 成品名称, mmst611.Mtr_Dim as 成品规格,  " & _
                " mmst414.Mtr_No as 半成品料号,mmst611_1.Mtr_Name AS 半成品名称, mmst611_1.Mtr_Dim AS 半成品规格, " & _
                " mmst414.Plan_No as 排程单号, mmst414.Order_No as 订单单号, " & _
                " mmst414.Cust_Order_No as 客户订单 " & _
          " FROM  mmst414 INNER JOIN " & _
                " mmst611 ON mmst414.Mtr_No = mmst611.Mtr_No INNER JOIN " & _
                " mmst611 mmst611_1 ON mmst414.Mtr_No = mmst611_1.Mtr_No " & _
                " AND " & g_Inv_Filter & _
           "ORDER BY mmst414.Mo_No,mmst414.P_Mtr,mmst414.Mtr_No DESC "
End If
            
w_rs.Open W_Str, G_Con
Set Grid1.DataSource = w_rs

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
Set FrmInvList = Nothing
g_Inv_Filter = ""
G_Filter = ""
End Sub

Private Sub Grid1_DblClick()
    Call Command1_Click(0)
End Sub

Private Sub Timer1_Timer()
If Grid1.Rows = 2 Then
    Call Command1_Click(0)
End If
End Sub



