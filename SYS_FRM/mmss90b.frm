VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form mmss90b 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "单据审核消息反馈授权(90b)"
   ClientHeight    =   6180
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   16335
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   907
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   16335
   Begin VSFlex7Ctl.VSFlexGrid TDBGrid1 
      Bindings        =   "mmss90b.frx":0000
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16335
      _cx             =   28813
      _cy             =   11033
      _ConvInfo       =   -1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483634
      ForeColorFixed  =   -2147483630
      BackColorSel    =   65280
      ForeColorSel    =   -2147483641
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"mmss90b.frx":0015
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   5
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
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   3240
         Top             =   1920
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "mmss90b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ST_Tmp As New ADODB.Recordset
Dim W_Text As String
Dim User_List As String
    Dim X() As String
Private Sub Form_Load()
Dim ST_901 As New ADODB.Recordset

With mmss90b
    .Left = Int(sys_main.Width - mmss90b.Width) / 2
    .Top = Int(sys_main.Height - mmss90b.Height - 1500) / 2
End With
Call Set_Color_Frm(Me)
Call Init_Grid

Call RefreshGrid

End Sub
Private Sub RefreshGrid() 'W_User_List

    Set ST_Tmp = Open_Rs("Select inv_type,user_list,reset_list,inv_list from mmst90b where pc_name='" & G_Pc_Name & "'")
    
    Set Adodc1.Recordset = ST_Tmp
    Set TDBGrid1.DataSource = Adodc1
    Call readactive
    For i = 1 To TDBGrid1.Rows
        TDBGrid1.RowHeight(i - 1) = 350
        If i < TDBGrid1.Rows Then
            TDBGrid1.TextMatrix(i, 0) = i
        End If
    Next i
    TDBGrid1.TextMatrix(0, 0) = " No"
    TDBGrid1.ColWidth(0) = 700

    TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
TDBGrid1.MergeCells = flexMergeFree
End Sub
Private Sub readactive()
With TDBGrid1
    TDBGrid1.TextMatrix(0, 1) = "单据名称"
    TDBGrid1.TextMatrix(0, 2) = "审核提示人员"
    TDBGrid1.TextMatrix(0, 3) = "重置提示人员"
    TDBGrid1.TextMatrix(0, 4) = "list_no"
'    TDBGrid1.TextMatrix(0, 4) = "更新人员"
    TDBGrid1.MergeCells = flexMergeNever
    Row_Height = 350
    With TDBGrid1
        .AutoResize = True
        For i = 1 To .Cols - 2
           .AutoSize (i)
        Next
        .ColHidden(4) = True
    End With
End With


End Sub

Private Sub Init_Grid()
 Dim T_INV As String
    Dim T_pRE_INV As String
    Dim Tmp_User As String
    Dim i As Double
    Dim Tmp_RB As New ADODB.Recordset
    
    G_Con.Execute "Delete from mmst90b where pc_name='" & G_Pc_Name & "'"
    '插入审核的
    Set ST_Tmp = Open_Rs("select d_list inv_list,inv_type from mmst905 order by inv_list ")
    
    Do Until ST_Tmp.EOF
            Tmp_User = ""
            T_pRE_INV = ST_Tmp!Inv_Type
            Set Tmp_RB = Open_Rs("Select user_name from mmst907 a inner join mmst901 b on a.user_list=b.list_no where inv_list='" & ST_Tmp!Inv_list & "' and type=0")
                Do Until Tmp_RB.EOF
                    Tmp_User = Tmp_User & Tmp_RB!user_name & ";"
                    Tmp_RB.MoveNext
                Loop
                    If Tmp_User <> "" Then
                    
                        Tmp_User = Left(Tmp_User, Len(Tmp_User) - 1)
                        
                    End If
                    
                        G_Con.Execute "INSERT INTO MMST90b(pc_name,INV_type,USER_LIST,inv_list) " & _
                                        " values ('" & G_Pc_Name & "', '" & T_pRE_INV & "','" & Tmp_User & "'," & ST_Tmp!Inv_list & ") "
             

        ST_Tmp.MoveNext
    Loop
    '插入重置的
    Set ST_Tmp = Open_Rs("select d_list inv_list,inv_type from mmst905 order by inv_list ")
    
    Do Until ST_Tmp.EOF
            Tmp_User = ""
            T_pRE_INV = ST_Tmp!Inv_Type
            Set Tmp_RB = Open_Rs("Select user_name from mmst907 a inner join mmst901 b on a.user_list=b.list_no where inv_list='" & ST_Tmp!Inv_list & "' and type=1")
                Do Until Tmp_RB.EOF
                    Tmp_User = Tmp_User & Tmp_RB!user_name & ";"
                    Tmp_RB.MoveNext
                Loop
                    If Tmp_User <> "" Then
                    
                        Tmp_User = Left(Tmp_User, Len(Tmp_User) - 1)
                    End If
                
                        G_Con.Execute "Update mmst90b set reset_list= '" & Tmp_User & "' where inv_list=" & ST_Tmp!Inv_list & " and pc_name='" & G_Pc_Name & "' "
             
        ST_Tmp.MoveNext
    Loop
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call ResizeListWindow(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
W_Text = ""
Set ST_Tmp = Nothing
End Sub
Private Sub TDBGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
TDBGrid1.ColComboList(2) = "..."
TDBGrid1.ColComboList(3) = "..."

End Sub

Private Sub TDBGrid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'显示当前系统内的用户

Dim i As Long
Dim user_str As String
Dim Tmp_User As String

If Col = 2 Then
    With FrmUserList
         .W_User_name = TDBGrid1.TextMatrix(Row, 2)
          User_List = TDBGrid1.TextMatrix(Row, 2)
         .Show vbModal
         If .cancel_status = False Then
            W_User_Count = .user_count
            user_str = ""
            For i = 0 To 100
                If W_User_List(i) <> "" Then
                    If user_str = "" Then
                        user_str = W_User_List(i)
                        
                    Else
                        user_str = user_str & ";" & W_User_List(i)
                    End If
                End If
            Next i
            '显示
                 TDBGrid1.TextMatrix(Row, 2) = user_str
                 
         End If
    End With

    If User_List <> TDBGrid1.TextMatrix(Row, 2) Then
    
        X = Split(TDBGrid1.TextMatrix(Row, 2), ";")
        
        For i = 0 To UBound(X)
             Tmp_User = Tmp_User & "'" & X(i) & "',"
'            G_Con.Execute "INSERT INTO mmst907(INV_LIST,USER_LIST,type) SELECT " & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & ",LIST_NO,0 FROM MMST901 WHERE USER_NAME='" & X(i) & "' AND LIST_NO NOT IN (SELECT USER_LIST FROM mmst907 WHERE INV_LIST =" & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & " and type=0) "
        Next
        
        If Tmp_User <> "" Then
            Tmp_User = "(" & Left(Tmp_User, Len(Tmp_User) - 1) & ")"
        
            '插入
            G_Con.Execute "INSERT INTO mmst907(INV_LIST,USER_LIST,type) SELECT " & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & ",LIST_NO,0 FROM MMST901 WHERE USER_NAME in " & Tmp_User & " AND LIST_NO NOT IN (SELECT USER_LIST FROM mmst907 WHERE INV_LIST =" & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & " and type=0) "
            '删除
            G_Con.Execute " Delete from  mmst907 WHERE user_list not in (select list_no from mmst901 where user_name in " & Tmp_User & " AND  INV_LIST =" & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & " and type=0) and INV_LIST =" & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & " and type=0"
        Else
            G_Con.Execute " Delete from  mmst907 WHERE inv_list=" & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & " and type=0"
            
        End If
     
        
    End If
Else
    With FrmUserList
         .W_User_name = TDBGrid1.TextMatrix(Row, 3)
          User_List = TDBGrid1.TextMatrix(Row, 3)
         .Show vbModal
         If .cancel_status = False Then
            W_User_Count = .user_count
            user_str = ""
            For i = 0 To 100
                If W_User_List(i) <> "" Then
                    If user_str = "" Then
                        user_str = W_User_List(i)
                        
                    Else
                        user_str = user_str & ";" & W_User_List(i)
                    End If
                End If
            Next i
            '显示
                 TDBGrid1.TextMatrix(Row, 3) = user_str
                 
         End If
    End With

    If User_List <> TDBGrid1.TextMatrix(Row, 3) Then
    
        X = Split(TDBGrid1.TextMatrix(Row, 3), ";")
        
        For i = 0 To UBound(X)
             Tmp_User = Tmp_User & "'" & X(i) & "',"
'            G_Con.Execute "INSERT INTO mmst907(INV_LIST,USER_LIST,type) SELECT " & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & ",LIST_NO,0 FROM MMST901 WHERE USER_NAME='" & X(i) & "' AND LIST_NO NOT IN (SELECT USER_LIST FROM mmst907 WHERE INV_LIST =" & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & " and type=0) "
        Next
        
        If Tmp_User <> "" Then
            Tmp_User = "(" & Left(Tmp_User, Len(Tmp_User) - 1) & ")"
        
            '插入
            G_Con.Execute "INSERT INTO mmst907(INV_LIST,USER_LIST,type) SELECT " & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & ",LIST_NO,1 FROM MMST901 WHERE USER_NAME in " & Tmp_User & " AND LIST_NO NOT IN (SELECT USER_LIST FROM mmst907 WHERE INV_LIST =" & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & " and type=1) "
            '删除
            G_Con.Execute " Delete from  mmst907 WHERE user_list not in (select list_no from mmst901 where user_name in " & Tmp_User & " AND  INV_LIST =" & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & " and type=1) and INV_LIST =" & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & " and type=1 "
        Else
            G_Con.Execute " Delete from  mmst907 WHERE inv_list=" & TDBGrid1.TextMatrix(Row, TDBGrid1.Cols - 1) & " and type=1"
            
        End If
     
        
    End If


End If
    
End Sub
