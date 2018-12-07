VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form Frm80XMx1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "新增/修改/删除规格资料"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   4305
      Left            =   -30
      ScaleHeight     =   4245
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   -30
      Width           =   5655
      Begin VB.CommandButton Cmd_NView 
         BackColor       =   &H00C0C0FF&
         Cancel          =   -1  'True
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4350
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "&Cancel"
         Top             =   3810
         Width           =   1215
      End
      Begin VB.CommandButton CmdView 
         BackColor       =   &H00C0C0FF&
         Caption         =   "确定"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2670
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "&OK"
         Top             =   3810
         Width           =   1215
      End
      Begin VSFlex7Ctl.VSFlexGrid TDBGrid1 
         Height          =   3675
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5595
         _cx             =   9869
         _cy             =   6482
         _ConvInfo       =   -1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
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
         FixedCols       =   1
         RowHeightMin    =   280
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
         AutoSearch      =   0
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
Attribute VB_Name = "Frm80XMx1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim W_Data1 As String
Dim W_Row As Double
Dim W_Col As Double
Dim W_Note_No As String
Dim W_Upd_Mode As Integer           '1.更改状态    2.查看状态
'相关变量的定义
Public W_Old_Mtr As String
Public W_Smtr_No As String
Public W_Sbom_No As String

Public Property Let Note_No(b As String)
    W_Note_No = b
End Property

Public Property Let Upd_Mode(b As Integer)
    W_Upd_Mode = b
End Property

Private Sub Cmd_NView_Click()
    G_Con.Execute " Delete From Mmii692 Where Pc_Name='" & G_Pc_Name & "' "
    Unload Me
End Sub

Private Sub CmdView_Click()
    
    '更新主表里面的临时表资料
    '******************************************************************
    '清除之前的历史数据
    G_Con.Execute " Delete From Mmst692 Where Old_Mtr_No='" & W_Old_Mtr & "' And Smtr_No='" & W_Smtr_No & "' And Sbom_no='" & W_Sbom_No & "' "
    '加入处理後的数据
    G_Con.Execute " Insert Into Mmst692(Old_Mtr_No,Smtr_No,Sbom_No,SbDim_No,Sbom_Dim,Defa_Status) " & _
                  " Select '" & W_Old_Mtr & "' As Old_Mtr_No,'" & W_Smtr_No & "' As Smtr_No,'" & W_Sbom_No & "' As sbom_no, " & _
                         " SbDim_No,Sbom_Dim,Defa_Status " & _
                  " From Mmii692 " & _
                  " Where Pc_Name='" & G_Pc_Name & "'  " & _
                  " Order By SbDim_No "
    '清除临时表里面的数据
    G_Con.Execute " Delete From Mmii692 Where Pc_Name='" & G_Pc_Name & "'"
    '更新主档资料
    G_Con.Execute " Update Mmst691 Set Sbom_Dim='' Where Old_Mtr_No='" & W_Old_Mtr & "' And Smtr_no='" & W_Smtr_No & "' And Sbom_No='" & W_Sbom_No & "'"
    G_Con.Execute " Update Mmst691 Set Sbom_Dim=b.Sbom_Dim " & _
                  " From MMst691 a Inner Join Mmst692 b On a.Old_Mtr_No=b.Old_Mtr_No And a.Smtr_no=b.Smtr_No " & _
                                                     " And a.Sbom_No=b.Sbom_No " & _
                  " Where a.Old_Mtr_No='" & W_Old_Mtr & "' And a.Smtr_no='" & W_Smtr_No & "' " & _
                    " And a.Sbom_No='" & W_Sbom_No & "' and  b.Defa_status=-1 "
    
    '关闭窗口
    Unload Me
End Sub

Private Sub Form_Load()
    Call CenterWindow(Me, Erp_File)
    If W_Upd_Mode = 1 Then
        Call Refresh_Grid1
    Else
        Call Refresh_Grid
    End If
    TDBGrid1.TabStop = True
End Sub
Private Sub Refresh_Grid()
Dim w_tmp As New ADODB.Recordset

w_tmp.Open " Select SbDim_No, Sbom_Dim,Defa_Status,List_No " & _
            " From Mmst692   " & _
            " Where Old_Mtr_No='" & W_Old_Mtr & "' And Smtr_No='" & W_Smtr_No & "' " & _
                  " And Sbom_No='" & W_Sbom_No & "' " & _
            " Order By List_no ", G_Con, adOpenDynamic, adLockOptimistic
            
Set TDBGrid1.DataSource = w_tmp
Call readactive
w_tmp.Close
TDBGrid1.Editable = flexEDNone
End Sub
Private Sub Refresh_Grid1()
Dim w_tmp1 As New ADODB.Recordset
                       
w_tmp1.Open " Select SbDim_No,Sbom_Dim,Defa_Status,List_No " & _
            " From Mmii692 " & _
            " Where Pc_Name='" & G_Pc_Name & "' " & _
            " Order By SbDim_No ", G_Con, adOpenDynamic, adLockOptimistic
           
Set TDBGrid1.DataSource = w_tmp1
Call ReadActive1
w_tmp1.Close
End Sub

Private Sub readactive()
With TDBGrid1
    .TextMatrix(0, 0) = "No."
    .TextMatrix(0, 1) = "规格编号"
    .TextMatrix(0, 2) = "规格描述"
    .TextMatrix(0, 3) = "预设标识"
    
    .ColWidth(0) = 550
    .ColWidth(4) = 0
    .Rows = .Rows
    '刷新全部 ROW 的高度 包括 HEADER
    For i = 1 To .Rows
        .RowHeight(i - 1) = 350
        If i < .Rows Then
            .TextMatrix(i, 0) = i
        End If
    Next i
    .ColAlignment(0) = flexAlignCenterCenter
    .ColDataType(3) = flexDTBoolean
End With

If TDBGrid1.Rows > 1 Then
    Call TDBGrid1_AfterRowColChange(0, 0, 1, 1)
End If
End Sub
Private Sub ReadActive1()
With TDBGrid1
    .TextMatrix(0, 0) = "No."
    .TextMatrix(0, 1) = "规格编号"
    .TextMatrix(0, 2) = "规格描述"
    .TextMatrix(0, 3) = "预设标识"
    
    .ColWidth(0) = 550
    .ColWidth(4) = 0
    .Rows = .Rows + 1
    '刷新全部 ROW 的高度 包括 HEADER
    For i = 1 To .Rows
        .RowHeight(i - 1) = 350
        If i < .Rows Then
            .TextMatrix(i, 0) = i
        End If
    Next i
    .ColAlignment(0) = flexAlignCenterCenter
    .ColDataType(3) = flexDTBoolean
End With

If TDBGrid1.Rows > 1 Then
    Call TDBGrid1_AfterRowColChange(0, 0, 1, 1)
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    W_Old_Mtr = ""
    W_Smtr_No = ""
    W_Sbom_No = ""
    Set Frm80XMx1 = Nothing
End Sub

Private Sub TDBGrid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
W_Data1 = TDBGrid1.TextMatrix(OldRow, OldCol)
W_Row = NewRow
W_Col = NewCol
If OldRow <> NewRow Then
    If NewRow >= 0 Then
        TDBGrid1.TextMatrix(OldRow, 0) = OldRow
        TDBGrid1.TextMatrix(NewRow, 0) = "★"
        TDBGrid1.Row = NewRow
        TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
    End If
End If
TDBGrid1.TextMatrix(0, 0) = " No"
TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
End Sub

Private Sub TDBGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim c_row1 As Double
Dim c_col1 As Double
Dim C_Row2 As Double
Dim C_Col2 As Double
'表格定位变量
With TDBGrid1
    c_row1 = .Row
    c_col1 = .Col
    C_Row2 = .TopRow
    C_Col2 = .LeftCol
End With


    '处理数据更新
    Dim w_tmp As New ADODB.Recordset
    If Col = 1 Then

        If Trim(TDBGrid1.TextMatrix(Row, Col)) <> W_Data1 Then
            If Trim(TDBGrid1.TextMatrix(Row, Col)) = "" Then
                MsgBox "规格编号不能为空!", 64, "提示"
                TDBGrid1.TextMatrix(Row, Col) = W_Data1
            Else
                    If Len(Trim(TDBGrid1.TextMatrix(Row, Col))) > 10 Then
                        Label2.Caption = "规格编号不能大於10位!"
                        TDBGrid1.TextMatrix(Row, Col) = W_Data1
                    Else
                        w_tmp.Open " Select SbDim_No  " & _
                                   " From Mmii692 " & _
                                   " Where Pc_Name='" & G_Pc_Name & "' And Sbdim_No='" & TDBGrid1.TextMatrix(Row, Col) & "'", G_Con
                        If w_tmp.EOF = False Then
                            MsgBox "同一物件下已有该编号的规格描述.", 64, "提示"
                            TDBGrid1.TextMatrix(Row, Col) = W_Data1
                        Else
                            Set w_tmp = Nothing
                            '*************************************************************
                            w_tmp.Open "Select Sbdim_No From Mmii692 Where List_no=" & Val(TDBGrid1.TextMatrix(Row, 4)), G_Con
                            If w_tmp.EOF = False Then
                                G_Con.Execute " Update Mmii692 Set Sbdim_No='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' " & _
                                              " Where list_no=" & Val(TDBGrid1.TextMatrix(Row, 4))
                            Else
                                G_Con.Execute " Insert Into Mmii692(Pc_Name,Sbdim_No,Sbom_Dim,Defa_Status) " & _
                                                           " Values( '" & G_Pc_Name & "' , " & _
                                                                   " '" & Trim(TDBGrid1.TextMatrix(Row, 1)) & "' ," & _
                                                                   " '" & Trim(TDBGrid1.TextMatrix(Row, 2)) & "' , " & _
                                                                   " '" & Val(TDBGrid1.TextMatrix(Row, 3)) & "' ) "
                            End If
                            G_Con.Execute " Delete From  Mmii692 Where Pc_Name='" & G_Pc_Name & "' And isnull(Sbdim_No,'')='' "
                            '表格定位处理
                            With TDBGrid1
                                Call Refresh_Grid1
                                .Row = c_row1
                                .Col = c_col1
                                .TopRow = C_Row2
                                .LeftCol = C_Col2
                            End With
                        End If
                        w_tmp.Close
                    End If
                    
              End If
              
        End If
    End If
    
    If Col = 2 Then

        If Trim(TDBGrid1.TextMatrix(Row, Col)) <> W_Data1 Then
            If Len(Trim(TDBGrid1.TextMatrix(Row, Col))) > 50 Then
                MsgBox "规格描述不能大於50位!", 64, "提示"
                TDBGrid1.TextMatrix(Row, Col) = W_Data1
            Else
                    Set w_tmp = Nothing
                    '*************************************************************
                    w_tmp.Open "Select Sbom_Dim From Mmii692 Where List_no=" & Val(TDBGrid1.TextMatrix(Row, 4)), G_Con
                    If w_tmp.EOF = False Then
                        G_Con.Execute " Update Mmii692 Set Sbom_Dim='" & Trim(TDBGrid1.TextMatrix(Row, Col)) & "' " & _
                                      " Where list_no=" & Val(TDBGrid1.TextMatrix(Row, 4))
                    Else
                        G_Con.Execute " Insert Into Mmii692(Pc_Name,Sbdim_No,Sbom_Dim,Defa_Status) " & _
                                                   " Values( '" & G_Pc_Name & "' , " & _
                                                           " '" & Trim(TDBGrid1.TextMatrix(Row, 1)) & "' ," & _
                                                           " '" & Trim(TDBGrid1.TextMatrix(Row, 2)) & "' , " & _
                                                           " '" & Val(TDBGrid1.TextMatrix(Row, 3)) & "' ) "
                    End If
                    G_Con.Execute " Delete From  Mmii692 Where Pc_Name='" & G_Pc_Name & "' And isnull(Sbdim_No,'')='' "
                    '表格定位处理
                        With TDBGrid1
                            Call Refresh_Grid1
                            .Row = c_row1
                            .Col = c_col1
                            .TopRow = C_Row2
                            .LeftCol = C_Col2
                        End With
                End If
                Set w_tmp = Nothing
        End If
    End If

    If Col = 3 Then


            If Val(TDBGrid1.TextMatrix(Row, Col)) = -1 Then
                '实现单选效果
                G_Con.Execute " Update Mmii692 Set Defa_Status=0 Where Pc_Name='" & G_Pc_Name & "' "
                G_Con.Execute " Update Mmii692 Set Defa_Status=-1 Where List_No=" & Val(TDBGrid1.TextMatrix(Row, 4)) & ""
            End If
                Set w_tmp = Nothing
                '*************************************************************
                w_tmp.Open "Select Defa_Status From Mmii692 Where List_no=" & Val(TDBGrid1.TextMatrix(Row, 4)), G_Con
                If w_tmp.EOF = False Then
                    G_Con.Execute " Update Mmii692 Set Defa_Status='" & Val(TDBGrid1.TextMatrix(Row, Col)) & "' " & _
                                  " Where list_no=" & Val(TDBGrid1.TextMatrix(Row, 4))
                Else
                    G_Con.Execute " Insert Into Mmii692(Pc_Name,Sbdim_No,Sbom_Dim,Defa_Status) " & _
                                               " Values( '" & G_Pc_Name & "' , " & _
                                                       " '" & Trim(TDBGrid1.TextMatrix(Row, 1)) & "' ," & _
                                                       " '" & Trim(TDBGrid1.TextMatrix(Row, 2)) & "' , " & _
                                                       " '" & Val(TDBGrid1.TextMatrix(Row, 3)) & "' ) "
                End If
                G_Con.Execute " Delete From  Mmii692 Where Pc_Name='" & G_Pc_Name & "' And isnull(Sbdim_No,'')='' "
                '表格定位处理
                With TDBGrid1
                    Call Refresh_Grid1
                    .Row = c_row1
                    .Col = c_col1
                    .TopRow = C_Row2
                    .LeftCol = C_Col2
                End With
    
                Set w_tmp = Nothing
            

    End If
'W_Grid3_Row = Row
'W_Grid3_Col = Col
End Sub

Private Sub TDBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If W_Upd_Mode = 1 And KeyCode = vbKeyDelete Then
            If MsgBox("你要删除此资料吗", vbYesNo, "提示") = vbYes Then
                G_Con.Execute " Delete  From Mmii692 Where Pc_Name='" & G_Pc_Name & "' And list_no=" & Val(TDBGrid1.TextMatrix(TDBGrid1.Row, 4))
                Call Refresh_Grid1
            End If
End If
End Sub
