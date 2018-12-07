VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F4732CE3-9A6C-11D2-8018-0080AD70A386}#5.7#0"; "ARESBUTTONPRO.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form mmss606 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "发票单制作(606)"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11850
   Tag             =   "Quotations"
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3690
      Top             =   7080
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   582
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
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   690
      Top             =   7110
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   582
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
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   0
      ScaleHeight     =   4815
      ScaleWidth      =   1965
      TabIndex        =   28
      Top             =   0
      Width           =   1995
      Begin VB.PictureBox lab_focus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   240
         ScaleHeight     =   360
         ScaleWidth      =   90
         TabIndex        =   49
         Top             =   330
         Visible         =   0   'False
         Width           =   90
      End
      Begin ARESBUTTONLib.AresButton cmd_find 
         Height          =   360
         Left            =   420
         TabIndex        =   15
         Top             =   360
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Find_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Find_Norm.bmp"
         ToolTipString   =   "查询符合条件的表单"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_print 
         Height          =   360
         Left            =   420
         TabIndex        =   16
         Top             =   773
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Print_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Print_Norm.bmp"
         ToolTipString   =   "列印报表"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_preview 
         Height          =   360
         Left            =   420
         TabIndex        =   17
         Top             =   1186
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Preview_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Preview_Norm.bmp"
         ToolTipString   =   "预览报表"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_save 
         Height          =   360
         Left            =   420
         TabIndex        =   18
         Top             =   1599
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Save_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Save_Norm.bmp"
         ToolTipString   =   "将表单存成文件"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_ok 
         Height          =   360
         Left            =   420
         TabIndex        =   19
         Top             =   2012
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Ok_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Ok_Norm.bmp"
         ToolTipString   =   "确认存盘"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_cancel 
         Height          =   360
         Left            =   420
         TabIndex        =   20
         Top             =   2425
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Cancel_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Cancel_Norm.bmp"
         ToolTipString   =   "放弃存盘"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_add 
         Height          =   360
         Left            =   420
         TabIndex        =   21
         Top             =   2838
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Add_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Add_Norm.bmp"
         ToolTipString   =   "增加一张单据"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   127107100
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_edit 
         Height          =   360
         Left            =   420
         TabIndex        =   22
         Top             =   3251
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Edit_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Edit_Norm.bmp"
         ToolTipString   =   "修改该单据"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_delete 
         Height          =   360
         Left            =   420
         TabIndex        =   23
         Top             =   3664
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Delete_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Delete_Norm.bmp"
         ToolTipString   =   "删除该单据"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmd_quit 
         Height          =   360
         Left            =   420
         TabIndex        =   24
         Top             =   4080
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Quit_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Quit_Norm.bmp"
         ToolTipString   =   "退出该程式"
         ToolTipShowTime =   0
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   67145324
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000014&
      ForeColor       =   &H00FF0000&
      Height          =   1965
      Left            =   -30
      TabIndex        =   30
      Top             =   4770
      Width           =   2025
      Begin ARESBUTTONLib.AresButton AresButton1 
         Height          =   1050
         Left            =   540
         TabIndex        =   31
         Top             =   510
         Width           =   1050
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\File.gif"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\File.gif"
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PrevPointer     =   217092436
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin VB.Label Help_txt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   180
         Width           =   1755
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6855
      Left            =   1980
      TabIndex        =   29
      Top             =   -90
      Width           =   90
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6795
      Left            =   2070
      ScaleHeight     =   6765
      ScaleWidth      =   9795
      TabIndex        =   1
      Top             =   -30
      Width           =   9825
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2265
         Left            =   0
         ScaleHeight     =   2235
         ScaleWidth      =   3465
         TabIndex        =   42
         Top             =   30
         Width           =   3495
         Begin VB.ComboBox inv_no 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1260
            Sorted          =   -1  'True
            TabIndex        =   0
            Top             =   750
            Width           =   1815
         End
         Begin VB.ComboBox pay_type 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1260
            Sorted          =   -1  'True
            TabIndex        =   3
            Top             =   1410
            Width           =   1815
         End
         Begin MSComCtl2.DTPicker inv_date 
            Height          =   315
            Left            =   1260
            TabIndex        =   2
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   87949312
            CurrentDate     =   37240
         End
         Begin MSComCtl2.DTPicker close_date 
            Height          =   315
            Left            =   1260
            TabIndex        =   4
            Top             =   1800
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   87949312
            CurrentDate     =   37249
         End
         Begin VB.Line Line8 
            X1              =   570
            X2              =   2670
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Line Line7 
            Index           =   0
            X1              =   570
            X2              =   2670
            Y1              =   390
            Y2              =   390
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   810
            TabIndex        =   47
            Top             =   1140
            Width           =   420
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice No:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   300
            TabIndex        =   46
            Top             =   810
            Width           =   930
         End
         Begin VB.Label lbCaption 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   960
            TabIndex        =   45
            Tag             =   "Quotations"
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   495
            TabIndex        =   44
            Top             =   1470
            Width           =   735
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Close Day:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   360
            TabIndex        =   43
            Top             =   1830
            Width           =   870
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2265
         Left            =   3480
         ScaleHeight     =   2235
         ScaleWidth      =   6255
         TabIndex        =   33
         Top             =   30
         Width           =   6285
         Begin VB.PictureBox cmd_brow 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   2130
            Picture         =   "mmss606.frx":0000
            ScaleHeight     =   330
            ScaleWidth      =   330
            TabIndex        =   6
            Top             =   60
            Width           =   330
         End
         Begin VB.TextBox cust_eaddr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   960
            TabIndex        =   8
            Top             =   420
            Width           =   5145
         End
         Begin VB.TextBox go_port 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4050
            TabIndex        =   14
            Top             =   1860
            Width           =   2055
         End
         Begin VB.TextBox lea_port 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1440
            TabIndex        =   13
            Top             =   1860
            Width           =   1935
         End
         Begin VB.TextBox boat_name 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   960
            TabIndex        =   11
            Top             =   1500
            Width           =   1815
         End
         Begin VB.TextBox boat_company 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   960
            TabIndex        =   10
            Top             =   1140
            Width           =   5145
         End
         Begin VB.TextBox cust_ename 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2490
            TabIndex        =   7
            Top             =   60
            Width           =   3615
         End
         Begin VB.TextBox lc_no 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   960
            TabIndex        =   9
            Top             =   780
            Width           =   5145
         End
         Begin MSComCtl2.DTPicker boat_date 
            Height          =   315
            Left            =   4080
            TabIndex        =   12
            Top             =   1500
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   87949312
            CurrentDate     =   37249
         End
         Begin VB.TextBox cust_no 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   960
            TabIndex        =   5
            Top             =   60
            Width           =   1485
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3630
            TabIndex        =   41
            Top             =   1950
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Addr:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   480
            TabIndex        =   40
            Top             =   480
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   660
            TabIndex        =   39
            Top             =   150
            Width           =   270
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shipment: From:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   38
            Top             =   1950
            Width           =   1290
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "L/C No:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   315
            TabIndex        =   37
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vassel:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   375
            TabIndex        =   36
            Top             =   1200
            Width           =   570
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ship Date:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3120
            TabIndex        =   35
            Top             =   1560
            Width           =   825
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sailing On:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   34
            Top             =   1560
            Width           =   870
         End
      End
      Begin VB.TextBox Remark 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   6420
         Width           =   8895
      End
      Begin VSFlex7Ctl.VSFlexGrid TDBGrid1 
         Height          =   4005
         Left            =   0
         TabIndex        =   25
         Top             =   2370
         Width           =   3465
         _cx             =   6112
         _cy             =   7064
         _ConvInfo       =   -1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
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
         BackColorSel    =   -2147483639
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483641
         GridColorFixed  =   -2147483641
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
         SelectionMode   =   0
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
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
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
         AllowUserFreezing=   3
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid TDBGrid2 
         Height          =   4005
         Left            =   3480
         TabIndex        =   48
         Top             =   2370
         Width           =   6285
         _cx             =   11086
         _cy             =   7064
         _ConvInfo       =   -1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
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
         BackColorSel    =   -2147483639
         ForeColorSel    =   -2147483641
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483641
         GridColorFixed  =   -2147483641
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
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
         MergeCompare    =   0
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
         AllowUserFreezing=   3
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   -30
         X2              =   9750
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "Remark:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   6450
         Width           =   615
      End
   End
End
Attribute VB_Name = "mmss606"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'*程序名称:成品出货维护(mmss606)
'*编写日期:
'*制作人员:
'*修改日期:
'*修改人员:
'***********************************************
'定义记录集与命令对象
Dim st_511 As ADODB.Recordset
Dim st_802 As New ADODB.Recordset
Dim com_511 As ADODB.Command
Dim Com_501 As ADODB.Command
Dim com_502 As ADODB.Command

'指示当前的inv_no
Dim W_Curr_MtrsNo As String

'存放TDBGRID1 的旧字符
Dim w_old_str1 As String
Dim w_old_str2 As String

'日期
Dim w_close_date As Date
Dim w_boat_date As Date
Dim W_inv_date As Date

'客户代号
Dim W_cust_No As String

'付款方式
Dim w_pay_type As String

'TDBGrid相关
Dim gridc_501(127) As Grid_Data '存放 Grid 属性值
Dim gridc_502(127) As Grid_Data '存放 Grid 属性值

Dim Row_Height1 As Double        'TDBGrid1 高度变量
Dim Row_Height2 As Double        'TDBGrid2 高度变量

'定义按钮变量
Dim C_Add As Boolean
Dim C_Edit As Boolean
Dim C_Delete As Boolean
Dim c_print As Boolean
Dim c_view As Boolean
Dim c_save As Boolean

'权限变量
Dim C_Off_Add As Boolean
Dim C_Off_Edit As Boolean
Dim C_Off_Delete As Boolean
Dim c_off_print As Boolean
Dim c_off_view As Boolean
Dim c_off_save As Boolean



'按存档
Private Sub Cmd_save_MouseClick()
Call vcontrol("S")
End Sub

Public Sub Form_Activate()
'当窗口激活时,刷新TDBGrid
Call GetVSGridSetting("mmss606", "TDBGrid1", gridc_501, g_CON_IniFile4)
Call GetVSGridSetting("mmss606", "TDBGrid2", gridc_502, g_CON_IniFile4)

Row_Height1 = gridc_501(0).Grid_RowHeight
Row_Height2 = gridc_502(0).Grid_RowHeight
Call RefreshGrid
End Sub

Private Sub Form_Load()
'装载按钮图片
Call load_picture

'表单接收键值优先
Me.KeyPreview = True

'将MDI子窗口置中
Call CenterWindow(Me, Erp_Deliv)

'打开记录集,因为st_031仅用来新增记录,故无必要将记录全部传下
Set st_511 = New ADODB.Recordset
With st_511
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .ActiveConnection = G_Con
    .Open "SELECT * FROM mmst511 WHERE inv_no=''", , , , adCmdText
End With

'com_511产生的单条记录集显示单据的表头内容,它会被反复执行.
Set com_511 = New ADODB.Command
With com_511
    .CommandType = adCmdText
    .CommandText = "SELECT mmst511.*,mmst021.cust_ename,mmst021.cust_caddr,mmst802.pay_scrpt " & _
                   "FROM mmst511,mmst021,mmst802 " & _
                        " WHERE mmst021.cust_no=mmst511.cust_no " & _
                            "AND mmst511.pay_type = mmst802.pay_type " & _
                            "AND mmst511.inv_no=?"
    .ActiveConnection = G_Con
    .Prepared = True '因为它会多次执行,将它预编绎.
End With

'com_501产生的记录集显示单据的明细,它会被反复执行.
Set Com_501 = New ADODB.Command
With Com_501
    .Name = "GetMtrsNo606"
    .CommandType = adCmdText
    .CommandText = "SELECT  deliv_no,deliv_date,Remark,upd_name,upd_date " & _
                   "FROM mmst501  WHERE inv_no=? and status='2'  ORDER BY  deliv_no"
    
    On Error Resume Next
    .ActiveConnection = G_Con
    .Prepared = True '因为它会多次执行,将它预编绎.
End With

'com_502产生的发票的明细,它会被反复执行.
Set com_502 = New ADODB.Command
With com_502
    .Name = "GetdelivDetail502"
    .CommandType = adCmdText
    .CommandText = "SELECT  a.order_no,b.cust_order,a.mtr_no,c.mtr_ename," & _
                           "c.mtr_dim,c.unit_name,a.mtr_amt,d.mtr_prs,d.money_no," & _
                           "d.mtr_prs * a.mtr_amt as mtr_total " & _
                    "FROM mmst502 a INNER JOIN " & _
                          "mmst011 b ON a.order_no = b.order_no INNER ON " & _
                          "mmsp611 c ON a.mtr_no = c.mtr_no INNER ON " & _
                          "mmst012 d ON a.order_no = d.order_no AND a.mtr_no = d.mtr_no " & _
                    "WHERE inv_no=?   ORDER BY  deliv_no"
    
    On Error Resume Next
    .ActiveConnection = G_Con
    .Prepared = True '因为它会多次执行,将它预编绎.
End With


st_802.Open "SELECT pay_scrpt FROM mmst802 order by pay_scrpt", G_Con, adOpenDynamic
Do While st_802.EOF = False
    pay_type.AddItem st_802!Pay_Scrpt
    st_802.MoveNext
Loop
st_802.Close



'将按钮变量赋初值
C_Add = False
C_Edit = False
C_Delete = False

C_Off_Add = False
C_Off_Edit = False
C_Off_Delete = False
c_off_print = False
c_off_save = False
c_off_view = False

'MDI子窗口按钮权限设订
C_Off_Add = lopcheck("A", "606", G_User_ID)
C_Off_Edit = lopcheck("U", "606", G_User_ID)
C_Off_Delete = lopcheck("D", "606", G_User_ID)
c_off_view = lopcheck("V", "606", G_User_ID)
c_off_print = lopcheck("P", "606", G_User_ID)
c_off_save = lopcheck("S", "606", G_User_ID)

'调用inv_no_Click
Call inv_no_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 34
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Form_KeyDown可照搬
If KeyCode = vbKeyReturn Then
    If ActiveControl.Name = "Remark" Then
        If ActiveControl.MultiLine = False Then
            SendKeys "{TAB}"
        End If
    Else
        SendKeys "{TAB}"
    End If
    Exit Sub
End If

If KeyCode = vbKeyM And Shift = 1 Then

    '如果用户手工改动了单据号
     If LCase(Me.ActiveControl.Name) = "inv_no" Then
         Call inv_no_LostFocus
     End If
    
    '如果没有单据
    If Trim(inv_no.Text) = "" Then
        Exit Sub
    End If
    
    '检查单据状态
    If check_ok = False Then
        Exit Sub
    End If
    '非新增修改状态按下 SHIFT + M 弹出明细菜单
    If Not (C_Add Or C_Edit Or c_delet) Then
        Call TDBGrid1_MouseUp(2, 0, TDBGrid1.Left + 200, TDBGrid1.Top + 50)
    Else
         Exit Sub
    End If
   
End If

If Shift = 0 Then
    Select Case KeyCode
    
    Case vbKeyF2               '新增
         If cmd_add.Enabled = True Then
             Call vcontrol("A")
             KeyCode = 0
         End If
    Case vbKeyF3               '编辑
        '如果用户手工改动了单据号
        If LCase(Me.ActiveControl.Name) = "inv_no" Then
            Call inv_no_LostFocus
        End If
     
        If cmd_edit.Enabled = True Then
             Call vcontrol("U")
             KeyCode = 0
        End If
        
        
    Case vbKeyF4               '删除
        
        '如果用户手工改动了单据号
        If LCase(Me.ActiveControl.Name) = "inv_no" Then
            Call inv_no_LostFocus
        End If
         
        If cmd_delete.Enabled = True Then
             Call vcontrol("D")
            KeyCode = 0
        End If
    Case vbKeyF5               '确认
         If cmd_ok.Enabled = True Then
             Call vcontrol("Y")
             KeyCode = 0
         End If
    Case vbKeyF6               '退出
         If Cmd_Quit.Enabled = True Then
             Call vcontrol("Q")
             KeyCode = 0
         End If
         
    Case vbKeyF7               '查询
         If cmd_find.Enabled = True Then
             Call vcontrol("F")
             KeyCode = 0
         End If
    Case vbKeyF8               '列印
        '如果用户手工改动了单据号
        If LCase(Me.ActiveControl.Name) = "inv_no" Then
            Call inv_no_LostFocus
        End If
         
        If cmd_print.Enabled = True Then
             Call vcontrol("P")
             KeyCode = 0
        End If
    Case vbKeyF9               '预览
        '如果用户手工改动了单据号
        If LCase(Me.ActiveControl.Name) = "inv_no" Then
            Call inv_no_LostFocus
        End If

        If cmd_preview.Enabled = True Then
             Call vcontrol("V")
             KeyCode = 0
        End If
    Case vbKeyF12              '存储
        '如果用户手工改动了单据号
        If LCase(Me.ActiveControl.Name) = "inv_no" Then
            Call inv_no_LostFocus
        End If
         
        If cmd_save.Enabled = True Then
             Call vcontrol("S")
             KeyCode = 0
        End If
    Case vbKeyEscape           '取消
         If cmd_cancel.Enabled = True Then
             Call vcontrol("N")
             KeyCode = 0
         End If
    End Select
End If
End Sub

Private Sub readshow()
'当新增时
If C_Add = True Then
    inv_no.Text = Creat_No
    Inv_date.Value = Date
    
    close_date.Value = Date
    
    cust_no.Text = ""
    cust_ename.Text = ""
    cust_eaddr.Text = ""
    lc_no.Text = ""
    boat_company.Text = ""
    boat_name.Text = ""
    boat_date.Value = Date
    lea_port.Text = ""
    go_port.Text = ""
    remark.Text = ""
    
    '刷新表格
    Call RefreshGrid
End If

'设定各按键的 Enabled 属性
If C_Add Or C_Edit Or C_Delete Then
    cmd_ok.Enabled = True
    cmd_cancel.Enabled = True
    
    cmd_add.Enabled = False
    cmd_edit.Enabled = False
    cmd_delete.Enabled = False
    cmd_print.Enabled = False
    cmd_save.Enabled = False
    cmd_preview.Enabled = False
    cmd_find.Enabled = False
Else
    cmd_ok.Enabled = False
    cmd_cancel.Enabled = False
    
    cmd_find.Enabled = True
    cmd_add.Enabled = True
    If inv_no.ListCount <= 0 Or inv_no.Text = "" Then
        cmd_edit.Enabled = False
        cmd_delete.Enabled = False
        cmd_print.Enabled = False
        cmd_save.Enabled = False
        cmd_preview.Enabled = False
     Else
        cmd_edit.Enabled = True
        cmd_delete.Enabled = True
        cmd_print.Enabled = True
        cmd_save.Enabled = True
        cmd_preview.Enabled = True
        '
    End If
End If

'通过权限设定按键的 Enabled 属性
If C_Off_Add = True Then
    cmd_add.Enabled = False
End If

If C_Off_Edit = True Then
    cmd_edit.Enabled = False
End If

If C_Off_Delete = True Then
    cmd_delete.Enabled = False
End If

If c_off_print = True Then
    cmd_print.Enabled = False
End If

If c_off_save = True Then
    cmd_save.Enabled = False
End If

If c_off_view = True Then
    cmd_preview.Enabled = False
End If

'设定各控件的 Locked 属性
Dim w_c As Control
For Each w_c In Me.Controls
    If TypeOf w_c Is ComboBox And UCase(w_c.Name) <> "inv_no" Then
        w_c.Locked = Not (C_Add Or C_Edit)
    
    ElseIf TypeOf w_c Is ComboBox And UCase(w_c.Name) = "inv_no" Then
        w_c.Locked = C_Edit
        
    ElseIf TypeOf w_c Is TextBox Then
        w_c.Locked = Not (C_Add Or C_Edit)
        
    ElseIf TypeOf w_c Is Frame And UCase(w_c.Name) = "FRAME3" Then
        w_c.Enabled = (C_Add Or C_Edit)
    
    ElseIf TypeOf w_c Is DTPicker Then
'        w_c.lock = (c_add Or c_edit)
        
    End If
Next w_c

If Not (C_Add Or C_Edit) Then
    inv_no.Locked = False
End If
cust_ename.Locked = True
cust_eaddr.Locked = True
If Not C_Add Then
   Cmd_brow.Enabled = False
   cust_no.Locked = True
Else
    Cmd_brow.Enabled = True
    cust_no.Locked = False
End If
End Sub

Private Sub readactive()
Dim w_rs501 As New ADODB.Recordset '从档

w_rs501.CursorLocation = adUseClient
'注意 从档的 SELECT 顺序一定和 INI 文件的顺序相同
W_Str = "SELECT deliv_no,deliv_date,Remark,upd_name,upd_date " & _
            "FROM mmst501 " & _
            " WHERE inv_no= '" & Trim(inv_no.Text) & "' AND inv_no <> ''" & _
            " ORDER BY  deliv_no"
                   
w_rs501.Open W_Str, G_Con, adOpenDynamic

Set Adodc1.Recordset = w_rs501

Set TDBGrid1.DataSource = Adodc1
Call SetVSGridSetting(TDBGrid1, gridc_501)

'刷新全部 ROW 的高度 包括 HEADER
For i = 1 To TDBGrid1.Rows
    TDBGrid1.Row = i - 1
    TDBGrid1.RowHeight(i - 1) = Row_Height1
    
    If i < TDBGrid1.Rows Then
        TDBGrid1.TextMatrix(i, 0) = i
    End If
Next i
TDBGrid1.ColAlignment(0) = flexAlignCenterCenter

Set TDBGrid2.DataSource = Adodc2
Call SetVSGridSetting(TDBGrid2, gridc_502)

'刷新全部 ROW 的高度 包括 HEADER
For i = 1 To TDBGrid2.Rows
    TDBGrid2.Row = i - 1
    TDBGrid2.RowHeight(i - 1) = Row_Height2

    If i < TDBGrid2.Rows Then
        TDBGrid2.TextMatrix(i, 0) = i
    End If
Next i
TDBGrid2.ColAlignment(0) = flexAlignCenterCenter
End Sub

'***********************************************************
'响应弹出菜单项单击的过程
Public Sub menu_add_Click()
If LockRecord("mmst511", "inv_no = '" & Trim(inv_no.Text) & "'") Then
    If check_ok = False Then
        Call UnLockRecord("mmst511", "inv_no = '" & Trim(inv_no.Text) & "'")
        Exit Sub
    End If

    Set FrmDelivno.CallForm = Me
    FrmDelivno.UpdateMode = 0 'UpdateMode=0表示新增
    FrmDelivno.Show vbModal
    
    '新增完毕解锁
    Call UnLockRecord("mmst511", "inv_no = '" & Trim(inv_no.Text) & "'")
    
    TDBGrid1.SetFocus
    TDBGrid1.Col = 1
    If TDBGrid1.Rows > 1 Then
        TDBGrid1.Row = 1
    End If
End If
End Sub

Public Sub menu_edit_Click()
'修改前加锁
If LockRecord("mmst511", "inv_no = '" & Trim(inv_no.Text) & "'") Then
    If check_ok() = False Then
        Call UnLockRecord("mmst511", "inv_no = '" & Trim(inv_no.Text) & "'")
        Exit Sub
    End If
    c_row = TDBGrid1.Row
    c_col = TDBGrid1.Col
    
With FrmDelivno
    .UpdateMode = 1 'UpdateMode=1表示修改
    
    Set .CallForm = Me
        .Deliv_No.Text = Adodc1.Recordset!Deliv_No
        .deliv_date.Text = Adodc1.Recordset!deliv_date
        .remark.Text = NullSetValue(Adodc1.Recordset!remark, "")
        .Show vbModal
    End With
        
    Call UnLockRecord("mmst511", "inv_no = '" & Trim(inv_no.Text) & "'")
    TDBGrid1.Row = c_row
    TDBGrid1.Col = c_col
End If

End Sub

Public Sub menu_delete_Click()
'删除前加锁
If LockRecord("mmst511", "inv_no = '" & Trim(inv_no.Text) & "'") Then
    '删除明细资料
    If MsgBox(g_CON_CDelete, vbYesNo + vbDefaultButton2 + vbInformation, g_CON_CTitle) = vbNo Then
        Call UnLockRecord("mmst511", "inv_no = '" & Trim(inv_no.Text) & "'")
        Exit Sub
    End If
    
    '判断当前单据是否已被审核
    If check_ok() = False Then
        Call UnLockRecord("mmst511", "inv_no = '" & Trim(inv_no.Text) & "'")
        Exit Sub
    End If
    
    '删除明细资料在出货单中将 INV_NO = ''
    
    G_Con.Execute "UPDATE mmst501 SET inv_no = '' WHERE deliv_no='" & Adodc1.Recordset!Deliv_No & "' "
    Call UnLockRecord("mmst511", "inv_no = '" & Trim(inv_no.Text) & "'")
    Call inv_no_Click
'    Call RefreshGrid
'
'    TDBGrid1.SetFocus
'    TDBGrid1.Col = 1
'    If TDBGrid1.Rows > 1 Then
'        TDBGrid1.Row = 1
'    End If

End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'当在新增或修改时,提示是否退出
If C_Add Or C_Edit Or C_Delete Then
    '当有数据改动时.询问是否要退出系统
    If MsgBox(g_CON_CQuit, vbQuestion + vbYesNo, g_CON_CTitle) = vbNo Then
        Cancel = 1
    Else
        '当有修改或删除时未解锁时,解除锁定
        If C_Edit Or C_Delete Then
            Call UnLockRecord("mmst511", "inv_no='" & inv_no.Text & "'")
        End If
        Cancel = 0
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

'退出时，存储 TDBGrid 属性
Call SaveGridSetting("mmss606", "TDBGrid1", gridc_501, g_CON_IniFile4)
Call SaveGridSetting("mmss606", "TDBGrid2", gridc_502, g_CON_IniFile4)

Set st_511 = Nothing
Set com_511 = Nothing
Set Com_501 = Nothing
Set com_502 = Nothing
Set TDBGrid1.DataSource = Nothing
Set TDBGrid2.DataSource = Nothing
Set mmss606 = Nothing
End Sub

'各按钮click事件
'************************************************************
'按确认
Sub cmd_ok_MouseClick()
Call vcontrol("Y")
End Sub

'按取消
Sub cmd_cancel_Mouseclick()
Call vcontrol("N")
End Sub

'按新增
Sub cmd_add_MouseClick()
Call vcontrol("A")
End Sub

'按修改
Sub cmd_edit_MouseClick()
Call vcontrol("U")
End Sub

'按删除
Sub cmd_delete_MouseClick()
Call vcontrol("D")
End Sub

'按预览
Private Sub Cmd_print_MouseClick()
Call vcontrol("P")
End Sub

'按打印
Private Sub Cmd_preview_MouseClick()
Call vcontrol("V")
End Sub



'按退出
Sub cmd_quit_MouseClick()
Call vcontrol("Q")
End Sub
'按查询
Private Sub cmd_find_MouseClick()
Call vcontrol("F")
End Sub

'VCONTROL 函数
Private Sub vcontrol(ByVal p_choice As String)
Dim W_add As Boolean

Select Case p_choice
    Case "Y"            '确定
        If check_ok() Then
            Call upd_data
            TDBGrid1.Enabled = True
        End If
    
    Case "N"            '取消
        '如果新增或修改时取消动作,则要解锁
        If C_Edit Or C_Delete Then
           Call UnLockRecord("mmst511", "inv_no='" & Trim(inv_no.Text) & "'")
        End If
        
        '当新增时取消动作
        If C_Add Then
            W_add = True
        End If
        C_Add = False
        C_Edit = False
        C_Delete = False
        TDBGrid1.Enabled = True
        
        inv_no.Text = W_Curr_MtrsNo
        
        '调用inv_no_Click
        Call inv_no_Click
        
        '新增或修改时自动刷新其ListIndex
        If W_add Then
            '用於自动执行一次下拉动作
            '显示DropDown
            Call SendMessage(inv_no.hWnd, CB_SHOWDROPDOWN, True, &O0)
            '隐藏DropDown
            Call SendMessage(inv_no.hWnd, CB_SHOWDROPDOWN, False, &O0)
            W_add = False
        End If
       
    Case "A"            ' 增加
        C_Add = True
        Call readshow
        
        TDBGrid1.Enabled = False
        inv_no.SetFocus
    
    Case "U"                    '修改
        If LockRecord("mmst511", "inv_no='" & Trim(inv_no.Text) & "'") Then
             '检查单据状态
            If check_ok() = False Then
                Call UnLockRecord("mmst511", "inv_no='" & Trim(inv_no.Text) & "'")
                Exit Sub
            End If
            C_Edit = True
            TDBGrid1.Enabled = False
            Call readshow
            
        End If
        Inv_date.SetFocus
        
    Case "D"                 '删除
        '加锁记录
        If LockRecord("mmst511", "inv_no='" & Trim(inv_no.Text) & "'") = True Then
            '删除当前记录
            If MsgBox(g_CON_CDelete, vbQuestion + vbYesNo, g_CON_CTitle) = vbNo Then
                Call UnLockRecord("mmst511", "inv_no='" & Trim(inv_no.Text) & "'")
                Exit Sub
            End If
            '判断是否可以删除
                        If check_ok = False Then
                Call UnLockRecord("mmst511", "inv_no='" & Trim(inv_no.Text) & "'")
                Exit Sub
            End If
            
            C_Delete = True
            
            '错误处理
            err.Clear
            On Error GoTo Del_Err
            '事务处理
            G_Con.BeginTrans
            '只是修改出货单中的INV_NO
            G_Con.Execute "UPDATE mmst501 SET inv_no = '' WHERE inv_no='" & Trim(inv_no.Text) & "'"
            G_Con.Execute "DELETE FROM mmst511 WHERE inv_no='" & Trim(inv_no.Text) & "'"
            G_Con.CommitTrans
            
            On Error GoTo 0
            Dim w_index As Integer
            
            '找到对应的 INDEX
            For w_index = 0 To inv_no.ListCount - 1
                 If inv_no.List(w_index) = Trim(inv_no.Text) Then
                    inv_no.RemoveItem (w_index)
                    Exit For
                 End If
            Next w_index
            
            C_Delete = False
            
            If inv_no.ListCount > 0 Then
                If inv_no.ListCount < w_index + 1 Then
                    w_index = w_index - 1
                End If
            Else
                w_index = -1
            End If
            
            If w_index <> -1 Then
                inv_no.ListIndex = w_index
            Else
                inv_no.Text = ""
                Call inv_no_Click
            End If
            Exit Sub
Del_Err:
            On Error Resume Next
            G_Con.RollbackTrans
            MsgBox "删除时出现错误!", 64, g_CON_CTitle
            Call UnLockRecord("mmst511", "inv_no='" & Trim(inv_no.Text) & "'")
       End If
       
    Case "P"    '打印
        c_print = True
        If check_ok = True Then
            Call sele_data
        End If
   
   Case "V"     '预览
        c_view = True
        If check_ok = True Then
            Call sele_data
        End If
    
    Case "S"
        c_save = True
        If check_ok = True Then
            Call sele_data
        End If
    
    Case "Q"    '退出
        Unload Me
    
    Case "F"   '查询
        With FrmpoInvSh
            .DefInvType = ""
            .DefTable = "mmst511"
            .DefField = "inv_no"
            .DefInvDate = "inv_date"
            .Label1(4).Caption = "发票单号"
            .Label1(2).Caption = "发票日期"
            .Caption = "发票制作查询"
        Set .CallControl = inv_no
            .cb_check.ListIndex = 0
            .cb_check.Enabled = False
            .Show vbModal
        If .ClickCancel = False Then
            If inv_no.ListCount > 0 Then
                inv_no.ListIndex = 0
            Else
                inv_no.Text = ""
                Call inv_no_Click
            End If
        End If
        End With
    End Select
End Sub

Private Function check_ok() As Boolean
Dim W_Rs As New ADODB.Recordset
Dim w_check As String

'当在网络情况更新数据时,先判断单据是否已审核(主档)或删除
'当状态为'2'时,只是不能异动单据,其它可以,如打印,存档等
w_check = CheckStatus("inv_no", Trim(inv_no.Text), "mmst511", "status")
If w_check = "9" Then
    If C_Edit Or (C_Add = False And C_Edit = False And C_Delete = False) Then
        MsgBox "当前单据已被其它用户删除,不能操作!", 64, g_CON_CTitle
        C_Edit = False
        check_ok = False
        Exit Function
    End If
End If

'当打印或预览或存档时判断是否有明细资料
If c_print Or c_save Or c_view Then
    If TDBGrid1.Rows < 0 Then
        MsgBox "此单据没有明细资料,请录入其明细!", vbInformation, g_CON_CTitle
        c_view = False
        c_print = False
        c_save = False
        check_ok = False
        Exit Function
    End If
    check_ok = True
    Exit Function
End If

'新增或修改时检查
If C_Add = True Then
    If Len(inv_no.Text) > 12 Then
        MsgBox "单据单号不能多於12个字符!", vbInformation, g_CON_CTitle
        check_ok = False
        inv_no.SetFocus
        Exit Function
    End If
    If Trim(inv_no.Text) = "" Then
        MsgBox "必须输入单据单号.", vbExclamation, g_CON_CTitle
        inv_no.SetFocus
        Exit Function
    Else
        W_Rs.CursorLocation = adUseClient
        W_Rs.Open "SELECT inv_no FROM mmst511 WHERE inv_no='" & inv_no.Text & "'", G_Con
        If W_Rs.EOF = False Then
            MsgBox "单据单号重复.", vbExclamation, g_CON_CTitle
            inv_no.SetFocus
            Exit Function
        End If
        W_Rs.Close
    End If
End If

'新增或修改主档时
If cust_no.Text = "" Then
    MsgBox "请选择客户代号.", vbExclamation, g_CON_CTitle
    cust_no.SetFocus
    check_ok = False
    Exit Function
Else
    W_Rs.Open "SELECT cust_no FROM mmst021 WHERE cust_no='" & cust_no.Text & "'", G_Con
    If W_Rs.EOF = True Then
        MsgBox "无此客户代号.", vbExclamation, g_CON_CTitle
        cust_no.SetFocus
        check_ok = False
        Exit Function
    Else
        W_cust_No = W_Rs(0)
    End If
    W_Rs.Close
End If

If pay_type.Text = "" Then
    MsgBox "请选择付款方式.", vbExclamation, g_CON_CTitle
    pay_type.SetFocus
    check_ok = False
    Exit Function
Else
    W_Rs.Open "SELECT pay_type FROM mmst802 WHERE pay_scrpt ='" & pay_type.Text & "'", G_Con, adOpenDynamic
    If W_Rs.EOF = True Then
        MsgBox "无此付款方式.", vbExclamation, g_CON_CTitle
        pay_type.SetFocus
        check_ok = False
        Exit Function
    Else
        w_pay_type = W_Rs!pay_type
    End If
    W_Rs.Close
End If

check_ok = True
End Function

Private Sub upd_data()
'upd_data将不再包含删除的过程
If C_Add = True Then
    With st_511
        .AddNew
        !inv_no = Trim(inv_no.Text)
        !Inv_date = Inv_date.Value
        !pay_type = w_pay_type
        !cust_no = W_cust_No
        !close_date = close_date.Value
        !lc_no = lc_no.Text
        !boat_company = boat_company.Text
        !boat_name = boat_name.Text
        !boat_date = boat_date.Value
        !lea_port = lea_port.Text
        !go_port = go_port.Text
        
        !remark = remark.Text
        !status = "0"
        !upd_name = Trim(G_User_Name)
        !upd_date = Format(Date, "yyyy-MM-dd")
        !lock = "No"
        .Update
    End With
    '刷新数据
    st_511.Requery
    '增加已经审核的 deliv_date <= Close_date 的deliv_no
'    G_Con.Execute "update mmst501 set inv_no = '" & inv_no.Text & "' WHERE cust_no = '" & W_cust_No & "' AND status = '2' AND deliv_date <= '" & close_date.Value & "'"
    C_Add = False
    '刷新成品ComboBox
    inv_no.AddItem inv_no.Text
    '用於自动执行一次下拉动作
    inv_no.ListIndex = inv_no.ListCount - 1
Else
    Dim st_511_1 As New ADODB.Recordset
    With st_511_1
        .CursorLocation = adUseClient
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .ActiveConnection = G_Con
        .Open "SELECT * FROM mmst511 WHERE inv_no='" & inv_no.Text & "'", , , , adCmdText
    End With
    
    With st_511_1
        
        !Inv_date = Inv_date.Value
        !pay_type = w_pay_type
        !cust_no = W_cust_No
        !close_date = close_date.Value
        !lc_no = lc_no.Text
        !boat_company = boat_company.Text
        !boat_name = boat_name.Text
        !boat_date = boat_date.Value
        !lea_port = lea_port.Text
        !go_port = go_port.Text
        
        !remark = remark.Text
        !status = "0"
        !upd_name = Trim(G_User_Name)
        !upd_date = Format(Date, "yyyy-MM-dd")
        !lock = "No"
        .Update
    End With
    C_Edit = False
    Call inv_no_Click
End If
End Sub



Private Sub inv_no_LostFocus()
If Not (C_Add Or C_Edit) Then
    If W_Curr_InvNo <> Trim(inv_no.Text) Then
        Call inv_no_Click
        If inv_no.Text <> "" Then
            Dim i As Long
            For i = 0 To inv_no.ListCount - 1
                If UCase(inv_no.List(i)) = UCase(inv_no.Text) Then
                    i = -10
                    Exit For
                End If
            Next
            If i <> -10 Then
                inv_no.AddItem inv_no.Text
                Call readshow
            End If
        End If
    End If
End If
End Sub



Private Sub TDBGrid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow Then
    If NewRow >= 0 Then
        On Error Resume Next
        TDBGrid1.TextMatrix(OldRow, 0) = w_old_str1
        w_old_str1 = TDBGrid1.TextMatrix(NewRow, 0)
        TDBGrid1.TextMatrix(NewRow, 0) = "★"
        TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
                
    End If

    Call refresh_grid2

End If
TDBGrid1.TextMatrix(0, 0) = " No"
TDBGrid1.ColAlignment(0) = flexAlignCenterCenter

End Sub

Private Sub TDBGrid1_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
If Col > 0 Then
    If ColIndex > gridc_501(0).Grid_Columns Then
        Cancel = 1
    Else
        If UCase(Mid(gridc_501(Col - 1).Grid_Visible, 1, 1)) = "F" Or gridc_501(Col - 1).Grid_Visible = "" Then
            Cancel = 1
        Else
            gridc_501(Col - 1).Grid_Width = TDBGrid1.ColWidth(Col)
        End If
    End If
End If

'移动ROW改变高度
If Row >= 0 Then
    w_cur_row = TDBGrid1.Row
    Row_Height1 = TDBGrid1.RowHeight(Row)
    Row_Height2 = TDBGrid1.RowHeight(Row)
    gridc_501(0).Grid_RowHeight = TDBGrid1.RowHeight(Row)
    gridc_502(0).Grid_RowHeight = TDBGrid1.RowHeight(Row)
    
    For i = 1 To TDBGrid1.Rows
        TDBGrid1.Row = i - 1
        TDBGrid1.RowHeight(i - 1) = Row_Height1
    Next i
    TDBGrid1.Row = w_cur_row

    For i = 1 To TDBGrid2.Rows
        TDBGrid2.Row = i - 1
        TDBGrid2.RowHeight(i - 1) = Row_Height2
    Next i

End If

End Sub


Private Sub TDBGrid2_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow Then
    If NewRow >= 0 Then
        TDBGrid2.TextMatrix(OldRow, 0) = w_old_str2
        w_old_str2 = TDBGrid2.TextMatrix(NewRow, 0)
        TDBGrid2.TextMatrix(NewRow, 0) = "★"
        TDBGrid2.ColAlignment(0) = flexAlignCenterCenter
                
    End If
End If
TDBGrid2.TextMatrix(0, 0) = " No"
TDBGrid2.ColAlignment(-1) = flexAlignCenterCenter

End Sub

Private Sub Tdbgrid2_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
If Col > 0 Then
    If ColIndex > gridc_502(0).Grid_Columns Then
        Cancel = 1
    Else
        If UCase(Mid(gridc_502(Col - 1).Grid_Visible, 1, 1)) = "F" Or gridc_502(Col - 1).Grid_Visible = "" Then
            Cancel = 1
        Else
            gridc_502(Col - 1).Grid_Width = TDBGrid2.ColWidth(Col)
        End If
    End If
End If

'移动ROW改变高度
If Row >= 0 Then
    w_cur_row = TDBGrid2.Row
    Row_Height2 = TDBGrid2.RowHeight(Row)
    gridc_502(0).Grid_RowHeight = TDBGrid2.RowHeight(Row)
    
    w_cur_row = TDBGrid1.Row
    Row_Height1 = TDBGrid2.RowHeight(Row)
    gridc_501(0).Grid_RowHeight = TDBGrid1.RowHeight(Row)
    
    For i = 1 To TDBGrid1.Rows
        TDBGrid1.Row = i - 1
        TDBGrid1.RowHeight(i - 1) = Row_Height1
    Next i
    TDBGrid1.Row = w_cur_row
End If
End Sub

Private Sub TDBGrid1_DblClick()

Call ViewTDBGridData(Adodc1.Recordset, gridc_501)
End Sub

Private Sub TDBGrid2_DblClick()
Call TDBGrid2_Click
Call ViewTDBGridData(Adodc2.Recordset, gridc_502)
End Sub

'弹出菜单,新增/修改或删除从档资料
Private Sub TDBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'如果不是左键
If Button <> 2 Then
    Exit Sub
End If
'如果没有inv_no
If Trim(inv_no.Text) = "" Then
    Exit Sub
End If
'检查单据状态
If check_ok() = False Then
    Exit Sub
End If
'这三个菜单项是整个系统共享的,应在此确保正确设置其enabled
Erp_Deliv.menu_add.Enabled = IIf(C_Off_Add, False, True)
Erp_Deliv.menu_delete.Enabled = IIf(C_Off_Delete, False, Adodc1.Recordset.EOF = False)
Erp_Deliv.menu_edit.Enabled = IIf(C_Off_Edit, False, Adodc1.Recordset.EOF = False)
PopupMenu Erp_Deliv.menu_modify
'菜单复位
Erp_Deliv.menu_add.Enabled = True
Erp_Deliv.menu_edit.Enabled = True
Erp_Deliv.menu_delete.Enabled = True
End Sub

Private Sub TDBGrid1_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
'鼠标点在HEADER上
If X > TDBGrid1.Left And Y < Row_Height1 Then
   
    '存储 TDBGrid 属性
    Call SaveVSGridSetting("mmss606", "TDBGrid1", gridc_501, g_CON_IniFile4)
    
    '调用 TDBGrid 属性设定
    With mmss_set
        Set .Parent_form = mmss606
        .Get_FormName = "mmss606"
        .Get_GridName = "TDBGrid1"
        .Gridc_File = g_CON_IniFile4
        .Show vbModal
    End With
    If TDBGrid1.Rows > 1 Then
        TDBGrid1.Col = 1
        TDBGrid1.Row = 1
    End If
End If
End Sub

Private Sub TDBGrid2_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
'鼠标点在HEADER上
If X + TDBGrid1.Width > TDBGrid2.Left And Y < Row_Height2 Then
   
    '存储 TDBGrid 属性
    Call SaveVSGridSetting("mmss606", "TDBGrid2", gridc_502, g_CON_IniFile4)
    
    '调用 TDBGrid 属性设定
    With mmss_set
        Set .Parent_form = mmss606
        .Get_FormName = "mmss606"
        .Get_GridName = "TDBGrid2"
        .Gridc_File = g_CON_IniFile4
        .Show vbModal
    End With
End If
End Sub

Private Sub TDBGrid1_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'不许更改第0行COl的宽度
If Col = 0 Then
    Cancel = True
End If
End Sub
Private Sub TDBGrid2_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'不许更改第0行COl的宽度
If Col = 0 Then
    Cancel = True
End If
End Sub

Private Sub TDBGrid2_Click()
On Error Resume Next
'当点击TDBGRID1 cell 时,移动 ADODC1.Recordset 指针
If Adodc2.Recordset.EOF = False Then
    Adodc2.Recordset.MoveFirst
    Adodc2.Recordset.Move TDBGrid2.Row - 1
    TDBGrid2.FocusRect = flexFocusNone
End If

End Sub

Private Sub inv_no_Click()
Dim w_rs511 As New ADODB.Recordset  '主档

If Not (C_Add Or C_Edit) Then
    W_Curr_MtrsNo = Trim(inv_no.Text)
    com_511.Parameters(0).Value = W_Curr_MtrsNo
    '重新执行原来的sql语句
    Set w_rs511 = com_511.Execute

    If w_rs511.EOF = False Then
        inv_no.Text = w_rs511!inv_no
        
        Inv_date.Value = NullSetValue(w_rs511!Inv_date, Date)
        pay_type.Text = NullSetValue(w_rs511!Pay_Scrpt, "")
        cust_no.Text = NullSetValue(w_rs511!cust_no, "")
        cust_ename.Text = NullSetValue(w_rs511!cust_ename, "")
        cust_eaddr.Text = NullSetValue(w_rs511!Cust_Caddr, "")
        
        close_date.Value = NullSetValue(w_rs511!close_date, Date)
        lc_no.Text = NullSetValue(w_rs511!lc_no, "")
        boat_date.Value = NullSetValue(w_rs511!boat_date, Date)
        boat_company = NullSetValue(w_rs511!boat_company, "")
        boat_name = NullSetValue(w_rs511!boat_name, "")
        lea_port = NullSetValue(w_rs511!lea_port, "")
        go_port = NullSetValue(w_rs511!go_port, "")
        remark.Text = NullSetValue(w_rs511!remark, "")
        
        w_close_date = NullSetValue(w_rs511!close_date, Date)
        W_inv_date = NullSetValue(w_rs511!Inv_date, Date)
        w_boat_date = NullSetValue(w_rs511!boat_date, Date)
        
    Else
        inv_no.Text = ""
        
        Inv_date.Value = Date
        pay_type.Text = ""
        cust_no.Text = ""
        cust_ename.Text = ""
        cust_eaddr.Text = ""
        
        close_date.Value = Date
        lc_no.Text = ""
        boat_date.Value = Date
        boat_company = ""
        boat_name = ""
        lea_port = ""
        go_port = ""
        remark.Text = ""
        w_close_date = Date
        W_inv_date = Date
        w_boat_date = Date
    End If
    w_rs511.Close
    Set w_rs511 = Nothing

    '刷新表格
    Call RefreshGrid
    Call readshow
    If Adodc1.Recordset.EOF = False Then
        TDBGrid1.Row = 1
        TDBGrid1.Col = 1
    Else
        Call TDBGrid1_AfterRowColChange(0, 0, 1, 1)
    End If
    
End If
End Sub

'刷新TDBGrid1,之所以定为public,是因为还会被表单frmcustquatmx调用
Public Sub RefreshGrid()

Call readactive
Set w_rs501 = Nothing
End Sub

'如果不是处于新增状态,则不让用户输入单号
Private Sub inv_no_KeyPress(KeyAscii As Integer)
'控制输入为字母或数字
If C_Add Then
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
           (KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ) Or _
           (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = vbKeyBack Or _
            KeyAscii = vbKeySpace) Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub cmd_brow_Click()
    If C_Add Or C_Edit Then
        With FrmList
            .G_Sql_Filter = "SELECT mmst021.Cust_No AS 客户代号, mmst021.Cust_Ename AS 英文名称, " & _
                                   "mmst021.Cust_caddr AS 客户地址, mmst802.Pay_Scrpt AS 付款方式 " & _
                              "FROM mmst802 RIGHT OUTER JOIN " & _
                                   "mmst021 ON mmst802.Pay_Type = mmst021.Pay_Type " & _
                              "WHERE mmst021.Cust_no like '" & Trim(cust_no.Text) & "%'"
                              
            .Caption = "客户列表"
            .Show vbModal
            If .Col_No1 <> "" Then
                cust_no.Text = .Col_No1
                cust_ename.Text = .Col_No2
                cust_eaddr.Text = .Col_No3
                pay_type.Text = NullSetValue(.Col_No4, "")
            End If
        End With
    End If
End Sub

Private Sub close_date_Change()
If Not (C_Add Or C_Edit) Then
    close_date.Value = w_close_date
End If
End Sub

Private Sub inv_date_Change()
If Not (C_Add Or C_Edit) Then
    Inv_date.Value = W_inv_date
End If
End Sub

Private Sub boat_date_Change()
If Not (C_Add Or C_Edit) Then
    boat_date.Value = w_boat_date
End If
End Sub


'自动生成单号 "前缀字符"I/O"+年份两位+月份+5位流水号
Private Function Creat_No()
Dim w_tmp As New ADODB.Recordset
Dim W_Str As String

Dim W_inv_No As String

W_inv_No = "V-"        '出库

W_inv_No = W_inv_No & Right(CStr(Year(Date)), 2) & Format(CStr(Month(Date)), "00") & Format(CStr(Day(Date)), "00")

W_Str = "SELECT Max(inv_no) As inv_no  FROM mmst511 WHERE inv_no like '" & W_inv_No & "%' "
            
w_tmp.Open W_Str, G_Con, adOpenForwardOnly, adLockReadOnly, adCmdText

If w_tmp.EOF = False Then
    If IsNull(w_tmp!inv_no) Then
        W_inv_No = W_inv_No & "001"
    Else
        W_inv_No = W_inv_No & Format(CStr(Val(Right(w_tmp!inv_no, 3)) + 1), "000")
    End If
Else
    W_inv_No = W_inv_No & "001"
End If

Creat_No = W_inv_No
End Function

'筛选打印数据并现实列印或预览效果
Private Sub sele_data()
Dim w_print As DAO.Recordset
Dim W_BookMark As Variant
Dim W_Rs As New ADODB.Recordset

W_Str = "SELECT  a.order_no,b.cust_order,a.mtr_no,c.mtr_name,c.mtr_ename," & _
                 "c.mtr_dim,c.unit_name,a.mtr_amt,case when f.deliv_type = '1' then 1 else -1 end * d.mtr_prs as mtr_prs,k.money_name," & _
                 "case when f.deliv_type = '1' then 1 else -1 end * d.mtr_prs * a.mtr_amt as mtr_total " & _
         "FROM mmst502 a INNER JOIN " & _
               "mmst501 f ON a.deliv_no = f.deliv_no INNER JOIN " & _
               "mmst011 b ON a.order_no = b.order_no INNER JOIN " & _
               "mmsp611 c ON a.mtr_no = c.mtr_no INNER JOIN " & _
               "mmst012 d ON a.order_no = d.order_no AND a.mtr_no = d.mtr_no INNER JOIN  " & _
               "mmst621 k ON b.money_no=k.money_no " & _
         "WHERE f.inv_no = '" & inv_no.Text & "'  ORDER BY f.deliv_type,f.deliv_no,a.mtr_no"

W_Rs.Open W_Str, G_Con, adOpenDynamic


'清除打印数据表
G_PrintDb.Execute "DELETE * FROM mmsr6061"
Set w_print = G_PrintDb.OpenRecordset("SELECT * FROM mmsr6061")

'选取数据
With w_print
    Do Until W_Rs.EOF
        .AddNew
        !loc_id = "A"
        !inv_no = inv_no.Text
        !Inv_date = CDate(Inv_date.Value)
        !pay_type = pay_type.Text
        !close_date = CDate(close_date.Value)
        
        !cust_no = cust_no.Text
        !cust_ename = cust_ename.Text
        !cust_eaddr = cust_eaddr.Text
        
        !lc_no = lc_no.Text
        !boat_company = boat_company.Text
        !boat_date = CDate(boat_date.Value)
        !boat_name = boat_name.Text
        !lea_port = lea_port.Text
        !go_port = go_port.Text
        
        !order_no = W_Rs!order_no
        !Cust_Order = W_Rs!Cust_Order
        !mtr_no = W_Rs!mtr_no
        !mtr_name = W_Rs!Mtr_ename & Space(1) & W_Rs!mtr_name
        
        !Unit_Name = W_Rs!Unit_Name
        !Mtr_Prs = W_Rs!Mtr_Prs
        !mtr_amt = W_Rs!mtr_amt
        !Amount = W_Rs!mtr_total
        !money_name = W_Rs!money_name
        .Update
        W_Rs.MoveNext
    Loop
End With
w_print.Close
W_Rs.Close

For i = 0 To 1000000

Next

If c_print Then
    c_print = False
    Call print_rpt(Erp_Deliv.Rpt1, "mmsr6061", "P")
End If
If c_view Then
    c_view = False
    Call print_rpt(Erp_Deliv.Rpt1, "mmsr6061", "V")
End If
If c_save Then
    c_save = False
    Set G_Rpt = Erp_Deliv.Rpt1
    G_Rpt_Name = "6061"
    mmssave.Show vbModal
End If
End Sub



'**********************************************************************
'装载按钮图片
'**********************************************************************

Private Sub load_picture()
cmd_find.PictureURL = App.Path + "\Picture\Norm\find_norm.bmp"
cmd_find.PictureDisableURL = App.Path + "\Picture\Dis\Find_dis.bmp"
cmd_find.PictureOverURL = App.Path + "\Picture\Over\Find_Over.bmp"

cmd_print.PictureURL = App.Path + "\Picture\Norm\print_norm.bmp"
cmd_print.PictureDisableURL = App.Path + "\Picture\Dis\print_dis.bmp"
cmd_print.PictureOverURL = App.Path + "\Picture\Over\print_Over.bmp"

cmd_preview.PictureURL = App.Path + "\Picture\Norm\preview_norm.bmp"
cmd_preview.PictureDisableURL = App.Path + "\Picture\Dis\preview_dis.bmp"
cmd_preview.PictureOverURL = App.Path + "\Picture\Over\preview_Over.bmp"

cmd_save.PictureURL = App.Path + "\Picture\Norm\save_norm.bmp"
cmd_save.PictureDisableURL = App.Path + "\Picture\Dis\save_dis.bmp"
cmd_save.PictureOverURL = App.Path + "\Picture\Over\save_Over.bmp"

cmd_ok.PictureURL = App.Path + "\Picture\Norm\Ok_norm.bmp"
cmd_ok.PictureDisableURL = App.Path + "\Picture\Dis\Ok_dis.bmp"
cmd_ok.PictureOverURL = App.Path + "\Picture\Over\Ok_Over.bmp"

cmd_cancel.PictureURL = App.Path + "\Picture\Norm\cancel_norm.bmp"
cmd_cancel.PictureDisableURL = App.Path + "\Picture\Dis\cancel_dis.bmp"
cmd_cancel.PictureOverURL = App.Path + "\Picture\Over\cancel_Over.bmp"

cmd_add.PictureURL = App.Path + "\Picture\Norm\add_norm.bmp"
cmd_add.PictureDisableURL = App.Path + "\Picture\Dis\add_dis.bmp"
cmd_add.PictureOverURL = App.Path + "\Picture\Over\add_Over.bmp"

cmd_edit.PictureURL = App.Path + "\Picture\Norm\edit_norm.bmp"
cmd_edit.PictureDisableURL = App.Path + "\Picture\Dis\edit_dis.bmp"
cmd_edit.PictureOverURL = App.Path + "\Picture\Over\edit_Over.bmp"

cmd_delete.PictureURL = App.Path + "\Picture\Norm\delete_norm.bmp"
cmd_delete.PictureDisableURL = App.Path + "\Picture\Dis\delete_dis.bmp"
cmd_delete.PictureOverURL = App.Path + "\Picture\Over\delete_Over.bmp"

Cmd_Quit.PictureURL = App.Path + "\Picture\Norm\Quit_norm.bmp"
Cmd_Quit.PictureDisableURL = App.Path + "\Picture\Dis\Quit_dis.bmp"
Cmd_Quit.PictureOverURL = App.Path + "\Picture\Over\Quit_Over.bmp"

AresButton1.PictureURL = App.Path + "\Picture\file.gif"
AresButton1.GifAnimationPlay
End Sub

'**********************************************************************
'更改提示符
'**********************************************************************

Private Sub cmd_find_MouseEnter()
Help_txt.Caption = cmd_find.ToolTipString
Help_txt.Refresh

End Sub

Private Sub cmd_find_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub cmd_print_MouseEnter()
Help_txt.Caption = cmd_print.ToolTipString
Help_txt.Refresh

End Sub

Private Sub cmd_print_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub cmd_preview_MouseEnter()
Help_txt.Caption = cmd_preview.ToolTipString
Help_txt.Refresh

End Sub

Private Sub cmd_preview_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub Cmd_save_MouseEnter()
Help_txt.Caption = cmd_save.ToolTipString
Help_txt.Refresh
End Sub
Private Sub Cmd_save_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub cmd_add_MouseEnter()
Help_txt.Caption = cmd_add.ToolTipString
Help_txt.Refresh

End Sub

Private Sub Cmd_Add_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub cmd_edit_MouseEnter()
Help_txt.Caption = cmd_edit.ToolTipString
Help_txt.Refresh
End Sub

Private Sub cmd_edit_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub cmd_delete_MouseEnter()
Help_txt.Caption = cmd_delete.ToolTipString
Help_txt.Refresh

End Sub

Private Sub cmd_delete_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub cmd_ok_MouseEnter()
Help_txt.Caption = cmd_ok.ToolTipString
Help_txt.Refresh

End Sub

Private Sub cmd_ok_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub cmd_cancel_MouseEnter()
Help_txt.Caption = cmd_cancel.ToolTipString
Help_txt.Refresh

End Sub

Private Sub cmd_cancel_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

Private Sub cmd_quit_MouseEnter()
Help_txt.Caption = Cmd_Quit.ToolTipString
Help_txt.Refresh

End Sub

Private Sub cmd_quit_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

'**********************************************************************
'按 ENTER调用 Mouse_click事件
'**********************************************************************
Private Sub Cmd_OK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_ok_MouseClick
End If
End Sub


Private Sub Cmd_cancel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_cancel_Mouseclick
End If
End Sub


Private Sub Cmd_find_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_find_MouseClick
End If
End Sub


Private Sub Cmd_print_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_print_MouseClick
End If
End Sub


Private Sub Cmd_preview_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_preview_MouseClick
End If
End Sub

Private Sub Cmd_Save_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_save_MouseClick
End If
End Sub

Private Sub Cmd_add_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_add_MouseClick
End If
End Sub

Private Sub Cmd_edit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_edit_MouseClick
End If
End Sub

Private Sub Cmd_delete_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_delete_MouseClick
End If
End Sub

Private Sub Cmd_quit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_quit_MouseClick
End If
End Sub




'**********************************************************************


Private Sub refresh_grid2()
  Dim w_rs502 As New ADODB.Recordset '从档
    Dim W_Str1 As String
    
    w_rs502.CursorLocation = adUseClient
    Dim w_cur_deli As String
    TDBGrid1.Col = 1
    w_cur_deli = TDBGrid1.Text
    
    '注意 从档的 SELECT 顺序一定和 INI 文件的顺序相同
    W_Str1 = "SELECT  a.order_no,b.cust_order,a.mtr_no,c.mtr_ename," & _
                               "c.mtr_dim,c.unit_name,a.mtr_amt,case when f.deliv_type = '1' then 1 else -1 end * d.mtr_prs as mtr_prs,b.money_no," & _
                               "case when f.deliv_type = '1' then 1 else -1 end * d.mtr_prs * a.mtr_amt as mtr_total " & _
                        "FROM mmst502 a INNER JOIN mmst501 f ON f.deliv_no = a.deliv_no INNER JOIN " & _
                              "mmst011 b ON a.order_no = b.order_no INNER JOIN " & _
                              "mmsp611 c ON a.mtr_no = c.mtr_no INNER JOIN " & _
                              "mmst012 d ON a.order_no = d.order_no AND a.mtr_no = d.mtr_no " & _
                " WHERE a.deliv_no= '" & TDBGrid1.Text & "'  " & _
                " ORDER BY  a.mtr_no"
                       
    w_rs502.Open W_Str1, G_Con, adOpenDynamic
    
    Set Adodc2.Recordset = w_rs502
    Set TDBGrid2.DataSource = Adodc2
    Set w_rs502 = Nothing
    Call SetVSGridSetting(TDBGrid2, gridc_502)
    '刷新全部 ROW 的高度 包括 HEADER
    For i = 1 To TDBGrid2.Rows
        TDBGrid2.Row = i - 1
        TDBGrid2.RowHeight(i - 1) = Row_Height2
        
        If i < TDBGrid2.Rows Then
            TDBGrid2.TextMatrix(i, 0) = i
        End If
    Next i
    TDBGrid2.ColAlignment(0) = flexAlignCenterCenter
End Sub


Private Sub cmd_find_SetFocus()
lab_focus.Visible = True
lab_focus.Top = cmd_find.Top
End Sub

Private Sub cmd_find_LeaveFocus()
lab_focus.Visible = False
End Sub

Private Sub cmd_print_SetFocus()
lab_focus.Visible = True
lab_focus.Top = cmd_print.Top
End Sub

Private Sub cmd_print_LeaveFocus()
lab_focus.Visible = False
End Sub

Private Sub cmd_preview_SetFocus()
lab_focus.Visible = True
lab_focus.Top = cmd_preview.Top
End Sub

Private Sub cmd_preview_LeaveFocus()
lab_focus.Visible = False
End Sub

Private Sub cmd_save_SetFocus()
lab_focus.Visible = True
lab_focus.Top = cmd_save.Top
End Sub

Private Sub cmd_save_LeaveFocus()
lab_focus.Visible = False
End Sub

Private Sub cmd_ok_SetFocus()
lab_focus.Visible = True
lab_focus.Top = cmd_ok.Top
End Sub

Private Sub cmd_ok_LeaveFocus()
lab_focus.Visible = False
End Sub

Private Sub cmd_cancel_SetFocus()
lab_focus.Visible = True
lab_focus.Top = cmd_cancel.Top
End Sub

Private Sub cmd_cancel_LeaveFocus()
lab_focus.Visible = False
End Sub

Private Sub cmd_add_SetFocus()
lab_focus.Visible = True
lab_focus.Top = cmd_add.Top
End Sub

Private Sub cmd_add_LeaveFocus()
lab_focus.Visible = False
End Sub

Private Sub cmd_edit_SetFocus()
lab_focus.Visible = True
lab_focus.Top = cmd_edit.Top
End Sub

Private Sub cmd_edit_LeaveFocus()
lab_focus.Visible = False
End Sub

Private Sub cmd_delete_SetFocus()
lab_focus.Visible = True
lab_focus.Top = cmd_delete.Top
End Sub

Private Sub cmd_delete_LeaveFocus()
lab_focus.Visible = False
End Sub

Private Sub cmd_quit_SetFocus()
lab_focus.Visible = True
lab_focus.Top = Cmd_Quit.Top
End Sub

Private Sub cmd_quit_LeaveFocus()
lab_focus.Visible = False
End Sub






