VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F4732CE3-9A6C-11D2-8018-0080AD70A386}#5.7#0"; "AresButtonPro.ocx"
Begin VB.Form mmss904 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����Ϻ����(904)"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   15165
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   2490
      ScaleHeight     =   2475
      ScaleWidth      =   12615
      TabIndex        =   11
      Top             =   810
      Width           =   12645
      Begin VB.CommandButton cmd_order 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2550
         TabIndex        =   42
         Top             =   120
         Width           =   300
      End
      Begin VB.TextBox order_no 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1290
         TabIndex        =   38
         Top             =   120
         Width           =   1560
      End
      Begin VB.TextBox remark 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2040
         Width           =   4455
      End
      Begin VB.TextBox re_amt 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7770
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton cmd_brow 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   10920
         TabIndex        =   26
         Top             =   585
         Width           =   300
      End
      Begin VB.CommandButton cmd_brow_type 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   9495
         TabIndex        =   25
         Top             =   135
         Width           =   300
      End
      Begin VB.TextBox Mtr_Name 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   7770
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1095
         Width           =   4455
      End
      Begin VB.TextBox Mtr_Dim 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   7770
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4455
      End
      Begin VB.CommandButton cmd_brow 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4440
         TabIndex        =   15
         Top             =   585
         Width           =   300
      End
      Begin VB.TextBox Mtr_Dim 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4455
      End
      Begin VB.TextBox Mtr_Name 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1100
         Width           =   4455
      End
      Begin VB.TextBox mtr_no 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1290
         MaxLength       =   26
         TabIndex        =   16
         Top             =   585
         Width           =   3465
      End
      Begin VB.TextBox type_name 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   7770
         TabIndex        =   28
         Top             =   120
         Width           =   2040
      End
      Begin VB.TextBox mtr_no 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   7770
         MaxLength       =   26
         TabIndex        =   27
         Top             =   585
         Width           =   3465
      End
      Begin VB.CommandButton cmd_brow_type 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   5430
         TabIndex        =   40
         Top             =   150
         Width           =   300
      End
      Begin VB.TextBox type_name 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   4170
         TabIndex        =   41
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   39
         Top             =   180
         Width           =   1020
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��   ע:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   37
         Top             =   2160
         Width           =   930
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   6630
         TabIndex        =   35
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ϴ���:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   6120
         TabIndex        =   33
         Tag             =   "Material Code:"
         Top             =   720
         Width           =   1980
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   6600
         TabIndex        =   32
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label bom_no 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bom_no"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   6480
         TabIndex        =   31
         Tag             =   "Material Code:"
         Top             =   840
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��        ��:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   6600
         TabIndex        =   30
         Top             =   180
         Width           =   795
      End
      Begin VB.Label mtr_type 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   6750
         TabIndex        =   29
         Top             =   315
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   6600
         TabIndex        =   24
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ϴ���:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Tag             =   "Material Code:"
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   1050
      End
      Begin VB.Label bom_no 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "bom_no"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Tag             =   "Material Code:"
         Top             =   840
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��        ��:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   3000
         TabIndex        =   18
         Top             =   180
         Width           =   795
      End
      Begin VB.Label mtr_type 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   17
         Top             =   315
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1050
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8325
      Left            =   0
      ScaleHeight     =   8295
      ScaleWidth      =   2475
      TabIndex        =   8
      Top             =   810
      Width           =   2505
      Begin ARESBUTTONLib.AresButton cmd_ok 
         Height          =   360
         Left            =   690
         TabIndex        =   0
         Top             =   420
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Ok_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Ok_Norm.bmp"
         ToolTipString   =   "ȷ�ϴ���"
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
         Height          =   420
         Left            =   690
         TabIndex        =   1
         Top             =   1200
         Width           =   495
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Cancel_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Cancel_Norm.bmp"
         ToolTipString   =   "��������"
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
         Height          =   420
         Left            =   690
         TabIndex        =   2
         Top             =   1995
         Width           =   495
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Add_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Add_Norm.bmp"
         ToolTipString   =   "����һ�ʼ�¼"
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
         Height          =   420
         Left            =   690
         TabIndex        =   3
         Top             =   2775
         Width           =   495
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Edit_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Edit_Norm.bmp"
         ToolTipString   =   "�޸ĸü�¼"
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
         Height          =   420
         Left            =   690
         TabIndex        =   4
         Top             =   3570
         Width           =   495
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Delete_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Delete_Norm.bmp"
         ToolTipString   =   "ɾ���ü�¼"
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
         Left            =   690
         TabIndex        =   5
         Top             =   4350
         Width           =   1125
         _Version        =   327687
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Quit_Norm.bmp"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\Norm\Quit_Norm.bmp"
         ToolTipString   =   "�˳��ó�ʽ"
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
      Begin ARESBUTTONLib.AresButton AresButton1 
         Height          =   825
         Left            =   780
         TabIndex        =   9
         Top             =   6270
         Width           =   1050
         _Version        =   327687
         BackGroundColor =   16777215
         PictureURL      =   "Y:\c_sys\sxc\XuSheng\Picture\erp_proj.gif"
         PictureBaseURL  =   "Y:\c_sys\sxc\XuSheng\Picture\erp_proj.gif"
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
         PrevPointer     =   90897964
         _ExtentX        =   873
         _ExtentY        =   741
         _StockProps     =   80
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   2490
         Y1              =   5070
         Y2              =   5070
      End
      Begin VB.Label Help_txt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   300
         TabIndex        =   10
         Top             =   5400
         Width           =   1755
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   30
      Top             =   6900
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
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
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VSFlex7Ctl.VSFlexGrid TDBGrid1 
      Bindings        =   "mmss904.frx":0000
      Height          =   5595
      Left            =   2520
      TabIndex        =   14
      Top             =   3480
      Width           =   12555
      _cx             =   22146
      _cy             =   9869
      _ConvInfo       =   -1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      FormatString    =   $"mmss904.frx":0015
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
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�����Ϻ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   465
      Left            =   7380
      TabIndex        =   12
      Top             =   180
      Width           =   2985
   End
End
Attribute VB_Name = "mmss904"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'*��������: �û����ϵ�(mmss904)
'*��д����: 2002��07��29��
'*������Ա: ��
'*�޸�����:
'*�޸���Ա:
'***********************************************
'�������򿪵����ݿ⼰���ݱ�����
Dim St_90d0 As New ADODB.Recordset
Dim St_90d1 As New ADODB.Recordset
'���尴ť����
Dim c_add As Boolean
Dim c_edit As Boolean
Dim c_delete As Boolean

'���TDBGRID1 �ľ��ַ�
Dim W_Old_Str As String
Dim W_Mtr_Type As String

Dim c_off_add As Boolean
Dim c_off_edit As Boolean
Dim c_off_delete As Boolean

'��¼��ǰ����
Dim W_col As Double
Dim W_Row As Double

'���崰��򿪱���
Dim Gridc_90d0(127) As Grid_Data '��� Grid ����ֵ
'Dim Gridc_90d1(127) As Grid_Data '��� Grid ����ֵ
Dim Row_Height As Double        'Grid �߶ȱ���
Private Sub cmd_brow_Click(Index As Integer)
If order_no.Text <> "" Then
If c_add Then
    With FrmSectList
'         .W_edit_able = True
'         .W_Form_name = "frm40bmx"
         .W_Field1 = "A.MTR_NO "
'         .W_Orderby = "a.mtr_type  "
         .Quer_status = True
         .W_Select_Data = " select a.mtr_no as ���ϱ��,mtr_name as Ʒ��,mtr_dim as ���,color_name as ��ɫ,a.mtr_amt as ��������," & _
                                " (isnull(ling_amt,0)+isnull(out_ling_amt,0)) as ��������,(a.mtr_amt-isnull(ling_amt,0)-isnull(out_ling_amt,0)+isnull(tui_amt,0)) as ��������,isnull(b.mtr_amt,0) as �������,c.bar_name as �ֱ�����,unit_name as ��λ,isnull(C.BAR_NO,a.bar_No)  " & _
                          " from mmsp012_mtr a left join mmst381 b on a.mtr_no=b.mtr_no  left join mmst903 c on c.bar_no=isnull(b.bar_no,a.bar_no)  " & _
                          " where a.order_no like '" & Trim(order_no.Text) & "' and a.mtr_no like '" & Trim(mtr_no(Index).Text) & "%' " & _
                          " AND (a.mtr_amt-isnull(ling_amt,0)-isnull(out_ling_amt,0)+isnull(tui_amt,0))>0  "
         .Grid1.Editable = flexEDKbdMouse
'         .Grid1.ColHidden(.Grid1.Cols - 1) = True
         .Show vbModal
         If .cancel_status = False And .List2 <> "" Then
            W_Bat = True
            mtr_no(Index).Text = .List1
            Mtr_Name(Index).Text = .List2
            Mtr_Dim(Index).Text = .List3
         
         End If
    End With

End If
    
Else

    With FrmMtrList
           .G_Type = ""
           .G_Mtr_No = Trim(mtr_no(Index).Text)
           .G_Mtr_Type = mtr_type(Index).Caption
        .Show vbModal
        If .mtr_no <> "" Then
    '        mtr_type(index).Text = .Mtr_Type_Tmp
            
            mtr_no(Index).Text = .mtr_no
            Mtr_Name(Index).Text = .Mtr_Name
            Mtr_Dim(Index).Text = .Mtr_Dim
            
            mtr_no(Index).SetFocus
        End If
    End With
End If
End Sub

Private Sub cmd_brow_type_Click(Index As Integer)
With FrmMtrType
    .Show vbModal
    type_name(Index).Text = .type_name
    mtr_type(Index).Caption = .Type_ID
    If Index = 0 Then
        type_name(1).Text = .type_name
        mtr_type(1).Caption = .Type_ID
    End If
End With
End Sub

Private Sub cmd_order_Click()
 With FrmSectList
'         .W_edit_able = False
         .Quer_status = False
         .W_Select_Data = " select order_no as �Ƶ�����,cust_no as �ͻ����,cust_name as �ͻ�����,cust_order_no as ��������,mtr_no as ��Ʒ�ͺ�,cust_mtr_no as �ͻ��ͺ�,mtr_amt as �Ƶ�����  " & _
                          " from mmsp011 " & _
                          " where status='2' and order_no like '" & Trim(order_no.Text) & "%' " & _
                          " order by order_no "
         .Grid1.Editable = flexEDNone
         .Show vbModal
         If .cancel_status = False And .List1 <> "" Then
            order_no.Text = .List1
'            Cust_Order_No.Text = .List4
'            old_mtr_no.Text = .List5
'            Cust_Mtr_No.Text = .List6
'            Text1.Text = .List7
            mtr_no(Index).SetFocus
         End If
    End With
End Sub

Public Sub Form_Activate()
'�����ڼ���ʱ,ˢ��TDBGrid
Call GetVSGridSetting("mmss904", "TDBGrid1", Gridc_90d0, g_CON_IniFile9)
Row_Height = Gridc_90d0(0).Grid_RowHeight

'Call GetVSGridSetting("mmss904", "TDBGrid2", Gridc_90d1, g_CON_IniFile9)
'Row_Height = Gridc_90d1(0).Grid_RowHeight
'Call readactive
'ˢ�±��
Call RefreshGrid
TDBGrid1.Col = 1
If TDBGrid1.Rows > 1 Then
    TDBGrid1.Row = 1
End If

End Sub

Private Sub Form_Load()
'װ��ͼƬ
Call load_picture
'����������
Call CenterWindow(mmss904, sys_main)

'����ť��������ֵ
c_add = False
c_edit = False
c_delete = False
c_off_add = False
c_off_edit = False
c_off_delete = False

'MDI�Ӵ��ڰ�ťȨ���趩
c_off_add = lopcheck("A", "904", G_User_ID)
c_off_edit = lopcheck("U", "904", G_User_ID)
c_off_delete = lopcheck("D", "904", G_User_ID)

If c_off_add = True Then
    cmd_add.Enabled = False
End If

If c_off_edit = True Then
    cmd_edit.Enabled = False
End If

If c_off_delete = True Then
    cmd_delete.Enabled = False
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
    KeyAscii = 34
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Form_KeyDown���հ�
If KeyCode = vbKeyReturn And Me.ActiveControl.Name <> "TDBGrid1" Then
    
    If ActiveControl.Name = "note" Then
        If ActiveControl.MultiLine = False Then
            SendKeys "{TAB}"
        End If
    Else
        SendKeys "{TAB}"
    End If
    Exit Sub
End If

If Shift = 0 Then
    Select Case KeyCode
        '�� ESC ʱȡ������
        Case vbKeyEscape
            If (c_add Or c_edit Or c_delete) Then
                Call vcontrol("N")
                KeyCode = 0
            End If
        '���� F2 ʱ,�¡���¼
        Case vbKeyF2
            If cmd_add.Enabled Then
                Call vcontrol("A")
                KeyCode = 0
            End If
        '���� F3 ʱ , �޸ļ�¼
        Case vbKeyF3
            If Me.ActiveControl.Name = "mtr_no" Then
'                Call mtr_no(0)_LostFocus
            End If

            If cmd_edit.Enabled Then
                Call vcontrol("U")
                KeyCode = 0
            End If
        '���� F4 ʱ , ɾ����¼
        Case vbKeyF4
            If Me.ActiveControl.Name = "mtr_no" Then
'                Call mtr_no(0)_LostFocus
            End If

            If cmd_delete.Enabled Then
                Call vcontrol("D")
                KeyCode = 0
            End If
        '���� F5 ʱ,�����¼
        Case vbKeyF5
            If cmd_ok.Enabled Then
                Call vcontrol("Y")
                KeyCode = 0
            End If
        '���� F6 ʱ,�˳�ϵͳ
        Case vbKeyF6
            If cmd_quit.Enabled Then
                Call vcontrol("Q")
                KeyCode = 0
            End If
    End Select
End If
End Sub

Sub readshow(Index As Integer)
'�Կؼ���ֵ
If c_add = True Or Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF Then
    mtr_no(0).Text = ""
    Mtr_Name(0).Text = ""
    Mtr_Dim(0).Text = ""
    type_name(0).Text = ""
    type_name(1).Text = ""
    
    mtr_type(0).Caption = ""
    mtr_type(1).Caption = ""
    
    mtr_no(1).Text = ""
    Mtr_Name(1).Text = ""
    Mtr_Dim(1).Text = ""
    remark.Text = ""
    re_amt.Text = ""
    order_no.Text = ""
    
Else

    order_no.Text = NullSetValue(St_90d0!order_no, "")
    mtr_no(Index).Text = St_90d0!mtr_no
    Mtr_Name(Index).Text = NullSetValue(St_90d0!Mtr_Name, "")
    Mtr_Dim(Index).Text = NullSetValue(St_90d0!Mtr_Dim, "")

    remark.Text = NullSetValue(St_90d0!remark, "")


    re_amt.Text = NullSetValue(St_90d0!re_amt, "")
    mtr_no(Index).Text = NullSetValue(St_90d0!mtr_no, "")

End If

'�趨������ Enabled ����
If c_add Or c_edit Or c_delete Then
    cmd_add.Enabled = False
    cmd_edit.Enabled = False
    cmd_delete.Enabled = False
    cmd_ok.Enabled = True
    cmd_cancel.Enabled = True
Else
    cmd_add.Enabled = True
    cmd_edit.Enabled = True
    cmd_delete.Enabled = True
    cmd_ok.Enabled = False
    cmd_cancel.Enabled = False
    
    '�����ݱ��޼�¼ʱ
    If St_90d0.EOF Then
        cmd_edit.Enabled = False
        cmd_delete.Enabled = False
    Else
        cmd_edit.Enabled = True
        cmd_delete.Enabled = True
    End If
End If

'ͨ��Ȩ���趨������ Enabled
If c_off_add = True Then
    cmd_add.Enabled = False
End If
    
If c_off_edit = True Then
    cmd_edit.Enabled = False
End If
    
If c_off_delete = True Then
    cmd_delete.Enabled = False
End If

If Not (c_add Or c_edit) Then
    Mtr_Name(Index).Locked = True
    type_name(Index).Locked = True
    Mtr_Dim(Index).Locked = True
    mtr_no(Index).Locked = True
    cmd_brow(Index).Enabled = False
    cmd_brow_type(Index).Enabled = False
remark.Locked = True
re_amt.Locked = True
Else
    Mtr_Name(Index).Locked = False
    type_name(Index).Locked = False
    Mtr_Dim(Index).Locked = False
    mtr_no(Index).Locked = False
    cmd_brow(Index).Enabled = True
    cmd_brow_type(Index).Enabled = True
remark.Locked = False
re_amt.Locked = False

End If
If c_edit Then
    mtr_no(Index).Locked = True
Else
    mtr_no(Index).Locked = False
End If
End Sub
'ˢ�±��
Private Sub RefreshGrid()
Call readactive

Call readshow(0)
Call readshow(1)
End Sub

Private Sub readactive()
Set St_90d0 = Nothing
With St_90d0
    .ActiveConnection = G_Con
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockPessimistic
    .Open "select a.order_no,a.Mtr_No  ,b.Mtr_Name,b.Mtr_Dim,re_mtr,c.mtr_name,c.mtr_dim ,a.re_amt,a.remark," & _
                 "" & _
                 "a.upd_name," & _
                 "a.upd_date " & _
            "FROM mmst90d a inner join mmst611 b on a.mtr_no=b.mtr_no inner join mmst611 c on c.mtr_no=re_mtr " & _
            " ORDER BY a.Mtr_No "

End With

'����tdbgrid1��������Դ
Set Adodc1.Recordset = St_90d0

Call SetVSGridSetting(TDBGrid1, Gridc_90d0)

'ˢ��ȫ�� ROW �ĸ߶� ���� HEADER
'For i = 1 To TDBGrid1.Rows
'    TDBGrid1.Row = i - 1
'    TDBGrid1.RowHeight(i - 1) = Row_Height
'
'    If i < TDBGrid1.Rows Then
'        TDBGrid1.TextMatrix(i, 0) = i
'    End If
'Next i
'TDBGrid1.ColAlignment(0) = flexAlignCenterCenter

End Sub

'������¼�
Private Sub Cmd_Add_MouseClick()
Call vcontrol("A")
End Sub

Private Sub Cmd_Edit_MouseClick()
Call vcontrol("U")
End Sub

Private Sub Cmd_Delete_MouseClick()
Call vcontrol("D")
End Sub

Private Sub Cmd_OK_MouseClick()
Call vcontrol("Y")
End Sub

Private Sub cmd_cancel_MouseClick()
Call vcontrol("N")
End Sub

Private Sub cmd_quit_MouseClick()
Call vcontrol("Q")
End Sub

Private Sub vcontrol(p_choice As String)
Select Case p_choice
    Case "Y"            'ȷ��
        If check_ok() = True Then
            Call upd_data
            TDBGrid1.Enabled = True
        End If
        
    Case "N"            'ȡ��
        '��������
        If c_edit Or c_delete Then
            Call UnLockRecord("mmst90d", "mtr_no='" & mtr_no(0).Text & "'")
        End If
        c_add = False
        c_edit = False
        c_delete = False
        
        TDBGrid1.Enabled = True
        Call readshow(0)
        
    Case "A"             '����
        c_add = True
        Call readshow(0)
        mtr_no(0).SetFocus
        TDBGrid1.Enabled = False
        
    Case "U"             '�޸�
        '����
        If LockRecord("mmst90d", "mtr_no='" & mtr_no(0).Text & "'") Then
            W_Row = TDBGrid1.Row
            W_col = TDBGrid1.Col
            
            c_edit = True
            TDBGrid1.Enabled = False
            Call readshow(0)
'            Mtr_Name(0).SetFocus
        End If
        
    Case "D"             'ɾ��
        '����
        If LockRecord("mmst90d", "mtr_no='" & mtr_no(0).Text & "'") = True Then
            If MsgBox(g_CON_CDelete, vbYesNo + vbDefaultButton2 + vbInformation, g_CON_CTitle) = vbNo Then
                Call UnLockRecord("mmst90d", "mtr_no='" & mtr_no(0).Text & "'")
                Exit Sub
            End If
            
            '�ж��Ƿ����ɾ��
            c_delete = True
            If check_ok = False Then
                Call UnLockRecord("mmst90d", "mtr_no(0)='" & mtr_no(0).Text & "'")
                c_delete = False
                Exit Sub
            End If
            
            'ɾ����¼
            G_Con.Execute "DELETE FROM mmst90d WHERE mtr_no='" & Trim(mtr_no(0).Text) & "'"
            c_delete = False
            'ˢ������
            
            Call RefreshGrid
            
            'ɾ�����ƶ�����һ�ʼ�¼
            TDBGrid1.Col = 1
            If TDBGrid1.Rows > 1 Then
                TDBGrid1.TopRow = 1
                TDBGrid1.Row = 1
            End If
            
        End If
    Case "Q"            '�˳�
        Unload Me
End Select
End Sub

'���޸Ļ�ɾ��������ʱ����һ�����ж�
Private Function check_ok() As Boolean
Dim w_tmp As New ADODB.Recordset
'If c_delete Then
'    If mtr_no(0) = "A001" Then
'        MsgBox "���û�Ϊϵͳ�û�,����ɾ��!", 64, "��ʾ��Ϣ"
'        mtr_no(0).SetFocus
'        check_ok = False
'        Exit Function
'    End If
'    check_ok = True
'End If
'����ʱ�ж�
If c_add = True Then
    If Trim(mtr_no(0).Text) = "" Then
        MsgBox "����������ϱ��", 64, "��ʾ��Ϣ"
        mtr_no(0).SetFocus
        check_ok = False
        Exit Function
    Else
        '�жϴ����Ƿ��ظ�
        w_tmp.CursorLocation = adUseClient
        w_tmp.Open "select Mtr_No from mmst611 where Mtr_No = '" & Trim(mtr_no(0).Text) & "'", G_Con, adOpenForwardOnly
        If w_tmp.EOF Then
            MsgBox "�����ϱ�Ų�����,������ȷ��!", 64, "��ʾ��Ϣ"
            mtr_no(0).SetFocus
            check_ok = False
            Set w_tmp = Nothing
            Exit Function
        End If
        Set w_tmp = Nothing
    End If
    
 If Trim(mtr_no(0).Text) = Trim(mtr_no(1).Text) Then
    MsgBox "�¾�������������ϺŶ���ͬ,�α��޸İ�", 64, "��ʾ��Ϣ"
    mtr_no(1).SetFocus
    check_ok = False
    Exit Function
End If
    
End If



'
''�������޸�ʱ�ж�
'If mtr_no(1).Text = "" Then
'    MsgBox "�����������ϱ��", 64, "��ʾ��Ϣ"
'    mtr_no(1).SetFocus
'    check_ok = False
'    Exit Function
'Else
'    If Trim(mtr_no(0).Text) = Trim(mtr_no(1).Text) Then
'
'        '�ж��û������Ƿ��ظ�
'        w_tmp.CursorLocation = adUseClient
'        w_tmp.Open "select Mtr_No from mmst90d where Mtr_No= '" & Trim(mtr_no(0).Text) & "' and  and order_no='" & Trim(order_no.Text) & "'", G_Con, adOpenForwardOnly
'
'        If w_tmp.EOF = False Then
'            MsgBox "�����ϱ���Ѿ�����!", 64, "��ʾ��Ϣ"
'            New_Mtr_No.SetFocus
'            Set w_tmp = Nothing
'            check_ok = False
'            Exit Function
'        End If
'        Set w_tmp = Nothing
'    End If
'End If


'If type_name(0).Text = "" Then
'    MsgBox "�������������", 64, "��ʾ��Ϣ"
'    type_name(0).SetFocus
'    check_ok = False
'    Exit Function
'Else
'    w_tmp.Open "select Mtr_Type from mmst603 where type_name='" & type_name(0).Text & "'", G_Con, , , adCmdText
'    If w_tmp.EOF = True Then
'        w_tmp.Close
'        MsgBox "�޴��������.", vbExclamation, g_CON_CTitle
'        type_name(0).SetFocus
'        Exit Function
'    Else
'        W_Mtr_Type = w_tmp!mtr_type
'    End If
'    w_tmp.Close
'End If


check_ok = True
End Function

'�����ݿ���и���
Private Sub upd_data()
Dim St_90d0_1 As New ADODB.Recordset
Dim st_905 As New ADODB.Recordset

Dim W_Find As String

W_Find = mtr_no(0).Text

On Error GoTo UpdateError
G_Con.BeginTrans

With St_90d0_1
    .ActiveConnection = G_Con
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockPessimistic
    .Open "select * from mmst90d where mtr_no='" & mtr_no(0).Text & "' and re_mtr='" & mtr_no(1).Text & "'"
End With

'����һ�ʼ�¼�����ݿ�
If c_add = True Then
    If St_90d0_1.EOF Then
        With St_90d0_1
            .AddNew
            !order_no = Trim(order_no.Text)
            !mtr_no = UCase(Trim(mtr_no(0).Text))
            !re_mtr = UCase(Trim(mtr_no(1).Text))
            !re_amt = Val(re_amt.Text)
            !upd_name = G_User_Name
            !upd_date = Get_SQLDATE
            !remark = remark.Text
            !lock = "No"
            .Update
        End With
    End If
    Set St_90d0_1 = Nothing
    c_add = False
End If

'����һ�ʼ�¼�����ݿ�
If c_edit = True Then
    If St_90d0_1.EOF = False Then
        With St_90d0_1
            !re_amt = Val(re_amt.Text)
            !upd_name = G_User_Name
            !upd_date = Get_SQLDATE
            !remark = remark.Text
            !lock = "No"
            .Update
        End With
    End If
    Set St_90d0_1 = Nothing
    c_edit = False
End If


G_Con.CommitTrans
'ˢ�����ݱ�
Call RefreshGrid

TDBGrid1.Row = TDBGrid1.FindRow(W_Find, 0, 2, False)
TDBGrid1.Col = W_col
TDBGrid1.TopRow = TDBGrid1.FindRow(W_Find, 0, 2, False)


Endx:
c_add = False

'Call RefreshGrid

Exit Sub

UpdateError:
G_Con.RollbackTrans
MsgBox "����ʱ��������!", 64, g_CON_CTitle
GoTo Endx

End Sub

'���� QueryUnload �� Unload �¼�
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If c_add Or c_edit Or c_delete Then
    '�������ݸĶ�ʱ.ѯ���Ƿ�Ҫ�˳�ϵͳ
    If MsgBox(g_CON_CQuit, vbQuestion + vbYesNo, g_CON_CTitle) = vbNo Then
        Cancel = 1
    Else
        '�����޸Ļ�ɾ��ʱδ����ʱ,�������
        If c_edit Or c_delete Then
            Call UnLockRecord("mmst90d", "mtr_no(0)='" & mtr_no(0).Text & "'")
        End If
        Cancel = 0
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

'�˳�ʱ���洢 TDBGrid ����
Call SaveGridSetting("mmss904", "TDBGrid1", Gridc_90d0, g_CON_IniFile9)

Set TDBGrid1.DataSource = Nothing
Set St_90d0 = Nothing
Set mmss904 = Nothing
End Sub
Private Sub TDBGrid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow Then
    If NewRow >= 0 Then
        TDBGrid1.TextMatrix(OldRow, 0) = W_Old_Str
        W_Old_Str = TDBGrid1.TextMatrix(NewRow, 0)
        TDBGrid1.TextMatrix(NewRow, 0) = "��"
        TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
                
    End If
    '�����TDBGRID1 cell ʱ,�ƶ� ADODC1.Recordset ָ��
    If Adodc1.Recordset.EOF = False Then
        Adodc1.Recordset.MoveFirst
        Adodc1.Recordset.Move TDBGrid1.Row - 1
        TDBGrid1.FocusRect = flexFocusRaised
    End If
    Call readshow(0)
    Call readshow(1)
'    Call readactive1
End If
TDBGrid1.TextMatrix(0, 0) = " No"
TDBGrid1.ColAlignment(0) = flexAlignCenterCenter
End Sub

Private Sub TDBGrid1_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'�ƶ�COl�ı���
If Col > 0 Then
    If Col > Gridc_90d0(0).Grid_Columns Then
        Cancel = 1
    Else
        If UCase(Mid(Gridc_90d0(Col - 1).Grid_Visible, 1, 1)) = "F" Or Gridc_90d0(Col - 1).Grid_Visible = "" Then
            Cancel = 1
        Else
            Gridc_90d0(Col - 1).Grid_Width = TDBGrid1.ColWidth(Col)
        End If
    End If
End If

'�ƶ�ROW�ı�߶�
If Row >= 0 Then
    w_cur_row = TDBGrid1.Row
    Row_Height = TDBGrid1.RowHeight(Row)
    Gridc_90d0(0).Grid_RowHeight = TDBGrid1.RowHeight(Row)
    
    For i = 1 To TDBGrid1.Rows
        TDBGrid1.Row = i - 1
        TDBGrid1.RowHeight(i - 1) = Row_Height
    Next i
    TDBGrid1.Row = w_cur_row
End If

End Sub

Private Sub TDBGrid1_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
'������HEADER��
If X > 0 And Y < Row_Height Then
   
    '�洢 TDBGrid ����
    Call SaveVSGridSetting("mmss904", "TDBGrid1", Gridc_90d0, g_CON_IniFile9)
    
    '���� TDBGrid �����趨
    With mmss_set
        Set .Parent_form = mmss904
        .Get_FormName = "mmss904"
        .Get_GridName = "TDBGrid1"
        .Gridc_File = g_CON_IniFile9
        .Show vbModal
    End With
End If
End Sub

Private Sub TDBGrid1_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'������ĵ�0��COl�Ŀ��
If Col = 0 Then
    Cancel = True
End If
End Sub


'**********************************************************************
'װ�ذ�ťͼƬ
'**********************************************************************
Private Sub load_picture()
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

cmd_quit.PictureURL = App.Path + "\Picture\Norm\Quit_norm.bmp"
cmd_quit.PictureDisableURL = App.Path + "\Picture\Dis\Quit_dis.bmp"
cmd_quit.PictureOverURL = App.Path + "\Picture\Over\Quit_Over.bmp"

AresButton1.PictureURL = App.Path + "\Picture\file.gif"
AresButton1.GifAnimationPlay
End Sub

'**********************************************************************
'������ʾ��
'**********************************************************************

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
Help_txt.Caption = cmd_quit.ToolTipString
Help_txt.Refresh

End Sub

Private Sub Cmd_quit_MouseLeave()
Help_txt.Caption = ""
Help_txt.Refresh
End Sub

'**********************************************************************
'�� ENTER���� Mouse_click�¼�
'**********************************************************************
Private Sub Cmd_OK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_OK_MouseClick
End If
End Sub


Private Sub Cmd_cancel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_cancel_MouseClick
End If
End Sub


Private Sub Cmd_add_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_Add_MouseClick
End If
End Sub

Private Sub Cmd_edit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_Edit_MouseClick
End If
End Sub

Private Sub Cmd_delete_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call Cmd_Delete_MouseClick
End If
End Sub

Private Sub Cmd_quit_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmd_quit_MouseClick
End If
End Sub
