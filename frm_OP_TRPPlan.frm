VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_OP_TRPPlan 
   Caption         =   " ��  ��  �@  �~"
   ClientHeight    =   7920
   ClientLeft      =   225
   ClientTop       =   900
   ClientWidth     =   15405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   15405
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   10800
      TabIndex        =   98
      Top             =   5520
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   -2147483624
      StartOfWeek     =   50266113
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38232
      MaxDate         =   2958455
   End
   Begin TabDlg.SSTab SSTAB1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "�@��ƨ��@�~"
      TabPicture(0)   =   "frm_OP_TRPPlan.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fam_SrcOrders"
      Tab(0).Control(1)=   "fam_SelectedOrders"
      Tab(0).Control(2)=   "fam_RouteData"
      Tab(0).Control(3)=   "txtReceipt_No"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "���u�s���C��"
      TabPicture(1)   =   "frm_OP_TRPPlan.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "dg_Tab1_Route"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "dg_Tab1_RouteOrders"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "�O�d�q��"
      TabPicture(2)   =   "frm_OP_TRPPlan.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmd_Tab2_ShowAll"
      Tab(2).Control(1)=   "cmd_Tab2_FilterAndSort"
      Tab(2).Control(2)=   "cmd_Tab2_Delete"
      Tab(2).Control(3)=   "cmd_Tab2_Reset"
      Tab(2).Control(4)=   "cmd_Tab2_Remove"
      Tab(2).Control(5)=   "dg_Tab2_ReservedOrders"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "���n�ϼжK�C�L"
      TabPicture(3)   =   "frm_OP_TRPPlan.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "fma_Tab3_OrderSum"
      Tab(3).Control(2)=   "dg_RouteData"
      Tab(3).ControlCount=   3
      Begin VB.Frame Frame5 
         Height          =   2370
         Left            =   -74880
         TabIndex        =   99
         Top             =   420
         Width           =   12465
         Begin VB.CommandButton cmd_Tab0_Print 
            BackColor       =   &H00C0FFC0&
            Caption         =   "�C�LBarCode"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   10200
            Picture         =   "frm_OP_TRPPlan.frx":0070
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   108
            Top             =   120
            Width           =   1035
         End
         Begin VB.CheckBox Check1 
            Caption         =   "�w���C�L"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   120
            TabIndex        =   107
            Top             =   2040
            Value           =   1  '�֨�
            Width           =   1425
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H00FFC0FF&
            Caption         =   "��  �}"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Index           =   1
            Left            =   11280
            Picture         =   "frm_OP_TRPPlan.frx":17F2
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   106
            Top             =   120
            Width           =   1065
         End
         Begin VB.TextBox DateS 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1155
            MaxLength       =   8
            TabIndex        =   105
            Top             =   240
            Width           =   1365
         End
         Begin VB.TextBox DateE 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2835
            MaxLength       =   8
            TabIndex        =   104
            Top             =   240
            Width           =   1365
         End
         Begin VB.CommandButton cmd_Query 
            BackColor       =   &H00FFFFC0&
            Caption         =   "��Ƭd��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   8040
            Picture         =   "frm_OP_TRPPlan.frx":1C34
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   103
            Top             =   120
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Excel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�� Excel"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   9120
            Picture         =   "frm_OP_TRPPlan.frx":24FE
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   102
            Top             =   120
            Width           =   1065
         End
         Begin VB.ListBox Storerkey 
            BeginProperty Font 
               Name            =   "�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   720
            Sorted          =   -1  'True
            Style           =   1  '���إ]�t�֨����
            TabIndex        =   101
            Top             =   720
            Width           =   2055
         End
         Begin VB.ListBox Area_Code 
            BeginProperty Font 
               Name            =   "�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   3600
            Style           =   1  '���إ]�t�֨����
            TabIndex        =   100
            Top             =   720
            Width           =   2535
         End
         Begin VB.Label sumlab 
            BeginProperty Font 
               Name            =   "�з���"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   1
            Left            =   1440
            TabIndex        =   114
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Lab_Storerkey 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�f�D"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   12
            Left            =   120
            TabIndex        =   113
            Top             =   720
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "����榡�Gyyyymmdd"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   24
            Left            =   4200
            TabIndex        =   112
            Top             =   360
            Width           =   2010
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�X�����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   25
            Left            =   135
            TabIndex        =   111
            Top             =   300
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   26
            Left            =   2565
            TabIndex        =   110
            Top             =   330
            Width           =   240
         End
         Begin VB.Label Lab_Storerkey 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�ϽX"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   0
            Left            =   3000
            TabIndex        =   109
            Top             =   720
            Width           =   480
         End
      End
      Begin VB.TextBox txtReceipt_No 
         Alignment       =   2  '�m�����
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00404000&
         Height          =   270
         Left            =   -64200
         TabIndex        =   95
         Top             =   3240
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.CommandButton cmd_Tab2_ShowAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "���J�����q��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65040
         Picture         =   "frm_OP_TRPPlan.frx":30C0
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   94
         Top             =   480
         Width           =   1320
      End
      Begin VB.CommandButton cmd_Tab2_FilterAndSort 
         BackColor       =   &H00C0E0FF&
         Caption         =   "�q��j�M"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65040
         Picture         =   "frm_OP_TRPPlan.frx":33CA
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   93
         Top             =   1440
         Width           =   1320
      End
      Begin VB.CommandButton cmd_Tab2_Delete 
         BackColor       =   &H00C0C0FF&
         Caption         =   "TMS�渹�R��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65040
         Picture         =   "frm_OP_TRPPlan.frx":3C94
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   91
         ToolTipText     =   "�R��"
         Top             =   5760
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Frame fma_Tab3_OrderSum 
         Appearance      =   0  '����
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  '�S���ؽu
         ForeColor       =   &H80000008&
         Height          =   6135
         Left            =   -74640
         TabIndex        =   80
         Top             =   600
         Visible         =   0   'False
         Width           =   6360
         Begin VB.CommandButton cmd_Tab3_Query 
            BackColor       =   &H00C0E0FF&
            Caption         =   "�d  ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   3840
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   86
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H00FFC0FF&
            Caption         =   "��  �}"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   2
            Left            =   5040
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   85
            Top             =   675
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab3_Cancel 
            BackColor       =   &H00C0FFFF&
            Caption         =   "��  ��"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   3840
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   84
            Top             =   675
            Width           =   1200
         End
         Begin VB.TextBox txt_Tab3_DeliveryDate_Start 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1020
            TabIndex        =   83
            Top             =   270
            Width           =   1125
         End
         Begin VB.TextBox txt_Tab3_DeliveryDate_End 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2415
            TabIndex        =   82
            Top             =   270
            Width           =   1125
         End
         Begin VB.CommandButton cmd_Tab3_Excel 
            BackColor       =   &H00FFFF80&
            Caption         =   "�� Exccel"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   5040
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   81
            Top             =   195
            Width           =   1200
         End
         Begin MSDataGridLib.DataGrid gd_Tab3_OrderSum 
            Height          =   2025
            Left            =   120
            TabIndex        =   89
            Top             =   1440
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   3572
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "�q��ƶq���R"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid gd_Tab3_Trp02wSum 
            Height          =   2025
            Left            =   120
            TabIndex        =   90
            Top             =   3840
            Width           =   6105
            _ExtentX        =   10769
            _ExtentY        =   3572
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "���ƨ��q����R"
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�e�f���"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   30
            Left            =   120
            TabIndex        =   88
            Top             =   315
            Width           =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   31
            Left            =   2175
            TabIndex        =   87
            Top             =   300
            Width           =   240
         End
      End
      Begin VB.CommandButton cmd_Tab2_Reset 
         Appearance      =   0  '����
         BackColor       =   &H00808080&
         Caption         =   "�q��Ƨ�"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65040
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   72
         Top             =   4800
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CommandButton cmd_Tab2_Remove 
         BackColor       =   &H00C0FFC0&
         Caption         =   "���ܫݱƨ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   -65040
         Picture         =   "frm_OP_TRPPlan.frx":3F9E
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   71
         Top             =   2400
         Width           =   1320
      End
      Begin MSDataGridLib.DataGrid dg_Tab2_ReservedOrders 
         Height          =   6570
         Left            =   -74895
         TabIndex        =   70
         Top             =   390
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   11589
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '����
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   1905
         Left            =   9375
         TabIndex        =   56
         Top             =   390
         Width           =   1995
         Begin VB.CommandButton cmd_Tab1_RouteNoQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "���u�s���d��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   105
            Picture         =   "frm_OP_TRPPlan.frx":42A8
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   58
            Top             =   975
            Width           =   1785
         End
         Begin VB.TextBox txt_Tab1_RouteNo 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   180
            MaxLength       =   10
            TabIndex        =   57
            Top             =   525
            Width           =   1605
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "���u�s��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   240
            Left            =   465
            TabIndex        =   59
            Top             =   195
            Width           =   1020
         End
      End
      Begin VB.Frame fam_RouteData 
         Height          =   585
         Left            =   -74895
         TabIndex        =   2
         Top             =   405
         Width           =   12540
         Begin VB.CommandButton cmd_Tab0_Clear 
            BackColor       =   &H008080FF&
            Caption         =   "�M��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3210
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   78
            Top             =   75
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmd_Tab0_Save 
            BackColor       =   &H00FF8080&
            Caption         =   "�s��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2595
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   77
            Top             =   75
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.CommandButton cmd_Tab0_Query 
            Appearance      =   0  '����
            BackColor       =   &H00808000&
            Caption         =   "�d��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1980
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   76
            Top             =   75
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox txt_Tab0_RouteNo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   540
            TabIndex        =   74
            Top             =   150
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.TextBox txt_Tab0_CarCheckInDate 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7380
            TabIndex        =   67
            Top             =   135
            Width           =   1140
         End
         Begin VB.TextBox txt_Tab0_CarCheckInTime 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9285
            MaxLength       =   4
            TabIndex        =   64
            Top             =   135
            Width           =   750
         End
         Begin VB.TextBox txt_Tab0_DockNo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5475
            TabIndex        =   62
            Top             =   135
            Width           =   1155
         End
         Begin VB.CommandButton cmd_Tab0_SelectedRemove_All 
            BackColor       =   &H000080FF&
            Caption         =   "�w��q�沾��(��)"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3840
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   53
            Top             =   60
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.CommandButton cmd_Tab0_CreateRoute 
            Appearance      =   0  '����
            BackColor       =   &H00FF8080&
            Caption         =   "�إ߸��u�s��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10080
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   52
            Top             =   75
            Width           =   1110
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00400040&
            FillStyle       =   0  '���
            Height          =   540
            Index           =   2
            Left            =   3840
            Top             =   45
            Width           =   1125
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00400040&
            FillStyle       =   0  '���
            Height          =   540
            Index           =   1
            Left            =   1950
            Top             =   45
            Width           =   1860
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   3
            Left            =   45
            Top             =   120
            Width           =   1890
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "���u�s��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   21
            Left            =   105
            TabIndex        =   75
            Top             =   150
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�w�p������"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   20
            Left            =   6720
            TabIndex        =   68
            Top             =   135
            Width           =   675
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   2
            Left            =   6675
            Top             =   105
            Width           =   1875
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   1
            Left            =   8565
            Top             =   105
            Width           =   1500
         End
         Begin VB.Shape Shape5 
            Height          =   450
            Index           =   0
            Left            =   4980
            Top             =   105
            Width           =   1680
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�w�p����ɶ�"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   19
            Left            =   8610
            TabIndex        =   63
            Top             =   135
            Width           =   675
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�X�Y�Ȧs"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   18
            Left            =   5040
            TabIndex        =   61
            Top             =   135
            Width           =   435
         End
      End
      Begin VB.Frame fam_SelectedOrders 
         Height          =   3315
         Left            =   -74880
         TabIndex        =   22
         Top             =   900
         Width           =   12585
         Begin VB.CheckBox chk_Tab0_Updatetrpw 
            Caption         =   "��s"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   405
            Left            =   1200
            TabIndex        =   116
            ToolTipText     =   "���J�P�ɧ�s�O�c����"
            Top             =   150
            Width           =   735
         End
         Begin VB.CommandButton cmd_Tab0_CreateRouteByAds 
            Appearance      =   0  '����
            BackColor       =   &H00FFFF00&
            Caption         =   "  �̦a�}  �ո��s"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11520
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   97
            Top             =   120
            Width           =   990
         End
         Begin VB.CheckBox ck_All 
            BackColor       =   &H80000012&
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   495
            Left            =   12480
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   96
            Top             =   2880
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H008080FF&
            Caption         =   "��  �}"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   10440
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   79
            Top             =   120
            Width           =   990
         End
         Begin VB.CommandButton cmd_Tab0_ImportOrders 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���J�ݱƨ��q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   30
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   73
            Top             =   105
            Width           =   1095
         End
         Begin VB.CommandButton cmd_Tab0_Reserve 
            BackColor       =   &H00FF8080&
            Caption         =   "�O�d�q��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7950
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   69
            Top             =   2880
            Width           =   1185
         End
         Begin VB.CommandButton cmd_Tab0_SelectedCancel 
            BackColor       =   &H00FF8080&
            Caption         =   "�ݿ����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9000
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   24
            Top             =   3000
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.CheckBox chk_Tab0_DriveTimes 
            Caption         =   "��ܨ���"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   405
            Left            =   5625
            TabIndex        =   60
            Top             =   150
            Width           =   750
         End
         Begin VB.CommandButton cmd_Tab0_srcOrderReset 
            Appearance      =   0  '����
            BackColor       =   &H00808080&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   11970
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   55
            Top             =   2910
            Width           =   495
         End
         Begin VB.CommandButton cmd_Tab0_srcOrderQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "�q��j�M"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10845
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   54
            Top             =   2910
            Width           =   1110
         End
         Begin VB.TextBox txt_Tab0_DeliveryPhone 
            Appearance      =   0  '����
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   8325
            TabIndex        =   48
            Top             =   315
            Width           =   1170
         End
         Begin VB.TextBox txt_Tab0_DeliveryCompany 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   6360
            TabIndex        =   46
            Top             =   315
            Width           =   825
         End
         Begin VB.TextBox txt_Tab0_DeliveryDriver 
            Appearance      =   0  '����
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   7170
            TabIndex        =   44
            Top             =   315
            Width           =   1170
         End
         Begin VB.TextBox txt_Tab0_DeliveryCarType 
            Alignment       =   2  '�m�����
            Appearance      =   0  '����
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   285
            Left            =   9480
            TabIndex        =   42
            Top             =   315
            Width           =   945
         End
         Begin VB.CommandButton cmd_Tab0_SelectCar 
            BackColor       =   &H00FFC0C0&
            Caption         =   "�H"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5190
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   41
            Top             =   150
            Width           =   330
         End
         Begin VB.TextBox txt_Tab0_DeliveryCarNo 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   4080
            TabIndex        =   40
            Top             =   150
            Width           =   1125
         End
         Begin VB.TextBox txt_Tab0_TRPDate 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2445
            TabIndex        =   38
            Top             =   150
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab0_Selected 
            BackColor       =   &H00FF8080&
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7095
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   36
            Top             =   2910
            Width           =   345
         End
         Begin VB.CommandButton cmd_Tab0_Remove 
            BackColor       =   &H008080FF&
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7455
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   35
            Top             =   2880
            Width           =   345
         End
         Begin VB.CommandButton cmd_Tab0_SelectedCancel_All 
            BackColor       =   &H00FF80FF&
            Caption         =   "�ݿ����(��)"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   23
            Top             =   2880
            Width           =   1530
         End
         Begin MSDataGridLib.DataGrid dg_Tab0_SelectedOrders 
            Height          =   2265
            Left            =   0
            TabIndex        =   25
            Top             =   600
            Width           =   11130
            _ExtentX        =   19632
            _ExtentY        =   3995
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame1 
            Enabled         =   0   'False
            Height          =   525
            Left            =   0
            TabIndex        =   26
            Top             =   2820
            Width           =   6915
            Begin VB.TextBox txt_Tab0_Selected_OTqty 
               Alignment       =   2  '�m�����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   990
               TabIndex        =   118
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Case 
               Alignment       =   2  '�m�����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   2310
               TabIndex        =   30
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Pallet 
               Alignment       =   2  '�m�����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   3540
               TabIndex        =   29
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Volumn 
               Alignment       =   2  '�m�����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   4785
               TabIndex        =   28
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_Selected_Weight 
               Alignment       =   2  '�m�����
               BackColor       =   &H00FFC0FF&
               ForeColor       =   &H000000C0&
               Height          =   270
               Left            =   6015
               TabIndex        =   27
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�c��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   22
               Left            =   1920
               TabIndex        =   117
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���q"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   4
               Left            =   5640
               TabIndex        =   34
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���n"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   5
               Left            =   4395
               TabIndex        =   33
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�O��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   6
               Left            =   3165
               TabIndex        =   32
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�֭p�G���"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   7
               Left            =   75
               TabIndex        =   31
               Top             =   210
               Width           =   900
            End
         End
         Begin VB.Shape Shape3 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00FF0000&
            FillStyle       =   0  '���
            Height          =   435
            Left            =   10815
            Top             =   2880
            Width           =   2880
         End
         Begin VB.Shape Shape2 
            BackStyle       =   1  '���z��
            BorderColor     =   &H00400040&
            FillStyle       =   0  '���
            Height          =   435
            Index           =   0
            Left            =   7920
            Top             =   2880
            Width           =   2790
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000002&
            BackStyle       =   1  '���z��
            Height          =   435
            Left            =   7050
            Top             =   2880
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�q  ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   8685
            TabIndex        =   49
            Top             =   120
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�B�餽�q"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   6360
            TabIndex        =   47
            Top             =   120
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�r�p�H"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   15
            Left            =   7455
            TabIndex        =   45
            Top             =   120
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��   ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   9630
            TabIndex        =   43
            Top             =   120
            Width           =   600
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "���P���X"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   13
            Left            =   3630
            TabIndex        =   39
            Top             =   165
            Width           =   420
         End
         Begin VB.Label Label1 
            BackStyle       =   0  '�z��
            Caption         =   "�X�����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   12
            Left            =   2010
            TabIndex        =   37
            Top             =   150
            Width           =   435
         End
      End
      Begin VB.Frame fam_SrcOrders 
         Height          =   2955
         Left            =   -74880
         TabIndex        =   1
         Top             =   4200
         Width           =   12540
         Begin VB.Frame Frame2 
            Enabled         =   0   'False
            Height          =   525
            Left            =   6960
            TabIndex        =   13
            Top             =   0
            Width           =   6915
            Begin VB.TextBox txt_Tab0_srcTotal_OTqty 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   990
               TabIndex        =   121
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcTotal_Weight 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   6000
               TabIndex        =   17
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcTotal_Volumn 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   4785
               TabIndex        =   16
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcTotal_Pallet 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   3540
               TabIndex        =   15
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcTotal_Case 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00404000&
               Height          =   270
               Left            =   2295
               TabIndex        =   14
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�c��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   27
               Left            =   1920
               TabIndex        =   122
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�`�p�G���"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   11
               Left            =   75
               TabIndex        =   21
               Top             =   210
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�O��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   10
               Left            =   3165
               TabIndex        =   20
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���n"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   9
               Left            =   4395
               TabIndex        =   19
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���q"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   8
               Left            =   5625
               TabIndex        =   18
               Top             =   210
               Width           =   360
            End
         End
         Begin VB.Frame fam_SelectedSum 
            Enabled         =   0   'False
            Height          =   525
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   6915
            Begin VB.TextBox txt_Tab0_srcSelected_OTqty 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   975
               TabIndex        =   119
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Case 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   2310
               TabIndex        =   7
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Pallet 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   3540
               TabIndex        =   6
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Volumn 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   4785
               TabIndex        =   5
               Top             =   165
               Width           =   840
            End
            Begin VB.TextBox txt_Tab0_srcSelected_Weight 
               Alignment       =   2  '�m�����
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   6015
               TabIndex        =   4
               Top             =   165
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�c��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   23
               Left            =   1920
               TabIndex        =   120
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���q"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   3
               Left            =   5640
               TabIndex        =   11
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "���n"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   2
               Left            =   4395
               TabIndex        =   10
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "�O��"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   1
               Left            =   3165
               TabIndex        =   9
               Top             =   210
               Width           =   360
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  '�z��
               Caption         =   "����G���"
               ForeColor       =   &H00800000&
               Height          =   180
               Index           =   0
               Left            =   75
               TabIndex        =   8
               Top             =   210
               Width           =   900
            End
         End
         Begin MSDataGridLib.DataGrid dg_TRP02W 
            Height          =   2280
            Left            =   45
            TabIndex        =   12
            Top             =   525
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   4022
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1028
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin MSDataGridLib.DataGrid dg_Tab1_RouteOrders 
         Height          =   2640
         Left            =   90
         TabIndex        =   50
         Top             =   4485
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   4657
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dg_Tab1_Route 
         Height          =   4065
         Left            =   90
         TabIndex        =   51
         Top             =   390
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   7170
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  '����
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   2100
         Left            =   9405
         TabIndex        =   65
         Top             =   2355
         Width           =   1980
         Begin VB.CommandButton cmdDeliveryDateFix 
            BackColor       =   &H000080FF&
            Caption         =   "��f�ɶ��w��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   120
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   92
            ToolTipText     =   "�q�榳�Ƶ���"
            Top             =   240
            Width           =   1785
         End
         Begin VB.CommandButton cmd_Tab1_RouteNoDelete 
            BackColor       =   &H00C0C0FF&
            Caption         =   "���u�s���R��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   105
            Picture         =   "frm_OP_TRPPlan.frx":45B2
            Style           =   1  '�Ϥ��~�[
            TabIndex        =   66
            ToolTipText     =   "�R��"
            Top             =   1200
            Width           =   1785
         End
      End
      Begin MSDataGridLib.DataGrid dg_RouteData 
         Height          =   4470
         Left            =   -74850
         TabIndex        =   115
         Top             =   2880
         Width           =   12435
         _ExtentX        =   21934
         _ExtentY        =   7885
         _Version        =   393216
         AllowUpdate     =   0   'False
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1028
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frm_OP_TRPPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form �]�p�ɴ�����
Private dbsrcFormWidth As Double     'Form �]�p�ɴ����e
Private rs_Access As ADODB.Recordset         '����C�L�� >> ���Ʀ� Access DB
Private MSAccessAP As access.Application
Private blTRP02WEventEnable As Boolean              '�ݿ���q�� Event Ĳ�o���ı���
Private blTab0SelectedOrderEventEnable As Boolean   '�w����q�� Event Ĳ�o���ı���
Private blTab1RouteEventEnable As Boolean           '���u�s���C�� Event Ĳ�o���ı���
Private blTab2ReservedEventEnable As Boolean        '�O�d�q��C�� Event Ĳ�o���ı���

Private blRouteModify As Boolean                    '�ƨ��@�~ >> ���u�s�� �d�ߡG���ĸ��u�s��
Private blRouteChange As Boolean                    '�ƨ��@�~ >> ���u�s�� ��Ʋ����ѧO�X��
Private strDispRouteNo As String                    '�ƨ��@�~ >> ���u�s�� �d�ߡG���u�s��

Private rs_TRP02W As ADODB.Recordset                '�ƨ��@�~�G�פJ���ݱƨ��q��
Private rs_Tab0_SelectedOrders As ADODB.Recordset   '�ƨ��@�~�G�w������ݱƨ��q��
Private rs_Tab1_Route As ADODB.Recordset            '���s�C��G���u�s���C��
Private rs_Tab1_RouteOrders As ADODB.Recordset      '���s�C��G���u�s�����ݤ��q��
Private rs_Tab2_ReservedOrders As ADODB.Recordset   '�O�d�q��
Private rs_Tab3_OrderSum As ADODB.Recordset         '�q����R�G�̦a�ϲέp
Private rs_Tab3_Trp02wSum As ADODB.Recordset        '���ƨ��q����R�G�̦a�ϲέp
Private rs_RouteData As ADODB.Recordset        '�@�����u�s���жK


Private strSourceFilter As String        '�ݱƨ��q��z��
Private strSourceOrderBy As String       '�ݱƨ��q��ƧǤ覡
Private dbsrcSelected_OTqty As Double     '�ݱƨ��q��: ������
Private dbsrcSelected_Case As Double     '�ݱƨ��q��: ����c��
Private dbsrcSelected_Pallet As Double   '�ݱƨ��q��: ����O��
Private dbsrcSelected_Volumn As Double   '�ݱƨ��q��: ������n
Private dbsrcSelected_Weight As Double   '�ݱƨ��q��: ������q
Private dbSelectedCount As Double        '����q�浧��
Private DelRecord

Private Sub ck_All_Click()
'On Error GoTo errer_handle
'Dim i As Integer
'rs_TRP02W.MoveFirst
'    For i = 0 To rs_TRP02W.RecordCount - 1
'        Call dg_TRP02W_RowColChange(i, 1)
'        rs_TRP02W.MoveNext
'    Next
'errer_handle:
End Sub

Private Sub cmd_Excel_Click()

    If rs_RouteData Is Nothing Then Exit Sub
    If rs_RouteData.RecordCount = 0 Then Exit Sub
    
    Recordset2Excel "�@�����s���", rs_RouteData

    '..�b���s��EXCEL
    With MyXlsApp
    End With
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
    '���}
    Unload Me
End Sub

Private Sub cmd_Tab0_AddressRoute_Click()

End Sub

Private Sub cmd_Query_Click()
    
    Screen.MousePointer = vbHourglass
    Dim str_Where As String, strTmp As String, i As Integer
    str_Where = ""
    
    '�X�f��
    If (Len(RTrim(DateS.Text)) > 0 And Len(RTrim(DateE.Text)) = 0) Or (Len(RTrim(DateS.Text)) = 0 And Len(RTrim(DateE.Text)) > 0) Then str_Where = str_Where & " and convert(char(8),t1.delivery_date,112) = '" & RTrim(DateS.Text) & RTrim(DateE.Text) & "' "
    If (Len(RTrim(DateS.Text)) > 0 And Len(RTrim(DateE.Text)) > 0) Then str_Where = " and convert(char(8),t1.delivery_date,112)  between '" & RTrim(DateS.Text) & "'and'" & RTrim(DateE.Text) & "' "

    '�f�D
    strTmp = ""
    For i = 0 To Storerkey.ListCount - 1
        If Storerkey.Selected(i) Then
                strTmp = strTmp & "'" & Storerkey.List(i) & "',"
        End If
    Next
    
    If Len(strTmp) > 0 Then str_Where = str_Where & " and o.storerkey in (" & Mid(strTmp, 1, Len(strTmp) - 1) & ")"

    '�ϽX
    strTmp = ""
    For i = 0 To Area_Code.ListCount - 1
        If Area_Code.Selected(i) Then
                strTmp = strTmp & "'" & Area_Code.List(i) & "',"
        End If
    Next
    
    If Len(strTmp) > 0 Then str_Where = str_Where & " and t1m.area_code in (" & Mid(strTmp, 1, Len(strTmp) - 1) & ")"
    '20160810 �ק�O�ƵL����i��
    str_SQL = "select " & _
            "�@���r�p�H=isnull(t9.driver,''), " & _
            "�X����=convert(char(8),t1.delivery_date,112), " & _
            "�@�����u�s��=t2.route_no, " & _
            "�Ȥ�²��=t1m.short_name, " & _
            " case when (ceiling(sum(t2.pallet_qty)))>=1 then ceiling(sum(t2.pallet_qty)) else 1 end as �O��, " & _
            "�c��=sum(t2.case_cnt) " & _
            "from orders o join trp02t t2 on o.orderkey = t2.c_receipt_no " & _
            "join trp01t t1 on t1.route_no = t2.route_no " & _
            "join trp09m t9 on t9.vehicle_id_no = t2.vehicle_id_no " & _
            "left join trp01m t1m on o.storerkey = t1m.storerkey and t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else o.consigneekey end " & _
            "where left(rtrim(t1.route_no),1)='F' and 1=1 " & str_Where & _
            "group by isnull(t9.driver,''),convert(char(8),t1.delivery_date,112),t2.route_no,t1m.short_name " & _
            "order by t2.route_no"
      
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
       Screen.MousePointer = vbDefault
       msg_text = "�d�ߵ��G�G�L�ŦX�����u�s�����"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Set dg_RouteData.DataSource = Nothing
       Exit Sub
    End If
    
    Call ReDim_Recordset(rs_RouteData)
    Call Replication_Recordset(tmp_Rs, rs_RouteData)
    tmp_Rs.Close
    
    With dg_RouteData
         .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
         .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
         .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
         .RowHeight = 270                '�]�wDataGrid ������Ҧ���ƦC����
    End With
    rs_RouteData.MoveFirst
    Set dg_RouteData.DataSource = rs_RouteData
    With dg_RouteData
        .RowHeight = 250
        .Columns(0).Width = 500        '�Ǹ�
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1200        '�@���r�p�H
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 1500        '�X�f��
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 1200        '���u�s��
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 2000       '�Ȥ�W��
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 1000       '�O��
        .Columns(5).Alignment = dbgRight
        .Columns(6).Width = 1000        '�c��
        .Columns(6).Alignment = dbgRight
    End With
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Clear_Click()
    '�ƨ��@�~ >> �M��
    If rs_Tab0_SelectedOrders.RecordCount <> 0 And blRouteModify = False Then
        '�s�W���u�s���Ҧ��G
        '�I�s [�w��q�沾��(��)] �ӳB�z�w�Q�Ȯɿ���� [�ݱƨ��q��] �٭�^ [�ݱƨ��q��]
        Call cmd_Tab0_SelectedRemove_All_Click
    End If
    If blRouteModify And blRouteChange Then
       '���ĸ��u�s�� & ��Ƥw�D���ʡA�n user �T�{�O�_�s��
        msg_text = "���u�s����ƬO�_�s�ɡH"
        If MsgBox(msg_text, vbOKCancel + vbCritical, msg_title) = vbOK Then
            '�I�s�s�ɵ{��
            Call cmd_Tab0_Save_Click
        Else
            '���s�ɡ��������s���J [�ݱƨ��q��] �w�٭� [���][����] �ާ@�� [�ݱƨ��q��] ���v�T
            Call cmd_Tab0_ImportOrders_Click
        End If
    End If
    '�M�����u�s�����ȡA�]�t�w��q��W�ӦC��
    Call Clear_RouteData
    txt_Tab0_RouteNo.Text = ""
End Sub

Private Sub cmd_Tab0_CreateRoute_Click()
    '�ƨ��@�~ >> �إ߸��u�s��
    If rs_Tab0_SelectedOrders.RecordCount = 0 Then
        msg_text = "��ƿ��~�G�L�˸����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    
    'add by Terry 20190614
    Dim str_CheckReceipt_No As String
    str_CheckReceipt_No = ""
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        str_CheckReceipt_No = str_CheckReceipt_No & "'" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "',"
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    
    str_CheckReceipt_No = str_CheckReceipt_No & "''"
    rs_Tab0_SelectedOrders.MoveFirst
    str_SQL = "select receipt_no  from trp02t where receipt_no in (" & str_CheckReceipt_No & ")"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        MsgBox ("���q��w�զ��@�����s�A�Э��s���J�ݱƨ��q��òM��[�w������@���q��]"), vbOKOnly + vbCritical
        tmp_Rs.Close
        Exit Sub
    End If
    tmp_Rs.Close
    
    
    
    '�ˮָ��u�s����ƬO�_���T�A���~�N�b Function ������� MessageBox
    If RouteData_Check() = False Then Exit Sub
    
    On Error GoTo err_Handle
    
    cmd_Tab0_CreateRoute.Enabled = False
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    '�ˬd�i�����q
    Dim intableWT, intableCBM As Long
    str_SQL = "select rtrim(isnull(LOADING_SIZE,0)),rtrim(isnull(MAX_CUBIC_CAPACITY,0)) from dbo.TRP09M where VEHICLE_ID_NO='" & Trim(txt_Tab0_DeliveryCarNo.Text) & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    intableWT = tmp_Rs.Fields(0).Value
    intableCBM = tmp_Rs.Fields(1).Value
    tmp_Rs.Close
    If intableWT < Val(txt_Tab0_Selected_Weight.Text) Then
        msg_text = "�ƨ����q�W�L�����i����,�����i����:" & intableWT
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        cmd_Tab0_CreateRoute.Enabled = True
        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Sub
    End If
    If intableCBM < Val(txt_Tab0_Selected_Volumn.Text) Then
        msg_text = "�ƨ����q�W�L�����i�����n,�����i�����n:" & intableCBM
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        cmd_Tab0_CreateRoute.Enabled = True
        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Sub
    End If
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    Dim intDriveTimes As Integer    '����
    Dim strRouteNo As String        '���u�s��
    
    '1.���ͨ���
    str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
              "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_Tab0_TRPDate.Text & "' and Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
    tmp_Rs.Close
    
    '2.���͸��u�s��
    str_SQL = "Select Isnull(Max(Cast(Right(Route_No,3) as integer))+1,1) as RouteSN " & _
              "From TRP01T Where Substring(Route_No,2,6)='" & Mid(txt_Tab0_TRPDate.Text, 3, 6) & "' and Left(Route_No,1) = 'F'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    strRouteNo = "F" & Mid(txt_Tab0_TRPDate, 3, 6) & Format(tmp_Rs.Fields("RouteSN").Value, "000")
    tmp_Rs.Close
    
    '3.Insert into TRP01T ���u�s���D��
    '  TRP01T.EXE_CONFIRM = '0' �s���͸��u�s���A�|���^�ǹL exe
    str_SQL = "Insert into TRP01T (Route_No,Delivery_Date,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Description,EXE_Confirm,AddWho) Values ('" & _
              strRouteNo & "','" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "'," & _
              txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'','0','" & User_id & "')"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '4.insert into TRP05T �����i�X�޲z
    str_SQL = "Insert into TRP05T (Route_No,Vehicle_ID_No,Drive_Times,Delivery_Date,Valid_Vehicle,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Dock_No,Expect_Time,Expect_Date) Values ('" & _
              strRouteNo & "','" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",'" & _
              Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "','1'," & _
              txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'" & _
              txt_Tab0_DockNo.Text & "','" & txt_Tab0_CarCheckInTime.Text & "','" & txt_Tab0_CarCheckInDate.Text & "')"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�Ѩ����D�ɧ�s�����������
    str_SQL = "Update TRP05T Set Driver = B.Driver , Driver_Phone = B.Driver_Phone, TRP_Company_Code = B.TRP_Company_Code " & _
              "From TRP05T A , TRP09M B Where a.Vehicle_ID_No = b.Vehicle_ID_No and a.Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "' and a.Route_No = '" & strRouteNo & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�g�� SSTab1.Tab 1 [���u�s���C��]
    blTab1RouteEventEnable = False
    rs_Tab1_Route.AddNew
    rs_Tab1_Route.Fields("�s��").Value = rs_Tab1_Route.RecordCount
    rs_Tab1_Route.Fields("���u�s��").Value = strRouteNo
    rs_Tab1_Route.Fields("�X�����").Value = txt_Tab0_TRPDate.Text
    rs_Tab1_Route.Fields("���P���X").Value = txt_Tab0_DeliveryCarNo.Text
    rs_Tab1_Route.Fields("����").Value = intDriveTimes
    rs_Tab1_Route.Fields("�r�p�H").Value = txt_Tab0_DeliveryDriver.Text
    rs_Tab1_Route.Fields("���").Value = txt_Tab0_Selected_OTqty.Text
    rs_Tab1_Route.Fields("�c��").Value = txt_Tab0_Selected_Case.Text
    rs_Tab1_Route.Fields("�O��").Value = txt_Tab0_Selected_Pallet.Text
    rs_Tab1_Route.Fields("���n").Value = txt_Tab0_Selected_Volumn.Text
    rs_Tab1_Route.Fields("���q").Value = txt_Tab0_Selected_Weight.Text
    rs_Tab1_Route.Fields("����").Value = txt_Tab0_DeliveryCarType.Text
    rs_Tab1_Route.Fields("�X�Y�Ȧs").Value = txt_Tab0_DockNo.Text
    rs_Tab1_Route.Fields("�w�p������").Value = txt_Tab0_CarCheckInDate.Text
    rs_Tab1_Route.Fields("�w�p����ɶ�").Value = txt_Tab0_CarCheckInTime.Text
    rs_Tab1_Route.Fields("EXE�^��").Value = "�s�ظ��s"
    rs_Tab1_Route.Fields("�ƨ���").Value = User_id
    rs_Tab1_Route.Update
    blTab1RouteEventEnable = True
    
    '5.insert into TRP02T [�ƨ��q����]
    '  �g�� SSTab1.Tab 1 [���u�s�����q��W�Ӫ�]
    blTab0SelectedOrderEventEnable = False
    rs_Tab1_RouteOrders.Filter = adFilterNone
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        'insert into TRP02T
        str_SQL = "Insert into TRP02T (Route_No,StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                  " Vehicle_ID_No,Drive_Times,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,otqty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                  "Select '" & strRouteNo & "',StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                  " '" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,otqty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                  "From TRP02W Where Receipt_No = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
'        '��Orders��s��ơA���}�l�ϥΦ]�����i��J����
'        str_SQL = "update trp02t set trp02t.otqty = orders.otqty from trp02t join orders on trp02t.receipt_no = orders.orderkey and trp02t.receipt_no = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '�g�� SSTab1.Tab 1 [���u�s�����q��W�Ӫ�]
        rs_Tab1_RouteOrders.AddNew
        rs_Tab1_RouteOrders.Fields("�s��").Value = rs_Tab1_RouteOrders.RecordCount
        rs_Tab1_RouteOrders.Fields("���u�s��").Value = strRouteNo
        rs_Tab1_RouteOrders.Fields("�e�f��").Value = rs_Tab0_SelectedOrders.Fields("�e�f��").Value
        rs_Tab1_RouteOrders.Fields("�q��s��").Value = rs_Tab0_SelectedOrders.Fields("�q��s��").Value
        rs_Tab1_RouteOrders.Fields("ZIP").Value = rs_Tab0_SelectedOrders.Fields("ZIP").Value
        rs_Tab1_RouteOrders.Fields("�Ȥ�W��").Value = rs_Tab0_SelectedOrders.Fields("�Ȥ�W��").Value
        rs_Tab1_RouteOrders.Fields("���").Value = rs_Tab0_SelectedOrders.Fields("���").Value
        rs_Tab1_RouteOrders.Fields("�c��").Value = rs_Tab0_SelectedOrders.Fields("�c��").Value
        rs_Tab1_RouteOrders.Fields("�O��").Value = rs_Tab0_SelectedOrders.Fields("�O��").Value
        rs_Tab1_RouteOrders.Fields("���n").Value = rs_Tab0_SelectedOrders.Fields("���n").Value
        rs_Tab1_RouteOrders.Fields("���q").Value = rs_Tab0_SelectedOrders.Fields("���q").Value
        rs_Tab1_RouteOrders.Fields("�q��Ƶ�").Value = rs_Tab0_SelectedOrders.Fields("�q��Ƶ�").Value
        rs_Tab1_RouteOrders.Fields("����").Value = rs_Tab0_SelectedOrders.Fields("����").Value
        rs_Tab1_RouteOrders.Fields("�S��ݨD1").Value = rs_Tab0_SelectedOrders.Fields("�S��ݨD1").Value
        rs_Tab1_RouteOrders.Fields("�S��ݨD2").Value = rs_Tab0_SelectedOrders.Fields("�S��ݨD2").Value
        rs_Tab1_RouteOrders.Fields("Receipt_No").Value = rs_Tab0_SelectedOrders.Fields("Receipt_No").Value
        rs_Tab1_RouteOrders.Fields("EXE�^��").Value = rs_Tab0_SelectedOrders.Fields("EXE�^��").Value
        rs_Tab1_RouteOrders.Fields("Area").Value = rs_Tab0_SelectedOrders.Fields("Area").Value
        rs_Tab1_RouteOrders.Fields("���A").Value = rs_Tab0_SelectedOrders.Fields("���A").Value
        rs_Tab1_RouteOrders.Fields("�Ȥ�²��").Value = rs_Tab0_SelectedOrders.Fields("�Ȥ�²��").Value
        rs_Tab1_RouteOrders.Update
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    
    '�T�{���u�s���� exe_confirm ���A
    '�D�n�ت��G�w�^�Ǥ����s�R����A���s���ͤ����s�A�Y�������O�w�^�ǭq��A�������s�]�w�� [�w�^��]
'Mark by Gemini @20111010
'    str_SQL = "Update TRP01T Set EXE_Confirm = (Select min(EXE_Confirm) From TRP02T Where TRP02T.Route_No = TRP01T.Route_No) " & _
'              "Where TRP01T.Route_No = '" & strRouteNo & "'"
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    cn.CommitTrans
    Tran_Level = 0
    
    If dg_Tab1_Route.SelBookmarks.Count > 0 Then
        dg_Tab1_Route.SelBookmarks.Remove 0
    End If
    dg_Tab1_Route.SelBookmarks.Add rs_Tab1_Route.Bookmark
    rs_Tab1_RouteOrders.Filter = " ���u�s�� = '" & strRouteNo & "' "
    blTab0SelectedOrderEventEnable = True
    
    '4.�� TRP02T Trigger [insert] �i��H�U�@�~
    '   a.�g�J TRP03T -- �ƨ��q�������
    '   b.�R�� TRP03W -- �ݱƨ��q�������
    '   c.�R�� TRP02W -- �ݱƨ��q��D��
    
    
    '5.�M�� [�w������ݱƨ��q��C��]
    blTab0SelectedOrderEventEnable = False
    '�ƨ��@�~�G�w������ݱƨ��q��C�� DBGrid �榡�]�w-ReSet
    Call CreateRS_Tab0_SelectedOrders
    '���s�p��w����q��G�c�ơA�O�ơA���n�A���q + �s�����s����
    Call Calculate_SelectedOrders
    blTab0SelectedOrderEventEnable = True
    
    '6.�M���ƨ��@�~����
    txt_Tab0_DockNo.Text = ""               '�X�Y�Ȧs
    txt_Tab0_CarCheckInDate.Text = ""       '�����w�p������
    txt_Tab0_CarCheckInTime.Text = ""       '�����w�p����ɶ�
    txt_Tab0_TRPDate.Text = ""              '�X�����
    txt_Tab0_DeliveryCarNo.Text = ""        '���P���X
    txt_Tab0_DeliveryCompany.Text = ""      '�B�餽�q
    txt_Tab0_DeliveryDriver.Text = ""       '�r�p�H
    txt_Tab0_DeliveryPhone.Text = ""        '�q��
    txt_Tab0_DeliveryCarType.Text = ""      '����
    
    cmd_Tab0_CreateRoute.Enabled = True
    
    'Call cmd_Tab0_ImportOrders_Click 'edit by Eric 20140729�A�קK���ƿz��S�w�q��A�����ݨD
    
    '�ݱƨ��q���`�p��T
    Call ReCaculate_OrderSum
    
    SSTab1.Tab = 1
    DoEvents: DoEvents
    '�w�ƨ�f�ɶ�
    Call cmdDeliveryDateFix_Click
    
    
On Error GoTo err_Handle2
    'Terry 20200212 �ƨ������JBestAPP Ĳ�o�����\�� �L�״��ϥ�
    cn.Execute "exec Andys_BestTMSOrderImport", RowsAffect, adExecuteNoRecords
    Dim HttpClient As Object

    Set HttpClient = CreateObject("Microsoft.XMLHTTP")
    HttpClient.Open "POST", "https://entrance-bestlog.azurewebsites.net/api/BestApp/BestAppTMS/InsertWaybillList", False
    HttpClient.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    HttpClient.Send
    
    
    Exit Sub

err_Handle2:
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
        Tran_Level = 0
        cn.RollbackTrans
   End If
   
   '�D�J���~���ܡGlocal �� Recordset [���u�s���C��] ��ƥ����R��
   '�]�� [���u�s���C��] ���� DB connection.transaction ����
   blTab1RouteEventEnable = False
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   rs_Tab1_Route.Filter = "���u�s��='" & strRouteNo & "'"
   If Not rs_Tab1_Route.EOF Then
      rs_Tab1_Route.Delete
   End If
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   blTab1RouteEventEnable = True
   
   rs_Tab1_RouteOrders.Filter = adFilterNone
   rs_Tab1_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   rs_Tab1_RouteOrders.Filter = "���u�s��='" & strRouteNo & "'"
   If Not rs_Tab1_RouteOrders.EOF Then
      Do While Not rs_Tab1_RouteOrders.EOF
         rs_Tab1_RouteOrders.Delete
         rs_Tab1_RouteOrders.MoveFirst
      Loop
   End If
   rs_Tab1_RouteOrders.Filter = adFilterNone
   rs_Tab1_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
      
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��@�~-�إ߸��u�s��", Me.Caption, "cmd_Tab0_CreateRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   cmd_Tab0_CreateRoute.Enabled = True
End Sub

Private Sub cmd_Tab0_CreateRouteByAds_Click()
   '�ƨ��@�~ >> �إ߸��u�s��
    If rs_Tab0_SelectedOrders.RecordCount = 0 Then
        msg_text = "��ƿ��~�G�L�˸����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    
    '�ˮָ��u�s����ƬO�_���T�A���~�N�b Function ������� MessageBox
    If RouteData_Check() = False Then Exit Sub
    
    On Error GoTo err_Handle
    
    cmd_Tab0_CreateRouteByAds.Enabled = False
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    '�ˬd�i�����q,��ƨ����ˬd
'    Dim intableWT, intableCBM As Long
'    str_SQL = "select rtrim(isnull(LOADING_SIZE,0)),rtrim(isnull(MAX_CUBIC_CAPACITY,0)) from dbo.TRP09M where VEHICLE_ID_NO='" & Trim(txt_Tab0_DeliveryCarNo.Text) & "'"
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    intableWT = tmp_Rs.Fields(0).Value
'    intableCBM = tmp_Rs.Fields(1).Value
'    tmp_Rs.Close
'    If intableWT < Val(txt_Tab0_Selected_Weight.Text) Then
'        msg_text = "�ƨ����q�W�L�����i����,�����i����:" & intableWT
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        cmd_Tab0_CreateRoute.Enabled = True
'        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
'        txt_Tab0_DeliveryCarNo.SetFocus
'        Exit Sub
'    End If
'    If intableCBM < Val(txt_Tab0_Selected_Volumn.Text) Then
'        msg_text = "�ƨ����q�W�L�����i�����n,�����i�����n:" & intableCBM
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        cmd_Tab0_CreateRoute.Enabled = True
'        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
'        txt_Tab0_DeliveryCarNo.SetFocus
'        Exit Sub
'    End If
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    Dim intDriveTimes As Integer    '����
    Dim strRouteNo As String        '���u�s��
    Dim strAddress As String        '���P�a�}���ͷs�����u�s��
    Dim strRouteNosum As String     '��sTRP01�BTRP05
    strAddress = ""
    rs_Tab1_RouteOrders.Filter = adFilterNone
    rs_Tab0_SelectedOrders.Sort = "Zip,�B�e�a�}"
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        
        If Trim(strAddress) <> Trim(rs_Tab0_SelectedOrders.Fields("�B�e�a�}").Value) Then '�a�}���@��
            
            strAddress = Trim(rs_Tab0_SelectedOrders.Fields("�B�e�a�}").Value)
            '1.���ͨ���
            str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
                      "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_Tab0_TRPDate.Text & "' and Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "'"
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
            tmp_Rs.Close
            
            '2.���͸��u�s��
            str_SQL = "Select Isnull(Max(Cast(Right(Route_No,3) as integer))+1,1) as RouteSN " & _
                      "From TRP01T Where Substring(Route_No,2,6)='" & Mid(txt_Tab0_TRPDate.Text, 3, 6) & "' and Left(Route_No,1) = 'F'"
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            strRouteNo = "F" & Mid(txt_Tab0_TRPDate, 3, 6) & Format(tmp_Rs.Fields("RouteSN").Value, "000")
            tmp_Rs.Close
            
            '�����������ͪ�����
            If Len(strRouteNosum) = 0 Then strRouteNosum = "'" & strRouteNo & "'" Else strRouteNosum = strRouteNosum & ",'" & strRouteNo & "'"
            
            '3.Insert into TRP01T ���u�s���D��
            '  TRP01T.EXE_CONFIRM = '0' �s���͸��u�s���A�|���^�ǹL exe
            str_SQL = "Insert into TRP01T (Route_No,Delivery_Date,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Description,EXE_Confirm,AddWho) Values ('" & _
                      strRouteNo & "','" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "'," & _
                      txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'','0','" & User_id & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '4.insert into TRP05T �����i�X�޲z
            str_SQL = "Insert into TRP05T (Route_No,Vehicle_ID_No,Drive_Times,Delivery_Date,Valid_Vehicle,Case_cnt,Pallet_Qty,Weight,Volumn_Weight,Dock_No,Expect_Time,Expect_Date) Values ('" & _
                      strRouteNo & "','" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",'" & _
                      Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "','1'," & _
                      txt_Tab0_Selected_Case.Text & "," & txt_Tab0_Selected_Pallet.Text & "," & txt_Tab0_Selected_Weight.Text & "," & txt_Tab0_Selected_Volumn.Text & ",'" & _
                      txt_Tab0_DockNo.Text & "','" & txt_Tab0_CarCheckInTime.Text & "','" & txt_Tab0_CarCheckInDate.Text & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '�Ѩ����D�ɧ�s�����������
            str_SQL = "Update TRP05T Set Driver = B.Driver , Driver_Phone = B.Driver_Phone, TRP_Company_Code = B.TRP_Company_Code " & _
                      "From TRP05T A , TRP09M B Where a.Vehicle_ID_No = b.Vehicle_ID_No and a.Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "' and a.Route_No = '" & strRouteNo & "' "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
        End If
        
     '5.insert into TRP02T [�ƨ��q����]
     '  �g�� SSTab1.Tab 1 [���u�s�����q��W�Ӫ�]
       blTab0SelectedOrderEventEnable = False
        'insert into TRP02T
        str_SQL = "Insert into TRP02T (Route_No,StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                  " Vehicle_ID_No,Drive_Times,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,otqty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                  "Select '" & strRouteNo & "',StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                  " '" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,otqty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                  "From TRP02W Where Receipt_No = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    
    '�T�{���u�s���� exe_confirm ���A
    '�D�n�ت��G�w�^�Ǥ����s�R����A���s���ͤ����s�A�Y�������O�w�^�ǭq��A�������s�]�w�� [�w�^��]
    
    '6. update trp01t,trp05t�A
    str_SQL = "update TRP01T set WEIGHT=(select sum(TRP02T.WEIGHT) from TRP02T where TRP02T.ROUTE_NO=TRP01T.ROUTE_NO) " & _
        "where   route_no in ( " & strRouteNosum & ") " & _
        "update TRP01T set CASE_CNT=(select sum(TRP02T.CASE_CNT) from TRP02T where TRP02T.ROUTE_NO=TRP01T.ROUTE_NO) " & _
        "where  route_no in ( " & strRouteNosum & ") " & _
        "update TRP01T set Pallet_Qty=(select sum(TRP02T.Pallet_Qty) from TRP02T where TRP02T.ROUTE_NO=TRP01T.ROUTE_NO) " & _
        "where route_no in ( " & strRouteNosum & ") " & _
        "update TRP01T set VOLUMN_WEIGHT=(select sum(TRP02T.VOLUMN_WEIGHT) from TRP02T where TRP02T.ROUTE_NO=TRP01T.ROUTE_NO) " & _
        "where route_no in ( " & strRouteNosum & ") "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "update TRP05T set WEIGHT=(select sum(TRP02T.WEIGHT) from TRP02T where TRP02T.ROUTE_NO=TRP05T.ROUTE_NO) " & _
        "where   route_no in ( " & strRouteNosum & ") " & _
        "update TRP05T set CASE_CNT=(select sum(TRP02T.CASE_CNT) from TRP02T where TRP02T.ROUTE_NO=TRP05T.ROUTE_NO) " & _
        "where  route_no in ( " & strRouteNosum & ") " & _
        "update TRP05T set Pallet_Qty=(select sum(TRP02T.Pallet_Qty) from TRP02T where TRP02T.ROUTE_NO=TRP05T.ROUTE_NO) " & _
        "where route_no in ( " & strRouteNosum & ") " & _
        "update TRP05T set VOLUMN_WEIGHT=(select sum(TRP02T.VOLUMN_WEIGHT) from TRP02T where TRP02T.ROUTE_NO=TRP05T.ROUTE_NO) " & _
        "where route_no in ( " & strRouteNosum & ") "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    
    cn.CommitTrans
    Tran_Level = 0
    
    If dg_Tab1_Route.SelBookmarks.Count > 0 Then
        dg_Tab1_Route.SelBookmarks.Remove 0
    End If
'    dg_Tab1_Route.SelBookmarks.Add rs_Tab1_Route.Bookmark
'    rs_Tab1_RouteOrders.Filter = " ���u�s�� = '" & strRouteNo & "' "
    blTab0SelectedOrderEventEnable = True
    
    '7.�� TRP02T Trigger [insert] �i��H�U�@�~
    '   a.�g�J TRP03T -- �ƨ��q�������
    '   b.�R�� TRP03W -- �ݱƨ��q�������
    '   c.�R�� TRP02W -- �ݱƨ��q��D��
    
    '8.�M�� [�w������ݱƨ��q��C��]
    blTab0SelectedOrderEventEnable = False
    '�ƨ��@�~�G�w������ݱƨ��q��C�� DBGrid �榡�]�w-ReSet
    Call CreateRS_Tab0_SelectedOrders
    '���s�p��w����q��G�c�ơA�O�ơA���n�A���q + �s�����s����
    Call Calculate_SelectedOrders
    blTab0SelectedOrderEventEnable = True
    
    '6.�M���ƨ��@�~����
    txt_Tab0_DockNo.Text = ""               '�X�Y�Ȧs
    txt_Tab0_CarCheckInDate.Text = ""       '�����w�p������
    txt_Tab0_CarCheckInTime.Text = ""       '�����w�p����ɶ�
    txt_Tab0_TRPDate.Text = ""              '�X�����
    txt_Tab0_DeliveryCarNo.Text = ""        '���P���X
    txt_Tab0_DeliveryCompany.Text = ""      '�B�餽�q
    txt_Tab0_DeliveryDriver.Text = ""       '�r�p�H
    txt_Tab0_DeliveryPhone.Text = ""        '�q��
    txt_Tab0_DeliveryCarType.Text = ""      '����
    
    cmd_Tab0_CreateRouteByAds.Enabled = True
    
    'Call cmd_Tab0_ImportOrders_Click
    
    '�ݱƨ��q���`�p��T
    Call ReCaculate_OrderSum
    
    SSTab1.Tab = 1
    DoEvents: DoEvents
    '�w�ƨ�f�ɶ�
    Call cmdDeliveryDateFix_Click
    
    '�d�߱ƨ����G
    '�]�w���u�s���C��
    blTab1RouteEventEnable = False
    Call CreateRS_Tab1_Route
    blTab1RouteEventEnable = True
    '�]�w���u�s�����q��C��
    Call CreateRS_Tab1_RouteOrders
    
    str_SQL = "Select ���u�s��,�X�����,���P���X,����,�r�p�H,���,�c��,�O��,���n,���q,����,�X�Y�Ȧs,�w�p������,�w�p����ɶ�,EXE�^��,�ƨ��� " & _
              "From TRPPlan_RouteData Where ���u�s�� in ( " & strRouteNosum & ") order by ���u�s��"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧸��u�s�����(TRP01T)"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    blTab1RouteEventEnable = False
    Do While Not tmp_Rs.EOF
        rs_Tab1_Route.AddNew
        rs_Tab1_Route.Fields("�s��").Value = rs_Tab1_Route.RecordCount
        rs_Tab1_Route.Fields("���u�s��").Value = tmp_Rs.Fields("���u�s��").Value
        rs_Tab1_Route.Fields("�X�����").Value = tmp_Rs.Fields("�X�����").Value
        rs_Tab1_Route.Fields("���P���X").Value = tmp_Rs.Fields("���P���X").Value
        rs_Tab1_Route.Fields("����").Value = tmp_Rs.Fields("����").Value
        rs_Tab1_Route.Fields("�r�p�H").Value = tmp_Rs.Fields("�r�p�H").Value
        rs_Tab1_Route.Fields("���").Value = tmp_Rs.Fields("���").Value
        rs_Tab1_Route.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
        rs_Tab1_Route.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
        rs_Tab1_Route.Fields("���n").Value = tmp_Rs.Fields("���n").Value
        rs_Tab1_Route.Fields("���q").Value = tmp_Rs.Fields("���q").Value
        rs_Tab1_Route.Fields("����").Value = tmp_Rs.Fields("����").Value
        rs_Tab1_Route.Fields("�X�Y�Ȧs").Value = tmp_Rs.Fields("�X�Y�Ȧs").Value
        rs_Tab1_Route.Fields("�w�p������").Value = tmp_Rs.Fields("�w�p������").Value
        rs_Tab1_Route.Fields("�w�p����ɶ�").Value = tmp_Rs.Fields("�w�p����ɶ�").Value
        rs_Tab1_Route.Fields("EXE�^��").Value = tmp_Rs.Fields("EXE�^��").Value
        rs_Tab1_Route.Fields("�ƨ���").Value = tmp_Rs.Fields("�ƨ���").Value
        rs_Tab1_Route.Update
        tmp_Rs.MoveNext
    Loop
    rs_Tab1_Route.MoveFirst
    blTab1RouteEventEnable = True
    tmp_Rs.Close
    'TRP03W
    str_SQL = "Select ���u�s��,�e�f��,�q��s��,ZIP,�Ȥ�W��,�Ȥ�a�},���,�c��,�O��,���n,���q,Receipt_No,EXE�^��,Area,���A,�Ȥ�²��,�q��Ƶ� " & _
              "From TRPPlan_RouteOrders " & _
               "Where ���u�s�� in ( " & strRouteNosum & ") Order by ���u�s��,Receipt_No"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���u�s�����q����(TRP02T)"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Do While Not tmp_Rs.EOF
        rs_Tab1_RouteOrders.AddNew
        rs_Tab1_RouteOrders.Fields("�s��").Value = rs_Tab1_RouteOrders.RecordCount
        rs_Tab1_RouteOrders.Fields("���u�s��").Value = tmp_Rs.Fields("���u�s��").Value
        rs_Tab1_RouteOrders.Fields("�e�f��").Value = tmp_Rs.Fields("�e�f��").Value
        rs_Tab1_RouteOrders.Fields("�q��s��").Value = tmp_Rs.Fields("�q��s��").Value
        rs_Tab1_RouteOrders.Fields("ZIP").Value = tmp_Rs.Fields("ZIP").Value
        rs_Tab1_RouteOrders.Fields("�Ȥ�W��").Value = tmp_Rs.Fields("�Ȥ�W��").Value
        rs_Tab1_RouteOrders.Fields("�a�}").Value = tmp_Rs.Fields("�Ȥ�a�}").Value
        rs_Tab1_RouteOrders.Fields("���").Value = tmp_Rs.Fields("���").Value
        rs_Tab1_RouteOrders.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
        rs_Tab1_RouteOrders.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
        rs_Tab1_RouteOrders.Fields("���n").Value = tmp_Rs.Fields("���n").Value
        rs_Tab1_RouteOrders.Fields("���q").Value = tmp_Rs.Fields("���q").Value
        rs_Tab1_RouteOrders.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
        rs_Tab1_RouteOrders.Fields("EXE�^��").Value = tmp_Rs.Fields("EXE�^��").Value
        rs_Tab1_RouteOrders.Fields("Area").Value = tmp_Rs.Fields("Area").Value
        rs_Tab1_RouteOrders.Fields("���A").Value = tmp_Rs.Fields("���A").Value
        rs_Tab1_RouteOrders.Fields("�Ȥ�²��").Value = tmp_Rs.Fields("�Ȥ�²��").Value
        rs_Tab1_RouteOrders.Fields("�q��Ƶ�").Value = tmp_Rs.Fields("�q��Ƶ�").Value
        rs_Tab1_RouteOrders.Update
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    rs_Tab1_RouteOrders.MoveFirst
    
    Screen.MousePointer = vbDefault
    
    
On Error GoTo err_Handle2
    'Terry 20200212 �ƨ������JBestAPP Ĳ�o�����\�� �L�״��ϥ�
    cn.Execute "exec Andys_BestTMSOrderImport", RowsAffect, adExecuteNoRecords
    Dim HttpClient As Object

    Set HttpClient = CreateObject("Microsoft.XMLHTTP")
    HttpClient.Open "POST", "https://entrance-bestlog.azurewebsites.net/api/BestApp/BestAppTMS/InsertWaybillList", False
    HttpClient.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    HttpClient.Send
    
    Exit Sub
    
err_Handle2:
Exit Sub
    
err_Handle:
   If Tran_Level <> 0 Then
        Tran_Level = 0
        cn.RollbackTrans
   End If
   
   '�D�J���~���ܡGlocal �� Recordset [���u�s���C��] ��ƥ����R��
   '�]�� [���u�s���C��] ���� DB connection.transaction ����
   blTab1RouteEventEnable = False
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   rs_Tab1_Route.Filter = "���u�s��='" & strRouteNo & "'"
   If Not rs_Tab1_Route.EOF Then
      rs_Tab1_Route.Delete
   End If
   rs_Tab1_Route.Filter = adFilterNone
   rs_Tab1_Route.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   blTab1RouteEventEnable = True
   
   rs_Tab1_RouteOrders.Filter = adFilterNone
   rs_Tab1_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
   rs_Tab1_RouteOrders.Filter = "���u�s��='" & strRouteNo & "'"
   If Not rs_Tab1_RouteOrders.EOF Then
      Do While Not rs_Tab1_RouteOrders.EOF
         rs_Tab1_RouteOrders.Delete
         rs_Tab1_RouteOrders.MoveFirst
      Loop
   End If
   rs_Tab1_RouteOrders.Filter = adFilterNone
   rs_Tab1_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
      
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��@�~-�̦a�}�إ߸��u�s��", Me.Caption, "cmd_Tab0_CreateRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   cmd_Tab0_CreateRouteByAds.Enabled = True
End Sub

Private Sub cmd_Tab0_ImportOrders_Click()
On Error GoTo err_Handle
Dim strReceiptNo As String
strReceiptNo = ""

    '��s�c�O�������
    If chk_Tab0_Updatetrpw.Value = 1 Then
        cn.Execute "exec gs_UpdateTRPW", RowsAffect, adExecuteNoRecords
    End If
    
'    '��sOrders���--�æ��ɭPQTY=0���a��
'    str_SQL = "update trp02w set trp02w.otqty = orders.otqty from trp02w join orders on trp02w.receipt_no = orders.orderkey and trp02w.OTConfirmuser is null and trp02w.OTQTY is null and orders.OTQTY is not null "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�ƨ��@�~>>�פJ�ݱƨ��q��
    Screen.MousePointer = vbHourglass
    DoEvents: DoEvents
    Set dg_TRP02W.DataSource = Nothing
    '�ƨ��@�~�G�ݱƨ��q��
    Call CreateRS_Tab0_TRP02W
    
    strSourceFilter = adFilterNone
    DoEvents
    
    '���w����q��̡G�߰� user �O�_�n�M��
    If rs_Tab0_SelectedOrders.RecordCount <> 0 Then
       msg_text = "���J�ݱƨ��q��G[�w����q��] �O�_�i��M��"
       If MsgBox(msg_text, vbYesNo + vbInformation + vbDefaultButton2, msg_title) = vbYes Then
          '�M�����u�s�����ȡA�]�t�w��q��W�ӦC��
          Call Clear_RouteData
          txt_Tab0_RouteNo.Text = ""
        Else
            dg_Tab0_SelectedOrders.Enabled = False
            rs_Tab0_SelectedOrders.MoveFirst
            Do While Not rs_Tab0_SelectedOrders.EOF
                strReceiptNo = strReceiptNo & rs_Tab0_SelectedOrders.Fields("Receipt_no") & "','"
                rs_Tab0_SelectedOrders.MoveNext
            Loop
            
            dg_Tab0_SelectedOrders.Enabled = True
            
       End If
    End If
    
    'dg_Tab0_SelectedOrders
    
    '�ݱƨ��q����J�G����p�p�G�k�s
    dbSelectedCount = 0
    dbsrcSelected_OTqty = 0: dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
    txt_Tab0_srcSelected_OTqty.Text = ""
    txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
    txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""
    
    '���^�ݱƨ��q��
    str_SQL = "Select ' ' as '��',�e�f��,�q��s��,���,�c��,�O��,���n,���q,�Ȥ�s��,isnull(ZIP,'') as ZIP,isnull(�Ȥ�²��,'') as �Ȥ�²��, " & _
              "isnull(Area,'') as Area ,isnull(���A,'') as ���A,isnull(�B�e�a�},'') as �B�e�a�},�q��Ƶ�,��s������,�t�e�ܧO,isnull(����,'') as ����, " & _
              "isnull(�S��ݨD1,'') as �S��ݨD1,isnull(�S��ݨD2,'') as �S��ݨD2,���,�M��,�N��,Receipt_No,C_Receipt_No,�f�D�渹,EXE�^��,isnull(�Ȥ�W��,'') as �Ȥ�W�� " & _
              " From TRPPlan_SourceOrder " & _
              " where Receipt_No not in ( '" & strReceiptNo & " ') " & _
              " Order by �q��s�� "
    strSourceOrderBy = " �q��s�� "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫ݱƨ��q����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    blTRP02WEventEnable = False
    dg_TRP02W.Visible = False
    Do While Not tmp_Rs.EOF
        rs_TRP02W.AddNew
        rs_TRP02W.Fields("�s��").Value = rs_TRP02W.RecordCount
        rs_TRP02W.Fields("�e�f��").Value = tmp_Rs.Fields("�e�f��").Value
        rs_TRP02W.Fields("�q��s��").Value = tmp_Rs.Fields("�q��s��").Value
        rs_TRP02W.Fields("���").Value = tmp_Rs.Fields("���").Value
        rs_TRP02W.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
        rs_TRP02W.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
        rs_TRP02W.Fields("���n").Value = tmp_Rs.Fields("���n").Value
        rs_TRP02W.Fields("���q").Value = tmp_Rs.Fields("���q").Value
        rs_TRP02W.Fields("�Ȥ�s��").Value = tmp_Rs.Fields("�Ȥ�s��").Value
        rs_TRP02W.Fields("Area").Value = tmp_Rs.Fields("Area").Value
        rs_TRP02W.Fields("zip").Value = tmp_Rs.Fields("zip").Value
        rs_TRP02W.Fields("�Ȥ�W��").Value = tmp_Rs.Fields("�Ȥ�W��").Value
        rs_TRP02W.Fields("�B�e�a�}").Value = tmp_Rs.Fields("�B�e�a�}").Value
        rs_TRP02W.Fields("�q��Ƶ�").Value = tmp_Rs.Fields("�q��Ƶ�").Value & ""
        rs_TRP02W.Fields("��s������").Value = RTrim(tmp_Rs.Fields("��s������").Value)
        rs_TRP02W.Fields("�t�e�ܧO").Value = RTrim(tmp_Rs.Fields("�t�e�ܧO").Value)
        rs_TRP02W.Fields("����").Value = tmp_Rs.Fields("����").Value
        rs_TRP02W.Fields("�S��ݨD1").Value = tmp_Rs.Fields("�S��ݨD1").Value
        rs_TRP02W.Fields("�S��ݨD2").Value = tmp_Rs.Fields("�S��ݨD2").Value
        rs_TRP02W.Fields("���").Value = tmp_Rs.Fields("���").Value
        rs_TRP02W.Fields("�M��").Value = tmp_Rs.Fields("�M��").Value
        rs_TRP02W.Fields("�N��").Value = tmp_Rs.Fields("�N��").Value
        rs_TRP02W.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
        rs_TRP02W.Fields("C_Receipt_No").Value = tmp_Rs.Fields("C_Receipt_No").Value
        rs_TRP02W.Fields("�f�D�渹").Value = tmp_Rs.Fields("�f�D�渹").Value
        rs_TRP02W.Fields("EXE�^��").Value = tmp_Rs.Fields("EXE�^��").Value
        rs_TRP02W.Fields("Area").Value = tmp_Rs.Fields("Area").Value
        rs_TRP02W.Fields("�Ȥ�²��").Value = tmp_Rs.Fields("�Ȥ�²��").Value
        rs_TRP02W.Fields("���A").Value = tmp_Rs.Fields("���A").Value
        rs_TRP02W.Update
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    rs_TRP02W.MoveFirst
    dg_TRP02W.Visible = True
    blTRP02WEventEnable = True
    
    '�ݱƨ��q���`�p��T
    Call Retrive_OrderSum
    Screen.MousePointer = vbDefault
    
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��C��-�פJ�ݱƨ��q��", Me.Caption, "cmd_Tab0_ImportOrders_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Print_Click()
'������f��>Barcode�C�L

On Error GoTo err_Handle
'1. ��Ƽg�X Access ��Ʈw >> �@�����s�жK
Dim i As Integer
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From �@�����s�жK"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access)
rs_Access.Open "�@�����s�жK", cnAccess, adOpenStatic, adLockOptimistic

    rs_RouteData.MoveFirst
    Do While Not rs_RouteData.EOF
            For i = 1 To Val(rs_RouteData.Fields("�O��"))
                str_SQL = "Insert into �@�����s�жK (line,�X����,�@�����u�s��,�@���r�p�H,�Ȥ�²��,�`�O��) " & _
                          "Values ('" & i & "','" & Left(Trim(rs_RouteData.Fields("�X����")), 4) & "/" & Mid(Trim(rs_RouteData.Fields("�X����")), 5, 2) & "/" & Right(Trim(rs_RouteData.Fields("�X����")), 2) & "','" & rs_RouteData.Fields("�@�����u�s��") & "','" & Trim(rs_RouteData.Fields("�@���r�p�H").Value) & "','" & rs_RouteData.Fields("�Ȥ�²��") & "','" & i & " / " & Trim(rs_RouteData.Fields("�O��").Value) & "')"
                cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
            Next
            rs_RouteData.MoveNext
    Loop
    cnAccess.CommitTrans
    Tran_Level = 0
    Call DB_Disconnect(cnAccess)
    
    '2. call Access �C�L����
    strAccessDBFileName_FullPath = GetAccessDBFileName
    Set MSAccessAP = New access.Application
    MSAccessAP.Visible = False
    MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)
    
    '[����C�L] �R�O�s -- �Q�� Access ����
    'If chk_Tab2_PreView.Value = vbChecked Then
    '�w���C�L
       MSAccessAP.Visible = True
       MSAccessAP.DoCmd.OpenReport "�@�����s�жK", acViewPreview
       Call Unload_RunLogForm
Exit Sub

err_Handle:
   If Tran_Level <> 0 Then cnAccess.RollbackTrans
   Tran_Level = 0
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
   End If
   Call Unload_RunLogForm
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--TIHI_LABEL", Me.Caption, "cmd_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab0_Query_Click()
    '�ƨ��@�~ >> �d��
    If Len(txt_Tab0_RouteNo.Text) = 0 Then Exit Sub
    
    '���ק蠟���s�G�O�_�w�^��WMS
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "Select EXE_CONFIRM From TRP01T Where Route_No = '" & txt_Tab0_RouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields("EXE_CONFIRM").Value = "2" Then
        msg_text = "�`�N�G�����u�s���w�^��WMS�A�L�k�ק�ΧR��!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
       tmp_Rs.Close
       
    '���ק蠟���s�G�O�_�w�X���T�{
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "Select sdnstatus From TRP05T Where Route_No = '" & txt_Tab0_RouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields("sdnstatus").Value = "1" Then
        msg_text = "�`�N�G�����u�s���X���T�{�A�L�k�ק�ΧR��!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
       tmp_Rs.Close
    
    If rs_Tab0_SelectedOrders.RecordCount <> 0 And blRouteModify = False Then
        '�s�W���u�s���Ҧ��G
        '�I�s [�w��q�沾��(��)] �ӳB�z�w�Q�Ȯɿ���� [�ݱƨ��q��] �٭�^ [�ݱƨ��q��]
        Call cmd_Tab0_SelectedRemove_All_Click
    End If
    If blRouteModify And blRouteChange Then
        '���ĸ��u�s�� & ��Ƥw�D���ʡA�n user �T�{�O�_�s��
        msg_text = "���u�s����ƬO�_�s�ɡH"
        If MsgBox(msg_text, vbOKCancel + vbCritical, msg_title) = vbOK Then
            '�I�s�s�ɵ{��
            Call cmd_Tab0_Save_Click
        Else
            '���s�ɡ��������s���J [�ݱƨ��q��] �w�٭� [���][����] �ާ@�� [�ݱƨ��q��] ���v�T
            Call cmd_Tab0_ImportOrders_Click
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    '�M�����u�s�����ȡA�]�t�w��q��W�ӦC��
    Call Clear_RouteData
    
    '���o���s���
    str_SQL = "Select �X�����,���P���X,�X�Y�Ȧs,�w�p������,�w�p����ɶ�,�B�餽�q,�r�p�H,�r�p�q��,����,�c��,�O��,���n,���q " & _
              "From TRPPlan_RouteQuery Where ���u�s�� = '" & txt_Tab0_RouteNo.Text & "'"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧱ƨ����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        
        '�M�����u�s�����ȡA�]�t�w��q��W�ӦC��
        Call Clear_RouteData
        
        Screen.MousePointer = vbDefault
        txt_Tab0_RouteNo.SelStart = 0: txt_Tab0_RouteNo.SelLength = Len(txt_Tab0_RouteNo.Text)
        txt_Tab0_RouteNo.SetFocus
        Exit Sub
    End If
    txt_Tab0_TRPDate.Text = tmp_Rs.Fields("�X�����").Value
    txt_Tab0_DeliveryCarNo.Text = tmp_Rs.Fields("���P���X").Value
    txt_Tab0_DockNo.Text = tmp_Rs.Fields("�X�Y�Ȧs").Value
    txt_Tab0_CarCheckInDate.Text = tmp_Rs.Fields("�w�p������").Value
    txt_Tab0_CarCheckInTime.Text = tmp_Rs.Fields("�w�p����ɶ�").Value
    txt_Tab0_DeliveryCompany.Text = tmp_Rs.Fields("�B�餽�q").Value
    txt_Tab0_DeliveryDriver.Text = tmp_Rs.Fields("�r�p�H").Value
    txt_Tab0_DeliveryPhone.Text = tmp_Rs.Fields("�r�p�q��").Value
    txt_Tab0_DeliveryCarType.Text = tmp_Rs.Fields("����").Value
    txt_Tab0_Selected_Case.Text = tmp_Rs.Fields("�c��").Value
    txt_Tab0_Selected_Pallet.Text = tmp_Rs.Fields("�O��").Value
    txt_Tab0_Selected_Volumn.Text = tmp_Rs.Fields("���n").Value
    txt_Tab0_Selected_Weight.Text = tmp_Rs.Fields("���q").Value
    tmp_Rs.Close
    
    '���o���s�q��
    str_SQL = "Select �e�f��,�q��s��,ZIP,Area,���A,�Ȥ�²��,�c��,�O��,���n,���q,����,�q��Ƶ�,�S��ݨD1,�S��ݨD2,Receipt_No,EXE�^��,�Ȥ�W��,�B�e�a�} " & _
              "From TRPPlan_RouteQueryOrders Where ���u�s�� = '" & txt_Tab0_RouteNo.Text & "' Order by Receipt_No "
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧭q��W�Ӹ��"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        
        '�M�����u�s�����ȡA�]�t�w��q��W�ӦC��
        Call Clear_RouteData
        
        Screen.MousePointer = vbDefault
        txt_Tab0_RouteNo.SelStart = 0: txt_Tab0_RouteNo.SelLength = Len(txt_Tab0_RouteNo.Text)
        txt_Tab0_RouteNo.SetFocus
        Exit Sub
    End If
    blTab0SelectedOrderEventEnable = False
    Do While Not tmp_Rs.EOF
        rs_Tab0_SelectedOrders.AddNew
        rs_Tab0_SelectedOrders.Fields("�s��").Value = rs_Tab0_SelectedOrders.RecordCount
        rs_Tab0_SelectedOrders.Fields("�e�f��").Value = tmp_Rs.Fields("�e�f��").Value
        rs_Tab0_SelectedOrders.Fields("�q��s��").Value = tmp_Rs.Fields("�q��s��").Value
        rs_Tab0_SelectedOrders.Fields("ZIP").Value = tmp_Rs.Fields("ZIP").Value
        rs_Tab0_SelectedOrders.Fields("Area").Value = tmp_Rs.Fields("Area").Value & ""
        rs_Tab0_SelectedOrders.Fields("���A").Value = tmp_Rs.Fields("���A").Value
        rs_Tab0_SelectedOrders.Fields("�Ȥ�²��").Value = tmp_Rs.Fields("�Ȥ�²��").Value
        rs_Tab0_SelectedOrders.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
        rs_Tab0_SelectedOrders.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
        rs_Tab0_SelectedOrders.Fields("���n").Value = tmp_Rs.Fields("���n").Value
        rs_Tab0_SelectedOrders.Fields("���q").Value = tmp_Rs.Fields("���q").Value
        rs_Tab0_SelectedOrders.Fields("����").Value = tmp_Rs.Fields("����").Value
        rs_Tab0_SelectedOrders.Fields("�q��Ƶ�").Value = tmp_Rs.Fields("�q��Ƶ�").Value
        rs_Tab0_SelectedOrders.Fields("�S��ݨD1").Value = tmp_Rs.Fields("�S��ݨD1").Value
        rs_Tab0_SelectedOrders.Fields("�S��ݨD2").Value = tmp_Rs.Fields("�S��ݨD2").Value
        rs_Tab0_SelectedOrders.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
        rs_Tab0_SelectedOrders.Fields("EXE�^��").Value = tmp_Rs.Fields("EXE�^��").Value
        rs_Tab0_SelectedOrders.Fields("�Ȥ�W��").Value = tmp_Rs.Fields("�Ȥ�W��").Value
        rs_Tab0_SelectedOrders.Fields("�B�e�a�}").Value = tmp_Rs.Fields("�B�e�a�}").Value
        rs_Tab0_SelectedOrders.Update
        tmp_Rs.MoveNext
    Loop
    rs_Tab0_SelectedOrders.MoveFirst
    rs_Tab0_SelectedOrders.Sort = " �s�� asc "
    blTab0SelectedOrderEventEnable = True
    tmp_Rs.Close
    blRouteModify = True
    blRouteChange = False
    strDispRouteNo = Trim(txt_Tab0_RouteNo.Text)
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��C��-�d��", Me.Caption, "cmd_Tab0_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Remove_Click()
    '�ƨ��@�~ >> �� �w����q�����
    If rs_TRP02W Is Nothing Then Exit Sub
    If rs_Tab0_SelectedOrders Is Nothing Then Exit Sub
    '�w����q��Y�L�ϥտ���GDisable �w��������ʧ@�A����~�R
    If dg_Tab0_SelectedOrders.SelBookmarks.Count = 0 Then Exit Sub
    
    blTab0SelectedOrderEventEnable = False
    
    '���������q��s�� Receipt_No
    Dim strReceiptNo As String
    strReceiptNo = rs_Tab0_SelectedOrders.Fields("Receipt_No").Value
       
    '�N���R���� [�w����q��] �[�J [�ݱƨ��q��]
    Call SelectedOrders_Removeto_TRP02W(strReceiptNo)
    '���s���� [�ݱƨ��q��] �� [�s��] ����
    Call ReSet_TRP02W_SeqNo
    
    '�R���ϥտ�����q��G�w����q�泡��
    rs_Tab0_SelectedOrders.Delete
    If Not rs_Tab0_SelectedOrders.EOF Then rs_Tab0_SelectedOrders.MoveFirst
    If dg_Tab0_SelectedOrders.SelBookmarks.Count > 0 Then dg_Tab0_SelectedOrders.SelBookmarks.Remove 0
    '���s�p��w����q��G�c�ơA�O�ơA���n�A���q + �s�����s����
    Call Calculate_SelectedOrders
    
    blTRP02WEventEnable = False
    rs_TRP02W.Filter = adFilterNone
    If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
    If rs_TRP02W.EOF Then
       strSourceFilter = adFilterNone
       rs_TRP02W.Filter = adFilterNone
    End If
    rs_TRP02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    blTRP02WEventEnable = True
    blTab0SelectedOrderEventEnable = True
    
    '���s�p�� [�ݱƨ��C��] ���`�p��T
    Call ReCaculate_OrderSum

End Sub

Private Sub cmd_Tab0_Reserve_Click()
    '�ݱƨ��q��G�O�d�q��
    cmd_Tab0_Reserve.Enabled = False
    
    '�ݱƨ��q��G����p�p�G�k�s
    dbSelectedCount = 0
    dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
    txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
    txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""
    
    '�٭�Ҧ��z��]�w�A�åH�w�] [�s��] �ƦC
    blTRP02WEventEnable = False
    rs_TRP02W.Filter = adFilterNone
    rs_TRP02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    
    Dim strRouteNo As String, intDriveTimes As Integer, dbOrderCnt As Double, iLoop As Double
    strRouteNo = "D"   '�S����u�s���A�κީҦ��O�d�q��
    intDriveTimes = 1
    dbOrderCnt = 0
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    blTab2ReservedEventEnable = False
    '�z��w�����
    rs_TRP02W.Filter = "��='V'"
    If Not rs_TRP02W.EOF Then
        Do While Not rs_TRP02W.EOF
            rs_Tab2_ReservedOrders.AddNew
            For iLoop = 0 To rs_TRP02W.Fields.Count - 1
                rs_Tab2_ReservedOrders.Fields(iLoop).Value = rs_TRP02W.Fields(iLoop).Value
            Next iLoop
            rs_Tab2_ReservedOrders.Fields(1).Value = " "
            rs_Tab2_ReservedOrders.Update
            
            'insert into TRP02T
            str_SQL = "Insert into TRP02T (Route_No,StorerKey,C_Receipt_No,Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                      " Vehicle_ID_No,Drive_Times,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,description,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                      "Select '" & strRouteNo & "',StorerKey,C_Receipt_No,Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                      " 'D'," & intDriveTimes & ",Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,description,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                      "From TRP02W Where Receipt_No = '" & rs_TRP02W.Fields("Receipt_No").Value & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '����� TRP02T Trigger [insert] �i��H�U�@�~
            '   a.�g�J TRP03T -- �ƨ��q�������
            '   b.�R�� TRP03W -- �ݱƨ��q�������
            '   c.�R�� TRP02W -- �ݱƨ��q��D��
            
            rs_TRP02W.MoveNext
        Loop
        '[�ݿ���q��] ���A�R���w������q��
        rs_TRP02W.MoveFirst
        Do While Not rs_TRP02W.EOF
            rs_TRP02W.Delete
            rs_TRP02W.MoveFirst
        Loop
    End If
    
    '��s trp01t & trp05t �� [�c��] [�O��] [���q] [���n]
    str_SQL = "EXEC ReservedOrders_Recalculate " & strRouteNo & " "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    If Tran_Level <> 0 Then
        cn.CommitTrans
        Tran_Level = 0
    End If
    blTab2ReservedEventEnable = True
    
    If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
    If rs_TRP02W.EOF Then
        strSourceFilter = adFilterNone
        rs_TRP02W.Filter = adFilterNone
    End If
    rs_TRP02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    '�����ϥտ�����A
    If dg_TRP02W.SelBookmarks.Count > 0 Then
        dg_TRP02W.SelBookmarks.Remove 0
    End If
    blTRP02WEventEnable = True
    cmd_Tab0_Reserve.Enabled = True
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��@�~-�إ߸��u�s��", Me.Caption, "cmd_Tab0_CreateRoute_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   cmd_Tab0_Reserve.Enabled = True
End Sub

Private Sub cmd_Tab0_Save_Click()
    '�ƨ��@�~ >> ���u�s���ק�Ҧ��s��
    If blRouteModify = False Then
        msg_text = "�D�g [�d��] �{�ǩ���ܤ����� [���u�s��]"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        Exit Sub
    End If
    If blRouteChange = False Then
        msg_text = "[���u�s��] ����ƨå����ʡA�������� [�s��] �{��"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        Exit Sub
    Else
       '�q���Ʀ����ʡA�B�������Q�����A���P�R��
        If rs_Tab0_SelectedOrders.RecordCount = 0 Then
            msg_text = "�����u�s���ثe�w�L�q��A�O�_�R�������s�H"
            If MsgBox(msg_text, vbOKCancel + vbCritical, msg_title) = vbOK Then
                Call Delete_RouteNo(strDispRouteNo)
                Call Clear_RouteData
                txt_Tab0_RouteNo.Text = ""
                Exit Sub
            End If
        End If
    End If
    '�ˮָ��u�s����ƬO�_���T��J
    If RouteData_Check = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    cmd_Tab0_Save.Enabled = False
    
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    Dim intDriveTimes As Integer    '����
    '1.�T�{ [�X�����] �P [���P���X] [�ק��v��] & [��ƬO�_�D����]
    '  �Y���ʫh�������s�p�⨮��
    str_SQL = "Select Rtrim(t05t.Vehicle_ID_No) as ���P���X,Convert(varchar(8),t01t.Delivery_Date,112) as �X�����,Rtrim(Isnull(t01t.AddWho,'')) as AddWho,t05t.Drive_Times as ���� " & _
              "From TRP05T t05t inner join TRP01T t01t on t01t.Route_No = t05t.Route_No " & _
              "Where t05t.Route_No = '" & strDispRouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "���u�s�� [" & strDispRouteNo & "] �w�䤣���ƤF"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    If UCase(tmp_Rs.Fields("AddWho").Value) <> UCase(User_id) And blRouteModifyControl Then
        tmp_Rs.Close
        msg_text = "�v�����ޡG���u�s�����ק�u���\�ѭ�ƨ��̰���"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault: cmd_Tab0_Save.Enabled = True
        Exit Sub
    End If
    intDriveTimes = tmp_Rs.Fields("����").Value
    
    If tmp_Rs.Fields("�X�����").Value <> txt_Tab0_TRPDate.Text Or UCase(tmp_Rs.Fields("���P���X").Value) <> txt_Tab0_DeliveryCarNo.Text Then
        '�X����� or ���P���X�D���ʡG�������s�p�⨮��
        tmp_Rs.Close
        str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
                  "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & txt_Tab0_TRPDate.Text & "' and Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "'"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
    End If
    tmp_Rs.Close
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    '2.��s TRP05T & TRP01T & TRP03T add TRP02T add By Gemini @20080313
    str_SQL = "Update TRP01T Set Delivery_Date = '" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "' " & _
              "Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update TRP05T Set Delivery_Date = '" & Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2) & "', " & _
              "   Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "',Drive_Times = " & intDriveTimes & ",Dock_No = '" & txt_Tab0_DockNo.Text & "',Expect_Date = '" & txt_Tab0_CarCheckInDate.Text & "'," & _
              "   Expect_Time = '" & txt_Tab0_CarCheckInTime.Text & "' " & _
              "Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update TRP03T Set Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "',Drive_Times = " & intDriveTimes & " " & _
              "Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Update TRP02T Set Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "',Drive_Times = " & intDriveTimes & " " & _
              "Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '3.�Ѩ����D�ɧ�s TRP05T �����������
    str_SQL = "Update TRP05T Set Driver = B.Driver , Driver_Phone = B.Driver_Phone, TRP_Company_Code = B.TRP_Company_Code " & _
              "From TRP05T A , TRP09M B Where a.Vehicle_ID_No = b.Vehicle_ID_No and a.Vehicle_ID_No = '" & txt_Tab0_DeliveryCarNo.Text & "' and a.Route_No = '" & strDispRouteNo & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '4..�N TRP02T ������s�Хܬ� [��s�X��] DeleteFlag = '1'
    str_SQL = "Update TRP02T Set DeleteFlag='1' Where Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '5.�N�q���s�X�� DeleteFalg �٭�^ 0
    '  �䤣�쪺�A��ܬO�s�[�J���A�i��s�W�{��
    blTab0SelectedOrderEventEnable = False
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        str_SQL = "Update TRP02T Set DeleteFlag='0' Where Receipt_No = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        If RowsAffect = 0 Then
            '�s�W�q��
            str_SQL = "Insert into TRP02T (Route_No,StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                      " Vehicle_ID_No,Drive_Times,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,OPQTY,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                      "Select '" & strDispRouteNo & "',StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Extern,Case_cnt,Pallet_Qty,Weight,Volumn_Weight," & _
                      " '" & txt_Tab0_DeliveryCarNo.Text & "'," & intDriveTimes & ",Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,EXE_Confirm,description,OTQTY,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                      "From TRP02W Where Receipt_No = '" & rs_Tab0_SelectedOrders.Fields("Receipt_No").Value & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        End If
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    blTab0SelectedOrderEventEnable = True
    
    '6.�N�����q���٭�^ TRP02W & TRP03W
    '(1).�N TRP03T �g�^ TRP03W >> �R�� TRP03T
    str_SQL = "Insert into TRP03W(" & _
              " STORERKEY,RECEIPT_NO,SEQ_NO,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,PALLET_QTY,WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,EXTERN) " & _
              "Select A.STORERKEY,A.RECEIPT_NO,A.SEQ_NO,A.PRODUCT_NO,A.SHIP_UNIT,A.ORDER_QTY,A.PALLET_QTY,A.WEIGHT,A.VOLUMN_WEIGHT,A.Description,A.EXTERN " & _
              "From TRP03T A INNER JOIN TRP02T B ON B.Receipt_No = a.Receipt_No and b.DeleteFlag = '1' and b.Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(2).�N TRP02T �g�^ TRP02W >> �R�� TRP02T
    str_SQL = "Insert into TRP02W(" & _
              " RECEIPT_NO,C_RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              " WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,OTQTY,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
              "Select RECEIPT_NO,C_RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              "  Weight,VOLUMN_WEIGHT,Description,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,OTQTY,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
              "From TRP02T Where DeleteFlag='1' and Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(3).�R�� TRP02T & TRP03T
    str_SQL = "Delete TRP03T FROM TRP02T Where TRP02T.Receipt_No = TRP03T.Receipt_No and TRP02T.DeleteFlag='1' and TRP02T.Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    str_SQL = "Delete From TRP02T Where DeleteFlag='1' and Route_No = '" & strDispRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '7.��s TRP01T & TRP05T ���έp����
    str_SQL = "exec  ReservedOrders_Recalculate " & strDispRouteNo & " "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    If Tran_Level <> 0 Then
       cn.CommitTrans
       Tran_Level = 0
    End If
    
    '�M���ù�����
    Call Clear_RouteData
    txt_Tab0_RouteNo.Text = ""
    cmd_Tab0_Save.Enabled = True
    
    '�ݱƨ��q���`�p��T
    Call Retrive_OrderSum
    
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cn.RollbackTrans
      Tran_Level = 0
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��@�~-���u�s���ק�s��", Me.Caption, "cmd_Tab0_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_SelectCar_Click()
    '�ƨ��@�~ >> �q�����
    If Len(txt_Tab0_TRPDate.Text) = 0 Then
        msg_text = "�Х���J�G�X�����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_TRPDate.SetFocus
        Exit Sub
    Else
        If chk_Tab0_DriveTimes.Value = vbChecked Then
            '��ܹB�e�����ݿ�M��--�]�t�w�Ʃw������������
            Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab0_SelectCar.Name & "1")
        Else
            '��ܹB�e�����ݿ�M��--����ܨ������
            Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Tab0_SelectCar.Name & "2")
        End If
    End If
End Sub

Private Sub cmd_Tab0_Selected_Click()
    '�ݱƨ��q��G���
    
    '�ݱƨ��q��G����p�p�G�k�s
    dbSelectedCount = 0
    dbsrcSelected_OTqty = 0: dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
    txt_Tab0_srcSelected_OTqty.Text = ""
    txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
    txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""
    
    '�٭�Ҧ��z��]�w�A�åH�w�] [�s��] �ƦC
    blTRP02WEventEnable = False
    rs_TRP02W.Filter = adFilterNone
    rs_TRP02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    
    '�z��w�����
    rs_TRP02W.Filter = "��='V'"
    If Not rs_TRP02W.EOF Then
        dg_Tab0_SelectedOrders.Visible = False
        blTab0SelectedOrderEventEnable = False
        Do While Not rs_TRP02W.EOF
                
            '�P�_�O�_�w�g����L
            rs_Tab0_SelectedOrders.Filter = adFilterNone
            rs_Tab0_SelectedOrders.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
            rs_Tab0_SelectedOrders.Filter = "Receipt_No = '" & rs_TRP02W.Fields("Receipt_No").Value & "'"
            '�p�G�O�d�ߩ���ܤ����ĸ��s�A�]�w���s�����ѧO�X��
            If blRouteModify Then blRouteChange = True
            If rs_Tab0_SelectedOrders.EOF Then
                '�s�W������q��
                rs_Tab0_SelectedOrders.AddNew
                rs_Tab0_SelectedOrders.Fields("�s��").Value = 999
                rs_Tab0_SelectedOrders.Fields("�e�f��").Value = rs_TRP02W.Fields("�e�f��").Value
                rs_Tab0_SelectedOrders.Fields("�q��s��").Value = rs_TRP02W.Fields("�q��s��").Value
                rs_Tab0_SelectedOrders.Fields("ZIP").Value = rs_TRP02W.Fields("ZIP").Value
                rs_Tab0_SelectedOrders.Fields("Area").Value = rs_TRP02W.Fields("Area").Value
                rs_Tab0_SelectedOrders.Fields("���").Value = rs_TRP02W.Fields("���").Value
                rs_Tab0_SelectedOrders.Fields("���A").Value = rs_TRP02W.Fields("���A").Value
                rs_Tab0_SelectedOrders.Fields("�Ȥ�²��").Value = rs_TRP02W.Fields("�Ȥ�²��").Value
                rs_Tab0_SelectedOrders.Fields("���").Value = rs_TRP02W.Fields("���").Value
                rs_Tab0_SelectedOrders.Fields("�c��").Value = rs_TRP02W.Fields("�c��").Value
                rs_Tab0_SelectedOrders.Fields("�O��").Value = rs_TRP02W.Fields("�O��").Value
                rs_Tab0_SelectedOrders.Fields("���n").Value = rs_TRP02W.Fields("���n").Value
                rs_Tab0_SelectedOrders.Fields("���q").Value = rs_TRP02W.Fields("���q").Value
                rs_Tab0_SelectedOrders.Fields("����").Value = rs_TRP02W.Fields("����").Value
                rs_Tab0_SelectedOrders.Fields("�q��Ƶ�").Value = rs_TRP02W.Fields("�q��Ƶ�").Value
                rs_Tab0_SelectedOrders.Fields("�S��ݨD1").Value = rs_TRP02W.Fields("�S��ݨD1").Value
                rs_Tab0_SelectedOrders.Fields("�S��ݨD2").Value = rs_TRP02W.Fields("�S��ݨD2").Value
                rs_Tab0_SelectedOrders.Fields("�q��Ƶ�").Value = rs_TRP02W.Fields("�q��Ƶ�").Value
                rs_Tab0_SelectedOrders.Fields("Receipt_No").Value = rs_TRP02W.Fields("Receipt_No").Value
                rs_Tab0_SelectedOrders.Fields("EXE�^��").Value = rs_TRP02W.Fields("EXE�^��").Value
                rs_Tab0_SelectedOrders.Fields("�Ȥ�W��").Value = rs_TRP02W.Fields("�Ȥ�W��").Value
                rs_Tab0_SelectedOrders.Fields("�B�e�a�}").Value = rs_TRP02W.Fields("�B�e�a�}").Value
                rs_Tab0_SelectedOrders.Update
            Else
                '��s������q����
                rs_Tab0_SelectedOrders.Fields("�e�f��").Value = rs_TRP02W.Fields("�e�f��").Value
                rs_Tab0_SelectedOrders.Fields("�q��s��").Value = rs_TRP02W.Fields("�q��s��").Value
                rs_Tab0_SelectedOrders.Fields("ZIP").Value = rs_TRP02W.Fields("ZIP").Value
                rs_Tab0_SelectedOrders.Fields("Area").Value = rs_TRP02W.Fields("Area").Value
                rs_Tab0_SelectedOrders.Fields("���A").Value = rs_TRP02W.Fields("���A").Value
                rs_Tab0_SelectedOrders.Fields("�Ȥ�²��").Value = rs_TRP02W.Fields("�Ȥ�²��").Value
                rs_Tab0_SelectedOrders.Fields("���").Value = rs_TRP02W.Fields("���").Value
                rs_Tab0_SelectedOrders.Fields("�c��").Value = rs_TRP02W.Fields("�c��").Value
                rs_Tab0_SelectedOrders.Fields("�O��").Value = rs_TRP02W.Fields("�O��").Value
                rs_Tab0_SelectedOrders.Fields("���n").Value = rs_TRP02W.Fields("���n").Value
                rs_Tab0_SelectedOrders.Fields("���q").Value = rs_TRP02W.Fields("���q").Value
                rs_Tab0_SelectedOrders.Fields("����").Value = rs_TRP02W.Fields("����").Value
                rs_Tab0_SelectedOrders.Fields("�q��Ƶ�").Value = rs_TRP02W.Fields("�q��Ƶ�").Value
                rs_Tab0_SelectedOrders.Fields("�S��ݨD1").Value = rs_TRP02W.Fields("�S��ݨD1").Value
                rs_Tab0_SelectedOrders.Fields("�S��ݨD2").Value = rs_TRP02W.Fields("�S��ݨD2").Value
                rs_Tab0_SelectedOrders.Fields("Receipt_No").Value = rs_TRP02W.Fields("Receipt_No").Value
                rs_Tab0_SelectedOrders.Fields("EXE�^��").Value = rs_TRP02W.Fields("EXE�^��").Value
                rs_Tab0_SelectedOrders.Fields("�Ȥ�W��").Value = rs_TRP02W.Fields("�Ȥ�W��").Value
                rs_Tab0_SelectedOrders.Fields("�B�e�a�}").Value = rs_TRP02W.Fields("�B�e�a�}").Value
            End If
            rs_TRP02W.MoveNext
        Loop
        '���s�� [�w����q��] ���� [�s��] �P������Ʋέp�G�c�ơA�O�ơA���n�A���q
        Call Calculate_SelectedOrders
        dg_Tab0_SelectedOrders.Visible = True
        blTab0SelectedOrderEventEnable = True
        
        '[�ݿ���q��] ���A�R���w������q��
        rs_TRP02W.MoveFirst
        Do While Not rs_TRP02W.EOF
            rs_TRP02W.Delete
            rs_TRP02W.MoveFirst
        Loop
    End If
    
    If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
    If rs_TRP02W.EOF Then
        rs_TRP02W.Filter = adFilterNone
        strSourceFilter = adFilterNone
    End If
    rs_TRP02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    '�����ϥտ�����A
    If dg_TRP02W.SelBookmarks.Count > 0 Then
        dg_TRP02W.SelBookmarks.Remove 0
    End If
    
    '���s�p�� [�ݱƨ��C��] ���`�p��T
    Call ReCaculate_OrderSum
    
    blTRP02WEventEnable = True

End Sub

Private Sub cmd_Tab0_SelectedCancel_All_Click()
    '�ƨ��@�~ >> X�ݿ��������
    
    '�ݱƨ��q��G����p�p�G�k�s
    dbSelectedCount = 0
    dbsrcSelected_Case = 0: dbsrcSelected_Pallet = 0: dbsrcSelected_Volumn = 0: dbsrcSelected_Weight = 0
    txt_Tab0_srcSelected_Case.Text = "": txt_Tab0_srcSelected_Pallet.Text = ""
    txt_Tab0_srcSelected_Volumn.Text = "": txt_Tab0_srcSelected_Weight.Text = ""
    
    '�٭�Ҧ��z��]�w�A�åH�w�] [�s��] �ƦC
    blTRP02WEventEnable = False
    rs_TRP02W.Filter = adFilterNone
    rs_TRP02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    
    '�z��w�����
    rs_TRP02W.Filter = "��='V'"
    If Not rs_TRP02W.EOF Then
        Do While Not rs_TRP02W.EOF
            rs_TRP02W.Fields("��").Value = " "
            rs_TRP02W.MoveNext
        Loop
    End If
    
    blTRP02WEventEnable = False
    If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
    If rs_TRP02W.EOF Then
        strSourceFilter = adFilterNone
        rs_TRP02W.Filter = adFilterNone
    End If
    rs_TRP02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    blTRP02WEventEnable = True
    
    '�����ϥտ�����A
    If dg_TRP02W.SelBookmarks.Count > 0 Then
        dg_TRP02W.SelBookmarks.Remove 0
    End If
    '�٭� [�ݱƨ��q��] �Ƨǳ]�w
    blTRP02WEventEnable = True
End Sub

Private Sub cmd_Tab0_SelectedCancel_Click()
    '�ƨ��@�~ >> X�ݿ����
    If rs_TRP02W Is Nothing Then Exit Sub
        '�ݿ���q��Y�L�ϥտ���GDisable �ݿ�����A����~�R
        If dg_TRP02W.SelBookmarks.Count = 0 Then Exit Sub
        
        If Trim(rs_TRP02W.Fields(1).Value) = "V" Then
        rs_TRP02W.Fields(1).Value = " "
        dbSelectedCount = dbSelectedCount - 1
        '�ݿ�w��G����p�p��s
        If dbSelectedCount = 0 Then
            dbsrcSelected_Case = 0
            dbsrcSelected_Pallet = 0
            dbsrcSelected_Volumn = 0
            dbsrcSelected_Weight = 0
            txt_Tab0_srcSelected_Case.Text = 0: txt_Tab0_srcSelected_Pallet.Text = 0
            txt_Tab0_srcSelected_Volumn.Text = 0: txt_Tab0_srcSelected_Weight.Text = 0
        Else
            dbsrcSelected_Case = dbsrcSelected_Case - rs_TRP02W.Fields("�c��").Value
            dbsrcSelected_Pallet = dbsrcSelected_Pallet - rs_TRP02W.Fields("�O��").Value
            dbsrcSelected_Volumn = dbsrcSelected_Volumn - rs_TRP02W.Fields("���n").Value
            dbsrcSelected_Weight = dbsrcSelected_Weight - rs_TRP02W.Fields("���q").Value
            txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
            txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
        End If
        '�����ϥտ�����A
        If dg_TRP02W.SelBookmarks.Count > 0 Then
            dg_TRP02W.SelBookmarks.Remove 0
        End If
    End If

End Sub


Private Sub cmd_Tab0_SelectedRemove_All_Click()
    '�ƨ��@�~ >> �� �w����q�����-����
    If rs_TRP02W Is Nothing Then Exit Sub
    If rs_Tab0_SelectedOrders Is Nothing Then Exit Sub
    If rs_Tab0_SelectedOrders.RecordCount = 0 Then Exit Sub
    '���u�s���d�ߡG���ĸ��u�s��
    '���U [�w��q�沾��(��) ���P��R�����u�s��
    If blRouteModify Then
        msg_text = "�T�w�n�R�������u�s�� [" & txt_Tab0_RouteNo.Text & "]"
        If MsgBox(msg_text, vbCritical + vbOKCancel, msg_title) = vbOK Then
            '�R�����w���u�s��
            Call Delete_RouteNo(strDispRouteNo)
            '�M�����u�s�����ȡA�]�t�w��q��W�ӦC��
            Call Clear_RouteData
            txt_Tab0_RouteNo.Text = ""
        End If
        Exit Sub
    End If
    
    blTab0SelectedOrderEventEnable = False
    
    '���������q��s�� Receipt_No
    Dim strReceiptNo As String
    '�v���g�^ [�ݱƨ��q�� TRP02W]
    rs_Tab0_SelectedOrders.Filter = adFilterNone
    rs_Tab0_SelectedOrders.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        strReceiptNo = rs_Tab0_SelectedOrders.Fields("Receipt_No").Value
        '�N���R���� [�w����q��] �[�J [�ݱƨ��q��]
        Call SelectedOrders_Removeto_TRP02W(strReceiptNo)
        rs_Tab0_SelectedOrders.MoveNext
    Loop
       
    '���s���� [�ݱƨ��q��] �� [�s��] ����
    Call ReSet_TRP02W_SeqNo
    
    '�ƨ��@�~�G�w������ݱƨ��q��C�� DBGrid �榡�]�w-ReSet
    Call CreateRS_Tab0_SelectedOrders
    
    '���s�p��w����q��G�c�ơA�O�ơA���n�A���q + �s�����s����
    Call Calculate_SelectedOrders
    '�ƧǤ覡
    
    blTRP02WEventEnable = False
    rs_TRP02W.Filter = adFilterNone
    If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
    If rs_TRP02W.EOF Then
        strSourceFilter = adFilterNone
        rs_TRP02W.Filter = adFilterNone
    End If
    rs_TRP02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    blTRP02WEventEnable = True
    
    '���s�p�� [�ݱƨ��C��] ���`�p��T
    Call ReCaculate_OrderSum
    
    blTab0SelectedOrderEventEnable = True
End Sub

Private Sub cmd_Tab0_srcOrderReset_Click()
    '�ƨ��@�~ >> �����ݱƨ��q��z��Ƨ�
    If rs_TRP02W Is Nothing Then Exit Sub
    '�����z�����A���]�ƧǨ̾�
     blTRP02WEventEnable = False
    '�z��w����̡G�������
    rs_TRP02W.Filter = "��='V'"
    If Not rs_TRP02W.EOF Then
        Do While Not rs_TRP02W.EOF
            rs_TRP02W.Fields(1).Value = " "
            rs_TRP02W.MoveNext
        Loop
    End If
    rs_TRP02W.Filter = adFilterNone
    strSourceFilter = adFilterNone
     'rs_TRP02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    rs_TRP02W.Sort = strSourceOrderBy
    blTRP02WEventEnable = True
    
    '���s�p�� [�ݱƨ��C��] ���`�p��T
    Call ReCaculate_OrderSum

End Sub

Private Sub cmd_Tab1_RouteNoDelete_Click()
    '���u�s���C�� >> ���u�s���R��
    If rs_Tab1_Route.RecordCount = 0 Then Exit Sub
    If dg_Tab1_Route.SelBookmarks.Count = 0 Then Exit Sub
    
    Dim strDeleteRouteNo As String, strCarno As String, dbDriveTimes As Double
    strDeleteRouteNo = Trim(rs_Tab1_Route.Fields("���u�s��").Value)
    strCarno = Trim(rs_Tab1_Route.Fields("���P���X").Value)
    dbDriveTimes = Trim(rs_Tab1_Route.Fields("����").Value)
     
    '���R�������s�G�O�_�w�^��WMS
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "Select isnull(Route,'') From " & strWMSDB & "..orders Where Route = '" & strDeleteRouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        msg_text = "�`�N�GWMS�t�Φ������u�s���ɡA�L�k�ק�ΧR��!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
       tmp_Rs.Close

    '���R�������s�G�O�_�w�X���T�{
    'str_SQL = "Select c_Route_No  From SDN01T Where c_Route_No = '" & strDeleteRouteNo & "'"
    'Terry 20191127 �אּ�ˬd�X�����A
    str_SQL = "Select Route_No  From TRP05T Where Route_No = '" & strDeleteRouteNo & "' and sdnstatus = '1' "
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�`�N�G�����u�s���w�X���T�{�A�L�k�R��! "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    tmp_Rs.Close

    '���R�������s�G�O�_�w�������� Add by Terry 20191127
    str_SQL = "Select Route_No  From SDN02W Where Route_No = '" & strDeleteRouteNo & "' "
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�`�N�G�����u�s���w�������աA�L�k�R��! "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    tmp_Rs.Close
    

    msg_text = "�T�{�R�����u�s���G" & strDeleteRouteNo
    If MsgBox(msg_text, vbYesNo + vbCritical + vbDefaultButton2, msg_title) = vbNo Then Exit Sub

    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)

    '���ұ��R�������s�A�ƨ��̬O�_�����ɵn�J���ϥΪ�
    str_SQL = "Select Rtrim(Isnull(AddWho,'')) as AddWho From TRP01T Where Route_No = '" & strDeleteRouteNo & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "��Ʋ��`�G�䤣����R�������u�s��!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    Else
        If UCase(tmp_Rs.Fields("AddWho").Value) <> UCase(User_id) And blRouteModifyControl = True Then
            tmp_Rs.Close
            msg_text = "�v�����ޡG���u�s�����R���u���\�ѭ�ƨ��̰���"
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            Exit Sub
        End If
    End If
    tmp_Rs.Close

    '�R�����s
    Call Delete_RouteNo(strDeleteRouteNo)
    
    '�R���d�ߵ��G���ӵ����u�s��--rs_Tab1_RouteOrders
    rs_Tab1_RouteOrders.Filter = adFilterNone
    rs_Tab1_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    rs_Tab1_RouteOrders.Filter = "���u�s��='" & strDeleteRouteNo & "'"
    If Not rs_Tab1_RouteOrders.EOF Then
        Do While Not rs_Tab1_RouteOrders.EOF
            rs_Tab1_RouteOrders.Delete
            rs_Tab1_RouteOrders.MoveFirst
        Loop
    End If
    rs_Tab1_RouteOrders.Filter = adFilterNone
    rs_Tab1_RouteOrders.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    
    '(7).�R���d�ߵ��G���ӵ����u�s��--rs_Tab1_Route
    rs_Tab1_Route.Delete
    If Not rs_Tab1_Route.EOF Then rs_Tab1_Route.MoveFirst
    
    blTab1RouteEventEnable = True
    Screen.MousePointer = vbDefault
    
    
On Error GoTo err_Handle2
    'Terry 20200220 �P�B�R��BestAPP�W�Ӹ��s���
    Dim HttpClient As Object
    Set HttpClient = CreateObject("Microsoft.XMLHTTP")
    HttpClient.Open "POST", "https://entrance-bestlog.azurewebsites.net/api/BestApp/BestAppTMS/DeleteRouteNoByWareHouse?Route_NO=" & strDeleteRouteNo & "&WareHouse=GYDC_BEST", False
    HttpClient.setRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    HttpClient.Send
    
    Exit Sub

err_Handle2:
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���u�s���C��-���u�s���R��", Me.Caption, "cmd_Tab1_RouteNoDelete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_RouteNoQuery_Click()
    '���u�s���C�� >> ���u�s���d��
    If Len(Trim(txt_Tab1_RouteNo.Text)) = 0 Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    On Error GoTo err_Handle
    
    '�]�w���u�s���C��
    blTab1RouteEventEnable = False
    Call CreateRS_Tab1_Route
    blTab1RouteEventEnable = True
    '�]�w���u�s�����q��C��
    Call CreateRS_Tab1_RouteOrders
    
    str_SQL = "Select  Rtrim(a1.Route_No) as ���u�s�� " & _
            ",Convert(varchar(8),a1.Delivery_Date,112) as �X����� " & _
            ",Rtrim(a2.Vehicle_ID_No) as ���P���X " & _
            ",a2.Drive_Times as ���� " & _
            ",Rtrim(Isnull(a2.Driver,'')) as �r�p�H " & _
            ",Round((select isnull(sum(t2.otqty),0) from trp02t t2 where a1.route_no = t2.route_no),2) as ��� " & _
            ",Round(a1.Case_cnt,2) as �c�� " & _
            ",Round(a1.Pallet_Qty,2) as �O�� " & _
            ",Round(a1.Volumn_Weight,2) as ���n " & _
            ",Round(a1.Weight,2) as ���q " & _
            ",Rtrim(Isnull(b1.VEHICLE_TYPE,'')) as ���� " & _
            ",Rtrim(Isnull(a2.Dock_No,'')) as �X�Y�Ȧs " & _
            ",Rtrim(Isnull(a2.Expect_Date,'')) as �w�p������ " & _
            ",Rtrim(Isnull(a2.Expect_time,'')) as �w�p����ɶ� " & _
            ",Case a1.EXE_Confirm When '0' Then '�s�ظ��s' When '1' Then '�]�w�^��' When '2' Then '�w�^��' When '9' Then '�w���z�f' else '�������A' End as EXE�^�� " & _
            ",Rtrim(Isnull(a1.AddWho,'')) as �ƨ��� " & _
            "From TRP01T a1 " & _
            "join TRP05T a2 on a2.Route_No = a1.Route_No " & _
            "join TRP09M b1 on b1.Vehicle_ID_No = a2.Vehicle_ID_No " & _
            "Left join TRP15M b2 on b2.Vehicle_Type = b1.Vehicle_Type " & _
            "Where Left(a1.Route_No,1) = 'F' and Rtrim(a1.Route_No) like '%" & RTrim(txt_Tab1_RouteNo.Text) & "%'  order by Rtrim(a1.Route_No)"
    
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧸��u�s�����(TRP01T)"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    blTab1RouteEventEnable = False
    Do While Not tmp_Rs.EOF
        rs_Tab1_Route.AddNew
        rs_Tab1_Route.Fields("�s��").Value = rs_Tab1_Route.RecordCount
        rs_Tab1_Route.Fields("���u�s��").Value = tmp_Rs.Fields("���u�s��").Value
        rs_Tab1_Route.Fields("�X�����").Value = tmp_Rs.Fields("�X�����").Value
        rs_Tab1_Route.Fields("���P���X").Value = tmp_Rs.Fields("���P���X").Value
        rs_Tab1_Route.Fields("����").Value = tmp_Rs.Fields("����").Value
        rs_Tab1_Route.Fields("�r�p�H").Value = tmp_Rs.Fields("�r�p�H").Value
        rs_Tab1_Route.Fields("���").Value = tmp_Rs.Fields("���").Value
        rs_Tab1_Route.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
        rs_Tab1_Route.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
        rs_Tab1_Route.Fields("���n").Value = tmp_Rs.Fields("���n").Value
        rs_Tab1_Route.Fields("���q").Value = tmp_Rs.Fields("���q").Value
        rs_Tab1_Route.Fields("����").Value = tmp_Rs.Fields("����").Value
        rs_Tab1_Route.Fields("�X�Y�Ȧs").Value = tmp_Rs.Fields("�X�Y�Ȧs").Value
        rs_Tab1_Route.Fields("�w�p������").Value = tmp_Rs.Fields("�w�p������").Value
        rs_Tab1_Route.Fields("�w�p����ɶ�").Value = tmp_Rs.Fields("�w�p����ɶ�").Value
        rs_Tab1_Route.Fields("EXE�^��").Value = tmp_Rs.Fields("EXE�^��").Value
        rs_Tab1_Route.Fields("�ƨ���").Value = tmp_Rs.Fields("�ƨ���").Value
        rs_Tab1_Route.Update
        tmp_Rs.MoveNext
    Loop
    rs_Tab1_Route.MoveFirst
    blTab1RouteEventEnable = True
    tmp_Rs.Close
    'TRP03W
    str_SQL = "Select  Rtrim(a1.Route_No) as ���u�s�� " & _
            ",Convert(varchar(8),a1.Arrive_Date,112) as �e�f�� " & _
            ",Rtrim(a1.Receipt_No) + '(' + Rtrim(a1.StorerKey)+'-'+Rtrim(a1.Extern)+')' as �q��s�� " & _
            ",Rtrim(a2.ZIP) as ZIP " & _
            ",Rtrim(a2.Full_Name) as �Ȥ�W�� " & _
            ",Rtrim(a2.address) as �Ȥ�a�} " & _
            ",��� =  isnull(a1.otqty,(select sum(isnull(o.otqty,0)) from orders o where o.orderkey = a1.c_receipt_no)) " & _
            ",Round(isnull(a1.Case_cnt,0),2) as �c�� " & _
            ",Round(isnull(a1.Pallet_Qty,0),2) as �O�� " & _
            ",Round(isnull(a1.Volumn_Weight,0),2) as ���n " & _
            ",Round(isnull(a1.Weight,0),2) as ���q " & _
            ",Rtrim(a1.Receipt_No) as Receipt_No " & _
            ",Case a1.EXE_Confirm When '0' Then '�s�ظ��s' When '1' Then '�]�w�^��' When '2' Then '�w�^��' When '9' Then '�w���z�f' else '�������A' End  AS EXE�^�� " & _
            ",Rtrim(Isnull(a2.Area_Code,'')) as Area " & _
            ",Rtrim(Isnull(a1.urgent_mark,'')) as ��� " & _
            ",Rtrim(Isnull(a1.Priority,'')) as ���A " & _
            ",Rtrim(Isnull(a2.Short_Name,'')) as �Ȥ�²�� " & _
            ",Rtrim(Isnull(a1.Priority,'')) as ���A " & _
            ",�q��Ƶ� = Rtrim(Isnull(a1.description,'')) " & _
            "From TRP02T a1 " & _
            "inner join TRP01M a2 on a2.ConsigneeKey = a1.ConsigneeKey and a2.storerkey = a1.storerkey " & _
            "Where Rtrim(a1.Route_No) like '%" & txt_Tab1_RouteNo.Text & "%' order by Rtrim(a1.Route_No),Rtrim(a1.Receipt_No)"
            
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���u�s�����q����(TRP02T)"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Do While Not tmp_Rs.EOF
        rs_Tab1_RouteOrders.AddNew
        rs_Tab1_RouteOrders.Fields("�s��").Value = rs_Tab1_RouteOrders.RecordCount
        rs_Tab1_RouteOrders.Fields("���u�s��").Value = tmp_Rs.Fields("���u�s��").Value
        rs_Tab1_RouteOrders.Fields("�e�f��").Value = tmp_Rs.Fields("�e�f��").Value
        rs_Tab1_RouteOrders.Fields("�q��s��").Value = tmp_Rs.Fields("�q��s��").Value
        rs_Tab1_RouteOrders.Fields("ZIP").Value = tmp_Rs.Fields("ZIP").Value
        rs_Tab1_RouteOrders.Fields("�Ȥ�W��").Value = tmp_Rs.Fields("�Ȥ�W��").Value
        rs_Tab1_RouteOrders.Fields("�a�}").Value = tmp_Rs.Fields("�Ȥ�a�}").Value
        rs_Tab1_RouteOrders.Fields("���").Value = tmp_Rs.Fields("���").Value
        rs_Tab1_RouteOrders.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
        rs_Tab1_RouteOrders.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
        rs_Tab1_RouteOrders.Fields("���n").Value = tmp_Rs.Fields("���n").Value
        rs_Tab1_RouteOrders.Fields("���q").Value = tmp_Rs.Fields("���q").Value
        rs_Tab1_RouteOrders.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
        rs_Tab1_RouteOrders.Fields("EXE�^��").Value = tmp_Rs.Fields("EXE�^��").Value
        rs_Tab1_RouteOrders.Fields("Area").Value = tmp_Rs.Fields("Area").Value
        rs_Tab1_RouteOrders.Fields("���A").Value = tmp_Rs.Fields("���A").Value
        rs_Tab1_RouteOrders.Fields("�Ȥ�²��").Value = tmp_Rs.Fields("�Ȥ�²��").Value
        rs_Tab1_RouteOrders.Fields("�q��Ƶ�").Value = tmp_Rs.Fields("�q��Ƶ�").Value
        rs_Tab1_RouteOrders.Update
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    rs_Tab1_RouteOrders.MoveFirst
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-���u�s���C��-���u�s���d��", Me.Caption, "cmd_Tab1_RouteNoQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Delete_Click()
    '�O�d�q�� >> ���� [�ݱƨ��q��]
    If rs_Tab2_ReservedOrders Is Nothing Then Exit Sub
    If rs_Tab2_ReservedOrders.RecordCount = 0 Then Exit Sub
    DelRecord = MsgBox("�R�����ƵL�k�_��,�T�w�n�R��? ", vbQuestion + vbYesNo, "�R��")
    If DelRecord = vbNo Then
        Exit Sub
    End If
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    blTab2ReservedEventEnable = False
    blTRP02WEventEnable = False
    cmd_Tab2_Delete.Enabled = False
    
    '�z��w�����
    rs_Tab2_ReservedOrders.Filter = "��='V'"
    If rs_Tab2_ReservedOrders.EOF Then
        rs_Tab2_ReservedOrders.Filter = adFilterNone
        rs_Tab2_ReservedOrders.Sort = "�s�� ASC"
        blTab2ReservedEventEnable = True
        blTRP02WEventEnable = True
        cmd_Tab2_Remove.Enabled = True
        Exit Sub
    End If
    
    Dim strRouteNo As String, iLoop As Double
    strRouteNo = "D"
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    rs_Tab2_ReservedOrders.MoveFirst
    Do While Not rs_Tab2_ReservedOrders.EOF
        str_SQL = "delete TRP02T where receipt_no ='" & rs_Tab2_ReservedOrders.Fields("TMS�渹").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "delete TRP03T where receipt_no ='" & rs_Tab2_ReservedOrders.Fields("TMS�渹").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "delete TRP02W where receipt_no ='" & rs_Tab2_ReservedOrders.Fields("TMS�渹").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "delete TRP03W where receipt_no ='" & rs_Tab2_ReservedOrders.Fields("TMS�渹").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        str_SQL = "delete TRP02W_TEMP where receipt_no ='" & rs_Tab2_ReservedOrders.Fields("TMS�渹").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        cn.Execute "delete status where orderkey ='" & rs_Tab2_ReservedOrders("TMS�渹") & "'", RowsAffect, adExecuteNoRecords
        
        str_SQL = "update orders set B_PHONE2='00',trafficCop=null,type='�R��' ,editdate = getdate()  where orderkey='" & rs_Tab2_ReservedOrders.Fields("TMS�渹").Value & "' and priority = '" & rs_Tab2_ReservedOrders.Fields("���A").Value & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        rs_Tab2_ReservedOrders.MoveNext
    Loop
    
    '[�ݿ���q��] ���A�R���w������q��
    rs_Tab2_ReservedOrders.MoveFirst
    Do While Not rs_Tab2_ReservedOrders.EOF
        rs_Tab2_ReservedOrders.Delete
        rs_Tab2_ReservedOrders.MoveFirst
    Loop
    
'    '��s trp01t & trp05t �� [�c��] [�O��] [���q] [���n]
'    str_SQL = "EXEC ReservedOrders_Recalculate " & strRouteNo & " "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    If Tran_Level <> 0 Then
        cn.CommitTrans
        Tran_Level = 0
    End If
    
    If Not (rs_TRP02W Is Nothing) Then
        If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
        If rs_TRP02W.EOF Then
            strSourceFilter = adFilterNone
            rs_TRP02W.Filter = adFilterNone
        End If
        rs_TRP02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    End If
    
    rs_Tab2_ReservedOrders.Filter = adFilterNone
    rs_Tab2_ReservedOrders.Sort = "�s�� ASC"
    blTab2ReservedEventEnable = True
    blTRP02WEventEnable = True
    
    cmd_Tab2_Delete.Enabled = True
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�O�d�q��-���ܫݱƨ��q��C��", Me.Caption, "cmd_Tab2_Remove_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   blTab2ReservedEventEnable = True
   blTRP02WEventEnable = True
   cmd_Tab2_Delete.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Remove_Click()

    '�O�d�q�� >> ���� [�ݱƨ��q��]
    If rs_Tab2_ReservedOrders Is Nothing Then Exit Sub
    If rs_Tab2_ReservedOrders.RecordCount = 0 Then Exit Sub
    
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    blTab2ReservedEventEnable = False
    blTRP02WEventEnable = False
    cmd_Tab2_Remove.Enabled = False
    
    '�z��w�����
    rs_Tab2_ReservedOrders.Filter = "��='V'"
    If rs_Tab2_ReservedOrders.EOF Then
        rs_Tab2_ReservedOrders.Filter = adFilterNone
        rs_Tab2_ReservedOrders.Sort = "�s�� ASC"
        blTab2ReservedEventEnable = True
        blTRP02WEventEnable = True
        cmd_Tab2_Remove.Enabled = True
        Exit Sub
    End If
    
    Dim strRouteNo As String, iLoop As Double
    strRouteNo = "D"
    
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    Call ReDim_Recordset(tmp_Rs)
    
    rs_Tab2_ReservedOrders.MoveFirst
    Do While Not rs_Tab2_ReservedOrders.EOF
    
    '�ˬd�q��O�_�R��
    tmp_Rs.Open "select receipt_no from trp02t where route_no = 'D' and receipt_no = '" & rs_Tab2_ReservedOrders("TMS�渹") & "' ", cn
    If Not tmp_Rs.EOF Then

       '(1).�N TRP03T �g�^ TRP03W >> �R�� TRP03T
       str_SQL = "Insert into TRP03W(" & _
                 "   STORERKEY,RECEIPT_NO,SEQ_NO,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,PALLET_QTY,WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,EXTERN) " & _
                 "Select A.STORERKEY,A.RECEIPT_NO,A.SEQ_NO,A.PRODUCT_NO,A.SHIP_UNIT,A.ORDER_QTY,A.PALLET_QTY,A.WEIGHT,A.VOLUMN_WEIGHT,A.Description,A.EXTERN " & _
                 "From TRP03T A Where a.Route_No = '" & strRouteNo & "' and Receipt_No = '" & rs_Tab2_ReservedOrders.Fields("TMS�渹").Value & "'"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

       If Not (rs_TRP02W Is Nothing) Then
            rs_TRP02W.AddNew
            For iLoop = 0 To rs_Tab2_ReservedOrders.Fields.Count - 1
                rs_TRP02W.Fields(iLoop).Value = rs_Tab2_ReservedOrders.Fields(iLoop).Value
            Next iLoop
            rs_TRP02W.Fields(0).Value = rs_TRP02W.RecordCount
            rs_TRP02W.Fields(1).Value = " "
            rs_TRP02W.Update
       End If
    
       '(2).�N TRP02T �g�^ TRP02W >> �R�� TRP02T
       str_SQL = "Insert into TRP02W(" & _
                 "   RECEIPT_NO,C_RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
                 "   WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
                 "Select RECEIPT_NO,C_RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
                 "  Weight,VOLUMN_WEIGHT,Description,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,OTQty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
                 "From TRP02T Where Route_No = '" & strRouteNo & "' and Receipt_No = '" & rs_Tab2_ReservedOrders.Fields("TMS�渹").Value & "'"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
       '(3).�R�� TRP02T & TRP03T
       str_SQL = "Delete From TRP03T Where Route_No = '" & strRouteNo & "' and Receipt_No = '" & rs_Tab2_ReservedOrders.Fields("TMS�渹").Value & "'"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
       
       str_SQL = "Delete From TRP02T Where Route_No = '" & strRouteNo & "' and Receipt_No = '" & rs_Tab2_ReservedOrders.Fields("TMS�渹").Value & "'"
       cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
       
    End If
    tmp_Rs.Close
    
       rs_Tab2_ReservedOrders.MoveNext
    Loop
    
    '[�ݿ���q��] ���A�R���w������q��
    rs_Tab2_ReservedOrders.MoveFirst
    Do While Not rs_Tab2_ReservedOrders.EOF
        rs_Tab2_ReservedOrders.Delete
        rs_Tab2_ReservedOrders.MoveFirst
    Loop
    
    '��s trp01t & trp05t �� [�c��] [�O��] [���q] [���n]
    str_SQL = "EXEC ReservedOrders_Recalculate " & strRouteNo & " "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    If Tran_Level <> 0 Then
        cn.CommitTrans
        Tran_Level = 0
    End If
    
    If Not (rs_TRP02W Is Nothing) Then
        If strSourceFilter <> "0" Then rs_TRP02W.Filter = strSourceFilter
        If rs_TRP02W.EOF Then
            strSourceFilter = adFilterNone
            rs_TRP02W.Filter = adFilterNone
        End If
        rs_TRP02W.Sort = strSourceOrderBy  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    End If
    
    rs_Tab2_ReservedOrders.Filter = adFilterNone
    rs_Tab2_ReservedOrders.Sort = "�s�� ASC"
    blTab2ReservedEventEnable = True
    blTRP02WEventEnable = True
    
    cmd_Tab2_Remove.Enabled = True
    
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�O�d�q��-���ܫݱƨ��q��C��", Me.Caption, "cmd_Tab2_Remove_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   blTab2ReservedEventEnable = True
   blTRP02WEventEnable = True
   cmd_Tab2_Remove.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_ShowAll_Click()
    '�ƨ��@�~>>��ܩҦ��O�d�q����
    Screen.MousePointer = vbHourglass
    DoEvents: DoEvents
    
    '�O�d�q��C��
    blTab2ReservedEventEnable = False
    Call CreateRS_Tab2_ReservedOrders
    DoEvents
    
    '���^�O�d�q����
    str_SQL = "Select ' ' as '��',�e�f��,�q��s��,���,�c��,�O��,���n,���q,�Ȥ�s��,ZIP,�Ȥ�²��,Area,���A,�B�e�a�},�q��Ƶ�,��s������,�t�e�ܧO,����,�S��ݨD1,�S��ݨD2,���,�M��,�N��,TMS�渹,C_RECEIPT_NO,�f�D�渹,EXE�^��,�Ȥ�W�� " & _
              "From TRPPlan_ReservedOrder Order by �q��s�� "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫O�d�q����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Dim iLoop As Double
    Do While Not tmp_Rs.EOF
        rs_Tab2_ReservedOrders.AddNew
        For iLoop = 1 To rs_Tab2_ReservedOrders.Fields.Count - 1
            rs_Tab2_ReservedOrders.Fields(iLoop).Value = tmp_Rs.Fields(iLoop - 1).Value & ""
        Next iLoop
        rs_Tab2_ReservedOrders.Fields(0).Value = rs_Tab2_ReservedOrders.RecordCount
        rs_Tab2_ReservedOrders.Update
        tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    blTab2ReservedEventEnable = True
    
    Screen.MousePointer = vbDefault
    Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�O�d�q��-��ܥ����q��", Me.Caption, "cmd_Tab2_ShowAll_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_FilterAndSort_Click()
    '�ƨ��@�~ >> �O�d�q��j�M
    If rs_Tab2_ReservedOrders Is Nothing Then Exit Sub
    If rs_Tab2_ReservedOrders.RecordCount = 0 Then Exit Sub
    
    strFormName_FilterAndSort = Me.Name
    strRSName_FilterAndSort = "rs_Tab2_ReservedOrders"
    
    If ShowForm_RS_FilterAndSort(rs_Tab2_ReservedOrders, "�O�d�q��", Me.Tag) = False Then
        MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    Me.WindowState = 2

End Sub

Private Sub cmd_Tab2_Reset_Click()
    '�ƨ��@�~ >> �����O�d�q��z��Ƨ�
    '�����z�����A���]�ƧǨ̾�
     blTab2ReservedEventEnable = False
     rs_Tab2_ReservedOrders.Filter = adFilterNone
     rs_TRP02W.Sort = "�s�� ASC"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
     blTab2ReservedEventEnable = True
End Sub

Private Sub cmd_Tab3_Cancel_Click()
    txt_Tab3_DeliveryDate_Start.Text = ""
    txt_Tab3_DeliveryDate_End.Text = ""
    Set gd_Tab3_OrderSum.DataSource = Nothing
    cmd_Tab3_Cancel.Enabled = False
    cmd_Tab3_Excel.Enabled = False
End Sub

Private Sub cmd_Tab3_Excel_Click()
    If gd_Tab3_OrderSum Is Nothing Then Exit Sub
    If rs_Tab3_OrderSum.RecordCount = 0 Then Exit Sub
    On Error GoTo err_Handle
    '�N��Ƽg�Jexcel��
    Dim MyXlsApp As Excel.Application   '�}��excel��
    Dim objFld As Field
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    '�s�WWookbooks
    MyXlsApp.Workbooks.Add
    '�s�WSheets
'    MyXlsApp.Sheets.Add
'    MyXlsApp.Sheets("Sheet1").Select
'    MyXlsApp.Sheets("Sheet1").Name = "�q����R"
    MyXlsApp.ActiveSheet.Name = "�q����R"
    
    Dim i As Integer
    i = 1
    ''�X����,���s,����,�q��,��ڤH,�渹,���,�ƶq,�������,���I���,��L���B,��],�ꦬ,��I,�_�I,���I,�Ƶ�
    MyXlsApp.Cells(i, 1).Value = "�a��"
    MyXlsApp.Cells(i, 2).Value = "�~�n"
    MyXlsApp.Cells(i, 3).Value = "���q"
    
    i = i + 1
    rs_Tab3_OrderSum.MoveFirst
    '���,����,�渹,�Z�O,�ɥX,�٤J
    Do While Not rs_Tab3_OrderSum.EOF
        MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '�x�s��榡 >> �Ʀr >> ���O = ��r
        MyXlsApp.Cells(i, 1).Value = Trim(rs_Tab3_OrderSum.Fields(1))
        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab3_OrderSum.Fields(2))
        MyXlsApp.Cells(i, 3).Value = Trim(rs_Tab3_OrderSum.Fields(3))
        rs_Tab3_OrderSum.MoveNext
        i = i + 1
    Loop
    
    '�����ϥ�
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '�e�uh
    MyXlsApp.Range("A1:C" & i - 1).Select
    MyXlsApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    MyXlsApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With MyXlsApp.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    MyXlsApp.Visible = True
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    'Exit Sub
    
    '���ƨ��q��
    '�N��Ƽg�Jexcel��
'    MyXlsApp.Sheets("Sheet2").Select
'    MyXlsApp.Sheets("Sheet2").Name = "���ƨ��q����R"
    MyXlsApp.ActiveSheet.Name = "���ƨ��q����R"
    
    'Dim i As Integer
    i = 1
    ''�X����,���s,����,�q��,��ڤH,�渹,���,�ƶq,�������,���I���,��L���B,��],�ꦬ,��I,�_�I,���I,�Ƶ�
    MyXlsApp.Cells(i, 1).Value = "�a��"
    MyXlsApp.Cells(i, 2).Value = "�~�n"
    MyXlsApp.Cells(i, 3).Value = "���q"
    
    i = i + 1
    rs_Tab3_Trp02wSum.MoveFirst
    '���,����,�渹,�Z�O,�ɥX,�٤J
    Do While Not rs_Tab3_Trp02wSum.EOF
        MyXlsApp.Cells(i, 1).NumberFormatLocal = "@" '�x�s��榡 >> �Ʀr >> ���O = ��r
        MyXlsApp.Cells(i, 1).Value = Trim(rs_Tab3_Trp02wSum.Fields(1))
        MyXlsApp.Cells(i, 2).Value = Trim(rs_Tab3_Trp02wSum.Fields(2))
        MyXlsApp.Cells(i, 3).Value = Trim(rs_Tab3_Trp02wSum.Fields(3))
        rs_Tab3_Trp02wSum.MoveNext
        i = i + 1
    Loop
    
    '�����ϥ�
    MyXlsApp.Cells.Select
    With MyXlsApp.Selection.Interior
        .ColorIndex = 2
        .Pattern = xlSolid
    End With
    '�e�uh
    MyXlsApp.Range("A1:C" & i - 1).Select
    MyXlsApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    MyXlsApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With MyXlsApp.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With MyXlsApp.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    MyXlsApp.Visible = True
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�ƨ��@�~-�q����R", Me.Caption, "cmd_Tab3_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab3_Query_Click()
    '�q����R
    Dim strSubwhere, str_Where, str_SQl1 As String
    strSubwhere = ""
    '�e�f���
    txt_Tab3_DeliveryDate_Start.Text = Trim(txt_Tab3_DeliveryDate_Start.Text)
    txt_Tab3_DeliveryDate_End.Text = Trim(txt_Tab3_DeliveryDate_End.Text)
    strSubwhere = ""
    If Len(txt_Tab3_DeliveryDate_Start.Text) > 0 And Len(txt_Tab3_DeliveryDate_End.Text) > 0 Then
        strSubwhere = "  Between '" & txt_Tab3_DeliveryDate_Start.Text & "' and '" & txt_Tab3_DeliveryDate_End.Text & "' "
    ElseIf Len(txt_Tab3_DeliveryDate_Start.Text) > 0 And Len(txt_Tab3_DeliveryDate_End.Text) = 0 Then
        strSubwhere = "  = '" & txt_Tab3_DeliveryDate_Start.Text & "' "
    ElseIf Len(txt_Tab3_DeliveryDate_Start.Text) = 0 And Len(txt_Tab3_DeliveryDate_End.Text) > 0 Then
        strSubwhere = "  = '" & txt_Tab3_DeliveryDate_End.Text & "' "
    End If
    If Len(strSubwhere) > 0 Then
        If Len(str_Where) = 0 Then
           str_Where = strSubwhere
        Else
           str_Where = str_Where & " and " & strSubwhere
        End If
    End If

    If Len(str_Where) = 0 Then
        msg_text = "�п�J���"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        str_SQL = "select case when t1m.AREA_CODE='C' then '�x�����y'  when t1m.AREA_CODE='S' then '����' else '�x�n' end as �a��, " & _
                "round(sum(cast(s.Busr4 as float)*od.OriginalQty),0) as �~�n,round(sum(cast(s.Stdgrosswgt as float)*od.OriginalQty),0) as ���q " & _
                "from orderdetail od " & _
                "inner join gv_skuxpack s on od.sku=s.sku and s.storerkey = od.storerkey " & _
                "inner join orders o on od.orderkey=o.orderkey " & _
                "inner join TRP01M t1m on t1m.ConsigneeKey = o.ConsigneeKey " & _
                "where t1m.AREA_CODE in ('C','S','S1')  and convert(char(8),o.DeliveryDate,112) " & str_Where & " " & _
                "group by t1m.AREA_CODE " & _
                "Union " & _
                "select  case when  left(t1m.SHORT_NAME,2)  in ('���F','�¦�','����')  then '�̪F�g�P'  when  left(t1m.SHORT_NAME,2)  in ('����','����','�p��')  then '�x�n�g�P' else '���ϸg�P' end as �a��, " & _
                "round(sum(cast(s.Busr4 as float)*od.OriginalQty),0) as �~�n,round(sum(cast(s.Stdgrosswgt as float)*od.OriginalQty),0) as ���q " & _
                "from orderdetail od " & _
                "inner join gv_skuxpack s on od.sku=s.sku and s.storerkey = od.storerkey " & _
                "inner join orders o on od.orderkey=o.orderkey " & _
                "inner join TRP01M t1m on t1m.ConsigneeKey = o.ConsigneeKey " & _
                "where t1m.AREA_CODE in ('w') " & _
                "and  left(t1m.SHORT_NAME,2)  in ('�͸�','����','����','���F','�¦�','����','�i��','�p��','�j�w','�O��') " & _
                "and convert(char(8),o.DeliveryDate,112) " & str_Where & " " & _
                "group by case when  left(t1m.SHORT_NAME,2)  in ('���F','�¦�','����')  then '�̪F�g�P'  when  left(t1m.SHORT_NAME,2)  in ('����','����','�p��')  then '�x�n�g�P' else '���ϸg�P' end "
                
        str_SQl1 = "select case when t1m.AREA_CODE='C' then '�x�����y'  when t1m.AREA_CODE='S' then '����' else '�x�n' end as �a��, " & _
                "round(sum(cast(s.Busr4 as float)*od.ORDER_QTY),0) as �~�n,round(sum(cast(s.Stdgrosswgt as float)*od.ORDER_QTY),0) as ���q " & _
                "from trp03w od " & _
                "inner join gv_skuxpack s on od.PRODUCT_NO=s.sku and s.storerkey = od.storerkey " & _
                "inner join trp02w o on od.RECEIPT_NO=o.RECEIPT_NO " & _
                "inner join TRP01M t1m on t1m.ConsigneeKey = o.ConsigneeKey " & _
                "where t1m.AREA_CODE in ('C','S','S1')  and  convert(char(8),o.ARRIVE_DATE,112) " & str_Where & "  " & _
                "group by t1m.AREA_CODE " & _
                "Union " & _
                "select  case when  left(t1m.SHORT_NAME,2)  in ('���F','�¦�','����')  then '�̪F�g�P'  when  left(t1m.SHORT_NAME,2)  in ('����','����','�p��')  then '�x�n�g�P' else '���ϸg�P' end as �a��, " & _
                "round(sum(cast(s.Busr4 as float)*od.ORDER_QTY),0) as �~�n,round(sum(cast(s.Stdgrosswgt as float)*od.ORDER_QTY),0) as ���q " & _
                "from trp03w od " & _
                "inner join gv_skuxpack s on od.PRODUCT_NO=s.sku and s.storerkey = od.storerkey " & _
                "inner join trp02w o on od.RECEIPT_NO=o.RECEIPT_NO " & _
                "inner join TRP01M t1m on t1m.ConsigneeKey = o.ConsigneeKey " & _
                "where t1m.AREA_CODE in ('w') " & _
                "and  left(t1m.SHORT_NAME,2)  in ('�͸�','����','����','���F','�¦�','����','�i��','�p��','�j�w','�O��') " & _
                "and  convert(char(8),o.ARRIVE_DATE,112) " & str_Where & "  " & _
                "group by case when  left(t1m.SHORT_NAME,2)  in ('���F','�¦�','����')  then '�̪F�g�P'  when  left(t1m.SHORT_NAME,2)  in ('����','����','�p��')  then '�x�n�g�P' else '���ϸg�P' end"
    End If
    On Error GoTo err_Handle
    
    '�q����R
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        Screen.MousePointer = vbDefault
        msg_text = "�d�ߵ��G�G�L�ŦX�j�M���󤧱ƨ����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Set gd_Tab3_OrderSum.DataSource = Nothing
        Exit Sub
    End If
    Call Replication_Recordset(tmp_Rs, rs_Tab3_OrderSum)
    Set gd_Tab3_OrderSum.DataSource = rs_Tab3_OrderSum
    tmp_Rs.Close
    With gd_Tab3_OrderSum
        .ColumnHeaders = True         '���D�����
        .RowHeight = 250
        .Columns(0).Width = 500       '�Ǹ�
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 1000       '�a��
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1200       '�~�n
        .Columns(2).Alignment = dbgRight
        .Columns(3).Width = 1200      '���q
        .Columns(3).Alignment = dbgRight
    End With
    DoEvents: DoEvents
    '���ƨ��q����R
    Call DB_CheckConnectStatus
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQl1, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        Screen.MousePointer = vbDefault
        msg_text = "�d�ߵ��G�G�L���ƨ����q��"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        Set gd_Tab3_Trp02wSum.DataSource = Nothing
        Exit Sub
    End If
    Call Replication_Recordset(tmp_Rs, rs_Tab3_Trp02wSum)
    Set gd_Tab3_Trp02wSum.DataSource = rs_Tab3_Trp02wSum
    tmp_Rs.Close
    With gd_Tab3_Trp02wSum
        .ColumnHeaders = True         '���D�����
        .RowHeight = 250
        .Columns(0).Width = 500       '�Ǹ�
        .Columns(0).Alignment = dbgLeft
        .Columns(1).Width = 1000       '�a��
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1200       '�~�n
        .Columns(2).Alignment = dbgRight
        .Columns(3).Width = 1200      '���q
        .Columns(3).Alignment = dbgRight
    End With
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    
    cmd_Tab3_Cancel.Enabled = True
    cmd_Tab3_Excel.Enabled = True
    Exit Sub
    
err_Handle:
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "�ƨ��@�~-�w����R", Me.Caption, "cmd_Tab3_Query_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDeliveryDateFix_Click()
    If rs_Tab1_Route.EOF Then Exit Sub
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "select * from trp02t t2 where len(rtrim(t2.description)) > 0 and t2.route_no = '" & rs_Tab1_Route("���u�s��") & "' "
    tmp_Rs.Open str_SQL, cn
    If tmp_Rs.EOF Then tmp_Rs.Close: Exit Sub
    tmp_Rs.Close
    strDeliveryDateFiRouteNo = rs_Tab1_Route.Fields("���u�s��").Value
    frm_DeliveryDateFix.Show vbModal
    
End Sub



Private Sub DateS_Click()
    
    If Trim(DateS.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(DateS.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(DateS.Text, 4) & "/" & Mid(DateS.Text, 5, 2) & "/" & Right(DateS.Text, 2))
        End If
    End If
    mvDate.Left = fam_SelectedOrders.Left + DateS.Left
    mvDate.Top = fam_SelectedOrders.Top + DateS.Top + DateS.Height
    mvDate.Visible = True
End Sub

Private Sub dg_Tab0_SelectedOrders_HeadClick(ByVal ColIndex As Integer)
    '�H�ƹ��I�� [�w����q��] dg_Tab0_SelectedOrders �����D�ϡG�Ƨ������
    Dim OrderFieldName As String
    If TypeName(rs_Tab0_SelectedOrders) <> "Nothing" Then
        OrderFieldName = "[" & dg_Tab0_SelectedOrders.Columns(ColIndex).Caption & "]"
        If strOrder = "ASC" Then
            strOrder = "DESC"
            rs_Tab0_SelectedOrders.Sort = OrderFieldName & " DESC "
        Else
            strOrder = "ASC"
            rs_Tab0_SelectedOrders.Sort = OrderFieldName & " ASC "
        End If
    End If
End Sub

Private Sub dg_Tab0_SelectedOrders_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    '�ƨ��@�~ >> �w����q�� DBGrid
    If blTab0SelectedOrderEventEnable Then
        With dg_Tab0_SelectedOrders
            '�ϥ���ܿ������ƦC
            If Not rs_Tab0_SelectedOrders.EOF Then
                dg_Tab0_SelectedOrders.SelBookmarks.Add rs_Tab0_SelectedOrders.Bookmark
            End If
        End With
    End If
    
'If Not rs_Tab0_SelectedOrders.EOF Then
'    txt_Tab0_DockNo = rs_Tab0_SelectedOrders("area")
'Else
'    txt_Tab0_DockNo = ""
'End If

End Sub

Private Sub dg_Tab1_Route_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    '���u�s���C��G�����
    If blTab1RouteEventEnable Then
        If Not rs_Tab1_Route.EOF Then
            dg_Tab1_Route.SelBookmarks.Add rs_Tab1_Route.Bookmark
            rs_Tab1_RouteOrders.Filter = " ���u�s�� = '" & rs_Tab1_Route.Fields("���u�s��").Value & "' "
        End If
    End If
End Sub

Private Sub dg_Tab2_ReservedOrders_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim objDataGrid As Object: Set objDataGrid = dg_Tab2_ReservedOrders
If Len(objDataGrid.Columns(ColIndex).DataField) = 0 Or objDataGrid.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, "�O�d�q��" & objDataGrid.Name, objDataGrid.Columns(ColIndex).DataField, objDataGrid.Columns(ColIndex).Width
End Sub

Private Sub dg_Tab2_ReservedOrders_HeadClick(ByVal ColIndex As Integer)
    '�H�ƹ��I�� [�O�d�q��] dg_Tab2_ReservedOrder �����D�ϡG�Ƨ������
    Dim OrderFieldName As String
    If TypeName(rs_Tab2_ReservedOrders) <> "Nothing" Then
        OrderFieldName = "[" & dg_Tab2_ReservedOrders.Columns(ColIndex).Caption & "]"
        If strOrder = "ASC" Then
            strOrder = "DESC"
            rs_Tab2_ReservedOrders.Sort = OrderFieldName & " DESC "
        Else
            strOrder = "ASC"
            rs_Tab2_ReservedOrders.Sort = OrderFieldName & " ASC "
        End If
    End If
End Sub

Private Sub dg_Tab2_ReservedOrders_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    '�ƨ��@�~ >> �O�d�q�� DBGrid
    If rs_Tab2_ReservedOrders.EOF Then Exit Sub
    If blTab2ReservedEventEnable Then
        With dg_Tab2_ReservedOrders
            '�I�@�U����A���I�h [����]
            If Trim(rs_Tab2_ReservedOrders.Fields(1).Value) = "" Then
                rs_Tab2_ReservedOrders.Fields(1).Value = "V"
            Else
                rs_Tab2_ReservedOrders.Fields(1).Value = " "
            End If
            '�ϥ���ܿ������ƦC
            If Not rs_Tab2_ReservedOrders.EOF Then
                dg_Tab2_ReservedOrders.SelBookmarks.Add rs_Tab2_ReservedOrders.Bookmark
            End If
        End With
    End If
End Sub

Private Sub dg_TRP02W_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim objDataGrid As Object: Set objDataGrid = dg_TRP02W
If Len(objDataGrid.Columns(ColIndex).DataField) = 0 Or objDataGrid.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, "�ݱƨ��q��" & objDataGrid.Name, objDataGrid.Columns(ColIndex).DataField, objDataGrid.Columns(ColIndex).Width
End Sub

Private Sub dg_TRP02W_DblClick()
If rs_TRP02W Is Nothing Then Exit Sub
If rs_TRP02W.RecordCount = 0 Then Exit Sub
If Len(RTrim(txtReceipt_no)) = 0 Then Exit Sub
frm_OrderDetail.Show vbModal
End Sub

Private Sub dg_TRP02W_HeadClick(ByVal ColIndex As Integer)
    '�H�ƹ��I�� [�ݱƨ��q��] dg_TRP02W �����D�ϡG�Ƨ������
    Dim OrderFieldName As String
    If TypeName(rs_TRP02W) <> "Nothing" Then
        '�קK���� [���] ���ʧ@
        blTRP02WEventEnable = False
        OrderFieldName = "[" & dg_TRP02W.Columns(ColIndex).Caption & "]"
        If strOrder = "ASC" Then
            strOrder = "DESC"
            rs_TRP02W.Sort = OrderFieldName & " DESC "
            strSourceOrderBy = OrderFieldName & " desc "
        Else
            strOrder = "ASC"
            rs_TRP02W.Sort = OrderFieldName & " ASC "
            strSourceOrderBy = OrderFieldName & " asc "
        End If
        blTRP02WEventEnable = True
    End If
End Sub

Private Sub dg_TRP02W_RowColChange(LastRow As Variant, ByVal LastCkmol As Integer)
    On Error GoTo err_Handle
    '�ƨ��@�~ >> �ݱƨ��q�� DBGrid
    If blTRP02WEventEnable Then
    txtReceipt_no = rs_TRP02W("receipt_no")
        With dg_TRP02W
            '�I��Y��ܿ���A��������H��L Button �M���B�z�G�]�������������K
            If Trim(rs_TRP02W.Fields(1).Value) = "" Then
                            
'                '���^��WMS�q��L�k�[�J�w�^�Ǹ��s
'                If rs_Tab0_SelectedOrders.Fields("EXE�^��") = "�w�^��" And rs_TRP02W("EXE�^��") = "�s�ظ��s" Then MsgBox "���^��WMS�q��A�L�k�[�J�w�^�Ǹ��s!!", 64, "�q����": Exit Sub

                rs_TRP02W.Fields(1).Value = "V"
                dbSelectedCount = dbSelectedCount + 1
                '����p�p��s
                dbsrcSelected_OTqty = dbsrcSelected_OTqty + rs_TRP02W.Fields("���").Value
                dbsrcSelected_Case = dbsrcSelected_Case + rs_TRP02W.Fields("�c��").Value
                dbsrcSelected_Pallet = dbsrcSelected_Pallet + rs_TRP02W.Fields("�O��").Value
                dbsrcSelected_Volumn = dbsrcSelected_Volumn + rs_TRP02W.Fields("���n").Value
                dbsrcSelected_Weight = dbsrcSelected_Weight + rs_TRP02W.Fields("���q").Value
                txt_Tab0_srcSelected_OTqty.Text = dbsrcSelected_OTqty
                txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
                txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
            Else
                dbSelectedCount = dbSelectedCount - 1
                rs_TRP02W.Fields(1).Value = " "
                '����p�p��s
                If dbSelectedCount = 0 Then
                    dbsrcSelected_OTqty = 0
                    dbsrcSelected_Case = 0
                    dbsrcSelected_Pallet = 0
                    dbsrcSelected_Volumn = 0
                    dbsrcSelected_Weight = 0
                    txt_Tab0_srcSelected_OTqty.Text = 0
                    txt_Tab0_srcSelected_Case.Text = 0: txt_Tab0_srcSelected_Pallet.Text = 0
                    txt_Tab0_srcSelected_Volumn.Text = 0: txt_Tab0_srcSelected_Weight.Text = 0
                Else
                    dbsrcSelected_OTqty = dbsrcSelected_OTqty - rs_TRP02W.Fields("���").Value
                    dbsrcSelected_Case = dbsrcSelected_Case - rs_TRP02W.Fields("�c��").Value
                    dbsrcSelected_Pallet = dbsrcSelected_Pallet - rs_TRP02W.Fields("�O��").Value
                    dbsrcSelected_Volumn = dbsrcSelected_Volumn - rs_TRP02W.Fields("���n").Value
                    dbsrcSelected_Weight = dbsrcSelected_Weight - rs_TRP02W.Fields("���q").Value
                    txt_Tab0_srcSelected_OTqty.Text = dbsrcSelected_OTqty
                    txt_Tab0_srcSelected_Case.Text = dbsrcSelected_Case: txt_Tab0_srcSelected_Pallet.Text = dbsrcSelected_Pallet
                    txt_Tab0_srcSelected_Volumn.Text = dbsrcSelected_Volumn: txt_Tab0_srcSelected_Weight.Text = dbsrcSelected_Weight
                End If
            End If
            '�ϥ���ܿ������ƦC
            If Not rs_TRP02W.EOF Then
                dg_TRP02W.SelBookmarks.Add rs_TRP02W.Bookmark
            End If
        End With
    End If
    Exit Sub
err_Handle:
End Sub

Private Sub cmd_Tab0_srcOrderQuery_Click()
    '�ƨ��@�~ >> �ݱƨ��q��j�M
    If rs_TRP02W Is Nothing Then Exit Sub
    If rs_TRP02W.RecordCount = 0 Then Exit Sub
    
    strFormName_FilterAndSort = Me.Name
    strRSName_FilterAndSort = "rs_TRP02W"
    
    If ShowForm_RS_FilterAndSort(rs_TRP02W, "�ݱƨ��q��", Me.Tag) = False Then
        MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    Me.WindowState = 2

End Sub

Private Sub Form_Activate()
    '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
    Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
    msg_title = "�ƨ��@�~"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '�d�I��Ӫ����L����ƥ�
    '�γ~�G�ϥΪ̫��U Esc �h���Ǧ^�����ơA�B��������������
    If KeyCode = vbKeyEscape Then
        mvDate.Visible = False
    End If
End Sub

Private Sub Form_Load()
    '�]�w Form �j�p�B��m
    dbsrcFormHeight = 7140
    dbsrcFormWidth = 11475
    Me.Height = 7650: Me.Width = 11600
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Left = 200
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300
    
    '�ƨ��@�~�G�ݱƨ��q��
    Call CreateRS_Tab0_TRP02W
    strSourceFilter = adFilterNone
    strSourceOrderBy = " �s�� asc "
    
    '�ƨ��@�~�G�w������ݱƨ��q��C�� DBGrid �榡�]�w
    Call CreateRS_Tab0_SelectedOrders
    
    '�w���ͤ����u�s���C��
    Call CreateRS_Tab1_Route
    Call CreateRS_Tab1_RouteOrders
    
    '�O�d�q��C��
    Call CreateRS_Tab2_ReservedOrders
    blTab2ReservedEventEnable = True
    SSTab1.Tab = 0
    
    Dim rsTmp As New ADODB.Recordset
    With rsTmp
        .CursorLocation = 3
        '�f�D
        .Open "select distinct storerkey from " & strWMSDB & "..storer (nolock) where type = 1 order by storerkey ", cn
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Storerkey.AddItem RTrim(rsTmp("storerkey"))
            rsTmp.MoveNext
        Loop
        .Close
        
        '�ϽX
        .Open "select distinct area_code from trp01m(nolock) where len(isnull(area_code,'')) > 0 order by area_code ", cn
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            Area_Code.AddItem RTrim(rsTmp("area_code"))
            rsTmp.MoveNext
        Loop
        .Close
    End With
    
    DateS = Format(Now() + 1, "YYYYMMDD")
    DateE = Format(Now() + 1, "YYYYMMDD")
    
End Sub

Private Sub Form_Resize()
    '�����j�p�ܰ�
    If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
    If Me.ScaleHeight < dbsrcFormHeight Then
        '�ܤp
        SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
        SSTab1.Width = (SSTab1.Width - (dbsrcFormWidth - Me.ScaleWidth))
        
        fam_SelectedOrders.Width = fam_SelectedOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        dg_Tab0_SelectedOrders.Width = dg_Tab0_SelectedOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        fam_SrcOrders.Height = fam_SrcOrders.Height - (dbsrcFormHeight - Me.ScaleHeight)
        fam_SrcOrders.Width = fam_SrcOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        dg_TRP02W.Height = dg_TRP02W.Height - (dbsrcFormHeight - Me.ScaleHeight)
        dg_TRP02W.Width = dg_TRP02W.Width - (dbsrcFormWidth - Me.ScaleWidth)
        
        Frame3.Left = Frame3.Left - (dbsrcFormWidth - Me.ScaleWidth)
        Frame4.Left = Frame4.Left - (dbsrcFormWidth - Me.ScaleWidth)
        dg_Tab1_Route.Width = dg_Tab1_Route.Width - (dbsrcFormWidth - Me.ScaleWidth)
        
        dg_Tab1_RouteOrders.Height = dg_Tab1_RouteOrders.Height - (dbsrcFormHeight - Me.ScaleHeight)
        dg_Tab1_RouteOrders.Width = dg_Tab1_RouteOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        
        dg_Tab2_ReservedOrders.Height = dg_Tab2_ReservedOrders.Height - (dbsrcFormHeight - Me.ScaleHeight)
        dg_Tab2_ReservedOrders.Width = dg_Tab2_ReservedOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_Remove.Left = cmd_Tab2_Remove.Left - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_FilterAndSort.Left = cmd_Tab2_FilterAndSort.Left - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_Reset.Left = cmd_Tab2_Reset.Left - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_ShowAll.Left = cmd_Tab2_ShowAll.Left - (dbsrcFormWidth - Me.ScaleWidth)
        cmd_Tab2_Delete.Left = cmd_Tab2_Delete.Left - (dbsrcFormWidth - Me.ScaleWidth)
        dbsrcFormHeight = Me.ScaleHeight
        dbsrcFormWidth = Me.ScaleWidth
    Else
       SSTab1.Height = (SSTab1.Height + (Me.ScaleHeight - dbsrcFormHeight))
       SSTab1.Width = (SSTab1.Width + (Me.ScaleWidth - dbsrcFormWidth))
       
       fam_SelectedOrders.Width = fam_SelectedOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       dg_Tab0_SelectedOrders.Width = dg_Tab0_SelectedOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       fam_SrcOrders.Width = fam_SrcOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       fam_SrcOrders.Height = fam_SrcOrders.Height + (Me.ScaleHeight - dbsrcFormHeight)
       dg_TRP02W.Height = dg_TRP02W.Height + (Me.ScaleHeight - dbsrcFormHeight)
       dg_TRP02W.Width = dg_TRP02W.Width + (Me.ScaleWidth - dbsrcFormWidth)
       
       Frame3.Left = Frame3.Left + (Me.ScaleWidth - dbsrcFormWidth)
       dg_Tab1_Route.Width = dg_Tab1_Route.Width + (Me.ScaleWidth - dbsrcFormWidth)
       Frame4.Left = Frame4.Left + (Me.ScaleWidth - dbsrcFormWidth)
       
       dg_Tab1_RouteOrders.Height = dg_Tab1_RouteOrders.Height + (Me.ScaleHeight - dbsrcFormHeight)
       dg_Tab1_RouteOrders.Width = dg_Tab1_RouteOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       
       dg_Tab2_ReservedOrders.Height = dg_Tab2_ReservedOrders.Height + (Me.ScaleHeight - dbsrcFormHeight)
       dg_Tab2_ReservedOrders.Width = dg_Tab2_ReservedOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_Remove.Left = cmd_Tab2_Remove.Left + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_FilterAndSort.Left = cmd_Tab2_FilterAndSort.Left + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_Reset.Left = cmd_Tab2_Reset.Left + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_ShowAll.Left = cmd_Tab2_ShowAll.Left + (Me.ScaleWidth - dbsrcFormWidth)
       cmd_Tab2_Delete.Left = cmd_Tab2_Delete.Left + (Me.ScaleWidth - dbsrcFormWidth)
       
       dbsrcFormHeight = Me.ScaleHeight
       dbsrcFormWidth = Me.ScaleWidth
    End If
End Sub

Private Sub Form_Terminate()
    '��s Menu [����]��[�w�}�����M��]
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
    Set frm_OP_TRPPlan = Nothing
End Sub

Private Sub CreateRS_Tab0_TRP02W()
    '�ƨ��@�~�G�ݱƨ��q��
    Call ReDim_Recordset(rs_TRP02W)
    With rs_TRP02W
        .Fields.Append "�s��", adDouble
        .Fields.Append "��", adVarChar, 2
        .Fields.Append "�e�f��", adVarChar, 10
        .Fields.Append "�q��s��", adVarChar, 60
        .Fields.Append "���", adDouble
        .Fields.Append "�c��", adDouble
        .Fields.Append "�O��", adDouble
        .Fields.Append "���n", adDouble
        .Fields.Append "���q", adDouble
        .Fields.Append "�Ȥ�s��", adVarChar, 30
        .Fields.Append "ZIP", adVarChar, 10
        .Fields.Append "�Ȥ�²��", adVarChar, 60
        .Fields.Append "Area", adVarChar, 10
        .Fields.Append "���A", adVarChar, 10
        .Fields.Append "�B�e�a�}", adVarChar, 120
        .Fields.Append "�q��Ƶ�", adVarChar, 300
        .Fields.Append "��s������", adVarChar, 4
        .Fields.Append "�t�e�ܧO", adVarChar, 20
        .Fields.Append "����", adVarChar, 10
        .Fields.Append "�S��ݨD1", adVarChar, 60
        .Fields.Append "�S��ݨD2", adVarChar, 60
        .Fields.Append "���", adVarChar, 10
        .Fields.Append "�M��", adVarChar, 10
        .Fields.Append "�N��", adVarChar, 10
        .Fields.Append "Receipt_No", adVarChar, 10
        .Fields.Append "C_Receipt_No", adVarChar, 10
        .Fields.Append "�f�D�渹", adVarChar, 40
        .Fields.Append "EXE�^��", adVarChar, 20
        .Fields.Append "�Ȥ�W��", adVarChar, 120
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open    '���ݳs������
    End With
    Set dg_TRP02W.DataSource = rs_TRP02W
    '�]�w������
    With dg_TRP02W
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .RowHeight = 250
'        .Columns(0).Width = 500         '�Ǹ�
'        .Columns(0).Alignment = dbgCenter
'        .Columns(1).Width = 300         '����ѧO���
'        .Columns(1).Alignment = dbgCenter
'        .Columns(2).Width = 800         '�e�f��
'        .Columns(2).Alignment = dbgCenter
'        .Columns(3).Width = 2500        '�q��s���G�q��s��+�f�D�渹+�f�D
'        .Columns(3).Alignment = dbgLeft
'        .Columns(4).Width = 600         '�c��
'        .Columns(4).Alignment = dbgRight
'        .Columns(5).Width = 600         '�O��
'        .Columns(5).Alignment = dbgRight
'        .Columns(6).Width = 600         '���n
'        .Columns(6).Alignment = dbgRight
'        .Columns(7).Width = 600         '���q
'        .Columns(7).Alignment = dbgRight
'        .Columns(8).Width = 1100        '�Ȥ�s��
'        .Columns(8).Alignment = dbgLeft
'        .Columns(9).Width = 400         'zip
'        .Columns(9).Alignment = dbgCenter
'        .Columns(10).Width = 1000       '�Ȥ�²��
'        .Columns(10).Alignment = dbgLeft
'        .Columns(11).Width = 450        'Area_Code
'        .Columns(11).Alignment = dbgCenter
'        .Columns(12).Width = 450        '���A�GPriority
'        .Columns(12).Alignment = dbgCenter
'        .Columns(13).Width = 2500       '�B�e�a�}
'        .Columns(13).Alignment = dbgLeft
'        .Columns(14).Width = 1400       '�q��Ƶ�
'        .Columns(14).Alignment = dbgLeft
'        .Columns(15).Width = 1000       '��s������
'        .Columns(15).Alignment = dbgLeft
'        .Columns(16).Width = 500        '����
'        .Columns(16).Alignment = dbgCenter
'        .Columns(17).Width = 1500       '�S��ݨD1
'        .Columns(17).Alignment = dbgLeft
'        .Columns(18).Width = 1500       '�S��ݨD2
'        .Columns(18).Alignment = dbgLeft
'        .Columns(19).Width = 500        '���
'        .Columns(19).Alignment = dbgCenter
'        .Columns(20).Width = 500        '�M��
'        .Columns(20).Alignment = dbgCenter
'        .Columns(21).Width = 500        '�N��
'        .Columns(21).Alignment = dbgCenter
'        .Columns(22).Width = 1100       'Receipt_No
'        .Columns(22).Alignment = dbgLeft
'        .Columns(23).Width = 1100       'C_Receipt_No
'        .Columns(23).Alignment = dbgLeft
'        .Columns(24).Width = 900        '�f�D�渹
'        .Columns(24).Alignment = dbgLeft
'        .Columns(25).Width = 1500       'EXE�^��
'        .Columns(25).Alignment = dbgLeft
'        .Columns(26).Width = 1500       '�Ȥ�W��
'        .Columns(26).Alignment = dbgLeft
    End With
    SetDataGridColWidth "�ݱƨ��q��", dg_TRP02W
End Sub

Private Sub CreateRS_Tab0_SelectedOrders()
    '�ƨ��@�~�G�w������ݱƨ��q��C��
    Call ReDim_Recordset(rs_Tab0_SelectedOrders)
    With rs_Tab0_SelectedOrders
        .Fields.Append "�s��", adDouble
        .Fields.Append "�e�f��", adVarChar, 20
        .Fields.Append "�q��s��", adVarChar, 60
        .Fields.Append "ZIP", adVarChar, 10
        .Fields.Append "Area", adVarChar, 10
        .Fields.Append "���", adVarChar, 10
        .Fields.Append "���A", adVarChar, 20
        .Fields.Append "�Ȥ�²��", adVarChar, 120
        .Fields.Append "���", adDouble
        .Fields.Append "�c��", adDouble
        .Fields.Append "�O��", adDouble
        .Fields.Append "���n", adDouble
        .Fields.Append "���q", adDouble
        .Fields.Append "����", adVarChar, 10
        .Fields.Append "�q��Ƶ�", adVarChar, 300
        .Fields.Append "�S��ݨD1", adVarChar, 60
        .Fields.Append "�S��ݨD2", adVarChar, 60
        .Fields.Append "Receipt_No", adVarChar, 10
        .Fields.Append "EXE�^��", adVarChar, 20
        .Fields.Append "�Ȥ�W��", adVarChar, 120
        .Fields.Append "�B�e�a�}", adVarChar, 120
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open    '���ݳs������
    End With
    Set dg_Tab0_SelectedOrders.DataSource = rs_Tab0_SelectedOrders
    '�]�w������
    With dg_Tab0_SelectedOrders
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .Columns(0).Width = 500        '�s��
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 800         '�e�f��
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 2500        '�q��s��
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 400         'ZIP
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 450         'Area
        .Columns(4).Alignment = dbgCenter
        .Columns(5).Width = 450         '���A�GOrders.Priority
        .Columns(5).Alignment = dbgCenter
        .Columns(6).Width = 450         '���A�GOrders.Priority
        .Columns(6).Alignment = dbgCenter
        .Columns(7).Width = 1000        '�Ȥ�²��
        .Columns(7).Alignment = dbgLeft
        .Columns(8).Width = 600         '���
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 600         '�c��
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 600         '�O��
        .Columns(10).Alignment = dbgRight
        .Columns(11).Width = 600         '���n
        .Columns(11).Alignment = dbgRight
        .Columns(12).Width = 600        '���q
        .Columns(12).Alignment = dbgRight
        .Columns(13).Width = 450        '����
        .Columns(13).Alignment = dbgLeft
        .Columns(14).Width = 1200       '�q��Ƶ�
        .Columns(14).Alignment = dbgLeft
        .Columns(15).Width = 1500       '�S��ݨD-1
        .Columns(15).Alignment = dbgLeft
        .Columns(16).Width = 1500       '�S��ݨD-2
        .Columns(16).Alignment = dbgLeft
        .Columns(17).Width = 1000       'Receipt_No
        .Columns(17).Alignment = dbgLeft
        .Columns(18).Width = 1000       'EXE�^��
        .Columns(18).Alignment = dbgLeft
        .Columns(19).Width = 1500       '�Ȥ�W��
        .Columns(19).Alignment = dbgLeft
    End With
End Sub

Private Sub CreateRS_Tab1_Route()
    '�ƨ��@�~�G�w�s�������u�s���C��
    Call ReDim_Recordset(rs_Tab1_Route)
    With rs_Tab1_Route
        .Fields.Append "�s��", adDouble
        .Fields.Append "���u�s��", adVarChar, 10
        .Fields.Append "�X�����", adVarChar, 8
        .Fields.Append "���P���X", adVarChar, 10
        .Fields.Append "����", adDouble
        .Fields.Append "�r�p�H", adVarChar, 20
        .Fields.Append "���", adDouble
        .Fields.Append "�c��", adDouble
        .Fields.Append "�O��", adDouble
        .Fields.Append "���n", adDouble
        .Fields.Append "���q", adDouble
        .Fields.Append "����", adVarChar, 10
        .Fields.Append "�X�Y�Ȧs", adVarChar, 10
        .Fields.Append "�w�p������", adVarChar, 8
        .Fields.Append "�w�p����ɶ�", adVarChar, 4
        .Fields.Append "EXE�^��", adVarChar, 20
        .Fields.Append "�ƨ���", adVarChar, 30
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open    '���ݳs������
    End With
    Set dg_Tab1_Route.DataSource = rs_Tab1_Route
    '�]�w������
    With dg_Tab1_Route
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .Columns(0).Width = 500         '�s��
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1000        '���u�s��
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 900         '�X�����
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 850         '���P���X
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Width = 500         '����
        .Columns(4).Alignment = dbgCenter
        .Columns(5).Width = 900         '�r�p�H
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 700         '���
        .Columns(6).Alignment = dbgRight
        .Columns(7).Width = 700         '�c��
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 700         '�O��
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 700         '���n
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 700         '���q
        .Columns(10).Alignment = dbgRight
        .Columns(11).Width = 450        '����
        .Columns(11).Alignment = dbgCenter
        .Columns(12).Width = 1000       '�X�Y�Ȧs
        .Columns(12).Alignment = dbgLeft
        .Columns(13).Width = 1400       '�w�p����������
        .Columns(13).Alignment = dbgLeft
        .Columns(14).Width = 1400       '�w�p��������ɶ�
        .Columns(14).Alignment = dbgLeft
        .Columns(15).Width = 900        'EXE �^�Ǫ��A
        .Columns(15).Alignment = dbgLeft
        .Columns(16).Width = 1200       '�ƨ���
        .Columns(16).Alignment = dbgLeft
    End With
End Sub

Private Sub CreateRS_Tab1_RouteOrders()
    '�ƨ��@�~�G�w�s�����s���q��C��
    Call ReDim_Recordset(rs_Tab1_RouteOrders)
    With rs_Tab1_RouteOrders
        .Fields.Append "�s��", adDouble
        .Fields.Append "���u�s��", adVarChar, 10
        .Fields.Append "�e�f��", adVarChar, 20
        .Fields.Append "�q��s��", adVarChar, 60
        .Fields.Append "ZIP", adVarChar, 10
        .Fields.Append "�Ȥ�W��", adVarChar, 120
        .Fields.Append "�a�}", adVarChar, 300
        .Fields.Append "���", adDouble
        .Fields.Append "�c��", adDouble
        .Fields.Append "�O��", adDouble
        .Fields.Append "���n", adDouble
        .Fields.Append "���q", adDouble
        .Fields.Append "�q��Ƶ�", adVarChar, 300
        .Fields.Append "����", adVarChar, 10
        .Fields.Append "�S��ݨD1", adVarChar, 60
        .Fields.Append "�S��ݨD2", adVarChar, 60
        .Fields.Append "Receipt_No", adVarChar, 60
        .Fields.Append "EXE�^��", adVarChar, 20
        .Fields.Append "Area", adVarChar, 10
        .Fields.Append "���A", adVarChar, 10
        .Fields.Append "�Ȥ�²��", adVarChar, 255
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open    '���ݳs������
    End With
    Set dg_Tab1_RouteOrders.DataSource = rs_Tab1_RouteOrders
    '�]�w������
    With dg_Tab1_RouteOrders
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .Columns(0).Width = 500         '�s��
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1050        '���u�s��
        .Columns(1).Alignment = dbgCenter
        .Columns(2).Width = 900         '�e�f��
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Width = 2150        '�q��s��
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 400         'ZIP
        .Columns(4).Alignment = dbgCenter
        .Columns(5).Width = 1500        '�Ȥ�W��
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 1500        '�Ȥ�W��
        .Columns(6).Alignment = dbgLeft
        .Columns(7).Width = 700         '���
        .Columns(7).Alignment = dbgRight
        .Columns(8).Width = 700         '�c��
        .Columns(8).Alignment = dbgRight
        .Columns(9).Width = 700         '�O��
        .Columns(9).Alignment = dbgRight
        .Columns(10).Width = 700         '���n
        .Columns(10).Alignment = dbgRight
        .Columns(11).Width = 700         '���q
        .Columns(11).Alignment = dbgRight
        .Columns(12).Width = 1500       '�q��Ƶ�
        .Columns(12).Alignment = dbgLeft
        .Columns(13).Width = 1200       '����
        .Columns(13).Alignment = dbgLeft
        .Columns(14).Width = 1500       '�S��ݨD-1
        .Columns(14).Alignment = dbgLeft
        .Columns(15).Width = 1500       '�S��ݨD-2
        .Columns(15).Alignment = dbgLeft
        .Columns(16).Width = 1100       'Receipt_No
        .Columns(16).Alignment = dbgLeft
        .Columns(17).Width = 1100       'EXE�^��
        .Columns(17).Alignment = dbgLeft
        .Columns(18).Width = 450        'Area
        .Columns(18).Alignment = dbgCenter
        .Columns(19).Width = 450        '���A
        .Columns(19).Alignment = dbgCenter
        .Columns(20).Width = 1100       '�Ȥ�²��
        .Columns(20).Alignment = dbgLeft
    End With
End Sub

Private Sub CreateRS_Tab2_ReservedOrders()
    '�ƨ��@�~�G�O�d�q��
    Call ReDim_Recordset(rs_Tab2_ReservedOrders)
    With rs_Tab2_ReservedOrders
         .Fields.Append "�s��", adDouble
         .Fields.Append "��", adVarChar, 2
         .Fields.Append "�e�f��", adVarChar, 10
         .Fields.Append "�q��s��", adVarChar, 60
         .Fields.Append "���", adDouble
         .Fields.Append "�c��", adDouble
         .Fields.Append "�O��", adDouble
         .Fields.Append "���n", adDouble
         .Fields.Append "���q", adDouble
         .Fields.Append "�Ȥ�s��", adVarChar, 30
         .Fields.Append "ZIP", adVarChar, 10
         .Fields.Append "�Ȥ�²��", adVarChar, 60
         .Fields.Append "Area", adVarChar, 10
         .Fields.Append "���A", adVarChar, 10
         .Fields.Append "�B�e�a�}", adVarChar, 120
         .Fields.Append "�q��Ƶ�", adVarChar, 300
         .Fields.Append "��s������", adVarChar, 4
         .Fields.Append "�t�e�ܧO", adVarChar, 20
         .Fields.Append "����", adVarChar, 10
         .Fields.Append "�S��ݨD1", adVarChar, 60
         .Fields.Append "�S��ݨD2", adVarChar, 60
         .Fields.Append "���", adVarChar, 10
         .Fields.Append "�M��", adVarChar, 10
         .Fields.Append "�N��", adVarChar, 10
         .Fields.Append "TMS�渹", adVarChar, 10
         .Fields.Append "C_RECEIPT_NO", adVarChar, 10
         .Fields.Append "�f�D�渹", adVarChar, 40
         .Fields.Append "EXE�^��", adVarChar, 20
         .Fields.Append "�Ȥ�W��", adVarChar, 120
         .CursorType = adOpenStatic
         .LockType = adLockOptimistic
         .Open    '���ݳs������
    End With
    Set dg_Tab2_ReservedOrders.DataSource = rs_Tab2_ReservedOrders
    '�]�w������
    With dg_Tab2_ReservedOrders
        .ColumnHeaders = True           '�M�w�O�_�b DataGrid �������ܸ�Ʀ�歺�C
        .HeadLines = 1.5                '��ܦb DataGrid �������Ʀ�歺������r��ơC
        .RowDividerStyle = dbgRaised    'DataGrid �����ƦC�����ؽu�˦��C
        .RowHeight = 250                '�]�wDataGrid ������Ҧ���ƦC����
        .RowHeight = 250
'        .Columns(0).Width = 500         '�Ǹ�
'        .Columns(0).Alignment = dbgCenter
'        .Columns(1).Width = 300         '����ѧO���
'        .Columns(1).Alignment = dbgCenter
'        .Columns(2).Width = 800         '�e�f��
'        .Columns(2).Alignment = dbgCenter
'        .Columns(3).Width = 2100        '�q��s���G�q��s��+�f�D�渹+�f�D
'        .Columns(3).Alignment = dbgLeft
'        .Columns(4).Width = 600         '�c��
'        .Columns(4).Alignment = dbgRight
'        .Columns(5).Width = 600         '�O��
'        .Columns(5).Alignment = dbgRight
'        .Columns(6).Width = 600         '���n
'        .Columns(6).Alignment = dbgRight
'        .Columns(7).Width = 600         '���q
'        .Columns(7).Alignment = dbgRight
'        .Columns(8).Width = 1100        '�Ȥ�s��
'        .Columns(8).Alignment = dbgLeft
'        .Columns(9).Width = 400         'zip
'        .Columns(9).Alignment = dbgCenter
'        .Columns(10).Width = 1000       '�Ȥ�²��
'        .Columns(10).Alignment = dbgLeft
'        .Columns(11).Width = 450        'Area_Code
'        .Columns(11).Alignment = dbgCenter
'        .Columns(12).Width = 450        '���A�GPriority
'        .Columns(12).Alignment = dbgCenter
'        .Columns(13).Width = 3000       '�B�e�a�}
'        .Columns(13).Alignment = dbgLeft
'        .Columns(14).Width = 1400       '�q��Ƶ�
'        .Columns(14).Alignment = dbgLeft
'        .Columns(15).Width = 1000       '��s������
'        .Columns(15).Alignment = dbgLeft
'        .Columns(16).Width = 500       '����
'        .Columns(16).Alignment = dbgLeft
'        .Columns(17).Width = 1500       '�S��ݨD1
'        .Columns(17).Alignment = dbgLeft
'        .Columns(18).Width = 1500       '�S��ݨD2
'        .Columns(18).Alignment = dbgLeft
'        .Columns(19).Width = 500        '���
'        .Columns(19).Alignment = dbgCenter
'        .Columns(20).Width = 500        '�M��
'        .Columns(20).Alignment = dbgCenter
'        .Columns(21).Width = 500        '�N��
'        .Columns(21).Alignment = dbgCenter
'        .Columns(22).Width = 1100       'Receipt_No
'        .Columns(22).Alignment = dbgLeft
'        .Columns(23).Width = 1100       'C_Receipt_No
'        .Columns(23).Alignment = dbgLeft
'        .Columns(24).Width = 900        '�f�D�渹
'        .Columns(24).Alignment = dbgLeft
'        .Columns(25).Width = 1500       'EXE�^��
'        .Columns(25).Alignment = dbgLeft
'        .Columns(26).Width = 1500       '�Ȥ�W��
'        .Columns(26).Alignment = dbgLeft
    End With
        SetDataGridColWidth "�O�d�q��", dg_Tab2_ReservedOrders
End Sub

Private Sub Calculate_SelectedOrders()
    '�@�~���e�G
    '1.�w��w����q��C��A�̭q��s�����s���� [�s��] ����
    '2.�p��w����q�椧�֭p���
    Dim dbSeqNo As Double
    dbSeqNo = 0
    txt_Tab0_Selected_OTqty.Text = ""
    txt_Tab0_Selected_Case.Text = ""
    txt_Tab0_Selected_Pallet.Text = ""
    txt_Tab0_Selected_Volumn.Text = ""
    txt_Tab0_Selected_Weight.Text = ""
    
    rs_Tab0_SelectedOrders.Filter = adFilterNone
    rs_Tab0_SelectedOrders.Sort = "Receipt_No asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    If Not rs_Tab0_SelectedOrders.EOF Then
       rs_Tab0_SelectedOrders.MoveFirst
    Else
        '�M�X�z�����A���L��ƪ̡A���� SubProgram ����
        Exit Sub
    End If
    Do While Not rs_Tab0_SelectedOrders.EOF
        dbSeqNo = dbSeqNo + 1
        rs_Tab0_SelectedOrders.Fields("�s��").Value = dbSeqNo
        txt_Tab0_Selected_OTqty.Text = Val(txt_Tab0_Selected_OTqty.Text) + rs_Tab0_SelectedOrders.Fields("���").Value
        txt_Tab0_Selected_Case.Text = Val(txt_Tab0_Selected_Case.Text) + rs_Tab0_SelectedOrders.Fields("�c��").Value
        txt_Tab0_Selected_Pallet.Text = Val(txt_Tab0_Selected_Pallet.Text) + rs_Tab0_SelectedOrders.Fields("�O��").Value
        txt_Tab0_Selected_Volumn.Text = Val(txt_Tab0_Selected_Volumn.Text) + rs_Tab0_SelectedOrders.Fields("���n").Value
        txt_Tab0_Selected_Weight.Text = Val(txt_Tab0_Selected_Weight.Text) + rs_Tab0_SelectedOrders.Fields("���q").Value
        rs_Tab0_SelectedOrders.MoveNext
    Loop
    rs_Tab0_SelectedOrders.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    If Not rs_Tab0_SelectedOrders.EOF Then rs_Tab0_SelectedOrders.MoveFirst
End Sub

Private Sub SelectedOrders_Removeto_TRP02W(ByVal strReceiptNo As String)
    '�N���w�� [�q��s��] �[�J [�ݱƨ��q��]
    blTRP02WEventEnable = False
    
    rs_TRP02W.Filter = adFilterNone
    rs_TRP02W.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    
    If rs_TRP02W.RecordCount > 0 Then
        rs_TRP02W.Filter = "Receipt_No = '" & strReceiptNo & "'"
        If Not rs_TRP02W.EOF Then
            '�q��s���w�s�b���ܡA���i��s�W�A�]����s
            rs_TRP02W.Filter = adFilterNone
            rs_TRP02W.Sort = "�s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
            blTRP02WEventEnable = True
            Exit Sub
        End If
    End If
    
    '���^�ݱƨ��q��
    If blRouteModify Then
        '�p�G�O�d�ߩ���ܤ����ĸ��s�A�]�w���s�����ѧO�X��
        blRouteChange = True
        '�g�Ѭd�߸��u�s���ұo���q����
        str_SQL = "Select �e�f��,�q��s��,���,�c��,�O��,���n,���q,�Ȥ�s��,ZIP,�Ȥ�W��,�B�e�a�},�q��Ƶ�,����,�S��ݨD1,�S��ݨD2,���,�M��,�N��,Receipt_No,c_receipt_no,�f�D�渹,EXE�^��,Area,�Ȥ�²��,��s������,�t�e�ܧO,���A " & _
                  "From TRPPlan_RouteQueryOrdersRemove Where Receipt_No = '" & strReceiptNo & "' Order by �q��s�� "
    Else
        str_SQL = "Select �e�f��,�q��s��,���,�c��,�O��,���n,���q,�Ȥ�s��,ZIP,�Ȥ�W��,�B�e�a�},�q��Ƶ�,����,�S��ݨD1,�S��ݨD2,���,�M��,�N��,Receipt_No,c_receipt_no,�f�D�渹,EXE�^��,Area,�Ȥ�²��,���A,��s������,�t�e�ܧO " & _
                  "From TRPPlan_SourceOrder Where Receipt_No = '" & strReceiptNo & "' Order by �q��s�� "
    End If
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫ݱƨ��q���ƥi�H�٭�^ [�ݿ���q��]"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Screen.MousePointer = vbDefault
        blTRP02WEventEnable = True
        Exit Sub
    End If
    
    rs_TRP02W.AddNew
    rs_TRP02W.Fields("�s��").Value = 999
    rs_TRP02W.Fields("�e�f��").Value = tmp_Rs.Fields("�e�f��").Value
    rs_TRP02W.Fields("�q��s��").Value = tmp_Rs.Fields("�q��s��").Value
    rs_TRP02W.Fields("���").Value = tmp_Rs.Fields("���").Value
    rs_TRP02W.Fields("�c��").Value = tmp_Rs.Fields("�c��").Value
    rs_TRP02W.Fields("�O��").Value = tmp_Rs.Fields("�O��").Value
    rs_TRP02W.Fields("���n").Value = tmp_Rs.Fields("���n").Value
    rs_TRP02W.Fields("���q").Value = tmp_Rs.Fields("���q").Value
    rs_TRP02W.Fields("�Ȥ�s��").Value = tmp_Rs.Fields("�Ȥ�s��").Value
    rs_TRP02W.Fields("Area").Value = tmp_Rs.Fields("Area").Value
    rs_TRP02W.Fields("zip").Value = tmp_Rs.Fields("zip").Value
    rs_TRP02W.Fields("�Ȥ�W��").Value = tmp_Rs.Fields("�Ȥ�W��").Value
    rs_TRP02W.Fields("�B�e�a�}").Value = tmp_Rs.Fields("�B�e�a�}").Value
    rs_TRP02W.Fields("�q��Ƶ�").Value = tmp_Rs.Fields("�q��Ƶ�").Value
    rs_TRP02W.Fields("����").Value = tmp_Rs.Fields("����").Value
    rs_TRP02W.Fields("�S��ݨD1").Value = tmp_Rs.Fields("�S��ݨD1").Value
    rs_TRP02W.Fields("�S��ݨD2").Value = tmp_Rs.Fields("�S��ݨD2").Value
    rs_TRP02W.Fields("���").Value = tmp_Rs.Fields("���").Value
    rs_TRP02W.Fields("�M��").Value = tmp_Rs.Fields("�M��").Value
    rs_TRP02W.Fields("�N��").Value = tmp_Rs.Fields("�N��").Value
    rs_TRP02W.Fields("Receipt_No").Value = tmp_Rs.Fields("Receipt_No").Value
    rs_TRP02W.Fields("C_Receipt_No").Value = tmp_Rs.Fields("C_Receipt_No").Value
    rs_TRP02W.Fields("�f�D�渹").Value = tmp_Rs.Fields("�f�D�渹").Value
    rs_TRP02W.Fields("EXE�^��").Value = tmp_Rs.Fields("EXE�^��").Value
    rs_TRP02W.Fields("Area").Value = tmp_Rs.Fields("Area").Value
    rs_TRP02W.Fields("�Ȥ�²��").Value = tmp_Rs.Fields("�Ȥ�²��").Value & ""
    rs_TRP02W.Fields("���A").Value = tmp_Rs.Fields("���A").Value
        rs_TRP02W.Fields("��s������").Value = tmp_Rs.Fields("��s������").Value & ""
    rs_TRP02W.Fields("�t�e�ܧO").Value = tmp_Rs.Fields("�t�e�ܧO").Value
    rs_TRP02W.Update
    tmp_Rs.Close
    
    rs_TRP02W.Filter = adFilterNone
    rs_TRP02W.Sort = "�q��s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    If Not rs_TRP02W.EOF Then rs_TRP02W.MoveFirst
    blTRP02WEventEnable = True
End Sub

Private Sub ReSet_TRP02W_SeqNo()
    '���s���� [�ݱƨ��q��] �� [�s��] ����
    blTRP02WEventEnable = False
    dg_TRP02W.Visible = False
    rs_TRP02W.Filter = adFilterNone
    rs_TRP02W.Sort = "�q��s�� asc"  '��l�ƧǡA�@���ƧǸ��Ѥp�ܤj
    If Not rs_TRP02W.EOF Then rs_TRP02W.MoveFirst
    
    Dim dbSeqNo As Double
    dbSeqNo = 0
    Do While Not rs_TRP02W.EOF
        dbSeqNo = dbSeqNo + 1
        rs_TRP02W.Fields("�s��").Value = dbSeqNo
        rs_TRP02W.MoveNext
    Loop
    If rs_TRP02W.RecordCount > 0 Then rs_TRP02W.MoveFirst
    dg_TRP02W.Visible = True
    blTRP02WEventEnable = True
End Sub



Private Sub mvDate_DateClick(ByVal DateClicked As Date)
    '������
    Select Case mvDate.Tag
           Case "�X�����"
                txt_Tab0_TRPDate.Text = Format(mvDate.Value, "yyyymmdd")
           Case "�w�p������"
                txt_Tab0_CarCheckInDate.Text = Format(mvDate.Value, "yyyymmdd")
           Case "�q����R�_"
                txt_Tab3_DeliveryDate_Start.Text = Format(mvDate.Value, "yyyymmdd")
           Case "�q����R��"
                txt_Tab3_DeliveryDate_End.Text = Format(mvDate.Value, "yyyymmdd")
    End Select
    mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Me.mvDate.Visible = False
    If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab
End Sub

Private Sub txt_Tab0_CarCheckInDate_Click()
    '�ƨ��@�~ >> �w�p������
    If Trim(txt_Tab0_CarCheckInDate.Text) = "" Then
        mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab0_CarCheckInDate.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab0_CarCheckInDate.Text, 4) & "/" & Mid(txt_Tab0_CarCheckInDate.Text, 5, 2) & "/" & Right(txt_Tab0_CarCheckInDate.Text, 2))
        End If
    End If
    mvDate.Left = fam_RouteData.Left + txt_Tab0_CarCheckInDate.Left
    mvDate.Top = fam_RouteData.Top + txt_Tab0_CarCheckInDate.Top + txt_Tab0_CarCheckInDate.Height
    mvDate.Tag = "�w�p������"
    mvDate.Visible = True
End Sub

Private Sub txt_Tab0_CarCheckInDate_KeyPress(KeyAscii As Integer)
    '�w�p������
    Select Case KeyAscii
         Case 97 To 122, 65 To 90   '�����\��J�r��
              KeyAscii = 0
         Case vbKeyReturn
              KeyAscii = 0
              txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
              txt_Tab0_CarCheckInTime.SetFocus
    End Select
End Sub

Private Sub txt_Tab0_CarCheckInTime_KeyPress(KeyAscii As Integer)
    '�w�p����ɶ�
    Select Case KeyAscii
         Case 97 To 122, 65 To 90   '�����\��J�r��
              KeyAscii = 0
    End Select
End Sub

Private Sub txt_Tab0_DeliveryCarNo_KeyPress(KeyAscii As Integer)
    '�ƨ��@�~ >> ���P���X
    Select Case KeyAscii
           Case 97 To 122   '�ഫ���j�g�r��
                KeyAscii = KeyAscii - 32
    End Select
End Sub

Private Sub txt_Tab0_DeliveryCarNo_LostFocus()  'daniel--20040928<���uuser��J���~������>
    If Len(txt_Tab0_DeliveryCarNo.Text) = 0 Then Exit Sub
    str_SQL = "Select Vehicle_ID_No from trp09m where Vehicle_ID_No='" & Trim(txt_Tab0_DeliveryCarNo.Text) & "' "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    cn.CommandTimeout = 120
    If tmp_Rs.EOF Then
        'tmp_rs.Close
        msg_text = "�L���������"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_DeliveryCarNo.SetFocus
    End If
    tmp_Rs.Close
End Sub

Private Sub txt_Tab0_DockNo_KeyPress(KeyAscii As Integer)
    '�X�Y�Ȧs
    Select Case KeyAscii
           Case 97 To 122   '�ഫ���j�g�r��
                KeyAscii = KeyAscii - 32
           Case vbKeyReturn
                KeyAscii = 0
                txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
                txt_Tab0_CarCheckInDate.SetFocus
    End Select
End Sub

Private Sub txt_Tab0_RouteNo_KeyPress(KeyAscii As Integer)
    '�ƨ��@�~ >> ���u�s��
    Select Case KeyAscii
        Case 97 To 122     '�p�g�r���אּ�j�g�r��
             KeyAscii = KeyAscii - 32
        Case vbKeyReturn
             cmd_Tab0_Query.SetFocus
    End Select
End Sub

Private Sub txt_Tab0_TRPDate_Click()
    '�ƨ��@�~ >> �X�����
    If Trim(txt_Tab0_TRPDate.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab0_TRPDate.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab0_TRPDate.Text, 4) & "/" & Mid(txt_Tab0_TRPDate.Text, 5, 2) & "/" & Right(txt_Tab0_TRPDate.Text, 2))
        End If
    End If
    mvDate.Left = fam_SelectedOrders.Left + txt_Tab0_TRPDate.Left
    mvDate.Top = fam_SelectedOrders.Top + txt_Tab0_TRPDate.Top + txt_Tab0_TRPDate.Height
    mvDate.Tag = "�X�����"
    mvDate.Visible = True
End Sub

Private Sub txt_Tab0_TRPDate_KeyPress(KeyAscii As Integer)
    '�ƨ��@�~ > [�X�����] ��Ʈ榡�Gyyyymmdd
    Select Case KeyAscii
         Case 97 To 122, 65 To 90   '�����\��J�r��
              KeyAscii = 0
         Case vbKeyReturn
              If Fun_ChkDateFormat(txt_Tab0_TRPDate.Text) = 1 Then
                 msg_text = "�X������G" & funRtn_msg
                 MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                 txt_Tab0_TRPDate.SelStart = 0: txt_Tab0_TRPDate.SelLength = Len(txt_Tab0_TRPDate.Text): txt_Tab0_TRPDate.SetFocus
                 Exit Sub
              Else
                 cmd_Tab0_SelectCar.SetFocus
              End If
    End Select
End Sub

Public Sub frm_OP_TRPPlan_rsFilterAndSort(ByVal strCode As String, ByVal strReturn As String)
    '��椽�ΰƵ{���A�� frm_RS_FilterAndSort ���I�s
    '�ǤJ�ȡGstrCode      �ʧ@�ѧO�X
    '                     [FILTER] �ۭq�z��    [SORT] �Ƨ�
    '        strReturn    �z�� or �Ƨ� ���]�w�r��
    
    Select Case strCode
           Case "FILTER"  '�ۭq�z��
                Select Case UCase(strRSName_FilterAndSort)
                       Case "RS_TRP02W"                '�ݱƨ��q����
                            blTRP02WEventEnable = False
                            '�z��w����̡G�������
                            rs_TRP02W.Filter = "��='V'"
                            If Not rs_TRP02W.EOF Then
                               Do While Not rs_TRP02W.EOF
                                  rs_TRP02W.Fields(1).Value = " "
                                  rs_TRP02W.MoveNext
                               Loop
                            End If
                            rs_TRP02W.Filter = adFilterNone
                            rs_TRP02W.Filter = strReturn
                            strSourceFilter = strReturn
                            If rs_TRP02W.RecordCount = 0 Then
                               msg_text = "��p���A�䤣��ŦX���󪺭q���"
                               MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                               rs_TRP02W.Filter = adFilterNone
                               strSourceFilter = adFilterNone
                               rs_TRP02W.Sort = strSourceOrderBy   '�٭�ƧǤ覡
                               blTRP02WEventEnable = True
                               Exit Sub
                            End If
                            '���s�p�� [�ݱƨ��C��] ���`�p��T
                            Call ReCaculate_OrderSum
                            blTRP02WEventEnable = True
                       Case "RS_TAB2_RESERVEDORDERS"   '�O�d�q��
                            blTab2ReservedEventEnable = False
                            rs_Tab2_ReservedOrders.Filter = adFilterNone
                            rs_Tab2_ReservedOrders.Filter = strReturn
                            If rs_Tab2_ReservedOrders.RecordCount = 0 Then
                               msg_text = "��p���A�䤣��ŦX���󪺫O�d�q���"
                               MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                               rs_Tab2_ReservedOrders.Filter = adFilterNone
                               rs_Tab2_ReservedOrders.Sort = strSourceOrderBy   '�٭�ƧǤ覡
                               blTRP02WEventEnable = True
                               Exit Sub
                            End If
                            blTab2ReservedEventEnable = True
                       
                End Select
           Case "SORT"    '�Ƨ�
                Select Case UCase(strRSName_FilterAndSort)
                       Case "RS_TRP02W"               '�ݱƨ��q����
                            If rs_TRP02W.EOF Then Exit Sub
                            blTRP02WEventEnable = False
                            rs_TRP02W.Sort = strReturn
                            strSourceOrderBy = strReturn
                            blTRP02WEventEnable = True
                       Case "RS_TAB2_RESERVEDORDERS"   '�O�d�q��
                            If rs_Tab2_ReservedOrders.EOF Then Exit Sub
                            blTab2ReservedEventEnable = False
                            rs_Tab2_ReservedOrders.Sort = strReturn
                            strSourceOrderBy = strReturn
                            blTab2ReservedEventEnable = True
                End Select
    End Select
End Sub

Private Sub txt_Tab1_RouteNo_KeyPress(KeyAscii As Integer)
    '���u�s���C�� >> ���u�s��
    Select Case KeyAscii
         Case 97 To 122   '�ഫ�j�g�r��
              KeyAscii = KeyAscii - 32
         Case vbKeyReturn
              cmd_Tab1_RouteNoQuery.SetFocus
    End Select
End Sub

Private Sub Clear_RouteData()
    '�ƨ��@�~�G�M�����u�s��������
    blRouteModify = False
    strDispRouteNo = ""
    blRouteChange = False
    
    blTab0SelectedOrderEventEnable = False
    '�ƨ��@�~�G�w������ݱƨ��q��C�� DBGrid �榡�]�w
    Call CreateRS_Tab0_SelectedOrders
    '���s�p��w����q��G�c�ơA�O�ơA���n�A���q + �s�����s����
    Call Calculate_SelectedOrders
    blTab0SelectedOrderEventEnable = True
    
    txt_Tab0_TRPDate.Text = ""
    txt_Tab0_DeliveryCarNo.Text = ""
    txt_Tab0_DockNo.Text = ""
    txt_Tab0_CarCheckInDate.Text = ""
    txt_Tab0_CarCheckInTime.Text = ""
    txt_Tab0_DeliveryCompany.Text = ""
    txt_Tab0_DeliveryDriver.Text = ""
    txt_Tab0_DeliveryPhone.Text = ""
    txt_Tab0_DeliveryCarType.Text = ""
    txt_Tab0_Selected_Case.Text = ""
    txt_Tab0_Selected_Pallet.Text = ""
    txt_Tab0_Selected_Volumn.Text = ""
    txt_Tab0_Selected_Weight.Text = ""
End Sub

Private Function RouteData_Check() As Boolean
    Dim Str_D_Orderkey As String
    Str_D_Orderkey = ""
    '�ˮָ��u�s����ƬO�_���T
    RouteData_Check = False
    
    If Len(Trim(txt_Tab0_TRPDate.Text)) = 0 Then
        msg_text = "��ƿ��~�G����J�X�����"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_TRPDate.SetFocus
        Exit Function
    End If
    If Len(Trim(txt_Tab0_DeliveryCarNo.Text)) = 0 Then
        msg_text = "��ƿ��~�G����J���P���X"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Function
    End If
    
    '����ˮ�
    'a1.�X������G�榡 yyyymmdd
    txt_Tab0_TRPDate.Text = Trim(txt_Tab0_TRPDate.Text)
    If Fun_ChkDateFormat(txt_Tab0_TRPDate.Text) = 1 Then
        msg_text = "�X������G" & funRtn_msg
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_TRPDate.SelStart = 0: txt_Tab0_TRPDate.SelLength = Len(txt_Tab0_TRPDate.Text): txt_Tab0_TRPDate.SetFocus
        Exit Function
    End If
'    'a2.�X����� >= ����
'    If txt_Tab0_TRPDate.Text < Format(Now, "yyyymmdd") Then
'        msg_text = "�X��������o�p�󤵤�"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        txt_Tab0_TRPDate.SelStart = 0: txt_Tab0_TRPDate.SelLength = Len(txt_Tab0_TRPDate.Text): txt_Tab0_TRPDate.SetFocus
'        Exit Function
'    End If
    
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    'b.�ˮ� [���P���X] �O�_����
    txt_Tab0_DeliveryCarNo.Text = Trim(txt_Tab0_DeliveryCarNo.Text)
    str_SQL = "Select Count(*) as RecCount From TRP09M Where Vehicle_ID_NO = '" & txt_Tab0_DeliveryCarNo.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "��ƿ��~�G���P���X " & txt_Tab0_DeliveryCarNo.Text & " ������"
        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
        txt_Tab0_DeliveryCarNo.SelStart = 0: txt_Tab0_DeliveryCarNo.SelLength = Len(txt_Tab0_DeliveryCarNo.Text)
        txt_Tab0_DeliveryCarNo.SetFocus
        Exit Function
    End If
    tmp_Rs.Close
'    '���w�X�Y�Ȧs�G������J
'    txt_Tab0_DockNo.Text = Trim(txt_Tab0_DockNo.Text)
'    If Len(Trim(txt_Tab0_DockNo.Text)) = 0 Then
'        msg_text = "��ƿ��~�G[�X�Y�Ȧs] ������J"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        txt_Tab0_DockNo.SetFocus
'        Exit Function
'    End If
    '�w�p������
    txt_Tab0_CarCheckInDate.Text = Trim(txt_Tab0_CarCheckInDate.Text)
    If Len(Trim(txt_Tab0_CarCheckInDate.Text)) <> 8 Then
        msg_text = "�w�p�������G��Ʈ榡 yyyymmdd "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
        txt_Tab0_CarCheckInDate.SetFocus
        Exit Function
    End If
    If Fun_ChkDateFormat(txt_Tab0_CarCheckInDate.Text) = 1 Then
        msg_text = "�w�p�������G��ƿ��~ yyyymmdd�A" & funRtn_msg
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text)
        txt_Tab0_CarCheckInDate.SetFocus
        Exit Function
    End If
'    'a2.�w�p������ >= ����
'    If txt_Tab0_CarCheckInDate.Text < Format(Now, "yyyymmdd") Then
'       msg_text = "�w�p���������o�p�󤵤�"
'       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'       txt_Tab0_CarCheckInDate.SelStart = 0: txt_Tab0_CarCheckInDate.SelLength = Len(txt_Tab0_CarCheckInDate.Text): txt_Tab0_CarCheckInDate.SetFocus
'       Exit Function
'    End If
    
    '�w�p����ɶ�
    txt_Tab0_CarCheckInTime.Text = Trim(txt_Tab0_CarCheckInTime.Text)
    If Len(Trim(txt_Tab0_CarCheckInTime.Text)) <> 4 Then
        msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
        txt_Tab0_CarCheckInTime.SetFocus
        Exit Function
    End If
    Select Case Left(txt_Tab0_CarCheckInTime.Text, 2)
           Case "00" To "23"
           Case Else
                msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
                txt_Tab0_CarCheckInTime.SetFocus
                Exit Function
    End Select
    Select Case Right(txt_Tab0_CarCheckInTime.Text, 2)
           Case "00" To "59"
           Case Else
                msg_text = "�w�p����ɶ��G��Ʈ榡 hhss "
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                txt_Tab0_CarCheckInTime.SelStart = 0: txt_Tab0_CarCheckInTime.SelLength = Len(txt_Tab0_CarCheckInTime.Text)
                txt_Tab0_CarCheckInTime.SetFocus
                Exit Function
    End Select
    
    
    '�ˬd�O�_���Q���ܫO�d�q�檺���by Eric 20141215
    If rs_Tab0_SelectedOrders Is Nothing Then Exit Function
    If rs_Tab0_SelectedOrders.RecordCount = 0 Then Exit Function
    rs_Tab0_SelectedOrders.MoveFirst
    Do While Not rs_Tab0_SelectedOrders.EOF
        Str_D_Orderkey = Str_D_Orderkey & "'" & rs_Tab0_SelectedOrders.Fields("Receipt_no") & "',"
        rs_Tab0_SelectedOrders.MoveNext
    Loop
        Call DB_CheckConnectStatus
        Call ReDim_Recordset(tmp_Rs)
        str_SQL = "select receipt_no from trp02t where route_no = 'D' and receipt_no in (" & Mid(Str_D_Orderkey, 1, Len(Str_D_Orderkey) - 1) & ")"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
                tmp_Rs.MoveFirst
                msg_text = "�o�{���ƨ���ƳQ���ܫO�d�q��" & Chr(13) + Chr(10) & "�Э��s���J�ݱƨ���ƦA�i��ƨ��@�~�C"
                MsgBox msg_text, vbOKOnly + vbCritical, msg_title
                Str_D_Orderkey = ""
                Do While Not tmp_Rs.EOF
                    Str_D_Orderkey = Str_D_Orderkey & "'" & tmp_Rs.Fields("receipt_no") & "',"
                    tmp_Rs.MoveNext
                Loop
                msg_text = "�Q���ܫO�d�q�檺TMS�q��:" & Chr(13) + Chr(10) & Mid(Str_D_Orderkey, 1, Len(Str_D_Orderkey) - 1)
                MsgBox msg_text, vbOKOnly + vbCritical, "�O�d�q�����ˬd"
                tmp_Rs.Close
                Exit Function
    End If
    tmp_Rs.Close
    
    RouteData_Check = True
End Function

Private Sub Delete_RouteNo(strRouteNo As String)
    Screen.MousePointer = vbHourglass
    blTab1RouteEventEnable = False
    Tran_Level = 0
    Tran_Level = cn.BeginTrans
    
    '�R�� TRP01T ���u�s���D��
    Call DB_CheckConnectStatus
    
    '(1).�N TRP03T �g�^ TRP03W >> �R�� TRP03T
    str_SQL = "Insert into TRP03W(" & _
              " STORERKEY,RECEIPT_NO,SEQ_NO,PRODUCT_NO,SHIP_UNIT,ORDER_QTY,PALLET_QTY,WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,EXTERN) " & _
              "Select A.STORERKEY,A.RECEIPT_NO,A.SEQ_NO,A.PRODUCT_NO,A.SHIP_UNIT,A.ORDER_QTY,A.PALLET_QTY,A.WEIGHT,A.VOLUMN_WEIGHT,A.Description,A.EXTERN " & _
              "From TRP03T A Where a.Route_No = '" & strRouteNo & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '(2).�N TRP02T �g�^ TRP02W >> �R�� TRP02T
    
    'EXE_CONFIRM=9��(�w���z�f)�A�g�^TRP02W�ɪ��A����
    cn.Execute "update TRP02T set EXE_CONFIRM = 0 Where Route_No = '" & strRouteNo & "' and exe_confirm <> 9 ", RowsAffect, adExecuteNoRecords
    
    str_SQL = "Insert into TRP02W(" & _
              " RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              " WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,c_receipt_no,otqty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey) " & _
              "Select RECEIPT_NO,RECEIPT_TYPE,TRP_TYPE,RECEIPT_DATE,ARRIVE_DATE,CONSIGNEEKEY,CASE_CNT,PALLET_QTY," & _
              "  Weight,VOLUMN_WEIGHT,Description,STORERKEY,EXTERN,URGENT_MARK,RESERVE_MARK,COLD_MARK,EXE_CONFIRM,Priority,c_receipt_no,otqty,OTConfirmUser,OTConfirmDate,OTPrintDate,OTPrintTimes,facility,bconsigneekey " & _
              "From TRP02T Where Route_No = '" & strRouteNo & "'" 'exe_confirm �]�w�� 0 by gemini 20080106
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
'    '(3).�R�� TRP02T & TRP03T
'    str_SQL = "Delete From TRP03T Where Route_No = '" & strRouteNo & "'"
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    str_SQL = "Delete From TRP02T Where Route_No = '" & strRouteNo & "'"
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    '(4).�R�� TRP05T
'    str_SQL = "Delete From TRP05T Where Route_No = '" & strRouteNo & "'"
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    '(5).�R�� TRP01T
'    str_SQL = "Delete From TRP01T Where Route_No = '" & strRouteNo & "'"
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords


    '�R��Trp1,2,3,5t �ϥΤ���edit by Eric 20141215
    str_SQL = "Delete From TRP03T Where Route_No = '" & strRouteNo & "';" & _
              "Delete From TRP02T Where Route_No = '" & strRouteNo & "';" & _
              "Delete From TRP05T Where Route_No = '" & strRouteNo & "';" & _
              "Delete From TRP01T Where Route_No = '" & strRouteNo & "';"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    
    '(6)��Ʈw���ʽT�{
    cn.CommitTrans
    Tran_Level = 0
    Screen.MousePointer = vbDefault
    
    Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      cn.RollbackTrans
      Tran_Level = 0
   End If
   Dim tmpString As String
   msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-�ƨ��@�~-���u�s���R��", Me.Caption, "Form ���� SubProgram Delete_RouteNo", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub Retrive_OrderSum()
    '�����ݱƨ��q��G�`�p��ƭ�
    txt_Tab0_srcTotal_OTqty.Text = ""
    txt_Tab0_srcTotal_Case.Text = ""
    txt_Tab0_srcTotal_Pallet.Text = ""
    txt_Tab0_srcTotal_Volumn.Text = ""
    txt_Tab0_srcTotal_Weight.Text = ""
    str_SQL = "Select Isnull(Round(sum(�c��),3),0) as �`�c��,Isnull(Round(sum(���q),3),0) as �`���q," & _
              "       Isnull(Round(sum(���n),3),0) as �`���n,Isnull(Round(sum(�O��),3),0) as �`�O��," & _
              "       Isnull(Round(sum(���),3),0) as �`��� " & _
              "From CutOrders_SourceOrder  "
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    cn.CommandTimeout = 0   '�L��������
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        txt_Tab0_srcTotal_OTqty.Text = tmp_Rs.Fields("�`���").Value
        txt_Tab0_srcTotal_Case.Text = tmp_Rs.Fields("�`�c��").Value
        txt_Tab0_srcTotal_Pallet.Text = tmp_Rs.Fields("�`�O��").Value
        txt_Tab0_srcTotal_Volumn.Text = tmp_Rs.Fields("�`���n").Value
        txt_Tab0_srcTotal_Weight.Text = tmp_Rs.Fields("�`���q").Value
    End If
    tmp_Rs.Close
End Sub

Private Sub ReCaculate_OrderSum()
    '�����ݱƨ��q��G�`�p��ƭ�  >>  �ثe�ݿ�C���`�p
    txt_Tab0_srcTotal_OTqty.Text = ""
    txt_Tab0_srcTotal_Case.Text = ""
    txt_Tab0_srcTotal_Pallet.Text = ""
    txt_Tab0_srcTotal_Volumn.Text = ""
    txt_Tab0_srcTotal_Weight.Text = ""
    
    If rs_TRP02W.RecordCount = 0 Then Exit Sub
    Dim dbTotalOTqty As Double
    Dim dbTotalCase As Double
    Dim dbTotalPallet As Double
    Dim dbTotalWeight As Double
    Dim dbTotalVolumn As Double
    dbTotalOTqty = 0: dbTotalCase = 0: dbTotalPallet = 0: dbTotalVolumn = 0: dbTotalWeight = 0
    blTRP02WEventEnable = False
    dg_TRP02W.Visible = False
    rs_TRP02W.MoveFirst
    Do While Not rs_TRP02W.EOF
        dbTotalOTqty = dbTotalOTqty + rs_TRP02W.Fields("���").Value
        dbTotalCase = dbTotalCase + rs_TRP02W.Fields("�c��").Value
        dbTotalPallet = dbTotalPallet + rs_TRP02W.Fields("�O��").Value
        dbTotalVolumn = dbTotalVolumn + rs_TRP02W.Fields("���n").Value
        dbTotalWeight = dbTotalWeight + rs_TRP02W.Fields("���q").Value
        rs_TRP02W.MoveNext
    Loop
    rs_TRP02W.MoveFirst
    txt_Tab0_srcTotal_OTqty.Text = dbTotalOTqty
    txt_Tab0_srcTotal_Case.Text = dbTotalCase
    txt_Tab0_srcTotal_Pallet.Text = dbTotalPallet
    txt_Tab0_srcTotal_Volumn.Text = dbTotalVolumn
    txt_Tab0_srcTotal_Weight.Text = dbTotalWeight
    
    dg_TRP02W.Visible = True
    blTRP02WEventEnable = True
End Sub


Private Sub txt_Tab3_DeliveryDate_End_Click()
    '�ƨ��@�~ >> �q����R
    If Trim(txt_Tab3_DeliveryDate_End.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab3_DeliveryDate_End.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab3_DeliveryDate_End.Text, 4) & "/" & Mid(txt_Tab3_DeliveryDate_End.Text, 5, 2) & "/" & Right(txt_Tab3_DeliveryDate_End.Text, 2))
        End If
    End If
    mvDate.Left = fma_Tab3_OrderSum.Left + txt_Tab3_DeliveryDate_End.Left
    mvDate.Top = fma_Tab3_OrderSum.Top + txt_Tab3_DeliveryDate_End.Top + txt_Tab3_DeliveryDate_End.Height
    mvDate.Tag = "�q����R��"
    mvDate.Visible = True
End Sub

Private Sub txt_Tab3_DeliveryDate_Start_Click()
    '�ƨ��@�~ >> �q����R
    If Trim(txt_Tab3_DeliveryDate_Start.Text) = "" Then
       mvDate.Value = Now
    Else
        If Fun_ChkDateFormat(txt_Tab3_DeliveryDate_Start.Text) = 1 Then
            mvDate.Value = Now
        Else
            mvDate.Value = CDate(Left(txt_Tab3_DeliveryDate_Start.Text, 4) & "/" & Mid(txt_Tab3_DeliveryDate_Start.Text, 5, 2) & "/" & Right(txt_Tab3_DeliveryDate_Start.Text, 2))
        End If
    End If
    mvDate.Left = fma_Tab3_OrderSum.Left + txt_Tab3_DeliveryDate_Start.Left
    mvDate.Top = fma_Tab3_OrderSum.Top + txt_Tab3_DeliveryDate_Start.Top + txt_Tab3_DeliveryDate_Start.Height
    mvDate.Tag = "�q����R�_"
    mvDate.Visible = True
End Sub
