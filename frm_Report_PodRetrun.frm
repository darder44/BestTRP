VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Report_PodRetrun 
   Caption         =   "POD�^��"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   BeginProperty Font 
      Name            =   "�ө���"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   8370
   WindowState     =   2  '�̤j��
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   240
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3720
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
      StartOfWeek     =   121831425
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton cmd_SendMail 
         BackColor       =   &H00C0C0FF&
         Caption         =   "��Excel�o�e"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   7320
         Picture         =   "frm_Report_PodRetrun.frx":0000
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   37
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_Msg 
         Alignment       =   2  '�m�����
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
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   120
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.TextBox txt_UnReciept 
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
         Left            =   1200
         TabIndex        =   35
         ToolTipText     =   "���ڵu���A�|���������q�渹�X"
         Top             =   1680
         Width           =   3285
      End
      Begin VB.CheckBox Check4 
         Caption         =   "D"
         Height          =   255
         Left            =   4680
         TabIndex        =   34
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         Caption         =   "C"
         Height          =   255
         Left            =   4680
         TabIndex        =   33
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "B"
         Height          =   255
         Left            =   4680
         TabIndex        =   32
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         Caption         =   "A"
         Height          =   255
         Left            =   4680
         TabIndex        =   31
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox cb_all 
         Caption         =   "�^�ǥ���"
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   2040
         Value           =   1  '�֨�
         Width           =   1335
      End
      Begin VB.OptionButton optIn 
         Caption         =   "�h�f��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optOut 
         Caption         =   "�X�f��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   2040
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.CheckBox chkPrintPreView 
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
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3000
         TabIndex        =   24
         Top             =   360
         Value           =   1  '�֨�
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.CommandButton cmdOTUpdate 
         BackColor       =   &H00C0E0FF&
         Caption         =   "���_�妸POD�^��"
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
         Left            =   5880
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   7
         ToolTipText     =   "�u�w����_"
         Top             =   1320
         Width           =   1065
      End
      Begin VB.ComboBox cboStorerkey 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   1485
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1065
         TabIndex        =   20
         Top             =   2340
         Visible         =   0   'False
         Width           =   3375
         Begin VB.OptionButton optYes 
            Caption         =   "�w�^��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton optNo 
            Caption         =   "���^��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton optAll 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.CommandButton cmd2Excel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "��X�^��"
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
         Left            =   5880
         Picture         =   "frm_Report_PodRetrun.frx":08CA
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtDeliveryDateS 
         Alignment       =   2  '�m�����
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
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   1
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtDeliveryDateE 
         Alignment       =   2  '�m�����
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
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   2
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtExternOrderkeyS 
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
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.TextBox txtExternOrderkeyE 
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
         Left            =   3000
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FF80FF&
         Caption         =   "���}"
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
         Left            =   7080
         Picture         =   "frm_Report_PodRetrun.frx":1BC4
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   12
         Top             =   1320
         Width           =   1065
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00C0FFC0&
         Caption         =   "���]"
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
         Left            =   7080
         Picture         =   "frm_Report_PodRetrun.frx":2B7D6
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   11
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�d��"
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
         Left            =   4680
         Picture         =   "frm_Report_PodRetrun.frx":2BAE8
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   9
         ToolTipText     =   "��f���180�Ѥ�"
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�f�D�s��"
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
         Left            =   120
         TabIndex        =   23
         Top             =   540
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "��������"
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
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1725
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
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
         Index           =   23
         Left            =   2655
         TabIndex        =   19
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "��f���"
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
         Index           =   22
         Left            =   120
         TabIndex        =   18
         Top             =   1005
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
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
         Index           =   3
         Left            =   2655
         TabIndex        =   17
         Top             =   1380
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�q�渹�X"
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
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   1365
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   7920
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   13
      Top             =   2520
      Width           =   8295
      Begin TabDlg.SSTab SSTab1 
         Height          =   3495
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6165
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Header"
         TabPicture(0)   =   "frm_Report_PodRetrun.frx":2BDF2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "dgMain_Header"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detail"
         TabPicture(1)   =   "frm_Report_PodRetrun.frx":2BE0E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dgMain_Detail"
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid dgMain_Header 
            Height          =   2295
            Left            =   120
            TabIndex        =   26
            Top             =   480
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   2
            RowHeight       =   20
            TabAction       =   1
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9.75
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
         Begin MSDataGridLib.DataGrid dgMain_Detail 
            Height          =   2295
            Left            =   -74880
            TabIndex        =   27
            Top             =   480
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   2
            RowHeight       =   20
            TabAction       =   1
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "�s�ө���"
               Size            =   9.75
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
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '������U��
      Height          =   390
      Left            =   0
      TabIndex        =   22
      Top             =   6630
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Text            =   "���A"
            TextSave        =   "���A"
            Object.ToolTipText     =   "���A"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   8149
            MinWidth        =   2646
            Object.ToolTipText     =   "��Ƶ���"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.ToolTipText     =   "�ϥΪ�"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Report_PodRetrun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blRouteT0Change As Boolean
Private rsMainHeader As ADODB.Recordset
Private rsMainDetail As ADODB.Recordset
Private rs_Receipt As ADODB.Recordset
Private rsMainReceitDetail As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private fso As Scripting.FileSystemObject
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cb_all_Click()
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub

If cb_all.Value = 1 Then
    '����
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        rsMainHeader.Fields("�O�_�^��") = "V"
        rsMainHeader.MoveNext
    Loop
    rsMainHeader.MoveFirst
Else
    '���P����
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        rsMainHeader.Fields("�O�_�^��") = " "
        rsMainHeader.MoveNext
    Loop
    rsMainHeader.MoveFirst
End If

End Sub

Private Sub Check1_Click()
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub

'�M��
rsMainHeader.MoveFirst
Do While Not rsMainHeader.EOF
    rsMainHeader.Fields("�O�_�^��") = " "
    rsMainHeader.MoveNext
Loop

rsMainHeader.MoveFirst

If Check1.Value = 1 Then
    '����
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        If RTrim(rsMainHeader.Fields("�����q�N��")) = "A" Then
            rsMainHeader.Fields("�O�_�^��") = "V"
        Else
            rsMainHeader.Fields("�O�_�^��") = " "
        End If
        rsMainHeader.MoveNext
    Loop
    rsMainHeader.MoveFirst
Else
    '���P����
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        rsMainHeader.Fields("�O�_�^��") = " "
        rsMainHeader.MoveNext
    Loop
    
End If
    rsMainHeader.MoveFirst
End Sub


Private Sub Check2_Click()
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub

'�M��
rsMainHeader.MoveFirst
Do While Not rsMainHeader.EOF
    rsMainHeader.Fields("�O�_�^��") = " "
    rsMainHeader.MoveNext
Loop

rsMainHeader.MoveFirst

If Check2.Value = 1 Then
    '����
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        If RTrim(rsMainHeader.Fields("�����q�N��")) = "B" Then
            rsMainHeader.Fields("�O�_�^��") = "V"
        Else
            rsMainHeader.Fields("�O�_�^��") = " "
        End If
        rsMainHeader.MoveNext
    Loop
    rsMainHeader.MoveFirst
Else
    '���P����
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        rsMainHeader.Fields("�O�_�^��") = " "
        rsMainHeader.MoveNext
    Loop
    
End If
    rsMainHeader.MoveFirst
End Sub


Private Sub Check3_Click()
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub

'�M��
rsMainHeader.MoveFirst
Do While Not rsMainHeader.EOF
    rsMainHeader.Fields("�O�_�^��") = " "
    rsMainHeader.MoveNext
Loop

rsMainHeader.MoveFirst

If Check3.Value = 1 Then
    '����
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        If RTrim(rsMainHeader.Fields("�����q�N��")) = "C" Then
            rsMainHeader.Fields("�O�_�^��") = "V"
        Else
            rsMainHeader.Fields("�O�_�^��") = " "
        End If
        rsMainHeader.MoveNext
    Loop
    rsMainHeader.MoveFirst
Else
    '���P����
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        rsMainHeader.Fields("�O�_�^��") = " "
        rsMainHeader.MoveNext
    Loop
    
End If
    rsMainHeader.MoveFirst
End Sub


Private Sub Check4_Click()
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub

'�M��
rsMainHeader.MoveFirst
Do While Not rsMainHeader.EOF
    rsMainHeader.Fields("�O�_�^��") = " "
    rsMainHeader.MoveNext
Loop

rsMainHeader.MoveFirst

If Check4.Value = 1 Then
    '����
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        If RTrim(rsMainHeader.Fields("�����q�N��")) = "D" Then
            rsMainHeader.Fields("�O�_�^��") = "V"
        Else
            rsMainHeader.Fields("�O�_�^��") = " "
        End If
        rsMainHeader.MoveNext
    Loop
    rsMainHeader.MoveFirst
Else
    '���P����
    rsMainHeader.MoveFirst
    Do While Not rsMainHeader.EOF
        rsMainHeader.Fields("�O�_�^��") = " "
        rsMainHeader.MoveNext
    Loop
    
End If
    rsMainHeader.MoveFirst
End Sub

Private Sub cmd_SendMail_Click()

On Error GoTo err_Handle
Dim strFrom As String, strTo As String, strCC As String, strBCC As String, strSubject As String, strTextbody As String, strAlways As String, strEmailID As String, strEmailPW As String, str_Date As String, strLMBO01Mail As String, strAddAttachment As String
'Ū��ini�Ѽ�
Dim objIni As New vbIniFile
str_Date = Format(Now(), "YYYY/MM/DD hh:mm:ss")
'objIni.FileName = App.Path & "/" & App.title & ".ini"
'
'strFrom = objIni.ReadData("INVCHECKEMAIL_LTKK01", "From", "")
'strTo = objIni.ReadData("INVCHECKEMAIL_LTKK01", "To", "")
'strCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "CC", "")
'strBCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "BCC", "")
'strSubject = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Subject", "")
'strTextbody = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Textbody", "")
'strEmailID = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailID", "")
'strEmailPW = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailPW", "")
'strAlways = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Always", "NO")

'�������w
strFrom = "autoreport@bestlog.com.tw"
strTo = "wanfang@maobao.com.tw;kellychen@maobao.com.tw;sara@maobao.com.tw;betty@maobao.com.tw;dennisren@maobao.com.tw;sharon@maobao.com.tw;alisa@maobao.com.tw;tia@maobao.com.tw;julia@maobao.com.tw;yolanda@maobao.com.tw"
'strTo = "gemini@bestlog.com.tw"
strCC = "tina.h@bestlog.com.tw;joane@bestlog.com.tw"
strSubject = str_Date & " POD Feedback Notice"
strTextbody = txt_Msg & Chr(13) & Chr(10) & "The letter sent automatically by the system, do not directly reply.Thanks" & Chr(13) & Chr(10) & "Time:" & str_Date
strEmailID = "autoreport"
strEmailPW = "bestauto"
strAlways = "NO"

If UCase(RTrim(strAlways)) <> "YES" Then strAlways = "NO"
Set objIni = Nothing

If Len(RTrim(strFrom)) > 0 Then '���H���
    strLMBO01Mail = "YES"
End If

If strLMBO01Mail = "YES" Then
Screen.MousePointer = 11
'�ǰe�l��
    Dim objEmail As Object
    Set objEmail = CreateObject("CDO.Message")

    objEmail.From = strFrom
    objEmail.To = strTo
    objEmail.CC = strCC   ' �ƥ�
    objEmail.BCC = strBCC ' �K��ƥ�
    objEmail.Subject = strSubject
    objEmail.TextBody = strTextbody
    'objEmail.AddAttachment strAddAttachment

    objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "bestlog.com.tw"
    objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    'SMTP ���A���ݭn���Ү�
    If Len(RTrim(strEmailID)) > 0 Then
        objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendusername") = strEmailID
        objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strEmailPW
    End If
    objEmail.Configuration.Fields.Update
    objEmail.Send

    Set objEmail = Nothing

End If

Exit Sub

err_Handle:
Screen.MousePointer = 0
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmd2Excel_Click()
On Error GoTo LogOnError
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub
    
    Screen.MousePointer = 11
    Dim FileName As String, txtpath As String, bl_Check As Boolean, str_Date As String, str_Orderkey As String
    Dim bl_CheckA As Boolean, bl_CheckB As Boolean, bl_CheckC As Boolean, bl_CheckD As Boolean, bl_CheckE As Boolean, bl_CheckSNRT As Boolean, bl_CheckOther As Boolean
    bl_CheckA = False: bl_CheckB = False: bl_CheckC = False: bl_CheckD = False: bl_CheckE = False: bl_CheckSNRT = False: bl_CheckOther = False
    
    If rsMainHeader.RecordCount = 0 Then Exit Sub
    bl_Check = False
    '�ˬd�O�_���Ŀ�A���ˬd�����Ǥ����q�O
    Do While Not rsMainHeader.EOF
        If rsMainHeader.Fields("�O�_�^��") = "V" Then
            bl_Check = True
            If RTrim(rsMainHeader.Fields("����")) = "SNRT" Then bl_CheckSNRT = True
            If RTrim(rsMainHeader.Fields("����")) = "Other" Then bl_CheckOther = True
            If RTrim(rsMainHeader.Fields("����")) = "A" Then bl_CheckA = True
            If RTrim(rsMainHeader.Fields("����")) = "B" Then bl_CheckB = True
            If RTrim(rsMainHeader.Fields("����")) = "C" Then bl_CheckC = True
            If RTrim(rsMainHeader.Fields("����")) = "D" Then bl_CheckD = True
            If Len(RTrim(rsMainHeader.Fields("����"))) = 0 Then bl_CheckE = True
        End If
        rsMainHeader.MoveNext
    Loop
    
    If bl_Check = False Then MsgBox "�S���Ŀ�^�Ǹ�ơA�нT�{�A�^��!", vbCritical + vbOKOnly, "�^���ˬd": Screen.MousePointer = 0: Exit Sub
    

    str_Date = Format(Now, "YYMMDDHHNNSS"): str_Orderkey = ""
        
    If bl_CheckSNRT = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\�p�_�j��o", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\�p�_�j��o", str_Date, "SNRT")
    End If
    
    If bl_CheckOther = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\��L", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\��L", str_Date, "Other")
    End If
    
    If bl_CheckA = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\�`���q", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\�`���q", str_Date, "A")
    End If
    
    If bl_CheckB = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\�_��", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\�_��", str_Date, "B")
    End If
    
    If bl_CheckC = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\�n��", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\�n��", str_Date, "C")
    End If
    
    If bl_CheckD = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\����", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\����", str_Date, "D")
    End If

    If bl_CheckE = True Then
        Call MBOrs2txt("C:\BEST\LMBO01\POD\���`", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\���`", str_Date, "")
    End If
    
    Screen.MousePointer = 0:
    rsMainHeader.Filter = ""

Exit Sub

LogOnError:
'    rsMainHeader.Close
'    rsMainDetail.Close
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub


Function MBOrs2txt(Str_Path1 As String, Str_Path2 As String, str_Date As String, Str_Company As String)
        Dim FileName As String, str_Orderkey As String, txtpath As String
'Str_Path1 = �����ƥ����|
'Str_Path2 = FTP�ƥ����|
'Str_RoutePath = �������s��Ƴƥ����|
'Str_FtpRoutePath = FTP���s��Ƴƥ����|

On Error GoTo LogOnError
            Dim ReturnOrders As Double, ReturnOrderdetail As Double, ReturnSignqty As Double
            Dim Str_RoutePath As String, Str_FtpRoutePath As String
'            Str_RoutePath = "C:\BEST\LMBO01\POD\�t�e���s"
'            Str_FtpRoutePath = "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\�t�e���s"
            ReturnOrders = 0: ReturnOrderdetail = 0: ReturnSignqty = 0
            rsMainHeader.MoveFirst
            '�q��D�����r��
            rsMainHeader.MoveFirst
            'strOrderNo = ""
            Set fso = New FileSystemObject
            FileName = Str_Company & "_rtb" & str_Date & ".txt"
            If Str_Company = "" Then FileName = "E_rtb" & str_Date & ".txt"
            If Dir(Str_Path1, vbDirectory) = "" Then MkDirs Str_Path1
            txtpath = Str_Path1 & "\" & FileName
            Open txtpath For Append As #1
            Do While Not rsMainHeader.EOF
                If rsMainHeader.Fields("�O�_�^��") = "V" And (rsMainHeader.Fields("����") = Str_Company Or Len(Str_Company) = 0) Then
                    ReturnOrders = ReturnOrders + 1
                    str_Orderkey = str_Orderkey & "'" & RTrim(rsMainHeader.Fields("TMS�渹")) & "',"
                    If Not IsNull(rsMainHeader.Fields(3)) Then Print #1, StrPadRightC(rsMainHeader.Fields(3), 1); Else Print #1, StrPadRightC(" ", 1);  '�����q�N��
                    If Not IsNull(rsMainHeader.Fields(4)) Then Print #1, StrPadRightC(rsMainHeader.Fields(4), 8); Else Print #1, StrPadRightC(" ", 8);  '�q�渹�X
                    If Not IsNull(rsMainHeader.Fields(5)) Then Print #1, StrPadRightC(rsMainHeader.Fields(5), 7); Else Print #1, StrPadRightC(" ", 7);  '�ĳ���
                    If Not IsNull(rsMainHeader.Fields(6)) Then Print #1, StrPadRightC(rsMainHeader.Fields(6), 10); Else Print #1, StrPadRightC(" ", 10);    '�o�����X
                    If Not IsNull(rsMainHeader.Fields(7)) Then Print #1, StrPadRightC(rsMainHeader.Fields(7), 2); Else Print #1, StrPadRightC(" ", 2);  '�o�����X�ˬd�X
                    If Not IsNull(rsMainHeader.Fields(8)) Then Print #1, StrPadRightC(rsMainHeader.Fields(8), 7); Else Print #1, StrPadRightC(" ", 7);  '�o�����
                    If Not IsNull(rsMainHeader.Fields(9)) Then Print #1, StrPadRightC(rsMainHeader.Fields(9), 8); Else Print #1, StrPadRightC(" ", 8);  '�Ȥ�s��
                    If Not IsNull(rsMainHeader.Fields(10)) Then Print #1, StrPadRightC(rsMainHeader.Fields(10), 50); Else Print #1, StrPadRightC(" ", 50);  '�Ȥ�W��
                    If Not IsNull(rsMainHeader.Fields(11)) Then Print #1, StrPadRightC(rsMainHeader.Fields(11), 3); Else Print #1, StrPadRightC(" ", 3);    '���N�N��
                    If Not IsNull(rsMainHeader.Fields(12)) Then Print #1, StrPadRightC(rsMainHeader.Fields(12), 2); Else Print #1, StrPadRightC(" ", 2);    '�U�f���{
                    If Not IsNull(rsMainHeader.Fields(13)) Then Print #1, StrPadRightC(rsMainHeader.Fields(13), 70); Else Print #1, StrPadRightC(" ", 70);  '�e�f�a�}
                    If Not IsNull(rsMainHeader.Fields(14)) Then Print #1, StrPadRightC(rsMainHeader.Fields(14), 1); Else Print #1, StrPadRightC(" ", 1);    '�p��
                    If Not IsNull(rsMainHeader.Fields(15)) Then Print #1, StrPadRightC(rsMainHeader.Fields(15), 8); Else Print #1, StrPadRightC(" ", 8);    '�Τ@�s��
                    If Not IsNull(rsMainHeader.Fields(16)) Then Print #1, StrPadLeft(rsMainHeader.Fields(16), 8); Else Print #1, StrPadLeft(" ", 8);    '�������B
                    If Not IsNull(rsMainHeader.Fields(17)) Then Print #1, StrPadLeft(rsMainHeader.Fields(17), 8); Else Print #1, StrPadLeft(" ", 8);    '�ƶq�������B
                    If Not IsNull(rsMainHeader.Fields(18)) Then Print #1, StrPadLeft(rsMainHeader.Fields(18), 8); Else Print #1, StrPadLeft(" ", 8);    '�S�O�������B
                    If Not IsNull(rsMainHeader.Fields(19)) Then Print #1, StrPadLeft(rsMainHeader.Fields(19), 8); Else Print #1, StrPadLeft(" ", 8);    '�{������
                    If Not IsNull(rsMainHeader.Fields(20)) Then Print #1, StrPadLeft(rsMainHeader.Fields(20), 10); Else Print #1, StrPadLeft(" ", 10);  '�f��
                    If Not IsNull(rsMainHeader.Fields(21)) Then Print #1, StrPadLeft(rsMainHeader.Fields(21), 10); Else Print #1, StrPadLeft(" ", 10);  '�|�e���B
                    If Not IsNull(rsMainHeader.Fields(22)) Then Print #1, StrPadLeft(rsMainHeader.Fields(22), 8); Else Print #1, StrPadLeft(" ", 8);    '�|�B
                    If Not IsNull(rsMainHeader.Fields(23)) Then Print #1, StrPadRightC(rsMainHeader.Fields(23), 70); Else Print #1, StrPadRightC(" ", 70);  '�Ƶ�
                    If Not IsNull(rsMainHeader.Fields(24)) Then Print #1, StrPadRightC(rsMainHeader.Fields(24), 25); Else Print #1, StrPadRightC(" ", 25);  '�Ȥ�q��s��
                    If Not IsNull(rsMainHeader.Fields(25)) Then Print #1, StrPadRightC(rsMainHeader.Fields(25), 1); Else Print #1, StrPadRightC(" ", 1);    '�H�f���o���X
                    If Not IsNull(rsMainHeader.Fields(26)) Then Print #1, StrPadRightC(rsMainHeader.Fields(26), 1); Else Print #1, StrPadRightC(" ", 1);    '�H�f���q��X
                    If Not IsNull(rsMainHeader.Fields(27)) Then Print #1, StrPadRightC(rsMainHeader.Fields(27), 1); Else Print #1, StrPadRightC(" ", 2);    '�p�⪫�y�O
                    If Not IsNull(rsMainHeader.Fields(28)) Then Print #1, StrPadRightC(rsMainHeader.Fields(28), 1); Else Print #1, StrPadRightC(" ", 1);    '�e�f�_
                    If Not IsNull(rsMainHeader.Fields(29)) Then Print #1, StrPadRightC(rsMainHeader.Fields(29), 2); Else Print #1, StrPadRightC(" ", 2);    '�q�����
                    If Not IsNull(rsMainHeader.Fields(30)) Then Print #1, StrPadRightC(rsMainHeader.Fields(30), 1); Else Print #1, StrPadRightC(" ", 1);    '�ꦬ�q�B�zMARK
                    If Not IsNull(rsMainHeader.Fields(31)) Then Print #1, StrPadRightC(rsMainHeader.Fields(31), 12); Else Print #1, StrPadRightC(" ", 12);  '�p���H
                    If Not IsNull(rsMainHeader.Fields(32)) Then Print #1, StrPadRightC(rsMainHeader.Fields(32), 20); Else Print #1, StrPadRightC(" ", 20);  '�q��
                    If Not IsNull(rsMainHeader.Fields(33)) Then Print #1, StrPadRightC(rsMainHeader.Fields(33), 12); Else Print #1, StrPadRightC(" ", 12);  '���N�m�W
                    If Not IsNull(rsMainHeader.Fields(34)) Then Print #1, StrPadRightC(rsMainHeader.Fields(34), 12); Else Print #1, StrPadRightC(" ", 12);  '�D�ީm�W
                    If Not IsNull(rsMainHeader.Fields(35)) Then Print #1, StrPadRightC(rsMainHeader.Fields(35), 50); Else Print #1, StrPadRightC(" ", 50);  '���e�Ȥ�
                    If Not IsNull(rsMainHeader.Fields(36)) Then Print #1, StrPadRightC(rsMainHeader.Fields(36), 7); Else Print #1, StrPadRightC(" ", 7);    '�w�p���
                    If Not IsNull(rsMainHeader.Fields(37)) Then Print #1, StrPadLeft(rsMainHeader.Fields(37), 8); Else Print #1, StrPadLeft(" ", 8);    '�B�O
                    If Not IsNull(rsMainHeader.Fields(38)) Then Print #1, StrPadRightC(rsMainHeader.Fields(38), 1); Else Print #1, StrPadRightC(" ", 1);    '�I�ڤ覡
                    If Not IsNull(rsMainHeader.Fields(39)) Then Print #1, StrPadRightC(rsMainHeader.Fields(39), 20); Else Print #1, StrPadRightC(" ", 20);  '�~�Ȥ��
                    If Not IsNull(rsMainHeader.Fields(40)) Then Print #1, StrPadRightC(rsMainHeader.Fields(40), 1); Else Print #1, StrPadRightC(" ", 1);    '�O�_���q�l�o��
                    If Not IsNull(rsMainHeader.Fields(41)) Then Print #1, StrPadRightC(rsMainHeader.Fields(41), 6); Else Print #1, StrPadRightC(" ", 6);    '�`���q
                    If Not IsNull(rsMainHeader.Fields(42)) Then Print #1, StrPadRightC(rsMainHeader.Fields(42), 4); Else Print #1, StrPadRightC(" ", 4);    '�H�d��4�X
                    If Not IsNull(rsMainHeader.Fields(43)) Then Print #1, StrPadLeft(rsMainHeader.Fields(43), 10); Else Print #1, StrPadLeft(" ", 10);  '�N���f��
                    If Not IsNull(rsMainHeader.Fields(44)) Then Print #1, StrPadRightC(rsMainHeader.Fields(44), 1); Else Print #1, StrPadRightC(" ", 1);    '�o���C�L�覡
                    If Not IsNull(rsMainHeader.Fields(45)) Then Print #1, StrPadRightC(rsMainHeader.Fields(45), 20); Else Print #1, StrPadRightC(" ", 20);  '�q��2
                    If Not IsNull(rsMainHeader.Fields(46)) Then Print #1, StrPadRightC(rsMainHeader.Fields(46), 8); Else Print #1, StrPadRightC(" ", 8);    '�έp�ﹳ
                    If Not IsNull(rsMainHeader.Fields(47)) Then Print #1, StrPadRightC(rsMainHeader.Fields(47), 3); Else Print #1, StrPadRightC(" ", 3);    '�����O
                    If Not IsNull(rsMainHeader.Fields(48)) Then Print #1, StrPadRightC(rsMainHeader.Fields(48), 3); Else Print #1, StrPadRightC(" ", 3);    '��F��
                    If Not IsNull(rsMainHeader.Fields(49)) Then Print #1, StrPadRightC(rsMainHeader.Fields(49), 2); Else Print #1, StrPadRightC(" ", 2);    '�Ӽh
                    If Not IsNull(rsMainHeader.Fields(50)) Then Print #1, StrPadRightC(rsMainHeader.Fields(50), 1); Else Print #1, StrPadRightC(" ", 1);    '�V�w�q��
                    If Not IsNull(rsMainHeader.Fields(51)) Then Print #1, StrPadRightC(rsMainHeader.Fields(51), 12); Else Print #1, StrPadRightC(" ", 12);  '���f��
                    If Not IsNull(rsMainHeader.Fields(52)) Then Print #1, StrPadRightC(rsMainHeader.Fields(52), 10); Else Print #1, StrPadRightC(" ", 10);  '�|��/�|�v
                    If Not IsNull(rsMainHeader.Fields(53)) Then Print #1, StrPadRightC(rsMainHeader.Fields(53), 40); Else Print #1, StrPadRightC(" ", 40);  '�Ȥ�²��
                    If Not IsNull(rsMainHeader.Fields(54)) Then Print #1, StrPadRightC(rsMainHeader.Fields(54), 7); Else Print #1, StrPadRightC(" ", 7);  '��ڨ�f��
                    If Not IsNull(rsMainHeader.Fields(55)) Then Print #1, StrPadRightC(rsMainHeader.Fields(55), 10); Else Print #1, StrPadRightC(" ", 10);  '���p�q�渹�X
                    Print #1, vbCrLf;
                End If
                rsMainHeader.MoveNext
            Loop
            Close #1
            rsMainHeader.MoveFirst

            '�q����������r��
            rsMainDetail.MoveFirst
            'strOrderNo = ""
            Set fso = New FileSystemObject
            FileName = Str_Company & "_rdb" & str_Date & ".txt"
            If Str_Company = "" Then FileName = "E_rdb" & str_Date & ".txt"
            If Dir(Str_Path1, vbDirectory) = "" Then MkDirs Str_Path1
            txtpath = Str_Path1 & "\" & FileName
            Open txtpath For Append As #1
            Do While Not rsMainHeader.EOF
              If rsMainHeader.Fields("�O�_�^��") = "V" Then
                rsMainDetail.Filter = "TMS�渹 = '" & rsMainHeader.Fields("TMS�渹") & "'"
                rsMainDetail.MoveFirst
                '�ư��ƶq���t�����~��
                Do While Not rsMainDetail.EOF
                '�p�G���X�q��h�n�^�ǡA�p�G���O�����X�q��h���^�Ǧ��t���~�����Ӷ� edit by Eric 20141001 Phil�q��
                    If rsMainHeader.Fields("TMS�渹") = rsMainDetail.Fields("TMS�渹") And rsMainHeader.Fields("�O�_�^��") = "V" And (rsMainHeader.Fields("����") = Str_Company Or Len(Str_Company) = 0) Then
                        ReturnOrderdetail = ReturnOrderdetail + 1
                        ReturnSignqty = ReturnSignqty + Val(rsMainDetail.Fields(7))
                        If Not IsNull(rsMainDetail.Fields(3)) Then Print #1, StrPadRightC(rsMainDetail.Fields(3), 8); Else Print #1, StrPadRightC(" ", 8);  '�q�渹�X
                        If Not IsNull(rsMainDetail.Fields(4)) Then Print #1, StrPadRightC(rsMainDetail.Fields(4), 16); Else Print #1, StrPadRightC(" ", 16);    '���~�s��
                        If Not IsNull(rsMainDetail.Fields(5)) Then Print #1, StrPadRightC(rsMainDetail.Fields(5), 60); Else Print #1, StrPadRightC(" ", 60);    '���~�W��
                        If Not IsNull(rsMainDetail.Fields(6)) Then Print #1, StrPadLeft(rsMainDetail.Fields(6), 10); Else Print #1, StrPadLeft(" ", 10);    '�q�f�q
                        If Not IsNull(rsMainDetail.Fields(8)) Then Print #1, StrPadLeft(rsMainDetail.Fields(8), 8); Else Print #1, StrPadLeft(" ", 8);  '���(���|)
                        If Not IsNull(rsMainDetail.Fields(9)) Then Print #1, StrPadLeft(rsMainDetail.Fields(9), 10); Else Print #1, StrPadLeft(" ", 10);    '�q�f���B(���|)
                        If Not IsNull(rsMainDetail.Fields(10)) Then Print #1, StrPadLeft(rsMainDetail.Fields(10), 8); Else Print #1, StrPadLeft(" ", 8);  '���(�t�|)
                        If Not IsNull(rsMainDetail.Fields(11)) Then Print #1, StrPadLeft(rsMainDetail.Fields(11), 10); Else Print #1, StrPadLeft(" ", 10);  '�q�f���B(�t�|)
                        If Not IsNull(rsMainDetail.Fields(7)) Then Print #1, StrPadLeft(rsMainDetail.Fields(7), 10); Else Print #1, StrPadLeft(" ", 10);  '�q�f�q-�ꦬ�q
                        If Not IsNull(rsMainDetail.Fields(12)) Then Print #1, StrPadRightC(rsMainDetail.Fields(12), 25); Else Print #1, StrPadRightC(" ", 25);  '��ڱ��X
                        If Not IsNull(rsMainDetail.Fields(13)) Then Print #1, StrPadLeft(rsMainDetail.Fields(13), 7); Else Print #1, StrPadLeft(" ", 7);    '�渹
                        If Not IsNull(rsMainDetail.Fields(14)) Then Print #1, StrPadRightC(rsMainDetail.Fields(14), 2); Else Print #1, StrPadRightC(" ", 2);    '���
                        If Not IsNull(rsMainDetail.Fields(15)) Then Print #1, StrPadRightC(rsMainDetail.Fields(15), 2); Else Print #1, StrPadRightC(" ", 2);    '�q�����
                        If Not IsNull(rsMainDetail.Fields(16)) Then Print #1, StrPadRightC(rsMainDetail.Fields(16), 1); Else Print #1, StrPadRightC(" ", 1);    '�o�����ӦC�L�_
                        If Not IsNull(rsMainDetail.Fields(17)) Then Print #1, StrPadRightC(rsMainDetail.Fields(17), 20); Else Print #1, StrPadRightC(" ", 20);  '������
                        Print #1, vbCrLf;
                    End If
                    rsMainDetail.MoveNext
                Loop
              End If
                rsMainHeader.MoveNext
                rsMainDetail.MoveFirst
            Loop
            Close #1
            rsMainHeader.MoveFirst
            rsMainDetail.MoveFirst

'Mark by Eric 20141210�A�p�_�M�j��o��ܥ_�ϭק�A�@�ֲ������\��
'    '��X�t�e���s
'        str_Orderkey = Mid(str_Orderkey, 1, Len(str_Orderkey) - 1)
'        Call Confirm_Recordset_Closed(tmp_Rs)
''        '�ɸ�ƪ�
''        str_SQL = "select �q�����=co.ordertype,�q�渹�X=co.externorderkey,���s�s��=s2.c_route_no,�e�f�_=co.DeliveryCode,�ﰪ���O�� =isnull(sum(sumreceivable),0) " & _
''                    "from sdn02t s2 join custorders co on s2.c_receipt_no = co.orderkey " & _
''                    "left join  sdn05t s5 on s5.sdn_no = s2.receipt_no and s5.costcode = 'forklift' " & _
''                    "where convert(char(8),co.adddate,112) between '20140801' and '20140901' " & _
''                    "group by co.OrderType,co.externorderkey,s2.c_route_no,co.DeliveryCode"
''
'        '���`��
'        str_SQL = "select �q�����=co.ordertype,�q�渹�X=co.externorderkey,���s�s��=s2.c_route_no,�e�f�_=co.DeliveryCode,�ﰪ���O�� =isnull(sum(sumreceivable),0) " & _
'                    "from sdn02t s2 join custorders co on s2.c_receipt_no = co.orderkey " & _
'                    "left join  sdn05t s5 on s5.sdn_no = s2.receipt_no and s5.costcode = 'forklift' " & _
'                    "where s2.c_receipt_no in (" & str_Orderkey & ") " & _
'                    "group by co.OrderType,co.externorderkey,s2.c_route_no,co.DeliveryCode "
'
'        tmp_Rs.Open str_SQL, cn
'
'            '���r��
'            tmp_Rs.MoveFirst
'            Set fso = New FileSystemObject
'            FileName = Str_Company & "_�t�e���s" & Str_Date & ".txt"
'            If Str_Company = "" Then FileName = "E_�t�e���s" & Str_Date & ".txt"
'            If Dir(Str_RoutePath, vbDirectory) = "" Then MkDirs Str_RoutePath
'            txtpath = Str_RoutePath & "\" & FileName
'            Open txtpath For Append As #1
'            Do While Not tmp_Rs.EOF
'                '�d�X�t�e���s
'                If Not IsNull(tmp_Rs.Fields(0)) Then Print #1, StrPadRightC(tmp_Rs.Fields(0), 2); Else Print #1, StrPadRightC(" ", 2);  '�q�����
'                If Not IsNull(tmp_Rs.Fields(1)) Then Print #1, StrPadRightC(tmp_Rs.Fields(1), 8); Else Print #1, StrPadRightC(" ", 8);    '�q�渹�X
'                If Not IsNull(tmp_Rs.Fields(2)) Then Print #1, StrPadRightC(tmp_Rs.Fields(2), 10); Else Print #1, StrPadRightC(" ", 10);    '���u�s��
'                If Not IsNull(tmp_Rs.Fields(3)) Then Print #1, StrPadRightC(tmp_Rs.Fields(3), 1); Else Print #1, StrPadLeft(" ", 1);    '�e�f�_
'                If Not IsNull(tmp_Rs.Fields(4)) Then Print #1, StrPadLeft(tmp_Rs.Fields(4), 9); Else Print #1, StrPadLeft(" ", 9);    '�ﰪ���O��
'                Print #1, vbCrLf;
'                tmp_Rs.MoveNext
'            Loop
'            Close #1
'            tmp_Rs.Close

'�������^�Ǫ��q�浧�ơA�q����ӼơA�q���`�^�Ƕq
txt_Msg.Text = txt_Msg.Text & Str_Company & "_rdb" & str_Date & ".txt : " & ReturnOrders & " Orders; " & ReturnOrderdetail & " Detail;Total Qty = " & ReturnSignqty & Chr(13) & Chr(10)

'�ƥ���FTP
'�ƥ��ɮ�
If Dir(Str_Path2, vbDirectory) = "" Then
    MkDirs Str_Path2
    If Len(Str_Company) = 0 Then Str_Company = "E"
    FileCopy Str_Path1 & "\" & Str_Company & "_rtb" & str_Date & ".txt", Str_Path2 & "\" & Str_Company & "_rtb" & str_Date & ".txt"
    FileCopy Str_Path1 & "\" & Str_Company & "_rdb" & str_Date & ".txt", Str_Path2 & "\" & Str_Company & "_rdb" & str_Date & ".txt"
    'FileCopy Str_RoutePath & "\" & Str_Company & "_�t�e���s" & Str_Date & ".txt", Str_FtpRoutePath & "\" & Str_Company & "_�t�e���s" & Str_Date & ".txt"
Else
    If Len(Str_Company) = 0 Then Str_Company = "E"
    FileCopy Str_Path1 & "\" & Str_Company & "_rtb" & str_Date & ".txt", Str_Path2 & "\" & Str_Company & "_rtb" & str_Date & ".txt"
    FileCopy Str_Path1 & "\" & Str_Company & "_rdb" & str_Date & ".txt", Str_Path2 & "\" & Str_Company & "_rdb" & str_Date & ".txt"
    'FileCopy Str_RoutePath & "\" & Str_Company & "_�t�e���s" & Str_Date & ".txt", Str_FtpRoutePath & "\" & Str_Company & "_�t�e���s" & Str_Date & ".txt"
End If

Exit Function
LogOnError:
'    rsMainHeader.Close
'    rsMainDetail.Close

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Function
Private Sub cmdOK_Click()
'
'If rsMain Is Nothing Then Exit Sub
'strOtQtyFixOrderkey = rsMain("TMS�渹")
'frm_OTQtyFix.Show vbModal
'
''��sDatagrid
'Call UpdateDatagrid

End Sub

Private Sub cmdOTUpdate_Click()

On Error GoTo err_Handle

'Dim Str_Date As String
'Str_Date = Format(Now(), "yyyymmddhhmmss")
'Call MBOrs2txt("C:\BEST\LMBO01\POD\�t�e���s", "\\192.168.200.200\ftp$\LMBO01\to_MaoBao\�t�e���s", Str_Date, "A")
Dim bl_Check1 As Boolean
bl_Check1 = False
txt_UnReciept.Text = ""
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.EOF Then Exit Sub
If rsMainDetail Is Nothing Then Exit Sub
If rsMainDetail.EOF Then Exit Sub

    Dim x As Integer, bl_Check As Boolean
    Screen.MousePointer = 11
    bl_Check = False
        
    '����datagrid
    dgMain_Header.Visible = False
    dgMain_Detail.Visible = False
    
    rsMainHeader.MoveFirst
    '�ˬd���L���A����檺�ܬO�_��i���w�^
     Do While Not rsMainHeader.EOF
                If rsMainHeader.Fields("�O�_�^��") = "V" Then
                '�ˬd���L����
                    str_SQL = "select  receipt_no,extern,sdnback from sdn02t where storerkey = 'LMBO01' and extern = '" & RTrim(rsMainHeader.Fields("�q�����")) & RTrim(rsMainHeader.Fields("�q�渹�X")) & "'"
                    Call Confirm_Recordset_Closed(tmp_Rs)
                    tmp_Rs.CursorLocation = adUseClient
                    tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic ' adOpenKeyset
                    tmp_Rs.MoveFirst
                    Do While Not tmp_Rs.EOF
                        If RTrim(tmp_Rs.Fields("sdnback")) = "0" Then
                            Screen.MousePointer = 0
                            dgMain_Header.Visible = True
                            dgMain_Detail.Visible = True
                            MsgBox "TMS�渹:" & RTrim(tmp_Rs.Fields("receipt_no")) & ",�f�D�渹:" & RTrim(tmp_Rs.Fields("extern")) & ",������檺�����S���^�ӡA�L�k�^�ǡC�нT�{!", vbOKOnly + vbCritical, "���w�^�ˬd"
                            tmp_Rs.Close
                            Exit Sub
                        End If
                        tmp_Rs.MoveNext
                    Loop
                    tmp_Rs.Close
                End If
            rsMainHeader.MoveNext
        Loop
        
'    rsMainHeader.MoveFirst
'    '�X�f�����A�p�G�����`ñ��h�n�ˬd���L�����A���X�q��h�n�P�_�O�_���X�f�A�����ܭn�����A�S�����ܤ���
'    If optOut.Value = True Then
'        Do While Not rsMainHeader.EOF
'            If rsMainHeader.Fields("�O�_�^��") = "V" And rsMainHeader.Fields("���A") = "���`�q��" Then
'            '�ˬd���L����
'                str_SQL = "select externreceiptkey from " & strWMSDB & "..receipt where storerkey = 'LMBO01' and status = '9' and receipttype = 'A' and externreceiptkey = '" & RTrim(rsMainHeader.Fields("�q�����")) & RTrim(rsMainHeader.Fields("�q�渹�X")) & "'"
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.CursorLocation = adUseClient
'                tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic ' adOpenKeyset
'                If tmp_Rs.EOF = True Then
'                '������
'                txt_UnReciept.Text = txt_UnReciept.Text & RTrim(rsMainHeader.Fields("�q�����")) & RTrim(rsMainHeader.Fields("�q�渹�X")) & ","
'                rsMainHeader.Fields("�O�_�^��") = " " '�����^��
'                End If
'                tmp_Rs.Close
'            End If
'            If rsMainHeader.Fields("�O�_�^��") = "V" And rsMainHeader.Fields("���A") = "���X�q��" Then
'            '���ˬd���L�t�m�A���h�n�P�_���L����
'                str_SQL = "select shippedqty=isnull(sum(shippedqty),0) from " & strWMSDB & "..orderdetail where storerkey = 'LMBO01' and status = '9' and externorderkey = '" & RTrim(rsMainHeader.Fields("�q�����")) & RTrim(rsMainHeader.Fields("�q�渹�X")) & "'"
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.CursorLocation = adUseClient
'                tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic ' adOpenKeyset
'                If Val(tmp_Rs.Fields("shippedqty")) > 0 Then
'                '���X�f�A�ˬd���L����
'                str_SQL = "select externreceiptkey from " & strWMSDB & "..receipt where storerkey = 'LMBO01' and status = '9' and receipttype = 'A' and externreceiptkey = '" & RTrim(rsMainHeader.Fields("�q�����")) & RTrim(rsMainHeader.Fields("�q�渹�X")) & "'"
'                Call Confirm_Recordset_Closed(rs_Receipt)
'                rs_Receipt.CursorLocation = adUseClient
'                rs_Receipt.Open str_SQL, cn, adOpenKeyset, adLockPessimistic ' adOpenKeyset
'                If rs_Receipt.EOF = True Then
'                '������
'                    txt_UnReciept.Text = txt_UnReciept.Text & RTrim(rsMainHeader.Fields("�q�����")) & RTrim(rsMainHeader.Fields("�q�渹�X")) & ","
'                    rsMainHeader.Fields("�O�_�^��") = " " '�����^��
'                End If
'                rs_Receipt.Close
'                End If
'                tmp_Rs.Close
'            End If
'
'            rsMainHeader.MoveNext
'        Loop
'    End If
'
'    If optIn.Value = True Then
' Do While Not rsMainHeader.EOF
'            If rsMainHeader.Fields("�O�_�^��") = "V" Then
'            '�ˬd���L����
'                str_SQL = "select externreceiptkey from " & strWMSDB & "..receipt where storerkey = 'LMBO01' and status = '9' and externreceiptkey = '" & RTrim(rsMainHeader.Fields("�q�����")) & RTrim(rsMainHeader.Fields("�q�渹�X")) & "'"
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.CursorLocation = adUseClient
'                tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic ' adOpenKeyset
'                If tmp_Rs.EOF = True Then
'                '������
'                txt_UnReciept.Text = txt_UnReciept.Text & RTrim(rsMainHeader.Fields("�q�����")) & RTrim(rsMainHeader.Fields("�q�渹�X")) & ","
'                rsMainHeader.Fields("�O�_�^��") = " " '�����^��
'                End If
'                tmp_Rs.Close
'            End If
'            rsMainHeader.MoveNext
'        Loop
'    End If
    rsMainHeader.MoveFirst
    '�ˬd���Ŀ�^��
    Do While Not rsMainHeader.EOF
        If rsMainHeader.Fields("�O�_�^��") = "V" Then
            bl_Check = True
        End If
        rsMainHeader.MoveNext
    Loop
    
    If bl_Check = False Then Screen.MousePointer = 0: dgMain_Header.Visible = True:     dgMain_Detail.Visible = False: Exit Sub
    rsMainHeader.Filter = "�O�_�^�� = 'V'"
    
    '���q��q�O�_����ñ��q
        rsMainHeader.MoveFirst
        rsMainDetail.MoveFirst
    Do While Not rsMainHeader.EOF
       If rsMainHeader.Fields("�O�_�^��") = "V" Then
       rsMainDetail.Filter = "TMS�渹 = '" & rsMainHeader.Fields("TMS�渹") & "'"
       rsMainDetail.MoveFirst
            Do While Not rsMainDetail.EOF
                If rsMainHeader.Fields("TMS�渹") = rsMainDetail.Fields("TMS�渹") Then
                    If Abs(Val(rsMainDetail.Fields("�q�f�q"))) <> Val(rsMainDetail.Fields("�q�f�q-�ꦬ�q")) Then
                        x = MsgBox("�q�渹�X:" & rsMainDetail.Fields("TMS�渹") & "����:" & rsMainDetail.Fields("TMS�渹����") & ":�q�f�q<>ñ���q�A�нT�{�O�_��s�^��?", vbQuestion + vbYesNo, "�ƶq�ˬd")
                        If x = 6 Then
                            '�O��
                        Else
                            rsMainHeader.Fields("�O�_�^��") = " "
                            GoTo next1
                            Exit Sub
                        End If
                    End If
                End If
                rsMainDetail.MoveNext
            Loop
       End If
next1:
        rsMainDetail.MoveFirst
        rsMainHeader.MoveNext
    Loop
'
'    rsMainHeader.MoveFirst
'    '�ˬd���L�^�ǤF
'    Do While Not rsMainHeader.EOF
'        If rsMainHeader.Fields("�O�_�^��") = "V" Then
'            str_SQL = "select returnstatus from sdn02t where c_receipt_no = '" & rsMainHeader.Fields("TMS�渹") & "'"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.CursorLocation = adUseClient
'            tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
'
'            If tmp_Rs.EOF = True Then
'                Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
'            Else
'                Do While Not tmp_Rs.EOF
'                    If RTrim(tmp_Rs.Fields("returnstatus")) = "1" Or RTrim(tmp_Rs.Fields("returnstatus")) = "2" Then MsgBox rsMainHeader.Fields("TMS�渹") & "���w�^�Ǹ��,�^�Ǥ���", vbCritical + vbOKOnly, "�^���ˬd": tmp_Rs.Close:    Screen.MousePointer = 0: Exit Sub
'                    tmp_Rs.MoveNext
'                Loop
'            End If
'            tmp_Rs.Close
'        End If
'        rsMainHeader.MoveNext
'    Loop
'
'    '�ˬd���O�_���X���T�{�F
'    rsMainHeader.MoveFirst
'    Do While Not rsMainHeader.EOF
'      If rsMainHeader.Fields("�O�_�^��") = "V" Then
'            str_SQL = "exec es_CheckConfirm '" & rsMainHeader.Fields("TMS�渹") & "'"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'            If Not tmp_Rs.EOF Then
'                MsgBox "TMS�渹:" & rsMainHeader.Fields("TMS�渹") & "�����渹:" & tmp_Rs.Fields("receipt_no") & "���X���A�нT�{��A�A�i��^�ǧ@�~�A�^�ǲפ�", vbCritical + vbOKOnly, "�^���ˬd"
'                tmp_Rs.Close: Screen.MousePointer = 0: Exit Sub
'            End If
'
'        End If
'                    rsMainHeader.MoveNext
'    Loop
'    rsMainHeader.MoveFirst: tmp_Rs.Close
    
'    '�ˬd�^�Ǫ�TMS�渹�����A�O�_�w�g�^�ӤF
'    rsMainHeader.MoveFirst
'    Do While Not rsMainHeader.EOF
'      If rsMainHeader.Fields("�O�_�^��") = "V" Then
'        str_SQL = "select ���TMS = s2.receipt_no,ñ�檬�A = s2.sdnback from sdn02t s2 where s2.c_receipt_no = '" & rsMainHeader.Fields("TMS�渹") & "'"
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        tmp_Rs.MoveFirst
'        Do While Not tmp_Rs.EOF
'            If tmp_Rs.Fields("ñ�檬�A") = "0" Then
'                MsgBox "TMS�渹:" & rsMainHeader.Fields("TMS�渹") & "�������渹:" & tmp_Rs.Fields("���TMS") & "���^�A�нT�{��A�A�i��^�ǧ@�~�A�^�ǲפ�", vbCritical + vbOKOnly, "�^���ˬd"
'                tmp_Rs.Close: Screen.MousePointer = 0:
'                dgMain_Header.Visible = True
'                dgMain_Detail.Visible = True
'                Exit Sub
'            End If
'            tmp_Rs.MoveNext
'        Loop
'        End If
'        rsMainHeader.MoveNext
'    Loop
'
    Tran_Level = cn.BeginTrans:
'
'    '�ˬdCO�q�檺receiptdetail�ꦬ�q�O�_���W��
'    rsMainHeader.MoveFirst
'    If RTrim(rsMainHeader.Fields("�q�����")) = "CO" Or RTrim(rsMainHeader.Fields("�q�����")) = "SC" Then
'        Do While Not rsMainHeader.EOF
'          If rsMainHeader.Fields("�O�_�^��") = "V" Then
'            str_SQL = "select �q�渹�X = isnull(a.externasnkey,r.externreceiptkey),�~��=isnull(a.sku,r.sku),�q���q =sum(isnull(a.qty,0)) ,�ꦬ�q = sum(isnull(r.qty,0)) " & _
'                        "from ( " & _
'                        "select a.externasnkey,ad.sku,qty=sum(isnull(ad.qtyordered,0)) " & _
'                        "From " & strWMSDB & "..asn a join " & strWMSDB & "..asndetail ad on a.asnkey = ad.asnkey " & _
'                        "where a.externasnkey = '" & RTrim(rsMainHeader.Fields("�q�����")) & RTrim(rsMainHeader.Fields("�q�渹�X")) & "' " & _
'                        "group by a.externasnkey,ad.sku " & _
'                        ") a full join " & _
'                        "( " & _
'                        "select r.externreceiptkey,rd.sku,qty=sum(isnull(rd.qtyreceived,0)) " & _
'                        "from " & strWMSDB & "..receipt r join " & strWMSDB & "..receiptdetail rd on r.receiptkey = rd.receiptkey " & _
'                        "where r.externreceiptkey = '" & RTrim(rsMainHeader.Fields("�q�����")) & RTrim(rsMainHeader.Fields("�q�渹�X")) & "' and r.status = '9' " & _
'                        "group by r.externreceiptkey,rd.sku " & _
'                        ") r on a.externasnkey = r.externreceiptkey and a.sku = r.sku " & _
'                        "group by isnull(a.externasnkey,r.externreceiptkey),isnull(a.sku,r.sku) "
'
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'              tmp_Rs.MoveFirst
'              Do While Not tmp_Rs.EOF
'                If Val(tmp_Rs.Fields("�q���q")) < Val(tmp_Rs.Fields("�ꦬ�q")) Then
'                    '���ͮt����
'                    tmp_Rs.MoveFirst
'                    Recordset2Excel "�t����", tmp_Rs
'                    Screen.MousePointer = 0
'                    cn.RollbackTrans: Tran_Level = 0
'                    dgMain_Header.Visible = True
'                    dgMain_Detail.Visible = True
'                    MsgBox "�q�渹�X:" & RTrim(tmp_Rs.Fields("�q�渹�X")) & " �~��:" & RTrim(tmp_Rs.Fields("�~��")) & " �q���q:" & RTrim(tmp_Rs.Fields("�q���q")) & " <> �ꦬ�q:" & RTrim(tmp_Rs.Fields("�ꦬ�q")) & "�A�L�k�^��", vbCritical + vbOKOnly, "�^���ˬd"
'                    tmp_Rs.Close
'                    Exit Sub
'                End If
'                tmp_Rs.MoveNext
'              Loop
'              tmp_Rs.Close
'              cn.Execute "update " & strWMSDB & "..asn set status = '9' where externasnkey = '" & RTrim(rsMainHeader.Fields("�q�����")) & RTrim(rsMainHeader.Fields("�q�渹�X")) & "'", RowsAffect, adExecuteNoRecords
'          End If
'              rsMainHeader.MoveNext
'        Loop
'
'    End If
    rsMainHeader.MoveFirst
    rsMainDetail.MoveFirst
    
    If rsMainHeader.EOF Then
        MsgBox "�d�L��ƥi�����ɡI", vbOKOnly + vbInformation, Me.Caption
        Screen.MousePointer = 0:
        cn.RollbackTrans: Tran_Level = 0
    Else
        Do While Not rsMainHeader.EOF
          If rsMainHeader.Fields("�O�_�^��") = "V" Then
                '�ˬdreturnstatus = 2���h���i�H��s��1
                str_SQL = "select returnstatus from sdn02t where c_receipt_no = '" & rsMainHeader.Fields("TMS�渹") & "'"
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                If tmp_Rs.Fields("returnstatus") = "2" Then Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0: dgMain_Header.Visible = True: dgMain_Detail.Visible = True: MsgBox "TMS�渹:" & rsMainHeader.Fields("TMS�渹") & "�w�^�ǡA�L�k�A�^��!�^�ǲפ�", vbOKOnly + vbCritical, "�^���ˬd": tmp_Rs.Close: Exit Sub
                tmp_Rs.Close
                '��sretrunstatus
                str_SQL = "update sdn02t set returnstatus = '1' where c_receipt_no = '" & rsMainHeader.Fields("TMS�渹") & "'"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
            rsMainHeader.MoveNext
        Loop
    End If
    rsMainHeader.MoveFirst
    
    '�^�Ǩ�FTP
    Call cmd2Excel_Click
    cn.CommitTrans: Tran_Level = 0
    cmdOTUpdate.Enabled = False
    dgMain_Header.Visible = True
    dgMain_Detail.Visible = True
    'Send Mail�q���Ȥ�
    Call cmd_SendMail_Click
    Screen.MousePointer = 0: rsMainHeader.Filter = "": rsMainDetail.Filter = "": rsMainHeader.Close: rsMainDetail.Close: txt_Msg = ""
    MsgBox "ñ���q�w�^��^_^�äwmail�q���Ȥ�C", vbOKOnly, "POD�^�Ǧ��\"
    
Exit Sub
err_Handle:
Screen.MousePointer = 0
    dgMain_Header.Visible = True
    dgMain_Detail.Visible = True
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Public Sub cmdPrintReport_Click()
'Dim i As Integer, j As Integer, k As Integer
'On Error GoTo err_Handle
'
'If rsMain Is Nothing Then MsgBox "�L��ƥi�ѦC�L�I", vbOKOnly + vbInformation, "����C�L": Exit Sub
'
''���w������
'If RTrim(strOtQtyFixOrderkey) <> "" Then
'    rsMain.Filter = "(TMS�渹 = " & strOtQtyFixOrderkey & ")"
'Else
'    rsMain.Filter = "(�� = 'V')"
'End If
'
'If rsMain.RecordCount = 0 Then rsMain.Filter = 0: MsgBox "�п�����C�L����ơC", 64, "�C�L": rsMain.Sort = "�s��": Exit Sub
'
'Screen.MousePointer = 11
'
''��Ƽg�J Access ��Ʈw
'Call AccessDB_Connect
'cnAccess.BeginTrans
'Tran_Level = cn.BeginTrans
'
'cnAccess.Execute "Delete From �X�f���", RowsAffect, adExecuteNoRecords
'
'Dim rs_Access As New ADODB.Recordset
'rs_Access.Open "�X�f���", cnAccess, adOpenStatic, adLockOptimistic
'
'rsMain.MoveFirst
'
'Do While Not rsMain.EOF
'    For j = 1 To rsMain("�X�f���") '�@��g�J�@��
'        rs_Access.AddNew
'
'        For i = 0 To rsMain.Fields.Count - 1 '�g�J�C�����
'            rs_Access.Fields(i).Value = rsMain.Fields(i).Value
'        Next i
'
'        rs_Access.Fields(i).Value = j
'        rs_Access.Fields(i + 1).Value = rsMain("�X�f���")
'        rs_Access.Update
'    Next j
'
'    'TRP02T��s���w�^��
'    str_SQL = "update trp02t set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS�渹")) & "' "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    'TRP02W��s���w�^��
'    str_SQL = "update TRP02w set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS�渹")) & "' "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    'ORT02T��s���w�^��
'    str_SQL = "update ort02t set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS�渹")) & "' "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    'ORT02W��s���w�^��
'    str_SQL = "update ort02w set otprinttimes = otprinttimes + 1 , otprintdate = getdate() where receipt_no = '" & RTrim(rsMain("TMS�渹")) & "' "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    rsMain("�C�L����") = rsMain("�C�L����") + 1
'    rsMain("�C�L�ɶ�") = Format(Now, "yyyy/mm/dd hh:mm:ss")
'
'   rsMain.MoveNext
'Loop
'
'cn.CommitTrans: Tran_Level = 0
'cnAccess.CommitTrans
'
'Call DB_Disconnect(cnAccess)
'
'strAccessDBFileName_FullPath = GetAccessDBFileName
'Dim MSAccessAP As New access.Application
'With MSAccessAP
'    .OpenCurrentDatabase (strAccessDBFileName_FullPath)
'
'    If chkPrintPreView.Value = vbChecked Then
'    '�w���C�L
'         .DoCmd.OpenReport "�X�f���", acViewPreview
'        .DoCmd.Maximize
'        .Visible = True
'    Else
'    '�����C�L�ܦL���
'        .Visible = False
'        .DoCmd.OpenReport "�X�f���", acViewNormal
'        .CloseCurrentDatabase
'        .Quit
'        Set MSAccessAP = Nothing
'End If
'
'End With
'rsMain.Filter = 0
'rsMain.Sort = "�s��"
'Screen.MousePointer = 0
'strOtQtyFixOrderkey = ""
'Exit Sub
'
'err_Handle:
'Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11

Set dgMain_Header.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ���q���ƦC"
Set dgMain_Detail.DataSource = Nothing: 'StatusBar.Panels(2).Text = "0 ����ƦC"

Dim chc_DeliveryDate As String, chc_ExternOrderkey, chc_Status As String, chc_Storerkey As String, chc_Carno As String, chc_Print As String, str_WhereExternorderkey As String
str_WhereExternorderkey = ""
''���ˬd���Lasn�аO���Φ^�Ǫ����
'If optIn = True Then
'    str_SQL = "update s2 set s2.returnstatus = '3' " & _
'        "from " & strWMSDB & "..asn a join custorders co on a.externasnkey = rtrim(co.ordertype) + rtrim(co.externorderkey) " & _
'        "join sdn02t s2 on s2.c_receipt_no = co.orderkey " & _
'        "where a.asntype = 'R' and a.status = 2  and s2.returnstatus <> 3 "
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'End If

'��f���
chc_DeliveryDate = ""
If Len(RTrim(txtDeliveryDateS.Text)) > 0 And Len(RTrim(txtDeliveryDateE.Text)) > 0 Then
   chc_DeliveryDate = "and convert(char(8),o.deliveryDate,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(RTrim(txtDeliveryDateS.Text)) > 0 And Len(RTrim(txtDeliveryDateE.Text)) = 0 Then
   chc_DeliveryDate = "and convert(char(8),o.deliveryDate,112) = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(RTrim(txtDeliveryDateS.Text)) = 0 And Len(RTrim(txtDeliveryDateE.Text)) > 0 Then
   chc_DeliveryDate = "and convert(char(8),o.deliveryDate,112) = '" & txtDeliveryDateE.Text & "' "
End If
'
''�f�D�渹
'chc_ExternOrderkey = ""
'If Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) > 0 Then
'   chc_ExternOrderkey = "and o.externorderkey between '" & txtExternOrderkeyS.Text & "' and '" & txtExternOrderkeyE.Text & "' "
'ElseIf Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) = 0 Then
'   chc_ExternOrderkey = "and o.externorderkey = '" & txtExternOrderkeyS.Text & "' "
'ElseIf Len(txtExternOrderkeyS.Text) = 0 And Len(txtExternOrderkeyE.Text) > 0 Then
'   chc_ExternOrderkey = "and o.externorderkey = '" & txtExternOrderkeyE.Text & "' "
'End If

'��ƪ��A
chc_Status = ""
If optNo = True Then chc_Status = chc_Status & "and s2.ReturnStatus = 0 "
If optYes = True Then chc_Status = chc_Status & "and s2.ReturnStatus <> 0 "
If optOut = True Then chc_Status = chc_Status & "and co.ordertype not in ('CO','SC') "
If optIn = True Then chc_Status = chc_Status & "and co.ordertype in ('CO','SC') "


'�f�D
chc_Storerkey = ""
If Len(RTrim(cboStorerkey.Text)) > 0 Then chc_Storerkey = " and o.storerkey = '" & RTrim(cboStorerkey.Text) & "' "
chc_Storerkey = "LMBO01"


If optOut.Value = True Then
'�X�f�����u�D�Xñ��q=�q��q���q��ñ��C edit by Eric 20141003�APhil�q��,20150122�y��F���^�ǡC
        str_SQL = "select externorderkey = rtrim(co.ordertype)+rtrim(co.ExternOrderkey) " & _
                    "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no and s2.storerkey = s3.storerkey " & _
                    "join CustOrders co on s2.c_receipt_no = co.orderkey " & _
                    "join orders o on co.orderkey = o.orderkey " & _
                    "join CustOrderdetail cod on  co.orderkey = cod.orderkey  and s3.seq_no = cod.orderlinenumber " & _
                    "where s2.storerkey = '" & chc_Storerkey & "' and s2.sdnback = '1'  " & chc_Status & chc_DeliveryDate & _
                    "and co.address not like  '%�x�F%' and address not like '%�Ὤ%' and address not like'%�y��%' " & _
                    "and co.City not in ('260','261','262','263','264','265','266','267','268','269','270','272','290','950','951','952','953','954','955','956','957','958','959','961','962','963','964','965','966','970','971','972','973','974','975','976','977','978','979','981','982','983') " & _
                    "and co.Administration not in ('015','016','017') " & _
                    "group by co.ExternOrderkey,co.ordertype " & _
                    "having sum(s3.sign_qty) = sum(cast(cod.originalqty as float)) "
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = adUseClient
        tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic   '
        
        If tmp_Rs.EOF = True Then
            tmp_Rs.Close: Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
        Else
            '�N�Ҧ��ŦX���󪺭q�渹�X��_��
            tmp_Rs.MoveFirst
            Do While Not tmp_Rs.EOF
                str_WhereExternorderkey = str_WhereExternorderkey & "'" & RTrim(tmp_Rs.Fields("externorderkey")) & "',"
                tmp_Rs.MoveNext
            Loop
            str_WhereExternorderkey = Mid(str_WhereExternorderkey, 1, Len(str_WhereExternorderkey) - 1)
            str_WhereExternorderkey = "(" & str_WhereExternorderkey & ")"
            tmp_Rs.Close: Set tmp_Rs = Nothing
        End If
Else
'�h�f����
'�ꦬ�q=�q��q���q��
        str_SQL = "select externorderkey = rtrim(co.ordertype)+rtrim(co.ExternOrderkey) ,�q���q=sum(cast(cod.originalqty as float)), " & _
                  "�ꦬ�q = (select isnull(sum(rd.QtyReceived),0) from " & strWMSDB & "..receipt r join " & strWMSDB & "..receiptdetail rd on r.receiptkey = rd.receiptkey and r.status = '9' and r.storerkey = 'LMBO01' and r.receipttype = 'R' where r.externreceiptkey =  rtrim(co.ordertype)+rtrim(co.ExternOrderkey)) " & _
                    "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no and s2.storerkey = s3.storerkey " & _
                    "join CustOrders co on s2.c_receipt_no = co.orderkey " & _
                    "join orders o on co.orderkey = o.orderkey " & _
                    "join CustOrderdetail cod on  co.orderkey = cod.orderkey  and s3.seq_no = cod.orderlinenumber " & _
                    "where s2.storerkey = '" & chc_Storerkey & "' and s2.sdnback = '1'  " & chc_Status & chc_DeliveryDate & _
                    "group by co.ExternOrderkey,co.ordertype"
                    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = adUseClient
        tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic   '
        
        If tmp_Rs.EOF = True Then
            tmp_Rs.Close: Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
        Else
            '�N�Ҧ��ŦX���󪺭q�渹�X��_��
            tmp_Rs.MoveFirst
            Do While Not tmp_Rs.EOF
                If Abs(Val(RTrim(tmp_Rs.Fields("�ꦬ�q")))) = Abs(Val(RTrim(tmp_Rs.Fields("�q���q")))) Then
                    str_WhereExternorderkey = str_WhereExternorderkey & "'" & RTrim(tmp_Rs.Fields("externorderkey")) & "',"
                End If
                tmp_Rs.MoveNext
            Loop
            If str_WhereExternorderkey = "" Then
                tmp_Rs.Close: Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
            Else
                str_WhereExternorderkey = Mid(str_WhereExternorderkey, 1, Len(str_WhereExternorderkey) - 1)
                str_WhereExternorderkey = "(" & str_WhereExternorderkey & ")"
                tmp_Rs.Close: Set tmp_Rs = Nothing
            End If
        End If
End If


'�զX�r��
str_SQL = "select distinct " & _
        "'�O�_�^��' = ' ' " & _
        ",TMS�渹=co.orderkey " & _
        ",�����q�N��=co.BranchId,�q�渹�X=co.ExternOrderkey,�q����=isnull(rtrim(cast(cast(convert(char(4),co.OrderDate,112) as int ) - 1911 as char)) + right(convert(char(8),co.OrderDate,112),4),'') " & _
        ",�o�����X=co.Invoice,�o�����X�ˬd�X=co.InvoiceCheck,�o�����=isnull(rtrim(cast(cast(convert(char(4),co.InvoiceDate,112) as int ) - 1911 as char)) + right(convert(char(8),co.InvoiceDate,112),4),'') " & _
        ",�Ȥ�s��=co.Consigneekey,�Ȥ�W��=co.Full_Name,�~�N�N��=co.SalesCode,�U�f���{=co.COD,�e�f�a�}=co.Address,�p��=co.Coupled,�Τ@�s��=co.VAT " & _
        ",�������B=cast(co.Allowance as float),�ƶq�������B=cast(co.QuantityAllowance as float),�S�O�������B=cast(co.SpecialAllowance as float),�{������=cast(co.CashAllowance as float) " & _
        ",�f��=cast(co.Amount as float),�|�e���B=cast(co.NetAmount as float),�|�B=cast(co.Tax as float),�Ƶ�=co.Notes,�Ȥ�q��s��=co.CustOrderkey,�H�f���o���X=co.InvoiceCode " & _
        ",�H�f���q��X=co.OrderCode,�p�⪫�y�O=co.LogisticsCode,�e�f�_=co.DeliveryCode,�q�����=co.OrderType,�ꦬ�q�B�zMARK=co.PaidMARK,�s���H=co.Contact " & _
        ",�q��=co.Phone1,�~�N�m�W=co.SalesName,�D�ީm�W=co.LeaderName,���e�Ȥ�=co.Address2,�w�p���=isnull(rtrim(cast(cast(convert(char(4),co.DeliveryDate,112) as int ) - 1911 as char)) + right(convert(char(8),co.DeliveryDate,112),4),'') " & _
        ",�B�O=co.Freight,�I�ڤ覡=co.Payment,�~�Ȥ��=co.SalesPhone,�O�_���q�l�o��=co.EInvoiceMark,�`���q=cast(co.TotalWeight as float),�H�d��4�X=co.Credit_Last4 " & _
        ",�N���f��=cast(co.Cash as float),�o���C�L�覡=co.InvoicePrint,�q��2=co.Phone2,�έp��H=co.ExternNumber,�����O=co.City,��F��=co.Administration,�Ӽh=co.Stairs " & _
        ",�V�w�q��=co.CrossCode,���f��=co.Storage,'�|��/�|�v'=co.InvoiceArea,�Ȥ�²��=co.short_name,��ڨ�f��=isnull(rtrim(cast(cast(convert(char(4),o.DeliveryDate,112) as int ) - 1911 as char)) + right(convert(char(8),o.DeliveryDate,112),4),''),���p�q��=rtrim(co.connectorderkey) " & _
        ",���A = isnull((select top 1 sdn.confirm_notes from sdn02t sdn where sdn.c_receipt_no  = s2.c_receipt_no and sdn.confirm_notes in ('���`�q��','���X�q��')),'���`�q��') " & _
        ",����=case when co.ExternNumber in ('10000545','10020700') then 'SNRT' else 'Other' end " & _
        "from sdn02t s2 join CustOrders co on s2.c_receipt_no = co.orderkey " & _
        "join orders o on co.orderkey = o.orderkey " & _
        "where s2.storerkey = '" & chc_Storerkey & "' and s2.sdnback = '1'  " & chc_Status & chc_DeliveryDate & _
        "and rtrim(co.ordertype) + rtrim(co.externorderkey) in " & str_WhereExternorderkey & _
        "order by co.orderkey "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic   '

If tmp_Rs.EOF = True Then
    tmp_Rs.Close: Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
Else
    Call Replication_Recordset(tmp_Rs, rsMainHeader)
    tmp_Rs.Close: Set tmp_Rs = Nothing
    
    Set dgMain_Header.DataSource = rsMainHeader: dgMain_Header.Visible = False
    rsMainHeader.MoveFirst
    
    With dgMain_Header
    Set dgMain_Header.DataSource = rsMainHeader
    
    '    .ColumnHeaders = True        '���D�����
    '    .RowHeight = 300
    '    .Columns(0).Alignment = dbgCenter
    '    .Columns(10).Alignment = dbgRight
    
    End With
End If
SetDataGridColWidth Me.Caption, dgMain_Header



'����
If optIn = True Then '�h�f
    str_SQL = "select " & _
            "TMS�渹=co.orderkey " & _
            ",TMS�渹���� = cod.orderlinenumber " & _
            ",�q�渹�X=cod.ExternOrderkey " & _
            ",���~�s��=cod.Sku " & _
            ",���~�W��=cod.Descr " & _
            ",�q�f�q=cast(cod.OriginalQty as float) " & _
            ",'�q�f�q-�ꦬ�q'= abs(cast(cod.OriginalQty as float)) " & _
            ",'���(���|)'=cast(cod.UnitNetPrice as float) " & _
            ",'�q�f���B(���|)'=cast(cod.NetPrice as float) " & _
            ",'���(�t�|)'=cast(cod.UnitGrossPrice as float) " & _
            ",'�q�f���B(�t�|)'=cod.GrossPrice " & _
            ",��ڱ��X=cod.BarCode " & _
            ",�渹=cast(cod.Externlineno as float) " & _
            ",���=cod.UOM " & _
            ",�q�����=cod.Ordertype " & _
            ",�o�����ӦC�L�_=cod.InvoicePCode " & _
            ",������=cod.Acceptance " & _
            "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no and s2.storerkey = s3.storerkey " & _
            "join CustOrders co on s2.c_receipt_no = co.orderkey join orders o on co.orderkey = o.orderkey  " & _
            "join CustOrderdetail cod on  co.orderkey = cod.orderkey  and s3.seq_no = cod.orderlinenumber " & _
            "where s2.storerkey = 'LMBO01' and s2.sdnback = '1'  " & chc_Status & chc_DeliveryDate & _
            "and rtrim(co.ordertype) + rtrim(co.externorderkey) in " & str_WhereExternorderkey & _
            "group by co.orderkey ,cod.orderlinenumber ,cod.ExternOrderkey ,cod.Sku ,cod.Descr ,cast(cod.OriginalQty as float) ,cast(cod.UnitNetPrice as float) , " & _
            "cast(cod.NetPrice as float) ,cod.GrossPrice,cast(cod.UnitGrossPrice as float) ,cod.BarCode ,cast(cod.Externlineno as float) ,cod.UOM ,cod.Ordertype ,cod.InvoicePCode ,cod.Acceptance,co.OrderType,co.ExternOrderkey,abs(cast(cod.OriginalQty as float)) order by co.orderkey,cod.sku"
            
Else
    str_SQL = "select " & _
            "TMS�渹=co.orderkey " & _
            ",TMS�渹���� = cod.orderlinenumber " & _
            ",�q�渹�X=cod.ExternOrderkey " & _
            ",���~�s��=cod.Sku " & _
            ",���~�W��=cod.Descr " & _
            ",�q�f�q=cast(cod.OriginalQty as float) " & _
            ",'�q�f�q-�ꦬ�q'=sum(s3.sign_qty) " & _
            ",'���(���|)'=cast(cod.UnitNetPrice as float) " & _
            ",'�q�f���B(���|)'=cast(cod.NetPrice as float) " & _
            ",'���(�t�|)'=cast(cod.UnitGrossPrice as float) " & _
            ",'�q�f���B(�t�|)'=cod.GrossPrice " & _
            ",��ڱ��X=cod.BarCode " & _
            ",�渹=cast(cod.Externlineno as float) " & _
            ",���=cod.UOM " & _
            ",�q�����=cod.Ordertype " & _
            ",�o�����ӦC�L�_=cod.InvoicePCode " & _
            ",������=cod.Acceptance " & _
            "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no and s2.storerkey = s3.storerkey " & _
            "join CustOrders co on s2.c_receipt_no = co.orderkey join orders o on co.orderkey = o.orderkey " & _
            "join CustOrderdetail cod on  co.orderkey = cod.orderkey  and s3.seq_no = cod.orderlinenumber " & _
            "where s2.storerkey = 'LMBO01' and s2.sdnback = '1'  " & chc_Status & chc_DeliveryDate & _
            "and rtrim(co.ordertype) + rtrim(co.externorderkey) in " & str_WhereExternorderkey & _
            "group by co.orderkey ,cod.orderlinenumber ,cod.ExternOrderkey ,cod.Sku ,cod.Descr ,cast(cod.OriginalQty as float) ,cast(cod.UnitNetPrice as float) , " & _
            "cast(cod.NetPrice as float) ,cod.GrossPrice,cast(cod.UnitGrossPrice as float) ,cod.BarCode ,cast(cod.Externlineno as float) ,cod.UOM ,cod.Ordertype ,cod.InvoicePCode ,cod.Acceptance order by co.orderkey"
            
End If
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic ' adOpenKeyset

If tmp_Rs.EOF = True Then
    tmp_Rs.Close: Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Exit Sub
Else
    Call Replication_Recordset(tmp_Rs, rsMainDetail)
    tmp_Rs.Close: Set tmp_Rs = Nothing
    
    Set dgMain_Detail.DataSource = rsMainDetail: dgMain_Detail.Visible = False
    rsMainDetail.MoveFirst
    
    With dgMain_Detail
    Set dgMain_Detail.DataSource = rsMainDetail
    
    '    .ColumnHeaders = True        '���D�����
    '    .RowHeight = 300
    '    .Columns(0).Alignment = dbgCenter
    '    .Columns(10).Alignment = dbgRight
    
    End With
    
    Call cb_all_Click
    cmdOTUpdate.Enabled = True
End If
'
'Dim str_Orderkey As String, str_externorderkey As String, Int_qty As Integer, bl_next As Boolean, Str_Sku As String
'str_externorderkey = "": str_Orderkey = "": Int_qty = 0: bl_next = True: Str_Sku = ""
'If optIn = True Then '�h�f
'    '��X�Ҧ����h���^�Ǫ��ꦬ���
'    Do While Not rsMainHeader.EOF
'        str_externorderkey = str_externorderkey & "'" & RTrim(rsMainHeader.Fields("�q�����")) & RTrim(rsMainHeader.Fields("�q�渹�X")) & "',"
'        rsMainHeader.MoveNext
'    Loop
'    rsMainHeader.MoveFirst
'    rsMainDetail.MoveFirst
'    str_externorderkey = Mid(str_externorderkey, 1, Len(str_externorderkey) - 1)
'    '��X���妸���ꦬ�q
'    Call Confirm_Recordset_Closed(rsMainReceitDetail)
'    rsMainReceitDetail.CursorLocation = adUseClient
'    str_SQL = "select r.externreceiptkey,rd.sku,rd.QtyReceived from " & strWMSDB & "..receipt r join " & strWMSDB & "..receiptdetail rd on r.receiptkey = rd.receiptkey where r.externreceiptkey in (" & str_externorderkey & ") order by r.externreceiptkey,rd.sku"
'    rsMainReceitDetail.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
'
'    '�H�n�^�Ǫ���Ƭ��D�A�ɶi�ꦬ�q
'    Do While Not rsMainDetail.EOF
'        '�z��f�D�渹���L�ꦬ�q
'        If bl_next = True Then
'            rsMainReceitDetail.Filter = "externreceiptkey = '" & RTrim(rsMainDetail.Fields("�q�����")) & RTrim(rsMainDetail.Fields("�q�渹�X")) & "'"
'            If rsMainReceitDetail.RecordCount = 0 Then rsMainDetail.Fields("�q�f�q-�ꦬ�q") = 0: GoTo next1         '�p�G���f�D�渹�S����ơA�h���U�@������
'
'            rsMainReceitDetail.Filter = "externreceiptkey = '" & RTrim(rsMainDetail.Fields("�q�����")) & RTrim(rsMainDetail.Fields("�q�渹�X")) & "' and sku = '" & RTrim(rsMainDetail.Fields("���~�s��")) & "'" '���h�D��S�w�~��
'            If rsMainReceitDetail.RecordCount = 0 Then rsMainDetail.Fields("�q�f�q-�ꦬ�q") = 0: GoTo next1         '�p�G���f�D�渹�S���~����ơA�h���U�@������
'        End If
'
'        If Str_Sku <> RTrim(rsMainDetail.Fields("���~�s��")) Then Str_Sku = RTrim(rsMainDetail.Fields("���~�s��")): Int_qty = Val(RTrim(rsMainReceitDetail.Fields("QtyReceived")))
'        If Int_qty >= Abs(Val(rsMainDetail.Fields("�q�f�q"))) Then
'            rsMainDetail.Fields("�q�f�q-�ꦬ�q") = Abs(Val(rsMainDetail.Fields("�q�f�q")))
'            Int_qty = Int_qty - Abs(Val(rsMainDetail.Fields("�q�f�q")))
'        Else
'            rsMainDetail.Fields("�q�f�q-�ꦬ�q") = Int_qty
'            Int_qty = 0
'        End If
'
'        If Int_qty = 0 Then bl_next = True Else bl_next = False
'next1:
'        rsMainDetail.MoveNext
'    Loop
'End If


rsMainHeader.MoveFirst
rsMainDetail.MoveFirst

SetDataGridColWidth Me.Caption, dgMain_Detail
StatusBar.Panels(2).Text = rsMainDetail.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgMain_Detail.Visible = True:: dgMain_Header.Visible = True
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
'Dim dg As Object: Set dg = dgMain
''�L��Ʃ���e�Ӥp�A���s�e��
'If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
'SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub



'Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'
'With dgMain
'
'If .DataSource Is Nothing Then Exit Sub
''If LastRow = Empty Then Exit Sub
'If .Row = -1 Or .Col <> 1 Then Exit Sub
'On Error GoTo err_Handle
'
'If .Col = 1 Then
'    If UCase(dgMain) <> "V" And Val(rsMain("�X�f���")) > 0 Then '������P��Ƥj��0
'        dgMain = "V"
'    Else
'        dgMain = " "
'
'    End If
'.Col = 0
'End If
'
'End With
'Exit Sub
'
'err_Handle:
'Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
'End Sub

Private Sub dgMain_Header_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

'�O�_�����
If rsMainHeader Is Nothing Then Exit Sub
If rsMainHeader.RecordCount = 0 Then Exit Sub

'If blRouteT0Change = False Then Exit Sub

'���
If dgMain_Header.Col = 1 Then

    If rsMainHeader("�O�_�^��") = " " Then
    
        rsMainHeader("�O�_�^��") = "V"
    Else
        rsMainHeader("�O�_�^��") = " "
    End If
    
    dgMain_Header.Col = 0

End If

''�ˬd�ƶq
'If rsMainHeader("�O�_�^��") = "V" Then
'            'rsMainDetail.Filter = "TMS�渹 = '" & rsMainHeader.Fields("TMS�渹") & "'"
'            rsMainDetail.MoveFirst
'
'            Do While Not rsMainDetail.EOF
'                If rsMainHeader.Fields("TMS�渹") = rsMainDetail.Fields("TMS�渹") Then
'                    If Abs(Val(rsMainDetail.Fields("�q�f�q"))) <> Val(rsMainDetail.Fields("�q�f�q-�ꦬ�q")) Then
'                        x = MsgBox("�q�渹�X:" & rsMainDetail.Fields("TMS�渹") & "����:" & rsMainDetail.Fields("TMS�渹����") & ":�q�f�q<>ñ���q�A�нT�{�O�_��s�^��?", vbQuestion + vbYesNo, "�ƶq�ˬd")
'                        If x = 6 Then
'                            '�O��
'                        Else
'                            rsMainHeader("�O�_�^��") = " "
'                            Screen.MousePointer = 0
'                            Exit Sub
'                        End If
'                    End If
'                End If
'                rsMainDetail.MoveNext
'            Loop
'End If
'rsMainDetail.MoveFirst

'�P�@����
If LastRow = Empty Then Exit Sub

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub


Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    Frame2.Height = Me.ScaleHeight - Frame1.Height - Frame1.Top - StatusBar.Height
    SSTab1.Height = Frame2.Height - 360
    dgMain_Header.Height = SSTab1.Height - 360
    dgMain_Detail.Height = SSTab1.Height - 360
    
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    Frame2.Width = Me.ScaleWidth
    SSTab1.Width = Frame2.Width - 240
    dgMain_Header.Width = SSTab1.Width - 240
    dgMain_Detail.Width = SSTab1.Width - 240
    
End If

End Sub

Private Sub cmdReset_Click()

'���]
Call ClearForm_AllField(Me)
optNo.Value = True
'optPrintNO.Value = True

End Sub

'Private Sub dgMain_HeadClick(ByVal ColIndex As Integer)
'
'If dgMain.Row = -1 Then Exit Sub
'If intColumnIndex = ColIndex Then
'    rsMain.Sort = dgMain.Columns(ColIndex).Caption & " DESC"
'    dgMain.ClearSelCols
'    intColumnIndex = 255
'
'Else
'    rsMain.Sort = dgMain.Columns(ColIndex).Caption
'    dgMain.ClearSelCols
'    intColumnIndex = ColIndex
'
'End If
'
'End Sub
Private Sub dgmain_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Call cmdOK_Click

End Sub

Private Sub cmdExit_Click()
Unload Me '�������{��
'End �������ε{��
End Sub

Private Sub Form_Load()
On Error GoTo err_Handle

optNo.Value = True

StatusBar.Panels(2).Text = "0 ����ƦC"
StatusBar.Panels(3).Text = User_id

Dim i As Integer


'�f�D
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(storerkey) from trp16M", cn, adOpenKeyset, adLockPessimistic
tmp_Rs.MoveFirst
For i = 0 To tmp_Rs.RecordCount - 1
    cboStorerkey.AddItem RTrim(tmp_Rs("storerkey"))
    tmp_Rs.MoveNext
Next
tmp_Rs.Close: Set tmp_Rs = Nothing
cboStorerkey.Text = "LMBO01"

txtDeliveryDateS = Format(Now - 2, "YYYYMMDD")

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMainHeader = Nothing
Set rsMainDetail = Nothing
End Sub



Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txtDeliveryDateS_Click()
Set objMvdateTarget = txtDeliveryDateS
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDeliveryDateE_Click()
Set objMvdateTarget = txtDeliveryDateE
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
'Private Sub txtOrderDateS_Click()
'Set objMvdateTarget = txtOrderDateS
'mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
'mvDate.Visible = True: mvDate.Value = Now
'
'End Sub
'Private Sub txtOrderDateE_Click()
'Set objMvdateTarget = txtOrderDateE
'mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
'mvDate.Visible = True: mvDate.Value = Now
'
'End Sub
'Private Sub txtOrderDateS_KeyPress(KeyAscii As Integer)
'If KeyAscii = 27 Then mvDate.Visible = False
'End Sub
'Private Sub txtOrderDateE_KeyPress(KeyAscii As Integer)
'If KeyAscii = 27 Then mvDate.Visible = False
'End Sub
Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateS_KeyPress(KeyAscii As Integer)
mvDate.Visible = False
End Sub
Private Sub txtDeliveryDateE_KeyPress(KeyAscii As Integer)
mvDate.Visible = False
End Sub
