VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Pallet_Match 
   Caption         =   "�̪O��ƽT�{"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9330
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   9330
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3960
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   -2147483624
      StartOfWeek     =   61341697
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin VB.Frame Frame3 
      Caption         =   "�\��"
      Height          =   5175
      Left            =   1560
      TabIndex        =   44
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
      Begin VB.CommandButton cmdPickSave 
         BackColor       =   &H00FFFF80&
         Caption         =   "�s��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   240
         Picture         =   "frm_Pallet_Match.frx":0000
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1065
      End
      Begin VB.CommandButton cmdPickDelete 
         BackColor       =   &H00FFC0FF&
         Caption         =   "�R��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   240
         Picture         =   "frm_Pallet_Match.frx":030A
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1065
      End
      Begin VB.CommandButton cmdPickEdit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�ק�"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   240
         Picture         =   "frm_Pallet_Match.frx":134C
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1065
      End
      Begin VB.CommandButton cmdPickAddNew 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�s�W"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   240
         Picture         =   "frm_Pallet_Match.frx":7B9E
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmdPickCancel 
         BackColor       =   &H00FF8080&
         Caption         =   "����"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   240
         Picture         =   "frm_Pallet_Match.frx":9CC8
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   45
         Top             =   4080
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "�g�P�Ӹ��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6255
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   4215
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
         Height          =   870
         Left            =   2760
         Picture         =   "frm_Pallet_Match.frx":1051A
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   50
         Top             =   3360
         Width           =   1065
      End
      Begin VB.CommandButton cmdAutoMatch 
         BackColor       =   &H00FFC0FF&
         Caption         =   "�۰ʤ��"
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   2760
         Picture         =   "frm_Pallet_Match.frx":3A12C
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   42
         Top             =   2400
         Width           =   1065
      End
      Begin VB.Frame Frame4 
         Height          =   615
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   3975
         Begin VB.OptionButton optAll 
            Caption         =   "����"
            Height          =   255
            Left            =   2280
            TabIndex        =   2
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optNo 
            Caption         =   "���T�{"
            Height          =   255
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optYes 
            Caption         =   "�w�T�{"
            Height          =   255
            Left            =   1200
            TabIndex        =   1
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�T�{�s��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   360
         Picture         =   "frm_Pallet_Match.frx":51E56
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   10
         Top             =   2400
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�����T�{"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   1560
         Picture         =   "frm_Pallet_Match.frx":53B50
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   11
         Top             =   2400
         Width           =   1065
      End
      Begin VB.TextBox txtKeyid 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2000
         Width           =   2085
      End
      Begin VB.ComboBox cboFloatCustomer 
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   31
         Text            =   "cboFloatCustomer"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00C0FFC0&
         Caption         =   "���]"
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   1560
         Picture         =   "frm_Pallet_Match.frx":6ABDA
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   9
         Top             =   3360
         Width           =   1065
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�d��"
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   360
         Picture         =   "frm_Pallet_Match.frx":6AEEC
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   8
         Top             =   3360
         Width           =   1065
      End
      Begin VB.TextBox txtE 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2820
         MaxLength       =   8
         TabIndex        =   4
         Top             =   900
         Width           =   1245
      End
      Begin VB.TextBox txtS 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
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
         TabIndex        =   3
         Top             =   900
         Width           =   1245
      End
      Begin VB.ComboBox cboCarno 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         ItemData        =   "frm_Pallet_Match.frx":6B1F6
         Left            =   1200
         List            =   "frm_Pallet_Match.frx":6B1F8
         TabIndex        =   6
         Top             =   1620
         Width           =   2085
      End
      Begin VB.ComboBox cboCustomer 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         ItemData        =   "frm_Pallet_Match.frx":6B1FA
         Left            =   1200
         List            =   "frm_Pallet_Match.frx":6B1FC
         TabIndex        =   5
         Top             =   1260
         Width           =   2085
      End
      Begin MSDataGridLib.DataGrid dgUtlcst 
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   4320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1508
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   20
         TabAction       =   1
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�ө���"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "KeyID"
         BeginProperty Font 
            Name            =   "�ө���"
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
         TabIndex        =   37
         Top             =   2085
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "ñ�����"
         BeginProperty Font 
            Name            =   "�ө���"
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
         TabIndex        =   28
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�ө���"
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
         Left            =   2460
         TabIndex        =   27
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�Ȥ�W��"
         BeginProperty Font 
            Name            =   "�ө���"
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
         TabIndex        =   25
         Top             =   1320
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�J�X�w���"
      Height          =   6255
      Left            =   4440
      TabIndex        =   29
      Top             =   0
      Width           =   4215
      Begin VB.CheckBox chkCustomer 
         Caption         =   "�ȯS��Ȥ�"
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   43
         ToolTipText     =   "�̪��޲z�t�Φ����ɪ��Ȥ�"
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox txtCheckNo 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   40
         Top             =   2400
         Width           =   2085
      End
      Begin VB.TextBox txtKeyidcst 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   20
         Top             =   2000
         Width           =   2085
      End
      Begin VB.Frame Frame5 
         Height          =   615
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   3975
         Begin VB.OptionButton optYescst 
            Caption         =   "�w�T�{"
            Height          =   255
            Left            =   1200
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optNocst 
            Caption         =   "���T�{"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optAllcst 
            Caption         =   "����"
            Height          =   255
            Left            =   2280
            TabIndex        =   15
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.ComboBox cboCustomercst 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         ItemData        =   "frm_Pallet_Match.frx":6B1FE
         Left            =   1200
         List            =   "frm_Pallet_Match.frx":6B200
         TabIndex        =   18
         Top             =   1260
         Width           =   2085
      End
      Begin VB.ComboBox cboCarnocst 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         ItemData        =   "frm_Pallet_Match.frx":6B202
         Left            =   1200
         List            =   "frm_Pallet_Match.frx":6B204
         TabIndex        =   19
         Top             =   1620
         Width           =   2085
      End
      Begin VB.TextBox txtScst 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
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
         TabIndex        =   16
         Top             =   900
         Width           =   1245
      End
      Begin VB.TextBox txtEcst 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2820
         MaxLength       =   8
         TabIndex        =   17
         Top             =   900
         Width           =   1245
      End
      Begin VB.CommandButton cmdQuerycst 
         BackColor       =   &H00C0FFFF&
         Caption         =   "�d��"
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   360
         Picture         =   "frm_Pallet_Match.frx":6B206
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   21
         Top             =   3360
         Width           =   1065
      End
      Begin VB.CommandButton cmdResetcst 
         BackColor       =   &H00C0FFC0&
         Caption         =   "���]"
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   1560
         Picture         =   "frm_Pallet_Match.frx":6B510
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   22
         Top             =   3360
         Width           =   1065
      End
      Begin MSDataGridLib.DataGrid dgCst 
         Height          =   855
         Left            =   120
         TabIndex        =   23
         Top             =   4320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1508
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   20
         TabAction       =   1
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�ө���"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "CDS�渹"
         BeginProperty Font 
            Name            =   "�ө���"
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
         TabIndex        =   41
         Top             =   2490
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "KeyID"
         BeginProperty Font 
            Name            =   "�ө���"
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
         Left            =   120
         TabIndex        =   39
         Top             =   2085
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�Ȥ�W��"
         BeginProperty Font 
            Name            =   "�ө���"
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
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�ө���"
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
         TabIndex        =   34
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   1
         Left            =   2460
         TabIndex        =   33
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "ñ�����"
         BeginProperty Font 
            Name            =   "�ө���"
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
         TabIndex        =   32
         Top             =   945
         Width           =   960
      End
   End
End
Attribute VB_Name = "frm_Pallet_Match"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsUtlcst As ADODB.Recordset
Private rsCst As ADODB.Recordset
Private objMvdateTarget As Object
Private intPickRow As Integer, intLastCol As Integer, intOrderRow As Integer, intSkuRow As Integer, intPickqty As Integer

Private Sub cboFloatCustomer_LostFocus()
cboFloatCustomer.Visible = False
End Sub

Private Sub cmdCancel_Click()
On Error GoTo err_Handle

Dim confirm As Integer
confirm = MsgBox("�T�w����?", vbQuestion + vbOKCancel, Me.Caption)
If confirm <> 1 Then Exit Sub

Dim strTmp As String
strTmp = rsUtlcst("keyid")

'��s��Ʈw
str_SQL = "update pallet_utlcst set checkuser = '',checkdate = null where keyid = '" & rsUtlcst("keyid") & "' " & _
          "update pallet_cst set checkuser = '',checkdate = null , keyid = '' where keyid = '" & rsUtlcst("keyid") & "' "
cn.BeginTrans
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
cn.CommitTrans

'��srsutlcst
rsUtlcst("�w�T�{") = " "
rsUtlcst("checkuser") = ""
rsUtlcst("checkdate") = ""
rsUtlcst.Update

'��srscst
If dgCst.DataSource Is Nothing = False Or rsCst Is Nothing = False Then
    rsCst.Filter = adFilterNone
    rsCst.Filter = "keyid = " & strTmp
    rsCst.MoveFirst
       Do While Not rsCst.EOF
            rsCst("�w�T�{") = " "
            rsCst("keyid") = ""
            rsCst("checkuser") = ""
            rsCst("checkdate") = ""
            rsCst.Update
            rsCst.MoveNext
        Loop

'    rsCst.Filter = adFilterNone
End If

cmdOK.Enabled = True: cmdCancel.Enabled = False
Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub cmdOK_Click()
On Error GoTo err_Handle

If dgCst.DataSource Is Nothing Then MsgBox "�п�����T�{���J�X�w��ơI", vbOKOnly, Me.Caption: Exit Sub
If rsCst("�w�T�{") = "��" Then MsgBox "�J�X�w��Ƥw�T�{�L�A���ˬd��A�T�{�I", vbOKOnly, Me.Caption: Exit Sub
If rsUtlcst("�ɤJ") <> rsCst("�ɤJ") Or rsUtlcst("�٦^") <> rsCst("�٦^") Then MsgBox "�ƶq���šA���ˬd�ƶq��A�T�{�I", vbOKOnly, Me.Caption: Exit Sub

'����Ȥᨮ������ˬd
If rsUtlcst("ñ�����") <> rsCst("ñ�����") Or rsUtlcst("�Ȥ�W��") <> rsCst("�Ȥ�W��") Or rsUtlcst("����") <> rsCst("����") Then
Dim confirm As Integer
confirm = MsgBox("����B�Ȥ�Ψ�����Ƥ��šA�T�w�s��?", vbQuestion + vbOKCancel, Me.Caption)
If confirm <> 1 Then Exit Sub
End If

'��s��Ʈw
Dim strNow As String: strNow = Format(Now(), "yyyy/mm/dd hh:MM:ss")
str_SQL = "update pallet_utlcst set checkuser = '" & User_id & "',checkdate = '" & strNow & "' where keyid = '" & rsUtlcst("keyid") & "' " & _
          "update pallet_cst set checkuser = '" & User_id & "',checkdate = '" & strNow & "' , keyid = '" & rsUtlcst("keyid") & "' where checkno = '" & RTrim(rsCst("�渹")) & "' and linenumber = '" & rsCst("����") & "' "
cn.BeginTrans
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
cn.CommitTrans

'��sdgutlcst
rsUtlcst("�w�T�{") = "��": rsUtlcst("checkuser") = User_id: rsUtlcst("checkdate") = strNow
rsUtlcst.Update
'��sdgcst
rsCst("�w�T�{") = "��": rsCst("checkuser") = User_id: rsCst("checkdate") = strNow: rsCst("Keyid") = rsUtlcst("Keyid")
rsCst.Update

cmdOK.Enabled = False: cmdCancel.Enabled = True

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub cmdPickAddNew_Click()
Dim i As Integer

With rsUtlcst
    i = 1
    If .RecordCount > 0 Then .MoveLast: i = .Fields("�s��") + 1
    .AddNew
    .Fields("�s��") = i
    .Fields("ñ�����") = Format(Now, "yyyymmdd")
    .Fields("�Ȥ�W��") = ""
    .Fields("�渹") = ""
    .Fields("����") = ""
    .Fields("�ɤJ") = "0"
    .Fields("�٦^") = "0"
End With

dgUtlcst.AllowUpdate = True
cmdPickSave.Enabled = True: cmdPickCancel.Enabled = True
cmdPickDelete.Enabled = False: cmdPickEdit.Enabled = False: cmdPickAddNew.Enabled = False
dgUtlcst.Col = 1: dgUtlcst.SetFocus
intPickRow = dgUtlcst.Row
intLastCol = dgUtlcst.Col

End Sub
Private Sub cmdPickEdit_Click()

If Len(rsUtlcst("checkuser")) > 0 Then MsgBox "�w�T�{��ƵL�k�ק�!!", vbInformation: Exit Sub

dgUtlcst.AllowUpdate = True
cmdPickSave.Enabled = True: cmdPickCancel.Enabled = True
cmdPickDelete.Enabled = False: cmdPickEdit.Enabled = False: cmdPickAddNew.Enabled = False
dgUtlcst.Col = 1: dgUtlcst.SetFocus
intPickRow = dgUtlcst.Row
intLastCol = dgUtlcst.Col

End Sub
Private Sub cmdPickDelete_Click()
On Error GoTo err_Handle
Dim confirm As Integer

If Len(rsUtlcst("checkuser")) > 0 Then MsgBox "�w�T�{��ƵL�k�R��!!", vbInformation: Exit Sub
confirm = MsgBox("�T�w�R��?", vbQuestion + vbOKCancel, Me.Caption)
If confirm <> 1 Then Exit Sub

str_SQL = "delete from pallet_utlcst where keyid = '" & rsUtlcst("keyid") & "' "
cn.BeginTrans
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
cn.CommitTrans

'��sdgUtlcst���
rsUtlcst.Delete: If rsUtlcst.EOF Then rsUtlcst.MovePrevious
cmdPickAddNew.SetFocus

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdPickSave_Click()
On Error GoTo err_Handle

If Len(RTrim(rsUtlcst("�Ȥ�W��") & "")) = 0 Then MsgBox "�п�J�Ȥ�W��!!", vbOKOnly + vbInformation, Me.Caption: dgUtlcst.Col = 2: dgUtlcst.SetFocus: Exit Sub
If Len(RTrim(rsUtlcst("����") & "")) = 0 Then MsgBox "�п�J����!!", vbOKOnly + vbInformation, Me.Caption: dgUtlcst.Col = 4: dgUtlcst.SetFocus: Exit Sub
If Val(Trim(rsUtlcst("�ɤJ"))) + Val(Trim(rsUtlcst("�٦^"))) = 0 Then MsgBox "�нT�{�ƶq!!", vbOKOnly + vbInformation, Me.Caption: dgUtlcst.Col = 5: dgUtlcst.SetFocus: Exit Sub

'�ˬd�O�_����
Dim rsTmp1 As New ADODB.Recordset
With rsTmp1
    .CursorLocation = adUseClient
    str_SQL = "select * from pallet_utlcst where keyid = '" & rsUtlcst("keyid") & "' "
    .Open str_SQL, cn, adOpenStatic, adLockOptimistic
        
    If .EOF Then
    
    Dim rsTmp As New ADODB.Recordset, keyid As String
    rsTmp.Open "select keyid = isnull(max(keyid),0) from pallet_utlcst", cn
    keyid = Format(Val(rsTmp("keyid")) + 1, "0000000000")
    
'        '�s�W��Ʈw���
            .AddNew
            .Fields("keyid") = keyid
            .Fields("Storer") = "UTL"
            .Fields("chargedate") = rsUtlcst("ñ�����")
            .Fields("customer") = rsUtlcst("�Ȥ�W��")
            .Fields("customersheetno") = rsUtlcst("�渹")
            .Fields("carno") = UCase(rsUtlcst("����"))
            .Fields("qtyin") = rsUtlcst("�ɤJ")
            .Fields("qtyout") = rsUtlcst("�٦^")
            .Fields("notes") = rsUtlcst.Fields("�Ƶ�")
            .Fields("Adduser") = User_id
            .Fields("Adddate") = Now()
            .Update
            
            '��sdgUtlcst
            rsUtlcst.Fields("keyid") = keyid
            rsUtlcst.Fields("Adduser") = User_id
            rsUtlcst.Fields("Adddate") = str(Now())
            rsUtlcst.Update
            
rsTmp.Close: Set rsTmp = Nothing
    Else

        '�ק���
            .Fields("chargedate") = rsUtlcst("ñ�����")
            .Fields("customer") = rsUtlcst("�Ȥ�W��")
            .Fields("customersheetno") = rsUtlcst("�渹")
            .Fields("carno") = UCase(rsUtlcst("����"))
            .Fields("qtyin") = rsUtlcst("�ɤJ")
            .Fields("qtyout") = rsUtlcst("�٦^")
            .Fields("notes") = rsUtlcst.Fields("�Ƶ�")
            .Fields("Edituser") = User_id
            .Fields("Editdate") = Now()
            .Update
            
            '��sdgUtlcst
            rsUtlcst.Fields("Edituser") = User_id
            rsUtlcst.Fields("Editdate") = str(Now())
            rsUtlcst.Update
'
    End If
End With

cmdPickAddNew.Enabled = True: cmdPickEdit.Enabled = True: cmdPickDelete.Enabled = True: dgUtlcst.AllowUpdate = False: cmdPickSave.Enabled = False: cmdPickCancel.Enabled = False
'Call Update
cmdPickAddNew.SetFocus
dgUtlcst.AllowUpdate = False

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdPickCancel_Click()

cmdPickSave.Enabled = False: cmdPickCancel.Enabled = False
cmdPickAddNew.Enabled = True
If rsUtlcst.RecordCount > 0 Then cmdPickEdit.Enabled = True: cmdPickDelete.Enabled = True
cmdPickAddNew.SetFocus
dgUtlcst.AllowUpdate = False

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
'Set dgUtlcst.DataSource = Nothing
Dim chc_Chargedate As String, chc_Carno As String, chc_Customer As String, chc_Check As String, chc_Keyid

'���X�Ȥ�̪O���
str_SQL = "select �w�T�{ = case when len(rtrim(isnull(checkuser,''))) > 0 then '��' else ' ' end " & _
          ", ñ����� = chargedate  " & _
          ", �Ȥ�W�� = customer " & _
          ", ���� = rtrim(carno) " & _
          ", �ɤJ= qtyin " & _
          ", �٦^ = qtyout " & _
          ", �Ƶ� = rtrim(notes) " & _
          ", �渹 = rtrim(customersheetno) " & _
          ", AddUser = rtrim(adduser) " & _
          ", Adddate = rtrim(convert( char(20) , adddate , 120 )) " & _
          ", CheckUser = rtrim(CheckUser) " & _
          ", Checkdate = rtrim(convert( char(20) , Checkdate , 120 )) " & _
          ", EditUser = rtrim(EditUser) " & _
          ", Editdate = rtrim(convert( char(20) , Editdate , 120 )) " & _
          ", KeyID " & _
          "from pallet_UTLcst "

'�Ȥ�W��
chc_Customer = ""
If Len(cboCustomer.Text) > 0 Then chc_Customer = "and Customer like '" & cboCustomer.Text & "%' "

'����
chc_Carno = ""
If Len(cboCarno.Text) > 0 Then chc_Carno = "and carno = '" & cboCarno.Text & "' "

'�ƥX���
chc_Chargedate = ""
If Len(txtS.Text) > 0 And Len(txtE.Text) > 0 Then
   chc_Chargedate = "and Chargedate between '" & txtS.Text & "' and '" & txtE.Text & "' "
ElseIf Len(txtS.Text) > 0 And Len(txtE.Text) = 0 Then
   chc_Chargedate = "and Chargedate = '" & txtS.Text & "' "
ElseIf Len(txtS.Text) = 0 And Len(txtE.Text) > 0 Then
   chc_Chargedate = "and Chargedate = '" & txtE.Text & "' "
End If

'�w�T�{
chc_Check = ""
If optNo = True Then chc_Check = "and len(rtrim(isnull(checkuser,''))) = 0 "
If optYes = True Then chc_Check = "and len(rtrim(isnull(checkuser,''))) > 0 "

'KeyID
chc_Keyid = ""
If Len(txtKeyid.Text) > 0 Then chc_Keyid = "and keyid = '" & txtKeyid.Text & "' "

'�զX�r��
str_SQL = str_SQL & "where 1 = 1 " & chc_Chargedate & chc_Carno & chc_Customer & chc_Check & chc_Keyid

Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.CursorLocation = adUseClient
tmp_rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_rs.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Set dgUtlcst.DataSource = Nothing: cmdOK.Enabled = False: Exit Sub
tmp_rs.Sort = "ñ�����,����"
Call Replication_Recordset(tmp_rs, rsUtlcst)
tmp_rs.Close: Set tmp_rs = Nothing

rsUtlcst.MoveFirst
Set dgUtlcst.DataSource = rsUtlcst

With dgUtlcst
Set dgUtlcst.DataSource = rsUtlcst
    .ColumnHeaders = True        '���D�����
    .RowHeight = 300
    .Columns(0).Width = 600:       .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 800:       .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1000:       .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1500:    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000:    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 600:    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 600:    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 2000:    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 1100:    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 1000:    .Columns(9).Alignment = dbgLeft
    .Columns(10).Width = 1500:    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 1000:    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1500:   .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 1000:    .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 1500:    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 1200:    .Columns(15).Alignment = dbgLeft
End With
cmdPickEdit.Enabled = True: cmdPickDelete.Enabled = True
Call dgUtlcst_RowColChange(Empty, Empty)
Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
'
'Private Sub dgUtlcst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim confirm As Integer
'
'If cmdPickSave.Enabled = True And dgUtlcst.ColContaining(X) = -1 And dgUtlcst.RowContaining(Y) <> intPickRow Then
'confirm = MsgBox("�O�_�s��!!", vbQuestion + vbOKCancel)
'If confirm = 1 Then cmdPickSave_Click
'cmdPickCancel_Click
'
'End If
'End Sub

Private Sub cmdQuerycst_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
'Set dgUtlcst.DataSource = Nothing
Dim chc_Chargedate As String, chc_Carno As String, chc_Customer As String, chc_Check As String, chc_Keyid As String, chc_CheckNo As String

'���X�Ȥ�̪O���
str_SQL = "select �w�T�{ = case when len(rtrim(isnull(checkuser,''))) > 0 then '��' else ' ' end " & _
          ", ñ����� = chargedate  " & _
          ", �Ȥ�W�� = rtrim(customer) " & _
          ", ���� = rtrim(carno) " & _
          ", �ɤJ = qtyin " & _
          ", �٦^ = qtyout " & _
          ", �Ƶ� = rtrim(notes) " & _
          ", �渹 = rtrim(checkno) " & _
          ", ���� = lineNumber " & _
          ", AddUser = rtrim(adduser) " & _
          ", Adddate = rtrim(convert( char(20) , adddate , 120 )) " & _
          ", CheckUser = rtrim(CheckUser) " & _
          ", Checkdate = rtrim(convert( char(20) , Checkdate , 120 )) " & _
          ", EditUser = rtrim(EditUser) " & _
          ", Editdate = rtrim(convert( char(20) , Editdate , 120 )) " & _
          ", KeyID = isnull(KeyID,'') " & _
          "from pallet_cst "

'�Ȥ�W��
chc_Customer = ""
If Len(cboCustomercst.Text) > 0 Then chc_Customer = "and Customer like '" & cboCustomercst.Text & "%' "

'�ȯS��Ȥ�
If chkCustomer.Value = 1 Then chc_Customer = chc_Customer + " and Customer in (select code from CodeLkup where listname= 'Cust_CDS') "

'����
chc_Carno = ""
If Len(cboCarnocst.Text) > 0 Then chc_Carno = "and carno = '" & cboCarnocst.Text & "' "

'�ƥX���
chc_Chargedate = ""
If Len(txtScst.Text) > 0 And Len(txtEcst.Text) > 0 Then
   chc_Chargedate = "and Chargedate between '" & txtScst.Text & "' and '" & txtEcst.Text & "' "
ElseIf Len(txtScst.Text) > 0 And Len(txtEcst.Text) = 0 Then
   chc_Chargedate = "and Chargedate = '" & txtScst.Text & "' "
ElseIf Len(txtScst.Text) = 0 And Len(txtEcst.Text) > 0 Then
   chc_Chargedate = "and Chargedate = '" & txtEcst.Text & "' "
End If

'�w�T�{
chc_Check = ""
If optNocst = True Then chc_Check = "and len(rtrim(isnull(checkuser,''))) = 0 "
If optYescst = True Then chc_Check = "and len(rtrim(isnull(checkuser,''))) > 0 "

'KeyID
chc_Keyid = ""
If Len(txtKeyidcst.Text) > 0 Then chc_Keyid = "and keyid = '" & txtKeyidcst.Text & "' "

'�渹
chc_CheckNo = ""
If Len(txtCheckNo.Text) > 0 Then chc_CheckNo = "and CheckNo = '" & txtCheckNo.Text & "' "

'�զX�r��
str_SQL = str_SQL & "where 1 = 1 " & chc_Chargedate & chc_Carno & chc_Customer & chc_Check & chc_Keyid & chc_CheckNo

Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.CursorLocation = adUseClient
tmp_rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_rs.EOF = True Then Screen.MousePointer = 0: MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Set dgCst.DataSource = Nothing: Call dgUtlcst_RowColChange(Empty, Empty): Exit Sub
tmp_rs.Sort = "ñ�����,����"
Call Replication_Recordset(tmp_rs, rsCst)
tmp_rs.Close: Set tmp_rs = Nothing

rsCst.MoveFirst
Set dgCst.DataSource = rsCst

With dgCst
Set dgCst.DataSource = rsCst
    .ColumnHeaders = True        '���D�����
    .RowHeight = 300
    .Columns(0).Width = 600:       .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 800:       .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1000:       .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1500:    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000:    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 600:    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 600:    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 2000:    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 1100:    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 600:    .Columns(9).Alignment = dbgCenter
    .Columns(10).Width = 1000:    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 1500:    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1000:    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 1500:   .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 1000:    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 1500:    .Columns(15).Alignment = dbgLeft
    .Columns(16).Width = 1200:    .Columns(16).Alignment = dbgLeft
End With
Call dgUtlcst_RowColChange(Empty, Empty)
Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub cmdResetcst_Click()
'���]
txtScst.Text = "": txtEcst.Text = ""
cboCustomercst.Text = "": cboCarnocst.Text = ""
optNocst = True
txtKeyidcst.Text = ""
txtCheckNo.Text = ""
chkCustomer.Value = 1
Set dgCst.DataSource = Nothing
End Sub

Private Sub cmdAutoMatch_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
'
''���X�Ȥ�̪O���
'str_SQL = "select distinct cst.checkno , cst.linenumber , cst.chargedate , cst.customer , cst.carno , cst.qtyin , cst.qtyout , Keyid = (select top 1 utl.keyid from pallet_utlcst utl where len(rtrim(isnull(utl.checkuser,'')))= 0 and cst.customer = utl.customer and cst.chargedate = utl.chargedate  and cst.qtyin = utl.qtyin and cst.qtyout = utl.qtyout order by keyid ) " & _
'            "from pallet_cst cst join pallet_utlcst utl on cst.carno = utl.carno and cst.customer = utl.customer and cst.chargedate = utl.chargedate  and cst.qtyin = utl.qtyin and cst.qtyout = utl.qtyout " & _
'            "where len(rtrim(isnull(cst.keyid,'')))= 0 " & _
'            "and len(rtrim(isnull(cst.checkuser,'')))= 0 " & _
'            "and len(rtrim(isnull(utl.checkuser,'')))= 0 "

Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.Open " exec gs_palletmatch '" & User_id & "'", cn
'If tmp_rs.EOF = True Then: Screen.MousePointer = 0: Exit Sub
'
'tmp_rs.MoveFirst
'Dim i As Integer
'Do While Not tmp_rs.EOF
'
'    str_SQL = "update pallet_utlcst set checkuser = '" & user_id & "',checkdate = getdate() where keyid = '" & tmp_rs("keyid") & "' " & _
'              "update pallet_cst set checkuser = '" & user_id & "',checkdate = getdate() , keyid = '" & tmp_rs("keyid") & "' where checkno = '" & RTrim(tmp_rs("checkno")) & "' and linenumber = '" & tmp_rs("linenumber") & "' "
'    cn.BeginTrans
'    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'    cn.CommitTrans
'
'    i = i + 1
'
'    tmp_rs.MoveNext
'
'Loop
MsgBox "�@���� " & tmp_rs("Matchcount") & " ����ƽT�{!!", vbOKOnly, "�۰ʤ��"

tmp_rs.Close: Set tmp_rs = Nothing

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub


Private Sub dgCst_DblClick()
If rsCst Is Nothing Then Exit Sub
If rsCst("�w�T�{") = "��" Then
    txtS.Text = "": txtE.Text = "": cboCustomer.Text = "": cboCarno.Text = "": optYes = True: txtKeyid.Text = rsCst("keyid")
Else
    txtScst.Text = "": txtEcst.Text = "": cboCarno.Text = rsCst("����"): cboCustomer.Text = rsCst("�Ȥ�W��"): optNo = True: txtKeyid.Text = "" ': cmdOK.Enabled = True: cmdCancel.Enabled = False
    '�Ȥ�W�٧t"-"
    If InStr(rsCst("�Ȥ�W��"), "-") > 0 Then cboCustomer.Text = Left(rsCst("�Ȥ�W��"), InStr(rsCst("�Ȥ�W��"), "-") - 1)

End If
End Sub

Private Sub dgCst_HeadClick(ByVal ColIndex As Integer)

With dgCst

    If .Row = -1 Then Exit Sub
    If intColIndex = ColIndex Then
        rsCst.Sort = .Columns(ColIndex).Caption & " DESC"
        .ClearSelCols
        intColIndex = 255
    
    Else
        rsCst.Sort = .Columns(ColIndex).Caption
        .ClearSelCols
        intColIndex = ColIndex
    
    End If

End With

End Sub

Private Sub dgCst_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

cboFloatCustomer.Visible = False
mvDate.Visible = False

If dgUtlcst.DataSource Is Nothing Then cmdOK.Enabled = False: cmdCancel.Enabled = False: Exit Sub

If rsUtlcst("�w�T�{") = "��" Then
    cmdOK.Enabled = False: cmdCancel.Enabled = True ': txtKeyidcst.Text = rsUtlcst("keyid"): optYescst = True
Else
    If dgCst.DataSource Is Nothing Then
    cmdOK.Enabled = False: cmdCancel.Enabled = False
    Else
    cmdOK.Enabled = True: cmdCancel.Enabled = False
    If rsCst.EOF = False Then If rsCst("�w�T�{") = "��" Then cmdOK.Enabled = False: cmdCancel.Enabled = False
    End If
End If


''�s�W���A�U�L�k�ܧ��ƦC
'If cmdPickSave.Enabled = True And LastRow <> Empty Then
'    dgUtlcst.Col = intLastCol
'    dgUtlcst.Row = intPickRow
'    Exit Sub
'End If

'    If dgUtlcst.Col = 2 And cmdPickSave.Enabled = True Then ShowList
'    If dgUtlcst.Col = 1 And cmdPickSave.Enabled = True Then
'    Set objMvdateTarget = dgUtlcst: mvDate.Visible = True: mvDate.Value = Now()
'    mvDate.Move dgUtlcst.Columns(dgUtlcst.Col).Left + dgUtlcst.Columns(dgUtlcst.Col).Width + dgUtlcst.Left + Frame2.Left, dgUtlcst.RowTop(dgUtlcst.Row) + dgUtlcst.Top + Frame2.Top
'    End If

'�����\���ܯS�w���
'If dgUtlcst.Col = 0 Or dgUtlcst.Col > 1 Then dgUtlcst.Col = Abs(LastCol): Exit Sub
'If dgUtlcst.Col = 4 Then
'    If LastCol = 3 Then dgUtlcst.Col = 5: Exit Sub
'    If LastCol = 5 Then dgUtlcst.Col = 2: Exit Sub
'    dgUtlcst.Col = IIf(LastCol = -1, 5, LastCol)
'End If
'��ƦC�O�_�ܧ�
If LastRow = Empty Then Exit Sub

Screen.MousePointer = 0

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub dgUtlcst_DblClick()
If rsUtlcst Is Nothing Then Exit Sub
    If rsUtlcst("�w�T�{") = "��" Then
        txtScst.Text = "": txtEcst.Text = "": cboCustomercst.Text = "": cboCarnocst.Text = "": optYescst = True: txtKeyidcst.Text = rsUtlcst("keyid"): txtCheckNo.Text = ""
    Else
        txtScst.Text = "": txtEcst.Text = "": cboCarnocst.Text = rsUtlcst("����"): cboCustomercst.Text = rsUtlcst("�Ȥ�W��"): optNocst = True: txtKeyid.Text = "" ': cmdOK.Enabled = True: cmdCancel.Enabled = False
        '�Ȥ�W�٧t"-"
        If InStr(rsUtlcst("�Ȥ�W��"), "-") > 0 Then cboCustomercst.Text = Left(rsUtlcst("�Ȥ�W��"), InStr(rsUtlcst("�Ȥ�W��"), "-") - 1)
    End If
End Sub

Private Sub dgUtlcst_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

If dgUtlcst.DataSource Is Nothing Then cmdOK.Enabled = False: cmdCancel.Enabled = False: Exit Sub

cboFloatCustomer.Visible = False
mvDate.Visible = False

'If dgUtlcst.Col = 1 Then

If rsUtlcst("�w�T�{") = "��" Then
    cmdOK.Enabled = False: cmdCancel.Enabled = True ': txtKeyidcst.Text = rsUtlcst("keyid"): optYescst = True
Else
    If dgCst.DataSource Is Nothing Then
    cmdOK.Enabled = False: cmdCancel.Enabled = False
    Else
    cmdOK.Enabled = True: cmdCancel.Enabled = False
    If rsCst("�w�T�{") = "��" Then cmdOK.Enabled = False: cmdCancel.Enabled = False
    End If
End If

'��ƦC�O�_�ܧ�
If LastRow = Empty Then Exit Sub

Screen.MousePointer = 0

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��

If (Me.ScaleWidth - 360) / 2 > txtE.Left + txtE.Width + 120 Then
    Frame1.Width = (Me.ScaleWidth - 360) / 2: Frame2.Width = Frame1.Width: Frame2.Left = Frame1.Left + Frame1.Width + 120
    dgUtlcst.Width = Frame1.Width - 240: dgCst.Width = dgUtlcst.Width
End If

If Me.ScaleHeight > dgUtlcst.Top Then
    Frame1.Height = Me.ScaleHeight - Frame1.Top - 120: Frame2.Height = Frame1.Height
    dgUtlcst.Height = Frame1.Height - dgUtlcst.Top - 120: dgCst.Height = dgUtlcst.Height

End If

End Sub

Private Sub cmdReset_Click()

'���]
txtS.Text = "": txtE.Text = ""
cboCustomer.Text = "": cboCarno.Text = ""
optNo = True
txtKeyid.Text = ""
Set dgUtlcst.DataSource = Nothing

End Sub

Private Sub dgUtlcst_HeadClick(ByVal ColIndex As Integer)

With dgUtlcst

    If .Row = -1 Then Exit Sub
    If intColIndex = ColIndex Then
        rsUtlcst.Sort = .Columns(ColIndex).Caption & " DESC"
        .ClearSelCols
        intColIndex = 255
    
    Else
        rsUtlcst.Sort = .Columns(ColIndex).Caption
        .ClearSelCols
        intColIndex = ColIndex
    
    End If

End With

End Sub
Private Sub dgUtlcst_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{tab}"

End Sub
Private Sub cboFloatCustomer_Click()

dgUtlcst.Text = cboFloatCustomer.Text

End Sub
Private Sub ShowList()

With dgUtlcst
.RowHeight = cboFloatCustomer.Height - 10
If .Col = 2 Then
    If .Columns(.Col).Left > 0 Then
            cboFloatCustomer.Visible = True
            cboFloatCustomer.Move .Left + .Columns(.Col).Left + 15, .Top + .RowTop(.Row), .Columns(.Col).Width
            If cboFloatCustomer.Left + cboFloatCustomer.Width > .Left + .Width Then '�p�G���W�XDataGrid����ܽd�򪺳B�z
                cboFloatCustomer.Width = cboFloatCustomer.Width + .Left + .Width - cboFloatCustomer.Left - cboFloatCustomer.Width
            End If
            cboFloatCustomer.Text = dgUtlcst.Text  '��sCombo����
    Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
        cboFloatCustomer.Visible = False
    End If
Else
    cboFloatCustomer.Visible = False
End If
End With
End Sub
Private Sub dgUtlcst_Scroll(Cancel As Integer)
ShowList
End Sub
Private Sub dgUtlcst_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
ShowList
End Sub
Private Sub dgUtlcst_RowResize(Cancel As Integer)
ShowList
End Sub
Private Sub cmdExit_Click()
Unload Me '�������{��
'End �������ε{��
End Sub

Private Sub Form_Load()
On Error GoTo err_Handle

'���X�Ȥ�W��
Call Confirm_Recordset_Closed(tmp_rs)
str_SQL = "select code from CodeLkup where listname='Cust_CDS'"

tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_rs.EOF Then
   Do While Not tmp_rs.EOF
      cboCustomer.AddItem Trim(tmp_rs.Fields("code"))
      cboCustomercst.AddItem Trim(tmp_rs.Fields("code"))
      cboFloatCustomer.AddItem Trim(tmp_rs.Fields("code"))
      tmp_rs.MoveNext
   Loop
End If
tmp_rs.Close

'���X�g�P�Ө���
cboCarno.Clear
str_SQL = "select distinct Carno = rtrim(carno) From pallet_utlcst"
Dim rsTmp As New ADODB.Recordset
rsTmp.CursorLocation = 3
rsTmp.Open str_SQL, cn ', adOpenForwardOnly, adLockPessimistic
rsTmp.Sort = "Carno"
If Not rsTmp.EOF Then
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If IsNull(rsTmp("carno")) = False Then cboCarno.AddItem rsTmp("carno")
            rsTmp.MoveNext
        Loop
End If
rsTmp.Close

'���X�J�X�b����
cboCarnocst.Clear
str_SQL = "select distinct Carno = rtrim(carno) From pallet_cds"
rsTmp.CursorLocation = 3
rsTmp.Open str_SQL, cn ', adOpenForwardOnly, adLockPessimistic
rsTmp.Sort = "Carno"
If Not rsTmp.EOF Then
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If IsNull(rsTmp("carno")) = False Then cboCarnocst.AddItem rsTmp("carno")
            rsTmp.MoveNext
        Loop
End If

rsTmp.Close: Set rsTmp = Nothing

cboCustomer.ListIndex = -1: cboFloatCustomer.ListIndex = -1
optNo = 1: optNocst = 1
chkCustomer.Value = 1

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")

End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txtScst_Click()

Set objMvdateTarget = txtScst
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width + Frame2.Left, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtEcst_Click()

Set objMvdateTarget = txtEcst
mvDate.Move objMvdateTarget.Left + Frame2.Left, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtS_Click()

Set objMvdateTarget = txtS
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtE_Click()

Set objMvdateTarget = txtE
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub
Private Sub txtScst_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then mvDate.Visible = False
End Sub
Private Sub txtEcst_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then mvDate.Visible = False
End Sub
Private Sub txtS_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then mvDate.Visible = False
End Sub
Private Sub txtE_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then mvDate.Visible = False
End Sub
