VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_OP_LoadSorting 
   Caption         =   "½�O�z�f�޲z"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
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
   ScaleHeight     =   9960
   ScaleWidth      =   11610
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "�Ы��ƹ��������@�Ӥ���Υk��������"
      Top             =   3720
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
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   -2147483624
      StartOfWeek     =   61800449
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin VB.Frame Frame2 
      Caption         =   "�Ȥ����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   5160
      TabIndex        =   15
      Top             =   2280
      Width           =   6135
      Begin MSDataGridLib.DataGrid dgLoadSorting 
         Height          =   2295
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         TabAction       =   1
         AllowDelete     =   -1  'True
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
      Height          =   860
      Left            =   360
      Picture         =   "frm_OP_LoadSorting.frx":0000
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   1060
   End
   Begin VB.Frame Frame4 
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
      Height          =   2175
      Left            =   5160
      TabIndex        =   24
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�ק�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   1320
         Picture         =   "frm_OP_LoadSorting.frx":0312
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   11
         Top             =   1080
         Width           =   1065
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
         Height          =   880
         Left            =   3720
         Picture         =   "frm_OP_LoadSorting.frx":6B64
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   13
         Top             =   1080
         Width           =   1065
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFC0FF&
         Caption         =   "�R��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   2520
         Picture         =   "frm_OP_LoadSorting.frx":30776
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   12
         Top             =   1080
         Width           =   1060
      End
      Begin VB.CommandButton cmdAddNew 
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
         Height          =   880
         Left            =   120
         Picture         =   "frm_OP_LoadSorting.frx":317B8
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   10
         Top             =   1080
         Width           =   1060
      End
      Begin VB.ComboBox cboCarno 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox txtDriver 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2640
         MaxLength       =   8
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  '�m�����
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   720
         MaxLength       =   8
         TabIndex        =   6
         Top             =   240
         Width           =   1245
      End
      Begin VB.TextBox txtLoadSortingKey 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   7
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�r�p"
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
         Left            =   2040
         TabIndex        =   29
         Top             =   645
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���"
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
         Index           =   10
         Left            =   120
         TabIndex        =   27
         Top             =   285
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
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
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   9
         Left            =   120
         TabIndex        =   26
         Top             =   645
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�渹"
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
         Index           =   7
         Left            =   2040
         TabIndex        =   25
         Top             =   285
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "�z�f���"
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
      TabIndex        =   23
      Top             =   2280
      Width           =   5055
      Begin MSDataGridLib.DataGrid dgRoute 
         Height          =   2295
         Left            =   120
         TabIndex        =   5
         Top             =   240
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
      Height          =   2175
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtAppend 
         Alignment       =   1  '�a�k���
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   37
         Top             =   1680
         Width           =   1245
      End
      Begin VB.TextBox txtStamp 
         Alignment       =   1  '�a�k���
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   720
         MaxLength       =   20
         TabIndex        =   35
         Top             =   1680
         Width           =   1245
      End
      Begin VB.TextBox txtSorting 
         Alignment       =   1  '�a�k���
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2640
         MaxLength       =   20
         TabIndex        =   33
         Top             =   1320
         Width           =   1245
      End
      Begin VB.TextBox txtPallet 
         Alignment       =   1  '�a�k���
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   720
         MaxLength       =   20
         TabIndex        =   31
         Top             =   1320
         Width           =   1245
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
         Height          =   880
         Left            =   3840
         Picture         =   "frm_OP_LoadSorting.frx":3362A
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   4
         Top             =   240
         Width           =   1060
      End
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
         Left            =   4680
         Picture         =   "frm_OP_LoadSorting.frx":33934
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   720
         MaxLength       =   20
         TabIndex        =   2
         Top             =   600
         Width           =   1365
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   3
         Top             =   600
         Width           =   1365
      End
      Begin VB.TextBox txtOrderDateE 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   1
         Top             =   240
         Width           =   1365
      End
      Begin VB.TextBox txtOrderDateS 
         Alignment       =   2  '�m�����
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   720
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1365
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3885
         Y1              =   1240
         Y2              =   1240
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�\��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   2085
         TabIndex        =   38
         Top             =   1740
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�K��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   165
         TabIndex        =   36
         Top             =   1740
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�z�f"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   2085
         TabIndex        =   34
         Top             =   1380
         Width           =   390
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "½�O"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   165
         TabIndex        =   32
         Top             =   1380
         Width           =   390
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
         Index           =   5
         Left            =   2055
         TabIndex        =   22
         Top             =   660
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�渹"
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
         TabIndex        =   21
         Top             =   645
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���"
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
         TabIndex        =   20
         Top             =   285
         Width           =   480
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
         Index           =   1
         Left            =   2055
         TabIndex        =   19
         Top             =   300
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '������U��
      Height          =   270
      Left            =   0
      TabIndex        =   18
      Top             =   9690
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   476
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
            Object.Width           =   13864
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
Attribute VB_Name = "frm_OP_LoadSorting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsRoute As ADODB.Recordset
Private rsLoadSorting As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object

Private Sub cmdAddnew_Click()

'�M���S��r��
Call myFormExCharFilter(Me)

'����ˬd
If Not myIsDate(txtDate) Then Call txtdate_Click: Exit Sub
'If Len(RTrim(txtLoadSortingKey)) = 0 Then MsgBox "�п�J�渹!!", vbOKOnly, Me.Caption: txtLoadSortingKey.SetFocus: Exit Sub
'If Len(RTrim(cboCarno)) = 0 Then MsgBox "�п�J����!!", vbOKOnly, Me.Caption: cboCarno.SetFocus: Exit Sub

On Error GoTo err_Handle
dgLoadSorting.Col = 0
Dim rsTmp As New ADODB.Recordset

'�渹�ˬd
rsTmp.Open "select checkno from gt_loadsorting where checkno = '" & RTrim(txtLoadSortingKey) & "' ", cn
If Not rsTmp.EOF Then MsgBox "�t�γ渹����!(" & RTrim(txtLoadSortingKey) & ")", 64, "�s�W����!": rsTmp.Close: Exit Sub
rsTmp.Close

''�����ˬd
'rsTmp.Open "select driver from trp09m where vehicle_id_no = '" & RTrim(cboCarno) & "' ", cn
'If rsTmp.EOF Then MsgBox "�t�εL������!(" & RTrim(cboCarno) & ")", 64, "�s�W����!": rsTmp.Close: Exit Sub

rsLoadSorting.Filter = "½�O > 0 or �z�f > 0 or �K�� > 0 or �\�� > 0"
If rsLoadSorting.EOF Then MsgBox "�Цܤֿ�J�@���z�f���!", 64, "�s�W���": rsLoadSorting.Filter = "": rsLoadSorting.Sort = "�s��": Exit Sub

rsLoadSorting.MoveFirst

cn.BeginTrans: Tran_Level = 1

'�`�p��h�����渹���ƶq
txtPallet = txtPallet - rsRoute("½�O")
txtSorting = txtSorting - rsRoute("�z�f")
txtStamp = txtStamp - rsRoute("�K��")
txtAppend = txtAppend - rsRoute("�\��")

'�����渹���ƶq�k0
rsRoute("½�O") = 0
rsRoute("�z�f") = 0
rsRoute("�K��") = 0
rsRoute("�\��") = 0

''�ˬd�X���T�{�ᨮ���O�_�ۦP
'rsTmp.Close
'rsTmp.Open "select carno = rtrim(c_vehicle_id_no) from sdn01t where c_route_no = '" & RTrim(txtLoadSortingKey) & "' ", cn
'
'If Not rsTmp.EOF Then '�������s
'    If rsTmp("carno") <> RTrim(cboCarno) Then '��������
'        If MsgBox("�̪O�渹�P���u�s�� (" & txtLoadSortingKey & ") �A�X���T�{��������!" & vbCrLf & "�O�_�P�B��s�X���T�{�����H", vbOKCancel, "�̪O��s�W") = vbOK Then cn.Execute "update sdn01t set c_vehicle_id_no = '" & RTrim(cboCarno) & "',driver = '" & strDriver & "',editdate = getdate() , edituser = '" & user_id & "' where c_route_no = '" & RTrim(txtLoadSortingKey) & "' ", RowsAffect, adExecuteNoRecords
'    End If
'End If
 
'�R����
str_SQL = "delete gt_loadsorting where checkno = '" & RTrim(txtLoadSortingKey) & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
'�g�J�z�f���
Dim intLineNumber As Integer
Do While Not rsLoadSorting.EOF
    If Val(rsLoadSorting("½�O")) + Val(rsLoadSorting("�z�f")) + Val(rsLoadSorting("�K��")) + Val(rsLoadSorting("�\��")) = 0 Then cn.RollbackTrans: Tran_Level = 0: MsgBox "�p�O�ƶq���o���� 0 ?!", 16, Me.Caption: Exit Sub
    If Val(rsLoadSorting("½�O")) > 0 And Val(rsLoadSorting("�t�e���q")) > 0 And (Val(rsLoadSorting("�t�e���q")) = Val(rsLoadSorting("�z�f"))) Then cn.RollbackTrans: Tran_Level = 0: MsgBox "��½�O�ƮɡA�z�f����������t�e���q?!", 16, "�s�W����": Exit Sub
    
    intLineNumber = intLineNumber + 1
    
    str_SQL = "insert into gt_loadsorting(checkno,route_no,storer,sortingdate,carno,consigneekey,company,linenumber,weightqty,palletqty,sortingqty,stampqty,appendqty,notes,adduser,edituser,adddate,editdate) " & _
            "values('" & RTrim(txtLoadSortingKey) & "','" & RTrim(rsLoadSorting("���u�s��")) & "','" & RTrim(rsLoadSorting("�f�D")) & "','" & RTrim(txtDate) & "','" & UCase(RTrim(cboCarno)) & "','" & rsLoadSorting("�Ȥ�s��") & "','" & rsLoadSorting("�Ȥ�W��") & "','" & intLineNumber & "'," & Val(rsLoadSorting("�t�e���q")) & "," & Val(rsLoadSorting("½�O")) & "," & Val(rsLoadSorting("�z�f")) & "," & Val(rsLoadSorting("�K��")) & "," & Val(rsLoadSorting("�\��")) & ",'" & rsLoadSorting("�Ƶ�") & "','" & User_id & "','" & User_id & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�`�p���ƶq�[�`
    txtPallet = txtPallet + rsLoadSorting("½�O")
    txtSorting = txtSorting + rsLoadSorting("�z�f")
    txtStamp = txtStamp + rsLoadSorting("�K��")
    txtAppend = txtAppend + rsLoadSorting("�\��")
    
    '�����渹���ƶq�[�`
    rsRoute("½�O") = rsRoute("½�O") + rsLoadSorting("½�O")
    rsRoute("�z�f") = rsRoute("�z�f") + rsLoadSorting("�z�f")
    rsRoute("�K��") = rsRoute("�K��") + rsLoadSorting("�K��")
    rsRoute("�\��") = rsRoute("�\��") + rsLoadSorting("�\��")
    
    rsLoadSorting.MoveNext
Loop

cn.CommitTrans: Tran_Level = 0

rsLoadSorting.Filter = "": rsLoadSorting.Sort = "�s��"

MsgBox "�s�W����!", 0, RTrim(txtLoadSortingKey)

rsRoute("���") = RTrim(txtDate)
rsRoute("���@") = "V"
rsRoute("�渹") = RTrim(txtLoadSortingKey)
rsRoute("����") = RTrim(cboCarno)
rsRoute("����") = User_id
rsRoute("���ʤ��") = Format(Now, "yyyy-MM-dd hh:mm:ss")
   
Call dgRoute_RowColChange(dgRoute.Row, dgRoute.Col)

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
Call cmdQueryDetail_Click

End Sub

Private Sub cmdDelete_Click()
On Error GoTo err_Handle

If rsRoute Is Nothing Then Exit Sub
If rsRoute.RecordCount = 0 Then Exit Sub
If Len(Trim(rsRoute("���@"))) = 0 Then Exit Sub
If MsgBox("�渹�G" & Trim(txtLoadSortingKey) & " �T�w�R���H", vbOKCancel, Me.Caption) <> vbOK Then Exit Sub

cn.BeginTrans: Tran_Level = 1
  
    '�R����
    str_SQL = "delete gt_loadsorting where checkno = '" & Trim(txtLoadSortingKey) & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

Call cmdQueryDetail_Click

'�`�p��h�����渹���ƶq
txtPallet = txtPallet - rsRoute("½�O")
txtSorting = txtSorting - rsRoute("�z�f")
txtStamp = txtStamp - rsRoute("�K��")
txtAppend = txtAppend - rsRoute("�\��")

'�����渹���ƶq�k0
rsRoute("½�O") = 0
rsRoute("�z�f") = 0
rsRoute("�K��") = 0
rsRoute("�\��") = 0

rsRoute("���@") = ""
rsRoute("����") = ""
rsRoute("���ʤ��") = ""

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgRoute.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Set dgLoadSorting.DataSource = Nothing
txtDate = "": cboCarno = "": txtDriver = "": txtLoadSortingKey = ""
Dim chc_PalletNo As String, chc_DeliveryDate As String, chc_Storerkey As String

'���
chc_DeliveryDate = ""
If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_DeliveryDate = "and ��� between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
   chc_DeliveryDate = "and ��� = '" & txtOrderDateS.Text & "' "
ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_DeliveryDate = "and ��� = '" & txtOrderDateE.Text & "' "
End If

'�渹
chc_PalletNo = ""
If Len(Text1.Text) > 0 And Len(Text2.Text) > 0 Then
   chc_PalletNo = "and �渹 between '" & Text1.Text & "' and '" & Text2.Text & "' "
ElseIf Len(Text1.Text) > 0 And Len(Text2.Text) = 0 Then
   chc_PalletNo = "and �渹 = '" & Text1.Text & "' "
ElseIf Len(Text1.Text) = 0 And Len(Text2.Text) > 0 Then
   chc_PalletNo = "and �渹 = '" & Text2.Text & "' "
End If

str_SQL = "select * from gv_LoadSortingSource where 1 = 1 " & chc_DeliveryDate & chc_PalletNo & "order by ���,�渹 "

Dim rsTmp As New ADODB.Recordset
rsTmp.Open str_SQL, cn, adOpenStatic, adLockPessimistic
If rsTmp.EOF = True Then MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Call cmdQueryDetail_Click

txtPallet = 0: txtSorting = 0: txtStamp = 0: txtAppend = 0

If Not rsTmp.EOF Then
    
    '�`�p
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        
        txtPallet = txtPallet + rsTmp("½�O")
        txtSorting = txtSorting + rsTmp("�z�f")
        txtStamp = txtStamp + rsTmp("�K��")
        txtAppend = txtAppend + rsTmp("�\��")
'        If rsTmp("½�O") + rsTmp("�z�f") + rsTmp("�K��") + rsTmp("�\��") > 0 Then rsTmp("���@") = "V"
        rsTmp.MoveNext
    
    Loop
    
    rsTmp.MoveFirst
End If

Set rsRoute = New ADODB.Recordset
rsRoute.CursorLocation = adUseClient

Call Replication_Recordset(rsTmp, rsRoute)
rsTmp.Close: Set rsTmp = Nothing

Set dgRoute.DataSource = rsRoute: dgRoute.Visible = False
If rsRoute.EOF = False Then rsRoute.MoveFirst

Set dgRoute.DataSource = rsRoute
If rsRoute.RecordCount = 1 Then Call dgRoute_RowColChange(1, 1)

SetDataGridColWidth Me.Caption, dgRoute
StatusBar.Panels(2).Text = rsRoute.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgRoute.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub cmdQueryDetail_Click()

If rsRoute Is Nothing Then Exit Sub
If rsRoute.RecordCount = 0 Then Exit Sub

On Error GoTo err_Handle
dgLoadSorting.Visible = False
Screen.MousePointer = 11

'�z�f����
'str_SQL = "select �f�D " & _
'",���u�s��,�ϽX,�Ȥ�s��,�Ȥ�W��,�t�e���q,½�O " & _
'",�z�f = case when left(�ϽX,1) = 'N' and �z�f is null then 0 " & _
'"            when left(�ϽX,1) in ('W','T','E') and �z�f is null then 0 " & _
'"            when �z�f is null then �t�e���q " & _
'"            else �z�f end " & _
'",�K��,�\��,�Ƶ� " & _
'"from gv_LoadSortingDetail where �渹 = '" & rsRoute("�渹") & "' order by �Ȥ�s�� , ���u�s�� "

Call Confirm_Recordset_Closed(Tmp_rs)
Tmp_rs.Open "exec gs_LoadSortingDetail '" & rsRoute("�渹") & "' ", cn, adOpenStatic, adLockPessimistic

Set rsLoadSorting = New ADODB.Recordset: rsLoadSorting.CursorLocation = 3

Call Replication_Recordset(Tmp_rs, rsLoadSorting)
Tmp_rs.Close: Set Tmp_rs = Nothing

Set dgLoadSorting.DataSource = rsLoadSorting
SetDataGridColWidth Me.Caption, dgLoadSorting

'dgLoadSorting.Columns.item(0).Visible = False
dgLoadSorting.Col = 4
Screen.MousePointer = 0: dgLoadSorting.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub cmdExit_Click()
Set rsRoute = Nothing
Set rsLoadSorting = Nothing
Unload Me
End Sub

Private Sub cmdEdit_Click()

'�M���S��r��
Call myFormExCharFilter(Me)

'����ˬd
If Not myIsDate(txtDate) Then Call txtdate_Click: Exit Sub
If Len(RTrim(txtLoadSortingKey)) = 0 Then MsgBox "�п�J�渹!!", vbOKOnly, Me.Caption: txtLoadSortingKey.SetFocus: Exit Sub
If Len(RTrim(cboCarno)) = 0 Then MsgBox "�п�J����!!", vbOKOnly, Me.Caption: cboCarno.SetFocus: Exit Sub

On Error GoTo err_Handle
dgLoadSorting.Col = 0
Dim rsTmp As New ADODB.Recordset

'�渹�ˬd
rsTmp.Open "select checkno from gt_loadsorting where checkno = '" & RTrim(txtLoadSortingKey) & "' ", cn
If rsTmp.EOF Then MsgBox "�t�εL���渹!(" & RTrim(txtLoadSortingKey) & ")", 64, "��s����!": rsTmp.Close: Exit Sub
rsTmp.Close

''�����ˬd
'rsTmp.Close
'rsTmp.Open "select driver from trp09m where vehicle_id_no = '" & RTrim(cboCarno) & "' ", cn
'If rsTmp.EOF Then MsgBox "�t�εL������!(" & RTrim(cboCarno) & ")", 64, "��s����!": rsTmp.Close: Exit Sub

rsLoadSorting.Filter = "½�O > 0 or �z�f > 0 or �K�� > 0 or �\�� > 0"
If rsLoadSorting.EOF Then MsgBox "�Цܤֿ�J�@���z�f���!", 64, "��s���": rsLoadSorting.Filter = "": rsLoadSorting.Sort = "�s��": Exit Sub

rsLoadSorting.MoveFirst

cn.BeginTrans: Tran_Level = 1

'�`�p��h�����渹���ƶq
txtPallet = txtPallet - rsRoute("½�O")
txtSorting = txtSorting - rsRoute("�z�f")
txtStamp = txtStamp - rsRoute("�K��")
txtAppend = txtAppend - rsRoute("�\��")

'�����渹���ƶq�k0
rsRoute("½�O") = 0
rsRoute("�z�f") = 0
rsRoute("�K��") = 0
rsRoute("�\��") = 0

'�R����
str_SQL = "delete gt_loadsorting where checkno = '" & RTrim(txtLoadSortingKey) & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
'�g�J�z�f���
Dim intLineNumber As Integer
Do While Not rsLoadSorting.EOF
If Val(rsLoadSorting("½�O")) + Val(rsLoadSorting("�z�f")) + Val(rsLoadSorting("�K��")) + Val(rsLoadSorting("�\��")) = 0 Then cn.RollbackTrans: Tran_Level = 0: MsgBox "�ƶq���o���� 0 ?!", 16, Me.Caption: Exit Sub
If Val(rsLoadSorting("½�O")) > 0 And Val(rsLoadSorting("�t�e���q")) > 0 And (Val(rsLoadSorting("�t�e���q")) = Val(rsLoadSorting("�z�f"))) Then cn.RollbackTrans: Tran_Level = 0: MsgBox "��½�O�ƮɡA�z�f����������t�e���q?!", 16, "�s�W����": Exit Sub
    intLineNumber = intLineNumber + 1
    
    str_SQL = "insert into gt_loadsorting(checkno,route_no,storer,sortingdate,carno,consigneekey,company,linenumber,weightqty,palletqty,sortingqty,stampqty,appendqty,notes,adduser,edituser,adddate,editdate) " & _
            "values('" & RTrim(txtLoadSortingKey) & "','" & RTrim(rsLoadSorting("���u�s��")) & "','" & RTrim(rsLoadSorting("�f�D")) & "','" & RTrim(txtDate) & "','" & UCase(RTrim(cboCarno)) & "','" & rsLoadSorting("�Ȥ�s��") & "','" & rsLoadSorting("�Ȥ�W��") & "','" & intLineNumber & "'," & Val(rsLoadSorting("�t�e���q")) & "," & Val(rsLoadSorting("½�O")) & "," & Val(rsLoadSorting("�z�f")) & "," & Val(rsLoadSorting("�K��")) & "," & Val(rsLoadSorting("�\��")) & ",'" & rsLoadSorting("�Ƶ�") & "','" & User_id & "','" & User_id & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�`�p���ƶq�[�`
    txtPallet = txtPallet + rsLoadSorting("½�O")
    txtSorting = txtSorting + rsLoadSorting("�z�f")
    txtStamp = txtStamp + rsLoadSorting("�K��")
    txtAppend = txtAppend + rsLoadSorting("�\��")
    
    '�����渹���ƶq�[�`
    rsRoute("½�O") = rsRoute("½�O") + rsLoadSorting("½�O")
    rsRoute("�z�f") = rsRoute("�z�f") + rsLoadSorting("�z�f")
    rsRoute("�K��") = rsRoute("�K��") + rsLoadSorting("�K��")
    rsRoute("�\��") = rsRoute("�\��") + rsLoadSorting("�\��")
    
    rsLoadSorting.MoveNext
Loop

cn.CommitTrans: Tran_Level = 0

MsgBox "��s����!", 0, RTrim(txtLoadSortingKey)

    '��s
    rsRoute("���") = RTrim(txtDate)
    rsRoute("���@") = "V"
    rsRoute("�渹") = RTrim(txtLoadSortingKey)
    rsRoute("����") = RTrim(cboCarno)
    rsRoute("����") = User_id
    rsRoute("���ʤ��") = Format(Now, "yyyy-MM-dd hh:mm:ss")
    
Call cmdQueryDetail_Click

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub dgLoadSorting_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle
If rsLoadSorting Is Nothing Then Exit Sub
If rsLoadSorting.RecordCount = 0 Then Exit Sub

With dgLoadSorting
    '�����\���ܯS�w���
    If .Col < 7 Or .Col > 12 Then .Col = Abs(LastCol): Exit Sub
'
'    '���O
'    If .Col = 2 Then
'        ShowUserType
'    '�Ȥ�
'    ElseIf .Col = 3 Then
'        ShowCustomer
'    '��L
'    Else
''        ShowText
'    End If

End With
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub dgroute_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgRoute
If dg.DataSource Is Nothing Then Exit Sub

'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgLoadSorting_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgLoadSorting
If dg.DataSource Is Nothing Then Exit Sub

'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgRoute_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

'�P�@����
If LastRow = Empty Then Exit Sub

'�O�_�����
If rsRoute Is Nothing Then Exit Sub
If rsRoute.RecordCount = 0 Then Exit Sub

Screen.MousePointer = 11

txtDate = rsRoute("���")
txtLoadSortingKey = rsRoute("�渹"): Frame4.Caption = rsRoute("�渹")
cboCarno = rsRoute("����")
'txtDriver = rsRoute("�r�p")

Call cmdQueryDetail_Click
Screen.MousePointer = 0

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    Frame3.Height = Me.ScaleHeight - Frame1.Height - Frame1.Top - StatusBar.Height - 60
    dgRoute.Height = Frame3.Height - 360
    Frame2.Height = Me.ScaleHeight - Frame4.Height - Frame4.Top - StatusBar.Height - 60
    dgLoadSorting.Height = Frame2.Height - dgLoadSorting.Top - 120

End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    Frame2.Width = Me.ScaleWidth - Frame3.Width - 120
    dgLoadSorting.Width = Frame2.Width - 240
    dgRoute.Width = Frame3.Width - 240
    
End If

End Sub

Private Sub cmdReset_Click()

'���]
Call ClearForm_AllField(Me)
Call cmdQueryDetail_Click
'Combo1.ListIndex = 0

End Sub

Private Sub dgroute_HeadClick(ByVal ColIndex As Integer)

If dgRoute.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsRoute.Sort = dgRoute.Columns(ColIndex).Caption & " DESC"
    dgRoute.ClearSelCols
    intColumnIndex = 255

Else
    rsRoute.Sort = dgRoute.Columns(ColIndex).Caption
    dgRoute.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub

Private Sub dgLoadSorting_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Dim intPLKG As Integer
    intPLKG = 750
    If rsLoadSorting("�f�D") = "LTHL01" Then intPLKG = 550
    If rsLoadSorting("�f�D") = "LNSL01" Then intPLKG = 550
    
    '���}½�O
    If dgLoadSorting.Col = 7 Then
        rsLoadSorting("�z�f") = rsLoadSorting("�t�e���q") - dgLoadSorting * intPLKG
        If rsLoadSorting("�z�f") < 0 Or Left(rsLoadSorting("�ϽX"), 1) = "N" Or Left(rsLoadSorting("�ϽX"), 1) = "W" Then rsLoadSorting("�z�f") = 0
    End If

    SendKeys "{tab}"
End If

End Sub

Private Sub Form_Load()
On Error GoTo err_Handle
Dim i As Integer
StatusBar.Panels(2).Text = "0 ����ƦC"
StatusBar.Panels(3).Text = User_id

'����
Call Confirm_Recordset_Closed(Tmp_rs)
Tmp_rs.CursorLocation = adUseClient
Tmp_rs.Open "select distinct(����) from gv_LoadSortingsource order by ���� ", cn, adOpenKeyset, adLockPessimistic

If Not Tmp_rs.EOF Then

    Tmp_rs.MoveFirst
    For i = 0 To Tmp_rs.RecordCount - 1
        cboCarno.AddItem Tmp_rs("����")
        Tmp_rs.MoveNext
    Next
    Tmp_rs.Close
    cboCarno.ListIndex = -1

End If

txtOrderDateS = Format(Now, "YYYYMMDD")
'txtOrderDateE = Format(Now + 3, "YYYYMMDD")
Set Tmp_rs = Nothing
    
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsRoute = Nothing
Set rsLoadSorting = Nothing
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txtdate_Click()

Set objMvdateTarget = txtDate
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width + Frame4.Left, objMvdateTarget.Top + objMvdateTarget.Height + Frame4.Top
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtOrderDateE_Click()

Set objMvdateTarget = txtOrderDateE
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtOrderDateS_Click()

Set objMvdateTarget = txtOrderDateS
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtDate_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub txtOrderDateE_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub
Private Sub txtOrderDateS_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False
If mvDate.Value < (Now - 90) Then objMvdateTarget.Text = Format(Now - 90, "yyyymmdd"): MsgBox "�ȯ�ק�90�Ѥ����!", 64, "�W�L����": Exit Sub

End Sub
