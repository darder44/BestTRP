VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_OP_PalletxSorting 
   Caption         =   "�̪O�޲z"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13275
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
   ScaleWidth      =   13275
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   120
      TabIndex        =   23
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
      StartOfWeek     =   278003713
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin VB.TextBox txtFlash1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      MaxLength       =   8
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox txtFlash 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6960
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.ComboBox cboCustomer1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_OP_PalletxSorting.frx":0000
      Left            =   2640
      List            =   "frm_OP_PalletxSorting.frx":0002
      TabIndex        =   39
      TabStop         =   0   'False
      Text            =   "cboCustomer"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboUserType1 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_OP_PalletxSorting.frx":0004
      Left            =   3840
      List            =   "frm_OP_PalletxSorting.frx":0006
      TabIndex        =   38
      TabStop         =   0   'False
      Text            =   "cboUserType2"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboCustomer 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_OP_PalletxSorting.frx":0008
      Left            =   2640
      List            =   "frm_OP_PalletxSorting.frx":000A
      TabIndex        =   37
      TabStop         =   0   'False
      Text            =   "cboCustomer"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboUserType 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_OP_PalletxSorting.frx":000C
      Left            =   3840
      List            =   "frm_OP_PalletxSorting.frx":000E
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "cboUserType"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "�̪O����"
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
      Left            =   6360
      TabIndex        =   21
      Top             =   2280
      Width           =   6135
      Begin VB.CommandButton cmdDeletePalletDetail 
         BackColor       =   &H00FFC0FF&
         Caption         =   "�R��"
         Height          =   375
         Left            =   1080
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdAddPalletDetail 
         BackColor       =   &H0000FFFF&
         Caption         =   "�s�W"
         Height          =   375
         Left            =   120
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid dgPalletDetail 
         Height          =   2295
         Left            =   120
         TabIndex        =   16
         Top             =   720
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
      Picture         =   "frm_OP_PalletxSorting.frx":0010
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   6480
      Visible         =   0   'False
      Width           =   1060
   End
   Begin VB.Frame Frame6 
      Caption         =   "½�O�P�z�f����"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4000
      Left            =   1800
      TabIndex        =   20
      Top             =   6000
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdDeleteSortingCost 
         BackColor       =   &H00FFC0FF&
         Caption         =   "�R��"
         Height          =   375
         Left            =   1080
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdAddSortingCost 
         BackColor       =   &H0000FFFF&
         Caption         =   "�s�W"
         Height          =   375
         Left            =   120
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid dgSortingCost 
         Height          =   2295
         Left            =   120
         TabIndex        =   19
         Top             =   720
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
      Left            =   6360
      TabIndex        =   30
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
         Picture         =   "frm_OP_PalletxSorting.frx":0322
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
         Picture         =   "frm_OP_PalletxSorting.frx":6B74
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
         Picture         =   "frm_OP_PalletxSorting.frx":30786
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
         Picture         =   "frm_OP_PalletxSorting.frx":317C8
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   10
         Top             =   1080
         Width           =   1060
      End
      Begin VB.ComboBox cboCarno 
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
      Begin VB.TextBox txtPalletKey 
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
         TabIndex        =   35
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   285
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "�̪O��"
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
      TabIndex        =   29
      Top             =   1320
      Width           =   6255
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
      Height          =   1215
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmd2Excel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "��Excel"
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
         Left            =   5040
         Picture         =   "frm_OP_PalletxSorting.frx":3363A
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   46
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
         Height          =   880
         Left            =   3840
         Picture         =   "frm_OP_PalletxSorting.frx":34934
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
         Picture         =   "frm_OP_PalletxSorting.frx":34C3E
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   34
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   300
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '�������U��
      Height          =   270
      Left            =   0
      TabIndex        =   24
      Top             =   9690
      Width           =   13275
      _ExtentX        =   23416
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
            Object.Width           =   16801
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
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�٤J=�ܮw�J�w"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   8
      Left            =   11400
      TabIndex        =   45
      Top             =   720
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�ɥX=�ܮw�X�w"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   6
      Left            =   11400
      TabIndex        =   44
      Top             =   480
      Width           =   1590
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�̪O�b�O�H�ܮw������I"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   4
      Left            =   11400
      TabIndex        =   43
      Top             =   240
      Width           =   2640
   End
End
Attribute VB_Name = "frm_OP_PalletxSorting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsRoute As ADODB.Recordset
Private rsPalletDetail As ADODB.Recordset
Private rsSortingCost As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
'Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long
Private Sub cmd2Excel_Click()

'��ƱƧ�
Recordset2Excel "�̪O��", rsRoute
'..�b���s��EXCEL
Set MyXlsApp = Nothing

End Sub
'Private Sub cboCustomer_Change()
'Call cboCustomer_Click
'End Sub
Private Sub cboCustomer1_Change()
Call cboCustomer1_Click
End Sub

'Private Sub cboUserType_Change()
'Call cboUserType_Click
'End Sub
'
'Private Sub cboUserType1_Change()
'Call cboUserType1_Click
'End Sub

Private Sub cmdAddnew_Click()

'�M���S���r��
Call myFormExCharFilter(Me)

'����ˬd
If Not myIsDate(txtDate) Then Call txtdate_Click: Exit Sub
If Len(RTrim(txtPalletKey)) = 0 Then MsgBox "�п�J�渹!!", vbOKOnly, Me.Caption: txtPalletKey.SetFocus: Exit Sub
If Len(RTrim(cboCarno)) = 0 Then MsgBox "�п�J����!!", vbOKOnly, Me.Caption: cboCarno.SetFocus: Exit Sub
'If rsPalletDetail.RecordCount + rsSortingCost.RecordCount = 0 Then MsgBox "�п�J���Ӹ��!!", vbOKOnly, Me.Caption: Exit Sub

On Error GoTo err_Handle
dgPalletDetail.Col = 0: dgSortingCost.Col = 0
Dim rsTmp As New ADODB.Recordset

'�渹�ˬd
rsTmp.Open "select checkno from pallet_cds where checkno = '" & RTrim(txtPalletKey) & "' ", cn
If Not rsTmp.EOF Then MsgBox "�̪O�渹����!(" & RTrim(txtPalletKey) & ")", 64, "�s�W����!": rsTmp.Close: Exit Sub
rsTmp.Close

'�����ˬd
rsTmp.Open "select driver = isnull(driver,'') ,receiver = isnull(receiver,driver) from trp09m where vehicle_id_no = '" & RTrim(cboCarno) & "' ", cn
If rsTmp.EOF Then MsgBox "�t�εL������!(" & RTrim(cboCarno) & ")", 64, "�s�W����!": rsTmp.Close: Exit Sub

'�Ȧs���
Dim strDriver As String: strDriver = rsTmp("driver")
Dim strReceiver As String: strReceiver = rsTmp("Receiver")
rsTmp.Close

'��J�ɶ��ˬd
If txtDate <> Format(Now, "YYYYMMDD") Then
    If MsgBox("��ڤ���P���ʤ�����šI" & vbCrLf & "�O�_�T�w�H(�N�|�v�T���v�b)", vbOKCancel, "") <> vbOK Then: Exit Sub
End If

cn.BeginTrans: Tran_Level = 1

'�ˬd�X���T�{�ᨮ���O�_�ۦP--Mark by Gemini @20150717
rsTmp.Open "select carno = rtrim(c_vehicle_id_no) from sdn01t where c_route_no = '" & RTrim(txtPalletKey) & "' ", cn

If Not rsTmp.EOF Then '�������s--mark by Gemini @20150717
    If rsTmp("carno") <> RTrim(cboCarno) Then '��������
        MsgBox "�̪O�渹�P���u�s�� (" & txtPalletKey & ") �A�X���T�{�������šA�нT�{!", 16, "�̪O���s"
'        If MsgBox("�̪O�渹�P���u�s�� (" & txtPalletKey & ") �A�X���T�{��������!" & vbCrLf & "�O�_�P�B��s�X���T�{�����H", vbOKCancel, "�̪O��s�W") = vbOK Then cn.Execute "update sdn01t set c_vehicle_id_no = '" & RTrim(cboCarno) & "',driver = '" & strDriver & "',receiver = '" & strReceiver & "',editdate = getdate() , edituser = '" & User_id & "' where c_route_no = '" & RTrim(txtPalletKey) & "' ", RowsAffect, adExecuteNoRecords
    End If
End If

'�g�J���Y���
str_SQL = "insert into pallet_cds(checkno,storer,carno,usertype,adddate,adduser,edituser,editdate) " & _
    "values('" & RTrim(txtPalletKey) & "','BEST','" & UCase(RTrim(cboCarno)) & "','" & Company_id & "','" & RTrim(txtDate) & "','" & User_id & "','" & User_id & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
'�R������
str_SQL = "delete pallet_cst where checkno = '" & RTrim(txtPalletKey) & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
'�g�J�̪O���
If rsPalletDetail.RecordCount > 0 Then
    rsPalletDetail.MoveFirst
   
    Do While Not rsPalletDetail.EOF
        '�ˬd���O
        '�ˬd�Ȥ�
        If Len(RTrim(rsPalletDetail("���O"))) = 0 Or Len(RTrim(rsPalletDetail("�Ȥ�"))) = 0 Then MsgBox "�п�J�̪O���O�ΫȤ�?!", vbOKOnly, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub
        If Val(rsPalletDetail("�ɥX")) = 0 And Val(rsPalletDetail("�٤J")) = 0 Then MsgBox "�ɥX�P�٤J�ƶq���o���� 0 ?!", 64, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub

        str_SQL = "insert into pallet_cst(checkno,linenumber,storer,carno,usertype,customer,customersheetno,chargedate,qtyin,qtyout,sortingqty,notes,adddate,adduser,edituser,keyindate,editdate) " & _
                "values('" & RTrim(txtPalletKey) & "','" & RTrim(rsPalletDetail("����")) & "','BEST','" & UCase(RTrim(cboCarno)) & "','" & rsPalletDetail("���O") & "','" & rsPalletDetail("�Ȥ�") & "','" & rsPalletDetail("�Ȥ�渹") & "','" & RTrim(txtDate) & "','" & Val(rsPalletDetail("�ɥX")) & "','" & Val(rsPalletDetail("�٤J")) & "',0,'" & rsPalletDetail("���ӳƵ�") & "','" & RTrim(txtDate) & "','" & User_id & "','" & User_id & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        rsPalletDetail.MoveNext
    Loop
End If
    
' '�g�J�z�f���
'If rsSortingCost.RecordCount > 0 Then
'    rsSortingCost.MoveFirst
'
'    Do While Not rsSortingCost.EOF
'        If Len(RTrim(rsSortingCost("���O"))) = 0 Then MsgBox "�п�J���O?!", vbOKOnly, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub
'        If Val(rsSortingCost("�p�O�ƶq")) = 0 Then MsgBox "�p�O�ƶq���o�� 0 ?!", 64, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub
'
'        str_SQL = "insert into pallet_cst(checkno,linenumber,storer,carno,usertype,customer,customersheetno,chargedate,qtyin,qtyout,sortingqty,notes,adddate,adduser,edituser,keyindate,editdate) " & _
'                "values('" & RTrim(txtPalletKey) & "','" & rsSortingCost("����") & "','BEST','" & UCase(RTrim(cboCarno)) & "','" & rsSortingCost("���O") & "','" & rsSortingCost("�Ȥ�") & "','" & rsSortingCost("�Ȥ�渹") & "','" & RTrim(txtDate) & "',0,0,'" & Val(rsSortingCost("�p�O�ƶq")) & "','" & rsSortingCost("���ӳƵ�") & "','" & RTrim(txtDate) & "','" & User_id & "','" & User_id & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'        rsSortingCost.MoveNext
'    Loop
'End If

cn.CommitTrans: Tran_Level = 0
MsgBox "�s�W����!", 0, RTrim(txtPalletKey)

'�Ȧs���
Dim strPalletKey As String, strDate As String, strCarno As String
strPalletKey = RTrim(txtPalletKey)
strDate = RTrim(txtDate)
strCarno = RTrim(cboCarno)

rsRoute.Find "�渹 = '" & RTrim(strPalletKey) & "'"
If rsRoute.EOF Then rsRoute.AddNew

rsRoute("���") = RTrim(strDate)
rsRoute("���@") = "V"
rsRoute("�渹") = RTrim(strPalletKey)
rsRoute("����") = RTrim(strCarno)
rsRoute("����") = User_id
rsRoute("���ʤ��") = Format(Now, "yyyy-MM-dd hh:mm:ss")

If rsPalletDetail.RecordCount = 0 Then rsRoute("���@") = "X"
    
Call dgRoute_RowColChange(dgRoute.Row, dgRoute.Col)

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdAddPalletDetail_Click()
If rsPalletDetail Is Nothing Then Exit Sub

'������
Dim i As Integer, j As Integer
If rsPalletDetail.RecordCount > 0 Then rsPalletDetail.MoveLast: i = rsPalletDetail("����")
'If rsSortingCost.RecordCount > 0 Then rsSortingCost.MoveLast: j = rsSortingCost("����")

'�s�W
rsPalletDetail.AddNew

If i > j Then
    rsPalletDetail("����") = i + 1
Else
    rsPalletDetail("����") = j + 1
End If

rsPalletDetail("����") = User_id
rsPalletDetail("���ʤ��") = Format(Now, "yyyy-MM-dd hh:mm:ss")

End Sub

Private Sub cmdAddSortingCost_Click()
If rsSortingCost Is Nothing Then Exit Sub

'������
Dim i As Integer, j As Integer
If rsPalletDetail.RecordCount > 0 Then rsPalletDetail.MoveLast: i = rsPalletDetail("����")
If rsSortingCost.RecordCount > 0 Then rsSortingCost.MoveLast: j = rsSortingCost("����")

'�s�W
rsSortingCost.AddNew

If i > j Then
    rsSortingCost("����") = i + 1
Else
    rsSortingCost("����") = j + 1
End If

rsSortingCost("����") = User_id
rsSortingCost("���ʤ��") = Format(Now, "yyyy-MM-dd hh:mm:ss")

End Sub

Private Sub cmdDelete_Click()
On Error GoTo err_Handle

If rsRoute Is Nothing Then Exit Sub
If rsRoute.RecordCount = 0 Then Exit Sub
If Len(Trim(rsRoute("���@"))) = 0 Then Exit Sub

'��J�ɶ��ˬd
If txtDate <> Format(Now, "YYYYMMDD") Then
    If MsgBox("��ڤ���P���ʤ�����šI" & vbCrLf & "�O�_�T�w�H(�N�|�v�T���v�b)", vbOKCancel, "�`�N") <> vbOK Then: Exit Sub
End If

If MsgBox("�渹�G" & Trim(txtPalletKey) & " �T�w�R���H", vbOKCancel, Me.Caption) <> vbOK Then Exit Sub

cn.BeginTrans: Tran_Level = 1

    '�R�����Y
    str_SQL = "delete pallet_cds where checkno = '" & Trim(txtPalletKey) & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '�R������
    str_SQL = "delete pallet_cst where checkno = '" & Trim(txtPalletKey) & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

Call cmdQueryDetail_Click

rsRoute("���@") = ""
rsRoute("����") = ""
rsRoute("���ʤ��") = ""

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdDeletePalletDetail_Click()

If dgPalletDetail.DataSource Is Nothing Then Exit Sub
If rsPalletDetail Is Nothing Then Exit Sub
If rsPalletDetail.RecordCount = 0 Then Exit Sub
rsPalletDetail.Delete

End Sub

Private Sub cmdDeleteSortingCost_Click()

If dgSortingCost.DataSource Is Nothing Then Exit Sub
If rsSortingCost Is Nothing Then Exit Sub
If rsSortingCost.RecordCount = 0 Then Exit Sub
rsSortingCost.Delete

End Sub

Private Sub cmdQuery_Click()

If Len(RTrim(txtOrderDateS)) + Len(RTrim(txtOrderDateE)) = 0 Then MsgBox "�п�J����϶�!", 16, Me.Caption: Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgRoute.DataSource = Nothing: StatusBar.Panels(2).Text = "0 ����ƦC"
Set dgPalletDetail.DataSource = Nothing: Set dgSortingCost.DataSource = Nothing
txtDate = "": cboCarno = "": txtDriver = "": txtPalletKey = ""
Dim chc_PalletNo As String, chc_DeliveryDate As String, chc_PalletNo1 As String, chc_DeliveryDate1 As String, chc_Storerkey As String

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

str_SQL = "select distinct " & _
            "��� " & _
            ",���@ " & _
            ",�渹 " & _
            ",�S�� = '           ' " & _
            ",���� " & _
            ",����,���ʤ�� " & _
            "From gv_PalletDetail where 1 = 1 " & chc_DeliveryDate & chc_PalletNo & " order by ���,�渹 "

Dim rsTmp As New ADODB.Recordset
rsTmp.Open str_SQL, cn, adOpenStatic, adLockPessimistic
If rsTmp.EOF = True Then MsgBox "�d�L��ơI", vbOKOnly + vbInformation, Me.Caption: Call cmdQueryDetail_Click

Set rsRoute = New ADODB.Recordset
rsRoute.CursorLocation = adUseClient

Call Replication_Recordset(rsTmp, rsRoute)
rsTmp.Close

'���
chc_DeliveryDate1 = ""
If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_DeliveryDate1 = "and s1.Delivery_Date between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
   chc_DeliveryDate1 = "and s1.Delivery_Date = '" & txtOrderDateS.Text & "' "
ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_DeliveryDate1 = "and s1.Delivery_Date = '" & txtOrderDateE.Text & "' "
End If

'�渹
chc_PalletNo1 = ""
If Len(Text1.Text) > 0 And Len(Text2.Text) > 0 Then
   chc_PalletNo1 = "and s1.c_route_no between '" & Text1.Text & "' and '" & Text2.Text & "' "
ElseIf Len(Text1.Text) > 0 And Len(Text2.Text) = 0 Then
   chc_PalletNo1 = "and s1.c_route_no = '" & Text1.Text & "' "
ElseIf Len(Text1.Text) = 0 And Len(Text2.Text) > 0 Then
   chc_PalletNo1 = "and s1.c_route_no = '" & Text2.Text & "' "
End If

'�L��ƫh������
If Not rsRoute.EOF Then

    '���q��S�����
    str_SQL = "select distinct s2.c_route_no " & _
    ",�S�� = case when s2.priority = 'A2B' and s2.storerkey = 'LABT01' then 'A2B,�Ȱ�' " & _
    "when s2.priority = 'A2B' then 'A2B' " & _
    "when s2.storerkey = 'LABT01' then '�Ȱ�' " & _
    "Else '' end from sdn02t s2 join sdn01t s1 on s1.c_route_no = s2.c_route_no where s2.priority = 'A2B' or s2.storerkey = 'LABT01' " & chc_DeliveryDate1 & chc_PalletNo1
    
    rsTmp.Open str_SQL, cn, adOpenStatic, adLockPessimistic
    
    If rsTmp.EOF = True Then
    
    Else
    
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
        
            rsRoute.MoveFirst
            Do While Not rsRoute.EOF
                If rsTmp("c_route_no") = rsRoute("�渹") Then rsRoute("�S��") = rsTmp("�S��")
                rsRoute.MoveNext
            Loop
            
        rsTmp.MoveNext
        Loop
    
    End If
    
    rsTmp.Close:
End If

Set rsTmp = Nothing

Set dgRoute.DataSource = rsRoute: dgRoute.Visible = False
If rsRoute.EOF = False Then rsRoute.MoveFirst

Set dgRoute.DataSource = rsRoute

SetDataGridColWidth Me.Caption, dgRoute
StatusBar.Panels(2).Text = rsRoute.RecordCount & " ����ƦC"
Screen.MousePointer = 0: dgRoute.Visible = True

Call dgRoute_RowColChange(1, 1)

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQueryDetail_Click()
On Error GoTo err_Handle
dgPalletDetail.Visible = False: dgSortingCost.Visible = False
Screen.MousePointer = 11

'�̪O����
str_SQL = "select " & _
            "���� " & _
            ",���O " & _
            ",�Ȥ� " & _
            ",�ɥX " & _
            ",�٤J " & _
            ",�Ȥ�渹 " & _
            ",���ӳƵ� " & _
            ",���� = ���Ӳ��� " & _
            ",���ʤ�� = ���Ӳ��ʤ�� " & _
            "From gv_PalletDetail where ���� > 0 and �渹 = '" & RTrim(txtPalletKey) & "' order by ���� "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockPessimistic

'If tmp_Rs.EOF Then
'tmp_Rs.Close

''���s�Ȥ�
'str_SQL = "select distinct " & _
'            "���� = ' ' " & _
'            ",���O = '                   ' " & _
'            ",�Ȥ� = cast('' as char(45)) " & _
'            ",�q��Ȥ� = rtrim(cust_name) " & _
'            ",�ɥX = 0 " & _
'            ",�٤J = 0 " & _
'            ",�Ȥ�渹 = ' ' " & _
'            ",���ӳƵ� = ' ' " & _
'            ",���� = ' ' " & _
'            ",���ʤ�� = ' ' " & _
'            "From sdn02t where c_route_no = '" & RTrim(txtPalletKey) & "' "
'
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockPessimistic
'
'End If

Set rsPalletDetail = New ADODB.Recordset: rsPalletDetail.CursorLocation = 3

Call Replication_Recordset(tmp_Rs, rsPalletDetail)
tmp_Rs.Close: Set tmp_Rs = Nothing

Set dgPalletDetail.DataSource = rsPalletDetail
SetDataGridColWidth Me.Caption, dgPalletDetail

'�Ȥ���
''�z�f����
'str_SQL = "select " & _
'            "���� " & _
'            ",���O " & _
'            ",�Ȥ� " & _
'            ",�p�O�ƶq " & _
'            ",�Ȥ�渹 " & _
'            ",���ӳƵ� " & _
'            ",���� = ���Ӳ��� " & _
'            ",���ʤ�� = ���Ӳ��ʤ�� " & _
'            "From gv_PalletDetail where ���� > 0 and �渹 = '" & RTrim(txtPalletKey) & "' and ���O in ('½�O��','�z�f��','�K��','�\��') order by ���� "
'
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockPessimistic
'
'Set rsSortingCost = New ADODB.Recordset: rsSortingCost.CursorLocation = 3
'
'Call Replication_Recordset(tmp_Rs, rsSortingCost)
'tmp_Rs.Close: Set tmp_Rs = Nothing
'
'Set dgSortingCost.DataSource = rsSortingCost
'SetDataGridColWidth Me.Caption, dgSortingCost
'dgSortingCost.Columns.item(0).Visible = False
'dgSortingCost.Visible = True

dgPalletDetail.Columns.item(0).Visible = False
cboCustomer.Visible = False: cboCustomer1.Visible = False
cboUserType.Visible = False: cboUserType1.Visible = False
txtFlash.Visible = False: txtFlash1.Visible = False
Screen.MousePointer = 0: dgPalletDetail.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub ShowUserType()

With dgPalletDetail
    .RowHeight = cboUserType.Height - 10
    If .Col = 2 Then
        If .Columns(.Col).Left > 0 Then
                cboUserType.Visible = True
                cboUserType.Move .Left + .Columns(.Col).Left + Frame2.Left + 15, .Top + .RowTop(.Row) + Frame2.Top, .Columns(.Col).Width
                If cboUserType.Left + cboUserType.Width - Frame2.Left > .Left + .Width Then '�p�G���W�XDataGrid����ܽd�򪺳B�z
                    cboUserType.Width = cboUserType.Width + .Left + .Width - cboUserType.Left - cboUserType.Width
                End If
                cboUserType.Text = rsPalletDetail("���O") '��sCombo����
        Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
            cboUserType.Visible = False
        End If
    Else
        cboUserType.Visible = False
    End If
    
End With
End Sub

Private Sub cboUserType_Click()
On Error GoTo err_Handle

If rsPalletDetail Is Nothing Then Exit Sub
If rsPalletDetail.RecordCount = 0 Then Exit Sub
rsPalletDetail("���O") = cboUserType.Text

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
MsgBox "�нT�{�O�_�W�L10�X���׭���!!", vbOKOnly, Me.Caption: cboUserType.SetFocus
End Sub
Private Sub ShowUserType1()

With dgSortingCost
    .RowHeight = cboUserType.Height - 10
    If .Col = 2 Then
        If .Columns(.Col).Left > 0 Then
                cboUserType1.Visible = True
                cboUserType1.Move .Left + .Columns(.Col).Left + Frame6.Left + 15, .Top + .RowTop(.Row) + Frame6.Top, .Columns(.Col).Width
                '�p�G���W�XDataGrid����ܽd�򪺳B�z
                If cboUserType1.Left + cboUserType1.Width - Frame2.Left > .Left + .Width Then
                    cboUserType1.Width = cboUserType1.Width + .Left + .Width - cboUserType1.Left - cboUserType1.Width
                End If
                cboUserType1.Text = rsSortingCost("���O")  '��sCombo����
        Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
            cboUserType1.Visible = False
        End If
    Else
        cboUserType1.Visible = False
    End If
End With
End Sub
Private Sub cboUserType1_Click()
If rsSortingCost Is Nothing Then Exit Sub
If rsSortingCost.RecordCount = 0 Then Exit Sub
rsSortingCost("���O") = cboUserType1.Text

End Sub

Private Sub ShowCustomer()

With dgPalletDetail
    .RowHeight = cboUserType.Height - 10
    If .Col = 3 Then
        If .Columns(.Col).Left > 0 Then
                cboCustomer.Visible = True
                cboCustomer.Move .Left + .Columns(.Col).Left + Frame2.Left + 15, .Top + .RowTop(.Row) + Frame2.Top, .Columns(.Col).Width
                If cboCustomer.Left + cboCustomer.Width - Frame2.Left > .Left + .Width Then '�p�G���W�XDataGrid����ܽd�򪺳B�z
                    cboCustomer.Width = cboCustomer.Width + .Left + .Width - cboCustomer.Left - cboCustomer.Width
                End If
                cboCustomer.Text = rsPalletDetail("�Ȥ�")  '��sCombo����
        Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
            cboCustomer.Visible = False
        End If
    Else
        cboCustomer.Visible = False
    End If
End With
End Sub

Private Sub cboCustomer_Click()

If rsPalletDetail Is Nothing Then Exit Sub
If rsPalletDetail.RecordCount = 0 Then Exit Sub
rsPalletDetail("�Ȥ�") = cboCustomer.Text

End Sub

Private Sub ShowCustomer1()

With dgSortingCost
    .RowHeight = cboUserType.Height - 10
    If .Col = 3 Then
        If .Columns(.Col).Left > 0 Then
                cboCustomer1.Visible = True
                cboCustomer1.Move .Left + .Columns(.Col).Left + Frame6.Left + 15, .Top + .RowTop(.Row) + Frame6.Top, .Columns(.Col).Width
                If cboCustomer1.Left + cboCustomer1.Width - Frame6.Left > .Left + .Width Then '�p�G���W�XDataGrid����ܽd�򪺳B�z
                    cboCustomer1.Width = cboCustomer1.Width + .Left + .Width - cboCustomer1.Left - cboCustomer1.Width
                End If
                cboCustomer1.Text = rsSortingCost("�Ȥ�")  '��sCombo����
        Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
            cboCustomer1.Visible = False
        End If
    Else
        cboCustomer1.Visible = False
    End If
End With
End Sub

Private Sub cboCustomer1_Click()

If rsSortingCost Is Nothing Then Exit Sub
If rsSortingCost.RecordCount = 0 Then Exit Sub
rsSortingCost("�Ȥ�") = cboCustomer1.Text

End Sub

Private Sub ShowText1()

With dgSortingCost
.RowHeight = txtFlash1.Height - 10
    If .Columns(.Col).Left > 0 Then
            txtFlash1.Visible = True
            txtFlash1.Move .Left + .Columns(.Col).Left + Frame6.Left + 15, .Top + .RowTop(.Row) + Frame6.Top, .Columns(.Col).Width
            If txtFlash1.Left + txtFlash1.Width - Frame6.Left > .Left + .Width Then '�p�G���W�XDataGrid����ܽd�򪺳B�z
                txtFlash1.Width = txtFlash1.Width + .Left + .Width - txtFlash1.Left - txtFlash.Width
            End If
            txtFlash1.Text = rsSortingCost.Fields(.Col)  '��stxt����
            txtFlash1.SetFocus
    Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
        txtFlash1.Visible = False
    End If

End With
End Sub

Private Sub cmdEdit_Click()

'�M���S���r��
Call myFormExCharFilter(Me)

'����ˬd
If Not myIsDate(txtDate) Then Call txtdate_Click: Exit Sub
If Len(RTrim(txtPalletKey)) = 0 Then MsgBox "�п�J�渹!!", vbOKOnly, Me.Caption: txtPalletKey.SetFocus: Exit Sub
If Len(RTrim(cboCarno)) = 0 Then MsgBox "�п�J����!!", vbOKOnly, Me.Caption: cboCarno.SetFocus: Exit Sub
'If rsPalletDetail.RecordCount + rsSortingCost.RecordCount = 0 Then MsgBox "�п�J���Ӹ��!!", vbOKOnly, Me.Caption: Exit Sub

On Error GoTo err_Handle
dgPalletDetail.Col = 0: dgSortingCost.Col = 0
Dim rsTmp As New ADODB.Recordset

'�渹�ˬd
rsTmp.Open "select checkno from pallet_cds where checkno = '" & RTrim(txtPalletKey) & "' ", cn
If rsTmp.EOF Then MsgBox "�t�εL���渹!(" & RTrim(txtPalletKey) & ")", 64, "��s����!": rsTmp.Close: Exit Sub
rsTmp.Close

'�����ˬd
rsTmp.Open "select driver = isnull(driver,''),Receiver=isnull(receiver,driver) from trp09m where vehicle_id_no = '" & RTrim(cboCarno) & "' ", cn
If rsTmp.EOF Then MsgBox "�t�εL������!(" & RTrim(cboCarno) & ")", 64, "��s����!": rsTmp.Close: Exit Sub

'�Ȧs���
Dim strDriver As String: strDriver = rsTmp("driver")
Dim strReceiver As String: strReceiver = rsTmp("Receiver")
rsTmp.Close

''��J�ɶ��ˬd
'If txtDate <> Format(Now, "YYYYMMDD") Then
'    MsgBox "�ȭ��ק蘆��s�W����ơI", 64, "�`�N"
'    Exit Sub
'End If

'��J�ɶ��ˬd
If txtDate <> Format(Now, "YYYYMMDD") Then
    If MsgBox("��ڤ���P���ʤ�����šI" & vbCrLf & "�O�_�T�w�H(�N�|�v�T���v�b)", vbOKCancel, "�`�N") <> vbOK Then: Exit Sub
End If

'��J�ɶ��ˬd
If MsgBox("�T�{�ק�I", vbOKCancel, "�`�N") <> vbOK Then Call cmdQueryDetail_Click: Exit Sub

cn.BeginTrans: Tran_Level = 1

'�ˬd�X���T�{�ᨮ���O�_�ۦP--mark by Gemini @20150717
rsTmp.Open "select carno = rtrim(c_vehicle_id_no) from sdn01t where c_route_no = '" & RTrim(txtPalletKey) & "' ", cn

If Not rsTmp.EOF Then '�������s
    If rsTmp("carno") <> RTrim(cboCarno) Then '��������
        MsgBox "�̪O�渹�P���u�s�� (" & txtPalletKey & ") �A�X���T�{�������šA�нT�{!", 16, "�̪O���s"
'        If MsgBox("�̪O�渹�P���u�s�� (" & txtPalletKey & ") �A�X���T�{��������!" & vbCrLf & "�O�_�P�B��s�X���T�{�����H", vbOKCancel, "�̪O���s") = vbOK Then
'            cn.Execute "update sdn01t set c_vehicle_id_no = '" & RTrim(cboCarno) & "',driver = '" & strDriver & "',receiver = '" & strReceiver & "',editdate = getdate() , edituser = '" & User_id & "' where c_route_no = '" & RTrim(txtPalletKey) & "' ", RowsAffect, adExecuteNoRecords
'
'        End If

    End If
End If

'��s���Y
    str_SQL = "update pallet_cds set " & _
              "carno = '" & UCase(RTrim(cboCarno)) & "' " & _
              ",adddate = '" & RTrim(txtDate) & "' " & _
              ",edituser = '" & User_id & "' " & _
              ",editdate = '" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "' " & _
              "where checkno = '" & RTrim(txtPalletKey) & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
'�R������
str_SQL = "delete pallet_cst where checkno = '" & RTrim(txtPalletKey) & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
'�g�J�̪O���
If rsPalletDetail.RecordCount > 0 Then
    rsPalletDetail.MoveFirst
   
    Do While Not rsPalletDetail.EOF
        If Len(RTrim(rsPalletDetail("���O"))) = 0 Or Len(RTrim(rsPalletDetail("�Ȥ�"))) = 0 Then MsgBox "�п�J�̪O���O�ΫȤ�?!", vbOKOnly, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub
        If Val(rsPalletDetail("�ɥX")) = 0 And Val(rsPalletDetail("�٤J")) = 0 Then MsgBox "�ɥX�P�٤J�ƶq���o���� 0 ?!", 64, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub

        str_SQL = "insert into pallet_cst(checkno,linenumber,storer,carno,usertype,customer,customersheetno,chargedate,qtyin,qtyout,sortingqty,notes,adddate,adduser,edituser,keyindate,editdate) " & _
                "values('" & RTrim(txtPalletKey) & "','" & RTrim(rsPalletDetail("����")) & "','BEST','" & UCase(RTrim(cboCarno)) & "','" & rsPalletDetail("���O") & "','" & rsPalletDetail("�Ȥ�") & "','" & rsPalletDetail("�Ȥ�渹") & "','" & RTrim(txtDate) & "','" & Val(rsPalletDetail("�ɥX")) & "','" & Val(rsPalletDetail("�٤J")) & "',0,'" & rsPalletDetail("���ӳƵ�") & "','" & RTrim(txtDate) & "','" & User_id & "','" & User_id & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        rsPalletDetail.MoveNext
    Loop
End If
    
 '�g�J�z�f���
'If rsSortingCost.RecordCount > 0 Then
'    rsSortingCost.MoveFirst
'
'    Do While Not rsSortingCost.EOF
'        If Len(RTrim(rsSortingCost("���O"))) = 0 Then MsgBox "�п�J���O?!", vbOKOnly, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub
'        If Val(rsSortingCost("�p�O�ƶq")) = 0 Then MsgBox "�p�O�ƶq���o�� 0 ?!", 64, Me.Caption: cn.RollbackTrans: Tran_Level = 1: Exit Sub
'
'        str_SQL = "insert into pallet_cst(checkno,linenumber,storer,carno,usertype,customer,customersheetno,chargedate,qtyin,qtyout,sortingqty,notes,adddate,adduser,edituser,keyindate,editdate) " & _
'                "values('" & RTrim(txtPalletKey) & "','" & rsSortingCost("����") & "','BEST','" & UCase(RTrim(cboCarno)) & "','" & rsSortingCost("���O") & "','" & rsSortingCost("�Ȥ�") & "','" & rsSortingCost("�Ȥ�渹") & "','" & RTrim(txtDate) & "',0,0,'" & Val(rsSortingCost("�p�O�ƶq")) & "','" & rsSortingCost("���ӳƵ�") & "','" & RTrim(txtDate) & "','" & User_id & "','" & User_id & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'        rsSortingCost.MoveNext
'    Loop
'End If

cn.CommitTrans: Tran_Level = 0
MsgBox "��s����!", 0, RTrim(txtPalletKey)

    '��s
    rsRoute("���") = RTrim(txtDate)
    rsRoute("���@") = "V"
    rsRoute("�渹") = RTrim(txtPalletKey)
    rsRoute("����") = RTrim(cboCarno)
    rsRoute("����") = User_id
    rsRoute("���ʤ��") = Format(Now, "yyyy-MM-dd hh:mm:ss")
    
If rsPalletDetail.RecordCount = 0 Then rsRoute("���@") = "X"

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dgPalletDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle
If rsPalletDetail Is Nothing Then Exit Sub
If rsPalletDetail.RecordCount = 0 Then Exit Sub

With dgPalletDetail
    '�����\���ܯS�w���
    If .Col < 2 Or .Col > 7 Then .Col = Abs(LastCol): Exit Sub
    cboCustomer.Visible = False: cboCustomer1.Visible = False
    cboUserType.Visible = False: cboUserType1.Visible = False
    txtFlash.Visible = False: txtFlash1.Visible = False
    
    '���O
    If .Col = 2 Then
        ShowUserType
    '�Ȥ�
    ElseIf .Col = 3 Then
        ShowCustomer
    '��L
    Else
'        ShowText
    End If

End With
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dgPalletDetail_Scroll(Cancel As Integer)
If cboUserType.Visible = True Then ShowUserType
If cboCustomer.Visible = True Then ShowCustomer
End Sub

Private Sub dgSortingCost_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle
If rsSortingCost Is Nothing Then Exit Sub
If rsSortingCost.RecordCount = 0 Then Exit Sub
cboCustomer.Visible = False: cboCustomer1.Visible = False
cboUserType.Visible = False: cboUserType1.Visible = False
txtFlash.Visible = False: txtFlash1.Visible = False
        
With dgSortingCost
    '�����\���ܯS�w���
    If .Col < 2 Or .Col > 6 Then .Col = Abs(LastCol): Exit Sub

    '���O
    If .Col = 2 Then
        ShowUserType1
    '�Ȥ�
    ElseIf .Col = 3 Then
        ShowCustomer1
    '��L
    Else
'        ShowText1
'        txtFlash1.SelStart = 0: txtFlash1.SelLength = Len(txtFlash1.Text)
'        txtFlash1.SetFocus
'        DoEvents: DoEvents
    End If

End With
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dgroute_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgRoute
'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgPalletDetail_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgPalletDetail

'�L��Ʃ���e�Ӥp�A���s�e��
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgSortingCost_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgSortingCost

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
If rsRoute.EOF Then Exit Sub

txtDate = rsRoute("���")
txtPalletKey = rsRoute("�渹"): Frame4.Caption = rsRoute("�渹")
cboCarno = rsRoute("����")
'txtDriver = rsRoute("�r�p")
Call cmdQueryDetail_Click

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dgSortingCost_Scroll(Cancel As Integer)
If cboUserType1.Visible = True Then ShowUserType1
If cboCustomer1.Visible = True Then ShowCustomer1
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '�̤p��

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    Frame3.Height = Me.ScaleHeight - Frame1.Height - Frame1.Top - StatusBar.Height - 60
    dgRoute.Height = Frame3.Height - 360
    dgPalletDetail.Height = Frame2.Height - dgPalletDetail.Top - 120
    dgSortingCost.Height = Frame6.Height - dgSortingCost.Top - 120
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    Frame2.Width = Me.ScaleWidth - Frame3.Width - 120: Frame6.Width = Frame2.Width
    dgPalletDetail.Width = Frame2.Width - 240: dgSortingCost.Width = Frame6.Width - 240
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

Private Sub dgPalletDetail_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub Form_Load()
On Error GoTo err_Handle
Dim i As Integer
StatusBar.Panels(2).Text = "0 ����ƦC"
StatusBar.Panels(3).Text = User_id

'�̪O���O
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
'tmp_Rs.Open "select distinct(���O) from gv_palletdetail order by ���O", cn, adOpenKeyset, adLockPessimistic
tmp_Rs.Open "select distinct(PalletType) as ���O from trp20m order by PalletType", cn, adOpenKeyset, adLockPessimistic
If Not tmp_Rs.EOF Then
tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        cboUserType.AddItem RTrim(tmp_Rs("���O"))
        tmp_Rs.MoveNext
    Next
    cboUserType.ListIndex = 0
End If
tmp_Rs.Close

''�z�f���O
'For i = 1 To 4
'cboUserType1.AddItem Choose(i, "½�O��", "�z�f��", "�K��", "�\��")
'Next
'cboUserType1.ListIndex = 0

'�Ȥ�
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
'tmp_Rs.Open "select distinct(�Ȥ�) from gv_palletdetail order by �Ȥ� ", cn, adOpenKeyset, adLockPessimistic
tmp_Rs.Open "select distinct(PalletCustomer) as �Ȥ� from trp21m order by PalletCustomer ", cn, adOpenKeyset, adLockPessimistic
If Not tmp_Rs.EOF Then
    tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        cboCustomer.AddItem RTrim(tmp_Rs("�Ȥ�"))
        cboCustomer1.AddItem RTrim(tmp_Rs("�Ȥ�"))
        tmp_Rs.MoveNext
    Next
    cboCustomer.ListIndex = 0: cboCustomer1.ListIndex = 0
End If
tmp_Rs.Close

'����
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(����) from gv_palletdetail order by ���� ", cn, adOpenKeyset, adLockPessimistic
If Not tmp_Rs.EOF Then
    tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        cboCarno.AddItem tmp_Rs("����")
        tmp_Rs.MoveNext
    Next
    cboCarno.ListIndex = -1
End If
tmp_Rs.Close

txtOrderDateS = Format(Now, "YYYYMMDD")
'txtOrderDateE = Format(Now + 3, "YYYYMMDD")
Set tmp_Rs = Nothing

'Call cmdQuery_Click
'Call cmdQueryDetail_Click
    
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsRoute = Nothing
Set rsPalletDetail = Nothing
Set rsSortingCost = Nothing
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txtdate_Click()

Set objMvdateTarget = txtDate
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width + Frame4.Left, objMvdateTarget.Top + objMvdateTarget.Height + Frame4.Top
mvDate.Visible = True: mvDate.Value = Now
    
End Sub

Private Sub txtFlash1_Change()

If rsSortingCost Is Nothing Then Exit Sub
rsSortingCost.Fields(dgSortingCost.Col) = txtFlash1.Text

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

End Sub