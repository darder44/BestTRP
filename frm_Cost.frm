VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Cost 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�B�O���@"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   14505
   StartUpPosition =   2  '�ù�����
   Begin VB.TextBox txt_Cartype 
      Alignment       =   1  '�a�k���
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   102
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txt_Stairs 
      Alignment       =   1  '�a�k���
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   100
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox txt_ReceiveCash 
      Alignment       =   1  '�a�k���
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   9480
      TabIndex        =   98
      ToolTipText     =   "��ڦ��{���B"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txt_Cash 
      Alignment       =   1  '�a�k���
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   9480
      TabIndex        =   96
      ToolTipText     =   "�U�f���{���B"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtB_City 
      Alignment       =   1  '�a�k���
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   95
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtChannelType 
      Alignment       =   1  '�a�k���
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   93
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtReserve_Mark 
      Alignment       =   1  '�a�k���
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   90
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtUrgent_Mark 
      Alignment       =   1  '�a�k���
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   89
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txt_TotalCost 
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   5040
      TabIndex        =   75
      ToolTipText     =   "�t�e���`�ҭl�ͥX���O�ΦX�p"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txt_TRPCost 
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1200
      TabIndex        =   74
      ToolTipText     =   "�t�e���`�ҭl�ͥX���t�e�O"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txt_SortingCost 
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   3240
      TabIndex        =   73
      ToolTipText     =   "�t�e���`�ҭl�ͥX���z�f�O"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdAbnormalCostUpdate 
      BackColor       =   &H00FF80FF&
      Caption         =   "���`�O�Χ�s"
      Height          =   375
      Left            =   6000
      TabIndex        =   72
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdReDeliveryCharge 
      BackColor       =   &H00FF80FF&
      Caption         =   "�A�t�p��"
      Height          =   375
      Left            =   8280
      TabIndex        =   71
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtCBM2 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   1320
      Width           =   960
   End
   Begin VB.TextBox txtOT2 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   2040
      Width           =   960
   End
   Begin VB.TextBox txtWT2 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   1680
      Width           =   960
   End
   Begin VB.TextBox txtCube2 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   960
      Width           =   960
   End
   Begin VB.TextBox txtCS2 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   600
      Width           =   960
   End
   Begin VB.TextBox txtEA2 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   2400
      Width           =   960
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "�s��"
      Height          =   375
      Left            =   3360
      TabIndex        =   62
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtEA1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   2400
      Width           =   960
   End
   Begin VB.TextBox txtCS1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   600
      Width           =   960
   End
   Begin VB.TextBox txtCube1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   960
      Width           =   960
   End
   Begin VB.TextBox txtWT1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   1680
      Width           =   960
   End
   Begin VB.TextBox txtOT1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   2040
      Width           =   960
   End
   Begin VB.TextBox txtCBM1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   1320
      Width           =   960
   End
   Begin VB.TextBox txtCBM 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   1320
      Width           =   960
   End
   Begin VB.TextBox txtOT 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   2040
      Width           =   960
   End
   Begin VB.CommandButton cmdCost 
      Caption         =   "�p�O�Ѧ�"
      Height          =   375
      Left            =   120
      TabIndex        =   46
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtWT 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   1680
      Width           =   960
   End
   Begin VB.TextBox txtCube 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   960
      Width           =   960
   End
   Begin VB.TextBox txtCS 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   600
      Width           =   960
   End
   Begin VB.TextBox txtEA 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   2400
      Width           =   960
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FF80FF&
      Caption         =   "���}"
      Height          =   375
      Left            =   4440
      TabIndex        =   36
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "�R��"
      Height          =   375
      Left            =   2280
      TabIndex        =   35
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "�s�W"
      Height          =   375
      Left            =   1200
      TabIndex        =   34
      Top             =   3240
      Width           =   975
   End
   Begin VB.ComboBox cboCostCode 
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
      ItemData        =   "frm_Cost.frx":0000
      Left            =   5400
      List            =   "frm_Cost.frx":0002
      TabIndex        =   33
      Text            =   "cboCostCode"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox cboCostKind 
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
      ItemData        =   "frm_Cost.frx":0004
      Left            =   6720
      List            =   "frm_Cost.frx":0006
      TabIndex        =   32
      Text            =   "cboCostKind"
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSaveExit 
      Caption         =   "�s�����}"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9840
      TabIndex        =   31
      Top             =   3240
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dgMain 
      Height          =   3615
      Left            =   120
      TabIndex        =   54
      Top             =   3720
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   6376
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.TextBox txt_OneOrder_OrderKey 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   120
         Width           =   1080
      End
      Begin VB.TextBox txt_OneOrder_FullName1 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   1320
         Width           =   2835
      End
      Begin VB.TextBox txt_OneOrder_ConsigneeKey1 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   1320
         Width           =   1080
      End
      Begin VB.TextBox txt_ZIP1 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   1620
         Width           =   1080
      End
      Begin VB.TextBox txt_OneOrder_Address1 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   1620
         Width           =   4155
      End
      Begin VB.TextBox txtArea1 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   1320
         Width           =   1320
      End
      Begin VB.TextBox txt_C_ROUTE_NO 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtArea 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   720
         Width           =   1320
      End
      Begin VB.TextBox txt_OneOrder_Receiver 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9600
         TabIndex        =   49
         Top             =   1620
         Width           =   1080
      End
      Begin VB.TextBox txt_OneOrder_Address 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1020
         Width           =   4155
      End
      Begin VB.TextBox txt_ZIP 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1020
         Width           =   1080
      End
      Begin VB.TextBox txt_OneOrder_ArriveDate 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1020
         Width           =   1080
      End
      Begin VB.TextBox txt_OneOrder_Description 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   420
         Width           =   5235
      End
      Begin VB.TextBox txt_OneOrder_ConsigneeKey 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox txt_OneOrder_OrderDate 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   1080
      End
      Begin VB.TextBox txt_OneOrder_VehicleID 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txt_OneOrder_Driver 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1620
         Width           =   1575
      End
      Begin VB.TextBox txt_OneOrder_TRPCompany 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1020
         Width           =   1575
      End
      Begin VB.TextBox txt_OneOrder_FullName 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   2835
      End
      Begin VB.TextBox txt_OneOrder_Status 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9600
         TabIndex        =   7
         Top             =   1320
         Width           =   1080
      End
      Begin VB.TextBox txt_OneOrder_StorerOrderKey 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7275
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   1440
      End
      Begin VB.TextBox txt_Priority 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   420
         Width           =   1080
      End
      Begin VB.TextBox txt_OneOrder_StorerKey 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   120
         Width           =   1080
      End
      Begin VB.TextBox txt_OneOrder_RouteNo 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   420
         Width           =   1575
      End
      Begin VB.TextBox txt_OneOrder_DeliveryDate 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txt_Storer 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   1875
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "TMS�渹"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   8760
         TabIndex        =   88
         Top             =   180
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���I"
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
         Height          =   195
         Index           =   23
         Left            =   3000
         TabIndex        =   86
         Top             =   1500
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�G�����s"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   56
         Left            =   120
         TabIndex        =   80
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�дڤH"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   8880
         TabIndex        =   50
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�_�I"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   55
         Left            =   3000
         TabIndex        =   30
         Top             =   960
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�q�渹�X"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   6435
         TabIndex        =   27
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�q��Ƶ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   26
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "��f���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   8760
         TabIndex        =   25
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�B�e����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   1380
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�r�p�m�W"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�B�餽�q"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "ñ�檬�A"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   8760
         TabIndex        =   21
         Top             =   1380
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�q����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   8760
         TabIndex        =   20
         Top             =   780
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�q�����O"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   54
         Left            =   8760
         TabIndex        =   19
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�f�D�s��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2640
         TabIndex        =   18
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���u�s��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�X�����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   16
         Top             =   780
         Width           =   780
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�p�O���O"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   195
      Index           =   29
      Left            =   6000
      TabIndex        =   103
      Top             =   2940
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "���_�Ӽh"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Index           =   28
      Left            =   4200
      TabIndex        =   101
      Top             =   2940
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�ꦬ�N���f��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   25
      Left            =   8160
      TabIndex        =   99
      Top             =   2580
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�����N���f��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   21
      Left            =   8160
      TabIndex        =   97
      Top             =   2220
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�q���O"
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
      Height          =   195
      Index           =   4
      Left            =   4440
      TabIndex        =   94
      Top             =   2580
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�M��"
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
      Height          =   195
      Index           =   27
      Left            =   2760
      TabIndex        =   92
      Top             =   2580
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "���"
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
      Height          =   195
      Index           =   26
      Left            =   720
      TabIndex        =   91
      Top             =   2580
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�O�ΦX�p"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   49
      Left            =   4200
      TabIndex        =   78
      Top             =   2220
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "���`�t�e�O"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   50
      Left            =   120
      TabIndex        =   77
      Top             =   2220
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "���`�z�f�O"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   51
      Left            =   2160
      TabIndex        =   76
      Top             =   2220
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "���p�O"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9960
      TabIndex        =   70
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��f��x�a�}"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   12360
      TabIndex        =   69
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "���sx�f�D"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   13440
      TabIndex        =   61
      Top             =   360
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "CBM"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   11040
      TabIndex        =   52
      Top             =   1380
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   11280
      TabIndex        =   48
      Top             =   2100
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�q��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   11760
      TabIndex        =   45
      Top             =   360
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   11280
      TabIndex        =   44
      Top             =   1740
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   11280
      TabIndex        =   42
      Top             =   1020
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�c"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   11280
      TabIndex        =   40
      Top             =   660
      Width           =   195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   11280
      TabIndex        =   38
      Top             =   2460
      Width           =   195
   End
End
Attribute VB_Name = "frm_Cost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DelRecord
Private rsMain As New ADODB.Recordset
Private intColumnIndex As Integer

Private Sub cmd2Excel_Click()

'��ƱƧ�
Recordset2Excel Me.Caption, rsMain

'..�b���s��EXCEL
If rsMain Is Nothing Then
Else
    With MyXlsApp
                
    End With
End If
Set MyXlsApp = Nothing

End Sub

Private Sub cboCostKind_Click()
rsMain("�д����O") = cboCostkind.Text
cboCostCode.Clear
Dim i As Integer, j As Integer

j = -1

'���p�O�N�X���
Dim rsTmp As New ADODB.Recordset
rsTmp.Open "select CostCode = rtrim(isnull(CostCode,'')) from trp17m where costkind = '" & rsMain("�д����O") & "' and storerkey = '" & txt_OneOrder_StorerKey.Text & "' and costkind <> 'notuse' ", cn

If Not rsTmp.EOF Then
    rsTmp.MoveFirst: i = 0
    Do While Not rsTmp.EOF
        If Right(rsTmp("CostCode"), 3) = RTrim(txt_Zip1.Text) Then j = i
        cboCostCode.AddItem rsTmp("CostCode")
        rsTmp.MoveNext: i = i + 1
    Loop
End If

rsTmp.Close: Set rsTmp = Nothing

cboCostCode.ListIndex = j
If Len(RTrim(rsMain("�Ƶ�"))) = 0 And rsMain("�д����O") = "�ﰪ���O�ɶK" Then rsMain("�Ƶ�") = "�ﰪ���O�ɶK"

End Sub

Private Sub cboCostCode_Click()
rsMain("�p�O�N�X") = cboCostCode.Text

'��scboCostCode
Dim rsTmp As New ADODB.Recordset
rsTmp.Open "select receivable=isnull(receivable,0) ,UOM=rtrim(isnull(uom,'')), payable=isnull(payable,0) ,areastart= rtrim(areastart),areaend = rtrim(areaend) from trp17m where costcode = '" & rsMain("�p�O�N�X") & "' and storerkey = '" & txt_OneOrder_StorerKey.Text & "' ", cn

If rsTmp.EOF Then Exit Sub
'rsMain("����") = rsMain("�p�O�N�X")
rsMain("���") = rsTmp("UOM")
rsMain("�������") = rsTmp("receivable")
rsMain("���I���") = rsTmp("payable")
rsMain("�з�����") = rsTmp("receivable")
rsMain("�з����I") = rsTmp("payable")
rsMain("�����`��") = rsTmp("receivable") * rsMain("�ƶq")
rsMain("���I�`��") = rsTmp("payable") * rsMain("�ƶq")
rsMain("�_�I") = rsTmp("areastart") & ""
rsMain("���I") = rsTmp("areaend") & ""

rsTmp.Close: Set rsTmp = Nothing

End Sub

Private Sub cmdAbnormalCostUpdate_Click()

frm_OP_SDNConfirm.txt_TRPCost.Text = txt_TRPCost.Text
frm_OP_SDNConfirm.txt_SortingCost.Text = txt_SortingCost.Text
frm_OP_SDNConfirm.txt_TotalCost.Text = txt_TotalCost.Text

str_SQL = "update sdn02t set trp_cost = '" & txt_TRPCost & "' ,sorting_cost = '" & txt_SortingCost & "',total_cost = '" & txt_TotalCost & "' where receipt_no = '" & frm_OP_SDNConfirm.txt_OneOrder_OrderKey & "' "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

End Sub

Private Sub cmdAddnew_Click()
dgMain.AllowAddNew = True
rsMain.AddNew
rsMain("�s��") = rsMain.RecordCount
dgMain.AllowAddNew = False
End Sub

Private Sub cmdCost_Click()

On Error GoTo err_Handle
'ñ��O�_�T�{
If Len(RTrim(frm_OP_SDNConfirm.txt_OneOrder_Status)) = 0 Then MsgBox "�|��ñ��T�{�I", 16, "�p�O": cmdSave.Enabled = False: cmdSaveExit.Enabled = False: Exit Sub

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select * from sdn05t (nolock) where sdn_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey) & "' ", cn
If Not tmp_Rs.EOF Then MsgBox "���q��w���B�O��ơA�t�Τ��A�p��B�O!!", 16, "�p�O": Exit Sub

Screen.MousePointer = 11

'�B�O�p��
cn.Execute "exec gs_Cost '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey) & "' ", RowsAffect, adExecuteNoRecords

'���X�q��p�O�ƶq��s��0
If RTrim(frm_OP_SDNConfirm.txt_OneOrder_Status) = "���X�q��" Then
    cn.Execute "update sdn05t set chargeqty = 0 , sumreceivable = 0 ,sumpayable = 0 where sdn_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_Status) & "' ", RowsAffect, adExecuteNoRecords
End If

'���p�O���
str_SQL = "select �д����O = rtrim(costkind) " & _
            ",�p�O�N�X = rtrim(costcode) " & _
            ",���� = rtrim(sdn_name) " & _
            ",��] = rtrim(rtrim(reason)) " & _
            ",������� = receivable " & _
            ",���I��� = payable " & _
            ",�ƶq = round(chargeqty,9) " & _
            ",��� = rtrim(uom) " & _
            ",�����`�� = sumreceivable " & _
            ",���I�`�� = sumpayable " & _
            ",ĳ�� = premiam " & _
            ",�p�O���� = isnull(vehicle_id_no,'') " & _
            ",�Ƶ� = rtrim(isnull(note,'')) " & _
            ",�_�I = rtrim(isnull(areastart,'')) " & _
            ",���I = rtrim(isnull(areaend,'')) " & _
            ",�з����� = stdreceivable " & _
            ",�з����I = stdpayable " & _
            "from sdn05t (nolock) where sdn_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' order by sdn_name , costkind , receivable desc"
            
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

Call Replication_Recordset(tmp_Rs, rsMain)

If Not rsMain.EOF Then rsMain.MoveFirst
Set dgMain.DataSource = rsMain

'�����e��
SetDataGridColWidth "�B�O���@", dgMain

cn.Execute "delete sdn05t where sdn_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey) & "' ", RowsAffect, adExecuteNoRecords

Screen.MousePointer = 0
Label1.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdDelete_Click()
If rsMain.EOF Or rsMain.RecordCount = 0 Then Exit Sub

dgMain.Col = 0
If InStr(1, rsMain("�Ƶ�"), "�M����") <> 0 Then
    If MsgBox("����t���M��ĳ���A�T�w�R�������p�O���Ӹ��?", vbOKCancel, Me.Caption) = vbOK Then
        rsMain.Delete
        MsgBox "�R���M���p���A�аȥ��T�O����M���p�����T�I", 16, "�`�N"
    End If
Else
    rsMain.Delete
End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo err_Handle

'If RTrim(txt_ReceiveCash.Text) <> "" Then
'    If Val(RTrim(txt_ReceiveCash)) <> Val(RTrim(txt_Cash)) Then
'        DelRecord = MsgBox("�����N���f��<>�ꦬ�N���f�ڡA�нT�{��ƬO�_���T?", vbQuestion + vbYesNo, "�p�O�s��")
'        If DelRecord = vbNo Then
'            Exit Sub
'        End If
'    End If
'End If

dgMain.Col = 0

cn.Execute "delete sdn05t where sdn_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' ", RowsAffect, adExecuteNoRecords
'If rsMain.RecordCount = 0 Then Call cmdExit_Click: Exit Sub

''��s�дڤH
'cn.Execute "update SDN01T set receiver = (select isnull(receiver,driver) from trp09m where vehicle_id_no = '" & mySplit(txt_OneOrder_VehicleID, "_", 0) & "') where c_route_no = '" & txt_C_ROUTE_NO & "'", RowsAffect, adExecuteNoRecords

rsMain.MoveFirst
Do While Not rsMain.EOF
    
    If Len(RTrim(rsMain("�д����O"))) > 0 Then
    
    If Round(rsMain("�������") * rsMain("�ƶq"), 0) <> Round(rsMain("�����`��"), 0) Then MsgBox "������� x �ƶq <> �����`��!!", 16, "�`�N"
    If Round(rsMain("���I���") * rsMain("�ƶq"), 0) <> Round(rsMain("���I�`��"), 0) Then MsgBox "���I��� x �ƶq <> ���I�`��!!", 16, "�`�N"
    If Len(RTrim(rsMain("�_�I"))) = 0 Or Len(RTrim(rsMain("���I"))) = 0 Then MsgBox "�_�I�Ψ��I����ť�!!", 16, "�`�N"
    If Len(RTrim(rsMain("�p�O�N�X"))) = 0 Or Len(RTrim(rsMain("�д����O"))) = 0 Then MsgBox "�д����O�έp�O�N�X����ť�!!", 16, "�`�N"
    
    str_SQL = "insert into sdn05t (c_route_no,uom,chargeqty,receivable,payable,premiam,reason,sumreceivable,sumpayable,areastart,areaend,note,sdn_name,sdn_no,costkind,costcode,stdreceivable,stdpayable,vehicle_id_no) " & _
    "values( '" & txt_C_ROUTE_NO.Text & "','" & rsMain("���") & "','" & rsMain("�ƶq") & "','" & rsMain("�������") & "','" & rsMain("���I���") & "','" & rsMain("ĳ��") & "','" & rsMain("��]") & "','" & rsMain("�����`��") & "','" & rsMain("���I�`��") & "','" & rsMain("�_�I") & "','" & rsMain("���I") & "','" & Trim(rsMain("�Ƶ�")) & "','" & rsMain("����") & _
    "','" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "','" & rsMain("�д����O") & "','" & rsMain("�p�O�N�X") & "','" & rsMain("�з�����") & "','" & rsMain("�з����I") & "','" & rsMain("�p�O����") & "' )"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
    End If
    rsMain.MoveNext
Loop

'�������u
cn.Execute "exec Es_ARnoDistribution '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' ", RowsAffect, adExecuteNoRecords
'��s�^Orders���ꦬ�A�p�G�O�ťաA�h�ꦬ=�����A���ȫh�ꦬ=�ꦬ
'If RTrim(txt_ReceiveCash.Text) = "" Then
'    cn.Execute "update o set o.receiveCash = '" & Val(RTrim(txt_Cash.Text)) & "' from orders o join sdn02t s2 on o.orderkey = s2.receipt_no where s2.receipt_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' and o.type <> '�R��'", RowsAffect, adExecuteNoRecords
'    txt_ReceiveCash.Text = Val(RTrim(txt_Cash.Text))
'Else
'    cn.Execute "update o set o.receiveCash = '" & Val(RTrim(txt_ReceiveCash.Text)) & "' from orders o join sdn02t s2 on o.orderkey = s2.receipt_no where s2.receipt_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' and o.type <> '�R��'", RowsAffect, adExecuteNoRecords
'    txt_ReceiveCash.Text = Val(RTrim(txt_ReceiveCash.Text))
'End If


Label1.Visible = False

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
   
End Sub

Private Sub cmdSaveExit_Click()

'dgMain.Col = 0
'
'cn.Execute "delete sdn05t where sdn_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' ", RowsAffect, adExecuteNoRecords
'If rsMain.RecordCount = 0 Then Call cmdExit_Click: Exit Sub
'
'rsMain.MoveFirst
'Do While Not rsMain.EOF
'
'    If Len(RTrim(rsMain("�д����O"))) > 0 Then
'
'    If Int(rsMain("�������") * rsMain("�ƶq")) <> Int(rsMain("�����`��")) Then MsgBox "�`�N!����������H�ƶq�����������`��!!"
'    If Int(rsMain("���I���") * rsMain("�ƶq")) <> Int(rsMain("���I�`��")) Then MsgBox "�`�N!���I������H�ƶq���������I�`��!!"
'    If Len(RTrim(rsMain("�_�I"))) = 0 Or Len(RTrim(rsMain("���I"))) = 0 Then MsgBox "�`�N!�_�I�Ψ��I����ť�!!"
'
'        str_SQL = "insert into sdn05t (c_route_no,uom,chargeqty,receivable,payable,premiam,reason,sumreceivable,sumpayable,areastart,areaend,note,sdn_name,sdn_no,costkind,costcode,stdreceivable,stdpayable) " & _
'        "values( '" & txt_C_Route_NO.Text & "','" & rsMain("���") & "','" & rsMain("�ƶq") & "','" & rsMain("�������") & "','" & rsMain("���I���") & "','" & rsMain("ĳ��") & "','" & rsMain("��]") & "','" & rsMain("�����`��") & "','" & rsMain("���I�`��") & "','" & rsMain("�_�I") & "','" & rsMain("���I") & "','" & rsMain("�Ƶ�") & "','" & rsMain("����") & _
'        "','" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "','" & rsMain("�д����O") & "','" & rsMain("�p�O�N�X") & "','" & rsMain("�з�����") & "','" & rsMain("�з����I") & "' )"
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    End If
'    rsMain.MoveNext
'Loop
Call cmdSave_Click
Call cmdExit_Click

End Sub

Private Sub cmdReDeliveryCharge_Click()

'ñ��O�_�T�{
If Len(RTrim(frm_OP_SDNConfirm.txt_OneOrder_Status)) = 0 Then MsgBox "�|��ñ��T�{�I", 16, "�p�O": cmdSave.Enabled = False: cmdSaveExit.Enabled = False: Exit Sub

If MsgBox("1.����ݥ��p�O�s��" & vbCr & vbLf & "2.���[�p�д����O�t�z�f�O���p�O�N�X" & vbCr & vbLf & "3.���[�p RePalletIs �P Cancel �p�O�N�X" & vbCr & vbLf & "4.�Ƶ��}�Y�G���t�e���[�p" & vbCr & vbLf & "5.�s�W���p�O�A�Ƶ��N���O�G���t�e", vbOKCancel, "�p�O����") <> vbOK Then Exit Sub

'�B�O�p��
cn.Execute "exec gs_ReDeliveryCharge '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' ", RowsAffect, adExecuteNoRecords

Call Form_Load
  
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

    If Len(dgMain.Columns(ColIndex).DataField) = 0 Then Exit Sub
    SaveSetting App.title, "�B�O���@" & "dgMain", dgMain.Columns(ColIndex).DataField, dgMain.Columns(ColIndex).Width
    
End Sub

Private Sub dgMain_DblClick()

    If dgMain.Columns(dgMain.Col).DataField = "�_�I" And Len(RTrim(dgMain.Columns(1))) > 0 Then dgMain = txtArea
    If dgMain.Columns(dgMain.Col).DataField = "���I" And Len(RTrim(dgMain.Columns(1))) > 0 Then dgMain = txtArea1
End Sub

Private Sub dgMain_HeadClick(ByVal ColIndex As Integer)
If dgMain.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMain.Sort = dgMain.Columns(ColIndex).Caption & " DESC"
    dgMain.ClearSelCols
    intColumnIndex = 255

Else
    rsMain.Sort = dgMain.Columns(ColIndex).Caption
    dgMain.ClearSelCols
    intColumnIndex = ColIndex

End If
End Sub

Private Sub ShowCostKind()

With dgMain
    .RowHeight = cboCostkind.Height - 10
    If .Col = 1 Then
        If .Columns(.Col).Left > 0 Then
                cboCostkind.Visible = True
                cboCostkind.Move .Left + .Columns(.Col).Left + 15, .Top + .RowTop(.Row), .Columns(.Col).Width
                If cboCostkind.Left + cboCostkind.Width > .Left + .Width Then '�p�G���W�XDataGrid����ܽd�򪺳B�z
                    cboCostkind.Width = cboCostkind.Width + .Left + .Width - cboCostkind.Left - cboCostkind.Width
                End If
                cboCostkind.Text = rsMain("�д����O")  '��sCombo����
        Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
            cboCostkind.Visible = False
        End If
    Else
        cboCostkind.Visible = False
    End If
End With

Call cboCostKind_Click

End Sub

Private Sub ShowCostCode()

With dgMain
    .RowHeight = cboCostCode.Height - 10
    If .Col = 2 Then
        If .Columns(.Col).Left > 0 Then
                cboCostCode.Visible = True
                cboCostCode.Move .Left + .Columns(.Col).Left + 15, .Top + .RowTop(.Row), .Columns(.Col).Width
                If cboCostCode.Left + cboCostCode.Width > .Left + .Width Then '�p�G���W�XDataGrid����ܽd�򪺳B�z
                    cboCostCode.Width = cboCostCode.Width + .Left + .Width - cboCostCode.Left - cboCostCode.Width
                End If
                cboCostCode.Text = rsMain("�p�O�N�X")  '��sCombo����
        Else '�p�G�α��b���ʥX�FDataGrid����ܽd��A�ȷ|�p��0
            cboCostCode.Visible = False
        End If
    Else
        cboCostCode.Visible = False
    End If
End With
End Sub

Private Sub dgmain_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{tab}"

If KeyAscii = 9 Then

    '���}�������
    If dgMain.Col = 5 Then rsMain("�����`��") = dgMain * rsMain("�ƶq")
    
    '���}���I���
    If dgMain.Col = 6 Then rsMain("���I�`��") = dgMain * rsMain("�ƶq")
    
    '���}�ƶq
    If dgMain.Col = 7 Then rsMain("�����`��") = rsMain("�������") * dgMain: rsMain("���I�`��") = rsMain("���I���") * dgMain
    
End If

End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

With dgMain
    cboCostkind.Visible = False: cboCostCode.Visible = False
    '
    'If dgMain.Col = 3 And cmdPickSave.Enabled = True Then ShowList
    
    '�����\���ܯS�w���
    If .Col = 9 Or .Col = 10 Or .Col > 15 Then .Col = Abs(LastCol): Exit Sub

    '    If LastCol = 3 Then dgPick.Col = 5: Exit Sub
    '    If LastCol = 5 Then dgPick.Col = 2: Exit Sub
    '    dgPick.Col = IIf(LastCol = -1, 5, LastCol)
    'End If
    ''��ƦC�O�_�ܧ�
    'If LastRow = Empty Then Exit Sub
    
    '�д����O
    If .Col = 1 Then ShowCostKind
    
    '�дڥN�X
    If .Col = 2 Then ShowCostCode
           
    Screen.MousePointer = 0

End With
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dgMain_Scroll(Cancel As Integer)
If cboCostkind.Visible = True Then ShowCostKind
If cboCostCode.Visible = True Then ShowCostCode
End Sub

Private Sub Form_Load()
On Error GoTo err_Handle
'If Len(RTrim(frm_OP_SDNConfirm.txt_ZIP.Text)) = 0 And (txt_Priority = "R" Or txt_Priority = "RC" Or txt_Priority = "A2B") Then MsgBox "���Ȥ�L�l���ϸ�", 64, "�`�N!": Exit Sub
'If Len(RTrim(frm_OP_SDNConfirm.txt_Zip1.Text)) = 0 And (txt_Priority <> "R" Or txt_Priority <> "RC" Or txt_Priority <> "A2B") Then MsgBox "���Ȥ�L�l���ϸ�", 64, "�`�N!": Exit Sub
Dim strConsigneeKey As String, strAddress As String
Dim Str_Skuxpack
Screen.MousePointer = 11

'���ӭq���`�ӽc�������
str_SQL = "select EA = Sum(IsNull(s3.ship_qty, 0)) " & _
                ",CS = sum(case when isnull(s.casecnt,0) = 0 then 0 else s3.ship_qty/s.casecnt end) " & _
                ",CUBE = sum(s3.ship_qty * s.stdcube) " & _
                ",CBM = sum(s3.ship_qty * s.stdcube) / 35.315 " & _
                ",WGT = sum(s3.ship_qty * s.stdgrosswgt) " & _
                ",OT = isnull((select otqty from trp02t (nolock) where s3.receipt_no = trp02t.receipt_no ),0) + isnull((select otqty from ort02t (nolock) where s3.receipt_no = ort02t.receipt_no ),0) " & _
                "from sdn03t s3 (nolock) join gv_skuxpack s(nolock) on s3.storerkey = s.storerkey and s3.product_no = s.sku " & _
                "where s3.receipt_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' group by s3.receipt_no "
            
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

txtEA = tmp_Rs("EA")
txtCS = Round(tmp_Rs("CS"), 9)
txtCube = Round(tmp_Rs("CUBE"), 9)
txtCBM = Round(tmp_Rs("CBM"), 9)
txtWT = Round(tmp_Rs("WGT"), 9)
txtOT = tmp_Rs("OT")

'�����sx�f�D�X�f���
str_SQL = "select CS = sum(case when sp.casecnt = 0 then 0 else (s3.Ship_qty /sp.casecnt) end) " & _
            ",Cube = sum(s3.ship_qty * sp.stdcube) " & _
            ",CBM = sum(s3.ship_qty * sp.stdcube) /35.315 " & _
            ",WGT = sum(s3.ship_qty * sp.stdgrosswgt) " & _
            ",EA = sum(s3.ship_qty) " & _
            ",OT = isnull((select sum(isnull(otqty,0)) from trp02t (nolock) where receipt_no in(select receipt_no from sdn02t (nolock) where storerkey = '" & frm_OP_SDNConfirm.txt_OneOrder_StorerKey.Text & "' and c_route_no = '" & RTrim(frm_OP_SDNConfirm.txt_C_ROUTE_NO.Text) & "')),0) + isnull((select sum(isnull(otqty,0)) from ort02t (nolock) where receipt_no in(select receipt_no from sdn02t (nolock) where storerkey = '" & frm_OP_SDNConfirm.txt_OneOrder_StorerKey.Text & "' and c_route_no = '" & RTrim(frm_OP_SDNConfirm.txt_C_ROUTE_NO.Text) & "')),0) " & _
            "from sdn03t s3 (nolock) join gv_skuxpack sp(nolock) on sp.storerkey = s3.storerkey and sp.sku = s3.product_no " & _
            "where s3.storerkey = '" & frm_OP_SDNConfirm.txt_OneOrder_StorerKey.Text & "' " & _
            "and s3.c_route_no = '" & RTrim(frm_OP_SDNConfirm.txt_C_ROUTE_NO.Text) & "' "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

txtEA1 = tmp_Rs("EA")
txtCS1 = Round(tmp_Rs("CS"), 9)
txtCube1 = Round(tmp_Rs("CUBE"), 9)
txtCBM1 = Round(tmp_Rs("CBM"), 9)
txtWT1 = Round(tmp_Rs("WGT"), 9)
txtOT1 = tmp_Rs("OT")

If RTrim(frm_OP_SDNConfirm.txt_Priority) = "R" Or RTrim(frm_OP_SDNConfirm.txt_Priority) = "RC" Then
    strConsigneeKey = frm_OP_SDNConfirm.txt_OneOrder_ConsigneeKey
Else
    strConsigneeKey = frm_OP_SDNConfirm.txt_OneOrder_ConsigneeKey1
End If

If RTrim(frm_OP_SDNConfirm.txt_Priority) = "R" Or RTrim(frm_OP_SDNConfirm.txt_Priority) = "RC" Then
    strAddress = frm_OP_SDNConfirm.txt_OneOrder_Address
Else
    strAddress = frm_OP_SDNConfirm.txt_OneOrder_Address1
End If

'����f��x�a�}
str_SQL = "select CS = sum(case when sp.casecnt = 0 then 0 else (s3.Ship_qty /sp.casecnt) end) " & _
            ",Cube = sum(s3.ship_qty * sp.stdcube) " & _
            ",CBM = sum(s3.ship_qty * sp.stdcube) /35.315 " & _
            ",WGT = sum(s3.ship_qty * sp.stdgrosswgt) " & _
            ",EA = sum(s3.ship_qty) " & _
            ",Address = t1m.address " & _
            "from sdn03t s3 (nolock) join gv_skuxpack sp(nolock) on sp.storerkey = s3.storerkey and sp.sku = s3.product_no " & _
            "join sdn02t s2 (nolock) on s2.receipt_no = s3.receipt_no and s2.priority = '" & frm_OP_SDNConfirm.txt_Priority & "' " & _
            "join orders o (nolock) on o.orderkey = s2.c_receipt_no and o.type <> '�R��' and o.storerkey = '" & frm_OP_SDNConfirm.txt_OneOrder_StorerKey.Text & "' " & _
            "join trp01m t1m (nolock) on t1m.storerkey = s2.storerkey and t1m.address = '" & strAddress & "' and t1m.storerkey = '" & frm_OP_SDNConfirm.txt_OneOrder_StorerKey.Text & "' " & _
            "where convert(char(8),s2.arrive_date,112) = '" & frm_OP_SDNConfirm.txt_OneOrder_ArriveDate & "' and t1m.consigneekey = case when o.priority = 'A2B' then o.b_company else o.consigneekey end " & _
            "group by t1m.address"

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

txtEA2 = tmp_Rs("EA")
txtCS2 = Round(tmp_Rs("CS"), 9)
txtCube2 = Round(tmp_Rs("CUBE"), 9)
txtCBM2 = Round(tmp_Rs("CBM"), 9)
txtWT2 = Round(tmp_Rs("WGT"), 9)

'����f��x�a�}�`���
str_SQL = "select OT = isnull((select sum(isnull(otqty,0)) from trp02t (nolock) join trp01m (nolock) on trp02t.storerkey = trp01m.storerkey and  case when rtrim(trp02t.priority) = 'A2B' then trp02t.bconsigneekey else trp02t.consigneekey end = trp01m.consigneekey and convert(char(8),trp02t.arrive_date,112) = '" & frm_OP_SDNConfirm.txt_OneOrder_ArriveDate & "' and trp01m.address = '" & tmp_Rs("Address") & "' and trp02t.storerkey = '" & frm_OP_SDNConfirm.txt_OneOrder_StorerKey.Text & "' ),0) " & _
          "+ isnull((select sum(isnull(otqty,0)) from ort02t (nolock) join trp01m (nolock) on ort02t.storerkey = trp01m.storerkey and case when rtrim(ort02t.priority) = 'A2B' then ort02t.bconsigneekey else ort02t.consigneekey end = trp01m.consigneekey and convert(char(8),ort02t.arrive_date,112) = '" & frm_OP_SDNConfirm.txt_OneOrder_ArriveDate & "' and trp01m.address = '" & tmp_Rs("Address") & "' and ort02t.storerkey = '" & frm_OP_SDNConfirm.txt_OneOrder_StorerKey.Text & "' ),0) "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

txtOT2 = tmp_Rs("OT")

If Val(txtEA2) > Val(txtEA1) Then txtEA2.BackColor = &HFF&

'���p�O���
str_SQL = "select �д����O = rtrim(costkind) " & _
            ",�p�O�N�X = rtrim(costcode) " & _
            ",���� = rtrim(sdn_name) " & _
            ",��] = rtrim(rtrim(reason)) " & _
            ",������� = receivable " & _
            ",���I��� = payable " & _
            ",�ƶq = round(chargeqty,9) " & _
            ",��� = rtrim(uom) " & _
            ",�����`�� = sumreceivable " & _
            ",���I�`�� = sumpayable " & _
            ",ĳ�� = premiam " & _
            ",�p�O���� = rtrim(isnull(vehicle_id_no,''))" & _
            ",�Ƶ� = rtrim(isnull(note,'')) " & _
            ",�_�I = rtrim(isnull(areastart,'')) " & _
            ",���I = rtrim(isnull(areaend,'')) " & _
            ",�з����� = stdreceivable " & _
            ",�з����I = stdpayable " & _
            "from sdn05t (nolock) where sdn_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' order by sdn_name , costkind , receivable desc"
            
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

Call Replication_Recordset(tmp_Rs, rsMain)

If Not rsMain.EOF Then rsMain.MoveFirst
Set dgMain.DataSource = rsMain

'�����e��
SetDataGridColWidth "�B�O���@", dgMain

txt_C_ROUTE_NO.Text = frm_OP_SDNConfirm.txt_C_ROUTE_NO.Text
txt_OneOrder_RouteNo.Text = frm_OP_SDNConfirm.txt_OneOrder_RouteNo.Text
txt_OneOrder_VehicleID.Text = frm_OP_SDNConfirm.txt_OneOrder_VehicleID.Text
txt_OneOrder_Driver.Text = frm_OP_SDNConfirm.txt_OneOrder_Driver.Text
txt_OneOrder_TRPCompany.Text = frm_OP_SDNConfirm.txt_OneOrder_TRPCompany.Text
txt_OneOrder_DeliveryDate.Text = frm_OP_SDNConfirm.txt_OneOrder_DeliveryDate.Text
txt_OneOrder_StorerKey.Text = frm_OP_SDNConfirm.txt_OneOrder_StorerKey.Text
txt_Storer.Text = frm_OP_SDNConfirm.txt_Storer.Text
txt_OneOrder_Description.Text = frm_OP_SDNConfirm.txt_OneOrder_Description.Text
txt_OneOrder_ConsigneeKey.Text = frm_OP_SDNConfirm.txt_OneOrder_ConsigneeKey.Text
txt_OneOrder_FullName.Text = frm_OP_SDNConfirm.txt_OneOrder_FullName.Text
txt_OneOrder_Address.Text = frm_OP_SDNConfirm.txt_OneOrder_Address.Text
txt_OneOrder_OrderKey.Text = frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text
txt_OneOrder_OrderDate.Text = frm_OP_SDNConfirm.txt_OneOrder_OrderDate.Text
txt_OneOrder_ArriveDate.Text = frm_OP_SDNConfirm.txt_OneOrder_ArriveDate.Text
txt_OneOrder_Status.Text = frm_OP_SDNConfirm.txt_OneOrder_Status.Text
txt_OneOrder_StorerOrderKey.Text = frm_OP_SDNConfirm.txt_OneOrder_StorerOrderKey.Text
txt_Priority.Text = frm_OP_SDNConfirm.txt_Priority.Text
txt_ZIP.Text = frm_OP_SDNConfirm.txt_ZIP.Text
txt_TRPCost.Text = frm_OP_SDNConfirm.txt_TRPCost.Text
txt_SortingCost.Text = frm_OP_SDNConfirm.txt_SortingCost.Text
txt_TotalCost.Text = frm_OP_SDNConfirm.txt_TotalCost.Text
txt_OneOrder_ConsigneeKey1 = frm_OP_SDNConfirm.txt_OneOrder_ConsigneeKey1
txt_Zip1 = frm_OP_SDNConfirm.txt_Zip1
txt_OneOrder_FullName1 = frm_OP_SDNConfirm.txt_OneOrder_FullName1
txt_OneOrder_Address1 = frm_OP_SDNConfirm.txt_OneOrder_Address1
txt_Cash = frm_OP_SDNConfirm.txt_Cash
txt_ReceiveCash = frm_OP_SDNConfirm.txt_ReceiveCash
txt_Stairs = frm_OP_SDNConfirm.txt_Stairs
txtUrgent_Mark = frm_OP_SDNConfirm.txt_UrgentMark
txtReserve_Mark = frm_OP_SDNConfirm.txt_ReserveMark
txt_Cartype = frm_OP_SDNConfirm.txt_Cartype


'��ƭק��v��
If Val(txt_OneOrder_ArriveDate) > lngDueDate Then
    cmdSave.Enabled = True: cmdSaveExit.Enabled = True
Else
    cmdSave.Enabled = False: cmdSaveExit.Enabled = False
End If

'���д����O
Dim rsTmp As New ADODB.Recordset
rsTmp.Open "select distinct costkind = rtrim(isnull(costkind,'')) from trp17m (nolock) where storerkey = '" & txt_OneOrder_StorerKey.Text & "' and costkind <> 'notuse' order by rtrim(isnull(costkind,'')) ", cn
If Not rsTmp.EOF Then
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
        cboCostkind.AddItem rsTmp("costkind")
        rsTmp.MoveNext
    Loop
End If
rsTmp.Close

'�ɽдڤH ?
str_SQL = "update sdn01t set sdn01t.receiver = isnull(trp09m.receiver,trp09m.driver) from sdn01t join trp09m on trp09m.vehicle_id_no = sdn01t.c_vehicle_id_no where C_ROUTE_NO = '" & RTrim(txt_C_ROUTE_NO.Text) & "' and sdn01t.receiver is null "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'���дڤH ?
rsTmp.Open "select receiver from sdn01t (nolock) where C_ROUTE_NO = '" & RTrim(txt_C_ROUTE_NO.Text) & "' ", cn
txt_OneOrder_Receiver = rsTmp("receiver") & ""
rsTmp.Close

'�����f����
If RTrim(txt_OneOrder_FullName) = "�ըƹF�_��" Then
    txtArea = "��饫�[����"
ElseIf RTrim(txt_OneOrder_FullName) = "�ըƹF����" Then
    txtArea = "�x������ٰ�"
ElseIf RTrim(txt_OneOrder_FullName) = "�ըƹF�n��" Then
    txtArea = "�������j����"
ElseIf RTrim(txt_OneOrder_FullName) = "" Then
    txtArea = ""
Else
    If RTrim(txt_ZIP) <> "" Then
        rsTmp.Open "select description from trp02m (nolock) where ZIP = '" & RTrim(txt_ZIP) & "' ", cn
        txtArea = rsTmp("description") & ""
        rsTmp.Close
    End If
End If

'����f����
If RTrim(txt_OneOrder_FullName1) = "�ըƹF�_��" Then
    txtArea1 = "��饫�[����"
ElseIf RTrim(txt_OneOrder_FullName1) = "�ըƹF����" Then
    txtArea1 = "�x������ٰ�"
ElseIf RTrim(txt_OneOrder_FullName1) = "�ըƹF�n��" Then
    txtArea1 = "�������j����"
ElseIf RTrim(txt_OneOrder_FullName1) = "" Then
    txtArea1 = ""
Else
    rsTmp.Open "select description from trp02m (nolock) where ZIP = '" & RTrim(txt_Zip1) & "' ", cn
    txtArea1 = rsTmp("description") & ""
    rsTmp.Close
End If

'�d�߮ɥ��d mark by Eric 20141206
'str_SQL = "select " & _
'          "Urgent_Mark = isnull(Urgent_Mark,'') " & _
'          ",Reserve_Mark = isnull(Reserve_Mark,'') " & _
'          ",B_city = isnull(B_city,'') " & _
'          "from orders o(nolock) where o.storerkey = '" & txt_OneOrder_StorerKey.Text & "' and o.orderkey = '" & txt_OneOrder_OrderKey & "' and o.type <> '�R��' "
'rsTmp.Open str_SQL, cn
'
'If Not rsTmp.EOF Then
'    txtUrgent_Mark = rsTmp("Urgent_Mark") & ""
'    txtReserve_Mark = rsTmp("Reserve_Mark") & ""
'    txtB_City = rsTmp("B_city") & ""
'End If
'
'rsTmp.Close
'
'Set rsTmp = Nothing

'�L�p�O��Ʈ� , �I�p�O�Ѧ�
If rsMain.RecordCount = 0 Then Call cmdCost_Click: Label1.Visible = True

'�������Y
Me.Caption = "�B�O���@" & "_" & frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text
Call cmdAddnew_Click

Screen.MousePointer = 0

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing

End Sub

Private Sub txtCBM_DblClick()
rsMain("�ƶq") = txtCBM
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtCS_DblClick()
rsMain("�ƶq") = txtCS
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtCube_DblClick()
rsMain("�ƶq") = txtCube
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtEA_DblClick()
rsMain("�ƶq") = txtEA
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtOT_DblClick()
rsMain("�ƶq") = txtOT
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtWT_DblClick()
rsMain("�ƶq") = txtWT
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub
Private Sub txtCBM1_DblClick()
rsMain("�ƶq") = txtCBM1
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtCS1_DblClick()
rsMain("�ƶq") = txtCS1
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtCube1_DblClick()
rsMain("�ƶq") = txtCube1
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtEA1_DblClick()
rsMain("�ƶq") = txtEA1
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtOT1_DblClick()
rsMain("�ƶq") = txtOT1
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtWT1_DblClick()
rsMain("�ƶq") = txtWT1
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub
Private Sub txtCBM2_DblClick()
rsMain("�ƶq") = txtCBM2
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtCS2_DblClick()
rsMain("�ƶq") = txtCS2
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtCube2_DblClick()
rsMain("�ƶq") = txtCube2
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtEA2_DblClick()
rsMain("�ƶq") = txtEA2
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtOT2_DblClick()
rsMain("�ƶq") = txtOT2
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub

Private Sub txtWT2_DblClick()
rsMain("�ƶq") = txtWT2
rsMain("�����`��") = rsMain("�������") * rsMain("�ƶq"): rsMain("���I�`��") = rsMain("���I���") * rsMain("�ƶq")
End Sub
