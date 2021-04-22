VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Cost 
   BorderStyle     =   1  '單線固定
   Caption         =   "運費維護"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   14505
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txt_Cartype 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
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
      Alignment       =   1  '靠右對齊
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
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
      Alignment       =   1  '靠右對齊
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
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
      ToolTipText     =   "實際收現金額"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txt_Cash 
      Alignment       =   1  '靠右對齊
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
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
      ToolTipText     =   "下貨收現金額"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtB_City 
      Alignment       =   1  '靠右對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
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
      Alignment       =   1  '靠右對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
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
      Alignment       =   1  '靠右對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
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
      Alignment       =   1  '靠右對齊
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
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
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
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
      ToolTipText     =   "配送異常所衍生出之費用合計"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txt_TRPCost 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
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
      ToolTipText     =   "配送異常所衍生出之配送費"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox txt_SortingCost 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
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
      ToolTipText     =   "配送異常所衍生出之理貨費"
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdAbnormalCostUpdate 
      BackColor       =   &H00FF80FF&
      Caption         =   "異常費用更新"
      Height          =   375
      Left            =   6000
      TabIndex        =   72
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdReDeliveryCharge 
      BackColor       =   &H00FF80FF&
      Caption         =   "再配計價"
      Height          =   375
      Left            =   8280
      TabIndex        =   71
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtCBM2 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
      Caption         =   "存檔"
      Height          =   375
      Left            =   3360
      TabIndex        =   62
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtEA1 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
      Caption         =   "計費參考"
      Height          =   375
      Left            =   120
      TabIndex        =   46
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtWT 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
      Caption         =   "離開"
      Height          =   375
      Left            =   4440
      TabIndex        =   36
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "刪除"
      Height          =   375
      Left            =   2280
      TabIndex        =   35
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "新增"
      Height          =   375
      Left            =   1200
      TabIndex        =   34
      Top             =   3240
      Width           =   975
   End
   Begin VB.ComboBox cboCostCode 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "新細明體"
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
         Name            =   "新細明體"
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
      Caption         =   "存檔離開"
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
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "TMS單號"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "迄點"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "二次路編"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "請款人"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "起點"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "訂單號碼"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "訂單備註"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "到貨日期"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "運送車號"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "駕駛姓名"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "運輸公司"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "簽單狀態"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "訂單日期"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "訂單類別"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "貨主編號"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "路線編號"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "出車日期"
         BeginProperty Font 
            Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "計費類別"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "毛寶樓層"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "實收代收貨款"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "應收代收貨款"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "通路別"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "專車"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "急單"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "費用合計"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "異常配送費"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "異常理貨費"
      BeginProperty Font 
         Name            =   "新細明體"
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
      Caption         =   "未計費"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "到貨日x地址"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "路編x貨主"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "CBM"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "件"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "訂單"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "重"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "材"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "箱"
      BeginProperty Font 
         Name            =   "新細明體"
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
      BackStyle       =   0  '透明
      Caption         =   "個"
      BeginProperty Font 
         Name            =   "新細明體"
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

'資料排序
Recordset2Excel Me.Caption, rsMain

'..在此編輯EXCEL
If rsMain Is Nothing Then
Else
    With MyXlsApp
                
    End With
End If
Set MyXlsApp = Nothing

End Sub

Private Sub cboCostKind_Click()
rsMain("請款類別") = cboCostkind.Text
cboCostCode.Clear
Dim i As Integer, j As Integer

j = -1

'取計費代碼資料
Dim rsTmp As New ADODB.Recordset
rsTmp.Open "select CostCode = rtrim(isnull(CostCode,'')) from trp17m where costkind = '" & rsMain("請款類別") & "' and storerkey = '" & txt_OneOrder_StorerKey.Text & "' and costkind <> 'notuse' ", cn

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
If Len(RTrim(rsMain("備註"))) = 0 And rsMain("請款類別") = "堆高機費補貼" Then rsMain("備註") = "堆高機費補貼"

End Sub

Private Sub cboCostCode_Click()
rsMain("計費代碼") = cboCostCode.Text

'更新cboCostCode
Dim rsTmp As New ADODB.Recordset
rsTmp.Open "select receivable=isnull(receivable,0) ,UOM=rtrim(isnull(uom,'')), payable=isnull(payable,0) ,areastart= rtrim(areastart),areaend = rtrim(areaend) from trp17m where costcode = '" & rsMain("計費代碼") & "' and storerkey = '" & txt_OneOrder_StorerKey.Text & "' ", cn

If rsTmp.EOF Then Exit Sub
'rsMain("項次") = rsMain("計費代碼")
rsMain("單位") = rsTmp("UOM")
rsMain("應收單價") = rsTmp("receivable")
rsMain("應付單價") = rsTmp("payable")
rsMain("標準應收") = rsTmp("receivable")
rsMain("標準應付") = rsTmp("payable")
rsMain("應收總價") = rsTmp("receivable") * rsMain("數量")
rsMain("應付總價") = rsTmp("payable") * rsMain("數量")
rsMain("起點") = rsTmp("areastart") & ""
rsMain("迄點") = rsTmp("areaend") & ""

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
rsMain("編號") = rsMain.RecordCount
dgMain.AllowAddNew = False
End Sub

Private Sub cmdCost_Click()

On Error GoTo err_Handle
'簽單是否確認
If Len(RTrim(frm_OP_SDNConfirm.txt_OneOrder_Status)) = 0 Then MsgBox "尚未簽單確認！", 16, "計費": cmdSave.Enabled = False: cmdSaveExit.Enabled = False: Exit Sub

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select * from sdn05t (nolock) where sdn_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey) & "' ", cn
If Not tmp_Rs.EOF Then MsgBox "此訂單已有運費資料，系統不再計算運費!!", 16, "計費": Exit Sub

Screen.MousePointer = 11

'運費計算
cn.Execute "exec gs_Cost '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey) & "' ", RowsAffect, adExecuteNoRecords

'未出訂單計費數量更新為0
If RTrim(frm_OP_SDNConfirm.txt_OneOrder_Status) = "未出訂單" Then
    cn.Execute "update sdn05t set chargeqty = 0 , sumreceivable = 0 ,sumpayable = 0 where sdn_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_Status) & "' ", RowsAffect, adExecuteNoRecords
End If

'取計費資料
str_SQL = "select 請款類別 = rtrim(costkind) " & _
            ",計費代碼 = rtrim(costcode) " & _
            ",項次 = rtrim(sdn_name) " & _
            ",原因 = rtrim(rtrim(reason)) " & _
            ",應收單價 = receivable " & _
            ",應付單價 = payable " & _
            ",數量 = round(chargeqty,9) " & _
            ",單位 = rtrim(uom) " & _
            ",應收總價 = sumreceivable " & _
            ",應付總價 = sumpayable " & _
            ",議價 = premiam " & _
            ",計費車號 = isnull(vehicle_id_no,'') " & _
            ",備註 = rtrim(isnull(note,'')) " & _
            ",起點 = rtrim(isnull(areastart,'')) " & _
            ",迄點 = rtrim(isnull(areaend,'')) " & _
            ",標準應收 = stdreceivable " & _
            ",標準應付 = stdpayable " & _
            "from sdn05t (nolock) where sdn_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' order by sdn_name , costkind , receivable desc"
            
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

Call Replication_Recordset(tmp_Rs, rsMain)

If Not rsMain.EOF Then rsMain.MoveFirst
Set dgMain.DataSource = rsMain

'取欄位寬度
SetDataGridColWidth "運費維護", dgMain

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
If InStr(1, rsMain("備註"), "專車價") <> 0 Then
    If MsgBox("此單含有專車議價，確定刪除此筆計費明細資料?", vbOKCancel, Me.Caption) = vbOK Then
        rsMain.Delete
        MsgBox "刪除專車計價，請務必確保此趟專車計價正確！", 16, "注意"
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
'        DelRecord = MsgBox("應收代收貨款<>實收代收貨款，請確認資料是否正確?", vbQuestion + vbYesNo, "計費存檔")
'        If DelRecord = vbNo Then
'            Exit Sub
'        End If
'    End If
'End If

dgMain.Col = 0

cn.Execute "delete sdn05t where sdn_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' ", RowsAffect, adExecuteNoRecords
'If rsMain.RecordCount = 0 Then Call cmdExit_Click: Exit Sub

''更新請款人
'cn.Execute "update SDN01T set receiver = (select isnull(receiver,driver) from trp09m where vehicle_id_no = '" & mySplit(txt_OneOrder_VehicleID, "_", 0) & "') where c_route_no = '" & txt_C_ROUTE_NO & "'", RowsAffect, adExecuteNoRecords

rsMain.MoveFirst
Do While Not rsMain.EOF
    
    If Len(RTrim(rsMain("請款類別"))) > 0 Then
    
    If Round(rsMain("應收單價") * rsMain("數量"), 0) <> Round(rsMain("應收總價"), 0) Then MsgBox "應收單價 x 數量 <> 應收總價!!", 16, "注意"
    If Round(rsMain("應付單價") * rsMain("數量"), 0) <> Round(rsMain("應付總價"), 0) Then MsgBox "應付單價 x 數量 <> 應付總價!!", 16, "注意"
    If Len(RTrim(rsMain("起點"))) = 0 Or Len(RTrim(rsMain("迄點"))) = 0 Then MsgBox "起點或迄點不能空白!!", 16, "注意"
    If Len(RTrim(rsMain("計費代碼"))) = 0 Or Len(RTrim(rsMain("請款類別"))) = 0 Then MsgBox "請款類別或計費代碼不能空白!!", 16, "注意"
    
    str_SQL = "insert into sdn05t (c_route_no,uom,chargeqty,receivable,payable,premiam,reason,sumreceivable,sumpayable,areastart,areaend,note,sdn_name,sdn_no,costkind,costcode,stdreceivable,stdpayable,vehicle_id_no) " & _
    "values( '" & txt_C_ROUTE_NO.Text & "','" & rsMain("單位") & "','" & rsMain("數量") & "','" & rsMain("應收單價") & "','" & rsMain("應付單價") & "','" & rsMain("議價") & "','" & rsMain("原因") & "','" & rsMain("應收總價") & "','" & rsMain("應付總價") & "','" & rsMain("起點") & "','" & rsMain("迄點") & "','" & Trim(rsMain("備註")) & "','" & rsMain("項次") & _
    "','" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "','" & rsMain("請款類別") & "','" & rsMain("計費代碼") & "','" & rsMain("標準應收") & "','" & rsMain("標準應付") & "','" & rsMain("計費車號") & "' )"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
    End If
    rsMain.MoveNext
Loop

'應收分攤
cn.Execute "exec Es_ARnoDistribution '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' ", RowsAffect, adExecuteNoRecords
'更新回Orders的實收，如果是空白，則實收=應收，有值則實收=實收
'If RTrim(txt_ReceiveCash.Text) = "" Then
'    cn.Execute "update o set o.receiveCash = '" & Val(RTrim(txt_Cash.Text)) & "' from orders o join sdn02t s2 on o.orderkey = s2.receipt_no where s2.receipt_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' and o.type <> '刪單'", RowsAffect, adExecuteNoRecords
'    txt_ReceiveCash.Text = Val(RTrim(txt_Cash.Text))
'Else
'    cn.Execute "update o set o.receiveCash = '" & Val(RTrim(txt_ReceiveCash.Text)) & "' from orders o join sdn02t s2 on o.orderkey = s2.receipt_no where s2.receipt_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' and o.type <> '刪單'", RowsAffect, adExecuteNoRecords
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
'    If Len(RTrim(rsMain("請款類別"))) > 0 Then
'
'    If Int(rsMain("應收單價") * rsMain("數量")) <> Int(rsMain("應收總價")) Then MsgBox "注意!應收單價乘以數量不等於應收總價!!"
'    If Int(rsMain("應付單價") * rsMain("數量")) <> Int(rsMain("應付總價")) Then MsgBox "注意!應付單價乘以數量不等於應付總價!!"
'    If Len(RTrim(rsMain("起點"))) = 0 Or Len(RTrim(rsMain("迄點"))) = 0 Then MsgBox "注意!起點或迄點不能空白!!"
'
'        str_SQL = "insert into sdn05t (c_route_no,uom,chargeqty,receivable,payable,premiam,reason,sumreceivable,sumpayable,areastart,areaend,note,sdn_name,sdn_no,costkind,costcode,stdreceivable,stdpayable) " & _
'        "values( '" & txt_C_Route_NO.Text & "','" & rsMain("單位") & "','" & rsMain("數量") & "','" & rsMain("應收單價") & "','" & rsMain("應付單價") & "','" & rsMain("議價") & "','" & rsMain("原因") & "','" & rsMain("應收總價") & "','" & rsMain("應付總價") & "','" & rsMain("起點") & "','" & rsMain("迄點") & "','" & rsMain("備註") & "','" & rsMain("項次") & _
'        "','" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "','" & rsMain("請款類別") & "','" & rsMain("計費代碼") & "','" & rsMain("標準應收") & "','" & rsMain("標準應付") & "' )"
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'    End If
'    rsMain.MoveNext
'Loop
Call cmdSave_Click
Call cmdExit_Click

End Sub

Private Sub cmdReDeliveryCharge_Click()

'簽單是否確認
If Len(RTrim(frm_OP_SDNConfirm.txt_OneOrder_Status)) = 0 Then MsgBox "尚未簽單確認！", 16, "計費": cmdSave.Enabled = False: cmdSaveExit.Enabled = False: Exit Sub

If MsgBox("1.此單需先計費存檔" & vbCr & vbLf & "2.不加計請款類別含理貨費的計費代碼" & vbCr & vbLf & "3.不加計 RePalletIs 與 Cancel 計費代碼" & vbCr & vbLf & "4.備註開頭二次配送不加計" & vbCr & vbLf & "5.新增的計費，備註將註記二次配送", vbOKCancel, "計費說明") <> vbOK Then Exit Sub

'運費計算
cn.Execute "exec gs_ReDeliveryCharge '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' ", RowsAffect, adExecuteNoRecords

Call Form_Load
  
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

    If Len(dgMain.Columns(ColIndex).DataField) = 0 Then Exit Sub
    SaveSetting App.title, "運費維護" & "dgMain", dgMain.Columns(ColIndex).DataField, dgMain.Columns(ColIndex).Width
    
End Sub

Private Sub dgMain_DblClick()

    If dgMain.Columns(dgMain.Col).DataField = "起點" And Len(RTrim(dgMain.Columns(1))) > 0 Then dgMain = txtArea
    If dgMain.Columns(dgMain.Col).DataField = "迄點" And Len(RTrim(dgMain.Columns(1))) > 0 Then dgMain = txtArea1
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
                If cboCostkind.Left + cboCostkind.Width > .Left + .Width Then '如果欄位超出DataGrid的顯示範圍的處理
                    cboCostkind.Width = cboCostkind.Width + .Left + .Width - cboCostkind.Left - cboCostkind.Width
                End If
                cboCostkind.Text = rsMain("請款類別")  '更新Combo的值
        Else '如果用捲軸捲動出了DataGrid的顯示範圍，值會小於0
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
                If cboCostCode.Left + cboCostCode.Width > .Left + .Width Then '如果欄位超出DataGrid的顯示範圍的處理
                    cboCostCode.Width = cboCostCode.Width + .Left + .Width - cboCostCode.Left - cboCostCode.Width
                End If
                cboCostCode.Text = rsMain("計費代碼")  '更新Combo的值
        Else '如果用捲軸捲動出了DataGrid的顯示範圍，值會小於0
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

    '離開應收單價
    If dgMain.Col = 5 Then rsMain("應收總價") = dgMain * rsMain("數量")
    
    '離開應付單價
    If dgMain.Col = 6 Then rsMain("應付總價") = dgMain * rsMain("數量")
    
    '離開數量
    If dgMain.Col = 7 Then rsMain("應收總價") = rsMain("應收單價") * dgMain: rsMain("應付總價") = rsMain("應付單價") * dgMain
    
End If

End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

With dgMain
    cboCostkind.Visible = False: cboCostCode.Visible = False
    '
    'If dgMain.Col = 3 And cmdPickSave.Enabled = True Then ShowList
    
    '不允許移至特定欄位
    If .Col = 9 Or .Col = 10 Or .Col > 15 Then .Col = Abs(LastCol): Exit Sub

    '    If LastCol = 3 Then dgPick.Col = 5: Exit Sub
    '    If LastCol = 5 Then dgPick.Col = 2: Exit Sub
    '    dgPick.Col = IIf(LastCol = -1, 5, LastCol)
    'End If
    ''資料列是否變更
    'If LastRow = Empty Then Exit Sub
    
    '請款類別
    If .Col = 1 Then ShowCostKind
    
    '請款代碼
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
'If Len(RTrim(frm_OP_SDNConfirm.txt_ZIP.Text)) = 0 And (txt_Priority = "R" Or txt_Priority = "RC" Or txt_Priority = "A2B") Then MsgBox "此客戶無郵遞區號", 64, "注意!": Exit Sub
'If Len(RTrim(frm_OP_SDNConfirm.txt_Zip1.Text)) = 0 And (txt_Priority <> "R" Or txt_Priority <> "RC" Or txt_Priority <> "A2B") Then MsgBox "此客戶無郵遞區號", 64, "注意!": Exit Sub
Dim strConsigneeKey As String, strAddress As String
Dim Str_Skuxpack
Screen.MousePointer = 11

'取該訂單總個箱材重資料
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

'取路編x貨主出貨資料
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

'取到貨日x地址
str_SQL = "select CS = sum(case when sp.casecnt = 0 then 0 else (s3.Ship_qty /sp.casecnt) end) " & _
            ",Cube = sum(s3.ship_qty * sp.stdcube) " & _
            ",CBM = sum(s3.ship_qty * sp.stdcube) /35.315 " & _
            ",WGT = sum(s3.ship_qty * sp.stdgrosswgt) " & _
            ",EA = sum(s3.ship_qty) " & _
            ",Address = t1m.address " & _
            "from sdn03t s3 (nolock) join gv_skuxpack sp(nolock) on sp.storerkey = s3.storerkey and sp.sku = s3.product_no " & _
            "join sdn02t s2 (nolock) on s2.receipt_no = s3.receipt_no and s2.priority = '" & frm_OP_SDNConfirm.txt_Priority & "' " & _
            "join orders o (nolock) on o.orderkey = s2.c_receipt_no and o.type <> '刪單' and o.storerkey = '" & frm_OP_SDNConfirm.txt_OneOrder_StorerKey.Text & "' " & _
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

'取到貨日x地址總件數
str_SQL = "select OT = isnull((select sum(isnull(otqty,0)) from trp02t (nolock) join trp01m (nolock) on trp02t.storerkey = trp01m.storerkey and  case when rtrim(trp02t.priority) = 'A2B' then trp02t.bconsigneekey else trp02t.consigneekey end = trp01m.consigneekey and convert(char(8),trp02t.arrive_date,112) = '" & frm_OP_SDNConfirm.txt_OneOrder_ArriveDate & "' and trp01m.address = '" & tmp_Rs("Address") & "' and trp02t.storerkey = '" & frm_OP_SDNConfirm.txt_OneOrder_StorerKey.Text & "' ),0) " & _
          "+ isnull((select sum(isnull(otqty,0)) from ort02t (nolock) join trp01m (nolock) on ort02t.storerkey = trp01m.storerkey and case when rtrim(ort02t.priority) = 'A2B' then ort02t.bconsigneekey else ort02t.consigneekey end = trp01m.consigneekey and convert(char(8),ort02t.arrive_date,112) = '" & frm_OP_SDNConfirm.txt_OneOrder_ArriveDate & "' and trp01m.address = '" & tmp_Rs("Address") & "' and ort02t.storerkey = '" & frm_OP_SDNConfirm.txt_OneOrder_StorerKey.Text & "' ),0) "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

txtOT2 = tmp_Rs("OT")

If Val(txtEA2) > Val(txtEA1) Then txtEA2.BackColor = &HFF&

'取計費資料
str_SQL = "select 請款類別 = rtrim(costkind) " & _
            ",計費代碼 = rtrim(costcode) " & _
            ",項次 = rtrim(sdn_name) " & _
            ",原因 = rtrim(rtrim(reason)) " & _
            ",應收單價 = receivable " & _
            ",應付單價 = payable " & _
            ",數量 = round(chargeqty,9) " & _
            ",單位 = rtrim(uom) " & _
            ",應收總價 = sumreceivable " & _
            ",應付總價 = sumpayable " & _
            ",議價 = premiam " & _
            ",計費車號 = rtrim(isnull(vehicle_id_no,''))" & _
            ",備註 = rtrim(isnull(note,'')) " & _
            ",起點 = rtrim(isnull(areastart,'')) " & _
            ",迄點 = rtrim(isnull(areaend,'')) " & _
            ",標準應收 = stdreceivable " & _
            ",標準應付 = stdpayable " & _
            "from sdn05t (nolock) where sdn_no = '" & RTrim(frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text) & "' order by sdn_name , costkind , receivable desc"
            
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

Call Replication_Recordset(tmp_Rs, rsMain)

If Not rsMain.EOF Then rsMain.MoveFirst
Set dgMain.DataSource = rsMain

'取欄位寬度
SetDataGridColWidth "運費維護", dgMain

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


'資料修改權限
If Val(txt_OneOrder_ArriveDate) > lngDueDate Then
    cmdSave.Enabled = True: cmdSaveExit.Enabled = True
Else
    cmdSave.Enabled = False: cmdSaveExit.Enabled = False
End If

'取請款類別
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

'補請款人 ?
str_SQL = "update sdn01t set sdn01t.receiver = isnull(trp09m.receiver,trp09m.driver) from sdn01t join trp09m on trp09m.vehicle_id_no = sdn01t.c_vehicle_id_no where C_ROUTE_NO = '" & RTrim(txt_C_ROUTE_NO.Text) & "' and sdn01t.receiver is null "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'取請款人 ?
rsTmp.Open "select receiver from sdn01t (nolock) where C_ROUTE_NO = '" & RTrim(txt_C_ROUTE_NO.Text) & "' ", cn
txt_OneOrder_Receiver = rsTmp("receiver") & ""
rsTmp.Close

'取提貨縣市
If RTrim(txt_OneOrder_FullName) = "佰事達北倉" Then
    txtArea = "桃園市觀音區"
ElseIf RTrim(txt_OneOrder_FullName) = "佰事達中倉" Then
    txtArea = "台中市西屯區"
ElseIf RTrim(txt_OneOrder_FullName) = "佰事達南倉" Then
    txtArea = "高雄市大社區"
ElseIf RTrim(txt_OneOrder_FullName) = "" Then
    txtArea = ""
Else
    If RTrim(txt_ZIP) <> "" Then
        rsTmp.Open "select description from trp02m (nolock) where ZIP = '" & RTrim(txt_ZIP) & "' ", cn
        txtArea = rsTmp("description") & ""
        rsTmp.Close
    End If
End If

'取到貨縣市
If RTrim(txt_OneOrder_FullName1) = "佰事達北倉" Then
    txtArea1 = "桃園市觀音區"
ElseIf RTrim(txt_OneOrder_FullName1) = "佰事達中倉" Then
    txtArea1 = "台中市西屯區"
ElseIf RTrim(txt_OneOrder_FullName1) = "佰事達南倉" Then
    txtArea1 = "高雄市大社區"
ElseIf RTrim(txt_OneOrder_FullName1) = "" Then
    txtArea1 = ""
Else
    rsTmp.Open "select description from trp02m (nolock) where ZIP = '" & RTrim(txt_Zip1) & "' ", cn
    txtArea1 = rsTmp("description") & ""
    rsTmp.Close
End If

'查詢時先查 mark by Eric 20141206
'str_SQL = "select " & _
'          "Urgent_Mark = isnull(Urgent_Mark,'') " & _
'          ",Reserve_Mark = isnull(Reserve_Mark,'') " & _
'          ",B_city = isnull(B_city,'') " & _
'          "from orders o(nolock) where o.storerkey = '" & txt_OneOrder_StorerKey.Text & "' and o.orderkey = '" & txt_OneOrder_OrderKey & "' and o.type <> '刪單' "
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

'無計費資料時 , 點計費參考
If rsMain.RecordCount = 0 Then Call cmdCost_Click: Label1.Visible = True

'視窗抬頭
Me.Caption = "運費維護" & "_" & frm_OP_SDNConfirm.txt_OneOrder_OrderKey.Text
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
rsMain("數量") = txtCBM
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtCS_DblClick()
rsMain("數量") = txtCS
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtCube_DblClick()
rsMain("數量") = txtCube
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtEA_DblClick()
rsMain("數量") = txtEA
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtOT_DblClick()
rsMain("數量") = txtOT
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtWT_DblClick()
rsMain("數量") = txtWT
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub
Private Sub txtCBM1_DblClick()
rsMain("數量") = txtCBM1
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtCS1_DblClick()
rsMain("數量") = txtCS1
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtCube1_DblClick()
rsMain("數量") = txtCube1
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtEA1_DblClick()
rsMain("數量") = txtEA1
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtOT1_DblClick()
rsMain("數量") = txtOT1
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtWT1_DblClick()
rsMain("數量") = txtWT1
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub
Private Sub txtCBM2_DblClick()
rsMain("數量") = txtCBM2
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtCS2_DblClick()
rsMain("數量") = txtCS2
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtCube2_DblClick()
rsMain("數量") = txtCube2
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtEA2_DblClick()
rsMain("數量") = txtEA2
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtOT2_DblClick()
rsMain("數量") = txtOT2
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub

Private Sub txtWT2_DblClick()
rsMain("數量") = txtWT2
rsMain("應收總價") = rsMain("應收單價") * rsMain("數量"): rsMain("應付總價") = rsMain("應付單價") * rsMain("數量")
End Sub
