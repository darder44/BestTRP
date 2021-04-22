VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOrders 
   Caption         =   "訂單處理"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11055
   Icon            =   "frmOrder.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   11055
   WindowState     =   2  '最大化
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   0
      TabIndex        =   51
      Top             =   1320
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabMaxWidth     =   4410
      ForeColor       =   10485760
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "清單"
      TabPicture(0)   =   "frmOrder.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "dgOrder"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "訂單"
      TabPicture(1)   =   "frmOrder.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame4 
         BorderStyle     =   0  '沒有框線
         Height          =   495
         Left            =   1680
         TabIndex        =   84
         Top             =   600
         Width           =   9015
         Begin VB.CheckBox chkPreview 
            Caption         =   "預覽列印"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3600
            TabIndex        =   87
            Top             =   40
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "上一筆"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   1095
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "下一筆"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   1200
            TabIndex        =   18
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrintPick 
            Caption         =   "明細表"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   4920
            TabIndex        =   19
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrintShip 
            Caption         =   "出貨單"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   6240
            TabIndex        =   20
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdReset1 
            Caption         =   "重設"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   7680
            TabIndex        =   21
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2775
         Left            =   120
         TabIndex        =   61
         Top             =   390
         Width           =   1455
         Begin VB.CommandButton cmdSave 
            Caption         =   "存檔"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "取消"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   16
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "刪除"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "修改"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddNew 
            Caption         =   "新增"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid dgOrder 
         Height          =   2175
         Left            =   -74955
         TabIndex        =   11
         Top             =   450
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   20
         TabAction       =   1
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
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
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
         Height          =   3015
         Left            =   1605
         TabIndex        =   62
         Top             =   390
         Width           =   9375
         Begin VB.TextBox txtRouteKey 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7560
            MaxLength       =   10
            TabIndex        =   33
            Top             =   1800
            Width           =   1485
         End
         Begin VB.ComboBox cboOrderCar 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   345
            ItemData        =   "frmOrder.frx":0902
            Left            =   7560
            List            =   "frmOrder.frx":0904
            TabIndex        =   36
            Top             =   2160
            Width           =   1485
         End
         Begin VB.TextBox txtWeight 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7560
            MaxLength       =   11
            TabIndex        =   24
            Text            =   "0"
            Top             =   720
            Width           =   1485
         End
         Begin VB.TextBox txtStatus 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4800
            MaxLength       =   3
            TabIndex        =   23
            Top             =   720
            Width           =   1485
         End
         Begin VB.ComboBox cboUsepallet 
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   345
            ItemData        =   "frmOrder.frx":0906
            Left            =   1680
            List            =   "frmOrder.frx":0908
            Style           =   2  '單純下拉式
            TabIndex        =   34
            Top             =   2160
            Width           =   1485
         End
         Begin VB.TextBox txtOrderkey 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   22
            Top             =   720
            Width           =   1485
         End
         Begin VB.TextBox txtAccountkey 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            MaxLength       =   9
            TabIndex        =   25
            Top             =   1080
            Width           =   1485
         End
         Begin VB.TextBox txtAccount 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            MaxLength       =   9
            TabIndex        =   28
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox txtCompanykey 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4800
            MaxLength       =   9
            TabIndex        =   26
            Top             =   1080
            Width           =   1485
         End
         Begin VB.TextBox txtCompany 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4800
            MaxLength       =   9
            TabIndex        =   29
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox txtTel 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            MaxLength       =   16
            TabIndex        =   37
            Top             =   2520
            Width           =   2325
         End
         Begin VB.TextBox txtAddress 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5160
            MaxLength       =   37
            TabIndex        =   38
            Top             =   2520
            Width           =   4605
         End
         Begin VB.TextBox txtOrderdate 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   31
            Top             =   1800
            Width           =   1485
         End
         Begin VB.TextBox txtDeliverydate 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4800
            MaxLength       =   10
            TabIndex        =   32
            Top             =   1800
            Width           =   1485
         End
         Begin VB.TextBox txtTransferkey 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7560
            MaxLength       =   6
            TabIndex        =   27
            Top             =   1080
            Width           =   1485
         End
         Begin VB.TextBox txtTransfer 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7560
            MaxLength       =   10
            TabIndex        =   30
            Top             =   1440
            Width           =   1485
         End
         Begin VB.TextBox txtExternOrderkey 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4800
            MaxLength       =   17
            TabIndex        =   35
            Top             =   2160
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "路線編號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   20
            Left            =   6360
            TabIndex        =   83
            Top             =   1860
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "重量"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   16
            Left            =   6360
            TabIndex        =   80
            Top             =   780
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "訂單狀態"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   14
            Left            =   3240
            TabIndex        =   77
            Top             =   780
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "訂單號碼"
            BeginProperty Font 
               Name            =   "新細明體"
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
            TabIndex        =   76
            Top             =   780
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "帳款客戶代號"
            BeginProperty Font 
               Name            =   "新細明體"
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
            TabIndex        =   75
            Top             =   1140
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "帳款客戶名稱"
            BeginProperty Font 
               Name            =   "新細明體"
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
            TabIndex        =   74
            Top             =   1500
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "送貨客戶代號"
            BeginProperty Font 
               Name            =   "新細明體"
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
            Left            =   3240
            TabIndex        =   73
            Top             =   1140
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "送貨客戶名稱"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   6
            Left            =   3240
            TabIndex        =   72
            Top             =   1500
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶電話"
            BeginProperty Font 
               Name            =   "新細明體"
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
            Left            =   120
            TabIndex        =   71
            Top             =   2580
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "送貨地址"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   8
            Left            =   4080
            TabIndex        =   70
            Top             =   2580
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "排出日期"
            BeginProperty Font 
               Name            =   "新細明體"
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
            TabIndex        =   69
            Top             =   1860
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "預出日期"
            BeginProperty Font 
               Name            =   "新細明體"
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
            Left            =   3240
            TabIndex        =   68
            Top             =   1860
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車商代號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   11
            Left            =   6360
            TabIndex        =   67
            Top             =   1140
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車商名稱"
            BeginProperty Font 
               Name            =   "新細明體"
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
            Left            =   6360
            TabIndex        =   66
            Top             =   1500
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運送車號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   13
            Left            =   6360
            TabIndex        =   65
            Top             =   2220
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "棧板使用"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   15
            Left            =   120
            TabIndex        =   64
            Top             =   2220
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶單號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   225
            Index           =   17
            Left            =   3240
            TabIndex        =   63
            Top             =   2220
            Width           =   960
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   1335
      Left            =   0
      TabIndex        =   53
      Top             =   -60
      Width           =   9975
      Begin VB.ComboBox cboTransferKey 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         ItemData        =   "frmOrder.frx":090A
         Left            =   5640
         List            =   "frmOrder.frx":090C
         Style           =   2  '單純下拉式
         TabIndex        =   85
         Top             =   180
         Width           =   1485
      End
      Begin VB.ComboBox cboCar 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         ItemData        =   "frmOrder.frx":090E
         Left            =   5640
         List            =   "frmOrder.frx":0910
         Style           =   2  '單純下拉式
         TabIndex        =   7
         Top             =   900
         Width           =   1485
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "離開"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   8640
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "重設"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   8640
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7320
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cboStatus 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         ItemData        =   "frmOrder.frx":0912
         Left            =   5640
         List            =   "frmOrder.frx":0914
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   540
         Width           =   1485
      End
      Begin VB.TextBox txt3E 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
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
         TabIndex        =   5
         Top             =   900
         Width           =   1485
      End
      Begin VB.TextBox txt3S 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
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
         TabIndex        =   4
         Top             =   900
         Width           =   1485
      End
      Begin VB.TextBox txt2S 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
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
         TabIndex        =   2
         Top             =   540
         Width           =   1485
      End
      Begin VB.TextBox txt2E 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
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
         TabIndex        =   3
         Top             =   540
         Width           =   1485
      End
      Begin VB.TextBox txt1E 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   1
         Top             =   180
         Width           =   1485
      End
      Begin VB.TextBox txt1S 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
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
         TabIndex        =   0
         Top             =   180
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "車商代號"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Left            =   4560
         TabIndex        =   86
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "車號"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Left            =   4560
         TabIndex        =   82
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "訂單狀態"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Left            =   4560
         TabIndex        =   60
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   18
         Left            =   2655
         TabIndex        =   59
         Top             =   960
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "新細明體"
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
         TabIndex        =   58
         Top             =   600
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "排出日期"
         BeginProperty Font 
            Name            =   "新細明體"
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
         TabIndex        =   57
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "預出日期"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   19
         Left            =   120
         TabIndex        =   56
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "出貨單號"
         BeginProperty Font 
            Name            =   "新細明體"
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
         TabIndex        =   55
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Left            =   2655
         TabIndex        =   54
         Top             =   240
         Width           =   360
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2895
      Left            =   0
      TabIndex        =   52
      Top             =   5280
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabMaxWidth     =   4410
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "訂單品項"
      TabPicture(0)   =   "frmOrder.frx":0916
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dgSku"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "揀貨"
      TabPicture(1)   =   "frmOrder.frx":0932
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboPick"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).Control(2)=   "dgPick"
      Tab(1).ControlCount=   3
      Begin VB.ComboBox cboPick 
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
         Left            =   -69000
         Style           =   2  '單純下拉式
         TabIndex        =   81
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame6 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   79
         Top             =   360
         Width           =   1455
         Begin VB.CommandButton cmdPickSave 
            Caption         =   "存檔"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   48
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdPickCancel 
            Caption         =   "取消"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   49
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdPickDelete 
            Caption         =   "刪除"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   47
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdPickEdit 
            Caption         =   "修改"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdPickAddnew 
            Caption         =   "新增"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2775
         Left            =   120
         TabIndex        =   78
         Top             =   360
         Width           =   1455
         Begin VB.CommandButton cmdSkuAddnew 
            Caption         =   "新增"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdSkuEdit 
            Caption         =   "修改"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdSkuDelete 
            Caption         =   "刪除"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   41
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdSkuCancel 
            Caption         =   "取消"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   43
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdSkuSave 
            Caption         =   "存檔"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   120
            TabIndex        =   42
            Top             =   1680
            Width           =   1215
         End
      End
      Begin MSDataGridLib.DataGrid dgSku 
         Height          =   2175
         Left            =   1605
         TabIndex        =   44
         Top             =   450
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   20
         TabAction       =   1
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
      Begin MSDataGridLib.DataGrid dgPick 
         Height          =   2175
         Left            =   -73395
         TabIndex        =   50
         Top             =   450
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3836
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   20
         TabAction       =   1
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
   End
End
Attribute VB_Name = "frmOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsOrder As ADODB.Recordset, rsSku As ADODB.Recordset, rsPick As ADODB.Recordset
Private intColumnIndex As Integer
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cboCar_GotFocus()

'取出車號
cboCar.Clear
strSql = "select distinct car From orders order by car"
Set rsTmp = New ADODB.Recordset

rsTmp.Open strSql, cnAccess ', adOpenStatic, adLockOptimistic
If Not rsTmp.EOF Then
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If IsNull(rsTmp("car")) = False Then cboCar.AddItem rsTmp("car")
            rsTmp.MoveNext
        Loop
End If
End Sub

Private Sub cboPick_LostFocus()
cboPick.Visible = False
End Sub

Private Sub cboTransferKey_GotFocus()

strSql = "select distinct TransferKey From orders order by TransferKey"
Set rsTmp = New ADODB.Recordset
cboTransferKey.Clear
rsTmp.Open strSql, cnAccess ', adOpenStatic, adLockOptimistic
If Not rsTmp.EOF Then
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If IsNull(rsTmp("TransferKey")) = False Then cboTransferKey.AddItem rsTmp("TransferKey")
            rsTmp.MoveNext
        Loop
End If

End Sub
Private Sub cboOrderCar_GotFocus()

strSql = "select distinct car From orders order by car"
Set rsTmp = New ADODB.Recordset
cboOrderCar.Clear
rsTmp.Open strSql, cnAccess ', adOpenStatic, adLockOptimistic
If Not rsTmp.EOF Then
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If IsNull(rsTmp("car")) = False Then cboOrderCar.AddItem rsTmp("car")
            rsTmp.MoveNext
        Loop
End If

End Sub
Private Sub cmdPickAddNew_Click()
Dim i As Integer

If rsSku("訂單數量") = rsSku("揀貨數量") Then MsgBox "該品項揀貨作業完成！", vbOKOnly + vbInformation, "揀貨維護": Exit Sub
SSTab2_Click (0)

With rsPick
    i = 1
    If .EOF = False Then .MoveLast: i = .Fields("揀貨項次") + 1
    .AddNew
    .Fields("揀貨項次") = i
    .Fields("棧板種類") = "PTA1W110140"
    .Fields("單位") = rsSku.Fields("單位")
    .Fields("新增日期") = Now()
    .Fields("修改日期") = Now()
End With

dgPick.AllowUpdate = True
cmdPickSave.Enabled = True: cmdPickCancel.Enabled = True
cmdPickDelete.Enabled = False: cmdPickEdit.Enabled = False: cmdPickAddnew.Enabled = False
intPickqty = IIf(IsNull(rsPick.Fields("揀貨數量")), 0, rsPick.Fields("揀貨數量"))
dgPick.Col = 1: dgPick.SetFocus
intPickRow = dgPick.Row
intSkuRow = dgSku.Row
intLastCol = dgPick.Col

End Sub
Private Sub cmdPickEdit_Click()

If txtStatus.Text = "9" Then MsgBox "已完成訂單無法修改揀貨明細!!", vbInformation ': Exit Sub

dgPick.AllowUpdate = True
cmdPickSave.Enabled = True: cmdPickCancel.Enabled = True
cmdPickDelete.Enabled = False: cmdPickEdit.Enabled = False: cmdPickAddnew.Enabled = False
dgPick.Col = 1: dgPick.SetFocus
intPickqty = rsPick.Fields("揀貨數量")
intPickRow = dgPick.Row
intSkuRow = dgSku.Row
intLastCol = dgPick.Col

End Sub
Private Sub cmdPickDelete_Click()
On Error GoTo err_Handle
Dim confirm As Integer

If txtStatus.Text = "9" Then MsgBox "已完成訂單無法刪除揀貨明細!!", vbInformation: Exit Sub
confirm = MsgBox("確定刪除?", vbQuestion + vbOKCancel, "揀貨明細維護")
If confirm <> 1 Then Exit Sub

rsPick.Delete

'strSql = "delete from pickdetail where orderkey = '" & rsOrder("出貨單號") & "' and linenumber = " & rsSku.Fields("項次") & " and sku = '" & rsSku.Fields("產品編號") & "' and picklinenumber = " & rsPick.Fields("揀貨項次") & " "
'cnAccess.BeginTrans
'    cnAccess.Execute strSql, RowsAffect, adExecuteNoRecords
'cnAccess.CommitTrans

'更新訂單明細資料
Call Update
cmdPickAddnew.SetFocus

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdPickSave_Click()
On Error GoTo err_Handle

If Len(RTrim(rsPick.Fields("批號") & " ")) = 0 Then MsgBox "請輸入批號!!", vbOKOnly + vbInformation, "揀貨明細維護": dgPick.Col = 1: dgPick.SetFocus: Exit Sub
If Len(RTrim(rsPick.Fields("棧板編號") & " ")) = 0 Then MsgBox "請輸入棧板編號!!", vbOKOnly + vbInformation, "揀貨明細維護": dgPick.Col = 2: dgPick.SetFocus: Exit Sub
If IsNull(rsPick.Fields("揀貨數量")) = True Or rsPick.Fields("揀貨數量") = 0 Then MsgBox "請輸入揀貨數量!!", vbOKOnly + vbInformation, "揀貨明細維護": dgPick.Col = 5: dgPick.SetFocus: Exit Sub

'揀貨量不得大於訂單量
If rsSku.Fields("訂單數量") - rsSku.Fields("揀貨數量") < rsPick.Fields("揀貨數量") - intPickqty Then
    MsgBox "揀貨量( " & rsSku.Fields("揀貨數量") + rsPick.Fields("揀貨數量") - intPickqty & " )不得大於訂單量( " & rsSku.Fields("訂單數量") & " )!!", vbOKOnly + vbInformation, "揀貨明細維護"
    dgPick.Col = 5
    dgPick.SetFocus
    dgPick.Text = rsSku.Fields("訂單數量") - rsSku.Fields("揀貨數量")
    Exit Sub
End If

'檢查是否重複
Set rsTmp = New ADODB.Recordset
With rsTmp
    .CursorLocation = adUseServer
    strSql = "select * from pickdetail where orderkey = '" & rsOrder.Fields("出貨單號") & "' and linenumber = " & rsSku.Fields("項次") & " and sku = '" & rsSku.Fields("產品編號") & "' and picklinenumber = " & rsPick.Fields("揀貨項次") & " "
    .Open strSql, cnAccess, adOpenStatic, adLockOptimistic
        
    If .EOF Then
        
        '新增資料
            .AddNew
            .Fields("orderkey") = rsOrder.Fields("出貨單號")
            .Fields("linenumber") = rsSku.Fields("項次")
            .Fields("sku") = UCase(rsSku.Fields("產品編號"))
            .Fields("picklinenumber") = rsPick.Fields("揀貨項次")
            .Fields("lot") = rsPick.Fields("批號")
            .Fields("palletid") = UCase(rsPick.Fields("棧板編號"))
            .Fields("UOM") = rsPick.Fields("單位")
            .Fields("pickqty") = rsPick.Fields("揀貨數量")
            .Fields("pallet") = UCase(rsPick.Fields("棧板種類"))
            .Fields("Adddate") = Now()
            .Fields("Editdate") = Now()
            .Update
    
    Else
     
        '修改資料
            .Fields("orderkey") = rsOrder.Fields("出貨單號")
            .Fields("linenumber") = rsSku.Fields("項次")
            .Fields("sku") = UCase(rsSku.Fields("產品編號"))
            .Fields("picklinenumber") = rsPick.Fields("揀貨項次")
            .Fields("lot") = UCase(rsPick.Fields("批號"))
            .Fields("palletid") = UCase(rsPick.Fields("棧板編號"))
            .Fields("UOM") = rsPick.Fields("單位")
            .Fields("pickqty") = rsPick.Fields("揀貨數量")
            .Fields("pallet") = UCase(rsPick.Fields("棧板種類"))
            .Fields("Editdate") = Now()
            .Update
    
    End If
End With

cmdPickAddnew.Enabled = True: cmdPickEdit.Enabled = True: cmdPickDelete.Enabled = True: dgPick.AllowUpdate = False: cmdPickSave.Enabled = False: cmdPickCancel.Enabled = False
Call Update
cmdPickAddnew.SetFocus
dgPick.AllowUpdate = False

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdPickCancel_Click()

cmdPickSave.Enabled = False: cmdPickCancel.Enabled = False
SSTab2_Click (0)
'cmdPickAddnew.SetFocus
dgPick.AllowUpdate = False

End Sub
Private Sub cmdSkuAddNew_Click()
On Error GoTo err_Handle
Dim i As Integer

Call dgOrder_RowColChange(1, 1)
i = 1
If rsSku.EOF = False Then rsSku.MoveLast: i = rsSku.Fields("項次") + 1
rsSku.AddNew
rsSku.Fields("項次") = i
rsSku.Fields("出貨原因") = "99"
rsSku.Fields("單位") = "DZ"
rsSku.Fields("訂單數量") = 0
rsSku.Fields("揀貨數量") = 0
rsSku.Fields("新增日期") = Now()
rsSku.Fields("修改日期") = Now()
dgSku.AllowUpdate = True
cmdSkuSave.Enabled = True: cmdSkuCancel.Enabled = True
cmdSkuDelete.Enabled = False: cmdSkuEdit.Enabled = False: cmdSkuAddnew.Enabled = False
dgSku.Col = 1: dgSku.SetFocus
intSkuRow = dgSku.Row
intLastCol = dgPick.Col

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdSkuEdit_Click()

If txtStatus.Text = "9" Then MsgBox "已完成訂單無法修改!!", vbInformation: Exit Sub

dgSku.AllowUpdate = True
cmdSkuSave.Enabled = True: cmdSkuCancel.Enabled = True
cmdSkuDelete.Enabled = False: cmdSkuEdit.Enabled = False: cmdSkuAddnew.Enabled = False: Frame6.Enabled = False
dgSku.Col = 1: dgSku.SetFocus
intSkuRow = dgSku.Row
intLastCol = dgPick.Col

End Sub
Private Sub cmdSkuDelete_Click()
On Error GoTo err_Handle
Dim confirm As Integer

If txtStatus.Text = "9" Then MsgBox "已完成訂單無法刪除!!", vbInformation: Exit Sub
confirm = MsgBox("確定刪除?", vbQuestion + vbOKCancel, "訂單明細維護")
If confirm <> 1 Then Exit Sub

strSql = "delete from orderdetail where orderkey = '" & rsOrder("出貨單號") & "' and linenumber = " & rsSku.Fields("項次") & " and sku = '" & rsSku.Fields("產品編號") & "'"
cnAccess.BeginTrans
    cnAccess.Execute strSql, RowsAffect, adExecuteNoRecords
    cnAccess.Execute "delete from pickdetail where orderkey = '" & rsOrder("出貨單號") & "' and linenumber = " & rsSku.Fields("項次") & " and sku = '" & rsSku.Fields("產品編號") & "'", RowsAffect, adExecuteNoRecords
cnAccess.CommitTrans
If dgSku.Row > 0 Then dgSku.Row = dgSku.Row - 1

Call Update
Call dgOrder_RowColChange(1, 1)

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdSkuSave_Click()
On Error GoTo err_Handle

If Len(RTrim(rsSku.Fields("產品編號") & " ")) = 0 Then MsgBox "請輸入產品編號!!", vbOKOnly + vbInformation, "訂單明細維護": dgSku.Col = 2: dgSku.SetFocus: Exit Sub
If IsNull(rsSku.Fields("訂單數量")) = True Or rsSku.Fields("訂單數量") < 1 Then MsgBox "請輸入訂單數量!!", vbOKOnly + vbInformation, "訂單明細維護": dgSku.Col = 5: dgSku.SetFocus: Exit Sub

'揀貨量不得大於訂單量
If IIf(IsNull(rsSku.Fields("訂單數量")), 0, rsSku.Fields("訂單數量")) < IIf(IsNull(rsSku.Fields("揀貨數量")), 0, rsSku.Fields("揀貨數量")) Then
    MsgBox "揀貨量不得大於訂單量!!", vbOKOnly + vbInformation, "訂單明細更新"
    dgSku.Col = 4
    dgSku.SetFocus
    Exit Sub
End If

'檢查是否重複
Set rsTmp = New ADODB.Recordset
With rsTmp
.CursorLocation = adUseServer
strSql = "select * from orderdetail where orderkey = '" & rsOrder.Fields("出貨單號") & "'and linenumber =" & rsSku.Fields("項次") & " and sku ='" & rsSku.Fields("產品編號") & "' "
.Open strSql, cnAccess, adOpenForwardOnly, adLockOptimistic
    
    If .EOF Then
    
        '新增資料
            .AddNew
            .Fields("orderkey") = rsOrder("出貨單號")
            .Fields("linenumber") = rsSku("項次")
            .Fields("shiptype") = rsSku("出貨原因")
            .Fields("sku") = UCase(rsSku("產品編號"))
            .Fields("descr") = UCase(rsSku("產品名稱"))
            .Fields("UOM") = rsSku("單位")
            .Fields("openqty") = rsSku("訂單數量")
            .Fields("pickqty") = rsSku("揀貨數量")
            .Fields("notes") = rsSku("備註")
            .Fields("Adddate") = Now()
            .Fields("Editdate") = Now()
            .Update
       
    Else
        '修改資料
            .Fields("orderkey") = rsOrder("出貨單號")
            .Fields("linenumber") = rsSku("項次")
            .Fields("shiptype") = rsSku("出貨原因")
            .Fields("sku") = UCase(rsSku("產品編號"))
            .Fields("descr") = UCase(rsSku("產品名稱"))
            .Fields("UOM") = rsSku("單位")
            .Fields("openqty") = rsSku("訂單數量")
            .Fields("pickqty") = rsSku("揀貨數量")
            .Fields("notes") = rsSku("備註")
            .Fields("Editdate") = Now()
            .Update
     
    End If
End With
Frame6.Enabled = True
cmdSkuAddnew.Enabled = True: cmdSkuEdit.Enabled = True: cmdSkuDelete.Enabled = True: dgSku.AllowUpdate = False: cmdSkuSave.Enabled = False: cmdSkuCancel.Enabled = False
Call Update
cmdSkuAddnew.SetFocus
dgSku.AllowUpdate = False

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdSkuCancel_Click()
On Error GoTo err_Handle

cmdSkuAddnew.Enabled = True: cmdSkuEdit.Enabled = True: cmdSkuDelete.Enabled = True: cmdSkuSave.Enabled = False: cmdSkuCancel.Enabled = False
dgSku.AllowUpdate = False
Call dgOrder_RowColChange(1, 1)

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub Update()
On Error GoTo err_Handle

If rsSku.EOF = False Then
'更新訂單明細(Pickqty)
Set rsTmp = New ADODB.Recordset
rsTmp.CursorLocation = adUseServer
strSql = "select sum (pickqty) as sumpickqty from pickdetail where orderkey = '" & rsOrder.Fields("出貨單號") & "' and linenumber = " & rsSku.Fields("項次") & " and sku = '" & rsSku.Fields("產品編號") & "'"
rsTmp.Open strSql, cnAccess

    cnAccess.BeginTrans
        strSql = "update orderdetail set pickqty = " & IIf(IsNull(rsTmp.Fields("sumpickqty")), 0, rsTmp.Fields("sumpickqty")) & ", editdate = '" & Now() & "' where orderkey = '" & rsOrder.Fields("出貨單號") & "' and linenumber = " & rsSku.Fields("項次") & " and sku = '" & rsSku.Fields("產品編號") & "'"
        cnAccess.Execute strSql, RowsAffect, adExecuteNoRecords
        rsSku("揀貨數量") = IIf(IsNull(rsTmp.Fields("sumpickqty")), 0, rsTmp.Fields("sumpickqty")): rsSku.Fields("修改日期") = Now() ': rsSku.Update
    cnAccess.CommitTrans
End If

'更新訂單狀態(Status)

Set rsTmp = New ADODB.Recordset
rsTmp.CursorLocation = adUseServer
strSql = "select * from orderdetail where orderkey = '" & rsOrder.Fields("出貨單號") & "' and openqty > pickqty and openqty > 0 "
rsTmp.Open strSql, cnAccess, adOpenForwardOnly, adLockReadOnly
    
'cnAccess.BeginTrans
If rsTmp.EOF = True And rsSku.EOF = False Then
'    cnAccess.Execute "update orders set status = 9 , editdate = '" & Now() & "' where orderkey = '" & rsOrder.Fields("出貨單號") & "'", RowsAffect, adExecuteNoRecords
    rsOrder.Fields("訂單狀態") = 9: txtStatus = 9: rsOrder.Fields("修改日期") = Now(): rsOrder.Update
Else
'    cnAccess.Execute "update orders set status = 0 , editdate = '" & Now() & "' where orderkey = '" & rsOrder.Fields("出貨單號") & "'", RowsAffect, adExecuteNoRecords
    rsOrder.Fields("訂單狀態") = 0: txtStatus = 0: rsOrder.Fields("修改日期") = Now(): rsOrder.Update
End If

'cnAccess.CommitTrans

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdPrintPick_Click()
On Error GoTo err_Handle
Dim confirm As Integer

If rsSku.EOF = True Then MsgBox "無明細資料可供列印!!", vbOKOnly + vbInformation, "報表列印": Exit Sub
confirm = MsgBox("確定列印訂單明細表？", vbQuestion + vbOKCancel, "報表列印")
If confirm <> 1 Then Exit Sub

strSql = "insert into pickinglist " & _
         "select o.orderkey , o.accountkey , o.account , o.companykey , o.company , o.transferkey , o.transfer , o.orderdate , o.deliverydate , o.car , o.usepallet , o.tel , o.address , od.shiptype , od.linenumber , od.sku , od.descr , od.UOM , o.weight, od.openqty , od.notes " & _
         "from orders o , orderdetail od " & _
         "Where o.orderkey = od.orderkey " & _
         "and o.orderkey = '" & rsOrder.Fields("出貨單號") & "' " & _
         "order by od.linenumber "
'寫入列印資料
cnAccess.BeginTrans
cnAccess.Execute "delete from pickinglist", RowsAffect, adExecuteNoRecords
cnAccess.Execute strSql, RowsAffect, adExecuteNoRecords
cnAccess.CommitTrans

'開啟報表列印
Set MSAccessAP = New Access.Application
With MSAccessAP
.Visible = False
.OpenCurrentDatabase (App.Path & "\" & App.Title & ".mdb")

If chkPreview.Value = vbChecked Then
   '預覽列印
    .Visible = True
    .DoCmd.OpenReport "PickingList", acViewPreview
   
Else
   '直接列印至印表機
    .Visible = False
    .DoCmd.OpenReport "PickingList", acViewNormal
    .CloseCurrentDatabase
    .Quit

End If

End With

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdPrintShip_Click()
On Error GoTo err_Handle
Dim confirm As Integer, h As Integer, i As Integer, j As Integer, k As Integer

If rsSku.EOF = True Then MsgBox "無資料可供列印!!", vbOKOnly + vbInformation, "報表列印": Exit Sub
confirm = MsgBox("確定列印出貨單？", vbQuestion + vbOKCancel, "報表列印")
If confirm <> 1 Then Exit Sub
If rsOrder.Fields("訂單狀態") <> 9 Then confirm = MsgBox("該訂單尚未完成，是否列印？", vbQuestion + vbOKCancel, "報表列印")
If confirm <> 1 Then Exit Sub

'取出列印資料
strSql = "select o.orderkey , o.accountkey , o.account , o.companykey , o.company , o.transferkey , o.transfer , o.car , o.usepallet , o.tel , o.address , o.orderdate , o.deliverydate , o.externorderkey , o.editdate " & _
        ", od.linenumber " & _
        ", od.shiptype " & _
        ", od.sku " & _
        ", od.descr " & _
        ", od.UOM " & _
        ", od.pickqty " & _
        ", 0 as eaqty  " & _
        ", o.weight " & _
        ", od.notes " & _
        "from orders o , orderdetail od " & _
        "Where o.orderkey = od.orderkey " & _
        "and o.orderkey = '" & rsOrder.Fields("出貨單號") & "' " & _
        "and od.pickqty > 0 " & _
        "order by od.linenumber "

Dim rs_Tmpshipped As New ADODB.Recordset
rs_Tmpshipped.CursorLocation = adUseServer
rs_Tmpshipped.Open strSql, cnAccess ', adOpenKeyset, adLockOptimistic
rs_Tmpshipped.MoveFirst

'開啟報表資料檔
Set rsTmp = New ADODB.Recordset
rsTmp.CursorLocation = adUseServer
rsTmp.Open "shippedlist", cnAccess, adOpenKeyset, adLockOptimistic

'寫入列印資料
cnAccess.BeginTrans
    cnAccess.Execute "delete * from shippedlist", RowsAffect, adExecuteNoRecords

    Do While Not rs_Tmpshipped.EOF
        rsTmp.AddNew
        
        i = 0
        
        For i = 0 To rs_Tmpshipped.Fields.Count - 1
            
            'DZ/EA判斷
            If i = 19 And rs_Tmpshipped.Fields(i) = "EA" Then
                rsTmp.Fields(i) = rs_Tmpshipped.Fields(i)
                i = i + 1
                rsTmp.Fields(i + 1) = rs_Tmpshipped.Fields(i)
                i = i + 1
            Else
                rsTmp.Fields(i) = rs_Tmpshipped.Fields(i)
            End If
        
        Next
          
        rs_Tmpshipped.MoveNext
    rsTmp.Update
    Loop
cnAccess.CommitTrans

'取出棧板資料
Set rsTmp = New ADODB.Recordset
rsTmp.CursorLocation = adUseServer

strSql = "select pd.sku , pd.lot , pd.palletid , pd.pallet " & _
        "from pickdetail pd " & _
        "where orderkey = '" & rsOrder("出貨單號") & "' " & _
        "order by sku , lot , palletid"

rsTmp.Open strSql, cnAccess, adOpenKeyset, adLockOptimistic
rsTmp.MoveFirst
Dim arr_tmp(10, 40) As String

h = 0: i = 0: j = 1
arr_tmp(h, 0) = rsTmp.RecordCount '棧板使用總數
arr_tmp(h, 1 + i) = rsTmp("sku")
arr_tmp(h, 2 + i) = rsTmp("lot")
arr_tmp(h, 3 + i) = rsTmp("palletid")

rsTmp.MoveNext
Do While Not rsTmp.EOF

If (i + 3) Mod 39 = 0 Then h = h + 1: i = 0: arr_tmp(h, 1 + i) = rsTmp("sku"): arr_tmp(h, 2 + i) = rsTmp("lot"): arr_tmp(h, 3 + i) = ""

'同品項同批號，板號置於同一列
If rsTmp("sku") = arr_tmp(h, 1 + i) And rsTmp("lot") = arr_tmp(h, 2 + i) Then
        
        j = j + 1
        
        If j > 10 Then
        
                GoTo NextLine '單行板號超過10筆換行
                
        Else
                arr_tmp(h, 3 + i) = arr_tmp(h, 3 + i) & " " & rsTmp("palletid")
                arr_tmp(h, 1 + i) = rsTmp("sku")
                arr_tmp(h, 2 + i) = rsTmp("lot")
                arr_tmp(h, 3 + i) = arr_tmp(h, 3 + i)

        End If
        
Else
        '資料確認
NextLine:
        i = i + 3: j = 1
        arr_tmp(h, 1 + i) = rsTmp("sku")
        arr_tmp(h, 2 + i) = rsTmp("lot")
        arr_tmp(h, 3 + i) = rsTmp("palletid")
End If
    
    rsTmp.MoveNext
Loop

'棧板種類數量計算

strSql = "select distinct (select count(*) from pickdetail pd where pd.orderkey = o.orderkey and pd.pallet = 'PTA1W110140' ) as PLA " & _
        ",(select count(*) from pickdetail pd where pd.orderkey = o.orderkey and pd.pallet = 'PTB1P110110' ) as PLB " & _
        ",(select count(*) from pickdetail pd where pd.orderkey = o.orderkey and pd.pallet = 'PTD1W110110' ) as PLC " & _
        ",(select count(*) from pickdetail pd where pd.orderkey = o.orderkey and pd.pallet = 'PTE1W110110' ) as PLD " & _
        ",(select count(*) from pickdetail pd where pd.orderkey = o.orderkey and pd.pallet = 'PTG1W100120' ) as PLE " & _
        ",(select count(*) from pickdetail pd where pd.orderkey = o.orderkey and pd.pallet = 'PTK1W' ) as PLF " & _
        ",(select count(*) from pickdetail pd where pd.orderkey = o.orderkey and pd.pallet = 'NONE' ) as NONE " & _
        "from pickdetail o " & _
        "where o.orderkey = '" & rsOrder("出貨單號") & "' "

Set rsTmp = New ADODB.Recordset
rsTmp.CursorLocation = 3
rsTmp.Open strSql, cnAccess

'填入資料
'棧板資料
k = h
For h = 0 To h
Set MSAccessAP = New Access.Application

With MSAccessAP
    .Visible = False
    .OpenCurrentDatabase (App.Path & "\" & App.Title & ".mdb")
    .DoCmd.OpenReport "Shippedlist", acViewDesign
'    .Reports("Shippedlist").[Label00].Caption = arr_tmp(h,0)
    .Reports("Shippedlist").[Label01].Caption = arr_tmp(h, 1): .Reports("Shippedlist").[Label02].Caption = arr_tmp(h, 2): .Reports("Shippedlist").[Label03].Caption = arr_tmp(h, 3)
    .Reports("Shippedlist").[Label04].Caption = arr_tmp(h, 4): .Reports("Shippedlist").[Label05].Caption = arr_tmp(h, 5): .Reports("Shippedlist").[Label06].Caption = arr_tmp(h, 6)
    .Reports("Shippedlist").[Label07].Caption = arr_tmp(h, 7): .Reports("Shippedlist").[Label08].Caption = arr_tmp(h, 8): .Reports("Shippedlist").[Label09].Caption = arr_tmp(h, 9)
    .Reports("Shippedlist").[Label10].Caption = arr_tmp(h, 10): .Reports("Shippedlist").[Label11].Caption = arr_tmp(h, 11): .Reports("Shippedlist").[Label12].Caption = arr_tmp(h, 12)
    .Reports("Shippedlist").[Label13].Caption = arr_tmp(h, 13): .Reports("Shippedlist").[Label14].Caption = arr_tmp(h, 14): .Reports("Shippedlist").[Label15].Caption = arr_tmp(h, 15)
    .Reports("Shippedlist").[Label16].Caption = arr_tmp(h, 16): .Reports("Shippedlist").[Label17].Caption = arr_tmp(h, 17): .Reports("Shippedlist").[Label18].Caption = arr_tmp(h, 18)
    .Reports("Shippedlist").[Label19].Caption = arr_tmp(h, 19): .Reports("Shippedlist").[Label20].Caption = arr_tmp(h, 20): .Reports("Shippedlist").[Label21].Caption = arr_tmp(h, 21)
    .Reports("Shippedlist").[Label22].Caption = arr_tmp(h, 22): .Reports("Shippedlist").[Label23].Caption = arr_tmp(h, 23): .Reports("Shippedlist").[Label24].Caption = arr_tmp(h, 24)
    .Reports("Shippedlist").[Label25].Caption = arr_tmp(h, 25): .Reports("Shippedlist").[Label26].Caption = arr_tmp(h, 26): .Reports("Shippedlist").[Label27].Caption = arr_tmp(h, 27)
    .Reports("Shippedlist").[Label28].Caption = arr_tmp(h, 28): .Reports("Shippedlist").[Label29].Caption = arr_tmp(h, 29): .Reports("Shippedlist").[Label30].Caption = arr_tmp(h, 30)
    .Reports("Shippedlist").[Label31].Caption = arr_tmp(h, 31): .Reports("Shippedlist").[Label32].Caption = arr_tmp(h, 32): .Reports("Shippedlist").[Label33].Caption = arr_tmp(h, 33)
    .Reports("Shippedlist").[Label34].Caption = arr_tmp(h, 34): .Reports("Shippedlist").[Label35].Caption = arr_tmp(h, 35): .Reports("Shippedlist").[Label36].Caption = arr_tmp(h, 36)
    .Reports("Shippedlist").[Label37].Caption = arr_tmp(h, 37): .Reports("Shippedlist").[Label38].Caption = arr_tmp(h, 38): .Reports("Shippedlist").[Label39].Caption = arr_tmp(h, 39)
    .Reports("Shippedlist").[Label52].Caption = h + 1 & " / " & k + 1 '頁次
    
If h > 0 Then GoTo PrintReport ' 第二張以後不列印棧板總計資料

'棧板資料
    Dim Plttype(6) As String, PltCount(6) As String
    Plttype(0) = "維他露木板*": PltCount(0) = rsTmp("PLA")
    Plttype(1) = "久津塑板*": PltCount(1) = rsTmp("PLB")
    Plttype(2) = "全台木板*": PltCount(2) = rsTmp("PLC")
    Plttype(3) = "捷盟木板*": PltCount(3) = rsTmp("PLD")
    Plttype(4) = "家福木板*": PltCount(4) = rsTmp("PLE")
    Plttype(5) = "金酒木板*": PltCount(5) = rsTmp("PLF")
    Plttype(6) = "": PltCount(6) = "0"
    .Reports("Shippedlist").[Label00].Caption = arr_tmp(0, 0) - rsTmp("NONE") '扣除未使用棧板數
    
    i = 0
a:
    If i = 6 Then GoTo PrintReport
    If PltCount(i) > 0 Then
    .Reports("Shippedlist").[Label40].Caption = Plttype(i): .Reports("Shippedlist").[Label41].Caption = PltCount(i)
    Else
    i = i + 1
    GoTo a
    End If
    i = i + 1
B:
    If i = 6 Then GoTo PrintReport
    If PltCount(i) > 0 Then
    .Reports("Shippedlist").[Label42].Caption = Plttype(i): .Reports("Shippedlist").[Label43].Caption = PltCount(i)
    Else
    i = i + 1
    GoTo B
    End If
    i = i + 1
C:
    If i = 6 Then GoTo PrintReport
    If PltCount(i) > 0 Then
    .Reports("Shippedlist").[Label44].Caption = Plttype(i): .Reports("Shippedlist").[Label45].Caption = PltCount(i)
    Else
    i = i + 1
    GoTo C
    End If
    i = i + 1
    
D:
    If i = 6 Then GoTo PrintReport
    If PltCount(i) > 0 Then
    .Reports("Shippedlist").[Label46].Caption = Plttype(i): .Reports("Shippedlist").[Label47].Caption = PltCount(i)
    Else
    i = i + 1
    GoTo D
    End If
    i = i + 1

E:
    If i = 6 Then GoTo PrintReport
    If PltCount(i) > 0 Then
    .Reports("Shippedlist").[Label48].Caption = Plttype(i): .Reports("Shippedlist").[Label49].Caption = PltCount(i)
    Else
    i = i + 1
    GoTo E
    End If
    i = i + 1
    
F:
    If i = 6 Then GoTo PrintReport
    If PltCount(i) > 0 Then
    .Reports("Shippedlist").[Label50].Caption = Plttype(i): .Reports("Shippedlist").[Label51].Caption = PltCount(i)
    Else
    i = i + 1
    GoTo F
    End If
    
PrintReport:
'開啟報表列印
If chkPreview.Value = vbChecked Then

   '預覽列印
    .Visible = True
    .DoCmd.OpenReport "Shippedlist", acViewPreview
   
Else
   '直接列印至印表機
    .Visible = False
    .DoCmd.OpenReport "Shippedlist", acViewNormal
    .CloseCurrentDatabase
    .Quit

End If
  
End With

Next h

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdPrevious_Click()

cmdNext.Enabled = True
rsOrder.MovePrevious
SSTab1_Click (0)

End Sub
Private Sub cmdNext_Click()

cmdPrevious.Enabled = True
rsOrder.MoveNext
SSTab1_Click (0)

End Sub
Private Sub cmdAddNew_Click()

SSTab1.TabCaption(1) = "訂單新增"
txtOrderkey.Text = "": txtOrderkey.BackColor = &HFFFFFF
txtStatus.Text = 0
txtAccountkey.Text = "": txtAccountkey.BackColor = &HFFFFFF
txtAccount.Text = "": txtAccount.BackColor = &HFFFFFF
txtCompanykey.Text = "": txtCompanykey.BackColor = &HFFFFFF
txtCompany.Text = "": txtCompany.BackColor = &HFFFFFF
txtTel.Text = "": txtTel.BackColor = &HFFFFFF
txtAddress.Text = "": txtAddress.BackColor = &HFFFFFF
txtOrderdate.Text = Format(Date, "YYYYMMDD"): txtOrderdate.BackColor = &HFFFFFF
txtDeliverydate.Text = Format(Date, "YYYYMMDD"): txtDeliverydate.BackColor = &HFFFFFF
txtTransferkey.Text = "": txtTransferkey.BackColor = &HFFFFFF
txtTransfer.Text = "": txtTransfer.BackColor = &HFFFFFF
cboOrderCar.Text = ""
txtExternOrderkey.Text = "": txtExternOrderkey.BackColor = &HFFFFFF
txtWeight.Text = "": txtWeight.BackColor = &HFFFFFF
txtRouteKey.Text = "": txtRouteKey.BackColor = &HFFFFFF
Frame2.Enabled = True
cmdSave.Enabled = True: cmdCancel.Enabled = True ': cmdReset1.Enabled = True
cmdDelete.Enabled = False: cmdEdit.Enabled = False: chkPreview.Enabled = False: cmdPrintPick.Enabled = False: cmdPrintShip.Enabled = False
cmdPrevious.Enabled = False
cmdNext.Enabled = False
cmdAddNew.Enabled = False
txtOrderkey.SetFocus
SSTab2.Enabled = False

End Sub
Private Sub cmdEdit_Click()

If txtStatus.Text = "9" Then MsgBox "已完成訂單確定修改?", vbQuestion ': Exit Sub

txtAccountkey.BackColor = &HFFFFFF
txtAccount.BackColor = &HFFFFFF
txtCompanykey.BackColor = &HFFFFFF
txtCompany.BackColor = &HFFFFFF
txtTel.BackColor = &HFFFFFF
txtAddress.BackColor = &HFFFFFF
txtOrderdate.BackColor = &HFFFFFF
txtDeliverydate.BackColor = &HFFFFFF
txtTransferkey.BackColor = &HFFFFFF
txtTransfer.BackColor = &HFFFFFF
txtExternOrderkey.BackColor = &HFFFFFF
txtWeight.BackColor = &HFFFFFF
txtRouteKey.BackColor = &HFFFFFF
Frame2.Enabled = True
txtOrderkey.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
'cmdReset1.Enabled = True
cmdDelete.Enabled = False: cmdEdit.Enabled = False: chkPreview.Enabled = False: cmdPrintPick.Enabled = False: cmdPrintShip.Enabled = False
cmdPrevious.Enabled = False
cmdNext.Enabled = False
cmdAddNew.Enabled = False
txtAccountkey.SetFocus

End Sub
Private Sub cmdDelete_Click()
On Error GoTo err_Handle
Dim confirm As Integer

If txtStatus.Text = "9" Then MsgBox "正在刪除已完成訂單?", vbQuestion, Me.Caption  ': Exit Sub

confirm = MsgBox("訂單刪除!!", vbQuestion + vbOKCancel, Me.Caption)
If confirm <> 1 Then Exit Sub

cnAccess.BeginTrans
    cnAccess.Execute "delete from orderdetail where orderkey = '" & rsOrder("出貨單號") & "'", RowsAffect, adExecuteNoRecords
    cnAccess.Execute "delete from pickdetail where orderkey = '" & rsOrder("出貨單號") & "'", RowsAffect, adExecuteNoRecords
cnAccess.CommitTrans
rsOrder.Delete

'更新表格資料
SSTab1_Click (0)
SSTab1.TabCaption(0) = "清單" & "( " & rsOrder.RecordCount & " 筆)"

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdSave_Click()
On Error GoTo err_Handle
Dim i As Integer

If Len(txtOrderkey.Text) = 0 Then MsgBox "請輸入出貨單號!!", vbOKOnly + vbInformation, "訂單新增": txtOrderkey.SetFocus: Exit Sub

'判斷dgorder是否null
If rsOrder Is Nothing Then
strSql = "select " & _
        "orderkey as 出貨單號 " & _
        ",status as 訂單狀態 " & _
        ",Accountkey as 帳款客戶代號 " & _
        ",Account as 帳款客戶名稱 " & _
        ",Companykey as 送貨客戶代號 " & _
        ",Company as 送貨客戶名稱 " & _
        ",Tel as 客戶電話 " & _
        ",Address as 客戶地址 " & _
        ",OrderDate as 排出日期 " & _
        ",DeliveryDate as 預出日期 " & _
        ",TransferKey as 車商代號 " & _
        ",Transfer as 車商名稱 " & _
        ",Car as 車號 " & _
        ",UsePallet as 棧板使用 " & _
        ",Weight as 重量 " & _
        ",ExternOrderkey as 客戶單號 " & _
        ",Routekey as 路線編號 " & _
        ",Adddate as 新增日期 " & _
        ",Editdate as 修改日期 " & _
        "from orders " & _
        "where 1 = 2 "
        
Set rsOrder = New ADODB.Recordset
rsOrder.CursorLocation = 3
rsOrder.Open strSql, cnAccess, adOpenForwardOnly, adLockPessimistic
Set dgOrder.DataSource = rsOrder

'標題行
For i = 0 To rsOrder.Fields.Count - 1
dgOrder.Columns(i).Caption = rsOrder.Fields(i).Name
Next
  
With dgOrder
    .ColumnHeaders = True        '標題行顯示
    .RowHeight = 300
    .Columns(0).Width = 1100:       .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 500:    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1000:    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1000:    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000:    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 1000:    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 1000:    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 1000:    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 1000:    .Columns(8).Alignment = dbgCenter
    .Columns(9).Width = 1000:    .Columns(9).Alignment = dbgCenter
    .Columns(10).Width = 1000:    .Columns(10).Alignment = dbgCenter
    .Columns(11).Width = 1000:    .Columns(11).Alignment = dbgCenter
    .Columns(12).Width = 1000:    .Columns(12).Alignment = dbgCenter
    .Columns(13).Width = 500:    .Columns(13).Alignment = dbgCenter
    .Columns(14).Width = 1000:    .Columns(14).Alignment = dbgRight
    .Columns(15).Width = 1600:    .Columns(15).Alignment = dbgLeft
    .Columns(16).Width = 1200:    .Columns(16).Alignment = dbgCenter
    .Columns(17).Width = 1600:    .Columns(17).Alignment = dbgCenter
    .Columns(18).Width = 1600:    .Columns(18).Alignment = dbgCenter

End With

End If

'新增或修改
Set rsTmp = New ADODB.Recordset
    rsTmp.CursorLocation = adUseServer
    strSql = "select * from orders where orderkey = '" & txtOrderkey.Text & "' "
    rsTmp.Open strSql, cnAccess, adOpenStatic, adLockOptimistic
With rsOrder
    If rsTmp.EOF = True Then
       
        '新增資料
            .AddNew
            .Fields("出貨單號") = UCase(txtOrderkey.Text)
            .Fields("訂單狀態") = txtStatus.Text
            .Fields("帳款客戶代號") = UCase(txtAccountkey.Text) & ""
            .Fields("帳款客戶名稱") = txtAccount.Text & ""
            .Fields("送貨客戶代號") = UCase(txtCompanykey.Text)
            .Fields("送貨客戶名稱") = txtCompany.Text
            .Fields("客戶電話") = txtTel.Text
            .Fields("客戶地址") = txtAddress.Text
            .Fields("車商代號") = UCase(txtTransferkey.Text)
            .Fields("車商名稱") = txtTransfer.Text
            .Fields("車號") = UCase(cboOrderCar.Text)
            .Fields("排出日期") = txtOrderdate.Text
            .Fields("預出日期") = txtDeliverydate.Text
            .Fields("棧板使用") = cboUsepallet.Text
            .Fields("客戶單號") = UCase(txtExternOrderkey.Text)
            .Fields("重量") = IIf(Len(txtWeight.Text) = 0, 0, txtWeight.Text)
            .Fields("路線編號") = txtRouteKey.Text
            .Fields("新增日期") = Now()
            .Fields("修改日期") = Now()
            .Update
         Else
     
        '訂單號碼是否重複
             If txtOrderkey.Enabled = True Then MsgBox "訂單號碼重複!!", vbOKOnly + vbInformation, "訂單新增": txtOrderkey.SetFocus: Exit Sub
        '修改資料
            .Fields("出貨單號") = UCase(txtOrderkey.Text)
            .Fields("帳款客戶代號") = UCase(txtAccountkey.Text)
            .Fields("帳款客戶名稱") = txtAccount.Text
            .Fields("送貨客戶代號") = UCase(txtCompanykey.Text)
            .Fields("送貨客戶名稱") = txtCompany.Text
            .Fields("客戶電話") = txtTel.Text
            .Fields("客戶地址") = txtAddress.Text
            .Fields("車商代號") = UCase(txtTransferkey.Text)
            .Fields("車商名稱") = txtTransfer.Text
            .Fields("車號") = UCase(cboOrderCar.Text)
            .Fields("排出日期") = txtOrderdate.Text
            .Fields("預出日期") = txtDeliverydate.Text
            .Fields("棧板使用") = cboUsepallet.Text
            .Fields("客戶單號") = UCase(txtExternOrderkey.Text)
            .Fields("重量") = IIf(Len(txtWeight.Text) = 0, 0, txtWeight.Text)
            .Fields("路線編號") = txtRouteKey.Text
            .Fields("修改日期") = Now()
            .Update
    
    End If
End With

SSTab1.TabCaption(0) = "清單" & "( " & rsOrder.RecordCount & " 筆)"
SSTab2.Enabled = True
cmdCancel_Click

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdCancel_Click()

txtOrderkey.BackColor = &HE0E0E0: txtOrderkey.Enabled = True
txtAccountkey.BackColor = &HE0E0E0
txtAccount.BackColor = &HE0E0E0
txtCompanykey.BackColor = &HE0E0E0
txtCompany.BackColor = &HE0E0E0
txtTel.BackColor = &HE0E0E0
txtAddress.BackColor = &HE0E0E0
txtOrderdate.BackColor = &HE0E0E0
txtDeliverydate.BackColor = &HE0E0E0
txtTransferkey.BackColor = &HE0E0E0
txtTransfer.BackColor = &HE0E0E0
cboOrderCar.ListIndex = -1
cboUsepallet.ListIndex = 0
txtExternOrderkey.BackColor = &HE0E0E0
txtWeight.BackColor = &HE0E0E0
txtRouteKey.BackColor = &HE0E0E0
cmdDelete.Enabled = False
cmdAddNew.Enabled = True
cmdEdit.Enabled = False
cmdCancel.Enabled = False
cmdSave.Enabled = False
'cmdReset1.Enabled = False
dgSku.AllowUpdate = False

SSTab1_Click (0)
Call dgOrder_RowColChange(1, 1)

End Sub
Private Sub cmdReset1_Click()

'重設
If txtOrderkey.Enabled Then txtOrderkey.Text = ""
chkPreview.Value = 0
txtStatus.Text = ""
txtAccountkey.Text = ""
txtAccount.Text = ""
txtCompanykey.Text = ""
txtCompany.Text = ""
txtTel.Text = ""
txtAddress.Text = ""
txtOrderdate.Text = ""
txtDeliverydate.Text = ""
txtTransferkey.Text = ""
txtTransfer.Text = ""
cboOrderCar.Text = ""
cboUsepallet.ListIndex = 0
txtExternOrderkey.Text = ""
txtWeight.Text = ""
txtRouteKey.Text = ""

End Sub
Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Dim chc_Orderby As String, chc_Orderkey As String, chc_Status As String, chc_Orderdate As String, chc_Deliverydate As String, chc_Car As String, chc_TransferKey As String
Dim i As Integer

SSTab1.Tab = 0

strSql = "select " & _
        "orderkey as 出貨單號 " & _
        ",status as 訂單狀態 " & _
        ",Accountkey as 帳款客戶代號 " & _
        ",Account as 帳款客戶名稱 " & _
        ",Companykey as 送貨客戶代號 " & _
        ",Company as 送貨客戶名稱 " & _
        ",Tel as 客戶電話 " & _
        ",Address as 客戶地址 " & _
        ",OrderDate as 排出日期 " & _
        ",DeliveryDate as 預出日期 " & _
        ",TransferKey as 車商代號 " & _
        ",Transfer as 車商名稱 " & _
        ",Car as 車號 " & _
        ",UsePallet as 棧板使用 " & _
        ",Weight as 重量 " & _
        ",ExternOrderkey as 客戶單號 " & _
        ",Routekey as 路線編號 " & _
        ",Adddate as 新增日期 " & _
        ",Editdate as 修改日期 " & _
        "from orders "
        
chc_Orderby = "order by orderkey"

'出貨單號
chc_Orderkey = ""
If Len(txt1S.Text) > 0 And Len(txt1E.Text) > 0 Then
   chc_Orderkey = "and orderkey between '" & txt1S.Text & "' and '" & txt1E.Text & "' "
ElseIf Len(txt1S.Text) > 0 And Len(txt1E.Text) = 0 Then
   chc_Orderkey = "and orderkey = '" & txt1S.Text & "' "
ElseIf Len(txt1S.Text) = 0 And Len(txt1E.Text) > 0 Then
   chc_Orderkey = "and orderkey = '" & txt1E.Text & "' "
End If

'訂單狀態
chc_Status = ""
Select Case Left(cboStatus.Text, 1)
        Case 0
            chc_Status = "and status = 0 "
        Case 9
            chc_Status = "and status = 9 "
End Select

'車商代號
chc_TransferKey = ""
If Len(cboTransferKey.Text) > 0 Then chc_TransferKey = "and TransferKey = '" & cboTransferKey.Text & "' "

'車號
chc_Car = ""
If Len(cboCar.Text) > 0 Then chc_Car = "and car = '" & cboCar.Text & "' "

'排出日期
chc_Orderdate = ""
If Len(txt2S.Text) > 0 And Len(txt2E.Text) > 0 Then
   chc_Orderdate = "and Orderdate between '" & txt2S.Text & "' and '" & txt2E.Text & "' "
ElseIf Len(txt2S.Text) > 0 And Len(txt2E.Text) = 0 Then
   chc_Orderdate = "and Orderdate = '" & txt2S.Text & "' "
ElseIf Len(txt2S.Text) = 0 And Len(txt2E.Text) > 0 Then
   chc_Orderdate = "and Orderdate = '" & txt2E.Text & "' "
End If

'預出日期
chc_Deliverydate = ""
If Len(txt3S.Text) > 0 And Len(txt3E.Text) > 0 Then
   chc_Deliverydate = "and DeliveryDate between '" & txt3S.Text & "' and '" & txt3E.Text & "' "
ElseIf Len(txt3S.Text) > 0 And Len(txt3E.Text) = 0 Then
   chc_Deliverydate = "and DeliveryDate = '" & txt3S.Text & "' "
ElseIf Len(txt3S.Text) = 0 And Len(txt3E.Text) > 0 Then
   chc_Deliverydate = "and DeliveryDate = '" & txt3E.Text & "' "
End If

'組合字串
strSql = strSql & "where 1 = 1 " & chc_Orderkey & chc_Orderdate & chc_Status & chc_Deliverydate & chc_Car & chc_TransferKey & chc_Orderby

Set rsOrder = New ADODB.Recordset
rsOrder.CursorLocation = 3
rsOrder.Open strSql, cnAccess, adOpenForwardOnly, adLockPessimistic
Set dgOrder.DataSource = rsOrder

'標題行
For i = 0 To rsOrder.Fields.Count - 1
dgOrder.Columns(i).Caption = rsOrder.Fields(i).Name
Next
  
With dgOrder
    .ColumnHeaders = True        '標題行顯示
    .RowHeight = 300
    .Columns(0).Width = 1100:       .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 500:    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1000:    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1000:    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000:    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 1000:    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 1500:    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 4000:    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 1000:    .Columns(8).Alignment = dbgCenter
    .Columns(9).Width = 1000:    .Columns(9).Alignment = dbgCenter
    .Columns(10).Width = 1000:    .Columns(10).Alignment = dbgCenter
    .Columns(11).Width = 1000:    .Columns(11).Alignment = dbgCenter
    .Columns(12).Width = 1000:    .Columns(12).Alignment = dbgCenter
    .Columns(13).Width = 500:    .Columns(13).Alignment = dbgCenter
    .Columns(14).Width = 1000:    .Columns(14).Alignment = dbgRight
    .Columns(15).Width = 1600:    .Columns(15).Alignment = dbgLeft
    .Columns(16).Width = 1200:    .Columns(16).Alignment = dbgCenter
    .Columns(17).Width = 1600:    .Columns(17).Alignment = dbgCenter
    .Columns(18).Width = 1600:    .Columns(18).Alignment = dbgCenter

End With
SSTab1.TabCaption(0) = "清單" & "( " & rsOrder.RecordCount & " 筆)"
If rsOrder.EOF = True Then Screen.MousePointer = 0: MsgBox "無資料可顯示！", vbOKOnly + vbInformation, Me.Caption: cmdCancel_Click: SSTab1.TabCaption(1) = "訂單": SSTab2.TabCaption(1) = "揀貨": cmdPickAddnew.Enabled = False: Exit Sub
cmdNext.Enabled = True
SSTab1.TabCaption(1) = "訂單" & "( " & rsOrder.Fields("出貨單號") & " )"

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub dgOrder_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle
Dim i As Integer

'訂單新增編輯狀態無作用
If cmdSave.Enabled = True Then Exit Sub

'上一筆與下一筆可用
'    cmdNext.Enabled = True:  cmdPrevious.Enabled = True
    If rsOrder.BOF = True Then cmdPrevious.Enabled = False: rsOrder.MoveFirst: cmdNext.Enabled = True: Exit Sub
    If rsOrder.EOF = True Then cmdNext.Enabled = False: rsOrder.MoveLast: cmdPrevious.Enabled = True: Exit Sub
    If rsOrder.RecordCount = 1 Then cmdNext.Enabled = False: cmdPrevious.Enabled = False

SSTab2.Enabled = True
Frame2.Enabled = False

'是否有資料
If dgOrder.Row = -1 Then
    txtOrderkey = "": cmdReset1_Click: cmdDelete.Enabled = False: cmdEdit.Enabled = False: cmdPrintShip.Enabled = False: chkPreview.Enabled = False: chkPreview.Enabled = False: cmdPrintPick.Enabled = False
    If rsOrder.BOF = True Then cmdPrevious.Enabled = False
    If rsOrder.EOF = True Then cmdNext.Enabled = False
    SSTab1.TabCaption(1) = "訂單": SSTab2.TabCaption(1) = "訂單": cmdSkuAddnew.Enabled = False: cmdSkuEdit.Enabled = False: cmdSkuDelete.Enabled = False: Set dgSku.DataSource = Nothing: Set dgSku.DataSource = Nothing
    Exit Sub
Else
    SSTab1.TabCaption(1) = "訂單" & "( " & rsOrder.Fields("出貨單號") & " )"
    cmdSkuAddnew.Enabled = True
End If

Screen.MousePointer = 11

cmdDelete.Enabled = True: chkPreview.Enabled = True: cmdPrintPick.Enabled = True: cmdPrintShip.Enabled = True: cmdEdit.Enabled = True: cmdDelete.Enabled = True

'更新訂單明細
strSql = "select " & _
        "od.linenumber as 項次 " & _
        ",od.shiptype as 出貨原因 " & _
        ",od.sku as 產品編號 " & _
        ",od.descr as 產品名稱 " & _
        ",od.UOM as 單位 " & _
        ",od.openqty as 訂單數量 " & _
        ",od.pickqty as 揀貨數量 " & _
        ",od.Notes as 備註 " & _
        ",od.adddate as 新增日期 " & _
        ",od.editdate as 修改日期 " & _
        "from orderdetail od " & _
        "where od.orderkey = '" & rsOrder.Fields("出貨單號") & "' " & _
        "order by od.linenumber"
        
Set rsSku = New ADODB.Recordset
rsSku.CursorLocation = 3
rsSku.Open strSql, cnAccess, adOpenKeyset, adLockPessimistic

Set dgSku.DataSource = rsSku

'欄位格式
With dgSku
    .ColumnHeaders = True        '標題行顯示
    .RowHeight = 300
    .Columns(0).Width = 500:       .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 500:    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1200:    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 3000:    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 500:    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 900:    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 900:    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 5000:    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 1600:    .Columns(8).Alignment = dbgCenter
    .Columns(9).Width = 1600:    .Columns(9).Alignment = dbgCenter
End With
    
cmdDelete.Enabled = True: cmdEdit.Enabled = True: chkPreview.Enabled = True: cmdPrintPick.Enabled = True: cmdPrintShip.Enabled = True: cmdSkuEdit.Enabled = True: cmdSkuDelete.Enabled = True: cmdSkuSave.Enabled = False: cmdSkuCancel.Enabled = False
If rsSku.EOF Then SSTab2.TabCaption(1) = "揀貨": cmdPickAddnew.Enabled = False: cmdSkuDelete.Enabled = False: cmdSkuEdit.Enabled = False
SSTab2.Tab = 0: dgPick.AllowUpdate = False
cmdPickCancel_Click

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub dgSku_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim confirm As Integer

If cmdSkuSave.Enabled = True And dgSku.ColContaining(X) = -1 And dgSku.RowContaining(Y) <> intSkuRow Then
confirm = MsgBox("是否存檔!!", vbQuestion + vbOKCancel)
If confirm = 1 Then cmdSkuSave_Click
intSkuRow = intSkuRow - 1
intLastCol = intLastCol + 1
cmdSkuCancel_Click

End If
End Sub

Private Sub dgPick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim confirm As Integer

If cmdPickSave.Enabled = True And dgPick.ColContaining(X) = -1 And dgPick.RowContaining(Y) <> intPickRow Then
confirm = MsgBox("是否存檔!!", vbQuestion + vbOKCancel)
If confirm = 1 Then cmdPickSave_Click
cmdPickCancel_Click

End If
End Sub

Private Sub dgSku_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

If dgOrder.Row = -1 Then Exit Sub

'新增狀態下無法變更資料列
If cmdSkuSave.Enabled = True And LastRow <> Empty Then
    
    dgSku.Col = intLastCol + 1
    dgSku.Row = intSkuRow
    
    Exit Sub
End If

'揀貨頁籤處理
If rsSku.EOF = False Then
    SSTab2.TabCaption(1) = "揀貨" & "( " & rsSku.Fields("產品編號") & " )"
Else
    SSTab2.TabCaption(1) = "揀貨" ': cmdPickAddnew.Enabled = False: cmdPickSave.Enabled = False: cmdPickCancel.Enabled = False
    Exit Sub
End If

'不允許移至特定欄位
If dgSku.Col = 0 Or dgSku.Col = 8 Or dgSku.Col = 9 Then dgSku.Col = Abs(LastCol): Exit Sub
If dgSku.Col = 6 Then
    If LastCol = 5 Then dgSku.Col = 7: Exit Sub
    If LastCol = 7 Then dgSku.Col = 5: Exit Sub
    dgSku.Col = LastCol
End If

'資料列是否變更
If LastRow = Empty Then Exit Sub

cmdSkuDelete.Enabled = True: cmdSkuEdit.Enabled = True
Screen.MousePointer = 0

SSTab2_Click (0)

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub dgPick_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

cboPick.Visible = False
'If dgSku.Row = -1 Then Exit Sub

'新增狀態下無法變更資料列
If cmdPickSave.Enabled = True And LastRow <> Empty Then
    dgPick.Col = intLastCol
    dgPick.Row = intPickRow
    
    Exit Sub
End If

If dgPick.Col = 3 And cmdPickSave.Enabled = True Then ShowList

'不允許移至特定欄位
If dgPick.Col = 0 Or dgPick.Col = 6 Or dgPick.Col = 7 Then dgPick.Col = Abs(LastCol): Exit Sub
If dgPick.Col = 4 Then
    If LastCol = 3 Then dgPick.Col = 5: Exit Sub
    If LastCol = 5 Then dgPick.Col = 2: Exit Sub
    dgPick.Col = IIf(LastCol = -1, 5, LastCol)
End If
'資料列是否變更
If LastRow = Empty Then Exit Sub

Screen.MousePointer = 0

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error GoTo err_Handle

'狀態排除
If PreviousTab = 1 Or dgOrder.Row = -1 Or cmdSave.Enabled = True Then Exit Sub 'cmdReset1_Click: txtOrderkey.Text = "":

cmdSave.Enabled = False: cmdCancel.Enabled = False ': cmdReset1.Enabled = False
cmdAddNew.Enabled = True: cmdEdit.Enabled = True: cmdDelete.Enabled = True

txtOrderkey.Text = RTrim(rsOrder.Fields("出貨單號")): txtOrderkey.BackColor = &HE0E0E0
txtStatus.Text = rsOrder.Fields("訂單狀態")
txtAccountkey.Text = RTrim(rsOrder.Fields("帳款客戶代號")) & "": txtAccountkey.BackColor = &HE0E0E0
txtAccount.Text = RTrim(rsOrder.Fields("帳款客戶名稱")) & "": txtAccount.BackColor = &HE0E0E0
txtCompanykey.Text = RTrim(rsOrder.Fields("送貨客戶代號")) & "": txtCompanykey.BackColor = &HE0E0E0
txtCompany.Text = RTrim(rsOrder.Fields("送貨客戶名稱")) & "": txtCompany.BackColor = &HE0E0E0
txtTel.Text = RTrim(rsOrder.Fields("客戶電話")) & "": txtTel.BackColor = &HE0E0E0
txtAddress.Text = RTrim(rsOrder.Fields("客戶地址")) & "": txtAddress.BackColor = &HE0E0E0
txtOrderdate.Text = RTrim(rsOrder.Fields("排出日期")) & "": txtOrderdate.BackColor = &HE0E0E0
txtDeliverydate.Text = RTrim(rsOrder.Fields("預出日期")) & "": txtDeliverydate.BackColor = &HE0E0E0
txtTransferkey.Text = RTrim(rsOrder.Fields("車商代號")) & "": txtTransferkey.BackColor = &HE0E0E0
txtTransfer.Text = RTrim(rsOrder.Fields("車商名稱")) & "": txtTransfer.BackColor = &HE0E0E0
cboOrderCar.Text = RTrim(rsOrder.Fields("車號") & " "): cboUsepallet.Text = RTrim(rsOrder.Fields("棧板使用") & " ")
txtExternOrderkey.Text = RTrim(rsOrder.Fields("客戶單號")) & "": txtExternOrderkey.BackColor = &HE0E0E0
txtWeight.Text = rsOrder.Fields("重量") & "": txtWeight.BackColor = &HE0E0E0
txtRouteKey.Text = rsOrder.Fields("路線編號") & "": txtRouteKey.BackColor = &HE0E0E0

Select Case RTrim(rsOrder.Fields("棧板使用"))
        Case "N"
            cboUsepallet.ListIndex = 0
        Case "Y"
            cboUsepallet.ListIndex = 1
        Case Else
            cboUsepallet.ListIndex = -1
        End Select

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
On Error GoTo err_Handle
Dim i As Integer

If PreviousTab = 1 Then Exit Sub
'訂單明細新增時不允許切換頁籤
If PreviousTab = 0 And cmdSkuSave.Enabled = True Then SSTab2.Tab = 0: dgSku.SetFocus: Exit Sub
'維護揀貨時切換頁籤保留揀貨資料，若切換後有更改品項則不保留
If cmdPickSave.Enabled = True And dgSku.Row > -1 And intSkuRow = dgSku.Row Then dgPick.SetFocus: Exit Sub
dgPick.AllowUpdate = False
If dgSku.Row = -1 Then cmdPickEdit.Enabled = False: cmdPickDelete.Enabled = False: cmdPickSave.Enabled = False: cmdPickCancel.Enabled = False: Set dgPick.DataSource = Nothing: Exit Sub

Screen.MousePointer = 11

'更新揀貨明細
strSql = "select " & _
        "picklinenumber as 揀貨項次 " & _
        ",lot as 批號 " & _
        ",palletid as 棧板編號 " & _
        ",pallet as 棧板種類 " & _
        ",UOM as 單位 " & _
        ",pickqty as 揀貨數量" & _
        ",adddate as 新增日期 " & _
        ",editdate as 修改日期 " & _
        "from pickdetail " & _
        "where orderkey = '" & rsOrder.Fields("出貨單號") & "' " & _
        "and sku = '" & rsSku.Fields("產品編號") & "' " & _
        "and linenumber = " & rsSku.Fields("項次") & " " & _
        "order by picklinenumber"
        
Set rsPick = New ADODB.Recordset
rsPick.CursorLocation = 3
rsPick.Open strSql, cnAccess, adOpenKeyset, adLockPessimistic

Set dgPick.DataSource = rsPick

With dgPick
    .ColumnHeaders = True        '標題行顯示
    .RowHeight = 300
    .Columns(0).Width = 900:       .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000:    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1000:    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1600:    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 500:    .Columns(4).Alignment = dbgCenter
    .Columns(5).Width = 1000:    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 1600:    .Columns(6).Alignment = dbgCenter
    .Columns(7).Width = 1600:    .Columns(7).Alignment = dbgCenter
    
End With

Frame6.Enabled = True: cmdPickAddnew.Enabled = True: cmdPickEdit.Enabled = True: cmdPickDelete.Enabled = True: cmdPickSave.Enabled = False: cmdPickCancel.Enabled = False
If dgPick.Row = -1 Then cmdPickEdit.Enabled = False: cmdPickDelete.Enabled = False: cmdPickSave.Enabled = False: cmdPickCancel.Enabled = False

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub Form_Load()

cboUsepallet.AddItem "N"
cboUsepallet.AddItem "Y"
cboStatus.AddItem "0--一般"
cboStatus.AddItem "9--完成"
cboPick.AddItem "PTA1W110140"
cboPick.AddItem "PTB1P110110"
cboPick.AddItem "PTD1W110110"
cboPick.AddItem "PTE1W110110"
cboPick.AddItem "PTG1W100120"
cboPick.AddItem "PTK1W"
cboPick.AddItem "NONE"
Set rsOrder = Nothing
cboUsepallet.ListIndex = 0
'cboStatus.ListIndex = 0
SSTab1.Left = 0
SSTab2.Left = 0
Frame1.Left = 0
SSTab1.Top = Frame1.Top + Frame1.Height
SSTab1.Tab = 0

End Sub
Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '最小化

If Me.ScaleHeight < Frame1.Top + Frame1.Height + 500 Then
    Exit Sub
Else
    
    SSTab1.Height = (Me.ScaleHeight - Frame1.Height) / 2
    Frame2.Height = SSTab1.Height - 420
    dgOrder.Height = SSTab1.Height - 500
    dgSku.Height = SSTab1.Height - 500
    dgPick.Height = SSTab1.Height - 500
    SSTab2.Top = SSTab1.Height + SSTab1.Top + 100
    SSTab2.Height = SSTab1.Height

End If

If Me.ScaleWidth < Frame1.Width + Frame1.Left Then

Exit Sub

    Else
    Frame2.Width = Me.ScaleWidth - 120 - (Frame3.Width + 100)
    dgOrder.Width = Me.ScaleWidth - 120
    dgSku.Width = Me.ScaleWidth - 120 - (Frame5.Width + 100)
    dgPick.Width = Me.ScaleWidth - 120 - (Frame6.Width + 100)
    SSTab1.Width = Me.ScaleWidth
    SSTab2.Width = Me.ScaleWidth

End If

End Sub

Private Sub cmdReset_Click()

'重設
txt1S.Text = "": txt1E.Text = ""
txt2S.Text = "": txt2E.Text = ""
txt3S.Text = "": txt3E.Text = ""
cboStatus.ListIndex = -1
cboCar.ListIndex = -1
cboTransferKey.ListIndex = -1

End Sub

Private Sub dgOrder_HeadClick(ByVal ColIndex As Integer)

If dgOrder.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsOrder.Sort = rsOrder.Fields(ColIndex).Name & " DESC"
    dgOrder.ClearSelCols
    intColumnIndex = rsOrder.Fields.Count

Else
    rsOrder.Sort = rsOrder.Fields(ColIndex).Name
    dgOrder.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub
Private Sub dgSku_HeadClick(ByVal ColIndex As Integer)

If dgSku.Row = -1 Or cmdSkuSave.Enabled = True Then Exit Sub
If intColumnIndex = ColIndex Then
    rsSku.Sort = dgSku.Columns(ColIndex).Caption & " DESC"
    dgSku.ClearSelCols
    intColumnIndex = 255

Else
    rsSku.Sort = dgSku.Columns(ColIndex).Caption
    dgSku.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub
Private Sub dgPick_HeadClick(ByVal ColIndex As Integer)

If dgPick.Row = -1 Or cmdPickSave.Enabled = True Then Exit Sub
If intColumnIndex = ColIndex Then
    rsPick.Sort = dgPick.Columns(ColIndex).Caption & " DESC"
    dgPick.ClearSelCols
    intColumnIndex = 255

Else
    rsPick.Sort = dgPick.Columns(ColIndex).Caption
    dgPick.ClearSelCols
    intColumnIndex = ColIndex

End If

End Sub
Private Sub dgPick_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{tab}"

End Sub
Private Sub cboPick_Click()

rsPick("棧板種類") = cboPick.Text

End Sub
Private Sub ShowList()

With dgPick
.RowHeight = cboPick.Height - 10
If .Col = 3 Then
    If .Columns(.Col).Left > 0 Then
            cboPick.Visible = True
            cboPick.Move .Left + .Columns(.Col).Left + 15, .Top + .RowTop(.Row), .Columns(.Col).Width
            If cboPick.Left + cboPick.Width > .Left + .Width Then '如果欄位超出DataGrid的顯示範圍的處理
                cboPick.Width = cboPick.Width + .Left + .Width - cboPick.Left - cboPick.Width
            End If
            cboPick.Text = rsPick("棧板種類")  '更新Combo的值
    Else '如果用捲軸捲動出了DataGrid的顯示範圍，值會小於0
        cboPick.Visible = False
    End If
Else
    cboPick.Visible = False
End If
End With
End Sub
Private Sub dgPick_Scroll(Cancel As Integer)
ShowList
End Sub
Private Sub dgPick_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
ShowList
End Sub
Private Sub dgPick_RowResize(Cancel As Integer)
ShowList
End Sub
Private Sub cmdExit_Click()
Unload Me '結束此程序
'End 結束應用程式
End Sub
Private Sub txtRouteKey_LostFocus()
txtRouteKey.Text = Format(txtRouteKey.Text, "0000000000")
End Sub

Private Sub txtWeight_LostFocus()

If IsNumeric(txtWeight.Text) = False Then MsgBox "你輸入的不是數字": txtWeight.SetFocus

End Sub
