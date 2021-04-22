VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form frm_OP_ManualOrders 
   Caption         =   "訂單維護作業 "
   ClientHeight    =   8310
   ClientLeft      =   240
   ClientTop       =   690
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   12480
   Begin MSDataGridLib.DataGrid dgMain 
      Height          =   2295
      Left            =   0
      TabIndex        =   104
      Top             =   4560
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   3
      RowHeight       =   20
      TabAction       =   2
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
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
   Begin VB.CommandButton cmdDelRs 
      BackColor       =   &H00FFFFC0&
      Caption         =   "刪除CTrl+D"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1080
      Style           =   1  '圖片外觀
      TabIndex        =   103
      Top             =   3960
      Width           =   1035
   End
   Begin VB.CommandButton cmdAddRs 
      BackColor       =   &H00FFFFC0&
      Caption         =   "新增CTrl+A"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   0
      Style           =   1  '圖片外觀
      TabIndex        =   102
      Top             =   3960
      Width           =   1035
   End
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   9360
      TabIndex        =   20
      Top             =   3960
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
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   -2147483624
      StartOfWeek     =   97189889
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin VB.Frame fam_Orders 
      BackColor       =   &H8000000A&
      Enabled         =   0   'False
      Height          =   3075
      Left            =   0
      TabIndex        =   30
      Top             =   840
      Width           =   11520
      Begin VB.CheckBox Chk_receive 
         Caption         =   "順收"
         Height          =   255
         Left            =   9720
         TabIndex        =   110
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txt_OtQty 
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
         Left            =   9495
         MaxLength       =   20
         TabIndex        =   108
         Top             =   1130
         Width           =   990
      End
      Begin VB.ComboBox txt_B_city 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frm_OP_ManualOrders.frx":0000
         Left            =   7500
         List            =   "frm_OP_ManualOrders.frx":0002
         TabIndex        =   106
         Top             =   1130
         Width           =   1575
      End
      Begin VB.ComboBox cmdFacility 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frm_OP_ManualOrders.frx":0004
         Left            =   6000
         List            =   "frm_OP_ManualOrders.frx":0014
         TabIndex        =   105
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txt_Description 
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
         TabIndex        =   100
         Top             =   840
         Width           =   8700
      End
      Begin VB.ComboBox cbo_Priority 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frm_OP_ManualOrders.frx":0041
         Left            =   960
         List            =   "frm_OP_ManualOrders.frx":0043
         TabIndex        =   79
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame frm_OP_ManualShipToOrders 
         BackColor       =   &H00004040&
         BorderStyle     =   0  '沒有框線
         Enabled         =   0   'False
         Height          =   1470
         Left            =   120
         TabIndex        =   60
         Top             =   1440
         Width           =   10365
         Begin VB.TextBox txt_ShipToAreaCode 
            Appearance      =   0  '平面
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
            Left            =   6360
            TabIndex        =   69
            Top             =   825
            Width           =   3100
         End
         Begin VB.TextBox txt_ShipToZIP 
            Appearance      =   0  '平面
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
            Left            =   6360
            TabIndex        =   68
            Top             =   540
            Width           =   3100
         End
         Begin VB.TextBox txt_ShipToExtraDemand2 
            Appearance      =   0  '平面
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
            Left            =   5250
            TabIndex        =   67
            Top             =   1125
            Width           =   4215
         End
         Begin VB.TextBox txt_ShipToExtraDemand1 
            Appearance      =   0  '平面
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
            Left            =   1005
            TabIndex        =   66
            Top             =   1125
            Width           =   4215
         End
         Begin VB.TextBox txt_ShipToFullName 
            Appearance      =   0  '平面
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
            Left            =   1005
            TabIndex        =   65
            Top             =   255
            Width           =   4215
         End
         Begin VB.TextBox txt_ShipToShortName 
            Appearance      =   0  '平面
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
            Left            =   6360
            TabIndex        =   64
            Top             =   240
            Width           =   3100
         End
         Begin VB.TextBox txt_ShipToAddress 
            Appearance      =   0  '平面
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
            Left            =   1005
            TabIndex        =   63
            Top             =   825
            Width           =   4215
         End
         Begin VB.TextBox txt_ShipToContact 
            Appearance      =   0  '平面
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
            Left            =   1005
            TabIndex        =   62
            Top             =   540
            Width           =   1635
         End
         Begin VB.TextBox txt_ShipToPhone 
            Appearance      =   0  '平面
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
            Left            =   3585
            TabIndex        =   61
            Top             =   540
            Width           =   1635
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "轉運到貨客戶"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   30
            Left            =   120
            TabIndex        =   78
            Top             =   60
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶名稱"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   77
            Top             =   315
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶簡稱"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   28
            Left            =   5535
            TabIndex        =   76
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運送地址"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   27
            Left            =   120
            TabIndex        =   75
            Top             =   870
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運送區域"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   26
            Left            =   5535
            TabIndex        =   74
            Top             =   900
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "特殊需求"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   73
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "郵遞區號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   24
            Left            =   5535
            TabIndex        =   72
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "聯絡人"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   71
            Top             =   585
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "電話"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   22
            Left            =   3150
            TabIndex        =   70
            Top             =   585
            Width           =   390
         End
      End
      Begin VB.TextBox txtShipToKey 
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
         Left            =   4680
         TabIndex        =   59
         Top             =   1140
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdShipToList 
         BackColor       =   &H00FFC0C0&
         Caption         =   "？"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6345
         Style           =   1  '圖片外觀
         TabIndex        =   58
         Top             =   1125
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.ComboBox cmbStorerkey 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frm_OP_ManualOrders.frx":0045
         Left            =   960
         List            =   "frm_OP_ManualOrders.frx":0047
         TabIndex        =   57
         Top             =   160
         Width           =   2175
      End
      Begin VB.TextBox txt_DeliveryDate 
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
         Left            =   9150
         TabIndex        =   56
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txt_OrderKey 
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
         Left            =   6900
         TabIndex        =   55
         Top             =   160
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Consigneelist 
         BackColor       =   &H00FFC0C0&
         Caption         =   "？"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2625
         Style           =   1  '圖片外觀
         TabIndex        =   54
         Top             =   1125
         Width           =   315
      End
      Begin VB.TextBox txt_Extern 
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
         Left            =   4050
         MaxLength       =   20
         TabIndex        =   53
         Top             =   160
         Width           =   1575
      End
      Begin VB.Frame fam_ConsigneeData 
         BackColor       =   &H00004040&
         BorderStyle     =   0  '沒有框線
         Enabled         =   0   'False
         Height          =   1470
         Left            =   120
         TabIndex        =   34
         Top             =   1485
         Width           =   10365
         Begin VB.ComboBox cmb_ExtraDemand2 
            BackColor       =   &H8000000B&
            Height          =   300
            Left            =   5235
            Style           =   2  '單純下拉式
            TabIndex        =   44
            Top             =   1125
            Width           =   4230
         End
         Begin VB.ComboBox cmb_ExtraDemand1 
            BackColor       =   &H8000000A&
            Height          =   300
            Left            =   1005
            Style           =   2  '單純下拉式
            TabIndex        =   43
            Top             =   1125
            Width           =   4230
         End
         Begin VB.ComboBox cmb_ZIP 
            BackColor       =   &H8000000B&
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
            Left            =   6360
            Style           =   2  '單純下拉式
            TabIndex        =   42
            Top             =   525
            Width           =   3100
         End
         Begin VB.TextBox txt_Phone 
            Appearance      =   0  '平面
            BackColor       =   &H8000000B&
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
            Left            =   3585
            TabIndex        =   41
            Top             =   540
            Width           =   1635
         End
         Begin VB.TextBox txt_Contact 
            Appearance      =   0  '平面
            BackColor       =   &H8000000B&
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
            Left            =   1005
            TabIndex        =   40
            Top             =   540
            Width           =   1635
         End
         Begin VB.TextBox txt_Address 
            Appearance      =   0  '平面
            BackColor       =   &H8000000B&
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
            Left            =   1005
            TabIndex        =   39
            Top             =   825
            Width           =   4215
         End
         Begin VB.TextBox txt_ShortName 
            Appearance      =   0  '平面
            BackColor       =   &H8000000B&
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
            Left            =   6360
            TabIndex        =   38
            Top             =   240
            Width           =   3100
         End
         Begin VB.TextBox txt_FullName 
            Appearance      =   0  '平面
            BackColor       =   &H8000000B&
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
            Left            =   1005
            TabIndex        =   37
            Top             =   255
            Width           =   4215
         End
         Begin VB.ComboBox cmb_AreaCode 
            BackColor       =   &H8000000B&
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
            Left            =   6360
            Style           =   2  '單純下拉式
            TabIndex        =   36
            Top             =   840
            Width           =   3100
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFC0C0&
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
            Height          =   240
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   35
            Top             =   60
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "電話"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   16
            Left            =   3150
            TabIndex        =   52
            Top             =   585
            Width           =   390
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "聯絡人"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   51
            Top             =   585
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "郵遞區號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   14
            Left            =   5535
            TabIndex        =   50
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "特殊需求"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   49
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運送區域"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   9
            Left            =   5535
            TabIndex        =   48
            Top             =   900
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運送地址"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   47
            Top             =   870
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶簡稱"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   7
            Left            =   5535
            TabIndex        =   46
            Top             =   300
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶名稱"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   45
            Top             =   315
            Width           =   780
         End
      End
      Begin VB.TextBox txt_ConsigneeKey 
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
         TabIndex        =   33
         Top             =   1140
         Width           =   1575
      End
      Begin VB.TextBox txt_OrderDate 
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
         Left            =   9150
         TabIndex        =   32
         Top             =   160
         Width           =   1335
      End
      Begin VB.TextBox txtType 
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3495
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   31
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "件數"
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
         Index           =   38
         Left            =   9120
         TabIndex        =   109
         Top             =   1180
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "組單類別"
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
         Index           =   37
         Left            =   6735
         TabIndex        =   107
         Top             =   1200
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
         Left            =   120
         TabIndex        =   101
         Top             =   900
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
         Index           =   17
         Left            =   120
         TabIndex        =   89
         Top             =   540
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "轉運到貨客戶編號"
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
         Index           =   21
         Left            =   3000
         TabIndex        =   88
         Top             =   1185
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "客戶採購編號"
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
         Left            =   5700
         TabIndex        =   87
         Top             =   225
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "訂單編號"
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
         Left            =   3195
         TabIndex        =   86
         Top             =   225
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "客戶編號"
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
         Index           =   4
         Left            =   120
         TabIndex        =   85
         Top             =   1185
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "送貨日"
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
         Left            =   8565
         TabIndex        =   84
         Top             =   540
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "訂單日"
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
         Left            =   8565
         TabIndex        =   83
         Top             =   225
         Width           =   585
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
         Left            =   120
         TabIndex        =   82
         Top             =   220
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "訂單狀態"
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
         Index           =   32
         Left            =   2640
         TabIndex        =   81
         Top             =   540
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "配送倉別"
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
         Index           =   33
         Left            =   5160
         TabIndex        =   80
         Top             =   540
         Width           =   780
      End
   End
   Begin VB.Frame fam_Header 
      Height          =   870
      Left            =   0
      TabIndex        =   16
      Top             =   -75
      Width           =   11520
      Begin VB.CommandButton cmd_AddNew 
         BackColor       =   &H00FF80FF&
         Caption         =   "新  增"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         Style           =   1  '圖片外觀
         TabIndex        =   3
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Modify 
         BackColor       =   &H00FF8080&
         Caption         =   "修  改"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Save 
         BackColor       =   &H008080FF&
         Caption         =   "存  檔"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7560
         Style           =   1  '圖片外觀
         TabIndex        =   4
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Delete 
         BackColor       =   &H000080FF&
         Caption         =   "刪  除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         Style           =   1  '圖片外觀
         TabIndex        =   6
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Exit 
         BackColor       =   &H00FFC0FF&
         Caption         =   "離  開"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   10200
         Style           =   1  '圖片外觀
         TabIndex        =   15
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Cancel 
         BackColor       =   &H0080FFFF&
         Caption         =   "取  消"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         Style           =   1  '圖片外觀
         TabIndex        =   5
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmd_OrdersQuery 
         BackColor       =   &H0080C0FF&
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
         Height          =   435
         Left            =   2640
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txt_QueryExternOrderKey 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1020
         MaxLength       =   10
         TabIndex        =   0
         Top             =   300
         Width           =   1590
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000080&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00008000&
         BorderWidth     =   2
         Height          =   585
         Left            =   3555
         Top             =   195
         Width           =   7920
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00004040&
         BackStyle       =   1  '不透明
         Height          =   495
         Left            =   2610
         Top             =   210
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "TMS單號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   195
         Left            =   105
         TabIndex        =   17
         Top             =   390
         Width           =   840
      End
   End
   Begin VB.Frame fam_OrderDetail 
      BackColor       =   &H8000000B&
      Height          =   4485
      Left            =   -240
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   16920
      Begin VB.CommandButton cmd_DetailVerify 
         BackColor       =   &H00808080&
         Caption         =   "細項驗證"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   9840
         Picture         =   "frm_OP_ManualOrders.frx":0049
         Style           =   1  '圖片外觀
         TabIndex        =   21
         ToolTipText     =   "新增"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmd_DetailCancel 
         BackColor       =   &H0080FFFF&
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10605
         Style           =   1  '圖片外觀
         TabIndex        =   12
         ToolTipText     =   "新增"
         Top             =   765
         Width           =   765
      End
      Begin VB.CommandButton cmd_DetailSave 
         BackColor       =   &H008080FF&
         Caption         =   "儲存"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9765
         Style           =   1  '圖片外觀
         TabIndex        =   11
         ToolTipText     =   "新增"
         Top             =   765
         Width           =   765
      End
      Begin VB.CommandButton cmd_DetailDel 
         BackColor       =   &H000080FF&
         Caption         =   "刪除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10605
         Style           =   1  '圖片外觀
         TabIndex        =   13
         ToolTipText     =   "刪除"
         Top             =   285
         Width           =   765
      End
      Begin VB.CommandButton cmd_DetailAddNew 
         BackColor       =   &H00FF80FF&
         Caption         =   "新增"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8925
         Style           =   1  '圖片外觀
         TabIndex        =   9
         ToolTipText     =   "新增"
         Top             =   285
         Width           =   765
      End
      Begin VB.CommandButton cmd_DetailModify 
         BackColor       =   &H00FF8080&
         Caption         =   "修改"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9765
         Style           =   1  '圖片外觀
         TabIndex        =   10
         ToolTipText     =   "新增"
         Top             =   285
         Width           =   765
      End
      Begin VB.Frame fam_DetailData 
         Appearance      =   0  '平面
         BackColor       =   &H00004000&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1200
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   8700
         Begin VB.TextBox txtCasecnt 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   585
            Locked          =   -1  'True
            TabIndex        =   98
            ToolTipText     =   "每箱入數"
            Top             =   840
            Width           =   915
         End
         Begin VB.TextBox txtLot5 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7635
            MaxLength       =   8
            TabIndex        =   96
            ToolTipText     =   "指定到期日"
            Top             =   840
            Width           =   915
         End
         Begin VB.TextBox txtLot4 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7635
            MaxLength       =   8
            TabIndex        =   94
            ToolTipText     =   "指定製造日"
            Top             =   480
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox txtOrderCS 
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
            Height          =   285
            Left            =   600
            TabIndex        =   92
            ToolTipText     =   "訂單箱數"
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox txtOrderEA 
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
            Height          =   285
            Left            =   2025
            TabIndex        =   90
            ToolTipText     =   "訂單個數"
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox txtLot3 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5280
            TabIndex        =   28
            ToolTipText     =   "生產批號"
            Top             =   480
            Width           =   1635
         End
         Begin VB.ComboBox cboLot6 
            BeginProperty Font 
               Name            =   "細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frm_OP_ManualOrders.frx":0913
            Left            =   3480
            List            =   "frm_OP_ManualOrders.frx":0915
            TabIndex        =   26
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtNotes 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   8
            ToolTipText     =   "商品備註"
            Top             =   825
            Width           =   4860
         End
         Begin VB.ComboBox cboSku 
            BeginProperty Font 
               Name            =   "細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frm_OP_ManualOrders.frx":0917
            Left            =   600
            List            =   "frm_OP_ManualOrders.frx":0919
            TabIndex        =   7
            Top             =   120
            Width           =   2895
         End
         Begin VB.TextBox txt_SkuDescr 
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
            Height          =   285
            Left            =   4080
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   120
            Width           =   4545
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "每箱"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   99
            Top             =   900
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "到期日"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   35
            Left            =   6960
            TabIndex        =   97
            Top             =   900
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "製造日"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   36
            Left            =   6960
            TabIndex        =   95
            Top             =   540
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "箱數"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   93
            Top             =   540
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "個數"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   18
            Left            =   1560
            TabIndex        =   91
            Top             =   540
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "批號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   34
            Left            =   4800
            TabIndex        =   29
            Top             =   540
            Width           =   420
         End
         Begin VB.Label lable14 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "倉別"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   17
            Left            =   3000
            TabIndex        =   27
            Top             =   555
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "備註"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   255
            Index           =   19
            Left            =   1560
            TabIndex        =   25
            Top             =   900
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貨號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   24
            Top             =   180
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "品名"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFC0&
            Height          =   195
            Index           =   31
            Left            =   3600
            TabIndex        =   22
            Top             =   180
            Width           =   420
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_OrderDetail 
         Height          =   2520
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   4445
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00008000&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00404080&
         BorderWidth     =   2
         Height          =   495
         Index           =   4
         Left            =   9720
         Top             =   720
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00404080&
         BorderWidth     =   2
         Height          =   495
         Index           =   1
         Left            =   10560
         Top             =   240
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00400000&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00800080&
         BorderWidth     =   2
         Height          =   495
         Index           =   0
         Left            =   10560
         Top             =   720
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404080&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00004040&
         BorderWidth     =   2
         Height          =   495
         Index           =   3
         Left            =   8880
         Top             =   240
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00404080&
         BorderWidth     =   2
         Height          =   495
         Index           =   2
         Left            =   9720
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "frm_OP_ManualOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form 設計時期的高
Private dbsrcFormWidth As Double     'Form 設計時期的寬
Private iLoop As Double              '迴圈計數

Private intsrcSKUNowRow As Double

Private arZip() As String               '郵遞區號
Private arZIPArea() As String           '郵遞區號檔設定之 AreaCode
Private arAreaCode() As String          '區域代碼
Private arExtraDemand() As String       '特殊需求
Private rsMain As ADODB.Recordset

Private Sub cmdAddRs_Click()
If rsMain Is Nothing Then Exit Sub
If dgMain.Enabled = False Then Exit Sub
Dim lngSeq As Long

'取項次
If rsMain.RecordCount > 0 Then
    rsMain.MoveLast
    lngSeq = Val(rsMain("項次"))
End If

'新增
rsMain.AddNew
rsMain("項次") = Format(lngSeq + 1, "00000")
'Call dgMain_RowColChange(1, 1)
dgMain.SetFocus
dgMain.Col = 1
End Sub

Private Sub cmdDelRs_Click()
If rsMain Is Nothing Then Exit Sub
If rsMain.RecordCount = 0 Then Exit Sub
If dgMain.Enabled = False Then Exit Sub
On Error GoTo err_Handle

If MsgBox("確定刪除此項次(" & RTrim(rsMain("項次")) & ")?", vbOKCancel, "確認刪除") <> vbOK Then: Exit Sub

'刪除
rsMain.Delete

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, Me.Caption & "_cmdDelRs_Click")

End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain

'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 200 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgmain_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then 'Ctrl+A
    Call cmdAddRs_Click
    dgMain.SetFocus: dgMain.Col = 0
End If

If KeyAscii = 4 Then 'Ctrl+D
    cmdDelRs.SetFocus
    Call cmdDelRs_Click
End If

If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle
Dim strDate As String

With dgMain

'同一行選取
'If LastRow = Empty Then

    'dtpDeliveryTime.Visible = False
    If .DataSource Is Nothing Then Exit Sub
    If rsMain.RecordCount = 0 Then Exit Sub
    If rsMain.EOF Then Exit Sub
    If .Col = -1 Then Exit Sub
    
        If LastCol <> -1 Then '別的物件回來
        
        '品號檢查
        If rsMain.Fields(LastCol).Name = "貨號" And Len(RTrim(rsMain("貨號"))) > 0 Then
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open "select descr,casecnt,innerpack,busr3=isnull(busr3,''),busr2 = isnull(busr2,''),busr1=isnull(busr1,'') from gv_skuxpack where sku = '" & rsMain("貨號") & "' and storerkey = '" & mySplit(cmbStorerkey, " ", 0) & "' ", cn, adOpenForwardOnly, adLockReadOnly
            If tmp_Rs.EOF = True Then
                tmp_Rs.Close: .Col = LastCol: MsgBox "無此貨號！", 64, Me.Caption: .SetFocus: Exit Sub
            Else
                rsMain("品名") = RTrim(tmp_Rs("descr"))
                rsMain("大單位名稱") = RTrim(tmp_Rs("busr3"))
                rsMain("中單位名稱") = RTrim(tmp_Rs("busr2"))
                rsMain("小單位名稱") = RTrim(tmp_Rs("busr1"))
                rsMain("大單位入數") = RTrim(tmp_Rs("casecnt"))
                rsMain("中單位入數") = RTrim(tmp_Rs("innerpack"))
                tmp_Rs.Close
                .Col = .Col + 1
            End If
        End If
    End If
    If rsMain("大單位數量") > 0 And rsMain("大單位入數") = 0 Then MsgBox "大單位入數為0，無法輸入大單位數量！", 16, "注意": rsMain("大單位數量") = 0: .SetFocus ': Exit Sub
    If rsMain("中單位數量") > 0 And rsMain("中單位入數") = 0 Then MsgBox "中單位入數為0，無法輸入中單位數量！", 16, "注意": rsMain("中單位數量") = 0: .SetFocus ': Exit Sub
    If RTrim(GetWord(Trim(rsMain("揀貨加工")), 1, 10)) <> Trim(rsMain("揀貨加工")) Then MsgBox "加工註記不得超過10個字元！", 16, "注意": .SetFocus: rsMain("揀貨加工") = GetWord(Trim(rsMain("揀貨加工")), 1, 10) ': Exit Sub
    
    '不允許移至
    If rsMain.Fields(.Col).Name = "品名" Or rsMain.Fields(.Col).Name = "大單位名稱" Or rsMain.Fields(.Col).Name = "中單位名稱" Or rsMain.Fields(.Col).Name = "中單位入數" Then .Col = .Col + 1
    If rsMain.Fields(.Col).Name = "小單位名稱" Then .Col = .Col + 3
    If rsMain.Fields(.Col).Name = "大單位入數" Then .Col = .Col + 2
'End If

End With
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
'從記憶體中移除表單，藉此引起 [Terminate] 事件
Set frm_OP_ManualOrders = Nothing
End Sub

Private Sub cbo_Priority_Click()
Label3(3).Caption = "到貨日"
frm_OP_ManualShipToOrders.ZOrder 1
Label3(21).Visible = False: txtShipToKey.Visible = False: cmdShipToList.Visible = False

If mySplit(Trim(cbo_Priority), " ", 0) = "R" Or mySplit(Trim(cbo_Priority), " ", 0) = "RC" Then

    Label3(3).Caption = "提貨日"
ElseIf mySplit(Trim(cbo_Priority), " ", 0) = "A2B" Then

    Label3(3).Caption = "提貨日"
    Label3(21).Visible = True: txtShipToKey.Visible = True: cmdShipToList.Visible = True
    
End If

End Sub

Private Sub cboSku_Click()

Dim rsTmp As New ADODB.Recordset, strLot6 As String

strLot6 = cboLot6.Text '紀錄倉別

rsTmp.Open "select casecnt,descr = rtrim(isnull(descr,'')) from gv_skuxpack where storerkey = '" & mySplit(cmbStorerkey, " ", 0) & "' and sku = '" & cboSku & "' ", cn
If Not rsTmp.EOF Then
    txtCasecnt = rsTmp("casecnt"): txt_SkuDescr = rsTmp("descr")
Else
    txtCasecnt = 0:: txt_SkuDescr = ""
End If
rsTmp.Close

'取倉別
str_SQL = "select distinct isnull(lotattribute.lottable06,'') as lottable06 from " & strWMSDB & "..lotxlocxid lotxlocxid join " & strWMSDB & "..lotattribute lotattribute on lotattribute.lot = lotattribute.lot where lotxlocxid.storerkey = '" & mySplit(RTrim(cmbStorerkey), " ", 0) & "' and lotattribute.sku = '" & cboSku & "' order by lotattribute.lottable06 "
rsTmp.Open str_SQL, cn

If Not rsTmp.EOF Then rsTmp.MoveFirst

cboLot6.Clear
Do While Not rsTmp.EOF
    cboLot6.AddItem Trim(rsTmp("lottable06"))
    rsTmp.MoveNext
Loop

cboLot6 = strLot6 '寫回倉別

End Sub

Private Sub cboSku_LostFocus()
Call cboSku_Click
'dg_OrderDetail.Col = 3: cboLot6.Text = dg_OrderDetail.Text  '倉別
End Sub

Private Sub cmbStorerkey_Click()

''取貨號
'Dim rsTmp As New ADODB.Recordset
'str_SQL = "select sku = rtrim(sku) from sku where storerkey = '" & mySplit(RTrim(cmbStorerkey), " ", 0) & "' order by sku"
'rsTmp.Open str_SQL, cn
'If rsTmp.EOF Then MsgBox "找不到該貨主商品主檔資料", vbOKOnly, Me.Caption: Exit Sub
'rsTmp.MoveFirst
'
''txt_ConsigneeKey = ""
''txtShipToKey = ""
'
'cboSku.Clear
'Do While Not rsTmp.EOF
'    cboSku.AddItem rsTmp("sku")
'    rsTmp.MoveNext
'Loop
'
'rsTmp.Close: Set rsTmp = Nothing

End Sub

Private Sub cmd_AddNew_Click()
'新增

'清除所有欄位值，包含 OrderDetail 部份
Call Clear_AllField
fam_Orders.Enabled = True
fam_Orders.BackColor = &HFF80FF
fam_OrderDetail.BackColor = &HFF80FF

txt_Extern.Enabled = True
txt_OrderKey.Enabled = True
txt_QueryExternOrderKey.Text = ""
cmd_Modify.Enabled = False
cmd_AddNew.Enabled = False
cmd_Delete.Enabled = False
cmd_Save.Enabled = True
cmd_Cancel.Enabled = True

cmd_DetailAddNew.Enabled = True
cmd_DetailModify.Enabled = False
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False
cbo_Priority = ""
cmbStorerkey.Enabled = True
dgMain.Enabled = True

'明細定義
str_SQL = "exec gs_ManualOrder_OrderDetail '" & txt_QueryExternOrderKey.Text & "' "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

Set rsMain = New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsMain)
tmp_Rs.Close

Set dgMain.DataSource = rsMain

SetDataGridColWidth Me.Caption, dgMain

txt_Extern.SetFocus
End Sub

Private Sub cmd_Cancel_Click()
'取消

fam_Orders.Enabled = False
fam_Orders.BackColor = &H8000000A
fam_OrderDetail.BackColor = &H8000000A
txt_Extern.Enabled = False
txt_OrderKey.Enabled = False
cmd_Modify.Enabled = True
cmd_AddNew.Enabled = True
cmd_Save.Enabled = False
cmd_Cancel.Enabled = False

fam_DetailData.Enabled = False
dgMain.Enabled = False
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = False
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False

Dim strTmp As String
strTmp = Trim(txt_QueryExternOrderKey)
'清除所有欄位值，包含 OrderDetail 部份
Call Clear_AllField

'若是修改模式，重新取回原訂單資料
If strTmp <> "" Then
   txt_QueryExternOrderKey.Text = strTmp
   Call cmd_OrdersQuery_Click
End If
End Sub

Private Sub cmd_Consigneelist_Click()
'顯示客戶待選清單
'Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Consigneelist.Name)

If cmbStorerkey = "" Then MsgBox "請先選取貨主！", 64, Me.Caption: Exit Sub

'呼叫視窗
strDataList_Caller = Me.Name & " " & "txt_ConsigneeKey" & " " & mySplit(cmbStorerkey, " ", 0)
frm_ConsigneekeyQuery.Show vbModal

End Sub

Private Sub cmd_Delete_Click()
'訂單 >> 刪除

Call cmd_OrdersQuery_Click

''利豐毛寶不可以手動改單
'If RTrim(cmbStorerkey) = "LLFA01" Then MsgBox "利豐貨主訂單，不可手動改單!", vbOKOnly + vbCritical, "訂單維護": Exit Sub
'If RTrim(cmbStorerkey) = "LMBO01" Then MsgBox "毛寶貨主訂單，不可手動改單!", vbOKOnly + vbCritical, "訂單維護": Exit Sub

If txtType = "刪單" Then Exit Sub

Dim rsTmp As New ADODB.Recordset, strRoute As String, Int_i As Long
Int_i = 0
rsTmp.Open "select orderkey from orders where orderkey = '" & txt_QueryExternOrderKey & "' ", cn
If rsTmp.EOF Then MsgBox "找不到此訂單!", vbOKOnly, "訂單刪除": rsTmp.Close: Exit Sub
rsTmp.Close

''檢查訂單是否已回傳WMS
'str_SQL = "select t2.receipt_no from trp02t t2 (nolock) join " & strWMSDB & "..orders o (nolock) on o.updatesource = t2.c_receipt_no where t2.c_receipt_no = '" & txt_QueryExternOrderKey & "' "
'rsTmp.Open str_SQL, cn
'If Not rsTmp.EOF Then MsgBox "此訂單已回傳WMS，訂單無法刪除!", vbOKOnly, "訂單刪除": rsTmp.Close: Exit Sub
'rsTmp.Close

'檢查是否已排車
str_SQL = "select route_no from trp02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' union select route_no from ort02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' "
rsTmp.Open str_SQL, cn
If Not rsTmp.EOF Then strRoute = rsTmp("route_no") & ""
If Not rsTmp.EOF Then MsgBox "已排路線編號" & rsTmp("route_no") & "，訂單無法刪除!", vbOKOnly, "訂單刪除": rsTmp.Close: Exit Sub
rsTmp.Close

If MsgBox("確定刪除此訂單" & txt_QueryExternOrderKey & " (含切割訂單)? ", vbQuestion + vbYesNo, "訂單刪除") <> vbYes Then Exit Sub
If mySplit(cmbStorerkey, " ", 0) = "LTKK01" Then If MsgBox("刪除台灣麒麟的訂單時，系統將自動發送MAIL通知貨主，是否繼續? ", vbQuestion + vbYesNo, "訂單刪除通知") <> vbYes Then Exit Sub

Call DB_CheckConnectStatus

Tran_Level = cn.BeginTrans
     
    cn.Execute "delete TRP03W where receipt_no in (select receipt_no from trp02w where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
    cn.Execute "delete TRP02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    
    cn.Execute "delete TRP03T where receipt_no in (select receipt_no from trp02t where route_no = 'D' and c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
    cn.Execute "delete TRP02T where route_no = 'D' and c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    
    cn.Execute "delete ORT03W where receipt_no in (select receipt_no from ORT02W where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
    cn.Execute "delete ORT02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    
    cn.Execute "delete ORT03T where receipt_no in (select receipt_no from ort02t where route_no = 'D' and c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
    cn.Execute "delete ORT02T where route_no = 'D' and c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
     
    cn.Execute "delete TRP02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    cn.Execute "delete ORT02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    
    cn.Execute "delete status where orderkey ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
    cn.Execute "update orders set B_PHONE2='00',trafficCop=null,type='刪單' ,editdate = getdate(),editwho = '" & User_id & "' where orderkey='" & txt_QueryExternOrderKey & "' ", RowsAffect, adExecuteNoRecords
    
    '更新路編材積重量
'    cn.Execute "exec gs_UpdateRoute '" & strRoute & "' ", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

txtType = "刪單"
cmbStorerkey.Enabled = True

'LTKK01刪單自動 Mail 通知
If mySplit(cmbStorerkey, " ", 0) = "LTKK01" Then Call SendMail(txt_QueryExternOrderKey)

Exit Sub

err_Handle:
If Tran_Level <> 0 Then Tran_Level = 0: cn.RollbackTrans

Dim tmpString As String
msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
CreateErrorLog Me.Name & "-訂單刪除", Me.Caption, "cmd_Tab2_Remove_Click", tmpString
MsgBox msg_text, vbOKOnly + vbInformation, msg_title

End Sub

'Private Sub DeleteOrder()
'
'Call DB_CheckConnectStatus
'
'Tran_Level = cn.BeginTrans
'
'     cn.Execute "delete TRP02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
'     cn.Execute "delete TRP03W where receipt_no in (select receipt_no from trp02w where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
'     cn.Execute "delete TRP02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
'     cn.Execute "delete ORT02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
'     cn.Execute "delete ORT03W where receipt_no in (select receipt_no from ORT02W where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
'     cn.Execute "delete ORT02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
'     cn.Execute "delete status where orderkey ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
'     cn.Execute "update orders set B_PHONE2='00',trafficCop=null,type='刪單' ,editdate = getdate()  where orderkey='" & txt_QueryExternOrderKey & "' and priority = '" & mySplit(Trim(cbo_Priority), " ", 0) & "' ", RowsAffect, adExecuteNoRecords
'
'cn.CommitTrans: Tran_Level = 0
'
'Exit Sub
'
'err_Handle:
'If Tran_Level <> 0 Then Tran_Level = 0: cn.RollbackTrans
'
'Dim tmpString As String
'msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
'tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
'CreateErrorLog Me.Name & "-訂單刪除", Me.Caption, "cmd_Tab2_Remove_Click", tmpString
'MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'
'End Sub

Private Sub cmd_DetailAddNew_Click()
'訂單細項 >> 新增
intsrcSKUNowRow = 0
fam_DetailData.Enabled = True
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = False
cmd_DetailSave.Enabled = True
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = True
txt_SkuDescr = ""

txtOrderCS.Text = "0": txtOrderEA.Text = "0": txtCasecnt.Text = "0": txtNotes.Text = "": cboLot6.Text = ""
cboSku.SetFocus

End Sub

Private Sub cmd_DetailCancel_Click()
'訂單細項 >> 取消
If intsrcSKUNowRow = 0 Then
   cmd_DetailModify.Enabled = False
   cboSku.Text = "": txt_SkuDescr.Text = "": txtOrderCS.Text = "0": txtOrderEA.Text = "0": txtCasecnt.Text = "0": txtNotes.Text = "": cboLot6.Text = ""
Else
   '修改模式之取消：取回原細項資料
   cmd_DetailModify.Enabled = True
   With dg_OrderDetail
        .Row = intsrcSKUNowRow
        .Col = 1: .Text = cboSku.Text '貨號
        .Col = 2: .Text = Trim(txt_SkuDescr.Text)    '品名
        .Col = 3: .Text = cboLot6.Text   '倉別
        .Col = 4: .Text = txtOrderCS '箱數
        .Col = 5: .Text = txtOrderEA '個數
        .Col = 6: .Text = txtCasecnt '每箱
        .Col = 7: .Text = txtNotes '備註
   End With
End If
cmd_DetailAddNew.Enabled = True
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False
fam_DetailData.Enabled = False
End Sub

Private Sub cmd_DetailModify_Click()
'訂單細項 >> 修改
If intsrcSKUNowRow = 0 Then Exit Sub
fam_DetailData.Enabled = True
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = False
cmd_DetailSave.Enabled = True
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = True

cboSku.SetFocus
End Sub

Private Sub cmd_DetailSave_Click()

txtNotes = myExCharFilter(txtNotes)

'訂單細項 >> 新增
If Len(Trim(cboSku.Text)) = 0 Then Exit Sub
If Len(RTrim(txtLot4)) > 0 And IsDate(Left(Trim(txtLot4), 4) & "/" & Mid(Trim(txtLot4), 5, 2) & "/" & Right(Trim(txtLot4), 2)) = False Then MsgBox "資料錯誤：製造日期格式錯誤!!": Exit Sub
If Len(RTrim(txtLot5)) > 0 And IsDate(Left(Trim(txtLot5), 4) & "/" & Mid(Trim(txtLot5), 5, 2) & "/" & Right(Trim(txtLot5), 2)) = False Then MsgBox "資料錯誤：到期日期格式錯誤!!": Exit Sub
If Len(Trim(cboLot6)) = 0 Then: MsgBox "資料錯誤：請輸入倉別!!": cboLot6.SetFocus: Exit Sub
'And (Left(cbo_Priority, 1) = "A" Or Left(cbo_Priority, 1) = "I")
If Len(RTrim(txtOrderCS.Text)) = 0 Then txtOrderCS.Text = "0"
If Len(RTrim(txtOrderEA.Text)) = 0 Then txtOrderEA.Text = "0"
If CheckSKU(cboSku.Text) = 1 Then Exit Sub

'配送倉別檢查
If Len(Trim(cmdFacility)) = 0 Then
    cmdFacility = "佰事達北倉"
    If UCase(Right(Trim(cboLot6), 2)) = "-C" Then cmdFacility = "佰事達中倉"
    If UCase(Right(Trim(cboLot6), 2)) = "-S" Then cmdFacility = "佰事達南倉"
Else
    If UCase(Right(Trim(cboLot6), 2)) <> "-C" And UCase(Right(Trim(cboLot6), 2)) <> "-S" And cmdFacility <> "佰事達北倉" Then: MsgBox "配送倉別與商品倉別不符!!", 64, "倉別錯誤!!": cboLot6.SetFocus: Exit Sub
    If UCase(Right(Trim(cboLot6), 2)) = "-C" And cmdFacility <> "佰事達中倉" Then: MsgBox "配送倉別與商品倉別不符!!", 64, "倉別錯誤!!": cboLot6.SetFocus: Exit Sub
    If UCase(Right(Trim(cboLot6), 2)) = "-S" And cmdFacility <> "佰事達南倉" Then: MsgBox "配送倉別與商品倉別不符!!", 64, "倉別錯誤!!": cboLot6.SetFocus: Exit Sub

End If

Dim lngOrderqty As Long
lngOrderqty = txtOrderCS.Text * txtCasecnt + txtOrderEA
If lngOrderqty = 0 Then
   msg_text = "資料錯誤：訂單數量不得為 0"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txtOrderCS.SetFocus
   Exit Sub
End If

'1. 選取資料存入的 ROW
With dg_OrderDetail
     If intsrcSKUNowRow <> 0 Then '修改模式：覆蓋原來資料
'        intsrcSKUNowRow = .Text
'        .Row = intsrcSKUNowRow
        If CheckSKU(cboSku.Text) = 1 Then Exit Sub
     Else
        Dim dbMaxNo As Double
        .Row = .Rows - 2: .Col = 0: dbMaxNo = Val(.Text)
        .Rows = .Rows + 1
        .Row = .Rows - 2
     End If
End With

'2. 資料存入 dg_srcSKU 指定的 Row
With dg_OrderDetail
     .Col = 0    '序號
     If intsrcSKUNowRow <> 0 Then
'        .Text = intsrcSKUNowRow    '貨物資料修改模式，沿用原序號
     Else
        .Text = Format(CInt(dbMaxNo + 1), "0000")         '貨物資料新增模式，產生新序號
     End If
     
     .Col = 1: .Text = UCase(RTrim(cboSku)) '貨號
     .Col = 2: .Text = txt_SkuDescr  '品名
     .Col = 3: .Text = cboLot6.Text   '倉別
     .Col = 4: .Text = txtOrderCS '箱數
     .Col = 5: .Text = txtOrderEA '個數
     .Col = 6: .Text = txtCasecnt '每箱
     .Col = 7: .Text = txtNotes '備註
     .Col = 8: .Text = txtLot3 '生產批號
     .Col = 9: .Text = txtLot4 '製造日
     .Col = 10: .Text = txtLot5 '到期日
     
End With

'重設修改模式識別旗標值：新增模式
intsrcSKUNowRow = 0

'訂單明細資料新增完畢，清除欄位值
cboSku = "": txtOrderCS.Text = "0": txtOrderEA.Text = "0": txtCasecnt.Text = "0": txtNotes.Text = "": cboLot6.Text = "": txt_SkuDescr = "": txtLot3 = "": txtLot4 = "": txtLot5 = ""
cboSku.SetFocus

intsrcSKUNowRow = 0
fam_DetailData.Enabled = False
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = True
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False

End Sub

Private Sub cmd_DetailDel_Click()
'訂單細項 >> 刪除
If intsrcSKUNowRow = 0 Then Exit Sub

Dim j As Integer
'1. 將刪除列資料由下一列資料取代
'   而後的資料列往上移一列
dg_OrderDetail.Visible = False
With dg_OrderDetail
     For iLoop = intsrcSKUNowRow To .Rows - 2   '會有多一行空白列
         .Row = iLoop
         For j = 0 To .Cols - 1
             .Col = j
             .Text = .TextArray((.Row + 1) * .Cols + .Col)
         Next j
         '防止最後第一列往上移給最後第二列時，會是弄白資料列，[序號] 欄位不能有值
         '有資料的列，[序號] 必須重新編號
         .Col = 0
         'If Val(.Text) = 0 Then .Text = "" Else .Text = .Row
     Next iLoop
'2. Grid 總列數 - 1
     .Rows = .Rows - 1
     .Row = 1
     For iLoop = 0 To .Cols - 1
         .ColSel = iLoop
     Next iLoop
End With
dg_OrderDetail.Visible = True
'3. Reset 變數
intsrcSKUNowRow = 0

'4. 清除欄位資料
cboSku.Text = "": txt_SkuDescr.Text = "": txtOrderCS.Text = "0": txtOrderEA.Text = "0": txtCasecnt.Text = "0": txtNotes.Text = "": cboLot6.Text = ""

fam_DetailData.Enabled = False
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = True
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False
fam_DetailData.Enabled = False

End Sub

'Private Sub cmd_DetailVerify_Click()
''細項驗證
'If dg_OrderDetail.Rows = 2 Then Exit Sub
'If Trim(cmbStorerkey.Text) = "" Then
'   msg_text = "未輸入 [貨主] 資料"
'   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'   Exit Sub
'End If
'
'Dim strSku As String, strErrorSKU As String
'Dim dbShipQty As Double
'
'Call DB_CheckConnectStatus
'Call ReDim_Recordset(tmp_Rs)
'strErrorSKU = ""
'With dg_OrderDetail
'     If .Rows = 2 Then Exit Sub
'     For iLoop = 1 To .Rows - 2
'         .Row = iLoop
'         .Col = 1: strSku = .Text
'         str_SQL = "Select *" & _
'                   "From BaseData_SKUPacking Where StorerKey = '" & cmbStorerkey.Text & "' and SKU = '" & strSku & "'"
'         tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'         If tmp_Rs.EOF Then
'            If strErrorSKU = "" Then
'               strErrorSKU = strSku
'            Else
'               strErrorSKU = strErrorSKU & "," & strSku
'            End If
'         Else
'            .Col = 3: .Text = tmp_Rs.Fields("品名").Value
''            .Col = 4: dbShipQty = .Text
''            .Col = 5: .Text = NumRound((dbShipQty / tmp_rs.Fields("板轉換").Value), 2)
''            .Col = 6: .Text = NumRound((dbShipQty * tmp_rs.Fields("單位材積").Value), 2)
''            .Col = 7: .Text = NumRound((dbShipQty * tmp_rs.Fields("單位重量").Value), 2)
'         End If
'         tmp_Rs.Close
'     Next iLoop
'End With
'End Sub

Private Sub cmd_Exit_Click(Index As Integer)
'離開
Unload Me
End Sub

Private Sub cmd_Modify_Click()
'修改
If Trim(txt_Extern.Text) = "" Then Exit Sub

'利豐不可以手動改單
If RTrim(cmbStorerkey) = "LLFA01" Then MsgBox "利豐貨主訂單，只允許修改到貨日與備註欄位，其他欄位請勿修改!", vbOKOnly + vbExclamation, "訂單維護"
'If RTrim(cmbStorerkey) = "LMBO01" Then MsgBox "毛寶貨主訂單，不可手動改單!", vbOKOnly + vbCritical, "訂單維護": Exit Sub

Call cmd_OrdersQuery_Click

fam_Orders.Enabled = True
fam_Orders.BackColor = &HFF8080
fam_OrderDetail.BackColor = &HFF8080

txt_Extern.Enabled = True
txt_OrderKey.Enabled = True
cmd_Modify.Enabled = False
cmd_Delete.Enabled = False
cmd_AddNew.Enabled = False
If txtType <> "刪單" Then cmd_Save.Enabled = True
cmd_Cancel.Enabled = True
txt_QueryExternOrderKey.Enabled = False
cmbStorerkey.Enabled = False
txt_OtQty.Enabled = True

'訂單細項編輯功能設定
cmd_DetailAddNew.Enabled = True
If intsrcSKUNowRow <> 0 Then
   cmd_DetailModify.Enabled = True
Else
   cmd_DetailModify.Enabled = False
End If
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False
Call cmbStorerkey_Click
dgMain.Enabled = True

'毛寶只能修改header，明細要鎖住
If RTrim(cmbStorerkey) = "LMBO01" Then dgMain.Enabled = False
End Sub

Private Sub cmd_OrdersQuery_Click()

'訂單查詢
fam_Orders.BackColor = &H8000000A
fam_OrderDetail.BackColor = &H8000000A
fam_Orders.Enabled = False
cmd_Modify.Enabled = False
cmd_AddNew.Enabled = True
cmd_Delete.Enabled = False
cmd_Save.Enabled = False
cmd_Cancel.Enabled = False
txt_QueryExternOrderKey.Enabled = True

fam_DetailData.Enabled = False
dgMain.Enabled = False
cmd_DetailAddNew.Enabled = False
cmd_DetailModify.Enabled = False
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False
cmbStorerkey.Enabled = True

txt_QueryExternOrderKey.Text = Trim(txt_QueryExternOrderKey.Text)
If Len(Trim(txt_QueryExternOrderKey.Text)) = 0 Then
   msg_text = "請先輸入查詢條件：[TMS單號]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_QueryExternOrderKey.SetFocus
   Exit Sub
End If

Dim strTmp As String
strTmp = Format(txt_QueryExternOrderKey.Text, "0000000000")

'清除所有欄位值，包含 OrderDetail 部份
Call Clear_AllField

txt_QueryExternOrderKey.Text = strTmp

On Error GoTo err_Handle
Screen.MousePointer = vbHourglass

'取得訂單 Header
'str_SQL = "Select 貨主單號,貨主,訂單日,送貨日,客戶編號,說明,客戶名稱,客戶簡稱,郵遞區號,運送區域,特殊需求1,特殊需求2,運送地址,聯絡人,電話,TMS單號,b_phone2,客戶單號,訂單類別,轉運到貨客戶編號,訂單狀態,配送倉別 " & _
'          "From ManualOrder_Orders Where TMS單號 = '" & txt_QueryExternOrderKey.Text & "'"

str_SQL = "Select Rtrim(a1.ExternOrderKey) As 貨主單號 " & _
            ",Rtrim(a1.StorerKey) as 貨主 " & _
            ",Convert(varchar(8),a1.OrderDate,112) as 訂單日 " & _
            ",Convert(varchar(8),a1.DeliveryDate,112) as 送貨日 " & _
            ",Rtrim(a1.ConsigneeKey) as 客戶編號 " & _
            ",Isnull(Rtrim(Cast(a1.Notes as varchar(300))),'') as 說明 " & _
            ",Case When b1.ConsigneeKey is null Then Rtrim(Isnull(a1.C_Company,'')) else Rtrim(Isnull(b1.Full_Name,'')) End as 客戶名稱 " & _
            ",Case When b1.ConsigneeKey is null Then '' else Rtrim(Isnull(b1.Short_Name,'')) End as 客戶簡稱 " & _
            ",Case When b1.ZIP is not null Then Rtrim(b1.Zip) else Rtrim(Isnull(a1.C_Zip,'')) End as 郵遞區號 " & _
            ",Rtrim(Isnull(b1.Area_Code,'')) as 運送區域 " & _
            ",Rtrim(Isnull(b1.Extra_Demand_Code,'')) as 特殊需求1 " & _
            ",Rtrim(Isnull(b1.Extra_Demand_Code2,'')) as 特殊需求2 " & _
            ",Case When b1.Address is not null then Rtrim(b1.Address) else Rtrim(Isnull(a1.C_Address1,'')) +Rtrim(Isnull(a1.C_Address2,'')) End as 運送地址 " & _
            ",Case When b1.Contact is not null Then Rtrim(b1.Contact) else Rtrim(Isnull(a1.C_Contact1,'')) End as 聯絡人 " & _
            ",Case When b1.phone is not null Then Rtrim(b1.phone) else Rtrim(Isnull(a1.C_Phone1,'')) End as 電話 " & _
            ",a1.OrderKey as TMS單號 " & _
            ",B_phone2 = isnull(B_phone2,'') " & _
            ",客戶單號 = rtrim(a1.customerorderkey) " & _
            ",訂單類別 = rtrim(a1.priority) " & _
            ",轉運到貨客戶編號 = rtrim(isnull(a1.b_company,'')) " & _
            ",訂單狀態 = rtrim(isnull(a1.type,'')) " & _
            ",配送倉別 = rtrim(isnull(a1.facility,'')) " & _
            ",組單類別 = rtrim(isnull(a1.b_city,''))" & _
            ",件數 = case when isnull(cast(a1.otqty as char),'') = '' then '' else rtrim(cast(a1.otqty as char)) end ,順收 = isnull(a1.GoodsBack,0) " & _
            "From Orders a1 Left outer join TRP01M b1 on b1.ConsigneeKey = a1.ConsigneeKey and b1.storerkey = a1.storerkey where a1.OrderKey = '" & txt_QueryExternOrderKey.Text & "'"

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "查詢結果：無符合設定條件之訂單資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmd_Modify.Enabled = False
   Screen.MousePointer = vbDefault
   Exit Sub
End If
'if tmp_rs("b_phone2") = 00 then label1.Caption ="已轉入排車系統，無法變更!!"
txt_Extern.Text = tmp_Rs.Fields("貨主單號").Value
txt_OrderKey.Text = tmp_Rs.Fields("客戶單號").Value
cmbStorerkey.Text = tmp_Rs.Fields("貨主").Value
txt_OrderDate.Text = tmp_Rs.Fields("訂單日").Value
txt_DeliveryDate.Text = tmp_Rs.Fields("送貨日").Value
txt_ConsigneeKey.Text = tmp_Rs.Fields("客戶編號").Value
txt_Description.Text = tmp_Rs.Fields("說明").Value
txt_FullName.Text = tmp_Rs.Fields("客戶名稱").Value
txt_Contact.Text = tmp_Rs.Fields("聯絡人").Value
txt_Phone.Text = tmp_Rs.Fields("電話").Value
txt_Address.Text = tmp_Rs.Fields("運送地址").Value
cbo_Priority.Text = tmp_Rs.Fields("訂單類別").Value
txtShipToKey.Text = tmp_Rs("轉運到貨客戶編號") & ""
txtType.Text = tmp_Rs("訂單狀態")
cmdFacility.Text = tmp_Rs("配送倉別")
txt_B_city.Text = tmp_Rs("組單類別")
txt_OtQty.Text = tmp_Rs("件數")

If RTrim(tmp_Rs("順收")) = "1" Then Chk_receive.Value = 1 Else Chk_receive.Value = 0

If RTrim(cbo_Priority) = "A2B" Then Label3(21).Visible = True: txtShipToKey.Visible = True: cmdShipToList.Visible = True

If Len(RTrim(tmp_Rs("特殊需求1"))) > 0 Then
    For iLoop = 0 To cmb_ExtraDemand1.ListCount - 1
        If arExtraDemand(iLoop) = tmp_Rs.Fields("特殊需求1").Value Then
           cmb_ExtraDemand1.ListIndex = iLoop
           Exit For
        End If
    Next iLoop
End If

If Len(RTrim(tmp_Rs("特殊需求2"))) > 0 Then
    For iLoop = 0 To cmb_ExtraDemand2.ListCount - 1
        If arExtraDemand(iLoop) = tmp_Rs.Fields("特殊需求2").Value Then
           cmb_ExtraDemand2.ListIndex = iLoop
           Exit For
        End If
    Next iLoop
End If

txt_ShortName.Text = tmp_Rs.Fields("客戶簡稱").Value

For iLoop = 0 To cmb_ZIP.ListCount - 1
    If arZip(iLoop) = tmp_Rs.Fields("郵遞區號").Value Then
       cmb_ZIP.ListIndex = iLoop
       Exit For
    End If
Next iLoop
DoEvents: DoEvents

For iLoop = 0 To cmb_AreaCode.ListCount - 1
    If arAreaCode(iLoop) = tmp_Rs.Fields("運送區域").Value Then
       cmb_AreaCode.ListIndex = iLoop
       Exit For
    End If
Next iLoop
tmp_Rs.Close

'取得訂單 Detail >> 以 OrderDetail 為主
'str_SQL = "exec gs_ManualOrder_OrderDetail '" & txt_QueryExternOrderKey.Text & "' "

str_SQL = "Select 項次 = isnull(od.OrderLineNumber,'') " & _
            ",貨號 = Rtrim(od.SKU),品名 = Rtrim(Isnull(s.Descr,'')),倉別 = rtrim(isnull(od.lottable06,'')) " & _
            ",大單位數量 = case when s.casecnt=0 then 0  else floor(od.originalQty/convert(int,s.casecnt)) end " & _
            ",大單位名稱 = rtrim(isnull(s.busr3,'')) " & _
            ",中單位數量 = case when s.casecnt=0 then case when s.InnerPack=0 then 0 else floor(convert(int,od.originalQty)/convert(int,s.InnerPack)) end else case when s.InnerPack=0  then 0 else floor((convert(int,od.originalQty)%convert(int,s.casecnt))/convert(int,s.InnerPack)) end end " & _
            ",中單位名稱 = rtrim(isnull(s.busr2,'')) " & _
            ",小單位數量 = case when s.casecnt=0 then  " & _
                            "case when s.InnerPack=0 then od.originalQty  " & _
                            "else convert(int,od.originalQty)%convert(int,s.InnerPack) end  " & _
                            "else case when s.InnerPack=0  then convert(int,od.originalQty)%convert(int,s.casecnt)  " & _
                            "else (convert(int,od.originalQty)%convert(int,s.casecnt))%convert(int,s.InnerPack) end End  " & _
            ",小單位名稱 = rtrim(isnull(s.busr1,'')),大單位入數 = s.casecnt " & _
            ",中單位入數 = s.innerpack,生產批號 = rtrim(isnull(lottable03,'')) " & _
            ",製造日 = rtrim(isnull(convert(char(8),lottable04,112),'')) " & _
            ",到期日 = rtrim(isnull(convert(char(8),lottable05,112),'')) " & _
            ",揀貨加工 = rtrim(isnull(updatesource,'')) " & _
            ",備註 = rtrim(isnull(od.notes,'')) " & _
            "From OrderDetail od (nolock) join gv_skuxpack s (nolock) on s.StorerKey = od.StorerKey and s.SKU = od.SKU where od.orderkey = '" & txt_QueryExternOrderKey.Text & "' " & _
            "Order by od.OrderLineNumber"

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "查詢結果：無符合設定條件之訂單明細資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmd_Modify.Enabled = False
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Set rsMain = New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsMain)
tmp_Rs.Close

rsMain.MoveFirst
Set dgMain.DataSource = rsMain

SetDataGridColWidth Me.Caption, dgMain

'Do While Not tmp_Rs.EOF
'   With dg_OrderDetail
'       .Rows = .Rows + 1
'       .Row = .Rows - 2
'       .Col = 0: .Text = Rtrim(tmp_Rs("項次"))
'       .Col = 1: .Text = tmp_Rs("貨號")
'       .Col = 2: .Text = tmp_Rs("品名")
'       .Col = 3: .Text = tmp_Rs("倉別") & ""
'       .Col = 4: .Text = tmp_Rs("箱數")
'       .Col = 5: .Text = tmp_Rs("個數")
'       .Col = 6: .Text = tmp_Rs("每箱")
'       .Col = 7: .Text = tmp_Rs("備註") & ""
'       .Col = 8: .Text = tmp_Rs("生產批號") & ""
'       .Col = 9: .Text = tmp_Rs("製造日") & ""
'       .Col = 10: .Text = tmp_Rs("到期日") & ""
'
'  End With
'  tmp_Rs.MoveNext
'Loop
'tmp_Rs.Close

If txtType <> "刪單" Then cmd_Modify.Enabled = True
If txtType <> "刪單" Then cmd_Delete.Enabled = True
cmd_AddNew.Enabled = True
cmd_Cancel.Enabled = False

'訂單細項功能鎖定
'fam_DetailData.Enabled = False
'cmd_DetailModify.Enabled = False
'cmd_DetailAddNew.Enabled = False
'cmd_DetailSave.Enabled = False
'cmd_DetailDel.Enabled = False
'fam_DetailData.Enabled = False
'intsrcSKUNowRow = 0

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單維護-訂單查詢", Me.Caption, "cmd_OrdersQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub
Sub SendMail(strOrderkey As String)

''LTKK01刪單自動 Mail 通知
'If mySplit(cmbStorerkey, " ", 0) = "LTKK01" Then
'
'    Dim strFrom As String, strTo As String, strCC As String, strBCC As String, strSubject As String, strTextbody As String, strAlways As String, strEmailID As String, strEmailPW As String, strAddAttachment As String
'
'    '讀取ini參數
'    Dim objIni As New vbIniFile
'    objIni.FileName = App.Path & "/" & App.title & ".ini"
'
'    strFrom = objIni.ReadData("INVCHECKEMAIL_LTKK01", "From", "")
'    strTo = objIni.ReadData("INVCHECKEMAIL_LTKK01", "To", "")
'    strCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "CC", "")
'    strBCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "BCC", "")
'    strSubject = "訂單刪除明細"
'    strTextbody = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Textbody", "")
'    strEmailID = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailID", "")
'    strEmailPW = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailPW", "")
'    strAlways = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Always", "NO")
'
'    '直接指定
'    strFrom = "Tkedi@bestlog.com.tw"
'    strTo = "jack@mail.kirin.com.tw,irene@mail.kirin.com.tw;ken@mail.kirin.com.tw;shiu@mail.kirin.com.tw;celine@mail.kirin.com.tw;simon@mail.kirin.com.tw"
'    strCC = "Tkedi@bestlog.com.tw"
'    strBCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "BCC", "")
''    strSubject = "訂單刪除明細"
'    strTextbody = "此為系統發送信件!!"
'    strEmailID = "tkedi"
'    strEmailPW = "tkedibl01"
'    strAlways = "NO"
'
'    Set objIni = Nothing
'
'    Dim rsTmp As New ADODB.Recordset
'
'    If Len(RTrim(strFrom)) > 0 Then '有寄件者
'
'        str_SQL = "select 倉別 = 'BL01' " & _
'                ",貨主代碼 = rtrim(o.storerkey) " & _
'                ",訂單號碼交貨單號 = rtrim(od.externorderkey) + rtrim(od.externlineno) " & _
'                ",地址別 = substring(o.consigneekey,5,20) " & _
'                ",料號 = isnull(ss.storersku,od.sku) " & _
'                ",倉別_儲位別 = 'BL01_'+ od.lottable06 " & _
'                ",最小單位數量 = isnull(od.originalqty,0) ,訂單日 = convert(varchar,o.orderdate,111) " & _
'                ",預計到貨日 =  convert(varchar,o.deliverydate,111) " & _
'                ",刪單日 = convert(varchar,o.editdate,111) " & _
'                ",客戶訂單號碼 = rtrim(o.customerorderkey) " & _
'                "From orders o join orderdetail od on o.orderkey = od.orderkey " & _
'                "left join storersku ss on ss.sku = od.sku and ss.storerkey = od.storerkey " & _
'                "Where o.type = '刪單' and o.orderkey = '" & strOrderkey & "' "
'
'        rsTmp.Open str_SQL, cn
'
'        '如果無資料也要mail
'        If Not rsTmp.EOF Or UCase(RTrim(strAlways)) = "YES" Then
'
'            strAddAttachment = "C:\LTKK01\訂單刪除明細\訂單刪除明細_" & Format(Now, "yyyymmddhhMMss") & ".xls"
'
'            Call Recordset2Excel("訂單刪除明細", rsTmp)
'            If Dir("C:\LTKK01\訂單刪除明細", vbDirectory) = "" Then MkDirs "C:\LTKK01\訂單刪除明細"
'            MyXlsApp.ActiveWorkbook.SaveAs strAddAttachment
'            MyXlsApp.Quit: Set MyXlsApp = Nothing
'
'            '傳送郵件
'            Dim objEmail As Object
'            Set objEmail = CreateObject("CDO.Message")
'
'            objEmail.From = strFrom
'            objEmail.To = strTo
'            objEmail.CC = strCC   ' 副本
'            objEmail.BCC = strBCC ' 密件副本
'            objEmail.Subject = strSubject
'            objEmail.TextBody = strTextbody
'            objEmail.AddAttachment strAddAttachment
'
'            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "bestlog.com.tw"
'            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
'            'SMTP 伺服器需要驗證時
'            If Len(RTrim(strEmailID)) > 0 Then
'                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendusername") = strEmailID
'                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strEmailPW
'            End If
'            objEmail.Configuration.Fields.Update
'            objEmail.Send
'
'            MsgBox "LTKK01刪單明細資料，系統已發Mail通知!", , "刪單明細資料"
'
'            Set objEmail = Nothing
'        End If
'    End If
'End If

End Sub
Private Sub cmd_Save_Click()
On Error GoTo err_Handle

'利豐不可以手動改單
'If RTrim(cmbStorerkey) = "LLFA01" Then MsgBox "利豐貨主訂單，不可手動改單!", vbOKOnly + vbCritical, "訂單維護": Exit Sub
'If RTrim(cmbStorerkey) = "LMBO01" Then MsgBox "毛寶貨主訂單，不可手動改單!", vbOKOnly + vbCritical, "訂單維護": Exit Sub

'Terry 20181220 user要求取消防呆
''檢查組單類別，目前只有亞培貨主再使用
'If Len(RTrim(txt_B_city)) > 0 And RTrim(Left(cmbStorerkey.Text, 6)) <> "LABT01" Then
'    MsgBox "組單類別目前只有亞培貨主使用，無法存檔。請確認檔案", vbOKOnly + vbCritical, "訂單維護": Exit Sub
'End If

'檢查件數是否為數字
If Len(RTrim(txt_OtQty.Text)) > 0 And Not IsNumeric(txt_OtQty.Text) Then MsgBox "件數欄位請輸入數字", vbOKOnly + vbCritical, "注意": txt_OtQty.SelStart = 0: txt_OtQty.SelLength = Len(txt_OtQty.Text): txt_OtQty.SetFocus: Exit Sub


rsMain.MoveFirst
Do While Not rsMain.EOF

If Trim(cbo_Priority) = "A2B" Then Exit Do
'配送倉別檢查
If Len(Trim(cmdFacility)) = 0 Then
    cmdFacility = "佰事達北倉"
    If UCase(Right(Trim(rsMain("倉別")), 2)) = "-C" Then cmdFacility = "佰事達中倉"
    If UCase(Right(Trim(rsMain("倉別")), 2)) = "-S" Then cmdFacility = "佰事達南倉"
Else
    If UCase(Right(Trim(rsMain("倉別")), 2)) <> "-C" And UCase(Right(Trim(rsMain("倉別")), 2)) <> "-S" And cmdFacility <> "佰事達北倉" Then: MsgBox "配送倉別與商品倉別不符!!", 64, "倉別錯誤!!": Exit Sub
    If UCase(Right(Trim(rsMain("倉別")), 2)) = "-C" And cmdFacility <> "佰事達中倉" Then: MsgBox "配送倉別與商品倉別不符!!", 64, "倉別錯誤!!": Exit Sub
    If UCase(Right(Trim(rsMain("倉別")), 2)) = "-S" And cmdFacility <> "佰事達南倉" Then: MsgBox "配送倉別與商品倉別不符!!", 64, "倉別錯誤!!": Exit Sub

End If

rsMain.MoveNext
Loop

cmd_Save.Enabled = False
'清除特殊字元
Call myFormExCharFilter(Me)

'檢核資料正確
If CheckOrdersData() = False Then cmd_Save.Enabled = True: Exit Sub

Tran_Level = cn.BeginTrans

'標記舊訂單為刪單
If txt_QueryExternOrderKey.Enabled = False And txtType <> "刪單" Then
     
'    Dim rsTmp As New ADODB.Recordset
'    rsTmp.Open "select route_no from trp02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' union select route_no from ort02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' ", cn
'    If Not rsTmp.EOF Then MsgBox "此訂單已排路線編號 " & rsTmp("route_no") & " ，無法修改!", vbOKOnly, "存檔": rsTmp.Close: Exit Sub
        
    Call DB_CheckConnectStatus
    
         cn.Execute "delete TRP03T where receipt_no in (select receipt_no from trp02t where route_no = 'D' and c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
         cn.Execute "delete TRP02T where route_no = 'D' and c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         
         cn.Execute "delete TRP03W where receipt_no in (select receipt_no from trp02w where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
         cn.Execute "delete TRP02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords

         cn.Execute "delete TRP02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         cn.Execute "delete ORT02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         
         cn.Execute "delete ORT03W where receipt_no in (select receipt_no from ORT02W where c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
         cn.Execute "delete ORT02W where c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         
         cn.Execute "delete ORT03T where receipt_no in (select receipt_no from ort02t where route_no = 'D' and c_receipt_no = '" & txt_QueryExternOrderKey & "') ", RowsAffect, adExecuteNoRecords
         cn.Execute "delete ORT02T where route_no = 'D' and c_receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         
         cn.Execute "delete TRP02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         cn.Execute "delete ORT02W_TEMP where receipt_no ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords

         cn.Execute "delete status where orderkey ='" & txt_QueryExternOrderKey & "'", RowsAffect, adExecuteNoRecords
         cn.Execute "update orders set B_PHONE2='00',trafficCop=null,type='刪單' ,editdate = getdate() , editwho= '" & User_id & "' where orderkey='" & txt_QueryExternOrderKey & "' ", RowsAffect, adExecuteNoRecords

If mySplit(cmbStorerkey, " ", 0) = "LTKK01" Then Call SendMail(txt_QueryExternOrderKey)

End If

'訂單資料存檔
If SaveOrdersData() = False Then
    cn.RollbackTrans: Tran_Level = 0
    MsgBox "存檔失敗！", 16, "錯誤"
    Exit Sub
Else
    cn.CommitTrans: Tran_Level = 0
    txtType = "刪單"
    MsgBox "訂單修改新增完成。", vbOKOnly, Me.Caption
End If

Call cmd_OrdersQuery_Click

'將緊急訂單標記在orders..urgent_mark欄位
'檢查是否有緊急訂單?
str_SQL = "select orderkey " & _
          "From orders " & _
          "where storerkey = 'LAPP01' and priority = 'I' and orderkey = '" & txt_QueryExternOrderKey & "' and type <> '刪單' and " & _
          "((convert(varchar(8),adddate,114) > '17:00:00' and convert(varchar(8),deliverydate,112) < = convert(varchar(8),getdate()+1,112) ) or " & _
          "(convert(varchar(8),adddate,114) > '17:30:00' and convert(varchar(8),deliverydate,112) < = convert(varchar(8),getdate()+2,112) ) or " & _
          "(convert(varchar(8),adddate,112) = convert(varchar(8),deliverydate,112)))"

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If Not tmp_Rs.EOF Then  '有回傳
    If MsgBox("發現緊急訂單，是否自動將訂單更新為緊急訂單?", vbQuestion + vbYesNo, "訂單維護") = vbYes Then
           '更新urgent_mark欄位V:緊急訂單
           cn.Execute "update orders set urgent_mark = 'V' where orderkey = '" & txt_QueryExternOrderKey & " ' ", RowsAffect, adExecuteNoRecords
    End If
End If

tmp_Rs.Close

fam_Orders.Enabled = False
fam_Orders.BackColor = &H8000000A
fam_OrderDetail.BackColor = &H8000000A
txt_Extern.Enabled = False
txt_OrderKey.Enabled = False
cmd_Modify.Enabled = True
cmd_AddNew.Enabled = True
cmd_Save.Enabled = False
cmd_Cancel.Enabled = False

fam_DetailData.Enabled = False
dgMain.Enabled = False
cmd_DetailModify.Enabled = False
cmd_DetailAddNew.Enabled = False
cmd_DetailSave.Enabled = False
cmd_DetailDel.Enabled = False
cmd_DetailCancel.Enabled = False
cmbStorerkey.Enabled = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdShipToList_Click()
'顯示客戶待選清單
'Call CallForm_BaseOP_DataList(Me.Name & "_" & cmd_Consigneelist.Name)

If cmbStorerkey = "" Then MsgBox "請先選取貨主！", 64, Me.Caption: Exit Sub

'呼叫視窗
strDataList_Caller = Me.Name & " " & "txtShipToKey" & " " & mySplit(cmbStorerkey, " ", 0)
frm_ConsigneekeyQuery.Show vbModal
End Sub

Private Sub Command1_Click()

'清除特殊字元
Call myFormExCharFilter(Me)

If Len(RTrim(cmbStorerkey)) = 0 Then MsgBox "請選擇貨主！", 64, Me.Caption: Exit Sub
If Len(RTrim(txt_ConsigneeKey)) = 0 Then MsgBox "請輸入客戶編號！", 64, Me.Caption: Exit Sub
If Len(RTrim(cmb_ZIP)) = 0 Then MsgBox "請輸入郵遞區號！", 64, Me.Caption: Exit Sub
If Len(RTrim(txt_FullName)) = 0 Then MsgBox "請輸入客戶名稱！", 64, Me.Caption: Exit Sub

Dim rsTmp As New ADODB.Recordset
rsTmp.Open "select * from trp01m where storerkey = '" & mySplit(cmbStorerkey, " ", 0) & "' and consigneekey = '" & RTrim(txt_ConsigneeKey) & "' ", cn

If rsTmp.EOF Then

    str_SQL = "insert into trp01m(storerkey,consigneekey,area_code,zip,full_name,short_name,contact,phone,address,extra_demand_code,extra_demand_code2,addwho,editwho,updatesource) values('" & _
               mySplit(cmbStorerkey, " ", 0) & "','" & txt_ConsigneeKey & "','" & mySplit(cmb_AreaCode, " ", 0) & "','" & mySplit(cmb_ZIP, " ", 0) & "','" & txt_FullName & "','" & txt_ShortName & "','" & txt_Contact & "','" & txt_Phone & "','" & txt_Address & "','" & mySplit(cmb_ExtraDemand1, " ", 0) & "','" & mySplit(cmb_ExtraDemand1, " ", 0) & "','" & User_id & "','" & User_id & "','Manual') "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

Else
    MsgBox "客戶編號重複，請確認！", 64, Me.Caption

End If

rsTmp.Close

End Sub

Private Sub dg_OrderDetail_Click()
cboSku = "": txt_SkuDescr = "": txtOrderCS.Text = "0": txtOrderEA.Text = "0": txtCasecnt.Text = "0": txtNotes.Text = "": txtLot3.Text = "": txtLot4.Text = "": txtLot5.Text = "": cboLot6.Text = ""

If fam_Orders.Enabled Then
   cmd_DetailModify.Enabled = False
   cmd_DetailAddNew.Enabled = True
   cmd_DetailSave.Enabled = False
   cmd_DetailDel.Enabled = False
   cmd_DetailCancel.Enabled = False
   fam_DetailData.Enabled = False
End If
With dg_OrderDetail
     intsrcSKUNowRow = 0
     .Col = 0    '項次
     If Len(.Text) = 0 Then Exit Sub
     If .Row = 0 Then Exit Sub
     .Col = 1: cboSku = .Text
     .Col = 2: txt_SkuDescr = .Text
     .Col = 3: cboLot6.Text = .Text
     .Col = 4: txtOrderCS.Text = .Text
     .Col = 5: txtOrderEA.Text = .Text
     .Col = 6: txtCasecnt.Text = .Text
     .Col = 7: txtNotes.Text = .Text
     .Col = 8: txtLot3.Text = .Text
     .Col = 9: txtLot4.Text = .Text
     .Col = 10: txtLot5.Text = .Text
     
     intsrcSKUNowRow = .Row
     .Col = 0
     For iLoop = 0 To .Cols - 1
         .ColSel = iLoop
     Next iLoop
     If fam_Orders.Enabled Then
        cmd_DetailModify.Enabled = True
        cmd_DetailAddNew.Enabled = True
        cmd_DetailSave.Enabled = False
        cmd_DetailDel.Enabled = True
        cmd_DetailCancel.Enabled = False
        fam_DetailData.Enabled = False
    End If
End With

End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "訂單維護作業"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'攔截整個表單鍵盤按鍵事件
'用途：使用者按下 Esc 則不傳回任何資料，且關閉日期選取視窗
If KeyCode = vbKeyEscape Then
   mvDate.Visible = False
End If

End Sub

Private Sub Form_Load()
'設定 Form 大小、位置
dbsrcFormHeight = 7140
dbsrcFormWidth = 11475
Me.Height = 7650: Me.Width = 11600
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

'設定訂單明細格式
Call SetGDFormat_OrderDetail

Dim tmp_cnt As Double
'取出所有郵遞區號 TRP02M
cmb_ZIP.Clear
str_SQL = "Select Rtrim(ZIP) as 'ZIP',Rtrim(Area_Code) as 'AreaCode',Rtrim(Isnull(Description,'')) as Descr  From TRP02M Order by ZIP"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arZip(1) As String
ReDim arZIPArea(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arZip(tmp_cnt) = tmp_Rs.Fields("ZIP").Value
      arZIPArea(tmp_cnt) = tmp_Rs.Fields("AreaCode").Value
      cmb_ZIP.AddItem tmp_Rs.Fields("ZIP").Value & Space(5 - Len(Trim(tmp_Rs.Fields("ZIP").Value))) & tmp_Rs.Fields("Descr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arZip) Then
         ReDim Preserve arZip(UBound(arZip) + 10) As String
         ReDim Preserve arZIPArea(UBound(arZIPArea) + 10) As String
      End If
   Loop
End If

'取出所有運送區域代碼 TRP03M
cmb_AreaCode.Clear
str_SQL = "Select Rtrim(Area_Code) as 'AreaCode',Rtrim(Isnull(Description,'')) as Descr  From TRP03M Order by Area_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arAreaCode(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arAreaCode(tmp_cnt) = tmp_Rs.Fields("AreaCode").Value
      cmb_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arAreaCode) Then
         ReDim Preserve arAreaCode(UBound(arAreaCode) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close

'取出所有特殊需求--TRP04M
cmb_ExtraDemand1.Clear: cmb_ExtraDemand2.Clear
str_SQL = "Select Rtrim(Extra_Demand_Code) as 'ECode',Isnull(Rtrim(Description),'') as 'ECodeDescr' From TRP04M Order by Extra_Demand_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arExtraDemand(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arExtraDemand(tmp_cnt) = tmp_Rs.Fields("ECode").Value
      cmb_ExtraDemand1.AddItem tmp_Rs.Fields("ECode").Value & Space(12 - Len(Trim(tmp_Rs.Fields("ECode").Value))) & tmp_Rs.Fields("ECodeDescr").Value
      cmb_ExtraDemand2.AddItem tmp_Rs.Fields("ECode").Value & Space(12 - Len(Trim(tmp_Rs.Fields("ECode").Value))) & tmp_Rs.Fields("ECodeDescr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arExtraDemand) Then
         ReDim Preserve arExtraDemand(UBound(arExtraDemand) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close

'取出所有貨主
str_SQL = "Select Rtrim(Storerkey) + ' ' + rtrim(short_name) as 'Storer' From TRP16M where storerkey <> 'LTKK01' Order by Storerkey "
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn
If Not tmp_Rs.EOF Then

   Do While Not tmp_Rs.EOF
      cmbStorerkey.AddItem tmp_Rs("Storer")
      tmp_Rs.MoveNext
   Loop
End If
cmbStorerkey.ListIndex = -1
tmp_Rs.Close

cbo_Priority.AddItem "I 出貨"
cbo_Priority.AddItem "R 退貨"
cbo_Priority.AddItem "A 轉倉"
cbo_Priority.AddItem "A2B 提貨配送"
cbo_Priority.AddItem "RC 提貨入庫"
cbo_Priority.AddItem "C 越庫"

'cbo_Priority.AddItem "RS 退貨品出庫"

txt_B_city.AddItem ""
txt_B_city.AddItem "杏一"
txt_B_city.AddItem "維康"

cbo_Priority = ""

End Sub

Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub

'fam_OrderDetail.Width = Me.ScaleWidth - 60
'dg_OrderDetail.Width = fam_OrderDetail.Width - 180
'
'fam_OrderDetail.Height = Me.ScaleHeight - fam_Orders.Height - fam_Header.Height ' - 360
'dg_OrderDetail.Height = fam_OrderDetail.Height - fam_DetailData.Height - 240

dgMain.Width = Me.ScaleWidth - 60
dgMain.Height = Me.ScaleHeight - fam_Orders.Height - fam_Orders.Top - fam_Header.Height - fam_Header.Top

End Sub

Private Sub Form_Terminate()
'更新 Menu [視窗]→[已開視窗清單]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub SetGDFormat_OrderDetail()
'名稱：SetGDFormat_OrdereDtail
'類別：副程式
'功能：清除並設定 [訂單名細資料] 顯示格式
'參數：傳入值：無
Dim sub_var1 As Integer, sub_var2 As Integer
dg_OrderDetail.Visible = False
With dg_OrderDetail
     .FixedRows = 1: .Cols = 11
     '設定允許整列選取
     .AllowBigSelection = True
     '設定列表之文字字型
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "新細明體": .CellFontSize = 9
         Next sub_var2
     Next sub_var1

     '設定列表之格式
     .Row = 0
     .Col = 0: .Text = "項次": .ColWidth(0) = 1000: .ColAlignment(0) = flexAlignLeftCenter
     .Col = 1: .Text = "貨號": .ColWidth(1) = 2400: .ColAlignment(1) = flexAlignLeftCenter
     .Col = 2: .Text = "品名": .ColWidth(2) = 3000: .ColAlignment(2) = flexAlignLeftCenter
     .Col = 3: .Text = "倉別": .ColWidth(3) = 600: .ColAlignment(3) = flexAlignCenterCenter
     .Col = 4: .Text = "箱數": .ColWidth(4) = 600: .ColAlignment(4) = flexAlignRightCenter
     .Col = 5: .Text = "個數": .ColWidth(5) = 600: .ColAlignment(5) = flexAlignRightCenter
     .Col = 6: .Text = "每箱": .ColWidth(6) = 600: .ColAlignment(6) = flexAlignRightCenter
     .Col = 7: .Text = "備註": .ColWidth(7) = 3000: .ColAlignment(7) = flexAlignLeftCenter
     .Col = 8: .Text = "生產批號": .ColWidth(8) = 1200: .ColAlignment(8) = flexAlignLeftCenter
     .Col = 9: .Text = "製造日": .ColWidth(9) = 800: .ColAlignment(9) = flexAlignLeftCenter
     .Col = 10: .Text = "到期日": .ColWidth(10) = 800: .ColAlignment(10) = flexAlignLeftCenter

     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_OrderDetail.Visible = True
End Sub

Private Sub Clear_AllField()
'清除所有欄位值
'txt_Extern.Text = ""
cmbStorerkey = ""
cboSku = ""
'cbo_Priority.ListIndex = 0
'txt_OrderDate.Text = ""
'txt_DeliveryDate.Text = ""
'txt_Description.Text = ""
'txt_OrderKey.Text = ""
'txt_ConsigneeKey.Text = ""
'txt_FullName.Text = "": txt_ShortName.Text = ""
'txt_Contact.Text = "": txt_Phone.Text = ""
'txt_Address.Text = "": cmb_AreaCode.ListIndex = -1: cmb_ZIP.ListIndex = -1
'cmb_ExtraDemand1.ListIndex = -1: cmb_ExtraDemand2.ListIndex = -1
cmdFacility = ""
txt_B_city = ""
txt_OtQty = ""
Call ClearForm_AllField(frm_OP_ManualOrders)
fam_ConsigneeData.ZOrder 0

'設定訂單明細格式
Call SetGDFormat_OrderDetail
cboSku.ListIndex = -1: txtOrderCS.Text = "0": txtOrderEA.Text = "0": txtCasecnt.Text = "0": txtNotes.Text = "": cboLot6.Text = ""
End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)
'日期選取
Select Case mvDate.Tag
       Case "訂單日"
            txt_OrderDate.Text = Format(mvDate.Value, "yyyymmdd")
       Case "送貨日"
            txt_DeliveryDate.Text = Format(mvDate.Value, "yyyymmdd")
End Select
mvDate.Visible = False
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text3_Change()

End Sub

Private Sub txt_ConsigneeKey_GotFocus()
    fam_ConsigneeData.ZOrder 0
    fam_ConsigneeData.Enabled = True
End Sub

Public Sub txt_ConsigneeKey_LostFocus()

'客戶編號
fam_ConsigneeData.Enabled = False
If Trim(txt_ConsigneeKey.Text) = "" Then Exit Sub
txt_ConsigneeKey = myExCharFilter(txt_ConsigneeKey.Text)

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
str_SQL = "Select * From TRP01M Where storerkey = '" & mySplit(cmbStorerkey, " ", 0) & "' and ConsigneeKey = '" & txt_ConsigneeKey.Text & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "資料錯誤：客戶編號 [" & txt_ConsigneeKey.Text & "] 未建檔"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Sub
End If
txt_FullName.Text = IIf(IsNull(tmp_Rs.Fields("Full_Name").Value), "", Trim(tmp_Rs.Fields("Full_Name").Value))
txt_ShortName.Text = IIf(IsNull(tmp_Rs.Fields("Short_Name").Value), "", Trim(tmp_Rs.Fields("Short_Name").Value))
txt_Contact.Text = IIf(IsNull(tmp_Rs.Fields("Contact").Value), "", Trim(tmp_Rs.Fields("Contact").Value))
txt_Phone.Text = IIf(IsNull(tmp_Rs.Fields("Phone").Value), "", Trim(tmp_Rs.Fields("Phone").Value))
txt_Address.Text = IIf(IsNull(tmp_Rs.Fields("Address").Value), "", Trim(tmp_Rs.Fields("Address").Value))
cmb_ZIP.ListIndex = -1
For iLoop = 0 To cmb_ZIP.ListCount - 1
   If arZip(iLoop) = Trim(tmp_Rs.Fields("ZIP").Value) Then
      cmb_ZIP.ListIndex = iLoop
      Exit For
   End If
Next iLoop
cmb_AreaCode.ListIndex = -1
For iLoop = 0 To cmb_AreaCode.ListCount - 1
    If arAreaCode(iLoop) = Trim(tmp_Rs.Fields("Area_Code").Value) Then
       cmb_AreaCode.ListIndex = iLoop
       Exit For
    End If
Next iLoop
cmb_ExtraDemand1.ListIndex = -1
For iLoop = 0 To cmb_ExtraDemand1.ListCount - 1
    If arExtraDemand(iLoop) = Trim(tmp_Rs.Fields("Extra_Demand_Code").Value) Then
       cmb_ExtraDemand1.ListIndex = iLoop
       Exit For
    End If
Next iLoop
cmb_ExtraDemand2.ListIndex = -1
For iLoop = 0 To cmb_ExtraDemand2.ListCount - 1
    If arExtraDemand(iLoop) = Trim(tmp_Rs.Fields("Extra_Demand_Code2").Value) Then
       cmb_ExtraDemand2.ListIndex = iLoop
       Exit For
    End If
Next iLoop
tmp_Rs.Close
End Sub

Private Sub txt_DeliveryDate_Click()
'送貨日
If Trim(txt_DeliveryDate.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_DeliveryDate.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_DeliveryDate.Text, 4) & "/" & Mid(txt_DeliveryDate.Text, 5, 2) & "/" & Right(txt_DeliveryDate.Text, 2))
   End If
End If
mvDate.Top = fam_Orders.Top + txt_DeliveryDate.Top + txt_DeliveryDate.Height
mvDate.Left = fam_Orders.Left + txt_DeliveryDate.Left + txt_DeliveryDate.Width
mvDate.Tag = "送貨日"
mvDate.Value = Now
mvDate.Visible = True
End Sub

Private Sub txt_OrderDate_Click()
'訂單日
If Trim(txt_OrderDate.Text) = "" Then
   mvDate.Value = Now
Else
   If Fun_ChkDateFormat(txt_OrderDate.Text) = 1 Then
      mvDate.Value = Now
   Else
      mvDate.Value = CDate(Left(txt_OrderDate.Text, 4) & "/" & Mid(txt_OrderDate.Text, 5, 2) & "/" & Right(txt_OrderDate.Text, 2))
   End If
End If
mvDate.Top = fam_Orders.Top + txt_OrderDate.Top + txt_OrderDate.Height
mvDate.Left = fam_Orders.Left + txt_OrderDate.Left + txt_OrderDate.Width
mvDate.Tag = "訂單日"
mvDate.Value = Now
mvDate.Visible = True
End Sub

Private Sub txtLot_Change()

End Sub

Private Sub txtOrderCS_GotFocus()
txtOrderCS.SelStart = 0: txtOrderCS.SelLength = Len(txtOrderCS.Text)
End Sub

Private Sub txtOrdercs_KeyPress(KeyAscii As Integer)
'訂單細項 >> 訂單量
Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          txtOrderEA.SelStart = 0: txtOrderEA.SelLength = Len(txtOrderEA.Text): txtOrderEA.SetFocus
End Select

End Sub

Private Sub txt_QueryExternOrderKey_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then Call cmd_OrdersQuery_Click

End Sub

Private Sub txtOrderEA_GotFocus()
txtOrderEA.SelStart = 0: txtOrderEA.SelLength = Len(txtOrderEA.Text)
End Sub

Private Sub txtOrderea_KeyPress(KeyAscii As Integer)
'訂單細項 >> 訂單量
Select Case KeyAscii
     Case 97 To 122, 65 To 90   '不允許輸入字元
          KeyAscii = 0
     Case vbKeyReturn
          cboLot6.SetFocus
End Select
End Sub

Private Sub txt_SKU_KeyPress(KeyAscii As Integer)
'訂單細項輸入 >> 貨號
If KeyAscii = vbKeyReturn Then
   txtOrderCS.SelStart = 0: txtOrderCS.SelLength = Len(txtOrderCS.Text): txtOrderCS.SetFocus
ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then '小寫字元改為大寫字元
       KeyAscii = KeyAscii - 32
End If
End Sub

Private Function CheckSKU(ByVal strSku As String) As Integer
'檢核貨號是否正確
CheckSKU = 1
If cmbStorerkey.Text = "" Then
   msg_text = "貨號檢核異常：尚未輸入 [貨主]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cmbStorerkey.SetFocus
   Exit Function
End If

'驗證貨號是否正確
str_SQL = "Select isnull(Rtrim(Descr),'') as 'Descr' From gv_SKUxpack Where StorerKey = '" & mySplit(cmbStorerkey.Text, " ", 0) & "' and SKU = '" & strSku & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "貨號驗證錯誤：Storer = [" & cmbStorerkey.Text & "] 無此貨號"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   cboSku.SetFocus
   Exit Function
End If
txt_SkuDescr.Text = tmp_Rs.Fields("Descr").Value
CheckSKU = 0
tmp_Rs.Close

End Function

Private Sub txt_StorerKey_KeyPress(KeyAscii As Integer)
'貨主
If KeyAscii >= 97 And KeyAscii <= 122 Then '小寫字元改為大寫字元
   KeyAscii = KeyAscii - 32
End If
End Sub

Private Function CheckOrdersData() As Boolean

'到貨日期檢查
rsMain.MoveFirst
Do While Not rsMain.EOF

    If Len(rsMain("製造日")) > 0 And Fun_ChkDateFormat(rsMain("製造日")) = 1 Then: MsgBox "製造日格式錯誤(YYYYMMDD)!", 16, Me.Caption: Exit Function
    If Len(rsMain("到期日")) > 0 And Fun_ChkDateFormat(rsMain("到期日")) = 1 Then: MsgBox "到期日格式錯誤(YYYYMMDD)!", 16, Me.Caption: Exit Function

    '資料檢驗--判斷SKU是否存在
    str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMain("貨號")) & "' and Storerkey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        MsgBox "貨號不存在 (" & Trim(rsMain("貨號")) & " )", 16, Me.Caption
        Exit Function
    End If

rsMain.MoveNext
Loop

'訂單維護作業執行存檔時，檢核訂單資料正確性
CheckOrdersData = False
If Trim(cmbStorerkey.Text) = "" Then
   msg_text = "資料錯誤：未輸入 [貨主] 資料"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
Else
   cmbStorerkey.Text = Trim(cmbStorerkey.Text)
End If
If Trim(cbo_Priority.Text) = "" Then
   msg_text = "資料錯誤：未輸入 [訂單類別] 資料"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
Else
   cbo_Priority.Text = Trim(cbo_Priority.Text)
End If

If Fun_ChkDateFormat(Trim(txt_OrderDate.Text)) = 1 Then: MsgBox "訂單日格式錯誤(YYYYMMDD)!", 16, Me.Caption: Exit Function
If Fun_ChkDateFormat(Trim(txt_DeliveryDate.Text)) = 1 Then: MsgBox "到貨日格式錯誤(YYYYMMDD)!", 16, Me.Caption: Exit Function
If Trim(txt_OrderDate.Text) = "" Then
   msg_text = "資料錯誤：未輸入 [訂單日] 資料"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
Else
   txt_OrderDate.Text = Trim(txt_OrderDate.Text)
End If
If Trim(txt_DeliveryDate.Text) = "" Then
   msg_text = "資料錯誤：未輸入 [到貨日] 資料"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
Else
   txt_DeliveryDate.Text = Trim(txt_DeliveryDate.Text)
End If
If Trim(txt_ConsigneeKey.Text) = "" Then
   msg_text = "資料錯誤：未輸入 [客戶編號] 資料"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
Else
   txt_ConsigneeKey.Text = Trim(txt_ConsigneeKey.Text)
End If

If Trim(txtShipToKey.Text) = "" And mySplit(Trim(cbo_Priority), " ", 0) = "A2B" Then
   msg_text = "資料錯誤：未輸入 [轉運到貨客戶編號] 資料"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
Else
   txtShipToKey.Text = Trim(txtShipToKey.Text)
End If

If txt_Extern.Enabled = True Then
   If Trim(txt_Extern.Text) = "" Then
      msg_text = "資料錯誤：新增訂單時必須輸入 [貨主單號] 資料"
      MsgBox msg_text, vbOKOnly + vbCritical, msg_title
      Exit Function
   Else
      txt_Extern.Text = Trim(txt_Extern.Text)
   End If
End If
If cmb_ZIP.ListIndex = -1 Then
   msg_text = "資料錯誤：客戶基本資料必須正確設定 [郵遞區號] 資料"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If

'If dg_OrderDetail.Rows <= 2 Then
'   msg_text = "資料錯誤：必須新增訂單 [明細資料] 資料"
'   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
'   Exit Function
'End If

If rsMain Is Nothing Then
   msg_text = "資料錯誤：必須新增訂單 [明細資料] 資料"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If

If rsMain.RecordCount = 0 Then
   msg_text = "資料錯誤：必須新增訂單 [明細資料] 資料"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If

rsMain.MoveFirst
Do While Not rsMain.EOF
    If rsMain("大單位數量") < 0 Or rsMain("中單位數量") < 0 Or rsMain("小單位數量") < 0 Then MsgBox "數量負數無效!", 16, "存檔": Exit Function
    If rsMain("大單位數量") > 0 And rsMain("大單位入數") = 0 Then MsgBox "大單位入數為0，輸入大單位數量無效!", 16, "存檔": Exit Function
    If rsMain("中單位數量") > 0 And rsMain("中單位入數") = 0 Then MsgBox "中單位入數為0，輸入中單位數量無效!", 16, "存檔": Exit Function
    If rsMain("大單位數量") <> Int(rsMain("大單位入數")) And rsMain("中單位數量") <> Int(rsMain("中單位數量")) And rsMain("小單位數量") <> Int(rsMain("小單位數量")) Then MsgBox "數量不能有小數點!", 16, "存檔": Exit Function
    If rsMain("大單位數量") * rsMain("大單位入數") + rsMain("中單位數量") * rsMain("中單位入數") + rsMain("小單位數量") = 0 Then MsgBox "訂單量不能為0!", 16, "存檔": Exit Function
rsMain.MoveNext
Loop

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)

'1.檢核客戶編號
str_SQL = "Select Count(*) AS RecCount From TRP01M Where storerkey = '" & mySplit(RTrim(cmbStorerkey), " ", 0) & "' and ConsigneeKey = '" & txt_ConsigneeKey.Text & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCount").Value = 0 Then
   tmp_Rs.Close
   msg_text = "資料錯誤：客戶編號 [" & txt_ConsigneeKey.Text & "] 未建檔"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If
tmp_Rs.Close

'是否已排路編
tmp_Rs.Open "select route_no from trp02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' union select route_no from ort02t where c_receipt_no = '" & txt_QueryExternOrderKey & "' and route_no <> 'D' ", cn
If Not tmp_Rs.EOF Then MsgBox "此訂單已排路線編號 " & tmp_Rs("route_no") & " ，無法修改!", vbOKOnly, "存檔": tmp_Rs.Close: Exit Function
tmp_Rs.Close

If mySplit(Trim(cmbStorerkey), " ", 0) = "LTKK01" Then

    '台灣麒麟訂單重複檢查
    If txt_QueryExternOrderKey.Enabled = True Then
       '新增訂單時
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and consigneekey = '" & txt_ConsigneeKey.Text & "' and convert(varchar(8),deliverydate,112) = '" & txt_DeliveryDate.Text & "' and rtrim(isnull(type,'')) <> '刪單' and priority ='" & mySplit(Trim(cbo_Priority.Text), " ", 0) & "' and cast(notes as varchar(300)) = '" & Trim(txt_Description) & "' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "台灣麒麟訂單重覆!(相同客戶編號、到貨日、訂單類別與訂單備註)", 64, "資料錯誤": tmp_Rs.Close: Exit Function
    Else
        '修改訂單時
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and consigneekey = '" & txt_ConsigneeKey.Text & "' and convert(varchar(8),deliverydate,112) = '" & txt_DeliveryDate.Text & "' and rtrim(isnull(type,'')) <> '刪單' and priority ='" & mySplit(Trim(cbo_Priority.Text), " ", 0) & "' and cast(notes as varchar(300)) = '" & txt_Description & "' and orderkey <> '" & txt_QueryExternOrderKey & "' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "台灣麒麟訂單重覆!(相同客戶編號、到貨日、訂單類別與訂單備註)", 64, "資料錯誤": tmp_Rs.Close: Exit Function
    End If

ElseIf mySplit(Trim(cmbStorerkey), " ", 0) = "LNIP01" Then
    If txt_QueryExternOrderKey.Enabled = True Then
       '新增訂單時
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '刪單' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "台灣立邦訂單號碼有重覆情形(允許重複)，請確認訂單資料無誤！", 64, "注意"
    Else
        '修改訂單時
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '刪單' and orderkey <> '" & txt_QueryExternOrderKey & "' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "台灣立邦訂單號碼有重覆情形(允許重複)，請確認訂單資料無誤！", 64, "注意"
    End If

ElseIf mySplit(Trim(cmbStorerkey), " ", 0) = "LABT01" Then
    If txt_QueryExternOrderKey.Enabled = True Then
       '新增訂單時
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '刪單' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "亞培訂單號碼重覆!", 64, "資料錯誤": tmp_Rs.Close: Exit Function
    Else
        '修改訂單時
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '刪單' and orderkey <> '" & txt_QueryExternOrderKey & "' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "亞培訂單號碼有重覆情形(允許重複)，請確認訂單資料無誤！並且另一貨主單號資料記得修正！", 64, "注意"
    End If
    
Else
     '其他貨主訂單重複檢查
    If txt_QueryExternOrderKey.Enabled = True Then
       '新增訂單時
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '刪單' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "貨主訂單號碼重覆!", 64, "資料錯誤": tmp_Rs.Close: Exit Function
    Else
        '修改訂單時
       tmp_Rs.Open "Select Count(*) as RecCount From Orders Where StorerKey = '" & mySplit(Trim(cmbStorerkey), " ", 0) & "' and ExternOrderKey = '" & txt_Extern.Text & "' and rtrim(isnull(type,'')) <> '刪單' and orderkey <> '" & txt_QueryExternOrderKey & "' ", cn
       If tmp_Rs.Fields("RecCount").Value <> 0 Then MsgBox "貨主訂單號碼重覆!", 64, "資料錯誤": tmp_Rs.Close: Exit Function
    End If
    
End If
tmp_Rs.Close

' mark by gemini
''3.修改：檢核狀態，貨主單號為主，所有訂單必須位於 [待排車狀態] 才可進行修改
'If txt_Extern.Enabled = False Then
'   str_SQL = "Select Count(*) as RecCount From TRP02T Where StorerKey = '" & txt_StorerKey.Text & "' and Extern = '" & txt_Extern.Text & "'"
'   tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'   If tmp_rs.Fields("RecCount").Value > 0 Then
'      tmp_rs.Close
'      msg_text = "資料錯誤：貨主訂單 [" & txt_Extern.Text & "] 狀態顯式 [已排車]，不允許修改"
'      MsgBox msg_text, vbOKOnly + vbCritical, msg_title
'      Exit Function
'   End If
'End If

''4.細項資料驗證，計算 [箱數]、[板數]、[材積]、[重量]
'Dim strSKU As String, strErrorSKU As String
'Dim dbShipQty As Double
'strErrorSKU = ""
'With dg_OrderDetail
'     If .Rows = 2 Then Exit Function
'     For iLoop = 1 To .Rows - 2
'         .Row = iLoop
'         .Col = 1: strSKU = .Text
'         If CheckSKU(.Text) = 1 Then Exit Function
'         .Col = 4: dbShipQty = Val(.Text)
'         str_SQL = "Select * " & _
'                   "From BaseData_SKUPacking Where StorerKey = '" & cmbStorerkey.Text & "' and SKU = '" & strSKU & "'"
'         tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'         If tmp_rs.EOF Then
'            If strErrorSKU = "" Then
'               strErrorSKU = strSKU
'            Else
'               strErrorSKU = strErrorSKU & "," & strSKU
'            End If
'         Else
'
'            .Col = 5
'            If tmp_rs.Fields("板轉換").Value = 0 Then
'            .Text = 0
'            Else
'            .Text = NumRound((dbShipQty / tmp_rs.Fields("板轉換").Value), 2)
'            End If
'
'            .Col = 6: .Text = NumRound((dbShipQty * tmp_rs.Fields("單位材積").Value), 2)
'            .Col = 7: .Text = NumRound((dbShipQty * tmp_rs.Fields("單位重量").Value), 2)
'         End If
'         tmp_rs.Close
'     Next iLoop
'End With

'Terry 20181217 user要求開放C單
''訂單類別檢查，C單只有花王、麒麟和亞培使用
'If mySplit(Trim(cbo_Priority), " ", 0) = "C" Then
'    If Left(RTrim(cmbStorerkey.Text), 6) <> "LKAO01" And Left(RTrim(cmbStorerkey.Text), 6) <> "LABT01" And Left(RTrim(cmbStorerkey.Text), 6) <> "LTKK01" And Left(RTrim(cmbStorerkey.Text), 6) <> "LCHF01" And Left(RTrim(cmbStorerkey.Text), 6) <> "LPSI01" And Left(RTrim(cmbStorerkey.Text), 6) <> "LNVA01" Then
'        msg_text = "資料錯誤： [訂單類別]=C，目前限定為亞培、麒麟、中祥、百事、妮維雅及花王貨主使用"
'        MsgBox msg_text, vbOKOnly + vbCritical, msg_title
'        Exit Function
'    End If
'End If

CheckOrdersData = True
Exit Function

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單維護-存檔-資料檢核", Me.Caption, "Form SubProgram CheckOrdersData", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Function

Private Function SaveOrdersData() As Boolean
'訂單資料存檔
On Error GoTo err_Handle
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
SaveOrdersData = False
Dim rsTmp As New ADODB.Recordset, intTmp As Long, strConsigneeKey As String, str_receive As String

If Chk_receive.Value = 1 Then str_receive = 1 Else str_receive = 0

'資料庫異動交易--起點
'Tran_Level = cn.BeginTrans

Dim strOrderkey As String

    '1.取新的訂單編號
    str_SQL = "select isnull(max(orderkey),0) from orders"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    strOrderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
    tmp_Rs.Close
    
    Dim intPointer As Integer
    intPointer = 1
    
    '3.訂單表頭資料存檔
    If Left(RTrim(cmbStorerkey.Text), 6) = "LMBO01" And mySplit(Trim(cbo_Priority), " ", 0) = "A2B" Then
    '毛寶A2B訂單，訂單資訊放B點的。其他貨主都是A點的
    Call txtShipToKey_LostFocus
        str_SQL = "Insert into Orders (OrderKey,StorerKey,ExternOrderKey,OrderDate,DeliveryDate,Priority,ConsigneeKey,C_Contact1,C_Company,C_Address1,C_Address2, " & _
                 " C_ZIP,C_Phone1,AddWho,EditWho,DoRoute,Notes,customerorderkey,type,b_company,facility,externconsigneekey,updatesource,b_city,otqty,GoodsBack) Values ('" & _
                 strOrderkey & "','" & mySplit(RTrim(cmbStorerkey.Text), " ", 0) & "','" & txt_Extern & "','" & _
                 Left(txt_OrderDate.Text, 4) & "/" & Mid(txt_OrderDate.Text, 5, 2) & "/" & Right(txt_OrderDate.Text, 2) & "','" & _
                 Left(txt_DeliveryDate.Text, 4) & "/" & Mid(txt_DeliveryDate.Text, 5, 2) & "/" & Right(txt_DeliveryDate.Text, 2) & "'," & _
                 "'" & mySplit(Trim(cbo_Priority), " ", 0) & "','" & txt_ConsigneeKey.Text & "','" & txt_ShipToContact.Text & "','" & txt_ShipToShortName.Text & "','" & myExCharFilter(Trim(GetWord(txt_ShipToAddress.Text, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(txt_ShipToAddress.Text, intPointer, 45))) & "','" & Trim(txt_ShipToZIP.Text) & "','" & _
                 txt_ShipToPhone.Text & "','" & User_id & "','" & User_id & "','Y','" & txt_Description.Text & "','" & UCase(txt_OrderKey.Text) & "','','" & txtShipToKey.Text & "','" & cmdFacility & "','" & txt_ConsigneeKey.Text & "','ManualOrder','" & RTrim(txt_B_city) & "','" & txt_OtQty & "','" & str_receive & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    Else
        str_SQL = "Insert into Orders (OrderKey,StorerKey,ExternOrderKey,OrderDate,DeliveryDate,Priority,ConsigneeKey,C_Contact1,C_Company,C_Address1,C_address2, " & _
                 " C_ZIP,C_Phone1,AddWho,EditWho,DoRoute,Notes,customerorderkey,type,b_company,facility,externconsigneekey,updatesource,b_city,otqty,GoodsBack) Values ('" & _
                 strOrderkey & "','" & mySplit(RTrim(cmbStorerkey.Text), " ", 0) & "','" & txt_Extern & "','" & _
                 Left(txt_OrderDate.Text, 4) & "/" & Mid(txt_OrderDate.Text, 5, 2) & "/" & Right(txt_OrderDate.Text, 2) & "','" & _
                 Left(txt_DeliveryDate.Text, 4) & "/" & Mid(txt_DeliveryDate.Text, 5, 2) & "/" & Right(txt_DeliveryDate.Text, 2) & "'," & _
                 "'" & mySplit(Trim(cbo_Priority), " ", 0) & "','" & txt_ConsigneeKey.Text & "','" & txt_Contact.Text & "','" & txt_ShortName.Text & "','" & myExCharFilter(Trim(GetWord(txt_Address.Text, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(txt_Address.Text, intPointer, 45))) & "','" & arZip(cmb_ZIP.ListIndex) & "','" & _
                 txt_Phone.Text & "','" & User_id & "','" & User_id & "','Y','" & txt_Description.Text & "','" & UCase(txt_OrderKey.Text) & "','','" & txtShipToKey.Text & "','" & cmdFacility & "','" & txt_ConsigneeKey.Text & "','ManualOrder','" & RTrim(txt_B_city) & "','" & txt_OtQty & "','" & str_receive & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    End If
    
'是否是修改
If txt_QueryExternOrderKey.Enabled = False And txt_QueryExternOrderKey <> strOrderkey Then

'毛寶訂單修改，則將TMS單號更新回Custorders中
 If RTrim(cmbStorerkey) = "LMBO01" Then
    cn.Execute "update custorders set orderkey = '" & strOrderkey & "' where orderkey = '" & txt_QueryExternOrderKey & "' ", RowsAffect, adExecuteNoRecords
    cn.Execute "update custorderdetail set orderkey = '" & strOrderkey & "' where orderkey = '" & txt_QueryExternOrderKey & "' ", RowsAffect, adExecuteNoRecords
 End If
 
'取原始訂單資料
    str_SQL = "select externordertype = isnull(o.externordertype,''), amount = isnull(o.amount,0) , billtokey=isnull(o.billtokey,'') , b_contact1 = isnull(o.b_contact1,'') ,  door = isnull(o.door,'') , stop = isnull(o.stop,'') , facility = isnull(o.facility,'') , o.invoiceno,o.externconsigneekey,o.b_vat,ordergroup = isnull(o.ordergroup,''),BuyerPo = isnull(o.BuyerPo,''),b_city = isnull(o.b_city,''),o.otqty " & _
    ", Cash = o.Cash " & _
    ", Bill = o.Bill " & _
    ", ReceiveCash = o.ReceiveCash " & _
    ", ReceiveBill = o.ReceiveBill " & _
    ", PayStatus = o.PayStatus " & _
    ", B_Contact2 = isnull(o.B_Contact2,'') " & _
    ",o.urgent_mark,o.reserve_mark , od.* " & _
    "from orders o join orderdetail od on o.orderkey = od.orderkey where o.orderkey = '" & txt_QueryExternOrderKey & "' "
    
    tmp_Rs.CursorLocation = 3
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

'更新其他欄位OTQTY,B_City, GoodsBack 現在顯示畫面上，所以用畫面上的為主。這邊不用原本的資料做更新

    str_SQL = "update orders set externordertype = '" & RTrim(tmp_Rs("externordertype")) & "'" & _
              ", amount = " & tmp_Rs("amount") & _
              ", billtokey = '" & tmp_Rs("billtokey") & "' " & _
              ", b_contact1 = '" & tmp_Rs("b_contact1") & "' " & _
              ", door = '" & tmp_Rs("door") & "' " & _
              ", stop = '" & tmp_Rs("stop") & "' " & _
              ", b_vat = '" & tmp_Rs("b_vat") & "' " & _
              ", Updatesource = '" & txt_QueryExternOrderKey & "' " & _
              ", invoiceno = '" & tmp_Rs("invoiceno") & "' " & _
              ", externconsigneekey = '" & tmp_Rs("externconsigneekey") & "' " & _
              ", ordergroup = '" & tmp_Rs("ordergroup") & "' " & _
              ", BuyerPo = '" & tmp_Rs("BuyerPo") & "' " & _
              ", Cash = " & tmp_Rs("Cash") & _
              ", Bill = " & tmp_Rs("Bill") & _
              ", ReceiveCash = " & tmp_Rs("ReceiveCash") & _
              ", ReceiveBill = " & tmp_Rs("ReceiveBill") & _
              ", PayStatus = '" & tmp_Rs("PayStatus") & "' " & _
              ", B_Contact2= '" & tmp_Rs("B_Contact2") & "' " & _
              ",urgent_mark= '" & tmp_Rs("urgent_mark") & "' " & _
              ",reserve_mark= '" & tmp_Rs("reserve_mark") & "' " & _
              "where orderkey = '" & strOrderkey & "' "
'    tmp_Rs.Close
    
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
End If

'4.處理訂單細項存檔
Dim strLineNo As String, strSku As String, dbOrderCS As Double, dbOrderIP As Double, dbOrderEA As Double, dbCasecnt As Double, dbInnerpack As Double, strTKLOC As String, strNotes As String, strLot3 As String, strLot4 As String, strLot5 As String

rsMain.MoveFirst
Do While Not rsMain.EOF
    strLineNo = Trim(myExCharFilter(rsMain("項次")))
    strSku = Trim(myExCharFilter(rsMain("貨號")))        '貨號
    strTKLOC = Trim(myExCharFilter(rsMain("倉別")))    'TKLOC
    dbOrderCS = Trim(myExCharFilter(rsMain("大單位數量")))      '箱數
    dbOrderIP = Trim(myExCharFilter(rsMain("中單位數量")))      '個數
    dbOrderEA = Trim(myExCharFilter(rsMain("小單位數量")))      '個數
    dbCasecnt = Trim(myExCharFilter(rsMain("大單位入數")))      '每箱
    dbInnerpack = Trim(myExCharFilter(rsMain("中單位入數")))      '每串
    strNotes = Trim(myExCharFilter(rsMain("備註")))      '備註
    strLot3 = Trim(myExCharFilter(rsMain("生產批號")))      '生產批號
    strLot4 = Trim(myExCharFilter(rsMain("製造日")))      '製造日
    strLot5 = Trim(myExCharFilter(rsMain("到期日")))      '到期日
    
    '新增：insert into OrderDetail
    str_SQL = "insert into OrderDetail(StorerKey,OrderKey,OrderLineNumber,ExternOrderKey,SKU,OriginalQty,OpenQty,ShippedQty,UOM,AddWho,EditWho,lottable06,notes,updatesource,lottable03,lottable04,lottable05) Values ('" & _
              mySplit(RTrim(cmbStorerkey.Text), " ", 0) & "','" & strOrderkey & "','" & strLineNo & "','" & txt_Extern & "','" & strSku & "'," & (dbOrderCS * dbCasecnt) + (dbOrderIP * dbInnerpack) + dbOrderEA & "," & (dbOrderCS * dbCasecnt) + (dbOrderIP * dbInnerpack) + dbOrderEA & ",0" & _
              ",'EA','" & User_id & "','" & User_id & "','" & UCase(strTKLOC) & "','" & strNotes & "','" & Trim(myExCharFilter(rsMain("揀貨加工"))) & "','" & strLot3 & "','" & strLot4 & "','" & strLot5 & "')"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    '無製造日則更新為null
    If Len(Trim(strLot4)) = 0 Then cn.Execute "update orderdetail set lottable04 = null where orderkey = '" & strOrderkey & "' and orderlinenumber = '" & strLineNo & "' ", RowsAffect, adExecuteNoRecords
    
    '無到期日則更新為null
    If Len(Trim(strLot5)) = 0 Then cn.Execute "update orderdetail set lottable05 = null where orderkey = '" & strOrderkey & "' and orderlinenumber = '" & strLineNo & "' ", RowsAffect, adExecuteNoRecords

'訂單修改時補明細資料
If txt_QueryExternOrderKey.Enabled = False And txt_QueryExternOrderKey <> strOrderkey Then

    '補訂單明細資料
    tmp_Rs.MoveFirst
    Do While Not tmp_Rs.EOF
        If strLineNo = RTrim(tmp_Rs("orderlinenumber")) Then
            cn.Execute "update orderdetail set ExternLineNo = '" & tmp_Rs("ExternLineNo") & "' " & _
            ",RetailSKU = '" & tmp_Rs("RetailSKU") & "' " & _
            ",PickCode = '" & tmp_Rs("PickCode") & "' " & _
            ",Facility = '" & tmp_Rs("Facility") & "' " & _
            ",UnitPrice = '" & tmp_Rs("UnitPrice") & "' " & _
            ",OtherUOM = '" & tmp_Rs("OtherUOM") & "' " & _
            ",UOM = '" & tmp_Rs("UOM") & "' " & _
            "where orderkey = '" & strOrderkey & "' and orderlinenumber = '" & strLineNo & "' ", RowsAffect, adExecuteNoRecords
        End If

    tmp_Rs.MoveNext
    Loop
        
End If

rsMain.MoveNext
Loop

'訂單修改時Close
If txt_QueryExternOrderKey.Enabled = False And txt_QueryExternOrderKey <> strOrderkey Then tmp_Rs.Close

'更新packkey
str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & strOrderkey & "' "

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'補客戶資料 LABT的資料代補齊
cn.Execute "exec gs_OrdersUpdate '" & mySplit(cmbStorerkey.Text, " ", 0) & "' ", RowsAffect, adExecuteNoRecords

'7. DB Transaction Commit
'cn.CommitTrans: Tran_Level = 0

'8.清除螢幕
Call Clear_AllField
txt_QueryExternOrderKey.Text = strOrderkey
SaveOrdersData = True
txt_QueryExternOrderKey.Enabled = True
Exit Function

err_Handle:
   If Tran_Level <> 0 Then cn.RollbackTrans: Tran_Level = 0

   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單維護-存檔-資料存檔", Me.Caption, "Form SubProgram SaveOrdersData", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Function

Private Sub txtShipToKey_GotFocus()
    frm_OP_ManualShipToOrders.ZOrder 0
    fam_ConsigneeData.Enabled = False
    Call txtShipToKey_LostFocus
End Sub

Public Sub txtShipToKey_LostFocus()

'轉運到貨客戶編號
If Trim(txtShipToKey.Text) = "" Then Exit Sub

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
str_SQL = "Select * From TRP01M Where storerkey = '" & mySplit(RTrim(cmbStorerkey.Text), " ", 0) & "' and ConsigneeKey = '" & txtShipToKey.Text & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "資料錯誤：客戶編號 [" & txtShipToKey.Text & "] 未建檔"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Sub
End If

txt_ShipToFullName.Text = IIf(IsNull(tmp_Rs.Fields("Full_Name").Value), "", Trim(tmp_Rs.Fields("Full_Name").Value))
txt_ShipToShortName.Text = IIf(IsNull(tmp_Rs.Fields("Short_Name").Value), "", Trim(tmp_Rs.Fields("Short_Name").Value))
txt_ShipToContact.Text = IIf(IsNull(tmp_Rs.Fields("Contact").Value), "", Trim(tmp_Rs.Fields("Contact").Value))
txt_ShipToPhone.Text = IIf(IsNull(tmp_Rs.Fields("Phone").Value), "", Trim(tmp_Rs.Fields("Phone").Value))
txt_ShipToAddress.Text = IIf(IsNull(tmp_Rs.Fields("Address").Value), "", Trim(tmp_Rs.Fields("Address").Value))
txt_ShipToZIP = tmp_Rs("zip") & ""
txt_ShipToAreaCode = tmp_Rs("Area_Code") & ""
txt_ShipToExtraDemand1 = tmp_Rs("Extra_Demand_code2") & ""
txt_ShipToExtraDemand2 = tmp_Rs("Extra_Demand_code2") & ""

tmp_Rs.Close

End Sub

Private Sub cboLot6_KeyPress(KeyAscii As Integer)
'小寫字元改為大寫字元
If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
End Sub

