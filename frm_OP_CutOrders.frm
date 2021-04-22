VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frm_OP_CutOrders 
   Caption         =   "ㄧ單多車訂單切割"
   ClientHeight    =   7140
   ClientLeft      =   210
   ClientTop       =   855
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11475
   Begin TabDlg.SSTab SSTab1 
      Height          =   7080
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12488
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "訂單列表"
      TabPicture(0)   =   "frm_OP_CutOrders.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Shape2(0)"
      Tab(0).Control(1)=   "Shape1"
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(5)=   "Label1(0)"
      Tab(0).Control(6)=   "Label1(19)"
      Tab(0).Control(7)=   "Shape2(1)"
      Tab(0).Control(8)=   "dg_TRP02W"
      Tab(0).Control(9)=   "cmd_Tab1_ResetRS"
      Tab(0).Control(10)=   "cmd_FilterAndSort"
      Tab(0).Control(11)=   "txt_Tab0_TotalCase"
      Tab(0).Control(12)=   "txt_Tab0_TotalPallet"
      Tab(0).Control(13)=   "txt_Tab0_TotalVolumn"
      Tab(0).Control(14)=   "txt_Tab0_TotalWeight"
      Tab(0).Control(15)=   "txt_Tab0_OrderCount"
      Tab(0).Control(16)=   "cmd_Tab0_DisplaySelectedOrder"
      Tab(0).Control(17)=   "cmd_Tab0_DisplayOrders"
      Tab(0).Control(18)=   "cmd_Exit(0)"
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "待切割訂單"
      TabPicture(1)   =   "frm_OP_CutOrders.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fam_Tab1_Orders"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fam_Tab1_OrderDetail"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "訂單切割明細 + 查詢"
      TabPicture(2)   =   "frm_OP_CutOrders.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dg_CutOrderDetail"
      Tab(2).Control(1)=   "dg_CutOrders"
      Tab(2).Control(2)=   "fam_Tab2_Delete"
      Tab(2).Control(3)=   "fam_Tab2_Qoery"
      Tab(2).ControlCount=   4
      Begin VB.Frame fam_Tab1_OrderDetail 
         BackColor       =   &H00808000&
         Caption         =   "訂單明細"
         ForeColor       =   &H00400040&
         Height          =   4470
         Left            =   270
         TabIndex        =   54
         Top             =   2400
         Width           =   10875
         Begin VB.TextBox txt_Tab1_SelectedPalletQty 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   4365
            TabIndex        =   63
            Top             =   465
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_SelectedVolumn 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   3420
            TabIndex        =   62
            Top             =   465
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_SelectedWeight 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   2475
            TabIndex        =   61
            Top             =   465
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_SelectedCaseQty 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   1515
            TabIndex        =   60
            Top             =   465
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_CutCaseQty 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   285
            Left            =   6480
            TabIndex        =   59
            Top             =   450
            Width           =   700
         End
         Begin VB.CommandButton cmd_Tab1_CutQty 
            BackColor       =   &H00C0C0FF&
            Caption         =   "數量切割"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   560
            Left            =   7200
            Style           =   1  '圖片外觀
            TabIndex        =   58
            Top             =   180
            Width           =   1200
         End
         Begin VB.TextBox txt_Tab1_CutPalletQty 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   285
            Left            =   6480
            TabIndex        =   57
            Top             =   150
            Width           =   700
         End
         Begin VB.CommandButton cmd_Tab1_CutOrders 
            BackColor       =   &H00FF8080&
            Caption         =   "訂單切割"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   560
            Left            =   9600
            Style           =   1  '圖片外觀
            TabIndex        =   56
            Top             =   180
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab1_ClearQty 
            BackColor       =   &H00FF80FF&
            Caption         =   "取消切割"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   560
            Left            =   8385
            Style           =   1  '圖片外觀
            TabIndex        =   55
            Top             =   180
            Width           =   1200
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_SelectedOrderDetail 
            Height          =   3645
            Left            =   45
            TabIndex        =   64
            Top             =   765
            Width           =   10740
            _ExtentX        =   18944
            _ExtentY        =   6429
            _Version        =   393216
            Cols            =   9
            _NumberOfBands  =   1
            _Band(0).Cols   =   9
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Index           =   18
            Left            =   1860
            TabIndex        =   71
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "板數"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Index           =   20
            Left            =   4650
            TabIndex        =   70
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "材積"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Index           =   21
            Left            =   3735
            TabIndex        =   69
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "重量"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Index           =   22
            Left            =   2805
            TabIndex        =   68
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "2.箱數切割"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0FF&
            Height          =   195
            Index           =   23
            Left            =   5475
            TabIndex        =   67
            Top             =   510
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "1.板數切割"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0FF&
            Height          =   195
            Index           =   24
            Left            =   5475
            TabIndex        =   66
            Top             =   225
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "選取項次小計"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   195
            Index           =   25
            Left            =   195
            TabIndex        =   65
            Top             =   510
            Width           =   1260
         End
      End
      Begin VB.CommandButton cmd_Exit 
         BackColor       =   &H00FFC0FF&
         Caption         =   "離  開"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Index           =   0
         Left            =   -64890
         Style           =   1  '圖片外觀
         TabIndex        =   53
         Top             =   495
         Width           =   1050
      End
      Begin VB.CommandButton cmd_Tab0_DisplayOrders 
         BackColor       =   &H00FF8080&
         Caption         =   "匯入待排車訂單"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -74670
         Style           =   1  '圖片外觀
         TabIndex        =   52
         Top             =   495
         Width           =   2250
      End
      Begin VB.CommandButton cmd_Tab0_DisplaySelectedOrder 
         BackColor       =   &H00C0E0FF&
         Caption         =   "訂單切割明細"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -72135
         Style           =   1  '圖片外觀
         TabIndex        =   50
         Top             =   495
         Width           =   2250
      End
      Begin VB.TextBox txt_Tab0_OrderCount 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   -73845
         TabIndex        =   49
         Top             =   6615
         Width           =   915
      End
      Begin VB.TextBox txt_Tab0_TotalWeight 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   -69630
         TabIndex        =   48
         Top             =   6615
         Width           =   1290
      End
      Begin VB.TextBox txt_Tab0_TotalVolumn 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   -67365
         TabIndex        =   47
         Top             =   6615
         Width           =   1290
      End
      Begin VB.TextBox txt_Tab0_TotalPallet 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   -65130
         TabIndex        =   46
         Top             =   6615
         Width           =   1290
      End
      Begin VB.TextBox txt_Tab0_TotalCase 
         Alignment       =   1  '靠右對齊
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   -71955
         TabIndex        =   45
         Top             =   6615
         Width           =   1290
      End
      Begin VB.Frame fam_Tab1_Orders 
         BackColor       =   &H8000000C&
         Caption         =   "訂單資料"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   1875
         Left            =   255
         TabIndex        =   11
         Top             =   480
         Width           =   10905
         Begin VB.TextBox txt_Tab1_Storer 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   990
            TabIndex        =   29
            Top             =   285
            Width           =   825
         End
         Begin VB.TextBox txt_Tab1_OrderKey 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   2745
            TabIndex        =   28
            Top             =   285
            Width           =   1050
         End
         Begin VB.TextBox txt_Tab1_Extern 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   4710
            TabIndex        =   27
            Top             =   285
            Width           =   1410
         End
         Begin VB.TextBox txt_Tab1_FullName 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   990
            TabIndex        =   26
            Top             =   585
            Width           =   5130
         End
         Begin VB.TextBox txt_Tab1_Address 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   990
            TabIndex        =   25
            Top             =   870
            Width           =   5130
         End
         Begin VB.TextBox txt_Tab1_ExtraDemand1 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   990
            TabIndex        =   24
            Top             =   1155
            Width           =   5130
         End
         Begin VB.TextBox txt_Tab1_ExtraDemand2 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   990
            TabIndex        =   23
            Top             =   1440
            Width           =   5130
         End
         Begin VB.TextBox txt_Tab1_ZIP 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   7170
            TabIndex        =   22
            Top             =   585
            Width           =   1680
         End
         Begin VB.TextBox txt_Tab1_AreaCode 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   9795
            TabIndex        =   21
            Top             =   585
            Width           =   825
         End
         Begin VB.TextBox txt_Tab1_VehicleType 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   7170
            TabIndex        =   20
            Top             =   870
            Width           =   1680
         End
         Begin VB.CheckBox chk_Tab1_MultiCustomer 
            BackColor       =   &H8000000C&
            Caption         =   "指送客戶"
            Height          =   180
            Left            =   8910
            TabIndex        =   19
            Top             =   1200
            Width           =   1260
         End
         Begin VB.TextBox txt_Tab1_ChannelType 
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   7170
            TabIndex        =   18
            Top             =   1155
            Width           =   1680
         End
         Begin VB.TextBox txt_Tab1_OrderDate 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   7185
            TabIndex        =   17
            Top             =   285
            Width           =   1170
         End
         Begin VB.TextBox txt_Tab1_DeliveryDate 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   9405
            TabIndex        =   16
            Top             =   270
            Width           =   1230
         End
         Begin VB.TextBox txt_Tab1_Weight 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   7650
            TabIndex        =   15
            Top             =   1440
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_Volumn 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   8595
            TabIndex        =   14
            Top             =   1440
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_PalletQty 
            Alignment       =   2  '置中對齊
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   270
            Left            =   9540
            TabIndex        =   13
            Top             =   1440
            Width           =   945
         End
         Begin VB.TextBox txt_Tab1_EXEConfirm 
            BackColor       =   &H00C0FFC0&
            Height          =   270
            Left            =   9795
            TabIndex        =   12
            Top             =   885
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貨        主"
            Height          =   180
            Index           =   4
            Left            =   225
            TabIndex        =   44
            Top             =   345
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "訂單編號"
            Height          =   180
            Index           =   5
            Left            =   1965
            TabIndex        =   43
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貨主單號"
            Height          =   180
            Index           =   6
            Left            =   3930
            TabIndex        =   42
            Top             =   345
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶名稱"
            Height          =   180
            Index           =   7
            Left            =   225
            TabIndex        =   41
            Top             =   645
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "送貨地址"
            Height          =   180
            Index           =   8
            Left            =   225
            TabIndex        =   40
            Top             =   915
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "特殊需求 1"
            Height          =   180
            Index           =   9
            Left            =   90
            TabIndex        =   39
            Top             =   1215
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "特殊需求 2"
            Height          =   180
            Index           =   10
            Left            =   90
            TabIndex        =   38
            Top             =   1485
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "郵遞區號"
            Height          =   180
            Index           =   11
            Left            =   6390
            TabIndex        =   37
            Top             =   645
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運送區碼"
            Height          =   180
            Index           =   12
            Left            =   9030
            TabIndex        =   36
            Top             =   660
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車種代碼"
            Height          =   180
            Index           =   13
            Left            =   6390
            TabIndex        =   35
            Top             =   915
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "通路型態"
            Height          =   180
            Index           =   14
            Left            =   6390
            TabIndex        =   34
            Top             =   1215
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "訂單日期"
            Height          =   180
            Index           =   15
            Left            =   6390
            TabIndex        =   33
            Top             =   345
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "出貨日期"
            Height          =   180
            Index           =   16
            Left            =   8625
            TabIndex        =   32
            Top             =   330
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "重量/材積/板數"
            Height          =   180
            Index           =   17
            Left            =   6405
            TabIndex        =   31
            Top             =   1515
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "EXE回傳"
            Height          =   180
            Index           =   26
            Left            =   9030
            TabIndex        =   30
            Top             =   960
            Width           =   690
         End
      End
      Begin VB.CommandButton cmd_FilterAndSort 
         BackColor       =   &H00FF80FF&
         Caption         =   "篩 選 排 序"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -69615
         Style           =   1  '圖片外觀
         TabIndex        =   9
         Top             =   510
         Width           =   2160
      End
      Begin VB.CommandButton cmd_Tab1_ResetRS 
         Appearance      =   0  '平面
         BackColor       =   &H00C0C0FF&
         Caption         =   "取消篩選排序"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -67410
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '圖片外觀
         TabIndex        =   8
         Top             =   510
         Width           =   2160
      End
      Begin VB.Frame fam_Tab2_Qoery 
         BackColor       =   &H00404000&
         Height          =   2160
         Left            =   -65790
         TabIndex        =   4
         Top             =   3270
         Width           =   1995
         Begin VB.CommandButton cmd_Tab2_ExternQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "貨主單號查詢"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   105
            Picture         =   "frm_OP_CutOrders.frx":0054
            Style           =   1  '圖片外觀
            TabIndex        =   6
            Top             =   1215
            Width           =   1785
         End
         Begin VB.TextBox txt_Tab2_Extern 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   180
            TabIndex        =   5
            Top             =   630
            Width           =   1605
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貨主單號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   240
            Left            =   465
            TabIndex        =   7
            Top             =   300
            Width           =   1020
         End
      End
      Begin VB.Frame fam_Tab2_Delete 
         Appearance      =   0  '平面
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -65805
         TabIndex        =   2
         Top             =   5520
         Visible         =   0   'False
         Width           =   1995
         Begin VB.CommandButton cmd_Tab2_CutOrderDelete 
            BackColor       =   &H00C0C0FF&
            Caption         =   "刪除切割訂單"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   90
            Picture         =   "frm_OP_CutOrders.frx":035E
            Style           =   1  '圖片外觀
            TabIndex        =   3
            ToolTipText     =   "刪除"
            Top             =   180
            Width           =   1800
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_CutOrders 
         Height          =   2775
         Left            =   -74790
         TabIndex        =   1
         Top             =   450
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   4895
         _Version        =   393216
         Cols            =   9
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
      End
      Begin MSDataGridLib.DataGrid dg_CutOrderDetail 
         Height          =   3600
         Left            =   -74790
         TabIndex        =   10
         Top             =   3285
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6350
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
      Begin MSDataGridLib.DataGrid dg_TRP02W 
         Height          =   5205
         Left            =   -74745
         TabIndex        =   51
         Top             =   1305
         Width           =   10890
         _ExtentX        =   19209
         _ExtentY        =   9181
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         HeadLines       =   1
         RowHeight       =   15
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
      Begin VB.Shape Shape2 
         BackColor       =   &H00004000&
         BackStyle       =   1  '不透明
         BorderColor     =   &H008080FF&
         BorderWidth     =   2
         Height          =   720
         Index           =   1
         Left            =   -74730
         Top             =   435
         Width           =   2385
      End
      Begin VB.Label Label1 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "訂單筆數"
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
         Index           =   19
         Left            =   -74745
         TabIndex        =   76
         Top             =   6690
         Width           =   840
      End
      Begin VB.Label Label1 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "總重量"
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
         Index           =   0
         Left            =   -70320
         TabIndex        =   75
         Top             =   6675
         Width           =   630
      End
      Begin VB.Label Label1 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "總材積"
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
         Index           =   1
         Left            =   -68055
         TabIndex        =   74
         Top             =   6675
         Width           =   630
      End
      Begin VB.Label Label1 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "總板數"
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
         Index           =   2
         Left            =   -65820
         TabIndex        =   73
         Top             =   6690
         Width           =   630
      End
      Begin VB.Label Label1 
         Appearance      =   0  '平面
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  '透明
         Caption         =   "總箱數"
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
         Index           =   3
         Left            =   -72645
         TabIndex        =   72
         Top             =   6675
         Width           =   630
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00404000&
         BackStyle       =   1  '不透明
         BorderColor     =   &H0000C0C0&
         BorderWidth     =   2
         Height          =   735
         Left            =   -69690
         Top             =   450
         Width           =   4530
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000001&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         Height          =   720
         Index           =   0
         Left            =   -72210
         Top             =   435
         Width           =   2385
      End
   End
End
Attribute VB_Name = "frm_OP_CutOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form 設計時期的高
Private dbsrcFormWidth As Double     'Form 設計時期的寬

Private blTRP02WEventEnable As Boolean
Private rs_TRP02W As ADODB.Recordset
Private rs_CutOrderDetail As ADODB.Recordset   '已完成訂單切割之訂單明細

Private dbCut_TotalCaseQty As Double
Private dbCut_TotalWeight As Double
Private dbCut_TotalVolumn As Double
Private dbCut_TotalPalletQty As Double

Private Sub cmd_FilterAndSort_Click()
'訂單列表 >> 篩選排序
If rs_TRP02W Is Nothing Then Exit Sub
If rs_TRP02W.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_TRP02W"

If ShowForm_RS_FilterAndSort(rs_TRP02W, "待排車訂單", Me.Tag) = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Me.WindowState = 2

End Sub

Private Sub cmd_Tab0_DisplayOrders_Click()
'訂單列表 >> 顯示待排車訂單
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_TRP02W.DataSource = Nothing
Set rs_TRP02W = Nothing

str_SQL = "Select 訂單編號,送貨日,客戶編號,貨主單號,箱數,重量,材積,板數,ZIP,區碼,客戶名稱,訂單日,貨主,識別,訂單備註,EXE回傳 " & _
          "From CutOrders_SourceOrder Order by 板數 DESC"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合設定條件之待排車訂單資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_TRP02W)
tmp_Rs.Close

With dg_TRP02W
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_TRP02W.MoveFirst
blTRP02WEventEnable = False
Set dg_TRP02W.DataSource = rs_TRP02W
With dg_TRP02W
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1100       '訂單編號
    .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 900        '送貨日
    .Columns(2).Alignment = dbgCenter
    .Columns(3).Width = 1200       '客戶編號
    .Columns(3).Alignment = dbgCenter
    .Columns(4).Width = 900        '貨主單號
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 800        '箱數
    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 800        '重量
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 800        '材積
    .Columns(7).Alignment = dbgRight
    .Columns(8).Width = 800        '板數
    .Columns(8).Alignment = dbgRight
    .Columns(9).Width = 500        'ZIP
    .Columns(9).Alignment = dbgCenter
    .Columns(10).Width = 500       '區碼
    .Columns(10).Alignment = dbgCenter
    .Columns(11).Width = 3500      '客戶名稱
    .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1000      '訂單日
    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 700       '貨主
    .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 1100      '識別
    .Columns(14).Alignment = dbgLeft
    .Columns(15).Width = 1500      '訂單備註
    .Columns(15).Alignment = dbgLeft
    .Columns(14).Width = 900       'EXE回傳
    .Columns(14).Alignment = dbgLeft
End With
blTRP02WEventEnable = True
'取的訂單所有細項總計資料值
str_SQL = "Select count(訂單編號) as 訂單筆數,sum(箱數) as 總箱數,sum(重量) as 總重量,sum(材積) as 總材積,sum(板數) as 總板數 " & _
          "From CutOrders_SourceOrder  "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
   txt_Tab0_OrderCount.Text = tmp_Rs.Fields("訂單筆數").Value
   txt_Tab0_TotalCase.Text = tmp_Rs.Fields("總箱數").Value
   txt_Tab0_TotalWeight.Text = tmp_Rs.Fields("總重量").Value
   txt_Tab0_TotalVolumn.Text = tmp_Rs.Fields("總材積").Value
   txt_Tab0_TotalPallet.Text = tmp_Rs.Fields("總板數").Value
End If
tmp_Rs.Close

'清欄位值
Call Clear_SelectedOrderData
'設定欲切割訂單之訂單名細
Call SetGrid_Format_SelectedOrderDetail
'設定切割訂單列表
Call SetGrid_Format_CutOrderList
'設定已完成切割訂單明細表
Call CreateRS_CutOrderDetail

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單列表-顯示待排車訂單", Me.Caption, "cmd_Tab0_DisplayOrders_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
'離開
Unload Me
End Sub

Private Sub cmd_Tab0_DisplaySelectedOrder_Click()
'訂單列表 >> 顯示訂單明細
If rs_TRP02W Is Nothing Then Exit Sub
If rs_TRP02W.RecordCount = 0 Then Exit Sub
If dg_TRP02W.SelBookmarks.Count = 0 Then
   msg_text = "無指定選取的訂單"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

Screen.MousePointer = vbHourglass
'設定切割訂單列表
Call SetGrid_Format_CutOrderList
'設定已完成切割訂單明細表
Call CreateRS_CutOrderDetail

'清欄位值
Call Clear_SelectedOrderData
SSTab1.Tab = 1
DoEvents: DoEvents

Dim strOrderkey As String
strOrderkey = rs_TRP02W.Fields("訂單編號").Value

str_SQL = "Select 貨主,訂單編號,貨主單號,訂單日,出貨日,區碼,客戶名稱,送貨地址,特殊需求1,特殊需求2,郵遞區號,車種代碼,通路,指送,重量,材積,板數,EXE回傳 " & _
          "From CutOrders_SelectedOrders Where 訂單編號 = '" & strOrderkey & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合設定條件之待排車訂單資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
txt_Tab1_Storer.Text = tmp_Rs.Fields("貨主").Value
txt_Tab1_OrderKey.Text = tmp_Rs.Fields("訂單編號").Value
txt_Tab1_Extern.Text = tmp_Rs.Fields("貨主單號").Value
txt_Tab1_OrderDate.Text = tmp_Rs.Fields("訂單日").Value
txt_Tab1_DeliveryDate.Text = tmp_Rs.Fields("出貨日").Value
txt_Tab1_FullName.Text = tmp_Rs.Fields("客戶名稱").Value
txt_Tab1_Address.Text = tmp_Rs.Fields("送貨地址").Value
txt_Tab1_ExtraDemand1.Text = tmp_Rs.Fields("特殊需求1").Value
txt_Tab1_ExtraDemand2.Text = tmp_Rs.Fields("特殊需求2").Value
txt_Tab1_ZIP.Text = tmp_Rs.Fields("郵遞區號").Value & ""
txt_Tab1_AreaCode.Text = tmp_Rs.Fields("區碼").Value
txt_Tab1_VehicleType.Text = tmp_Rs.Fields("車種代碼").Value
txt_Tab1_ChannelType.Text = tmp_Rs.Fields("通路").Value
If tmp_Rs.Fields("指送").Value = "Y" Then
   chk_Tab1_MultiCustomer.Value = vbChecked
Else
   chk_Tab1_MultiCustomer.Value = vbUnchecked
End If
txt_Tab1_Weight.Text = tmp_Rs.Fields("重量").Value
txt_Tab1_Volumn.Text = tmp_Rs.Fields("材積").Value
txt_Tab1_PalletQty.Text = tmp_Rs.Fields("板數").Value
txt_Tab1_EXEConfirm.Text = tmp_Rs.Fields("EXE回傳").Value
tmp_Rs.Close

'設定欲切割訂單之訂單名細
Call SetGrid_Format_SelectedOrderDetail
str_SQL = "Select 項次,貨號,品名,訂單量,箱數,重量,材積,板數,每板箱數,每箱個數 " & _
          "From CutOrders_SelectedOrderDetail Where 訂單編號 = '" & strOrderkey & "' order by 項次"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   Screen.MousePointer = vbDefault
   msg_text = "查詢結果：無符合設定條件之待排車訂單明細資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Do While Not tmp_Rs.EOF
   With dg_SelectedOrderDetail
        .Rows = .Rows + 1
        .Row = .Rows - 2
        .Col = 0    '訂單明細項次
        .Text = tmp_Rs.Fields("項次").Value
        .Col = 2    '貨號
        .Text = tmp_Rs.Fields("貨號").Value
        .Col = 3    '品名
        .Text = tmp_Rs.Fields("品名").Value
        .Col = 4    '箱數
        .Text = tmp_Rs.Fields("箱數").Value
        .Col = 5    '重量
        .Text = tmp_Rs.Fields("重量").Value
        .Col = 6    '材積
        .Text = tmp_Rs.Fields("材積").Value
        .Col = 7    '板數
        .Text = tmp_Rs.Fields("板數").Value
        .Col = 13   '每板箱數
        .Text = tmp_Rs.Fields("每板箱數").Value
        '每箱個數
        .Col = 14: .Text = tmp_Rs("每箱個數")
   End With
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
Set tmp_Rs = Nothing
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單列表-顯示訂單名細", Me.Caption, "cmd_Tab0_DisplaySelectedOrder_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   If Not (tmp_Rs Is Nothing) Then
      Set tmp_Rs = Nothing
   End If
End Sub

Private Sub cmd_Tab1_ClearQty_Click()
'待切割訂單 >> 清除 [板數切割][箱數切割] 欄位值
txt_Tab1_CutPalletQty.Text = ""
txt_Tab1_CutCaseQty.Text = ""
'清除切割個數
dg_SelectedOrderDetail.Col = 12: dg_SelectedOrderDetail.Text = ""
txt_Tab1_CutCaseQty.SetFocus
'RUN Button [數量切割] Click
Call cmd_Tab1_CutQty_Click
End Sub

Private Sub cmd_Tab1_CutOrders_Click()
'待切割訂單 >> 切割訂單
If rs_TRP02W Is Nothing Then Exit Sub
If rs_TRP02W.RecordCount = 0 Then Exit Sub

Dim intTRP02WBookMark As String     '正在進行 [訂單切割作業] 之訂單資料列
Dim strCutOrder_SrcKey As String    '正在進行 [訂單切割作業] 之訂單編號
Dim dbMaxKey As Double              '新訂單編號：尾碼 key
Dim strCutOrder_NewKey As String    '新切割出來之訂單其 [訂單編號]
Dim i As Double
Dim int_CS As Integer, int_CutCS As Integer

On Error GoTo err_Handle
If Len(Trim(txt_Tab1_OrderKey.Text)) = 0 Then Exit Sub

If dg_TRP02W.SelBookmarks.Count = 0 Then
   msg_text = "程序錯誤：未選取訂單"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

'檢查是有點選欲切割之訂單細項
dg_SelectedOrderDetail.Visible = False
Dim dbCount As Double
dbCount = 0
With dg_SelectedOrderDetail
     For i = 1 To .Rows - 2
         .Row = i: .Col = 1
         If Len(Trim(.Text)) <> 0 Then
            dbCount = dbCount + 1
         End If
     Next i
End With
dg_SelectedOrderDetail.Visible = True
If dbCount = 0 Then
   msg_text = "資料錯誤：未選取欲切割之訂單喔"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Sub
End If

'檢查總數量是否等於切割數量，防止拆單拆到原單號不見 edit by eric
For i = 1 To dg_SelectedOrderDetail.Rows - 1
    dg_SelectedOrderDetail.Row = i:
    dg_SelectedOrderDetail.Col = 4: int_CS = int_CS + Val(dg_SelectedOrderDetail.Text)
    dg_SelectedOrderDetail.Col = 8: int_CutCS = int_CutCS + Val(dg_SelectedOrderDetail.Text)
Next

If int_CutCS = int_CS Then
   msg_text = "總切割箱數 等於 訂單總箱數，請確認!"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Sub
End If

Screen.MousePointer = vbHourglass
'資料庫異動交易--起點
Tran_Level = 0
Tran_Level = cn.BeginTrans

'為新切割出來之訂單決定其 [訂單編號]
strCutOrder_SrcKey = txt_Tab1_OrderKey.Text
str_SQL = "Select Cast(Code as integer) as AvailNo From CodeLKUP Where ListName = 'CUTORDERSNO'  "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   strCutOrder_NewKey = "CT" & Format(1, "00000000")
   str_SQL = "Insert into CodeLKUP (ListName,Code,Description,AddWho,EditWho) Values ('CUTORDERSNO',2,'ㄧ單多車重新產生訂單號碼','" & User_id & "','" & User_id & "')"
Else
   strCutOrder_NewKey = "CT" & Format(tmp_Rs.Fields("AvailNo").Value, "00000000")
   str_SQL = "Update CodeLKUP Set Code = " & (tmp_Rs.Fields("AvailNo").Value + 1) & " Where ListName = 'CUTORDERSNO'"
End If
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
tmp_Rs.Close

'阻斷原始訂單列表 DBGrid 的 Event 執行
blTRP02WEventEnable = False
rs_TRP02W.Filter = adFilterNone
rs_TRP02W.Filter = "訂單編號 = '" & strCutOrder_SrcKey & "'"
If rs_TRP02W.RecordCount = 0 Then
   msg_text = "抱歉ㄟ，找不到符合條件的原訂單資料喔"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   rs_TRP02W.Filter = adFilterNone
   rs_TRP02W.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
   Exit Sub
Else
   '產生切割訂單-Header
   intTRP02WBookMark = rs_TRP02W.Bookmark
   With dg_CutOrders
        .Rows = .Rows + 1
        .Row = .Rows - 2
        .Col = 0
        .Text = .Rows - 2
        .Col = 1: .Text = strCutOrder_NewKey '訂單編號
        .Col = 2: .Text = rs_TRP02W.Fields("送貨日").Value '送貨日
        .Col = 3: .Text = rs_TRP02W.Fields("客戶編號").Value '客戶編號
        .Col = 4: .Text = rs_TRP02W.Fields("貨主單號").Value '貨主單號
        .Col = 5: .Text = txt_Tab1_SelectedCaseQty.Text '箱數
        .Col = 6: .Text = txt_Tab1_SelectedWeight.Text '重量
        .Col = 7: .Text = txt_Tab1_SelectedVolumn.Text '材積
        .Col = 8: .Text = txt_Tab1_SelectedPalletQty.Text '板數
        .Col = 9: .Text = rs_TRP02W.Fields("ZIP").Value '郵遞區號
        .Col = 10: .Text = rs_TRP02W.Fields("區碼").Value '區碼
        .Col = 11: .Text = rs_TRP02W.Fields("客戶名稱").Value '客戶名稱
        .Col = 12: .Text = rs_TRP02W.Fields("訂單日").Value '訂單日
        .Col = 13: .Text = rs_TRP02W.Fields("貨主").Value '貨主
        .Col = 14: .Text = "切割訂單" '識別
        .Col = 15: .Text = rs_TRP02W.Fields("EXE回傳").Value 'EXE回傳
        .Col = 0
        For i = 0 To .Cols - 1
            .ColSel = i
        Next i
        
        '產生新的訂單資料--TRP02W
        str_SQL = "Insert into TRP02W (StorerKey,Receipt_No,C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Description,Case_cnt,Weight,Volumn_Weight,Pallet_Qty,EXTERN,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,exe_confirm) " & _
                  "Select StorerKey,'" & strCutOrder_NewKey & "',C_Receipt_No,Receipt_Type,TRP_Type,Receipt_Date,Arrive_Date,ConsigneeKey,Description," & _
                  txt_Tab1_SelectedCaseQty.Text & "," & txt_Tab1_SelectedWeight.Text & "," & txt_Tab1_SelectedVolumn.Text & "," & txt_Tab1_SelectedPalletQty.Text & ",EXTERN,Urgent_Mark,Reserve_Mark,Cold_Mark,Priority,exe_confirm " & _
                  "From TRP02W Where StorerKey = '" & txt_Tab1_Storer.Text & "' and Receipt_No = '" & strCutOrder_SrcKey & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '更新原訂單之統計數字--TRP02W
        str_SQL = "Update TRP02W Set Case_cnt=Case_cnt-" & txt_Tab1_SelectedCaseQty.Text & "," & _
                  "Weight=Weight-" & txt_Tab1_SelectedWeight.Text & ",Volumn_Weight=Volumn_Weight-" & txt_Tab1_SelectedVolumn.Text & ",Pallet_Qty=Pallet_Qty-" & txt_Tab1_SelectedPalletQty.Text & " " & _
                  "Where StorerKey = '" & txt_Tab1_Storer.Text & "' and Receipt_No = '" & strCutOrder_SrcKey & "'"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   End With
End If
rs_TRP02W.Filter = adFilterNone
rs_TRP02W.Sort = "訂單編號 ASC"
Do While Not rs_TRP02W.EOF
   If rs_TRP02W.Bookmark = intTRP02WBookMark Then
        Do While dg_TRP02W.SelBookmarks.Count <> 0
           dg_TRP02W.SelBookmarks.Remove 0
        Loop
        '反白顯示正在進行 [訂單切割作業] 之訂單資料列
        dg_TRP02W.SelBookmarks.Add rs_TRP02W.Bookmark
      Exit Do
   End If
   rs_TRP02W.MoveNext
Loop
blTRP02WEventEnable = True

'切割訂單之 OrderDetail
Dim dbsrcQty As Double, dbCutQty As Double, dbSeqNo As Double, dbCutEAQty As Long, dbCasecntQty As Integer
dbSeqNo = 0
dg_SelectedOrderDetail.Visible = False
With dg_SelectedOrderDetail
     For i = 1 To .Rows - 2
         .Row = i: .Col = 1
         If .Text <> "" Then   '細項被選取進行切割
            .Col = 0: dbSeqNo = .Text          '保留原訂單項次編號以為對應
            .Col = 4: dbsrcQty = Val(.Text)    '原訂單箱數
            .Col = 8: dbCutQty = Val(.Text)    '切割箱數
            .Col = 12: dbCutEAQty = Val(.Text) '切割個數
            .Col = 14: dbCasecntQty = Val(.Text) '品項每箱個數
            If dbsrcQty = dbCutQty Then        '若全項次箱數進行切割，註記準備後續刪除此細項
               .Col = 1: .Text = "X"
               Call InsertInto_CutOrderDetail(strCutOrder_NewKey, dbSeqNo)
               str_SQL = "Update TRP03W Set Receipt_No = '" & strCutOrder_NewKey & "' " & _
                         "Where Receipt_No = '" & strCutOrder_SrcKey & "' and SEQ_NO = " & dbSeqNo & ""
               cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            Else
               '更新 待切割訂單明細
               .Col = 1: .Text = ""        '清除註記，更新欄位值
               .Col = 4: dbsrcQty = Val(.Text)    '原訂單箱數
               .Col = 8: dbCutQty = Val(.Text)    '切割箱數
               .Col = 4: .Text = dbsrcQty - dbCutQty
               .Col = 5: dbsrcQty = Val(.Text)    '原訂單重量
               .Col = 9: dbCutQty = Val(.Text)    '切割重量
               .Col = 5: .Text = dbsrcQty - dbCutQty
               .Col = 6: dbsrcQty = Val(.Text)    '原訂單材積
               .Col = 10: dbCutQty = Val(.Text)   '切割材積
               .Col = 6: .Text = dbsrcQty - dbCutQty
               .Col = 7: dbsrcQty = Val(.Text)    '原訂單板數
               .Col = 11: dbCutQty = Val(.Text)   '切割板數
               .Col = 7: .Text = dbsrcQty - dbCutQty
               Call InsertInto_CutOrderDetail(strCutOrder_NewKey, dbSeqNo)
               
               '更新 TRP03W 原數量
               str_SQL = "Update TRP03W Set Order_Qty =  "
'               .Col = 4: str_SQL = str_SQL & .Text & ",Weight = "'mark by gemini
               .Col = 4: str_SQL = str_SQL & (.Text * dbCasecntQty) & ",Weight = " 'add by gemini
               .Col = 5: str_SQL = str_SQL & .Text & ",Volumn_Weight = "
               .Col = 6: str_SQL = str_SQL & .Text & ",Pallet_Qty = "
               .Col = 7: str_SQL = str_SQL & .Text & " "
               str_SQL = str_SQL & "Where Receipt_No = '" & strCutOrder_SrcKey & "' and SEQ_NO = " & dbSeqNo & ""
               cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
               
'               '將箱數換算回個數 by gemini 20071212
'               str_SQL = "Update TRP03W Set TRP03W.Order_Qty = TRP03W.Order_Qty * s1.casecnt " & _
'                        "from trp03w trp03w join sku s on trp03w.product_no = s.sku " & _
'                        "join pack s1 on s1.packkey = s.packkey and s.storerkey = trp03w.storerkey " & _
'                        "Where Receipt_No = '" & strCutOrder_SrcKey & "' and SEQ_NO = " & dbSeqNo & ""
'                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        
               '新增新訂單之訂單細項
               str_SQL = "Insert into TRP03W (StorerKey,Extern,Receipt_No,Seq_No,Product_No,Ship_Unit,Order_Qty,Weight,Volumn_Weight,Pallet_Qty,Description) " & _
                         "Select StorerKey,Extern,'" & strCutOrder_NewKey & "',Seq_No,Product_No,Ship_Unit,"
'               .Col = 8: str_SQL = str_SQL & .Text & ","'mark by gemini
               str_SQL = str_SQL & (dbCutEAQty) & "," 'add by gemini
               .Col = 9: str_SQL = str_SQL & .Text & ","
               .Col = 10: str_SQL = str_SQL & .Text & ","
               .Col = 11: str_SQL = str_SQL & .Text & ","
               str_SQL = str_SQL & "Description From TRP03W Where Receipt_No = '" & strCutOrder_SrcKey & "' and SEQ_NO = " & dbSeqNo & ""
               cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
               
'               '將箱數換算回個數 by gemini 20071212
'               str_SQL = "Update TRP03W Set TRP03W.Order_Qty = TRP03W.Order_Qty * s1.casecnt " & _
'                        "from trp03w trp03w join sku s on trp03w.product_no = s.sku " & _
'                        "join pack s1 on s1.packkey = s.packkey and s.storerkey = trp03w.storerkey " & _
'                        "Where Receipt_No = '" & strCutOrder_NewKey & "' and SEQ_NO = " & dbSeqNo & ""
'               cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
               .Col = 8: .Text = ""   '切割箱數
               .Col = 9: .Text = ""   '切割重量
               .Col = 10: .Text = ""  '切割材積
               .Col = 11: .Text = ""  '切割板數
               .Col = 12: .Text = ""  '切割個數
            End If
         End If
     Next i
End With

'刪除以全數量切割之訂單細項
Dim j As Double
With dg_SelectedOrderDetail
     For i = 1 To .Rows - 2
         For j = 1 To .Rows - 2
             .Row = j: .Col = 1
             If .Text = "X" Then
                Call Delete_GridRow(dg_SelectedOrderDetail, j)
                Exit For
             End If
         Next j
     Next i
     '重新產生訂單加總統計資料
     txt_Tab1_Weight.Text = 0
     txt_Tab1_Volumn.Text = 0
     txt_Tab1_PalletQty.Text = 0
     For i = 1 To .Rows - 2
         .Row = i
         .Col = 5: txt_Tab1_Weight.Text = Val(txt_Tab1_Weight.Text) + Val(.Text)
         .Col = 6: txt_Tab1_Volumn.Text = Val(txt_Tab1_Volumn.Text) + Val(.Text)
         .Col = 7: txt_Tab1_PalletQty.Text = Val(txt_Tab1_PalletQty.Text) + Val(.Text)
     Next i
End With
dg_SelectedOrderDetail.Visible = True

'清除欄位值-選取項次統計
txt_Tab1_SelectedCaseQty.Text = ""
dbCut_TotalCaseQty = 0
txt_Tab1_SelectedWeight.Text = ""
dbCut_TotalWeight = 0
txt_Tab1_SelectedVolumn.Text = ""
dbCut_TotalVolumn = 0
txt_Tab1_SelectedPalletQty.Text = ""
dbCut_TotalPalletQty = 0

'細項切割數量欄位：板數，箱數
txt_Tab1_CutCaseQty.Text = ""
txt_Tab1_CutPalletQty.Text = ""
If dg_SelectedOrderDetail.Rows = 2 And txt_Tab1_Weight.Text = 0 And txt_Tab1_Volumn.Text = 0 And txt_Tab1_PalletQty.Text = 0 Then
   
   '已全部切割之訂單：刪除 TRP02W & TRP03W
   str_SQL = "Delete From TRP02W Where StorerKey = '" & txt_Tab1_Storer.Text & "' and Receipt_No = '" & strCutOrder_SrcKey & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   str_SQL = "Delete From TRP03W Where StorerKey = '" & txt_Tab1_Storer.Text & "' and Receipt_No = '" & strCutOrder_SrcKey & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   '清除 [待切割訂單] 頁面之訂單欄位資料值
   Call Clear_SelectedOrderData
   DoEvents
   SSTab1.Tab = 0
   DoEvents
End If

cn.CommitTrans: Tran_Level = 0
Screen.MousePointer = vbDefault

'拆單之訂單個數檢查
Dim rsTmp As New ADODB.Recordset
'str_SQL = " select * from gv_CheckOrderQty g join trp02w t2 on g.tms單號 = t2.c_receipt_no where TMS單號 = '" & strCutOrder_SrcKey & "' "

'抓取原本單號的c_receipt_no edit by Eric 20141230
str_SQL = "select " & _
            "TMS單號 = od.orderkey " & _
            ",訂單量 = sum(od.originalqty) " & _
            ",待排車量 = isnull((select sum(isnull(t3.order_qty,0)) from trp03w t3(nolock) join trp02w t2(nolock) on t3.receipt_no = t2.receipt_no where t3.receipt_no = t2.receipt_no and t2.c_receipt_no = od.orderkey ),0) " & _
            ",已排車量 = isnull((select sum(isnull(t3.order_qty,0)) from trp03t t3(nolock) join trp02t t2(nolock) on t3.receipt_no = t2.receipt_no where t3.receipt_no = t2.receipt_no and t2.c_receipt_no = od.orderkey ),0) " & _
            "from orderdetail od(nolock) join orders o(nolock) on o.orderkey = od.orderkey and isnull(o.type,'')<>'刪單' and priority not in ( 'R','A2B','RC') " & _
            "where od.orderkey in (select c_receipt_no from trp02w(nolock) where receipt_no = '" & strCutOrder_SrcKey & "') " & _
            "group by od.orderkey " & _
            "having sum(od.originalqty)<>isnull((select sum(isnull(t3.order_qty,0)) from trp03w t3(nolock) join trp02w t2(nolock) on t3.receipt_no = t2.receipt_no where t3.receipt_no = t2.receipt_no and t2.c_receipt_no = od.orderkey ),0)+isnull((select sum(isnull(t3.order_qty,0)) from trp03t t3(nolock) join trp02t t2(nolock) on t3.receipt_no = t2.receipt_no where t3.receipt_no = t2.receipt_no and t2.c_receipt_no = od.orderkey ),0) "
rsTmp.Open str_SQL, cn

If Not rsTmp.EOF Then MsgBox "客戶原始訂單量與拆單訂單量不符，請確認!", vbOKOnly, Me.Caption
rsTmp.Close: Set rsTmp = Nothing

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   
   dg_SelectedOrderDetail.Visible = True
   blTRP02WEventEnable = True
   
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-待切割訂單-訂單切割", Me.Caption, "cmd_Tab1_CutOrders_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_CutQty_Click()
'待切割訂單 >> 數量切割
If rs_TRP02W Is Nothing Then Exit Sub
If rs_TRP02W.RecordCount = 0 Then Exit Sub
If txt_Tab1_SelectedPalletQty = 0 Then MsgBox "每板箱數為0，無法進行訂單切割。", 64, Me.Caption: Exit Sub

Dim tmpQty As Double

cmd_Tab1_CutOrders.Enabled = False
cmd_Tab1_ClearQty.Enabled = False
If Val(txt_Tab1_CutCaseQty.Text) = 0 And Val(txt_Tab1_CutPalletQty.Text) = 0 Then
   '把數量清除表示：不選取此項
   dg_SelectedOrderDetail.Col = 1: dg_SelectedOrderDetail.Text = ""   '取消選取
End If

If Val(txt_Tab1_CutPalletQty.Text) > 0 Then
   dg_SelectedOrderDetail.Col = 7: tmpQty = Val(dg_SelectedOrderDetail.Text)
   If Val(txt_Tab1_CutPalletQty.Text) > tmpQty Then
      msg_text = "資料錯誤：切割板數 大於 品項總板數"
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
      cmd_Tab1_ClearQty.Enabled = True
      cmd_Tab1_CutOrders.Enabled = True
      Exit Sub
   End If
   '有輸入切割板數：以板數為準，清除 [切割箱數] 欄位值
   dg_SelectedOrderDetail.Col = 11
   dg_SelectedOrderDetail.Text = txt_Tab1_CutPalletQty.Text
   dg_SelectedOrderDetail.Col = 8
   dg_SelectedOrderDetail.Text = ""
   
   '計算切割個數
   dg_SelectedOrderDetail.Col = 13: tmpQty = Val(dg_SelectedOrderDetail.Text) * Val(txt_Tab1_CutPalletQty)
   dg_SelectedOrderDetail.Col = 14: tmpQty = Val(dg_SelectedOrderDetail.Text) * tmpQty
   dg_SelectedOrderDetail.Col = 12: dg_SelectedOrderDetail.Text = tmpQty: If dg_SelectedOrderDetail.Text = 0 Then dg_SelectedOrderDetail.Text = ""
Else
   dg_SelectedOrderDetail.Col = 4: tmpQty = Val(dg_SelectedOrderDetail.Text)
   If Val(txt_Tab1_CutCaseQty.Text) > tmpQty Then
      msg_text = "資料錯誤：切割箱數 大於 品項總箱數"
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
      cmd_Tab1_ClearQty.Enabled = True
      cmd_Tab1_CutOrders.Enabled = True
      Exit Sub
   End If

   '輸入切割箱數：箱數
   dg_SelectedOrderDetail.Col = 11
   dg_SelectedOrderDetail.Text = ""
   dg_SelectedOrderDetail.Col = 8
   dg_SelectedOrderDetail.Text = txt_Tab1_CutCaseQty.Text
   
    '計算切割個數
   dg_SelectedOrderDetail.Col = 14: tmpQty = Val(dg_SelectedOrderDetail.Text) * Val(txt_Tab1_CutCaseQty)
   dg_SelectedOrderDetail.Col = 12: dg_SelectedOrderDetail.Text = tmpQty: If dg_SelectedOrderDetail.Text = 0 Then dg_SelectedOrderDetail.Text = ""
End If

'檢查切割個數是否為整數
dg_SelectedOrderDetail.Col = 12
If Val(dg_SelectedOrderDetail.Text) <> Int(Val(dg_SelectedOrderDetail.Text)) Then MsgBox "切割個數不能有小數點!", vbOKOnly, Me.Caption: Call cmd_Tab1_ClearQty_Click

'計算選取之訂單細項之加總 [箱數] [重量] [才積] [板數]
Call Calculate_SelectedPrderDetail

'清除切割量欄位值
txt_Tab1_CutCaseQty.Text = ""
txt_Tab1_CutPalletQty.Text = ""
cmd_Tab1_ClearQty.Enabled = True
cmd_Tab1_CutOrders.Enabled = True
End Sub

Private Sub cmd_Tab1_ResetRS_Click()
'取消篩選排序
If rs_TRP02W Is Nothing Then Exit Sub

'移除篩選條件，重設排序依據
 blTRP02WEventEnable = False
 rs_TRP02W.Filter = adFilterNone
 rs_TRP02W.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
 blTRP02WEventEnable = True
End Sub

Private Sub cmd_Tab2_CutOrderDelete_Click()
'訂單切割明細 >> 刪除

Dim dbDeleteRow As Double, strOrderkey As String, strStorerkey As String, strExtern As String, strC_Receipt_no As String
With dg_CutOrders
     dbDeleteRow = .Row
     .Col = 1: strOrderkey = .Text      '訂單編號 Receipt_No
     .Col = 4: strExtern = .Text        '貨主單號 Extern
     .Col = 13: strStorerkey = .Text    '貨主  StorerKey
     .Col = 16: strC_Receipt_no = .Text '原始TMS單號 C_Receipt_no
     
     If .Text = "" Then Exit Sub
     If Left(strOrderkey, 2) <> "CT" Then MsgBox "非切割訂單無法刪除!", vbOKOnly, Me.Caption: Exit Sub
     
     msg_text = "刪除作業：確認刪除選取之切割訂單：" & strOrderkey
     If MsgBox(msg_text, vbOKCancel + vbInformation, msg_title) = vbCancel Then Exit Sub
End With

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)

'檢核欲刪除之訂單：以原始TMS單號為查詢條件
str_SQL = "Select Count(*) as RecCount From TRP02W Where c_receipt_no = '" & strC_Receipt_no & "' and StorerKey = '" & strStorerkey & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCount").Value = 1 Then
   tmp_Rs.Close
   msg_text = "訂單編號：" & strOrderkey & " 不允許刪除，因其原始TMS單號只對應ㄧ筆訂單資料!"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
ElseIf tmp_Rs.Fields("RecCount").Value = 0 Then
   tmp_Rs.Close
   msg_text = "訂單編號：" & strOrderkey & " 已不存在，請重新執行查詢!"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
tmp_Rs.Close

'取最小訂單編號：接收被刪除訂單所有之項目、數量
Dim strToOrderKey As String
str_SQL = "Select Min(Receipt_No) as 接收訂單編號 From TRP02W Where C_Receipt_no = '" & strC_Receipt_no & "' and StorerKey = '" & strStorerkey & "' and Receipt_No <> '" & strOrderkey & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
   strToOrderKey = tmp_Rs.Fields("接收訂單編號").Value
Else
   tmp_Rs.Close
   msg_text = "找不到可以接收欲刪除之訂單項次的目標訂單!"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
tmp_Rs.Close

'以下未完成 Gemini @20080324
Tran_Level = 0
Tran_Level = cn.BeginTrans
'更新接收訂單之相關資料 TRP02W
With dg_CutOrders
     .Row = dbDeleteRow
     str_SQL = "Update TRP02W Set Case_cnt=Case_cnt+"
     .Col = 5: str_SQL = str_SQL & .Text & ",Weight=Weight+"
     .Col = 6: str_SQL = str_SQL & .Text & ",Volumn_Weight=Volumn_Weight+"
     .Col = 7: str_SQL = str_SQL & .Text & ",Pallet_Qty=Pallet_Qty+"
     .Col = 5: str_SQL = str_SQL & .Text & " "
     str_SQL = str_SQL & "Where StorerKey = '" & strStorerkey & "' and Receipt_No = '" & strToOrderKey & "'"
     cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
End With

'更新接收訂單之相關資料 TRP03W
rs_CutOrderDetail.Filter = adFilterNone
rs_CutOrderDetail.Filter = "訂單編號 = '" & strOrderkey & "'"
If rs_CutOrderDetail.EOF Then
   msg_text = "抱歉ㄟ，找不到符合條件的子訂單明細資料喔"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   rs_CutOrderDetail.Filter = adFilterNone
   rs_CutOrderDetail.Sort = "訂單編號,項次 ASC"  '原始排序，一般資料序號由小至大
   Exit Sub
Else
   Do While Not rs_CutOrderDetail.EOF
      '找找看接收訂單編號有無相同項次、貨號的訂單細項 TRP03W
      str_SQL = "Select Count(*) AS RecCount From TRP03W " & _
                "Where StorerKey = '" & strStorerkey & "' and Receipt_No = '" & strToOrderKey & "' and " & _
                "      Seq_No = " & rs_CutOrderDetail.Fields("項次").Value & " and Product_No = '" & rs_CutOrderDetail.Fields("貨號").Value & "'"
      tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
      If tmp_Rs.Fields("RecCount").Value = 0 Then
         '新增細項 TRP03W
         str_SQL = "Insert into TRP03W (StorerKey,EXTERN,Receipt_No,Seq_No,Product_No,Ship_Unit,Order_Qty,Weight,Volumn_Weight,Pallet_Qty,Description) " & _
                   "Select StorerKey,EXTERN,'" & strToOrderKey & "',Seq_No,Product_No,Ship_Unit,Order_Qty,Weight,Volumn_Weight,Pallet_Qty,Description " & _
                   "From TRP03W Where StorerKey = '" & strStorerkey & "' and Receipt_No = '" & strOrderkey & "' and " & _
                   "      Seq_No = " & rs_CutOrderDetail.Fields("項次").Value & " and Product_No = '" & rs_CutOrderDetail.Fields("貨號").Value & "'"
      Else
         '更新細項 TRP03W
         str_SQL = "Update TRP03W Set Order_Qty = Order_Qty + " & rs_CutOrderDetail.Fields("箱數").Value & "," & _
                   "Weight = Weight + " & rs_CutOrderDetail.Fields("重量").Value & "," & _
                   "Volumn_Weight = Volumn_Weight + " & rs_CutOrderDetail.Fields("材積").Value & "," & _
                   "Pallet_Qty = Pallet_Qty + " & rs_CutOrderDetail.Fields("板數").Value & " " & _
                   "Where StorerKey = '" & strStorerkey & "' and Receipt_No = '" & strToOrderKey & "' and " & _
                   "      Seq_No = " & rs_CutOrderDetail.Fields("項次").Value & " and Product_No = '" & rs_CutOrderDetail.Fields("貨號").Value & "'"
      End If
      cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
      tmp_Rs.Close
      rs_CutOrderDetail.MoveNext
   Loop
   '刪除細項
   rs_CutOrderDetail.MoveFirst
   Do While Not rs_CutOrderDetail.EOF
      rs_CutOrderDetail.Delete
      rs_CutOrderDetail.MoveFirst
   Loop
   str_SQL = "Delete From TRP03W Where StorerKey = '" & strStorerkey & "' and Receipt_No = '" & strOrderkey & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
End If
rs_CutOrderDetail.Filter = adFilterNone
rs_CutOrderDetail.Sort = "訂單編號,項次 ASC"  '原始排序，一般資料序號由小至大

'刪除子訂單表頭
Call Delete_GridRow(dg_CutOrders, dbDeleteRow)
'刪除訂單主檔 TRP02W
str_SQL = "Delete From TRP02W Where StorerKey = '" & strStorerkey & "' and Receipt_No = '" & strOrderkey & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

cn.CommitTrans
Tran_Level = 0
Screen.MousePointer = vbDefault

'拆單之訂單個數檢查
Dim rsTmp As New ADODB.Recordset
str_SQL = " select * from gv_CheckOrderQty g join trp02w t2 on g.tms單號 = t2.c_receipt_no where TMS單號 = '" & strToOrderKey & "' "
rsTmp.Open str_SQL, cn
If Not rsTmp.EOF Then MsgBox "客戶原始訂單量與拆單訂單量不符，請確認!", vbOKOnly, Me.Caption
rsTmp.Close: Set rsTmp = Nothing

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單切割明細-刪除", Me.Caption, "cmd_Tab2_CutOrderDelete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_ExternQuery_Click()
'訂單切割明細+查詢 >> 查詢
If Len(Trim(txt_Tab2_Extern.Text)) = 0 Then Exit Sub

Screen.MousePointer = vbHourglass
On Error GoTo err_Handle

'設定切割訂單列表
Call SetGrid_Format_CutOrderList
'設定已完成切割訂單明細表
Call CreateRS_CutOrderDetail

str_SQL = "Select 訂單編號,送貨日,客戶編號,貨主單號,箱數,重量,材積,板數,ZIP,區碼,客戶名稱,訂單日,貨主,識別,EXE回傳,原始TMS單號 " & _
          "From CutOrders_SourceOrder Where 貨主單號 like '" & Trim(txt_Tab2_Extern.Text) & "%' Order by 訂單編號 "
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "查詢結果：無符合設定條件之待排車訂單資料(TRP02W)"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Do While Not tmp_Rs.EOF
   With dg_CutOrders
        .Rows = .Rows + 1
        .Row = .Rows - 2
        .Col = 0
        .Text = .Rows - 2
        .Col = 1    '訂單編號
        .Text = tmp_Rs.Fields("訂單編號").Value
        .Col = 2    '送貨日
        .Text = tmp_Rs.Fields("送貨日").Value
        .Col = 3    '客戶編號
        .Text = tmp_Rs.Fields("客戶編號").Value
        .Col = 4    '貨主單號
        .Text = tmp_Rs.Fields("貨主單號").Value
        .Col = 5    '箱數
        .Text = tmp_Rs.Fields("箱數").Value
        .Col = 6    '重量
        .Text = tmp_Rs.Fields("重量").Value
        .Col = 7    '材積
        .Text = tmp_Rs.Fields("材積").Value
        .Col = 8    '板數
        .Text = tmp_Rs.Fields("板數").Value
        .Col = 9    '郵遞區號
        .Text = tmp_Rs.Fields("ZIP").Value
        .Col = 10   '區碼
        .Text = tmp_Rs.Fields("區碼").Value
        .Col = 11   '客戶名稱
        .Text = tmp_Rs.Fields("客戶名稱").Value
        .Col = 12   '訂單日
        .Text = tmp_Rs.Fields("訂單日").Value
        .Col = 13   '貨主
        .Text = tmp_Rs.Fields("貨主").Value
        .Col = 14   '識別
        .Text = tmp_Rs.Fields("識別").Value
        .Col = 15   'EXE回傳
        .Text = tmp_Rs.Fields("EXE回傳").Value
        .Col = 16: .Text = tmp_Rs.Fields("原始TMS單號").Value '原始TMS單號
        .Col = 0
   End With
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close

'TRP03W
str_SQL = "Select 訂單編號,項次,貨號,品名,箱數,重量,材積,板數 " & _
          "From CutOrders_SelectedOrderDetail Where 貨主單號 like '" & Trim(txt_Tab2_Extern.Text) & "%' order by 訂單編號,項次"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "查詢結果：無符合設定條件之待排車訂單明細資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Do While Not tmp_Rs.EOF
   rs_CutOrderDetail.AddNew
   rs_CutOrderDetail.Fields("訂單編號").Value = tmp_Rs.Fields("訂單編號").Value
   rs_CutOrderDetail.Fields("項次").Value = tmp_Rs.Fields("項次").Value
   rs_CutOrderDetail.Fields("貨號").Value = tmp_Rs.Fields("貨號").Value
   rs_CutOrderDetail.Fields("品名").Value = tmp_Rs.Fields("品名").Value
   rs_CutOrderDetail.Fields("箱數").Value = tmp_Rs.Fields("箱數").Value
   rs_CutOrderDetail.Fields("重量").Value = tmp_Rs.Fields("重量").Value
   rs_CutOrderDetail.Fields("材積").Value = tmp_Rs.Fields("材積").Value
   rs_CutOrderDetail.Fields("板數").Value = tmp_Rs.Fields("板數").Value
   rs_CutOrderDetail.Update
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-訂單切割名細+查詢-貨主單號查詢", Me.Caption, "cmd_Tab2_ExternQuery_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub dg_CutOrders_Click()
'已完成切割之子訂單列表
'點一次：選取
Dim i As Double
With dg_CutOrders
     .Col = 1   '訂單編號
     If Len(Trim(.Text)) = 0 Then Exit Sub
     .Col = 0
     For i = 0 To .Cols - 1
         .ColSel = i
     Next i
End With
End Sub

Private Sub dg_SelectedOrderDetail_Click()
'待切割之訂單：訂單明細項次
'點一次：選取，除非清除 [切割數量] 否則ㄧ直保持 [選取] 狀態

txt_Tab1_CutPalletQty.Text = ""
txt_Tab1_CutCaseQty.Text = ""

Dim i As Integer
Dim tmpQty As Double
With dg_SelectedOrderDetail
     .Col = 2   '貨號
     If Len(Trim(.Text)) = 0 Then Exit Sub
     .Col = 1
     If Len(.Text) = 0 Then
        .Text = "V"
        .Col = 4   '顯示所選取之箱數
        tmpQty = .Text
        dbCut_TotalCaseQty = dbCut_TotalCaseQty + .Text
        txt_Tab1_SelectedCaseQty.Text = dbCut_TotalCaseQty
        .Col = 8: .Text = tmpQty
        txt_Tab1_CutCaseQty.Text = tmpQty
        
        '計算所選取的個數
        .Col = 14: tmpQty = .Text * tmpQty
        .Col = 12: .Text = tmpQty
        
        .Col = 5   '顯示所選取之重量
        tmpQty = .Text
        dbCut_TotalWeight = dbCut_TotalWeight + .Text
        txt_Tab1_SelectedWeight.Text = dbCut_TotalWeight
        .Col = 9: .Text = tmpQty
        
        .Col = 6   '顯示所選取之材積
        tmpQty = .Text
        dbCut_TotalVolumn = dbCut_TotalVolumn + .Text
        txt_Tab1_SelectedVolumn.Text = dbCut_TotalVolumn
        .Col = 10: .Text = tmpQty
        
        .Col = 7   '顯示所選取之板數
        tmpQty = .Text
        dbCut_TotalPalletQty = dbCut_TotalPalletQty + .Text
        txt_Tab1_SelectedPalletQty.Text = dbCut_TotalPalletQty
        .Col = 11: .Text = tmpQty
        txt_Tab1_CutPalletQty.Text = tmpQty
     Else
        .Col = 11   '切割之板數
        If Val(.Text) <> 0 Then
           txt_Tab1_CutPalletQty.Text = .Text
        End If
        .Col = 8   '切割之箱數
        If Val(.Text) <> 0 Then
           txt_Tab1_CutCaseQty.Text = .Text
        End If
     End If
     '反白選取之資料行
     .Col = 0
     For i = 0 To .Cols - 1
         .ColSel = i
     Next i
End With
End Sub

Private Sub dg_TRP02W_HeadClick(ByVal ColIndex As Integer)
'以滑鼠點選 dg_TRP02W 欄位標題區
Dim OrderFieldName As String
If TypeName(rs_TRP02W) <> "Nothing" Then
   OrderFieldName = "[" & dg_TRP02W.Columns(ColIndex).Caption & "]"
   If strOrder = "ASC" Then
      strOrder = "DESC"
      rs_TRP02W.Sort = OrderFieldName & " DESC "
   Else
      strOrder = "ASC"
      rs_TRP02W.Sort = OrderFieldName & " ASC "
   End If
End If
End Sub

Private Sub dg_TRP02W_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If blTRP02WEventEnable Then
   With dg_TRP02W
        'Do While .SelBookmarks.Count <> 0
        '   dg_TRP02W.SelBookmarks.Remove 0
        'Loop
        '反白顯示選取之資料列
        dg_TRP02W.SelBookmarks.Add rs_TRP02W.Bookmark
   End With
End If
End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "ㄧ單多車訂單切割"
End Sub

Private Sub Form_Load()
'設定 Form 大小、位置
dbsrcFormHeight = 7140
dbsrcFormWidth = 11475
Me.Height = 7650: Me.Width = 11600
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

'設定欲切割訂單之訂單名細
Call SetGrid_Format_SelectedOrderDetail

'設定已完成切割訂單列表
Call SetGrid_Format_CutOrderList
'設定已完成切割訂單明細表
Call CreateRS_CutOrderDetail

End Sub

Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '變小
   SSTab1.Height = (SSTab1.Height - (dbsrcFormHeight - Me.ScaleHeight))
   SSTab1.Width = (SSTab1.Width - (dbsrcFormWidth - Me.ScaleWidth))
   dg_TRP02W.Width = dg_TRP02W.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_TRP02W.Height = dg_TRP02W.Height - (dbsrcFormHeight - Me.ScaleHeight)
   Label1(19).Top = Label1(19).Top - (dbsrcFormHeight - Me.ScaleHeight)
   txt_Tab0_OrderCount.Top = txt_Tab0_OrderCount.Top - (dbsrcFormHeight - Me.ScaleHeight)
   Label1(3).Top = Label1(3).Top - (dbsrcFormHeight - Me.ScaleHeight)
   txt_Tab0_TotalCase.Top = txt_Tab0_TotalCase.Top - (dbsrcFormHeight - Me.ScaleHeight)
   Label1(0).Top = Label1(0).Top - (dbsrcFormHeight - Me.ScaleHeight)
   txt_Tab0_TotalWeight.Top = txt_Tab0_TotalWeight.Top - (dbsrcFormHeight - Me.ScaleHeight)
   Label1(1).Top = Label1(1).Top - (dbsrcFormHeight - Me.ScaleHeight)
   txt_Tab0_TotalVolumn.Top = txt_Tab0_TotalVolumn.Top - (dbsrcFormHeight - Me.ScaleHeight)
   Label1(2).Top = Label1(2).Top - (dbsrcFormHeight - Me.ScaleHeight)
   txt_Tab0_TotalPallet.Top = txt_Tab0_TotalPallet.Top - (dbsrcFormHeight - Me.ScaleHeight)
   
   fam_Tab1_Orders.Left = fam_Tab1_Orders.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2)
   fam_Tab1_OrderDetail.Height = fam_Tab1_OrderDetail.Height - (dbsrcFormHeight - Me.ScaleHeight)
   fam_Tab1_OrderDetail.Width = fam_Tab1_OrderDetail.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_SelectedOrderDetail.Height = dg_SelectedOrderDetail.Height - (dbsrcFormHeight - Me.ScaleHeight)
   dg_SelectedOrderDetail.Width = dg_SelectedOrderDetail.Width - (dbsrcFormWidth - Me.ScaleWidth)
   
   dg_CutOrders.Width = dg_CutOrders.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_CutOrderDetail.Width = dg_CutOrderDetail.Width - (dbsrcFormWidth - Me.ScaleWidth)
   dg_CutOrderDetail.Height = dg_CutOrderDetail.Height - (dbsrcFormHeight - Me.ScaleHeight)
   fam_Tab2_Qoery.Left = fam_Tab2_Qoery.Left - (dbsrcFormWidth - Me.ScaleWidth)
   fam_Tab2_Delete.Left = fam_Tab2_Delete.Left - (dbsrcFormWidth - Me.ScaleWidth)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   SSTab1.Height = (SSTab1.Height + (Me.ScaleHeight - dbsrcFormHeight))
   SSTab1.Width = (SSTab1.Width + (Me.ScaleWidth - dbsrcFormWidth))
   dg_TRP02W.Width = dg_TRP02W.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_TRP02W.Height = dg_TRP02W.Height + (Me.ScaleHeight - dbsrcFormHeight)
   Label1(19).Top = Label1(19).Top + (Me.ScaleHeight - dbsrcFormHeight)
   txt_Tab0_OrderCount.Top = txt_Tab0_OrderCount.Top + (Me.ScaleHeight - dbsrcFormHeight)
   Label1(3).Top = Label1(3).Top + (Me.ScaleHeight - dbsrcFormHeight)
   txt_Tab0_TotalCase.Top = txt_Tab0_TotalCase.Top + (Me.ScaleHeight - dbsrcFormHeight)
   Label1(0).Top = Label1(0).Top + (Me.ScaleHeight - dbsrcFormHeight)
   txt_Tab0_TotalWeight.Top = txt_Tab0_TotalWeight.Top + (Me.ScaleHeight - dbsrcFormHeight)
   Label1(1).Top = Label1(1).Top + (Me.ScaleHeight - dbsrcFormHeight)
   txt_Tab0_TotalVolumn.Top = txt_Tab0_TotalVolumn.Top + (Me.ScaleHeight - dbsrcFormHeight)
   Label1(2).Top = Label1(2).Top + (Me.ScaleHeight - dbsrcFormHeight)
   txt_Tab0_TotalPallet.Top = txt_Tab0_TotalPallet.Top + (Me.ScaleHeight - dbsrcFormHeight)
   
   fam_Tab1_Orders.Left = fam_Tab1_Orders.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2)
   fam_Tab1_OrderDetail.Height = fam_Tab1_OrderDetail.Height + (Me.ScaleHeight - dbsrcFormHeight)
   fam_Tab1_OrderDetail.Width = fam_Tab1_OrderDetail.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_SelectedOrderDetail.Height = dg_SelectedOrderDetail.Height + (Me.ScaleHeight - dbsrcFormHeight)
   dg_SelectedOrderDetail.Width = dg_SelectedOrderDetail.Width + (Me.ScaleWidth - dbsrcFormWidth)
   
   dg_CutOrders.Width = dg_CutOrders.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_CutOrderDetail.Width = dg_CutOrderDetail.Width + (Me.ScaleWidth - dbsrcFormWidth)
   dg_CutOrderDetail.Height = dg_CutOrderDetail.Height + (Me.ScaleHeight - dbsrcFormHeight)
   fam_Tab2_Qoery.Left = fam_Tab2_Qoery.Left + (Me.ScaleWidth - dbsrcFormWidth)
   fam_Tab2_Delete.Left = fam_Tab2_Delete.Left + (Me.ScaleWidth - dbsrcFormWidth)
   
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
End If
End Sub

Private Sub Form_Terminate()
'更新 Menu [視窗]→[已開視窗清單]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'從記憶體中移除表單，藉此引起 [Terminate] 事件
Set frm_OP_CutOrders = Nothing
End Sub

Private Sub Clear_SelectedOrderData()
'清除 [待切割訂單] Orders 資料欄位
dbCut_TotalCaseQty = 0
txt_Tab1_SelectedCaseQty.Text = ""
dbCut_TotalWeight = 0
txt_Tab1_SelectedWeight.Text = ""
dbCut_TotalVolumn = 0
txt_Tab1_SelectedVolumn.Text = ""
dbCut_TotalPalletQty = 0
txt_Tab1_SelectedPalletQty.Text = ""

txt_Tab1_CutCaseQty.Text = ""
txt_Tab1_CutPalletQty.Text = ""

txt_Tab1_Storer.Text = ""
txt_Tab1_OrderKey.Text = ""
txt_Tab1_Extern.Text = ""
txt_Tab1_OrderDate.Text = ""
txt_Tab1_DeliveryDate.Text = ""
txt_Tab1_FullName.Text = ""
txt_Tab1_Address.Text = ""
txt_Tab1_ExtraDemand1.Text = ""
txt_Tab1_ExtraDemand2.Text = ""
txt_Tab1_ZIP.Text = ""
txt_Tab1_AreaCode.Text = ""
txt_Tab1_VehicleType.Text = ""
txt_Tab1_ChannelType.Text = ""
chk_Tab1_MultiCustomer.Value = vbUnchecked
txt_Tab1_Weight.Text = ""
txt_Tab1_Volumn.Text = ""
txt_Tab1_PalletQty.Text = ""
 txt_Tab1_EXEConfirm.Text = ""
End Sub

Private Sub SetGrid_Format_SelectedOrderDetail()
'選取作為待切割訂單之項目明細
Dim sub_var1 As Integer, sub_var2 As Integer
dg_SelectedOrderDetail.Visible = False
With dg_SelectedOrderDetail
     .Rows = 2: .Cols = 15
     .FixedRows = 1
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
     '設定列表之欄位寬度
     .ColWidth(0) = 1000
     .ColWidth(1) = 300
     .ColWidth(2) = 800
     .ColWidth(3) = 2200
     .ColWidth(4) = 600
     .ColWidth(5) = 700
     .ColWidth(6) = 700
     .ColWidth(7) = 700
     .ColWidth(8) = 850
     .ColWidth(9) = 850
     .ColWidth(10) = 850
     .ColWidth(11) = 850
     .ColWidth(12) = 850
     .ColWidth(13) = 850
     .ColWidth(14) = 850
     '設定列表之標題
     .Row = 0
     .Col = 0: .Text = "項次"
     .Col = 1: .Text = "※"
     .Col = 2: .Text = "貨號"
     .Col = 3: .Text = "品名"
     .Col = 4: .Text = "箱數"
     .Col = 5: .Text = "重量"
     .Col = 6: .Text = "材積"
     .Col = 7: .Text = "板數"
     .Col = 8: .Text = "切割箱數"
     .Col = 9: .Text = "切割重量"
     .Col = 10: .Text = "切割材積"
     .Col = 11: .Text = "切割板數"
     .Col = 12: .Text = "切割個數"
     .Col = 13: .Text = "每板箱數"
     .Col = 14: .Text = "每箱個數"
     '設定列表之文字對齊
     .ColAlignment(0) = flexAlignLeftCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignRightCenter
     .ColAlignment(5) = flexAlignRightCenter
     .ColAlignment(6) = flexAlignRightCenter
     .ColAlignment(7) = flexAlignRightCenter
     .ColAlignment(8) = flexAlignRightCenter
     .ColAlignment(9) = flexAlignRightCenter
     .ColAlignment(10) = flexAlignRightCenter
     .ColAlignment(11) = flexAlignRightCenter
     .ColAlignment(12) = flexAlignRightCenter
     .ColAlignment(13) = flexAlignRightCenter
     .ColAlignment(14) = flexAlignRightCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignLeftCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_SelectedOrderDetail.Visible = True
End Sub

Private Sub Delete_GridRow(ByRef dgDataGrid As MSHFlexGrid, ByVal intRow As Double)
'待切割訂單項次(Detail) 資料刪除
If intRow = 0 Then Exit Sub

Dim i As Double, j As Integer

'1. 將刪除列資料由下一列資料取代
'   而後的資料列往上移一列
With dgDataGrid
     For i = intRow To .Rows - 2   '會有多一行空白列
         .Row = i
         For j = 0 To .Cols - 1
             .Col = j
             .Text = .TextArray((.Row + 1) * .Cols + .Col)
         Next j
         '防止最後第一列往上移給最後第二列時，會是弄白資料列，[序號] 欄位不能有值
         '有資料的列，[序號] 必須重新編號
         .Col = 0
         If Val(.Text) = 0 Then .Text = ""   'Else .Text = .Row
     Next i
'2. Grid 總列數 - 1
     .Rows = .Rows - 1
     .Row = 1
     For i = 0 To .Cols - 1
         .ColSel = i
     Next i
End With

End Sub

Private Sub Calculate_SelectedPrderDetail()
'計算選取之訂單細項：箱數，重量，才積，板數

dbCut_TotalCaseQty = 0
txt_Tab1_SelectedCaseQty.Text = ""
dbCut_TotalWeight = 0
txt_Tab1_SelectedWeight.Text = ""
dbCut_TotalVolumn = 0
txt_Tab1_SelectedVolumn.Text = ""
dbCut_TotalPalletQty = 0
txt_Tab1_SelectedPalletQty.Text = ""

Dim dbCaseQty As Double, dbWeight As Double, dbVolumn As Double, dbPalletQty As Double, tmpQty As Long, dbCutEAQty As Long, dbTiHiQty As Long, dbCasecntQty As Integer
Dim dbCutPLQty As Double, dbCutCSQty As Double
Dim i As Double
With dg_SelectedOrderDetail
     For i = 1 To .Rows - 2
         .Row = i
         .Col = 1
         If .Text <> "" Then   '被選取
            .Col = 4: dbCaseQty = Val(.Text)     '箱數
            .Col = 5: dbWeight = Val(.Text)      '重量
            .Col = 6: dbVolumn = Val(.Text)      '材積
            .Col = 7: dbPalletQty = Val(.Text)   '板數
            .Col = 12: dbCutEAQty = Val(.Text) '切割個數
            .Col = 13: dbTiHiQty = Val(.Text) '每板箱數
            .Col = 14: dbCasecntQty = Val(.Text) '每箱個數
            .Col = 11   '切割板數
            If Val(.Text) <> 0 Then '有切割板數
               dbCutPLQty = Val(.Text)
               '切割板數換算之箱數
               .Col = 8: .Text = dbCutEAQty / dbCasecntQty
               dbCut_TotalCaseQty = dbCut_TotalCaseQty + .Text
               
              '切割板數換算之重量
              .Col = 9: .Text = ((dbCutPLQty / dbPalletQty) * dbWeight)
               dbCut_TotalWeight = dbCut_TotalWeight + .Text
               
               '切割箱數換算之材積
               .Col = 10: .Text = ((dbCutPLQty / dbPalletQty) * dbVolumn)
               dbCut_TotalVolumn = dbCut_TotalVolumn + .Text
               
               dbCut_TotalPalletQty = dbCut_TotalPalletQty + dbCutPLQty
            Else
               .Col = 8   '切割箱數
               If Val(.Text) <> 0 Then
                  dbCutCSQty = Val(.Text)
                  dbCut_TotalCaseQty = dbCut_TotalCaseQty + dbCutCSQty
                 .Col = 9   '切割箱數換算之重量
                 .Text = ((dbCutCSQty / dbCaseQty) * dbWeight)
                  dbCut_TotalWeight = dbCut_TotalWeight + ((dbCutCSQty / dbCaseQty) * dbWeight)
                 .Col = 10   '切割箱數換算之材積
                 .Text = ((dbCutCSQty / dbCaseQty) * dbVolumn)
                  dbCut_TotalVolumn = dbCut_TotalVolumn + ((dbCutCSQty / dbCaseQty) * dbVolumn)
                 
                 '切割箱數換算之板數
                 .Col = 11: .Text = (dbCutEAQty / dbTiHiQty / dbCasecntQty)
                  dbCut_TotalPalletQty = dbCut_TotalPalletQty + .Text
               End If
            End If
         Else
            .Col = 9: .Text = ""
            .Col = 10: .Text = ""
         End If
     Next i
End With
'顯示選取之細項各欄位之加總值
txt_Tab1_SelectedCaseQty.Text = dbCut_TotalCaseQty
txt_Tab1_SelectedWeight.Text = dbCut_TotalWeight
txt_Tab1_SelectedVolumn.Text = dbCut_TotalVolumn
txt_Tab1_SelectedPalletQty.Text = dbCut_TotalPalletQty

End Sub
Private Sub SetGrid_Format_CutOrderList()
'已執行切割之訂單列表
Dim sub_var1 As Integer, sub_var2 As Integer
dg_CutOrders.Visible = False
With dg_CutOrders
     .Rows = 2: .Cols = 17
     .FixedRows = 1
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
     '設定列表之欄位寬度
     .ColWidth(0) = 500
     .ColWidth(1) = 1100
     .ColWidth(2) = 900
     .ColWidth(3) = 1200
     .ColWidth(4) = 800
     .ColWidth(5) = 800
     .ColWidth(6) = 800
     .ColWidth(7) = 800
     .ColWidth(8) = 800
     .ColWidth(9) = 400
     .ColWidth(10) = 500
     .ColWidth(11) = 3500
     .ColWidth(12) = 1000
     .ColWidth(13) = 700
     .ColWidth(14) = 800
     .ColWidth(15) = 800
     .ColWidth(16) = 1200
     '設定列表之標題
     .Row = 0
     .Col = 0: .Text = "項次"
     .Col = 1: .Text = "訂單編號"
     .Col = 2: .Text = "送貨日"
     .Col = 3: .Text = "客戶編號"
     .Col = 4: .Text = "貨主單號"
     .Col = 5: .Text = "箱數"
     .Col = 6: .Text = "重量"
     .Col = 7: .Text = "材積"
     .Col = 8: .Text = "板數"
     .Col = 9: .Text = "ZIP"
     .Col = 10: .Text = "區碼"
     .Col = 11: .Text = "客戶名稱"
     .Col = 12: .Text = "訂單日"
     .Col = 13: .Text = "貨主"
     .Col = 14: .Text = "識別"
     .Col = 15: .Text = "EXE回傳"
     .Col = 16: .Text = "原始TMS單號"
     '設定列表之文字對齊
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignCenterCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignLeftCenter
     .ColAlignment(5) = flexAlignRightCenter
     .ColAlignment(6) = flexAlignRightCenter
     .ColAlignment(7) = flexAlignRightCenter
     .ColAlignment(8) = flexAlignRightCenter
     .ColAlignment(9) = flexAlignCenterCenter
     .ColAlignment(10) = flexAlignCenterCenter
     .ColAlignment(11) = flexAlignLeftCenter
     .ColAlignment(12) = flexAlignLeftCenter
     .ColAlignment(13) = flexAlignLeftCenter
     .ColAlignment(14) = flexAlignLeftCenter
     .ColAlignment(15) = flexAlignLeftCenter
     .ColAlignment(16) = flexAlignLeftCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_CutOrders.Visible = True
End Sub

Private Sub CreateRS_CutOrderDetail()
'已執行切割之訂單明細
Call ReDim_Recordset(rs_CutOrderDetail)
With rs_CutOrderDetail
     .Fields.Append "訂單編號", adVarChar, 10
     .Fields.Append "項次", adDouble
     .Fields.Append "貨號", adVarChar, 20
     .Fields.Append "品名", adVarChar, 60
     .Fields.Append "箱數", adDouble
     .Fields.Append "重量", adDouble
     .Fields.Append "材積", adDouble
     .Fields.Append "板數", adDouble
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With
Set dg_CutOrderDetail.DataSource = rs_CutOrderDetail
'設定顯示欄位
With dg_CutOrderDetail
    .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
    .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
    .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
    .RowHeight = 250                '設定DataGrid 控制項中所有資料列的高
    .Columns(0).Width = 1000        '訂單編號
    .Columns(0).Alignment = dbgLeft
    .Columns(1).Width = 800         '項次
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 2000         '貨號
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 2400        '品名
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 800         '箱數
    .Columns(4).Alignment = dbgRight
    .Columns(5).Width = 800         '重量
    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 800         '材積
    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 800         '板數
    .Columns(7).Alignment = dbgRight
End With
End Sub

Private Sub InsertInto_CutOrderDetail(strOrderkey As String, SeqNo As Double)
'將 [待切割訂單]-訂單明細 目前資料列
'  轉入 [切割訂單明細] 之明細項次 Recordset
rs_CutOrderDetail.AddNew
rs_CutOrderDetail.Fields("訂單編號").Value = strOrderkey
rs_CutOrderDetail.Fields("項次").Value = SeqNo
With dg_SelectedOrderDetail
     .Col = 2
     rs_CutOrderDetail.Fields("貨號").Value = .Text
     .Col = 3
     rs_CutOrderDetail.Fields("品名").Value = .Text
     .Col = 8
     rs_CutOrderDetail.Fields("箱數").Value = .Text
     .Col = 9
     rs_CutOrderDetail.Fields("重量").Value = .Text
     .Col = 10
     rs_CutOrderDetail.Fields("材積").Value = .Text
     .Col = 11
     rs_CutOrderDetail.Fields("板數").Value = .Text
End With
rs_CutOrderDetail.Update
End Sub

Public Sub frm_OP_CutOrders_rsFilterAndSort(ByVal strCode As String, ByVal strReturn As String)
'表單公用副程式，由 frm_RS_FilterAndSort 表單呼叫
'傳入值：strCode      動作識別碼
'                     [FILTER] 自訂篩選    [SORT] 排序
'        strReturn    篩選 or 排序 之設定字串

Select Case strCode
       Case "FILTER"  '自訂篩選
            Select Case UCase(strRSName_FilterAndSort)
                   Case "RS_TRP02W"   '庫存查詢明細資料
                        blTRP02WEventEnable = False
                        rs_TRP02W.Filter = adFilterNone
                        rs_TRP02W.Filter = strReturn
                        If rs_TRP02W.RecordCount = 0 Then
                           msg_text = "抱歉ㄟ，找不到符合條件的資料喔"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_TRP02W.Filter = adFilterNone
                           rs_TRP02W.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
                           blTRP02WEventEnable = True
                           Exit Sub
                        End If
                        blTRP02WEventEnable = True
            End Select
       Case "SORT"    '排序
            Select Case UCase(strRSName_FilterAndSort)
                   Case "RS_TRP02W"   '倉租計算明細資料
                        blTRP02WEventEnable = False
                        rs_TRP02W.Sort = strReturn
                        blTRP02WEventEnable = True
            End Select
End Select
End Sub

Private Sub txt_Tab2_Extern_KeyPress(KeyAscii As Integer)
'訂單切割明細 + 查詢 >> 貨主單號
If KeyAscii = vbKeyReturn Then
   cmd_Tab2_ExternQuery.SetFocus
End If
End Sub
