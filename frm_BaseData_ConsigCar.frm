VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_BaseData_ConsigCar 
   Caption         =   "客戶/車輛/貨運公司 基本資料維護作業"
   ClientHeight    =   8130
   ClientLeft      =   285
   ClientTop       =   750
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15630
   ScaleWidth      =   28560
   Begin TabDlg.SSTab SSTab1 
      Height          =   8040
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   14182
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   14215660
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "客戶資料"
      TabPicture(0)   =   "frm_BaseData_ConsigCar.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmd_Tab0_ConsigneeShow"
      Tab(0).Control(1)=   "fam_Tab0_Consignee"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "cmd_Tab0_2Excel"
      Tab(0).Control(4)=   "cmd_Tab0_ConsigneeQuery"
      Tab(0).Control(5)=   "dg_Tab0_ConsigneeList"
      Tab(0).Control(6)=   "Label1(0)"
      Tab(0).Control(7)=   "Shape1(1)"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "車輛資料"
      TabPicture(1)   =   "frm_BaseData_ConsigCar.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmd_Tab1_CarShow"
      Tab(1).Control(1)=   "fam_Tab1_Car"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "cmd_Tab1_2Excel"
      Tab(1).Control(4)=   "cmd_Tab1_CarQuery"
      Tab(1).Control(5)=   "dg_Tab1_CarList"
      Tab(1).Control(6)=   "Label1(1)"
      Tab(1).Control(7)=   "Shape1(0)"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "貨運公司"
      TabPicture(2)   =   "frm_BaseData_ConsigCar.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fam_Tab2_FunctionArea"
      Tab(2).Control(1)=   "Frame1"
      Tab(2).Control(2)=   "fam_Tab2_Company"
      Tab(2).Control(3)=   "dg_Tab2_TRPCompanyList"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   " "
      TabPicture(3)   =   "frm_BaseData_ConsigCar.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "Frame6"
      Tab(3).Control(2)=   "fam_Tab3_Sku"
      Tab(3).Control(3)=   "dg_Tab3_SkuList"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "允收天數資料"
      TabPicture(4)   =   "frm_BaseData_ConsigCar.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "dg_Tab4_AcceptableList"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "fam_Tab4_Acceptable"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame4"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame8"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "CmnDialog"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Frame7"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "允收天數匯入"
      TabPicture(5)   =   "frm_BaseData_ConsigCar.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "gd_Tab5_AaccessTable"
      Tab(5).Control(1)=   "Frame"
      Tab(5).ControlCount=   2
      Begin VB.Frame Frame7 
         Appearance      =   0  '平面
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   550
         Left            =   8280
         TabIndex        =   1
         Top             =   2880
         Width           =   2025
         Begin VB.CommandButton cmd_Tab4_SaveToExcel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "查詢結果轉Excel"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   50
            Style           =   1  '圖片外觀
            TabIndex        =   2
            Top             =   120
            Width           =   1950
         End
      End
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   7680
         Top             =   2880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_Tab0_ConsigneeShow 
         BackColor       =   &H00FFC0C0&
         Caption         =   "顯示所有客戶"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -73005
         Style           =   1  '圖片外觀
         TabIndex        =   223
         Top             =   390
         Width           =   1830
      End
      Begin VB.CommandButton cmd_Tab1_CarShow 
         BackColor       =   &H00FFC0FF&
         Caption         =   "顯示所有車輛"
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
         Left            =   -73020
         Style           =   1  '圖片外觀
         TabIndex        =   222
         Top             =   405
         Width           =   1830
      End
      Begin VB.Frame fam_Tab0_Consignee 
         Appearance      =   0  '平面
         BackColor       =   &H8000000C&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   6360
         Left            =   -71160
         TabIndex        =   161
         Top             =   1545
         Width           =   7440
         Begin VB.ComboBox cmb_Tab0_ExtraDemand2 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   1020
            Style           =   2  '單純下拉式
            TabIndex        =   192
            Top             =   3615
            Width           =   6000
         End
         Begin VB.ComboBox cmb_Tab0_ExtraDemand1 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   1020
            Style           =   2  '單純下拉式
            TabIndex        =   191
            Top             =   3210
            Width           =   6000
         End
         Begin VB.ComboBox cmb_Tab0_VehicleType 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   1020
            Style           =   2  '單純下拉式
            TabIndex        =   190
            Top             =   2805
            Width           =   6000
         End
         Begin VB.TextBox txt_Tab0_Phone 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   3180
            MaxLength       =   30
            TabIndex        =   189
            Top             =   2430
            Width           =   1575
         End
         Begin VB.TextBox txt_Tab0_Contact 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            TabIndex        =   188
            Top             =   2430
            Width           =   1575
         End
         Begin VB.TextBox txt_Tab0_Fax 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   5355
            MaxLength       =   18
            TabIndex        =   187
            Top             =   2430
            Width           =   1695
         End
         Begin VB.TextBox txt_Tab0_CodeDate1 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            MaxLength       =   10
            TabIndex        =   186
            ToolTipText     =   "麒麟啤酒允收期;P&G允收期"
            Top             =   5160
            Width           =   840
         End
         Begin VB.TextBox txt_Tab0_CodeDate2 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   2925
            MaxLength       =   10
            TabIndex        =   185
            ToolTipText     =   "麒麟清酒允收期"
            Top             =   5160
            Width           =   840
         End
         Begin VB.TextBox txt_Tab0_Notes 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            TabIndex        =   184
            Top             =   5880
            Width           =   5880
         End
         Begin VB.TextBox txt_Tab0_PalletSpec 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   2925
            MaxLength       =   20
            TabIndex        =   183
            Top             =   5520
            Width           =   3960
         End
         Begin VB.TextBox txt_Tab0_PalletType 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            MaxLength       =   20
            TabIndex        =   182
            Top             =   5520
            Width           =   840
         End
         Begin VB.TextBox txt_Tab0_Stamp 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1005
            MaxLength       =   1
            TabIndex        =   181
            Top             =   4800
            Width           =   600
         End
         Begin VB.TextBox txt_Tab0_Penalties 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   2445
            MaxLength       =   1
            TabIndex        =   180
            Top             =   4800
            Width           =   600
         End
         Begin VB.TextBox txt_Tab0_Channel 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   3060
            MaxLength       =   20
            TabIndex        =   179
            Top             =   4440
            Width           =   1200
         End
         Begin VB.TextBox txt_Tab0_ChannelType 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            MaxLength       =   10
            TabIndex        =   178
            Top             =   4440
            Width           =   1200
         End
         Begin VB.TextBox txt_Tab0_ConsigneeKey 
            BackColor       =   &H8000000F&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1020
            TabIndex        =   177
            Top             =   240
            Width           =   1980
         End
         Begin VB.ComboBox cmb_Tab0_Storer 
            BackColor       =   &H00C0FFC0&
            Enabled         =   0   'False
            Height          =   300
            Left            =   3525
            Style           =   2  '單純下拉式
            TabIndex        =   176
            Top             =   240
            Width           =   3495
         End
         Begin VB.TextBox txt_Tab0_FullName 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            MaxLength       =   60
            TabIndex        =   175
            Top             =   600
            Width           =   6000
         End
         Begin VB.ComboBox cmb_Tab0_Zip 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   1020
            Style           =   2  '單純下拉式
            TabIndex        =   174
            Top             =   960
            Width           =   1995
         End
         Begin VB.TextBox txt_Tab0_Class 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   3885
            MaxLength       =   10
            TabIndex        =   173
            ToolTipText     =   "樓層補貼欄位請輸入數字，系統預設=0"
            Top             =   960
            Width           =   705
         End
         Begin VB.TextBox txt_Tab0_GridCode 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   5580
            MaxLength       =   5
            TabIndex        =   172
            Top             =   960
            Width           =   1395
         End
         Begin VB.TextBox txt_Tab0_ShortName 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            MaxLength       =   60
            TabIndex        =   171
            Top             =   1680
            Width           =   6000
         End
         Begin VB.ComboBox cmb_Tab0_AreaCode 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   1020
            Style           =   2  '單純下拉式
            TabIndex        =   170
            Top             =   1320
            Width           =   6000
         End
         Begin VB.TextBox txt_Tab0_Address 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1020
            TabIndex        =   169
            Top             =   2040
            Width           =   6000
         End
         Begin VB.ComboBox cmb_Tab0_PickTool 
            BackColor       =   &H00C0FFC0&
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
            Left            =   1020
            Style           =   2  '單純下拉式
            TabIndex        =   168
            Top             =   4080
            Width           =   1890
         End
         Begin VB.TextBox txt_Tab0_UnLoad 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   4080
            TabIndex        =   167
            Top             =   4080
            Width           =   420
         End
         Begin VB.CheckBox chk_Tab0_MultiCustomer 
            BackColor       =   &H8000000C&
            Caption         =   "指送客戶"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   4440
            TabIndex        =   166
            Top             =   4800
            Width           =   1140
         End
         Begin VB.CheckBox chkDC 
            BackColor       =   &H8000000C&
            Caption         =   "統倉客戶"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   3240
            TabIndex        =   165
            ToolTipText     =   "麒麟不足公斤補貼標準"
            Top             =   4800
            Width           =   1260
         End
         Begin VB.ComboBox cmdCodeDateRate 
            BackColor       =   &H00C0FFC0&
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
            ItemData        =   "frm_BaseData_ConsigCar.frx":00A8
            Left            =   6480
            List            =   "frm_BaseData_ConsigCar.frx":00B2
            TabIndex        =   164
            ToolTipText     =   "若未指定，雀巢eOrder轉入時預設為1/2效期"
            Top             =   5160
            Width           =   930
         End
         Begin VB.ComboBox cmb_Tab0_Group 
            BackColor       =   &H00C0FFC0&
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
            Left            =   5100
            TabIndex        =   163
            Text            =   "cmb_Tab0_Group"
            Top             =   4420
            Width           =   1875
         End
         Begin VB.TextBox txt_Tab0_CodeDate3 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   4845
            MaxLength       =   10
            TabIndex        =   162
            ToolTipText     =   "麒麟飲料允收期"
            Top             =   5160
            Width           =   840
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "特殊需求 2"
            Height          =   180
            Index           =   12
            Left            =   150
            TabIndex        =   221
            Top             =   3690
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "特殊需求 1"
            Height          =   180
            Index           =   11
            Left            =   150
            TabIndex        =   220
            Top             =   3285
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車種代碼"
            Height          =   180
            Index           =   10
            Left            =   285
            TabIndex        =   219
            Top             =   2880
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "電話"
            Height          =   180
            Index           =   9
            Left            =   2790
            TabIndex        =   218
            Top             =   2475
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "聯絡人"
            Height          =   180
            Index           =   8
            Left            =   465
            TabIndex        =   217
            Top             =   2475
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "傳真"
            Height          =   180
            Index           =   55
            Left            =   4980
            TabIndex        =   216
            Top             =   2475
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "允收期1"
            Height          =   180
            Index           =   57
            Left            =   120
            TabIndex        =   215
            Top             =   5205
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "允收期2"
            Height          =   180
            Index           =   58
            Left            =   1920
            TabIndex        =   214
            Top             =   5205
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "備註"
            Height          =   180
            Index           =   59
            Left            =   240
            TabIndex        =   213
            Top             =   5925
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "棧板規格"
            Height          =   180
            Index           =   60
            Left            =   2160
            TabIndex        =   212
            Top             =   5565
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "棧板材質"
            Height          =   180
            Index           =   61
            Left            =   240
            TabIndex        =   211
            Top             =   5565
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貼標"
            Height          =   180
            Index           =   62
            Left            =   600
            TabIndex        =   210
            Top             =   4845
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "罰款客戶"
            Height          =   180
            Index           =   63
            Left            =   1680
            TabIndex        =   209
            Top             =   4845
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "通路別"
            Height          =   180
            Index           =   45
            Left            =   2400
            TabIndex        =   208
            Top             =   4485
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "通路型態"
            Height          =   180
            Index           =   13
            Left            =   240
            TabIndex        =   207
            Top             =   4485
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶編號"
            Height          =   180
            Index           =   16
            Left            =   240
            TabIndex        =   206
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貨主"
            Height          =   180
            Index           =   17
            Left            =   3120
            TabIndex        =   205
            Top             =   300
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶名稱"
            Height          =   180
            Index           =   4
            Left            =   240
            TabIndex        =   204
            Top             =   660
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "郵遞區號"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   203
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "樓層補貼"
            Height          =   180
            Index           =   7
            Left            =   3120
            TabIndex        =   202
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "矩陣圖碼"
            Height          =   180
            Index           =   18
            Left            =   4800
            TabIndex        =   201
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶簡稱"
            Height          =   180
            Index           =   6
            Left            =   240
            TabIndex        =   200
            Top             =   1740
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運送區碼"
            Height          =   180
            Index           =   3
            Left            =   240
            TabIndex        =   199
            Top             =   1380
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運送地址"
            Height          =   180
            Index           =   5
            Left            =   240
            TabIndex        =   198
            Top             =   2085
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "搬運工具"
            Height          =   180
            Index           =   43
            Left            =   240
            TabIndex        =   197
            Top             =   4170
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "卸貨難易度"
            Height          =   180
            Index           =   15
            Left            =   3120
            TabIndex        =   196
            Top             =   4140
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "允收期限"
            Height          =   180
            Index           =   64
            Left            =   5760
            TabIndex        =   195
            Top             =   5205
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶群組"
            Height          =   180
            Index           =   69
            Left            =   4320
            TabIndex        =   194
            Top             =   4485
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "允收期3"
            Height          =   180
            Index           =   70
            Left            =   3840
            TabIndex        =   193
            Top             =   5205
            Width           =   630
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  '平面
         BackColor       =   &H00004000&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -71145
         TabIndex        =   154
         Top             =   345
         Width           =   7440
         Begin VB.CommandButton cmd_Tab0_AddNew 
            BackColor       =   &H00C0FFC0&
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
            Height          =   450
            Left            =   1290
            Style           =   1  '圖片外觀
            TabIndex        =   160
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab0_Modify 
            BackColor       =   &H00C0E0FF&
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
            Height          =   450
            Left            =   75
            Style           =   1  '圖片外觀
            TabIndex        =   159
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab0_Save 
            BackColor       =   &H00C0C0FF&
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
            Height          =   450
            Left            =   2505
            Style           =   1  '圖片外觀
            TabIndex        =   158
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab0_Delete 
            BackColor       =   &H000080FF&
            Caption         =   "刪  除"
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
            Height          =   450
            Left            =   4935
            Style           =   1  '圖片外觀
            TabIndex        =   157
            Top             =   195
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
            Height          =   450
            Index           =   0
            Left            =   6150
            Style           =   1  '圖片外觀
            TabIndex        =   156
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab0_Cancel 
            BackColor       =   &H00C0FFFF&
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
            Height          =   450
            Left            =   3720
            Style           =   1  '圖片外觀
            TabIndex        =   155
            Top             =   195
            Width           =   1200
         End
      End
      Begin VB.Frame fam_Tab1_Car 
         Appearance      =   0  '平面
         BackColor       =   &H8000000C&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   5400
         Left            =   -71145
         TabIndex        =   105
         Top             =   1545
         Width           =   7440
         Begin VB.ComboBox cmb_Tab1_AreaCode 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '單純下拉式
            TabIndex        =   127
            Top             =   1500
            Width           =   6375
         End
         Begin VB.ComboBox cmb_Tab1_ZIP 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '單純下拉式
            TabIndex        =   126
            Top             =   1110
            Width           =   1995
         End
         Begin VB.ComboBox cmb_Tab1_Company 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '單純下拉式
            TabIndex        =   125
            Top             =   2280
            Width           =   6375
         End
         Begin VB.TextBox txt_Tab1_WeightCapacity 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   945
            MaxLength       =   7
            TabIndex        =   124
            Top             =   4260
            Width           =   885
         End
         Begin VB.TextBox txt_Tab1_VolumnCapacity 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   3360
            MaxLength       =   7
            TabIndex        =   123
            Top             =   4260
            Width           =   885
         End
         Begin VB.TextBox txt_Tab1_PalletCapacity 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6060
            MaxLength       =   7
            TabIndex        =   122
            Top             =   4260
            Width           =   885
         End
         Begin VB.TextBox txt_Tab1_CarWeight 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   945
            TabIndex        =   121
            Top             =   2670
            Width           =   885
         End
         Begin VB.TextBox txt_Tab1_CarHeight 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   3375
            TabIndex        =   120
            Top             =   2670
            Width           =   885
         End
         Begin VB.ComboBox cmb_Tab1_CarBox 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '單純下拉式
            TabIndex        =   119
            Top             =   3045
            Width           =   2715
         End
         Begin VB.ComboBox cmb_Tab1_EmployType 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '單純下拉式
            TabIndex        =   118
            Top             =   3435
            Width           =   2715
         End
         Begin VB.ComboBox cmb_Tab1_UnloadType 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '單純下拉式
            TabIndex        =   117
            Top             =   3840
            Width           =   2715
         End
         Begin VB.ComboBox cmb_Tab1_VehicleType 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            Left            =   945
            Style           =   2  '單純下拉式
            TabIndex        =   116
            Top             =   1890
            Width           =   6375
         End
         Begin VB.TextBox txt_Tab1_Description 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   930
            TabIndex        =   115
            Top             =   4635
            Width           =   6255
         End
         Begin VB.ComboBox cmb_Tab1_CarType 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            ItemData        =   "frm_BaseData_ConsigCar.frx":00C0
            Left            =   5685
            List            =   "frm_BaseData_ConsigCar.frx":00D3
            TabIndex        =   114
            Top             =   1125
            Width           =   1635
         End
         Begin VB.TextBox txt_Tab1_CarID 
            BackColor       =   &H8000000F&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   945
            TabIndex        =   113
            Top             =   240
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab1_Driver 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   975
            TabIndex        =   112
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txt_Tab1_Phone 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   2940
            TabIndex        =   111
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txt_Tab1_Receiver 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   5700
            TabIndex        =   110
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkPND 
            BackColor       =   &H8000000C&
            Caption         =   "PND到貨追蹤"
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
            Left            =   3000
            TabIndex        =   109
            ToolTipText     =   "是否使用PND到貨追蹤系統"
            Top             =   300
            Width           =   1620
         End
         Begin VB.TextBox txtAPFix 
            Alignment       =   1  '靠右對齊
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   3780
            TabIndex        =   108
            Top             =   1140
            Width           =   735
         End
         Begin VB.TextBox txtAdd 
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   915
            TabIndex        =   107
            ToolTipText     =   "使用者 / 時間"
            Top             =   5040
            Width           =   2805
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H8000000E&
            Enabled         =   0   'False
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   4395
            TabIndex        =   106
            ToolTipText     =   "使用者 / 時間"
            Top             =   5040
            Width           =   2805
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運送區碼"
            Height          =   180
            Index           =   23
            Left            =   150
            TabIndex        =   153
            Top             =   1575
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貨運公司"
            Height          =   180
            Index           =   24
            Left            =   150
            TabIndex        =   152
            Top             =   2370
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "裝載重量"
            Height          =   180
            Index           =   25
            Left            =   150
            TabIndex        =   151
            Top             =   4320
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "裝載材積"
            Height          =   180
            Index           =   26
            Left            =   2610
            TabIndex        =   150
            Top             =   4335
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "裝載板數"
            Height          =   180
            Index           =   27
            Left            =   5340
            TabIndex        =   149
            Top             =   4320
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "總重"
            Height          =   180
            Index           =   28
            Left            =   510
            TabIndex        =   148
            Top             =   2730
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車床高度"
            Height          =   180
            Index           =   29
            Left            =   2595
            TabIndex        =   147
            Top             =   2730
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車廂形式"
            Height          =   180
            Index           =   30
            Left            =   150
            TabIndex        =   146
            Top             =   3135
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "僱用方式"
            Height          =   180
            Index           =   31
            Left            =   150
            TabIndex        =   145
            Top             =   3525
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "裝卸方式"
            Height          =   180
            Index           =   32
            Left            =   150
            TabIndex        =   144
            Top             =   3930
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車    種"
            Height          =   180
            Index           =   33
            Left            =   330
            TabIndex        =   143
            Top             =   1980
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "說明"
            Height          =   180
            Index           =   34
            Left            =   495
            TabIndex        =   142
            Top             =   4695
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "郵遞區號"
            Height          =   180
            Index           =   22
            Left            =   150
            TabIndex        =   141
            Top             =   1185
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "計費類別"
            Height          =   180
            Index           =   44
            Left            =   4920
            TabIndex        =   140
            Top             =   1185
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "車牌號碼"
            Height          =   180
            Index           =   19
            Left            =   120
            TabIndex        =   139
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "駕駛人"
            Height          =   180
            Index           =   20
            Left            =   360
            TabIndex        =   138
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "電話"
            Height          =   180
            Index           =   21
            Left            =   2520
            TabIndex        =   137
            Top             =   780
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "請款人"
            Height          =   180
            Index           =   56
            Left            =   5040
            TabIndex        =   136
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "KG"
            Height          =   180
            Index           =   14
            Left            =   1920
            TabIndex        =   135
            Top             =   2730
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "KG"
            Height          =   180
            Index           =   65
            Left            =   1920
            TabIndex        =   134
            Top             =   4320
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "材"
            Height          =   180
            Index           =   66
            Left            =   4320
            TabIndex        =   133
            Top             =   4320
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "板"
            Height          =   180
            Index           =   67
            Left            =   7080
            TabIndex        =   132
            Top             =   4320
            Width           =   180
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "CM"
            Height          =   180
            Index           =   68
            Left            =   4320
            TabIndex        =   131
            Top             =   2730
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "運費調整"
            Height          =   180
            Index           =   71
            Left            =   3000
            TabIndex        =   130
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "新增"
            Height          =   180
            Index           =   72
            Left            =   480
            TabIndex        =   129
            Top             =   5100
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "異動"
            Height          =   180
            Index           =   73
            Left            =   3960
            TabIndex        =   128
            Top             =   5100
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '平面
         BackColor       =   &H00404080&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -71130
         TabIndex        =   98
         Top             =   360
         Width           =   7485
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
            Height          =   450
            Index           =   1
            Left            =   6195
            Style           =   1  '圖片外觀
            TabIndex        =   104
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab1_Delete 
            BackColor       =   &H000080FF&
            Caption         =   "刪  除"
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
            Height          =   450
            Left            =   4980
            Style           =   1  '圖片外觀
            TabIndex        =   103
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab1_Save 
            BackColor       =   &H00C0C0FF&
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
            Height          =   450
            Left            =   2535
            Style           =   1  '圖片外觀
            TabIndex        =   102
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab1_Modify 
            BackColor       =   &H00C0E0FF&
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
            Height          =   450
            Left            =   90
            Style           =   1  '圖片外觀
            TabIndex        =   101
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab1_AddNew 
            BackColor       =   &H00C0FFC0&
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
            Height          =   450
            Left            =   1320
            Style           =   1  '圖片外觀
            TabIndex        =   100
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab1_Cancel 
            BackColor       =   &H00C0FFFF&
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
            Height          =   450
            Left            =   3765
            Style           =   1  '圖片外觀
            TabIndex        =   99
            Top             =   195
            Width           =   1200
         End
      End
      Begin VB.CommandButton cmd_Tab0_2Excel 
         Appearance      =   0  '平面
         BackColor       =   &H00808080&
         Caption         =   "Excel"
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
         Left            =   -73650
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '圖片外觀
         TabIndex        =   97
         Top             =   420
         Width           =   585
      End
      Begin VB.CommandButton cmd_Tab0_ConsigneeQuery 
         BackColor       =   &H00C0FFC0&
         Caption         =   "客戶搜尋"
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
         Left            =   -74775
         Style           =   1  '圖片外觀
         TabIndex        =   96
         Top             =   420
         Width           =   1110
      End
      Begin VB.CommandButton cmd_Tab1_2Excel 
         Appearance      =   0  '平面
         BackColor       =   &H00808080&
         Caption         =   "Excel"
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
         Left            =   -73650
         MaskColor       =   &H00FFC0C0&
         Style           =   1  '圖片外觀
         TabIndex        =   95
         Top             =   420
         Width           =   585
      End
      Begin VB.CommandButton cmd_Tab1_CarQuery 
         BackColor       =   &H00C0FFC0&
         Caption         =   "車輛搜尋"
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
         Left            =   -74775
         Style           =   1  '圖片外觀
         TabIndex        =   94
         Top             =   420
         Width           =   1110
      End
      Begin VB.Frame fam_Tab2_FunctionArea 
         Appearance      =   0  '平面
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -74910
         TabIndex        =   90
         Top             =   435
         Width           =   3795
         Begin VB.CommandButton cmd_Tab2_CompanyShow 
            BackColor       =   &H00FFC0FF&
            Caption         =   "顯示所有公司"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   75
            Style           =   1  '圖片外觀
            TabIndex        =   93
            Top             =   180
            Width           =   1830
         End
         Begin VB.CommandButton cmd_Tab2_CarQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "資料搜尋"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1950
            Style           =   1  '圖片外觀
            TabIndex        =   92
            Top             =   180
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab2_TRPCompanyReset 
            Appearance      =   0  '平面
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
            Height          =   495
            Left            =   3105
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   91
            Top             =   180
            Width           =   585
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '平面
         BackColor       =   &H00004040&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -71100
         TabIndex        =   83
         Top             =   435
         Width           =   7440
         Begin VB.CommandButton cmd_Tab2_Cancel 
            BackColor       =   &H00C0FFFF&
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
            Height          =   450
            Left            =   3735
            Style           =   1  '圖片外觀
            TabIndex        =   89
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab2_AddNew 
            BackColor       =   &H00C0FFC0&
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
            Height          =   450
            Left            =   1305
            Style           =   1  '圖片外觀
            TabIndex        =   88
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab2_Modify 
            BackColor       =   &H00C0E0FF&
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
            Height          =   450
            Left            =   75
            Style           =   1  '圖片外觀
            TabIndex        =   87
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab2_Save 
            BackColor       =   &H00C0C0FF&
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
            Height          =   450
            Left            =   2520
            Style           =   1  '圖片外觀
            TabIndex        =   86
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab2_Delete 
            BackColor       =   &H000080FF&
            Caption         =   "刪  除"
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
            Height          =   450
            Left            =   4935
            Style           =   1  '圖片外觀
            TabIndex        =   85
            Top             =   195
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
            Height          =   450
            Index           =   2
            Left            =   6165
            Style           =   1  '圖片外觀
            TabIndex        =   84
            Top             =   195
            Width           =   1200
         End
      End
      Begin VB.Frame fam_Tab2_Company 
         BackColor       =   &H8000000C&
         Height          =   2355
         Left            =   -74730
         TabIndex        =   66
         Top             =   1215
         Width           =   10890
         Begin VB.TextBox txt_Tab2_CompanyCode 
            BackColor       =   &H8000000F&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1050
            TabIndex        =   74
            Top             =   285
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab2_CName 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   73
            Top             =   780
            Width           =   4620
         End
         Begin VB.TextBox txt_Tab2_EName 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   72
            Top             =   1155
            Width           =   4620
         End
         Begin VB.TextBox txt_Tab2_Address 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   71
            Top             =   1530
            Width           =   4620
         End
         Begin VB.TextBox txt_Tab2_ShortName 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   70
            Top             =   1530
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab2_Contact 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   69
            Top             =   780
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab2_Phone 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   68
            Top             =   1155
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab2_Descr 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   67
            Top             =   1920
            Width           =   4620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "公司代碼"
            Height          =   180
            Index           =   35
            Left            =   225
            TabIndex        =   82
            Top             =   405
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "中文名稱"
            Height          =   180
            Index           =   36
            Left            =   225
            TabIndex        =   81
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "英文名稱"
            Height          =   180
            Index           =   37
            Left            =   225
            TabIndex        =   80
            Top             =   1215
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "地  址"
            Height          =   180
            Index           =   38
            Left            =   495
            TabIndex        =   79
            Top             =   1590
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "簡  稱"
            Height          =   180
            Index           =   39
            Left            =   6240
            TabIndex        =   78
            Top             =   1590
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "聯絡人"
            Height          =   180
            Index           =   40
            Left            =   6150
            TabIndex        =   77
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "電  話"
            Height          =   180
            Index           =   41
            Left            =   6240
            TabIndex        =   76
            Top             =   1215
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "說  明"
            Height          =   180
            Index           =   42
            Left            =   495
            TabIndex        =   75
            Top             =   1980
            Width           =   450
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  '平面
         BackColor       =   &H00004040&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -71070
         TabIndex        =   59
         Top             =   480
         Width           =   7440
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
            Height          =   450
            Index           =   3
            Left            =   6165
            Style           =   1  '圖片外觀
            TabIndex        =   65
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab3_Delete 
            BackColor       =   &H000080FF&
            Caption         =   "刪  除"
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
            Height          =   450
            Left            =   4935
            Style           =   1  '圖片外觀
            TabIndex        =   64
            Top             =   195
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab3_Save 
            BackColor       =   &H00C0C0FF&
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
            Height          =   450
            Left            =   2520
            Style           =   1  '圖片外觀
            TabIndex        =   63
            Top             =   195
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab3_Modify 
            BackColor       =   &H00C0E0FF&
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
            Height          =   450
            Left            =   75
            Style           =   1  '圖片外觀
            TabIndex        =   62
            Top             =   195
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab3_AddNew 
            BackColor       =   &H00C0FFC0&
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
            Height          =   450
            Left            =   1305
            Style           =   1  '圖片外觀
            TabIndex        =   61
            Top             =   195
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab3_Cancel 
            BackColor       =   &H00C0FFFF&
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
            Height          =   450
            Left            =   3735
            Style           =   1  '圖片外觀
            TabIndex        =   60
            Top             =   195
            Visible         =   0   'False
            Width           =   1200
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  '平面
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   -74880
         TabIndex        =   55
         Top             =   480
         Width           =   3795
         Begin VB.CommandButton cmd_Tab2_SkuReset 
            Appearance      =   0  '平面
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
            Height          =   495
            Left            =   3105
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   58
            Top             =   180
            Width           =   585
         End
         Begin VB.CommandButton cmd_Tab2_SkuQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "資料搜尋"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1950
            Style           =   1  '圖片外觀
            TabIndex        =   57
            Top             =   180
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab2_SkuShow 
            BackColor       =   &H00FFC0FF&
            Caption         =   "顯示全部"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   75
            Style           =   1  '圖片外觀
            TabIndex        =   56
            Top             =   180
            Width           =   1830
         End
      End
      Begin VB.Frame fam_Tab3_Sku 
         BackColor       =   &H8000000C&
         Height          =   2355
         Left            =   -74820
         TabIndex        =   36
         Top             =   1260
         Width           =   10890
         Begin VB.TextBox txt_Tab3_Sku 
            BackColor       =   &H8000000F&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1050
            TabIndex        =   45
            Top             =   285
            Width           =   1365
         End
         Begin VB.TextBox txt_Tab3_DESCR 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   44
            Top             =   780
            Width           =   4620
         End
         Begin VB.TextBox txt_Tab3_NOTES1 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   43
            Top             =   1155
            Width           =   4620
         End
         Begin VB.TextBox txt_Tab3_NOTES2 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   42
            Top             =   1530
            Width           =   4620
         End
         Begin VB.TextBox txt_Tab3_BUSR1 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   41
            Top             =   1530
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab3_STDGROSSWGT 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   40
            Top             =   780
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab3_BUSR4 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   39
            Top             =   1155
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab3_SKUGROUP 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   6765
            TabIndex        =   38
            Top             =   1920
            Width           =   1965
         End
         Begin VB.TextBox txt_Tab3_StorerKey 
            BackColor       =   &H8000000E&
            ForeColor       =   &H00FF0000&
            Height          =   270
            Left            =   1050
            TabIndex        =   37
            Top             =   1920
            Width           =   1965
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貨號代碼"
            Height          =   180
            Index           =   53
            Left            =   225
            TabIndex        =   54
            Top             =   405
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "中文名稱"
            Height          =   180
            Index           =   52
            Left            =   225
            TabIndex        =   53
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "備註一"
            Height          =   180
            Index           =   51
            Left            =   225
            TabIndex        =   52
            Top             =   1215
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "備註二"
            Height          =   180
            Index           =   50
            Left            =   225
            TabIndex        =   51
            Top             =   1590
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "產品別"
            Height          =   180
            Index           =   49
            Left            =   6150
            TabIndex        =   50
            Top             =   1590
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "每箱重"
            Height          =   180
            Index           =   48
            Left            =   6150
            TabIndex        =   49
            Top             =   840
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "每箱材"
            Height          =   180
            Index           =   47
            Left            =   6150
            TabIndex        =   48
            Top             =   1215
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "類別"
            Height          =   180
            Index           =   46
            Left            =   6150
            TabIndex        =   47
            Top             =   1980
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貨主"
            Height          =   180
            Index           =   54
            Left            =   225
            TabIndex        =   46
            Top             =   1980
            Width           =   360
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  '平面
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   0
         TabIndex        =   32
         Top             =   360
         Width           =   3795
         Begin VB.CommandButton cmd_Tab4_AcceptableShow 
            BackColor       =   &H00FFC0FF&
            Caption         =   "顯示所有資料"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   75
            Style           =   1  '圖片外觀
            TabIndex        =   35
            Top             =   180
            Width           =   1830
         End
         Begin VB.CommandButton cmd_Tab4_AcceptableQuery 
            BackColor       =   &H00C0FFC0&
            Caption         =   "資料搜尋"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1920
            Style           =   1  '圖片外觀
            TabIndex        =   34
            Top             =   180
            Width           =   1110
         End
         Begin VB.CommandButton cmd_Tab4_AcceptableReset 
            Appearance      =   0  '平面
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
            Height          =   495
            Left            =   3105
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   33
            Top             =   180
            Width           =   585
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  '平面
         BackColor       =   &H00004000&
         ForeColor       =   &H80000008&
         Height          =   750
         Left            =   3720
         TabIndex        =   25
         Top             =   360
         Width           =   7440
         Begin VB.CommandButton cmd_Tab4_Cancel 
            BackColor       =   &H00C0FFFF&
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
            Height          =   450
            Left            =   3720
            Style           =   1  '圖片外觀
            TabIndex        =   31
            Top             =   195
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
            Height          =   450
            Index           =   4
            Left            =   6150
            Style           =   1  '圖片外觀
            TabIndex        =   30
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab4_Delete 
            BackColor       =   &H000080FF&
            Caption         =   "刪  除"
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
            Height          =   450
            Left            =   4935
            Style           =   1  '圖片外觀
            TabIndex        =   29
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab4_Save 
            BackColor       =   &H00C0C0FF&
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
            Height          =   450
            Left            =   2505
            Style           =   1  '圖片外觀
            TabIndex        =   28
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab4_Modify 
            BackColor       =   &H00C0E0FF&
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
            Height          =   450
            Left            =   75
            Style           =   1  '圖片外觀
            TabIndex        =   27
            Top             =   195
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab4_AddNew 
            BackColor       =   &H00C0FFC0&
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
            Height          =   450
            Left            =   1290
            Style           =   1  '圖片外觀
            TabIndex        =   26
            Top             =   195
            Width           =   1200
         End
      End
      Begin VB.Frame fam_Tab4_Acceptable 
         Appearance      =   0  '平面
         BackColor       =   &H8000000C&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2385
         Left            =   0
         TabIndex        =   11
         Top             =   1200
         Width           =   11100
         Begin VB.TextBox txt_Tab4_ConsigneeKey 
            BackColor       =   &H8000000F&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1080
            TabIndex        =   17
            Top             =   915
            Width           =   1980
         End
         Begin VB.ComboBox cmb_Tab4_Storer 
            BackColor       =   &H00C0FFC0&
            Height          =   300
            ItemData        =   "frm_BaseData_ConsigCar.frx":00F9
            Left            =   1080
            List            =   "frm_BaseData_ConsigCar.frx":00FB
            Style           =   2  '單純下拉式
            TabIndex        =   16
            Top             =   330
            Width           =   1935
         End
         Begin VB.TextBox txt_Tab4_FullName 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  '沒有框線
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   340
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   925
            Width           =   6000
         End
         Begin VB.TextBox txt_Tab4_Sku 
            BackColor       =   &H8000000F&
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
            ForeColor       =   &H00404000&
            Height          =   360
            Left            =   1080
            TabIndex        =   14
            Top             =   1320
            Width           =   1980
         End
         Begin VB.TextBox txt_Tab4_AllowDays 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   1080
            MaxLength       =   4
            TabIndex        =   13
            Top             =   1800
            Width           =   1080
         End
         Begin VB.TextBox txt_Tab4_DESCR 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  '沒有框線
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   340
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   1330
            Width           =   6000
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "貨        主"
            Height          =   180
            Index           =   79
            Left            =   285
            TabIndex        =   24
            Top             =   390
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶編號"
            Height          =   180
            Index           =   78
            Left            =   285
            TabIndex        =   23
            Top             =   1035
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "客戶名稱"
            Height          =   180
            Index           =   77
            Left            =   3525
            TabIndex        =   22
            Top             =   1020
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "產品名稱"
            Height          =   180
            Index           =   76
            Left            =   3525
            TabIndex        =   21
            Top             =   1425
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "產品編號"
            Height          =   180
            Index           =   75
            Left            =   285
            TabIndex        =   20
            Top             =   1440
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "允收天數"
            Height          =   180
            Index           =   74
            Left            =   285
            TabIndex        =   19
            Top             =   1940
            Width           =   720
         End
         Begin VB.Label Label2 
            BackStyle       =   0  '透明
            Caption         =   "客 戶 允 收 天 數 資 料 維 護 作 業"
            BeginProperty Font 
               Name            =   "標楷體"
               Size            =   15.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   375
            Left            =   4230
            TabIndex        =   18
            Top             =   320
            Width           =   6135
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "允收天數匯入"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   11040
         Begin VB.CommandButton cmdImportT5 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3720
            Style           =   1  '圖片外觀
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT5 
            Height          =   300
            Left            =   135
            TabIndex        =   8
            ToolTipText     =   "Local Drive List"
            Top             =   240
            Width           =   2640
         End
         Begin VB.DirListBox dirLocalDirT5 
            Height          =   1560
            Left            =   135
            TabIndex        =   7
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   5640
         End
         Begin VB.FileListBox filLocalFileT5 
            Height          =   1890
            Left            =   5880
            Pattern         =   "*.xls"
            TabIndex        =   6
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   240
            Width           =   4950
         End
         Begin VB.CommandButton cmdOpenFilesT5 
            BackColor       =   &H0080FFFF&
            Caption         =   "開啟"
            Height          =   375
            Left            =   4800
            Style           =   1  '圖片外觀
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
      End
      Begin MSDataGridLib.DataGrid gd_Tab5_AaccessTable 
         Height          =   3945
         Left            =   -74880
         TabIndex        =   3
         Top             =   2760
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   6959
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid dg_Tab4_AcceptableList 
         Height          =   3360
         Left            =   0
         TabIndex        =   10
         Top             =   3600
         Width           =   11100
         _ExtentX        =   19579
         _ExtentY        =   5927
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
      Begin MSDataGridLib.DataGrid dg_Tab0_ConsigneeList 
         Height          =   7080
         Left            =   -74820
         TabIndex        =   224
         Top             =   810
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   12488
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
      Begin MSDataGridLib.DataGrid dg_Tab1_CarList 
         Height          =   6120
         Left            =   -74820
         TabIndex        =   225
         Top             =   810
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   10795
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid dg_Tab2_TRPCompanyList 
         Height          =   3345
         Left            =   -74715
         TabIndex        =   226
         Top             =   3585
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   5900
         _Version        =   393216
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
      Begin MSDataGridLib.DataGrid dg_Tab3_SkuList 
         Height          =   3345
         Left            =   -74805
         TabIndex        =   227
         Top             =   3630
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   5900
         _Version        =   393216
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "客 戶 基 本 資 料 維 護 作 業"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   0
         Left            =   -69915
         TabIndex        =   229
         Top             =   1185
         Width           =   5070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "運 輸 車 輛 基 本 資 料 維 護 作 業"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   15.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   1
         Left            =   -70395
         TabIndex        =   228
         Top             =   1185
         Width           =   6120
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00400040&
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         Height          =   435
         Index           =   0
         Left            =   -74820
         Top             =   405
         Width           =   1800
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  '不透明
         BorderColor     =   &H00400040&
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         Height          =   465
         Index           =   1
         Left            =   -74805
         Top             =   390
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frm_BaseData_ConsigCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private dbsrcFormHeight As Double    'Form 設計時期的高
Private dbsrcFormWidth As Double     'Form 設計時期的寬

Private blTab0ConsignEventEnable As Boolean     '客戶資料 List 之事件 Enable 控制
Private blTab1CarEventEnable As Boolean         '車輛資料 List 之事件 Enable 控制
Private blTab2CompanyEventEnable As Boolean     '貨運公司 List 之事件 Enable 控制
Private blTab3skuEventEnable As Boolean         '貨運公司 List 之事件 Enable 控制
Private blTab4AcceptableEventEnable As Boolean  '客戶允收天數 List 之事件 Enable 控制

Private arStorer() As String            '貨主
Private arZip() As String               '郵遞區號
Private arZIPArea() As String           '郵遞區號檔設定之 AreaCode
Private arAreaCode() As String          '區域代碼
Private arVehicleType() As String       '車種類型
Private arExtraDemand() As String       '特殊需求
Private arPickTool() As String          '搬運工具
Private arCompany() As String           '車輛：貨運公司
Private arCarBox() As String            '車輛：車廂形式
Private arEmployType() As String        '車輛：僱用方式
Private arUnloadType() As String        '車輛：裝卸方式

Private rs_Tab0_ConsigneeList As ADODB.Recordset       '顯示所有客戶資料
Private rs_Tab1_CarList As ADODB.Recordset             '顯示車輛基本資料
Private rs_Tab2_TRPCompanyList As ADODB.Recordset      '顯示運輸公司基本資料
Private rs_Tab3_SkuList As ADODB.Recordset             '顯示貨號基本資料
Private rs_Tab4_AcceptableList As ADODB.Recordset      '顯示所有客戶允收天數資料

Private MyXlsAppV2 As Excel.Application     '允收天數資料轉Excel
Private rsMain As ADODB.Recordset
Private fso As Scripting.FileSystemObject

Private Sub cmb_Tab0_Zip_Click()
'客戶資料 >> 郵遞區號
If fam_Tab0_Consignee.Enabled = True Then
    If cmb_Tab0_Zip.ListIndex <> -1 Then
        Dim i As Double
        For i = 0 To cmb_Tab0_AreaCode.ListCount - 1
            If arAreaCode(i) = UCase(arZIPArea(cmb_Tab0_Zip.ListIndex)) Then
                cmb_Tab0_AreaCode.ListIndex = i
                Exit Sub
            End If
        Next i
    End If
End If
End Sub

Private Sub cmb_Tab1_ZIP_Click()
'車輛資料 >> 郵遞區號
If fam_Tab1_Car.Enabled = True Then
    If cmb_Tab1_ZIP.ListIndex <> -1 Then
        Dim i As Double
        For i = 0 To cmb_Tab1_AreaCode.ListCount - 1
            If arAreaCode(i) = arZIPArea(cmb_Tab1_ZIP.ListIndex) Then
                cmb_Tab1_AreaCode.ListIndex = i
                Exit Sub
            End If
        Next i
    End If
End If
End Sub

Private Sub cmd_Tab0_2Excel_Click()

Dim rsTmp As New ADODB.Recordset
Screen.MousePointer = 11
rsTmp.Open "select * from gv_customer ", cn
Recordset2Excel "客戶主檔", rsTmp
Set MyXlsApp = Nothing
rsTmp.Close: Set rsTmp = Nothing
Screen.MousePointer = 0

End Sub
Private Sub cmd_Tab1_2Excel_Click()

Dim rsTmp As New ADODB.Recordset
Screen.MousePointer = 11

Recordset2Excel "車輛主檔", rs_Tab1_CarList
Set MyXlsApp = Nothing

Screen.MousePointer = 0

End Sub

Private Sub cmd_Tab0_AddNew_Click()
'客戶資料 >> 轉換至新增模式
If Not rs_Tab0_ConsigneeList Is Nothing Then
    If dg_Tab0_ConsigneeList.SelBookmarks.Count > 0 Then dg_Tab0_ConsigneeList.SelBookmarks.Remove 0
End If
fam_Tab0_Consignee.BackColor = &HC0FFC0
fam_Tab0_Consignee.Enabled = True
txt_Tab0_ConsigneeKey.Enabled = True
cmb_Tab0_Storer.Enabled = True
Call Clear_ConsigneeData
cmd_Tab0_Save.Enabled = True
cmd_Tab0_Cancel.Enabled = True
cmd_Tab0_AddNew.Enabled = False
cmd_Tab0_Modify.Enabled = False
cmd_Tab0_Delete.Enabled = False
End Sub

Private Sub cmd_Tab0_Cancel_Click()
'客戶資料 >> 取消修改
Call Clear_ConsigneeData
If txt_Tab0_ConsigneeKey.Enabled = False Then
    If Not rs_Tab0_ConsigneeList Is Nothing Then
        dg_Tab0_ConsigneeList.SelBookmarks.Add rs_Tab0_ConsigneeList.Bookmark
        Call Display_SelectedConsignData(rs_Tab0_ConsigneeList.Fields("貨主").Value, rs_Tab0_ConsigneeList.Fields("客戶編號").Value)
    End If
End If
fam_Tab0_Consignee.BackColor = &H8000000C
fam_Tab0_Consignee.Enabled = False
cmd_Tab0_Cancel.Enabled = False
cmd_Tab0_Save.Enabled = False
cmd_Tab0_AddNew.Enabled = True
cmd_Tab0_Modify.Enabled = True
cmd_Tab0_Delete.Enabled = True
End Sub

Private Sub cmd_Tab0_ConsigneeQuery_Click()
'客戶資料 >> 客戶搜尋
If rs_Tab0_ConsigneeList Is Nothing Then Exit Sub
If rs_Tab0_ConsigneeList.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_Tab0_ConsigneeList"

If ShowForm_RS_FilterAndSort(rs_Tab0_ConsigneeList, "客戶資料", Me.Tag) = False Then
    MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
    Exit Sub
End If
Me.WindowState = vbNormal

End Sub

Private Sub cmd_Tab0_ConsigneeReset_Click()
'客戶基本資料 >> 取消篩選排序
'移除篩選條件，重設排序依據
If rs_Tab0_ConsigneeList Is Nothing Then Exit Sub
 blTab0ConsignEventEnable = False
 rs_Tab0_ConsigneeList.Filter = adFilterNone
 rs_Tab0_ConsigneeList.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
 blTab0ConsignEventEnable = True

End Sub

Private Sub cmd_Tab0_ConsigneeShow_Click()
'客戶資料 >> 顯示所有客戶
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab0_ConsigneeList.DataSource = Nothing
Call ReDim_Recordset(rs_Tab0_ConsigneeList)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "SELECT Rtrim(t1.StorerKey) as 貨主 , Rtrim(t1.ConsigneeKey) as 客戶編號 , Rtrim(Isnull(t1.Full_Name,'')) as 客戶名稱  " & _
          "From TRP01M t1 join trp16m t16 on t1.storerkey = t16.storerkey and t16.storer_status <> '0' Order by t1.StorerKey,ConsigneeKey"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    msg_text = "資料錯誤：查詢結果傳回 0 列客戶資料"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab0_ConsigneeList)
tmp_Rs.Close

blTab0ConsignEventEnable = False
With dg_Tab0_ConsigneeList
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab0_ConsigneeList.MoveFirst
Set dg_Tab0_ConsigneeList.DataSource = rs_Tab0_ConsigneeList
With dg_Tab0_ConsigneeList
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 600        '貨主
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       '客戶編號
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 3000       '客戶名稱
    .Columns(3).Alignment = dbgLeft
End With
blTab0ConsignEventEnable = True
Call Clear_ConsigneeData
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-客戶資料-顯示所有資料", Me.Caption, "cmd_Tab0-ConsignShow_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Delete_Click()
'客戶資料 >> 刪除
Dim blDelete As Boolean
blDelete = True
msg_text = ""

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus
Screen.MousePointer = vbHourglass

'1.檢核 TRP02W 是否有此客戶訂單資料
str_SQL = "Select Count(*) as RecCnt From TRP02W Where ConsigneeKey = '" & Trim(txt_Tab0_ConsigneeKey.Text) & "' and storerkey = '" & Left(cmb_Tab0_Storer.Text, InStr(cmb_Tab0_Storer.Text + " ", " ") - 1) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
    blDelete = False
    If msg_text = "" Then
        msg_text = "   待排車訂單 [TRP02W] 有此客戶訂單資料"
    Else
        msg_text = msg_text & vbCrLf & "   待排車訂單 [TRP02W] 有此客戶訂單資料"
    End If
End If
tmp_Rs.Close
'2.檢核 TRP02T 是否有此客戶訂單資料
str_SQL = "Select Count(*) as RecCnt From TRP02T Where ConsigneeKey = '" & Trim(txt_Tab0_ConsigneeKey.Text) & "' and storerkey = '" & Left(cmb_Tab0_Storer.Text, InStr(cmb_Tab0_Storer.Text + " ", " ") - 1) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
    blDelete = False
    If msg_text = "" Then
        msg_text = "   已排車訂單 [TRP02T] 有此客戶訂單資料"
    Else
        msg_text = msg_text & vbCrLf & "   已排車訂單 [TRP02T] 有此客戶訂單資料"
    End If
End If
tmp_Rs.Close
'3.檢核 Orders 是否有此客戶訂單資料
str_SQL = "Select Count(*) as RecCnt From Orders Where ConsigneeKey = '" & Trim(txt_Tab0_ConsigneeKey.Text) & "' and storerkey = '" & Left(cmb_Tab0_Storer.Text, InStr(cmb_Tab0_Storer.Text + " ", " ") - 1) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   訂單主檔 [Orders] 有此客戶訂單資料"
   Else
      msg_text = msg_text & vbCrLf & "   訂單主檔 [Orders] 有此客戶訂單資料"
   End If
End If
tmp_Rs.Close

'檢核是否允許進行刪除旗標值
If blDelete = False Then
   msg_text = "客戶資料無法刪除：" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

'允許刪除
str_SQL = "Delete From TRP01M Where ConsigneeKey = '" & Trim(txt_Tab0_ConsigneeKey.Text) & "' and storerkey = '" & Left(cmb_Tab0_Storer.Text, InStr(cmb_Tab0_Storer.Text + " ", " ") - 1) & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

fam_Tab0_Consignee.BackColor = &H8000000C
fam_Tab0_Consignee.Enabled = False
cmd_Tab0_Cancel.Enabled = False
cmd_Tab0_Save.Enabled = False
cmd_Tab0_AddNew.Enabled = True
cmd_Tab0_Modify.Enabled = False
cmd_Tab0_Delete.Enabled = False
'重新顯示所有客戶資料
Call cmd_Tab0_ConsigneeShow_Click

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-客戶資料-刪除", Me.Caption, "cmd_Tab0_Delete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Modify_Click()
'客戶資料 >> 轉換修改模式
'確認選取客戶資料方允許 [修改] 功能
If rs_Tab0_ConsigneeList Is Nothing Then Exit Sub
If dg_Tab0_ConsigneeList.SelBookmarks.Count <> 0 Then
   fam_Tab0_Consignee.BackColor = &HC0E0FF
   fam_Tab0_Consignee.Enabled = True
   txt_Tab0_ConsigneeKey.Enabled = False
   cmd_Tab0_Save.Enabled = True
   cmd_Tab0_Cancel.Enabled = True
   cmd_Tab0_AddNew.Enabled = False
   cmd_Tab0_Modify.Enabled = False
   cmd_Tab0_Delete.Enabled = False
End If
End Sub

Private Sub cmd_Tab0_Save_Click()
'客戶資料 >> 客戶資料存檔

If Len(RTrim((cmb_Tab0_Storer.Text))) = 0 Or Len(RTrim(txt_Tab0_ConsigneeKey)) = 0 Then MsgBox "請輸入貨主與客戶編號", 16, "注意": Exit Sub

'清除特殊字元
Call myFormExCharFilter(Me)

On Error GoTo err_Handle

'樓層補貼的檢查，一定要數字。為輸入以0計算
If Len(RTrim(txt_Tab0_Class.Text)) = 0 Then
            MsgBox "樓層補貼沒輸入，系統預設帶0", 64, "注意"
            txt_Tab0_Class.Text = 0
End If
If Not IsNumeric(txt_Tab0_Class.Text) Then
            MsgBox "樓層補貼欄位請輸入數字", 64, "注意"
            Exit Sub
End If

'判斷不可為負數
If Left(RTrim(txt_Tab0_Class.Text), 1) = "-" Then MsgBox "樓層補貼請勿輸入負號", 64, "注意": Exit Sub
'取小數點2位無條件進位，*1為了補0
If Val(txt_Tab0_Class.Text) < 1 Then
    txt_Tab0_Class.Text = ("0" & Format(txt_Tab0_Class.Text, ".##")) * 1
Else
    txt_Tab0_Class.Text = Format(txt_Tab0_Class.Text, ".##") * 1
End If

'客戶編號重複檢查
If txt_Tab0_ConsigneeKey.Enabled = True Then
    Dim rsTmp As New ADODB.Recordset
    rsTmp.Open "select consigneekey from trp01m where rtrim(consigneekey) = '" & RTrim(txt_Tab0_ConsigneeKey) & "' and rtrim(storerkey) = '" & Left(cmb_Tab0_Storer.Text, InStr(cmb_Tab0_Storer.Text + " ", " ") - 1) & "' ", cn
    If rsTmp.EOF = False Then
        MsgBox "同一貨主，新增客戶編號重複!!", 64, "注意"
           txt_Tab0_ConsigneeKey.SelStart = 0: txt_Tab0_ConsigneeKey.SelLength = Len(txt_Tab0_ConsigneeKey.Text)
       txt_Tab0_ConsigneeKey.SetFocus
       Exit Sub
    End If
End If

'存檔資料檢核
If Check_ComsigneeData = False Then Exit Sub

'LTKK01地址別重複判定
If Left(cmb_Tab0_Storer, 6) = "LTKK01" Then
    str_SQL = "select * from trp01m " & _
                "where consigneekey <> '" & RTrim(txt_Tab0_ConsigneeKey) & "' " & _
                "and substring(consigneekey , 5,20) = '" & RTrim(Mid(txt_Tab0_ConsigneeKey, 5, 20)) & "' and storerkey = 'LTKK01'"
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn
    If Not tmp_Rs.EOF Then MsgBox "貨主 LTKK01 地址別編碼太短或編號重複，已存在客戶編號(" & RTrim(tmp_Rs("consigneekey")) & ") ，客戶名稱(" & RTrim(tmp_Rs("short_name")) & ")。", vbOKOnly, "客戶主檔新增": Exit Sub
End If

Screen.MousePointer = vbHourglass
If Not (tmp_Cmd Is Nothing) Then
   Set tmp_Cmd = Nothing
End If
Set tmp_Cmd = New ADODB.Command
If tmp_para Is Nothing Then
   Set tmp_para = New ADODB.Parameter
End If

tmp_Cmd.ActiveConnection = cn
tmp_Cmd.CommandTimeout = 0    '執行時間設定：無限期等待
tmp_Cmd.CommandType = adCmdStoredProc
tmp_Cmd.CommandText = "Master_ConsigneeData_UPDATE"

'貨主
Set tmp_para = tmp_Cmd.CreateParameter("StorerKey", adChar, adParamInput, 15)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("StorerKey").Value = arStorer(cmb_Tab0_Storer.ListIndex)

'客戶編號
Set tmp_para = tmp_Cmd.CreateParameter("ConsigneeKey", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("ConsigneeKey").Value = Trim(txt_Tab0_ConsigneeKey.Text)

'郵遞區號
Set tmp_para = tmp_Cmd.CreateParameter("ZIP", adVarChar, adParamInput, 18)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab0_Zip.ListIndex <> -1 Then
   tmp_Cmd.Parameters("ZIP").Value = arZip(cmb_Tab0_Zip.ListIndex)
Else
   tmp_Cmd.Parameters("ZIP").Value = ""
End If


'運送區碼
Set tmp_para = tmp_Cmd.CreateParameter("Area_Code", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab0_AreaCode.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Area_Code").Value = arAreaCode(cmb_Tab0_AreaCode.ListIndex)
Else
   tmp_Cmd.Parameters("Area_Code").Value = Null
End If

'運送地址
Set tmp_para = tmp_Cmd.CreateParameter("Address", adVarChar, adParamInput, 200)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_Address.Text) = "" Then
   tmp_Cmd.Parameters("Address").Value = ""
Else
   tmp_Cmd.Parameters("Address").Value = Trim(txt_Tab0_Address.Text)
End If

'聯絡人 'Terry 20180123 contact 長度由30改為80
Set tmp_para = tmp_Cmd.CreateParameter("Contact", adVarChar, adParamInput, 80)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_Contact.Text) = "" Then
   tmp_Cmd.Parameters("Contact").Value = ""
Else
   tmp_Cmd.Parameters("Contact").Value = Trim(txt_Tab0_Contact.Text)
End If

'電話
Set tmp_para = tmp_Cmd.CreateParameter("Phone", adVarChar, adParamInput, 30)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_Phone.Text) = "" Then
   tmp_Cmd.Parameters("Phone").Value = ""
Else
   tmp_Cmd.Parameters("Phone").Value = Trim(txt_Tab0_Phone.Text)
End If

'客戶等級
Set tmp_para = tmp_Cmd.CreateParameter("Class", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_Class.Text) = "" Then
   tmp_Cmd.Parameters("Class").Value = Null
Else
   tmp_Cmd.Parameters("Class").Value = Trim(txt_Tab0_Class.Text)
End If

'特殊需求 1
Set tmp_para = tmp_Cmd.CreateParameter("Extra_Demand_Code", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab0_ExtraDemand1.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Extra_Demand_Code").Value = arExtraDemand(cmb_Tab0_ExtraDemand1.ListIndex)
Else
   tmp_Cmd.Parameters("Extra_Demand_Code").Value = Null
End If

'特殊需求 2
Set tmp_para = tmp_Cmd.CreateParameter("Extra_Demand_Code2", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab0_ExtraDemand2.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Extra_Demand_Code2").Value = arExtraDemand(cmb_Tab0_ExtraDemand2.ListIndex)
Else
   tmp_Cmd.Parameters("Extra_Demand_Code2").Value = Null
End If

'客戶名稱
Set tmp_para = tmp_Cmd.CreateParameter("Full_Name", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_FullName.Text) = "" Then
   tmp_Cmd.Parameters("Full_Name").Value = ""
Else
   tmp_Cmd.Parameters("Full_Name").Value = Trim(txt_Tab0_FullName.Text)
End If

'客戶簡稱
Set tmp_para = tmp_Cmd.CreateParameter("Short_Name", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_ShortName.Text) = "" Then
   tmp_Cmd.Parameters("Short_Name").Value = ""
Else
   tmp_Cmd.Parameters("Short_Name").Value = Trim(txt_Tab0_ShortName.Text)
End If

'通路型態
Set tmp_para = tmp_Cmd.CreateParameter("Channel_Type", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_ChannelType.Text) > 0 Then
   tmp_Cmd.Parameters("Channel_Type").Value = Trim(txt_Tab0_ChannelType.Text)
Else
   tmp_Cmd.Parameters("Channel_Type").Value = Null
End If

'拆櫃難易度
Set tmp_para = tmp_Cmd.CreateParameter("Unload_Type", adVarChar, adParamInput, 3)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_UnLoad.Text) = "" Then
   tmp_Cmd.Parameters("Unload_Type").Value = Null
Else
   tmp_Cmd.Parameters("Unload_Type").Value = Trim(txt_Tab0_UnLoad.Text)
End If

'Billing_Type
Set tmp_para = tmp_Cmd.CreateParameter("BILLING_TYPE", adVarChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("BILLING_TYPE").Value = Null

'Payment_Type
Set tmp_para = tmp_Cmd.CreateParameter("Payment_Type", adVarChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Payment_Type").Value = Null

'Special_Charge
Set tmp_para = tmp_Cmd.CreateParameter("Special_Charge", adVarChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Special_Charge").Value = Null

'指送客戶
Set tmp_para = tmp_Cmd.CreateParameter("Multi_Customer", adVarChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
If chk_Tab0_MultiCustomer.Value = vbChecked Then
   tmp_Cmd.Parameters("Multi_Customer").Value = "Y"
Else
   tmp_Cmd.Parameters("Multi_Customer").Value = "N"
End If

'統倉客戶
Set tmp_para = tmp_Cmd.CreateParameter("dc", adVarChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
If chkDC.Value = vbChecked Then
   tmp_Cmd.Parameters("dc").Value = "Y"
Else
   tmp_Cmd.Parameters("dc").Value = "N"
End If

'Grid_Code 矩陣圖碼
Set tmp_para = tmp_Cmd.CreateParameter("Grid_Code", adVarChar, adParamInput, 5)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_GridCode.Text) = "" Then
   tmp_Cmd.Parameters("Grid_Code").Value = Null
Else
   tmp_Cmd.Parameters("Grid_Code").Value = Trim(txt_Tab0_GridCode.Text)
End If

'車種代碼
Set tmp_para = tmp_Cmd.CreateParameter("Vehicle_Type", adVarChar, adParamInput, 2)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab0_VehicleType.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Vehicle_Type").Value = arVehicleType(cmb_Tab0_VehicleType.ListIndex)
Else
   tmp_Cmd.Parameters("Vehicle_Type").Value = Null
End If

'搬運工具
Set tmp_para = tmp_Cmd.CreateParameter("PICK_TOOL", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab0_PickTool.ListIndex <> -1 Then
   tmp_Cmd.Parameters("PICK_TOOL").Value = arPickTool(cmb_Tab0_PickTool.ListIndex)
Else
   tmp_Cmd.Parameters("PICK_TOOL").Value = Null
End If

'通路別
Set tmp_para = tmp_Cmd.CreateParameter("Channel", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab0_Channel.Text) > 0 Then
   tmp_Cmd.Parameters("Channel").Value = Trim(txt_Tab0_Channel.Text)
Else
   tmp_Cmd.Parameters("Channel").Value = Null
End If

'傳真
Set tmp_para = tmp_Cmd.CreateParameter("fax", adChar, adParamInput, 18)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("fax").Value = Trim(txt_Tab0_Fax.Text)

'Codedate1
Set tmp_para = tmp_Cmd.CreateParameter("codedate1", adChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("codedate1").Value = Trim(txt_Tab0_CodeDate1.Text)

'Codedate2
Set tmp_para = tmp_Cmd.CreateParameter("codedate2", adChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("codedate2").Value = Trim(txt_Tab0_CodeDate2.Text)

'貼標
Set tmp_para = tmp_Cmd.CreateParameter("stamp", adChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("stamp").Value = Trim(txt_Tab0_Stamp.Text)

'罰款客戶
Set tmp_para = tmp_Cmd.CreateParameter("Penalties", adChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Penalties").Value = Trim(txt_Tab0_Penalties.Text)

'棧板材質
Set tmp_para = tmp_Cmd.CreateParameter("PalletType", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("PalletType").Value = Trim(txt_Tab0_PalletType.Text)

'棧板規格
Set tmp_para = tmp_Cmd.CreateParameter("Palletspec", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Palletspec").Value = Trim(txt_Tab0_PalletSpec.Text)

'備註
Set tmp_para = tmp_Cmd.CreateParameter("Notes", adChar, adParamInput, 255)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Notes").Value = Trim(txt_Tab0_Notes.Text)

'客戶群組
Set tmp_para = tmp_Cmd.CreateParameter("CustGroup", adChar, adParamInput, 255)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("CustGroup").Value = Trim(cmb_Tab0_Group.Text)

'允收期
Set tmp_para = tmp_Cmd.CreateParameter("CodeDateRate", adChar, adParamInput, 255)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("CodeDateRate").Value = Trim(cmdCodeDateRate)


'Codedate3
Set tmp_para = tmp_Cmd.CreateParameter("codedate3", adChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("codedate3").Value = Trim(txt_Tab0_CodeDate3.Text)


Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'非同步執行
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   DoEvents: DoEvents  '讓 [執行中] 訊息視窗有 [更新] 時間
Loop
Set tmp_Cmd = Nothing

fam_Tab0_Consignee.BackColor = &H8000000C
fam_Tab0_Consignee.Enabled = False
cmd_Tab0_Cancel.Enabled = False
cmd_Tab0_Save.Enabled = False
cmd_Tab0_AddNew.Enabled = True
cmd_Tab0_Modify.Enabled = True
cmd_Tab0_Delete.Enabled = False

'紀錄EditWho
str_SQL = "update trp01m set editwho = '" & User_id & "' , editdate = getdate() where storerkey = '" & mySplit(cmb_Tab0_Storer, " ", 0) & "' and consigneekey = '" & txt_Tab0_ConsigneeKey & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'紀錄ADDWho
str_SQL = "update trp01m set addwho = '" & User_id & "' where addwho is null and storerkey = '" & mySplit(cmb_Tab0_Storer, " ", 0) & "' and consigneekey = '" & txt_Tab0_ConsigneeKey & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'LTKK01客戶主檔異動自動 Mail 通知
If mySplit(cmb_Tab0_Storer, " ", 0) = "LTKK01" Then Call SendMail(txt_Tab0_ConsigneeKey)

If rs_Tab0_ConsigneeList Is Nothing = False Then rs_Tab0_ConsigneeList("客戶名稱") = Trim(txt_Tab0_FullName.Text)

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-客戶資料-存檔", Me.Caption, "cmd_Tab0_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_AddNew_Click()
'車輛資料 >> 新增模式轉換
If Not rs_Tab1_CarList Is Nothing Then
   If dg_Tab1_CarList.SelBookmarks.Count > 0 Then dg_Tab1_CarList.SelBookmarks.Remove 0
End If
fam_Tab1_Car.BackColor = &HC0FFC0
fam_Tab1_Car.Enabled = True
txt_Tab1_CarID.Enabled = True
Call Clear_CarData
cmd_Tab1_Save.Enabled = True
cmd_Tab1_Cancel.Enabled = True
cmd_Tab1_AddNew.Enabled = False
cmd_Tab1_Modify.Enabled = False
cmd_Tab1_Delete.Enabled = False
End Sub

Private Sub cmd_Tab1_Cancel_Click()
'車輛資料 >> 取消
Call Clear_CarData
If txt_Tab1_CarID.Enabled = False Then
   If Not rs_Tab1_CarList Is Nothing Then
      dg_Tab1_CarList.SelBookmarks.Add rs_Tab1_CarList.Bookmark
      Call Display_SelectedCarData(rs_Tab1_CarList.Fields("車牌號碼").Value)
   End If
End If
fam_Tab1_Car.BackColor = &H8000000C
chkPND.BackColor = &H8000000C
fam_Tab1_Car.Enabled = False
cmd_Tab1_Cancel.Enabled = False
cmd_Tab1_Save.Enabled = False
cmd_Tab1_AddNew.Enabled = True
cmd_Tab1_Modify.Enabled = True
cmd_Tab1_Delete.Enabled = True
End Sub

Private Sub cmd_Tab1_CarQuery_Click()
'車輛資料 >> 車輛搜尋
If rs_Tab1_CarList Is Nothing Then Exit Sub
If rs_Tab1_CarList.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_Tab1_CarList"

If ShowForm_RS_FilterAndSort(rs_Tab1_CarList, "車輛資料", Me.Tag) = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Me.WindowState = vbNormal

End Sub

Private Sub cmd_Tab1_CarReset_Click()
'車輛基本資料 >> 取消篩選排序
'移除篩選條件，重設排序依據
If rs_Tab1_CarList Is Nothing Then Exit Sub
 blTab1CarEventEnable = False
 rs_Tab1_CarList.Filter = adFilterNone
 rs_Tab1_CarList.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
 blTab1CarEventEnable = True
End Sub

Private Sub cmd_Tab1_CarShow_Click()

'車輛資料 >> 顯示所有車輛
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab1_CarList.DataSource = Nothing
Call ReDim_Recordset(rs_Tab1_CarList)
Call Confirm_Recordset_Closed(tmp_Rs)

str_SQL = "Select Rtrim(a1.Vehicle_ID_No) as 車牌號碼 , Rtrim(Isnull(a1.Driver,'')) as 駕駛人 , Rtrim(Isnull(b1.Description,'')) as 車種 , Rtrim(Isnull(a1.receiver,'')) as 請款人   " & _
          "From TRP09M a1 Left outer join TRP15M b1 on b1.Vehicle_Type = a1.Vehicle_Type Order by A1.Vehicle_ID_No"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列車輛資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Call Replication_Recordset(tmp_Rs, rs_Tab1_CarList)
tmp_Rs.Close

blTab1CarEventEnable = False
With dg_Tab1_CarList
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With

rs_Tab1_CarList.MoveFirst
Set dg_Tab1_CarList.DataSource = rs_Tab1_CarList
With dg_Tab1_CarList
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 900        '車牌號碼
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       '司機
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 3000       '車種
    .Columns(3).Alignment = dbgLeft
End With

blTab1CarEventEnable = True
Call Clear_CarData
Screen.MousePointer = vbDefault
Call Display_SelectedCarData(rs_Tab1_CarList.Fields("車牌號碼").Value)
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-車輛資料-顯示所有資料", Me.Caption, "cmd_Tab1-CarShow_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_AddNew_Click()
'貨運公司資料 >> 新增
If Not rs_Tab2_TRPCompanyList Is Nothing Then
   If dg_Tab2_TRPCompanyList.SelBookmarks.Count > 0 Then dg_Tab2_TRPCompanyList.SelBookmarks.Remove 0
End If
fam_Tab2_Company.BackColor = &HC0FFC0
fam_Tab2_Company.Enabled = True
txt_Tab2_CompanyCode.Enabled = True
Call Clear_CompanyData
cmd_Tab2_Save.Enabled = True
cmd_Tab2_Cancel.Enabled = True
cmd_Tab2_AddNew.Enabled = False
cmd_Tab2_Modify.Enabled = False
cmd_Tab2_Delete.Enabled = False
End Sub

Private Sub cmd_Tab2_Cancel_Click()
'貨運公司資料 >> 取消
Call Clear_CompanyData
If txt_Tab2_CompanyCode.Enabled = False Then
   If Not rs_Tab2_TRPCompanyList Is Nothing Then
      dg_Tab2_TRPCompanyList.SelBookmarks.Add rs_Tab2_TRPCompanyList.Bookmark
      Call Display_SelectedCompanyData(rs_Tab2_TRPCompanyList.Fields("公司代碼").Value)
   End If
End If
fam_Tab2_Company.BackColor = &H8000000C
fam_Tab2_Company.Enabled = False
cmd_Tab2_Cancel.Enabled = False
cmd_Tab2_Save.Enabled = False
cmd_Tab2_AddNew.Enabled = True
cmd_Tab2_Modify.Enabled = True
cmd_Tab2_Delete.Enabled = True
End Sub

Private Sub cmd_Tab2_CarQuery_Click()
'貨運公司資料 >> 車輛搜尋
If rs_Tab2_TRPCompanyList Is Nothing Then Exit Sub
If rs_Tab2_TRPCompanyList.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_Tab2_TRPCompanyList"

If ShowForm_RS_FilterAndSort(rs_Tab2_TRPCompanyList, "貨運公司資料", Me.Tag) = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Me.WindowState = vbNormal
End Sub

Private Sub cmd_Tab2_CompanyShow_Click()
'貨運公司資料 >> 顯示所有貨運公司
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab2_TRPCompanyList.DataSource = Nothing
Call ReDim_Recordset(rs_Tab2_TRPCompanyList)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "Select Rtrim(Company_Code) as 公司代碼,Rtrim(Isnull(C_Name,'')) as 中文名稱,Rtrim(Isnull(E_Name,'')) as 英文名稱,Rtrim(Isnull(Short_Name,'')) as 簡稱," & _
          "   Rtrim(isnull(Phone,'')) as 電話 , Rtrim(Isnull(Contact,'')) as 聯絡人 ,Rtrim(Isnull(Address,'')) as 地址 , Rtrim(Isnull(Description,'')) as Descr  " & _
          "From TRP08M Order by Company_Code"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列貨運公司資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab2_TRPCompanyList)
tmp_Rs.Close

blTab2CompanyEventEnable = False
With dg_Tab2_TRPCompanyList
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab2_TRPCompanyList.MoveFirst
Set dg_Tab2_TRPCompanyList.DataSource = rs_Tab2_TRPCompanyList
With dg_Tab2_TRPCompanyList
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 900        '貨運公司代碼
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 2500       '中文名稱
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000       '英文名稱
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1500       '簡稱
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 1100       '電話
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 800        '聯絡人
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 2500       '地址
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 2000       '說明
    .Columns(8).Alignment = dbgLeft
End With
blTab2CompanyEventEnable = True
Call Clear_CompanyData
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-貨運公司資料-顯示所有資料", Me.Caption, "cmd_Tab2-CompanyShow_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Delete_Click()
'車輛資料 >> 刪除
Dim blDelete As Boolean
blDelete = True
msg_text = ""

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus
Screen.MousePointer = vbHourglass

'1.檢核 TRP05T 是否有此車輛裝載出車資料
str_SQL = "Select Count(*) as RecCnt From TRP05T Where Vehicle_ID_No = '" & Trim(txt_Tab1_CarID.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   路線編號運送資料 [TRP05T] 有此車輛裝載出車資料"
   Else
      msg_text = msg_text & vbCrLf & "   路線編號運送資料 [TRP05T] 有此車輛裝載出車資料"
   End If
End If
tmp_Rs.Close

'2.檢核 TRP02T 是否有此車輛裝載出車資料
str_SQL = "Select Count(*) as RecCnt From TRP02T Where Vehicle_ID_No = '" & Trim(txt_Tab1_CarID.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   已排車訂單 [TRP02T] 有此車輛裝載出車資料"
   Else
      msg_text = msg_text & vbCrLf & "   已排車訂單 [TRP02T] 有此車輛裝載出車資料"
   End If
End If
tmp_Rs.Close

'3.檢核 SDN02T 是否有此車輛裝載出車資料
str_SQL = "Select Count(*) as RecCnt From SDN02T Where Vehicle_ID_No = '" & Trim(txt_Tab1_CarID.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   已出車訂單 [SDN02T] 有此車輛裝載出車資料"
   Else
      msg_text = msg_text & vbCrLf & "   已排車訂單 [SDN02T] 有此車輛裝載出車資料"
   End If
End If
tmp_Rs.Close

'4.檢核 SDN01T 是否有此車輛裝載出車資料
str_SQL = "Select Count(*) as RecCnt From SDN01T Where C_Vehicle_ID_No = '" & Trim(txt_Tab1_CarID.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   已出車訂單 [SDN01T] 有此車輛裝載出車資料"
   Else
      msg_text = msg_text & vbCrLf & "   已排車訂單 [SDN01T] 有此車輛裝載出車資料"
   End If
End If
tmp_Rs.Close

'5.檢核 ORT02T 是否有此車輛裝載出車資料
str_SQL = "Select Count(*) as RecCnt From ORT02T Where Vehicle_ID_No = '" & Trim(txt_Tab1_CarID.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   已出車訂單 [ORT02T] 有此車輛裝載出車資料"
   Else
      msg_text = msg_text & vbCrLf & "   已排車訂單 [ORT02T] 有此車輛裝載出車資料"
   End If
End If
tmp_Rs.Close

'檢核是否允許進行刪除旗標值
If blDelete = False Then
   msg_text = "車輛資料無法刪除：" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

'允許刪除
str_SQL = "Delete From TRP09M Where Vehicle_ID_No = '" & Trim(txt_Tab1_CarID.Text) & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

fam_Tab1_Car.BackColor = &H8000000C
fam_Tab1_Car.Enabled = False
cmd_Tab1_Cancel.Enabled = False
cmd_Tab1_Save.Enabled = False
cmd_Tab1_AddNew.Enabled = True
cmd_Tab1_Modify.Enabled = False
cmd_Tab1_Delete.Enabled = False
'重新顯示所有客戶資料
Call cmd_Tab1_CarShow_Click

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-車輛資料-刪除", Me.Caption, "cmd_Tab1_Delete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Modify_Click()
'車輛資料 >> 修改
'確認選取車輛資料方允許 [修改] 功能
If rs_Tab1_CarList Is Nothing Then Exit Sub
'If txt_Tab1_Driver = "" Then Exit Sub mark by Gemini @20151019
If dg_Tab1_CarList.SelBookmarks.Count <> 0 Then
   fam_Tab1_Car.BackColor = &HC0E0FF
   chkPND.BackColor = &HC0E0FF
   fam_Tab1_Car.Enabled = True
   txt_Tab1_CarID.Enabled = False
   cmd_Tab1_Save.Enabled = True
   cmd_Tab1_Cancel.Enabled = True
   cmd_Tab1_AddNew.Enabled = False
   cmd_Tab1_Modify.Enabled = False
   cmd_Tab1_Delete.Enabled = False
End If

End Sub

Private Sub cmd_Tab1_Save_Click()
'車輛資料 >> 存檔

'清除特殊字元
Call myFormExCharFilter(Me)

On Error GoTo err_Handle

'存檔資料檢核
If Check_CarData = False Then Exit Sub

Screen.MousePointer = vbHourglass
If Not (tmp_Cmd Is Nothing) Then
   Set tmp_Cmd = Nothing
End If
Set tmp_Cmd = New ADODB.Command
If tmp_para Is Nothing Then
   Set tmp_para = New ADODB.Parameter
End If

tmp_Cmd.ActiveConnection = cn
tmp_Cmd.CommandTimeout = 0    '執行時間設定：無限期等待
tmp_Cmd.CommandType = adCmdStoredProc
tmp_Cmd.CommandText = "Master_CarDara_Update"

'車牌號碼
Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_ID_NO", adChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("VEHICLE_ID_NO").Value = Trim(txt_Tab1_CarID.Text)

'郵遞區號
Set tmp_para = tmp_Cmd.CreateParameter("ZIP", adVarChar, adParamInput, 18)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_ZIP.ListIndex <> -1 Then
   tmp_Cmd.Parameters("ZIP").Value = arZip(cmb_Tab1_ZIP.ListIndex)
Else
   tmp_Cmd.Parameters("ZIP").Value = Null
End If

'運送區碼
Set tmp_para = tmp_Cmd.CreateParameter("Area_Code", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_AreaCode.ListIndex <> -1 Then
   tmp_Cmd.Parameters("Area_Code").Value = arAreaCode(cmb_Tab1_AreaCode.ListIndex)
Else
   tmp_Cmd.Parameters("Area_Code").Value = Null
End If

'貨運公司
Set tmp_para = tmp_Cmd.CreateParameter("TRP_COMPANY_CODE", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_Company.ListIndex <> -1 Then
   tmp_Cmd.Parameters("TRP_COMPANY_CODE").Value = arCompany(cmb_Tab1_Company.ListIndex)
Else
   tmp_Cmd.Parameters("TRP_COMPANY_CODE").Value = Null
End If

'車種
Set tmp_para = tmp_Cmd.CreateParameter("VEHICLE_TYPE", adVarChar, adParamInput, 2)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_VehicleType.ListIndex <> -1 Then
   tmp_Cmd.Parameters("VEHICLE_TYPE").Value = arVehicleType(cmb_Tab1_VehicleType.ListIndex)
Else
   tmp_Cmd.Parameters("VEHICLE_TYPE").Value = Null
End If

'可承載重量
Set tmp_para = tmp_Cmd.CreateParameter("LOADING_SIZE", adDouble, adParamInput)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_WeightCapacity.Text) = "" Then
   tmp_Cmd.Parameters("LOADING_SIZE").Value = Null
Else
   tmp_Cmd.Parameters("LOADING_SIZE").Value = Trim(txt_Tab1_WeightCapacity.Text)
End If

'可承載材積
Set tmp_para = tmp_Cmd.CreateParameter("MAX_CUBIC_CAPACITY", adDouble, adParamInput)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_VolumnCapacity.Text) = "" Then
   tmp_Cmd.Parameters("MAX_CUBIC_CAPACITY").Value = Null
Else
   tmp_Cmd.Parameters("MAX_CUBIC_CAPACITY").Value = Trim(txt_Tab1_VolumnCapacity.Text)
End If

'司機
Set tmp_para = tmp_Cmd.CreateParameter("DRIVER", adVarChar, adParamInput, 30)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_Driver.Text) = "" Then
   tmp_Cmd.Parameters("DRIVER").Value = Null
Else
   tmp_Cmd.Parameters("DRIVER").Value = Trim(txt_Tab1_Driver.Text)
End If

'電話
Set tmp_para = tmp_Cmd.CreateParameter("DRIVER_PHONE", adVarChar, adParamInput, 18)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_Phone.Text) = "" Then
   tmp_Cmd.Parameters("DRIVER_PHONE").Value = Null
Else
   tmp_Cmd.Parameters("DRIVER_PHONE").Value = Trim(txt_Tab1_Phone.Text)
End If

'說明
Set tmp_para = tmp_Cmd.CreateParameter("Description", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_Description.Text) = "" Then
   tmp_Cmd.Parameters("Description").Value = Null
Else
   tmp_Cmd.Parameters("Description").Value = Trim(txt_Tab1_Description.Text)
End If

'可裝載板數
Set tmp_para = tmp_Cmd.CreateParameter("PALLET_CAPACITY", adDouble, adParamInput)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_PalletCapacity.Text) = "" Then
   tmp_Cmd.Parameters("PALLET_CAPACITY").Value = Null
Else
   tmp_Cmd.Parameters("PALLET_CAPACITY").Value = Trim(txt_Tab1_PalletCapacity.Text)
End If

'車重
Set tmp_para = tmp_Cmd.CreateParameter("CAR_WIEGHT", adDouble, adParamInput)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_CarWeight.Text) = "" Then
   tmp_Cmd.Parameters("CAR_WIEGHT").Value = "0"
Else
   tmp_Cmd.Parameters("CAR_WIEGHT").Value = Trim(txt_Tab1_CarWeight.Text)
End If

'車廂形式
Set tmp_para = tmp_Cmd.CreateParameter("CARBOX_TYPE", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_CarBox.ListIndex <> -1 Then
   tmp_Cmd.Parameters("CARBOX_TYPE").Value = arCarBox(cmb_Tab1_CarBox.ListIndex)
Else
   tmp_Cmd.Parameters("CARBOX_TYPE").Value = Null
End If

'車床高度
Set tmp_para = tmp_Cmd.CreateParameter("CAR_HEIGHT", adDouble, adParamInput)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab1_CarHeight.Text) = "" Then
   tmp_Cmd.Parameters("CAR_HEIGHT").Value = Null   'Trim(txt_Tab1_CarHeight.Text)
Else
   tmp_Cmd.Parameters("CAR_HEIGHT").Value = Trim(txt_Tab1_CarHeight.Text)
End If

'裝卸方式
Set tmp_para = tmp_Cmd.CreateParameter("UNLAODING_TYPE", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_UnloadType.ListIndex <> -1 Then
   tmp_Cmd.Parameters("UNLAODING_TYPE").Value = arUnloadType(cmb_Tab1_UnloadType.ListIndex)
Else
   tmp_Cmd.Parameters("UNLAODING_TYPE").Value = Null
End If

'僱用方式
Set tmp_para = tmp_Cmd.CreateParameter("EMPLOY_TYPE", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If cmb_Tab1_EmployType.ListIndex <> -1 Then
   tmp_Cmd.Parameters("EMPLOY_TYPE").Value = arEmployType(cmb_Tab1_EmployType.ListIndex)
Else
   tmp_Cmd.Parameters("EMPLOY_TYPE").Value = Null
End If

'計費類別
Set tmp_para = tmp_Cmd.CreateParameter("CAR_TYPE", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(cmb_Tab1_CarType.Text)) > 0 Then
   tmp_Cmd.Parameters("CAR_TYPE").Value = cmb_Tab1_CarType.Text
Else
   tmp_Cmd.Parameters("CAR_TYPE").Value = Null
End If

'請款人
Set tmp_para = tmp_Cmd.CreateParameter("Receiver", adVarChar, adParamInput, 50)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(txt_Tab1_Receiver)) > 0 Then
   tmp_Cmd.Parameters("Receiver").Value = txt_Tab1_Receiver
Else
   tmp_Cmd.Parameters("Receiver").Value = Null
End If

'PND
Set tmp_para = tmp_Cmd.CreateParameter("PND", adVarChar, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
If chkPND.Value = vbChecked Then
   tmp_Cmd.Parameters("PND").Value = "Y"
Else
   tmp_Cmd.Parameters("PND").Value = "N"
End If

'APFix
Set tmp_para = tmp_Cmd.CreateParameter("APFix", adDouble, adParamInput, 1)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("APFix").Value = txtAPFix

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'非同步執行
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   DoEvents: DoEvents  '讓 [執行中] 訊息視窗有 [更新] 時間
Loop

'紀錄EditWho
str_SQL = "update trp09m set editwho = '" & User_id & "' , editdate = getdate() where VEHICLE_ID_NO = '" & txt_Tab1_CarID & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'紀錄ADDWho
str_SQL = "update trp09m set addwho = '" & User_id & "' where addwho is null and VEHICLE_ID_NO = '" & txt_Tab1_CarID & "' "
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'異動紀錄
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "Select * From TRP09M Where VEHICLE_ID_NO = '" & txt_Tab1_CarID & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If Not tmp_Rs.EOF Then
    Dim str As String, i As Integer
    For i = 0 To tmp_Rs.Fields.Count - 1
        str = str & RTrim(tmp_Rs.Fields(i)) & ","
    Next i
    
    '寫入資料庫紀錄
    str_SQL = "Insert into gt_Logs(APName,APVer,APCaption,Code,Description,Notes,ComputerName,AddWho) Values ('" & _
                    App.EXEName & "','" & App.Major & "." & App.Minor & "." & App.Revision & "','" & Me.Caption & "','0','車輛主檔異動紀錄','" & str & "','" & strComputerName & "','" & User_id & "')"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
End If

fam_Tab1_Car.BackColor = &H8000000C
chkPND.BackColor = &H8000000C
fam_Tab1_Car.Enabled = False
cmd_Tab1_Cancel.Enabled = False
cmd_Tab1_Save.Enabled = False
cmd_Tab1_AddNew.Enabled = True
cmd_Tab1_Modify.Enabled = False
cmd_Tab1_Delete.Enabled = False

'重新顯示所有客戶資料
'Call cmd_Tab1_CarShow_Click'改為不重新查詢，避免USER從頭查找
If Not rs_Tab1_CarList Is Nothing Then
    If rs_Tab1_CarList("車牌號碼") = txt_Tab1_CarID Then '非新增時更新清單資料
        rs_Tab1_CarList("駕駛人") = txt_Tab1_Driver
        rs_Tab1_CarList("車種") = cmb_Tab1_VehicleType
        Call Display_SelectedCarData(rs_Tab1_CarList.Fields("車牌號碼").Value)
    End If
End If
Screen.MousePointer = vbDefault

Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-車輛資料-存檔", Me.Caption, "cmd_Tab1_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Delete_Click()
'貨運公司 >> 刪除
Dim blDelete As Boolean
blDelete = True
msg_text = ""

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus
Screen.MousePointer = vbHourglass

'1.檢核 TRP05T 是否有此貨運公司裝載出車資料
str_SQL = "Select Count(*) as RecCnt From TRP05T Where TRP_Company_Code = '" & Trim(txt_Tab2_CompanyCode.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   路線編號運送資料 [TRP05T] 有此貨運公司裝載出車資料"
   Else
      msg_text = msg_text & vbCrLf & "   路線編號運送資料 [TRP05T] 有此貨運公司裝載出車資料"
   End If
End If
tmp_Rs.Close

'2.檢核 ORT05T 是否有此貨運公司裝載出車資料
str_SQL = "Select Count(*) as RecCnt From ORT05T Where TRP_Company_Code = '" & Trim(txt_Tab2_CompanyCode.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   路線編號運送資料 [ORT05T] 有此貨運公司裝載出車資料"
   Else
      msg_text = msg_text & vbCrLf & "   路線編號運送資料 [TRP05T] 有此貨運公司裝載出車資料"
   End If
End If
tmp_Rs.Close

'3.檢核 TRP09M 是否有此貨運公司裝載出車資料
str_SQL = "Select Count(*) as RecCnt From TRP09M Where TRP_Company_Code = '" & Trim(txt_Tab2_CompanyCode.Text) & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.Fields("RecCnt").Value > 0 Then
   blDelete = False
   If msg_text = "" Then
      msg_text = "   車輛基本資料檔 [TRP09M] 有此貨運公司車輛資料"
   Else
      msg_text = msg_text & vbCrLf & "   車輛基本資料檔 [TRP09M] 有此貨運公司車輛資料"
   End If
End If
tmp_Rs.Close

'檢核是否允許進行刪除旗標值
If blDelete = False Then
   msg_text = "貨運公司資料無法刪除：" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

'允許刪除
str_SQL = "Delete From TRP08M Where Company_Code = '" & Trim(txt_Tab2_CompanyCode.Text) & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

fam_Tab2_Company.BackColor = &H8000000C
fam_Tab2_Company.Enabled = False
cmd_Tab2_Cancel.Enabled = False
cmd_Tab2_Save.Enabled = False
cmd_Tab2_AddNew.Enabled = False
cmd_Tab2_Modify.Enabled = False
cmd_Tab2_Delete.Enabled = False
'重新顯示所有客戶資料
Call cmd_Tab1_CarShow_Click

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-貨運公司資料-刪除", Me.Caption, "cmd_Tab2_Delete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_Modify_Click()
'貨運公司資料 >> 修改
'確認選取車輛資料方允許 [修改] 功能
If rs_Tab2_TRPCompanyList Is Nothing Then Exit Sub
If dg_Tab2_TRPCompanyList.SelBookmarks.Count <> 0 Then
   fam_Tab2_Company.BackColor = &HC0E0FF
   fam_Tab2_Company.Enabled = True
   txt_Tab2_CompanyCode.Enabled = False
   cmd_Tab2_Save.Enabled = True
   cmd_Tab2_Cancel.Enabled = True
   cmd_Tab2_AddNew.Enabled = False
   cmd_Tab2_Modify.Enabled = False
   cmd_Tab2_Delete.Enabled = False
End If

End Sub

Private Sub cmd_Tab2_Save_Click()

'清除特殊字元
Call myFormExCharFilter(Me)

'貨運公司資料 >> 存檔
On Error GoTo err_Handle

'存檔資料檢核
If Check_CompanyData = False Then Exit Sub

Screen.MousePointer = vbHourglass
If Not (tmp_Cmd Is Nothing) Then
   Set tmp_Cmd = Nothing
End If
Set tmp_Cmd = New ADODB.Command
If tmp_para Is Nothing Then
   Set tmp_para = New ADODB.Parameter
End If

tmp_Cmd.ActiveConnection = cn
tmp_Cmd.CommandTimeout = 0    '執行時間設定：無限期等待
tmp_Cmd.CommandType = adCmdStoredProc
tmp_Cmd.CommandText = "Master_CompanyData_Update"
'公司代碼
Set tmp_para = tmp_Cmd.CreateParameter("Company_Code", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Company_Code").Value = Trim(txt_Tab2_CompanyCode.Text)
'中文名稱
Set tmp_para = tmp_Cmd.CreateParameter("C_Name", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("C_Name").Value = Trim(txt_Tab2_CName.Text)
'英文名稱
Set tmp_para = tmp_Cmd.CreateParameter("E_Name", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(txt_Tab2_EName.Text)) > 0 Then
   tmp_Cmd.Parameters("E_Name").Value = Trim(txt_Tab2_EName.Text)
Else
   tmp_Cmd.Parameters("E_Name").Value = Null
End If
'地址
Set tmp_para = tmp_Cmd.CreateParameter("Address", adVarChar, adParamInput, 45)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(txt_Tab2_Address.Text)) <> 0 Then
   tmp_Cmd.Parameters("Address").Value = Trim(txt_Tab2_Address.Text)
Else
   tmp_Cmd.Parameters("Address").Value = Null
End If
'簡稱
Set tmp_para = tmp_Cmd.CreateParameter("Short_Name", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(txt_Tab2_ShortName.Text)) <> 0 Then
   tmp_Cmd.Parameters("Short_Name").Value = Trim(txt_Tab2_ShortName.Text)
Else
   tmp_Cmd.Parameters("Short_Name").Value = Null
End If
'聯絡人 'Terry 20180123 contact 長度由30改為80
Set tmp_para = tmp_Cmd.CreateParameter("Contact", adVarChar, adParamInput, 80)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab2_Contact.Text) = "" Then
   tmp_Cmd.Parameters("Contact").Value = Null
Else
   tmp_Cmd.Parameters("Contact").Value = Trim(txt_Tab2_Contact.Text)
End If
'電話
Set tmp_para = tmp_Cmd.CreateParameter("Phone", adVarChar, adParamInput, 18)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Phone").Value = Null
If Trim(txt_Tab2_Contact.Text) = "" Then
   tmp_Cmd.Parameters("Phone").Value = Null
Else
   tmp_Cmd.Parameters("Phone").Value = Trim(txt_Tab2_Phone.Text)
End If
'說明
Set tmp_para = tmp_Cmd.CreateParameter("Description", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
If Trim(txt_Tab2_Descr.Text) = "" Then
   tmp_Cmd.Parameters("Description").Value = Null
Else
   tmp_Cmd.Parameters("Description").Value = Trim(txt_Tab2_Descr.Text)
End If

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'非同步執行
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   DoEvents: DoEvents  '讓 [執行中] 訊息視窗有 [更新] 時間
Loop

fam_Tab2_Company.BackColor = &H8000000C
fam_Tab2_Company.Enabled = False
cmd_Tab2_Cancel.Enabled = False
cmd_Tab2_Save.Enabled = False
cmd_Tab2_AddNew.Enabled = True
cmd_Tab2_Modify.Enabled = False
cmd_Tab2_Delete.Enabled = False
'重新顯示所有客戶資料
Call cmd_Tab1_CarShow_Click
Screen.MousePointer = vbDefault

Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-運輸公司資料-存檔", Me.Caption, "cmd_Tab2_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_SkuQuery_Click()
'貨號資料 >> 貨號資料搜尋
If rs_Tab3_SkuList Is Nothing Then Exit Sub
If rs_Tab3_SkuList.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_Tab3_SkuList"

If ShowForm_RS_FilterAndSort(rs_Tab3_SkuList, "貨號資料", Me.Tag) = False Then
    MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
    Exit Sub
End If
Me.WindowState = vbNormal
End Sub

Private Sub cmd_Tab2_SkuShow_Click()
'貨號資料 >> 顯示所有貨號
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab2_TRPCompanyList.DataSource = Nothing
Call ReDim_Recordset(rs_Tab2_TRPCompanyList)
Call Confirm_Recordset_Closed(tmp_Rs)

str_SQL = "select StorerKey as 貨主, Sku as 貨號, DESCR as 中文名稱,  STDGROSSWGT as 每箱重, busr4  as 每箱材, " & _
        "isnull(SKUGROUP,'') as 類別,rtrim(SUSR1) as 產品別,NOTES1 as 備註一,NOTES2 as 備註二  from gv_SKUxpack"

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
   msg_text = "資料錯誤：查詢結果傳回 0 列貨號資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab3_SkuList)
tmp_Rs.Close

blTab3skuEventEnable = False
With dg_Tab3_SkuList
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab3_SkuList.MoveFirst
Set dg_Tab3_SkuList.DataSource = rs_Tab3_SkuList
With dg_Tab3_SkuList
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 900        '貨主
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 900       '貨號
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 2500       '中文名稱
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 900       '每箱重
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 900       '每箱材
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 900        '類別
    .Columns(6).Alignment = dbgLeft
    .Columns(7).Width = 900       '產品別
    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 2000       '英文名稱
    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 2000       '說明
    .Columns(9).Alignment = dbgLeft
End With
blTab3skuEventEnable = True
'Call Clear_CompanyData
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-貨號資料-顯示所有資料", Me.Caption, "cmd_Tab3_SKUShow_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab2_TRPCompanyReset_Click()
'貨運公司基本資料 >> 取消篩選排序
'移除篩選條件，重設排序依據
If rs_Tab2_TRPCompanyList Is Nothing Then Exit Sub
 blTab2CompanyEventEnable = False
 rs_Tab2_TRPCompanyList.Filter = adFilterNone
 rs_Tab2_TRPCompanyList.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
 blTab2CompanyEventEnable = True

End Sub

Private Sub cmd_Tab3_AddNew_Click()
'貨運公司資料 >> 新增
If Not rs_Tab3_SkuList Is Nothing Then
   If dg_Tab3_SkuList.SelBookmarks.Count > 0 Then dg_Tab3_SkuList.SelBookmarks.Remove 0
End If
fam_Tab3_Sku.BackColor = &HC0FFC0
fam_Tab3_Sku.Enabled = True
txt_Tab3_Sku.Enabled = True
Call Clear_SkuData
cmd_Tab3_Save.Enabled = True
cmd_Tab3_Cancel.Enabled = True
cmd_Tab3_AddNew.Enabled = False
cmd_Tab3_Modify.Enabled = False
cmd_Tab3_Delete.Enabled = False
End Sub

Private Sub cmd_Tab3_Cancel_Click()
'貨號資料 >> 取消
Call Clear_SkuData
If txt_Tab3_Sku.Enabled = False Then
   If Not rs_Tab3_SkuList Is Nothing Then
      dg_Tab3_SkuList.SelBookmarks.Add rs_Tab3_SkuList.Bookmark
      Call Display_SelectedSkuData(rs_Tab3_SkuList.Fields("貨號").Value)
   End If
End If
fam_Tab3_Sku.BackColor = &H8000000C
fam_Tab3_Sku.Enabled = False
cmd_Tab3_Cancel.Enabled = False
cmd_Tab3_Save.Enabled = False
cmd_Tab3_AddNew.Enabled = True
cmd_Tab3_Modify.Enabled = True
cmd_Tab3_Delete.Enabled = True
End Sub

Private Sub cmd_Tab3_Modify_Click()
'貨號公司資料 >> 修改
'確認選取車輛資料方允許 [修改] 功能
If rs_Tab3_SkuList Is Nothing Then Exit Sub
If dg_Tab3_SkuList.SelBookmarks.Count <> 0 Then
   fam_Tab3_Sku.BackColor = &HC0E0FF
   fam_Tab3_Sku.Enabled = True
   txt_Tab3_Sku.Enabled = False
   cmd_Tab3_Save.Enabled = True
   cmd_Tab3_Cancel.Enabled = True
   cmd_Tab3_AddNew.Enabled = False
   cmd_Tab3_Modify.Enabled = False
   cmd_Tab3_Delete.Enabled = False
End If
End Sub

Private Sub cmd_Tab3_Save_Click()
'貨運公司資料 >> 存檔

'清除特殊字元
Call myFormExCharFilter(Me)

On Error GoTo err_Handle
'select StorerKey as 貨主, Sku as 貨號, DESCR as 中文名稱, STDGROSSWGT as 每箱重,rtrim(BUSR4) as 每箱材, " & _
        "SKUGROUP as 類別,rtrim(BUSR1) as 產品別,NOTES1 as 英文名稱,NOTES2 as 說明  from dbo.SKU"
'存檔資料檢核
If Check_SkuData = False Then Exit Sub

Screen.MousePointer = vbHourglass
If Not (tmp_Cmd Is Nothing) Then
   Set tmp_Cmd = Nothing
End If
Set tmp_Cmd = New ADODB.Command
If tmp_para Is Nothing Then
   Set tmp_para = New ADODB.Parameter
End If

tmp_Cmd.ActiveConnection = cn
tmp_Cmd.CommandTimeout = 0    '執行時間設定：無限期等待
tmp_Cmd.CommandType = adCmdStoredProc
tmp_Cmd.CommandText = "Master_SkuData_Update"
'貨主
Set tmp_para = tmp_Cmd.CreateParameter("StorerKey", adChar, adParamInput, 15)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("StorerKey").Value = Trim(txt_Tab3_StorerKey.Text)
'貨號
Set tmp_para = tmp_Cmd.CreateParameter("Sku", adVarChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("Sku").Value = Trim(txt_Tab3_Sku.Text)
'中文名稱
Set tmp_para = tmp_Cmd.CreateParameter("DESCR", adVarChar, adParamInput, 60)
tmp_Cmd.Parameters.Append tmp_para
'If Len(Trim(txt_Tab3_DESCR.Text)) > 0 Then
   tmp_Cmd.Parameters("DESCR").Value = Trim(txt_Tab3_DESCR.Text)
'Else
'   tmp_cmd.Parameters("DESCR").Value = Null
'End If
'每箱重
Set tmp_para = tmp_Cmd.CreateParameter("STDGROSSWGT", adVarChar, adParamInput, 30)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(txt_Tab3_STDGROSSWGT.Text)) <> 0 Then
   tmp_Cmd.Parameters("STDGROSSWGT").Value = Trim(txt_Tab3_STDGROSSWGT.Text)
Else
   tmp_Cmd.Parameters("STDGROSSWGT").Value = 0
End If
'每箱材
Set tmp_para = tmp_Cmd.CreateParameter("BUSR4", adVarChar, adParamInput, 30)
tmp_Cmd.Parameters.Append tmp_para
If Len(Trim(txt_Tab3_BUSR4.Text)) <> 0 Then
   tmp_Cmd.Parameters("BUSR4").Value = Trim(txt_Tab3_BUSR4.Text)
Else
   tmp_Cmd.Parameters("BUSR4").Value = 0
End If
'類別
Set tmp_para = tmp_Cmd.CreateParameter("SKUGROUP", adVarChar, adParamInput, 10)
tmp_Cmd.Parameters.Append tmp_para
'If Trim(txt_Tab3_SKUGROUP.Text) = "" Then
'   tmp_cmd.Parameters("SKUGROUP").Value = Null
'Else
   tmp_Cmd.Parameters("SKUGROUP").Value = Trim(txt_Tab3_SKUGROUP.Text)
'End If
'產品別
Set tmp_para = tmp_Cmd.CreateParameter("BUSR1", adVarChar, adParamInput, 30)
tmp_Cmd.Parameters.Append tmp_para
'If Trim(txt_Tab3_BUSR1.Text) = "" Then
'   tmp_cmd.Parameters("BUSR1").Value = Null
'Else
   tmp_Cmd.Parameters("BUSR1").Value = Trim(txt_Tab3_BUSR1.Text)
'End If
'英文名稱
Set tmp_para = tmp_Cmd.CreateParameter("NOTES1", adVarChar, adParamInput, 40)
tmp_Cmd.Parameters.Append tmp_para
'If Trim(txt_Tab3_NOTES1.Text) = "" Then
'   tmp_cmd.Parameters("NOTES1").Value = Null
'Else
   tmp_Cmd.Parameters("NOTES1").Value = Trim(txt_Tab3_NOTES1.Text)
'End If
'說明
Set tmp_para = tmp_Cmd.CreateParameter("NOTES2", adVarChar, adParamInput, 40)
tmp_Cmd.Parameters.Append tmp_para
'If Trim(txt_Tab3_NOTES2.Text) = "" Then
'   tmp_cmd.Parameters("NOTES2").Value = Null
'Else
   tmp_Cmd.Parameters("NOTES2").Value = Trim(txt_Tab3_NOTES2.Text)
'End If
Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'非同步執行
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   DoEvents: DoEvents  '讓 [執行中] 訊息視窗有 [更新] 時間
Loop

fam_Tab3_Sku.BackColor = &H8000000C
fam_Tab3_Sku.Enabled = False
cmd_Tab3_Cancel.Enabled = False
cmd_Tab3_Save.Enabled = False
cmd_Tab3_AddNew.Enabled = True
cmd_Tab3_Modify.Enabled = False
cmd_Tab3_Delete.Enabled = False
'重新顯示所有客戶資料
Call cmd_Tab2_SkuShow_Click
Screen.MousePointer = vbDefault

Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "貨號資料-存檔", Me.Caption, "cmd_Tab3_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Public Sub dg_Tab0_ConsigneeList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'客戶資料列表：整行選取
If rs_Tab0_ConsigneeList Is Nothing Then Exit Sub
If blTab0ConsignEventEnable Then
   If Not rs_Tab0_ConsigneeList.EOF Then
      dg_Tab0_ConsigneeList.SelBookmarks.Add rs_Tab0_ConsigneeList.Bookmark
      Call Display_SelectedConsignData(rs_Tab0_ConsigneeList.Fields("貨主").Value, rs_Tab0_ConsigneeList.Fields("客戶編號").Value)
      fam_Tab0_Consignee.BackColor = &H8000000C
      fam_Tab0_Consignee.Enabled = False
      cmd_Tab0_Cancel.Enabled = False
      cmd_Tab0_Save.Enabled = False
      cmd_Tab0_AddNew.Enabled = True
      cmd_Tab0_Modify.Enabled = True
      cmd_Tab0_Delete.Enabled = True
   End If
End If
End Sub

Private Sub dg_Tab1_CarList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'車輛資料列表：整行選取
If blTab1CarEventEnable Then
   If Not rs_Tab1_CarList.EOF Then
      dg_Tab1_CarList.SelBookmarks.Add rs_Tab1_CarList.Bookmark
      Call Display_SelectedCarData(rs_Tab1_CarList.Fields("車牌號碼").Value)
      fam_Tab1_Car.BackColor = &H8000000C
      chkPND.BackColor = &H8000000C
      fam_Tab1_Car.Enabled = False
      cmd_Tab1_Cancel.Enabled = False
      cmd_Tab1_Save.Enabled = False
      cmd_Tab1_AddNew.Enabled = True
      cmd_Tab1_Modify.Enabled = True
      cmd_Tab1_Delete.Enabled = True
   End If
End If
End Sub

Private Sub dg_Tab2_TRPCompanyList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'車輛資料列表：整行選取
If blTab2CompanyEventEnable Then
   If Not rs_Tab2_TRPCompanyList.EOF Then
      dg_Tab2_TRPCompanyList.SelBookmarks.Add rs_Tab2_TRPCompanyList.Bookmark
      Call Display_SelectedCompanyData(rs_Tab2_TRPCompanyList.Fields("公司代碼").Value)
      fam_Tab2_Company.BackColor = &H8000000C
      fam_Tab2_Company.Enabled = False
      cmd_Tab2_Cancel.Enabled = False
      cmd_Tab2_Save.Enabled = False
      cmd_Tab2_AddNew.Enabled = True
      cmd_Tab2_Modify.Enabled = True
      cmd_Tab2_Delete.Enabled = True
   End If
End If
End Sub

Private Sub dg_Tab3_SkuList_Click()
'車輛資料列表：整行選取
If blTab3skuEventEnable Then
   If Not rs_Tab3_SkuList.EOF Then
      dg_Tab3_SkuList.SelBookmarks.Add rs_Tab3_SkuList.Bookmark
'      Call Display_SelectedSkuData(rs_Tab3_SkuList.Fields("貨號").Value)
      fam_Tab3_Sku.BackColor = &H8000000C
      fam_Tab3_Sku.Enabled = False
      cmd_Tab3_Cancel.Enabled = False
      cmd_Tab3_Save.Enabled = False
      cmd_Tab3_AddNew.Enabled = True
      cmd_Tab3_Modify.Enabled = True
      cmd_Tab3_Delete.Enabled = True
   End If
End If
End Sub

Private Sub dg_Tab3_SkuList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

'同一行選取
If LastRow = Empty Then Exit Sub

'是否有資料
If rs_Tab3_SkuList Is Nothing Then Exit Sub
If rs_Tab3_SkuList.RecordCount = 0 Then Exit Sub

Call Display_SelectedSkuData(rs_Tab3_SkuList.Fields("貨號").Value)

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub drvLocalDriveT5_Change()
    On Error GoTo DriveError
    dirLocalDirT5.Path = drvLocalDriveT5.Drive
    Exit Sub
DriveError:
    MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
    Resume Next
End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "客戶/車輛基本資料維護作業"
End Sub

Private Sub Form_Load()
'設定 Form 大小、位置
dbsrcFormHeight = 7140
dbsrcFormWidth = 11475

Me.Height = 7650: Me.Width = 11600
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Left = 200
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

'取出所有貨主資料--TRP16M
Dim tmp_cnt As Integer
cmb_Tab0_Storer.Clear
cmb_Tab4_Storer.Clear   '客戶允收天數貨主
str_SQL = "Select Rtrim(StorerKey) as 'StorerKey',Isnull(Rtrim(Short_Name),'') as 'StorerName' From TRP16M Order by StorerKey"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arStorer(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arStorer(tmp_cnt) = tmp_Rs.Fields("StorerKey").Value
      cmb_Tab0_Storer.AddItem tmp_Rs.Fields("StorerKey").Value & Space(7 - Len(Trim(tmp_Rs.Fields("StorerKey").Value))) & tmp_Rs.Fields("StorerName").Value
      cmb_Tab4_Storer.AddItem tmp_Rs.Fields("StorerKey").Value & Space(7 - Len(Trim(tmp_Rs.Fields("StorerKey").Value))) & tmp_Rs.Fields("StorerName").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arStorer) Then
         ReDim Preserve arStorer(UBound(arStorer) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close

'取出所有郵遞區號 TRP02M
cmb_Tab0_Zip.Clear: cmb_Tab1_ZIP.Clear
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
      cmb_Tab0_Zip.AddItem tmp_Rs.Fields("ZIP").Value & Space(5 - Len(Trim(tmp_Rs.Fields("ZIP").Value))) & tmp_Rs.Fields("Descr").Value
      cmb_Tab1_ZIP.AddItem tmp_Rs.Fields("ZIP").Value & Space(5 - Len(Trim(tmp_Rs.Fields("ZIP").Value))) & tmp_Rs.Fields("Descr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arZip) Then
         ReDim Preserve arZip(UBound(arZip) + 10) As String
         ReDim Preserve arZIPArea(UBound(arZIPArea) + 10) As String
      End If
   Loop
End If

'取出所有運送區域代碼 TRP03M
cmb_Tab0_AreaCode.Clear: cmb_Tab1_AreaCode.Clear
str_SQL = "Select Rtrim(Area_Code) as 'AreaCode',Rtrim(Isnull(Description,'')) as Descr  From TRP03M Order by Area_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arAreaCode(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arAreaCode(tmp_cnt) = tmp_Rs.Fields("AreaCode").Value
      cmb_Tab0_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      cmb_Tab1_AreaCode.AddItem tmp_Rs.Fields("AreaCode").Value & Space(10 - Len(Trim(tmp_Rs.Fields("AreaCode").Value))) & tmp_Rs.Fields("Descr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arAreaCode) Then
         ReDim Preserve arAreaCode(UBound(arAreaCode) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close

'取出所有車種資料--TRP15M
cmb_Tab0_VehicleType.Clear: cmb_Tab1_VehicleType.Clear
str_SQL = "Select Rtrim(Vehicle_Type) as 'VType',Isnull(Rtrim(Description),'') as 'VTypeDescr' From TRP15M Order by Vehicle_Type"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arVehicleType(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arVehicleType(tmp_cnt) = tmp_Rs.Fields("VType").Value
      cmb_Tab0_VehicleType.AddItem tmp_Rs.Fields("VType").Value & Space(4 - Len(Trim(tmp_Rs.Fields("VType").Value))) & tmp_Rs.Fields("VTypeDescr").Value
      cmb_Tab1_VehicleType.AddItem tmp_Rs.Fields("VType").Value & Space(4 - Len(Trim(tmp_Rs.Fields("VType").Value))) & tmp_Rs.Fields("VTypeDescr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arVehicleType) Then
         ReDim Preserve arVehicleType(UBound(arVehicleType) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close

'取出所有特殊需求--TRP04M
cmb_Tab0_ExtraDemand1.Clear: cmb_Tab0_ExtraDemand2.Clear
str_SQL = "Select Rtrim(Extra_Demand_Code) as 'ECode',Isnull(Rtrim(Description),'') as 'ECodeDescr' From TRP04M Order by Extra_Demand_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arExtraDemand(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arExtraDemand(tmp_cnt) = tmp_Rs.Fields("ECode").Value
      cmb_Tab0_ExtraDemand1.AddItem tmp_Rs.Fields("ECode").Value & Space(12 - Len(Trim(tmp_Rs.Fields("ECode").Value))) & tmp_Rs.Fields("ECodeDescr").Value
      cmb_Tab0_ExtraDemand2.AddItem tmp_Rs.Fields("ECode").Value & Space(12 - Len(Trim(tmp_Rs.Fields("ECode").Value))) & tmp_Rs.Fields("ECodeDescr").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arExtraDemand) Then
         ReDim Preserve arExtraDemand(UBound(arExtraDemand) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close
'取出所有貨運公司--TRP09M
cmb_Tab1_Company.Clear
str_SQL = "Select Rtrim(Company_Code) as 'CCode',Isnull(Rtrim(C_Name),'') as 'CName' From TRP08M Order by Company_Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arCompany(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arCompany(tmp_cnt) = tmp_Rs.Fields("CCode").Value
      cmb_Tab1_Company.AddItem tmp_Rs.Fields("CCode").Value & Space(5 - Len(Trim(tmp_Rs.Fields("CCode").Value))) & tmp_Rs.Fields("CName").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arCompany) Then
         ReDim Preserve arCompany(UBound(arCompany) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close
'取出所有車廂形式--CODELKUP.ListName = [CARBOXTYPE]
cmb_Tab1_CarBox.Clear
str_SQL = "SELECT RTRIM(Code) AS 代碼, RTRIM(Description) AS 車廂形式 " & _
          "From CodeLKUP Where ListName = 'CARBOXTYPE'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arCarBox(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arCarBox(tmp_cnt) = tmp_Rs.Fields("代碼").Value
      cmb_Tab1_CarBox.AddItem tmp_Rs.Fields("代碼").Value & Space(5 - Len(Trim(tmp_Rs.Fields("代碼").Value))) & tmp_Rs.Fields("車廂形式").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arCarBox) Then
         ReDim Preserve arCarBox(UBound(arCarBox) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close
'取出所有僱用方式--CODELKUP.ListName = [EMPLOYTYPE]
cmb_Tab1_EmployType.Clear
str_SQL = "SELECT RTRIM(Code) AS 代碼, RTRIM(Description) AS 僱用方式 " & _
          "From CodeLKUP Where ListName = 'EMPLOYTYPE'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arEmployType(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arEmployType(tmp_cnt) = tmp_Rs.Fields("代碼").Value
      cmb_Tab1_EmployType.AddItem tmp_Rs.Fields("代碼").Value & Space(5 - Len(Trim(tmp_Rs.Fields("代碼").Value))) & tmp_Rs.Fields("僱用方式").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arEmployType) Then
         ReDim Preserve arEmployType(UBound(arEmployType) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close
'取出所有裝卸方式--CODELKUP.ListName = [LOADUNLOADTYPE]
cmb_Tab1_UnloadType.Clear
str_SQL = "SELECT RTRIM(Code) AS 代碼, RTRIM(Description) AS 裝卸方式 " & _
          "From CodeLKUP Where ListName = 'LOADUNLOADTYPE'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arUnloadType(1) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arUnloadType(tmp_cnt) = tmp_Rs.Fields("代碼").Value
      cmb_Tab1_UnloadType.AddItem tmp_Rs.Fields("代碼").Value & Space(5 - Len(Trim(tmp_Rs.Fields("代碼").Value))) & tmp_Rs.Fields("裝卸方式").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arUnloadType) Then
         ReDim Preserve arUnloadType(UBound(arUnloadType) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close

'取得 搬運工具
cmb_Tab0_PickTool.Clear: tmp_cnt = 0
ReDim arPickTool(1) As String
str_SQL = "SELECT RTRIM(Code) AS 代碼, RTRIM(Description) AS 搬運工具 " & _
          "From CodeLKUP Where ListName = 'MOVETOOL'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
Do While Not tmp_Rs.EOF
   cmb_Tab0_PickTool.AddItem tmp_Rs.Fields("代碼").Value & "  " & tmp_Rs.Fields("搬運工具").Value
   tmp_cnt = tmp_cnt + 1
   If UBound(arPickTool) < tmp_cnt Then
      ReDim Preserve arPickTool(tmp_cnt) As String
   End If
   arPickTool(tmp_cnt - 1) = tmp_Rs.Fields("代碼").Value
   tmp_Rs.MoveNext
Loop
tmp_Rs.Close
cmb_Tab0_PickTool.ListIndex = -1

'取出所有計費代碼 --TRP09M.Car_Type
'cmb_Tab1_CarType.Clear
'str_SQL = "Select distinct Rtrim(isnull(Car_Type,'')) as 'Car_Type' from TRP09M Order by Car_Type"
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
'If Not tmp_Rs.EOF Then
'   Do While Not tmp_Rs.EOF
'      cmb_Tab1_CarType.AddItem tmp_Rs.Fields("Car_Type").Value
'      tmp_Rs.MoveNext
'   Loop
'End If
'tmp_Rs.Close

'取出所有通路體系--TRP18M.consigneekey
cmb_Tab0_Group.Clear
str_SQL = "Select distinct custgroup as 'custGroup' from TRP01M Order by custgroup"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
If Not tmp_Rs.EOF Then
   Do While Not tmp_Rs.EOF
      cmb_Tab0_Group.AddItem RTrim(tmp_Rs.Fields("custGroup").Value)
      tmp_Rs.MoveNext
   Loop
End If
cmb_Tab0_Group = ""
tmp_Rs.Close

blTab0ConsignEventEnable = True

SSTab1.Tab = 0

End Sub
Private Sub Form_Resize()
If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
If Me.ScaleHeight < dbsrcFormHeight Then
   '變小
   SSTab1.Top = (SSTab1.Top - ((dbsrcFormHeight - Me.ScaleHeight) / 2))
   SSTab1.Left = (SSTab1.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2))
     
   dbsrcFormHeight = Me.ScaleHeight
   dbsrcFormWidth = Me.ScaleWidth
Else
   SSTab1.Top = (SSTab1.Top + ((Me.ScaleHeight - dbsrcFormHeight) / 2))
   SSTab1.Left = (SSTab1.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2))
   
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
Set frm_BaseData_ConsigCar = Nothing
End Sub

Private Sub cmd_Exit_Click(Index As Integer)
'離開
Unload Me
End Sub

Private Sub Display_SelectedConsignData(ByVal strStorerkey As String, ByVal strConsigneeKey As String)
'顯示傳入之客戶資料
Call Clear_ConsigneeData

str_SQL = "Select * From TRP01M Where ConsigneeKey = '" & strConsigneeKey & "' and Storerkey = '" & strStorerkey & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "查詢結果：無符合設定條件之客戶基本資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Dim i As Double
txt_Tab0_ConsigneeKey.Text = Trim(tmp_Rs.Fields("ConsigneeKey").Value)
DoEvents: DoEvents
For i = 0 To cmb_Tab0_Storer.ListCount - 1
    If arStorer(i) = Trim(tmp_Rs.Fields("StorerKey").Value) Then
       cmb_Tab0_Storer.ListIndex = i
       Exit For
    End If
Next i
If IsNull(tmp_Rs.Fields("ZIP").Value) Then
   cmb_Tab0_Zip.ListIndex = -1
Else
DoEvents: DoEvents
   For i = 0 To cmb_Tab0_Zip.ListCount - 1
       If arZip(i) = Trim(tmp_Rs.Fields("ZIP").Value) Then
          cmb_Tab0_Zip.ListIndex = i
          Exit For
       End If
   Next i
End If
txt_Tab0_Class.Text = IIf(IsNull(tmp_Rs.Fields("Class").Value), "", Trim(tmp_Rs.Fields("Class").Value))
If IsNull(tmp_Rs.Fields("Area_Code").Value) Then
   cmb_Tab0_AreaCode.ListIndex = -1
Else
DoEvents: DoEvents
   For i = 0 To cmb_Tab0_AreaCode.ListCount - 1
       If arAreaCode(i) = Trim(tmp_Rs.Fields("Area_Code").Value) Then
          cmb_Tab0_AreaCode.ListIndex = i
          Exit For
       End If
   Next i
End If
txt_Tab0_FullName.Text = IIf(IsNull(tmp_Rs.Fields("Full_Name").Value), "", Trim(tmp_Rs.Fields("Full_Name").Value))
txt_Tab0_Address.Text = IIf(IsNull(tmp_Rs.Fields("Address").Value), "", Trim(tmp_Rs.Fields("Address").Value))
txt_Tab0_Contact.Text = IIf(IsNull(tmp_Rs.Fields("Contact").Value), "", Trim(tmp_Rs.Fields("Contact").Value))
txt_Tab0_Phone.Text = IIf(IsNull(tmp_Rs.Fields("Phone").Value), "", Trim(tmp_Rs.Fields("Phone").Value))
txt_Tab0_ShortName.Text = IIf(IsNull(tmp_Rs.Fields("Short_Name").Value), "", Trim(tmp_Rs.Fields("Short_Name").Value))
txt_Tab0_GridCode.Text = IIf(IsNull(tmp_Rs.Fields("Grid_Code").Value), "", Trim(tmp_Rs.Fields("Grid_Code").Value))
If IsNull(tmp_Rs.Fields("Vehicle_Type").Value) Then
   cmb_Tab0_VehicleType.ListIndex = -1
Else
   For i = 0 To cmb_Tab0_VehicleType.ListCount - 1
       If arVehicleType(i) = Trim(tmp_Rs.Fields("Vehicle_Type").Value) Then
          cmb_Tab0_VehicleType.ListIndex = i
          Exit For
       End If
   Next i
End If
If IsNull(tmp_Rs.Fields("Extra_Demand_Code").Value) Then
   cmb_Tab0_ExtraDemand1.ListIndex = -1
Else
   For i = 0 To cmb_Tab0_ExtraDemand1.ListCount - 1
       If arExtraDemand(i) = Trim(tmp_Rs.Fields("Extra_Demand_Code").Value) Then
          cmb_Tab0_ExtraDemand1.ListIndex = i
          Exit For
       End If
   Next i
End If
If IsNull(tmp_Rs.Fields("Extra_Demand_Code2").Value) Then
   cmb_Tab0_ExtraDemand2.ListIndex = -1
Else
   For i = 0 To cmb_Tab0_ExtraDemand2.ListCount - 1
       If arExtraDemand(i) = Trim(tmp_Rs.Fields("Extra_Demand_Code2").Value) Then
          cmb_Tab0_ExtraDemand2.ListIndex = i
          Exit For
       End If
   Next i
End If
txt_Tab0_ChannelType.Text = IIf(IsNull(tmp_Rs.Fields("Channel_Type").Value), "", Trim(tmp_Rs.Fields("Channel_Type").Value))
txt_Tab0_Channel.Text = IIf(IsNull(tmp_Rs.Fields("Channel").Value), "", Trim(tmp_Rs.Fields("Channel").Value))
txt_Tab0_UnLoad.Text = IIf(IsNull(tmp_Rs.Fields("Unload_Type").Value), "", Trim(tmp_Rs.Fields("Unload_Type").Value))

If IsNull(tmp_Rs.Fields("Multi_Customer").Value) Then
   chk_Tab0_MultiCustomer.Value = vbUnchecked
Else
   If Trim(tmp_Rs.Fields("Multi_Customer").Value) = "N" Then
      chk_Tab0_MultiCustomer.Value = vbUnchecked
   Else
      chk_Tab0_MultiCustomer.Value = vbChecked
   End If
End If

If IsNull(tmp_Rs.Fields("DC").Value) Then
   chk_Tab0_MultiCustomer.Value = vbUnchecked
Else
   If Trim(tmp_Rs.Fields("DC").Value) = "N" Then
      chkDC.Value = vbUnchecked
   Else
      chkDC.Value = vbChecked
   End If
End If


If IsNull(tmp_Rs.Fields("PICK_TOOL").Value) Then
   cmb_Tab0_PickTool.ListIndex = -1
Else
   For i = 0 To cmb_Tab0_PickTool.ListCount - 1
       If arPickTool(i) = Trim(tmp_Rs.Fields("pick_tool").Value) Then
          cmb_Tab0_PickTool.ListIndex = i
          Exit For
       End If
   Next i
End If

'通路體系
cmb_Tab0_Group = RTrim(tmp_Rs.Fields("CustGroup"))

txt_Tab0_CodeDate1.Text = IIf(IsNull(tmp_Rs.Fields("CodeDate1").Value), "", Trim(tmp_Rs.Fields("CodeDate1").Value))
txt_Tab0_CodeDate2.Text = IIf(IsNull(tmp_Rs.Fields("CodeDate2").Value), "", Trim(tmp_Rs.Fields("CodeDate2").Value))
txt_Tab0_CodeDate3.Text = IIf(IsNull(tmp_Rs.Fields("CodeDate3").Value), "", Trim(tmp_Rs.Fields("CodeDate3").Value))
txt_Tab0_Fax.Text = IIf(IsNull(tmp_Rs.Fields("fax").Value), "", Trim(tmp_Rs.Fields("fax").Value))
txt_Tab0_Stamp.Text = IIf(IsNull(tmp_Rs.Fields("stamp").Value), "", Trim(tmp_Rs.Fields("stamp").Value))
txt_Tab0_Penalties.Text = IIf(IsNull(tmp_Rs.Fields("Penalties").Value), "", Trim(tmp_Rs.Fields("Penalties").Value))
txt_Tab0_PalletType.Text = IIf(IsNull(tmp_Rs.Fields("PalletType").Value), "", Trim(tmp_Rs.Fields("PalletType").Value))
txt_Tab0_PalletSpec.Text = IIf(IsNull(tmp_Rs.Fields("PalletSpec").Value), "", Trim(tmp_Rs.Fields("PalletSpec").Value))
txt_Tab0_Notes.Text = IIf(IsNull(tmp_Rs.Fields("Notes").Value), "", Trim(tmp_Rs.Fields("Notes").Value))

'允收期
cmdCodeDateRate = tmp_Rs("CodeDateRate")

tmp_Rs.Close

End Sub

Private Sub Clear_ConsigneeData()
'清除 客戶資料 畫面之欄位值
txt_Tab0_ConsigneeKey.Text = ""
cmb_Tab0_Storer.ListIndex = -1
cmb_Tab0_Zip.ListIndex = -1
txt_Tab0_Class.Text = ""
cmb_Tab0_AreaCode.ListIndex = -1
txt_Tab0_FullName.Text = ""
txt_Tab0_Address.Text = ""
txt_Tab0_Contact.Text = ""
txt_Tab0_Phone.Text = ""
txt_Tab0_ShortName.Text = ""
txt_Tab0_GridCode.Text = ""
cmb_Tab0_VehicleType.ListIndex = -1
cmb_Tab0_ExtraDemand1.ListIndex = -1: cmb_Tab0_ExtraDemand2.ListIndex = -1
txt_Tab0_Channel.Text = ""
txt_Tab0_ChannelType.Text = ""
txt_Tab0_UnLoad.Text = ""
chk_Tab0_MultiCustomer.Value = vbUnchecked
chkDC.Value = vbUnchecked
txt_Tab0_Fax.Text = ""
txt_Tab0_CodeDate1.Text = ""
txt_Tab0_CodeDate2.Text = ""
txt_Tab0_CodeDate3.Text = ""
txt_Tab0_Stamp.Text = ""
txt_Tab0_Penalties.Text = ""
txt_Tab0_PalletType.Text = ""
txt_Tab0_PalletSpec.Text = ""
txt_Tab0_Notes.Text = ""
cmb_Tab0_Group.Text = ""
cmdCodeDateRate = ""
End Sub

Private Function Check_ComsigneeData() As Boolean
'客戶基本資料檢核

Check_ComsigneeData = False
msg_text = ""

If cmb_Tab0_Zip.ListIndex = -1 Then
   If msg_text = "" Then
      msg_text = "未輸入郵遞區號"
   Else
      msg_text = msg_text & vbCrLf & "未輸入郵遞區號"
   End If
End If

If Len(Trim(txt_Tab0_FullName.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入客戶名稱"
   Else
      msg_text = msg_text & vbCrLf & "未輸入客戶名稱"
   End If
End If

If Len(Trim(txt_Tab0_ShortName.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入客戶簡稱"
   Else
      msg_text = msg_text & vbCrLf & "未輸入客戶簡稱"
   End If
End If

If Len(Trim(txt_Tab0_Address.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入運送地址"
   Else
      msg_text = msg_text & vbCrLf & "未輸入運送地址"
   End If
End If

If txt_Tab0_ConsigneeKey.Enabled = True And UCase(Left(Trim(txt_Tab0_ConsigneeKey.Text), 4)) = "BEST" Then MsgBox "客編開頭不能使用""BEST""保留字!!", 16, "注意": Exit Function

If Trim(cmdCodeDateRate) = "1/2" Or Trim(cmdCodeDateRate) = "2/3" Or Trim(cmdCodeDateRate) = "" Then
Else
   If msg_text = "" Then
      msg_text = "錯誤資料 [允收期限]，允收期限只允許選擇1/2、2/3與空白 "
   Else
      msg_text = msg_text & vbCrLf & "允收期限錯誤"
   End If
End If

If Len(Trim(txt_Tab0_CodeDate1.Text)) = 0 And mySplit(cmb_Tab0_Storer, " ", 0) = "LTKK01" Then
   If msg_text = "" Then
      msg_text = "未輸入 [啤酒允收期]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入客戶允收期(錯誤的允收期將影響配貨正確性)"
   End If
End If

If Len(Trim(txt_Tab0_CodeDate2.Text)) = 0 And mySplit(cmb_Tab0_Storer, " ", 0) = "LTKK01" Then
   If msg_text = "" Then
      msg_text = "未輸入 [清酒允收期]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入客戶允收期(錯誤的允收期將影響配貨正確性)"
   End If
End If

If Len(Trim(txt_Tab0_CodeDate3.Text)) = 0 And mySplit(cmb_Tab0_Storer, " ", 0) = "LTKK01" Then
   If msg_text = "" Then
      msg_text = "未輸入 [飲料允收期]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入客戶允收期(錯誤的允收期將影響配貨正確性)"
   End If
End If

If msg_text = "" Then
   Check_ComsigneeData = True
Else
   msg_text = "客戶資料異常，請修正後再執行 [存 檔]：" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If
End Function

Private Sub Display_SelectedCarData(ByVal strCarID As String)
'顯示傳入之車輛基本資料
Call Clear_CarData

str_SQL = "Select * From TRP09M Where Vehicle_ID_No = '" & strCarID & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "查詢結果：無符合設定條件之車輛基本資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Dim i As Double
txt_Tab1_CarID.Text = Trim(tmp_Rs.Fields("Vehicle_ID_No").Value)
txt_Tab1_Driver.Text = IIf(IsNull(tmp_Rs.Fields("Driver").Value), "", Trim(tmp_Rs.Fields("Driver").Value))
txt_Tab1_Phone.Text = IIf(IsNull(tmp_Rs.Fields("DRIVER_PHONE").Value), "", Trim(tmp_Rs.Fields("DRIVER_PHONE").Value))
If IsNull(tmp_Rs.Fields("ZIP").Value) Then
   cmb_Tab1_ZIP.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_ZIP.ListCount - 1
       If arZip(i) = Trim(tmp_Rs.Fields("ZIP").Value) Then
          cmb_Tab1_ZIP.ListIndex = i
          Exit For
       End If
   Next i
End If
If IsNull(tmp_Rs.Fields("Area_Code").Value) Then
   cmb_Tab1_AreaCode.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_AreaCode.ListCount - 1
       If arAreaCode(i) = Trim(tmp_Rs.Fields("Area_Code").Value) Then
          cmb_Tab1_AreaCode.ListIndex = i
          Exit For
       End If
   Next i
End If
If IsNull(tmp_Rs.Fields("Vehicle_Type").Value) Then
   cmb_Tab1_VehicleType.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_VehicleType.ListCount - 1
       If arVehicleType(i) = Trim(tmp_Rs.Fields("Vehicle_Type").Value) Then
          cmb_Tab1_VehicleType.ListIndex = i
          Exit For
       End If
   Next i
End If
If IsNull(tmp_Rs.Fields("TRP_COMPANY_CODE").Value) Then
   cmb_Tab1_Company.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_Company.ListCount - 1
       If arCompany(i) = Trim(tmp_Rs.Fields("TRP_COMPANY_CODE").Value) Then
          cmb_Tab1_Company.ListIndex = i
          Exit For
       End If
   Next i
End If
txt_Tab1_CarWeight.Text = IIf(IsNull(tmp_Rs.Fields("CAR_WIEGHT").Value), "", Trim(tmp_Rs.Fields("CAR_WIEGHT").Value))
txt_Tab1_CarHeight.Text = IIf(IsNull(tmp_Rs.Fields("CAR_HEIGHT").Value), "", Trim(tmp_Rs.Fields("CAR_HEIGHT").Value))
If IsNull(tmp_Rs.Fields("CARBOX_TYPE").Value) Then
   cmb_Tab1_CarBox.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_CarBox.ListCount - 1
   DoEvents: DoEvents
       If arCarBox(i) = Trim(tmp_Rs.Fields("CARBOX_TYPE").Value) Then
          cmb_Tab1_CarBox.ListIndex = i
          Exit For
       End If
   Next i
End If

If IsNull(tmp_Rs.Fields("EMPLOY_TYPE").Value) Then
   cmb_Tab1_EmployType.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_EmployType.ListCount - 1
       If arEmployType(i) = Trim(tmp_Rs.Fields("EMPLOY_TYPE").Value) Then
          cmb_Tab1_EmployType.ListIndex = i
          Exit For
       End If
   Next i
End If

If IsNull(tmp_Rs.Fields("UNLAODING_TYPE").Value) Then
   cmb_Tab1_UnloadType.ListIndex = -1
Else
   For i = 0 To cmb_Tab1_UnloadType.ListCount - 1
       If arUnloadType(i) = Trim(tmp_Rs.Fields("UNLAODING_TYPE").Value) Then
          cmb_Tab1_UnloadType.ListIndex = i
          Exit For
       End If
   Next i
End If

If IsNull(tmp_Rs.Fields("PND").Value) Then
   chkPND.Value = vbUnchecked
Else
   If Trim(tmp_Rs.Fields("pnd").Value) = "N" Then
      chkPND.Value = vbUnchecked
   Else
      chkPND.Value = vbChecked
   End If
End If

txtAPFix = (tmp_Rs.Fields("APFix").Value)

txt_Tab1_WeightCapacity.Text = IIf(IsNull(tmp_Rs.Fields("LOADING_SIZE").Value), "", Trim(tmp_Rs.Fields("LOADING_SIZE").Value))
txt_Tab1_VolumnCapacity.Text = IIf(IsNull(tmp_Rs.Fields("MAX_CUBIC_CAPACITY").Value), "", Trim(tmp_Rs.Fields("MAX_CUBIC_CAPACITY").Value))
txt_Tab1_PalletCapacity.Text = IIf(IsNull(tmp_Rs.Fields("PALLET_CAPACITY").Value), "", Trim(tmp_Rs.Fields("PALLET_CAPACITY").Value))
txt_Tab1_Description.Text = IIf(IsNull(tmp_Rs.Fields("Description").Value), "", Trim(tmp_Rs.Fields("Description").Value))
cmb_Tab1_CarType.Text = IIf(IsNull(tmp_Rs.Fields("CAR_TYPE").Value), "", Trim(tmp_Rs.Fields("CAR_TYPE").Value))
txt_Tab1_Receiver = tmp_Rs("receiver") & ""
txt_Tab1_Receiver = tmp_Rs("receiver") & ""
txtAdd = tmp_Rs("addwho") & " / " & tmp_Rs("adddate")
txtEdit = tmp_Rs("editwho") & " / " & tmp_Rs("editdate")
tmp_Rs.Close

End Sub
Private Sub Display_SelectedCompanyData(ByVal strCompanyCode As String)
'顯示傳入之貨運公司基本資料
Call Clear_CompanyData

str_SQL = "Select * From TRP08M Where Company_Code = '" & strCompanyCode & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "查詢結果：無符合設定條件之貨運公司基本資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Dim i As Double
txt_Tab2_CompanyCode.Text = Trim(tmp_Rs.Fields("company_code").Value)
txt_Tab2_CName.Text = IIf(IsNull(tmp_Rs.Fields("C_NAME").Value), "", Trim(tmp_Rs.Fields("C_NAME").Value))
txt_Tab2_EName.Text = IIf(IsNull(tmp_Rs.Fields("E_NAME").Value), "", Trim(tmp_Rs.Fields("E_NAME").Value))
txt_Tab2_Address.Text = IIf(IsNull(tmp_Rs.Fields("Address").Value), "", Trim(tmp_Rs.Fields("Address").Value))
txt_Tab2_Descr.Text = IIf(IsNull(tmp_Rs.Fields("Description").Value), "", Trim(tmp_Rs.Fields("Description").Value))
txt_Tab2_Contact.Text = IIf(IsNull(tmp_Rs.Fields("Contact").Value), "", Trim(tmp_Rs.Fields("Contact").Value))
txt_Tab2_Phone.Text = IIf(IsNull(tmp_Rs.Fields("Phone").Value), "", Trim(tmp_Rs.Fields("Phone").Value))
txt_Tab2_ShortName.Text = IIf(IsNull(tmp_Rs.Fields("Short_Name").Value), "", Trim(tmp_Rs.Fields("Short_Name").Value))

tmp_Rs.Close

End Sub

Private Sub Display_SelectedSkuData(ByVal strSkuCode As String)
'顯示傳入之貨號基本資料
Call Clear_SkuData

str_SQL = "Select StorerKey,Sku,DESCR, STDGROSSWGT, busr4," & _
    "SKUGROUP,BUSR1,NOTES1,NOTES2 From gv_Skuxpack Where Sku = '" & strSkuCode & "'"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "查詢結果：無符合設定條件之資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Dim i As Double
txt_Tab3_StorerKey.Text = Trim(tmp_Rs.Fields("StorerKey").Value)
txt_Tab3_Sku.Text = IIf(IsNull(tmp_Rs.Fields("Sku").Value), "", Trim(tmp_Rs.Fields("Sku").Value))
txt_Tab3_DESCR.Text = IIf(IsNull(tmp_Rs.Fields("DESCR").Value), "", Trim(tmp_Rs.Fields("DESCR").Value))
txt_Tab3_STDGROSSWGT.Text = IIf(IsNull(tmp_Rs.Fields("STDGROSSWGT").Value), "", Trim(tmp_Rs.Fields("STDGROSSWGT").Value))
txt_Tab3_BUSR4.Text = IIf(IsNull(tmp_Rs.Fields("busr4").Value), "", Trim(tmp_Rs.Fields("busr4").Value))
txt_Tab3_SKUGROUP.Text = IIf(IsNull(tmp_Rs.Fields("SKUGROUP").Value), "", Trim(tmp_Rs.Fields("SKUGROUP").Value))
txt_Tab3_BUSR1.Text = IIf(IsNull(tmp_Rs.Fields("BUSR1").Value), "", Trim(tmp_Rs.Fields("BUSR1").Value))
txt_Tab3_NOTES1.Text = IIf(IsNull(tmp_Rs.Fields("NOTES1").Value), "", Trim(tmp_Rs.Fields("NOTES1").Value))
txt_Tab3_NOTES2.Text = IIf(IsNull(tmp_Rs.Fields("NOTES2").Value), "", Trim(tmp_Rs.Fields("NOTES2").Value))
tmp_Rs.Close

End Sub

Private Sub Clear_CarData()
'清除 車輛資料 畫面之欄位值
txt_Tab1_CarID.Text = ""
txt_Tab1_Driver.Text = ""
txt_Tab1_Phone.Text = ""
cmb_Tab1_ZIP.ListIndex = -1
cmb_Tab1_AreaCode.ListIndex = -1
cmb_Tab1_Company.ListIndex = -1
txt_Tab1_CarWeight.Text = ""
txt_Tab1_CarHeight.Text = ""
cmb_Tab1_VehicleType.ListIndex = -1
cmb_Tab1_CarBox.ListIndex = -1
cmb_Tab1_EmployType.ListIndex = -1
cmb_Tab1_UnloadType.ListIndex = -1
txt_Tab1_VolumnCapacity.Text = ""
txt_Tab1_WeightCapacity.Text = ""
txt_Tab1_PalletCapacity.Text = ""
txt_Tab1_Receiver.Text = ""
cmb_Tab1_CarType = ""
chkPND = vbUnchecked
txtAPFix = ""
End Sub

Private Function Check_CarData() As Boolean
'車輛基本資料檢核
Check_CarData = False
msg_text = ""
If Len(Trim(txt_Tab1_CarID.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入 [車牌號碼]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入車牌號碼"
   End If
End If

'資料檢核
If Len(RTrim(cmb_Tab1_CarType)) = 0 Then msg_text = msg_text & vbCrLf & "未選擇""計費類別""！"
If Len(RTrim(cmb_Tab1_Company)) = 0 Then msg_text = msg_text & vbCrLf & "未選擇""貨運公司""！"
If Len(RTrim(txt_Tab1_WeightCapacity)) = 0 Then msg_text = msg_text & vbCrLf & "未輸入""可裝載重量""！"
If Len(RTrim(txt_Tab1_VolumnCapacity)) = 0 Then msg_text = msg_text & vbCrLf & "未輸入""可裝載材積""！"
If IsNumeric(txtAPFix) = False Then msg_text = msg_text & vbCrLf & "未輸入或格式錯誤""運費調整%""！"
If Val(txtAPFix) < 0 Then msg_text = msg_text & vbCrLf & "不得為負值""運費調整%""！"

If msg_text = "" Then
   Check_CarData = True
Else
   msg_text = "車輛資料異常，請修正後再執行 [存 檔]：" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If

End Function

Private Sub Clear_CompanyData()
'清除 貨運公司 畫面之欄位值
txt_Tab2_CompanyCode.Text = ""
txt_Tab2_CName.Text = ""
txt_Tab2_EName.Text = ""
txt_Tab2_Address.Text = ""
txt_Tab2_Descr.Text = ""
txt_Tab2_Contact.Text = ""
txt_Tab2_Phone.Text = ""
txt_Tab2_ShortName.Text = ""
End Sub

Private Sub Clear_SkuData()
'清除 貨號 畫面之欄位值
txt_Tab3_StorerKey.Text = ""
txt_Tab3_Sku.Text = ""
txt_Tab3_DESCR.Text = ""
txt_Tab3_STDGROSSWGT.Text = ""
txt_Tab3_BUSR4.Text = ""
txt_Tab3_SKUGROUP.Text = ""
txt_Tab3_BUSR1.Text = ""
txt_Tab3_NOTES1.Text = ""
txt_Tab3_NOTES2.Text = ""
End Sub


Private Function Check_CompanyData() As Boolean
'貨運公司基本資料檢核
Check_CompanyData = False
msg_text = ""
If Len(Trim(txt_Tab2_CompanyCode.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入 [公司代碼]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入公司代碼"
   End If
End If
If Len(Trim(txt_Tab2_CName.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入 [公司中文名稱]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入 [公司中文名稱]"
   End If
End If

If msg_text = "" Then
   Check_CompanyData = True
Else
   msg_text = "貨運公司資料異常，請修正後再執行 [存 檔]：" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If
End Function

Private Function Check_SkuData() As Boolean
'貨號資料檢核
Check_SkuData = False
msg_text = ""
If Len(Trim(txt_Tab3_Sku.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入 [貨號]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入貨號"
   End If
End If
If Len(Trim(txt_Tab3_DESCR.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入 [中文名稱]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入 [中文名稱]"
   End If
End If
If Len(Trim(txt_Tab3_STDGROSSWGT.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入 [每箱重]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入 [每箱重]"
   End If
End If
If Len(Trim(txt_Tab3_BUSR4.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入 [每箱材]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入 [每箱材]"
   End If
End If

If msg_text = "" Then
   Check_SkuData = True
Else
   msg_text = "商品資料異常，請修正後再執行 [存 檔]：" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If
End Function

Public Sub frm_BaseData_ConsigCar_rsFilterAndSort(ByVal strCode As String, ByVal strReturn As String)
'表單公用副程式，由 frm_RS_FilterAndSort 表單呼叫
'傳入值：strCode      動作識別碼
'                     [FILTER] 自訂篩選    [SORT] 排序
'        strReturn    篩選 or 排序 之設定字串

Select Case strCode
       Case "FILTER"  '自訂篩選
            Select Case UCase(strRSName_FilterAndSort)
                   Case "RS_TAB0_CONSIGNEELIST"   '客戶基本資料
                        blTab0ConsignEventEnable = False
                        rs_Tab0_ConsigneeList.Filter = adFilterNone
                        rs_Tab0_ConsigneeList.Filter = strReturn
                        If rs_Tab0_ConsigneeList.RecordCount = 0 Then
                           msg_text = "抱歉ㄟ，找不到符合條件的資料喔"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_Tab0_ConsigneeList.Filter = adFilterNone
                           rs_Tab0_ConsigneeList.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
                           blTab0ConsignEventEnable = True
                           Exit Sub
                        End If
                        blTab0ConsignEventEnable = True
                   Case "RS_TAB1_CARLIST"         '車輛基本資料
                        blTab1CarEventEnable = False
                        rs_Tab1_CarList.Filter = adFilterNone
                        rs_Tab1_CarList.Filter = strReturn
                        If rs_Tab1_CarList.RecordCount = 0 Then
                           msg_text = "抱歉ㄟ，找不到符合條件的資料喔"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_Tab1_CarList.Filter = adFilterNone
                           rs_Tab1_CarList.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
                           blTab1CarEventEnable = True
                           Exit Sub
                        End If
                        blTab1CarEventEnable = True
                   Case "RS_TAB2_TRPCOMPANYLIST"         '運輸公司基本資料
                        blTab2CompanyEventEnable = False
                        rs_Tab2_TRPCompanyList.Filter = adFilterNone
                        rs_Tab2_TRPCompanyList.Filter = strReturn
                        If rs_Tab2_TRPCompanyList.RecordCount = 0 Then
                           msg_text = "抱歉ㄟ，找不到符合條件的資料喔"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_Tab2_TRPCompanyList.Filter = adFilterNone
                           rs_Tab2_TRPCompanyList.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
                           blTab2CompanyEventEnable = True
                           Exit Sub
                        End If
                        blTab2CompanyEventEnable = True
                    Case "RS_TAB4_ACCEPTABLELIST"   '客戶允收天數基本資料
                        blTab4AcceptableEventEnable = False
                        rs_Tab4_AcceptableList.Filter = adFilterNone
                        rs_Tab4_AcceptableList.Filter = strReturn
                        If rs_Tab4_AcceptableList.RecordCount = 0 Then
                           msg_text = "抱歉ㄟ，找不到符合條件的資料喔"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_Tab4_AcceptableList.Filter = adFilterNone
                           rs_Tab4_AcceptableList.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
                           blTab4AcceptableEventEnable = True
                           Exit Sub
                        End If
                        blTab4AcceptableEventEnable = True
                    Case "RS_TAB3_SKULIST"   '貨號基本資料
                        blTab3skuEventEnable = False
                        rs_Tab3_SkuList.Filter = adFilterNone
                        rs_Tab3_SkuList.Filter = strReturn
                        If rs_Tab3_SkuList.RecordCount = 0 Then
                           msg_text = "抱歉ㄟ，找不到符合條件的資料喔"
                           MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                           rs_Tab3_SkuList.Filter = adFilterNone
                           rs_Tab3_SkuList.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
                           blTab3skuEventEnable = True
                           Exit Sub
                        End If
                        blTab3skuEventEnable = True
            End Select
       Case "SORT"    '排序
            Select Case UCase(strRSName_FilterAndSort)
                   Case "RS_TAB0_CONSIGNEELIST"   '客戶基本資料
                        blTab0ConsignEventEnable = False
                        rs_Tab0_ConsigneeList.Sort = strReturn
                        blTab0ConsignEventEnable = True
                   Case "RS_TAB1_CARLIST"        '車輛基本資料
                        blTab1CarEventEnable = False
                        rs_Tab1_CarList.Sort = strReturn
                        blTab1CarEventEnable = True
                   Case "RS_TAB2_TRPCOMPANYLIST"    '貨運公司基本資料
                        blTab2CompanyEventEnable = False
                        rs_Tab2_TRPCompanyList.Sort = strReturn
                        blTab2CompanyEventEnable = True
                   Case "RS_TAB4_ACCEPTABLELIST"   '客戶允收天數基本資料
                        blTab4AcceptableEventEnable = False
                        rs_Tab4_AcceptableList.Sort = strReturn
                        blTab4AcceptableEventEnable = True
                   Case "RS_TAB3_SKULIST"   '貨號基本資料
                        blTab3skuEventEnable = False
                        rs_Tab3_SkuList.Sort = strReturn
                        blTab3skuEventEnable = True
            End Select
End Select
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab
End Sub

Private Sub txt_Tab1_CarID_LostFocus()

Dim rsTmp As New ADODB.Recordset
rsTmp.Open "select vehicle_id_no from trp09m where rtrim(vehicle_id_no) = '" & RTrim(txt_Tab1_CarID) & "' ", cn
If rsTmp.EOF = False Then
    MsgBox "新增車號重複!!", 64, "注意"
       txt_Tab1_CarID.SelStart = 0: txt_Tab1_CarID.SelLength = Len(txt_Tab1_CarID.Text)
   txt_Tab1_CarID.SetFocus

End If

End Sub
Sub SendMail(strConsigneeKey As String)

'LTKK01客戶主檔異動自動 Mail 通知
'Gary edit strto 20170424 irene@mail.kirin.com.tw;ken@mail.kirin.com.tw;shiu@mail.kirin.com.tw;celine@mail.kirin.com.tw;pinkhsu@mail.kirin.com.tw
If mySplit(cmb_Tab0_Storer, " ", 0) = "LTKK01" Then
    
    Dim strFrom As String, strTo As String, strCC As String, strBCC As String, strSubject As String, strTextbody As String, strAlways As String, strEmailID As String, strEmailPW As String, strAddAttachment As String
    
    '讀取ini參數
    Dim objIni As New vbIniFile
    objIni.FileName = App.Path & "/" & App.title & ".ini"
    
    strFrom = objIni.ReadData("INVCHECKEMAIL_LTKK01", "From", "")
    strTo = objIni.ReadData("INVCHECKEMAIL_LTKK01", "To", "")
    strCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "CC", "")
    strBCC = objIni.ReadData("INVCHECKEMAIL_LTKK01", "BCC", "")
    strSubject = "客戶主檔異動(" & strConsigneeKey & "-" & txt_Tab0_ShortName & ")"
    strTextbody = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Textbody", "")
    strEmailID = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailID", "")
    strEmailPW = objIni.ReadData("INVCHECKEMAIL_LTKK01", "EmailPW", "")
    strAlways = objIni.ReadData("INVCHECKEMAIL_LTKK01", "Always", "NO")
    
    
    strFrom = "Tkedi@bestlog.com.tw"
    strTo = "ken@mail.kirin.com.tw;shiu@mail.kirin.com.tw;celine@mail.kirin.com.tw;pinkhsu@mail.kirin.com.tw"
    strCC = "Tkedi@bestlog.com.tw"
'     strTo = "gemini@bestlog.com.tw"
     strCC = ""
    
    Set objIni = Nothing
    
    Dim rsTmp As New ADODB.Recordset
    
    If Len(RTrim(strFrom)) > 0 Then '有寄件者
    
        str_SQL = "select * from gv_webCustomer where 客戶編號 + 地址別 = '" & strConsigneeKey & "' "

        rsTmp.Open str_SQL, cn
        
        '如果無資料也要mail
        If Not rsTmp.EOF Or UCase(RTrim(strAlways)) = "YES" Then
            
            strAddAttachment = "C:\BEST\DYDC_Best\LTKK01\客戶主檔異動\客戶主檔異動_" & Format(Now, "yyyymmddhhMMss") & ".xls"
            
            Call Recordset2Excel("客戶主檔異動", rsTmp)
            If Dir("C:\BEST\DYDC_Best\LTKK01\客戶主檔異動", vbDirectory) = "" Then MkDirs "C:\BEST\DYDC_Best\LTKK01\客戶主檔異動"
            MyXlsApp.ActiveWorkbook.SaveAs strAddAttachment
            MyXlsApp.Quit: Set MyXlsApp = Nothing
    
            '傳送郵件
            Dim objEmail As Object
            Set objEmail = CreateObject("CDO.Message")
        
            objEmail.From = strFrom
            objEmail.To = strTo
            objEmail.CC = strCC   ' 副本
            objEmail.BCC = strBCC ' 密件副本
            objEmail.Subject = strSubject
            objEmail.TextBody = strTextbody
            objEmail.AddAttachment strAddAttachment
        
            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "bestlog.com.tw"
            objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            'SMTP 伺服器需要驗證時
            If Len(RTrim(strEmailID)) > 0 Then
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendusername") = strEmailID
                objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strEmailPW
            End If
            objEmail.Configuration.Fields.Update
            objEmail.Send
        
            MsgBox "LTKK01客戶主檔異動，系統已發Mail通知貨主!", , "自動Mail通知"
        
            Set objEmail = Nothing
        End If
    End If
End If

End Sub


Private Sub cmd_Tab4_AcceptableReset_Click()
'客戶允收天數基本資料 >> 取消篩選排序
'移除篩選條件，重設排序依據
If rs_Tab4_AcceptableList Is Nothing Then Exit Sub
 blTab4AcceptableEventEnable = False
 rs_Tab4_AcceptableList.Filter = adFilterNone
 rs_Tab4_AcceptableList.Sort = "編號 ASC"  '原始排序，一般資料序號由小至大
 blTab4AcceptableEventEnable = True
End Sub

Private Sub cmd_Tab4_AcceptableShow_Click()
'允收天數資料 >> 顯示所有客戶+貨號
On Error GoTo err_Handle
Screen.MousePointer = vbHourglass
DoEvents: DoEvents
Set dg_Tab4_AcceptableList.DataSource = Nothing
Call ReDim_Recordset(rs_Tab4_AcceptableList)
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select distinct Rtrim(a.storerkey) as 貨主,Rtrim(a.Customer) as 客戶編號," & _
" (select top 1 Rtrim(t1m.Full_Name) from TRP01M t1m where t1m.ConsigneeKey = a.customer order by  Rtrim(t1m.Full_Name) desc)  as 客戶名稱,Rtrim(a.ItemNo) as 產品編號,Rtrim(s.descr) as 產品名稱,a.allowdays as 允收天數" & _
" from Acceptable a " & _
"inner join " & strWMSDB & "..sku s on s.sku=a.itemno " & _
"order by Rtrim(a.Customer),Rtrim(a.ItemNo) "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF Then
    msg_text = "資料錯誤：查詢結果傳回 0 列客戶允收天數資料"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
    Set rs_Tab4_AcceptableList = Nothing
    Exit Sub
End If
Call Replication_Recordset(tmp_Rs, rs_Tab4_AcceptableList)
tmp_Rs.Close

blTab4AcceptableEventEnable = False
With dg_Tab4_AcceptableList
     .ColumnHeaders = True           '決定是否在 DataGrid 控制項中顯示資料行行首。
     .HeadLines = 1.5                '顯示在 DataGrid 控制項的資料行行首中的文字行數。
     .RowDividerStyle = dbgRaised    'DataGrid 控制項資料列間的框線樣式。
     .RowHeight = 270                '設定DataGrid 控制項中所有資料列的高
End With
rs_Tab4_AcceptableList.MoveFirst
Set dg_Tab4_AcceptableList.DataSource = rs_Tab4_AcceptableList
With dg_Tab4_AcceptableList
    .RowHeight = 250
    .Columns(0).Width = 500        '序號
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 800        '貨主
    .Columns(1).Alignment = dbgLeft
    .Columns(2).Width = 1000       '客戶編號
    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 3000       '客戶名稱
    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000       '產品編號
    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 3000       '產品名稱
    .Columns(5).Alignment = dbgLeft
    .Columns(6).Width = 800        '允收天數
    .Columns(6).Alignment = dbgLeft
End With
blTab4AcceptableEventEnable = True
Call Clear_AcceptableData
Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-客戶允收天數資料-顯示所有資料", Me.Caption, "cmd_Tab4_AcceptableShow_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4_AddNew_Click()
'客戶允收天數資料 >> 轉換至新增模式
If Not rs_Tab4_AcceptableList Is Nothing Then
    If dg_Tab4_AcceptableList.SelBookmarks.Count > 0 Then dg_Tab4_AcceptableList.SelBookmarks.Remove 0
End If
fam_Tab4_Acceptable.BackColor = &HC0FFC0
fam_Tab4_Acceptable.Enabled = True
txt_Tab4_ConsigneeKey.Enabled = True
txt_Tab4_Sku.Enabled = True
Call Clear_AcceptableData
cmd_Tab4_Save.Enabled = True
cmd_Tab4_Cancel.Enabled = True
cmd_Tab4_AddNew.Enabled = False
cmd_Tab4_Modify.Enabled = False
cmd_Tab4_Delete.Enabled = False
cmb_Tab4_Storer.ListIndex = 0
txt_Tab4_ConsigneeKey.BackColor = &HFFFFFF
txt_Tab4_Sku.BackColor = &HFFFFFF
dg_Tab4_AcceptableList.Enabled = False
cmd_Tab4_AcceptableShow.Enabled = False
cmd_Tab4_AcceptableQuery.Enabled = False
cmd_Tab4_AcceptableReset.Enabled = False
End Sub

Private Sub cmd_Tab4_Cancel_Click()
'客戶允收天數資料 >> 取消修改
Call Clear_AcceptableData
If txt_Tab4_ConsigneeKey.Enabled = False And txt_Tab4_Sku.Enabled = False Then
    If Not rs_Tab4_AcceptableList Is Nothing Then
        dg_Tab4_AcceptableList.SelBookmarks.Add rs_Tab4_AcceptableList.Bookmark
        Call Display_SelectedAcceptableData(rs_Tab4_AcceptableList.Fields("客戶編號").Value, rs_Tab4_AcceptableList.Fields("產品編號").Value)
    End If
End If
fam_Tab4_Acceptable.BackColor = &H8000000C
fam_Tab4_Acceptable.Enabled = False
cmd_Tab4_Cancel.Enabled = False
cmd_Tab4_Save.Enabled = False
cmd_Tab4_AddNew.Enabled = True
cmd_Tab4_Modify.Enabled = True
cmd_Tab4_Delete.Enabled = True
txt_Tab4_ConsigneeKey.BackColor = &H8000000F
txt_Tab4_Sku.BackColor = &H8000000F
dg_Tab4_AcceptableList.Enabled = True
cmd_Tab4_AcceptableShow.Enabled = True
cmd_Tab4_AcceptableQuery.Enabled = True
cmd_Tab4_AcceptableReset.Enabled = True
End Sub

Private Sub cmd_Tab4_Delete_Click()
'客戶允收天數資料 >> 刪除
Dim blDelete As Boolean
blDelete = True
msg_text = ""

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus
Screen.MousePointer = vbHourglass

If Len(RTrim(txt_Tab4_ConsigneeKey.Text)) = 0 Or Len(RTrim(txt_Tab4_Sku.Text)) = 0 Then
   msg_text = "請選擇欲刪除的客戶資料"
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

'檢核是否允許進行刪除旗標值
If blDelete = False Then
   msg_text = "客戶資料無法刪除：" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
Else
   Dim CheckDelete As Integer
   msg_text = "確定要刪除此筆客戶允收天數資料？"
   CheckDelete = MsgBox(msg_text, vbOKCancel + vbQuestion, msg_title)
End If

If CheckDelete = 1 Then
    '允許刪除
    str_SQL = "Delete From Acceptable Where Customer = '" & Trim(txt_Tab4_ConsigneeKey.Text) & "' and ItemNo='" & Trim(txt_Tab4_Sku.Text) & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
Else
    Screen.MousePointer = vbDefault
    Exit Sub
End If

fam_Tab4_Acceptable.BackColor = &H8000000C
fam_Tab4_Acceptable.Enabled = False
cmd_Tab4_Cancel.Enabled = False
cmd_Tab4_Save.Enabled = False
cmd_Tab4_AddNew.Enabled = True
cmd_Tab4_Modify.Enabled = False
cmd_Tab4_Delete.Enabled = False
txt_Tab4_ConsigneeKey.BackColor = &H8000000F
txt_Tab4_Sku.BackColor = &H8000000F

'重新顯示所有客戶允收天數資料
Call cmd_Tab4_AcceptableShow_Click

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-客戶允收天數資料-刪除", Me.Caption, "cmd_Tab4_Delete_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4_Modify_Click()
'客戶允收天數資料 >> 轉換修改模式
'確認選取客戶允收天數資料方允許 [修改] 功能
If rs_Tab4_AcceptableList Is Nothing Then Exit Sub
If dg_Tab4_AcceptableList.SelBookmarks.Count <> 0 Then
   fam_Tab4_Acceptable.BackColor = &HC0E0FF
   fam_Tab4_Acceptable.Enabled = True
   txt_Tab4_ConsigneeKey.Enabled = False
   txt_Tab4_Sku.Enabled = False
   cmd_Tab4_Save.Enabled = True
   cmd_Tab4_Cancel.Enabled = True
   cmd_Tab4_AddNew.Enabled = False
   cmd_Tab4_Modify.Enabled = False
   cmd_Tab4_Delete.Enabled = False
   txt_Tab4_ConsigneeKey.BackColor = &H8000000F
   txt_Tab4_Sku.BackColor = &H8000000F
End If
End Sub



Private Sub cmd_Tab4_Save_Click()
'客戶允收天數資料 >> 客戶允收天數資料存檔
On Error GoTo err_Handle

Dim rsTmp As New ADODB.Recordset

'客戶編號 檢查
If txt_Tab4_ConsigneeKey.Enabled = True And Len(Trim(txt_Tab4_Sku.Text)) <> 0 Then
    rsTmp.Open "select consigneekey,Rtrim(Isnull(Full_Name,'')) as full_name from trp01m where Left(rtrim(consigneekey),8) = '" & Trim(txt_Tab4_ConsigneeKey.Text) & "' and rtrim(storerkey) = '" & Left(cmb_Tab4_Storer.Text, InStr(cmb_Tab4_Storer.Text + " ", " ") - 1) & "' ", cn
    If rsTmp.EOF = True Then
       MsgBox "「客戶編號」不存在系統，請確認資料!!", 64, "注意"
           txt_Tab4_ConsigneeKey.SelStart = 0: txt_Tab4_ConsigneeKey.SelLength = Len(txt_Tab4_ConsigneeKey.Text)
       txt_Tab4_ConsigneeKey.SetFocus
       rsTmp.Close
       Exit Sub
    Else
       rsTmp.Close
    End If
End If

'產品編號 檢查
If txt_Tab4_Sku.Enabled = True And Len(Trim(txt_Tab4_Sku.Text)) <> 0 Then
    rsTmp.Open "select Sku,Rtrim(Isnull(DESCR,'')) as DESCR from " & strWMSDB & "..Sku where rtrim(Sku) = '" & Trim(txt_Tab4_Sku.Text) & "' and rtrim(storerkey) = '" & Left(cmb_Tab4_Storer.Text, InStr(cmb_Tab4_Storer.Text + " ", " ") - 1) & "' ", cn
    If rsTmp.EOF = True Then
       MsgBox "「產品編號」不存在系統，請確認資料!!", 64, "注意"
           txt_Tab4_Sku.SelStart = 0: txt_Tab4_Sku.SelLength = Len(txt_Tab4_Sku.Text)
       txt_Tab4_Sku.SetFocus
       rsTmp.Close
       Exit Sub
    Else
       rsTmp.Close
    End If
End If

'客戶編號+產品編號 重複檢查
If txt_Tab4_ConsigneeKey.Enabled = True And txt_Tab4_Sku.Enabled = True Then
    rsTmp.Open "select customer,itemno from Acceptable where rtrim(customer) = '" & Trim(txt_Tab4_ConsigneeKey.Text) & "' and rtrim(itemno)='" & Trim(txt_Tab4_Sku.Text) & "' and rtrim(storerkey) = '" & Left(cmb_Tab4_Storer.Text, InStr(cmb_Tab4_Storer.Text + " ", " ") - 1) & "' ", cn
    If rsTmp.EOF = False Then
        MsgBox "同一貨主，新增「客戶編號+產品編號」重複!!", 64, "注意"
           txt_Tab4_ConsigneeKey.SelStart = 0: txt_Tab4_ConsigneeKey.SelLength = Len(txt_Tab4_ConsigneeKey.Text)
       txt_Tab4_ConsigneeKey.SetFocus
       rsTmp.Close
       Exit Sub
    Else
        rsTmp.Close
    End If
End If

'存檔資料檢核
If Check_AcceptableData = False Then Exit Sub

Screen.MousePointer = vbHourglass
If Not (tmp_Cmd Is Nothing) Then
   Set tmp_Cmd = Nothing
End If
Set tmp_Cmd = New ADODB.Command
If tmp_para Is Nothing Then
   Set tmp_para = New ADODB.Parameter
End If

tmp_Cmd.ActiveConnection = cn
tmp_Cmd.CommandTimeout = 0    '執行時間設定：無限期等待
tmp_Cmd.CommandType = adCmdStoredProc
tmp_Cmd.CommandText = "Master_AcceptableData_Update"

'貨主
Set tmp_para = tmp_Cmd.CreateParameter("StorerKey", adChar, adParamInput, 15)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("StorerKey").Value = arStorer(cmb_Tab4_Storer.ListIndex)

'客戶編號
Set tmp_para = tmp_Cmd.CreateParameter("ConsigneeKey", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("ConsigneeKey").Value = Trim(txt_Tab4_ConsigneeKey.Text)

'產品編號
Set tmp_para = tmp_Cmd.CreateParameter("SKU", adChar, adParamInput, 20)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("SKU").Value = Trim(txt_Tab4_Sku.Text)

'允收天數
Set tmp_para = tmp_Cmd.CreateParameter("AllowDays", adInteger, adParamInput)
tmp_Cmd.Parameters.Append tmp_para
tmp_Cmd.Parameters("AllowDays").Value = Trim(txt_Tab4_AllowDays.Text)

Call Confirm_Recordset_Closed(tmp_Rs)
Call DB_CheckConnectStatus

'非同步執行
Set tmp_Rs = tmp_Cmd.Execute(, , adAsyncExecute)
Do While tmp_Cmd.State = adStateExecuting
   DoEvents: DoEvents  '讓 [執行中] 訊息視窗有 [更新] 時間
Loop
Set tmp_Cmd = Nothing

fam_Tab4_Acceptable.BackColor = &H8000000C
fam_Tab4_Acceptable.Enabled = False
cmd_Tab4_Cancel.Enabled = False
cmd_Tab4_Save.Enabled = False
cmd_Tab4_AddNew.Enabled = True
cmd_Tab4_Modify.Enabled = False
cmd_Tab4_Delete.Enabled = False
dg_Tab4_AcceptableList.Enabled = True
cmd_Tab4_AcceptableShow.Enabled = True
cmd_Tab4_AcceptableQuery.Enabled = True
cmd_Tab4_AcceptableReset.Enabled = True

'重新顯示所有客戶資料
Call cmd_Tab4_AcceptableShow_Click

Screen.MousePointer = vbDefault
Exit Sub

err_Handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-客戶允收天數資料-存檔", Me.Caption, "cmd_Tab4_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab4_SaveToExcel_Click()
'查詢結果>> 轉 EXCEL
If blTab4AcceptableEventEnable = False Then
    msg_text = "無資料不能轉Excel喔！"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Exit Sub
End If
    
    If rs_Tab4_AcceptableList Is Nothing Then Exit Sub
    If rs_Tab4_AcceptableList.RecordCount = 0 Then Exit Sub
    
blTab4AcceptableEventEnable = False '避免每行資料選取

    rs_Tab4_AcceptableList.MoveFirst
    
    Recordset2ExcelV2 "允收天數資料", "允收天數資料", rs_Tab4_AcceptableList
    
    Set MyXlsAppV2 = Nothing
    
blTab4AcceptableEventEnable = True
End Sub

Private Sub cmdOpenFilesT5_Click()

On Error GoTo err_Handle
gd_Tab5_AaccessTable.Enabled = False

Dim str As String, strFieldName As String, strFilePath As String, strSheetName As String, str_storekey As String, str_soldcode As String, Str_Sku As String, str_allowdays As Integer
'確認路徑是否帶"\"
If Right(filLocalFileT5.Path, 1) = "\" Then
    strFilePath = filLocalFileT5.Path
Else
    strFilePath = filLocalFileT5.Path & "\"
End If
'建立欄位名稱陣列
strFieldName = ""
If Right(filLocalFileT5.Path, 1) <> "\" Then
    strFilePath = filLocalFileT5.Path & "\"
Else
    strFilePath = filLocalFileT5.Path
End If
Set rsMain = New ADODB.Recordset
strSheetName = "DATA"
Call Excel2Recordset(strFilePath & filLocalFileT5.FileName, strSheetName, strFieldName, rsMain)
rsMain.MoveFirst
Set gd_Tab5_AaccessTable.DataSource = rsMain

'若AcceptableTemp不為空table刪除之
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "if object_id ('AcceptableTemp') is not null drop TABLE  AcceptableTemp"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If rsMain.EOF = False Then
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "CREATE TABLE AcceptableTemp ( ItemNo char(20) , Customer varchar(20) , AllowDays int, Storerkey char(15) );"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    End If
    Do While Not rsMain.EOF
        gd_Tab5_AaccessTable.Col = 0: str_storekey = Trim(gd_Tab5_AaccessTable.Text)
        gd_Tab5_AaccessTable.Col = 1: str_soldcode = Trim(gd_Tab5_AaccessTable.Text)
        gd_Tab5_AaccessTable.Col = 2: Str_Sku = Trim(gd_Tab5_AaccessTable.Text)
        gd_Tab5_AaccessTable.Col = 3: str_allowdays = Trim(gd_Tab5_AaccessTable.Text)
        str_SQL = "INSERT INTO AcceptableTemp (storerkey,Customer,ItemNo,allowdays) VALUES ('" & str_storekey & "','" & str_soldcode & "','" & Str_Sku & "','" & str_allowdays & "')"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        rsMain.MoveNext
Loop
rsMain.MoveFirst

'

If rsMain Is Nothing Then
    MsgBox "查無資料!", 64, "Excel2Recordset"
Else
    SetDataGridColWidth Me.Caption, gd_Tab5_AaccessTable
    MsgBox "此工作表共 " & rsMain.RecordCount & "筆資料，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    cmdImportT5.Enabled = True
End If

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub
Private Sub cmdImportT5_Click()

    On Error GoTo err_Handle
    gd_Tab5_AaccessTable.Enabled = False: cmdImportT5.Enabled = False
    Dim existmark As Integer
    
    Call Confirm_Recordset_Closed(tmp_Rs)
'    '檢查貨主是否是LCHF01
'    str_SQL = "select * from AcceptableTemp where Storerkey <> 'LCHF01'"
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If Not tmp_Rs.EOF Then
'       msg_text = "貨主不為LCHF01:" & Trim(tmp_Rs.Fields("Customer").Value) & "," & Trim(tmp_Rs.Fields("ItemNo").Value) & "，請確認資料，謝謝。"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
'        tmp_Rs.Close
'        Exit Sub
'    End If
    
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    '檢查客戶編號是否存在
    str_SQL = "select  customer,* from AcceptableTemp where customer NOT in (select  distinct consigneekey from trp01m where storerkey='LCHF01')"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
       msg_text = "有客戶編號(soldcode):" & Trim(tmp_Rs.Fields("Customer").Value) & "不存在，請確認資料，謝謝。"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
        tmp_Rs.Close
        Exit Sub
    End If

    Call Confirm_Recordset_Closed(tmp_Rs)
    '檢查sku是否存在
    str_SQL = "select  ItemNo,* from AcceptableTemp where  ItemNo NOT in (select  distinct sku from " & strWMSDB & "..sku where storerkey='LCHF01')"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
       msg_text = "有品號(sku):" & Trim(tmp_Rs.Fields("ItemNo").Value) & "不存在，請確認資料，謝謝。"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
        tmp_Rs.Close
        Exit Sub
    End If

'    Call Confirm_Recordset_Closed(tmp_Rs)
'    '檢查是否有allowdays<=0
'    str_SQL = "select * from AcceptableTemp where allowdays=0"
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If Not tmp_Rs.EOF Then
'       msg_text = "有日期為0:" & Trim(tmp_Rs.Fields("Customer").Value) & "," & Trim(tmp_Rs.Fields("ItemNo").Value) & "，請確認資料，謝謝。"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
'        tmp_Rs.Close
'        Exit Sub
'    End If
    

    Call Confirm_Recordset_Closed(tmp_Rs)
    '檢查是否有重覆soldcode+sku
    str_SQL = "select * from AcceptableTemp where ItemNo+Customer in (select ItemNo+Customer from AcceptableTemp group by Customer,ItemNo having count(*)>1) order by ItemNo+Customer "
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        msg_text = "有重覆的SoldCode+SKU:" & Trim(tmp_Rs.Fields("Customer").Value) & "," & Trim(tmp_Rs.Fields("ItemNo").Value) & "，請確認資料，謝謝。"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
        tmp_Rs.Close
        Exit Sub
    End If
'    Call Confirm_Recordset_Closed(tmp_rs)
'      str_SQL = "select * from  Acceptable where RTRim(Customer)+Rtrim(ItemNo) in (select Rtrim(u.customer)+Rtrim(u.ItemNo) from AcceptableTemp u inner join Acceptable a on u.customer=a.Customer and u.ItemNo=a.ItemNo)"
'    tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If Not tmp_rs.EOF Then
'        msg_text = "有重覆的SoldCode+SKU2，請確認資料，謝謝。"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
'        tmp_rs.Close
'        Exit Sub
'    End If
    existmark = 0
    '檢查是否有已存在Acceptable的資料
    Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select a.*,u.* from AcceptableTemp u inner join Acceptable a on u.customer=a.Customer and u.ItemNo=a.ItemNo"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If Not tmp_Rs.EOF Then
        existmark = 1
        tmp_Rs.Close
        'Exit Sub
    End If


        Tran_Level = cn.BeginTrans
        If existmark = 1 Then
            str_SQL = "delete Acceptable where RTRim(Customer)+Rtrim(ItemNo) in (select Rtrim(u.Customer)+Rtrim(u.ItemNo) from AcceptableTemp u inner join Acceptable a on u.Customer=a.Customer and u.ItemNo=a.ItemNo)"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        End If
        str_SQL = "Insert Acceptable(Storerkey, Customer, ItemNo, AllowDays,addwho,adddate,editwho,editdate) select Storerkey,Customer,ItemNo,AllowDays,'" & User_id & "' , getdate() , '" & User_id & "' , getdate() from AcceptableTemp  where  ItemNo in (select  distinct sku from " & strWMSDB & "..sku where storerkey='LCHF01')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        cn.CommitTrans: Tran_Level = 0
        '備份檔案
        Dim fl_file As Scripting.File
        Set fso = New FileSystemObject
        Dim strExcelFileName As String, str_ArPath As String, str_file As String, arrstr_file
        strExcelFileName = filLocalFileT5.Path & "\" & filLocalFileT5.FileName
        If fso.FileExists(strExcelFileName) = True Then
        str_ArPath = "D:\LCHF01\AcceptTable\"
        str_file = filLocalFileT5.FileName
            If Dir(str_ArPath, vbDirectory) = "" Then MkDirs str_ArPath
            Set fl_file = fso.GetFile(strExcelFileName) '原檔案路徑
            arrstr_file = Split(str_file, ".")
            fl_file.copy (str_ArPath & arrstr_file(0) & "_" & Format(Now, "YYMMDDHHMMSS") & ".xls")

            If fso.FileExists(str_ArPath & arrstr_file(0) & "_" & Format(Now, "YYMMDDHHMMSS") & ".xls") = True Then
                    fl_file.Delete
            End If

        End If
        Set rsMain = Nothing
        filLocalFileT5.Refresh
        msg_text = "匯入成功，檔案備份於D:\LCHF01\AcceptTable。"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
    Exit Sub

err_Handle:
    Dim tmpString As String
    gd_Tab5_AaccessTable.Enabled = True: cmdImportT5.Enabled = True
    If Tran_Level = 1 Then cn.RollbackTrans
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "允收天數匯入-存檔", Me.Caption, "cmd_Tab0_Save_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault

End Sub
Private Sub Display_SelectedAcceptableData(ByVal strConsigneeKey As String, strSku As String)
'顯示傳入之客戶允收天數資料
Call Clear_AcceptableData

str_SQL = "select distinct Rtrim(a.storerkey) as 貨主,Rtrim(a.Customer) as 客戶編號" & _
",(select top 1 Rtrim(t1m.Full_Name) from TRP01M t1m where Left(t1m.ConsigneeKey,8)=a.customer order by  Rtrim(t1m.Full_Name) desc)  as 客戶名稱,Rtrim(a.ItemNo) as 產品編號,Rtrim(s.descr) as 產品名稱,a.allowdays as 允收天數 " & _
" from Acceptable a " & _
"inner join " & strWMSDB & "..sku s on s.sku=a.itemno " & _
"Where Customer = '" & strConsigneeKey & "' and ItemNo='" & strSku & "'" & _
" order by Rtrim(a.Customer),Rtrim(a.ItemNo) "

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
cn.CommandTimeout = 0   '無限期等待
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "查詢結果：無符合設定條件之客戶允收天數基本資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

Dim i As Double
For i = 0 To cmb_Tab4_Storer.ListCount - 1
    If arStorer(i) = Trim(tmp_Rs.Fields("貨主").Value) Then
       cmb_Tab4_Storer.ListIndex = i
       Exit For
    End If
Next i

txt_Tab4_ConsigneeKey.Text = Trim(tmp_Rs.Fields("客戶編號").Value)
txt_Tab4_FullName.Text = IIf(IsNull(tmp_Rs.Fields("客戶名稱").Value), "", Trim(tmp_Rs.Fields("客戶名稱").Value))
txt_Tab4_Sku.Text = IIf(IsNull(tmp_Rs.Fields("產品編號").Value), "", Trim(tmp_Rs.Fields("產品編號").Value))
txt_Tab4_DESCR.Text = IIf(IsNull(tmp_Rs.Fields("產品名稱").Value), "", Trim(tmp_Rs.Fields("產品名稱").Value))
txt_Tab4_AllowDays.Text = IIf(IsNull(tmp_Rs.Fields("允收天數").Value), "", Trim(tmp_Rs.Fields("允收天數").Value))

tmp_Rs.Close

End Sub

Private Sub Clear_AcceptableData()
'清除 客戶允收天數資料 畫面之欄位值
cmb_Tab4_Storer.ListIndex = -1
txt_Tab4_ConsigneeKey.Text = ""
txt_Tab4_FullName.Text = ""
txt_Tab4_Sku.Text = ""
txt_Tab4_DESCR.Text = ""
txt_Tab4_AllowDays.Text = ""
txt_Tab4_ConsigneeKey.BackColor = &H8000000F
txt_Tab4_Sku.BackColor = &H8000000F
End Sub
Private Function Check_AcceptableData() As Boolean
'客戶允收天數基本資料檢核
Check_AcceptableData = False
msg_text = ""
If cmb_Tab4_Storer.ListIndex = -1 Then
   If msg_text = "" Then
      msg_text = "未輸入 [貨主]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入 [貨主]"
   End If
End If
If Len(Trim(txt_Tab4_ConsigneeKey.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入 [客戶編號]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入 [客戶編號]"
   End If
End If
If Len(Trim(txt_Tab4_Sku.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入 [產品編號]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入 [產品編號]"
   End If
End If
If Len(Trim(txt_Tab4_AllowDays.Text)) = 0 Then
   If msg_text = "" Then
      msg_text = "未輸入 [允收天數]"
   Else
      msg_text = msg_text & vbCrLf & "未輸入 [允收天數]"
   End If
End If

If msg_text = "" Then
   Check_AcceptableData = True
Else
   msg_text = "客戶允收天數資料異常，請修正後再執行 [存 檔]：" & vbCrLf & msg_text
   MsgBox msg_text, vbOKOnly + vbCritical, msg_title
   Exit Function
End If
End Function
Private Sub txt_Tab4_ConsigneeKey_LostFocus()
'客戶編號 檢查
If txt_Tab4_ConsigneeKey.Enabled = True And Len(Trim(txt_Tab4_ConsigneeKey.Text)) <> 0 Then
    Dim rsTmp As New ADODB.Recordset
    rsTmp.Open "select top 1 consigneekey,Rtrim(Isnull(Full_Name,'')) as full_name from trp01m where Left(rtrim(consigneekey),8) = '" & RTrim(txt_Tab4_ConsigneeKey.Text) & "' and rtrim(storerkey) = '" & Left(cmb_Tab4_Storer.Text, InStr(cmb_Tab4_Storer.Text + " ", " ") - 1) & "' order by Rtrim(Isnull(Full_Name,'')) desc", cn
    If rsTmp.EOF = True Then
       MsgBox "「客戶編號」不存在系統，請確認資料!!", 64, "注意"
           txt_Tab4_ConsigneeKey.SelStart = 0: txt_Tab4_ConsigneeKey.SelLength = Len(txt_Tab4_ConsigneeKey.Text)
       txt_Tab4_FullName.Text = ""
       txt_Tab4_ConsigneeKey.SetFocus
       Exit Sub
    Else
       txt_Tab4_FullName.Text = rsTmp.Fields("full_name").Value
    End If
rsTmp.Close
End If
End Sub
Private Sub dirLocalDirT5_Change()
    filLocalFileT5.Path = dirLocalDirT5.Path
End Sub

Private Sub txt_Tab4_Sku_LostFocus()
'產品編號 檢查
If txt_Tab4_Sku.Enabled = True And Len(Trim(txt_Tab4_Sku.Text)) <> 0 Then
    Dim rsTmp As New ADODB.Recordset
    rsTmp.Open "select Sku,Rtrim(Isnull(DESCR,'')) as DESCR from " & strWMSDB & "..Sku where rtrim(Sku) = '" & RTrim(txt_Tab4_Sku.Text) & "' and rtrim(storerkey) = '" & Left(cmb_Tab4_Storer.Text, InStr(cmb_Tab4_Storer.Text + " ", " ") - 1) & "' ", cn
    If rsTmp.EOF = True Then
       MsgBox "「產品編號」不存在系統，請確認資料!!", 64, "注意"
           txt_Tab4_Sku.SelStart = 0: txt_Tab4_Sku.SelLength = Len(txt_Tab4_Sku.Text)
       txt_Tab4_DESCR.Text = ""
       txt_Tab4_Sku.SetFocus
       Exit Sub
    Else
       txt_Tab4_DESCR.Text = rsTmp.Fields("DESCR").Value
    End If
rsTmp.Close
End If
End Sub

Private Sub txt_Tab4_AllowDays_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyBack        '0 - 9,BACKSPACE處理
'        Case vbKeyDelete, vbKeyDecimal          '小數點處理
'            If InStr(1, txt_Tab4_AllowDays.Text, ".") <> 0 Then KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Sub Recordset2ExcelV2(str As String, FileName As String, rs As Object)
On Error GoTo err_Handle
If rs Is Nothing Then MsgBox "無資料可供轉檔！", vbOKOnly + vbInformation, "Save2Excel": Exit Sub
Screen.MousePointer = 11
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String
Dim tmp_col As Double, tmp_row As Double
Dim tmp_letter As String, tmp_RangNo As String, tmpI As Integer

    Dim ExcelTitle As String, saved As Integer, msg_text As String
    msg_text = ""
    saved = 0
    Call DocStoreDirectory(strDocPath)

    Dim strTranFileName As String           'Excel 檔案名稱
    CmnDialog.DialogTitle = "轉存 Excel 檔"
    CmnDialog.InitDir = "C:\my documents"
    CmnDialog.FileName = FileName & "-" & Format(Now, "YYYYMMDDHHNNSS")
    CmnDialog.Filter = "Excel檔案(*.xls)|*.xls"
    CmnDialog.FilterIndex = 1
    CmnDialog.CancelError = True
    CmnDialog.Flags = cdlOFNHideReadOnly    '隱藏唯讀核取方塊
    On Error Resume Next
    CmnDialog.ShowSave

    If err.Number = cdlCancel Then          '於 [存檔] 對話方塊中，按下 [取消] 鈕
       msg_text = "您選擇 [取消] 按鈕，必須於 Excel 中自行存檔！"
       MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
       strTranFileName = ""
    Else
       strTranFileName = CmnDialog.FileName
       If Dir(strTranFileName) <> "" Then
          Kill strTranFileName
       End If
    End If

Screen.MousePointer = 11
'開啟EXCEL物件
Set MyXlsAppV2 = CreateObject("Excel.Application")

With MyXlsAppV2
    .Visible = False
    
    If Dir(App.Path & "\XLT\" & str & ".xlt") = "" Then '找不到本機範例檔
        
        '取範例檔路徑
        Dim objIni As vbIniFile, arrTmp, strXltPath As String
        Set objIni = New vbIniFile
        
        With objIni
        
            .FileName = striniFileName_FullPath
            strXltPath = RTrim(.ReadData("EXCEL", "XLTPATH", "")) '不支援中文資料夾名稱
            
        End With
        Set objIni = Nothing

    End If

    '無指定路徑不使用範例檔
    If Len(RTrim(strXltPath)) = 0 Then strXltPath = App.Path & "\XLT"
    If Dir(strXltPath, vbDirectory) = "" Then GoTo Run
    
    '範例檔
    If Dir(strXltPath & "\" & str & ".xlt") <> "" Then
'        If MsgBox("是否使用範例檔?(" & strXltPath & "\" & str & ".xlt), vbQuestion + vbYesNo, "轉Excel") = vbNo Then GoTo Run
        
        '開啟範例檔
        .Workbooks.Open (strXltPath & "\" & str & ".xlt")
        
        '尋找DATA工作表
        For i = 1 To .Sheets.Count
            If UCase(RTrim(.Sheets(i).Name)) = "DATA" Then .Sheets("Data").Select: Exit For '選定DATA工作表
        Next
        
        '找不到新增DATA工作表
        If UCase(RTrim(.ActiveSheet.Name)) <> "DATA" Then .Sheets.Add: .ActiveSheet.Name = "DATA":
        
        '搜尋存放儲存格
        For k = 65 To 66 '90
            For j = 1 To 100
                tmp_row = j
                If UCase(.Range(Chr(k) & j).Value) = "BESTLOG" Then GoTo NextStep
                
            Next j
        Next k
        k = 65: j = 1

        '寫入標題列
        For i = 0 To rs.Fields.Count - 1
            l = i Mod 26
            .Range(strCol & Chr(k + l) & j).Value = rs.Fields(i).Name
            '欄位超過26
            If Chr(65 + l) = "Z" Then
                If strCol = "" Then
                    strCol = "A"
                Else
                    strCol = Chr(Asc(strCol) + 1)
                End If
            End If
        Next i
        
NextStep:

        '資料寫入
        '.ActiveSheet.Cells(2, 1).CopyFromRecordset rs
        '.Range(Chr(k) & j).CopyFromRecordset rs
        
'        tmp_row = 2
        Do While Not rs.EOF
            DoEvents
            '判斷使用者是否取消轉檔作業
            For tmp_col = 0 To rs.Fields.Count - 1
                tmp_letter = Chr(65 + tmp_col)      ' A 之 ascii code
                If Asc(tmp_letter) > 90 Then        ' > Z 則變成 AA 起始
                   tmp_letter = "A" & Chr(Asc(tmp_letter) - 90 + 64)
                End If
                tmp_RangNo = tmp_letter & (tmp_row)
                '設定格式
'                With excelAP.Range(tmp_RangNo)
'                    .NumberFormatLocal = "@"      '儲存格格式 >> 數字 >> 類別 = 文字
'                    '.Font.Name = "新細明體"       '儲存格格式 >> 字型 >> 字型 = Times New Roman
'                    '.Font.FontStyle = "標準"      '儲存格格式 >> 字型 >> 外型樣式 = 標準
'                    '.Font.Size = 12               '儲存格格式 >> 字型 >> 大小 = 12
'                End With
                .Range(tmp_RangNo) = Trim(rs.Fields(tmp_col).Value)
            Next tmp_col
            rs.MoveNext
            tmp_row = tmp_row + 1
        Loop
        
        
    Else '不使用範例檔

Run:
        '新增Excel
        .Workbooks.Add: .Sheets("Sheet1").Select: .Sheets("Sheet1").Name = str
        
        '寫入標題列
        For i = 0 To rs.Fields.Count - 1
            l = i Mod 26
            .Range(strCol & Chr(65 + l) & "1").Value = rs.Fields(i).Name
            '欄位超過26
            If Chr(65 + l) = "Z" Then
                If strCol = "" Then
                    strCol = "A"
                Else
                    strCol = Chr(Asc(strCol) + 1)
                End If
            End If
        Next i
        
        '資料寫入
        '.Range("A2").CopyFromRecordset rs
        '.ActiveSheet.Cells(2, 1).CopyFromRecordset rs

        tmp_row = 2
        Do While Not rs.EOF
            DoEvents
            '判斷使用者是否取消轉檔作業
            For tmp_col = 0 To rs.Fields.Count - 1
                tmp_letter = Chr(65 + tmp_col)      ' A 之 ascii code
                If Asc(tmp_letter) > 90 Then        ' > Z 則變成 AA 起始
                   tmp_letter = "A" & Chr(Asc(tmp_letter) - 90 + 64)
                End If
                tmp_RangNo = tmp_letter & (tmp_row)
                '設定格式
'                With excelAP.Range(tmp_RangNo)
'                    .NumberFormatLocal = "@"      '儲存格格式 >> 數字 >> 類別 = 文字
'                    '.Font.Name = "新細明體"       '儲存格格式 >> 字型 >> 字型 = Times New Roman
'                    '.Font.FontStyle = "標準"      '儲存格格式 >> 字型 >> 外型樣式 = 標準
'                    '.Font.Size = 12               '儲存格格式 >> 字型 >> 大小 = 12
'                End With
                .Range(tmp_RangNo) = Trim(rs.Fields(tmp_col).Value)
            Next tmp_col
            rs.MoveNext
            tmp_row = tmp_row + 1
        Loop
    
    End If
      
        If Len(RTrim(strTranFileName)) > 0 Then
           .ActiveWorkbook.Author = User_id
           .ActiveWorkbook.SaveAs FileName:=strTranFileName, FileFormat:=xlNormal
           .ActiveWindow.Close
           .Visible = False
        Else
           .ActiveWorkbook.Author = User_id
           .Visible = True
        End If

    saved = 1
    
End With

    If Len(RTrim(strTranFileName)) = 0 And saved = 1 Then
       Screen.MousePointer = vbDefault
    ElseIf Len(RTrim(strTranFileName)) > 0 And saved = 1 Then
       Screen.MousePointer = vbDefault
       msg_text = "轉存作業完成，檔案存放位置：" & strTranFileName
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    End If
        
Exit Sub

err_Handle:
   If err.Number = 0 Then Exit Sub
   Dim tmpString As String
   Screen.MousePointer = vbDefault
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "--允收天數轉EXCEL", Me.Caption, "cmd_SaveToExcel_Tab4", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub dg_Tab4_AcceptableList_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'客戶允收天數資料列表：整行選取
If blTab4AcceptableEventEnable Then
   If Not rs_Tab4_AcceptableList.EOF Then
      dg_Tab4_AcceptableList.SelBookmarks.Add rs_Tab4_AcceptableList.Bookmark
      Call Display_SelectedAcceptableData(rs_Tab4_AcceptableList.Fields("客戶編號").Value, rs_Tab4_AcceptableList.Fields("產品編號").Value)
      fam_Tab4_Acceptable.BackColor = &H8000000C
      fam_Tab4_Acceptable.Enabled = False
      cmd_Tab4_Cancel.Enabled = False
      cmd_Tab4_Save.Enabled = False
      cmd_Tab4_AddNew.Enabled = True
      cmd_Tab4_Modify.Enabled = True
      cmd_Tab4_Delete.Enabled = True
   End If
End If
End Sub

Private Sub cmd_Tab4_AcceptableQuery_Click()
'客戶允收天數資料 >> 允收資料搜尋
If rs_Tab4_AcceptableList Is Nothing Then Exit Sub
If rs_Tab4_AcceptableList.RecordCount = 0 Then Exit Sub

strFormName_FilterAndSort = Me.Name
strRSName_FilterAndSort = "rs_Tab4_AcceptableList"

If ShowForm_RS_FilterAndSort(rs_Tab4_AcceptableList, "客戶允收天數資料", Me.Tag) = False Then
    MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
    Exit Sub
End If
Me.WindowState = vbNormal
End Sub


