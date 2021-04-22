VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_FTP 
   Caption         =   "訂單接收"
   ClientHeight    =   8775
   ClientLeft      =   210
   ClientTop       =   750
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8775
   ScaleWidth      =   11280
   WindowState     =   2  '最大化
   Begin InetCtlsObjects.Inet ITC 
      Left            =   240
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8160
      Left            =   120
      TabIndex        =   0
      Top             =   15
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   14393
      _Version        =   393216
      Tabs            =   23
      Tab             =   18
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frm_FTP.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmd_Import"
      Tab(0).Control(1)=   "fraControls"
      Tab(0).Control(2)=   "fraRemoteFiles"
      Tab(0).Control(3)=   "fraLoginInfo"
      Tab(0).Control(4)=   "fraStatus"
      Tab(0).Control(5)=   "fraLocalFiles"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frm_FTP.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dg_CustInv"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Vitalon訂單匯入"
      TabPicture(2)   =   "frm_FTP.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "dgMainT2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   " "
      TabPicture(3)   =   "frm_FTP.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).Control(1)=   "int3"
      Tab(3).Control(2)=   "dgMainT3"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "LKAO01"
      TabPicture(4)   =   "frm_FTP.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "dgMainT4"
      Tab(4).Control(1)=   "Frame7"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   " "
      TabPicture(5)   =   "frm_FTP.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "dgMainT5"
      Tab(5).Control(1)=   "Frame8"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "朝日訂單匯入"
      TabPicture(6)   =   "frm_FTP.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "dgMainT6"
      Tab(6).Control(1)=   "Frame9"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "百事I訂單匯入"
      TabPicture(7)   =   "frm_FTP.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "dgMainT7"
      Tab(7).Control(1)=   "Frame10"
      Tab(7).ControlCount=   2
      TabCaption(8)   =   " "
      TabPicture(8)   =   "frm_FTP.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame11"
      Tab(8).Control(1)=   "dgMainT8"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   " "
      TabPicture(9)   =   "frm_FTP.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame12"
      Tab(9).Control(1)=   "dgMainT9"
      Tab(9).ControlCount=   2
      TabCaption(10)  =   " "
      TabPicture(10)  =   "frm_FTP.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame13"
      Tab(10).Control(1)=   "dgMainT10"
      Tab(10).ControlCount=   2
      TabCaption(11)  =   " "
      TabPicture(11)  =   "frm_FTP.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "dgMainT11"
      Tab(11).Control(1)=   "Frame3"
      Tab(11).ControlCount=   2
      TabCaption(12)  =   "百事RC訂單匯入"
      TabPicture(12)  =   "frm_FTP.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "dgMainT12"
      Tab(12).Control(1)=   "Frame4"
      Tab(12).ControlCount=   2
      TabCaption(13)  =   "--其他貨主訂單"
      TabPicture(13)  =   "frm_FTP.frx":016C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "dgMainT13"
      Tab(13).Control(1)=   "Frame5"
      Tab(13).ControlCount=   2
      TabCaption(14)  =   "Excel訂單匯入"
      TabPicture(14)  =   "frm_FTP.frx":0188
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "dgMainT14"
      Tab(14).Control(1)=   "Frame6"
      Tab(14).ControlCount=   2
      TabCaption(15)  =   "中祥訂單匯入"
      TabPicture(15)  =   "frm_FTP.frx":01A4
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "dgMainT15"
      Tab(15).Control(1)=   "Frame14"
      Tab(15).ControlCount=   2
      TabCaption(16)  =   "毛寶訂單匯入"
      TabPicture(16)  =   "frm_FTP.frx":01C0
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "Frame15"
      Tab(16).Control(1)=   "SSTab2"
      Tab(16).ControlCount=   2
      TabCaption(17)  =   " "
      TabPicture(17)  =   "frm_FTP.frx":01DC
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "Frame16"
      Tab(17).Control(1)=   "SSTab3"
      Tab(17).ControlCount=   2
      TabCaption(18)  =   "百事A2B、C單匯入"
      TabPicture(18)  =   "frm_FTP.frx":01F8
      Tab(18).ControlEnabled=   -1  'True
      Tab(18).Control(0)=   "dgMainT18"
      Tab(18).Control(0).Enabled=   0   'False
      Tab(18).Control(1)=   "Frame17"
      Tab(18).Control(1).Enabled=   0   'False
      Tab(18).ControlCount=   2
      TabCaption(19)  =   " "
      TabPicture(19)  =   "frm_FTP.frx":0214
      Tab(19).ControlEnabled=   0   'False
      Tab(19).Control(0)=   "Frame18"
      Tab(19).Control(1)=   "dgMainT19"
      Tab(19).ControlCount=   2
      TabCaption(20)  =   " 特力屋訂單匯入"
      TabPicture(20)  =   "frm_FTP.frx":0230
      Tab(20).ControlEnabled=   0   'False
      Tab(20).Control(0)=   "Frame19"
      Tab(20).Control(1)=   "dgMainT20"
      Tab(20).ControlCount=   2
      TabCaption(21)  =   "中祥RC訂單匯入"
      TabPicture(21)  =   "frm_FTP.frx":024C
      Tab(21).ControlEnabled=   0   'False
      Tab(21).Control(0)=   "Frame20"
      Tab(21).Control(1)=   "dgMainT21"
      Tab(21).ControlCount=   2
      TabCaption(22)  =   " "
      TabPicture(22)  =   "frm_FTP.frx":0268
      Tab(22).ControlEnabled=   0   'False
      Tab(22).Control(0)=   "Frame21"
      Tab(22).Control(1)=   "dgMainT22"
      Tab(22).ControlCount=   2
      Begin VB.Frame Frame21 
         Caption         =   "LYFY09退貨訂單匯入"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   194
         Top             =   1320
         Width           =   9840
         Begin VB.CommandButton cmdImportT22 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   199
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT22 
            Height          =   300
            Left            =   135
            TabIndex        =   198
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT22 
            Height          =   1560
            Left            =   135
            TabIndex        =   197
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT22 
            Height          =   1530
            Left            =   4560
            Pattern         =   "PG退貨*.xls"
            TabIndex        =   196
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT22 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   195
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   200
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "中祥RC訂單匯入"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   190
         Top             =   1320
         Width           =   9840
         Begin VB.ComboBox cboSheetT21 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   210
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT21 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   209
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT21 
            Height          =   1560
            Left            =   120
            TabIndex        =   208
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT21 
            Height          =   300
            Left            =   120
            TabIndex        =   207
            ToolTipText     =   "Local Drive List"
            Top             =   360
            Width           =   2040
         End
         Begin VB.CommandButton cmdOpenFilesT21 
            BackColor       =   &H0080FFFF&
            Caption         =   "開啟"
            Height          =   375
            Left            =   3480
            Style           =   1  '圖片外觀
            TabIndex        =   206
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdImportT21 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   2400
            Style           =   1  '圖片外觀
            TabIndex        =   191
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   193
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "特力屋訂單匯入"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   186
         Top             =   1320
         Width           =   9840
         Begin VB.CommandButton cmdOpenFilesT20 
            BackColor       =   &H0080FFFF&
            Caption         =   "開啟"
            Height          =   375
            Left            =   3360
            Style           =   1  '圖片外觀
            TabIndex        =   215
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cboSheetT20 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5280
            Style           =   2  '單純下拉式
            TabIndex        =   214
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT20 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   213
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT20 
            Height          =   1560
            Left            =   120
            TabIndex        =   212
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.CommandButton cmdImportT20 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   2280
            Style           =   1  '圖片外觀
            TabIndex        =   188
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT20 
            Height          =   300
            Left            =   135
            TabIndex        =   187
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4440
            TabIndex        =   211
            Top             =   300
            Width           =   720
         End
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   181
         Top             =   3720
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7646
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "訂單資料(Format)"
         TabPicture(0)   =   "frm_FTP.frx":0284
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "dgMainT17"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "出貨備註(RawHerder)"
         TabPicture(1)   =   "frm_FTP.frx":02A0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dgMainT17_1"
         Tab(1).ControlCount=   1
         Begin MSDataGridLib.DataGrid dgMainT17 
            Height          =   3855
            Left            =   120
            TabIndex        =   182
            Top             =   360
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
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
         Begin MSDataGridLib.DataGrid dgMainT17_1 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   183
            Top             =   360
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   6800
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
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
      Begin VB.Frame Frame18 
         Caption         =   "LAPP01-退貨"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   172
         Top             =   1320
         Width           =   9840
         Begin VB.CommandButton cmdImportT19 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   177
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT19 
            Height          =   300
            Left            =   135
            TabIndex        =   176
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT19 
            Height          =   1560
            Left            =   135
            TabIndex        =   175
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT19 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   174
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT19 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   173
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   178
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "百事A2B訂單匯入"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   120
         TabIndex        =   164
         Top             =   1320
         Width           =   9840
         Begin VB.ComboBox cboSheetT18 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   169
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT18 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   168
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT18 
            Height          =   1560
            Left            =   135
            TabIndex        =   167
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT18 
            Height          =   300
            Left            =   135
            TabIndex        =   166
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT18 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   165
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   170
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "LAPP01-訂單匯入"
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
         Left            =   -74760
         TabIndex        =   158
         Top             =   1320
         Width           =   9840
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "轉Excel"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2280
            Style           =   1  '圖片外觀
            TabIndex        =   202
            Top             =   240
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.ComboBox cboSheetT17 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   184
            Top             =   240
            Visible         =   0   'False
            Width           =   4365
         End
         Begin VB.CommandButton CmbStartT17 
            Caption         =   "開啟檔案"
            Height          =   375
            Left            =   4560
            TabIndex        =   180
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdImportT17 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   162
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT17 
            Height          =   300
            Left            =   135
            TabIndex        =   161
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT17 
            Height          =   1560
            Left            =   135
            TabIndex        =   160
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT17 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   159
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   185
            Top             =   300
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   163
            Top             =   300
            Width           =   720
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   149
         Top             =   3780
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7646
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "訂單主檔"
         TabPicture(0)   =   "frm_FTP.frx":02BC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "dgMainT16"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "訂單明細檔"
         TabPicture(1)   =   "frm_FTP.frx":02D8
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dgMainT16_1"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   " "
         TabPicture(2)   =   "frm_FTP.frx":02F4
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "dgMainT16_2"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   " "
         TabPicture(3)   =   "frm_FTP.frx":0310
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "dgMainT16_3"
         Tab(3).ControlCount=   1
         Begin MSDataGridLib.DataGrid dgMainT16 
            Height          =   3975
            Left            =   120
            TabIndex        =   150
            Top             =   360
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
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
         Begin MSDataGridLib.DataGrid dgMainT16_1 
            Height          =   3975
            Left            =   -74880
            TabIndex        =   151
            Top             =   360
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
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
         Begin MSDataGridLib.DataGrid dgMainT16_2 
            Height          =   3975
            Left            =   -74880
            TabIndex        =   152
            Top             =   360
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
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
         Begin MSDataGridLib.DataGrid dgMainT16_3 
            Height          =   3975
            Left            =   -74880
            TabIndex        =   153
            Top             =   360
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   7011
            _Version        =   393216
            AllowUpdate     =   0   'False
            Enabled         =   -1  'True
            HeadLines       =   2
            RowHeight       =   20
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "新細明體"
               Size            =   9
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
      Begin VB.Frame Frame15 
         Caption         =   "LMBO01-訂單匯入"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74760
         TabIndex        =   144
         Top             =   1260
         Width           =   9840
         Begin VB.CommandButton Command5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "轉Excel"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2280
            Style           =   1  '圖片外觀
            TabIndex        =   205
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton Cmb_Import2 
            Caption         =   "銷貨匯入"
            Height          =   375
            Left            =   5760
            TabIndex        =   157
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Cmb_Import4 
            Caption         =   "寄庫銷貨匯入"
            Height          =   375
            Left            =   8160
            TabIndex        =   156
            Top             =   240
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Cmb_Import3 
            Caption         =   "轉撥匯入"
            Height          =   375
            Left            =   6960
            TabIndex        =   155
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton Cmb_Import1 
            Caption         =   "開啟檔案"
            Height          =   375
            Left            =   4560
            TabIndex        =   154
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Cmd_Impdata 
            BackColor       =   &H0080FFFF&
            Caption         =   "訂單匯入"
            Height          =   375
            Left            =   3480
            Style           =   1  '圖片外觀
            TabIndex        =   148
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT16 
            Enabled         =   0   'False
            Height          =   300
            Left            =   120
            TabIndex        =   147
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT16 
            Enabled         =   0   'False
            Height          =   1560
            Left            =   135
            TabIndex        =   146
            ToolTipText     =   "Local Directory"
            Top             =   720
            Visible         =   0   'False
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT16 
            Enabled         =   0   'False
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   145
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Visible         =   0   'False
            Width           =   5190
         End
         Begin VB.Label lab_Orderdetail 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   6960
            TabIndex        =   204
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lab_Orders 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   5640
            TabIndex        =   203
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "中祥訂單匯入"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   136
         Top             =   1380
         Width           =   9840
         Begin VB.CommandButton cmdOpenFilesT15 
            BackColor       =   &H0080FFFF&
            Caption         =   "開啟"
            Height          =   375
            Left            =   3480
            Style           =   1  '圖片外觀
            TabIndex        =   216
            Top             =   240
            Width           =   975
         End
         Begin VB.ComboBox cboSheetT15 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   141
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT15 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   140
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT15 
            Height          =   1560
            Left            =   135
            TabIndex        =   139
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT15 
            Height          =   300
            Left            =   135
            TabIndex        =   138
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT15 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   2400
            Style           =   1  '圖片外觀
            TabIndex        =   137
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   142
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Excel訂單匯入"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   128
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboSheetT14 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   133
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT14 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   132
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT14 
            Height          =   1560
            Left            =   120
            TabIndex        =   131
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT14 
            Height          =   300
            Left            =   135
            TabIndex        =   130
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT14 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   129
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   134
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "貨主"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   118
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboStorerkeyT13 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   126
            Text            =   "cboStorerkeyT13"
            Top             =   240
            Width           =   1605
         End
         Begin VB.CommandButton cmdImportT13 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3480
            Style           =   1  '圖片外觀
            TabIndex        =   123
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT13 
            Height          =   300
            Left            =   1815
            TabIndex        =   122
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   1560
         End
         Begin VB.DirListBox dirLocalDirT13 
            Height          =   1560
            Left            =   135
            TabIndex        =   121
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT13 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   120
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT13 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   119
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   124
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "百事RC提貨訂單匯入"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   110
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboSheetT12 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   115
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT12 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   114
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT12 
            Height          =   1560
            Left            =   135
            TabIndex        =   113
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT12 
            Height          =   300
            Left            =   135
            TabIndex        =   112
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT12 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3480
            Style           =   1  '圖片外觀
            TabIndex        =   111
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   116
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "立邦訂單匯入"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   100
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboSheetT11 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   105
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT11 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   104
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT11 
            Height          =   1560
            Left            =   135
            TabIndex        =   103
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT11 
            Height          =   300
            Left            =   135
            TabIndex        =   102
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT11 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   101
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   106
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "LNSL01-PX退貨"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   88
         Top             =   1380
         Width           =   9840
         Begin VB.CommandButton cmdImportT10 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   93
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT10 
            Height          =   300
            Left            =   135
            TabIndex        =   92
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT10 
            Height          =   1560
            Left            =   135
            TabIndex        =   91
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT10 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   90
            ToolTipText     =   "僅顯示 ""*.xls"" 類型檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT10 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   89
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   94
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "LNSL01-PX訂單"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   80
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboSheetT9 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   85
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT9 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   84
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT9 
            Height          =   1560
            Left            =   135
            TabIndex        =   83
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT9 
            Height          =   300
            Left            =   135
            TabIndex        =   82
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT9 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   81
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   86
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "LNSL01-一般退貨"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   72
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboSheetT8 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   79
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT8 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   76
            ToolTipText     =   "僅顯示 ""*.xls"" 類型檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT8 
            Height          =   1560
            Left            =   135
            TabIndex        =   75
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT8 
            Height          =   300
            Left            =   135
            TabIndex        =   74
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT8 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   73
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   77
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "LPSI01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   64
         Top             =   1380
         Width           =   9840
         Begin VB.CommandButton Command2 
            BackColor       =   &H0080FFFF&
            Caption         =   "手開單加允收期"
            Height          =   495
            Left            =   2160
            Style           =   1  '圖片外觀
            TabIndex        =   127
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdImportT7 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   69
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT7 
            Height          =   300
            Left            =   135
            TabIndex        =   68
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT7 
            Height          =   1560
            Left            =   135
            TabIndex        =   67
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT7 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   66
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT7 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   65
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   70
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "LSJR01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   56
         Top             =   1380
         Width           =   9840
         Begin VB.ComboBox cboSheetT6 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   61
            Top             =   240
            Width           =   4365
         End
         Begin VB.FileListBox filLocalFileT6 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   60
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.DirListBox dirLocalDirT6 
            Height          =   1560
            Left            =   135
            TabIndex        =   59
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.DriveListBox drvLocalDriveT6 
            Height          =   300
            Left            =   135
            TabIndex        =   58
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.CommandButton cmdImportT6 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   57
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            Left            =   4560
            TabIndex        =   62
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "LKAO01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   48
         Top             =   1380
         Width           =   9840
         Begin VB.CommandButton cmd2Excel 
            BackColor       =   &H00FFC0C0&
            Caption         =   "轉Excel"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2280
            Style           =   1  '圖片外觀
            TabIndex        =   109
            Top             =   240
            Width           =   1065
         End
         Begin VB.CommandButton cmdImportT4 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3480
            Style           =   1  '圖片外觀
            TabIndex        =   53
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT4 
            Height          =   300
            Left            =   135
            TabIndex        =   52
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT4 
            Height          =   1560
            Left            =   135
            TabIndex        =   51
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT4 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   50
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT4 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   49
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            TabIndex        =   54
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "LNIP01-退貨"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   40
         Top             =   1380
         Width           =   9840
         Begin VB.CommandButton cmdImportT5 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   45
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT5 
            Height          =   300
            Left            =   135
            TabIndex        =   44
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT5 
            Height          =   1560
            Left            =   135
            TabIndex        =   43
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT5 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   42
            ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
            Top             =   720
            Width           =   5190
         End
         Begin VB.ComboBox cboSheetT5 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   41
            Top             =   240
            Width           =   4365
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            TabIndex        =   46
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "LVTL01"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   35
         Top             =   1380
         Width           =   9840
         Begin VB.CommandButton cmdImportT2 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   39
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT2 
            Height          =   300
            Left            =   135
            TabIndex        =   38
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT2 
            Height          =   1560
            Left            =   135
            TabIndex        =   37
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT2 
            Height          =   2070
            Left            =   4560
            Pattern         =   "*.txt"
            TabIndex        =   36
            ToolTipText     =   "僅顯示 ""*.txt"" 檔案"
            Top             =   240
            Width           =   5190
         End
      End
      Begin VB.CommandButton cmd_Import 
         BackColor       =   &H0080FFFF&
         Caption         =   "匯入"
         Height          =   375
         Left            =   -71400
         Style           =   1  '圖片外觀
         TabIndex        =   9
         ToolTipText     =   "訂單匯入"
         Top             =   6000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame fraControls 
         Caption         =   "上下傳"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   -71550
         TabIndex        =   31
         Top             =   4920
         Width           =   6750
         Begin VB.CommandButton Command1 
            Caption         =   "下載"
            Height          =   375
            Left            =   3840
            TabIndex        =   32
            Top             =   225
            Width           =   615
         End
         Begin VB.Image imgSendFile 
            DragIcon        =   "frm_FTP.frx":032C
            Enabled         =   0   'False
            Height          =   345
            Left            =   135
            Picture         =   "frm_FTP.frx":076E
            Stretch         =   -1  'True
            ToolTipText     =   "Send Selected File"
            Top             =   270
            Width           =   390
         End
         Begin VB.Label lblSendFile 
            AutoSize        =   -1  'True
            Caption         =   "上傳檔案"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   555
            TabIndex        =   34
            Top             =   360
            Width           =   735
         End
         Begin VB.Image imgReceiveFile 
            Enabled         =   0   'False
            Height          =   345
            Left            =   1665
            Picture         =   "frm_FTP.frx":0BB0
            Stretch         =   -1  'True
            ToolTipText     =   "Recieve Selected File"
            Top             =   270
            Width           =   390
         End
         Begin VB.Label lblReceiveFile 
            AutoSize        =   -1  'True
            Caption         =   "Alc下載並匯入"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2085
            TabIndex        =   33
            Top             =   360
            Width           =   1185
         End
      End
      Begin VB.Frame fraRemoteFiles 
         Caption         =   "FTP 檔案"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         Left            =   -71550
         TabIndex        =   23
         Top             =   1800
         Width           =   6750
         Begin VB.ListBox lstRemoteFile 
            Enabled         =   0   'False
            Height          =   2040
            ItemData        =   "frm_FTP.frx":0FF2
            Left            =   120
            List            =   "frm_FTP.frx":0FF4
            TabIndex        =   30
            ToolTipText     =   "Remote Files"
            Top             =   840
            Width           =   6450
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "刪除"
            Enabled         =   0   'False
            Height          =   320
            Left            =   1800
            Style           =   1  '圖片外觀
            TabIndex        =   29
            ToolTipText     =   "Delete"
            Top             =   315
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.CommandButton cmdUpFolder 
            Enabled         =   0   'False
            Height          =   320
            Left            =   2775
            Picture         =   "frm_FTP.frx":0FF6
            Style           =   1  '圖片外觀
            TabIndex        =   28
            ToolTipText     =   "Move Up One Folder"
            Top             =   315
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.CommandButton cmdNewFolder 
            Caption         =   "新增資料夾"
            Enabled         =   0   'False
            Height          =   320
            Left            =   3720
            Style           =   1  '圖片外觀
            TabIndex        =   27
            ToolTipText     =   "New Folder"
            Top             =   315
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.CommandButton cmd_shp 
            BackColor       =   &H00FF8080&
            Caption         =   "SHP"
            Height          =   320
            Left            =   120
            Style           =   1  '圖片外觀
            TabIndex        =   26
            Top             =   315
            Width           =   495
         End
         Begin VB.CommandButton cmd_Alc 
            BackColor       =   &H00FF8080&
            Caption         =   "ALC"
            Height          =   320
            Left            =   660
            Style           =   1  '圖片外觀
            TabIndex        =   25
            Top             =   315
            Width           =   495
         End
         Begin VB.CommandButton cmd_CFM 
            BackColor       =   &H00FF8080&
            Caption         =   "CFM"
            Height          =   320
            Left            =   1200
            Style           =   1  '圖片外觀
            TabIndex        =   24
            Top             =   315
            Width           =   495
         End
      End
      Begin VB.Frame fraLoginInfo 
         Caption         =   "佰事達物流"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         Left            =   -74880
         TabIndex        =   10
         Top             =   1800
         Width           =   3210
         Begin VB.TextBox txtServer 
            Height          =   285
            Left            =   180
            TabIndex        =   15
            ToolTipText     =   "FTP Server Name"
            Top             =   765
            Width           =   2805
         End
         Begin VB.TextBox txtUserName 
            Height          =   285
            Left            =   180
            TabIndex        =   14
            ToolTipText     =   "User Name"
            Top             =   1395
            Width           =   2805
         End
         Begin VB.TextBox txtPassword 
            Height          =   285
            IMEMode         =   3  '暫止
            Left            =   180
            PasswordChar    =   "*"
            TabIndex        =   13
            ToolTipText     =   "Password"
            Top             =   2025
            Width           =   2805
         End
         Begin VB.CommandButton cmdLogOn 
            BackColor       =   &H0000FF00&
            Caption         =   "登  入"
            Height          =   420
            Left            =   225
            Style           =   1  '圖片外觀
            TabIndex        =   12
            ToolTipText     =   "Log On"
            Top             =   2475
            Width           =   1320
         End
         Begin VB.CommandButton cmdLogOff 
            BackColor       =   &H008080FF&
            Caption         =   "登  出"
            Enabled         =   0   'False
            Height          =   420
            Left            =   1665
            Style           =   1  '圖片外觀
            TabIndex        =   11
            ToolTipText     =   "Log Off"
            Top             =   2475
            Width           =   1320
         End
         Begin VB.Label lblServer 
            AutoSize        =   -1  'True
            Caption         =   "Server:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   495
            Width           =   630
         End
         Begin VB.Label lblUserName 
            AutoSize        =   -1  'True
            Caption         =   "User Name:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   17
            Top             =   1125
            Width           =   1005
         End
         Begin VB.Label lblPassword 
            AutoSize        =   -1  'True
            Caption         =   "Password:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   1755
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "坊元訂單匯入"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   3
         Top             =   1380
         Width           =   9885
         Begin VB.ComboBox cboSheetT3 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5400
            Style           =   2  '單純下拉式
            TabIndex        =   98
            Top             =   240
            Width           =   4365
         End
         Begin VB.CommandButton cmdImportT3 
            BackColor       =   &H0080FFFF&
            Caption         =   "匯入"
            Height          =   375
            Left            =   3000
            Style           =   1  '圖片外觀
            TabIndex        =   7
            Top             =   240
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDriveT3 
            Height          =   300
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDirT3 
            Height          =   1560
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFileT3 
            Height          =   1530
            Left            =   4560
            Pattern         =   "*.xls"
            TabIndex        =   4
            ToolTipText     =   "Local Files"
            Top             =   720
            Width           =   5190
         End
         Begin VB.Label Label1 
            Alignment       =   2  '置中對齊
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "工作表"
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
            TabIndex        =   99
            Top             =   300
            Width           =   720
         End
      End
      Begin VB.Frame fraStatus 
         Caption         =   "狀態"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   -74880
         TabIndex        =   1
         Top             =   4920
         Width           =   3210
         Begin VB.Label lblStatus 
            Alignment       =   2  '置中對齊
            Caption         =   "Not Connected"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   240
            TabIndex        =   2
            ToolTipText     =   "Connection Status"
            Top             =   360
            Width           =   2895
         End
      End
      Begin MSDataGridLib.DataGrid dg_CustInv 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   8
         Top             =   1440
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   11668
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
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
      Begin InetCtlsObjects.Inet int3 
         Left            =   -65520
         Top             =   4560
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
      End
      Begin VB.Frame fraLocalFiles 
         Caption         =   "TK訂單匯入"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   -74880
         TabIndex        =   19
         Top             =   5760
         Width           =   10080
         Begin VB.CommandButton cmdOpenFile 
            BackColor       =   &H0080FFFF&
            Caption         =   "開啟檔案"
            Height          =   375
            Left            =   2280
            Style           =   1  '圖片外觀
            TabIndex        =   108
            ToolTipText     =   "開啟其他文件專用"
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.DriveListBox drvLocalDrive 
            Height          =   300
            Left            =   135
            TabIndex        =   22
            ToolTipText     =   "Local Drive List"
            Top             =   270
            Width           =   2040
         End
         Begin VB.DirListBox dirLocalDir 
            Height          =   1560
            Left            =   135
            TabIndex        =   21
            ToolTipText     =   "Local Directory"
            Top             =   720
            Width           =   4335
         End
         Begin VB.FileListBox filLocalFile 
            Height          =   2070
            Left            =   4560
            TabIndex        =   20
            ToolTipText     =   "Local Files"
            Top             =   240
            Width           =   5415
         End
      End
      Begin MSDataGridLib.DataGrid dgMainT5 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   47
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT4 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   55
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT6 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   63
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT7 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   71
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT8 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   78
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT9 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   87
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT10 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   95
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT2 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   96
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid dgMainT3 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   97
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
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
      Begin MSDataGridLib.DataGrid dgMainT11 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   107
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT12 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   117
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT13 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   125
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT14 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   135
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT15 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   143
         Top             =   3960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT18 
         Height          =   4335
         Left            =   120
         TabIndex        =   171
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT19 
         Height          =   3375
         Left            =   -74880
         TabIndex        =   179
         Top             =   3900
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT20 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   189
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT21 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   192
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainT22 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   201
         Top             =   3960
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm_FTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private dbsrcFormHeight As Double    'Form 設計時期的高
Private dbsrcFormWidth As Double     'Form 設計時期的寬

Private strTranFileName As String
Private str_file As String
Private strOrderNo As String
Private str_Orderkey As String
Private str_CustomerID As String
Private str_CSKU, str_Note, str_DESCR As String
Private str_ExternOrderkey As String
'Private i As Double

Private RecievingSize As Boolean
Private rs_Src As ADODB.Recordset           '原始訂單資料
Private rs_Head As ADODB.Recordset          '切割後之訂單表頭資料
Private rs_Detail As ADODB.Recordset        '切割後之訂單明細資料
'Private int_Repeat As Integer
'Private int_Order As Integer
'Private int_OrderLine As Integer
Private cn_Self As ADODB.Connection
Private fso As Scripting.FileSystemObject
Private rsMainTK As ADODB.Recordset
Private rsMainT2 As ADODB.Recordset
Private rsMainT3 As ADODB.Recordset
Private rsMainT4 As ADODB.Recordset
Private rsMainT5 As ADODB.Recordset
Private rsMainT6 As ADODB.Recordset
Private rsMainT7 As ADODB.Recordset
Private rsMainT8 As ADODB.Recordset
Private rsMainT9 As ADODB.Recordset
Private rsMainT10 As ADODB.Recordset
Private rsMainT11 As ADODB.Recordset
Private rsMainT12 As ADODB.Recordset
Private rsMainT13 As ADODB.Recordset
Private rsMainT14 As ADODB.Recordset
Private rsMainT15 As ADODB.Recordset
Private rsMainT16 As ADODB.Recordset
Private rsMainT16_1 As ADODB.Recordset
Private rsMainT16_2 As ADODB.Recordset
Private rsMainT16_3 As ADODB.Recordset
Private rsMainT17 As ADODB.Recordset
Private rsMainT17_1 As ADODB.Recordset
Private rsMainT18 As ADODB.Recordset
Private rsMainT19 As ADODB.Recordset
Private rsMainT20 As ADODB.Recordset
Private rsMainT21 As ADODB.Recordset
Private rsMainT22 As ADODB.Recordset
Private strTranFileNameT4 As String
Private blDo As Boolean
Private ConfirmYN
Private Str_updatesource1 As String '主檔
Private Str_updatesource2 As String '明細
Private strFileHeader As String
Private strFileDetail As String
Private arrTmp


Private Sub cboSheetT15_Click()
'On Error GoTo err_Handle
'
'Dim strFilePath As String, strFieldName As String
'If Right(filLocalFileT15.Path, 1) <> "\" Then
'    strFilePath = filLocalFileT15.Path & "\"
'Else
'    strFilePath = filLocalFileT15.Path
'End If
'
'
'Set rsMainT15 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, cboSheetT15, strFieldName, rsMainT15)
'
'Set dgMainT15.DataSource = rsMainT15
'
'If rsMainT15 Is Nothing Then
'
'    MsgBox "查無資料!", 64, "Excel2Recordset"
''
'Else
'
'rsMainT15.Sort = "出貨單號"
'
'    SetDataGridColWidth Me.Caption, dgMainT15
'    MsgBox "此工作表共 " & rsMainT15.RecordCount & "筆明細", 64, "Excel2Recordset"
'
'End If
'
'Exit Sub
'err_Handle:
'
'Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cboSheetT17_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String
On Error GoTo err_Handle

'確認路徑是否帶"\"
If Right(filLocalFileT17.Path, 1) = "\" Then
    strFilePath = filLocalFileT17.Path
Else
    strFilePath = filLocalFileT17.Path & "\"
End If

'建立欄位名稱陣列
strFieldName = "DC代號" & Chr(9) & "客戶代號" & Chr(9) & "EXE文件編號" & Chr(9) & "SAP DN NO." & Chr(9) & "單據類別" & Chr(9) & "單據產生日" & Chr(9) & "預計出貨日" & Chr(9) & "項次" & Chr(9) & "商品編碼" & Chr(9) & "商品訂購數量" & Chr(9) & "銷售別" & Chr(9) & "商品最小數量" & Chr(9) & "客戶進價" & Chr(9) & _
              "折讓進額" & Chr(9) & "實際出貨數量" & Chr(9) & "實際出貨倉別" & Chr(9) & "稅別" & Chr(9) & "批次" & Chr(9) & "實際檢貨日期" & Chr(9) & "貨主" & Chr(9) & "送貨地址" & Chr(9) & "銷售組織" & Chr(9) & "營業所" & Chr(9) & "業務組長" & Chr(9) & "王安單號" & Chr(9) & "SAP訂貨單位" & Chr(9) & "原因" & Chr(9) & "客戶名稱" & Chr(9) & "備注" & Chr(9) & "客戶通路別" & Chr(9)

'"DC代號" & Chr(9) & "客戶代號" & Chr(9) & "SAP DN NO." & Chr(9) & "單據產生日" & Chr(9) & "預計出貨日" & Chr(9) & "項次" & Chr(9) & "商品編碼" & Chr(9) & "商品訂購數量" & Chr(9) & "送貨地址" & Chr(9) & "客戶名稱" & Chr(9) & "備注" & Chr(9) & "客戶通路別" & Chr(9)
If Right(filLocalFileT17.Path, 1) <> "\" Then
    strFilePath = filLocalFileT17.Path & "\"
Else
    strFilePath = filLocalFileT17.Path
End If

Set rsMainT17 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, cboSheetT15, strFieldName, rsMainT15)
'''''''''''''''''''''''''''
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT17.FileName)   '打開路徑
    
    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = (cboSheetT17) Then .Sheets(i).Select: Exit For '選定工作表
    Next

    'k = 1 '預設由第一列開始匯入
    
    '若無來源欄位名稱
    If strFieldName = "" Then
        '取欄位名稱
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '由第二列開始匯入
    End If
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "DC代號" Then k = i: Exit For
    Next i
    
    If Trim(.Cells(i, 1)) <> "DC代號" Then MsgBox "找不到""DC代號""欄位名稱，檔案載入終止!", 64, "金盛世訂單匯入": GoTo endsub
    
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT17 = Nothing: GoTo endsub
    
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsMainT17.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT17.CursorType = adOpenKeyset
    rsMainT17.LockType = adLockOptimistic
    rsMainT17.Open
    
    '寫入Recordset  '從這邊開始往下寫
    Do While Len(RTrim(.Cells(k + 1, 1))) > 0
    rsMainT17.AddNew
        For j = 1 To UBound(arrTmp)
            rsMainT17(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
        Next j
    rsMainT17.Update
    k = k + 1
    Loop
    
    If rsMainT17.RecordCount > 0 Then rsMainT17.MoveFirst

    'Call OffLineRecordset(rsMainT15, rs)
    
   ' rsMainT15.Close: Set rsMainT15 = Nothing
  
End With
'''''''''''''''''''''''''''
Set dgMainT17.DataSource = rsMainT17

If rsMainT17 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT17
    MsgBox "此工作表共 " & rsMainT17.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
End If

endsub:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Exit Sub

err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "金盛世訂單匯入")

End Sub

Private Sub cboSheetT18_Click()
On Error GoTo err_Handle

Dim strFilePath As String, strFieldName As String
If Right(filLocalFileT18.Path, 1) <> "\" Then
    strFilePath = filLocalFileT18.Path & "\"
Else
    strFilePath = filLocalFileT18.Path
End If

'strFieldName = "Delivery" & Chr(9) & "Sold-to.Pt" & Chr(9) & "Name.of.sold-to" & Chr(9) & "Ship-To.Pt" & Chr(9) & "Name.of.the.ship-to.Party" & Chr(9) & "Item" & Chr(9) & "Material" & Chr(9) & "Plnt" & Chr(9) & "SLoc" & Chr(9) & "Batch" & Chr(9) & "Route" & Chr(9) & "Deliv.date" & Chr(9) & "Qty(stckpg.unit)" & Chr(9) & "BUn" & Chr(9) & "Delivery.qty" & Chr(9) & "SU" & Chr(9) & "order.no" & Chr(9) & "po.no" & Chr(9) & "remarks" & Chr(9)

Set rsMainT18 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT18.FileName, cboSheetT18, strFieldName, rsMainT18)

Set dgMainT18.DataSource = rsMainT18

If rsMainT18 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
'
Else

rsMainT18.Sort = "交貨"

    SetDataGridColWidth Me.Caption, dgMainT18
    MsgBox "此工作表共 " & rsMainT18.RecordCount & "筆明細", 64, "Excel2Recordset"

End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub


Private Sub cboSheetT19_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String
On Error GoTo err_Handle

'確認路徑是否帶"\"
If Right(filLocalFileT19.Path, 1) = "\" Then
    strFilePath = filLocalFileT19.Path
Else
    strFilePath = filLocalFileT19.Path & "\"
End If

'建立欄位名稱陣列
strFieldName = "交貨單號" & Chr(9) & "單據類別" & Chr(9) & "收貨倉別" & Chr(9) & "收貨庫存地" & Chr(9) & "項次" & Chr(9) & "品號" & Chr(9) & "最小單位" & Chr(9) & "商品訂購數量" & Chr(9) & "預計交貨日期" & Chr(9) & "銷售組織" & Chr(9) & "營業所" & Chr(9) & "客戶代號" & Chr(9) & "SAP單位" & Chr(9) & "收貨地址" & Chr(9) & "客戶名稱" & Chr(9) & "通路" & Chr(9)
'交貨單號    單據類別    收貨倉別    收貨庫存地  項次    品號    最小單位     商品訂購數量   預計交貨日期    銷售組織    營業所  客戶代號    SAP單位 收貨地址    客戶名稱    通路


If Right(filLocalFileT19.Path, 1) <> "\" Then
    strFilePath = filLocalFileT19.Path & "\"
Else
    strFilePath = filLocalFileT19.Path
End If

Set rsMainT19 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, cboSheetT15, strFieldName, rsMainT15)
'''''''''''''''''''''''''''
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT19.FileName)   '打開路徑
    
    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = (cboSheetT19) Then .Sheets(i).Select: Exit For '選定工作表
    Next

    'k = 1 '預設由第一列開始匯入
    
    '若無來源欄位名稱
    If strFieldName = "" Then
        '取欄位名稱
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '由第二列開始匯入
    End If
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "交貨單號" Then k = i: Exit For
    Next i
    
    If Trim(.Cells(i, 1)) <> "交貨單號" Then MsgBox "找不到""交貨單號""欄位名稱，檔案載入終止!", 64, "金盛世手退貨匯入": GoTo endsub
    
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT19 = Nothing: GoTo endsub
    
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsMainT19.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT19.CursorType = adOpenKeyset
    rsMainT19.LockType = adLockOptimistic
    rsMainT19.Open
    
    '寫入Recordset  '從這邊開始往下寫
    Do While Len(RTrim(.Cells(k + 1, 1))) > 0
    rsMainT19.AddNew
        For j = 1 To UBound(arrTmp)
            rsMainT19(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
        Next j
    rsMainT19.Update
    k = k + 1
    Loop
    
    If rsMainT19.RecordCount > 0 Then rsMainT19.MoveFirst

    'Call OffLineRecordset(rsMainT15, rs)
    
   ' rsMainT15.Close: Set rsMainT15 = Nothing
  
End With
'''''''''''''''''''''''''''
Set dgMainT19.DataSource = rsMainT19

rsMainT19.Sort = "交貨單號"

If rsMainT19 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT19
    MsgBox "此工作表共 " & rsMainT19.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
End If

endsub:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Exit Sub

err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "金盛世退貨匯入")

End Sub

Private Sub cboSheetT20_Change()
'On Error GoTo err_Handle
'
'Dim strFilePath As String, strFieldName As String
'If Right(filLocalFileT20.Path, 1) <> "\" Then
'    strFilePath = filLocalFileT20.Path & "\"
'Else
'    strFilePath = filLocalFileT20.Path
'End If
'
'
'Set rsMainT20 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT20.FileName, cboSheetT20, strFieldName, rsMainT20)
'
'Set dgMainT20.DataSource = rsMainT20
'
'If rsMainT20 Is Nothing Then
'
'    MsgBox "查無資料!", 64, "Excel2Recordset"
''
'Else
'
'rsMainT20.Sort = "出貨單號"
'
'    SetDataGridColWidth Me.Caption, dgMainT20
'    MsgBox "此工作表共 " & rsMainT20.RecordCount & "筆明細", 64, "Excel2Recordset"
'
'End If
'
'Exit Sub
'err_Handle:
'
'Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cboSheetT21_Click()
'Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
'Dim str As String, strFieldName As String, strFilePath As String
'On Error GoTo err_Handle
'
''確認路徑是否帶"\"
'If Right(filLocalFileT21.Path, 1) = "\" Then
'    strFilePath = filLocalFileT21.Path
'Else
'    strFilePath = filLocalFileT21.Path & "\"
'End If
'
''建立欄位名稱陣列
'strFieldName = "調撥日" & Chr(9) & "產品代號" & Chr(9) & "品名" & Chr(9) & "數量" & Chr(9) & "調撥單號" & Chr(9) & "撥出倉庫名稱" & Chr(9) & "撥入倉庫名稱" & Chr(9) & "備註" & Chr(9) & "袋" & Chr(9) & "個" & Chr(9) & "數量" & Chr(9)
''訂單單號    預計到貨日  客戶代號    客戶名稱    單店代號    備註    品號    箱  袋  個  數量
'
'If Right(filLocalFileT21.Path, 1) <> "\" Then
'    strFilePath = filLocalFileT21.Path & "\"
'Else
'    strFilePath = filLocalFileT21.Path
'End If
'
'Set rsMainT21 = New ADODB.Recordset
''''''''''''''''''''''''''''
'Set MyXlsApp = CreateObject("Excel.Application")
'
'With MyXlsApp
'    .Visible = False
'    .Workbooks.Open (strFilePath & filLocalFileT21.FileName)   '打開路徑
'
'    '尋找指定工作表
'    For i = 1 To .Sheets.Count
'        If (.Sheets(i).Name) = (cboSheetT21) Then .Sheets(i).Select: Exit For '選定工作表
'    Next
'
'    'k = 1 '預設由第一列開始匯入
'
'    '若無來源欄位名稱
'    If strFieldName = "" Then
'        '取欄位名稱
'        For i = 1 To 255
'            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
'               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
'        Next i
'        'k = 2 '由第二列開始匯入
'    End If
'    For i = 1 To 255
'            If Trim(.Cells(i, 1)) = "訂單單號" Then k = i: Exit For
'    Next i
'
'    '訂單單號    預計到貨日  客戶代號    客戶名稱    單店代號    備註    品號    箱  袋  個  數量   '手key單一定要防止欄位錯誤
'    If Trim(.Cells(i, 1)) <> "調撥日" Then MsgBox "找不到""調撥日""欄位名稱，檔案載入終止!", 64, "中祥RC訂單匯入": GoTo endsub
'    If Trim(.Cells(i, 2)) <> "產品代號" Then MsgBox "找不到""產品代號""欄位名稱，檔案載入終止!", 64, "中祥RC訂單匯入": GoTo endsub
'    If Trim(.Cells(i, 3)) <> "品名" Then MsgBox "找不到""品名""欄位名稱，檔案載入終止!", 64, "中祥RC訂單匯入": GoTo endsub
'    If Trim(.Cells(i, 4)) <> "數量" Then MsgBox "找不到""數量""欄位名稱，檔案載入終止!", 64, "中祥RC訂單匯入": GoTo endsub
'    If Trim(.Cells(i, 5)) <> "調撥單號" Then MsgBox "找不到""調撥單號""欄位名稱，檔案載入終止!", 64, "中祥RC訂單匯入": GoTo endsub
'    If Trim(.Cells(i, 6)) <> "撥出倉庫名稱" Then MsgBox "找不到""撥出倉庫名稱""欄位名稱，檔案載入終止!", 64, "中祥RC訂單匯入": GoTo endsub
'    If Trim(.Cells(i, 7)) <> "撥入倉庫名稱" Then MsgBox "找不到""撥入倉庫名稱""欄位名稱，檔案載入終止!", 64, "中祥RC訂單匯入": GoTo endsub
'    If Trim(.Cells(i, 8)) <> "備註" Then MsgBox "找不到""備註""欄位名稱，檔案載入終止!", 64, "中祥RC訂單匯入": GoTo endsub
'
'
'    '切割欄位名稱
'    arrTmp = Split(strFieldName, Chr(9))
'
'
'    If UBound(arrTmp) < 1 Then Set rsMainT21 = Nothing: GoTo endsub
'
'    '建立Recordset
'    For i = 0 To UBound(arrTmp) - 1
'        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
'        rsMainT21.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
'    Next i
'
'    rsMainT21.CursorType = adOpenKeyset
'    rsMainT21.LockType = adLockOptimistic
'    rsMainT21.Open
'
'    '寫入Recordset  '從這邊開始往下寫
'    Do While Len(RTrim(.Cells(k + 1, 1))) > 0
'    rsMainT21.AddNew
'        For j = 1 To UBound(arrTmp)
'            rsMainT21(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
'        Next j
'    rsMainT21.Update
'    k = k + 1
'    Loop
'
'    If rsMainT21.RecordCount > 0 Then rsMainT21.MoveFirst
'
'
'End With
''''''''''''''''''''''''''''
'Set dgMainT21.DataSource = rsMainT21
'
'rsMainT21.Sort = "調撥單號"
'
'
'If rsMainT21 Is Nothing Then
'
'    MsgBox "查無資料!", 64, "Excel2Recordset"
'
'Else
'    SetDataGridColWidth Me.Caption, dgMainT21
'    MsgBox "此工作表共 " & rsMainT21.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
'
'End If
'
'endsub:
'MyXlsApp.Quit: Set MyXlsApp = Nothing
'Exit Sub
'
'err_Handle:
'
'Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "中祥RC訂單匯入")
End Sub


Private Sub cboSheetT22_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String
On Error GoTo err_Handle

'確認路徑是否帶"\"
If Right(filLocalFileT22.Path, 1) = "\" Then
    strFilePath = filLocalFileT22.Path
Else
    strFilePath = filLocalFileT22.Path & "\"
End If

'建立欄位名稱陣列
strFieldName = "退貨單號" & Chr(9) & "預計收退日" & Chr(9) & "客戶代號" & Chr(9) & "客戶名稱" & Chr(9) & "單店代號" & Chr(9) & "備註" & Chr(9) & "品號" & Chr(9) & "箱" & Chr(9) & "袋" & Chr(9) & "個" & Chr(9) & "數量" & Chr(9)
'退貨單號    預計收退日  客戶代號    客戶名稱    單店代號    備註    品號    箱  袋  個  數量

If Right(filLocalFileT22.Path, 1) <> "\" Then
    strFilePath = filLocalFileT22.Path & "\"
Else
    strFilePath = filLocalFileT22.Path
End If

Set rsMainT22 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, cboSheetT15, strFieldName, rsMainT15)
'''''''''''''''''''''''''''
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT22.FileName)   '打開路徑
    
    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = (cboSheetT22) Then .Sheets(i).Select: Exit For '選定工作表
    Next

    'k = 1 '預設由第一列開始匯入
    
    '若無來源欄位名稱
    If strFieldName = "" Then
        '取欄位名稱
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '由第二列開始匯入
    End If
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "退貨單號" Then k = i: Exit For
    Next i

    '退貨單號    預計收退日  客戶代號    客戶名稱    單店代號    備註    品號    箱  袋  個  數量   '手key單一定要防止欄位錯誤
    If Trim(.Cells(i, 1)) <> "退貨單號" Then MsgBox "找不到""退貨單號""欄位名稱，檔案載入終止!", 64, "永豐餘P&G訂單匯入": GoTo endsub
    If Trim(.Cells(i, 2)) <> "預計收退日" Then MsgBox "找不到""預計收退日""欄位名稱，檔案載入終止!", 64, "永豐餘P&G訂單匯入": GoTo endsub
    If Trim(.Cells(i, 3)) <> "客戶代號" Then MsgBox "找不到""客戶代號""欄位名稱，檔案載入終止!", 64, "永豐餘P&G訂單匯入": GoTo endsub
    If Trim(.Cells(i, 4)) <> "客戶名稱" Then MsgBox "找不到""客戶名稱""欄位名稱，檔案載入終止!", 64, "永豐餘P&G訂單匯入": GoTo endsub
    If Trim(.Cells(i, 5)) <> "單店代號" Then MsgBox "找不到""單店代號""欄位名稱，檔案載入終止!", 64, "永豐餘P&G訂單匯入": GoTo endsub
    If Trim(.Cells(i, 6)) <> "備註" Then MsgBox "找不到""備註""欄位名稱，檔案載入終止!", 64, "永豐餘P&G訂單匯入": GoTo endsub
    If Trim(.Cells(i, 7)) <> "品號" Then MsgBox "找不到""品號""欄位名稱，檔案載入終止!", 64, "永豐餘P&G訂單匯入": GoTo endsub
    If Trim(.Cells(i, 8)) <> "箱" Then MsgBox "找不到""箱""欄位名稱，檔案載入終止!", 64, "永豐餘P&G訂單匯入": GoTo endsub
    If Trim(.Cells(i, 9)) <> "袋" Then MsgBox "找不到""袋""欄位名稱，檔案載入終止!", 64, "永豐餘P&G訂單匯入": GoTo endsub
    If Trim(.Cells(i, 10)) <> "個" Then MsgBox "找不到""個""欄位名稱，檔案載入終止!", 64, "永豐餘P&G訂單匯入": GoTo endsub
    If Trim(.Cells(i, 11)) <> "數量" Then MsgBox "找不到""數量""欄位名稱，檔案載入終止!", 64, "永豐餘P&G訂單匯入": GoTo endsub
    
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT19 = Nothing: GoTo endsub
    
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsMainT22.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT22.CursorType = adOpenKeyset
    rsMainT22.LockType = adLockOptimistic
    rsMainT22.Open
    
    '寫入Recordset  '從這邊開始往下寫
    Do While Len(RTrim(.Cells(k + 1, 1))) > 0
    rsMainT22.AddNew
        For j = 1 To UBound(arrTmp)
            rsMainT22(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
        Next j
    rsMainT22.Update
    k = k + 1
    Loop
    
    If rsMainT22.RecordCount > 0 Then rsMainT22.MoveFirst

    'Call OffLineRecordset(rsMainT15, rs)
    
   ' rsMainT15.Close: Set rsMainT15 = Nothing
  
End With
'''''''''''''''''''''''''''
Set dgMainT22.DataSource = rsMainT22

rsMainT22.Sort = "退貨單號"


If rsMainT22 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT22
    MsgBox "此工作表共 " & rsMainT22.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
End If

endsub:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Exit Sub

err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "永豐餘P&G訂單匯入")
End Sub


Private Sub Cmb_Import1_Click()
 
Dim strFileName As String, strFieldName As String, Str_Filter As String
Str_Filter = ""

''斷開目前的
'LCDisConnect "bestprepares\IPC$"
'LCDisConnect "192.168.200.200\IPC$"
''重新連線
'LCConnect "192.168.200.200", "LMBO01", "34245356"

On Error GoTo err_Handle

'MsgBox "請依照下列方式開始執行訂單開啟作業:" & Chr(13) & "1.先點選訂單主檔(ST開頭)進行訂單開啟" & Chr(13) & "2.再點選訂單明細(SD開頭)進行明細開啟" & Chr(13) & "3.請按確定後開始進行^_^", vbOKOnly + vbInformation, "毛寶訂單開啟"

'匯入訂單主檔
With dlgCommonDialog
    .DialogTitle = "毛寶訂單主檔匯入"
    .CancelError = True
    '.InitDir = App.Path
    .InitDir = "\\192.168.200.200\ftp$\LMBO01\to_Best"
    '.InitDir = "ftp://LMBO01:34245356@192.168.2.202"
    'ToDo: 設定通用對話方塊控制項的旗標及屬性
    .Filter = "ST*.txt|ST*.txt"
    '.Filter = "rtb*.txt|rtb*.txt"
    .ShowOpen
    strFileName = .FileName
    
    If err.Number = cdlCancel Then strFileName = "": Exit Sub
    
    If Len(strFileName) = 0 Then Exit Sub

End With

strFileHeader = strFileName
arrTmp = Split(strFileName, "\")
strFileName = arrTmp(UBound(arrTmp)) '取得count
Str_Filter = Mid(strFileName, 3, Len(strFileName))

If Dir(strFileName) = "" Then MsgBox "找不到檔案！", vbOKOnly + vbInformation, "毛寶訂單開啟": Exit Sub '找不到檔案

Call FixedLenghtText2Recordset(strFileName, "分分司代號,訂單號碼,訂單日期,發票號碼,發票號碼檢查碼,發票日期,客戶編號,客戶名稱,業代代號,下貨收現,送貨地址,聯式,統一編號,折讓金額,數量折讓金額,特別折讓金額,現金折讓,貨款,稅前金額,稅額,備註," & _
                                            "客戶訂單編號,隨貨附發票碼,隨貨附訂單碼,計算物流費,送貨否,訂單種類,實收量處理MARK,連絡人,電話,業代姓名,主管姓名,指送客戶,預計日期,運費,付款方式,業務手機,是否為電子發票,總重量,信卡後4碼,代收貨款,發票列印方式,電話2,統計對象," & _
                                            "縣市別,行政區,樓層,越庫訂單,提貨倉,稅區/稅率,客戶簡稱,訂單窗口,關聯訂單號碼", "1,8,7,10,2,7,8,50,3,2,70,1,8,8,8,8,8,10,10,8,70,25,1,1,1,1,2,1,12,20,12,12,50,7,8,1,20,1,6,4,10,1,20,8,3,3,2,1,12,10,40,10,10", rsMainT16)

'提貨倉取左邊三碼 add by Gemini @ 20160425
rsMainT16.MoveFirst
Do While Not rsMainT16.EOF
    rsMainT16("提貨倉") = Left(rsMainT16("提貨倉"), 3)
    rsMainT16.MoveNext
Loop

Set dgMainT16.DataSource = rsMainT16

'Recordset2Excel "TEST", rsMainT16


'紀錄檔名
Str_updatesource1 = strFileName

''匯入訂單明細檔
'With dlgCommonDialog
'    .DialogTitle = "毛寶訂單明細檔匯入"
'    .CancelError = True
'    '.InitDir = App.Path
'    .InitDir = "\\192.168.200.200\ftp$\LMBO01\to_Best"
'    'ToDo: 設定通用對話方塊控制項的旗標及屬性
'    '.Filter = "SD*.txt|SD*.txt"
'    .Filter = "SD" & Str_Filter & "|" & "SD" & Str_Filter
''    .Filter = "rdb*.txt|rdb*.txt"
'    .ShowOpen
'    strFileName = .FileName
'
'    If err.Number = cdlCancel Then strFileName = "": Exit Sub
'
'    If Len(strFileName) = 0 Then Exit Sub
'
'End With

strFileDetail = Replace(strFileHeader, strFileName, "SD" & Str_Filter)

strFileName = "SD" & Str_Filter

'arrTmp = Split(strFileName, "\")
'strFileName = arrTmp(UBound(arrTmp)) '取得count


If Dir(strFileName) = "" Then MsgBox "找不到檔案！", vbOKOnly + vbInformation, "毛寶訂單明細檔匯入": Exit Sub '找不到檔案

Call FixedLenghtText2Recordset(strFileName, "訂單號碼,產品編號,產品名稱,訂貨量,單價(未稅),訂貨金額(未稅),單價(含稅),訂貨金額(含稅),訂貨量-實收量,國際條碼,行號,單位,訂單種類,發票明細列印否,允收期", "8,16,60,10,8,10,8,10,10,25,7,2,2,1,20", rsMainT16_1)

Set dgMainT16_1.DataSource = rsMainT16_1

'MsgBox "毛寶訂單主檔開啟:" & rsMainT16.RecordCount & "筆，請確認筆數是否正確!", vbOKOnly + vbInformation, "毛寶訂單開啟"

MsgBox "毛寶訂單主檔開啟:" & rsMainT16.RecordCount & "筆，訂單明細檔開啟:" & rsMainT16_1.RecordCount & "筆，請確認筆數是否正確!", vbOKOnly + vbInformation, "毛寶訂單明細檔開啟"

Str_updatesource2 = strFileName

rsMainT16.Sort = "訂單號碼,訂單種類"
rsMainT16_1.Sort = "訂單號碼,訂單種類,行號"


lab_Orders.Caption = "訂單:" & rsMainT16.RecordCount & "筆":
lab_Orderdetail.Caption = "明細:" & rsMainT16_1.RecordCount & "筆":

'Recordset2Excel "TEST", rsMainT16_1

'排序


'馬玉山的程式碼
'datagrid1=出貨總表;銷貨明細表;轉撥明細表;寄庫銷貨明細表
'Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
'Dim str As String, strFieldName As String, strFilePath As String, str_TmpSQL As String, S As String, Str_Sku As String
'Dim bl_Check As Boolean '檢查匯入的資料~有無出現在總表中~沒有就stop
'bl_Check = True
'S = "": Str_Sku = ""
''S記錄上一筆消貨單號,Str_sku記錄上一筆品號,如果空白則帶上一筆
'
'On Error GoTo err_Handle
'SSTab2.Tab = 0: SSTab2.Enabled = False: Cmb_Import1.Enabled = False
'
'Call DB_Connect_Self(cn_string) '建立新連線
'
''確認路徑是否帶"\"
'If Right(filLocalFileT16.Path, 1) = "\" Then
'    strFilePath = filLocalFileT16.Path
'Else
'    strFilePath = filLocalFileT16.Path & "\"
'End If
'
''建立欄位名稱陣列
'strFieldName = "單據名稱" & Chr(9) & "銷貨單別" & Chr(9) & "銷貨單號" & Chr(9) & "單據日期" & Chr(9) & "指定日期" & Chr(9) & "客戶代號" & Chr(9) & "客戶簡稱" & Chr(9) & "件數" & Chr(9) & "送貨地址" & Chr(9) & "備註" & Chr(9) '出貨總表
'If Right(filLocalFileT16.Path, 1) <> "\" Then
'    strFilePath = filLocalFileT16.Path & "\"
'Else
'    strFilePath = filLocalFileT16.Path
'End If
'
'Set rsMainT16 = New ADODB.Recordset
'
'Set MyXlsApp = CreateObject("Excel.Application")
'
'With MyXlsApp
'    .Visible = False
'    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '打開路徑
'    '尋找指定工作表
'    For i = 1 To .Sheets.Count
'        If (.Sheets(i).Name) = "出貨總表" Then .Sheets(i).Select: Exit For '選定工作表
'    Next
'
'    If (.ActiveSheet.Name) <> "出貨總表" Then MsgBox "找不到出貨總表工作表!!", 16, "開啟檔案中止": GoTo endsub
'
'    'k = 1 '預設由第一列開始匯入
'
'    '若無來源欄位名稱
'    If strFieldName = "" Then
'        '取欄位名稱
'        For i = 1 To 255
'            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
'               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
'        Next i
'        'k = 2 '由第二列開始匯入
'    End If
'
'    For i = 1 To 255
'            If Trim(.Cells(i, 1)) = "單據名稱" Then k = i: Exit For
'    Next i
'
'    '切割欄位名稱
'    arrTmp = Split(strFieldName, Chr(9))
'
'    'Dim rsMainT15 As New ADODB.Recordset
'
'    If UBound(arrTmp) < 1 Then Set rsMainT16 = Nothing: GoTo endsub
'
'    '建立Recordset
'    For i = 0 To UBound(arrTmp) - 1
'        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "出貨總表工作表，第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
'        rsMainT16.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
'    Next i
'
'    rsMainT16.CursorType = adOpenKeyset
'    rsMainT16.LockType = adLockOptimistic
'    rsMainT16.Open
'
'    '寫入Recordset  '從這邊開始往下寫
'    Do While Len(RTrim(.Cells(k + 1, 8))) > 0
'    rsMainT16.AddNew
'        For j = 1 To UBound(arrTmp)
'            rsMainT16(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
'        Next j
'    rsMainT16.Update
'    k = k + 1
'    Loop
'
'    If rsMainT16.RecordCount > 0 Then rsMainT16.MoveFirst
'
'End With
'''以下將datagrid的資料存入佔存table中
''    '新增tabel
''    cn_Self.Execute "if object_id ('tempdb..##all_data') is not null drop table tempdb..##all_data ", RowsAffect, adExecuteNoRecords
''    str_TmpSQL = "CREATE TABLE tempdb..##all_data(單據名稱 varchar(30),銷貨單別 varchar(30),銷貨單號 varchar(30),單據日期 varchar(30),指定日期 varchar(30),客戶代號 varchar(30),客戶簡稱 varchar(30),件數 varchar(30),送貨地址 varchar(80),備註 varchar(60))"
''    Call Confirm_Recordset_Closed(tmp_Rs)
''    tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
''    '放入datagrid資料到table
''    dgMainT16.Visible = False
''    Do While Not rsMainT16.EOF
''        '將匯入的excel資料 存入暫存的資料表##SkuCompare中
''        str_TmpSQL = "INSERT INTO tempdb..##all_data (單據名稱,銷貨單別,銷貨單號,單據日期,指定日期,客戶代號,客戶簡稱,件數,送貨地址,備註) " & _
''                     "VALUES ('" & Trim(rsMainT16("單據名稱").Value) & "','" & Trim(rsMainT16("銷貨單別").Value) & "','" & Trim(rsMainT16("銷貨單號").Value) & "','" & _
''                     "" & Trim(rsMainT16("單據日期").Value) & "','" & Trim(rsMainT16("指定日期").Value) & "','" & Trim(rsMainT16("客戶代號").Value) & "','" & Trim(rsMainT16("客戶簡稱").Value) & "','" & _
''                     "" & Trim(rsMainT16("件數").Value) & "','" & Trim(rsMainT16("送貨地址").Value) & "','" & Trim(rsMainT16("備註").Value) & "')"
''
''        Call Confirm_Recordset_Closed(tmp_Rs)
''        tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
''        rsMainT16.MoveNext
''    Loop
''    dgMainT16.Visible = True
'
'Set dgMainT16.DataSource = rsMainT16
'
'If rsMainT16 Is Nothing Then
'
'    MsgBox "查無資料!", 64, "Excel2Recordset"
'
'Else
'    SetDataGridColWidth Me.Caption, dgMainT16 '設定欄寬
'End If
'
''如果沒有出貨總表則跳離
'If rsMainT16 Is Nothing Then MsgBox "找不到出貨總表工作表!!", 16, "開啟檔案中止": SSTab2.Enabled = True: Cmb_Import1.Enabled = True: Exit Sub
'If rsMainT16.EOF Then MsgBox "找不到出貨總表工作表!!", 16, "開啟檔案中止": SSTab2.Enabled = True: Cmb_Import1.Enabled = True: Exit Sub
'
''/////////////////////////////////////////////////////////////////////////////////匯入銷貨明細檔/////////////////////////////////////////////////////////////////////////////////
'SSTab2.Tab = 1
'strFieldName = "銷貨單號" & Chr(9) & "銷貨日期" & Chr(9) & "指定日期" & Chr(9) & "客戶代號" & Chr(9) & "客戶簡稱" & Chr(9) & "品號" & Chr(9) & "品名" & Chr(9) & "銷貨數量" & Chr(9) & "贈/備品量" & Chr(9) & "單位" & Chr(9) & "批號" & Chr(9) & "備註" & Chr(9) & "件數" & Chr(9) & "送貨地址" & Chr(9)  '銷貨明細表
'
'Set rsMainT16_1 = New ADODB.Recordset
'
''Set MyXlsApp = CreateObject("Excel.Application")
'
'With MyXlsApp
'    .Visible = False
''    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '打開路徑
'    '尋找指定工作表
'    For i = 1 To .Sheets.Count
'        If (.Sheets(i).Name) = "銷貨明細表" Then .Sheets(i).Select: Exit For '選定工作表
'    Next
'
'    If (.ActiveSheet.Name) <> "銷貨明細表" Then MsgBox "找不到銷貨明細表工作表!!", 16, "開啟檔案中止": GoTo endsub
'
'    'k = 1 '預設由第一列開始匯入
'
'    '若無來源欄位名稱
'    If strFieldName = "" Then
'        '取欄位名稱
'        For i = 1 To 255
'            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
'               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
'        Next i
'        'k = 2 '由第二列開始匯入
'    End If
'
'    For i = 1 To 255
'            If Trim(.Cells(i, 1)) = "銷貨單號" Then k = i: Exit For
'    Next i
'
'    '切割欄位名稱
'    arrTmp = Split(strFieldName, Chr(9))
'
'    If UBound(arrTmp) < 1 Then Set rsMainT16 = Nothing: GoTo endsub
'
'    '建立Recordset
'    For i = 0 To UBound(arrTmp) - 1
'        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "銷貨明細表工作表，第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
'        rsMainT16_1.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
'    Next i
'
'    rsMainT16_1.CursorType = adOpenKeyset
'    rsMainT16_1.LockType = adLockOptimistic
'    rsMainT16_1.Open
'    rsMainT16.MoveFirst: S = ""
'
'    '寫入Recordset  '從這邊開始往下寫
'    Do While Len(RTrim(.Cells(k + 1, 8))) > 0   '挑準數量為空值則停止
''    If RTrim(.Cells(k + 1, 6)) = "60400119" Then '排除運費
''    Else
'        rsMainT16_1.AddNew
'            For j = 1 To UBound(arrTmp)
'                If Len(Trim(rsMainT16_1("銷貨單號").Value)) = 0 Then rsMainT16_1("銷貨單號").Value = S
'                If Len(Trim(rsMainT16_1("品號").Value)) = 0 Then rsMainT16_1("品號").Value = Str_Sku
'                If j = UBound(arrTmp) Then    '送貨地址欄位
'                rsMainT16.MoveFirst
'                    Do While Not rsMainT16.EOF
'                        Str_Sku = Trim(rsMainT16_1("品號").Value)
'                        If Trim(rsMainT16_1("銷貨單號").Value) = Trim(rsMainT16("銷貨單別").Value) & "-" & Trim(rsMainT16("銷貨單號").Value) Then rsMainT16_1("送貨地址").Value = Trim(rsMainT16("送貨地址").Value):  rsMainT16_1("指定日期").Value = Trim(rsMainT16("指定日期").Value): rsMainT16_1("備註").Value = Trim(rsMainT16("備註").Value): rsMainT16_1("件數").Value = Trim(rsMainT16("件數").Value): S = Trim(rsMainT16_1("銷貨單號").Value): bl_Check = False: Exit Do
'                        rsMainT16.MoveNext
'                    Loop
'                    If bl_Check = True Then MsgBox "出貨總表查無:" & Trim(rsMainT16_1("銷貨單號").Value) & "資料!", 64, "銷貨明細表匯入中止": SSTab2.Enabled = True: Cmb_Import1.Enabled = True: GoTo endsub
'                Else
'                    rsMainT16_1(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
'                End If
'            Next j
'        rsMainT16_1.Update
''    End If
'    k = k + 1
'
'    Loop
'
'    If rsMainT16_1.RecordCount > 0 Then rsMainT16_1.MoveFirst
'
'End With
'
'Set dgMainT16_1.DataSource = rsMainT16_1
'
'If rsMainT16_1 Is Nothing Then
'    MsgBox "銷貨明細表查無資料!", 64, "Excel2Recordset"
'Else
'    SetDataGridColWidth Me.Caption, dgMainT16_1
'End If
''////////////////////////////////////////////////////////////////////////////////匯入轉撥明細表/////////////////////////////////////////////////////////////////////////////////////
'SSTab2.Tab = 2: S = "": Str_Sku = ""
'
''建立欄位名稱陣列
'strFieldName = "單別-單號" & Chr(9) & "單據日期" & Chr(9) & "轉入庫別" & Chr(9) & "品號" & Chr(9) & "品名" & Chr(9) & "轉撥數量" & Chr(9) & "單位" & Chr(9) & "批號" & Chr(9) & "備註" & Chr(9) & "指定日期" & Chr(9) & "件數" & Chr(9) & "送貨地址" & Chr(9) '轉撥明細表
'If Right(filLocalFileT16.Path, 1) <> "\" Then
'    strFilePath = filLocalFileT16.Path & "\"
'Else
'    strFilePath = filLocalFileT16.Path
'End If
'
'Set rsMainT16_2 = New ADODB.Recordset
'
''Set MyXlsApp = CreateObject("Excel.Application")
'
'With MyXlsApp
'    .Visible = False
''    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '打開路徑
'    '尋找指定工作表
'    For i = 1 To .Sheets.Count
'        If (.Sheets(i).Name) = "轉撥明細表" Then .Sheets(i).Select: Exit For '選定工作表
'    Next
'
'    If (.ActiveSheet.Name) <> "轉撥明細表" Then MsgBox "找不到轉撥明細表工作表!!", 16, "開啟檔案中止": GoTo endsub
'
'    'k = 1 '預設由第一列開始匯入
'
'    '若無來源欄位名稱
'    If strFieldName = "" Then
'        '取欄位名稱
'        For i = 1 To 255
'            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
'               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
'        Next i
'        'k = 2 '由第二列開始匯入
'    End If
'
'    For i = 1 To 255
'            If Trim(.Cells(i, 1)) = "單別-單號" Then k = i: Exit For
'    Next i
'
'    '切割欄位名稱
'    arrTmp = Split(strFieldName, Chr(9))
'
'    If UBound(arrTmp) < 1 Then Set rsMainT16_2 = Nothing: GoTo endsub
'
'    '建立Recordset
'    For i = 0 To UBound(arrTmp) - 1
'        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "轉撥明細表工作表，第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
'        rsMainT16_2.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
'    Next i
'
'    rsMainT16_2.CursorType = adOpenKeyset
'    rsMainT16_2.LockType = adLockOptimistic
'    rsMainT16_2.Open
'    rsMainT16.MoveFirst: S = ""
'    '寫入Recordset  '從這邊開始往下寫
'    Do While Len(RTrim(.Cells(k + 1, 6))) > 0   '挑準數量為空值則停止
'    rsMainT16_2.AddNew
'            For j = 1 To UBound(arrTmp)
'                If Len(Trim(rsMainT16_2("單別-單號").Value)) = 0 Then rsMainT16_2("單別-單號").Value = S
'                If Len(Trim(rsMainT16_2("品號").Value)) = 0 Then rsMainT16_2("品號").Value = Str_Sku
'                If j = UBound(arrTmp) Then    '送貨地址欄位
'                    rsMainT16.MoveFirst
'                    Do While Not rsMainT16.EOF
'                        Str_Sku = Trim(rsMainT16_2("品號").Value)
'                        If Trim(rsMainT16_2("單別-單號").Value) = Trim(rsMainT16("銷貨單別").Value) & "-" & Trim(rsMainT16("銷貨單號").Value) Then
'                            rsMainT16_2("送貨地址").Value = Trim(rsMainT16("送貨地址").Value): rsMainT16_2("件數").Value = Trim(rsMainT16("件數").Value): rsMainT16_2("指定日期").Value = Trim(rsMainT16("指定日期").Value): rsMainT16_2("備註").Value = Trim(rsMainT16("備註").Value): rsMainT16_2("單據日期").Value = Trim(rsMainT16("單據日期").Value): S = Trim(rsMainT16_2("單別-單號").Value): bl_Check = False: Exit Do
'                        End If
'                        rsMainT16.MoveNext
'                    Loop
'                    If bl_Check = True Then MsgBox "出貨總表查無:" & Trim(rsMainT16_2("單別-單號").Value) & "資料!", 64, "轉撥明細表匯入中止":  SSTab2.Enabled = True: Cmb_Import1.Enabled = True: GoTo endsub
'                Else
'                    rsMainT16_2(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
'                End If
'                 bl_Check = True
'            Next j
'    rsMainT16_2.Update
'    k = k + 1
'    Loop
'
'    If rsMainT16_2.RecordCount > 0 Then rsMainT16_2.MoveFirst
'
'End With
'
'Set dgMainT16_2.DataSource = rsMainT16_2
'
'If rsMainT16_2 Is Nothing Then
'    MsgBox "轉撥明細表查無資料!", 64, "Excel2Recordset"
'Else
'    SetDataGridColWidth Me.Caption, dgMainT16_2
'End If
'
''///////////////////////////////////////////////////////////////////////////////////匯入寄庫銷貨明細表//////////////////////////////////////////////////////////////////////////////////////////
'SSTab2.Tab = 3: S = "": Str_Sku = ""
'
''建立欄位名稱陣列
'strFieldName = "銷貨單號" & Chr(9) & "銷貨日期" & Chr(9) & "指定日期" & Chr(9) & "客戶代號" & Chr(9) & "客戶簡稱" & Chr(9) & "品號" & Chr(9) & "品名" & Chr(9) & "銷貨數量" & Chr(9) & "贈/備品量" & Chr(9) & "單位" & Chr(9) & "批號" & Chr(9) & "備註" & Chr(9) & "件數" & Chr(9) & "送貨地址" & Chr(9) '寄庫銷貨明細表
'If Right(filLocalFileT16.Path, 1) <> "\" Then
'    strFilePath = filLocalFileT16.Path & "\"
'Else
'    strFilePath = filLocalFileT16.Path
'End If
'
'Set rsMainT16_3 = New ADODB.Recordset
'
''Set MyXlsApp = CreateObject("Excel.Application")
'
'With MyXlsApp
'    .Visible = False
''    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '打開路徑
'
'    '尋找指定工作表
'    For i = 1 To .Sheets.Count
'        If (.Sheets(i).Name) = "寄庫銷貨明細表" Then .Sheets(i).Select: Exit For '選定工作表
'    Next
'
'    If (.ActiveSheet.Name) <> "寄庫銷貨明細表" Then MsgBox "找不到寄庫銷貨明細表工作表!!", 16, "開啟檔案中止": GoTo endsub
'
'    'k = 1 '預設由第一列開始匯入
'
'    '若無來源欄位名稱
'    If strFieldName = "" Then
'        '取欄位名稱
'        For i = 1 To 255
'            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
'               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
'        Next i
'        'k = 2 '由第二列開始匯入
'    End If
'
'    For i = 1 To 255
'            If Trim(.Cells(i, 1)) = "銷貨單號" Then k = i: Exit For
'    Next i
'
'    '切割欄位名稱
'    arrTmp = Split(strFieldName, Chr(9))
'
'    'Dim rsMainT15 As New ADODB.Recordset
'
'    If UBound(arrTmp) < 1 Then Set rsMainT16_3 = Nothing: GoTo endsub
'    '建立Recordset
'    For i = 0 To UBound(arrTmp) - 1
'        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "寄庫銷貨明細表工作表，第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
'        rsMainT16_3.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
'    Next i
'
'    rsMainT16_3.CursorType = adOpenKeyset
'    rsMainT16_3.LockType = adLockOptimistic
'    rsMainT16_3.Open
'    rsMainT16.MoveFirst: S = ""
'
'    '寫入Recordset  '從這邊開始往下寫
'    Do While Len(RTrim(.Cells(k + 1, 8))) > 0   '挑準數量為空值則停止
'    rsMainT16_3.AddNew
'
'            For j = 1 To UBound(arrTmp)
'                If Len(Trim(rsMainT16_3("銷貨單號").Value)) = 0 Then rsMainT16_3("銷貨單號").Value = S
'                If Len(Trim(rsMainT16_3("品號").Value)) = 0 Then rsMainT16_3("品號").Value = Str_Sku
'                If j = UBound(arrTmp) Then    '送貨地址欄位
'                    If Trim(rsMainT16_3("客戶代號")) = "1201011001" Or Trim(rsMainT16_3("客戶代號")) = "1201011002" Or Trim(rsMainT16_3("客戶代號")) = "1201011003" Or Trim(rsMainT16_3("客戶代號")) = "1201011006" Or Trim(rsMainT16_3("客戶代號")) = "1201011009" Or Trim(rsMainT16_3("客戶代號")) = "1201011004" Or Trim(rsMainT16_3("客戶代號")) = "1201011010" Or Trim(rsMainT16_3("客戶代號")) = "1201011011" Then
'                        S = Trim(rsMainT16_3("銷貨單號").Value): Str_Sku = Trim(rsMainT16_3("品號").Value)
'
'                        '透過客戶編號查出系統中的到貨地址,客戶名稱 ;因為主檔沒有寄庫銷貨明細表的細項
'                        Call Confirm_Recordset_Closed(tmp_Rs)
'                        str_SQL = "select full_name,address from trp01m where storerkey = 'LMYS01' and consigneekey = '" & Trim(rsMainT16_3("客戶代號").Value) & "'"
'                        tmp_Rs.Open str_SQL, cn
'
'                        If Not tmp_Rs.EOF Then rsMainT16_3("客戶簡稱") = Trim(tmp_Rs("full_name")): rsMainT16_3("送貨地址") = Trim(tmp_Rs("address"))
'                        tmp_Rs.Close
'
'                    Else
'                        MsgBox "發現非所屬客戶代號: " & Trim(rsMainT16_3("客戶代號")) & " 請確認是否為好事多、是否於客戶主檔新增資料、並請通知資訊部修改程式!", 64, "寄庫銷貨明細表匯入終止": GoTo endsub   '發現6個指定的客戶代號以外的明細,則停止匯入
'                    End If
'               Else
'                    rsMainT16_3(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j))):
'                End If
'            Next j
'    rsMainT16_3.Update
'    k = k + 1
'    Loop
'
'    If rsMainT16_3.RecordCount > 0 Then rsMainT16_3.MoveFirst
'
'End With
'
'SSTab2.Enabled = True: SSTab2.Tab = 0: Cmb_Import1.Enabled = True
'Set dgMainT16_3.DataSource = rsMainT16_3
'
'If rsMainT16_3 Is Nothing Then
'    MsgBox "寄庫銷貨明細表查無資料!", 64, "馬玉山訂單開啟"
'Else
'    SetDataGridColWidth Me.Caption, dgMainT16_3
'    MsgBox "此訂單檔共匯入:" & Chr(13) & "收貨總表:" & rsMainT16.RecordCount & "筆明細" & Chr(13) & "" & _
'                                          "銷貨明細表:" & rsMainT16_1.RecordCount & "筆明細" & Chr(13) & "" & _
'                                          "轉撥明細表:" & rsMainT16_2.RecordCount & "筆明細" & Chr(13) & "" & _
'                                          "寄庫銷貨明細表:" & rsMainT16_3.RecordCount & "筆明細" & Chr(13) & "" & _
'                                          "請確認筆數是否正確!", 64, "馬玉山訂單開啟"
'End If
'
''如果有出貨總表，其他三個工作表沒有資料則提示，但不擋
'If rsMainT16_1.RecordCount = 0 And rsMainT16_2.RecordCount = 0 And rsMainT16_3.RecordCount = 0 Then MsgBox "此訂單無細項資料，請確認此訂單是否正確!", vbCritical, "馬玉山訂單開啟"
'
'endsub:
'SSTab2.Enabled = True: Cmb_Import1.Enabled = True
'MyXlsApp.Quit: Set MyXlsApp = Nothing

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
Set rsMainT16_1 = Nothing: Set rsMainT16 = Nothing
SSTab2.Enabled = True: Cmb_Import1.Enabled = True
End Sub

Private Sub Cmb_Import2_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String, str_TmpSQL As String, S As String

On Error GoTo err_Handle
SSTab2.Tab = 1
Call DB_Connect_Self(cn_string) '建立新連線
'確認路徑是否帶"\"
If Right(filLocalFileT16.Path, 1) = "\" Then
    strFilePath = filLocalFileT16.Path
Else
    strFilePath = filLocalFileT16.Path & "\"
End If

'建立欄位名稱陣列
strFieldName = "銷貨單號" & Chr(9) & "銷貨日期" & Chr(9) & "指定日期" & Chr(9) & "客戶代號" & Chr(9) & "客戶簡稱" & Chr(9) & "品號" & Chr(9) & "品名" & Chr(9) & "銷貨數量" & Chr(9) & "贈/備品量" & Chr(9) & "單位" & Chr(9) & "批號" & Chr(9) & "備註" & Chr(9) & "件數" & Chr(9) & "送貨地址" & Chr(9)  '銷貨明細表
If Right(filLocalFileT16.Path, 1) <> "\" Then
    strFilePath = filLocalFileT16.Path & "\"
Else
    strFilePath = filLocalFileT16.Path
End If

Set rsMainT16_1 = New ADODB.Recordset

Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '打開路徑
    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = "銷貨明細表" Then .Sheets(i).Select: Exit For '選定工作表
    Next

    'k = 1 '預設由第一列開始匯入
    
    '若無來源欄位名稱
    If strFieldName = "" Then
        '取欄位名稱
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '由第二列開始匯入
    End If
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "銷貨單號" Then k = i: Exit For
    Next i
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT16 = Nothing: GoTo endsub
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsMainT16_1.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT16_1.CursorType = adOpenKeyset
    rsMainT16_1.LockType = adLockOptimistic
    rsMainT16_1.Open
    
    '寫入Recordset  '從這邊開始往下寫
    Do While Len(RTrim(.Cells(k + 1, 8))) > 0   '挑準數量為空值則停止
    If RTrim(.Cells(k + 1, 7)) = "運費" Then
    Else
        rsMainT16_1.AddNew
            For j = 1 To UBound(arrTmp)
                rsMainT16_1(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
            Next j
        rsMainT16_1.Update
    End If
    k = k + 1
    Loop
    
    If rsMainT16_1.RecordCount > 0 Then rsMainT16_1.MoveFirst

endsub:
End With
    
Set dgMainT17.DataSource = rsMainT16_1

If rsMainT16_1 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT17
    MsgBox "此工作表共 " & rsMainT16_1.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
End If
 Call Cmb_Import3_Click:
Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub Cmb_Import3_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String, str_TmpSQL As String

On Error GoTo err_Handle
Call DB_Connect_Self(cn_string) '建立新連線
SSTab2.Tab = 2
'確認路徑是否帶"\"
If Right(filLocalFileT16.Path, 1) = "\" Then
    strFilePath = filLocalFileT16.Path
Else
    strFilePath = filLocalFileT16.Path & "\"
End If

'建立欄位名稱陣列
strFieldName = "單別-單號" & Chr(9) & "單據日期" & Chr(9) & "轉入庫別" & Chr(9) & "品號" & Chr(9) & "品名" & Chr(9) & "轉撥數量" & Chr(9) & "單位" & Chr(9) & "批號" & Chr(9) & "備註" & Chr(9) '轉撥明細表
If Right(filLocalFileT16.Path, 1) <> "\" Then
    strFilePath = filLocalFileT16.Path & "\"
Else
    strFilePath = filLocalFileT16.Path
End If

Set rsMainT16_2 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, cboSheetT15, strFieldName, rsMainT15)
'''''''''''''''''''''''''''
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '打開路徑
    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = "轉撥明細表" Then .Sheets(i).Select: Exit For '選定工作表
    Next

    'k = 1 '預設由第一列開始匯入
    
    '若無來源欄位名稱
    If strFieldName = "" Then
        '取欄位名稱
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '由第二列開始匯入
    End If
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "單別-單號" Then k = i: Exit For
    Next i
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT16_2 = Nothing: GoTo endsub
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsMainT16_2.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT16_2.CursorType = adOpenKeyset
    rsMainT16_2.LockType = adLockOptimistic
    rsMainT16_2.Open
    
    '寫入Recordset  '從這邊開始往下寫
    Do While Len(RTrim(.Cells(k + 1, 6))) > 0   '挑準數量為空值則停止
    rsMainT16_2.AddNew
        For j = 1 To UBound(arrTmp)
            rsMainT16_2(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
        Next j
    rsMainT16_2.Update
    k = k + 1
    Loop
    
    If rsMainT16_2.RecordCount > 0 Then rsMainT16_2.MoveFirst

    'Call OffLineRecordset(rsMainT15, rs)
    
   ' rsMainT15.Close: Set rsMainT15 = Nothing
  
endsub:
End With

''以下將datagrid的資料存入佔存table中
'    '新增tabel
'    cn_Self.Execute "if object_id ('tempdb..##data2') is not null drop table tempdb..##data2 ", RowsAffect, adExecuteNoRecords
'    str_TmpSQL = "CREATE TABLE tempdb..##data2(單別單號 varchar(50),單據日期 varchar(50),轉入庫別 varchar(50),品號 varchar(50),品名 varchar(80),轉撥數量 varchar(50),單位 varchar(80),批號 varchar(50),備註 varchar(80))"
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
'    '放入datagrid資料到table
'    dgMainT18.Visible = False
'    Do While Not rsMainT16_2.EOF
'        '將匯入的excel資料 存入暫存的資料表##SkuCompare中
'        str_TmpSQL = "INSERT INTO tempdb..##data2 (單別單號,單據日期,轉入庫別,品號,品名,轉撥數量,單位,批號,備註) " & _
'                     "VALUES ('" & Trim(rsMainT16_2("單別-單號").Value) & "','" & Trim(rsMainT16_2("單據日期").Value) & "','" & Trim(rsMainT16_2("轉入庫別").Value) & "','" & _
'                     "" & Trim(rsMainT16_2("品號").Value) & "','" & Trim(rsMainT16_2("品名").Value) & "','" & Trim(rsMainT16_2("轉撥數量").Value) & "','" & Trim(rsMainT16_2("單位").Value) & "','" & _
'                     "" & Trim(rsMainT16_2("批號").Value) & "','" & Trim(rsMainT16_2("備註").Value) & "')"
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
'        rsMainT16_2.MoveNext
'    Loop
'    dgMainT18.Visible = True
    
Set dgMainT16_2.DataSource = rsMainT16_2

If rsMainT16_2 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT16_2
    MsgBox "此工作表共 " & rsMainT16_2.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
End If
 Call Cmb_Import4_Click:
Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub


Private Sub Cmb_Import4_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String, str_TmpSQL As String
On Error GoTo err_Handle
Call DB_Connect_Self(cn_string) '建立新連線
SSTab2.Tab = 3
'確認路徑是否帶"\"
If Right(filLocalFileT16.Path, 1) = "\" Then
    strFilePath = filLocalFileT16.Path
Else
    strFilePath = filLocalFileT16.Path & "\"
End If

'建立欄位名稱陣列
strFieldName = "銷貨單號" & Chr(9) & "銷貨日期" & Chr(9) & "指定日期" & Chr(9) & "客戶代號" & Chr(9) & "客戶簡稱" & Chr(9) & "品號" & Chr(9) & "品名" & Chr(9) & "銷貨數量" & Chr(9) & "贈/備品量" & Chr(9) & "單位" & Chr(9) & "批號" & Chr(9) & "備註" & Chr(9) & "件數" & Chr(9) '寄庫銷貨明細表
If Right(filLocalFileT16.Path, 1) <> "\" Then
    strFilePath = filLocalFileT16.Path & "\"
Else
    strFilePath = filLocalFileT16.Path
End If

Set rsMainT16_3 = New ADODB.Recordset
'Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, cboSheetT15, strFieldName, rsMainT15)
'''''''''''''''''''''''''''
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '打開路徑
    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = "寄庫銷貨明細表" Then .Sheets(i).Select: Exit For '選定工作表
    Next

    'k = 1 '預設由第一列開始匯入
    
    '若無來源欄位名稱
    If strFieldName = "" Then
        '取欄位名稱
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '由第二列開始匯入
    End If
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "銷貨單號" Then k = i: Exit For
    Next i
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT16_3 = Nothing: GoTo endsub
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsMainT16_3.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT16_3.CursorType = adOpenKeyset
    rsMainT16_3.LockType = adLockOptimistic
    rsMainT16_3.Open
    
    '寫入Recordset  '從這邊開始往下寫
    Do While Len(RTrim(.Cells(k + 1, 8))) > 0   '挑準數量為空值則停止
    rsMainT16_3.AddNew
        For j = 1 To UBound(arrTmp)
            rsMainT16_3(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
        Next j
    rsMainT16_3.Update
    k = k + 1
    Loop
    
    If rsMainT16_3.RecordCount > 0 Then rsMainT16_3.MoveFirst
  
endsub:
End With


''以下將datagrid的資料存入佔存table中
'    '新增tabel
'    cn_Self.Execute "if object_id ('tempdb..##data3') is not null drop table tempdb..##data3 ", RowsAffect, adExecuteNoRecords
'    str_TmpSQL = "CREATE TABLE tempdb..##data3(銷貨單號 varchar(50),銷貨日期 varchar(50),指定日期 varchar(50),客戶代號 varchar(50),客戶簡稱 varchar(50),品號 varchar(50),品名 varchar(80),銷貨數量 varchar(50),贈備品量 varchar(50),單位 varchar(50),批號 varchar(50),備註 varchar(80),件數 varchar(50))"
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
'    '放入datagrid資料到table
'    dgMainT19.Visible = False
'    Do While Not rsMainT16_3.EOF
'        '將匯入的excel資料 存入暫存的資料表##SkuCompare中
'        str_TmpSQL = "INSERT INTO tempdb..##data3 (銷貨單號,銷貨日期,指定日期,客戶代號,客戶簡稱,品號,品名,銷貨數量,贈備品量,單位,批號,備註,件數) " & _
'                     "VALUES ('" & Trim(rsMainT16_3("銷貨單號").Value) & "','" & Trim(rsMainT16_3("銷貨日期").Value) & "','" & Trim(rsMainT16_3("指定日期").Value) & "','" & _
'                     "" & Trim(rsMainT16_3("客戶代號").Value) & "','" & Trim(rsMainT16_3("客戶簡稱").Value) & "','" & Trim(rsMainT16_3("品號").Value) & "','" & Trim(rsMainT16_3("品名").Value) & "','" & _
'                     "" & Trim(rsMainT16_3("銷貨數量").Value) & "','" & Trim(rsMainT16_3("贈/備品量").Value) & "','" & Trim(rsMainT16_3("單位").Value) & "','" & Trim(rsMainT16_3("批號").Value) & "','" & Trim(rsMainT16_3("備註").Value) & "','" & Trim(rsMainT16_3("件數").Value) & "')"
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
'        rsMainT16_3.MoveNext
'    Loop
'    dgMainT19.Visible = True
    
Set dgMainT16_3.DataSource = rsMainT16_3

If rsMainT16_3 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT16_3
    MsgBox "此工作表共 " & rsMainT16_3.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub



Private Sub Cmd_Impdata_Click()
Dim bl_OrderCheck As Boolean '檢查訂單是否出現在明細中
Dim str_Storerkey As String, str_Priority As String, Str_Address As String, Str_Lot05 As String
Dim Int_RC As Integer: Dim Int_C As Integer: Dim Int_i As Integer: Dim Int_otqty As Integer
Dim Str_AllOrderkey As String
Str_AllOrderkey = ""
Int_RC = 0: Int_C = 0: Int_i = 0 '計算訂單類別的筆數
Int_otqty = 0 '計算訂單件數
str_Storerkey = "LMBO01"

bl_OrderCheck = False
If rsMainT16 Is Nothing Then Exit Sub
If rsMainT16.EOF Then Exit Sub
If rsMainT16_1 Is Nothing Then Exit Sub
If rsMainT16_1.EOF Then Exit Sub

'GoTo copy:
On Error GoTo err_Handle

SSTab2.Enabled = False: Cmd_Impdata.Enabled = False

'資料檢驗--判斷檔案是否已轉入

Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where storerkey = '" & str_Storerkey & "' and rtrim(updatesource)='" & Str_updatesource1 & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

'檢查主檔中的訂單號碼，有無出現在訂單明細中
rsMainT16.MoveFirst: rsMainT16_1.MoveFirst
Do While Not rsMainT16.EOF
        Do While Not rsMainT16_1.EOF
            If RTrim(rsMainT16.Fields("訂單號碼")) + RTrim(rsMainT16.Fields("訂單種類")) = RTrim(rsMainT16_1.Fields("訂單號碼")) + RTrim(rsMainT16_1.Fields("訂單種類")) Then
            '正確，有出現資料。
                bl_OrderCheck = True
                Exit Do
            End If
            rsMainT16_1.MoveNext
        Loop
    If bl_OrderCheck = False Then MsgBox "訂單號碼+訂單種類:" & RTrim(rsMainT16.Fields("訂單號碼")) & RTrim(rsMainT16.Fields("訂單種類")) & " 未出現在明細檔中，請確認訂單檔資料是否正確，訂單轉入中止", vbOKOnly + vbCritical, "毛寶訂單匯入": SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: bl_OrderCheck = False: Exit Sub
    bl_OrderCheck = False
    rsMainT16.MoveNext
    rsMainT16_1.MoveFirst
Loop

rsMainT16_1.MoveFirst
Do While Not rsMainT16_1.EOF
    If RTrim(rsMainT16_1.Fields("允收期")) = "1" Or RTrim(rsMainT16_1.Fields("允收期")) = "0" Then
    '允收期不可以為0或1
    MsgBox "訂單號碼:" & Trim(rsMainT16_1("訂單號碼")) & "，允收期=" & Trim(rsMainT16("允收期")) & "，不可為1或0值，訂單匯入中止", vbCritical + vbOKOnly
    Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: Exit Sub
    End If
    rsMainT16_1.MoveNext
Loop
rsMainT16_1.MoveFirst

'檢查預計日期不可小於今日
    rsMainT16.MoveFirst
    Do While Not rsMainT16.EOF
        If Val(Left(Trim(rsMainT16("預計日期")), 3)) + 1911 & "/" & Mid(Trim(rsMainT16("預計日期")), 4, 2) & "/" & Right(Trim(rsMainT16("預計日期")), 2) < Format(Now, "YYYY/MM/DD") Then
            MsgBox "預計日期:" & Trim(rsMainT16("預計日期")) & "小於今日，請確認預計日期是否錯誤，訂單匯入中止", vbCritical + vbOKOnly
            Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: Exit Sub
        End If
        rsMainT16.MoveNext
    Loop
    
'檢查提貨倉不等於 154,160,161,162,163,165
    rsMainT16.MoveFirst
    Do While Not rsMainT16.EOF
        If Trim(rsMainT16("提貨倉")) <> "154" And Trim(rsMainT16("提貨倉")) <> "160" And Trim(rsMainT16("提貨倉")) <> "161" And Trim(rsMainT16("提貨倉")) <> "162" And Trim(rsMainT16("提貨倉")) <> "163" And Trim(rsMainT16("提貨倉")) <> "165" Then
            MsgBox "提貨倉:" & Trim(rsMainT16("提貨倉")) & "，不等於毛寶使用倉別，請確認格式有無問題!", vbCritical + vbOKOnly
            Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: Exit Sub
        End If
        rsMainT16.MoveNext
    Loop
    rsMainT16.MoveFirst

'檢查提貨倉162，單別不等於CO和SC的
    Do While Not rsMainT16.EOF
        If (Trim(rsMainT16("提貨倉")) <> "162" And Trim(rsMainT16("訂單種類")) = "SC") Or (Trim(rsMainT16("提貨倉")) <> "162" And Trim(rsMainT16("訂單種類")) = "CO") Then
            MsgBox "退貨訂單種類:" & Trim(rsMainT16("訂單種類")) & "的提貨倉:" & Trim(rsMainT16("提貨倉")) & "不等於162，請確認資料有無問題!", vbCritical + vbOKOnly
            Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: Exit Sub
        End If
        rsMainT16.MoveNext
    Loop
    rsMainT16.MoveFirst
'檢查品號是否存在
    rsMainT16_1.MoveFirst
    Do While Not rsMainT16_1.EOF
        '檢查是否有此品號
        str_SQL = "select sku from " & strWMSDB & "..sku(nolock) where storerkey = '" & str_Storerkey & "' and sku = '" & Trim(rsMainT16_1("產品編號")) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn

        If tmp_Rs.EOF Then
            '無此品號
            MsgBox "系統找不到品號:" & Trim(rsMainT16_1("產品編號")) & "的資料，請先建立商品主檔資料，匯入中止", vbCritical + vbOKOnly, "品號檢查"
            Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: Exit Sub
        End If
        rsMainT16_1.MoveNext
    Loop
    
rsMainT16.MoveFirst: rsMainT16_1.MoveFirst

Tran_Level = cn.BeginTrans: Cmd_Impdata.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String, strInvoiceDate As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String

'取最後客戶編號
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = '" & str_Storerkey & "' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))

tmp_Rs.Close
'開始匯入
Do While Not rsMainT16.EOF
    DoEvents: DoEvents
        '新增訂單資料
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        Str_Address = RTrim(rsMainT16.Fields("送貨地址"))
        
        '檢查是否有此客戶編號
        str_SQL = "select top 1 consigneekey from trp01m where storerkey = '" & str_Storerkey & "' and rtrim(consigneekey) = '" & RTrim(rsMainT16.Fields("客戶編號")) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn

        If tmp_Rs.EOF Then
            '無此客戶編號則新增
            '新增客戶主檔
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
            " values('" & str_Storerkey & "','','" & Trim(rsMainT16("客戶編號")) & "','" & myExCharFilter(Trim(rsMainT16("客戶名稱"))) & "','" & myExCharFilter(Trim(rsMainT16("客戶簡稱"))) & "','" & myExCharFilter(Trim(rsMainT16("連絡人"))) & "','" & myExCharFilter(Trim(rsMainT16("電話"))) & "','" & Str_Address & "','','" & User_id & "','" & User_id & "' ) ", RowsAffect, adExecuteNoRecords
            '紀錄新增之客戶編號
            strNewConsigneekey = strNewConsigneekey & Trim(rsMainT16("客戶編號")) & "','"
            strConsigneeKey = Trim(rsMainT16("客戶編號"))
        Else
            '比對客戶名稱，簡稱與到貨地址是否相符
            Call Confirm_Recordset_Closed(rsTmp)
            str_SQL = "select consigneekey from trp01m(nolock) " & _
                        "where storerkey = '" & str_Storerkey & "' and full_name = '" & myExCharFilter(Trim(rsMainT16("客戶名稱"))) & "' and short_name = '" & myExCharFilter(Trim(rsMainT16("客戶簡稱"))) & "' " & _
                        "and rtrim(address) = '" & Str_Address & "' "
            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn

                If rsTmp.EOF Then
                    '聯絡人、電話與到貨地址不符
                    intTmp = intTmp + 1
                    strConsigneeKey = "BEST" & Format(intTmp, "000000") '待確認BEST

                    '新增客戶主檔
                    cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
                    " values('" & str_Storerkey & "','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16("客戶名稱"))) & "','" & myExCharFilter(Trim(rsMainT16("客戶簡稱"))) & "','" & myExCharFilter(Trim(rsMainT16("連絡人"))) & "','" & myExCharFilter(Trim(rsMainT16("電話"))) & "','" & Str_Address & "','','" & User_id & "','" & User_id & "' ) ", RowsAffect, adExecuteNoRecords
            
                    '紀錄新增之客戶編號
                    strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
                Else '相符沿用舊客編
                    strConsigneeKey = Trim(rsTmp("consigneekey"))
                    blCustomerMatch = True

                End If
            rsTmp.Close
        End If
        tmp_Rs.Close

        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select ExternOrderKey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT16("訂單種類"))) & myExCharFilter(Trim(rsMainT16("訂單號碼"))) & "' and externordertype = '" & myExCharFilter(Trim(rsMainT16("訂單種類"))) & "' and storerkey = '" & str_Storerkey & "' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            '紀錄所有orderkey,最後一次批次更新packkey
            Str_AllOrderkey = Str_AllOrderkey & "'" & str_Orderkey & "',"
            '配送倉別判斷
            strFacility = "佰事達北倉"

            If Len(Trim(rsMainT16("預計日期"))) = 0 Then    '如果沒有指定日期,則帶隔日一天
                strDate = Format(Now + 1, "YYYY/MM/DD")
            Else
                strDate = Val(Left(Trim(rsMainT16("預計日期")), 3)) + 1911 & "/" & Mid(Trim(rsMainT16("預計日期")), 4, 2) & "/" & Right(Trim(rsMainT16("預計日期")), 2)
            End If
            
            strOrderDate = Val(Left(Trim(rsMainT16("訂單日期")), 3)) + 1911 & "/" & Mid(Trim(rsMainT16("訂單日期")), 4, 2) & "/" & Right(Trim(rsMainT16("訂單日期")), 2)
            

            
            Dim intPointer As Integer
            intPointer = 1
            
            '訂單類別轉換
            If myExCharFilter(Trim(rsMainT16("提貨倉"))) = "154" Then
                str_Priority = "A2B"
            Else
                If myExCharFilter(Trim(rsMainT16("訂單種類"))) = "CO" Or myExCharFilter(Trim(rsMainT16("訂單種類"))) = "SC" Then
                    str_Priority = "R"
                Else
                    '需用提貨倉攔位判斷是否維I或是A2B
                    str_Priority = "I"
                End If
            End If
            '寫入毛寶專用Table，紀錄訂單所有資訊
            If Len(Trim(rsMainT16("發票日期"))) = 0 Then
                 str_SQL = "insert CustOrders(Storerkey,orderkey,BranchId,ExternOrderkey,OrderDate,Invoice,InvoiceCheck,InvoiceDate,Consigneekey, " & _
                        "Full_Name,SalesCode,COD,Address,Coupled,VAT,Allowance,QuantityAllowance,SpecialAllowance, " & _
                        "CashAllowance,Amount,NetAmount,Tax,Notes,CustOrderkey,InvoiceCode,OrderCode, " & _
                        "LogisticsCode , DeliveryCode, OrderType, PaidMARK, Contact, Phone1, SalesName, LeaderName, " & _
                        "Address2,DeliveryDate,Freight,Payment,SalesPhone,EInvoiceMark,TotalWeight,Credit_Last4,Cash, " & _
                        "InvoicePrint , Phone2, ExternNumber, City, Administration, Stairs, CrossCode, Storage, InvoiceArea,Short_name,keyinuser,ConnectOrderkey,addwho,updatesource) " & _
                        "values ('LMBO01','" & str_Orderkey & "','" & Trim(rsMainT16("分分司代號")) & "','" & Trim(rsMainT16("訂單號碼")) & "','" & strOrderDate & "','" & _
                        Trim(rsMainT16("發票號碼")) & "','" & Trim(rsMainT16("發票號碼檢查碼")) & "',null,'" & _
                        Trim(rsMainT16("客戶編號")) & "','" & Trim(rsMainT16("客戶名稱")) & "','" & Trim(rsMainT16("業代代號")) & "','" & _
                        Trim(rsMainT16("下貨收現")) & "','" & Trim(rsMainT16("送貨地址")) & "','" & Trim(rsMainT16("聯式")) & "','" & _
                        Trim(rsMainT16("統一編號")) & "','" & Trim(rsMainT16("折讓金額")) & "','" & Trim(rsMainT16("數量折讓金額")) & "','" & _
                        Trim(rsMainT16("特別折讓金額")) & "','" & Trim(rsMainT16("現金折讓")) & "','" & Trim(rsMainT16("貨款")) & "','" & _
                        Trim(rsMainT16("稅前金額")) & "','" & Trim(rsMainT16("稅額")) & "','" & Trim(rsMainT16("備註")) & "','" & _
                        Trim(rsMainT16("客戶訂單編號")) & "','" & Trim(rsMainT16("隨貨附發票碼")) & "','" & Trim(rsMainT16("隨貨附訂單碼")) & "','" & _
                        Trim(rsMainT16("計算物流費")) & "','" & Trim(rsMainT16("送貨否")) & "','" & Trim(rsMainT16("訂單種類")) & "','" & _
                        Trim(rsMainT16("實收量處理MARK")) & "','" & Trim(rsMainT16("連絡人")) & "','" & Trim(rsMainT16("電話")) & "','" & _
                        Trim(rsMainT16("業代姓名")) & "','" & Trim(rsMainT16("主管姓名")) & "','" & Trim(rsMainT16("指送客戶")) & "','" & _
                        strDate & "','" & Trim(rsMainT16("運費")) & "','" & Trim(rsMainT16("付款方式")) & "','" & Trim(rsMainT16("業務手機")) & "','" & _
                        Trim(rsMainT16("是否為電子發票")) & "','" & Trim(rsMainT16("總重量")) & "','" & Trim(rsMainT16("信卡後4碼")) & "','" & _
                        Trim(rsMainT16("代收貨款")) & "','" & Trim(rsMainT16("發票列印方式")) & "','" & Trim(rsMainT16("電話2")) & "','" & Trim(rsMainT16("統計對象")) & "','" & _
                        Trim(rsMainT16("縣市別")) & "','" & Trim(rsMainT16("行政區")) & "','" & Trim(rsMainT16("樓層")) & "','" & Trim(rsMainT16("越庫訂單")) & "','" & Trim(rsMainT16("提貨倉")) & "','" & Trim(rsMainT16("稅區/稅率")) & "','" & Trim(rsMainT16("客戶簡稱")) & "','" & Trim(rsMainT16("訂單窗口")) & "','" & Trim(rsMainT16("關聯訂單號碼")) & "','" & User_id & "','" & Str_updatesource1 & "')"

            Else
                strInvoiceDate = Val(Left(Trim(rsMainT16("發票日期")), 3)) + 1911 & "/" & Mid(Trim(rsMainT16("發票日期")), 4, 2) & "/" & Right(Trim(rsMainT16("發票日期")), 2)
                str_SQL = "insert CustOrders(Storerkey,orderkey,BranchId,ExternOrderkey,OrderDate,Invoice,InvoiceCheck,InvoiceDate,Consigneekey, " & _
                        "Full_Name,SalesCode,COD,Address,Coupled,VAT,Allowance,QuantityAllowance,SpecialAllowance, " & _
                        "CashAllowance,Amount,NetAmount,Tax,Notes,CustOrderkey,InvoiceCode,OrderCode, " & _
                        "LogisticsCode , DeliveryCode, OrderType, PaidMARK, Contact, Phone1, SalesName, LeaderName, " & _
                        "Address2,DeliveryDate,Freight,Payment,SalesPhone,EInvoiceMark,TotalWeight,Credit_Last4,Cash, " & _
                        "InvoicePrint , Phone2, ExternNumber, City, Administration, Stairs, CrossCode, Storage, InvoiceArea,Short_name,keyinuser,ConnectOrderkey,addwho,updatesource) " & _
                        "values ('LMBO01','" & str_Orderkey & "','" & Trim(rsMainT16("分分司代號")) & "','" & Trim(rsMainT16("訂單號碼")) & "','" & strOrderDate & "','" & _
                        Trim(rsMainT16("發票號碼")) & "','" & Trim(rsMainT16("發票號碼檢查碼")) & "','" & strInvoiceDate & "','" & _
                        Trim(rsMainT16("客戶編號")) & "','" & Trim(rsMainT16("客戶名稱")) & "','" & Trim(rsMainT16("業代代號")) & "','" & _
                        Trim(rsMainT16("下貨收現")) & "','" & Trim(rsMainT16("送貨地址")) & "','" & Trim(rsMainT16("聯式")) & "','" & _
                        Trim(rsMainT16("統一編號")) & "','" & Trim(rsMainT16("折讓金額")) & "','" & Trim(rsMainT16("數量折讓金額")) & "','" & _
                        Trim(rsMainT16("特別折讓金額")) & "','" & Trim(rsMainT16("現金折讓")) & "','" & Trim(rsMainT16("貨款")) & "','" & _
                        Trim(rsMainT16("稅前金額")) & "','" & Trim(rsMainT16("稅額")) & "','" & Trim(rsMainT16("備註")) & "','" & _
                        Trim(rsMainT16("客戶訂單編號")) & "','" & Trim(rsMainT16("隨貨附發票碼")) & "','" & Trim(rsMainT16("隨貨附訂單碼")) & "','" & _
                        Trim(rsMainT16("計算物流費")) & "','" & Trim(rsMainT16("送貨否")) & "','" & Trim(rsMainT16("訂單種類")) & "','" & _
                        Trim(rsMainT16("實收量處理MARK")) & "','" & Trim(rsMainT16("連絡人")) & "','" & Trim(rsMainT16("電話")) & "','" & _
                        Trim(rsMainT16("業代姓名")) & "','" & Trim(rsMainT16("主管姓名")) & "','" & Trim(rsMainT16("指送客戶")) & "','" & _
                        strDate & "','" & Trim(rsMainT16("運費")) & "','" & Trim(rsMainT16("付款方式")) & "','" & Trim(rsMainT16("業務手機")) & "','" & _
                        Trim(rsMainT16("是否為電子發票")) & "','" & Trim(rsMainT16("總重量")) & "','" & Trim(rsMainT16("信卡後4碼")) & "','" & _
                        Trim(rsMainT16("代收貨款")) & "','" & Trim(rsMainT16("發票列印方式")) & "','" & Trim(rsMainT16("電話2")) & "','" & Trim(rsMainT16("統計對象")) & "','" & _
                        Trim(rsMainT16("縣市別")) & "','" & Trim(rsMainT16("行政區")) & "','" & Trim(rsMainT16("樓層")) & "','" & Trim(rsMainT16("越庫訂單")) & "','" & Trim(rsMainT16("提貨倉")) & "','" & Trim(rsMainT16("稅區/稅率")) & "','" & Trim(rsMainT16("客戶簡稱")) & "','" & Trim(rsMainT16("訂單窗口")) & "','" & Trim(rsMainT16("關聯訂單號碼")) & "','" & User_id & "','" & Str_updatesource1 & "')"
    
            End If
           
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            '寫入Orders
            If str_Priority = "A2B" Then
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,b_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externordertype,cash) " & _
                "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT16("訂單種類")) & myExCharFilter(Trim(rsMainT16("訂單號碼"))) & "','" & str_Priority & "','" & str_Storerkey & "','" & strOrderDate & "','" & strDate & "','" & strFacility & "','LMBO01-154" & _
                "','" & myExCharFilter(Trim(rsMainT16("客戶名稱"))) & "','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16("連絡人"))) & "','','','" & myExCharFilter(Trim(rsMainT16("電話"))) & "','','" & myExCharFilter(Trim(GetWord(Str_Address, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(Str_Address, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT16("備註"))) & "','" & Str_updatesource1 & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT16("發票號碼"))) & "','" & myExCharFilter(Trim(rsMainT16("訂單種類"))) & "','" & myExCharFilter(Trim(rsMainT16("代收貨款"))) & "') "
            Else
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externordertype,cash) " & _
                "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT16("訂單種類")) & myExCharFilter(Trim(rsMainT16("訂單號碼"))) & "','" & str_Priority & "','" & str_Storerkey & "','" & strOrderDate & "','" & strDate & "','" & strFacility & "','" & _
                strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16("客戶名稱"))) & "','" & myExCharFilter(Trim(rsMainT16("連絡人"))) & "','','','" & myExCharFilter(Trim(rsMainT16("電話"))) & "','','" & myExCharFilter(Trim(GetWord(Str_Address, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(Str_Address, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT16("備註"))) & "','" & Str_updatesource1 & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT16("發票號碼"))) & "','" & myExCharFilter(Trim(rsMainT16("訂單種類"))) & "','" & myExCharFilter(Trim(rsMainT16("代收貨款"))) & "') "
            End If
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1


            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
            'If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & Str_storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT16("訂單種類")) & myExCharFilter(Trim(rsMainT16("訂單號碼"))) & "','" 'Trim(rsMainT16("訂單種類"))
            blDuplicationOrder = True

        End If
        
        '訂單重複檢查
        If blDuplicationOrder = False Then
            rsMainT16_1.Filter = "訂單號碼 = '" & rsMainT16.Fields("訂單號碼") & "' and 訂單種類 = '" & rsMainT16.Fields("訂單種類") & "'"
            rsMainT16_1.Sort = "行號"
            rsMainT16_1.MoveFirst
            Do While Not rsMainT16_1.EOF
                '增加明細
                int_orderlinenuber = int_orderlinenuber + 1
                lngCasecnt = 1

                '效期換算 lot05 = 到貨日+(允收期X有效天數)
                    '取箱包轉換率
'                    str_SQL = "select susr2=isnull(susr2,0) from " & strWMSDB & "..sku where sku = '" & myExCharFilter(Trim(rsMainT16_1("產品編號"))) & "'"
'
'                    Call Confirm_Recordset_Closed(tmp_Rs)
'                    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'                    Str_Lot05 = Format(strDate + (Val(rsMainT16_1.Fields("允收期")) * Val(tmp_Rs.Fields("susr2"))), "YYYYMMDD")
'                    tmp_Rs.Close
                    
                intQTY = Abs(Val(rsMainT16_1("訂貨量")))
                strLot06 = RTrim(rsMainT16.Fields("提貨倉"))
                
                '紀錄毛寶專用訂單明細
                str_SQL = "insert CustOrderdetail(Storerkey,orderkey,orderlinenumber,ExternOrderkey,Sku,Descr,OriginalQty,UnitNetPrice,NetPrice,UnitGrossPrice,GrossPrice,RefusalQty,BarCode,Externlineno,UOM,Ordertype,InvoicePCode,Acceptance,addwho) " & _
                "values('LMBO01','" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Trim(rsMainT16_1.Fields("訂單號碼")) & "','" & Trim(rsMainT16_1.Fields("產品編號")) & "','" & Trim(rsMainT16_1.Fields("產品名稱")) & _
                "','" & Trim(rsMainT16_1.Fields("訂貨量")) & "','" & Trim(rsMainT16_1.Fields("單價(未稅)")) & "','" & Trim(rsMainT16_1.Fields("訂貨金額(未稅)")) & _
                "','" & Trim(rsMainT16_1.Fields("單價(含稅)")) & "','" & Trim(rsMainT16_1.Fields("訂貨金額(含稅)")) & "','" & Trim(rsMainT16_1.Fields("訂貨量-實收量")) & _
                "','" & Trim(rsMainT16_1.Fields("國際條碼")) & "','" & Trim(rsMainT16_1.Fields("行號")) & "','" & Trim(rsMainT16_1.Fields("單位")) & _
                "','" & Trim(rsMainT16_1.Fields("訂單種類")) & "','" & Trim(rsMainT16_1.Fields("發票明細列印否")) & "','" & Trim(rsMainT16_1.Fields("允收期")) & "','" & User_id & "')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                '訂單明細資料新增
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice)" & _
                "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT16_1("行號"))) & "','" & Trim(rsMainT16("訂單種類")) & myExCharFilter(Trim(rsMainT16_1("訂單號碼"))) & "','" & myExCharFilter(Trim(rsMainT16_1("產品編號"))) & "','" & str_Storerkey & "'," & _
                "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT16_1("單位"))) & "','0')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
                int_OrderLine = int_OrderLine + 1
                rsMainT16_1.MoveNext
            Loop
        End If

        rsMainT16.MoveNext
        rsMainT16_1.MoveFirst
Loop

'批次更新packkey
If Str_AllOrderkey <> "" Then '加入Str_AllOrderkey <> "" 判斷 by Gemini @20160704
    str_SQL = "update orderdetail " & _
    "Set orderdetail.packkey = sku.packkey " & _
    "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
    "where orderkey in (" & Mid(Str_AllOrderkey, 1, Len(Str_AllOrderkey) - 1) & ") "

    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

End If

cn.Execute "exec gs_ordersupdate 'LMBO01'", RowsAffect, adExecuteNoRecords
cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入 " & int_OrderLine & " 筆明細" & Chr(13) & Chr(13) & "TMS單號： " & strOrderKeyS & "~" & str_Orderkey & Chr(13)
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "，檔案 " & Str_updatesource1)

'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & Str_updatesource1 & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 品號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "'"

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption

    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\" & str_Storerkey & "\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing

End If

copy:

'備份檔案到本機
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & Str_updatesource1) = "" Then
    FileCopy strFileHeader, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & Str_updatesource1
    FileCopy strFileDetail, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & Str_updatesource2
    FileCopy strFileHeader, "\\192.168.200.200\Backup$\" & str_Storerkey & "\Orders\Backup\" & Str_updatesource1
    FileCopy strFileDetail, "\\192.168.200.200\Backup$\" & str_Storerkey & "\Orders\Backup\" & Str_updatesource2
Else
    FileCopy strFileHeader, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(Str_updatesource1, ".", 0) & "." & mySplit(Str_updatesource1, ".", -1)
    FileCopy strFileDetail, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(Str_updatesource2, ".", 0) & "." & mySplit(Str_updatesource2, ".", -1)
    FileCopy strFileHeader, "\\192.168.200.200\Backup$\" & str_Storerkey & "\Orders\Backup\" & mySplit(Str_updatesource1, ".", 0) & "." & mySplit(Str_updatesource1, ".", -1)
    FileCopy strFileDetail, "\\192.168.200.200\Backup$\" & str_Storerkey & "\Orders\Backup\" & mySplit(Str_updatesource2, ".", 0) & "." & mySplit(Str_updatesource2, ".", -1)
End If

Kill strFileHeader
Kill strFileDetail

''斷開目前連線連線
'LCDisConnect "192.168.200.200\IPC$"
'LCDisConnect "bestprepares\IPC$"
'
''重新連線公共資料夾
'LCConnect "bestprepares", "share", "share"
''LCConnect "192.168.200.200", "share", "share"

filLocalFileT16.Refresh:
Screen.MousePointer = 0: Cmd_Impdata.Enabled = True: SSTab2.Enabled = True
Exit Sub




'以下為馬玉山訂單匯入的程式碼
'Dim Int_RC As Integer: Dim Int_C As Integer: Dim Int_I As Integer: Dim Int_otqty As Integer
'Int_RC = 0: Int_C = 0: Int_I = 0 '計算訂單類別的筆數
'Int_otqty = 0 '計算訂單件數
'
'If rsMainT16 Is Nothing Then Exit Sub
'If rsMainT16.EOF Then Exit Sub
'
'On Error GoTo err_Handle
'SSTab2.Enabled = False: Cmd_Impdata.Enabled = False
'strTranFileName = filLocalFileT16.Path & "\" & filLocalFileT16.FileName
'
''資料檢驗--判斷檔案是否已轉入
'Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT16.FileName & "' "
'
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'If tmp_Rs.EOF = False Then SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
'tmp_Rs.Close
'
'Dim arrTmp
'
'If rsMainT16_1.RecordCount = 0 Or rsMainT16_1 Is Nothing Then
'Else
'rsMainT16_1.MoveFirst
'Do While Not rsMainT16_1.EOF
'    '到貨日期檢查
'    If Len(Trim(rsMainT16_1("指定日期"))) = 0 Then
'    Else
'        arrTmp = Split(Trim(rsMainT16_1("指定日期")), "/")
'        If Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "指定日期小於今日，訂單轉入終止!", 16, Me.Caption: SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Exit Sub
'    End If
'
'    '數量檢查
'    If Val(rsMainT16_1("銷貨數量")) + Val(rsMainT16_1("贈/備品量")) < 1 Then
'        MsgBox "訂單數量小於1，" & Trim(rsMainT16_1("銷貨單號")) & "-品號：" & Trim(rsMainT16_1("品號")) & "(" & Trim(rsMainT16_1("品名")) & ")，訂單轉入終止!!請確認!!", , "訂單檔匯入": SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Exit Sub
'        Exit Sub
'    End If
'
''    資料檢驗 --判斷SKU是否存在
'    If Trim(rsMainT16_1("品號")) <> "60400119" Then
'        str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT16_1("品號")) & "' and Storerkey = 'LMYS01' "
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If tmp_Rs.EOF Then
'            MsgBox "訂單發現新品號 (" & Trim(rsMainT16_1("品號")) & " ) " & Trim(rsMainT16_1("品名")) & "，訂單轉入終止!!": Cmd_Impdata.Enabled = True: Screen.MousePointer = 0
'            SSTab2.Enabled = True: Cmd_Impdata.Enabled = True
'            Exit Sub
'        End If
'
'    End If
'
'    rsMainT16_1.MoveNext
'Loop
'rsMainT16_1.MoveFirst
'End If
'
'If rsMainT16_2.RecordCount = 0 Or rsMainT16_2 Is Nothing Then
'Else
'rsMainT16_2.MoveFirst
'Do While Not rsMainT16_2.EOF
'    '到貨日期檢查
'    If Len(Trim(rsMainT16_2("指定日期"))) = 0 Then
'    Else
'        arrTmp = Split(Trim(rsMainT16_2("指定日期")), "/")
'        If Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "指定日期小於今日，訂單轉入終止!", 16, Me.Caption: SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Exit Sub
'    End If
'
'    '數量檢查
'    If Val(rsMainT16_2("轉撥數量")) < 1 Then
'        MsgBox "訂單數量小於1，" & Trim(rsMainT16_2("單別-單號")) & "-品號：" & Trim(rsMainT16_2("品號")) & "(" & Trim(rsMainT16_2("品名")) & ")，訂單轉入終止!!請確認!!", , "訂單檔匯入": SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Exit Sub
'        Exit Sub
'    End If
'
''    資料檢驗 --判斷SKU是否存在
'    If Trim(rsMainT16_2("品號")) <> "60400119" Then
'        str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT16_2("品號")) & "' and Storerkey = 'LMYS01' "
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If tmp_Rs.EOF Then
'            MsgBox "訂單發現新品號 (" & Trim(rsMainT16_2("品號")) & " ) " & Trim(rsMainT16_2("品名")) & "，訂單轉入終止!!": Cmd_Impdata.Enabled = True: Screen.MousePointer = 0
'            SSTab2.Enabled = True: Cmd_Impdata.Enabled = True
'            Exit Sub
'        End If
'    End If
'
'    rsMainT16_2.MoveNext
'Loop
'rsMainT16_2.MoveFirst
'End If
'
'If rsMainT16_3.RecordCount = 0 Or rsMainT16_3 Is Nothing Then
'Else
'rsMainT16_3.MoveFirst
'Do While Not rsMainT16_3.EOF
'    '到貨日期檢查
'    If Len(Trim(rsMainT16_3("指定日期"))) = 0 Then
'    Else
'        arrTmp = Split(Trim(rsMainT16_3("指定日期")), "/")
'        If Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "指定日期小於今日，訂單轉入終止!", 16, Me.Caption: SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Exit Sub
'    End If
'
'    '數量檢查
'    If Val(rsMainT16_3("銷貨數量")) + Val(rsMainT16_3("贈/備品量")) < 1 Then
'        MsgBox "訂單數量小於1，" & Trim(rsMainT16_3("銷貨單號")) & "-品號：" & Trim(rsMainT16_3("品號")) & "(" & Trim(rsMainT16_3("品名")) & ")，訂單轉入終止!!請確認!!", , "訂單檔匯入": SSTab2.Enabled = True: Cmd_Impdata.Enabled = True: Exit Sub
'        Exit Sub
'    End If
'
''    資料檢驗 --判斷SKU是否存在
'    If Trim(rsMainT16_3("品號")) <> "60400119" Then
'        str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT16_3("品號")) & "' and Storerkey = 'LMYS01' "
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If tmp_Rs.EOF Then
'            MsgBox "訂單發現新品號 (" & Trim(rsMainT16_3("品號")) & " ) " & Trim(rsMainT16_3("品名")) & "，訂單轉入終止!!": Cmd_Impdata.Enabled = True: Screen.MousePointer = 0
'            SSTab2.Enabled = True: Cmd_Impdata.Enabled = True
'            Exit Sub
'        End If
'    End If
'
'    rsMainT16_3.MoveNext
'Loop
'rsMainT16_3.MoveFirst
'End If
'
'Tran_Level = cn.BeginTrans: Cmd_Impdata.Enabled = False
'
'Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long
'Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
'Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
'Dim strDate As String, strOrderType As String, strSku As String
'
''開始匯入 銷貨明細表
'If rsMainT16_1 Is Nothing Then GoTo next18
'If rsMainT16_1.RecordCount = 0 Then GoTo next18
'
''取最後客戶編號
'Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LMYS01' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
'tmp_Rs.Open str_SQL, cn
'
'If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))
'
'tmp_Rs.Close
'
'Do While Not rsMainT16_1.EOF
'    DoEvents: DoEvents
'
''    資料檢驗--判斷是否屬佰事達訂單-->跳下一筆
'    If Trim(rsMainT16_1("品號")) = "60400119" Then intNotBest = intNotBest + 1: GoTo nextRow17
'
'    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
'    If strOrderNo <> UCase(Trim(rsMainT16_1("銷貨單號"))) Then
'        strOrderNo = UCase(Trim(rsMainT16_1("銷貨單號")))
'        int_orderlinenuber = 0
'        blDuplicationOrder = False
'
'        '檢查是否有此客戶編號
'        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LMYS01' and consigneekey = '" & myExCharFilter(Trim(rsMainT16_1("客戶代號"))) & "' order by len(consigneekey),consigneekey "
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.CursorLocation = 3
'        tmp_Rs.Open str_SQL, cn
'
'        If tmp_Rs.EOF Then
'            '無此客戶編號則新增
'            intTmp = intTmp + 1
'            strConsigneeKey = "BEST" & Format(intTmp, "000000")
'            'strConsigneeKey = myExCharFilter(Trim(rsMainT16_1("客戶代號")))
'
'            '新增客戶主檔
'            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
'            " values('LMYS01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_1("客戶簡稱"))) & "','" & myExCharFilter(Trim(rsMainT16_1("客戶簡稱"))) & "','','','" & myExCharFilter(Trim(rsMainT16_1("送貨地址"))) & "','','" & User_id & "','" & User_id & "' ) ", RowsAffect, adExecuteNoRecords
'            '紀錄新增之客戶編號
'            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'        Else
'            '比對聯絡人、電話與到貨地址是否相符
'            Call Confirm_Recordset_Closed(rsTmp)
'            str_SQL = "select * from trp01m " & _
'                        "where storerkey = 'LMYS01' and full_name = '" & myExCharFilter(Trim(rsMainT16_1("客戶簡稱"))) & "' " & _
'                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT16_1("送貨地址"))) & "' "
'            rsTmp.CursorLocation = 3
'            rsTmp.Open str_SQL, cn
'
'                If rsTmp.EOF Then
'                    '聯絡人、電話與到貨地址不符
'                    intTmp = intTmp + 1
'                    strConsigneeKey = "BEST" & Format(intTmp, "000000") '待確認BEST
'
'                    '新增客戶主檔
'                    cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes,updatesource,addwho,editwho) " & _
'                    " values('LMYS01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_1("客戶簡稱"))) & "','" & myExCharFilter(Trim(rsMainT16_1("客戶簡稱"))) & "','','','" & myExCharFilter(Trim(rsMainT16_1("送貨地址"))) & "','','','" & User_id & "','" & User_id & "' ) ", RowsAffect, adExecuteNoRecords
'
'                    '紀錄新增之客戶編號
'                    strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'                Else '相符沿用舊客編
'                    strConsigneeKey = Trim(rsTmp("consigneekey"))
'                    blCustomerMatch = True
'
'                End If
'            rsTmp.Close
'        End If
'        tmp_Rs.Close
'
'        '資料檢驗--判斷訂單是否重複，重複不增加
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT16_1("銷貨單號"))) & "' and storerkey = 'LMYS01' and isnull(type,'') <> '刪單' "
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'        If tmp_Rs.EOF Then
'
'            '取訂單號碼
'            str_SQL = "select isnull(max(orderkey),0) from orders"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
'            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
'            tmp_Rs.Close
'
'            '配送倉別判斷
''           strFacility = Trim(rsMainT16_1("倉庫"))          倉庫待確認
'            strFacility = "佰事達北倉"
'
'            arrTmp = Split(Trim(rsMainT16_1("指定日期")), "/")
'            If Len(Trim(rsMainT16_1("指定日期"))) = 0 Then    '如果沒有指定日期,則帶隔日一天
'                strDate = Format(Now + 1, "YYYY/MM/DD")
'            Else
'                strDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
'            End If
'
'            arrTmp = Split(Trim(rsMainT16_1("銷貨日期")), "/")
'            strOrderDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
'            Dim intPointer As Integer
'            intPointer = 1
'            Int_otqty = Int_otqty + Val(Trim(rsMainT16_1("件數")))
'            'updatesource
'            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,otqty,externconsigneekey) " & _
'            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT16_1("銷貨單號"))) & "','C','LMYS01','" & strOrderDate & "','" & strDate & "','" & strFacility & "','" & _
'            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_1("客戶簡稱"))) & "','','','','','','" & myExCharFilter(Trim(GetWord(rsMainT16_1("送貨地址"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT16_1("送貨地址"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT16_1("備註"))) & "','" & filLocalFileT16.FileName & "','','" & User_id & "','" & User_id & "','','" & myExCharFilter(Trim(rsMainT16_1("件數"))) & "','" & myExCharFilter(Trim(rsMainT16_1("客戶代號"))) & "') "
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'            int_Order = int_Order + 1
'            Int_C = Int_C + 1
'
'            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LMYS01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
'        Else
'            '訂單重複
'            Call FTPlog("訂單重複" & str_SQL)
'            '紀錄重複
'            strReOrderkey = strReOrderkey & Trim(rsMainT16_1("銷貨單號")) & "','"
'            blDuplicationOrder = True
'
'        End If
'    End If
'
'        '訂單重複檢查
'        If blDuplicationOrder = False Then
'
'            '增加明細
'            int_orderlinenuber = int_orderlinenuber + 1
'
'            lngCasecnt = 1
'
'            '單位換算
'            If Left(myExCharFilter(Trim(rsMainT16_1("單位"))), 1) = "箱" Then
'
'                '取箱包轉換率
'                str_SQL = "select casecnt = case when p.casecnt = 0 then 1 else p.casecnt end from " & strWMSDB & "..sku s (nolock) join " & strWMSDB & "..pack p (nolock) on s.packkey = p.packkey " & _
'                "where s.storerkey = 'LMYS01' and s.sku = '" & myExCharFilter(Trim(rsMainT16_1("品號"))) & "' "
'
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'                lngCasecnt = tmp_Rs("casecnt")
'                tmp_Rs.Close
'
'            End If
'
'            intQTY = (Val(rsMainT16_1("銷貨數量")) + Val(rsMainT16_1("贈/備品量"))) * lngCasecnt
'            strLot06 = "R01" '預設R01 , 中倉R01-C
'
'            '訂單明細資料新增
'            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
'            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT16_1("銷貨單號"))) & "','" & myExCharFilter(Trim(rsMainT16_1("品號"))) & "','LMYS01'," & _
'            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT16_1("單位"))) & "','0','" & myExCharFilter(Trim(rsMainT16_1("備註"))) & "')"
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'            '更新packkey
'            str_SQL = "update orderdetail " & _
'            "Set orderdetail.packkey = sku.packkey " & _
'            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
'            "where orderkey = '" & str_Orderkey & "' "
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'            int_OrderLine = int_OrderLine + 1
'        End If
'
'nextRow17:
'        rsMainT16_1.MoveNext
'Loop
'
'next18:
'
'If rsMainT16_2 Is Nothing Then GoTo next19
'If rsMainT16_2.RecordCount = 0 Then GoTo next19
'
''取最後客戶編號
'Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LMYS01' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
'tmp_Rs.Open str_SQL, cn
'
'If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))
'
'tmp_Rs.Close
'
''開始匯入 轉撥明細表
'strOrderNo = ""
'Do While Not rsMainT16_2.EOF
'
'    '資料檢驗--判斷是否屬佰事達訂單-->跳下一筆
'    If Trim(rsMainT16_2("品號")) = "60400119" Then intNotBest = intNotBest + 1: GoTo nextRow18
'
'    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
'    If strOrderNo <> UCase(Trim(rsMainT16_2("單別-單號"))) Then
'        strOrderNo = UCase(Trim(rsMainT16_2("單別-單號")))
'        int_orderlinenuber = 0
'        strLot06 = ""
'        blDuplicationOrder = False
'
'        '檢查是否有此客戶名稱
'        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LMYS01' and rtrim(full_name) = '" & myExCharFilter(Trim(rsMainT16_2("轉入庫別"))) & "' order by len(consigneekey),consigneekey "
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.CursorLocation = 3
'        tmp_Rs.Open str_SQL, cn
'
'        If tmp_Rs.EOF Then
'            '無此客戶名稱則新增
'            intTmp = intTmp + 1
'            strConsigneeKey = "BEST" & Format(intTmp, "000000") '待確認
'
'            '新增客戶主檔 待確認,updatesource要帶甚麼
'            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
'            " values('LMYS01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_2("轉入庫別"))) & "','" & myExCharFilter(Trim(rsMainT16_2("轉入庫別"))) & "','','','" & myExCharFilter(Trim(rsMainT16_2("送貨地址"))) & "','','" & User_id & "','" & User_id & "') ", RowsAffect, adExecuteNoRecords
'
'            '紀錄新增之客戶編號
'            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'        Else
'            '比對聯絡人、電話與到貨地址是否相符
'            Call Confirm_Recordset_Closed(rsTmp)
'            str_SQL = "select * from trp01m " & _
'                        "where storerkey = 'LMYS01' and full_name = '" & myExCharFilter(Trim(rsMainT16_2("轉入庫別"))) & "' " & _
'                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT16_2("送貨地址"))) & "' "
'            rsTmp.CursorLocation = 3
'            rsTmp.Open str_SQL, cn
'
'            If rsTmp.EOF Then
'                '聯絡人、電話與到貨地址不符
'                intTmp = intTmp + 1
'                strConsigneeKey = "BEST" & Format(intTmp, "000000") '待確認
'
'                '新增客戶主檔
'                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes,updatesource,addwho,editwho) " & _
'                " values('LMYS01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_2("轉入庫別"))) & "','" & myExCharFilter(Trim(rsMainT16_2("轉入庫別"))) & "','','','" & myExCharFilter(Trim(rsMainT16_2("送貨地址"))) & "','" & myExCharFilter(Trim(rsMainT16_2("備註"))) & "','','" & User_id & "','" & User_id & "') ", RowsAffect, adExecuteNoRecords
'
'                '紀錄新增之客戶編號
'                strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'            Else '相符沿用舊客編
'                strConsigneeKey = Trim(rsTmp("consigneekey"))
'                blCustomerMatch = True
'
'            End If
'            rsTmp.Close
'        End If
'        tmp_Rs.Close
'
'        '資料檢驗--判斷訂單是否重複，重複不增加
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT16_2("單別-單號"))) & "' and storerkey = 'LMYS01' and isnull(type,'') <> '刪單' "
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'        If tmp_Rs.EOF Then
'
'            '取訂單號碼
'            str_SQL = "select isnull(max(orderkey),0) from orders"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
'            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
'            tmp_Rs.Close
'
'            '配送倉別判斷
'            'strFacility = Trim(rsMainT16_2("轉入庫別"))
'            strFacility = "佰事達北倉"
'
'            arrTmp = Split(Trim(rsMainT16_2("指定日期")), "/")
'            If Len(Trim(rsMainT16_2("指定日期"))) = 0 Then    '如果沒有指定日期,則帶隔日一天
'                strDate = Format(Now + 1, "YYYY/MM/DD")
'            Else
'                strDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
'            End If
'
'            arrTmp = Split(Trim(rsMainT16_2("單據日期")), "/")
'            strOrderDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
'            intPointer = 1
'
'            strOrderType = "C"
'            If Trim(rsMainT16_2("轉入庫別")) = "物流倉/佰事達北倉" Or Trim(rsMainT16_2("轉入庫別")) = "物流倉/佰事達中倉" Then strOrderType = "RC"
'            If strOrderType = "C" Then
'                Int_C = Int_C + 1
'            Else
'                Int_RC = Int_RC + 1
'            End If
'            strLot06 = IIf(UCase(Trim(rsMainT16_2("轉入庫別"))) = "物流倉/佰事達中倉", "R01-C", "R01")
'            strFacility = IIf(UCase(Trim(rsMainT16_2("轉入庫別"))) = "物流倉/佰事達中倉", "佰事達中倉", "佰事達北倉")
'
'            If UCase(Trim(rsMainT16_2("轉入庫別"))) = "物流倉/佰事達北倉" Or UCase(Trim(rsMainT16_2("轉入庫別"))) = "物流倉/佰事達中倉" Then
'                'RC類訂單不累加件數
'            Else
'                Int_otqty = Int_otqty + Val(Trim(rsMainT16_2("件數")))
'            End If
'
'            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,otqty) " & _
'            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT16_2("單別-單號"))) & "','" & strOrderType & "','LMYS01','" & strOrderDate & "','" & strDate & "','" & strFacility & "','" & _
'            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_2("轉入庫別"))) & "','','','','','','" & myExCharFilter(Trim(GetWord(rsMainT16_2("送貨地址"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT16_2("送貨地址"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT16_2("備註"))) & "','" & filLocalFileT16.FileName & "','','" & User_id & "','" & User_id & "','','" & myExCharFilter(Trim(rsMainT16_2("件數"))) & "')"
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'            int_Order = int_Order + 1
'
'            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LMYS01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
'        Else
'            '訂單重複
'            Call FTPlog("訂單重複" & str_SQL)
'            '紀錄重複
'            strReOrderkey = strReOrderkey & Trim(rsMainT16_2("單別-單號")) & "','"
'            blDuplicationOrder = True
'
'        End If
'    End If
'
'        '訂單重複檢查
'        If blDuplicationOrder = False Then
'            '增加明細
'            int_orderlinenuber = int_orderlinenuber + 1
'
''            lngCasecnt = 1
''            '單位換算
''            If Left(myExCharFilter(Trim(rsMainT16_2("單位"))), 1) = "箱" Then
''
''                '取箱包轉換率
''                str_SQL = "select casecnt = case when p.casecnt = 0 then 1 else p.casecnt end from " & strWMSDB & "..sku s (nolock) join " & strWMSDB & "..pack p (nolock) on s.packkey = p.packkey " & _
''                "where s.storerkey = 'LMYS01' and s.sku = '" & myExCharFilter(Trim(rsMainT16_2("品號"))) & "' "
''
''                Call Confirm_Recordset_Closed(tmp_Rs)
''                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
''                lngCasecnt = tmp_Rs("casecnt")
''                tmp_Rs.Close
''            End If
'
'            intQTY = Val(rsMainT16_2("轉撥數量")) '* lngCasecnt
'
'            '訂單明細資料新增
'            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
'            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT16_2("單別-單號"))) & "','" & myExCharFilter(Trim(rsMainT16_2("品號"))) & "','LMYS01'," & _
'            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT16_2("單位"))) & "','0','" & myExCharFilter(Trim(rsMainT16_2("備註"))) & "')"
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'            '更新packkey
'            str_SQL = "update orderdetail " & _
'            "Set orderdetail.packkey = sku.packkey " & _
'            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
'            "where orderkey = '" & str_Orderkey & "' "
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'            int_OrderLine = int_OrderLine + 1
'        End If
'
'nextRow18:
'        rsMainT16_2.MoveNext
'Loop
'
'next19:
'
''開始匯入寄庫銷貨明細表
'If rsMainT16_3 Is Nothing Then GoTo nextend
'If rsMainT16_3.RecordCount = 0 Then GoTo nextend
'
''取最後客戶編號
'Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LMYS01' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
'tmp_Rs.Open str_SQL, cn
'
'If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))
'
'tmp_Rs.Close
'
'Do While Not rsMainT16_3.EOF
'
'    '資料檢驗--判斷是否屬佰事達訂單-->跳下一筆
'    If Trim(rsMainT16_3("品號")) = "60400119" Then intNotBest = intNotBest + 1: GoTo nextRow19
'
'    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
'    If strOrderNo <> UCase(Trim(rsMainT16_3("銷貨單號"))) Then
'        strOrderNo = UCase(Trim(rsMainT16_3("銷貨單號")))
'        int_orderlinenuber = 0
'        blDuplicationOrder = False
'
'        '檢查是否有此客戶名稱
'        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LMYS01' and rtrim(consigneekey) = '" & myExCharFilter(Trim(rsMainT16_3("客戶代號"))) & "' order by len(consigneekey),consigneekey "
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.CursorLocation = 3
'        tmp_Rs.Open str_SQL, cn
'
'        If tmp_Rs.EOF Then
'            '無此客戶名稱則新增
'            intTmp = intTmp + 1
'            'strConsigneeKey = "BEST" & Format(intTmp, "000000")
'            strConsigneeKey = myExCharFilter(Trim(rsMainT16_3("客戶代號")))
'
'            '新增客戶主檔
'            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
'            " values('LMYS01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_3("客戶簡稱"))) & "','" & myExCharFilter(Trim(rsMainT16_3("客戶簡稱"))) & "','','','" & myExCharFilter(Trim(rsMainT16_3("送貨地址"))) & "','','" & User_id & "','" & User_id & "' ) ", RowsAffect, adExecuteNoRecords
'            '紀錄新增之客戶編號
'            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'        Else
'            '比對聯絡人、電話與到貨地址是否相符
'            Call Confirm_Recordset_Closed(rsTmp)
'            str_SQL = "select * from trp01m " & _
'                        "where storerkey = 'LMYS01' and full_name = '" & myExCharFilter(Trim(rsMainT16_3("客戶簡稱"))) & "' " & _
'                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT16_3("送貨地址"))) & "' "
'            rsTmp.CursorLocation = 3
'            rsTmp.Open str_SQL, cn
'
'                If rsTmp.EOF Then
'                    '聯絡人、與到貨地址不符
'                    intTmp = intTmp + 1
'                    strConsigneeKey = "BEST" & Format(intTmp, "000000") '待確認BEST
'
'                    '新增客戶主檔
'                    cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes,updatesource,addwho,editwho) " & _
'                    " values('LMYS01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_3("客戶簡稱"))) & "','" & myExCharFilter(Trim(rsMainT16_3("客戶簡稱"))) & "','','','" & myExCharFilter(Trim(rsMainT16_3("送貨地址"))) & "','','" & myExCharFilter(Trim(rsMainT16_3("客戶代號"))) & "','" & User_id & "','" & User_id & "' ) ", RowsAffect, adExecuteNoRecords
'
'                    '紀錄新增之客戶編號
'                    strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'                Else '相符沿用舊客編
'                    strConsigneeKey = Trim(rsTmp("consigneekey"))
'                    blCustomerMatch = True
'
'                End If
'            rsTmp.Close
'        End If
'        tmp_Rs.Close
'
'        '資料檢驗--判斷訂單是否重複，重複不增加
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT16_3("銷貨單號"))) & "' and storerkey = 'LMYS01' and isnull(type,'') <> '刪單' "
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'        If tmp_Rs.EOF Then
'
'            '取訂單號碼
'            str_SQL = "select isnull(max(orderkey),0) from orders"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
'            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
'            tmp_Rs.Close
'
'            '配送倉別判斷 20130404修改
'            'strFacility = IIf(UCase(Trim(rsMainT16_3("客戶代號"))) = "1201011004", "佰事達中倉", "佰事達北倉")
'            strFacility = "佰事達北倉"
'            arrTmp = Split(Trim(rsMainT16_3("指定日期")), "/")
'            If Len(Trim(rsMainT16_3("指定日期"))) = 0 Then    '如果沒有指定日期,則帶隔日一天
'                strDate = Format(Now + 1, "YYYY/MM/DD")
'            Else
'                strDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
'            End If
'
'            arrTmp = Split(Trim(rsMainT16_3("銷貨日期")), "/")
'            strOrderDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
'            intPointer = 1
'            'Int_otqty = Int_otqty + Val(Trim(rsMainT16_3("件數"))) I類訂單不加總件數
'
'            'updatesource 要帶filLocalFileT11.FileName 還是 客戶代號 待確認
'            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,otqty,externconsigneekey) " & _
'            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT16_3("銷貨單號"))) & "','I','LMYS01','" & strOrderDate & "','" & strDate & "','" & strFacility & "','" & _
'            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT16_3("客戶簡稱"))) & "','','','','','','" & myExCharFilter(Trim(GetWord(rsMainT16_3("送貨地址"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT16_3("送貨地址"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT16_3("備註"))) & "','" & filLocalFileT16.FileName & "','','" & User_id & "','" & User_id & "','','" & Val(myExCharFilter(Trim(rsMainT16_3("件數")))) & "','" & myExCharFilter(Trim(rsMainT16_3("客戶代號"))) & "') "
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'            int_Order = int_Order + 1
'            Int_I = Int_I + 1
'
'            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LMYS01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
'        Else
'            '訂單重複
'            Call FTPlog("訂單重複" & str_SQL)
'            '紀錄重複
'            strReOrderkey = strReOrderkey & Trim(rsMainT16_3("銷貨單號")) & "','"
'            blDuplicationOrder = True
'
'        End If
'    End If
'
'        '訂單重複檢查
'        If blDuplicationOrder = False Then
'            '增加明細
'            int_orderlinenuber = int_orderlinenuber + 1
'
''            lngCasecnt = 1
''
''            '單位換算
''            If Left(myExCharFilter(Trim(rsMainT16_3("單位"))), 1) = "箱" Then
''
''                '取箱包轉換率
''                str_SQL = "select casecnt = case when p.casecnt = 0 then 1 else p.casecnt end from " & strWMSDB & "..sku s (nolock) join " & strWMSDB & "..pack p (nolock) on s.packkey = p.packkey " & _
''                "where s.storerkey = 'LMYS01' and s.sku = '" & myExCharFilter(Trim(rsMainT16_3("品號"))) & "' "
''
''                Call Confirm_Recordset_Closed(tmp_Rs)
''                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
''                lngCasecnt = tmp_Rs("casecnt")
''                tmp_Rs.Close
''            End If
'
'            intQTY = (Val(rsMainT16_3("銷貨數量")) + Val(rsMainT16_3("贈/備品量"))) '* lngCasecnt
'
'            strLot06 = "R01" '預設R01 , 中倉R01-C ,20130408修改 所有為R01,佰事達北倉
'
''            If Trim(rsMainT16_3("客戶代號")) = "1201011004" Then strLot06 = "R01-C"
''            strLot06 = IIf(UCase(Trim(rsMainT16_3("客戶代號"))) = "1201011004", "R01-C", "R01")
'
'            '訂單明細資料新增
'            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
'            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT16_3("銷貨單號"))) & "','" & myExCharFilter(Trim(rsMainT16_3("品號"))) & "','LMYS01'," & _
'            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT16_3("單位"))) & "','0','" & myExCharFilter(Trim(rsMainT16_3("備註"))) & "')"
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'            '更新packkey
'            str_SQL = "update orderdetail " & _
'            "Set orderdetail.packkey = sku.packkey " & _
'            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
'            "where orderkey = '" & str_Orderkey & "' "
'
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'            int_OrderLine = int_OrderLine + 1
'        End If
'
'nextRow19:
'        rsMainT16_3.MoveNext
'Loop
'
'nextend:
'
''加總C類訂單 新增一筆A2B訂單
'        Dim ExternOrderKey As String
'        strDate = Format(Now, "YYYYMMDD") '目前時間
'        int_orderlinenuber = 0
'        int_orderlinenuber = int_orderlinenuber + 1
'        ExternOrderKey = "A2B" & strDate
'        '檢查externorderkey是否已經有相同的 , 同一天匯入兩次以上訂單資料
'
'        str_SQL = "select top 1 externorderkey from orders where storerkey = 'LMYS01' and externorderkey like '" & ExternOrderKey & "%' order by externorderkey desc"
'        Call Confirm_Recordset_Closed(rsTmp)
'        rsTmp.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If rsTmp.EOF Then
'            GoTo A2B
'        End If
'
'        If ExternOrderKey = Left(Trim(rsTmp("externorderkey").Value), 11) And Mid(Trim(rsTmp("externorderkey").Value), 12, 1) = "-" Then
'                ExternOrderKey = ExternOrderKey & "-" & mySplit(Trim(rsTmp("externorderkey").Value), "-", 1) + 1
'        Else
'                ExternOrderKey = ExternOrderKey & "-1"
'        End If
'
'A2B:
'        '取訂單號碼
'        str_SQL = "select isnull(max(orderkey),0) from orders"
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'        str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
'        If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
'        tmp_Rs.Close: rsTmp.Close
'        '取客戶主檔資料
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        str_SQL = "select * from trp01m where storerkey = 'LMYS01' and consigneekey = 'BEST000016' " 'BEST000016=馬玉山高雄 ;BEST000017 = 物流倉\佰事達北倉
'        tmp_Rs.Open str_SQL, cn
'        '寫入orders
'        str_SQL = "INSERT orders (OrderKey,StorerKey,ExternOrderKey,OrderDate,Deliverydate,Priority,ConsigneeKey,c_contact1,c_company,c_address1,c_zip,c_phone1,b_company,UpdateSource,type,door,route,stop,Notes,adddate,addwho,editdate,editwho,doroute,CustomerOrderkey,externconsigneekey,otqty) " & _
'                  "values('" & str_Orderkey & "','LMYS01','" & ExternOrderKey & "','" & strDate & "','" & strDate & "','A2B','" & Trim(tmp_Rs("ConsigneeKey").Value) & "','" & Trim(tmp_Rs("Contact").Value) & "','" & Trim(tmp_Rs("full_name").Value) & "','" & Trim(tmp_Rs("address").Value) & "'," & _
'                         "'" & Trim(tmp_Rs("zip").Value) & "','" & Trim(tmp_Rs("phone").Value) & "','BEST000017','" & filLocalFileT16.FileName & "','','99','99','99','高雄馬玉山提貨至觀音倉 共計" & Int_otqty & "件',getdate(),'" & User_id & "',getdate(),'" & User_id & "','Y','" & ExternOrderKey & "','" & Trim(tmp_Rs("consigneekey").Value) & "','" & Int_otqty & "')"
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'        '寫入orderdetail
'        str_SQL = "insert into orderdetail (orderkey,orderlinenumber,externorderkey,sku,storerkey,originalqty,openqty,uom,packkey,status,adddate,addwho,editwho,lottable06) " & _
'                  "values ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "0000") & "','" & ExternOrderKey & "','OT','LMYS01','" & Int_otqty & "','" & Int_otqty & "','EA','OT','0',getdate(),'" & User_id & "','" & User_id & "','R01') "
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'        tmp_Rs.Close
''-----------------------
'
'cn.CommitTrans: Tran_Level = 0
'
''訊息顯示
'    msg_text = "匯入 " & int_Order & " 筆訂單" & Chr(13) & "轉運出貨(C): " & Int_C & " 筆 : 提貨入庫(RC): " & Int_RC & " 筆 : 一般出貨(I) : " & Int_I & " 筆" & Chr(13) & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & Chr(13) & "匯入 非佰事達訂單 " & intNotBest & " 筆明細" & Chr(13) & Chr(13) & "TMS單號： " & strOrderKeyS & "~" & str_Orderkey & Chr(13) & Chr(13) & "系統產生1筆A2B訂單及明細，訂單號碼:" & str_Orderkey
'    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'    Call FTPlog("匯入 " & int_Order & " 筆訂單" & Chr(13) & "轉運出貨(C): " & Int_C & " 筆 : 提貨入庫(RC): " & Int_RC & " 筆 : 一般出貨(I) : " & Int_I & " 筆" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "非佰事達訂單 " & intNotBest & " 筆明細，檔案 " & filLocalFileT16.FileName)
'
''訂單重複顯示
'If Len(strReOrderkey & strRePoOrderkey) > 0 Then
'
'    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT16.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 品號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
'        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = 'LMYS01'"
'
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
'
'    Call Recordset2Excel("訂單重複", tmp_Rs)
'    If Dir("C:\LMYS01\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\LMYS01\訂單重複"
'    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
'    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LMYS01\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
'    Set MyXlsApp = Nothing
'
'End If
'
''備份至FTP
'If Dir("O:\LMYS01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LMYS01\OrdersBackup"
'FileCopy strTranFileName, "O:\LMYS01\OrdersBackup\" & filLocalFileT16.FileName
'
''備份檔案
'If Dir("C:\BEST\LMYS01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LMYS01\Orders\Backup"
'If Dir("C:\BEST\LMYS01\Orders\Backup\" & filLocalFileT16.FileName) = "" Then
'    FileCopy strTranFileName, "C:\BEST\LMYS01\Orders\Backup\" & filLocalFileT16.FileName
'Else
'    FileCopy strTranFileName, "C:\BEST\LMYS01\Orders\Backup\" & mySplit(filLocalFileT16.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT16.FileName, ".", -1)
'End If
'
'Kill strTranFileName
'
'filLocalFileT16.Refresh:
'Screen.MousePointer = 0: Cmd_Impdata.Enabled = True: SSTab2.Enabled = True
'Exit Sub

err_Handle:
    Set dgMainT16.DataSource = Nothing: Set dgMainT16_1.DataSource = Nothing
    Cmd_Impdata.Enabled = True: Screen.MousePointer = 0: SSTab2.Enabled = True
    Call ErrorMsgbox(App.title, err.Number, err.Description, "主檔檔案名稱： " & Str_updatesource1 & "明細檔案名稱：" & Str_updatesource2)
End Sub

Private Sub cmdImportT15_Click()
Dim bl_Error As Boolean '記錄有小數點的旗標
Dim str_Error As String '記錄有小數點錯誤的資料
Dim x As Long
bl_Error = False: str_Error = ""
Dim str_Storerkey As String

If rsMainT15 Is Nothing Then Exit Sub
If rsMainT15.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT15.Enabled = False: cmdImportT15.Enabled = False
strTranFileName = filLocalFileT15.Path & "\" & filLocalFileT15.FileName

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT15.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT15.Enabled = True: dgMainT15.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT15.RecordCount = 0 Or rsMainT15 Is Nothing Then
Else
rsMainT15.MoveFirst
str_Storerkey = "LCHF01"
Do While Not rsMainT15.EOF
    '到貨日期檢查
    If Len(Trim(rsMainT15("訂單預交日"))) = 0 Then
        If Trim(rsMainT15("單別")) = "出退" Or Trim(rsMainT15("單別")) = "代退" Then
        Else
            MsgBox "訂單號碼:" & Trim(rsMainT15("出貨單號")) & "的到貨日為空白，訂單轉入終止!", 16, Me.Caption: dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
        End If
    ElseIf Len(Trim(rsMainT15("訂單預交日"))) > 0 And Len(Trim(rsMainT15("訂單預交日"))) < 8 Then
         MsgBox "訂單號碼:" & Trim(rsMainT15("出貨單號")) & "的到貨日:" & Trim(rsMainT15("訂單預交日")) & "，格式不對，請補齊8碼，訂單轉入終止!", 16, Me.Caption: dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
    Else
        '檢查到貨日不可小於今日
        If Trim(rsMainT15("訂單預交日")) < Format(Now, "YYYYMMDD") Then
            If blAdmin = True Then
            
            '最高權限不檢查到貨日
                 x = MsgBox("到貨日小於今日，你確定要繼續嗎?", vbQuestion + vbYesNo, "最高權限到貨日檢查") '紀錄按下的是確定或是取消
                    If x = 6 Then
                        '繼續
                    Else
                        '離開
                         dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
                    End If
            Else
            
            End If
        End If
    End If
    
'    '貨主檢查
'    If Len(Trim(rsMainT15("貨主"))) = 0 Or Trim(rsMainT15("貨主")) <> "LCHF01" Then
'        MsgBox "訂單號碼:" & Trim(rsMainT15("出貨單號")) & "的貨主有誤，訂單轉入終止!", 16, Me.Caption: dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
'    End If
    
    
    '訂單日檢查
    If Len(Trim(rsMainT15("日期"))) = 0 Then
        MsgBox "訂單號碼:" & Trim(rsMainT15("出貨單號")) & "的訂單日為空白，訂單轉入終止!", 16, Me.Caption: dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT15("日期"))) > 0 And Len(Trim(rsMainT15("日期"))) < 8 Then
         MsgBox "訂單號碼:" & Trim(rsMainT15("出貨單號")) & "的訂單日:" & Trim(rsMainT15("日期")) & "，格式不對，請補齊8碼，訂單轉入終止!", 16, Me.Caption: dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
    Else
        If Trim(rsMainT15("單別")) = "出退" Or Trim(rsMainT15("單別")) = "代退" Then
        Else
            If Trim(rsMainT15("日期")) > Trim(rsMainT15("訂單預交日")) Then MsgBox "訂單號碼:" & Trim(rsMainT15("出貨單號")) & "的訂單日:" & Trim(rsMainT15("日期")) & "，大於到貨日，訂單轉入終止!", 16, Me.Caption: dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
        End If
    End If
    
    '數量檢查
    If Val(rsMainT15("數量")) < 1 Then
        If Trim(rsMainT15("單別")) = "出退" Or Trim(rsMainT15("單別")) = "代退" Then
        Else
            MsgBox "數量小於1，" & Trim(rsMainT15("出貨單號")) & "-品號：" & Trim(rsMainT15("產品代號")) & "，訂單轉入終止!!請確認!!", , "訂單檔匯入": dgMainT15.Enabled = True: cmdImportT15.Enabled = True: Exit Sub
            Exit Sub
        End If
    End If
    
        '資料檢驗 --判斷SKU是否存在
        str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMainT15("產品代號")) & "' and Storerkey = '" & str_Storerkey & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '按鈕那些要改
            MsgBox "訂單發現新品號 (" & Trim(rsMainT15("產品代號")) & ")，請先於商品主檔新建商品資料，訂單轉入終止!!": cmdImportT15.Enabled = True: Screen.MousePointer = 0
            dgMainT15.Enabled = True: cmdImportT15.Enabled = True
            Exit Sub
        End If
        
        '檢查A2B訂單以外的客戶編號是否存在
        str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT15("收貨客戶")) & "' and Storerkey = '" & str_Storerkey & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '按鈕那些要改
            MsgBox "訂單發現新客戶編號 ( " & Trim(rsMainT15("收貨客戶")) & " )，請先於客戶主檔新建客戶資料，訂單轉入終止!!": cmdImportT15.Enabled = True: Screen.MousePointer = 0
            dgMainT15.Enabled = True: cmdImportT15.Enabled = True
            Exit Sub
        End If
        
        '檢查數量有無小數點
        If InStr(Trim(rsMainT15("數量")), ".") <> 0 Then
            str_Error = "訂單號碼:" & Trim(rsMainT15("出貨單號")) & "，品號:" & Trim(rsMainT15("產品代號")) & Chr(13) & str_Error
            bl_Error = True
        End If
        
'        '檢查貨主--mark by Gemini @20160602
'        If UCase(Trim(rsMainT15("貨主"))) <> "LABT01" And UCase(Trim(rsMainT15("貨主"))) <> "LLFA01" Then
'            MsgBox "訂單發現非亞培的貨主: " & Trim(rsMainT15("貨主")) & " )，此匯入程式僅供匯入亞培及利豐訂單，請確認後再匯入，訂單轉入終止!!": cmdImportT15.Enabled = True: Screen.MousePointer = 0
'            dgMainT15.Enabled = True: cmdImportT15.Enabled = True
'            Exit Sub
'        End If

        '判斷單別
        If Trim(rsMainT15("單別")) = "A2B" Then
            MsgBox "單別為A2B:" & Trim(rsMainT15("單別")) & "，A2B訂單請由公版EXCEL訂單匯入，訂單轉入終止!!": cmdImportT15.Enabled = True: Screen.MousePointer = 0
                dgMainT15.Enabled = True: cmdImportT15.Enabled = True
            Exit Sub
        End If
        
        If Trim(rsMainT15("單別")) = "出貨" Or Trim(rsMainT15("單別")) = "出退" Or Trim(rsMainT15("單別")) = "代退" Or Trim(rsMainT15("單別")) = "代銷" Then
        Else
            MsgBox "系統無此單別:" & Trim(rsMainT15("單別")) & "，請確認此訂單類別是否正確，訂單轉入終止!!": cmdImportT15.Enabled = True: Screen.MousePointer = 0
                dgMainT15.Enabled = True: cmdImportT15.Enabled = True
            Exit Sub
        End If
        
    rsMainT15.MoveNext
Loop
rsMainT15.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "資料有小數點！請重新匯入"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT15.Enabled = True: cmdImportT15.Enabled = True
                Exit Sub
End If


Tran_Level = cn.BeginTrans: cmdImportT15.Enabled = False: dgMainT15.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String
Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String


'開始匯入
Do While Not rsMainT15.EOF
    DoEvents: DoEvents
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT15("出貨單號"))) Then
        strOrderNo = UCase(Trim(rsMainT15("出貨單號")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '已經檢查過，則不比對，直接抓出客戶主檔中的 客戶編號，郵遞區號，連絡人，電話，地址
        '單別為A2B則，抓提貨客編，非A2B則抓到貨客編
        str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & str_Storerkey & "' and consigneekey = '" & Trim(rsMainT15("收貨客戶")) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        
        Dim str_Priority As String
        '相符沿用舊客編
        strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
        strZip = myExCharFilter(Trim(tmp_Rs("zip")))
        strContact = myExCharFilter(Trim(tmp_Rs("contact")))
        strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
        strAddress = myExCharFilter(Trim(tmp_Rs("address")))
        strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
        
        '將單別代成I or R
        If myExCharFilter(Trim(rsMainT15("單別"))) = "出退" Or myExCharFilter(Trim(rsMainT15("單別"))) = "代退" Then
            str_Priority = "R"
        Else
            str_Priority = "I"
        End If
        
        blCustomerMatch = True
        tmp_Rs.Close

        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select externorderkey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT15("出貨單號"))) & "' and storerkey = '" & str_Storerkey & "' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '取訂單號碼
            tmp_Rs.Close
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
'            If UCase(Right(Trim(rsMainT15("倉別")), 2)) = "-C" Then
'                strFacility = "佰事達中倉"
'            ElseIf UCase(Right(Trim(rsMainT15("倉別")), 2)) = "-S" Then
'                strFacility = "佰事達南倉"
'            Else
            strFacility = "佰事達北倉"
'            End If
            
'            If Trim(rsMainT15("倉別")) = "" Then strFacility = ""

            strOrderDate = Trim(rsMainT15("日期"))
            Dim intPointer As Integer
            intPointer = 1
            
            Call Confirm_Recordset_Closed(tmp_Rs)
            
            If (Trim(rsMainT15("單別")) = "出退" Or Trim(rsMainT15("單別")) = "代退") And Len(Trim(rsMainT15("訂單預交日"))) = 0 Then
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT15("出貨單號")) & "','" & str_Priority & "','" & str_Storerkey & "','" & strOrderDate & "','" & strOrderDate & "','" & strFacility & "','" & _
                Trim(rsMainT15("收貨客戶")) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & Trim(GetWord(strAddress, intPointer, 58)) & "','" & Trim(GetWord(strAddress, intPointer, 45)) & "','" & Trim(rsMainT15(["客戶訂單(訂單)"])) & "','" & Trim(rsMainT15("單據備註")) & "','" & filLocalFileT15.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
            Else
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT15("出貨單號")) & "','" & str_Priority & "','" & str_Storerkey & "','" & strOrderDate & "','" & Trim(rsMainT15("訂單預交日")) & "','" & strFacility & "','" & _
                Trim(rsMainT15("收貨客戶")) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & Trim(GetWord(strAddress, intPointer, 58)) & "','" & Trim(GetWord(strAddress, intPointer, 45)) & "','" & Trim(rsMainT15(["客戶訂單(訂單)"])) & "','" & Trim(rsMainT15("單據備註")) & "','" & filLocalFileT15.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
            End If
            
            
'            If (Trim(rsMainT15("單別")) = "出退" Or Trim(rsMainT15("單別")) = "代退") And Len(Trim(rsMainT15("訂單預交日"))) = 0 Then
'                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
'                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT15("出貨單號"))) & "','" & str_Priority & "','" & Str_storerkey & "','" & strOrderDate & "','" & strOrderDate & "','" & strFacility & "','" & _
'                myExCharFilter(Trim(rsMainT15("收貨客戶"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT15("單據備註"))) & "','" & filLocalFileT15.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
'            Else
'                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
'                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT15("出貨單號"))) & "','" & str_Priority & "','" & Str_storerkey & "','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT15("訂單預交日"))) & "','" & strFacility & "','" & _
'                myExCharFilter(Trim(rsMainT15("收貨客戶"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT15("單據備註"))) & "','" & filLocalFileT15.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
'            End If
            
            
            
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
'            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & Str_Storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT15("出貨單號")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
        
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            
            intQTY = Abs(Val(rsMainT15("數量")))
            
            If Trim(rsMainT15("單別")) = "出退" Or Trim(rsMainT15("單別")) = "代退" Then
                '訂單明細資料新增 退貨
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM) " & _
                "values ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Trim(rsMainT15("出貨單號")) & "','" & Trim(rsMainT15("產品代號")) & "','LCHF01'," & _
                "'" & intQTY & "','" & intQTY & "','R01','" & strFacility & "','" & Trim(rsMainT15("數量單位")) & "')"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            Else
                '訂單明細資料新增
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Lottable03)" & _
                "select  '" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Trim(rsMainT15("出貨單號")) & "','" & Trim(rsMainT15("產品代號")) & "','LCHF01'," & intQTY & " * p.casecnt ," & intQTY & " * p.casecnt " & _
                ",'R01','" & strFacility & "','" & Trim(rsMainT15("數量單位")) & "','" & Trim(rsMainT15("批號")) & "'" & _
                "from " & strWMSDB & "..sku s join " & strWMSDB & "..pack p on s.packkey = p.packkey and s.sku = '" & Trim(rsMainT15("產品代號")) & "' and s.storerkey = 'LCHF01'"
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            End If
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT15.MoveNext
Loop

'更新允收期
cn.Execute "exec gs_ordersupdate 'LCHF01'", RowsAffect, adExecuteNoRecords

'檢查有無異常訂單  倉別庫別錯誤 es_Checklot06_by_storer '貨主','訂單檔名'
'str_SQL = "exec es_Checklot06_by_storer '" & Str_Storerkey & "','" & filLocalFileT18.FileName & "'"
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'If Not tmp_Rs.EOF Then  '有回傳錯誤的訂單資料，產生excel
'    Recordset2Excel "配送倉別與明細倉別不符的訂單資料", tmp_Rs
'End If
'
'tmp_Rs.Close

cn.CommitTrans: Tran_Level = 0: dgMainT15.Enabled = True


'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "TMS單號： " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "檔案 " & filLocalFileT15.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT15.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 品號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'備份檔案
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT15.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT15.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT15.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT15.FileName, ".", -1)
End If

'備份至FTP
If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT15.FileName


Kill strTranFileName
    
filLocalFileT15.Refresh:
Screen.MousePointer = 0: cmdImportT15.Enabled = True: dgMainT15.Enabled = True
Exit Sub

err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
   Call ErrorMsgbox(App.title, err.Number, err.Description, "檔案名稱： " & strTranFileName)
    cmdImportT15.Enabled = True: Screen.MousePointer = 0: dgMainT15.Enabled = True

End Sub



Private Sub cmdImportT17_Click()
Dim bl_Error As Boolean '記錄有小數點的旗標
Dim str_Error As String '記錄有小數點錯誤的資料
Dim x As Integer '記錄按下確定或是取消，決定是否更新緊急訂單urgent_mark
Dim Str_packkey As String '紀錄packkey
bl_Error = False: str_Error = "": Str_packkey = ""

If rsMainT17 Is Nothing Then Exit Sub
If rsMainT17.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT17.Enabled = False: cmdImportT17.Enabled = False
strTranFileName = filLocalFileT17.Path & "\" & filLocalFileT17.FileName

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where rtrim(updatesource)='" & filLocalFileT17.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then dgMainT17.Enabled = True: cmdImportT17.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT17.RecordCount = 0 Or rsMainT17 Is Nothing Then
Else
rsMainT17.MoveFirst
Do While Not rsMainT17.EOF
    '到貨日期檢查
    If Len(Trim(rsMainT17("預計出貨日"))) = 0 Then
    Else
        arrTmp = Split(Trim(rsMainT17("預計出貨日")), "/")
        If Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "預計出貨日小於今日，訂單轉入終止!", 16, Me.Caption: dgMainT17.Enabled = True: cmdImportT17.Enabled = True: Exit Sub
    End If
    
    '數量檢查
    If Val(rsMainT17("商品訂購數量")) < 1 Then
        MsgBox "商品訂購數量小於1，" & Trim(rsMainT17("SAP DN NO.")) & "-品號：" & Trim(rsMainT17("商品編碼")) & "，訂單轉入終止!!請確認!!", , "訂單檔匯入": dgMainT17.Enabled = True: cmdImportT17.Enabled = True: Exit Sub
        Exit Sub
    End If
    
    '資料檢驗 --判斷SKU是否存在
        str_SQL = "select sku from " & strWMSDB & "..sku(nolock) where Storerkey = 'LAPP01' and sku='" & Trim(rsMainT17("商品編碼")) & "'"
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '按鈕那些要改
            MsgBox "訂單發現新品號 (" & Trim(rsMainT17("商品編碼")) & ")，訂單轉入終止!!": cmdImportT17.Enabled = True: Screen.MousePointer = 0
            dgMainT17.Enabled = True: cmdImportT17.Enabled = True
            Exit Sub
        End If
    '檢查數量有無小數點
        If InStr(Trim(rsMainT17("商品訂購數量")), ".") <> 0 Then
            str_Error = "訂單號碼:" & Trim(rsMainT17("SAP DN NO.")) & "，品號:" & Trim(rsMainT17("商品編碼")) & Chr(13) & str_Error
            bl_Error = True
        End If
    rsMainT17.MoveNext
Loop
rsMainT17.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "資料有小數點！請重新匯入"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT17.Enabled = True: cmdImportT17.Enabled = True
                Exit Sub
End If

Tran_Level = cn.BeginTrans: cmdImportT17.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String

'開始匯入
If rsMainT17 Is Nothing Then GoTo next18
If rsMainT17.RecordCount = 0 Then GoTo next18

'取最後客戶編號
'Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LAPP01' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
'tmp_Rs.Open str_SQL, cn
'
'If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))
'
'tmp_Rs.Close


Do While Not rsMainT17.EOF
    DoEvents: DoEvents
    
'    資料檢驗--判斷是否屬佰事達訂單-->跳下一筆
'    If Trim(rsMainT17("品號")) = "60400119" Then intNotBest = intNotBest + 1: GoTo nextRow17
                  
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT17("SAP DN NO."))) Then
        strOrderNo = UCase(Trim(rsMainT17("SAP DN NO.")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '檢查是否有此客戶資料
        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LAPP01' and consigneekey = '" & Right(myExCharFilter(Trim(rsMainT17("客戶代號"))), 7) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        
        If tmp_Rs.EOF Then
            '無此客戶名稱則新增
            strConsigneeKey = Right(myExCharFilter(Trim(rsMainT17("客戶代號"))), 7)
            
            '新增客戶主檔
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho,channel) " & _
            " values('LAPP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT17("客戶名稱"))) & "','" & myExCharFilter(Trim(rsMainT17("客戶名稱"))) & "','','','" & myExCharFilter(Trim(rsMainT17("送貨地址"))) & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT17("客戶通路別"))) & "' ) ", RowsAffect, adExecuteNoRecords
            '紀錄新增之客戶編號
            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
        Else
'            '比對聯絡人、電話與到貨地址是否相符
'            Call Confirm_Recordset_Closed(rsTmp)
'            str_SQL = "select * from trp01m " & _
'                        "where storerkey = 'LAPP01' and full_name = '" & myExCharFilter(Trim(rsMainT17("客戶名稱"))) & "' " & _
'                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT17("送貨地址"))) & "' "
'            rsTmp.CursorLocation = 3
'            rsTmp.Open str_SQL, cn
'
'                If rsTmp.EOF Then
'                    '聯絡人、電話與到貨地址不符
'                strConsigneekey = myExCharFilter(Trim(rsMainT17("客戶代號")))
'
'                    '新增客戶主檔
'                    cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes,updatesource,addwho,editwho,channel) " & _
'                    " values('LAPP01','','" & strConsigneekey & "','" & myExCharFilter(Trim(rsMainT17("客戶名稱"))) & "','" & myExCharFilter(Trim(rsMainT17("客戶名稱"))) & "','','','" & myExCharFilter(Trim(rsMainT17("送貨地址"))) & "','','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT17("客戶通路別"))) & "' ) ", RowsAffect, adExecuteNoRecords
'
'                    '紀錄新增之客戶編號
'                    strNewConsigneekey = strNewConsigneekey & strConsigneekey & "','"
'                Else
                    
                    '相符沿用舊客編
                    strConsigneeKey = Trim(tmp_Rs("consigneekey"))
                    blCustomerMatch = True
'
'                End If
'            rsTmp.Close
        End If
        tmp_Rs.Close
    
        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select orderkey from orders where storerkey = 'LAPP01' and isnull(type,'') <> '刪單' and rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT17("SAP DN NO."))) & "'"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
            strFacility = "佰事達北倉"
            If Right(UCase(Trim(rsMainT17("DC代號"))), 1) = "C" Then strFacility = "佰事達中倉"
            If Right(UCase(Trim(rsMainT17("DC代號"))), 1) = "S" Then strFacility = "佰事達南倉"

            
            arrTmp = Split(Trim(rsMainT17("預計出貨日")), "/")
            If Len(Trim(rsMainT17("預計出貨日"))) = 0 Then
                cn.RollbackTrans: Tran_Level = 0
                msg_text = "訂單號碼:" & Trim(rsMainT17("SAP DN NO.")) & "，品號:" & Trim(rsMainT17("商品編碼")) & Chr(13) & "資料沒有預計出貨日！請向廠商確認"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT17.Enabled = True: cmdImportT17.Enabled = True
                Exit Sub
            Else
                strDate = Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2)
            End If
            
            arrTmp = Split(Trim(rsMainT17("單據產生日")), "/")
            strOrderDate = Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2)
            Dim intPointer As Integer
            intPointer = 1
            'updatesource
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT17("SAP DN NO."))) & "','I','LAPP01','" & strOrderDate & "','" & strDate & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT17("客戶名稱"))) & "','','','','','','" & myExCharFilter(Trim(GetWord(rsMainT17("送貨地址"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT17("送貨地址"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT17("備注"))) & "','" & filLocalFileT17.FileName & "','','" & User_id & "','" & User_id & "','','" & Right(myExCharFilter(Trim(rsMainT17("客戶代號"))), 7) & "') "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

            
            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
            'Mark by Eric因為gs_ordersupdate就會更新zip了
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LAPP01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT17("SAP DN NO.")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
        
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            
            lngCasecnt = 1
            
'            '單位換算
'            If Left(myExCharFilter(Trim(rsMainT17("單位"))), 1) = "箱" Then
'
                '取箱包轉換率
                str_SQL = "select p.casecnt, p.innerpack,p.packkey from " & strWMSDB & "..sku s (nolock) join " & strWMSDB & "..pack p (nolock) on s.packkey = p.packkey where s.storerkey = 'LAPP01' and s.sku = '" & myExCharFilter(Trim(rsMainT17("商品編碼"))) & "'"
                
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                lngCasecnt = tmp_Rs("casecnt")      '大單位入數
                lngInnerpack = tmp_Rs("innerpack")  '中單位入數
                Str_packkey = tmp_Rs("packkey")  'packkey
                tmp_Rs.Close
'
'            End If
            
            
'            If UCase(rsMainT17("SAP訂貨單位")) = "BDL" Then intQTY = Val(rsMainT17("商品訂購數量")) * lngInnerpack  '中單位入數，找出pack資料表
'            If UCase(rsMainT17("SAP訂貨單位")) = "KAR" Then intQTY = Val(rsMainT17("商品訂購數量")) * lngCasecnt  '大單位入數，

            
            intQTY = Val(rsMainT17("商品訂購數量"))

            '取trp19m倉別對照表
            str_SQL = "select bestlot06 from trp19m(nolock) where storerkey = 'LAPP01' and storerlot06 = '" & rsMainT17("DC代號") & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

            If tmp_Rs.EOF Then
            '沒有找到對應的倉別
            '沿用訂單上的倉別
                strLot06 = rsMainT17("DC代號") '待確認
            Else
            '有找到對應的倉別
                strLot06 = Trim(tmp_Rs("bestlot06"))
            End If

            tmp_Rs.Close
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,Packkey)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT17("SAP DN NO."))) & "','" & myExCharFilter(Trim(rsMainT17("商品編碼"))) & "','LAPP01'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT17("SAP訂貨單位"))) & "','0','" & Str_packkey & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            'Mark by Eric 20141216抓取箱入數時順便抓取packkey寫入
'            '更新packkey
'            str_SQL = "update orderdetail " & _
'            "Set orderdetail.packkey = sku.packkey " & _
'            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
'            "where orderkey = '" & str_Orderkey & "' "
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT17.MoveNext
Loop

next18:

'執行gs_ordersupdate   用客戶主檔更新訂單資料

cn.Execute "exec gs_ordersupdate 'LAPP01'", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

'將緊急訂單標記在orders..urgent_mark欄位
'檢查是否有緊急訂單?
str_SQL = "select orderkey " & _
"From orders(nolock) " & _
"where storerkey = 'LAPP01' and priority = 'I' and updatesource = '" & filLocalFileT17.FileName & "' and " & _
"((convert(varchar(8),adddate,114) > '17:00:00' and convert(varchar(8),deliverydate,112) = convert(varchar(8),getdate()+1,112)) or " & _
"(convert(varchar(8),adddate,114) > '17:30:00' and convert(varchar(8),deliverydate,112) = convert(varchar(8),getdate()+2,112) ) or " & _
"(convert(varchar(8),adddate,112) = convert(varchar(8),deliverydate,112)) " & _
") "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If Not tmp_Rs.EOF Then  '有回傳
    x = MsgBox("發現緊急訂單，是否自動將訂單更新為緊急訂單?", vbQuestion + vbYesNo, "APP訂單匯入") '按下的是確定6或是取消
    If x = 6 Then
           '更新urgent_mark欄位V:緊急訂單
           cn.Execute "exec es_update_urgent_mark 'LAPP01','" & filLocalFileT17.FileName & "'", RowsAffect, adExecuteNoRecords
    End If
End If

'檢查有無異常訂單  倉別庫別錯誤 es_Checklot06_by_storer '貨主','訂單檔名'
str_SQL = "exec es_Checklot06_by_storer 'LAPP01','" & filLocalFileT17.FileName & "'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If Not tmp_Rs.EOF Then  '有回傳錯誤的訂單資料，產生excel
    Recordset2Excel "配送倉別與明細倉別不符的訂單資料", tmp_Rs
End If

tmp_Rs.Close


'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "匯入 非佰事達訂單 " & intNotBest & " 筆明細" & Chr(13) & "TMS單號： " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "非佰事達訂單 " & intNotBest & " 筆明細，檔案 " & filLocalFileT17.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT17.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 品號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = 'LAPP01' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\LAPP01\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\LAPP01\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LAPP01\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'備份至FTP
If Dir("O:\LAPP01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LAPP01\OrdersBackup"
FileCopy strTranFileName, "O:\LAPP01\OrdersBackup\" & filLocalFileT17.FileName

'備份檔案
If Dir("C:\BEST\LAPP01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LAPP01\Orders\Backup"
If Dir("C:\BEST\LAPP01\Orders\Backup\" & filLocalFileT17.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LAPP01\Orders\Backup\" & filLocalFileT17.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LAPP01\Orders\Backup\" & mySplit(filLocalFileT17.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT17.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT17.Refresh:
Screen.MousePointer = 0: cmdImportT17.Enabled = True: dgMainT17.Enabled = True
Exit Sub

err_Handle:
   Call ErrorMsgbox(App.title, err.Number, err.Description, "檔案名稱： " & strTranFileName)
    cmdImportT17.Enabled = True: Screen.MousePointer = 0: dgMainT17.Enabled = True
End Sub

Private Sub cmdImportT18_Click()
Dim bl_Error As Boolean '記錄有小數點的旗標
Dim str_Error As String '記錄有小數點錯誤的資料
Dim x As Long
bl_Error = False: str_Error = ""
Dim str_Storerkey As String

If rsMainT18 Is Nothing Then Exit Sub
If rsMainT18.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT18.Enabled = False: cmdImportT18.Enabled = False
strTranFileName = filLocalFileT18.Path & "\" & filLocalFileT18.FileName

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where rtrim(updatesource)='" & filLocalFileT18.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT18.RecordCount = 0 Or rsMainT18 Is Nothing Then
Else
rsMainT18.MoveFirst
str_Storerkey = "LPSI01"

Do While Not rsMainT18.EOF
    '到貨日期檢查
    If Len(Trim(rsMainT18("交貨日期"))) = 0 Then
         MsgBox "訂單號碼:" & Trim(rsMainT18("交貨")) & "的到貨日為空白，訂單轉入終止!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT18("交貨日期"))) > 0 And Len(Trim(rsMainT18("交貨日期"))) < 8 Then
         MsgBox "訂單號碼:" & Trim(rsMainT18("交貨")) & "的到貨日:" & Trim(rsMainT18("交貨日期")) & "，格式不對，請補齊8碼，訂單轉入終止!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
    ElseIf IsDate(Left(Trim(rsMainT18("交貨日期")), 4) + "/" + Mid(Trim(rsMainT18("交貨日期")), 6, 2) + "/" + Right(Trim(rsMainT18("交貨日期")), 2)) = False Then
         MsgBox "訂單號碼:" & Trim(rsMainT18("交貨")) & "的到貨日:" & Trim(rsMainT18("交貨日期")) & "，不是一個正常日期，訂單轉入終止!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
    Else
        '檢查到貨日不可小於今日
        If Trim(rsMainT18("交貨日期")) < Format(Now, "YYYY.MM.DD") Then
            If blAdmin = True Then
            
            '最高權限不檢查到貨日
                 x = MsgBox("到貨日小於今日，你確定要繼續嗎?", vbQuestion + vbYesNo, "最高權限到貨日檢查") '紀錄按下的是確定或是取消
                    If x = 6 Then
                        '繼續
                    Else
                        '離開
                         dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
                    End If
            Else
            
            End If
        End If
    End If

    '訂單日檢查
    If Len(Trim(rsMainT18("供貨日期"))) = 0 Then
         MsgBox "訂單號碼:" & Trim(rsMainT18("交貨")) & "的訂單日為空白，訂單轉入終止!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT18("供貨日期"))) > 0 And Len(Trim(rsMainT18("供貨日期"))) < 8 Then
         MsgBox "訂單號碼:" & Trim(rsMainT18("交貨")) & "的訂單日:" & Trim(rsMainT18("供貨日期")) & "，格式不對，請補齊8碼，訂單轉入終止!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
    ElseIf IsDate(Left(Trim(rsMainT18("供貨日期")), 4) + "/" + Mid(Trim(rsMainT18("供貨日期")), 6, 2) + "/" + Right(Trim(rsMainT18("供貨日期")), 2)) = False Then
         MsgBox "訂單號碼:" & Trim(rsMainT18("交貨")) & "的訂單日:" & Trim(rsMainT18("供貨日期")) & "，不是一個正常日期，訂單轉入終止!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub

    Else
        If Trim(rsMainT18("供貨日期")) > Trim(rsMainT18("交貨日期")) Then MsgBox "訂單號碼:" & Trim(rsMainT18("交貨")) & "的訂單日:" & Trim(rsMainT18("供貨日期")) & "，大於到貨日，訂單轉入終止!", 16, Me.Caption: dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
    End If
    
    '數量檢查
    If Val(rsMainT18("交貨數量")) < 1 Then
        MsgBox "數量小於1，" & Trim(rsMainT18("交貨")) & "-品號：" & Trim(rsMainT18("物料")) & "，訂單轉入終止!!請確認!!", , "訂單檔匯入": dgMainT18.Enabled = True: cmdImportT18.Enabled = True: Exit Sub
        Exit Sub
    End If
    
        '資料檢驗 --判斷SKU是否存在
        str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMainT18("物料")) & "' and Storerkey = 'LPSI01' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '按鈕那些要改
            MsgBox "訂單發現新品號 (" & Trim(rsMainT18("物料")) & ")，請先於商品主檔新建商品資料，訂單轉入終止!!": cmdImportT18.Enabled = True: Screen.MousePointer = 0
            dgMainT18.Enabled = True: cmdImportT18.Enabled = True
            Exit Sub
        End If
        
        '檢查提貨客戶編號是否存在
            str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT18("工廠")) & "' and Storerkey = 'LPSI01' "
        
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        
            If tmp_Rs.EOF Then  '按鈕那些要改
                MsgBox "訂單發現新客戶編號 ( " & Trim(rsMainT18("工廠")) & " )，請先於客戶主檔新建客戶資料，訂單轉入終止!!": cmdImportT18.Enabled = True: Screen.MousePointer = 0
                dgMainT18.Enabled = True: cmdImportT18.Enabled = True
                Exit Sub
            End If
        
        '檢查到貨客戶編號是否存在
        str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT18("收貨人")) & "' and Storerkey = 'LPSI01' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '按鈕那些要改
            MsgBox "訂單發現新客戶編號 ( " & Trim(rsMainT18("收貨人")) & " )，請先於客戶主檔新建客戶資料，訂單轉入終止!!": cmdImportT18.Enabled = True: Screen.MousePointer = 0
            dgMainT18.Enabled = True: cmdImportT18.Enabled = True
            Exit Sub
        End If
        
        '檢查數量有無小數點
        If InStr(Trim(rsMainT18("交貨數量")), ".") <> 0 Then
            str_Error = "訂單號碼:" & Trim(rsMainT18("交貨")) & "，品號:" & Trim(rsMainT18("物料")) & Chr(13) & str_Error
            bl_Error = True
        End If
        
    rsMainT18.MoveNext
Loop
rsMainT18.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "資料有小數點！請重新匯入"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT18.Enabled = True: cmdImportT18.Enabled = True
                Exit Sub
End If

Tran_Level = cn.BeginTrans: cmdImportT18.Enabled = False: dgMainT18.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String
Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String

'開始匯入
Do While Not rsMainT18.EOF
    DoEvents: DoEvents
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT18("交貨"))) Then
        strOrderNo = UCase(Trim(rsMainT18("交貨")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '已經檢查過，則不比對，直接抓出客戶主檔中的 客戶編號，郵遞區號，連絡人，電話，地址
        If myExCharFilter(Trim(rsMainT18("出貨位置"))) = "北區" Then
            '單別為C
            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = 'LPSI01' and consigneekey = '" & myExCharFilter(Trim(rsMainT18("收貨人"))) & "'"
        Else
            '單別為A2B則，抓提貨客編，非A2B則抓到貨客編
            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = 'LPSI01' and consigneekey = '" & myExCharFilter(Trim(rsMainT18("工廠"))) & "'"
        End If
        
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn

        '相符沿用舊客編
        strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
        strZip = myExCharFilter(Trim(tmp_Rs("zip")))
        strContact = myExCharFilter(Trim(tmp_Rs("contact")))
        strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
        strAddress = myExCharFilter(Trim(tmp_Rs("address")))
        strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
        blCustomerMatch = True
        tmp_Rs.Close
        
        strFacility = strShort_name
        
        'C單取配送倉別
        If myExCharFilter(Trim(rsMainT18("出貨位置"))) = "北區" Then
'            str_SQL = "select short_name=isnull(short_name,'') from trp01m where storerkey = 'LPSI01' and consigneekey = '" & myExCharFilter(Trim(rsMainT18("工廠"))) & "'"
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.CursorLocation = 3
'        tmp_Rs.Open str_SQL, cn
        
        strFacility = "佰事達北倉"
        
'        tmp_Rs.Close
        End If

        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select externorderkey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT18("交貨"))) & "' and storerkey = 'LPSI01' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
                        
            strOrderDate = Trim(rsMainT18("供貨日期"))
            Dim intPointer As Integer
            intPointer = 1
            'updatesource
            
            If myExCharFilter(Trim(rsMainT18("出貨位置"))) = "北區" Then
            
                '北區用C單
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,b_company,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT18("交貨"))) & "','C','LPSI01','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT18("交貨日期"))) & "','" & strFacility & "','" & _
                myExCharFilter(Trim(rsMainT18("收貨人"))) & "','','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','" & Trim(rsMainT18("採購文件")) & "','','" & filLocalFileT18.FileName & "','','" & User_id & "','" & User_id & "','','','','0') "
            
            Else
                '中南區用A2B，多紀錄一個B點的客編B_company
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,b_company,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT18("交貨"))) & "','A2B','LPSI01','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT18("交貨日期"))) & "','" & strFacility & "','" & _
                myExCharFilter(Trim(rsMainT18("工廠"))) & "','" & myExCharFilter(Trim(rsMainT18("收貨人"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','" & Trim(rsMainT18("採購文件")) & "','','" & filLocalFileT18.FileName & "','','" & User_id & "','" & User_id & "','','','','0') "
            End If
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

            
'            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & Str_Storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT18("交貨")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
        
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            
            intQTY = Val(rsMainT18("交貨數量"))
            
            
            '訂單明細資料新增
'            If Trim(rsMainT18("單位名稱")) = "箱" Or Trim(rsMainT18("單位名稱")) = "CS" Or Trim(rsMainT18("單位名稱")) = "CASE" Then
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM)" & _
            " select '" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT18("交貨"))) & "','" & myExCharFilter(Trim(rsMainT18("物料"))) & "','LPSI01'," & _
            "'" & intQTY & "' * p.casecnt ,'" & intQTY & "' * p.casecnt,'','" & strFacility & "',''" & _
            "from " & strWMSDB & "..sku s join " & strWMSDB & "..pack p on s.packkey = p.packkey and s.sku = '" & Trim(rsMainT18("物料")) & "' and s.storerkey = 'LPSI01' "
'            Else
'                 str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM)" & _
'                "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT18("訂單號碼"))) & "','" & myExCharFilter(Trim(rsMainT18("品號"))) & "','" & myExCharFilter(Trim(rsMainT18("貨主"))) & "'," & _
'                "'" & intQTY & "','" & intQTY & "','" & myExCharFilter(Trim(rsMainT18("倉別"))) & "','','')"
'           End If
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT18.MoveNext
Loop

'執行gs_ordersupdate   用客戶主檔更新訂單資料

'cn.Execute "exec gs_ordersupdate 'LPSI01'", RowsAffect, adExecuteNoRecords

'檢查有無異常訂單  倉別庫別錯誤 es_Checklot06_by_storer '貨主','訂單檔名'
'str_SQL = "exec es_Checklot06_by_storer '" & Str_Storerkey & "','" & filLocalFileT18.FileName & "'"
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'If Not tmp_Rs.EOF Then  '有回傳錯誤的訂單資料，產生excel
'    Recordset2Excel "配送倉別與明細倉別不符的訂單資料", tmp_Rs
'End If
'
'tmp_Rs.Close

cn.CommitTrans: Tran_Level = 0: dgMainT18.Enabled = True


'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "TMS單號： " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "檔案 " & filLocalFileT18.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT18.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 品號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'備份檔案
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT18.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT18.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT18.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT18.FileName, ".", -1)
End If

'備份至FTP
If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT18.FileName



Kill strTranFileName
    
filLocalFileT18.Refresh:
Screen.MousePointer = 0: cmdImportT18.Enabled = True: dgMainT18.Enabled = True
Exit Sub

err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
   Call ErrorMsgbox(App.title, err.Number, err.Description, "檔案名稱： " & strTranFileName)
    cmdImportT18.Enabled = True: Screen.MousePointer = 0: dgMainT18.Enabled = True

End Sub

Private Sub cmdImportT19_Click()
'
'APP退貨訂單的匯入程式 create by Eric 20130613
'開頭檢查匯入的訂單:
'1.到貨日是否大於今日 2.數量是否<0 3.數量是否有小數點 4.品號是否存在
'檢查訂單客戶編號 , 系統是否存在, 存在則帶出系統客戶主檔中的資料, 不存在則使用訂單上的客戶編號進行新增 (不比對客戶主檔資料，進行重編)
'程式結束後 , 會執行ordersupdate, 系統客戶主檔, 更新訂單資料

Dim str_Error As String
Dim bl_Error As Boolean
On Error GoTo err_Handle
strTranFileName = filLocalFileT19.Path & "\" & filLocalFileT19.FileName
If Len(RTrim(cboSheetT19)) = 0 Then MsgBox "請選擇工作表", 64, Me.Caption: Exit Sub
If rsMainT19.EOF Or rsMainT19 Is Nothing Then Exit Sub

Screen.MousePointer = 11: SSTab2.Enabled = False: cmdImportT19.Enabled = False


'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where rtrim(updatesource)='" & filLocalFileT19.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp
rsMainT19.MoveFirst

Do While Not rsMainT19.EOF

    '到貨日期檢查
    arrTmp = Split(Trim(rsMainT19("預計交貨日期")), ".")
    If Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then:  MsgBox "預計交貨日期小於今日，訂單轉入終止!", 16, Me.Caption: cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0: Exit Sub
    
    '數量檢查
    If Trim(rsMainT19("商品訂購數量")) < 1 Then
        MsgBox "發現商品訂購數量小於1，" & "交貨單號:" & Trim(rsMainT19("交貨單號")) & "-客戶代號:" & Trim(rsMainT19("客戶代號")) & "-客戶名稱:" & Trim(rsMainT19("客戶名稱")) & "-商品訂購數量:" & Trim(rsMainT19("商品訂購數量")) & ")，訂單轉入終止!!", , "退貨單匯入": cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0: Exit Sub
        Exit Sub
    End If
    
    '檢查數量有無小數點
    If InStr(Trim(rsMainT19("商品訂購數量")), ".") <> 0 Then
        str_Error = "訂單號碼:" & Trim(rsMainT19("交貨單號")) & "，品號:" & Trim(rsMainT19("品號")) & Chr(13) & str_Error
        bl_Error = True
    End If
    
    '資料檢驗 --判斷SKU是否存在
    str_SQL = "select sku from " & strWMSDB & "..sku(nolock) where sku='" & Trim(rsMainT19("品號")) & "' and Storerkey = 'LAPP01' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

    If tmp_Rs.EOF Then  '按鈕那些要改
        MsgBox "訂單發現新品號 (" & Trim(rsMainT19("品號")) & ")，訂單轉入終止!!":
        cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0:
        Exit Sub
    End If
    
    rsMainT19.MoveNext
Loop

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "資料有小數點！請重新匯入"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0:
                Exit Sub
End If

Tran_Level = cn.BeginTrans: cmdImportT19.Enabled = False: dgMainT19.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim str_Fullname As String, str_Contact As String, str_Phone As String, Str_Address As String, str_Channel As String
            
'取最後客戶編號
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m(nolock) where storerkey = 'LAPP01' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))

tmp_Rs.Close

rsMainT19.MoveFirst
Do While Not rsMainT19.EOF
    DoEvents: DoEvents
    
'    '資料檢驗--判斷是否屬佰事達訂單-->跳下一筆
'    If UCase(Trim(rsMainT19("倉庫"))) = "" Then
''        MsgBox "客戶單號：" & Trim(rsMainT4("銷貨單號")) & "( " & Trim(rsMainT4("倉庫")) & " )" & vbCrLf & "請通知客戶，確認該筆訂單是否有誤!?", vbOKOnly, "非佰事達之訂單不轉入"
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
        '資料檢驗--判斷SKU是否存在
'    str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT19("品號")) & "' and Storerkey = 'LAPP01' "
'
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If tmp_Rs.EOF Then
'        cn.RollbackTrans: Tran_Level = 0
'        MsgBox "訂單發現新品號 (" & Trim(rsMainT19("品號")) & " ) " & Trim(rsMainT19("品名")) & "，訂單轉入終止!!": cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0
'        Exit Sub
'    End If

'    End If
                  
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT19("交貨單號"))) Then
        strOrderNo = UCase(Trim(rsMainT19("交貨單號")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '檢查是否有此客戶編號，(帶出系統 客戶名稱、聯絡人、電話、地址、通路)
        str_SQL = "select consigneekey,full_name,contact=isnull(contact,''),phone=isnull(phone,''),address,channel=isnull(channel,'') from trp01m(nolock) where  storerkey = 'LAPP01'  and consigneekey = '" & myExCharFilter(Trim(rsMainT19("客戶代號"))) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn

        
        If tmp_Rs.EOF Then
            '無此客戶名稱則新增
            intTmp = intTmp + 1
            strConsigneeKey = myExCharFilter(Trim(rsMainT19("客戶代號")))
            
            '新增客戶主檔
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho,channel) " & _
            " values('LAPP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT19("客戶名稱"))) & "','" & myExCharFilter(Trim(rsMainT19("客戶名稱"))) & "','','','" & myExCharFilter(Trim(rsMainT19("收貨地址"))) & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT19("通路"))) & "') ", RowsAffect, adExecuteNoRecords
'
            '紀錄新增之客戶編號
            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
        Else
'            '確認程式是否要用訂單地址，要的話則重編客戶主檔，否則則直接使用系統客戶編號的資料
'            '比對 (電話、到貨地址) 是否相符
'            Call Confirm_Recordset_Closed(rsTmp)
'            str_SQL = "select * from trp01m " & _
'                        "where storerkey = 'LAPP01' and full_name = '" & myExCharFilter(Trim(rsMainT19("客戶名稱"))) & "' " & _
'                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT19("收貨地址"))) & "' "
'            rsTmp.CursorLocation = 3
'            rsTmp.Open str_SQL, cn
'
'            If rsTmp.EOF Then
'                '電話與地址不符則重編
'                intTmp = intTmp + 1
'                strConsigneeKey = "BEST" & Format(intTmp, "000000")
'
'                '新增客戶主檔
'                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes,updatesource,addwho,editwho) " & _
'                " values('LAPP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT19("客戶名稱"))) & "','" & myExCharFilter(Trim(rsMainT19("客戶名稱"))) & "','','','" & myExCharFilter(Trim(rsMainT19("收貨地址"))) & "','','','" & User_id & "','" & User_id & "') ", RowsAffect, adExecuteNoRecords
'
'                '紀錄新增之客戶編號
'                strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'            Else '相符沿用舊客編
'
'                str_Fullname = Trim(tmp_Rs("full_name").Value)
'                str_Contact = Trim(tmp_Rs("contact").Value)
'                str_Phone = Trim(tmp_Rs("phone").Value)
'                str_Address = Trim(tmp_Rs("address").Value)
'                str_Channel = Trim(tmp_Rs("channel").Value)
'                strConsigneeKey = Trim(rsTmp("consigneekey"))
'                blCustomerMatch = True
'
'            End If
'                rsTmp.Close
                str_Fullname = Trim(tmp_Rs("full_name").Value)
                str_Contact = Trim(tmp_Rs("contact").Value)
                str_Phone = Trim(tmp_Rs("phone").Value)
                Str_Address = Trim(tmp_Rs("address").Value)
                str_Channel = Trim(tmp_Rs("channel").Value)
                strConsigneeKey = Trim(tmp_Rs("consigneekey"))
                blCustomerMatch = True

        End If
        tmp_Rs.Close
    
'        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select orderkey from orders where storerkey = 'LAPP01' and isnull(type,'') <> '刪單' and rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT19("交貨單號"))) & "'"
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
    
            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '收貨庫存地  --確認後開放
            strFacility = "佰事達北倉"
            If Right(myExCharFilter(Trim(rsMainT19("收貨庫存地"))), 1) = "C" Then strFacility = "佰事達中倉"
            If Right(myExCharFilter(Trim(rsMainT19("收貨庫存地"))), 1) = "S" Then strFacility = "佰事達南倉"
            
            arrTmp = Split(Trim(rsMainT19("預計交貨日期")), ".")
            strDeliveryDate = Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2)   '到貨日期
            strOrderDate = Format(Now, "YYYY/MM/DD")    '訂單日期
            Dim intPointer As Integer
            intPointer = 1
            
            '訂單日期待確認
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT19("交貨單號"))) & "','R','LAPP01','" & strOrderDate & "','" & strDeliveryDate & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & str_Fullname & "','" & str_Contact & "','','','" & str_Phone & "','','" & myExCharFilter(Trim(GetWord(Str_Address, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(Str_Address, intPointer, 45))) & "','','','" & filLocalFileT19.FileName & "','','" & User_id & "','" & User_id & "','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
            'If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LAPP01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT19("交貨單號")) & "','"
            blDuplicationOrder = True

        End If
    End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Abs(Trim(rsMainT19("商品訂購數量")))
            strLot06 = (Trim(rsMainT19("收貨庫存地")))
            
            '取trp19m倉別對照表
            str_SQL = "select bestlot06 from trp19m(nolock) where storerkey = 'LAPP01' and storerlot06 = '" & Trim(rsMainT19("收貨庫存地")) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

            If tmp_Rs.EOF Then
            '沒有找到對應的倉別
            '沿用訂單上的倉別
                strLot06 = (Trim(rsMainT19("收貨庫存地")))
            Else
            '有找到對應的倉別
                strLot06 = Trim(tmp_Rs("bestlot06"))
            End If

            tmp_Rs.Close
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,addwho,editwho)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT19("交貨單號"))) & "','" & myExCharFilter(Trim(rsMainT19("品號"))) & "','LAPP01'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','','0','" & User_id & "','" & User_id & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT19.MoveNext
Loop

'執行gs_ordersupdate   用客戶主檔更新訂單資料

cn.Execute "exec gs_ordersupdate 'LAPP01'", RowsAffect, adExecuteNoRecords

cmdImportT19.Enabled = True: dgMainT19.Enabled = True: Screen.MousePointer = 0:

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細 (共 " & rsMainT19.RecordCount & " 筆明細)，TMS單號： " & strOrderKeyS & "~" & str_Orderkey & "，檔案 " & filLocalFileT19.FileName & " 備份於 C:\BEST\LAPP01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細，檔案 " & filLocalFileT19.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT19.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 料號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = 'LAPP01' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\LAPP01\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\LAPP01\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LAPP01\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'備份至FTP
If Dir("O:\LAPP01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LAPP01\OrdersBackup"
FileCopy strTranFileName, "O:\LAPP01\OrdersBackup\" & filLocalFileT19.FileName

'備份檔案
If Dir("C:\BEST\LAPP01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LAPP01\Orders\Backup"
If Dir("C:\BEST\LAPP01\Orders\Backup\" & filLocalFileT19.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LAPP01\Orders\Backup\" & filLocalFileT19.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LAPP01\Orders\Backup\" & mySplit(filLocalFileT19.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT19.FileName, ".", -1)
End If

filLocalFileT19.Refresh: cboSheetT19.Clear

Kill strTranFileName

Screen.MousePointer = 0: cmdImportT19.Enabled = True: dgMainT19.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    cmdImportT19.Enabled = True: Screen.MousePointer = 0: dgMainT19.Enabled = True
    Call ErrorMsgbox(App.title, err.Number, err.Description, "檔案名稱： " & strTranFileName)
End Sub

Private Sub cmdImportT20_Click() 'Terry 20180825 特力屋訂單匯入
Dim bl_Error As Boolean '記錄有小數點的旗標
Dim str_Error As String '記錄有小數點錯誤的資料
Dim x As Long
bl_Error = False: str_Error = ""
Dim str_Storerkey As String

If rsMainT20 Is Nothing Then Exit Sub
If rsMainT20.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT20.Enabled = False: cmdImportT20.Enabled = False
strTranFileName = filLocalFileT20.Path & "\" & filLocalFileT20.FileName

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT20.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT20.Enabled = True: dgMainT20.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT20.RecordCount = 0 Or rsMainT20 Is Nothing Then
Else
rsMainT20.MoveFirst
str_Storerkey = "LTRI03"
Do While Not rsMainT20.EOF
    '到貨日期檢查
    If Len(Trim(rsMainT20("到貨日"))) = 0 Then
    
        'Terry 需確認單別資料 修改中
        If Trim(rsMainT20("訂單類別")) = "R" Then
        Else
            MsgBox "訂單號碼:" & Trim(rsMainT20("訂單號碼")) & "的到貨日為空白，訂單轉入終止!", 16, Me.Caption: dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
        End If
    ElseIf Len(Trim(rsMainT20("到貨日"))) > 0 And Len(Trim(rsMainT20("到貨日"))) < 8 Then
         MsgBox "訂單號碼:" & Trim(rsMainT20("訂單號碼")) & "的到貨日:" & Trim(rsMainT20("到貨日")) & "，格式不對，請補齊8碼，訂單轉入終止!", 16, Me.Caption: dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
    Else
        '檢查到貨日不可小於今日
        If Trim(rsMainT20("到貨日")) < Format(Now, "YYYYMMDD") Then
            If blAdmin = True Then
            
            '最高權限不檢查到貨日
                 x = MsgBox("到貨日小於今日，你確定要繼續嗎?", vbQuestion + vbYesNo, "最高權限到貨日檢查") '紀錄按下的是確定或是取消
                    If x = 6 Then
                        '繼續
                    Else
                        '離開
                         dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
                    End If
            Else
            
            End If
        End If
    End If
    
'    '貨主檢查
'    If Len(Trim(rsMainT20("貨主"))) = 0 Or Trim(rsMainT20("貨主")) <> "LCHF01" Then
'        MsgBox "訂單號碼:" & Trim(rsMainT20("訂單號碼")) & "的貨主有誤，訂單轉入終止!", 16, Me.Caption: dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
'    End If
    
    
    '訂單日檢查
    If Len(Trim(rsMainT20("訂單日"))) = 0 Then
        MsgBox "訂單號碼:" & Trim(rsMainT20("訂單號碼")) & "的訂單日為空白，訂單轉入終止!", 16, Me.Caption: dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT20("訂單日"))) > 0 And Len(Trim(rsMainT20("訂單日"))) < 8 Then
         MsgBox "訂單號碼:" & Trim(rsMainT20("訂單號碼")) & "的訂單日:" & Trim(rsMainT20("訂單日")) & "，格式不對，請補齊8碼，訂單轉入終止!", 16, Me.Caption: dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
    ElseIf Trim(rsMainT20("訂單日")) > Trim(rsMainT20("到貨日")) Then
        MsgBox "訂單號碼:" & Trim(rsMainT20("訂單號碼")) & "的訂單日:" & Trim(rsMainT20("訂單日")) & "，大於到貨日，訂單轉入終止!", 16, Me.Caption: dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
    End If
    
    '數量檢查
    If Val(rsMainT20("數量")) < 1 Then
        MsgBox "數量小於1，" & Trim(rsMainT20("訂單號碼")) & "-品號：" & Trim(rsMainT20("品號")) & "，訂單轉入終止!!請確認!!", , "訂單檔匯入": dgMainT20.Enabled = True: cmdImportT20.Enabled = True: Exit Sub
        Exit Sub
    End If
    
        '資料檢驗 --判斷SKU是否存在
        str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMainT20("品號")) & "' and Storerkey = '" & str_Storerkey & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

        If tmp_Rs.EOF Then  '按鈕那些要改
            MsgBox "訂單發現新品號 (" & Trim(rsMainT20("品號")) & ")，請先於商品主檔新建商品資料，訂單轉入終止!!": cmdImportT20.Enabled = True: Screen.MousePointer = 0
            dgMainT20.Enabled = True: cmdImportT20.Enabled = True
            Exit Sub
        End If
        
        'Terry 特力屋沒有客戶主檔
'        '檢查A2B訂單以外的客戶編號是否存在
'        str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT20("到貨客戶編號")) & "' and Storerkey = '" & str_Storerkey & "' "
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
'        If tmp_Rs.EOF Then  '按鈕那些要改
'            MsgBox "訂單發現新客戶編號 ( " & Trim(rsMainT20("到貨客戶編號")) & " )，請先於客戶主檔新建客戶資料，訂單轉入終止!!": cmdImportT20.Enabled = True: Screen.MousePointer = 0
'            dgMainT20.Enabled = True: cmdImportT20.Enabled = True
'            Exit Sub
'        End If
        
        '檢查數量有無小數點
        If InStr(Trim(rsMainT20("數量")), ".") <> 0 Then
            str_Error = "訂單號碼:" & Trim(rsMainT20("訂單號碼")) & "，品號:" & Trim(rsMainT20("品號")) & Chr(13) & str_Error
            bl_Error = True
        End If
        
        
        'Terry 需確認單別資料 修改中
'        '判斷單別
'        If Trim(rsMainT20("訂單類別")) = "A2B" Then
'            MsgBox "訂單類別為A2B:" & Trim(rsMainT20("訂單類別")) & "，A2B訂單請由公版EXCEL訂單匯入，訂單轉入終止!!": cmdImportT20.Enabled = True: Screen.MousePointer = 0
'                dgMainT20.Enabled = True: cmdImportT20.Enabled = True
'            Exit Sub
'        End If
'
'
'        If Trim(rsMainT20("訂單類別")) = "出貨" Or Trim(rsMainT20("訂單類別")) = "出退" Or Trim(rsMainT20("訂單類別")) = "代退" Or Trim(rsMainT20("訂單類別")) = "代銷" Then
'        Else
'            MsgBox "系統無此單別:" & Trim(rsMainT20("訂單類別")) & "，請確認此訂單類別是否正確，訂單轉入終止!!": cmdImportT20.Enabled = True: Screen.MousePointer = 0
'                dgMainT20.Enabled = True: cmdImportT20.Enabled = True
'            Exit Sub
'        End If
        
    rsMainT20.MoveNext
Loop
rsMainT20.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "資料有小數點！請重新匯入"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT20.Enabled = True: cmdImportT20.Enabled = True
                Exit Sub
End If


Tran_Level = cn.BeginTrans: cmdImportT20.Enabled = False: dgMainT20.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String
Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String


'開始匯入
Do While Not rsMainT20.EOF
    DoEvents: DoEvents
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT20("訂單號碼"))) Then
        strOrderNo = UCase(Trim(rsMainT20("訂單號碼")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        
        'Terry 特力屋沒有客戶主檔
        '已經檢查過，則不比對，直接抓出客戶主檔中的 客戶編號，郵遞區號，連絡人，電話，地址
        '單別為A2B則，抓提貨客編，非A2B則抓到貨客編
        str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & str_Storerkey & "' and consigneekey = 'LTRI03'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn

        Dim str_Priority As String
        '相符沿用舊客編
        strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
        strZip = myExCharFilter(Trim(tmp_Rs("zip")))
        strContact = myExCharFilter(Trim(tmp_Rs("contact")))
        strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
        strAddress = myExCharFilter(Trim(tmp_Rs("address")))
        strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
        
'        'Terry 暫時代空
'        strConsigneeKey = ""
'        strZip = ""
'        strContact = ""
'        strPhone = ""
'        strAddress = ""
'        strShort_name = ""


        
        'Terry 需確認單別資料 修改中
'        '將單別代成I or R
'        If myExCharFilter(Trim(rsMainT20("訂單類別"))) = "出退" Or myExCharFilter(Trim(rsMainT20("訂單類別"))) = "代退" Then
'            str_Priority = "R"
'        Else
'            str_Priority = "I"
'        End If
        
'        blCustomerMatch = True
        tmp_Rs.Close

        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select externorderkey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT20("訂單號碼"))) & "' and storerkey = '" & str_Storerkey & "' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '取訂單號碼
            tmp_Rs.Close
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
'            If UCase(Right(Trim(rsMainT20("倉別")), 2)) = "-C" Then
'                strFacility = "佰事達中倉"
'            ElseIf UCase(Right(Trim(rsMainT20("倉別")), 2)) = "-S" Then
'                strFacility = "佰事達南倉"
'            Else
            strFacility = "佰事達北倉"
'            End If
            
'            If Trim(rsMainT20("倉別")) = "" Then strFacility = ""

            strOrderDate = Trim(rsMainT20("訂單日"))
            Dim intPointer As Integer
            intPointer = 1
            
            Call Confirm_Recordset_Closed(tmp_Rs)
            
            
'            'Terry 需確認單別資料 修改中
'            If (Trim(rsMainT20("訂單類別")) = "出退" Or Trim(rsMainT20("訂單類別")) = "代退") And Len(Trim(rsMainT20("到貨日"))) = 0 Then
'                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
'                "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT20("訂單號碼")) & "','" & str_Priority & "','" & str_Storerkey & "','" & strOrderDate & "','" & strOrderDate & "','" & strFacility & "','" & _
'                Trim(rsMainT20("到貨客戶編號")) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & Trim(GetWord(strAddress, intPointer, 58)) & "','" & Trim(GetWord(strAddress, intPointer, 45)) & "','" & Trim(rsMainT20(["客戶訂單(訂單)"])) & "','" & Trim(rsMainT20("訂單備註")) & "','" & filLocalFileT20.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
'            Else
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT20("訂單號碼")) & "','" & str_Priority & "','" & str_Storerkey & "','" & strOrderDate & "','" & Trim(rsMainT20("到貨日")) & "','" & strFacility & "','" & _
                "LTRI03','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & Trim(GetWord(strAddress, intPointer, 58)) & "','" & Trim(GetWord(strAddress, intPointer, 45)) & "','','" & Trim(rsMainT20("訂單備註")) & "','" & filLocalFileT20.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
'            End If
            
            
'            If (Trim(rsMainT20("訂單類別")) = "出退" Or Trim(rsMainT20("訂單類別")) = "代退") And Len(Trim(rsMainT20("到貨日"))) = 0 Then
'                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
'                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT20("訂單號碼"))) & "','" & str_Priority & "','" & Str_storerkey & "','" & strOrderDate & "','" & strOrderDate & "','" & strFacility & "','" & _
'                myExCharFilter(Trim(rsMainT20("到貨客戶編號"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT20("訂單備註"))) & "','" & filLocalFileT20.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
'            Else
'                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
'                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT20("訂單號碼"))) & "','" & str_Priority & "','" & Str_storerkey & "','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT20("到貨日"))) & "','" & strFacility & "','" & _
'                myExCharFilter(Trim(rsMainT20("到貨客戶編號"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT20("訂單備註"))) & "','" & filLocalFileT20.FileName & "','','" & User_id & "','" & User_id & "','','','','') "
'            End If
            
            
            
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
'            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & Str_Storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT20("訂單號碼")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
        
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            
            intQTY = Abs(Val(rsMainT20("數量")))
            
            '訂單明細資料新增
            If RTrim(rsMainT20("單位名稱")) = "箱" Or RTrim(rsMainT20("單位名稱")) = "CS" Then
                '箱數
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Lottable03)" & _
                "select  '" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Trim(rsMainT20("訂單號碼")) & "','" & Trim(rsMainT20("品號")) & "','LTRI03'," & intQTY & " * p.casecnt ," & intQTY & " * p.casecnt " & _
                ",'R01','" & strFacility & "','" & Trim(rsMainT20("單位名稱")) & "','" & Trim(rsMainT20("目的儲位")) & "' " & _
                "from " & strWMSDB & "..sku s join " & strWMSDB & "..pack p on s.packkey = p.packkey and s.sku = '" & Trim(rsMainT20("品號")) & "' and s.storerkey = 'LTRI03'"
            Else
                '個數
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Lottable03) " & _
                "values ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Trim(rsMainT20("訂單號碼")) & "','" & Trim(rsMainT20("品號")) & "','LTRI03'," & _
                "'" & intQTY & "','" & intQTY & "','R01','" & strFacility & "','" & Trim(rsMainT20("單位名稱")) & "','" & Trim(rsMainT20("目的儲位")) & "')"
            End If
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT20.MoveNext
Loop

'更新允收期
'cn.Execute "exec gs_ordersupdate 'LTRI03'", RowsAffect, adExecuteNoRecords

'檢查有無異常訂單  倉別庫別錯誤 es_Checklot06_by_storer '貨主','訂單檔名'
'str_SQL = "exec es_Checklot06_by_storer '" & Str_Storerkey & "','" & filLocalFileT14.FileName & "'"
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'If Not tmp_Rs.EOF Then  '有回傳錯誤的訂單資料，產生excel
'    Recordset2Excel "配送倉別與明細倉別不符的訂單資料", tmp_Rs
'End If
'
'tmp_Rs.Close

cn.CommitTrans: Tran_Level = 0: dgMainT20.Enabled = True


'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "TMS單號： " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "檔案 " & filLocalFileT20.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT20.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 品號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'備份檔案
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT20.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT20.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT20.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT20.FileName, ".", -1)
End If


'備份至FTP
If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT20.FileName


Kill strTranFileName
    
filLocalFileT20.Refresh:
Screen.MousePointer = 0: cmdImportT20.Enabled = True: dgMainT20.Enabled = True
Exit Sub

err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
   Call ErrorMsgbox(App.title, err.Number, err.Description, "檔案名稱： " & strTranFileName)
    cmdImportT20.Enabled = True: Screen.MousePointer = 0: dgMainT20.Enabled = True
End Sub

Private Sub cmdImportT21_Click()
    If rsMainT21 Is Nothing Then Exit Sub
    On Error GoTo err_Handle
    dgMainT21.Enabled = False: cmdImportT21.Enabled = False
    Dim str_ASNkey As String, int_orderlinenuber As Integer, str_Storerkey As String, strKeycount As String, str_Lottable06 As String
    Dim rsKeycount As New ADODB.Recordset
    Dim bl_Error As Boolean '記錄有小數點的旗標
    Dim str_Error As String '記錄有小數點錯誤的資料
    Dim x As Long
    bl_Error = False: str_Error = ""
    
    str_Storerkey = "LCHF01"
    str_ASNkey = ""
    dgMainT21.Enabled = False: cmdImportT21.Enabled = False
    strTranFileName = filLocalFileT21.Path & "\" & filLocalFileT21.FileName
    
    Do While Not rsMainT21.EOF
        str_SQL = "select externorderkey from orders(nolock) where storerkey = '" & str_Storerkey & "' and externorderkey = '" & Trim(rsMainT21("調撥單號")) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn
        If Not tmp_Rs.EOF Then
            msg_text = "訂單重複:" & Trim(rsMainT21("調撥單號")) & "，請確認資料，謝謝。"
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            dgMainT21.Enabled = True: cmdImportT21.Enabled = True
            tmp_Rs.Close
            Exit Sub
        End If
        rsMainT21.MoveNext
    Loop
    
    rsMainT21.MoveFirst
    
    Do While Not rsMainT21.EOF
        '到貨日期檢查
        If Len(Trim(rsMainT21("調撥日"))) = 0 Then
             MsgBox "交易單號:" & Trim(rsMainT21("調撥單號")) & "的供貨日期為空白，訂單轉入終止!", 16, Me.Caption: dgMainT21.Enabled = True: cmdImportT21.Enabled = True: Exit Sub
        Else
            '檢查到貨日不可小於今日
            If Format(Trim(rsMainT21("調撥日")), "YYYYMMDD") < Format(Now, "YYYYMMDD") Then
                If blAdmin = True Then
                
                '最高權限不檢查到貨日
                     x = MsgBox("調撥日小於今日，你確定要繼續嗎?", vbQuestion + vbYesNo, "最高權限到貨日檢查") '紀錄按下的是確定或是取消
                        If x = 6 Then
                            '繼續
                        Else
                            '離開
                             dgMainT21.Enabled = True: cmdImportT21.Enabled = True: Exit Sub
                        End If
                Else
                
                End If
            End If
        End If
        
        '訂單日檢查
        If Len(Trim(rsMainT21("調撥日"))) = 0 Then
             MsgBox "供貨日期為空白，訂單轉入終止!", 16, Me.Caption: dgMainT21.Enabled = True: cmdImportT21.Enabled = True: Exit Sub
        End If
        
        '數量檢查
        If Val(rsMainT21("數量")) < 1 Then
            MsgBox "數量小於1，" & Trim(rsMainT21("調撥單號")) & "-品號：" & Trim(rsMainT21("產品代號")) & "，訂單轉入終止!!請確認!!", , "訂單檔匯入": dgMainT21.Enabled = True: cmdImportT21.Enabled = True: Exit Sub
            Exit Sub
        End If
        
        '資料檢驗 --判斷SKU是否存在
        If InStr(1, Trim(rsMainT21("品名")), "棧板") Then
        Else
            str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMainT21("產品代號")) & "' and Storerkey = '" & str_Storerkey & "'"
            
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        
            If tmp_Rs.EOF Then
                MsgBox "訂單發現新品號 (" & Trim(rsMainT21("產品代號")) & ")，請先於商品主檔新建商品資料，訂單轉入終止!!": cmdImportT21.Enabled = True: Screen.MousePointer = 0
                dgMainT21.Enabled = True: cmdImportT21.Enabled = True
                Exit Sub
            End If
        End If
        '檢查數量有無小數點
        If InStr(Trim(rsMainT21("數量")), ".") <> 0 Then
            str_Error = "交易單號:" & Trim(rsMainT21("調撥單號")) & "，品號:" & Trim(rsMainT21("產品代號")) & Chr(13) & str_Error
            bl_Error = True
        End If

            
        rsMainT21.MoveNext
    Loop
    rsMainT21.MoveFirst
    
    If bl_Error = True Then
                    msg_text = str_Error & Chr(13) & "資料有小數點！請重新匯入"
                    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                    dgMainT21.Enabled = True: cmdImportT21.Enabled = True
                    Exit Sub
    End If
    
    rsMainT21.MoveFirst
    Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long, IntCasecnt As Integer
    Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim strDate As String, strOrderType As String, strSku As String
    Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String
    Tran_Level = cn.BeginTrans: cmdImportT12.Enabled = False: dgMainT12.Enabled = False
    Do While Not rsMainT21.EOF
        If InStr(1, Trim(rsMainT21("品名")), "棧板") Then
            GoTo NextRow1
        End If
        If str_ASNkey <> Trim(rsMainT21("調撥單號")) Then
            str_ASNkey = Trim(rsMainT21("調撥單號"))
            int_orderlinenuber = 0
            
            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & str_Storerkey & "' and consigneekey = 'LCHF01'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.CursorLocation = 3
            tmp_Rs.Open str_SQL, cn
    '        End If
            '相符沿用舊客編
            strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
            strZip = myExCharFilter(Trim(tmp_Rs("zip")))
            strContact = myExCharFilter(Trim(tmp_Rs("contact")))
            strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
            strAddress = myExCharFilter(Trim(tmp_Rs("address")))
            strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
            tmp_Rs.Close
            
            '資料檢驗--判斷訂單是否重複，重複不增加
            Call Confirm_Recordset_Closed(tmp_Rs)
            str_SQL = "select externorderkey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT21("調撥單號"))) & "' and storerkey = '" & str_Storerkey & "' and isnull(type,'') <> '刪單' "
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            If tmp_Rs.EOF Then
    
                '取訂單號碼
                str_SQL = "select isnull(max(orderkey),0) from orders"
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
                If strOrderKeyS = "" Then
                    strOrderKeyS = str_Orderkey
                End If
                
                tmp_Rs.Close
                
                Dim intPointer As Integer
                intPointer = 1
                
            
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                            "VALUES ('" & str_Orderkey & "','" & RTrim(rsMainT21("調撥單號")) & "','" & "RC" & "','" & str_Storerkey & "',convert(char(10),'" & Trim(rsMainT21("調撥日")) & "',111)," & _
                            "convert(char(10),'" & Trim(rsMainT21("調撥日")) & "',111) ,'" & strFacility & "','LCHF01','" & strShort_name & "','" & strContact & "','','','" & strPhone & "'," & _
                            "'" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & _
                            "','','" & "" & "','" & filLocalFileT12.FileName & "','','" & User_id & "','" & User_id & "','','','" & "RC" & "','" & "" & "') "
    
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                int_Order = int_Order + 1
                int_orderlinenuber = 1
            Else
                tmp_Rs.Close
                
                '訂單重複
                Call FTPlog("訂單重複" & str_SQL)
                '紀錄重複
                strReOrderkey = strReOrderkey & Trim(rsMainT21("調撥單號")) & "','"
                GoTo NextRow1
                
            End If
        End If

        If Trim(rsMainT21("撥出倉庫名稱")) = "良品倉" Then
            str_Lottable06 = "R01"
        End If
        '寫入表身
        str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,Externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable03,Lottable06, Facility,UOM,otherUOM)" & _
                    "select '" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Format(int_orderlinenuber, "00000") & "','" & RTrim(rsMainT21("調撥單號")) & "','" & Trim(rsMainT21("產品代號")) & "','" & str_Storerkey & "'," & _
                    "cast('" & Trim(rsMainT21("數量")) & "' as Int) * p.casecnt,cast('" & Trim(rsMainT21("數量")) & "' as Int) * p.casecnt,CONVERT(CHAR(8),CONVERT(DATETIME,'" & Trim(rsMainT21("備註")) & "',111),112),'R01','" & strFacility & "','',''" & _
                    "FROM " & strWMSDB & "..sku s join " & strWMSDB & "..pack p on s.packkey = p.packkey where s.sku = '" & Trim(rsMainT21("產品代號")) & "' and s.storerkey = '" & str_Storerkey & "'"

        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        '更新packkey
        str_SQL = "update orderdetail " & _
        "Set orderdetail.packkey = sku.packkey " & _
        "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
        "where orderkey = '" & str_Orderkey & "' "
        
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        int_orderlinenuber = int_orderlinenuber + 1
        
        int_OrderLine = int_OrderLine + 1
NextRow1:

        rsMainT21.MoveNext

        Loop
        'Close RecordSet
        rsMainT21.Close: Set rsMainT21 = Nothing
                
                
cn.Execute "exec gs_ordersupdate 'LCHF01'", RowsAffect, adExecuteNoRecords


cn.CommitTrans: Tran_Level = 0: dgMainT21.Enabled = True


'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "TMS單號： " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "檔案 " & filLocalFileT12.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT21.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 品號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'備份檔案
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT21.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT21.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT21.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT21.FileName, ".", -1)
End If

'備份至FTP
If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT21.FileName


Kill strTranFileName
    
filLocalFileT21.Refresh:
Screen.MousePointer = 0: cmdImportT21.Enabled = True: dgMainT21.Enabled = True
Exit Sub

err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
   Call ErrorMsgbox(App.title, err.Number, err.Description, "檔案名稱： " & strTranFileName)
    cmdImportT21.Enabled = True: Screen.MousePointer = 0: dgMainT21.Enabled = True

End Sub


Private Sub cmdImportT22_Click()
Dim bl_Error As Boolean '記錄有小數點的旗標
Dim str_Error As String '記錄有小數點錯誤的資料
Dim x As Integer '記錄按下確定或是取消，決定是否更新緊急訂單urgent_mark
bl_Error = False: str_Error = ""
Dim str_Storerkey As String

str_Storerkey = "LYFY09"    '貨主

If rsMainT22 Is Nothing Then Exit Sub
If rsMainT22.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT22.Enabled = False: cmdImportT22.Enabled = False
strTranFileName = filLocalFileT22.Path & "\" & filLocalFileT22.FileName

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT22.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then dgMainT22.Enabled = True: cmdImportT22.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT22.RecordCount = 0 Or rsMainT22 Is Nothing Then
Else
rsMainT22.MoveFirst
Do While Not rsMainT22.EOF
    '到貨日期檢查
    If Len(Trim(rsMainT22("預計收退日"))) = 0 Then
    Else
        arrTmp = Split(Trim(rsMainT22("預計收退日")), "/")
        If Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "預計收退日小於今日，訂單轉入終止!", 16, Me.Caption: dgMainT22.Enabled = True: cmdImportT22.Enabled = True: Exit Sub
    End If
    
    '數量檢查
    If Val(rsMainT22("數量")) < 1 Then
        MsgBox "數量小於1，" & Trim(rsMainT22("退貨單號")) & "-品號：" & Trim(rsMainT22("品號")) & "，訂單轉入終止!!請確認!!", , "訂單檔匯入": dgMainT22.Enabled = True: cmdImportT22.Enabled = True: Exit Sub
        Exit Sub
    End If
    
    '資料檢驗 --判斷SKU是否存在
        str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT22("品號")) & "' and Storerkey = '" & str_Storerkey & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '按鈕那些要改
            MsgBox "訂單發現新品號 (" & Trim(rsMainT22("品號")) & ")，請先於商品主檔新建商品資料，訂單轉入終止!!": cmdImportT22.Enabled = True: Screen.MousePointer = 0
            dgMainT22.Enabled = True: cmdImportT22.Enabled = True
            Exit Sub
        End If
        
       '資料檢驗 --判斷consigneekey是否存在
        str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT22("單店代號")) & "' and Storerkey = '" & str_Storerkey & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '按鈕那些要改
            MsgBox "訂單發現新客戶編號 ( " & Trim(rsMainT22("單店代號")) & " )，請先於客戶主檔新建客戶資料，訂單轉入終止!!": cmdImportT22.Enabled = True: Screen.MousePointer = 0
            dgMainT22.Enabled = True: cmdImportT22.Enabled = True
            Exit Sub
        End If
        
    '檢查數量有無小數點
        If InStr(Trim(rsMainT22("數量")), ".") <> 0 Then
            str_Error = "訂單號碼:" & Trim(rsMainT22("退貨單號")) & "，品號:" & Trim(rsMainT22("品號")) & Chr(13) & str_Error
            bl_Error = True
        End If
    rsMainT22.MoveNext
Loop
rsMainT22.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "資料有小數點！請重新匯入"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT22.Enabled = True: cmdImportT22.Enabled = True
                Exit Sub
End If

Tran_Level = cn.BeginTrans: cmdImportT22.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String
Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String

'開始匯入
If rsMainT22 Is Nothing Then GoTo next18
If rsMainT22.RecordCount = 0 Then GoTo next18

'取最後客戶編號
'Call Confirm_Recordset_Closed(tmp_Rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LAPP01' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
'tmp_Rs.Open str_SQL, cn
'
'If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))
'
'tmp_Rs.Close


Do While Not rsMainT22.EOF
    DoEvents: DoEvents
    
'    資料檢驗--判斷是否屬佰事達訂單-->跳下一筆
'    If Trim(rsMainT22("品號")) = "60400119" Then intNotBest = intNotBest + 1: GoTo nextRow17
                  
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT22("退貨單號"))) Then
        strOrderNo = UCase(Trim(rsMainT22("退貨單號")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '已經檢查過，則不比對，直接抓出客戶主檔中的 客戶編號，郵遞區號，連絡人，電話，地址
        str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & str_Storerkey & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT22("單店代號"))) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
'
'        If tmp_Rs.EOF Then
'            '無此客戶名稱則新增
'            strConsigneeKey = myExCharFilter(Trim(rsMainT22("SINGLE_SHOP_CODE")))
'
'            '新增客戶主檔
'            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
'            " values('" & Str_Storerkey & "','" & myExCharFilter(Trim(rsMainT22("POSTAL_CODE"))) & "','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT22("CUSTOMER_NAME"))) & "','" & myExCharFilter(Trim(rsMainT22("SUPPLIER_CODE"))) & "','" & myExCharFilter(Trim(rsMainT22("CONTACT"))) & "','" & myExCharFilter(Trim(rsMainT22("TELEPHONE"))) & "','" & myExCharFilter(Trim(rsMainT22("REQUEST_ADDRESS"))) & "','','" & User_id & "','" & User_id & "') ", RowsAffect, adExecuteNoRecords
'            '紀錄新增之客戶編號
'            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'        Else
''            '比對聯絡人、電話與到貨地址是否相符
'            Call Confirm_Recordset_Closed(rsTmp)
'            str_SQL = "select * from trp01m " & _
'                        "where storerkey = '" & Str_Storerkey & "' and full_name = '" & myExCharFilter(Trim(rsMainT22("CUSTOMER_NAME"))) & "' " & _
'                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT22("REQUEST_ADDRESS"))) & "' "
'            rsTmp.CursorLocation = 3
'            rsTmp.Open str_SQL, cn
                    

'                If rsTmp.EOF Then
'                    '聯絡人、電話與到貨地址不符
'                    strConsigneeKey = myExCharFilter(Trim(rsMainT22("SINGLE_SHOP_CODE")))
'
'                    '新增客戶主檔
'                    cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,updatesource,addwho,editwho) " & _
'                    " values('" & Str_Storerkey & "','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT22("SINGLE_SHOP_CODE"))) & "','" & myExCharFilter(Trim(rsMainT22("CUSTOMER_NAME"))) & "','" & myExCharFilter(Trim(rsMainT22("SUPPLIER_CODE"))) & "','" & myExCharFilter(Trim(rsMainT22("CONTACT"))) & "','" & myExCharFilter(Trim(rsMainT22("TELEPHONE"))) & "','" & myExCharFilter(Trim(rsMainT22("REQUEST_ADDRESS"))) & "','','" & User_id & "','" & User_id & "') ", RowsAffect, adExecuteNoRecords
'                    '紀錄新增之客戶編號
'                    strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
'                Else
'
'                    '相符沿用舊客編
'                    strConsigneeKey = myExCharFilter(Trim(rsMainT22("POSTAL_CODE")))
'                    blCustomerMatch = True
''
'                End If
''            rsTmp.Close
'        End If

        '相符沿用舊客編
        
        strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
        strZip = myExCharFilter(Trim(tmp_Rs("zip")))
        strContact = myExCharFilter(Trim(tmp_Rs("contact")))
        strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
        strAddress = myExCharFilter(Trim(tmp_Rs("address")))
        strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
        blCustomerMatch = True
        tmp_Rs.Close

        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT22("退貨單號"))) & "' and storerkey = '" & str_Storerkey & "' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
            strFacility = "佰事達北倉"
'            If Right(UCase(Trim(rsMainT22("DC代號"))), 1) = "C" Then strFacility = "佰事達中倉"
'            If Right(UCase(Trim(rsMainT22("DC代號"))), 1) = "S" Then strFacility = "佰事達南倉"

            
            arrTmp = Split(Trim(rsMainT22("預計收退日")), "/")
            If Len(Trim(rsMainT22("預計收退日"))) = 0 Then
                cn.RollbackTrans: Tran_Level = 0
                msg_text = "訂單號碼:" & Trim(rsMainT22("退貨單號")) & "，品號:" & Trim(rsMainT22("品號")) & Chr(13) & "資料沒有預計收退日！請向廠商確認"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT22.Enabled = True: cmdImportT22.Enabled = True
                Exit Sub
            Else
                strDate = Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2)
            End If
            
'            arrTmp = Split(Trim(rsMainT22("ORDERED_DATE")), "/")
'            strOrderDate = Val(arrTmp(0)) & "/" & arrTmp(1) & "/" & arrTmp(2)
            
            strOrderDate = Format(Now, "YYYY/MM/DD")
            Dim intPointer As Integer
            intPointer = 1
            'updatesource
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT22("退貨單號"))) & "','R','" & str_Storerkey & "','" & strOrderDate & "','" & strDate & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT22("備註"))) & "','" & filLocalFileT22.FileName & "','','" & User_id & "','" & User_id & "','','" & myExCharFilter(Trim(rsMainT22("退貨單號"))) & "') "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

            
            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & str_Storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT22("退貨單號")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
        
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            
'            lngCasecnt = 1
            
'            '單位換算
'            If Left(myExCharFilter(Trim(rsMainT22("單位"))), 1) = "箱" Then
'
                '取箱包轉換率
'                str_SQL = "select p.casecnt, p.innerpack from " & strWMSDB & "..sku s join " & strWMSDB & "..pack p (nolock) on s.packkey = p.packkey where s.storerkey = 'LAPP01' and s.sku = '" & myExCharFilter(Trim(rsMainT22("商品編碼"))) & "'"
'
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'                lngCasecnt = tmp_Rs("casecnt")      '大單位入數
'                lngInnerpack = tmp_Rs("innerpack")  '中單位入數
'                tmp_Rs.Close
'
'            End If
            
            
'            If UCase(rsMainT22("SAP訂貨單位")) = "BDL" Then intQTY = Val(rsMainT22("商品訂購數量")) * lngInnerpack  '中單位入數，找出pack資料表
'            If UCase(rsMainT22("SAP訂貨單位")) = "KAR" Then intQTY = Val(rsMainT22("商品訂購數量")) * lngCasecnt  '大單位入數，

            
            intQTY = Val(rsMainT22("數量"))
            
            strLot06 = "R01"
            
'            '取trp19m倉別對照表
'            str_SQL = "select bestlot06 from trp19m where storerkey = 'LAPP01' and storerlot06 = '" & rsMainT22("DC代號") & "'"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'            If tmp_Rs.EOF Then
'            '沒有找到對應的倉別
'            '沿用訂單上的倉別
'                strLot06 = rsMainT22("DC代號") '待確認
'            Else
'            '有找到對應的倉別
'                strLot06 = Trim(tmp_Rs("bestlot06"))
'            End If
'
'            tmp_Rs.Close
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT22("退貨單號"))) & "','" & myExCharFilter(Trim(rsMainT22("品號"))) & "','" & str_Storerkey & "'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT22.MoveNext
Loop

next18:

'執行gs_ordersupdate   用客戶主檔更新訂單資料

cn.Execute "exec gs_ordersupdate 'LYFY09'", RowsAffect, adExecuteNoRecords

'檢查有無異常訂單  倉別庫別錯誤 es_Checklot06_by_storer '貨主','訂單檔名'
str_SQL = "exec es_Checklot06_by_storer '" & str_Storerkey & "','" & filLocalFileT22.FileName & "'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If Not tmp_Rs.EOF Then  '有回傳錯誤的訂單資料，產生excel
    Recordset2Excel "配送倉別與明細倉別不符的訂單資料", tmp_Rs
End If

tmp_Rs.Close

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "TMS單號： " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "檔案 " & filLocalFileT22.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT22.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 品號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'備份至FTP
If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT22.FileName

'備份檔案
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT22.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT22.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT22.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT22.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT22.Refresh:
Screen.MousePointer = 0: cmdImportT22.Enabled = True: dgMainT22.Enabled = True
Exit Sub

err_Handle:
   Call ErrorMsgbox(App.title, err.Number, err.Description, "檔案名稱： " & strTranFileName)
    cmdImportT22.Enabled = True: Screen.MousePointer = 0: dgMainT22.Enabled = True
End Sub

Private Sub cmdOpenFile_Click()

'Dim strFullFileName As String, strFileName As String, ExecuteDOSCommand
'
'cmdOpenFile.Enabled = False
'strFullFileName = filLocalFile.Path & "\" & filLocalFile.FileName
'If Len(Trim(filLocalFile.FileName)) = 0 Then Exit Sub
'strFileName = filLocalFile.FileName
'If UCase(Left(strFullFileName, 1)) <> "T" Then cmdOpenFile.Enabled = True: MsgBox "請由T:磁碟機匯入!", 64, Me.Caption: Exit Sub
'
'If strFullFileName = "" Then Exit Sub
'
'On Error GoTo err_Handle
'
''複製檔案
'If Dir("C:\LTKK01\Document", vbDirectory) = "" Then MkDirs "C:\LTKK01\Document"
'FileCopy strFullFileName, "C:\LTKK01\Document\" & strFileName
'
''備份至FTP
'If Dir("O:\Kirin\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\Kirin\OrdersBackup"
'FileCopy strFullFileName, "O:\Kirin\OrdersBackup\" & strFileName
'
'MsgBox "檔案(" & strFileName & ")" & vbCrLf & "已複製到：" & vbCrLf & "C:\LTKK01\Document\" & vbCrLf & "O:\LTKK01\OrdersBackup\", 64, Me.Caption
'
''資料庫記錄
'str_SQL = "insert into gt_filelog(storerkey,filename,filedate,filelen,addwho) values('LTKK01','" & strFileName & "','" & Format(FileDateTime(strFullFileName), "YYYYMMDD hh:mm:ss") & "','" & FileLen(strFullFileName) & "','" & User_id & "')"
'cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'Dim strFrom As String, strTo As String, strCC As String, strBCC As String, strSubject As String, strTextbody As String, strAlways As String, strEmailID As String, strEmailPW As String, strAddAttachment As String
''接取檔資料
'Call ReDim_Recordset(tmp_Rs)
'Call Confirm_Recordset_Closed(tmp_Rs)
'
'str_SQL = "select 檔案時間 = filename , 檔案時間 = filedate, 取檔時間 = gettime, 差異時間 = convert(char(20),gettime - filedate,20) from gv_FileTime where filename = '" & strFileName & "' "
'tmp_Rs.Open str_SQL, cn
'
'If Not tmp_Rs.EOF Then
'    strTextbody = "系統取檔：" & strFileName & "-檔案時間：" & tmp_Rs("檔案時間") & " 檔案大小：" & FileLen(strFullFileName) & " 時間差：" & ((Mid(tmp_Rs("差異時間"), 9, 2) - 1) * 24) + Mid(tmp_Rs("差異時間"), 12, 2) & Mid(tmp_Rs("差異時間"), 14, 6)
'Else
'    strTextbody = "系統取檔：" & strFileName & "-檔案時間：無 檔案大小：" & FileLen(strFullFileName) & " 時間差：無"
'End If
'
'tmp_Rs.Close
'
''LTKK01接取檔自動 Mail 通知
''直接指定
'strFrom = "Tkedi@bestlog.com.tw"
'strTo = "jack@mail.kirin.com.tw,irene@mail.kirin.com.tw;ken@mail.kirin.com.tw;shiu@mail.kirin.com.tw;celine@mail.kirin.com.tw;simon@mail.kirin.com.tw"
'strCC = "Tkedi@bestlog.com.tw"
''strTo = "eric_huang@bestlog.com.tw"
''strCC = ""
'strBCC = strBCC
'strSubject = "取檔通知(" & strFileName & ")"
'strTextbody = strTextbody
'strEmailID = "tkedi"
'strEmailPW = "tkedibl01"
'strAlways = "NO"
'
''傳送郵件
'Dim objEmail As Object
'Set objEmail = CreateObject("CDO.Message")
'
'objEmail.From = strFrom
'objEmail.To = strTo
'objEmail.CC = strCC   ' 副本
'objEmail.BCC = strBCC ' 密件副本
'objEmail.Subject = RTrim(strSubject)
'objEmail.TextBody = strTextbody
'objEmail.AddAttachment strAddAttachment
'
'objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "bestlog.com.tw"
'objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
''SMTP 伺服器需要驗證時
'If Len(RTrim(strEmailID)) > 0 Then
'    objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'    objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendusername") = strEmailID
'    objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strEmailPW
'End If
'objEmail.Configuration.Fields.Update
'objEmail.Send
'
'Set objEmail = Nothing
'
'MsgBox "取檔通知Email完成", 64, strFileName
'
'ExecuteDOSCommand = Shell("cmd /c start C:\LTKK01\Document\" & strFileName, 0)
'
'
''刪除來源檔案
'Kill strFullFileName
'filLocalFile.Refresh
'cmdOpenFile.Enabled = True

Exit Sub

err_Handle:
cmdOpenFile.Enabled = True
Call ErrorMsgbox(App.title, err.Number, err.Description, "檔案名稱： " & strTranFileName)

End Sub

Private Sub CmbStartT17_Click()
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
Dim str As String, strFieldName As String, strFilePath As String, str_TmpSQL As String, S As String, Str_Sku As String
Dim bl_Check As Boolean '檢查匯入的資料~有無出現在總表中~沒有就stop
bl_Check = True
S = "": Str_Sku = ""
'S記錄上一筆消貨單號,Str_sku記錄上一筆品號,如果空白則帶上一筆

On Error GoTo err_Handle
SSTab3.Tab = 1: SSTab3.Enabled = False: CmbStartT17.Enabled = False: cmdImportT17.Enabled = False

Call DB_Connect_Self(cn_string) '建立新連線

'確認路徑是否帶"\"
If Right(filLocalFileT17.Path, 1) = "\" Then
    strFilePath = filLocalFileT17.Path
Else
    strFilePath = filLocalFileT17.Path & "\"
End If

'建立欄位名稱陣列
strFieldName = "交貨單號" & Chr(9) & "出貨備註" & Chr(9)

If Right(filLocalFileT17.Path, 1) <> "\" Then
    strFilePath = filLocalFileT17.Path & "\"
Else
    strFilePath = filLocalFileT17.Path
End If

Set rsMainT17_1 = New ADODB.Recordset

Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFilePath & filLocalFileT17.FileName)   '打開路徑
    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = "RawHeader" Then .Sheets(i).Select: Exit For '選定工作表
    Next
    
    If (.ActiveSheet.Name) <> "RawHeader" Then MsgBox "找不到RawHeader工作表!!", 16, "開啟檔案中止": GoTo endsub
    
    'k = 1 '預設由第一列開始匯入
    
    '若無來源欄位名稱
    If strFieldName = "" Then
        '取欄位名稱
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '由第二列開始匯入
    End If
    
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "交貨單號" Then k = i: Exit For
    Next i
    
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    'Dim rsMainT15 As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rsMainT17_1 = Nothing: GoTo endsub
    
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "RawHeader工作表，第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsMainT17_1.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT17_1.CursorType = adOpenKeyset
    rsMainT17_1.LockType = adLockOptimistic
    rsMainT17_1.Open
    
    '寫入Recordset  '從這邊開始往下寫
    Do While Len(RTrim(.Cells(k + 1, 1))) > 0
    rsMainT17_1.AddNew
        For j = 1 To UBound(arrTmp)
            rsMainT17_1(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
        Next j
    rsMainT17_1.Update
    k = k + 1
    Loop
    
    If rsMainT17_1.RecordCount > 0 Then rsMainT17_1.MoveFirst

End With


Set dgMainT17_1.DataSource = rsMainT17_1

'If rsMainT17_1 Is Nothing Then
'
'    MsgBox "查無資料!", 64, "Excel2Recordset"
'
'Else
'    SetDataGridColWidth Me.Caption, rsMainT17_1 '設定欄寬
'End If

'如果沒有出貨總表則跳離
If rsMainT17_1 Is Nothing Then MsgBox "找不到RawHeader工作表!!", 16, "開啟檔案中止": SSTab3.Enabled = True: CmbStartT17.Enabled = True: Exit Sub
If rsMainT17_1.EOF Then MsgBox "找不到RawHeader工作表!!", 16, "開啟檔案中止": SSTab3.Enabled = True: CmbStartT17.Enabled = True: Exit Sub

'/////////////////////////////////////////////////////////////////////////////////匯入Format////////////////////////////////////////////////////////////////////////////////
SSTab3.Tab = 0
'建立欄位名稱陣列
strFieldName = "DC代號" & Chr(9) & "客戶代號" & Chr(9) & "EXE文件編號" & Chr(9) & "SAP DN NO." & Chr(9) & "單據類別" & Chr(9) & "單據產生日" & Chr(9) & "預計出貨日" & Chr(9) & "項次" & Chr(9) & "商品編碼" & Chr(9) & "商品訂購數量" & Chr(9) & "銷售別" & Chr(9) & "商品最小數量" & Chr(9) & "客戶進價" & Chr(9) & _
              "折讓進額" & Chr(9) & "實際出貨數量" & Chr(9) & "實際出貨倉別" & Chr(9) & "稅別" & Chr(9) & "批次" & Chr(9) & "實際檢貨日期" & Chr(9) & "貨主" & Chr(9) & "送貨地址" & Chr(9) & "銷售組織" & Chr(9) & "營業所" & Chr(9) & "業務組長" & Chr(9) & "王安單號" & Chr(9) & "SAP訂貨單位" & Chr(9) & "原因" & Chr(9) & "客戶名稱" & Chr(9) & "備注" & Chr(9) & "客戶通路別" & Chr(9)

Set rsMainT17 = New ADODB.Recordset

'Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
'    .Workbooks.Open (strFilePath & filLocalFileT16.FileName)   '打開路徑
    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = "Format" Then .Sheets(i).Select: Exit For '選定工作表
    Next
    
    If (.ActiveSheet.Name) <> "Format" Then MsgBox "找不到Format工作表!!", 16, "開啟檔案中止": GoTo endsub

    'k = 1 '預設由第一列開始匯入
    
    '若無來源欄位名稱
    If strFieldName = "" Then
        '取欄位名稱
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        'k = 2 '由第二列開始匯入
    End If
    
    For i = 1 To 255
            If Trim(.Cells(i, 1)) = "DC代號" Then k = i: Exit For
    Next i
    
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    If UBound(arrTmp) < 1 Then Set rsMainT17 = Nothing: GoTo endsub
    
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "Format工作表，第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsMainT17.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsMainT17.CursorType = adOpenKeyset
    rsMainT17.LockType = adLockOptimistic
    rsMainT17.Open
    
    rsMainT17_1.MoveFirst
    '寫入Recordset  '從這邊開始往下寫
    Do While Len(RTrim(.Cells(k + 1, 1))) > 0   '挑準數量為空值則停止
'    If RTrim(.Cells(k + 1, 6)) = "60400119" Then '排除運費
'    Else
        rsMainT17.AddNew
            For j = 1 To UBound(arrTmp)
                If j = UBound(arrTmp) Then    '備注欄位
                    bl_Check = True
                    rsMainT17_1.MoveFirst
                    Do While Not rsMainT17_1.EOF
                        If Trim(rsMainT17_1("交貨單號").Value) = Trim(rsMainT17("SAP DN NO.").Value) Then rsMainT17("備注").Value = Trim(rsMainT17_1("出貨備註").Value): bl_Check = False: Exit Do
                        rsMainT17_1.MoveNext
                    Loop
                    rsMainT17(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
                    If bl_Check = True Then MsgBox "RawHeader查無訂單號碼:" & Trim(rsMainT17("SAP DN NO.").Value) & "的出貨備註資料!", 64, "Format匯入中止": SSTab3.Enabled = True: CmbStartT17.Enabled = True: GoTo endsub
                Else
                    rsMainT17(j - 1) = RTrim(myExCharFilter(.Cells(k + 1, j)))
                End If
            Next j
        rsMainT17.Update
'    End If
    k = k + 1
    
    Loop
    
    If rsMainT17.RecordCount > 0 Then rsMainT17.MoveFirst

End With
    
Set dgMainT17.DataSource = rsMainT17

rsMainT17.Sort = "[SAP DN NO.]"

If rsMainT17 Is Nothing Then
    MsgBox "Format查無資料!", 64, "Excel2Recordset"
Else
    SetDataGridColWidth Me.Caption, dgMainT17
    MsgBox "此訂單檔共匯入:" & Chr(13) & "Format:" & rsMainT17.RecordCount & "筆明細" & Chr(13) & "" & _
                                          "RawHeader:" & rsMainT17_1.RecordCount & "筆明細" & Chr(13) & "" & _
                                          "請確認筆數是否正確!", 64, "金盛世訂單開啟"
    cmdImportT17.Enabled = True
End If

'如果有出貨總表，其他三個工作表沒有資料則提示，但不擋
'If rsMainT16_1.RecordCount = 0 And rsMainT16_2.RecordCount = 0 And rsMainT16_3.RecordCount = 0 Then MsgBox "此訂單無細項資料，請確認此訂單是否正確!", vbCritical, "馬玉山訂單開啟"

endsub:
SSTab3.Enabled = True: CmbStartT17.Enabled = True:
MyXlsApp.Quit: Set MyXlsApp = Nothing

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
SSTab3.Enabled = True: CmbStartT17.Enabled = True
End Sub


Private Sub cmdOpenFilesT15_Click()
If Trim(cboSheetT15) = "" Then Exit Sub
On Error GoTo err_Handle
dgMainT15.Enabled = False: cmdImportT15.Enabled = False

Dim str As String, strFieldName As String, strFilePath As String, strSheetName As String, str_storekey As String, Str_Sku As String
'確認路徑是否帶"\"
If Right(filLocalFileT15.Path, 1) = "\" Then
    strFilePath = filLocalFileT15.Path
Else
    strFilePath = filLocalFileT15.Path & "\"
End If
'建立欄位名稱陣列
strFieldName = ""
If Right(filLocalFileT15.Path, 1) <> "\" Then
    strFilePath = filLocalFileT15.Path & "\"
Else
    strFilePath = filLocalFileT15.Path
End If
Set rsMainT15 = New ADODB.Recordset
strSheetName = cboSheetT15
Call Excel2Recordset(strFilePath & filLocalFileT15.FileName, strSheetName, strFieldName, rsMainT15)
rsMainT15.MoveFirst

Set dgMainT15.DataSource = rsMainT15

If rsMainT15 Is Nothing Then
    MsgBox "查無資料!", 64, "Excel2Recordset"
Else
    'SetDataGridColWidth Me.Caption, dgMainT15
    MsgBox "此工作表共 " & rsMainT15.RecordCount & "筆資料，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    cmdImportT15.Enabled = True
End If
rsMainT15.Sort = "出貨單號"
rsMainT15.MoveFirst
dgMainT15.Enabled = True: cmdImportT15.Enabled = True

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdOpenFilesT20_Click()
If Trim(cboSheetT20) = "" Then Exit Sub
On Error GoTo err_Handle
dgMainT20.Enabled = False: cmdImportT20.Enabled = False

Dim str As String, strFieldName As String, strFilePath As String, strSheetName As String, str_storekey As String, Str_Sku As String
'確認路徑是否帶"\"
If Right(filLocalFileT20.Path, 1) = "\" Then
    strFilePath = filLocalFileT20.Path
Else
    strFilePath = filLocalFileT20.Path & "\"
End If
'建立欄位名稱陣列
strFieldName = ""
If Right(filLocalFileT20.Path, 1) <> "\" Then
    strFilePath = filLocalFileT20.Path & "\"
Else
    strFilePath = filLocalFileT20.Path
End If
Set rsMainT20 = New ADODB.Recordset
strSheetName = cboSheetT20
Call Excel2Recordset(strFilePath & filLocalFileT20.FileName, strSheetName, strFieldName, rsMainT20)
rsMainT20.MoveFirst

Set dgMainT20.DataSource = rsMainT20

If rsMainT20 Is Nothing Then
    MsgBox "查無資料!", 64, "Excel2Recordset"
Else
    'SetDataGridColWidth Me.Caption, dgMainT20
    MsgBox "此工作表共 " & rsMainT20.RecordCount & "筆資料，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    cmdImportT20.Enabled = True
End If

rsMainT20.Sort = "訂單號碼"
rsMainT20.MoveFirst
dgMainT20.Enabled = True: cmdImportT20.Enabled = True

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdOpenFilesT21_Click()
If Trim(cboSheetT21) = "" Then Exit Sub
On Error GoTo err_Handle
dgMainT21.Enabled = False: cmdImportT21.Enabled = False

Dim str As String, strFieldName As String, strFilePath As String, strSheetName As String, str_storekey As String, Str_Sku As String
'確認路徑是否帶"\"
If Right(filLocalFileT21.Path, 1) = "\" Then
    strFilePath = filLocalFileT21.Path
Else
    strFilePath = filLocalFileT21.Path & "\"
End If
'建立欄位名稱陣列
strFieldName = ""
If Right(filLocalFileT21.Path, 1) <> "\" Then
    strFilePath = filLocalFileT21.Path & "\"
Else
    strFilePath = filLocalFileT21.Path
End If
Set rsMainT21 = New ADODB.Recordset
strSheetName = cboSheetT21
Call Excel2Recordset(strFilePath & filLocalFileT21.FileName, strSheetName, strFieldName, rsMainT21)
rsMainT21.MoveFirst

Set dgMainT21.DataSource = rsMainT21

If rsMainT21 Is Nothing Then
    MsgBox "查無資料!", 64, "Excel2Recordset"
Else
    'SetDataGridColWidth Me.Caption, dgMainT21
    MsgBox "此工作表共 " & rsMainT21.RecordCount & "筆資料，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    cmdImportT21.Enabled = True
End If
rsMainT21.Sort = "調撥單號"
rsMainT21.MoveFirst

dgMainT21.Enabled = True: cmdImportT21.Enabled = True

Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

'Private Sub cmdUpload3_Click()
'On Error GoTo err_Handle
'
'If cmdLogOff3.Enabled = False Then MsgBox "請先登入伺服器！", 64, Me.Caption: Exit Sub
'
'cmdUpload3.Enabled = False
'
'If int3Ready(True) = True Then
'
'    int3.Execute , "Put " & Chr(34) & filLocal3.Path & "\" & filLocal3.FileName & Chr(34) & " ""XRSLUPL.TXT"""
'    lblStatus3 = "上傳中請稍後...."
'End If
'
'    Do While int3.StillExecuting
'        DoEvents: DoEvents: DoEvents
'    Loop
'
'    lstRemoteFile3.Clear
'    int3.Execute , "DIR"
'
'    Do While int3.StillExecuting
'    DoEvents: DoEvents: DoEvents
'    Loop
'
'    '上傳確認
'    For i = 0 To lstRemoteFile3.ListCount - 1
'        If lstRemoteFile3.List(i) = "XRSLUPL.TXT" Then
'            '上傳完成刪除本機上之檔案
'            Kill filLocal3.Path & "\" & "XRSLUPL.TXT"
'            lblStatus3 = "檔案上傳完成！"
'            filLocal3.Refresh
'            GoTo Step1
'        End If
'    Next i
'
'    lblStatus3 = "上傳失敗！"
'    MsgBox "檔案上傳失敗，請重新上傳！", 64, "Error"
'Step1:
'
'cmdUpload3.Enabled = True
'Exit Sub
'
'err_Handle:
'    Dim tmpString As String
'    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
'    tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
'    CreateErrorLog Me.Name & "-上傳", Me.Caption, "cmdUpload3_Click", tmpString
'    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'End Sub

Private Sub Command1_Click()
    Dim i As Double
    '暫不使用
    If ITCReady(True) = True Then
        'Check that they are not recieving a folder
        If Right(lstRemoteFile.Text, 1) = "/" Then
            MsgBox lstRemoteFile.Text & " is a folder and cannot be sent.", vbCritical + vbOKOnly, "Error"
            Exit Sub
        End If
        'Check that the file does not already exist on the computer, if it does exit sub
        For i = 0 To filLocalFile.ListCount
            If lstRemoteFile.Text = filLocalFile.List(i) Then
                MsgBox "檔案 " & Right(lstRemoteFile.Text, 18) & " 已存在.", vbInformation + vbOKOnly, "Recieve"
                Exit Sub
            End If
        Next i
        str_file = Trim(Right(lstRemoteFile.Text, 18))
        ITC.Execute , "GET " & Chr(34) & str_file & Chr(34) & " " & Chr(34) & filLocalFile.Path & "\" & str_file & Chr(34)
        lblStatus = "下載中請稍後...."
        Do Until ITCReady(False)
            DoEvents: DoEvents: DoEvents: DoEvents
        Loop
        filLocalFile.Refresh
        lblStatus = "已連線"
    End If
    
End Sub

Private Sub cmd_import_Click()
    'TK訂單一筆訂單2筆明細的備註如果不同，會分成兩張訂單進行匯入
    
    Dim strTranFileName As String, strFileName As String, strRePoOrderkey As String
    '開始匯入檔案
    
    strTranFileName = filLocalFile.Path & "\" & filLocalFile.FileName
    If Len(Trim(filLocalFile.FileName)) = 0 Then Exit Sub
    If UCase(Left(strTranFileName, 1)) <> "T" Then MsgBox "請由T磁碟機匯入!", 64, Me.Caption: Exit Sub
    strFileName = filLocalFile.FileName
    If strTranFileName = "" Then Exit Sub
    
    On Error GoTo err_Handle
    If FileLen(strTranFileName) = 0 Then MsgBox "檔案大小 = 0 , 檔名: " & filLocalFile.FileName, vbOKOnly + vbInformation, Me.Caption: Exit Sub
      
    cmd_Import.Enabled = False: Screen.MousePointer = 11: dg_CustInv.Enabled = False
    Set dg_CustInv.DataSource = Nothing
    DoEvents: DoEvents
    
    Dim strRow As String    '讀取每一行文字
    Dim strField As String  '讀取每個區隔的資料
    Dim intPointer As Integer
    Set rs_Src = New Recordset
    
    With rs_Src
        .Fields.Append "訂單號碼", adChar, 30, adFldUpdatable
        .Fields.Append "訂單項次", adChar, 10, adFldUpdatable
        .Fields.Append "交貨單號", adChar, 20, adFldUpdatable
        .Fields.Append "訂單日期", adChar, 8, adFldUpdatable
        .Fields.Append "訂單類別", adChar, 10, adFldUpdatable
        .Fields.Append "地址別", adChar, 30, adFldUpdatable
        .Fields.Append "料號", adChar, 35, adFldUpdatable
        .Fields.Append "中文名稱", adChar, 100, adFldUpdatable
        .Fields.Append "最小單位數量", adDouble, adFldUpdatable
        .Fields.Append "訂單數量", adDouble, adFldUpdatable
        .Fields.Append "訂單單位", adChar, 10, adFldUpdatable
        .Fields.Append "單價", adDouble, adFldUpdatable
        .Fields.Append "到貨日期", adChar, 8, adFldUpdatable
        .Fields.Append "客戶單號", adChar, 60, adFldUpdatable
        .Fields.Append "備註", adChar, 255, adFldUpdatable
        .Fields.Append "貼標", adChar, 10, adFldUpdatable
        .Fields.Append "倉別", adChar, 18, adFldUpdatable
        .Fields.Append "儲位", adChar, 18, adFldUpdatable
        .CursorType = adOpenKeyset
        .LockType = adLockOptimistic
        .Open
        
        Dim arrTmp, arrTmp1, strLineTmp As String, intArrline As Integer, i As Double
        
        '開啟檔案
        Open strTranFileName For Input As #1
        Do While Not EOF(1)
        '匯入檔案
            Line Input #1, strLineTmp '取資料行
            arrTmp = Split(strLineTmp, Chr(10)) '先切筆數
            If UBound(arrTmp) = -1 Then GoTo NextStep
            If UBound(arrTmp) > 0 Then
                For intArrline = 0 To UBound(arrTmp) - 1
                    '切筆數再切欄位
                    arrTmp1 = Split(arrTmp(intArrline), ",")
                    .AddNew
                        For i = 0 To .Fields.Count - 1
                            .Fields(i) = Trim(arrTmp1(i))
                        Next i
                    .Update
                    
                Next intArrline
            Else '直接切欄位
                    arrTmp1 = Split(arrTmp(intArrline), ",") '切欄位
                    .AddNew
                        For i = 0 To .Fields.Count - 1
                            If i = 15 Then
                                .Fields(i) = GetWord(Trim(arrTmp1(i)), 1, 10) & "" '貼標od.updatesource取10碼
                            Else
                                .Fields(i) = Trim(arrTmp1(i)) & ""
                            End If
                        Next i
                    .Update
            End If
NextStep:
        Loop
        
        Close #1
        .MoveFirst
        .Sort = "訂單號碼,地址別,到貨日期,訂單類別,備註,訂單項次"
End With
 
 '資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select updatesource from orders where storerkey = 'LTKK01' and rtrim(updatesource)='" & filLocalFile.FileName & "' "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True: Exit Sub

'檢查訂單資料是否正確
rs_Src.MoveFirst
Do While Not rs_Src.EOF
    '到貨日期檢查
    If Trim(rs_Src("到貨日期")) < Format(Now, "YYYYMMDD") Then: MsgBox "到貨日小於今日，訂單轉入終止!", 16, Me.Caption: cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True: Exit Sub
    
    '貨主貨號 edit by eric 20140923@檢查有無品號，順便更新之後要用的品號。
    str_SQL = "select sku from gv_skuxpack where storerkey = 'LTKK01' and (storersku = '" & Trim(rs_Src("料號")) & "' or sku = '" & Trim(rs_Src("料號")) & "') "
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
        MsgBox "訂單發現新料號 (" & Trim(rs_Src("料號")) & " ) " & Trim(rs_Src("中文名稱")) & "，訂單轉入終止!!"
        tmp_Rs.Close
        Exit Sub
    End If
    'if料號長度>20則使用資料庫的SKU，否則直接使用訂單上的料號
    If Len(Trim(rs_Src("料號"))) > 20 Then rs_Src("料號") = tmp_Rs("sku")
    
'    '檢查客戶編號
'    str_SQL = "select consigneekey from trp01m where storerkey = 'LTKK01' and len(rtrim(consigneekey))>5 and substring(rtrim(consigneekey),5,20) = '" & Trim(rs_Src("地址別")) & "'"
'
'    Call Confirm_Recordset_Closed(rsMainTK)
'    rsMainTK.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If rsMainTK.EOF Then
'        cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
'        MsgBox "訂單發現新客戶單號:" & Trim(rs_Src("地址別")) & Chr(13) & "請先連絡客戶，於系統建立第五碼開始為:" & Trim(rs_Src("地址別")) & "的客戶單號" & Chr(13) & "EX: XXXX" & Trim(rs_Src("地址別")) & Chr(13) & "請連絡客戶，新增客戶主檔資料!!", vbCritical, "訂單轉入終止!!"
'        Exit Sub
'    End If
    
    '檢查單別
    If Trim(rs_Src("訂單類別")) = "C" Then
        MsgBox "發現不明訂單類別，訂單訂單號碼:" & Trim(rs_Src("訂單號碼")) & " 訂單類別:" & Trim(rs_Src("訂單類別")) & "，訂單轉入終止!!請確認!!"
        cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
        Exit Sub
    End If
    
    '檢查TK倉別 Create by Gemini @20080521
    If (UCase(Trim(rs_Src("倉別"))) = "BL01" Or UCase(Trim(rs_Src("倉別"))) = "BLR68" Or UCase(Trim(rs_Src("倉別"))) = "BL02") = False Then
        MsgBox "訂單檔案：" & filLocalFile.FileName & " (TK單號：" & Trim(rs_Src("交貨單號")) & ")。" & vbCrLf & "請通知客戶，TK倉庫別不符，請確認該筆訂單是否有誤!?", vbOKOnly, "發現TK倉別非佰事達倉別 BL01, BLR68 之訂單明細!"
        cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
        Exit Sub
    End If
    
    '資料檢驗 --判斷客戶PO訂單號碼是否重複, 重複時入系統並紀錄 長度>0才檢查 edit by eric20140923
    If Len(StrNoCH(Trim(rs_Src("客戶單號")))) > 0 Then
         Call Confirm_Recordset_Closed(tmp_Rs)
         str_SQL = "select o.orderkey from orders o(nolock) join orderdetail od(nolock) on o.orderkey = od.orderkey " & _
             "where o.externorderkey = '" & Trim(rs_Src("訂單號碼")) & "' " & _
             "and od.externlineno = '" & Trim(rs_Src("訂單項次")) & Trim(rs_Src("交貨單號")) & "' " & _
             "and rtrim(substring(o.consigneekey,5,20)) = '" & Trim(rs_Src("地址別")) & "' " & _
             "and rtrim(o.b_phone1)='" & StrNoCH(Trim(rs_Src("客戶單號"))) & "' " & _
             "and len(rtrim(isnull(o.b_phone1,''))) > 0 " & _
             "and isnull(o.type,'') <> '刪單' " & _
             "and o.priority = '" & Trim(rs_Src("訂單類別")) & "' and o.storerkey = 'LTKK01' "
    
         tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
         If tmp_Rs.EOF = False Then strRePoOrderkey = strRePoOrderkey & StrNoCH(Trim(rs_Src("客戶單號"))) & "','"
    End If
    rs_Src.MoveNext
Loop

Tran_Level = cn.BeginTrans
Dim int_OrderLine As Integer, int_Order As Integer, int_Repeat As Integer, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strTKLOC As String, strFacility As String, strSku As String, intOrderLinenumber As Integer
rs_Src.MoveFirst
Do While Not rs_Src.EOF
DoEvents: DoEvents
'
'    '貨主貨號
'    str_SQL = "select sku from gv_skuxpack where storerkey = 'LTKK01' and (storersku = '" & Trim(rs_Src("料號")) & "' or sku = '" & Trim(rs_Src("料號")) & "') "
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If tmp_Rs.EOF Then
'        cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
'        MsgBox "訂單發現新料號 (" & Trim(rs_Src("料號")) & " ) " & Trim(rs_Src("中文名稱")) & "，訂單轉入終止!!"
'        Exit Sub
'    End If
    
strSku = Trim(rs_Src("料號"))

'If Len(Trim(rs_Src("料號"))) > 20 Then strSku = tmp_Rs("sku")
        
            '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
            If strOrderNo <> Trim(rs_Src("訂單號碼")) & Trim(rs_Src("地址別")) & Trim(rs_Src("到貨日期")) & Trim(rs_Src("訂單類別")) & Trim(rs_Src("備註")) Then
                    strOrderNo = Trim(rs_Src("訂單號碼")) & Trim(rs_Src("地址別")) & Trim(rs_Src("到貨日期")) & Trim(rs_Src("訂單類別")) & Trim(rs_Src("備註"))
                    
                    '訂單主檔新增一筆
                    str_SQL = "select isnull(max(orderkey),0) from orders"
                    Call Confirm_Recordset_Closed(tmp_Rs)
                    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                    str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
                    tmp_Rs.Close
                    
                    'edit by Eric 20150119新增南倉
                    If Right(UCase(Trim(rs_Src("儲位"))), 2) = "-S" Then
                        strFacility = "佰事達南倉"
                    Else
                        strFacility = "佰事達北倉"
                    End If
                    intOrderLinenumber = 0
                    
                    str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,DeliveryDate,ConsigneeKey,CustomerOrderkey,Notes,UpdateSource,Facility,b_phone1,type,addwho,editwho) " & _
                    "VALUES ('" & str_Orderkey & "','" & Trim(rs_Src("訂單號碼")) & "','" & Trim(rs_Src("訂單類別")) & "','LTKK01','" & Trim(rs_Src("訂單日期")) & "','" & Trim(rs_Src("到貨日期")) & "', " & _
                    "'" & Trim(rs_Src("地址別")) & "','" & Trim(rs_Src("客戶單號")) & "','" & Trim(rs_Src("備註")) & "','" & filLocalFile.FileName & "','" & strFacility & "','" & StrNoCH(Trim(rs_Src("客戶單號"))) & "','','" & User_id & "','" & User_id & "')"
                    
                    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                    int_Order = int_Order + 1
            End If
            
            '資料檢驗--判斷訂單明細是否重複，重複不增加明細，跳下一筆資料，所以放在最前面也無用。
            Call Confirm_Recordset_Closed(tmp_Rs)
            str_SQL = "select o.orderkey from ORDERDETAIL od(nolock) join orders o(nolock) on o.orderkey = od.orderkey where o.storerkey = 'LTKK01' and o.ExternOrderKey='" & Trim(rs_Src("訂單號碼")) & "' and rtrim(o.priority)= '" & Trim(rs_Src("訂單類別")) & "' and isnull(type,'') <> '刪單' and od.ExternlineNO= '" & Trim(rs_Src("訂單項次")) & "_" & Trim(rs_Src("交貨單號")) & "' "
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            If tmp_Rs.EOF Then
                 
                 If Trim(rs_Src("儲位")) = "04A" Or Trim(rs_Src("儲位")) = "03A" Then
                    strTKLOC = "R" & Trim(rs_Src("儲位"))
                 Else
                    strTKLOC = Trim(rs_Src("儲位"))
                 End If
                 
                 intOrderLinenumber = intOrderLinenumber + 1
                                     
                '訂單明細資料新增
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber, ExternlineNO ,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice ,CartonGroup,notes,updatesource)" & _
                "VALUES ('" & str_Orderkey & "','" & Format(intOrderLinenumber, "00000") & "','" & Trim(rs_Src("訂單項次")) & "_" & Trim(rs_Src("交貨單號")) & "','" & Trim(rs_Src("訂單號碼")) & "','" & strSku & "','LTKK01'," & _
                "'" & Trim(rs_Src("最小單位數量")) & "','" & Trim(rs_Src("最小單位數量")) & "','" & strTKLOC & "','" & Trim(rs_Src("倉別")) & "','" & Trim(rs_Src("訂單單位")) & "','" & Trim(rs_Src("單價")) & "','" & Trim(rs_Src("貼標")) & "','" & Trim(rs_Src("備註")) & "','" & IIf(UCase(Trim(rs_Src("貼標"))) = "Y", "蓋章", "") & "') "
                
                cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                int_OrderLine = int_OrderLine + 1
                
            Else
                 int_Repeat = int_Repeat + 1
                Call FTPlog("訂單明細重複" & str_SQL)
                
                '紀錄重複
                strReOrderkey = strReOrderkey & Trim(rs_Src("訂單號碼")) & Trim(rs_Src("訂單項次")) & Trim(rs_Src("交貨單號")) & "','"
'                GoTo Nextstep

            End If
    
'    Else
'        '訂單重複
'        Call FTPlog("訂單重複" & str_SQL)
'        '紀錄重複
'        strReOrderkey = strReOrderkey & RTrim(tmp_rs("externorderkey")) & "','"
'    End If
           
'nextloop:
rs_Src.MoveNext
Loop

'補客戶資料
cn.Execute "exec gs_ordersupdate 'LTKK01' ", RowsAffect, adExecuteNoRecords

'新客戶檢查
str_SQL = "select 貨主=storerkey , 貨主訂單單號=externorderkey , 訂單類別 = priority , 訂單日期=orderdate , 到貨日期=deliverydate , 客戶編號=consigneekey , 檢查日期 = getdate() " & _
        "from orders o (nolock) " & _
        "Where o.storerkey = 'LTKK01' and o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select t1m.consigneekey from trp01m t1m where t1m.storerkey = 'LTKK01') "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel顯示
    Call Recordset2Excel("缺客戶資料", tmp_Rs)
    If Dir("C:\LTKK01\缺客戶資料", vbDirectory) = "" Then MkDirs "C:\LTKK01\缺客戶資料"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\LTKK01\缺客戶資料\缺客戶資料_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True: cn.RollbackTrans: Tran_Level = 0
    MsgBox "發現新客戶，訂單轉入中止!", vbOKOnly, Me.Caption
    
    Exit Sub
End If

cn.CommitTrans: Tran_Level = 0

'顯示轉入資料
Set dg_CustInv.DataSource = rs_Src

'取欄位寬度
SetDataGridColWidth Me.Caption, dg_CustInv

With dg_CustInv
      .Columns(8).Alignment = dbgRight
      .Columns(9).Alignment = dbgRight
      .Columns(11).Alignment = dbgRight
 End With

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，檔案 " & strTranFileName & " 備份於 C:\Orders\LTKK01\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，檔案:" & strTranFileName)
    cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
    
    If int_Repeat > 0 Then
        msg_text = "有" & int_Repeat & " 筆訂單明細重複轉檔!"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    End If

'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then
'
'    str_SQL = "select 重複類別 = 'TK訂單明細重複-不轉入' , 轉入檔案名稱 = '" & filLocalFile.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 料號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
'        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(od.externorderkey)+rtrim(od.orderlinenumber) in ('" & strReOrderkey & "') " & _
'        "Union " & _
'        "select 重複類別 = '客戶訂單號碼重複-已轉入' , 轉入檔案名稱 = '" & filLocalFile.FileName & "' ,訂單號碼 = o.externorderkey ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 料號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource ,檢查時間 = getdate() " & _
'        "From orders o join orderdetail od on o.orderkey = od.orderkey and isnull(o.type,'') <> '刪單' where rtrim(b_phone1) in ('" & strRePoOrderkey & "') and len(rtrim(isnull(b_phone1,''))) > 0 "

    str_SQL = "select 重複類別 = 'TK訂單明細重複-不轉入' , 轉入檔案名稱 = '" & filLocalFile.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = rtrim(replace(od.externlineno,'_','')) , 料號 = isnull(s.storersku,s.sku) , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o(nolock) join orderdetail od(nolock) on o.orderkey = od.orderkey left join gv_skuxpack s(nolock) on s.sku = od.sku and od.storerkey = s.storerkey where rtrim(od.externorderkey)+rtrim(od.orderlinenumber) in ('" & strReOrderkey & "') or rtrim(od.externorderkey)+rtrim(replace(od.externlineno,'_','')) in ('" & strReOrderkey & "') " & _
        "union " & _
        "select 重複類別 = '客戶訂單號碼重複-已轉入' , 轉入檔案名稱 = '" & filLocalFile.FileName & "' ,訂單號碼 = o.externorderkey ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 料號 = isnull(s.storersku,s.sku) , 數量 = od.openqty , 上次檔案名稱 = o.updatesource ,檢查時間 = getdate() " & _
        "From orders o(nolock) join orderdetail od(nolock) on o.orderkey = od.orderkey and isnull(o.type,'') <> '刪單' left join gv_skuxpack s(nolock) on s.sku = od.sku and od.storerkey = s.storerkey where rtrim(b_phone1) in ('" & strRePoOrderkey & "') and len(rtrim(isnull(b_phone1,''))) > 0 "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\LTKK01\訂單重複", vbDirectory) = "" Then MkDirs "C:\LTKK01\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\LTKK01\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'取檔資料庫記錄
str_SQL = "insert into gt_filelog(storerkey,filename,filedate,filelen,addwho) values('LTKK01','" & filLocalFile.FileName & "','" & Format(FileDateTime(strTranFileName), "YYYY/MM/DD hh:mm:ss") & "','" & FileLen(strTranFileName) & "','" & User_id & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

Dim strFrom As String, strTo As String, strCC As String, strBCC As String, strSubject As String, strTextbody As String, strAlways As String, strEmailID As String, strEmailPW As String, strAddAttachment As String

'接取檔資料
Call ReDim_Recordset(tmp_Rs)
Call Confirm_Recordset_Closed(tmp_Rs)

str_SQL = "select 檔案時間 = filename , 檔案時間 = filedate, 取檔時間 = gettime, 差異時間 = convert(char(20),gettime - filedate,20) from gv_FileTime where filename = '" & strFileName & "' "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then
    strTextbody = strFileName & "-檔案時間：" & tmp_Rs("檔案時間") & " 檔案大小：" & FileLen(strTranFileName) & " 時間差：" & ((Mid(tmp_Rs("差異時間"), 9, 2) - 1) * 24) + Mid(tmp_Rs("差異時間"), 12, 2) & Mid(tmp_Rs("差異時間"), 14, 6) & "：(匯入)"
Else
    strTextbody = strFileName & "-檔案時間：無 檔案大小：" & FileLen(strTranFileName) & " 時間差：無" & "：(匯入)"
End If

''LTKK01自動 Mail 通知
''直接指定
''Exit Sub
'strFrom = "Tkedi@bestlog.com.tw"
'strTo = "jack@mail.kirin.com.tw,irene@mail.kirin.com.tw;ken@mail.kirin.com.tw;shiu@mail.kirin.com.tw;celine@mail.kirin.com.tw;simon@mail.kirin.com.tw"
'strCC = "tkedi@bestlog.com.tw"
'strBCC = strBCC
'strSubject = "取檔通知(" & filLocalFile.FileName & ")"
'strTextbody = strTextbody
'strEmailID = "tkedi"
'strEmailPW = "tkedibl01"
'strAlways = "NO"
'
''傳送郵件
'Dim objEmail As Object
'Set objEmail = CreateObject("CDO.Message")
'
'objEmail.From = strFrom
'objEmail.To = strTo
'objEmail.CC = strCC   ' 副本
'objEmail.BCC = strBCC ' 密件副本
'objEmail.Subject = RTrim(strSubject)
'objEmail.TextBody = strTextbody
'objEmail.AddAttachment strAddAttachment
'
'objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "bestlog.com.tw"
'objEmail.Configuration.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
''SMTP 伺服器需要驗證時
'If Len(RTrim(strEmailID)) > 0 Then
'    objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
'    objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendusername") = strEmailID
'    objEmail.Configuration("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strEmailPW
'End If
'objEmail.Configuration.Fields.Update
'objEmail.Send
'
'Set objEmail = Nothing
'
'MsgBox "取檔通知Email完成", 64, strFileName
'
'tmp_Rs.Close

'備份檔案
If Dir("C:\LTKK01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\LTKK01\Orders\Backup"
If strTranFileName <> "C:\LTKK01\Orders\Backup\" & filLocalFile.FileName Then FileCopy strTranFileName, "C:\LTKK01\Orders\Backup\" & filLocalFile.FileName: Kill strTranFileName

'備份至FTP
If Dir("O:\KIRIN\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\KIRIN\OrdersBackup"
FileCopy "C:\LTKK01\Orders\Backup\" & filLocalFile.FileName, "O:\KIRIN\OrdersBackup\" & filLocalFile.FileName

filLocalFile.Refresh
SSTab1.Tab = 1
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-匯入", Me.Caption, "cmd_Import_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmd_Import.Enabled = True: Screen.MousePointer = 0: dg_CustInv.Enabled = True
End Sub

'Private Sub cmdImport3_Click()
'On Error GoTo err_Handle
'
'If FileLen(filLocal3.Path & "\" & filLocal3.FileName) = 0 Then MsgBox "檔案大小 = 0,檔名:" & str_file, vbOKOnly + vbInformation, Me.Caption: Exit Sub
'
'If UCase(filLocal3.FileName) <> "XRSLDNL.TXT" Then
'    ConfirmYN = MsgBox("正匯入檔案名稱非 xrsldnl.txt 的檔案，確定匯入?", vbQuestion + vbYesNo, "Warning")
'    If ConfirmYN = vbNo Then Exit Sub
'End If
'
'cmdImport3.Enabled = False: Screen.MousePointer = 11: dg3.Enabled = False
'
''開始匯入檔案
''        If Len(Trim(filLocal3.FileName)) = 0 Then
''            Exit Sub
''        End If
'''Dim fl_file As Scripting.File
'''Set fso = New FileSystemObject
'''If fso.FileExists(strTranFileName) = False Then Exit Sub
'''
'''SSTab1.Tab = 3
''        If strTranFileName = "" Then
''            'Get_CustInv = False
''            Exit Sub
''        End If
'Dim rsReturnOrders As New ADODB.Recordset           '退貨訂單資料
'Dim strRow As String    '讀取每一行文字
'Dim strField As String  '讀取每個區隔的資料
'Dim intPointer As Integer
''Set rsReturnOrders = New Recordset
'With rsReturnOrders
'    .Fields.Append "LoadNo", adChar, 12, adFldUpdatable
'    .Fields.Append "OrderNo", adChar, 12, adFldUpdatable
'    .Fields.Append "OrderDate", adChar, 8, adFldUpdatable
'    .Fields.Append "DeliveryDate", adChar, 8, adFldUpdatable
'    .Fields.Append "CustomerID", adChar, 10, adFldUpdatable
'    .Fields.Append "CustomerName1", adChar, 30, adFldUpdatable
'    .Fields.Append "CustomerName2", adChar, 30, adFldUpdatable
'    .Fields.Append "CustomerName3", adChar, 30, adFldUpdatable
'    .Fields.Append "Address1", adChar, 30, adFldUpdatable
'    .Fields.Append "Address2", adChar, 30, adFldUpdatable
'    .Fields.Append "Address3", adChar, 30, adFldUpdatable
'    .Fields.Append "DeliveryADRS1", adChar, 30, adFldUpdatable
'    .Fields.Append "DeliveryADRS2", adChar, 30, adFldUpdatable
'    .Fields.Append "DeliveryADRS3", adChar, 30, adFldUpdatable
'    .Fields.Append "ZIP", adChar, 3, adFldUpdatable
'    .Fields.Append "SKU", adChar, 14, adFldUpdatable
'    .Fields.Append "SKUDescription", adChar, 30, adFldUpdatable
'    .Fields.Append "OrderCS", adChar, 7, adFldUpdatable
'    .Fields.Append "OrderCT", adChar, 7, adFldUpdatable
'    .Fields.Append "OrderEA", adChar, 7, adFldUpdatable
'    .Fields.Append "OTotalEA", adChar, 7, adFldUpdatable
'    .Fields.Append "CS", adChar, 4, adFldUpdatable
'    .Fields.Append "Casecnt", adChar, 4, adFldUpdatable
'    .Fields.Append "EA", adChar, 4, adFldUpdatable
'    .Fields.Append "PickCS", adChar, 7, adFldUpdatable
'    .Fields.Append "PickCT", adChar, 7, adFldUpdatable
'    .Fields.Append "PickEA", adChar, 7, adFldUpdatable
'    .Fields.Append "PTotalEA", adChar, 7, adFldUpdatable
'    .Fields.Append "OrderComments", adChar, 60, adFldUpdatable
'    .Fields.Append "AssignedDate", adChar, 8, adFldUpdatable
'    .Fields.Append "OrderLine", adChar, 2, adFldUpdatable
'    .Fields.Append "Warehouse", adChar, 12, adFldUpdatable
'    .Fields.Append "DEPOTNO", adChar, 1, adFldUpdatable
'    .Fields.Append "REPACKAGE", adChar, 1, adFldUpdatable
'    .Fields.Append "WT", adChar, 11, adFldUpdatable
'    .Fields.Append "CBM", adChar, 11, adFldUpdatable
'    .Fields.Append "CARNO", adChar, 10, adFldUpdatable
'    .Fields.Append "AssignedFlag", adChar, 1, adFldUpdatable
'    .Fields.Append "CustomerPO", adChar, 12, adFldUpdatable
'    .Fields.Append "OrderType", adChar, 1, adFldUpdatable
'    .Fields.Append "RETURNCT", adChar, 7, adFldUpdatable
'    .Fields.Append "UTLOrderLine", adChar, 4, adFldUpdatable
'    .Fields.Append "ULTCustPO", adChar, 16, adFldUpdatable
'    .CursorType = adOpenKeyset
'    .LockType = adLockOptimistic
'    .Open
'End With
'
''txt to rs
'Open filLocal3.Path & "\" & filLocal3.FileName For Input As #1
'
'Do Until EOF(1)
'Line Input #1, strRow
'intPointer = 1
'With rsReturnOrders
'    .AddNew
'    .Fields(0) = Trim(GetWord(strRow, intPointer, 12))
'    .Fields(1) = Trim(GetWord(strRow, intPointer, 12))
'    .Fields(2) = Trim(GetWord(strRow, intPointer, 8))
'    .Fields(3) = Trim(GetWord(strRow, intPointer, 8))
'    .Fields(4) = Trim(GetWord(strRow, intPointer, 10))
'    .Fields(5) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(6) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(7) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(8) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(9) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(10) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(11) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(12) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(13) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(14) = Trim(GetWord(strRow, intPointer, 3))
'    .Fields(15) = Trim(GetWord(strRow, intPointer, 14))
'    .Fields(16) = Trim(GetWord(strRow, intPointer, 30))
'    .Fields(17) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(18) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(19) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(20) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(21) = Trim(GetWord(strRow, intPointer, 4))
'    .Fields(22) = Trim(GetWord(strRow, intPointer, 4))
'    .Fields(23) = Trim(GetWord(strRow, intPointer, 4))
'    .Fields(24) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(25) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(26) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(27) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(28) = Trim(GetWord(strRow, intPointer, 60))
'    .Fields(29) = Trim(GetWord(strRow, intPointer, 8))
'    .Fields(30) = Trim(GetWord(strRow, intPointer, 2))
'    .Fields(31) = Trim(GetWord(strRow, intPointer, 12))
'    .Fields(32) = Trim(GetWord(strRow, intPointer, 1))
'    .Fields(33) = Trim(GetWord(strRow, intPointer, 1))
'    .Fields(34) = Trim(GetWord(strRow, intPointer, 11))
'    .Fields(35) = Trim(GetWord(strRow, intPointer, 11))
'    .Fields(36) = Trim(GetWord(strRow, intPointer, 10))
'    .Fields(37) = Trim(GetWord(strRow, intPointer, 1))
'    .Fields(38) = Trim(GetWord(strRow, intPointer, 12))
'    .Fields(39) = Trim(GetWord(strRow, intPointer, 1))
'    .Fields(40) = Trim(GetWord(strRow, intPointer, 7))
'    .Fields(41) = Trim(GetWord(strRow, intPointer, 4))
'    .Fields(42) = Trim(GetWord(strRow, intPointer, 16))
'
'    .Update
'    .MoveFirst
'End With
'
'Loop
'Close
'
'rsReturnOrders.MoveFirst
'Set dg3.DataSource = rsReturnOrders
'With dg3
'    .Columns(17).Alignment = dbgRight
'    .Columns(18).Alignment = dbgRight
'    .Columns(19).Alignment = dbgRight
'    .Columns(20).Alignment = dbgRight
'    .Columns(21).Alignment = dbgRight
'    .Columns(22).Alignment = dbgRight
'    .Columns(23).Alignment = dbgRight
'    .Columns(24).Alignment = dbgRight
'    .Columns(25).Alignment = dbgRight
'    .Columns(26).Alignment = dbgRight
'    .Columns(27).Alignment = dbgRight
'
'End With
'
'rsReturnOrders.MoveFirst
'strOrderNo = ""
'int_repeat = 0
'int_order = 0
'int_orderline = 0
'
'cn.BeginTrans
'Do While Not rsReturnOrders.EOF
'    DoEvents: DoEvents
'
'    If strOrderNo <> rsReturnOrders("OrderNo") Then '資料檢驗--判斷訂單編號已訂是否要在 [訂單主檔] 中新增一筆
'        strOrderNo = rsReturnOrders("OrderNo")
'
'        '資料檢驗--判斷訂單是否重複
'        Call Confirm_Recordset_Closed(tmp_rs)
'        str_ExternOrderKey = "R" & Trim(rsReturnOrders("OrderNO"))
'        str_SQL = "select ExternOrderKey from orders where ExternOrderKey='" & str_ExternOrderKey & "' "
'        tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If tmp_rs.EOF Then
'        '訂單主檔新增一筆
'            str_SQL = "select isnull(max(orderkey),0) from orders"
'            Call Confirm_Recordset_Closed(tmp_rs)
'            tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'            str_Orderkey = StrPadLeft(Val(Trim(tmp_rs.Fields(0))) + 1, 10, 0)
'            tmp_rs.Close
'            str_CustomerID = Trim(rsReturnOrders.Fields(4).Value)
'            '0 "LoadNo",1"OrderNo",2"OrderDate", 3"DeliveryDate",4"CustomerID",5"CustomerName1",6"CustomerName2",7"CustomerName3",8"Address1",9"Address2",10"Address3",
'            '11"DeliveryADRS1",12"DeliveryADRS2",13"DeliveryADRS3",14"ZIP",15"SKU",16"SKUDescription",17"OrderCS",18"OrderCT",19"OrderEA",20"OTotalEA",
'            '21"CS",22"CT",23"EA",24"PickCS",25"PickCT",26"PickEA",27"PTotalEA",28"OrderComments",29"AssignedDate",30"OrderLine",
'            '31"CPNO",32"DEPOTNO",33"REPACKAGE",34"RLNO",35"RDNO",36"CARNO",37"AssignedFlag",37"CustomerPO",39"OrderType",40"RETURNCT",
'            '41"UTLOrderLine",42, adChar, 16, adFldUpdatable
'
'            'OrderKey,StorerKey,ExternOrderKey,OrderDate,DeliveryDate,Priority,ConsigneeKey,C_contact1,C_Contact2,C_Company,C_Address1,C_Address2,C_Address3,C_Address4,C_City,
'            'C_State,C_Zip,C_Country,C_ISOCntryCode,C_Phone1,C_Phone2,C_Fax1,C_Fax2,C_vat,BuyerPO, BillToKey, B_contact1, B_Contact2, B_Company, B_Address1, B_Address2, B_Address3, B_Address4, B_City, B_State, B_Zip, B_Country,
'            'B_ISOCntryCode, B_Phone1, B_Phone2, B_Fax1, B_Fax2, B_Vat, IncoTerm, PmtTerm, OpenQty, Status, DischargePlace, DeliveryPlace, IntermodalVehicle, CountryOfOrigin, CountryDestination, UpdateSource, Type, OrderGroup, Door,
'            'Route, Stop, Notes, EffectiveDate, AddDate, AddWho, EditDate, EditWho, TrafficCop, ArchiveCop, ContainerType, ContainerQty, BilledContainerQty, DoRoute, CustomerOrderkey, PIVNO, Amount
'            str_Note = Trim(rsReturnOrders.Fields(28).Value)
'            str_Note = Replace(str_Note, "'", " ")
'            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,DeliveryDate,ConsigneeKey,C_Company,C_Zip,C_Address1,C_Address2,C_Address3,CustomerOrderkey,Notes,UpdateSource,DischargePlace,B_PHONE2,Type,OrderGroup,ContainerQty,BilledContainerQty,b_fax2) " & _
'                "VALUES ('" & str_Orderkey & "','" & str_ExternOrderKey & "','" & Trim(rsReturnOrders.Fields(39)) & "','UTL','" & Trim(rsReturnOrders.Fields(2)) & "','" & Trim(rsReturnOrders.Fields(3)) & "', " & _
'                "'" & str_CustomerID & "','" & Trim(rsReturnOrders.Fields(5)) & "','" & Trim(rsReturnOrders.Fields(14)) & "','" & Trim(rsReturnOrders.Fields(11)) & "','" & Trim(rsReturnOrders.Fields(12)) & "'," & _
'                "'" & Trim(rsReturnOrders.Fields(13)) & "','" & Trim(rsReturnOrders.Fields(38)) & "','" & str_Note & "','" & filLocal3.FileName & "','" & Trim(rsReturnOrders.Fields(31)) & "','01','" & Trim(rsReturnOrders.Fields(39)) & "','" & Trim(rsReturnOrders.Fields(33)) & "','" & Trim(rsReturnOrders.Fields(18)) & "','" & Trim(rsReturnOrders.Fields(18)) & "','0')"
'            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'            int_order = int_order + 1
'        End If
'End If
'
'        '資料檢驗--判斷訂單明細是否重複
'        Call Confirm_Recordset_Closed(tmp_rs)
'        str_SQL = "select ExternOrderKey from ORDERDETAIL where ExternOrderKey='R" & Trim(rsReturnOrders.Fields(1)) & "' and OrderLineNumber= '" & Trim(rsReturnOrders.Fields(41)) & "'"
'        tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If Not tmp_rs.EOF Then
'            int_repeat = int_repeat + 1
'            Call FTPlog("訂單明細重複" & str_SQL)
'            GoTo nextloop
'        End If
'
'        '訂單明細資料新增
'        'OrderKey, OrderLineNumber, OrderDetailSysId, ExternOrderKey, ExternLineNo, Sku, StorerKey, ManufacturerSku, RetailSku, AltSku, OriginalQty, OpenQty, ShippedQty, AdjustedQty, QtyPreAllocated, QtyAllocated, QtyPicked, UOM, PackKey, PickCode, CartonGroup, Lot, ID, Facility, Status, UnitPrice, Tax01, Tax02, ExtendedPrice, UpdateSource, Lottable01, Lottable02, Lottable03, Lottable04, Lottable05, EffectiveDate, AddDate, AddWho, EditDate, EditWho, TrafficCop, ArchiveCop, TariffKey, Lottable06, Lottable07, Lottable08, Lottable09, Lottable10, Lottable11, Beginqty
'        str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber, ExternOrderKey,ExternLineNo,Sku,StorerKey,OriginalQty,openqty,qtypreallocated,shippedqty,qtyallocated,adjustedqty,qtypicked,Lottable01, Lottable02,PackKey,Lottable03,pickcode)" & _
'         "VALUES ('" & str_Orderkey & "','" & Trim(rsReturnOrders.Fields(41)) & "','" & str_ExternOrderKey & "','" & Trim(rsReturnOrders.Fields(40)) & "','" & Trim(rsReturnOrders.Fields(15)) & "','UTL', " & _
'         "'" & Round(CLng(rsReturnOrders.Fields(20)) / IIf(CLng(rsReturnOrders.Fields(22)) = 0, 1, CLng(rsReturnOrders.Fields(22))), 3) & "','" & Trim(rsReturnOrders.Fields(17)) & "','" & Trim(rsReturnOrders.Fields(17)) & "','" & Trim(rsReturnOrders.Fields(19)) & "','" & Trim(rsReturnOrders.Fields(19)) & "','" & Trim(rsReturnOrders.Fields(20)) & "','" & Trim(rsReturnOrders.Fields(20)) & "','" & Trim(rsReturnOrders.Fields(37)) & "','" & Trim(rsReturnOrders.Fields(29)) & "','" & Trim(rsReturnOrders.Fields(15)) & "','" & Trim(rsReturnOrders.Fields(27)) & "','" & Trim(rsReturnOrders.Fields(22)) & "')"
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'        '資料檢驗--判斷SKU否存在
'        str_SQL = "select sku from sku where sku='" & Trim(rsReturnOrders("sku")) & "'"
'        Call Confirm_Recordset_Closed(tmp_rs)
'        tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If tmp_rs.EOF Then
'            str_DESCR = Trim(rsReturnOrders("SKUDescription"))
'            str_DESCR = Replace(str_DESCR, "'", " ")
'            'StorerKey, Sku, DESCR, SUSR1, SUSR2, SUSR3, SUSR4, SUSR5, MANUFACTURERSKU, RETAILSKU, ALTSKU, PACKKey, STDGROSSWGT, STDNETWGT, STDCUBE, TARE, CLASS, ACTIVE, SKUGROUP, Tariffkey, BUSR1, BUSR2, BUSR3, BUSR4, BUSR5, LOTTABLE01LABEL, LOTTABLE02LABEL, LOTTABLE03LABEL, LOTTABLE04LABEL, LOTTABLE05LABEL, NOTES1, NOTES2, PickCode, StrategyKey, CartonGroup, PutCode, PutawayLoc, PutawayZone, InnerPack, Cube, GrossWgt, NetWgt, ABC, CycleCountFrequency, LastCycleCount, ReorderPoint, ReorderQty, StdOrderCost, CarryCost, Price, Cost, ReceiptHoldCode, ReceiptInspectionLoc, TrafficCop, ArchiveCop, IOFlag, TareWeight, LotxIdDetailOtherlabel1, LotxIdDetailOtherlabel2, LotxIdDetailOtherlabel3, AvgCaseWeight, TolerancePct, SkuRotat01, SkuRotat02, SkuRotat03, SkuRotat04, SkuRotat05, DefaultRotation, AllocParm, onreceiptcopypackkey, replenishmentQty
'            str_CSKU = StringCleaner(Trim(rsReturnOrders.Fields(15)), "'")
'            str_SQL = "INSERT sku (StorerKey, Sku, DESCR,PACKKey,STDGROSSWGT,BUSR4) " & _
'            "VALUES ('UTL','" & Trim(rsReturnOrders.Fields(15)) & "','" & str_DESCR & "','" & Trim(rsReturnOrders.Fields(15)) & "','" & Trim(rsReturnOrders.Fields(34)) / 1000 & "','" & Trim(rsReturnOrders.Fields(35)) / 1000 & "' )"
'            cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
'            'PackKey, PackDescr, PackUOM1, CaseCnt, ISWHQty1, ReplenishUOM1, ReplenishZone1, CartonizeUOM1, LengthUOM1, WidthUOM1, HeightUOM1, CubeUOM1, PackUOM2, InnerPack, ISWHQty2, ReplenishUOM2, ReplenishZone2, CartonizeUOM2, LengthUOM2, WidthUOM2, HeightUOM2, CubeUOM2, PackUOM3, Qty, ISWHQty3, ReplenishUOM3, ReplenishZone3, CartonizeUOM3, LengthUOM3, WidthUOM3, HeightUOM3, CubeUOM3, PackUOM4, Pallet, ISWHQty4, ReplenishUOM4, ReplenishZone4, CartonizeUOM4, LengthUOM4, WidthUOM4, HeightUOM4, CubeUOM4, PalletWoodLength, PalletWoodWidth, PalletWoodHeight, PalletTI, PalletHI, PackUOM5, Cube, ISWHQty5, PackUOM6, GrossWgt, ISWHQty6, PackUOM7, NetWgt, ISWHQty7, PackUOM8, OtherUnit1, ISWHQty8, ReplenishUOM8, ReplenishZone8, CartonizeUOM8, LengthUOM8, WidthUOM8, HeightUOM8, PackUOM9, OtherUnit2, ISWHQty9, ReplenishUOM9, ReplenishZone9, CartonizeUOM9, LengthUOM9, WidthUOM9, HeightUOM9, AddDate, AddWho, EditDate, EditWho, TrafficCop, ArchiveCop, Timestamp
'            str_SQL = "INSERT pack (PACKKey,PackDescr,GrossWgt, PalletHI, PalletTI,CaseCnt) " & _
'            "VALUES ('" & Trim(rsReturnOrders.Fields(15)) & "','" & str_DESCR & "','" & Trim(rsReturnOrders.Fields(34)) / 1000 & "', " & _
'             "'0','0','1')"
'            cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
'        End If
'
'        tmp_rs.Close
'        int_orderline = int_orderline + 1
'nextloop:
'
'    rsReturnOrders.MoveNext
'Loop
'
'cn.CommitTrans
'
''備份檔案
'If Dir("C:\From_ids\Backup\UTLR", vbDirectory) = "" Then MkDir "C:\From_ids\Backup\UTLR"
'
'If filLocal3.Path = "C:\From_ids\Backup\UTLR" Then
'Else
'    FileCopy filLocal3.Path & "\" & filLocal3.FileName, "C:\From_ids\Backup\UTLR\xrsldnl" & Format(Now(), "yyyymmddhhmmss") & ".txt"
'    Kill filLocal3.Path & "\" & filLocal3.FileName
'End If
'
'If int_repeat > 0 Then
'    msg_text = "有 " & int_repeat & " 筆訂單明細重複轉檔"
'    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'End If
'
'msg_text = "匯入 " & int_order & " 筆訂單， " & int_orderline & " 筆明細，文字檔 " & filLocal3.FileName & " 備份於 C:\From_ids\Backup\UTLR\"
'MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'Call FTPlog("匯入 " & int_order & " 筆訂單， " & int_orderline & " 筆明細，文字檔 " & filLocal3.FileName)
'filLocal3.Refresh
'
'cmdImport3.Enabled = True: Screen.MousePointer = 0: dg3.Enabled = True
'Exit Sub
'
'err_Handle:
'Close #1
''cn.RollbackTrans
'Dim tmpString As String
'msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
'tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
'CreateErrorLog Me.Name & "-匯入", Me.Caption, "cmdImport3_Click", tmpString
'MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'cmdImport3.Enabled = True: Screen.MousePointer = 0: dg3.Enabled = True
'End Sub

Private Sub cmdLogOn_Click()
    On Error GoTo LogOnError
    
    If txtServer = "" Or txtPassword = "" Then
        MsgBox "你必須輸入ftp Server與密碼", vbInformation + vbOKOnly, "LogOn Failure"
        Exit Sub
    End If
    
    'Set status label
    lblStatus = "Connecting"
    'Set protocol and server
    ITC.Protocol = icFTP
    ITC.url = txtServer
    'If no username is entered default to anonymous
    If txtUserName = "" Then
        ITC.UserName = "anonymous"
    Else
        ITC.UserName = txtUserName
    End If
    ITC.Cancel
    cmdLogOn.Enabled = False
    'Set the password and connect
    ITC.PassWord = txtPassword
    ITC.RequestTimeout = 20
    
    ITC.Execute , "CD " & Chr(34) & "/Bestg/Alc" & Chr(34)
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
        
    ITC.Execute , "DIR"
    Do While ITC.StillExecuting
        DoEvents: DoEvents: DoEvents
    Loop
    
    'Set status label, disable the log on button, and enable the log off button
    lblStatus = "已連線至alc資料夾"
    cmdLogOn.Enabled = False
    cmdNewFolder.Enabled = True
    cmdDelete.Enabled = True
    cmdUpFolder.Enabled = True
    imgSendFile.Enabled = True
    imgReceiveFile.Enabled = True
    lstRemoteFile.Enabled = True
    cmdLogOff.Enabled = True
    
    Exit Sub
    
LogOnError:
    'If logon fails alert the user
    MsgBox "登入錯誤....", vbOKOnly + vbInformation, "登入錯誤"
    ITC.Cancel
    lblStatus = "Not Connected"
    cmdLogOn.Enabled = True
    cmdNewFolder.Enabled = False
    cmdDelete.Enabled = False
    cmdUpFolder.Enabled = False
    imgSendFile.Enabled = False
    imgReceiveFile.Enabled = False
    lstRemoteFile.Enabled = False
    cmdLogOff.Enabled = False
End Sub

Private Sub cmdLogOff_Click()
'Clear the list of remote files and log off
lstRemoteFile.Clear
ITC.Cancel

Do Until ITCReady(False)
    DoEvents: DoEvents: DoEvents: DoEvents
Loop

lblStatus = "Closing Connection"

If ITCReady(False) Then
    ITC.Execute , "CLOSE"
Else
    ITC.Cancel
End If
lblStatus = "Not Connected"
cmdLogOn.Enabled = False
cmdNewFolder.Enabled = False
cmdDelete.Enabled = False
cmdUpFolder.Enabled = False
imgSendFile.Enabled = False
imgReceiveFile.Enabled = False
lstRemoteFile.Enabled = False
cmdLogOff.Enabled = False
cmdLogOn.Enabled = True
End Sub
'Private Sub cmdLogOff3_Click()
'If int3.StillExecuting Then MsgBox "請稍等.  FTP伺服器執行中", vbInformation + vbOKOnly, "忙碌中": Exit Sub
'    'Clear the list of remote files and log off
'    lstRemoteFile3.Clear
'
''    Do Until int3Ready(True)
''        DoEvents: DoEvents: DoEvents: DoEvents
''    Loop
''
'    lblStatus3 = "Closing Connection"
'
'    If int3Ready(False) Then
'        int3.Execute , "CLOSE"
'    Else
'        int3.Cancel
'    End If
'    lblStatus3 = "Not Connected"
'    cmdLogon3.Enabled = False
''    cmdNewFolder.Enabled = False
''    cmdDelete.Enabled = False
''    cmdUpFolder.Enabled = False
''    imgSendFile.Enabled = False
'    cmdImport3.Enabled = False
'    lstRemoteFile3.Enabled = False
'    cmdLogOff3.Enabled = False
'    cmdLogon3.Enabled = True
'    int3.Cancel
'End Sub

Private Sub cmdDelete_Click()
'暫不使用
'If the itc is ready, ask user if they want to delete it, if so then delete
If ITCReady(True) Then
    If MsgBox("確定刪除 " & lstRemoteFile.Text & " ?", vbQuestion + vbOKCancel, "Delete") = vbOK Then
        If Right(lstRemoteFile.Text, 1) <> "/" Then
            ITC.Execute , "DELETE " & Chr(34) & lstRemoteFile.Text & Chr(34)
        Else
            ITC.Execute , "RMDIR " & Chr(34) & lstRemoteFile.Text & Chr(34)
        End If
        
        Do Until ITCReady(False)
            DoEvents: DoEvents: DoEvents: DoEvents
        Loop
        lstRemoteFile.Clear
        ITC.Execute , "DIR"
        lblStatus = "已連線"
    End If
End If
End Sub

Private Sub cmdNewFolder_Click()
    '暫不使用
    'If the itc is ready then check to make sure the folder doesn't already exist
    Dim FolderName As String, i As Double
    If ITCReady(True) Then
        FolderName = InputBox("Enter new folder name", "New Folder")
        For i = 0 To lstRemoteFile.ListCount
            If FolderName & "/" = lstRemoteFile.List(i) Then
                MsgBox "Folder " & FolderName & " already exists.", vbInformation + vbOKOnly, "New Folder"
                Exit Sub
            End If
        Next i
        'Create the new folder then refresh the remote file list
        ITC.Execute , "MKDIR " & Chr(34) & FolderName & Chr(34)
        Do Until ITCReady(False)
            DoEvents: DoEvents: DoEvents: DoEvents
        Loop
        
        lstRemoteFile.Clear
        ITC.Execute , "DIR"
        lblStatus = "已連線"
    End If
End Sub

Private Sub cmdUpFolder_Click()
'暫不使用
'If the itc is ready then move up one directory and refresh the remote files list
If ITCReady(True) Then
    ITC.Execute , "CDUP"
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    lstRemoteFile.Clear
    ITC.Execute , "DIR"
    lblStatus = "已連線"
    
End If
End Sub

Private Sub Command2_Click()

strTranFileName = filLocalFileT7.Path & "\" & filLocalFileT7.FileName
If Len(RTrim(cboSheetT7)) = 0 Then MsgBox "請選擇工作表", 64, Me.Caption: Exit Sub
If rsMainT7.EOF Or rsMainT7 Is Nothing Then Exit Sub
On Error GoTo err_Handle

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT7.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT7.Enabled = True: dgMainT7.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

'到貨日期檢查
rsMainT7.MoveFirst
Do While Not rsMainT7.EOF

If Format(myExCharFilter(Trim(rsMainT7("Deliv.date"))), "YYYY/MM/DD") < Format(Now, "YYYY/MM/DD") Then: MsgBox "到貨日小於今日，訂單轉入終止!", 16, Me.Caption: Exit Sub

    rsMainT7.MoveNext
Loop

Tran_Level = cn.BeginTrans
cmdImportT7.Enabled = False: dgMainT7.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlineNumber As Integer, intTmp As Integer, lngInnerpack As Long
Dim strOrderNo As String, strWHOrderNo As String, strWH As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String, strLot06 As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset

rsMainT7.MoveFirst

''CT or BEST判斷
'strWHOrderNo = UCase(Trim(rsMainT7("Delivery")))
'
'If (Trim(rsMainT7("sloc")) = "0007" Or Trim(rsMainT7("sloc")) = "0008" Or Trim(rsMainT7("sloc")) = "0009" Or Trim(rsMainT7("sloc")) = "0010") Then
'    strWH = "BEST"
'Else
'    strWH = "CT"
'End If

Do While Not rsMainT7.EOF

    '資料檢驗--判斷訂單數是否為0
    If Trim(rsMainT7("Qty (stckpg unit)")) = 0 Then intNotBest = intNotBest + 1: GoTo next1

'    DoEvents: DoEvents
    
    'CT & BEST雙倉出貨判斷
    If strWHOrderNo = UCase(Trim(rsMainT7("Delivery"))) Then
        If (strWH = "BEST" And (Trim(rsMainT7("sloc")) = "0001" Or Trim(rsMainT7("sloc")) = "0005")) Or (strWH = "CT" And (Trim(rsMainT7("sloc")) <> "0001" Or Trim(rsMainT7("sloc")) <> "0005") = False) Then
            cmdImportT7.Enabled = True: Screen.MousePointer = 0: dgMainT7.Enabled = True: cn.RollbackTrans: Tran_Level = 0
            MsgBox "發現 CT & BEST 雙倉出貨訂單，請與客戶確認訂單正確無誤！", 16, "訂單轉入終止 "
            Exit Sub
        End If
    End If
    
    '是否為佰事達廠別倉別-->跳下一筆
    If Trim(rsMainT7("plnt")) = "1119" And (Trim(rsMainT7("sloc")) = "0001" Or Trim(rsMainT7("sloc")) = "0005") Then
        
        strWH = "CT"
        '檢查訂單量是否出現小數點
        If Val(rsMainT7("Qty (stckpg unit)")) <> Round(Val(rsMainT7("Qty (stckpg unit)")), 0) Then
            cmdImportT7.Enabled = True: Screen.MousePointer = 0: dgMainT7.Enabled = True: cn.RollbackTrans: Tran_Level = 0
            MsgBox "發現訂單量出現小數點，請與客戶確認正確訂單量！", 16, "訂單轉入終止 "
            Exit Sub
        End If
        
        '資料檢驗--判斷SKU是否存在
        str_SQL = "select sku,innerpack from gv_skuxpack where sku='" & Trim(rsMainT7("Material")) & "' and Storerkey = 'LNSL01' "
        
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
            cn.RollbackTrans: Tran_Level = 0: cmdImportT7.Enabled = True: dgMainT7.Enabled = True: Screen.MousePointer = 0
            MsgBox "訂單發現新品號 (" & Trim(rsMainT7("Material")) & ") ，訂單轉入終止!!"
            Exit Sub
        End If
        lngInnerpack = tmp_Rs("Innerpack")

    Else
        intNotBest = intNotBest + 1
        strWH = "BEST"
        GoTo next1

    End If
                  
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT7("Delivery"))) Then
        strOrderNo = UCase(Trim(rsMainT7("Delivery")))
        int_orderlineNumber = 0
        blDuplicationOrder = False
            
        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT7("Delivery"))) & "' and storerkey = 'LNSL01' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
            strFacility = "佰事達北倉"
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,BillToKey,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT7("Delivery"))) & "','I','LNSL01',getdate(),'" & myExCharFilter(Trim(rsMainT7("Deliv.date"))) & "','" & strFacility & "','" & _
            myExCharFilter(Trim(rsMainT7("Ship-To"))) & "','" & GetWord(myExCharFilter(Trim(rsMainT7("Name of the ship-to party"))), 1, 45) & "','','','','','','" & myExCharFilter(Trim(rsMainT7("Sold-to pt"))) & "','" & myExCharFilter(Trim(rsMainT7("po no"))) & "','" & myExCharFilter(Trim(rsMainT7("remarks"))) & "','" & filLocalFileT7.FileName & "','','" & User_id & "','" & User_id & "','' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT7("Delivery")) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '訂單重複不增加明細
        If blDuplicationOrder = False Then
            '增加明細
            int_orderlineNumber = int_orderlineNumber + 1
            
            '數量轉換
            intQTY = Trim(rsMainT7("Qty (stckpg unit)"))
'            If Trim(rsMainT7("Material")) = "12129314" Then intQTY = intQTY * IIf(lngInnerpack = 0, 1, lngInnerpack)
            'If lngInnerpack > 0 Then intQTY = intQTY * lngInnerpack
            
            '倉別轉換
            strLot06 = myExCharFilter(Trim(rsMainT7("sloc")))
            
            If strLot06 = "0001" Then
               strLot06 = "R01"
            ElseIf strLot06 = "0002" Then
               strLot06 = "R01"
            ElseIf strLot06 = "0005" Then
               strLot06 = "R08"
            End If
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,Externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable05,Lottable06, Facility,UOM,Unitprice,notes) " & _
            "VALUES ('" & str_Orderkey & "','" & int_orderlineNumber & "','" & myExCharFilter(Trim(rsMainT7("Item"))) & "','" & myExCharFilter(Trim(rsMainT7("Delivery"))) & "','" & myExCharFilter(Trim(rsMainT7("material"))) & "','LNSL01'," & _
            "'" & intQTY & "','" & intQTY & "','" & myExCharFilter(Trim(rsMainT7("Batch"))) & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT7("BUn"))) & "','0','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:

        If (Trim(rsMainT7("sloc")) = "0001" Or Trim(rsMainT7("sloc")) = "0005") Then
            strWH = "CT"
        Else
            strWH = "BEST"
        End If

        strWHOrderNo = UCase(Trim(rsMainT7("Delivery")))
        rsMainT7.MoveNext
        
Loop

'補客戶資料
cn.Execute "exec gs_ordersupdate 'LNSL01' ", RowsAffect, adExecuteNoRecords

'新客戶檢查
str_SQL = "select 貨主=storerkey , 貨主訂單單號=externorderkey , 訂單類別 = priority , 訂單日期=orderdate , 到貨日期=deliverydate , 客戶編號=consigneekey , 檢查日期 = getdate() " & _
        "from orders o " & _
        "Where o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel顯示
    Call Recordset2Excel("缺客戶資料", tmp_Rs)
    If Dir("C:\BEST\LNSL01\缺客戶資料", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\缺客戶資料"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\缺客戶資料\缺客戶資料_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmdImportT7.Enabled = True: dgMainT7.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
    MsgBox "發現新客戶，訂單轉入中止!", 16, Me.Caption
    
    Exit Sub
End If

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細 (共 " & rsMainT7.RecordCount & " 筆明細)，TMS單號： " & strOrderKeyS & "~" & str_Orderkey & "，檔案 " & filLocalFileT7.FileName & " 備份於 C:\BEST\LNSL01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細，檔案 " & filLocalFileT7.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT7.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 料號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\LNSL01\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''備份至FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT7.FileName

'備份檔案
If Dir("C:\BEST\LNSL01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\Orders\Backup"
If Dir("C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT7.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT7.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & mySplit(filLocalFileT7.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT7.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT7.Refresh: cboSheetT7.Clear
Screen.MousePointer = 0: cmdImportT7.Enabled = True: dgMainT7.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "手開單加允收期-匯入", Me.Caption, "cmd2_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT7.Enabled = True: Screen.MousePointer = 0: dgMainT7.Enabled = True
End Sub

Private Sub Command3_Click()

'資料排序
Recordset2Excel "TEST", rsMainT17_1

'..在此編輯EXCEL
With MyXlsApp
    
End With

Set MyXlsApp = Nothing
End Sub



Private Sub Command5_Click()

'資料排序
Recordset2Excel "訂單主檔", rsMainT16
Recordset2Excel "訂單明細檔", rsMainT16_1

'..在此編輯EXCEL
With MyXlsApp
    
End With

Set MyXlsApp = Nothing

End Sub

Private Sub dg_CustInv_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
'DataGrid ColResize事件中加入下段程式碼，用以記錄欄寬
If Len(dg_CustInv.Columns(ColIndex).DataField) = 0 Then Exit Sub
SaveSetting App.title, Me.Caption & "dg_CustInv", dg_CustInv.Columns(ColIndex).DataField, dg_CustInv.Columns(ColIndex).Width
End Sub

Private Sub dgMainT16_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT16
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT16_1_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT16_1
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT16_2_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT16_2
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT16_3_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT16_3
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT17_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT17
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dgMainT18_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT18
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT15_Change()
    filLocalFileT15.Path = dirLocalDirT15.Path
End Sub

Private Sub dirLocalDirT16_Change()
    filLocalFileT16.Path = dirLocalDirT16.Path
End Sub


Private Sub dirLocalDirT17_Change()
    filLocalFileT17.Path = dirLocalDirT17.Path
End Sub

Private Sub dirLocalDirT18_Change()
    filLocalFileT18.Path = dirLocalDirT18.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo DriveError
    dirLocalDirT20.Path = drvLocalDriveT20.Drive
    Exit Sub
DriveError:
    MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
    Resume Next
End Sub

Private Sub dirLocalDirT19_Change()
    filLocalFileT19.Path = dirLocalDirT19.Path
End Sub

Private Sub dirLocalDirT20_Change()
       filLocalFileT20.Path = dirLocalDirT20.Path
End Sub


Private Sub dirLocalDirT21_Change()
    filLocalFileT21.Path = dirLocalDirT21.Path
End Sub

Private Sub dirLocalDirT22_Change()
    filLocalFileT22.Path = dirLocalDirT22.Path
End Sub


Private Sub drvLocalDrive_Change()
    On Error GoTo DriveError
    dirLocalDir.Path = drvLocalDrive.Drive
    Exit Sub
DriveError:
    MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
    Resume Next
End Sub

Private Sub dirLocalDir_Change()
    filLocalFile.Path = dirLocalDir.Path
End Sub

Private Sub drvLocalDriveT15_Change()
On Error GoTo DriveError
dirLocalDirT15.Path = drvLocalDriveT15.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub drvLocalDriveT16_Change()
On Error GoTo DriveError
dirLocalDirT16.Path = drvLocalDriveT16.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub


Private Sub drvLocalDriveT17_Change()
On Error GoTo DriveError
dirLocalDirT17.Path = drvLocalDriveT17.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub drvLocalDriveT18_Change()
On Error GoTo DriveError
dirLocalDirT18.Path = drvLocalDriveT18.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub

Private Sub drvLocalDriveT19_Change()
On Error GoTo DriveError
dirLocalDirT19.Path = drvLocalDriveT19.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub



Private Sub drvLocalDriveT20_Click()
On Error GoTo DriveError
dirLocalDirT20.Path = drvLocalDriveT20.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub




Private Sub drvLocalDriveT20_Change()
    On Error GoTo DriveError
    dirLocalDirT20.Path = drvLocalDriveT20.Drive
    Exit Sub
DriveError:
    MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
    Resume Next
End Sub

Private Sub drvLocalDriveT21_Change()
On Error GoTo DriveError
dirLocalDirT21.Path = drvLocalDriveT21.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub

Private Sub drvLocalDriveT22_Change()
On Error GoTo DriveError
dirLocalDirT22.Path = drvLocalDriveT22.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub


Private Sub filLocalFileT15_Click()

On Error GoTo err_Handle
Set rsMainT15 = Nothing: Set dgMainT15.DataSource = rsMainT15
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT15.Path, 1) = "\" Then
    strFilePath = filLocalFileT15.Path
Else
    strFilePath = filLocalFileT15.Path & "\"
End If

If Dir(strFilePath & filLocalFileT15.FileName) = "" Then: filLocalFileT15.Refresh: Exit Sub

cboSheetT15.Clear

If UCase(mySplit(filLocalFileT15.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT15.FileName)
    MyXlsApp.DisplayAlerts = False
  
    '列出所有工作表
    blDo = False
    cboSheetT15.Clear
    
    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT15.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT15.ListIndex = -1
    
    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT15.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub filLocalFileT16_Click()

On Error GoTo err_Handle
Set rsMainT16 = Nothing: Set dgMainT16.DataSource = rsMainT16
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT16.Path, 1) = "\" Then
    strFilePath = filLocalFileT16.Path
Else
    strFilePath = filLocalFileT16.Path & "\"
End If

If Dir(strFilePath & filLocalFileT16.FileName) = "" Then: filLocalFileT16.Refresh: Exit Sub

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub filLocalFileT17_Click()

On Error GoTo err_Handle
Set rsMainT17 = Nothing: Set dgMainT16.DataSource = rsMainT17
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT17.Path, 1) = "\" Then
    strFilePath = filLocalFileT17.Path
Else
    strFilePath = filLocalFileT17.Path & "\"
End If

If Dir(strFilePath & filLocalFileT17.FileName) = "" Then: filLocalFileT17.Refresh: Exit Sub

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub filLocalFileT18_Click()

On Error GoTo err_Handle
Set rsMainT18 = Nothing: Set dgMainT18.DataSource = rsMainT18
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT18.Path, 1) = "\" Then
    strFilePath = filLocalFileT18.Path
Else
    strFilePath = filLocalFileT18.Path & "\"
End If

If Dir(strFilePath & filLocalFileT18.FileName) = "" Then: filLocalFileT18.Refresh: Exit Sub

cboSheetT18.Clear

If UCase(mySplit(filLocalFileT18.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT18.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT18.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT18.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT18.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT18.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub


Private Sub filLocalFileT19_Click()

On Error GoTo err_Handle
Set rsMainT19 = Nothing: Set dgMainT19.DataSource = rsMainT19
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT19.Path, 1) = "\" Then
    strFilePath = filLocalFileT19.Path
Else
    strFilePath = filLocalFileT19.Path & "\"
End If

If Dir(strFilePath & filLocalFileT19.FileName) = "" Then: filLocalFileT19.Refresh: Exit Sub

cboSheetT19.Clear

If UCase(mySplit(filLocalFileT19.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT19.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT19.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT19.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT19.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT19.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub
Private Sub filLocalFileT20_Click()
On Error GoTo err_Handle
Set rsMainT20 = Nothing: Set dgMainT20.DataSource = rsMainT20
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT20.Path, 1) = "\" Then
    strFilePath = filLocalFileT20.Path
Else
    strFilePath = filLocalFileT20.Path & "\"
End If

If Dir(strFilePath & filLocalFileT20.FileName) = "" Then: filLocalFileT20.Refresh: Exit Sub

cboSheetT20.Clear

If UCase(mySplit(filLocalFileT20.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT20.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT20.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        
        cboSheetT20.AddItem MyXlsApp.Sheets(i).Name
  
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT20.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT20.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub filLocalFileT21_Click()

On Error GoTo err_Handle
Set rsMainT21 = Nothing: Set dgMainT21.DataSource = rsMainT21
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT21.Path, 1) = "\" Then
    strFilePath = filLocalFileT21.Path
Else
    strFilePath = filLocalFileT21.Path & "\"
End If

If Dir(strFilePath & filLocalFileT21.FileName) = "" Then: filLocalFileT21.Refresh: Exit Sub

cboSheetT21.Clear

If UCase(mySplit(filLocalFileT21.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT21.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT21.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        
        cboSheetT21.AddItem MyXlsApp.Sheets(i).Name
  
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT21.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT21.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub filLocalFileT22_Click()

On Error GoTo err_Handle
Set rsMainT22 = Nothing: Set dgMainT22.DataSource = rsMainT22
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT22.Path, 1) = "\" Then
    strFilePath = filLocalFileT22.Path
Else
    strFilePath = filLocalFileT22.Path & "\"
End If

If Dir(strFilePath & filLocalFileT22.FileName) = "" Then: filLocalFileT22.Refresh: Exit Sub

cboSheetT22.Clear

If UCase(mySplit(filLocalFileT22.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT22.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT22.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT22.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT22.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT22.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub


Private Sub lstRemoteFile_DblClick()
'If the itc is ready, check that the selected is a folder and change the directory
If ITCReady(True) Then
    If Right(lstRemoteFile.Text, 1) = "/" Then
        ITC.Execute , "CD " & Chr(34) & lstRemoteFile.Text & Chr(34)
    Else
        Call imgReceiveFile_Click
        Exit Sub
    End If
    
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    
    lstRemoteFile.Clear
    ITC.Execute , "DIR"
    lblStatus = "已連線"
End If
End Sub

Private Sub imgReceiveFile_Click()
Dim i As Double
    'If the ITC is not still executing then receive the file
    If ITCReady(True) = True Then
        'Check that they are not recieving a folder
        If Right(lstRemoteFile.Text, 1) = "/" Then
            MsgBox lstRemoteFile.Text & " is a folder and cannot be sent.", vbCritical + vbOKOnly, "Error"
            Exit Sub
        End If
        'Check that the file does not already exist on the computer, if it does exit sub
        For i = 0 To filLocalFile.ListCount
            If lstRemoteFile.Text = filLocalFile.List(i) Then
                MsgBox "檔案 " & Right(lstRemoteFile.Text, 18) & " 已存在.", vbInformation + vbOKOnly, "Recieve"
                Exit Sub
            End If
        Next i
        str_file = Trim(Right(lstRemoteFile.Text, 18))
        ITC.Execute , "GET " & Chr(34) & str_file & Chr(34) & " " & Chr(34) & filLocalFile.Path & "\" & str_file & Chr(34)
        lblStatus = "下載中請稍後...."
        Do Until ITCReady(False)
            DoEvents: DoEvents: DoEvents: DoEvents
        Loop
        filLocalFile.Refresh
        lblStatus = "已連線"
        
        '開始匯入檔案
        strTranFileName = filLocalFile.Path & "\" & str_file
        If Len(Trim(strTranFileName)) = 0 Then
            Exit Sub
        End If
        If FileLen(strTranFileName) = 0 Then
            msg_text = "檔案大小=0,檔名:" & str_file
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
            Exit Sub
        End If
        SSTab1.Tab = 1
        DoEvents: DoEvents
        Dim strRow As String    '讀取每一行文字
        Dim strField As String  '讀取每個區隔的資料
        Dim intPointer As Integer, int_OrderLine As Integer, int_Order As Integer, int_Repeat As Integer
        Set rs_Src = New Recordset
        With rs_Src
            .Fields.Append "OrderNo", adChar, 12, adFldUpdatable
            .Fields.Append "OrderType", adChar, 1, adFldUpdatable
            .Fields.Append "Division", adChar, 3, adFldUpdatable
            .Fields.Append "OrderDate", adChar, 8, adFldUpdatable
            .Fields.Append "DeliveryDate", adChar, 8, adFldUpdatable
            .Fields.Append "CustomerID", adChar, 10, adFldUpdatable
            .Fields.Append "CustomerName", adChar, 45, adFldUpdatable
            .Fields.Append "ZIP", adChar, 3, adFldUpdatable
            .Fields.Append "Address1", adChar, 45, adFldUpdatable
            .Fields.Append "Address2", adChar, 45, adFldUpdatable
            .Fields.Append "Address3", adChar, 45, adFldUpdatable
            .Fields.Append "CustomerPO", adChar, 12, adFldUpdatable
            .Fields.Append "OrderComments", adChar, 60, adFldUpdatable
            .Fields.Append "OrderLine", adChar, 4, adFldUpdatable
            .Fields.Append "SKU", adChar, 14, adFldUpdatable
            .Fields.Append "SKUDescription", adChar, 60, adFldUpdatable
            .Fields.Append "AllocateQTY", adChar, 11, adFldUpdatable
            .Fields.Append "Ship QTY", adChar, 11, adFldUpdatable
            .Fields.Append "Weight", adChar, 11, adFldUpdatable
            .Fields.Append "MSR", adChar, 11, adFldUpdatable
            .Fields.Append "AssignedFlag", adChar, 1, adFldUpdatable
            .Fields.Append "AssignedDate", adChar, 8, adFldUpdatable
            .Fields.Append "HI", adChar, 5, adFldUpdatable
            .Fields.Append "TI", adChar, 5, adFldUpdatable
            .CursorType = adOpenKeyset
            .LockType = adLockOptimistic
            .Open
        End With
        ' 給予起始值。
        intPointer = 1
        Open strTranFileName For Input As #1
        Do Until EOF(1)
            Line Input #1, strRow
            With rs_Src
                .AddNew
                .Fields(0) = Trim(GetWord(strRow, intPointer, 12))
                .Fields(1) = Trim(GetWord(strRow, intPointer, 1))
                .Fields(2) = Trim(GetWord(strRow, intPointer, 3))
                .Fields(3) = Trim(GetWord(strRow, intPointer, 8))
                .Fields(4) = Trim(GetWord(strRow, intPointer, 8))
                .Fields(5) = Trim(GetWord(strRow, intPointer, 10))
                .Fields(6) = Trim(GetWord(strRow, intPointer, 45))
                .Fields(7) = Trim(GetWord(strRow, intPointer, 3))
                .Fields(8) = Trim(GetWord(strRow, intPointer, 45))
                .Fields(9) = Trim(GetWord(strRow, intPointer, 45))
                .Fields(10) = Trim(GetWord(strRow, intPointer, 45))
                .Fields(11) = Trim(GetWord(strRow, intPointer, 12))
                .Fields(12) = Trim(GetWord(strRow, intPointer, 60))
                .Fields(13) = Trim(GetWord(strRow, intPointer, 4))
                .Fields(14) = Trim(GetWord(strRow, intPointer, 14))
                .Fields(15) = Trim(GetWord(strRow, intPointer, 60))
                .Fields(16) = Trim(GetWord(strRow, intPointer, 11))
                .Fields(17) = Trim(GetWord(strRow, intPointer, 11))
                .Fields(18) = Trim(GetWord(strRow, intPointer, 11))
                .Fields(19) = Trim(GetWord(strRow, intPointer, 11))
                .Fields(20) = Trim(GetWord(strRow, intPointer, 1))
                .Fields(21) = Trim(GetWord(strRow, intPointer, 8))
                .Fields(22) = Trim(GetWord(strRow, intPointer, 5))
                .Fields(23) = Trim(GetWord(strRow, intPointer, 5))
                .Update
                .MoveFirst
            End With
            intPointer = 1
        Loop
        Close
        rs_Src.MoveFirst
        Set dg_CustInv.DataSource = rs_Src
        With dg_CustInv
              .Columns(15).Alignment = dbgRight
              .Columns(16).Alignment = dbgRight
              .Columns(17).Alignment = dbgRight
              .Columns(18).Alignment = dbgRight
         End With
         
        rs_Src.MoveFirst
        strOrderNo = ""
        int_Repeat = 0
        int_Order = 0
        int_OrderLine = 0
        Do While Not rs_Src.EOF
           If strOrderNo <> rs_Src.Fields("OrderNo").Value Then '資料檢驗--判斷訂單編號已訂是否要在 [訂單主檔] 中新增一筆
                strOrderNo = rs_Src.Fields("OrderNo").Value
                '資料檢驗--判斷訂單是否重複
                Call Confirm_Recordset_Closed(tmp_Rs)
                str_SQL = "select ExternOrderKey from Logictown.dbo.orders where ExternOrderKey='" & Trim(rs_Src.Fields(0)) & "' "
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                If tmp_Rs.EOF Then
                    '訂單主檔新增一筆
                    str_SQL = "select isnull(max(orderkey),0) from Logictown.dbo.orders"
                    Call Confirm_Recordset_Closed(tmp_Rs)
                    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                    str_Orderkey = StrPadLeft(Val(tmp_Rs.Fields(0)) + 1, 10, 0)
                    str_CustomerID = Trim(rs_Src.Fields(5).Value)
                    If Left(rs_Src.Fields(5).Value, 3) = "GO" Or Left(rs_Src.Fields(5).Value, 3) = "SO" Or Left(rs_Src.Fields(5).Value, 3) = "NO" Or Len(rs_Src.Fields(5).Value) = 0 Then
                        str_SQL = "select isnull(max(ConsigneeKey),0) from Logictown.dbo.orders where left(ConsigneeKey,3)='UTL'"
                        Call Confirm_Recordset_Closed(tmp_Rs)
                        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                        str_CustomerID = "UTL" & StrPadLeft(Val(Right(tmp_Rs.Fields(0), 7)) + 1, 7, 0)
                        tmp_Rs.Close
                    End If
                    'tmp_rs.Close
                    cn.BeginTrans
                        'OrderNo,OrderType,Division,OrderDate,DeliveryDate,CustomerID,CustomerName,ZIP,Address1,Address2,Address3,CustomerPO,OrderComments
                        'OrderLine ,SKU,SKUDescription,AllocateQTY,Ship QTY,Weight,MSR,AssignedFlag,AssignedDate,HI,TI,
                        
                        'OrderKey,StorerKey,ExternOrderKey,OrderDate,DeliveryDate,Priority,ConsigneeKey,C_contact1,C_Contact2,C_Company,C_Address1,C_Address2,C_Address3,C_Address4,C_City,
                        'C_State,C_Zip,C_Country,C_ISOCntryCode,C_Phone1,C_Phone2,C_Fax1,C_Fax2,C_vat,BuyerPO, BillToKey, B_contact1, B_Contact2, B_Company, B_Address1, B_Address2, B_Address3, B_Address4, B_City, B_State, B_Zip, B_Country,
                        'B_ISOCntryCode, B_Phone1, B_Phone2, B_Fax1, B_Fax2, B_Vat, IncoTerm, PmtTerm, OpenQty, Status, DischargePlace, DeliveryPlace, IntermodalVehicle, CountryOfOrigin, CountryDestination, UpdateSource, Type, OrderGroup, Door,
                        'Route, Stop, Notes, EffectiveDate, AddDate, AddWho, EditDate, EditWho, TrafficCop, ArchiveCop, ContainerType, ContainerQty, BilledContainerQty, DoRoute, CustomerOrderkey, PIVNO, Amount
                        str_SQL = "INSERT Logictown.dbo.orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,DeliveryDate,ConsigneeKey,C_Company,C_Zip,C_Address1,C_Address2,C_Address3,CustomerOrderkey,Notes,UpdateSource) " & _
                            "VALUES ('" & str_Orderkey & "','" & Trim(rs_Src.Fields(0)) & "','" & Trim(rs_Src.Fields(1)) & "','" & Trim(rs_Src.Fields(2)) & "','" & Trim(rs_Src.Fields(3)) & "','" & Trim(rs_Src.Fields(4)) & "', " & _
                            "'" & str_CustomerID & "','" & Trim(rs_Src.Fields(6)) & "','" & Trim(rs_Src.Fields(7)) & "','" & Trim(rs_Src.Fields(8)) & "','" & Trim(rs_Src.Fields(9)) & "'," & _
                            "'" & Trim(rs_Src.Fields(10)) & "','" & Trim(rs_Src.Fields(11)) & "','" & Trim(rs_Src.Fields(12)) & "','" & Trim(filLocalFile.FileName) & "')"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        int_Order = int_Order + 1
                    cn.CommitTrans
                End If
                tmp_Rs.Close
           End If
           '資料檢驗--判斷訂單明細是否重複
           Call Confirm_Recordset_Closed(tmp_Rs)
           str_SQL = "select ExternOrderKey from Logictown.dbo.ORDERDETAIL where ExternOrderKey='" & Trim(rs_Src.Fields(0)) & "' and OrderLineNumber= '" & Trim(rs_Src.Fields(13)) & "'"
           tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
           If Not tmp_Rs.EOF Then
                int_Repeat = int_Repeat + 1
                GoTo nextloop
           End If
           '訂單明細資料新增
           cn.BeginTrans
                'OrderNo,OrderType,Division,OrderDate,DeliveryDate,CustomerID,CustomerName,ZIP,Address1,Address2,Address3,CustomerPO,OrderComments
                'OrderLine ,SKU,SKUDescription,AllocateQTY,Ship QTY,Weight,MSR,AssignedFlag,AssignedDate,HI,TI,
                
                'OrderKey, OrderLineNumber, OrderDetailSysId, ExternOrderKey, ExternLineNo, Sku, StorerKey, ManufacturerSku, RetailSku, AltSku, OriginalQty, OpenQty, ShippedQty, AdjustedQty, QtyPreAllocated, QtyAllocated, QtyPicked, UOM, PackKey, PickCode, CartonGroup, Lot, ID, Facility, Status, UnitPrice, Tax01, Tax02, ExtendedPrice, UpdateSource, Lottable01, Lottable02, Lottable03, Lottable04, Lottable05, EffectiveDate, AddDate, AddWho, EditDate, EditWho, TrafficCop, ArchiveCop, TariffKey, Lottable06, Lottable07, Lottable08, Lottable09, Lottable10, Lottable11, Beginqty
            str_SQL = "INSERT Logictown.dbo.ORDERDETAIL (OrderKey,OrderLineNumber, ExternOrderKey,Sku,StorerKey,OriginalQty,Lottable01, Lottable02,PackKey)" & _
                     "VALUES ('" & str_Orderkey & "','" & Trim(rs_Src.Fields(13)) & "','" & Trim(rs_Src.Fields(0)) & "','" & Trim(rs_Src.Fields(14)) & "','" & Trim(rs_Src.Fields(2)) & "', " & _
                     "'" & Trim(rs_Src.Fields(16)) / 1000 & "','" & Trim(rs_Src.Fields(20)) & "','" & Trim(rs_Src.Fields(21)) & "','" & Trim(rs_Src.Fields(14)) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            cn.CommitTrans
           '資料檢驗--判斷SKU否存在
            str_SQL = "select sku from Logictown.dbo.sku where sku='" & Trim(rs_Src.Fields(14)) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            If tmp_Rs.EOF Then
                cn.BeginTrans
                    'StorerKey, Sku, DESCR, SUSR1, SUSR2, SUSR3, SUSR4, SUSR5, MANUFACTURERSKU, RETAILSKU, ALTSKU, PACKKey, STDGROSSWGT, STDNETWGT, STDCUBE, TARE, CLASS, ACTIVE, SKUGROUP, Tariffkey, BUSR1, BUSR2, BUSR3, BUSR4, BUSR5, LOTTABLE01LABEL, LOTTABLE02LABEL, LOTTABLE03LABEL, LOTTABLE04LABEL, LOTTABLE05LABEL, NOTES1, NOTES2, PickCode, StrategyKey, CartonGroup, PutCode, PutawayLoc, PutawayZone, InnerPack, Cube, GrossWgt, NetWgt, ABC, CycleCountFrequency, LastCycleCount, ReorderPoint, ReorderQty, StdOrderCost, CarryCost, Price, Cost, ReceiptHoldCode, ReceiptInspectionLoc, TrafficCop, ArchiveCop, IOFlag, TareWeight, LotxIdDetailOtherlabel1, LotxIdDetailOtherlabel2, LotxIdDetailOtherlabel3, AvgCaseWeight, TolerancePct, SkuRotat01, SkuRotat02, SkuRotat03, SkuRotat04, SkuRotat05, DefaultRotation, AllocParm, onreceiptcopypackkey, replenishmentQty
                    str_CSKU = StringCleaner(Trim(rs_Src.Fields(15)), "'")
                    str_SQL = "INSERT Logictown.dbo.sku (StorerKey, Sku, DESCR,PACKKey,STDGROSSWGT,BUSR4) " & _
                        "VALUES ('" & Trim(rs_Src.Fields(2)) & "','" & Trim(rs_Src.Fields(14)) & "','" & str_CSKU & "','" & Trim(rs_Src.Fields(14)) & "','" & Trim(rs_Src.Fields(18)) / 10000 & "','" & Trim(rs_Src.Fields(19)) / 10000 & "' )"
                    cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
                    'PackKey, PackDescr, PackUOM1, CaseCnt, ISWHQty1, ReplenishUOM1, ReplenishZone1, CartonizeUOM1, LengthUOM1, WidthUOM1, HeightUOM1, CubeUOM1, PackUOM2, InnerPack, ISWHQty2, ReplenishUOM2, ReplenishZone2, CartonizeUOM2, LengthUOM2, WidthUOM2, HeightUOM2, CubeUOM2, PackUOM3, Qty, ISWHQty3, ReplenishUOM3, ReplenishZone3, CartonizeUOM3, LengthUOM3, WidthUOM3, HeightUOM3, CubeUOM3, PackUOM4, Pallet, ISWHQty4, ReplenishUOM4, ReplenishZone4, CartonizeUOM4, LengthUOM4, WidthUOM4, HeightUOM4, CubeUOM4, PalletWoodLength, PalletWoodWidth, PalletWoodHeight, PalletTI, PalletHI, PackUOM5, Cube, ISWHQty5, PackUOM6, GrossWgt, ISWHQty6, PackUOM7, NetWgt, ISWHQty7, PackUOM8, OtherUnit1, ISWHQty8, ReplenishUOM8, ReplenishZone8, CartonizeUOM8, LengthUOM8, WidthUOM8, HeightUOM8, PackUOM9, OtherUnit2, ISWHQty9, ReplenishUOM9, ReplenishZone9, CartonizeUOM9, LengthUOM9, WidthUOM9, HeightUOM9, AddDate, AddWho, EditDate, EditWho, TrafficCop, ArchiveCop, Timestamp
                    str_SQL = "INSERT Logictown.dbo.pack (PACKKey,PackDescr,GrossWgt, PalletHI, PalletTI,CaseCnt) " & _
                        "VALUES ('" & Trim(rs_Src.Fields(14)) & "','" & str_CSKU & "','" & Trim(rs_Src.Fields(18)) / 10000 & "', " & _
                         "'" & Trim(rs_Src.Fields(22)) & "','" & Trim(rs_Src.Fields(23)) & "','1')"
                    cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
                    
                cn.CommitTrans
            Else
                cn.BeginTrans
                    str_CSKU = StringCleaner(Trim(rs_Src.Fields(15)), "'")
                    str_SQL = "UPDATE Logictown.dbo.sku set StorerKey='" & Trim(rs_Src.Fields(2)) & "',  DESCR ='" & str_CSKU & "',PACKKey='" & Trim(rs_Src.Fields(14)) & "' " & _
                        "where Sku='" & Trim(rs_Src.Fields(14)) & "'"
                    cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
                    
                    str_SQL = "UPDATE Logictown.dbo.pack set PackDescr='" & str_CSKU & "',GrossWgt='" & Trim(rs_Src.Fields(18)) / 10000 & "',CubeUOM1='" & Trim(rs_Src.Fields(19)) / 10000 & "' " & _
                        ", PalletHI='" & Trim(rs_Src.Fields(22)) & "', PalletTI='" & Trim(rs_Src.Fields(23)) & "',CaseCnt='1' " & _
                        "where PACKKey='" & Trim(rs_Src.Fields(14)) & "'"
                    cn.Execute (str_SQL), RowsAffect, adExecuteNoRecords
                    
                cn.CommitTrans
            End If
            tmp_Rs.Close
            int_OrderLine = int_OrderLine + 1
nextloop:
            rs_Src.MoveNext
        Loop
        '備份檔案
        Dim fl_file As Scripting.File

        Set fso = New FileSystemObject
        If fso.FileExists(strTranFileName) = True Then
            Set fl_file = fso.GetFile(strTranFileName)
            fl_file.copy ("C:\from_ids\backup\Alc\" & str_file)
            If fso.FileExists("C:\from_ids\backup\Alc\" & str_file) = True Then
                fl_file.Delete
            End If
        End If
        filLocalFile.Refresh
        '刪出ftp上之檔案
        ITC.Execute , "DELETE " & Chr(34) & str_file & Chr(34)
        Do Until ITCReady(False)
            DoEvents: DoEvents: DoEvents: DoEvents
        Loop
        lstRemoteFile.Clear
        ITC.Execute , "DIR"
        lblStatus = "已連線置alc資料夾"
        
        If int_Repeat > 0 Then
            msg_text = "有" & int_Repeat & "筆訂單明細重複轉檔"
            MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        End If
        msg_text = "匯入" & int_Order & "筆訂單 " & int_OrderLine & "筆明細,文字檔備份於C:\from_ids\backup\Alc\"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        
    End If
End Sub

'Private Sub imgReceiveFile3_Click()
'Call Command3_Click
''Call cmdImport3_Click
'End Sub

Private Sub imgSendFile_Click()
Dim i As Double
'If the ITC is not still executing then send the file
If ITCReady(True) = True Then
    'Check that a file has been selected
    If Trim(filLocalFile.FileName) = "" Then
        MsgBox "請點選要上傳的檔案", vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If
    
    'Check that the file does not already exist on the server
    For i = 0 To lstRemoteFile.ListCount
        If filLocalFile.FileName = lstRemoteFile.List(i) Then
            If MsgBox("檔案 " & filLocalFile.FileName & " 已經存在" & vbCrLf & "要覆蓋嗎?", vbQuestion + vbYesNo, "Overwrite") = vbNo Then
                Exit Sub
            End If
        End If
    Next i
           
    'Send the file and update the remote file list box
    ITC.Execute , "PUT " & Chr(34) & filLocalFile.Path & "\" & filLocalFile.FileName & Chr(34) & " " & Chr(34) & filLocalFile.FileName & Chr(34)
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    lstRemoteFile.Clear
    ITC.Execute , "DIR"
    lblStatus = "已連線"
End If
End Sub

Private Sub ITC_StateChanged(ByVal State As Integer)
    'Check the state of the itc, and change the status accordingly
    Dim Data1, RemoteFiles, i As Double
    Dim RemoteFileName As String
    Select Case State
        Case icResolvingHost
            lblStatus = "Finding Host IP Address"
        Case icHostResolved
            lblStatus = "IP Address Found"
        Case icConnecting
            lblStatus = "Connecting To Host"
        Case icConnected
            lblStatus = "已連線"
        Case icRequesting
            lblStatus = "Sending Request"
        Case icRequestSent
            lblStatus = "Request Sent"
        Case icReceivingResponse
            lblStatus = "Receiving Response"
        Case icResponseReceived
            lblStatus = "Response Receiving"
        Case icDisconnecting
            lblStatus = "Disconnecting"
        Case icDisconnected
            lblStatus = "Not Connected"
        Case icError
            If ITC.ResponseCode = 12030 Then
                lblStatus = "Not Connected"
                cmdLogOn.Enabled = False
                cmdNewFolder.Enabled = False
                cmdDelete.Enabled = False
                'cmdRename.Enabled = False
                'cmdSize.Enabled = False
                cmdUpFolder.Enabled = False
                imgSendFile.Enabled = False
                imgReceiveFile.Enabled = False
                lstRemoteFile.Enabled = False
                cmdLogOff.Enabled = False
                cmdLogOn.Enabled = True
                ITC.Cancel
            End If
            If ITC.ResponseCode <> 87 Then
                MsgBox ITC.ResponseCode & " " & ITC.ResponseInfo, vbOKOnly + vbCritical, "Error"
            End If
        Case icResponseCompleted
            'loop until you get all data
            Do While True
                Data1 = ITC.GetChunk(4096, icString)
                If Len(Data1) = 0 Then Exit Do
                DoEvents
                RemoteFiles = RemoteFiles & Data1
            Loop
            'Beep
            'If it is recieving size data tell the user the size and then exit the sub
            If RecievingSize Then
                MsgBox "The size of file " & lstRemoteFile.Text & " is " & RemoteFiles & " bytes", vbInformation + vbOKOnly, "Size"
                Exit Sub
            End If
            'Loop through, check for carriage returns to get each file name and add to listbox
            For i = 1 To Len(RemoteFiles)
                If Mid(RemoteFiles, i, 1) = Chr(13) Then
                    If Trim(RemoteFileName) <> "" Then
                        lstRemoteFile.AddItem RemoteFileName
                        RemoteFileName = ""
                    End If
                Else
                    If Mid(RemoteFiles, i, 1) <> Chr(10) Then
                        RemoteFileName = RemoteFileName & Mid(RemoteFiles, i, 1)
                    End If
                End If
            Next i
    End Select
End Sub

'Private Sub int3_StateChanged(ByVal State As Integer)
'    'Check the state of the itc, and change the status accordingly
'    Dim Data1, RemoteFiles
'    Dim RemoteFileName As String
'    Select Case State
'        Case icResolvingHost
'            lblStatus3 = "Finding Host IP Address"
'        Case icHostResolved
'            lblStatus3 = "IP Address Found"
'        Case icConnecting
'            lblStatus3 = "Connecting To Host"
'        Case icConnected
'            lblStatus3 = "已連線"
'        Case icRequesting
'            lblStatus3 = "Sending Request"
'        Case icRequestSent
'            lblStatus3 = "Request Sent"
'        Case icReceivingResponse
'            lblStatus3 = "Receiving Response"
'        Case icResponseReceived
'            lblStatus3 = "Response Receiving"
'        Case icDisconnecting
'            lblStatus3 = "Disconnecting"
'        Case icDisconnected
'            lblStatus3 = "Not Connected"
'        Case icError
'            If int3.ResponseCode = 12030 Then
'                lblStatus3 = "Not Connected"
'                cmdLogon3.Enabled = False
''                cmdNewFolder.Enabled = False
''                cmdDelete.Enabled = False
'                'cmdRename.Enabled = False
'                'cmdSize.Enabled = False
''                cmdUpFolder.Enabled = False
''                imgSendFile.Enabled = False
'                imgReceiveFile3.Enabled = False
'                lstRemoteFile3.Enabled = False
'                cmdLogOff3.Enabled = False
'                cmdLogon3.Enabled = True
'                int3.Cancel
'            End If
'            If int3.ResponseCode <> 87 Then
'                MsgBox int3.ResponseCode & " " & int3.ResponseInfo, vbOKOnly + vbCritical, "Error"
'            End If
'        Case icResponseCompleted
'            'loop until you get all data
'            Do While True
'                Data1 = int3.GetChunk(4096, icString)
'                If Len(Data1) = 0 Then Exit Do
'                DoEvents
'                RemoteFiles = RemoteFiles & Data1
'            Loop
'            'Beep
'            'If it is recieving size data tell the user the size and then exit the sub
'            If RecievingSize Then
'                MsgBox "The size of file " & lstRemoteFile3.Text & " is " & RemoteFiles & " bytes", vbInformation + vbOKOnly, "Size"
'                Exit Sub
'            End If
'            'Loop through, check for carriage returns to get each file name and add to listbox
'            For i = 1 To Len(RemoteFiles)
'                If Mid(RemoteFiles, i, 1) = Chr(13) Then
'                    If Trim(RemoteFileName) <> "" Then
'                        lstRemoteFile3.AddItem RemoteFileName
'                        RemoteFileName = ""
'                    End If
'                Else
'                    If Mid(RemoteFiles, i, 1) <> Chr(10) Then
'                        RemoteFileName = RemoteFileName & Mid(RemoteFiles, i, 1)
'                    End If
'                End If
'            Next i
'    End Select
'End Sub

Public Function GetWord(ByVal strData As String, ByRef intStart As Integer, ByVal intLen As Integer) As String
    Dim intloop, z As Integer
    Dim strTemp As String
    '字串不足時補空白
    If Len(strData) < intStart + intLen Then
        strData = Left(strData & String(intStart + intLen, " "), intStart + intLen)
    End If
    
    intloop = 0
    z = Len(strData)
    Do While intloop <= intLen - 1
        strTemp = Mid(strData, intStart + intloop, 1)
        If intloop = intLen - 1 Then        '判斷最後一碼是否為中文,因為當字串前面為英文時切割後可能會多一格
            If Asc(strTemp) < 0 Then        '如果字元是中文
                intLen = intLen - 1
                GetWord = GetWord & " "     '字串直接加一格空白,不再加中文
            Else
                GetWord = GetWord + strTemp
            End If
            intloop = intloop + 1
        Else
            If Asc(strTemp) < 0 Then
                intLen = intLen - 1                         '如果字元是中文
            End If
            GetWord = GetWord + strTemp
            intloop = intloop + 1
        End If
            
    Loop
    intStart = intStart + intloop
End Function

Public Function StringCleaner(S As String, Search As String) As String
    Dim i As Integer, res As String
    res = S
    Do While InStr(res, Search)
        i = InStr(res, Search)
        res = Left(res, i - 1) & Mid(res, i + 1)
    Loop
    StringCleaner = res
End Function

Private Sub FTPlog(ByVal strActionName As String)
    '產生程式執行記錄
    Dim fso As Scripting.FileSystemObject
    Dim ts_LogFile As Scripting.TextStream
    Dim strTmp As String
    
    If Dir(App.Path & "\Log", vbDirectory) = "" Then MkDirs App.Path & "\Log"
    
    '取得程式執行記錄檔，若不存在則自動新增
    Set fso = New Scripting.FileSystemObject
    If fso.FileExists(App.Path & "\Log\Import.log") Then
       Set ts_LogFile = fso.OpenTextFile(App.Path & "\log\Import.log", ForAppending)  'open TextStream Object
    Else
       Set ts_LogFile = fso.CreateTextFile(App.Path & "\log\Import.log", True)       'create TextStream Object
    End If
    '寫入狀態值
    strTmp = Format(Now, "yyyy-mm-dd ttttt") & "，" & strActionName & " 匯入者 : " & User_id
    
    ts_LogFile.WriteLine (strTmp)
    ts_LogFile.WriteBlankLines (1)
    ts_LogFile.Close
    Set ts_LogFile = Nothing
    Set fso = Nothing
End Sub

Private Sub cmd_Alc_Click()
    lstRemoteFile.Clear
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    ITC.Execute , "CD " & Chr(34) & "/Bestg/Alc" & Chr(34)
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
        
    ITC.Execute , "DIR"
    'ITC.Execute , "LS"
    Do While ITC.StillExecuting
        DoEvents: DoEvents: DoEvents
    Loop
    lblStatus = "已連線至alc資料夾"
End Sub

Private Sub cmd_shp_Click()
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    lstRemoteFile.Clear
    ITC.Execute , "CD " & Chr(34) & "/Bestg/shp" & Chr(34)
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
        
    ITC.Execute , "DIR"
    'ITC.Execute , "LS"
    Do While ITC.StillExecuting
        DoEvents: DoEvents: DoEvents
    Loop
    lblStatus = "已連線至shp資料夾"
End Sub

Private Sub cmd_CFM_Click()
    lstRemoteFile.Clear
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    ITC.Execute , "CD " & Chr(34) & "/Bestg/CFM" & Chr(34)
    Do Until ITCReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
        
    ITC.Execute , "DIR"
    'ITC.Execute , "LS"
    Do While ITC.StillExecuting
        DoEvents: DoEvents: DoEvents
    Loop
    lblStatus = "已連線至cfm資料夾"
End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "FTP上下傳"
End Sub

Private Sub Form_Load()
    '設定 Form 大小、位置
    dbsrcFormHeight = 7140
    dbsrcFormWidth = 11475
    Me.Height = 7650: Me.Width = 11600
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Left = 200
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300
    Me.txtServer = "10.190.1.2"
    'Me.txtServer = "202.145.173.225"
    Me.txtUserName = "NLJL"
    Me.txtPassword = "NLJL"
    
'    txtServer3 = "chutney.unilever.com"
'    chutney.unilever.com
'    txtUserName3 = "jean.chang"
'    txtPassWord3 = "re7ated"
    
    '如果是系統管理員，則開啟立邦
    If UCase(User_id) = "ADMINISTRATOR" Then
        SSTab1.Tab = 5: SSTab1.Caption = "立邦退貨訂單"
        SSTab1.Tab = 11: SSTab1.Caption = "立邦訂單匯入"
    End If
    SSTab1.Tab = 18
    
    If Dir("C:\LTKK01\Orders", vbDirectory) = "" Then MkDirs "C:\LTKK01\Orders"
    If Dir("C:\BEST\LVTL01\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\LVTL01\Orders"
    If Dir("C:\BEST\LFYY01\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\LFYY01\Orders"
    If Dir("C:\BEST\LKAO01\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\LKAO01\Orders"
    If Dir("C:\BEST\LSJR01\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\LSJR01\Orders"
    If Dir("C:\BEST\LNSL01\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\Orders"
    If Dir("C:\BEST\LNIP01\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\LNIP01\Orders"
    If Dir("C:\BEST\Other\Orders", vbDirectory) = "" Then MkDirs "C:\BEST\Other\Orders"
       
    drvLocalDrive.Drive = "C:": dirLocalDir.Path = "C:\LTKK01\Orders"
    drvLocalDriveT2.Drive = "C:": dirLocalDirT2.Path = "C:\BEST\LVTL01\Orders"
    drvLocalDriveT3.Drive = "C:": dirLocalDirT3.Path = "C:\BEST\LFYY01\Orders"
    drvLocalDriveT4.Drive = "C:": dirLocalDirT4.Path = "C:\BEST\LKAO01\Orders"
    drvLocalDriveT5.Drive = "C:": dirLocalDirT5.Path = "C:\BEST\LNIP01\Orders"
    drvLocalDriveT6.Drive = "C:": dirLocalDirT6.Path = "C:\BEST\LSJR01\Orders"
    drvLocalDriveT7.Drive = "C:": dirLocalDirT7.Path = "C:\BEST\LNSL01\Orders"
    drvLocalDriveT8.Drive = "C:": dirLocalDirT8.Path = "C:\BEST\LNSL01\Orders"
    drvLocalDriveT9.Drive = "C:": dirLocalDirT9.Path = "C:\BEST\LNSL01\Orders"
    drvLocalDriveT10.Drive = "C:": dirLocalDirT10.Path = "C:\BEST\LNSL01\Orders"
    drvLocalDriveT11.Drive = "C:": dirLocalDirT11.Path = "C:\BEST\LNIP01\Orders"
    drvLocalDriveT12.Drive = "C:": dirLocalDirT12.Path = "C:\BEST\LKAO01\Orders"
    drvLocalDriveT13.Drive = "C:": dirLocalDirT13.Path = "C:\BEST\Other\Orders"
    drvLocalDriveT14.Drive = "C:": dirLocalDirT14.Path = "C:\BEST\LNSL01\Orders"
'    drvLocalDriveT15.Drive = "C:": dirLocalDirT15.Path = "C:\BEST\LMYS01\Orders"
'    drvLocalDriveT16.Drive = "C:": dirLocalDirT16.Path = "C:\BEST\LMYS01\Orders"
'    drvLocalDriveT17.Drive = "C:": dirLocalDirT17.Path = "C:\BEST\LAPP01\Orders"
'    drvLocalDriveT18.Drive = "C:": dirLocalDirT18.Path = "C:\BEST\LAPP01\Orders"
    '取出所有貨主資料--TRP16M
    Dim tmp_cnt As Integer
    cboStorerkeyT13.Clear
    str_SQL = "Select Rtrim(StorerKey) as 'StorerKey',Isnull(Rtrim(Short_Name),'') as 'StorerName' From TRP16M where storer_status > 0 Order by StorerKey"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
    Do While Not tmp_Rs.EOF
       cboStorerkeyT13.AddItem tmp_Rs.Fields("StorerKey")
    tmp_Rs.MoveNext
    Loop
    tmp_Rs.Close
    
    cboStorerkeyT13 = "LHPT01"

    RecievingSize = False
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '最小化

If Me.ScaleHeight > SSTab1.Top + Frame1.Top + Frame1.Height + 500 Then
    SSTab1.Height = (Me.ScaleHeight)
    dg_CustInv.Height = SSTab1.Height - SSTab1.Top - 1200
    dgMainT2.Height = SSTab1.Height - Frame1.Top - Frame1.Height - 240
    dgMainT3.Height = SSTab1.Height - Frame2.Top - Frame2.Height - 240
    dgMainT4.Height = SSTab1.Height - Frame7.Top - Frame7.Height - 240
    dgMainT5.Height = SSTab1.Height - Frame8.Top - Frame8.Height - 240
    dgMainT6.Height = SSTab1.Height - Frame9.Top - Frame9.Height - 240
    dgMainT7.Height = SSTab1.Height - Frame10.Top - Frame10.Height - 240
    dgMainT8.Height = SSTab1.Height - Frame11.Top - Frame11.Height - 240
    dgMainT9.Height = SSTab1.Height - Frame12.Top - Frame12.Height - 240
    dgMainT10.Height = SSTab1.Height - Frame13.Top - Frame13.Height - 240
    dgMainT11.Height = SSTab1.Height - Frame3.Top - Frame3.Height - 240
    dgMainT12.Height = SSTab1.Height - Frame4.Top - Frame4.Height - 240
    dgMainT13.Height = SSTab1.Height - Frame5.Top - Frame5.Height - 240
    dgMainT14.Height = SSTab1.Height - Frame6.Top - Frame6.Height - 240
    dgMainT15.Height = SSTab1.Height - Frame14.Top - Frame14.Height - 240
    dgMainT16.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 240
    dgMainT16_1.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 240
    dgMainT16_2.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 240
    dgMainT16_3.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 240
    dgMainT17.Height = SSTab1.Height - Frame16.Top - Frame16.Height - 240
    dgMainT17_1.Height = SSTab1.Height - Frame16.Top - Frame16.Height - 240
    dgMainT18.Height = SSTab1.Height - Frame17.Top - Frame17.Height - 240
    dgMainT19.Height = SSTab1.Height - Frame18.Top - Frame18.Height - 240
    dgMainT20.Height = SSTab1.Height - Frame19.Top - Frame19.Height - 240
    dgMainT21.Height = SSTab1.Height - Frame20.Top - Frame18.Height - 240
    dgMainT22.Height = SSTab1.Height - Frame21.Top - Frame19.Height - 240
End If

If Me.ScaleHeight > SSTab2.Top + Frame1.Top + Frame1.Height + 500 Then
    SSTab2.Height = (Me.ScaleHeight)
    dgMainT15.Height = SSTab1.Height - Frame14.Top - Frame14.Height - 800
    dgMainT16.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT16_1.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT16_2.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT16_3.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT17.Height = SSTab1.Height - Frame16.Top - Frame16.Height - 800
    dgMainT17_1.Height = SSTab1.Height - Frame16.Top - Frame16.Height - 800
    dgMainT18.Height = SSTab1.Height - Frame17.Top - Frame17.Height - 800
    dgMainT19.Height = SSTab1.Height - Frame18.Top - Frame18.Height - 800
    dgMainT20.Height = SSTab1.Height - Frame19.Top - Frame19.Height - 800
    dgMainT21.Height = SSTab1.Height - Frame20.Top - Frame18.Height - 800
    dgMainT22.Height = SSTab1.Height - Frame21.Top - Frame19.Height - 800
End If

If Me.ScaleHeight > SSTab3.Top + Frame1.Top + Frame1.Height + 500 Then
    SSTab3.Height = (Me.ScaleHeight)
    dgMainT15.Height = SSTab1.Height - Frame14.Top - Frame14.Height - 800
    dgMainT16.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT16_1.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT16_2.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT16_3.Height = SSTab1.Height - Frame15.Top - Frame15.Height - 800
    dgMainT17.Height = SSTab1.Height - Frame16.Top - Frame16.Height - 800
    dgMainT17_1.Height = SSTab1.Height - Frame16.Top - Frame16.Height - 800
    dgMainT18.Height = SSTab1.Height - Frame17.Top - Frame17.Height - 800
    dgMainT19.Height = SSTab1.Height - Frame18.Top - Frame18.Height - 800
    dgMainT20.Height = SSTab1.Height - Frame19.Top - Frame19.Height - 800
    dgMainT21.Height = SSTab1.Height - Frame20.Top - Frame20.Height - 800
    dgMainT22.Height = SSTab1.Height - Frame21.Top - Frame21.Height - 800
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    SSTab1.Width = (Me.ScaleWidth - 240)
    dg_CustInv.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT2.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT3.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT4.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT5.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT6.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT7.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT8.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT9.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT10.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT11.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT12.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT13.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT14.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT15.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT16.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT16_1.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT16_2.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT16_3.Width = SSTab1.Width - SSTab1.Left - 400
    dgMainT17.Width = SSTab1.Width - SSTab1.Left - 600
    dgMainT17_1.Width = SSTab1.Width - SSTab1.Left - 600
    dgMainT18.Width = SSTab1.Width - SSTab1.Left - 600
    dgMainT19.Width = SSTab1.Width - SSTab1.Left - 600
    dgMainT20.Width = SSTab1.Width - SSTab1.Left - 600
    dgMainT21.Width = SSTab1.Width - SSTab1.Left - 600
    dgMainT22.Width = SSTab1.Width - SSTab1.Left - 600
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    SSTab2.Width = (Me.ScaleWidth - 850)
    dgMainT15.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16_1.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16_2.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16_3.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT17.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT17_1.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT18.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT19.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT20.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT21.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT22.Width = SSTab1.Width - SSTab1.Left - 750
    
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    SSTab3.Width = (Me.ScaleWidth - 850)
    dgMainT15.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16_1.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16_2.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT16_3.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT17.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT17_1.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT18.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT19.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT20.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT21.Width = SSTab1.Width - SSTab1.Left - 750
    dgMainT22.Width = SSTab1.Width - SSTab1.Left - 750
End If
End Sub

Private Sub Form_Terminate()
'更新 Menu [視窗]→[已開視窗清單]
ITC.Cancel
'int3.Cancel
'Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Cancel any tasks that the itc is doing
ITC.Cancel
'int3.Cancel
'從記憶體中移除表單，藉此引起 [Terminate] 事件
Set frm_FTP = Nothing
Set rsMainT2 = Nothing
Set rsMainT3 = Nothing
Set rsMainT4 = Nothing
Set rsMainT5 = Nothing
Set rsMainT6 = Nothing
Set rsMainT7 = Nothing
Set rsMainT8 = Nothing
Set rsMainT9 = Nothing
Set rsMainT10 = Nothing
Set rsMainT11 = Nothing
Set rsMainT12 = Nothing
Set rsMainT13 = Nothing
Set rsMainT14 = Nothing
Set rsMainT15 = Nothing
Set rsMainT16 = Nothing
Set rsMainT16_1 = Nothing
Set rsMainT16_2 = Nothing
Set rsMainT16_3 = Nothing
Set rsMainT17 = Nothing
Set rsMainT18 = Nothing
Set rsMainT19 = Nothing
Set rsMainT20 = Nothing
Set rsMainT21 = Nothing
Set rsMainT22 = Nothing
End Sub

Private Function ITCReady(ShowMessage As Boolean)
'Check the state of itc, if it is not executing return true
If ITC.StillExecuting Then
    ITCReady = False
    If ShowMessage Then
        MsgBox "請稍等.  FTP伺服器執行中", vbInformation + vbOKOnly, "忙碌中"
    End If
Else
    ITCReady = True
End If
End Function

'Private Function int3Ready(ShowMessage As Boolean)
''Check the state of itc, if it is not executing return true
'If int3.StillExecuting Then
'    int3Ready = False
'    If ShowMessage Then
'        MsgBox "請稍等.  FTP伺服器執行中", vbInformation + vbOKOnly, "忙碌中"
'    End If
'Else
'    int3Ready = True
'End If
'End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
If Left(Trim(SSTab1.Caption), 2) = "--" Then SSTab1.Tab = PreviousTab
filLocalFile.Refresh
dirLocalDirT3.Refresh
If Len(Trim(SSTab1.Caption)) = 0 Then SSTab1.Tab = PreviousTab

End Sub
Private Sub cmdImportT2_Click()
'會紀錄到訂單明細，因為VTL的出貨丹會顯示明細，其他貨主不會

strTranFileName = filLocalFileT2.Path & "\" & filLocalFileT2.FileName
If Len(Trim(filLocalFileT2.FileName)) = 0 Then Exit Sub
On Error GoTo err_Handle

Dim strLineTmp As String, i As Integer, j As Integer, k As Integer, arrLen, blDuplicationOrder As Boolean, Intcheck As Integer, Str_check As String
cmdImportT2.Enabled = False: dgMainT2.Enabled = False: Screen.MousePointer = 11
Intcheck = 0
Str_check = ""

Set rsMainT2 = Nothing

    arrLen = Array(11, 9, 9, 9, 9, 16, 37, 6, 9, 9, 9, 2, 11, 3, 3, 12, 31, 3, 9, 17, 255)

    Open filLocalFileT2.Path & "\" & filLocalFileT2.FileName For Input As #1
    
    Set rsMainT2 = New ADODB.Recordset
    With rsMainT2
        .Fields.Append "出貨單號", adChar, arrLen(0), adFldUpdatable
        .Fields.Append "帳款客戶", adChar, arrLen(1), adFldUpdatable
        .Fields.Append "帳款客戶名稱", adChar, arrLen(2), adFldUpdatable
        .Fields.Append "送貨客戶", adChar, arrLen(3), adFldUpdatable
        .Fields.Append "送貨客戶名稱", adChar, arrLen(4), adFldUpdatable
        .Fields.Append "客戶電話", adChar, arrLen(5), adFldUpdatable
        .Fields.Append "客戶地址", adChar, arrLen(6), adFldUpdatable
        .Fields.Append "車商代號", adChar, arrLen(7), adFldUpdatable
        .Fields.Append "車商名稱", adChar, arrLen(8), adFldUpdatable
        .Fields.Append "預出日期", adChar, arrLen(9), adFldUpdatable
        .Fields.Append "排出日期", adChar, arrLen(10), adFldUpdatable
        .Fields.Append "使用棧板", adChar, arrLen(11), adFldUpdatable
        .Fields.Append "噸數", adDouble, arrLen(12), adFldUpdatable
        .Fields.Append "項次", adUnsignedSmallInt, arrLen(13), adFldUpdatable
        .Fields.Append "出貨原因", adChar, arrLen(14), adFldUpdatable
        .Fields.Append "產品編號", adChar, arrLen(15), adFldUpdatable
        .Fields.Append "產品名稱", adChar, arrLen(16), adFldUpdatable
        .Fields.Append "單位", adChar, arrLen(17), adFldUpdatable
        .Fields.Append "數量", adDouble, arrLen(18), adFldUpdatable
        .Fields.Append "客戶單號", adChar, arrLen(19), adFldUpdatable
        .Fields.Append "備註", adChar, arrLen(20), adFldUpdatable
        .Fields.Append "類別", adChar, 10, adFldUpdatable
        .Fields.Append "倉別", adChar, 18, adFldUpdatable
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
  
        '開啟檔案
        k = 0
        Do While Not EOF(1)
                k = k + 1
                Line Input #1, strLineTmp
                If Len(RTrim(strLineTmp)) > 135 Then
                    .AddNew
                    j = 1
                    
                    For i = 0 To rsMainT2.Fields.Count - 3
                        .Fields(i) = RTrim(GetWord(strLineTmp, j, arrLen(i)))
                    Next
                    
                    '訂單類別-是否為入庫
                    rsMainT2("類別") = "I": rsMainT2("倉別") = "R01"
                    If UCase(Trim(rsMainT2("送貨客戶"))) = "DW327" Or UCase(Trim(rsMainT2("送貨客戶"))) = "DW328" Or UCase(Trim(rsMainT2("送貨客戶"))) = "DP1324" Then rsMainT2("類別") = "RC": rsMainT2("倉別") = "R01"
                    If UCase(Trim(rsMainT2("送貨客戶"))) = "DW427" Then rsMainT2("類別") = "RC": rsMainT2("倉別") = "R01-C"
                    If UCase(Trim(rsMainT2("送貨客戶"))) = "DW815" Then rsMainT2("類別") = "RC": rsMainT2("倉別") = "R01-S"
                    
                    '訂單出庫倉別
                    If RTrim(rsMainT2("類別")) = "I" Then
                        If UCase(Trim(rsMainT2("車商代號"))) = "W3270" Or UCase(Trim(rsMainT2("車商代號"))) = "W3280" Or UCase(Trim(rsMainT2("車商代號"))) = "P1324" Or UCase(Trim(rsMainT2("車商代號"))) = "P2324" Then rsMainT2("倉別") = "R01" '北倉出貨
                        If UCase(Trim(rsMainT2("車商代號"))) = "W4270" Then rsMainT2("倉別") = "R01-C" '中倉出貨
                        If UCase(Trim(rsMainT2("車商代號"))) = "W8150" Then rsMainT2("倉別") = "R01-S" '南倉出貨
                    End If
                End If
                
        Loop
            Close #1
        
           .MoveFirst
    
    End With
    rsMainT2.Sort = "出貨單號,車商代號,項次"
    Set dgMainT2.DataSource = rsMainT2
    
    With dgMainT2
    
    For i = 0 To rsMainT2.Fields.Count - 1
    .Columns(i).Caption = rsMainT2.Fields(i).Name
    Next
    
        .ColumnHeaders = True        '標題行顯示
        .RowHeight = 300

    End With
    
    SetDataGridColWidth Me.Caption, dgMainT2
'
'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders(nolock) where storerkey = 'LVTL01' and rtrim(updatesource)='" & filLocalFileT2.FileName & "'"
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: Exit Sub

'訂單資料檢查 add by Eric
rsMainT2.MoveFirst

Do While Not rsMainT2.EOF
    '單位檢查DZ,EA
    If UCase(Trim(rsMainT2("單位"))) <> "DZ" And UCase(Trim(rsMainT2("單位"))) <> "EA" Then
        MsgBox "訂單有EA,DZ以外的單位，請確認檔案格式是否正確。", vbOKOnly, Me.Caption: cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: Exit Sub
    End If
    
    '到貨日期檢查
    If Trim(rsMainT2("預出日期")) < Format(Now, "YYYYMMDD") Then: MsgBox "到貨日小於今日，訂單轉入終止!", 16, Me.Caption: cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: Exit Sub

    '資料檢驗--判斷SKU是否存在
    str_SQL = "select sku from " & strWMSDB & "..sku(nolock) where Storerkey = 'LVTL01' and sku = '" & Trim(rsMainT2("產品編號")) & "'"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        MsgBox "訂單發現新品號 (" & Trim(rsMainT2("產品編號")) & " ) " & Trim(rsMainT2("產品名稱")) & "，訂單轉入終止!!": cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: Exit Sub
    End If

    '資料檢驗--判斷是否屬佰事達訂單
    If UCase(Trim(rsMainT2("送貨客戶"))) = "DW327" Or UCase(Trim(rsMainT2("送貨客戶"))) = "DW328" Or UCase(Trim(rsMainT2("送貨客戶"))) = "DW427" Or UCase(Trim(rsMainT2("送貨客戶"))) = "DW815" Or UCase(Trim(rsMainT2("送貨客戶"))) = "DP1324" Or UCase(Trim(rsMainT2("車商代號"))) = "W3270" Or UCase(Trim(rsMainT2("車商代號"))) = "W3280" Or UCase(Trim(rsMainT2("車商代號"))) = "W4270" Or UCase(Trim(rsMainT2("車商代號"))) = "W8150" Or UCase(Trim(rsMainT2("車商代號"))) = "WA500" Or UCase(Trim(rsMainT2("車商代號"))) = "WB500" Or UCase(Trim(rsMainT2("車商代號"))) = "WD500" Or UCase(Trim(rsMainT2("車商代號"))) = "P1324" Or UCase(Trim(rsMainT2("車商代號"))) = "P2324" Then
    Else
        MsgBox "客戶單號：" & Trim(rsMainT2("出貨單號")) & vbCrLf & "請通知客戶，確認該筆訂單是否有誤!?", vbOKOnly, "發現非佰事達之訂單資料"
        cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: Exit Sub
    End If

    '新客戶檢查1
    str_SQL = "select storerkey from trp01m(nolock) where storerkey = 'LVTL01' and consigneekey = '" & Trim(rsMainT2("帳款客戶")) & "'"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close
        MsgBox "發現新客戶，訂單轉入中止!", vbOKOnly, Me.Caption
        Exit Sub
    End If

    '新客戶檢查2
    str_SQL = "select storerkey from trp01m(nolock) where storerkey = 'LVTL01' and consigneekey = '" & Trim(rsMainT2("送貨客戶")) & "'"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
            cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close
            MsgBox "發現新客戶:" & Trim(rsMainT2("送貨客戶")) & "，訂單轉入中止!", vbOKOnly, Me.Caption
            Exit Sub
    End If
    
    '比對訂單項次
    If Trim(rsMainT2("出貨單號")) <> Str_check Then
        Str_check = Trim(rsMainT2("出貨單號"))
        Intcheck = 1
        If Val(Trim(rsMainT2("項次"))) <> Intcheck Then
                cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0
                MsgBox "發現訂單項次有誤，訂單轉入中止!", vbOKOnly, Me.Caption
                Exit Sub
        End If
        Intcheck = Intcheck + 1
    Else
        If Val(Trim(rsMainT2("項次"))) <> Intcheck Then
                cmdImportT2.Enabled = True: dgMainT2.Enabled = True: Screen.MousePointer = 0
                MsgBox "發現訂單項次有誤，訂單轉入中止!", vbOKOnly, Me.Caption
                Exit Sub
        End If
        Intcheck = Intcheck + 1
    End If
    rsMainT2.MoveNext
Loop

'開始匯入
Tran_Level = cn.BeginTrans
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, intQTY As Long, strFacility As String

rsMainT2.MoveFirst
Do While Not rsMainT2.EOF
DoEvents: DoEvents

'資料檢驗--來源訂單相同單號判斷，不同增加HEAD
If strOrderNo <> UCase(Trim(rsMainT2("出貨單號"))) Then
    strOrderNo = UCase(Trim(rsMainT2("出貨單號")))
    blDuplicationOrder = False

    '資料檢驗--判斷訂單是否重複，重複不增加
    Call Confirm_Recordset_Closed(tmp_Rs)
    str_SQL = "select orderkey from orders where rtrim(ExternOrderKey) ='" & Trim(rsMainT2("出貨單號")) & "' and storerkey = 'LVTL01' and isnull(type,'') <> '刪單' "
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then

        If RTrim(rsMainT2("類別")) = "RC" Then int_Asn = int_Asn + 1
        '取訂單號碼
        str_SQL = "select isnull(max(orderkey),0) from orders"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
        tmp_Rs.Close

        strFacility = UCase(Trim(rsMainT2("車商名稱")))
        If UCase(Trim(rsMainT2("車商代號"))) = "W3270" Or UCase(Trim(rsMainT2("車商代號"))) = "W3280" Or UCase(Trim(rsMainT2("車商代號"))) = "P1324" Or UCase(Trim(rsMainT2("車商代號"))) = "P2324" Then strFacility = "佰事達北倉"
        If UCase(Trim(rsMainT2("車商代號"))) = "W4270" Then strFacility = "佰事達中倉"
        If UCase(Trim(rsMainT2("車商代號"))) = "W8150" Then strFacility = "佰事達南倉"
        If UCase(Trim(rsMainT2("送貨客戶"))) = "DW327" Or UCase(Trim(rsMainT2("送貨客戶"))) = "DW328" Or UCase(Trim(rsMainT2("送貨客戶"))) = "DP1324" Then strFacility = "佰事達北倉"
        If UCase(Trim(rsMainT2("送貨客戶"))) = "DW427" Then strFacility = "佰事達中倉"
        If UCase(Trim(rsMainT2("送貨客戶"))) = "DW815" Then strFacility = "佰事達南倉"
        
        If Len(Trim(rsMainT2("預出日期"))) = 0 Then rsMainT2("預出日期") = Format(Now() + 1, "YYYYMMDD")

        str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,DeliveryDate,Stop,Door,Facility,ConsigneeKey,billtokey,c_company,b_contact1,c_phone1,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,amount,addwho,editwho) " & _
        "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT2("出貨單號")) & "','" & rsMainT2("類別") & "','LVTL01','" & Trim(rsMainT2("排出日期")) & "','" & Trim(rsMainT2("預出日期")) & "','" & Trim(rsMainT2("車商代號")) & "','" & Trim(rsMainT2("車商名稱")) & "','" & strFacility & "', " & _
        "'" & Trim(rsMainT2("送貨客戶")) & "','" & Trim(rsMainT2("帳款客戶")) & "','" & Trim(rsMainT2("送貨客戶名稱")) & "','" & Trim(rsMainT2("帳款客戶名稱")) & "','" & Trim(rsMainT2("客戶電話")) & "',substring('" & Trim(rsMainT2("客戶地址")) & "', 1, 60),substring('" & Trim(rsMainT2("客戶地址")) & "', 61, 45),'" & Trim(rsMainT2("客戶單號")) & "','" & Trim(rsMainT2("備註")) & "','" & filLocalFileT2.FileName & "','','" & Trim(rsMainT2("噸數")) & "','" & User_id & "','" & User_id & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        int_Order = int_Order + 1
    Else
        '訂單重複
        Call FTPlog("訂單重複" & str_SQL)
        '紀錄重複
        strReOrderkey = strReOrderkey & Trim(rsMainT2("出貨單號")) & Trim(rsMainT2("項次")) & "','"
        blDuplicationOrder = True

    End If
End If

'    '資料檢驗--判斷訂單明細是否重複，重複不增加明細，跳下一筆資料
'    Call Confirm_Recordset_Closed(tmp_Rs)
'    str_SQL = "select o.orderkey from ORDERDETAIL od (nolock) join orders o (nolock) on o.orderkey = od.orderkey where od.editdate >getdate()-1 and rtrim(o.ExternOrderKey) + rtrim(od.OrderLineNumber) ='" & Trim(rsMainT2("出貨單號")) & Trim(rsMainT2("項次")) & "' and o.storerkey = 'LVTL01' and isnull(o.type,'') <> '刪單' "
'    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If tmp_Rs.EOF And blDuplicationOrder = False Then
    If blDuplicationOrder = False Then
         intQTY = Trim(rsMainT2("數量"))
         If UCase(Trim(rsMainT2("單位"))) = "DZ" Then intQTY = Trim(rsMainT2("數量")) * 12

        '訂單明細資料新增
        str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber, ExternlineNO ,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice ,CartonGroup,notes)" & _
        "VALUES ('" & str_Orderkey & "','" & Trim(rsMainT2("項次")) & "','" & Trim(rsMainT2("出貨原因")) & "','" & Trim(rsMainT2("出貨單號")) & "','" & Trim(rsMainT2("產品編號")) & "','LVTL01'," & _
        "'" & intQTY & "','" & intQTY & "','" & rsMainT2("倉別") & "','','" & Trim(rsMainT2("單位")) & "','0','" & Trim(rsMainT2("使用棧板")) & "','" & Trim(rsMainT2("備註")) & "')"
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        
        int_OrderLine = int_OrderLine + 1
'
'    Else
'        '訂單明細重複
'        Call FTPlog("訂單明細重複" & str_SQL)
'        '紀錄重複
'        strReOrderkey = strReOrderkey & Trim(rsMainT2("出貨單號")) & Trim(rsMainT2("項次")) & "','"
'    End If
    End If
    
    rsMainT2.MoveNext
Loop

'補客戶資料
cn.Execute "exec gs_ordersupdate 'LVTL01' ", RowsAffect, adExecuteNoRecords


'到貨通知單 1.不入排車系統資料 2.產生預收採購單
Dim rsTmp As New ADODB.Recordset
rsTmp.CursorLocation = 3

'orderlinenumber 寫入podetail.externpokey，指定空白節省子查詢速度 edit by Eric 20140311
str_SQL = "select od.storerkey " & _
            ",o.orderkey " & _
            ", od.externorderkey " & _
            ", o.priority " & _
            ", orderlinenumber = ' '" & _
            ", od.sku " & _
            ", s.descr " & _
            ", openqty = sum(od.openqty) " & _
            ", notes = cast(o.notes as varchar(300)) " & _
            ", o.consigneekey " & _
            ", o.c_company " & _
            ", ContainerKey = case when len(rtrim(o.customerorderkey)) = 0 then od.externorderkey else o.customerorderkey end " & _
            "from orders o (nolock) join orderdetail od(nolock) on o.orderkey = od.orderkey " & _
            "join " & strWMSDB & "..sku s(nolock) on s.sku = od.sku and s.storerkey = od.storerkey " & _
            "where o.storerkey = 'LVTL01' and o.priority = 'RC' and o.B_PHONE2 is null " & _
            "group by od.storerkey ,o.orderkey , od.externorderkey , o.priority ,od.storerkey ,o.orderkey , od.externorderkey , o.priority , o.consigneekey , o.c_company ,od.sku, s.descr,cast(o.notes as varchar(300)) , o.customerorderkey " & _
            "order by od.storerkey , o.orderkey "
            
rsTmp.Open str_SQL, cn
If Not rsTmp.EOF Then

    Dim strKeycount As String, strOrderkey As String, intLineNumber As Integer
        
    rsTmp.MoveFirst
    Do While Not rsTmp.EOF
    
        '寫入WMS
        If Trim(rsTmp("orderkey")) <> strOrderkey Then
            intLineNumber = 1
            strOrderkey = Trim(rsTmp("orderkey"))
    
            '取系統PO單號
            Dim rsKeycount As New ADODB.Recordset
            rsKeycount.Open "select keycount = isnull(keycount,0) From " & strWMSDB & "..NCOUNTER where keyname='po' ", cn
            '單號+1
            cn.Execute "update " & strWMSDB & "..NCOUNTER set keycount='" & rsKeycount("Keycount") + 1 & "' where keyname= 'po'", RowsAffect, adExecuteNoRecords
            strKeycount = Format(rsKeycount("Keycount") + 1, "0000000000")
            rsKeycount.Close: Set rsKeycount = Nothing
    
            '寫入表頭
            str_SQL = "insert into " & strWMSDB & "..po (poKey,StorerKey,BuyersReference ,  BuyerVAT , sellername,selleraddress1,externpokey,potype,notes) " & _
                        "values( '" & strKeycount & "','" & rsTmp("StorerKey") & "','" & rsTmp("ExternOrderKey") & "','" & RTrim(rsTmp("ContainerKey")) & "','" & rsTmp("consigneekey") & "','" & rsTmp("C_company") & "','" & rsTmp("OrderKey") & "','" & rsTmp("priority") & "','" & rsTmp("notes") & "') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '註記已轉排車系統
            cn.Execute "update orders set B_PHONE2='00',trafficCop=null where orderkey = '" & rsTmp("OrderKey") & "' ", RowsAffect, adExecuteNoRecords
            
        End If
    
            '寫入表身
            str_SQL = "insert into " & strWMSDB & "..podetail (poKey,PoLineNumber,ExternLineNo,SKU,SkuDescription,StorerKey,QtyOrdered) " & _
                    "values( '" & strKeycount & "','" & Format(intLineNumber, "00000") & "','" & rsTmp("OrderLineNumber") & "','" & rsTmp("SKU") & "','" & rsTmp("descr") & "','" & rsTmp("StorerKey") & "','" & rsTmp("openqty") & "') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
            intLineNumber = intLineNumber + 1
    
        rsTmp.MoveNext

    Loop
    
End If
rsTmp.Close: Set rsTmp = Nothing

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，檔案 " & filLocalFileT2.FileName & " 備份於 C:\BEST\LVTL01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，檔案:" & strTranFileName)
    
    If int_Asn > 0 Then MsgBox "有 " & int_Asn & " 筆到貨通知單轉入!", vbOKOnly + vbInformation, Me.Caption: Call FTPlog("匯入 " & int_Asn & " 筆到貨通知訂單，檔案:" & strTranFileName)
    If int_Repeat > 0 Then MsgBox "有 " & int_Repeat & " 筆訂單明細重複轉檔!", vbOKOnly + vbInformation, Me.Caption: Call FTPlog("匯入 " & int_Repeat & " 筆重複訂單明細，檔案:" & strTranFileName)

'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT2.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 料號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o(nolock) join orderdetail od(nolock) on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) + rtrim(od.OrderLineNumber) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\LVTL01\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\LVTL01\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LVTL01\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''備份至FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT2.FileName

'備份檔案
If Dir("C:\BEST\LVTL01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LVTL01\Orders\Backup"
If Dir("C:\BEST\LVTL01\Orders\Backup\" & filLocalFileT2.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LVTL01\Orders\Backup\" & filLocalFileT2.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LVTL01\Orders\Backup\" & mySplit(filLocalFileT2.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & ".txt"
End If

Kill strTranFileName
    
filLocalFileT2.Refresh
Screen.MousePointer = 0: cmdImportT2.Enabled = True: dgMainT2.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-匯入", Me.Caption, "cmd_Import_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT2.Enabled = True: Screen.MousePointer = 0: dgMainT2.Enabled = True

End Sub

Private Sub dgMainT2_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT2
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, SSTab1.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT2_Change()
    filLocalFileT2.Path = dirLocalDirT2.Path
End Sub

Private Sub drvLocalDriveT2_Change()

On Error GoTo DriveError
dirLocalDirT2.Path = drvLocalDriveT2.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub
Private Sub cmdImportT4_Click()

strTranFileName = filLocalFileT4.Path & "\" & filLocalFileT4.FileName
If Len(RTrim(cboSheetT4)) = 0 Then MsgBox "請選擇工作表", 64, Me.Caption: Exit Sub
If rsMainT4.EOF Or rsMainT4 Is Nothing Then Exit Sub
Dim strStorerkey As String
On Error GoTo err_Handle

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT4.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

'貨主編號
strStorerkey = "LKAO01"
cmdImportT4.Enabled = False: dgMainT4.Enabled = False

'到貨日期檢查,判斷SKU是否存在
rsMainT4.MoveFirst
Do While Not rsMainT4.EOF
'
'    If Replace(myExCharFilter(Trim(rsMainT4("交貨日"))), ".", "/") < Format(Now - 1, "YYYY/MM/DD") Then: MsgBox "到貨日小於今日，訂單轉入終止!", 16, Me.Caption: Exit Sub

    str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & myExCharFilter(Trim(rsMainT4("商品代號"))) & "' and Storerkey = '" & strStorerkey & "' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        MsgBox "訂單發現新品號 (" & myExCharFilter(Trim(rsMainT4("商品代號"))) & " ) " & Trim(rsMainT4("商品名稱")) & "，訂單轉入終止!!": cmdImportT4.Enabled = True: dgMainT4.Enabled = True
        tmp_Rs.Close
        Exit Sub

'        '新增SKU
'        '檢查Packkey大於10碼重編Packkey
'        If Len(strSku) > 10 Then
'
'            '取Packkey流水號
'            Call Confirm_Recordset_Closed(tmp_rs)
'            'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = '" & strstorerkey & "' and left(consigneekey,4) = '" & strstorerkey & "' order by consigneekey desc "
'            'tmp_rs.Open str_SQL, cn
'            '
'            'If Not tmp_rs.EOF Then intTmp = Val(tmp_rs("consigneekey"))
'            '
'            'tmp_rs.Close
'
'        Else
'            strPackkey = strSku
'        End If
'
'        lngCasecnt = Val(rsMainT4("送貨 PCs")) / Val(rsMainT4("訂購/出貨數量"))
'        dblStdCube = Round(Val(rsMainT4("體積")) / Val(rsMainT4("送貨 PCs")) / 28316, 10)
'        dblStdGrossWGT = Val(rsMainT4("淨重")) / Val(rsMainT4("送貨 PCs"))
'
'        str_SQL = "insert into sku(Storerkey,SKU,SKUGROUP,DESCR,STDCUBE,STDGROSSWGT,SUSR1,SUSR2,SUSR3,SUSR4,SUSR5,BUSR1,BUSR2,BUSR3,BUSR4,BUSR5,Packkey,AllocParm,DefaultRotation,IOFlag,PickCode,PutAwayLoc,PutCode,PutAwayZone,ReceiptInspectionLoc,SKURotat01,StrategyKey,LOTTABLE01LABEL,LOTTABLE02LABEL,LOTTABLE03LABEL,LOTTABLE04LABEL,LOTTABLE05LABEL,LOTTABLE06LABEL,LOTTABLE07LABEL,LOTTABLE08LABEL,LOTTABLE09LABEL,LOTTABLE10LABEL,LOTTABLE11LABEL) Values " & _
'                  "('LKAO01','" & strSku & "','STD000N',convert(char(60),'" & myExCharFilter(Trim(rsMainT4("商品名稱"))) & "')," & dblStdCube & "," & dblStdGrossWGT & ",'',0,'',5000,1,'" & myExCharFilter(Trim(rsMainT4("BUn"))) & "','','" & myExCharFilter(Trim(rsMainT4("SU"))) & "','','','" & strPackkey & "','FLOAT PICK','FIFO','N','NSPFIFO','UNKNOWN','NSPPASTD','RACK','QC','LOTTABLE05','ZONEA','Pack Key','交運單號','生產批號','製造日','到期日','倉別','棧板ID','棧板類別','','','') "
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
'
'        '新增PACK
'        str_SQL = "insert into pack(Packkey,Packdescr,PackUOM1,Casecnt,LengthUOM1,WidthUOM1,HeightUOM1,CubeUOM1,PackUOM2,Innerpack,PackUOM3,Qty,LengthUOM3,WidthUOM3,HeightUOM3,CubeUOM3,PackUOM4,Pallet,PalletTI,PalletHI,ADDDate,ADDWho,EditDate,EditWho,replenishzone1,replenishzone2,replenishzone3,replenishzone4,replenishzone8,replenishzone9,CartonizeUOM3) Values " & _
'                  "('" & strSku & "','" & strPackkey & "','CS','" & lngCasecnt & "',0,0,0,0,'IP',0,'EA',1,0,0,0,0,'PL',0,0,0,getdate(),'SA',getdate(),'SA','CASE','PICK','PICK','PICK','PICK','PICK','Y') "
'        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

    End If
    tmp_Rs.Close

rsMainT4.MoveNext
Loop

Tran_Level = cn.BeginTrans
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, dblStdCube As Double, dblStdGrossWGT As Double, lngCasecnt As Long, intPointer As Integer
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strOrderKeyS As String, strPackkey As String, strSku As String, strAddress As String

Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
            
Screen.MousePointer = 11
            
'取最後客戶編號
'Call Confirm_Recordset_Closed(tmp_rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = '" & strstorerkey & "' and left(consigneekey,4) = '" & strstorerkey & "' order by consigneekey desc "
'tmp_rs.Open str_SQL, cn
'
'If Not tmp_rs.EOF Then intTmp = Val(tmp_rs("consigneekey"))
'
'tmp_rs.Close

rsMainT4.MoveFirst
Do While Not rsMainT4.EOF
'    DoEvents: DoEvents
    
    intPointer = 1
    strSku = myExCharFilter(Trim(rsMainT4("商品代號")))
    strAddress = myExCharFilter(Trim(rsMainT4("城市"))) & myExCharFilter(Trim(rsMainT4("住址"))) & myExCharFilter(Trim(rsMainT4("門牌")))
                  
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT4("訂購單號"))) Then
        strOrderNo = UCase(Trim(rsMainT4("訂購單號")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
'        '無此客戶編號新增
'        str_SQL = "select * from trp01m where storerkey = '" & strStorerkey & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT4("收貨客戶"))) & "' "
'        Call Confirm_Recordset_Closed(tmp_rs)
'        tmp_rs.CursorLocation = 3
'        tmp_rs.Open str_SQL, cn
'
'        If tmp_rs.EOF Then
'
'            '新增客戶主檔
'            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address) " & _
'            " values('" & strStorerkey & "','','" & myExCharFilter(Trim(rsMainT4("收貨客戶"))) & "','" & myExCharFilter(Trim(rsMainT4("客戶名稱"))) & "','" & myExCharFilter(Trim(rsMainT4("客戶名稱"))) & "','','','" & myExCharFilter(Trim(rsMainT4("城市"))) & myExCharFilter(Trim(rsMainT4("門牌"))) & "' ) ", RowsAffect, adExecuteNoRecords
'
'        End If
    
        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT4("訂購單號"))) & "' and storerkey = '" & strStorerkey & "' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
                     
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,B_company,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT4("訂購單號"))) & "','C','" & strStorerkey & "','" & myExCharFilter(Trim(rsMainT4("運送日期"))) & "','" & myExCharFilter(Trim(rsMainT4("交貨日"))) & "','','" & _
            myExCharFilter(Trim(rsMainT4("收貨客戶"))) & "','" & myExCharFilter(Trim(rsMainT4("收貨客戶"))) & "','" & myExCharFilter(Trim(rsMainT4("客戶名稱"))) & "','','','','" & GetWord(strAddress, intPointer, 58) & "','" & GetWord(strAddress, intPointer, 45) & "','','','" & filLocalFileT4.FileName & "','','" & User_id & "','" & User_id & "','' )"

'            '花王訂單新增customerorderkey,notes
'            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,B_company,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
'            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT4("訂購單號"))) & "','A2B','" & strStorerkey & "','" & myExCharFilter(Trim(rsMainT4("運送日期"))) & "','" & myExCharFilter(Trim(rsMainT4("交貨日"))) & "','','" & _
'            myExCharFilter(Trim(rsMainT4("收貨客戶"))) & "','" & myExCharFilter(Trim(rsMainT4("收貨客戶"))) & "','" & myExCharFilter(Trim(rsMainT4("客戶名稱"))) & "','','','','" & GetWord(strAddress, intPointer, 58) & "','" & GetWord(strAddress, intPointer, 45) & "','" & myExCharFilter(Trim(rsMainT4("採購單號"))) & "','" & myExCharFilter(Trim(rsMainT4("備註"))) & "','" & filLocalFileT4.FileName & "','','" & User_id & "','" & User_id & "','' )"
'
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & myExCharFilter(Trim(rsMainT4("訂購單號"))) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Val(myExCharFilter(Trim(rsMainT4("送貨 PCs"))))
            
            strLot06 = "R01"
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT4("訂購單號"))) & "','" & myExCharFilter(Trim(rsMainT4("商品代號"))) & "','" & strStorerkey & "'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT4("BUn"))) & "','0','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If

        rsMainT4.MoveNext
Loop

'新客戶檢查
str_SQL = "select 貨主=storerkey , 貨主訂單單號=externorderkey , 訂單類別 = priority , 訂單日期=orderdate , 到貨日期=deliverydate , 客戶編號=consigneekey , 客戶名稱=c_company , 檢查日期 = getdate() " & _
        "from orders o " & _
        "Where o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel顯示
    Call Recordset2Excel("缺客戶資料", tmp_Rs)
    If Dir("C:\BEST\" & strStorerkey & "\缺客戶資料", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\缺客戶資料"
    MyXlsApp.Range("h:h").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & strStorerkey & "\缺客戶資料\缺客戶資料_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmdImportT4.Enabled = True: dgMainT4.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
    MsgBox "發現新客戶，訂單轉入中止!", vbOKOnly, Me.Caption
    Exit Sub
End If

'補客戶資料
cn.Execute "exec gs_ordersupdate '" & strStorerkey & "' ", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，(共 " & rsMainT4.RecordCount & " 筆明細)，TMS單號： " & strOrderKeyS & "~" & str_Orderkey & "，檔案 " & filLocalFileT4.FileName & " 備份於 C:\BEST\" & strStorerkey & "\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，檔案 " & filLocalFileT4.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 本次檔案名稱 = '" & filLocalFileT4.FileName & "' , 上次檔案名稱 = o.updatesource ,重複訂單號碼 = rtrim(o.externorderkey) ,上次客戶單號 = rtrim(o.customerorderkey) ,  上次訂單日期 = convert(varchar,o.orderdate,111) , 上次到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 上次料號 = od.sku , 上次數量 = od.openqty ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\" & strStorerkey & "\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & strStorerkey & "\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''備份至FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT4.FileName

'備份檔案
If Dir("C:\BEST\" & strStorerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\Orders\Backup"
If Dir("C:\BEST\" & strStorerkey & "\Orders\Backup\" & filLocalFileT4.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & strStorerkey & "\Orders\Backup\" & filLocalFileT4.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & strStorerkey & "\Orders\Backup\" & mySplit(filLocalFileT4.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT4.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT4.Refresh: cboSheetT4.Clear
Screen.MousePointer = 0: cmdImportT4.Enabled = True: dgMainT4.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-匯入", Me.Caption, "cmdImportT4_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT4.Enabled = True: Screen.MousePointer = 0: dgMainT4.Enabled = True

End Sub

Private Sub dgMainT4_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT4
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT4_Change()
    filLocalFileT4.Path = dirLocalDirT4.Path
End Sub
Private Sub drvLocalDriveT4_Change()

On Error GoTo DriveError
dirLocalDirT4.Path = drvLocalDriveT4.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT4_Click()

On Error GoTo err_Handle
Set rsMainT4 = Nothing: Set dgMainT4.DataSource = rsMainT4
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT4.Path, 1) = "\" Then
    strFilePath = filLocalFileT4.Path
Else
    strFilePath = filLocalFileT4.Path & "\"
End If

If Dir(strFilePath & filLocalFileT4.FileName) = "" Then: filLocalFileT4.Refresh: Exit Sub

cboSheetT4.Clear

If UCase(mySplit(filLocalFileT4.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT4.FileName)
  
    '列出所有工作表
    blDo = False
    cboSheetT4.Clear
    
    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT4.AddItem MyXlsApp.Sheets(i).Name
        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT4.ListIndex = -1
    
    MyXlsApp.DisplayAlerts = False:  MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT4.FileName
    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT4.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cboSheetT4_Click()

On Error GoTo err_Handle
Dim str As String, strFieldName As String, strFilePath As String, IntI As Integer
IntI = 1
'確認路徑是否帶"\"
If Right(filLocalFileT4.Path, 1) = "\" Then
    strFilePath = filLocalFileT4.Path
Else
    strFilePath = filLocalFileT4.Path & "\"
End If

'建立欄位名稱陣列
strFieldName = ""

If Right(filLocalFileT4.Path, 1) <> "\" Then
    strFilePath = filLocalFileT4.Path & "\"
Else
    strFilePath = filLocalFileT4.Path
End If

''建立欄位名稱陣列
'strFieldName = "訂購單號" & Chr(9) & "出貨單號" & Chr(9) & "運送文件" & Chr(9) & "運送日期" & Chr(9) & "交貨日" & Chr(9) & "收貨客戶" & Chr(9) & "客戶名稱" & Chr(9) & "城市" & Chr(9) & "住址" & Chr(9) & "門牌" & Chr(9) & "商品代號" & Chr(9) & "商品名稱" & Chr(9) & "訂購/出貨數量" & Chr(9) & "SU" & Chr(9) & "送貨PCs" & Chr(9) & "BUn" & Chr(9) & "淨重" & Chr(9) & "WUn" & Chr(9) & "體積" & Chr(9) & "VUn" & Chr(9) & "庫別" & Chr(9)
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_Rs)
Set rsMainT4 = New ADODB.Recordset
'Call Excel2RecordsetT4(strFilePath & filLocalFileT4.FileName, cboSheetT4, strFieldName, rsMainT4)
Call Excel2Recordset(strFilePath & filLocalFileT4.FileName, cboSheetT4, strFieldName, tmp_Rs)
tmp_Rs.MoveFirst

Call Replication_Recordset(tmp_Rs, rsMainT4)
tmp_Rs.Close


Set dgMainT4.DataSource = rsMainT4

If rsMainT4 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT4
    MsgBox "此工作表共 " & rsMainT4.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
End If

rsMainT4.Sort = "訂購單號,編號"  'add by Eric 按照電子檔的順序去排序

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Sub Excel2RecordsetT4(strFileName As String, strSheetName As String, strFieldName As String, ByRef rs As ADODB.Recordset)
'**************************************************
'Create by Gemini @20090312 4 Excel匯入Recordset
'使用說明
'1.如果來源Excel工作表不帶欄位名稱，請於strFieldName指定，並以char(9)作為分隔符號
'strFieldName = "客戶" & Chr(9) & "統編" & Chr(9) & "銷貨日期" & Chr(9) & "銷貨單號" & Chr(9) & "聯絡人" & Chr(9) & "電話" & Chr(9) & "送貨地址" & Chr(9) & "發票號碼" & Chr(9) & "業務員" & Chr(9) & "備註" & Chr(9) & "品號" & Chr(9) & "品名" & Chr(9) & "倉庫" & Chr(9) & "數量" & Chr(9) & "單位" & Chr(9) & "前置單據/備註/客戶單號" & Chr(9)

'參數說明
'strFileName:來源檔案名稱路徑
'strSheetName:來源工作表
'strFieldName:欄位名稱
'rs:回傳的Recordset
'範例
'call Excel2Recordset ("C:\book1.xls","Sheet1", "客戶" & Chr(9) & "統編" & Chr(9) & "銷貨日期" & Chr(9),rsMain)
'**************************************************
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
If Dir(strFileName) = "" Then MsgBox "找不到檔案！", vbOKOnly + vbInformation, "Excel2Recordset": Exit Sub '找不到檔案

On Error GoTo err_Handle
Screen.MousePointer = 11

'開啟EXCEL物件
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '選定工作表
        If i = 1 Then strSheetName1 = RTrim(.Sheets(i).Name) '取第一個工作表名稱
        .Sheets(.Sheets(i).Name).Select
    Next
    
    '找不到指定工作表，選用第一個
    If RTrim(.ActiveSheet.Name) <> strSheetName Then .Sheets(strSheetName1).Select
    
    k = 1 '預設由第一列開始匯入
    
    '若無來源欄位名稱
    If strFieldName = "" Then
        '取欄位名稱
        For i = 1 To 255
            If Len(Trim(.Cells(1, i) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        k = 2 '由第二列開始匯入
    End If
    
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rs = Nothing: GoTo endsub
    '建立Recordset
    For i = 0 To UBound(arrTmp)
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    k = 7 '由7列開始
    
    '寫入Recordset
    Do While Len(RTrim(.Cells(k, 2))) > 0
    rsTmp.AddNew
        For j = 2 To UBound(arrTmp) + 2 ''由B7儲存格開始
            rsTmp(j - 2) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst

    Call OffLineRecordset(rsTmp, rs)
    
    rsTmp.Close: Set rsTmp = Nothing
  
endsub:
Screen.MousePointer = 0
.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

Exit Sub
err_Handle:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Set rs = Nothing

Dim str As String

If err.Number = 3367 Then
    str = "欄位名稱( " & arrTmp(i) & ")重複！"
    
ElseIf err.Number = -2147217887 Then
    str = "請確認儲存格(" & Chr(64 + j) & k & ")，資料是否有誤！"

Else
     str = "Exceed2Recordset"
End If

Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, str)

End Sub
Private Sub cmdImportT5_Click()

On Error GoTo err_Handle
strTranFileName = filLocalFileT5.Path & "\" & filLocalFileT5.FileName
If Len(RTrim(cboSheetT5)) = 0 Then MsgBox "請選擇工作表", 64, Me.Caption: Exit Sub
If rsMainT5.EOF Or rsMainT5 Is Nothing Then Exit Sub

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT5.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT5.Enabled = True: dgMainT5.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp
rsMainT5.MoveFirst
Do While Not rsMainT5.EOF

    '到貨日期檢查
    arrTmp = Split(Trim(rsMainT5("銷貨日期")), "/")
    If Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "到貨日小於今日，訂單轉入終止!", 16, Me.Caption: Exit Sub
    
    '數量檢查
    If Trim(rsMainT5("數量")) > 0 Then
        MsgBox "發現訂單數量大於0，" & Trim(rsMainT5("品號")) & "-" & Trim(rsMainT5("品名")) & "(" & Trim(rsMainT5("數量")) & Trim(rsMainT5("單位")) & ")，訂單轉入終止!!", , "退貨單匯入": Exit Sub
        Exit Sub
    End If
    
    rsMainT5.MoveNext
Loop

Tran_Level = cn.BeginTrans: cmdImportT5.Enabled = False: dgMainT5.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
            
'取最後客戶編號
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LNIP01' and left(consigneekey,4) = 'LNIP' order by consigneekey desc "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))

tmp_Rs.Close

rsMainT5.MoveFirst
Do While Not rsMainT5.EOF
    DoEvents: DoEvents
    
'    '資料檢驗--判斷是否屬佰事達訂單-->跳下一筆
'    If UCase(Trim(rsMaint5("倉庫"))) = "" Then
''        MsgBox "客戶單號：" & Trim(rsMainT4("銷貨單號")) & "( " & Trim(rsMainT4("倉庫")) & " )" & vbCrLf & "請通知客戶，確認該筆訂單是否有誤!?", vbOKOnly, "非佰事達之訂單不轉入"
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
        '資料檢驗--判斷SKU是否存在
    str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT5("品號")) & "' and Storerkey = 'LNIP01' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        cn.RollbackTrans: Tran_Level = 0
        MsgBox "訂單發現新品號 (" & Trim(rsMainT5("品號")) & " ) " & Trim(rsMainT5("品名")) & "，訂單轉入終止!!": cmdImportT5.Enabled = True: dgMainT5.Enabled = True: Screen.MousePointer = 0
        Exit Sub
    End If

'    End If
                  
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT5("銷貨單號"))) Then
        strOrderNo = UCase(Trim(rsMainT5("銷貨單號")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '檢查是否有此客戶名稱
        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LNIP01' and rtrim(full_name) = '" & myExCharFilter(Trim(rsMainT5("客戶"))) & "' order by len(consigneekey),consigneekey "
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        
        If tmp_Rs.EOF Then
            '無此客戶名稱則新增
            intTmp = intTmp + 1
            strConsigneeKey = "LNIP" & Format(intTmp, "000000")
            
            '新增客戶主檔
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address) " & _
            " values('LNIP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT5("客戶"))) & "','" & myExCharFilter(Trim(rsMainT5("客戶"))) & "','" & myExCharFilter(Trim(rsMainT5("聯絡人"))) & "','" & myExCharFilter(Trim(rsMainT5("電話"))) & "','" & myExCharFilter(Trim(rsMainT5("送貨地址"))) & "' ) ", RowsAffect, adExecuteNoRecords
'
            '紀錄新增之客戶編號
            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
        Else
            '比對聯絡人、電話與到貨地址是否相符
            Call Confirm_Recordset_Closed(rsTmp)
            str_SQL = "select * from trp01m " & _
                        "where storerkey = 'LNIP01' and full_name = '" & myExCharFilter(Trim(rsMainT5("客戶"))) & "' " & _
                        "and rtrim(contact) = '" & myExCharFilter(Trim(rsMainT5("聯絡人"))) & "' " & _
                        "and rtrim(phone) = '" & myExCharFilter(Trim(rsMainT5("電話"))) & "' " & _
                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT5("送貨地址"))) & "' "
            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn
            
            If rsTmp.EOF Then
                '聯絡人、電話與到貨地址不符
                intTmp = intTmp + 1
                strConsigneeKey = "LNIP" & Format(intTmp, "000000")
                
                '新增客戶主檔
                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes) " & _
                " values('LNIP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT5("客戶"))) & "','" & myExCharFilter(Trim(rsMainT5("客戶"))) & "','" & myExCharFilter(Trim(rsMainT5("聯絡人"))) & "','" & myExCharFilter(Trim(rsMainT5("電話"))) & "','" & myExCharFilter(Trim(rsMainT5("送貨地址"))) & "','" & myExCharFilter(Trim(tmp_Rs("notes"))) & "' ) ", RowsAffect, adExecuteNoRecords
                 
                '紀錄新增之客戶編號
                strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
            Else '相符沿用舊客編
                strConsigneeKey = Trim(rsTmp("consigneekey"))
                blCustomerMatch = True

            End If
            rsTmp.Close
        End If
        tmp_Rs.Close
    
'        '資料檢驗--判斷訂單是否重複，重複不增加
'        Call Confirm_Recordset_Closed(tmp_rs)
'        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMaint5("銷貨單號"))) & "' and storerkey = 'LNIP01' and isnull(type,'') <> '刪單' "
'        tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'        If tmp_rs.EOF Then
    
            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
            strFacility = Trim(rsMainT5("倉庫"))
            strFacility = "佰事達北倉"
            arrTmp = Split(Trim(rsMainT5("銷貨日期")), "/")
            strOrderDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT5("銷貨單號"))) & "','R','LNIP01','" & strOrderDate & "','" & CDate(strOrderDate) + 1 & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT5("客戶"))) & "','" & myExCharFilter(Trim(rsMainT5("聯絡人"))) & "','" & myExCharFilter(Trim(rsMainT5("業務員"))) & "','" & myExCharFilter(Trim(rsMainT5("統編"))) & "','" & myExCharFilter(Trim(rsMainT5("電話"))) & "','','" & myExCharFilter(Trim(GetWord(rsMainT5("送貨地址"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT5("送貨地址"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT5("備註"))) & "','" & filLocalFileT5.FileName & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT5("發票號碼"))) & "' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LNIP01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
'        Else
'            '訂單重複
'            Call FTPlog("訂單重複" & str_SQL)
'            '紀錄重複
'            strReOrderkey = strReOrderkey & Trim(rsMaint5("銷貨單號")) & "','"
'            blDuplicationOrder = True
'
'        End If
    End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Abs(Trim(rsMainT5("數量")))
            strLot06 = IIf(UCase(Trim(rsMainT5("倉庫"))) = "A06", "A06-S", Trim(rsMainT5("倉庫")))
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT5("銷貨單號"))) & "','" & myExCharFilter(Trim(rsMainT5("品號"))) & "','LNIP01'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','" & myExCharFilter(rsMainT5("倉庫")) & "','" & myExCharFilter(Trim(rsMainT5("單位"))) & "','0','" & myExCharFilter(Trim(rsMainT5("前置單據/備註/客戶單號"))) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT5.MoveNext
Loop

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細 (共 " & rsMainT5.RecordCount & " 筆明細)，TMS單號： " & strOrderKeyS & "~" & str_Orderkey & "，檔案 " & filLocalFileT5.FileName & " 備份於 C:\BEST\LNIP01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細，檔案 " & filLocalFileT5.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT5.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 料號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\LNIP01\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\LNIP01\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNIP01\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''備份至FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT4.FileName

'備份檔案
If Dir("C:\BEST\LNIP01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LNIP01\Orders\Backup"
If Dir("C:\BEST\LNIP01\Orders\Backup\" & filLocalFileT5.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LNIP01\Orders\Backup\" & filLocalFileT5.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LNIP01\Orders\Backup\" & mySplit(filLocalFileT5.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT5.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT5.Refresh: cboSheetT5.Clear
Screen.MousePointer = 0: cmdImportT5.Enabled = True: dgMainT5.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-匯入", Me.Caption, "cmdImportt5_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT5.Enabled = True: Screen.MousePointer = 0: dgMainT5.Enabled = True

End Sub
'Sub ExcelSheet2Recordset()
'On Error GoTo err_Handle
'Dim strExcel As String, arrTmp, strFilePath As String
'
''確認路徑是否帶"\"
'If Right(filLocalFileT4.Path, 1) = "\" Then
'    strFilePath = filLocalFileT4.Path
'Else
'    strFilePath = filLocalFileT4.Path & "\"
'End If
'
''建立欄位名稱陣列
'arrTmp = Array("客戶", "統編", "銷貨日期", "銷貨單號", "聯絡人", "電話", "送貨地址", "發票號碼", "業務員", "備註", "品號", "品名", "倉庫", "數量", "單位", "前置單據/備註/客戶單號")
'
''建立 Excel 報表資料庫連接
'Dim cnExcel As New ADODB.Connection ':Set cnExcel = New ADODB.Connection
'cnExcel.Provider = "Microsoft.Jet.OLEDB.4.0"
'cnExcel.ConnectionString = "Data Source=" & filLocalFileT4.Path & "\" & filLocalFileT4.FileName & ";Extended Properties=""Excel 8.0; HDR=No;ReadOnly=True;"""
'cnExcel.Open
'
'Call ReDim_Recordset(tmp_rs)
'
'tmp_rs.CursorLocation = 3
'tmp_rs.Open "select * from [" & cboSheetT4 & "$] ", cnExcel ', adOpenStatic, adLockOptimistic
'tmp_rs.Sort = "F4,F16"
'
'Set rsMainT4 = New ADODB.Recordset
'
''將傳入 tmp_rs 完整複製至 rsMainT4
'Dim fldcnt As Integer, reccnt As Double
'
''建立 Recordset 的 Table 架構 (在記憶體中的 ADO Recordset)
'rsMainT4.Fields.Append "編號", adDouble
'For fldcnt = 0 To tmp_rs.Fields.Count - 1
'    rsMainT4.Fields.Append arrTmp(fldcnt), tmp_rs.Fields(fldcnt).Type, tmp_rs.Fields(fldcnt).DefinedSize
'Next fldcnt
'
'With rsMainT4
'     .CursorType = adOpenStatic
'     .LockType = adLockOptimistic
'     .Open    '不需連接物件
'End With
'
'reccnt = 0
'Do While Not tmp_rs.EOF
'   reccnt = reccnt + 1
'   rsMainT4.AddNew
'   rsMainT4.Fields(0).Value = reccnt
'   For fldcnt = 0 To tmp_rs.Fields.Count - 1
'    rsMainT4.Fields(fldcnt + 1).Value = tmp_rs.Fields(fldcnt).Value & ""
'   Next fldcnt
'   rsMainT4.Update
'   tmp_rs.MoveNext
'Loop
'
'tmp_rs.Close: Set tmp_rs = Nothing
'
'Set dgMainT4.DataSource = rsMainT4: dgMainT4.Visible = False
'
'rsMainT4.MoveFirst
'
'SetDataGridColWidth Me.Caption, dgMainT4
'dgMainT4.RowHeight = 300
'Screen.MousePointer = 0: dgMainT4.Visible = True
'MsgBox "此工作表共 " & rsMainT4.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "工作表開啟"
'
'cnExcel.Close: Set cnExcel = Nothing
'
'Exit Sub
'err_Handle:
'Set cnExcel = Nothing
'Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")

'End Sub

Private Sub dgMainT5_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT5
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT5_Change()
    filLocalFileT5.Path = dirLocalDirT5.Path
End Sub
Private Sub drvLocalDriveT5_Change()

On Error GoTo DriveError
dirLocalDirT5.Path = drvLocalDriveT5.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT5_Click()

On Error GoTo err_Handle
Set rsMainT5 = Nothing: Set dgMainT5.DataSource = rsMainT5
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT5.Path, 1) = "\" Then
    strFilePath = filLocalFileT5.Path
Else
    strFilePath = filLocalFileT5.Path & "\"
End If

If Dir(strFilePath & filLocalFileT5.FileName) = "" Then: filLocalFileT5.Refresh: Exit Sub

cboSheetT5.Clear

If UCase(mySplit(filLocalFileT5.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT5.FileName)
    MyXlsApp.DisplayAlerts = False
  
    '列出所有工作表
    blDo = False
    cboSheetT5.Clear
    
    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT5.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT5.ListIndex = -1
    
    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT5.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cboSheetT5_Click()

On Error GoTo err_Handle
Dim str As String, strFieldName As String, strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT5.Path, 1) = "\" Then
    strFilePath = filLocalFileT5.Path
Else
    strFilePath = filLocalFileT5.Path & "\"
End If

'建立欄位名稱陣列
strFieldName = "客戶" & Chr(9) & "統編" & Chr(9) & "銷貨日期" & Chr(9) & "銷貨單號" & Chr(9) & "聯絡人" & Chr(9) & "電話" & Chr(9) & "送貨地址" & Chr(9) & "發票號碼" & Chr(9) & "業務員" & Chr(9) & "備註" & Chr(9) & "品號" & Chr(9) & "品名" & Chr(9) & "倉庫" & Chr(9) & "數量" & Chr(9) & "單位" & Chr(9) & "前置單據/備註/客戶單號" & Chr(9)

If Right(filLocalFileT5.Path, 1) <> "\" Then
    strFilePath = filLocalFileT5.Path & "\"
Else
    strFilePath = filLocalFileT5.Path
End If

Set rsMainT5 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT5.FileName, cboSheetT5, strFieldName, rsMainT5)

Set dgMainT5.DataSource = rsMainT5

If rsMainT5 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT5
    MsgBox "此工作表共 " & rsMainT5.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Sub ExcelSheet2RecordsetT5()
On Error GoTo err_Handle
Dim strExcel As String, arrTmp, strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT5.Path, 1) = "\" Then
    strFilePath = filLocalFileT5.Path
Else
    strFilePath = filLocalFileT5.Path & "\"
End If

'建立 Excel 報表資料庫連接
Dim cnExcel As New ADODB.Connection ':Set cnExcel = New ADODB.Connection
cnExcel.Provider = "Microsoft.Jet.OLEDB.4.0"
cnExcel.ConnectionString = "Data Source=" & filLocalFileT5.Path & "\" & filLocalFileT5.FileName & ";Extended Properties=""Excel 8.0;ReadOnly=True;"""
cnExcel.Open

Call ReDim_Recordset(tmp_Rs)

tmp_Rs.CursorLocation = 3
tmp_Rs.Open "select * from [" & cboSheetT5 & "$] where len(rtrim(客戶)) > 0 ", cnExcel ', adOpenStatic, adLockOptimistic
tmp_Rs.Sort = "銷貨單號,品號"

Set rsMainT5 = New ADODB.Recordset

'將傳入 tmp_rs 完整複製至 rsMainT5
Dim fldcnt As Integer, reccnt As Double

'建立 Recordset 的 Table 架構 (在記憶體中的 ADO Recordset)
rsMainT5.Fields.Append "編號", adDouble
For fldcnt = 0 To tmp_Rs.Fields.Count - 1
    rsMainT5.Fields.Append tmp_Rs.Fields(fldcnt).Name, tmp_Rs.Fields(fldcnt).Type, tmp_Rs.Fields(fldcnt).DefinedSize
Next fldcnt

With rsMainT5
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With

reccnt = 0
Do While Not tmp_Rs.EOF
   reccnt = reccnt + 1
   rsMainT5.AddNew
   rsMainT5.Fields(0).Value = reccnt
   For fldcnt = 0 To tmp_Rs.Fields.Count - 1
    rsMainT5.Fields(fldcnt + 1).Value = myExCharFilter(tmp_Rs.Fields(fldcnt).Value & "")
   Next fldcnt
   rsMainT5.Update
   tmp_Rs.MoveNext
Loop

tmp_Rs.Close: Set tmp_Rs = Nothing

Set dgMainT5.DataSource = rsMainT5: dgMainT5.Visible = False

rsMainT5.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT5

Screen.MousePointer = 0: dgMainT5.Visible = True
MsgBox "此工作表共 " & rsMainT5.RecordCount & " 筆明細", 64, "工作表開啟"

cnExcel.Close: Set cnExcel = Nothing

Exit Sub
err_Handle:
Set cnExcel = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT6_Click()

strTranFileName = filLocalFileT6.Path & "\" & filLocalFileT6.FileName
If Len(RTrim(cboSheetT6)) = 0 Then MsgBox "請選擇工作表", 64, Me.Caption: Exit Sub
If rsMainT6.EOF Or rsMainT6 Is Nothing Then Exit Sub
Dim strStorerkey As String, strSku As String, strPackkey As String, lngCasecnt As Long, lngPallet As Long, dblStdCube As Double, dblStdGrossWGT As Double
Dim rsTmp As New ADODB.Recordset
On Error GoTo err_Handle

'貨主編號
strStorerkey = "LSJR01"

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select updatesource from orders where rtrim(updatesource)='" & filLocalFileT6.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

'資料檢驗
rsMainT6.MoveFirst
Do While Not rsMainT6.EOF

    '到貨日期檢查
'    If Replace(myExCharFilter(Trim(rsMainT6("客戶指定送貨日"))), ".", "/") < Format(Now - 1, "YYYY/MM/DD") Then: MsgBox "到貨日小於今日，訂單轉入終止!", 16, Me.Caption: Exit Sub

    '判斷出貨倉是否存在
    strSku = RTrim(myExCharFilter(Trim(rsMainT6("出貨倉"))))
    str_SQL = "select * from trp01m where consigneekey='" & RTrim(myExCharFilter(Trim(rsMainT6("出貨倉")))) & "' and Storerkey = '" & strStorerkey & "' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        MsgBox "無出貨倉資料 (" & myExCharFilter(Trim(rsMainT6("出貨倉"))) & ")，請新增後再轉入。 ", 16, "訂單轉入終止!!"
        tmp_Rs.Close: Exit Sub
    End If
    tmp_Rs.Close

    '判斷SKU是否存在
    strSku = RTrim(myExCharFilter(Trim(rsMainT6("料號"))))
    str_SQL = "select * from " & strWMSDB & "..sku where sku='" & strSku & "' and Storerkey = '" & strStorerkey & "' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
'        MsgBox "訂單發現新品號 (" & myExCharFilter(Trim(rsMainT6("商品代號"))) & " ) " & Trim(rsMainT6("商品名稱")) & "，訂單轉入終止!!"
'        tmp_Rs.Close: Exit Sub

       '新增SKU
       '重編Packkey
        Call Confirm_Recordset_Closed(rsTmp)
        str_SQL = "select top 1 substring(packkey,7,20) as packkey from " & strWMSDB & "..sku where storerkey = '" & strStorerkey & "' and left(packkey,6) = '" & strStorerkey & "' order by substring(packkey,7,20) desc "
        rsTmp.Open str_SQL, cn
        
        If rsTmp.EOF Then
            strPackkey = strStorerkey & "0001"
        Else
            strPackkey = strStorerkey & Format(Val(rsTmp("packkey")) + 1, "0000")
        End If
        
        rsTmp.Close

        lngPallet = Val(rsMainT6("大單位個數")) * Val(rsMainT6("每板箱數"))
        lngCasecnt = Val(rsMainT6("大單位個數"))
        dblStdCube = Round(Val(rsMainT6("小單位材積") / rsMainT6("大單位個數") / 28316), 10)
        dblStdGrossWGT = Val(rsMainT6("小單位重量") / rsMainT6("大單位個數"))
        
        '新增SKU
        str_SQL = "insert into sku(Storerkey,SKU,SKUGROUP,DESCR,STDCUBE,STDGROSSWGT,SUSR1,SUSR2,SUSR3,SUSR4,SUSR5,BUSR1,BUSR2,BUSR3,BUSR4,BUSR5,Packkey,AllocParm,DefaultRotation,IOFlag,PickCode,PutAwayLoc,PutCode,PutAwayZone,ReceiptInspectionLoc,SKURotat01,StrategyKey,LOTTABLE01LABEL,LOTTABLE02LABEL,LOTTABLE03LABEL,LOTTABLE04LABEL,LOTTABLE05LABEL,LOTTABLE06LABEL,LOTTABLE07LABEL,LOTTABLE08LABEL,LOTTABLE09LABEL,LOTTABLE10LABEL,LOTTABLE11LABEL) Values " & _
                  "('" & strStorerkey & "','" & strSku & "','STD000N',convert(char(60),'" & myExCharFilter(Trim(rsMainT6("產品名稱"))) & "')," & dblStdCube & "," & dblStdGrossWGT & ",'',0,'',5000,1,'EA','','CS','','','" & strPackkey & "','FLOAT PICK','FIFO','N','NSPFIFO','UNKNOWN','NSPPASTD','RACK','QC','LOTTABLE05','ZONEA','Pack Key','交運單號','生產批號','製造日','到期日','倉別','棧板ID','棧板類別','','','') "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

        '新增PACK
        str_SQL = "insert into pack(Packkey,Packdescr,PackUOM1,Casecnt,LengthUOM1,WidthUOM1,HeightUOM1,CubeUOM1,PackUOM2,Innerpack,PackUOM3,Qty,LengthUOM3,WidthUOM3,HeightUOM3,CubeUOM3,PackUOM4,Pallet,PalletTI,PalletHI,ADDDate,ADDWho,EditDate,EditWho,replenishzone1,replenishzone2,replenishzone3,replenishzone4,replenishzone8,replenishzone9,CartonizeUOM3) Values " & _
                  "('" & strPackkey & "','" & strPackkey & "_" & strSku & "','CS','" & lngCasecnt & "',0,0,0,0,'IP',0,'EA',1,0,0,0,0,'PL'," & lngPallet & ",0,0,getdate(),'SA',getdate(),'SA','CASE','PICK','PICK','PICK','PICK','PICK','Y') "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    Else '更新商品主檔
    
        lngPallet = Val(rsMainT6("大單位個數")) * Val(rsMainT6("每板箱數"))
        lngCasecnt = Val(rsMainT6("大單位個數"))
        dblStdCube = Round(Val(rsMainT6("小單位材積") / rsMainT6("大單位個數") / 28316), 10)
        dblStdGrossWGT = Val(rsMainT6("小單位重量") / rsMainT6("大單位個數"))
        
        '更新sku
        str_SQL = "update " & strWMSDB & "..sku " & _
                  "set DESCR = '" & myExCharFilter(Trim(rsMainT6("產品名稱"))) & "' " & _
                  ",STDCUBE = " & dblStdCube & " " & _
                  ",STDGROSSWGT = " & dblStdGrossWGT & " " & _
                  "where storerkey = '" & strStorerkey & "' and sku = '" & strSku & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

        '更新PACK
        str_SQL = "update " & strWMSDB & "..pack " & _
        "set Casecnt = " & lngCasecnt & " " & _
        ",Pallet = " & lngPallet & " " & _
        "where packkey = '" & tmp_Rs("packkey") & "' "
        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

    End If

rsMainT6.MoveNext
Loop

Tran_Level = cn.BeginTrans: cmdImportT6.Enabled = False: dgMainT6.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, intPointer As Integer
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strOrderKeyS As String, strAddress As String, strNotes As String, strDeleteOrder As String

Dim rsTmp1 As New ADODB.Recordset
            
Screen.MousePointer = 11
            
'取最後客戶編號
'Call Confirm_Recordset_Closed(tmp_rs)
'str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = '" & strstorerkey & "' and left(consigneekey,4) = '" & strstorerkey & "' order by consigneekey desc "
'tmp_rs.Open str_SQL, cn
'
'If Not tmp_rs.EOF Then intTmp = Val(tmp_rs("consigneekey"))
'
'tmp_rs.Close

rsMainT6.MoveFirst
Do While Not rsMainT6.EOF
    
    '刪單檢查
    If UCase(myExCharFilter(Trim(rsMainT6("訂單狀態")))) = "DELETE" Then

    strDeleteOrder = strDeleteOrder + UCase(Trim(rsMainT6("貨主訂單號碼"))) & "','"

    GoTo nextLine
    End If
    
    intPointer = 1
    strSku = myExCharFilter(Trim(rsMainT6("料號")))
    strAddress = myExCharFilter(Trim(rsMainT6("到貨地址")))
                     
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT6("貨主訂單號碼"))) Then
        strOrderNo = UCase(Trim(rsMainT6("貨主訂單號碼")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
           
        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select externorderkey from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT6("貨主訂單號碼"))) & "' and storerkey = '" & strStorerkey & "' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '取TMS單號
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '無此客戶編號新增
            str_SQL = "select consigneekey from trp01m where storerkey = '" & strStorerkey & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT6("客戶編號"))) & "' "
            Call Confirm_Recordset_Closed(rsTmp)
            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn
    
            If rsTmp.EOF Then
                '新增客戶主檔
                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Fax,channel_type,Address,updatesource) " & _
                " values('" & strStorerkey & "','" & myExCharFilter(Trim(rsMainT6("郵遞區號"))) & "','" & myExCharFilter(Trim(rsMainT6("客戶編號"))) & "','" & myExCharFilter(Trim(rsMainT6("客戶名稱"))) & "','" & myExCharFilter(Trim(rsMainT6("客戶簡稱"))) & "','" & myExCharFilter(Trim(rsMainT6("聯絡人"))) & "','" & myExCharFilter(Trim(rsMainT6("電話"))) & "','" & myExCharFilter(Trim(rsMainT6("傳真"))) & "','" & myExCharFilter(Trim(rsMainT6("通路別"))) & "','" & strAddress & "','" & strOrderKeyS & "') ", RowsAffect, adExecuteNoRecords
            End If
            rsTmp.Close
            
            If Len(myExCharFilter(Trim(rsMainT6("指定到貨時間")))) > 0 Then
                strNotes = "指定到貨時間:" & myExCharFilter(Trim(rsMainT6("指定到貨時間"))) & ";"
            End If
            
            strNotes = strNotes & myExCharFilter(Trim(rsMainT6("訂單備註")))
            
            '新增訂單表頭
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,b_company) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT6("貨主訂單號碼"))) & "','A2B','" & strStorerkey & "','" & myExCharFilter(Trim(rsMainT6("訂單日期"))) & "','" & myExCharFilter(Trim(rsMainT6("到貨日"))) & "','','" & _
            myExCharFilter(Trim(rsMainT6("出貨倉"))) & "','','','','','','','" & myExCharFilter(Trim(rsMainT6("客戶單號"))) & "','" & strNotes & "','" & filLocalFileT6.FileName & "','','" & User_id & "','" & User_id & "','','" & myExCharFilter(Trim(rsMainT6("客戶編號"))) & "')"
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & myExCharFilter(Trim(rsMainT6("貨主訂單號碼"))) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Val(myExCharFilter(Trim(rsMainT6("訂單數量")))) * Val(myExCharFilter(Trim(rsMainT6("大單位個數"))))
            
            strLot06 = "R01"
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternLineNo,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable04,Lottable05,Lottable06, Facility,UOM,Unitprice,notes)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT6("項次"))) & "','" & myExCharFilter(Trim(rsMainT6("貨主訂單號碼"))) & "','" & myExCharFilter(Trim(rsMainT6("料號"))) & "','" & strStorerkey & "'," & _
            "'" & intQTY & "','" & intQTY & "','" & myExCharFilter(Trim(rsMainT6("指定製造日"))) & "','" & myExCharFilter(Trim(rsMainT6("指定到期日"))) & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT6("小單位名稱"))) & "','" & myExCharFilter(Trim(rsMainT6("單價"))) & "','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If

nextLine:
        rsMainT6.MoveNext
Loop

'刪單通知
If Len(RTrim(strDeleteOrder)) > 0 Then

MsgBox "請注意有刪單通知！", 64, "訂單轉入"

str_SQL = "select 貨主=storerkey ,TMS單號 = rtrim(o.orderkey),貨主訂單單號=externorderkey , 訂單類別 = priority , 訂單日期=orderdate , 到貨日期=deliverydate , 客戶編號=consigneekey , 客戶名稱=c_company,訂單狀態 = rtrim(o.type) ,通知檔案 = '" & filLocalFileT6.FileName & "', 檢查日期 = getdate() " & _
        "from orders o " & _
        "Where o.externorderkey in ('" & strDeleteOrder & "') "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF = False Then
    
        'Excel顯示
        Call Recordset2Excel("刪單通知", tmp_Rs)
        If Dir("C:\BEST\" & strStorerkey & "\刪單通知", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\刪單通知"
        MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & strStorerkey & "\刪單通知\刪單通知_" & Format(Now, "yyyymmddhhMMss") & ".xls"
        Set MyXlsApp = Nothing: tmp_Rs.Close
    
    End If

End If

''新客戶檢查
'str_SQL = "select 貨主=storerkey , 訂單單號=externorderkey , 訂單類別 = priority , 訂單日期=orderdate , 到貨日期=deliverydate , 客戶編號=consigneekey , 客戶名稱=c_company , 檢查日期 = getdate() " & _
'        "from orders o " & _
'        "Where o.b_phone2 Is Null " & _
'        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
'
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'If tmp_Rs.EOF = False Then
'
'    'Excel顯示
'    Call Recordset2Excel("缺客戶資料", tmp_Rs)
'    If Dir("C:\BEST\" & strStorerkey & "\缺客戶資料", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\缺客戶資料"
'    MyXlsApp.Range("h:h").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
'    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & strStorerkey & "\缺客戶資料\缺客戶資料_" & Format(Now, "yyyymmddhhMMss") & ".xls"
'    Set MyXlsApp = Nothing: tmp_Rs.Close
'
'    cmdImportT6.Enabled = True: dgMainT6.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
'    MsgBox "發現新客戶，訂單轉入中止!", vbOKOnly, Me.Caption
'    Exit Sub
'End If

'補客戶資料
cn.Execute "exec gs_ordersupdate '" & strStorerkey & "' ", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，(共 " & rsMainT6.RecordCount & " 筆明細)，TMS單號： " & strOrderKeyS & "~" & str_Orderkey & "，檔案 " & filLocalFileT6.FileName & " 備份於 C:\BEST\" & strStorerkey & "\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，檔案 " & filLocalFileT6.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 本次檔案名稱 = '" & filLocalFileT6.FileName & "' , 上次檔案名稱 = o.updatesource ,重複訂單號碼 = rtrim(o.externorderkey) ,上次客戶單號 = rtrim(o.customerorderkey) ,  上次訂單日期 = convert(varchar,o.orderdate,111) , 上次到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 上次料號 = od.sku , 上次數量 = od.openqty ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\" & strStorerkey & "\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & strStorerkey & "\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''備份至FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFilet6.FileName

'備份檔案
If Dir("C:\BEST\" & strStorerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & strStorerkey & "\Orders\Backup"
If Dir("C:\BEST\" & strStorerkey & "\Orders\Backup\" & filLocalFileT6.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & strStorerkey & "\Orders\Backup\" & filLocalFileT6.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & strStorerkey & "\Orders\Backup\" & mySplit(filLocalFileT6.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT6.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT6.Refresh: cboSheetT6.Clear
Screen.MousePointer = 0: cmdImportT6.Enabled = True: dgMainT6.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-匯入", Me.Caption, "cmdImportt6_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT6.Enabled = True: Screen.MousePointer = 0: dgMainT6.Enabled = True

End Sub

Private Sub dgMainT6_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT6
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT6_Change()
    filLocalFileT6.Path = dirLocalDirT6.Path
End Sub
Private Sub drvLocalDriveT6_Change()

On Error GoTo DriveError
dirLocalDirT6.Path = drvLocalDriveT6.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT6_Click()

On Error GoTo err_Handle
Set rsMainT6 = Nothing: Set dgMainT6.DataSource = rsMainT6
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT6.Path, 1) = "\" Then
    strFilePath = filLocalFileT6.Path
Else
    strFilePath = filLocalFileT6.Path & "\"
End If

If Dir(strFilePath & filLocalFileT6.FileName) = "" Then: filLocalFileT6.Refresh: Exit Sub

cboSheetT6.Clear

If UCase(mySplit(filLocalFileT6.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT6.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT6.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT6.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT6.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT6.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cboSheetT6_Click()

On Error GoTo err_Handle

Dim strFilePath As String
If Right(filLocalFileT6.Path, 1) <> "\" Then
    strFilePath = filLocalFileT6.Path & "\"
Else
    strFilePath = filLocalFileT6.Path
End If

Set rsMainT6 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT6.FileName, cboSheetT6, "", rsMainT6)

Set dgMainT6.DataSource = rsMainT6

If rsMainT6 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT6
    MsgBox "此工作表共 " & rsMainT6.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
rsMainT6.Sort = "貨主訂單號碼,項次"
    
End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
'Private Sub cboSheetT7_Click()
'If blDo = True Then Call ExcelSheet2RecordsetT7
'End Sub


Sub Excel2Recordset2(strFileName As String, strSheetName As String, strFieldName As String, ByRef rs As ADODB.Recordset, intRow As Integer)
'**************************************************
'因為欄位名稱重新定義，所以獨立此副程式，為了跳過第一個欄位名稱
'**************************************************
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
If Dir(strFileName) = "" Then MsgBox "找不到檔案！", vbOKOnly + vbInformation, "Excel2Recordset": Exit Sub '找不到檔案

On Error GoTo err_Handle
Screen.MousePointer = 11

'開啟EXCEL物件
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '選定工作表
        If i = 1 Then strSheetName1 = RTrim(.Sheets(i).Name) '取第一個工作表名稱
        .Sheets(.Sheets(i).Name).Select
    Next
    
    '找不到指定工作表，選用第一個
    If RTrim(.ActiveSheet.Name) <> strSheetName Then .Sheets(strSheetName1).Select
    
    k = intRow '由指定列開始匯入
    
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rs = Nothing: GoTo endsub
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    '寫入Recordset
    Do While Len(RTrim(.Cells(k, 1))) > 0
    rsTmp.AddNew
        For j = 1 To UBound(arrTmp)
            rsTmp(j - 1) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst

    Call OffLineRecordset(rsTmp, rs)
    
    rsTmp.Close: Set rsTmp = Nothing
  
endsub:
Screen.MousePointer = 0
.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

Exit Sub
err_Handle:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Set rs = Nothing

Dim str As String

If err.Number = 3367 Then
    str = "欄位名稱( " & arrTmp(i) & ")重複！"
    
ElseIf err.Number = -2147217887 Then
    str = "請確認儲存格(" & Chr(64 + j) & k & ")，資料是否有誤！"

Else
     str = "Exceed2Recordset"
End If

Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, str)
End Sub
Private Sub cboSheetT7_Click()

On Error GoTo err_Handle

Dim strFilePath As String, strFieldName As String
If Right(filLocalFileT7.Path, 1) <> "\" Then
    strFilePath = filLocalFileT7.Path & "\"
Else
    strFilePath = filLocalFileT7.Path
End If

'strFieldName = "訂單編號" & Chr(9) & "收貨客戶代號" & Chr(9) & "貨號" & Chr(9) & "出貨箱數" & Chr(9) & "出貨包數" & Chr(9) & "單位" & Chr(9) & "出貨日" & Chr(9) & "批次" & Chr(9) & "PO" & Chr(9)

Set rsMainT7 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT7.FileName, cboSheetT7, strFieldName, rsMainT7)

Set dgMainT7.DataSource = rsMainT7

If rsMainT7 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"

Else

'rsMainT7.Sort = "單據號碼,產品品號"

    SetDataGridColWidth Me.Caption, dgMainT7
    MsgBox "此工作表共 " & rsMainT7.RecordCount & "筆明細", 64, "Excel2Recordset"

End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT7_Click()

strTranFileName = filLocalFileT7.Path & "\" & filLocalFileT7.FileName
If Len(RTrim(cboSheetT7)) = 0 Then MsgBox "請選擇工作表", 64, Me.Caption: Exit Sub
If rsMainT7.EOF Or rsMainT7 Is Nothing Then Exit Sub
On Error GoTo err_Handle

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT7.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT7.Enabled = True: dgMainT7.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

'資料檢驗
rsMainT7.MoveFirst
Do While Not rsMainT7.EOF

    '到貨日期檢查
    If Format(myExCharFilter(Trim(rsMainT7("到貨日期"))), "YYYY/MM/DD") < Format(Now, "YYYY/MM/DD") Then: MsgBox "到貨日小於今日，訂單轉入終止!", 16, Me.Caption: Exit Sub
    
    '判斷訂單數是否為0
    If Val(myExCharFilter(Trim(rsMainT7("出貨數量")))) = 0 Then MsgBox "出貨數量為 0，訂單轉入終止!!": Exit Sub
    
    '判斷SKU是否存在
    str_SQL = "select sku,casecnt from gv_skuxpack where sku='" & Trim(rsMainT7("產品品號")) & "' and Storerkey = 'LPSI01' "
'
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        cmdImportT7.Enabled = True: dgMainT7.Enabled = True: Screen.MousePointer = 0
        MsgBox "訂單發現新品號 (" & Trim(rsMainT7("產品品號")) & ") ，訂單轉入終止!!"
        Exit Sub
    End If

    '轉換率檢查
    If Val(rsMainT7("出貨數量")) * Val(tmp_Rs("casecnt")) <> Val(rsMainT7("出貨包數")) Then MsgBox "訂單出貨箱數與出貨包數不符(箱包轉換率與客戶不同)，訂單轉入終止!!", 16, Me.Caption: Exit Sub

    rsMainT7.MoveNext
Loop

Tran_Level = cn.BeginTrans
cmdImportT7.Enabled = False: dgMainT7.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlineNumber As Integer, intTmp As Integer, lngInnerpack As Long
Dim strOrderNo As String, strWHOrderNo As String, strWH As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String, strLot06 As String, strPickMark As String, strPono As String, strTmp As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset

rsMainT7.MoveFirst

Do While Not rsMainT7.EOF
                  
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT7("單據號碼"))) Then
        strOrderNo = UCase(Trim(rsMainT7("單據號碼")))
        int_orderlineNumber = 0
        blDuplicationOrder = False
            
        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select OrderKey from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT7("單據號碼"))) & "' and storerkey = 'LPSI01' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close

            '配送倉別判斷
            strFacility = "佰事達北倉"
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders(OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,BillToKey,CustomerOrderkey,BuyerPO,Notes,UpdateSource,type,addwho,editwho,b_phone1,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT7("單據號碼"))) & "','I','LPSI01',getdate(),'" & myExCharFilter(Trim(rsMainT7("到貨日期"))) & "','" & strFacility & "','" & _
            myExCharFilter(Trim(rsMainT7("客戶代號"))) & "','','','','','','','','" & myExCharFilter(Trim(rsMainT7("採購訂單號碼"))) & "','" & myExCharFilter(Trim(rsMainT7("訂單號碼"))) & "','','" & filLocalFileT7.FileName & "','','" & User_id & "','" & User_id & "','','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT7("單據號碼")) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '訂單重複不增加明細
        If blDuplicationOrder = False Then
            
            '增加明細
            int_orderlineNumber = int_orderlineNumber + 1
            
            intQTY = Val(myExCharFilter(Trim(rsMainT7("出貨包數")))) 'Val(myExCharFilter(Trim(rsMainT7("出貨數量"))))
            strLot06 = "FG01"
'            str_Orderkey = StrPadLeft(int_orderlineNumber, 10, 0)
                        
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,Externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable03,Lottable06, Facility,UOM,Unitprice,notes,updatesource) " & _
            "VALUES ('" & str_Orderkey & "','" & StrPadLeft(Val(int_orderlineNumber), 5, 0) & "','" & StrPadLeft(Val(int_orderlineNumber), 5, 0) & "','" & myExCharFilter(Trim(rsMainT7("單據號碼"))) & "','" & myExCharFilter(Trim(rsMainT7("產品品號"))) & "','LPSI01'," & _
            "'" & intQTY & "','" & intQTY & "','" & Left(myExCharFilter(Trim(rsMainT7("批次"))), 8) & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT7("單位編碼"))) & "','0','','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' and OrderLineNumber = '" & StrPadLeft(Val(int_orderlineNumber), 5, 0) & "' "
                       
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If

        strWHOrderNo = UCase(Trim(rsMainT7("單據號碼")))
        rsMainT7.MoveNext
        
Loop

'補客戶資料
cn.Execute "exec gs_Ordersupdate 'LPSI01' ", RowsAffect, adExecuteNoRecords

'新客戶檢查
str_SQL = "select 貨主=storerkey , 貨主訂單單號=externorderkey , 訂單類別 = priority , 訂單日期=orderdate , 到貨日期=deliverydate , 客戶編號=consigneekey , 檢查日期 = getdate() " & _
        "from orders o " & _
        "Where o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel顯示
    Call Recordset2Excel("缺客戶資料", tmp_Rs)
    If Dir("C:\BEST\LPSI01\缺客戶資料", vbDirectory) = "" Then MkDirs "C:\BEST\LPSI01\缺客戶資料"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LPSI01\缺客戶資料\缺客戶資料_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmdImportT7.Enabled = True: dgMainT7.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
    MsgBox "發現新客戶，訂單轉入中止!", 16, Me.Caption
    rsMainT7.MoveFirst
    Exit Sub
End If

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，TMS單號： " & strOrderKeyS & "~" & str_Orderkey & "，檔案 " & filLocalFileT7.FileName & " 備份於 C:\BEST\LPSI01\Orders\Backup " & strTmp
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，檔案 " & filLocalFileT7.FileName & strTmp)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT7.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 料號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\LPSI01\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\LPSI01\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LPSI01\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''備份至FTP
'If Dir("O:\LPSI01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LPSI01\OrdersBackup"
'FileCopy strTranFileName, "O:\LPSI01\OrdersBackup\" & filLocalFileT7.FileName

'備份檔案
If Dir("C:\BEST\LPSI01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LPSI01\Orders\Backup"
If Dir("C:\BEST\LPSI01\Orders\Backup\" & filLocalFileT7.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LPSI01\Orders\Backup\" & filLocalFileT7.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LPSI01\Orders\Backup\" & mySplit(filLocalFileT7.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT7.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT7.Refresh: cboSheetT7.Clear
Screen.MousePointer = 0: cmdImportT7.Enabled = True: dgMainT7.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-匯入", Me.Caption, "cmdImportT7_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT7.Enabled = True: Screen.MousePointer = 0: dgMainT7.Enabled = True

End Sub

Private Sub dgMainT7_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT7
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT7_Change()
    filLocalFileT7.Path = dirLocalDirT7.Path
End Sub

Private Sub drvLocalDriveT7_Change()

On Error GoTo DriveError
dirLocalDirT7.Path = drvLocalDriveT7.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT7_Click()

On Error GoTo err_Handle
Set rsMainT7 = Nothing: Set dgMainT7.DataSource = rsMainT7
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT7.Path, 1) = "\" Then
    strFilePath = filLocalFileT7.Path
Else
    strFilePath = filLocalFileT7.Path & "\"
End If

If Dir(strFilePath & filLocalFileT7.FileName) = "" Then: filLocalFileT7.Refresh: Exit Sub

cboSheetT7.Clear

If UCase(mySplit(filLocalFileT7.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT7.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT7.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT7.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT7.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT7.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cboSheetT8_Click()
If blDo = True Then Call ExcelSheet2RecordsetT8
End Sub
Private Sub ExcelSheet2RecordsetT8()

Dim strFilePath As String
If Right(filLocalFileT8.Path, 1) <> "\" Then
    strFilePath = filLocalFileT8.Path & "\"
Else
    strFilePath = filLocalFileT8.Path
End If

Set rsMainT8 = New ADODB.Recordset
Call Excel2RecordsetT8(strFilePath & filLocalFileT8.FileName, cboSheetT8, "Document.date" & Chr(9) & "Requested.delivery.d" & Chr(9) & "Sold-to.party" & Chr(9) & "Sold.to.name" & Chr(9) & "Ship.to" & Chr(9) & "Ship.to.name" & Chr(9) & "PO.number" & Chr(9) & "Sales.document" & Chr(9) & "Delivery.number" & Chr(9) & "Billing.Doc.no" & Chr(9) & "Order.Type" & Chr(9) & "Material" & Chr(9) & "Material.Description" & Chr(9) & "Order.Quantity" & Chr(9) & "Order.Confirmed.Quan" & Chr(9) & "Sales.unit" & Chr(9) & "Order.Reason" & Chr(9) & "Description" & Chr(9) & "Reason.for.Rejection" & Chr(9) & "Created.By" & Chr(9) & "Remarks" & Chr(9), rsMainT8)
rsMainT8.Sort = "Sales.document"

Set dgMainT8.DataSource = rsMainT8

If rsMainT8 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"

Else
    SetDataGridColWidth Me.Caption, dgMainT8
    MsgBox "此工作表共 " & rsMainT8.RecordCount & "筆明細", 64, "Excel2Recordset"

End If

End Sub

Sub ExcelSheet2RecordsetT8_old()
On Error GoTo err_Handle
Dim strExcel As String, arrTmp, strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT8.Path, 1) = "\" Then
    strFilePath = filLocalFileT8.Path
Else
    strFilePath = filLocalFileT8.Path & "\"
End If

'建立欄位名稱陣列
arrTmp = Array("Document.date", "Requested.delivery.d", "Sold-to.party", "Sold.to.name", "Ship.to", "Ship.to.name", "PO.number", "Sales.document", "Delivery.number", "Billing.Doc.no", "Order.Type", "Material", "Material.Description", "Order.Quantity", "Order.Confirmed.Quan", "Sales.unit", "Order.Reason", "Description", "Reason.for.Rejection", "Created.By", "Remarks")

'建立 Excel 報表資料庫連接
Dim cnExcel As New ADODB.Connection ':Set cnExcel = New ADODB.Connection
cnExcel.Provider = "Microsoft.Jet.OLEDB.4.0"
cnExcel.ConnectionString = "Data Source=" & filLocalFileT8.Path & "\" & filLocalFileT8.FileName & ";Extended Properties=""Excel 8.0; HDR=YES;ReadOnly=True;"""
cnExcel.Open

Call ReDim_Recordset(tmp_Rs)

tmp_Rs.CursorLocation = 3
tmp_Rs.Open "select * from [" & cboSheetT8 & "$] ", cnExcel ', adOpenStatic, adLockOptimistic
'tmp_rs.Sort = "[Sales document],Material"

Set rsMainT8 = New ADODB.Recordset

'將傳入 tmp_rs 完整複製至 rsMainT8
Dim fldcnt As Integer, reccnt As Double

'建立 Recordset 的 Table 架構 (在記憶體中的 ADO Recordset)
rsMainT8.Fields.Append "編號", adDouble
For fldcnt = 0 To tmp_Rs.Fields.Count - 1
    rsMainT8.Fields.Append arrTmp(fldcnt), tmp_Rs.Fields(fldcnt).Type, tmp_Rs.Fields(fldcnt).DefinedSize
Next fldcnt

With rsMainT8
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With

reccnt = 0
Do While Not tmp_Rs.EOF
If IsNull(tmp_Rs.Fields(0).Value) = True Then GoTo NextRow 'Add by Gemini @20110721
   reccnt = reccnt + 1
   rsMainT8.AddNew
   rsMainT8.Fields(0).Value = reccnt
   For fldcnt = 0 To tmp_Rs.Fields.Count - 1
    rsMainT8.Fields(fldcnt + 1).Value = tmp_Rs.Fields(fldcnt).Value & ""
   Next fldcnt
   rsMainT8.Update
NextRow:
   tmp_Rs.MoveNext
Loop

tmp_Rs.Close: Set tmp_Rs = Nothing

Set dgMainT8.DataSource = rsMainT8: dgMainT8.Visible = False

rsMainT8.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT8
dgMainT8.RowHeight = 300
Screen.MousePointer = 0: dgMainT8.Visible = True
MsgBox "此工作表共 " & rsMainT8.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "工作表開啟"

cnExcel.Close: Set cnExcel = Nothing

Exit Sub
err_Handle:
Set cnExcel = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT8_Click()

strTranFileName = filLocalFileT8.Path & "\" & filLocalFileT8.FileName
If Len(RTrim(cboSheetT8)) = 0 Then MsgBox "請選擇工作表", 64, Me.Caption: Exit Sub
If rsMainT8.EOF Or rsMainT8 Is Nothing Then Exit Sub

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT8.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT8.Enabled = True: dgMainT8.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

On Error GoTo err_Handle

'到貨日期檢查
rsMainT8.MoveFirst
Do While Not rsMainT8.EOF

'    If Format(myExCharFilter(Trim(rsMainT8("Document.date"))), "YYYY/MM/DD") < Format(Now, "YYYY/MM/DD") Then: MsgBox "到貨日小於今日，訂單轉入終止!", 16, Me.Caption: Exit Sub

    '資料檢驗--判斷SKU是否存在
    str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT8("Material")) & "' and Storerkey = 'LNSL01' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        MsgBox "訂單發現新品號 (" & Trim(rsMainT8("Material")) & " )" & Trim(rsMainT8("Material.Description")) & "，訂單轉入終止!!": cmdImportT8.Enabled = True: dgMainT8.Enabled = True: Screen.MousePointer = 0
        Exit Sub
    End If

    rsMainT8.MoveNext
Loop

Tran_Level = cn.BeginTrans
cmdImportT8.Enabled = False: dgMainT8.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlineNumber As Integer, intTmp As Integer, lngCase As Long, lngInnerpack As Long
Dim strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset

rsMainT8.MoveFirst
Do While Not rsMainT8.EOF
    DoEvents: DoEvents
    
'    '資料檢驗--判斷是否訂單數是否為0-->跳下一筆
'    If Trim(rsMainT8("Qty(stckpg.unit)")) = 0 Then
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
'    End If
                  
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT8("Sales.document"))) Then
        strOrderNo = UCase(Trim(rsMainT8("Sales.document")))
        int_orderlineNumber = 0
        blDuplicationOrder = False
            
        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT8("Sales.document"))) & "' and storerkey = 'LNSL01' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
            strFacility = "佰事達北倉"
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,billtokey,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT8("Sales.document"))) & "','R','LNSL01','" & myExCharFilter(Trim(rsMainT8("Document.date"))) & "','" & myExCharFilter(Trim(rsMainT8("Requested.delivery.d"))) & "','" & strFacility & "','" & _
            myExCharFilter(Trim(rsMainT8("Ship.To"))) & "','" & myExCharFilter(Trim(rsMainT8("ship.to.name"))) & "','','','','','','" & myExCharFilter(Trim(rsMainT8("Sold-to.party"))) & "','" & myExCharFilter(Trim(rsMainT8("po.number"))) & "','" & myExCharFilter(Trim(rsMainT8("Remarks"))) & "','" & filLocalFileT8.FileName & "','','" & User_id & "','" & User_id & "','' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT8("Sales.document")) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '訂單重複不增加明細
        If blDuplicationOrder = False Then
            '增加明細
            int_orderlineNumber = int_orderlineNumber + 1
            
            '取商品主檔資料
            str_SQL = "select casecnt = isnull(s.casecnt,0) ,innerpack from gv_skuxpack s where s.storerkey = 'LNSL01' and s.sku = '" & myExCharFilter(Trim(rsMainT8("material"))) & "' "
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            lngCase = tmp_Rs("casecnt")
            lngInnerpack = tmp_Rs("innerpack")
            tmp_Rs.Close
                        
            If Trim(rsMainT8("Sales.unit")) = "CS" Then
                intQTY = Trim(rsMainT8("order.quantity")) * IIf(lngCase = 0, 1, lngCase)
            Else
                intQTY = Trim(rsMainT8("order.quantity"))
'                If myExCharFilter(Trim(rsMainT8("material"))) = "12129314" Then intQTY = intQTY * IIf(lngInnerpack = 0, 1, lngInnerpack)
                If lngInnerpack > 0 Then intQTY = intQTY * lngInnerpack
            End If
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,Externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable03,Lottable06, Facility,UOM,Unitprice,notes) " & _
            "VALUES ('" & str_Orderkey & "','" & int_orderlineNumber & "','" & int_orderlineNumber & "','" & myExCharFilter(Trim(rsMainT8("Sales.document"))) & "','" & myExCharFilter(Trim(rsMainT8("material"))) & "','LNSL01'," & _
            "'" & intQTY & "','" & intQTY & "',null,null,'','" & myExCharFilter(Trim(rsMainT8("Sales.unit"))) & "','0','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT8.MoveNext

Loop

'補客戶資料
cn.Execute "exec gs_ordersupdate 'LNSL01' ", RowsAffect, adExecuteNoRecords

'新客戶檢查
str_SQL = "select 貨主=storerkey , 貨主訂單單號=externorderkey , 訂單類別 = priority , 訂單日期=orderdate , 到貨日期=deliverydate , 客戶編號=consigneekey , 檢查日期 = getdate() " & _
        "from orders o " & _
        "Where o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel顯示
    Call Recordset2Excel("缺客戶資料", tmp_Rs)
    If Dir("C:\BEST\LNSL01\缺客戶資料", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\缺客戶資料"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\缺客戶資料\缺客戶資料_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmdImportT8.Enabled = True: dgMainT8.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
    MsgBox "發現新客戶，訂單轉入中止!", vbOKOnly, Me.Caption
    
    Exit Sub
End If

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細 (共 " & rsMainT8.RecordCount & " 筆明細)，TMS單號： " & strOrderKeyS & "~" & str_Orderkey & "，檔案 " & filLocalFileT8.FileName & " 備份於 C:\BEST\LNSL01\Orders\LNSL01\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細，檔案 " & filLocalFileT8.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT8.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 料號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\LNSL01\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''備份至FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT8.FileName

'備份檔案
If Dir("C:\BEST\LNSL01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\Orders\Backup"
If Dir("C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT8.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT8.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & mySplit(filLocalFileT8.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT8.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT8.Refresh: cboSheetT8.Clear
Screen.MousePointer = 0: cmdImportT8.Enabled = True: dgMainT8.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-匯入", Me.Caption, "cmdImportT8_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT8.Enabled = True: Screen.MousePointer = 0: dgMainT8.Enabled = True

End Sub

Private Sub dgMainT8_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT8
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT8_Change()
    filLocalFileT8.Path = dirLocalDirT8.Path
End Sub
Private Sub drvLocalDriveT8_Change()

On Error GoTo DriveError
dirLocalDirT8.Path = drvLocalDriveT8.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT8_Click()

On Error GoTo err_Handle
Set rsMainT8 = Nothing: Set dgMainT8.DataSource = rsMainT8
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT8.Path, 1) = "\" Then
    strFilePath = filLocalFileT8.Path
Else
    strFilePath = filLocalFileT8.Path & "\"
End If

If Dir(strFilePath & filLocalFileT8.FileName) = "" Then: filLocalFileT8.Refresh: Exit Sub

cboSheetT8.Clear

If UCase(mySplit(filLocalFileT8.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT8.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT8.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT8.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT8.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT8.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cboSheetT9_Click()
If blDo = True Then Call ExcelSheet2RecordsetT9
End Sub
Private Sub ExcelSheet2RecordsetT9()

Dim strFilePath As String
If Right(filLocalFileT9.Path, 1) <> "\" Then
    strFilePath = filLocalFileT9.Path & "\"
Else
    strFilePath = filLocalFileT9.Path
End If

Set rsMainT9 = New ADODB.Recordset
Call Excel2RecordsetT9(strFilePath & filLocalFileT9.FileName, cboSheetT9, "Delivery" & Chr(9) & "Expr1" & Chr(9) & "TO.Number" & Chr(9) & "DlvTy" & Chr(9) & "Sold-to.pt" & Chr(9) & "Ship-To.Pt" & Chr(9) & "Deliv.date" & Chr(9) & "Expr2" & Chr(9) & "Item" & Chr(9) & "Material" & Chr(9) & "Act.qty(dest)" & Chr(9) & "BUn" & Chr(9) & "Actual.qty" & Chr(9) & "AUn" & Chr(9) & "Plnt" & Chr(9) & "SLoc" & Chr(9) & "Batch" & Chr(9) & "SLED/BBD" & Chr(9) & "Route" & Chr(9) & "Notes" & Chr(9), rsMainT9)
rsMainT9.Sort = "[Delivery]"

Set dgMainT9.DataSource = rsMainT9

If rsMainT9 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"

Else
    SetDataGridColWidth Me.Caption, dgMainT9
    MsgBox "此工作表共 " & rsMainT9.RecordCount & "筆明細", 64, "Excel2Recordset"

End If

End Sub
Sub ExcelSheet2RecordsetT9_old()
On Error GoTo err_Handle
Dim strExcel As String, arrTmp, strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT9.Path, 1) = "\" Then
    strFilePath = filLocalFileT9.Path
Else
    strFilePath = filLocalFileT9.Path & "\"
End If

'建立欄位名稱陣列
arrTmp = Array("Delivery", "Expr1", "TO.Number", "DlvTy", "Sold-to.pt", "Ship-To.Pt", "Deliv.date", "Expr2", "Item", "Material", "Act.qty(dest)", "BUn", "Actual.qty", "AUn", "Plnt", "SLoc", "Batch", "SLED/BBD", "Route")

'建立 Excel 報表資料庫連接
Dim cnExcel As New ADODB.Connection ':Set cnExcel = New ADODB.Connection
cnExcel.Provider = "Microsoft.Jet.OLEDB.4.0"
cnExcel.ConnectionString = "Data Source=" & filLocalFileT9.Path & "\" & filLocalFileT9.FileName & ";Extended Properties=""Excel 8.0; HDR=YES;ReadOnly=True;"""
cnExcel.Open

Call ReDim_Recordset(tmp_Rs)

tmp_Rs.CursorLocation = 3
tmp_Rs.Open "select * from [" & cboSheetT9 & "$] where len(rtrim(Delivery)) > 0 ", cnExcel ', adOpenStatic, adLockOptimistic
tmp_Rs.Sort = "[Delivery]"

Set rsMainT9 = New ADODB.Recordset

'將傳入 tmp_rs 完整複製至 rsMainT9
Dim fldcnt As Integer, reccnt As Double

'建立 Recordset 的 Table 架構 (在記憶體中的 ADO Recordset)
rsMainT9.Fields.Append "編號", adDouble
For fldcnt = 0 To 18
    rsMainT9.Fields.Append arrTmp(fldcnt), tmp_Rs.Fields(fldcnt).Type, tmp_Rs.Fields(fldcnt).DefinedSize
Next fldcnt

With rsMainT9
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With

reccnt = 0
Do While Not tmp_Rs.EOF
   reccnt = reccnt + 1
   rsMainT9.AddNew
   rsMainT9.Fields(0).Value = reccnt
   For fldcnt = 0 To 18
    rsMainT9.Fields(fldcnt + 1).Value = tmp_Rs.Fields(fldcnt).Value & ""
   Next fldcnt
   rsMainT9.Update
   tmp_Rs.MoveNext
Loop

tmp_Rs.Close: Set tmp_Rs = Nothing

Set dgMainT9.DataSource = rsMainT9: dgMainT9.Visible = False

rsMainT9.MoveFirst

SetDataGridColWidth Me.Caption, dgMainT9
dgMainT9.RowHeight = 300
Screen.MousePointer = 0: dgMainT9.Visible = True
MsgBox "此工作表共 " & rsMainT9.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "工作表開啟"

cnExcel.Close: Set cnExcel = Nothing

Exit Sub
err_Handle:
Set cnExcel = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT9_Click()

strTranFileName = filLocalFileT9.Path & "\" & filLocalFileT9.FileName
If Len(RTrim(filLocalFileT9.FileName)) = 0 Then MsgBox "請選擇檔案", 64, Me.Caption: Exit Sub
If Len(RTrim(cboSheetT9)) = 0 Then MsgBox "請選擇工作表", 64, Me.Caption: Exit Sub
If rsMainT9.EOF Or rsMainT9 Is Nothing Then Exit Sub

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT9.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT9.Enabled = True: dgMainT9.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

On Error GoTo err_Handle

'到貨日期檢查
rsMainT9.MoveFirst
Do While Not rsMainT9.EOF

If Format(myExCharFilter(Trim(rsMainT9("Deliv.date"))), "YYYY/MM/DD") < Format(Now - 1, "YYYY/MM/DD") Then: MsgBox "到貨日小於今日，訂單轉入終止!", 16, Me.Caption: Exit Sub

    rsMainT9.MoveNext
Loop

Tran_Level = cn.BeginTrans
cmdImportT9.Enabled = False: dgMainT9.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlineNumber As Integer, intTmp As Integer, lngInnerpack As Long
Dim strOrderNo As String, strWHOrderNo As String, strWH As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String, strLot06 As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset

rsMainT9.MoveFirst

Do While Not rsMainT9.EOF

    '資料檢驗--判斷訂單數是否為0
    If Trim(rsMainT9("Act.qty(dest)")) = 0 Then intNotBest = intNotBest + 1: GoTo next1

    DoEvents: DoEvents
    
'    'CT & BEST雙倉出貨判斷
'    If strWHOrderNo = UCase(Trim(rsMainT9("Delivery"))) Then
'        If (strWH = "BEST" And (Trim(rsMainT9("sloc")) = "0001" Or Trim(rsMainT9("sloc")) = "0005")) Or (strWH = "CT" And (Trim(rsMainT9("sloc")) <> "0001" Or Trim(rsMainT9("sloc")) <> "0005") = False) Then
'            cmdImportT9.Enabled = True: Screen.MousePointer = 0: dgMainT9.Enabled = True: cn.RollbackTrans: Tran_Level = 0
'            MsgBox "發現 CT & BEST 雙倉出貨訂單，請與客戶確認訂單正確無誤！", 16, "訂單轉入終止 "
'            Exit Sub
'        End If
'    End If
    
    '是否為佰事達廠別倉別-->跳下一筆
    If Trim(rsMainT9("plnt")) = "1119" And (Trim(rsMainT9("sloc")) = "0001" Or Trim(rsMainT9("sloc")) = "0002" Or Trim(rsMainT9("sloc")) = "0005" Or Trim(rsMainT9("sloc")) = "0007" Or Trim(rsMainT9("sloc")) = "0010") Then
        
        strWH = "CT"
        '檢查訂單量是否出現小數點
        If Val(rsMainT9("Act.qty(dest)")) <> Round(Val(rsMainT9("Act.qty(dest)")), 0) Then
            cmdImportT9.Enabled = True: Screen.MousePointer = 0: dgMainT9.Enabled = True: cn.RollbackTrans: Tran_Level = 0
            MsgBox "發現訂單量出現小數點，請與客戶確認正確訂單量！", 16, "訂單轉入終止 "
            Exit Sub
        End If
        
        '資料檢驗--判斷SKU是否存在
        str_SQL = "select sku,innerpack from gv_skuxpack where sku='" & Trim(rsMainT9("Material")) & "' and Storerkey = 'LNSL01' "
        
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
            cn.RollbackTrans: Tran_Level = 0: cmdImportT9.Enabled = True: dgMainT9.Enabled = True: Screen.MousePointer = 0
            MsgBox "訂單發現新品號 (" & Trim(rsMainT9("Material")) & ") ，訂單轉入終止!!"
            Exit Sub
        End If
        lngInnerpack = tmp_Rs("Innerpack")

    Else
        intNotBest = intNotBest + 1
        strWH = "BEST"
        GoTo next1

    End If
                  
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT9("Delivery"))) Then
        strOrderNo = UCase(Trim(rsMainT9("Delivery")))
        int_orderlineNumber = 0
        blDuplicationOrder = False
            
        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT9("Delivery"))) & "' and storerkey = 'LNSL01' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
            strFacility = "佰事達北倉"
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,BillToKey,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT9("Delivery"))) & "','I','LNSL01', '" & myExCharFilter(Trim(rsMainT9("ExPr1"))) & "',cast('" & myExCharFilter(Trim(rsMainT9("Deliv.date"))) & "' as datetime)+1,'" & strFacility & "','" & _
            myExCharFilter(Trim(rsMainT9("Route"))) & "','','','','','','','" & myExCharFilter(Trim(rsMainT9("Sold-to.pt"))) & "','" & myExCharFilter(Trim(rsMainT9("to.number"))) & "','" & myExCharFilter(Trim(rsMainT9("Notes"))) & "','" & filLocalFileT9.FileName & "','','" & User_id & "','" & User_id & "','' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT9("Delivery")) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '訂單重複不增加明細
        If blDuplicationOrder = False Then
        
            '增加明細
            int_orderlineNumber = int_orderlineNumber + 1
            
            '數量轉換
            intQTY = Trim(rsMainT9("Act.qty(dest)"))
'            If Trim(rsMainT9("Material")) = "12129314" Then intQTY = intQTY * IIf(lngInnerpack = 0, 1, lngInnerpack)
            If lngInnerpack > 0 Then intQTY = intQTY * lngInnerpack
            
            '倉別轉換
            strLot06 = myExCharFilter(Trim(rsMainT9("sloc")))
            
            If strLot06 = "0001" Then
               strLot06 = "R01"
            ElseIf strLot06 = "0002" Then
               strLot06 = "R02"
            ElseIf strLot06 = "0005" Then
               strLot06 = "R05"
            ElseIf strLot06 = "0007" Then
               strLot06 = "R07"
            ElseIf strLot06 = "0010" Then
               strLot06 = "R10"
            End If
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,Externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable03,Lottable06, Facility,UOM,Unitprice,notes,updatesource) " & _
            "VALUES ('" & str_Orderkey & "','" & int_orderlineNumber & "','" & myExCharFilter(Trim(rsMainT9("Item"))) & "','" & myExCharFilter(Trim(rsMainT9("Delivery"))) & "','" & myExCharFilter(Trim(rsMainT9("material"))) & "','LNSL01'," & _
            "'" & intQTY & "','" & intQTY & "','" & myExCharFilter(Trim(rsMainT9("Batch"))) & "','" & strLot06 & "','','" & myExCharFilter(Trim(rsMainT9("BUn"))) & "','0','','')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:

        If (Trim(rsMainT9("sloc")) = "0001" Or Trim(rsMainT9("sloc")) = "0005") Then
            strWH = "CT"
        Else
            strWH = "BEST"
        End If

        strWHOrderNo = UCase(Trim(rsMainT9("Delivery")))
        rsMainT9.MoveNext
        
Loop

'補客戶資料
cn.Execute "exec gs_ordersupdate 'LNSL01' ", RowsAffect, adExecuteNoRecords

'新客戶檢查
str_SQL = "select 貨主=storerkey , 貨主訂單單號=externorderkey , 訂單類別 = priority , 訂單日期=orderdate , 到貨日期=deliverydate , 客戶編號=consigneekey , 檢查日期 = getdate() " & _
        "from orders o " & _
        "Where o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel顯示
    Call Recordset2Excel("缺客戶資料", tmp_Rs)
    If Dir("C:\BEST\LNSL01\缺客戶資料", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\缺客戶資料"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\缺客戶資料\缺客戶資料_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmdImportT9.Enabled = True: dgMainT9.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
    MsgBox "發現新客戶，訂單轉入中止!", 16, Me.Caption
    
    Exit Sub
End If

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細 (共 " & rsMainT9.RecordCount & " 筆明細)，TMS單號： " & strOrderKeyS & "~" & str_Orderkey & "，檔案 " & filLocalFileT9.FileName & " 備份於 C:\BEST\LNSL01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細，檔案 " & filLocalFileT9.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT9.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 料號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\LNSL01\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''備份至FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFilet9.FileName

'備份檔案
If Dir("C:\BEST\LNSL01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\Orders\Backup"
If Dir("C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT9.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT9.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & mySplit(filLocalFileT9.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT9.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT9.Refresh: cboSheetT9.Clear
Screen.MousePointer = 0: cmdImportT9.Enabled = True: dgMainT9.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-匯入", Me.Caption, "cmdImportt9_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT9.Enabled = True: Screen.MousePointer = 0: dgMainT9.Enabled = True

End Sub

Private Sub dgMainT9_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT9
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT9_Change()
    filLocalFileT9.Path = dirLocalDirT9.Path
End Sub
Private Sub drvLocalDriveT9_Change()

On Error GoTo DriveError
dirLocalDirT9.Path = drvLocalDriveT9.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT9_Click()

On Error GoTo err_Handle
Set rsMainT9 = Nothing: Set dgMainT9.DataSource = rsMainT9
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT9.Path, 1) = "\" Then
    strFilePath = filLocalFileT9.Path
Else
    strFilePath = filLocalFileT9.Path & "\"
End If

If Dir(strFilePath & filLocalFileT9.FileName) = "" Then: filLocalFileT9.Refresh: Exit Sub

cboSheetT9.Clear

If UCase(mySplit(filLocalFileT9.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT9.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT9.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT9.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT9.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT9.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cboSheetT10_Click()
If blDo = True Then Call ExcelSheet2Recordsett10
End Sub

Sub ExcelSheet2Recordsett10()
On Error GoTo err_Handle
Dim strExcel As String, arrTmp, strFilePath As String, i As Integer, strQty As String

'確認路徑是否帶"\"
If Right(filLocalFileT10.Path, 1) = "\" Then
    strFilePath = filLocalFileT10.Path
Else
    strFilePath = filLocalFileT10.Path & "\"
End If

'建立欄位名稱陣列
arrTmp = Array("Ship.to(Return.Code)", "Ship.to.name", "PO.number", "ZRR收件日", "品號", "件數", "Sales完成簽核", "E-mail.to僑泰", "備註")

'建立 Excel 報表資料庫連接
Dim cnExcel As New ADODB.Connection ':Set cnExcel = New ADODB.Connection
cnExcel.Provider = "Microsoft.Jet.OLEDB.4.0"
cnExcel.ConnectionString = "Data Source=" & filLocalFileT10.Path & "\" & filLocalFileT10.FileName & ";Extended Properties=""Excel 8.0; HDR=YES;ReadOnly=True;"""
cnExcel.Open

Call ReDim_Recordset(tmp_Rs)

tmp_Rs.CursorLocation = 3
tmp_Rs.Open "select * from [" & cboSheetT10 & "$] ", cnExcel ', adOpenStatic, adLockOptimistic
'tmp_rs.Sort = "[Ship to(Return Code)]"

Set rsMainT10 = New ADODB.Recordset

'將傳入 tmp_rs 完整複製至 rsMaint10
Dim fldcnt As Integer, reccnt As Double

'建立 Recordset 的 Table 架構 (在記憶體中的 ADO Recordset)
rsMainT10.Fields.Append "編號", adDouble
For fldcnt = 0 To 8
'    rsMainT10.Fields.Append arrTmp(fldcnt), tmp_rs.Fields(fldcnt).Type, tmp_rs.Fields(fldcnt).DefinedSize
rsMainT10.Fields.Append arrTmp(fldcnt), adVarChar, 255
Next fldcnt

With rsMainT10
     .CursorType = adOpenStatic
     .LockType = adLockOptimistic
     .Open    '不需連接物件
End With

reccnt = 0
tmp_Rs.MoveFirst
Do While Not tmp_Rs.EOF
    If IsNull(tmp_Rs("Ship to(Return Code)")) = False Or Len(Trim(tmp_Rs("Ship to(Return Code)"))) > 0 Then
        reccnt = reccnt + 1
        rsMainT10.AddNew
        rsMainT10.Fields(0).Value = reccnt
        For fldcnt = 0 To 8
         rsMainT10.Fields(fldcnt + 1).Value = Trim(tmp_Rs.Fields(fldcnt)) & ""
        Next fldcnt
        rsMainT10.Update
    End If
   tmp_Rs.MoveNext
Loop

tmp_Rs.Close: Set tmp_Rs = Nothing

rsMainT10.MoveFirst

'註記合併儲存格
i = 1
strQty = rsMainT10("件數")
Do While Not rsMainT10.EOF
    
    If rsMainT10("件數") = "" Then
        i = i + 1
        rsMainT10("件數") = "**" & i & "/" & strQty
        
    Else
        strQty = rsMainT10("件數")
        i = 1
    End If

rsMainT10.MoveNext
Loop

i = 0
'統計分筆數與總件數
rsMainT10.MoveLast
Do While Not rsMainT10.BOF

        If Left(rsMainT10("件數"), 2) = "**" Then
        rsMainT10("備註") = "(" & Val(Replace(mySplit(rsMainT10("件數"), "/", 0), "**", "")) & "/" & Val(Replace(mySplit(rsMainT10("件數"), "/", 0), "**", "")) + i & ")共" & mySplit(rsMainT10("件數"), "/", -1) & "件，" & rsMainT10("備註")
        i = i + 1
        rsMainT10("件數") = 1
    Else
        If i <> 0 Then rsMainT10("備註") = "(1/" & i + 1 & ")共" & rsMainT10("件數") & "件，" & rsMainT10("備註")
        strQty = rsMainT10("件數")
        i = 0
    End If

rsMainT10.MovePrevious
Loop

Set dgMainT10.DataSource = rsMainT10: dgMainT10.Visible = False


SetDataGridColWidth Me.Caption, dgMainT10
dgMainT10.RowHeight = 300
Screen.MousePointer = 0: dgMainT10.Visible = True
MsgBox "此工作表共 " & rsMainT10.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "工作表開啟"

cnExcel.Close: Set cnExcel = Nothing

Exit Sub
err_Handle:
Set cnExcel = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT10_Click()

On Error GoTo err_Handle
strTranFileName = filLocalFileT10.Path & "\" & filLocalFileT10.FileName
If Len(RTrim(cboSheetT10)) = 0 Then MsgBox "請選擇工作表", 64, Me.Caption: Exit Sub
If rsMainT10.EOF Or rsMainT10 Is Nothing Then Exit Sub

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT10.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT10.Enabled = True: dgMainT10.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

'到貨日期檢查
rsMainT10.MoveFirst
'Do While Not rsMainT10.EOF
'
'    If myExCharFilter(Trim(rsMainT10("ZRR收件日"))) < Format(Now - 1, "YYYYMMDD") Then: MsgBox "到貨日小於今日，訂單轉入終止!", 16, Me.Caption: Exit Sub
'    rsMainT10.MoveNext
'Loop

Tran_Level = cn.BeginTrans
cmdImportT10.Enabled = False: dgMainT10.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlineNumber As Integer, intTmp As Integer
Dim strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset

rsMainT10.MoveFirst
Do While Not rsMainT10.EOF
    DoEvents: DoEvents
    
    '資料檢驗--判斷是否品號是否為D-->跳下一筆
'    If UCase(Trim(rsMainT10("品號"))) <> "D" Then
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
        '資料檢驗--判斷SKU是否存在
        str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT10("品號")) & "' and Storerkey = 'LNSL01' "
        
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
            MsgBox "訂單發現新品號 (" & Trim(rsMainT10("品號")) & ")" & "，訂單轉入終止!!": cmdImportT10.Enabled = True: dgMainT10.Enabled = True: Screen.MousePointer = 0
            cn.RollbackTrans: Tran_Level = 0
            Exit Sub
        End If

'    End If
                  
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT10("PO.number"))) Then
        strOrderNo = UCase(Trim(rsMainT10("PO.number")))
        int_orderlineNumber = 0
        blDuplicationOrder = False
            
        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT10("PO.number"))) & "' and storerkey = 'LNSL01' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
            strFacility = "佰事達北倉"
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,billtokey,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT10("PO.number"))) & "','R','LNSL01','" & myExCharFilter(Trim(rsMainT10("ZRR收件日"))) & "',cast(" & "'" & myExCharFilter(Trim(rsMainT10("ZRR收件日"))) & "' as datetime)+1,'" & strFacility & "','" & _
            myExCharFilter(Trim(rsMainT10("Ship.to(Return.Code)"))) & "','" & myExCharFilter(Trim(rsMainT10("ship.to.name"))) & "','','','','','','','','" & myExCharFilter(Trim(rsMainT10("備註"))) & "','" & filLocalFileT10.FileName & "','','" & User_id & "','" & User_id & "','' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT10("PO.number")) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '訂單重複不增加明細
        If blDuplicationOrder = False Then
            '增加明細
            int_orderlineNumber = int_orderlineNumber + 1

            intQTY = Trim(rsMainT10("件數"))
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,Externlineno,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable03,Lottable06, Facility,UOM,Unitprice,notes) " & _
            "VALUES ('" & str_Orderkey & "','" & int_orderlineNumber & "','" & int_orderlineNumber & "','" & myExCharFilter(Trim(rsMainT10("PO.number"))) & "','" & myExCharFilter(Trim(rsMainT10("品號"))) & "','LNSL01'," & _
            "'" & intQTY & "','" & intQTY & "',null,null,'','EA','0','" & myExCharFilter(Trim(rsMainT10("備註"))) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT10.MoveNext

Loop

'補客戶資料
cn.Execute "exec gs_ordersupdate 'LNSL01' ", RowsAffect, adExecuteNoRecords

'新客戶檢查
str_SQL = "select 貨主=storerkey , 貨主訂單單號=externorderkey , 訂單類別 = priority , 訂單日期=orderdate , 到貨日期=deliverydate , 客戶編號=consigneekey , 檢查日期 = getdate() " & _
        "from orders o " & _
        "Where o.b_phone2 Is Null " & _
        "and o.consigneekey not in (select trp01m.consigneekey from trp01m where trp01m.storerkey = o.storerkey) "
        
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then

    'Excel顯示
    Call Recordset2Excel("缺客戶資料", tmp_Rs)
    If Dir("C:\BEST\LNSL01\缺客戶資料", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\缺客戶資料"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\缺客戶資料\缺客戶資料_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing: tmp_Rs.Close
    
    cmdImportT10.Enabled = True: dgMainT10.Enabled = True: Screen.MousePointer = 0: cn.RollbackTrans: Tran_Level = 0
    MsgBox "發現新客戶，訂單轉入中止!", vbOKOnly, Me.Caption
    
    Exit Sub
End If

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細 (共 " & rsMainT10.RecordCount & " 筆明細)，TMS單號： " & strOrderKeyS & "~" & str_Orderkey & "，檔案 " & filLocalFileT10.FileName & " 備份於 C:\BEST\LNSL01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細，檔案 " & filLocalFileT10.FileName)
    cmd_Import.Enabled = True: Screen.MousePointer = 0: dgMainT10.Enabled = True
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT10.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 料號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\LNSL01\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNSL01\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''備份至FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFilet10.FileName

'備份檔案
If Dir("C:\BEST\LNSL01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LNSL01\Orders\Backup"
If Dir("C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT10.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & filLocalFileT10.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LNSL01\Orders\Backup\" & mySplit(filLocalFileT10.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT10.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT10.Refresh: cboSheetT10.Clear
Screen.MousePointer = 0: cmdImportT10.Enabled = True: dgMainT10.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-匯入", Me.Caption, "cmdImportt10_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT10.Enabled = True: Screen.MousePointer = 0: dgMainT10.Enabled = True

End Sub

Private Sub dgMainT10_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT10
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT10_Change()
    filLocalFileT10.Path = dirLocalDirT10.Path
End Sub
Private Sub drvLocalDriveT10_Change()

On Error GoTo DriveError
dirLocalDirT10.Path = drvLocalDriveT10.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT10_Click()

On Error GoTo err_Handle
Set rsMainT10 = Nothing: Set dgMainT10.DataSource = rsMainT10
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT10.Path, 1) = "\" Then
    strFilePath = filLocalFileT10.Path
Else
    strFilePath = filLocalFileT10.Path & "\"
End If

If Dir(strFilePath & filLocalFileT10.FileName) = "" Then: filLocalFileT10.Refresh: Exit Sub

cboSheetT10.Clear

If UCase(mySplit(filLocalFileT10.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT10.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT10.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT10.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT10.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT10.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cboSheetT3_Click()
If blDo = True Then Call ExcelSheet2RecordsetT3
End Sub

Sub ExcelSheet2RecordsetT3()
On Error GoTo err_Handle
Screen.MousePointer = 11

Dim strFilePath As String
If Right(filLocalFileT3.Path, 1) <> "\" Then
    strFilePath = filLocalFileT3.Path & "\"
Else
    strFilePath = filLocalFileT3.Path
End If

Set rsMainT3 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT3.FileName, cboSheetT3, "", rsMainT3)

rsMainT3.Sort = "訂單編號"

Set dgMainT3.DataSource = rsMainT3

If rsMainT3 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT3
    MsgBox "此工作表共 " & rsMainT3.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
End If

Screen.MousePointer = 0
Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT3_Click()

strTranFileName = filLocalFileT3.Path & "\" & filLocalFileT3.FileName
If Len(RTrim(cboSheetT3)) = 0 Then MsgBox "請選擇工作表", 64, Me.Caption: Exit Sub
If rsMainT3.EOF Or rsMainT3 Is Nothing Then Exit Sub

On Error GoTo err_Handle
'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT3.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT3.Enabled = True: dgMainT3.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

''到貨日期檢查
'rsMainT3.MoveFirst
'Do While Not rsMainT3.EOF
'
'    If myExCharFilter(Trim(rsMainT3("應送日期")))< Format(Now, "YYYY/MM/DD") Then: MsgBox "到貨日小於今日，訂單轉入終止!", 16, Me.Caption: Exit Sub
'
'    rsMainT3.MoveNext
'Loop

Tran_Level = cn.BeginTrans: cmdImportT3.Enabled = False: dgMainT3.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, arrTmp, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
            
'取最後客戶編號
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select top 1 substring(consigneekey,7,20) as consigneekey from trp01m where storerkey = 'LFYY01' and left(consigneekey,6) = 'LFYY00' order by consigneekey desc "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))

tmp_Rs.Close

rsMainT3.MoveFirst
Do While Not rsMainT3.EOF
    DoEvents: DoEvents
    
    '資料檢驗--判斷是否屬佰事達訂單-->跳下一筆
'    If UCase(Trim(rsMainT3("倉庫"))) <> "A02" And UCase(Trim(rsMainT3("倉庫"))) <> "A02A" And UCase(Trim(rsMainT3("倉庫"))) <> "A02C" Then
''        MsgBox "客戶單號：" & Trim(rsMainT4("銷貨單號")) & "( " & Trim(rsMainT4("倉庫")) & " )" & vbCrLf & "請通知客戶，確認該筆訂單是否有誤!?", vbOKOnly, "非佰事達之訂單不轉入"
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
        '資料檢驗--判斷SKU是否存在
        str_SQL = "select casecnt = case when isnull(casecnt,0) = 0 then 1 else casecnt end from gv_skuxpack where sku='" & Trim(rsMainT3("捷盟品號")) & "' and Storerkey = 'LFYY01' "
        
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
            MsgBox "訂單發現新品號 (" & Trim(rsMainT3("捷盟品號")) & " ) " & Trim(rsMainT3("商品名稱")) & "，訂單轉入終止!!"
             cmdImportT3.Enabled = True: dgMainT3.Enabled = True: Screen.MousePointer = 0
            cn.RollbackTrans: Tran_Level = 0
            tmp_Rs.Close
            Exit Sub
        End If
        tmp_Rs.Close
        
'    End If
                  
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT3("訂單編號"))) Then
        strOrderNo = UCase(Trim(rsMainT3("訂單編號")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '檢查是否有此客戶名稱
        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LFYY01' and rtrim(full_name) = '" & myExCharFilter(Trim(rsMainT3("DC名稱"))) & "' order by len(consigneekey),consigneekey "
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        
        If tmp_Rs.EOF Then
            '無此客戶名稱則新增
            intTmp = intTmp + 1
            strConsigneeKey = "LFYY" & Format(intTmp, "000000")
            
            '新增客戶主檔
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address) " & _
            " values('LFYY01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT3("DC名稱"))) & "','" & myExCharFilter(Trim(rsMainT3("DC名稱"))) & "','" & myExCharFilter(Trim(rsMainT3("訂貨人員"))) & "','','" & myExCharFilter(Trim(rsMainT3("DC住址"))) & "' ) ", RowsAffect, adExecuteNoRecords
'
            '紀錄新增之客戶編號
            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
        Else
            '比對聯絡人、電話與到貨地址是否相符
            Call Confirm_Recordset_Closed(rsTmp)
            str_SQL = "select * from trp01m " & _
                        "where storerkey = 'LFYY01' and full_name = '" & myExCharFilter(Trim(rsMainT3("DC名稱"))) & "' " & _
                        "and rtrim(contact) = '" & myExCharFilter(Trim(rsMainT3("訂貨人員"))) & "' " & _
                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT3("DC住址"))) & "' "
            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn
            
            If rsTmp.EOF Then
                '聯絡人與到貨地址不符
                intTmp = intTmp + 1
                strConsigneeKey = "LFYY" & Format(intTmp, "000000")
                
                '新增客戶主檔
                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address) " & _
                " values('LFYY01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT3("DC名稱"))) & "','" & myExCharFilter(Trim(rsMainT3("DC名稱"))) & "','" & myExCharFilter(Trim(rsMainT3("訂貨人員"))) & "','','" & myExCharFilter(Trim(rsMainT3("DC住址"))) & "' ) ", RowsAffect, adExecuteNoRecords
                 
                '紀錄新增之客戶編號
                strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
            Else '相符沿用舊客編
                strConsigneeKey = Trim(rsTmp("consigneekey"))
                blCustomerMatch = True

            End If
        End If
        tmp_Rs.Close
    
        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT3("訂單編號"))) & "' and storerkey = 'LFYY01' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
            strFacility = "佰事達北倉"
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT3("訂單編號"))) & "','I','LFYY01','" & myExCharFilter(Trim(rsMainT3("訂貨日期"))) & "','" & myExCharFilter(Trim(rsMainT3("到貨日期"))) & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT3("DC名稱"))) & "','" & myExCharFilter(Trim(rsMainT3("訂貨人員"))) & "','','','" & myExCharFilter(Trim(GetWord(rsMainT3("DC住址"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT3("DC住址"), intPointer, 45))) & "','','','" & filLocalFileT3.FileName & "','','" & User_id & "','" & User_id & "','','" & myExCharFilter(Trim(rsMainT3("DC編號"))) & "' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LFYY01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT3("訂單編號")) & "','"
            blDuplicationOrder = True
            
        End If
    End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
        
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Val(rsMainT3("訂貨數"))
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes,addwho,editwho)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & Trim(rsMainT3("訂單編號")) & "','" & Trim(rsMainT3("捷盟品號")) & "','LFYY01'," & _
            "'" & intQTY & "','" & intQTY & "','R01','','EA','0','','" & User_id & "','" & User_id & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT3.MoveNext
Loop

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細 (共 " & rsMainT3.RecordCount & " 筆明細)，TMS單號： " & strOrderKeyS & "~" & str_Orderkey & "，檔案 " & filLocalFileT3.FileName & " 備份於 C:\BEST\LFYY01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細，檔案 " & filLocalFileT3.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT3.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 料號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\LFYY01\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\LFYY01\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LFYY01\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''備份至FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT4.FileName

'備份檔案
If Dir("C:\BEST\LFYY01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LFYY01\Orders\Backup"
If Dir("C:\BEST\LFYY01\Orders\Backup\" & filLocalFileT3.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LFYY01\Orders\Backup\" & filLocalFileT3.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LFYY01\Orders\Backup\" & mySplit(filLocalFileT3.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT3.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT3.Refresh: cboSheetT3.Clear
Screen.MousePointer = 0: cmdImportT3.Enabled = True: dgMainT3.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-匯入", Me.Caption, "cmdImportT3_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT3.Enabled = True: Screen.MousePointer = 0: dgMainT3.Enabled = True

End Sub

Private Sub dgMainT3_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT3
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT3_Change()
    filLocalFileT3.Path = dirLocalDirT3.Path
End Sub
Private Sub drvLocalDriveT3_Change()

On Error GoTo DriveError
dirLocalDirT3.Path = drvLocalDriveT3.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT3_Click()

On Error GoTo err_Handle
Set rsMainT3 = Nothing: Set dgMainT3.DataSource = rsMainT10
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT3.Path, 1) = "\" Then
    strFilePath = filLocalFileT3.Path
Else
    strFilePath = filLocalFileT3.Path & "\"
End If

If Dir(strFilePath & filLocalFileT3.FileName) = "" Then: filLocalFileT3.Refresh: Exit Sub

cboSheetT3.Clear

If UCase(mySplit(filLocalFileT3.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT3.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT3.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT3.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT3.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT3.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cboSheetT11_Click()

On Error GoTo err_Handle
Dim str As String, strFieldName As String, strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT11.Path, 1) = "\" Then
    strFilePath = filLocalFileT11.Path
Else
    strFilePath = filLocalFileT11.Path & "\"
End If

'建立欄位名稱陣列
strFieldName = "客戶" & Chr(9) & "統編" & Chr(9) & "銷貨日期" & Chr(9) & "銷貨單號" & Chr(9) & "聯絡人" & Chr(9) & "電話" & Chr(9) & "送貨地址" & Chr(9) & "發票號碼" & Chr(9) & "業務員" & Chr(9) & "備註" & Chr(9) & "品號" & Chr(9) & "品名" & Chr(9) & "倉庫" & Chr(9) & "數量" & Chr(9) & "單位" & Chr(9) & "前置單據/備註/客戶單號" & Chr(9)

If Right(filLocalFileT11.Path, 1) <> "\" Then
    strFilePath = filLocalFileT11.Path & "\"
Else
    strFilePath = filLocalFileT11.Path
End If

Set rsMainT11 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT11.FileName, cboSheetT11, strFieldName, rsMainT11)

Set dgMainT11.DataSource = rsMainT11

If rsMainT11 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT11
    MsgBox "此工作表共 " & rsMainT11.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT11_Click()

On Error GoTo err_Handle
strTranFileName = filLocalFileT11.Path & "\" & filLocalFileT11.FileName
If Len(RTrim(cboSheetT11)) = 0 Then MsgBox "請選擇工作表", 64, Me.Caption: Exit Sub
If rsMainT11.EOF Or rsMainT11 Is Nothing Then Exit Sub

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select * from orders where rtrim(updatesource)='" & filLocalFileT11.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT11.Enabled = True: dgMainT11.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp
rsMainT11.MoveFirst
Do While Not rsMainT11.EOF

    '到貨日期檢查
    arrTmp = Split(Trim(rsMainT11("銷貨日期")), "/")
    If Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then: MsgBox "到貨日小於今日，訂單轉入終止!", 16, Me.Caption: Exit Sub
            
    '數量檢查
    If Val(rsMainT11("數量")) < 0 Then
        MsgBox "訂單數量小於1，" & Trim(rsMainT11("品號")) & "-" & Trim(rsMainT11("品名")) & "(" & Trim(rsMainT11("數量")) & Trim(rsMainT11("單位")) & ")，訂單轉入終止!!", , "訂單檔匯入": Exit Sub
        Exit Sub
    End If
                 
    rsMainT11.MoveNext
Loop

Tran_Level = cn.BeginTrans: cmdImportT11.Enabled = False: dgMainT11.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
            
'取最後客戶編號
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = 'LNIP01' and left(consigneekey,4) = 'LNIP' order by consigneekey desc "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))

tmp_Rs.Close

rsMainT11.MoveFirst
Do While Not rsMainT11.EOF
    DoEvents: DoEvents
    
'    '資料檢驗--判斷是否屬佰事達訂單-->跳下一筆
'    If UCase(Trim(rsMainT11("倉庫"))) = "" Then
''        MsgBox "客戶單號：" & Trim(rsMainT4("銷貨單號")) & "( " & Trim(rsMainT4("倉庫")) & " )" & vbCrLf & "請通知客戶，確認該筆訂單是否有誤!?", vbOKOnly, "非佰事達之訂單不轉入"
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
        '資料檢驗--判斷SKU是否存在
    str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT11("品號")) & "' and Storerkey = 'LNIP01' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        cn.RollbackTrans: Tran_Level = 0
        MsgBox "訂單發現新品號 (" & Trim(rsMainT11("品號")) & " ) " & Trim(rsMainT11("品名")) & "，訂單轉入終止!!": cmdImportT11.Enabled = True: dgMainT11.Enabled = True: Screen.MousePointer = 0
        Exit Sub
    End If

'    End If
                  
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT11("銷貨單號"))) Then
        strOrderNo = UCase(Trim(rsMainT11("銷貨單號")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '檢查是否有此客戶名稱
        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = 'LNIP01' and rtrim(full_name) = '" & myExCharFilter(Trim(rsMainT11("客戶"))) & "' order by len(consigneekey),consigneekey "
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        
        If tmp_Rs.EOF Then
            '無此客戶名稱則新增
            intTmp = intTmp + 1
            strConsigneeKey = "LNIP" & Format(intTmp, "000000")
            
            '新增客戶主檔
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address) " & _
            " values('LNIP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT11("客戶"))) & "','" & myExCharFilter(Trim(rsMainT11("客戶"))) & "','" & myExCharFilter(Trim(rsMainT11("聯絡人"))) & "','" & myExCharFilter(Trim(rsMainT11("電話"))) & "','" & myExCharFilter(Trim(rsMainT11("送貨地址"))) & "' ) ", RowsAffect, adExecuteNoRecords
'
            '紀錄新增之客戶編號
            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
        Else
            '比對聯絡人、電話與到貨地址是否相符
            Call Confirm_Recordset_Closed(rsTmp)
            str_SQL = "select * from trp01m " & _
                        "where storerkey = 'LNIP01' and full_name = '" & myExCharFilter(Trim(rsMainT11("客戶"))) & "' " & _
                        "and rtrim(contact) = '" & myExCharFilter(Trim(rsMainT11("聯絡人"))) & "' " & _
                        "and rtrim(phone) = '" & myExCharFilter(Trim(rsMainT11("電話"))) & "' " & _
                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT11("送貨地址"))) & "' "
            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn
            
            If rsTmp.EOF Then
                '聯絡人、電話與到貨地址不符
                intTmp = intTmp + 1
                strConsigneeKey = "LNIP" & Format(intTmp, "000000")
                
                '新增客戶主檔
                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes) " & _
                " values('LNIP01','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT11("客戶"))) & "','" & myExCharFilter(Trim(rsMainT11("客戶"))) & "','" & myExCharFilter(Trim(rsMainT11("聯絡人"))) & "','" & myExCharFilter(Trim(rsMainT11("電話"))) & "','" & myExCharFilter(Trim(rsMainT11("送貨地址"))) & "','" & myExCharFilter(Trim(tmp_Rs("notes"))) & "' ) ", RowsAffect, adExecuteNoRecords
                 
                '紀錄新增之客戶編號
                strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
            Else '相符沿用舊客編
                strConsigneeKey = Trim(rsTmp("consigneekey"))
                blCustomerMatch = True

            End If
            rsTmp.Close
        End If
        tmp_Rs.Close
    
'        '資料檢驗--判斷訂單是否重複，重複不增加
'        Call Confirm_Recordset_Closed(tmp_rs)
'        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT11("銷貨單號"))) & "' and storerkey = 'LNIP01' and isnull(type,'') <> '刪單' "
'        tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'        If tmp_rs.EOF Then
    
            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
            strFacility = Trim(rsMainT11("倉庫"))
            strFacility = "佰事達北倉"
            arrTmp = Split(Trim(rsMainT11("銷貨日期")), "/")
            strOrderDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT11("銷貨單號"))) & "','I','LNIP01','" & strOrderDate & "','" & CDate(strOrderDate) + 1 & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT11("客戶"))) & "','" & myExCharFilter(Trim(rsMainT11("聯絡人"))) & "','" & myExCharFilter(Trim(rsMainT11("業務員"))) & "','" & myExCharFilter(Trim(rsMainT11("統編"))) & "','" & myExCharFilter(Trim(rsMainT11("電話"))) & "','','" & myExCharFilter(Trim(GetWord(rsMainT11("送貨地址"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT11("送貨地址"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT11("備註"))) & "','" & filLocalFileT11.FileName & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT11("發票號碼"))) & "' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = 'LNIP01') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
'        Else
'            '訂單重複
'            Call FTPlog("訂單重複" & str_SQL)
'            '紀錄重複
'            strReOrderkey = strReOrderkey & Trim(rsMainT11("銷貨單號")) & "','"
'            blDuplicationOrder = True
'
'        End If
    End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Trim(rsMainT11("數量"))
            strLot06 = IIf(UCase(Trim(rsMainT11("倉庫"))) = "A06", "A06-S", Trim(rsMainT11("倉庫")))
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT11("銷貨單號"))) & "','" & myExCharFilter(Trim(rsMainT11("品號"))) & "','LNIP01'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','" & myExCharFilter(rsMainT11("倉庫")) & "','" & myExCharFilter(Trim(rsMainT11("單位"))) & "','0','" & myExCharFilter(Trim(rsMainT11("前置單據/備註/客戶單號"))) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT11.MoveNext
Loop

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細 (共 " & rsMainT11.RecordCount & " 筆明細)，TMS單號： " & strOrderKeyS & "~" & str_Orderkey & "，檔案 " & filLocalFileT11.FileName & " 備份於 C:\BEST\LNIP01\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細，檔案 " & filLocalFileT11.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT11.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 料號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\LNIP01\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\LNIP01\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\LNIP01\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''備份至FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT4.FileName

'備份檔案
If Dir("C:\BEST\LNIP01\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\LNIP01\Orders\Backup"
If Dir("C:\BEST\LNIP01\Orders\Backup\" & filLocalFileT11.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\LNIP01\Orders\Backup\" & filLocalFileT11.FileName
Else
    FileCopy strTranFileName, "C:\BEST\LNIP01\Orders\Backup\" & mySplit(filLocalFileT11.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT11.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT11.Refresh: cboSheetT11.Clear
Screen.MousePointer = 0: cmdImportT11.Enabled = True: dgMainT11.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-匯入", Me.Caption, "cmdImportT11_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT11.Enabled = True: Screen.MousePointer = 0: dgMainT11.Enabled = True

End Sub

Private Sub dgMainT11_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT11
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT11_Change()
    filLocalFileT11.Path = dirLocalDirT11.Path
End Sub
Private Sub drvLocalDriveT11_Change()

On Error GoTo DriveError
dirLocalDirT11.Path = drvLocalDriveT11.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT11_Click()

On Error GoTo err_Handle
Set rsMainT11 = Nothing: Set dgMainT11.DataSource = rsMainT11
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT11.Path, 1) = "\" Then
    strFilePath = filLocalFileT11.Path
Else
    strFilePath = filLocalFileT11.Path & "\"
End If

If Dir(strFilePath & filLocalFileT11.FileName) = "" Then: filLocalFileT11.Refresh: Exit Sub

cboSheetT11.Clear

If UCase(mySplit(filLocalFileT11.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT11.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT11.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT11.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT11.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT11.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT12_Click()
Dim bl_Error As Boolean '記錄有小數點的旗標
Dim str_Error As String '記錄有小數點錯誤的資料
Dim x As Long

bl_Error = False: str_Error = ""
Dim str_Storerkey As String

If rsMainT12 Is Nothing Then Exit Sub
If rsMainT12.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT12.Enabled = False: cmdImportT12.Enabled = False
strTranFileName = filLocalFileT12.Path & "\" & filLocalFileT12.FileName

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where rtrim(updatesource)='" & filLocalFileT12.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT12.RecordCount = 0 Or rsMainT12 Is Nothing Then
Else
rsMainT12.MoveFirst
str_Storerkey = "LPSI01"

Do While Not rsMainT12.EOF
    '到貨日期檢查
    If Len(Trim(rsMainT12("供貨日期"))) = 0 Then
         MsgBox "交易單號:" & Trim(rsMainT12("交貨")) & "的供貨日期為空白，訂單轉入終止!", 16, Me.Caption: dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT12("供貨日期"))) > 0 And Len(Trim(rsMainT12("供貨日期"))) < 8 Then
         MsgBox "交易單號:" & Trim(rsMainT12("交貨")) & "的供貨日期:" & Trim(rsMainT12("供貨日期")) & "，格式不對，請補齊8碼，訂單轉入終止!", 16, Me.Caption: dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
    Else
        '檢查到貨日不可小於今日
        If Trim(rsMainT12("供貨日期")) < Format(Now, "YYYYMMDD") Then
            If blAdmin = True Then
            
            '最高權限不檢查到貨日
                 x = MsgBox("供貨日期小於今日，你確定要繼續嗎?", vbQuestion + vbYesNo, "最高權限到貨日檢查") '紀錄按下的是確定或是取消
                    If x = 6 Then
                        '繼續
                    Else
                        '離開
                         dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
                    End If
            Else
            
            End If
        End If
    End If
    
    '訂單日檢查
    If Len(Trim(rsMainT12("供貨日期"))) = 0 Then
         MsgBox "供貨日期為空白，訂單轉入終止!", 16, Me.Caption: dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT12("供貨日期"))) > 0 And Len(Trim(rsMainT12("供貨日期"))) < 8 Then
         MsgBox "供貨日期:" & Trim(rsMainT12("供貨日期")) & "，格式不對，請補齊8碼，訂單轉入終止!", 16, Me.Caption: dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
'    Else
'        If Trim(rsMainT12("供貨日期")) > Trim(rsMainT12("供貨日期")) Then MsgBox "訂單號碼:" & Trim(rsMainT12("訂單號碼")) & "的訂單日:" & Trim(rsMainT12("供貨日期")) & "，大於到貨日，訂單轉入終止!", 16, Me.Caption: dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
    End If
    
    '數量檢查
    If Val(rsMainT12("交貨數量")) < 1 Then
        MsgBox "數量小於1，" & Trim(rsMainT12("交易單號")) & "-品號：" & Trim(rsMainT12("物料")) & "，訂單轉入終止!!請確認!!", , "訂單檔匯入": dgMainT12.Enabled = True: cmdImportT12.Enabled = True: Exit Sub
        Exit Sub
    End If
    
        '資料檢驗 --判斷SKU是否存在
        str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMainT12("物料")) & "' and Storerkey = '" & str_Storerkey & "'"
        
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '按鈕那些要改
            MsgBox "訂單發現新品號 (" & Trim(rsMainT12("物料")) & ")，請先於商品主檔新建商品資料，訂單轉入終止!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
            dgMainT12.Enabled = True: cmdImportT12.Enabled = True
            Exit Sub
        End If
        
'        If Trim(rsMainT12("訂單類別")) = "A2B" Then
'        '檢查A2B訂單客戶編號是否存在
'                str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT12("提貨客戶編號")) & "' and Storerkey = '" & Str_storerkey & "'"
'
'                Call Confirm_Recordset_Closed(tmp_Rs)
'                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'                If tmp_Rs.EOF Then  '按鈕那些要改
'                    MsgBox "訂單發現新客戶編號 ( " & Trim(rsMainT12("提貨客戶編號")) & " )，請先於客戶主檔新建客戶資料，訂單轉入終止!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
'                    dgMainT12.Enabled = True: cmdImportT12.Enabled = True
'                    Exit Sub
'                End If
'        End If
'
'        '檢查A2B訂單以外的客戶編號是否存在
'        str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT12("到貨客戶編號")) & "' and Storerkey = '" & Str_storerkey & "'"
'
'        Call Confirm_Recordset_Closed(tmp_Rs)
'        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'        If tmp_Rs.EOF Then  '按鈕那些要改
'            MsgBox "訂單發現新客戶編號 ( " & Trim(rsMainT12("到貨客戶編號")) & " )，請先於客戶主檔新建客戶資料，訂單轉入終止!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
'            dgMainT12.Enabled = True: cmdImportT12.Enabled = True
'            Exit Sub
'        End If
        
        '檢查數量有無小數點
        If InStr(Trim(rsMainT12("交貨數量")), ".") <> 0 Then
            str_Error = "交易單號:" & Trim(rsMainT12("交貨")) & "，品號:" & Trim(rsMainT12("物料")) & Chr(13) & str_Error
            bl_Error = True
        End If
        
'        '檢查貨主--mark by Gemini @20160602
'        If UCase(Trim(rsMainT12("貨主"))) <> "LABT01" And UCase(Trim(rsMainT12("貨主"))) <> "LLFA01" Then
'            MsgBox "訂單發現非亞培的貨主: " & Trim(rsMainT12("貨主")) & " )，此匯入程式僅供匯入亞培及利豐訂單，請確認後再匯入，訂單轉入終止!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
'            dgMainT12.Enabled = True: cmdImportT12.Enabled = True
'            Exit Sub
'        End If

        '判斷單別--add by  Gemini @20160602
'        If Trim(rsMainT12("訂單類別")) <> "C" Or Trim(rsMainT12("訂單類別")) <> "R" Or Trim(rsMainT12("訂單類別")) <> "RC" Or Trim(rsMainT12("訂單類別")) <> "I" Or Trim(rsMainT12("訂單類別")) <> "A2B" Then
'        Else
'            MsgBox "系統無此單別:" & Trim(rsMainT12("訂單類別")) & "，請確認此訂單類別是否正確，訂單轉入終止!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
'                dgMainT12.Enabled = True: cmdImportT12.Enabled = True
'            Exit Sub
'        End If
        
        '檢查貨主是否存在
        str_SQL = "select storerkey from trp16m where storerkey = '" & str_Storerkey & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then  '按鈕那些要改
            MsgBox "訂單發現新貨主 ( " & str_Storerkey & " )，請先於貨主主檔新建貨主資料，訂單轉入終止!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
            dgMainT12.Enabled = True: cmdImportT12.Enabled = True
            Exit Sub
        End If
        
        '判斷C單
'        If Trim(rsMainT12("訂單類別")) = "C" And (Trim(rsMainT12("貨主")) = "LKAO01" Or Trim(rsMainT12("貨主")) = "LABT01") Then
'        Else
'            MsgBox "此貨主:" & Trim(rsMainT12("貨主")) & "之訂單類別不可為:" & Trim(rsMainT12("訂單類別")) & "，請確認此訂單類別是否正確，訂單轉入終止!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
'                dgMainT12.Enabled = True: cmdImportT12.Enabled = True
'            Exit Sub
'        End If
        
'        '判斷是否為A2B或R單
'        If Trim(rsMainT12("訂單類別")) <> "A2B" And Trim(rsMainT12("訂單類別")) <> "R" And UCase(Trim(rsMainT12("貨主"))) = "LABT01" Then
'            MsgBox "此貨主:" & Trim(rsMainT12("貨主")) & "之訂單類別不可為:" & Trim(rsMainT12("訂單類別")) & "，請確認此訂單類別是否正確，訂單轉入終止!!": cmdImportT12.Enabled = True: Screen.MousePointer = 0
'                dgMainT12.Enabled = True: cmdImportT12.Enabled = True
'            Exit Sub
'        End If
        
    rsMainT12.MoveNext
Loop
rsMainT12.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "資料有小數點！請重新匯入"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT12.Enabled = True: cmdImportT12.Enabled = True
                Exit Sub
End If

Tran_Level = cn.BeginTrans: cmdImportT12.Enabled = False: dgMainT12.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long, IntCasecnt As Integer
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String
Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String

'開始匯入
Do While Not rsMainT12.EOF
    DoEvents: DoEvents
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT12("交貨"))) Then
        strOrderNo = UCase(Trim(rsMainT12("交貨")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '已經檢查過，則不比對，直接抓出客戶主檔中的 客戶編號，郵遞區號，連絡人，電話，地址
        '單別為A2B則，抓提貨客編，非A2B則抓到貨客編
'        If myExCharFilter(Trim(rsMainT12("訂單類別"))) = "A2B" Then
'            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & myExCharFilter(Trim(rsMainT12("貨主"))) & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT12("提貨客戶編號"))) & "'"
'            Call Confirm_Recordset_Closed(tmp_Rs)
'            tmp_Rs.CursorLocation = 3
'            tmp_Rs.Open str_SQL, cn
'        Else
            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & str_Storerkey & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT12("工廠"))) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.CursorLocation = 3
            tmp_Rs.Open str_SQL, cn
'        End If
        '相符沿用舊客編
        strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
        strZip = myExCharFilter(Trim(tmp_Rs("zip")))
        strContact = myExCharFilter(Trim(tmp_Rs("contact")))
        strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
        strAddress = myExCharFilter(Trim(tmp_Rs("address")))
        strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
        blCustomerMatch = True
        tmp_Rs.Close

        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select externorderkey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT12("交貨"))) & "' and storerkey = '" & str_Storerkey & "' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
            strFacility = "佰事達北倉"
'            If UCase(Right(Trim(rsMainT12("倉別")), 2)) = "-C" Then
'                strFacility = "佰事達中倉"
'            ElseIf UCase(Right(Trim(rsMainT12("倉別")), 2)) = "-S" Then
'                strFacility = "佰事達南倉"
'            Else
'                strFacility = "佰事達北倉"
'            End If
            
'            If Trim(rsMainT12("倉別")) = "" Then strFacility = ""
            
            strOrderDate = Trim(rsMainT12("供貨日期"))
            Dim intPointer As Integer
            intPointer = 1
            'updatesource
'            If myExCharFilter(Trim(rsMainT12("訂單類別"))) = "A2B" Then
            'A2B，多紀錄一個B點的客編B_company
'                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,b_company,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
'                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT12("訂單號碼"))) & "','" & myExCharFilter(Trim(rsMainT12("訂單類別"))) & "','" & myExCharFilter(Trim(rsMainT12("貨主"))) & "','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT12("到貨日"))) & "','" & strFacility & "','" & _
'                myExCharFilter(Trim(rsMainT12("提貨客戶編號"))) & "','" & myExCharFilter(Trim(rsMainT12("到貨客戶編號"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT12("訂單備註"))) & "','" & filLocalFileT12.FileName & "','','" & User_id & "','" & User_id & "','','','" & Trim(rsMainT12("組單類別")) & "','" & Val(Trim(rsMainT12("件數"))) & "') "
'            Else
            'not A2B
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT12("交貨"))) & "','" & "RC" & "','" & str_Storerkey & "','" & myExCharFilter(Trim(rsMainT12("供貨日期"))) & "','" & myExCharFilter(Trim(rsMainT12("供貨日期"))) & "','" & strFacility & "','" & _
                myExCharFilter(Trim(rsMainT12("工廠"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','','" & "" & "','" & filLocalFileT12.FileName & "','','" & User_id & "','" & User_id & "','','','" & "RC" & "','" & "" & "') "
'            End If
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

            
'            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & Str_Storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT12("交貨")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
        
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            
            str_SQL = "select CaseCnt as CN, sku from gv_skuxpack where sku='" & Trim(rsMainT12("物料")) & "' and Storerkey = '" & str_Storerkey & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            
            IntCasecnt = Trim(tmp_Rs("CN"))
            intQTY = Val(rsMainT12("交貨數量") * IntCasecnt)
            
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable03,Lottable06, Facility,UOM,otherUOM)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT12("交貨"))) & "','" & myExCharFilter(Trim(rsMainT12("物料"))) & "','" & str_Storerkey & "'," & _
            "'" & intQTY & "','" & intQTY & "','" & myExCharFilter(Trim(rsMainT12("批次"))) & "','','" & strFacility & "','','" & myExCharFilter(Trim(rsMainT12("BUn"))) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT12.MoveNext
Loop

'執行gs_ordersupdate   用客戶主檔更新訂單資料
'
'cn.Execute "exec gs_ordersupdate 'LYFY09'", RowsAffect, adExecuteNoRecords

'執行gs_ordersupdate
cn.Execute "exec gs_ordersupdate 'LPSI01'", RowsAffect, adExecuteNoRecords

'檢查有無異常訂單  倉別庫別錯誤 es_Checklot06_by_storer '貨主','訂單檔名'
'str_SQL = "exec es_Checklot06_by_storer '" & Str_Storerkey & "','" & filLocalFileT12.FileName & "'"
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'If Not tmp_Rs.EOF Then  '有回傳錯誤的訂單資料，產生excel
'    Recordset2Excel "配送倉別與明細倉別不符的訂單資料", tmp_Rs
'End If
'
'tmp_Rs.Close

cn.CommitTrans: Tran_Level = 0: dgMainT12.Enabled = True


'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "TMS單號： " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "檔案 " & filLocalFileT12.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT12.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 品號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'備份檔案
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT12.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT12.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT12.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT12.FileName, ".", -1)
End If

''備份至FTP
'If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
'FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT12.FileName


Kill strTranFileName
    
filLocalFileT12.Refresh:
Screen.MousePointer = 0: cmdImportT12.Enabled = True: dgMainT12.Enabled = True
Exit Sub

err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
   Call ErrorMsgbox(App.title, err.Number, err.Description, "檔案名稱： " & strTranFileName)
    cmdImportT12.Enabled = True: Screen.MousePointer = 0: dgMainT12.Enabled = True

End Sub
Private Sub dgMainT12_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT12
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT12_Change()
    filLocalFileT12.Path = dirLocalDirT12.Path
End Sub
Private Sub drvLocalDriveT12_Change()

On Error GoTo DriveError
dirLocalDirT12.Path = drvLocalDriveT12.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT12_Click()

On Error GoTo err_Handle
Set rsMainT12 = Nothing: Set dgMainT12.DataSource = rsMainT12
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT12.Path, 1) = "\" Then
    strFilePath = filLocalFileT12.Path
Else
    strFilePath = filLocalFileT12.Path & "\"
End If

If Dir(strFilePath & filLocalFileT12.FileName) = "" Then: filLocalFileT12.Refresh: Exit Sub

cboSheetT12.Clear

If UCase(mySplit(filLocalFileT12.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT12.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT12.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT12.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT12.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT12.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cboSheetT12_Click()

On Error GoTo err_Handle

Dim strFilePath As String, strFieldName As String
If Right(filLocalFileT12.Path, 1) <> "\" Then
    strFilePath = filLocalFileT12.Path & "\"
Else
    strFilePath = filLocalFileT12.Path
End If

'strFieldName = "Delivery" & Chr(9) & "Sold-to.Pt" & Chr(9) & "Name.of.sold-to" & Chr(9) & "Ship-To.Pt" & Chr(9) & "Name.of.the.ship-to.Party" & Chr(9) & "Item" & Chr(9) & "Material" & Chr(9) & "Plnt" & Chr(9) & "SLoc" & Chr(9) & "Batch" & Chr(9) & "Route" & Chr(9) & "Deliv.date" & Chr(9) & "Qty(stckpg.unit)" & Chr(9) & "BUn" & Chr(9) & "Delivery.qty" & Chr(9) & "SU" & Chr(9) & "order.no" & Chr(9) & "po.no" & Chr(9) & "remarks" & Chr(9)

Set rsMainT12 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT12.FileName, cboSheetT12, strFieldName, rsMainT12)

Set dgMainT12.DataSource = rsMainT12

If rsMainT12 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
'
Else

rsMainT12.Sort = "交貨"

    SetDataGridColWidth Me.Caption, dgMainT12
    MsgBox "此工作表共 " & rsMainT12.RecordCount & "筆明細", 64, "Excel2Recordset"

End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Sub Excel2RecordsetT12(strFileName As String, strSheetName As String, strFieldName As String, ByRef rs As ADODB.Recordset)
'**************************************************
'因為欄位名稱重複，所以獨立此副程式，為了跳過第一個欄位名稱
'**************************************************
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
If Dir(strFileName) = "" Then MsgBox "找不到檔案！", vbOKOnly + vbInformation, "Excel2Recordset": Exit Sub '找不到檔案

On Error GoTo err_Handle
Screen.MousePointer = 11

'開啟EXCEL物件
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '選定工作表
        If i = 1 Then strSheetName1 = RTrim(.Sheets(i).Name) '取第一個工作表名稱
        .Sheets(.Sheets(i).Name).Select
    Next
    
    '找不到指定工作表，選用第一個
    If RTrim(.ActiveSheet.Name) <> strSheetName Then .Sheets(strSheetName1).Select
    
    k = 2 '由第二列開始匯入
    
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rs = Nothing: GoTo endsub
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    '寫入Recordset
    Do While Len(RTrim(.Cells(k, 1))) > 0
    rsTmp.AddNew
        For j = 1 To UBound(arrTmp)
            rsTmp(j - 1) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst

    Call OffLineRecordset(rsTmp, rs)
    
    rsTmp.Close: Set rsTmp = Nothing
  
endsub:
Screen.MousePointer = 0
.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

Exit Sub
err_Handle:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Set rs = Nothing

Dim str As String

If err.Number = 3367 Then
    str = "欄位名稱( " & arrTmp(i) & ")重複！"
    
ElseIf err.Number = -2147217887 Then
    str = "請確認儲存格(" & Chr(64 + j) & k & ")，資料是否有誤！"

Else
     str = "Exceed2Recordset"
End If

Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, str)

End Sub
Private Sub cboSheetT13_Click()

On Error GoTo err_Handle
Dim str As String, strFieldName As String, strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT13.Path, 1) = "\" Then
    strFilePath = filLocalFileT13.Path
Else
    strFilePath = filLocalFileT13.Path & "\"
End If

'建立欄位名稱陣列
'strFieldName = "客戶" & Chr(9) & "統編" & Chr(9) & "銷貨日期" & Chr(9) & "銷貨單號" & Chr(9) & "聯絡人" & Chr(9) & "電話" & Chr(9) & "送貨地址" & Chr(9) & "發票號碼" & Chr(9) & "業務員" & Chr(9) & "備註" & Chr(9) & "品號" & Chr(9) & "品名" & Chr(9) & "倉庫" & Chr(9) & "數量" & Chr(9) & "單位" & Chr(9) & "前置單據/備註/客戶單號" & Chr(9)

If Right(filLocalFileT13.Path, 1) <> "\" Then
    strFilePath = filLocalFileT13.Path & "\"
Else
    strFilePath = filLocalFileT13.Path
End If

Set rsMainT13 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT13.FileName, cboSheetT13, strFieldName, rsMainT13)
rsMainT13.Sort = "銷貨單號,品號"

Set dgMainT13.DataSource = rsMainT13

If rsMainT13 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMainT13
    MsgBox "此工作表共 " & rsMainT13.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT13_Click()

If Len(RTrim(cboStorerkeyT13)) = 0 Then MsgBox "請選貨主編號！", 64, "訂單轉入": Exit Sub

On Error GoTo err_Handle

strTranFileName = filLocalFileT13.Path & "\" & filLocalFileT13.FileName
If Len(RTrim(cboSheetT13)) = 0 Then MsgBox "請選擇工作表", 64, Me.Caption: Exit Sub
If rsMainT13.EOF Or rsMainT13 Is Nothing Then Exit Sub
Dim strStorerkey As String

strStorerkey = cboStorerkeyT13

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where rtrim(updatesource) = '" & filLocalFileT13.FileName & "' and storerkey = '" & strStorerkey & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then cmdImportT13.Enabled = True: dgMainT13.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp
rsMainT13.MoveFirst
Do While Not rsMainT13.EOF

    '到貨日期檢查
    arrTmp = Split(Trim(rsMainT13("銷貨日期")), "/")
    If UBound(arrTmp) < 2 Then MsgBox "銷貨日期格式有誤(YYYY/MM/DD)，訂單轉入終止!", 16, Me.Caption: Exit Sub
    If IsDate(Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)) = False Then MsgBox "銷貨日期有誤(" & Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) & ")，訂單轉入終止!", 16, Me.Caption: Exit Sub
    If Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2) < Format(Now, "YYYY/MM/DD") Then MsgBox "到貨日小於今日，訂單轉入終止!", 16, Me.Caption: Exit Sub
            
    '數量檢查
    If Val(rsMainT13("數量")) < 0 Then
        MsgBox "訂單數量小於1，" & Trim(rsMainT13("品號")) & "-" & Trim(rsMainT13("品名")) & "(" & Trim(rsMainT13("數量")) & Trim(rsMainT13("單位")) & ")，訂單轉入終止!!", , "訂單檔匯入": Exit Sub
        Exit Sub
    End If
                 
    rsMainT13.MoveNext
Loop

Tran_Level = cn.BeginTrans: cmdImportT13.Enabled = False: dgMainT13.Enabled = False
Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
            
'取最後客戶編號
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select top 1 substring(consigneekey,5,20) as consigneekey from trp01m where storerkey = '" & strStorerkey & "' and left(consigneekey,4) = 'BEST' order by consigneekey desc "
tmp_Rs.Open str_SQL, cn

If Not tmp_Rs.EOF Then intTmp = Val(tmp_Rs("consigneekey"))

tmp_Rs.Close

rsMainT13.MoveFirst
Do While Not rsMainT13.EOF
    DoEvents: DoEvents
    
'    '資料檢驗--判斷是否屬佰事達訂單-->跳下一筆
'    If UCase(Trim(rsMainT11("倉庫"))) = "" Then
''        MsgBox "客戶單號：" & Trim(rsMainT4("銷貨單號")) & "( " & Trim(rsMainT4("倉庫")) & " )" & vbCrLf & "請通知客戶，確認該筆訂單是否有誤!?", vbOKOnly, "非佰事達之訂單不轉入"
'        intNotBest = intNotBest + 1
'        GoTo next1
'    Else
        '資料檢驗--判斷SKU是否存在
    str_SQL = "select sku from " & strWMSDB & "..sku where sku='" & Trim(rsMainT13("品號")) & "' and Storerkey = '" & strStorerkey & "' "
    
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        cn.RollbackTrans: Tran_Level = 0
        MsgBox "訂單發現新品號 (" & Trim(rsMainT13("品號")) & " ) " & Trim(rsMainT13("品名")) & "，訂單轉入終止!!": cmdImportT13.Enabled = True: dgMainT13.Enabled = True: Screen.MousePointer = 0
        tmp_Rs.Close
        Exit Sub
    End If
    tmp_Rs.Close
    
'    End If
                  
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT13("銷貨單號"))) Then
        strOrderNo = UCase(Trim(rsMainT13("銷貨單號")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '檢查是否有此客戶名稱
        str_SQL = "select top 1 consigneekey ,notes = isnull(notes,'') from trp01m where storerkey = '" & strStorerkey & "' and rtrim(full_name) = '" & myExCharFilter(Trim(rsMainT13("客戶"))) & "' order by len(consigneekey),consigneekey "
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.CursorLocation = 3
        tmp_Rs.Open str_SQL, cn
        
        If tmp_Rs.EOF Then
            '無此客戶名稱則新增
            intTmp = intTmp + 1
            strConsigneeKey = "BEST" & Format(intTmp, "000000")
            
            '新增客戶主檔
            cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address) " & _
            " values('" & strStorerkey & "','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT13("客戶"))) & "','" & myExCharFilter(Trim(rsMainT13("客戶"))) & "','" & myExCharFilter(Trim(rsMainT13("聯絡人"))) & "','" & myExCharFilter(Trim(rsMainT13("電話"))) & "','" & myExCharFilter(Trim(rsMainT13("送貨地址"))) & "' ) ", RowsAffect, adExecuteNoRecords
'
            '紀錄新增之客戶編號
            strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
        Else
            '比對聯絡人、電話與到貨地址是否相符
            Call Confirm_Recordset_Closed(rsTmp)
            str_SQL = "select * from trp01m " & _
                        "where storerkey = '" & strStorerkey & "' and full_name = '" & myExCharFilter(Trim(rsMainT13("客戶"))) & "' " & _
                        "and rtrim(contact) = '" & myExCharFilter(Trim(rsMainT13("聯絡人"))) & "' " & _
                        "and rtrim(phone) = '" & myExCharFilter(Trim(rsMainT13("電話"))) & "' " & _
                        "and rtrim(address) = '" & myExCharFilter(Trim(rsMainT13("送貨地址"))) & "' "
            rsTmp.CursorLocation = 3
            rsTmp.Open str_SQL, cn
            
            If rsTmp.EOF Then
                '聯絡人、電話與到貨地址不符
                intTmp = intTmp + 1
                strConsigneeKey = Left(strStorerkey, 4) & Format(intTmp, "000000")
                
                '新增客戶主檔
                cn.Execute "insert into trp01m(Storerkey,zip,consigneekey,Full_Name,Short_Name,Contact,Phone,Address,notes) " & _
                " values('" & strStorerkey & "','','" & strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT13("客戶"))) & "','" & myExCharFilter(Trim(rsMainT13("客戶"))) & "','" & myExCharFilter(Trim(rsMainT13("聯絡人"))) & "','" & myExCharFilter(Trim(rsMainT13("電話"))) & "','" & myExCharFilter(Trim(rsMainT13("送貨地址"))) & "','" & myExCharFilter(Trim(tmp_Rs("notes"))) & "' ) ", RowsAffect, adExecuteNoRecords
                 
                '紀錄新增之客戶編號
                strNewConsigneekey = strNewConsigneekey & strConsigneeKey & "','"
            Else '相符沿用舊客編
                strConsigneeKey = Trim(rsTmp("consigneekey"))
                blCustomerMatch = True

            End If
            rsTmp.Close
        End If
        tmp_Rs.Close
    
        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select * from orders where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT13("銷貨單號"))) & "' and storerkey = '" & strStorerkey & "' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then
    
            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
            strFacility = "佰事達北倉"
            arrTmp = Split(Trim(rsMainT13("銷貨日期")), "/")
            strOrderDate = Val(arrTmp(0)) + 1911 & "/" & arrTmp(1) & "/" & arrTmp(2)
            Dim intPointer As Integer
            intPointer = 1
            
            str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno) " & _
            "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT13("銷貨單號"))) & "','I','" & strStorerkey & "','" & strOrderDate & "','" & CDate(strOrderDate) + 1 & "','" & strFacility & "','" & _
            strConsigneeKey & "','" & myExCharFilter(Trim(rsMainT13("客戶"))) & "','" & myExCharFilter(Trim(rsMainT13("聯絡人"))) & "','" & myExCharFilter(Trim(rsMainT13("業務員"))) & "','" & myExCharFilter(Trim(rsMainT13("統編"))) & "','" & myExCharFilter(Trim(rsMainT13("電話"))) & "','','" & myExCharFilter(Trim(GetWord(rsMainT13("送貨地址"), intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(rsMainT13("送貨地址"), intPointer, 45))) & "','','" & myExCharFilter(Trim(rsMainT13("備註"))) & "','" & filLocalFileT13.FileName & "','','" & User_id & "','" & User_id & "','" & myExCharFilter(Trim(rsMainT13("發票號碼"))) & "' )"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1
            
            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & strStorerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
            blCustomerMatch = False
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT13("銷貨單號")) & "','"
            blDuplicationOrder = True

        End If
    End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            intQTY = Trim(rsMainT13("數量"))
            strLot06 = Trim(rsMainT13("倉庫"))
            
            '訂單明細資料新增
            str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM,Unitprice,notes)" & _
            "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT13("銷貨單號"))) & "','" & myExCharFilter(Trim(rsMainT13("品號"))) & "','" & strStorerkey & "'," & _
            "'" & intQTY & "','" & intQTY & "','" & strLot06 & "','F1','" & myExCharFilter(Trim(rsMainT13("單位"))) & "','0','" & myExCharFilter(Trim(rsMainT13("前置單據/備註/客戶單號"))) & "')"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from " & strWMSDB & "..sku sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
next1:
        rsMainT13.MoveNext
Loop

cn.CommitTrans: Tran_Level = 0

'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細 (共 " & rsMainT13.RecordCount & " 筆明細)，TMS單號： " & strOrderKeyS & "~" & str_Orderkey & "，檔案 " & filLocalFileT13.FileName & " 備份於 C:\BEST\Other\Orders\Backup "
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單， " & int_OrderLine & " 筆明細，非佰事達訂單 " & intNotBest & " 筆明細，檔案 " & filLocalFileT13.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 本次檔案名稱 = '" & filLocalFileT13.FileName & "' , 上次檔案名稱 = o.updatesource ,重複訂單號碼 = rtrim(o.externorderkey) ,上次客戶單號 = rtrim(o.customerorderkey) ,  上次訂單日期 = convert(varchar,o.orderdate,111) , 上次到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 上次品號 = od.sku , 上次數量 = od.openqty ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\Best\Other\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\Other\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\Other\訂單重複\" & strStorerkey & "訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

''備份至FTP
'If Dir("O:\LVTL01\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\LVTL01\OrdersBackup"
'FileCopy strTranFileName, "O:\LVTL01\OrdersBackup\" & filLocalFileT4.FileName

'備份檔案
If Dir("C:\BEST\Other\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\Other\Orders\Backup"
If Dir("C:\BEST\Other\Orders\Backup\" & filLocalFileT11.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\Other\Orders\Backup\" & filLocalFileT13.FileName
Else
    FileCopy strTranFileName, "C:\BEST\Other\Orders\Backup\" & mySplit(filLocalFileT13.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT13.FileName, ".", -1)
End If

Kill strTranFileName
    
filLocalFileT13.Refresh: cboSheetT13.Clear
Screen.MousePointer = 0: cmdImportT13.Enabled = True: dgMainT13.Enabled = True
Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-匯入", Me.Caption, "cmdImportT13_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    cmdImportT13.Enabled = True: Screen.MousePointer = 0: dgMainT13.Enabled = True

End Sub

Private Sub dgMainT13_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT13
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT13_Change()
    filLocalFileT13.Path = dirLocalDirT13.Path
End Sub
Private Sub drvLocalDriveT13_Change()

On Error GoTo DriveError
dirLocalDirT13.Path = drvLocalDriveT13.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT13_Click()

On Error GoTo err_Handle
Set rsMainT13 = Nothing: Set dgMainT13.DataSource = rsMainT13
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT13.Path, 1) = "\" Then
    strFilePath = filLocalFileT13.Path
Else
    strFilePath = filLocalFileT13.Path & "\"
End If

If Dir(strFilePath & filLocalFileT13.FileName) = "" Then: filLocalFileT13.Refresh: Exit Sub

cboSheetT13.Clear

If UCase(mySplit(filLocalFileT13.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT13.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT13.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
        DoEvents: DoEvents
        cboSheetT13.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT13.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT13.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub
Private Sub cmd2Excel_Click()

'資料排序
Recordset2Excel "TEST", rsMainT4

'..在此編輯EXCEL
With MyXlsApp
    
End With

Set MyXlsApp = Nothing

End Sub
Private Sub cboSheetT14_Click()

On Error GoTo err_Handle

Dim strFilePath As String, strFieldName As String
If Right(filLocalFileT14.Path, 1) <> "\" Then
    strFilePath = filLocalFileT14.Path & "\"
Else
    strFilePath = filLocalFileT14.Path
End If

'strFieldName = "Delivery" & Chr(9) & "Sold-to.Pt" & Chr(9) & "Name.of.sold-to" & Chr(9) & "Ship-To.Pt" & Chr(9) & "Name.of.the.ship-to.Party" & Chr(9) & "Item" & Chr(9) & "Material" & Chr(9) & "Plnt" & Chr(9) & "SLoc" & Chr(9) & "Batch" & Chr(9) & "Route" & Chr(9) & "Deliv.date" & Chr(9) & "Qty(stckpg.unit)" & Chr(9) & "BUn" & Chr(9) & "Delivery.qty" & Chr(9) & "SU" & Chr(9) & "order.no" & Chr(9) & "po.no" & Chr(9) & "remarks" & Chr(9)

Set rsMainT14 = New ADODB.Recordset
Call Excel2Recordset(strFilePath & filLocalFileT14.FileName, cboSheetT14, strFieldName, rsMainT14)

Set dgMainT14.DataSource = rsMainT14

If rsMainT14 Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
'
Else

rsMainT14.Sort = "貨主,訂單號碼"

    SetDataGridColWidth Me.Caption, dgMainT14
    MsgBox "此工作表共 " & rsMainT14.RecordCount & "筆明細", 64, "Excel2Recordset"

End If

Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdImportT14_Click()
Dim bl_Error As Boolean '記錄有小數點的旗標
Dim str_Error As String '記錄有小數點錯誤的資料
Dim x As Long
bl_Error = False: str_Error = ""
Dim str_Storerkey As String

If rsMainT14 Is Nothing Then Exit Sub
If rsMainT14.EOF Then Exit Sub

On Error GoTo err_Handle

dgMainT14.Enabled = False: cmdImportT14.Enabled = False
strTranFileName = filLocalFileT14.Path & "\" & filLocalFileT14.FileName

'資料檢驗--判斷檔案是否已轉入
Call Confirm_Recordset_Closed(tmp_Rs)
str_SQL = "select orderkey from orders where rtrim(updatesource)='" & filLocalFileT14.FileName & "' "

tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If tmp_Rs.EOF = False Then dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Screen.MousePointer = 0: tmp_Rs.Close: MsgBox "檔案名稱相同，請確認是否重複轉入!", vbOKOnly, Me.Caption: Exit Sub
tmp_Rs.Close

Dim arrTmp

If rsMainT14.RecordCount = 0 Or rsMainT14 Is Nothing Then
Else
rsMainT14.MoveFirst
str_Storerkey = myExCharFilter(Trim(rsMainT14("貨主")))

Do While Not rsMainT14.EOF
    '到貨日期檢查
    If Len(Trim(rsMainT14("到貨日"))) = 0 Then
         MsgBox "訂單號碼:" & Trim(rsMainT14("訂單號碼")) & "的到貨日為空白，訂單轉入終止!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT14("到貨日"))) > 0 And Len(Trim(rsMainT14("到貨日"))) < 8 Then
         MsgBox "訂單號碼:" & Trim(rsMainT14("訂單號碼")) & "的到貨日:" & Trim(rsMainT14("到貨日")) & "，格式不對，請補齊8碼，訂單轉入終止!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
    ElseIf IsDate(Left(Trim(rsMainT14("到貨日")), 4) + "/" + Mid(Trim(rsMainT14("到貨日")), 5, 2) + "/" + Right(Trim(rsMainT14("到貨日")), 2)) = False Then
         MsgBox "訂單號碼:" & Trim(rsMainT14("訂單號碼")) & "的到貨日:" & Trim(rsMainT14("到貨日")) & "，不是一個正常日期，訂單轉入終止!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
    Else
        '檢查到貨日不可小於今日
        If Trim(rsMainT14("到貨日")) < Format(Now, "YYYYMMDD") Then
            If blAdmin = True Then
            
            '最高權限不檢查到貨日
                 x = MsgBox("到貨日小於今日，你確定要繼續嗎?", vbQuestion + vbYesNo, "最高權限到貨日檢查") '紀錄按下的是確定或是取消
                    If x = 6 Then
                        '繼續
                    Else
                        '離開
                         dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
                    End If
            Else
            
            End If
        End If
    End If
    
    '訂單日檢查
    If Len(Trim(rsMainT14("訂單日"))) = 0 Then
         MsgBox "訂單號碼:" & Trim(rsMainT14("訂單號碼")) & "的訂單日為空白，訂單轉入終止!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
    ElseIf Len(Trim(rsMainT14("訂單日"))) > 0 And Len(Trim(rsMainT14("訂單日"))) < 8 Then
         MsgBox "訂單號碼:" & Trim(rsMainT14("訂單號碼")) & "的訂單日:" & Trim(rsMainT14("訂單日")) & "，格式不對，請補齊8碼，訂單轉入終止!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
    ElseIf IsDate(Left(Trim(rsMainT14("訂單日")), 4) + "/" + Mid(Trim(rsMainT14("訂單日")), 5, 2) + "/" + Right(Trim(rsMainT14("訂單日")), 2)) = False Then
         MsgBox "訂單號碼:" & Trim(rsMainT14("訂單號碼")) & "的訂單日:" & Trim(rsMainT14("到貨日")) & "，不是一個正常日期，訂單轉入終止!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub

    Else
        If Trim(rsMainT14("訂單日")) > Trim(rsMainT14("到貨日")) Then MsgBox "訂單號碼:" & Trim(rsMainT14("訂單號碼")) & "的訂單日:" & Trim(rsMainT14("訂單日")) & "，大於到貨日，訂單轉入終止!", 16, Me.Caption: dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
    End If
    
    '數量檢查
    If Val(rsMainT14("數量")) < 1 Then
        MsgBox "數量小於1，" & Trim(rsMainT14("訂單號碼")) & "-品號：" & Trim(rsMainT14("品號")) & "，訂單轉入終止!!請確認!!", , "訂單檔匯入": dgMainT14.Enabled = True: cmdImportT14.Enabled = True: Exit Sub
        Exit Sub
    End If
    
        '資料檢驗 --判斷SKU是否存在
        str_SQL = "select sku from gv_skuxpack where sku='" & Trim(rsMainT14("品號")) & "' and Storerkey = '" & Trim(rsMainT14("貨主")) & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '按鈕那些要改
            MsgBox "訂單發現新品號 (" & Trim(rsMainT14("品號")) & ")，請先於商品主檔新建商品資料，訂單轉入終止!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
            dgMainT14.Enabled = True: cmdImportT14.Enabled = True
            Exit Sub
        End If
        
        If Trim(rsMainT14("訂單類別")) = "A2B" Then
        '檢查A2B訂單客戶編號是否存在
                str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT14("提貨客戶編號")) & "' and Storerkey = '" & Trim(rsMainT14("貨主")) & "' "
            
                Call Confirm_Recordset_Closed(tmp_Rs)
                tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            
                If tmp_Rs.EOF Then  '按鈕那些要改
                    MsgBox "訂單發現新客戶編號 ( " & Trim(rsMainT14("提貨客戶編號")) & " )，請先於客戶主檔新建客戶資料，訂單轉入終止!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
                    dgMainT14.Enabled = True: cmdImportT14.Enabled = True
                    Exit Sub
                End If
        End If
        
        '檢查A2B訂單以外的客戶編號是否存在
        str_SQL = "select consigneekey from trp01m where consigneekey='" & Trim(rsMainT14("到貨客戶編號")) & "' and Storerkey = '" & Trim(rsMainT14("貨主")) & "' "
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    
        If tmp_Rs.EOF Then  '按鈕那些要改
            MsgBox "訂單發現新客戶編號 ( " & Trim(rsMainT14("到貨客戶編號")) & " )，請先於客戶主檔新建客戶資料，訂單轉入終止!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
            dgMainT14.Enabled = True: cmdImportT14.Enabled = True
            Exit Sub
        End If
        
        '檢查數量有無小數點
        If InStr(Trim(rsMainT14("數量")), ".") <> 0 Then
            str_Error = "訂單號碼:" & Trim(rsMainT14("訂單號碼")) & "，品號:" & Trim(rsMainT14("品號")) & Chr(13) & str_Error
            bl_Error = True
        End If
        
'        '檢查貨主--mark by Gemini @20160602
'        If UCase(Trim(rsMainT14("貨主"))) <> "LABT01" And UCase(Trim(rsMainT14("貨主"))) <> "LLFA01" Then
'            MsgBox "訂單發現非亞培的貨主: " & Trim(rsMainT14("貨主")) & " )，此匯入程式僅供匯入亞培及利豐訂單，請確認後再匯入，訂單轉入終止!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
'            dgMainT14.Enabled = True: cmdImportT14.Enabled = True
'            Exit Sub
'        End If

        '判斷單別--add by  Gemini @20160602
        If Trim(rsMainT14("訂單類別")) <> "C" Or Trim(rsMainT14("訂單類別")) <> "R" Or Trim(rsMainT14("訂單類別")) <> "RC" Or Trim(rsMainT14("訂單類別")) <> "I" Or Trim(rsMainT14("訂單類別")) <> "A2B" Then
        Else
            MsgBox "系統無此單別:" & Trim(rsMainT14("訂單類別")) & "，請確認此訂單類別是否正確，訂單轉入終止!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
                dgMainT14.Enabled = True: cmdImportT14.Enabled = True
            Exit Sub
        End If
        
        '檢查貨主是否存在
        str_SQL = "select storerkey from trp16m where storerkey = '" & Trim(rsMainT14("貨主")) & "'"
        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then  '按鈕那些要改
            MsgBox "訂單發現新貨主 ( " & Trim(rsMainT14("貨主")) & " )，請先於貨主主檔新建貨主資料，訂單轉入終止!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
            dgMainT14.Enabled = True: cmdImportT14.Enabled = True
            Exit Sub
        End If
        
        '判斷C單
'        If Trim(rsMainT14("訂單類別")) = "C" And (Trim(rsMainT14("貨主")) = "LKAO01" Or Trim(rsMainT14("貨主")) = "LABT01") Then
'        Else
'            MsgBox "此貨主:" & Trim(rsMainT14("貨主")) & "之訂單類別不可為:" & Trim(rsMainT14("訂單類別")) & "，請確認此訂單類別是否正確，訂單轉入終止!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
'                dgMainT14.Enabled = True: cmdImportT14.Enabled = True
'            Exit Sub
'        End If
        
'        '判斷是否為A2B或R單
'        If Trim(rsMainT14("訂單類別")) <> "A2B" And Trim(rsMainT14("訂單類別")) <> "R" And UCase(Trim(rsMainT14("貨主"))) = "LABT01" Then
'            MsgBox "此貨主:" & Trim(rsMainT14("貨主")) & "之訂單類別不可為:" & Trim(rsMainT14("訂單類別")) & "，請確認此訂單類別是否正確，訂單轉入終止!!": cmdImportT14.Enabled = True: Screen.MousePointer = 0
'                dgMainT14.Enabled = True: cmdImportT14.Enabled = True
'            Exit Sub
'        End If
        
    rsMainT14.MoveNext
Loop
rsMainT14.MoveFirst
End If

If bl_Error = True Then
                msg_text = str_Error & Chr(13) & "資料有小數點！請重新匯入"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                dgMainT14.Enabled = True: cmdImportT14.Enabled = True
                Exit Sub
End If

Tran_Level = cn.BeginTrans: cmdImportT14.Enabled = False: dgMainT14.Enabled = False

Dim int_OrderLine As Integer, int_Order As Integer, int_Asn As Integer, int_Repeat As Integer, intQTY As Long, intNotBest As Integer, int_orderlinenuber As Integer, intTmp As Long, lngCasecnt As Long, lngInnerpack As Long
Dim strLot06 As String, strOrderNo As String, strDeliveryDate As String, strReOrderkey As String, strRePoOrderkey As String, strTKLOC As String, strFacility As String, blDuplicationOrder As Boolean, blCustomerMatch As Boolean, strOrderDate As String, strConsigneeKey As String, strNewConsigneekey As String, strOrderKeyS As String
Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
Dim strDate As String, strOrderType As String, strSku As String
Dim strZip As String, strContact As String, strPhone As String, strAddress As String, strShort_name As String

'開始匯入
Do While Not rsMainT14.EOF
    DoEvents: DoEvents
    '資料檢驗--來源訂單相同單號判斷，不同增加HEAD
    If strOrderNo <> UCase(Trim(rsMainT14("訂單號碼"))) Then
        strOrderNo = UCase(Trim(rsMainT14("訂單號碼")))
        int_orderlinenuber = 0
        blDuplicationOrder = False
        
        '已經檢查過，則不比對，直接抓出客戶主檔中的 客戶編號，郵遞區號，連絡人，電話，地址
        '單別為A2B則，抓提貨客編，非A2B則抓到貨客編
        If myExCharFilter(Trim(rsMainT14("訂單類別"))) = "A2B" Then
            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & myExCharFilter(Trim(rsMainT14("貨主"))) & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT14("提貨客戶編號"))) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.CursorLocation = 3
            tmp_Rs.Open str_SQL, cn
        Else
            str_SQL = "select consigneekey,zip=isnull(zip,''),contact=isnull(contact,''),phone=isnull(phone,''),address=isnull(address,''),short_name=isnull(short_name,'') from trp01m where storerkey = '" & myExCharFilter(Trim(rsMainT14("貨主"))) & "' and consigneekey = '" & myExCharFilter(Trim(rsMainT14("到貨客戶編號"))) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.CursorLocation = 3
            tmp_Rs.Open str_SQL, cn
        End If
        '相符沿用舊客編
        strConsigneeKey = myExCharFilter(Trim(tmp_Rs("consigneekey")))
        strZip = myExCharFilter(Trim(tmp_Rs("zip")))
        strContact = myExCharFilter(Trim(tmp_Rs("contact")))
        strPhone = myExCharFilter(Trim(tmp_Rs("phone")))
        strAddress = myExCharFilter(Trim(tmp_Rs("address")))
        strShort_name = myExCharFilter(Trim(tmp_Rs("short_name")))
        blCustomerMatch = True
        tmp_Rs.Close

        '資料檢驗--判斷訂單是否重複，重複不增加
        Call Confirm_Recordset_Closed(tmp_Rs)
        str_SQL = "select externorderkey from orders(nolock) where rtrim(ExternOrderKey) ='" & myExCharFilter(Trim(rsMainT14("訂單號碼"))) & "' and storerkey = '" & myExCharFilter(Trim(rsMainT14("貨主"))) & "' and isnull(type,'') <> '刪單' "
        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        If tmp_Rs.EOF Then

            '取訂單號碼
            str_SQL = "select isnull(max(orderkey),0) from orders"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_Orderkey = StrPadLeft(Val(Trim(tmp_Rs.Fields(0))) + 1, 10, 0)
            If strOrderKeyS = "" Then strOrderKeyS = str_Orderkey
            tmp_Rs.Close
            
            '配送倉別判斷
            If UCase(Right(Trim(rsMainT14("倉別")), 2)) = "-C" Then
                strFacility = "佰事達中倉"
            ElseIf UCase(Right(Trim(rsMainT14("倉別")), 2)) = "-S" Then
                strFacility = "佰事達南倉"
            Else
                strFacility = "佰事達北倉"
            End If
            
            If Trim(rsMainT14("倉別")) = "" Then strFacility = ""
            
            strOrderDate = Trim(rsMainT14("訂單日"))
            Dim intPointer As Integer
            intPointer = 1
            'updatesource
            If myExCharFilter(Trim(rsMainT14("訂單類別"))) = "A2B" Then
                
            'A2B，多紀錄一個B點的客編B_company
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,b_company,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT14("訂單號碼"))) & "','" & myExCharFilter(Trim(rsMainT14("訂單類別"))) & "','" & myExCharFilter(Trim(rsMainT14("貨主"))) & "','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT14("到貨日"))) & "','" & strFacility & "','" & _
                myExCharFilter(Trim(rsMainT14("提貨客戶編號"))) & "','" & myExCharFilter(Trim(rsMainT14("到貨客戶編號"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','" & Trim(rsMainT14("採購單號")) & "','" & myExCharFilter(Trim(rsMainT14("訂單備註"))) & "','" & filLocalFileT14.FileName & "','','" & User_id & "','" & User_id & "','','','" & Trim(rsMainT14("組單類別")) & "','" & Val(Trim(rsMainT14("件數"))) & "') "
            Else
            'not A2B
                str_SQL = "INSERT orders (OrderKey,ExternOrderKey,Priority,StorerKey,OrderDate,Deliverydate,Facility,ConsigneeKey,c_company,c_contact1,c_contact2,c_vat,c_phone1,c_zip,c_address1,c_address2,CustomerOrderkey,Notes,UpdateSource,type,addwho,editwho,invoiceno,externconsigneekey,b_city,otqty) " & _
                "VALUES ('" & str_Orderkey & "','" & myExCharFilter(Trim(rsMainT14("訂單號碼"))) & "','" & myExCharFilter(Trim(rsMainT14("訂單類別"))) & "','" & myExCharFilter(Trim(rsMainT14("貨主"))) & "','" & strOrderDate & "','" & myExCharFilter(Trim(rsMainT14("到貨日"))) & "','" & strFacility & "','" & _
                myExCharFilter(Trim(rsMainT14("到貨客戶編號"))) & "','" & strShort_name & "','" & strContact & "','','','" & strPhone & "','" & strZip & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 58))) & "','" & myExCharFilter(Trim(GetWord(strAddress, intPointer, 45))) & "','" & Trim(rsMainT14("採購單號")) & "','" & myExCharFilter(Trim(rsMainT14("訂單備註"))) & "','" & filLocalFileT14.FileName & "','','" & User_id & "','" & User_id & "','','','" & Trim(rsMainT14("組單類別")) & "','" & Val(Trim(rsMainT14("件數"))) & "') "
            End If
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_Order = int_Order + 1

            
'            '如果客戶主檔相符，更新訂單郵遞區號，以免需客戶確認
'            If blCustomerMatch = True Then cn.Execute "update orders set c_zip = (select zip from trp01m where consigneekey = '" & strConsigneeKey & "' and storerkey = '" & Str_Storerkey & "') where orderkey = '" & str_Orderkey & "' ", RowsAffect, adExecuteNoRecords
'            blCustomerMatch = False
        Else
            '訂單重複
            Call FTPlog("訂單重複" & str_SQL)
            '紀錄重複
            strReOrderkey = strReOrderkey & Trim(rsMainT14("訂單號碼")) & "','"
            blDuplicationOrder = True

        End If
   End If
    
        '訂單重複檢查
        If blDuplicationOrder = False Then
        
            '增加明細
            int_orderlinenuber = int_orderlinenuber + 1
            
            intQTY = Val(rsMainT14("數量"))
            
            
            '訂單明細資料新增
            If Trim(rsMainT14("單位名稱")) = "箱" Or Trim(rsMainT14("單位名稱")) = "CS" Or Trim(rsMainT14("單位名稱")) = "CASE" Then
                str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM)" & _
                " select '" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT14("訂單號碼"))) & "','" & myExCharFilter(Trim(rsMainT14("品號"))) & "','" & myExCharFilter(Trim(rsMainT14("貨主"))) & "'," & _
                "'" & intQTY & "' * p.casecnt ,'" & intQTY & "' * p.casecnt,'" & myExCharFilter(Trim(rsMainT14("倉別"))) & "','',''" & _
                "from " & strWMSDB & "..sku s join " & strWMSDB & "..pack p on s.packkey = p.packkey and s.sku = '" & Trim(rsMainT14("品號")) & "' and s.storerkey = '" & myExCharFilter(Trim(rsMainT14("貨主"))) & "' "
            Else
                 str_SQL = "INSERT ORDERDETAIL (OrderKey,OrderLineNumber,ExternOrderKey,Sku,StorerKey,OriginalQty,OpenQty,Lottable06, Facility,UOM)" & _
                "VALUES ('" & str_Orderkey & "','" & Format(int_orderlinenuber, "00000") & "','" & myExCharFilter(Trim(rsMainT14("訂單號碼"))) & "','" & myExCharFilter(Trim(rsMainT14("品號"))) & "','" & myExCharFilter(Trim(rsMainT14("貨主"))) & "'," & _
                "'" & intQTY & "','" & intQTY & "','" & myExCharFilter(Trim(rsMainT14("倉別"))) & "','','')"
           End If
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            '更新packkey
            str_SQL = "update orderdetail " & _
            "Set orderdetail.packkey = sku.packkey " & _
            "from gv_skuxpack sku join orderdetail on orderdetail.sku = sku.sku and sku.storerkey = orderdetail.storerkey " & _
            "where orderkey = '" & str_Orderkey & "' "
            
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            int_OrderLine = int_OrderLine + 1
        End If
        
nextRow17:
        rsMainT14.MoveNext
Loop

'執行gs_ordersupdate   用客戶主檔更新訂單資料
'
'cn.Execute "exec gs_ordersupdate 'LYFY09'", RowsAffect, adExecuteNoRecords

'檢查有無異常訂單  倉別庫別錯誤 es_Checklot06_by_storer '貨主','訂單檔名'
'str_SQL = "exec es_Checklot06_by_storer '" & Str_Storerkey & "','" & filLocalFileT14.FileName & "'"
'Call Confirm_Recordset_Closed(tmp_Rs)
'tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'
'If Not tmp_Rs.EOF Then  '有回傳錯誤的訂單資料，產生excel
'    Recordset2Excel "配送倉別與明細倉別不符的訂單資料", tmp_Rs
'End If
'
'tmp_Rs.Close

cn.CommitTrans: Tran_Level = 0: dgMainT14.Enabled = True


'訊息顯示
    msg_text = "匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "TMS單號： " & strOrderKeyS & "~" & str_Orderkey
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Call FTPlog("匯入 " & int_Order & " 筆訂單" & Chr(13) & "匯入" & int_OrderLine & " 筆明細" & Chr(13) & "檔案 " & filLocalFileT14.FileName)
    
'訂單重複顯示
If Len(strReOrderkey & strRePoOrderkey) > 0 Then

    str_SQL = "select 重複類別 = '訂單重複-不轉入' , 轉入檔案名稱 = '" & filLocalFileT14.FileName & "' ,訂單號碼 = rtrim(o.externorderkey) ,客戶單號 = rtrim(o.customerorderkey) ,  訂單日期 = convert(varchar,o.orderdate,111) , 到貨日 = convert(varchar,o.deliverydate,111) , 項次 = od.orderlinenumber , 品號 = od.sku , 數量 = od.openqty , 上次檔案名稱 = o.updatesource  ,檢查時間 = getdate() " & _
        "From orders o join orderdetail od on o.orderkey = od.orderkey where rtrim(o.ExternOrderKey) in ('" & strReOrderkey & "') and o.storerkey = '" & str_Storerkey & "' "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    MsgBox "發現訂單重複!!", vbOKOnly, Me.Caption
    
    Call Recordset2Excel("訂單重複", tmp_Rs)
    If Dir("C:\BEST\" & str_Storerkey & "\訂單重複", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\訂單重複"
    MyXlsApp.Range("g:g").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    MyXlsApp.ActiveWorkbook.SaveAs "C:\BEST\" & str_Storerkey & "\訂單重複\訂單重複_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    Set MyXlsApp = Nothing
       
End If

'備份檔案
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup", vbDirectory) = "" Then MkDirs "C:\BEST\" & str_Storerkey & "\Orders\Backup"
If Dir("C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT14.FileName) = "" Then
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & filLocalFileT14.FileName
Else
    FileCopy strTranFileName, "C:\BEST\" & str_Storerkey & "\Orders\Backup\" & mySplit(filLocalFileT14.FileName, ".", 0) & Format(Now, "yyyyMMddhhmmss") & "." & mySplit(filLocalFileT14.FileName, ".", -1)
End If

'備份至FTP
If Dir("O:\" & str_Storerkey & "\OrdersBackup", vbDirectory) = "" Then MkDirs "O:\" & str_Storerkey & "\OrdersBackup"
FileCopy strTranFileName, "O:\" & str_Storerkey & "\OrdersBackup\" & filLocalFileT14.FileName



Kill strTranFileName
    
filLocalFileT14.Refresh:
Screen.MousePointer = 0: cmdImportT14.Enabled = True: dgMainT14.Enabled = True
Exit Sub

err_Handle:
    If Tran_Level <> 0 Then cn.RollbackTrans
    Close #1
   Call ErrorMsgbox(App.title, err.Number, err.Description, "檔案名稱： " & strTranFileName)
    cmdImportT14.Enabled = True: Screen.MousePointer = 0: dgMainT14.Enabled = True

End Sub

Private Sub dgMainT14_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMainT14
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub dirLocalDirT14_Change()
    filLocalFileT14.Path = dirLocalDirT14.Path
End Sub

Private Sub drvLocalDriveT14_Change()

On Error GoTo DriveError
dirLocalDirT14.Path = drvLocalDriveT14.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next

End Sub

Private Sub filLocalFileT14_Click()

On Error GoTo err_Handle
Set rsMainT14 = Nothing: Set dgMainT14.DataSource = rsMainT14
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFileT14.Path, 1) = "\" Then
    strFilePath = filLocalFileT14.Path
Else
    strFilePath = filLocalFileT14.Path & "\"
End If

If Dir(strFilePath & filLocalFileT14.FileName) = "" Then: filLocalFileT14.Refresh: Exit Sub

cboSheetT14.Clear

If UCase(mySplit(filLocalFileT14.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFileT14.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheetT14.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        cboSheetT14.AddItem MyXlsApp.Sheets(i).Name
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheetT14.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheetT14.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Sub Excel2RecordsetT8(strFileName As String, strSheetName As String, strFieldName As String, ByRef rs As ADODB.Recordset)
'**************************************************
'因為欄位名稱重複，所以獨立此副程式，為了跳過第一個欄位名稱
'**************************************************
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
If Dir(strFileName) = "" Then MsgBox "找不到檔案！", vbOKOnly + vbInformation, "Excel2Recordset": Exit Sub '找不到檔案

On Error GoTo err_Handle
Screen.MousePointer = 11

'開啟EXCEL物件
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '選定工作表
        If i = 1 Then strSheetName1 = RTrim(.Sheets(i).Name) '取第一個工作表名稱
        .Sheets(.Sheets(i).Name).Select
    Next
    
    '找不到指定工作表，選用第一個
    If RTrim(.ActiveSheet.Name) <> strSheetName Then .Sheets(strSheetName1).Select
    
    k = 2 '由第二列開始匯入
    
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rs = Nothing: GoTo endsub
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    '寫入Recordset
    Do While Len(RTrim(.Cells(k, 1))) > 0
    rsTmp.AddNew
        For j = 1 To UBound(arrTmp)
            rsTmp(j - 1) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst

    Call OffLineRecordset(rsTmp, rs)
    
    rsTmp.Close: Set rsTmp = Nothing
  
endsub:
Screen.MousePointer = 0
.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

Exit Sub
err_Handle:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Set rs = Nothing

Dim str As String

If err.Number = 3367 Then
    str = "欄位名稱( " & arrTmp(i) & ")重複！"
    
ElseIf err.Number = -2147217887 Then
    str = "請確認儲存格(" & Chr(64 + j) & k & ")，資料是否有誤！"

Else
     str = "Exceed2Recordset"
End If

Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, str)

End Sub
Sub Excel2RecordsetT9(strFileName As String, strSheetName As String, strFieldName As String, ByRef rs As ADODB.Recordset)
'**************************************************
'因為欄位名稱重複，所以獨立此副程式，為了跳過第一個欄位名稱
'**************************************************
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
If Dir(strFileName) = "" Then MsgBox "找不到檔案！", vbOKOnly + vbInformation, "Excel2Recordset": Exit Sub '找不到檔案

On Error GoTo err_Handle
Screen.MousePointer = 11

'開啟EXCEL物件
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If RTrim(.Sheets(i).Name) = strSheetName Then .Sheets(strSheetName).Select: Exit For '選定工作表
        If i = 1 Then strSheetName1 = RTrim(.Sheets(i).Name) '取第一個工作表名稱
        .Sheets(.Sheets(i).Name).Select
    Next
    
    '找不到指定工作表，選用第一個
    If RTrim(.ActiveSheet.Name) <> strSheetName Then .Sheets(strSheetName1).Select
    
    k = 2 '由第二列開始匯入
    
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rs = Nothing: GoTo endsub
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    '寫入Recordset
    Do While Len(RTrim(.Cells(k, 1))) > 0
    rsTmp.AddNew
        For j = 1 To UBound(arrTmp)
            rsTmp(j - 1) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst

    Call OffLineRecordset(rsTmp, rs)
    
    rsTmp.Close: Set rsTmp = Nothing
  
endsub:
Screen.MousePointer = 0
.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

Exit Sub
err_Handle:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Set rs = Nothing

Dim str As String

If err.Number = 3367 Then
    str = "欄位名稱( " & arrTmp(i) & ")重複！"
    
ElseIf err.Number = -2147217887 Then
    str = "請確認儲存格(" & Chr(64 + j) & k & ")，資料是否有誤！"

Else
     str = "Exceed2Recordset"
End If

Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, str)

End Sub

Private Sub DB_Connect_Self(connection_string As String)
'ADO [Connection] Object connect
On Error GoTo err_Handle
Set cn_Self = New ADODB.Connection
cn_Self.CommandTimeout = 300
cn_Self.ConnectionTimeout = 20
cn_Self.ConnectionString = connection_string
cn_Self.Open Options:=adAsyncConnect
Do While cn_Self.State = adStateConnecting
   DoEvents: DoEvents
Loop
Exit Sub

err_Handle:
   msg_text = "連線錯誤：無法與資料庫建立連線，請通知 資訊部 "
   MsgBox msg_text, vbOKOnly + vbInformation, ""
   End
End Sub
Sub XML2Recordset(FileName As String)
Dim i As Integer, arrLen
arrLen = Array(12, 6, 12, 2, 50, 15, 19, 12, 6, 19, 60, 15, 15, 15, 10, 100, 20, 60, 8, 6, 3, 3, 10, 19, 20, 76, 76, 3, 10, 10, 60, 255, 19, 20)

'檔案長度為0
If FileLen(FileName) = 0 Then Call ErrorMsgbox(Me.Caption, err.Number, err.Description, FileName & "檔案長度為 0 "): Exit Sub

Set rs_Src = Nothing

Dim objXMLDOM As New MSXML2.DOMDocument40
Dim objNodes As IXMLDOMNodeList
Dim objBookNode As IXMLDOMNode

'開始讀取xml
objXMLDOM.async = False

'開啟xml檔錯誤
If Not objXMLDOM.Load(FileName) Then Call ErrorMsgbox(Me.Caption, err.Number, objXMLDOM.parseError.reason, FileName): Exit Sub

Set objNodes = objXMLDOM.selectNodes("/ROWSET/ROW")
If objNodes.Length <= 0 Then Call ErrorMsgbox(Me.Caption, err.Number, "objNodes.length = " & objNodes.Length, FileName): Exit Sub

Set rs_Src = New ADODB.Recordset
With rs_Src
    .Fields.Append "DELIVERY_NO", adChar, 12, adFldUpdatable
    .Fields.Append "GOODS_OWNER", adChar, 6, adFldUpdatable
    .Fields.Append "DELIVERY_DETAIL_ID", adChar, 12, adFldUpdatable
    .Fields.Append "ORDER_TYPE", adChar, 2, adFldUpdatable
    .Fields.Append "TRANSACTION_TYPE_DESC", adChar, 50, adFldUpdatable
    .Fields.Append "SALESREP", adChar, 15, adFldUpdatable
    .Fields.Append "ORDERED_DATE", adChar, 19, adFldUpdatable
    .Fields.Append "SALES_ORDER_NUMBER", adChar, 12, adFldUpdatable
    .Fields.Append "SEQUENCE_NO", adChar, 6, adFldUpdatable
    .Fields.Append "NOTIFY_DATE", adChar, 19, adFldUpdatable
    .Fields.Append "CUSTOMER_NAME", adChar, 60, adFldUpdatable
    .Fields.Append "CONTACT", adChar, 15, adFldUpdatable
    .Fields.Append "TELEPHONE", adChar, 15, adFldUpdatable
    .Fields.Append "SINGLE_SHOP_CODE", adChar, 15, adFldUpdatable
    .Fields.Append "SUPPLIER_CODE", adChar, 10, adFldUpdatable
    .Fields.Append "REQUEST_ADDRESS", adChar, 100, adFldUpdatable
    .Fields.Append "ITEM_NAME", adChar, 20, adFldUpdatable
    .Fields.Append "ITEM_DESCRIPTION", adChar, 60, adFldUpdatable
    .Fields.Append "QTY1", adChar, 8, adFldUpdatable
    .Fields.Append "QTY2", adChar, 6, adFldUpdatable
    .Fields.Append "UOM1", adChar, 3, adFldUpdatable
    .Fields.Append "UOM2", adChar, 3, adFldUpdatable
    .Fields.Append "SELLING_PRICE", adChar, 10, adFldUpdatable
    .Fields.Append "SCHEDULE_SHIP_DATE", adChar, 19, adFldUpdatable
    .Fields.Append "CUST_PO_NUMBER", adChar, 20, adFldUpdatable
    .Fields.Append "DELIVERY_INFORMATIONS", adChar, 76, adFldUpdatable
    .Fields.Append "SHIPPING_INSTRUCTIONS", adChar, 76, adFldUpdatable
    .Fields.Append "ORGANIZATION_CODE", adChar, 3, adFldUpdatable
    .Fields.Append "SUBINVENTORY", adChar, 10, adFldUpdatable
    .Fields.Append "POSTAL_CODE", adChar, 10, adFldUpdatable
    .Fields.Append "SHIP_TO_CUSTOMER_NAME", adChar, 60, adFldUpdatable
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open

'寫入rs

For Each objBookNode In objNodes

       .AddNew
       For i = 0 To rs_Src.Fields.Count - 1
         .Fields(i) = Replace(GetWord((objBookNode.selectSingleNode(rs_Src.Fields(i).Name).nodeTypedValue), 1, arrLen(i)), "'", """")
       Next i
       .Update

Next objBookNode

.Sort = "SALES_ORDER_NUMBER,SEQUENCE_NO"
.MoveFirst

Set dgMainT20.DataSource = rs_Src

End With

End Sub
Sub ExceltoRecordset(strFileName As String, strSheetName As String, strFieldName As String, ByRef rs As ADODB.Recordset)
'**************************************************
'Create by Gemini @20090312 4 Excel匯入Recordset
'使用說明
'1.如果來源Excel工作表不帶欄位名稱，請於strFieldName指定，並以char(9)作為分隔符號
'strFieldName = "客戶" & Chr(9) & "統編" & Chr(9) & "銷貨日期" & Chr(9) & "銷貨單號" & Chr(9) & "聯絡人" & Chr(9) & "電話" & Chr(9) & "送貨地址" & Chr(9) & "發票號碼" & Chr(9) & "業務員" & Chr(9) & "備註" & Chr(9) & "品號" & Chr(9) & "品名" & Chr(9) & "倉庫" & Chr(9) & "數量" & Chr(9) & "單位" & Chr(9) & "前置單據/備註/客戶單號" & Chr(9)

'參數說明
'strFileName:來源檔案名稱路徑
'strSheetName:來源工作表
'strFieldName:欄位名稱
'rs:回傳的Recordset
'範例
'call Excel2Recordset ("C:\book1.xls","Sheet1", "客戶" & Chr(9) & "統編" & Chr(9) & "銷貨日期" & Chr(9),rsMain)
'**************************************************
Dim i As Integer, j As Integer, k As Long, arrTmp, strSheetName1 As String, strSheetNamex As String
If Dir(strFileName) = "" Then MsgBox "找不到檔案！", vbOKOnly + vbInformation, "Excel2Recordset": Exit Sub '找不到檔案

On Error GoTo err_Handle
Screen.MousePointer = 11

'開啟EXCEL物件
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)

    '尋找指定工作表
    For i = 1 To .Sheets.Count
        If (.Sheets(i).Name) = (strSheetName) Then .Sheets(strSheetName).Select: Exit For '選定工作表
    Next
    
    If (.ActiveSheet.Name) <> (strSheetName) Then
        MsgBox "找不到 " & strSheetName & " 工作表！", vbOKOnly + vbInformation, "Excel2Recordset"
        .Quit: Set MyXlsApp = Nothing
        Exit Sub
    End If

    k = 1 '預設由第一列開始匯入
    
    '若無來源欄位名稱
    If strFieldName = "" Then
        '取欄位名稱
        For i = 1 To 255
            If Len(Trim(.Cells(i, 1) & "")) = 0 Then Exit For
               strFieldName = strFieldName & myExCharFilter(Trim(.Cells(1, i))) & Chr(9)
        Next i
        k = 2 '由第二列開始匯入
    End If
    
    '切割欄位名稱
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then Set rs = Nothing: GoTo endsub
    '建立Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "第 " & i & " 欄位名稱 (" & arrTmp(i) & ") 有誤，檔案載入終止!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adVarChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    '寫入Recordset  '從這邊開始往下寫
    Do While Len(RTrim(.Cells(k, 1))) > 0
    rsTmp.AddNew
        For j = 1 To UBound(arrTmp)
            rsTmp(j - 1) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst

    Call OffLineRecordset(rsTmp, rs)
    
    rsTmp.Close: Set rsTmp = Nothing
  
endsub:
Screen.MousePointer = 0
.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

Exit Sub
err_Handle:
MyXlsApp.Quit: Set MyXlsApp = Nothing
Set rs = Nothing

Dim str As String

If err.Number = 3367 Then
    str = "欄位名稱( " & arrTmp(i) & ")重複！"
    
ElseIf err.Number = -2147217887 Then
    str = "請確認儲存格(" & k & Chr(64 + j) & ")，資料是否有誤！"
    
ElseIf err.Number = 13 Then
    str = "請確認儲存格(" & k & Chr(64 + j) & ")，資料是否有誤！"

Else
     str = "請確認儲存格(" & k & Chr(64 + j) & ")，資料是否有誤！"
End If

Call ErrorMsgbox("Recordset2Excel", err.Number, err.Description, str)

End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
If Trim(SSTab2.Caption) = "" Then SSTab2.Tab = PreviousTab
End Sub


Public Sub LCDisConnect(IpAddress)
 
Dim str1 As String, strRun As String
 'yes不詢問
str1 = "NET use \\" & IpAddress & " /Delete /yes" & vbCrLf
strRun = str1
Shell "cmd.exe /c " & strRun, vbHide
 
End Sub

Public Sub LCConnect(IpAddress As String, ACC As String, PassWord As String)
 '重新連線
 'LCConnect "192.168.2.202", "share", "share"
Dim str1 As String, strRun As String
str1 = "NET use \\" & IpAddress & " " & PassWord & " /user:" & IpAddress & "\" & ACC & " /PERSISTENT:NO" & vbCrLf
strRun = str1
Shell "cmd.exe /c " & strRun, vbHide
 
End Sub
