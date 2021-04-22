VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frm_SystemUpdate 
   Caption         =   " 系 統 更 新"
   ClientHeight    =   5745
   ClientLeft      =   690
   ClientTop       =   1875
   ClientWidth     =   10875
   Icon            =   "frm_SystemUpdate.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   10875
   Begin TabDlg.SSTab SSTab1 
      Height          =   5490
      Left            =   135
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9684
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "主程式更新"
      TabPicture(0)   =   "frm_SystemUpdate.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "報表檔更新"
      TabPicture(1)   =   "frm_SystemUpdate.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "條碼字型"
      TabPicture(2)   =   "frm_SystemUpdate.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Excel 檔案更新"
      TabPicture(3)   =   "frm_SystemUpdate.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame8"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame8 
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   4470
         Left            =   -70365
         TabIndex        =   42
         Top             =   645
         Width           =   5655
         Begin VB.TextBox txt_Tab4Path 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1275
            TabIndex        =   48
            Top             =   360
            Width           =   4140
         End
         Begin VB.ListBox lst_Tab4UpdateInfo 
            Height          =   2220
            Left            =   1290
            TabIndex        =   47
            Top             =   900
            Width           =   4095
         End
         Begin VB.CommandButton cmd_Tab4GetUpdateInfo 
            BackColor       =   &H00C0E0FF&
            Caption         =   "更新資訊"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   1845
            Picture         =   "frm_SystemUpdate.frx":037A
            Style           =   1  '圖片外觀
            TabIndex        =   46
            Top             =   3405
            Width           =   975
         End
         Begin VB.CommandButton cmd_Tab4Exit 
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
            Height          =   840
            Left            =   4290
            Picture         =   "frm_SystemUpdate.frx":0684
            Style           =   1  '圖片外觀
            TabIndex        =   45
            Top             =   3420
            Width           =   975
         End
         Begin VB.CommandButton cmd_Tab4SysInfo 
            BackColor       =   &H00C0FFFF&
            Caption         =   "系統資訊"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   630
            Picture         =   "frm_SystemUpdate.frx":0AC6
            Style           =   1  '圖片外觀
            TabIndex        =   44
            Top             =   3405
            Width           =   975
         End
         Begin VB.CommandButton cmd_Tab4Update 
            BackColor       =   &H00C0FFC0&
            Caption         =   "系統更新"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   3090
            Picture         =   "frm_SystemUpdate.frx":1688
            Style           =   1  '圖片外觀
            TabIndex        =   43
            Top             =   3405
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "檔案路徑"
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
            Height          =   195
            Index           =   7
            Left            =   330
            TabIndex        =   50
            Top             =   420
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "檔案名稱"
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
            Height          =   195
            Index           =   6
            Left            =   345
            TabIndex        =   49
            Top             =   960
            Width           =   840
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   4470
         Left            =   -74715
         TabIndex        =   39
         Top             =   645
         Width           =   4200
         Begin VB.ListBox lst_Tab4SystemInfo 
            Height          =   3660
            Left            =   75
            TabIndex        =   40
            Top             =   660
            Width           =   4050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "目前使用之 Excel 檔案資訊"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   300
            TabIndex        =   41
            Top             =   315
            Width           =   2475
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   4470
         Left            =   -70365
         TabIndex        =   30
         Top             =   645
         Width           =   5655
         Begin VB.CommandButton cmd_Tab3Update 
            BackColor       =   &H00C0FFC0&
            Caption         =   "系統更新"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   3090
            Picture         =   "frm_SystemUpdate.frx":1F52
            Style           =   1  '圖片外觀
            TabIndex        =   36
            Top             =   3405
            Width           =   975
         End
         Begin VB.CommandButton cmd_Tab3SysInfo 
            BackColor       =   &H00C0FFFF&
            Caption         =   "系統資訊"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   630
            Picture         =   "frm_SystemUpdate.frx":281C
            Style           =   1  '圖片外觀
            TabIndex        =   35
            Top             =   3405
            Width           =   975
         End
         Begin VB.CommandButton cmd_Tab3Exit 
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
            Height          =   840
            Left            =   4290
            Picture         =   "frm_SystemUpdate.frx":33DE
            Style           =   1  '圖片外觀
            TabIndex        =   34
            Top             =   3420
            Width           =   975
         End
         Begin VB.CommandButton cmd_Tab3GetUpdateInfo 
            BackColor       =   &H00C0E0FF&
            Caption         =   "更新資訊"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   1860
            Picture         =   "frm_SystemUpdate.frx":3820
            Style           =   1  '圖片外觀
            TabIndex        =   33
            Top             =   3405
            Width           =   975
         End
         Begin VB.ListBox lst_Tab3UpdateInfo 
            Height          =   2220
            Left            =   1290
            TabIndex        =   32
            Top             =   900
            Width           =   4095
         End
         Begin VB.TextBox txt_Tab3Path 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1275
            TabIndex        =   31
            Top             =   360
            Width           =   4140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "檔案名稱"
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
            Height          =   195
            Index           =   5
            Left            =   345
            TabIndex        =   38
            Top             =   960
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "檔案路徑"
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
            Height          =   195
            Index           =   4
            Left            =   330
            TabIndex        =   37
            Top             =   420
            Width           =   840
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   4470
         Left            =   -74715
         TabIndex        =   27
         Top             =   645
         Width           =   4200
         Begin VB.ListBox lst_Tab3SystemInfo 
            Height          =   3480
            Left            =   255
            TabIndex        =   28
            Top             =   660
            Width           =   3690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "已安裝之條碼字型資訊"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   300
            TabIndex        =   29
            Top             =   315
            Width           =   2100
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   4470
         Left            =   -74715
         TabIndex        =   24
         Top             =   645
         Width           =   4200
         Begin VB.ListBox lst_Tab2SystemInfo 
            Height          =   3480
            Left            =   255
            TabIndex        =   25
            Top             =   660
            Width           =   3690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "目前使用之報表檔資訊"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   300
            TabIndex        =   26
            Top             =   315
            Width           =   2100
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   4470
         Left            =   -70365
         TabIndex        =   14
         Top             =   645
         Width           =   5655
         Begin VB.TextBox txt_Tab2Path 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1275
            TabIndex        =   21
            Top             =   360
            Width           =   4140
         End
         Begin VB.TextBox txt_Tab2FileName 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1275
            TabIndex        =   20
            Top             =   885
            Width           =   4140
         End
         Begin VB.ListBox lst_Tab2UpdateInfo 
            Height          =   1680
            Left            =   1290
            TabIndex        =   19
            Top             =   1425
            Width           =   3690
         End
         Begin VB.CommandButton cmd_Tab2GetUpdateInfo 
            BackColor       =   &H00C0E0FF&
            Caption         =   "更新資訊"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   1860
            Picture         =   "frm_SystemUpdate.frx":3B2A
            Style           =   1  '圖片外觀
            TabIndex        =   18
            Top             =   3405
            Width           =   975
         End
         Begin VB.CommandButton cmd_Tab2Exit 
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
            Height          =   840
            Left            =   4290
            Picture         =   "frm_SystemUpdate.frx":3E34
            Style           =   1  '圖片外觀
            TabIndex        =   17
            Top             =   3420
            Width           =   975
         End
         Begin VB.CommandButton cmd_Tab2SysInfo 
            BackColor       =   &H00C0FFFF&
            Caption         =   "系統資訊"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   630
            Picture         =   "frm_SystemUpdate.frx":4276
            Style           =   1  '圖片外觀
            TabIndex        =   16
            Top             =   3405
            Width           =   975
         End
         Begin VB.CommandButton cmd_Tab2Update 
            BackColor       =   &H00C0FFC0&
            Caption         =   "系統更新"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   3090
            Picture         =   "frm_SystemUpdate.frx":4E38
            Style           =   1  '圖片外觀
            TabIndex        =   15
            Top             =   3405
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "檔案路徑"
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
            Height          =   195
            Index           =   3
            Left            =   330
            TabIndex        =   23
            Top             =   420
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "檔案名稱"
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
            Height          =   195
            Index           =   2
            Left            =   345
            TabIndex        =   22
            Top             =   960
            Width           =   840
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   4470
         Left            =   4635
         TabIndex        =   4
         Top             =   645
         Width           =   5655
         Begin VB.CommandButton cmd_Tab1Update 
            BackColor       =   &H00C0FFC0&
            Caption         =   "系統更新"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   3090
            Picture         =   "frm_SystemUpdate.frx":5702
            Style           =   1  '圖片外觀
            TabIndex        =   13
            Top             =   3405
            Width           =   975
         End
         Begin VB.CommandButton cmd_Tab1SysInfo 
            BackColor       =   &H00C0FFFF&
            Caption         =   "系統資訊"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   615
            Picture         =   "frm_SystemUpdate.frx":5FCC
            Style           =   1  '圖片外觀
            TabIndex        =   12
            Top             =   3405
            Width           =   975
         End
         Begin VB.CommandButton cmd_Tab1Exit 
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
            Height          =   840
            Left            =   4290
            Picture         =   "frm_SystemUpdate.frx":6B8E
            Style           =   1  '圖片外觀
            TabIndex        =   11
            Top             =   3405
            Width           =   975
         End
         Begin VB.CommandButton cmd_Tab1GetUpdateInfo 
            BackColor       =   &H00C0E0FF&
            Caption         =   "更新資訊"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   1845
            Picture         =   "frm_SystemUpdate.frx":6FD0
            Style           =   1  '圖片外觀
            TabIndex        =   10
            Top             =   3405
            Width           =   975
         End
         Begin VB.ListBox lst_Tab1UpdateInfo 
            Height          =   1680
            Left            =   1290
            TabIndex        =   9
            Top             =   1425
            Width           =   3690
         End
         Begin VB.TextBox txt_Tab1FileName 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1275
            TabIndex        =   8
            Top             =   885
            Width           =   4140
         End
         Begin VB.TextBox txt_Tab1Path 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   1275
            TabIndex        =   7
            Top             =   360
            Width           =   4140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "檔案名稱"
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
            Height          =   195
            Index           =   1
            Left            =   345
            TabIndex        =   6
            Top             =   960
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "檔案路徑"
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
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   5
            Top             =   420
            Width           =   840
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   4470
         Left            =   285
         TabIndex        =   1
         Top             =   645
         Width           =   4200
         Begin VB.ListBox lst_Tab1SystemInfo 
            Height          =   3480
            Left            =   255
            TabIndex        =   3
            Top             =   630
            Width           =   3690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "系統資訊"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   195
            Left            =   300
            TabIndex        =   2
            Top             =   315
            Width           =   840
         End
      End
   End
End
Attribute VB_Name = "frm_SystemUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private intloop As Integer

Private fso As Scripting.FileSystemObject
Private fd_Path As Scripting.Folder
Private fl_file As Scripting.File
Private ts_File As Scripting.TextStream

Private Sub cmd_Tab1GetUpdateInfo_Click()
'取得系統更新資訊
Dim objIni As vbIniFile
Dim strPath As String, strFileName As String

Set objIni = New vbIniFile
If Dir(striniFileName_FullPath, vbHidden + vbReadOnly) = "" Then
   msg_text = "指定設定檔 [" & striniFileName_FullPath & " 不存在" & vbCrLf & _
              "請通知系統維護人員 "
   MsgBox msg_text, vbOKOnly, msg_title
   Exit Sub
End If

'指定 INI 檔案存放位置與檔案名稱
objIni.FileName = striniFileName_FullPath
'取得 FilePath , FileName
strPath = objIni.ReadData("SYSTEMUPDATE", "FILEPATH", "0")
strFileName = objIni.ReadData("SYSTEMUPDATE", "FILENAME", "0")
'顯示取得之系統更新資訊
txt_Tab1Path.Text = strPath
txt_Tab1FileName.Text = strFileName
Set objIni = Nothing

lst_Tab1UpdateInfo.Clear
'設定欄位顯示格式
Dim clnvalue(2) As Long, lstString As String
clnvalue(0) = 70        '欄位說明
clnvalue(1) = 1000      '欄位資料
Call ListBox_SetTabStops(lst_Tab1UpdateInfo.hwnd, 2, clnvalue)

Set fso = New FileSystemObject
If fso.FileExists(strPath & "\" & strFileName) = False Then
   msg_text = "系統執行檔：" & strPath & "\" & strFileName & " 不存在" & vbCrLf & "請通知系統維護人員"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Set fl_file = fso.GetFile(strPath & "\" & strFileName)
lst_Tab1UpdateInfo.AddItem "日期時間" & vbTab & fl_file.DateCreated
lst_Tab1UpdateInfo.AddItem ""
lst_Tab1UpdateInfo.AddItem "最近存取" & vbTab & fl_file.DateLastAccessed
lst_Tab1UpdateInfo.AddItem ""
lst_Tab1UpdateInfo.AddItem "最後修改" & vbTab & fl_file.DateLastModified
lst_Tab1UpdateInfo.AddItem ""
lst_Tab1UpdateInfo.AddItem "檔案大小" & vbTab & fl_file.Size & " Bytes"
Set fl_file = Nothing
Set fso = Nothing

End Sub

Private Sub cmd_Tab1Exit_Click()
Unload Me
End Sub

Private Sub cmd_Tab1SysInfo_Click()
lst_Tab1SystemInfo.Clear

'設定欄位顯示格式
Dim clnvalue(2) As Long, lstString As String
clnvalue(0) = 70        '欄位說明
clnvalue(1) = 1000      '欄位資料
Call ListBox_SetTabStops(lst_Tab1SystemInfo.hwnd, 2, clnvalue)

lst_Tab1SystemInfo.AddItem "系統名稱" & vbTab & App.ProductName
lst_Tab1SystemInfo.AddItem "公司名稱" & vbTab & App.CompanyName
lst_Tab1SystemInfo.AddItem "註解說明" & vbTab & App.Comments
lst_Tab1SystemInfo.AddItem "商標註記" & vbTab & App.LegalTrademarks
lst_Tab1SystemInfo.AddItem "著 作 權" & vbTab & App.LegalCopyright
lst_Tab1SystemInfo.AddItem "檔案名稱" & vbTab & App.EXEName & ".exe"
lst_Tab1SystemInfo.AddItem "程式路徑" & vbTab & App.Path
lst_Tab1SystemInfo.AddItem ""

Set fso = New FileSystemObject
If fso.FileExists(App.Path & "\" & App.EXEName & ".exe") = False Then Exit Sub
Set fl_file = fso.GetFile(App.Path & "\" & App.EXEName & ".exe")
lst_Tab1SystemInfo.AddItem "日期時間" & vbTab & fl_file.DateCreated
lst_Tab1SystemInfo.AddItem ""
lst_Tab1SystemInfo.AddItem "最近存取" & vbTab & fl_file.DateLastAccessed
lst_Tab1SystemInfo.AddItem ""
lst_Tab1SystemInfo.AddItem "最後修改" & vbTab & fl_file.DateLastModified
lst_Tab1SystemInfo.AddItem ""
lst_Tab1SystemInfo.AddItem "檔案大小" & vbTab & fl_file.Size & " Bytes"
Set fl_file = Nothing
Set fso = Nothing

End Sub

Private Sub cmd_Tab1Update_Click()
Dim strSrcFullPath As String, strSrcBackName As String, strNewFullPath As String
strSrcFullPath = App.Path & "\" & App.EXEName & ".exe"           '原執行檔檔名
strSrcBackName = App.Path & "\" & App.EXEName & "_Backup.exe"    '原執行檔備份檔名
strNewFullPath = txt_Tab1Path.Text & "\" & txt_Tab1FileName.Text '更新執行檔檔名

On Error GoTo err_Handle
Set fso = New FileSystemObject
If fso.FileExists(strSrcBackName) Then
   fso.DeleteFile (strSrcBackName)
End If
If fso.FileExists(strNewFullPath) = False Then
   msg_text = "檔案：" & strNewFullPath & " 不存在" & vbCrLf & "請通知系統維護人員"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Set fl_file = fso.GetFile(strSrcFullPath)
fl_file.Name = App.EXEName & "_Backup.exe"
Set fl_file = fso.GetFile(strNewFullPath)
fl_file.Copy (strSrcFullPath)
Set fl_file = Nothing
Set fso = Nothing
msg_text = "更新狀態：成功完成系統執行檔更新作業"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'更新 [系統資訊] 資料
Call cmd_Tab1SysInfo_Click
Exit Sub

err_Handle:
   msg_text = "系統更新錯誤，錯誤訊息如下" & vbCrLf & err.Number & "：" & err.Description
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab2Exit_Click()
Unload Me
End Sub

Private Sub cmd_Tab2GetUpdateInfo_Click()
lst_Tab2UpdateInfo.Clear
'取得系統更新資訊
Dim objIni As vbIniFile
Dim strPath As String, strFileName As String

Set objIni = New vbIniFile
If Dir(striniFileName_FullPath, vbHidden + vbReadOnly) = "" Then
   msg_text = "指定設定檔 [" & striniFileName_FullPath & " 不存在" & vbCrLf & _
              "請通知系統維護人員 "
   MsgBox msg_text, vbOKOnly, msg_title
   Exit Sub
End If

'指定 INI 檔案存放位置與檔案名稱
objIni.FileName = striniFileName_FullPath
'取得 FilePath , FileName
strPath = objIni.ReadData("REPORT", "FILEPATH", "0")
strFileName = objIni.ReadData("REPORT", "FILENAME", "0")
'顯示取得之系統更新資訊
txt_Tab2Path.Text = strPath
txt_Tab2FileName.Text = strFileName
Set objIni = Nothing

lst_Tab1UpdateInfo.Clear
'設定欄位顯示格式
Dim clnvalue(2) As Long, lstString As String
clnvalue(0) = 70        '欄位說明
clnvalue(1) = 1000      '欄位資料
Call ListBox_SetTabStops(lst_Tab2UpdateInfo.hwnd, 2, clnvalue)

Set fso = New FileSystemObject
If fso.FileExists(strPath & "\" & strFileName) = False Then
   msg_text = "更新報表檔：" & strPath & "\" & strFileName & " 不存在" & vbCrLf & "請通知系統維護人員"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Set fl_file = fso.GetFile(strPath & "\" & strFileName)
lst_Tab2UpdateInfo.AddItem "日期時間" & vbTab & fl_file.DateCreated
lst_Tab2UpdateInfo.AddItem ""
lst_Tab2UpdateInfo.AddItem "最近存取" & vbTab & fl_file.DateLastAccessed
lst_Tab2UpdateInfo.AddItem ""
lst_Tab2UpdateInfo.AddItem "最後修改" & vbTab & fl_file.DateLastModified
lst_Tab2UpdateInfo.AddItem ""
lst_Tab2UpdateInfo.AddItem "檔案大小" & vbTab & fl_file.Size & " Bytes"
Set fl_file = Nothing
Set fso = Nothing
End Sub

Private Sub cmd_Tab2SysInfo_Click()
lst_Tab2SystemInfo.Clear

'設定欄位顯示格式
Dim clnvalue(2) As Long, lstString As String
clnvalue(0) = 70       '欄位說明
clnvalue(1) = 1000      '欄位資料
Call ListBox_SetTabStops(lst_Tab2SystemInfo.hwnd, 2, clnvalue)

Set fso = New FileSystemObject
If fso.FileExists(GetAccessDBFileName) = False Then
   msg_text = "檔案遺失：報表檔不見了，請通知系統維護人員"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
Set fl_file = fso.GetFile(GetAccessDBFileName)
lst_Tab2SystemInfo.AddItem "檔案名稱" & vbTab & fl_file.Name
lst_Tab2SystemInfo.AddItem ""
lst_Tab2SystemInfo.AddItem "存放目錄" & vbTab & fl_file.ParentFolder
lst_Tab2SystemInfo.AddItem ""
lst_Tab2SystemInfo.AddItem "日期時間" & vbTab & fl_file.DateCreated
lst_Tab2SystemInfo.AddItem ""
lst_Tab2SystemInfo.AddItem "最近存取" & vbTab & fl_file.DateLastAccessed
lst_Tab2SystemInfo.AddItem ""
lst_Tab2SystemInfo.AddItem "最後修改" & vbTab & fl_file.DateLastModified
lst_Tab2SystemInfo.AddItem ""
lst_Tab2SystemInfo.AddItem "檔案大小" & vbTab & fl_file.Size & " Bytes"
Set fl_file = Nothing
Set fso = Nothing
End Sub

Private Sub cmd_Tab2Update_Click()
Dim strSrcFullPath As String, strSrcBackName As String, strNewFullPath As String
strSrcFullPath = GetAccessDBFileName            '原 Access DB 檔檔名
strSrcBackName = GetAccessDBFileName & "_bk"    '原 Access DB 檔備份檔名
strNewFullPath = txt_Tab2Path.Text & "\" & txt_Tab2FileName.Text '更新執行檔檔名

On Error GoTo err_Handle
Set fso = New FileSystemObject
If fso.FileExists(strSrcBackName) Then
   fso.DeleteFile (strSrcBackName)
End If
If fso.FileExists(strNewFullPath) = False Then
   msg_text = "報表更新檔：" & strNewFullPath & " 不存在" & vbCrLf & "請通知系統維護人員"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If fso.FileExists(strSrcFullPath) Then
   Set fl_file = fso.GetFile(strSrcFullPath)
   fl_file.Copy (strSrcBackName)
   fl_file.Delete
End If
Set fl_file = fso.GetFile(strNewFullPath)
fl_file.Copy (strSrcFullPath)
Set fl_file = Nothing
Set fso = Nothing
msg_text = "更新狀態：成功完成報表檔更新作業"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'更新 [報表檔資訊] 資料
Call cmd_Tab2SysInfo_Click
Exit Sub

err_Handle:
   msg_text = "系統更新錯誤，錯誤訊息如下" & vbCrLf & err.Number & "：" & err.Description
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab3Exit_Click()
Unload Me
End Sub

Private Sub cmd_Tab3GetUpdateInfo_Click()
lst_Tab3UpdateInfo.Clear

'設定欄位顯示格式
Dim clnvalue(2) As Long, lstString As String
clnvalue(0) = 80        '欄位說明
clnvalue(1) = 1000      '欄位資料
Call ListBox_SetTabStops(lst_Tab3UpdateInfo.hwnd, 2, clnvalue)

Dim tmpI As Integer, fso As FileSystemObject, strNewFontPath As String
Call Get_BardcodeFoneName(BardCode_FontName(), strNewFontPath) '取得條碼字型資訊
txt_Tab3Path.Text = strNewFontPath
'顯示 EXC-DATA Server 上各個條碼字型檔相關資訊：檔案名稱，Size
Set fso = New FileSystemObject
For tmpI = LBound(BardCode_FontName) To UBound(BardCode_FontName)
    If fso.FileExists(txt_Tab3Path.Text & "\" & BardCode_FontName(tmpI)) Then
       lstString = BardCode_FontName(tmpI) & vbTab & fso.GetFile(txt_Tab3Path.Text & "\" & BardCode_FontName(tmpI)).Size & " Bytes"
       lst_Tab3UpdateInfo.AddItem lstString
    Else
       lstString = BardCode_FontName(tmpI) & vbTab & "無此檔案"
       lst_Tab3UpdateInfo.AddItem lstString
    End If
Next tmpI
Set fso = Nothing

End Sub

Private Sub cmd_Tab3SysInfo_Click()
lst_Tab3SystemInfo.Clear

'設定欄位顯示格式
Dim clnvalue(2) As Long, lstString As String
clnvalue(0) = 80        '欄位說明
clnvalue(1) = 1000      '欄位資料
Call ListBox_SetTabStops(lst_Tab3SystemInfo.hwnd, 2, clnvalue)

Dim strFontPath As String, tmpI As Integer, fso As FileSystemObject, strNewFontPath As String
Call Get_BardcodeFoneName(BardCode_FontName(), strNewFontPath) '取得條碼字型資訊
'顯示目前已安裝之條碼字型資訊
Set fso = New FileSystemObject
strFontPath = Get_SystemFolder("FONTS")            '取得系統字型目錄
For tmpI = LBound(BardCode_FontName) To UBound(BardCode_FontName)
    If fso.FileExists(strFontPath & "\" & BardCode_FontName(tmpI)) Then
       lstString = BardCode_FontName(tmpI) & vbTab & fso.GetFile(strFontPath & "\" & BardCode_FontName(tmpI)).Size & " Bytes"
       lst_Tab3SystemInfo.AddItem lstString
    Else
       lstString = BardCode_FontName(tmpI) & vbTab & "無此檔案"
       lst_Tab3SystemInfo.AddItem lstString
    End If
Next tmpI
Set fso = Nothing
End Sub

Private Sub cmd_Tab3Update_Click()
On Error GoTo err_Handle
If Not CheckLoginUser Then
   msg_text = "先生！非常抱歉，您沒有執行此作業之權限，請通知系統維護人員"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Unload Me
End If

'條碼字型檔更新
Dim strFontPath As String, tmpI As Integer, fso As FileSystemObject, strNewFontPath As String
Set fso = New FileSystemObject
Call Get_BardcodeFoneName(BardCode_FontName(), strNewFontPath) '取得條碼字型資訊
strFontPath = Get_SystemFolder("FONTS")                        '取得系統字型目錄
'MsgBox "System Font Path : " & strFontPath & "  New Font Path : " & strNewFontPath
For tmpI = LBound(BardCode_FontName) To UBound(BardCode_FontName)
    If fso.FileExists(strFontPath & "\" & BardCode_FontName(tmpI)) Then
       '刪除系統舊字型檔
       fso.DeleteFile (strFontPath & "\" & BardCode_FontName(tmpI))
    End If
    fso.CopyFile strNewFontPath & "\" & BardCode_FontName(tmpI), strFontPath & "\", True
Next tmpI
msg_text = "更新狀態：成功完成條碼字型檔更新作業"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title
Set fso = Nothing
'更新 [已安裝之條碼字型資訊] 資料
Call cmd_Tab3SysInfo_Click
Exit Sub

err_Handle:
   msg_text = "系統更新錯誤，錯誤訊息如下" & vbCrLf & err.Number & "：" & err.Description
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab4Exit_Click()
Unload Me
End Sub

Private Sub cmd_Tab4GetUpdateInfo_Click()
lst_Tab4UpdateInfo.Clear
'取得系統更新資訊
Dim objIni As vbIniFile
Dim strPath As String, strFileName As String, arFileName() As String

Set objIni = New vbIniFile
If Dir(striniFileName_FullPath, vbHidden + vbReadOnly) = "" Then
   msg_text = "指定設定檔 [" & striniFileName_FullPath & " 不存在" & vbCrLf & _
              "請通知系統維護人員 "
   MsgBox msg_text, vbOKOnly, msg_title
   Exit Sub
End If

'指定 INI 檔案存放位置與檔案名稱
objIni.FileName = striniFileName_FullPath
'取得 FilePath , FileName
strFileName = objIni.ReadData("EXCEL", "FILENAME", "0")
strPath = objIni.ReadData("EXCEL", "FILEPATH", "0")
txt_Tab4Path.Text = strPath
arFileName = Split(strFileName, ",", -1, vbTextCompare)

'設定欄位顯示格式
Dim clnvalue(2) As Long, lstString As String
clnvalue(0) = 60        '欄位說明
clnvalue(1) = 1000      '欄位資料
Call ListBox_SetTabStops(lst_Tab4UpdateInfo.hwnd, 2, clnvalue)
'顯示 EXC-DATA Server 上各個 Excel 範本檔相關資訊：說明，CreateDate，Size
Set fso = New FileSystemObject
For intloop = LBound(arFileName) To UBound(arFileName)
    lst_Tab4UpdateInfo.AddItem arFileName(intloop) & ".xls"
    lst_Tab4UpdateInfo.AddItem "    " & objIni.ReadData("EXCEL", arFileName(intloop), "0")
    If fso.FileExists(strPath & "\" & arFileName(intloop) & ".xls") Then
       Set fl_file = fso.GetFile(strPath & "\" & arFileName(intloop) & ".xls")
       lst_Tab4UpdateInfo.AddItem "    " & fl_file.DateCreated
       lst_Tab4UpdateInfo.AddItem "    " & fl_file.Size & " Bytes"
    Else
       lst_Tab4UpdateInfo.AddItem "    " & "檔案不存在"
    End If
    lst_Tab4UpdateInfo.AddItem ""
Next intloop
Set fl_file = Nothing
Set fso = Nothing

End Sub

Private Sub cmd_Tab4SysInfo_Click()
lst_Tab4SystemInfo.Clear
'取得系統更新資訊
Dim objIni As vbIniFile
Dim strTmp As String, strFileName As String, arFileName() As String

Set objIni = New vbIniFile
If Dir(striniFileName_FullPath, vbHidden + vbReadOnly) = "" Then
   msg_text = "指定設定檔 [" & striniFileName_FullPath & " 不存在" & vbCrLf & _
              "請通知系統維護人員"
   MsgBox msg_text, vbOKOnly, msg_title
   Exit Sub
End If

'指定 INI 檔案存放位置與檔案名稱
objIni.FileName = striniFileName_FullPath
'取得 FilePath , FileName
strFileName = objIni.ReadData("EXCEL", "FILENAME", "0")
arFileName = Split(strFileName, ",", -1, vbTextCompare)

'設定欄位顯示格式
Dim clnvalue(2) As Long, lstString As String
clnvalue(0) = 60        '欄位說明
clnvalue(1) = 1000      '欄位資料
Call ListBox_SetTabStops(lst_Tab4SystemInfo.hwnd, 2, clnvalue)
     
Set fso = New FileSystemObject
For intloop = LBound(arFileName) To UBound(arFileName)
    lst_Tab4SystemInfo.AddItem arFileName(intloop) & ".xls"
    lst_Tab4SystemInfo.AddItem "    " & objIni.ReadData("EXCEL", arFileName(intloop), "0")
    If fso.FileExists(App.Path & "\" & arFileName(intloop) & ".xls") Then
       Set fl_file = fso.GetFile(App.Path & "\" & arFileName(intloop) & ".xls")
       lst_Tab4SystemInfo.AddItem "    " & fl_file.DateCreated
       lst_Tab4SystemInfo.AddItem "    " & fl_file.Size & " Bytes"
    Else
       lst_Tab4SystemInfo.AddItem "    " & "檔案不存在"
    End If
    lst_Tab4SystemInfo.AddItem " "
Next intloop
Set fl_file = Nothing
Set fso = Nothing

End Sub

Private Sub cmd_Tab4Update_Click()
Dim objIni As vbIniFile
Dim strPath As String, strFileName As String, arFileName() As String

On Error GoTo err_Handle
Set objIni = New vbIniFile
If Dir(striniFileName_FullPath, vbHidden + vbReadOnly) = "" Then
   msg_text = "指定設定檔 [" & striniFileName_FullPath & " 不存在" & vbCrLf & _
              "請通知系統維護人員"
   MsgBox msg_text, vbOKOnly, msg_title
   Exit Sub
End If

'指定 INI 檔案存放位置與檔案名稱
objIni.FileName = striniFileName_FullPath
'取得 FilePath , FileName
strFileName = objIni.ReadData("EXCEL", "FILENAME", "0")
strPath = objIni.ReadData("EXCEL", "FILEPATH", "0")
arFileName = Split(strFileName, ",", -1, vbTextCompare)

Set fso = New FileSystemObject
For intloop = LBound(arFileName) To UBound(arFileName)
    If fso.FileExists(strPath & "\" & arFileName(intloop) & ".xls") Then
       fso.CopyFile strPath & "\" & arFileName(intloop) & ".xls", App.Path & "\", True
    Else
       msg_text = "Excel 範本檔 " & arFileName(intloop) & ".xls 不存在，請通知系統維護人員"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    End If
Next intloop
Set fso = Nothing
msg_text = "更新狀態：成功完成 Excel 範本檔更新作業"
MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'重新顯示 [系統資訊]
Call cmd_Tab4SysInfo_Click
Exit Sub

err_Handle:
   msg_text = "系統更新錯誤，錯誤訊息如下" & vbCrLf & err.Number & "：" & err.Description
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "系統更新作業"
End Sub

Private Sub Form_Load()
Me.Height = 6150: Me.Width = 11000
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200

End Sub

Private Sub Form_Terminate()
'更新 Menu [視窗]→[已開視窗清單]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'從記憶體中移除表單，藉此引起 [Terminate] 事件
Set frm_SystemUpdate = Nothing
End Sub

Private Function CheckLoginUser() As Boolean
Dim urName As String
urName = Get_LoginUserName   '取得 Login User Name
'需為 Administrator 方能執行 [Barcode 字型檔更新]
Select Case UCase(urName)
       Case "ADMINISTRATOR", "MINGSON", "CHWEN", "DANIEL"
            CheckLoginUser = True
       Case Else
            CheckLoginUser = False
End Select
End Function

