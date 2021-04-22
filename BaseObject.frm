VERSION 5.00
Begin VB.Form BaseObject 
   Caption         =   "BaseObject"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   11280
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame7 
      Caption         =   "功能"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   9615
      Begin VB.CommandButton cmdPreView 
         BackColor       =   &H00C0E0FF&
         Caption         =   "預覽"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   2880
         Picture         =   "BaseObject.frx":0000
         Style           =   1  '圖片外觀
         TabIndex        =   14
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0FFC0&
         Caption         =   "列印"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   4200
         Picture         =   "BaseObject.frx":1708A
         Style           =   1  '圖片外觀
         TabIndex        =   13
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdExport 
         BackColor       =   &H00FFC0C0&
         Caption         =   "轉Excel"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   5520
         Picture         =   "BaseObject.frx":17394
         Style           =   1  '圖片外觀
         TabIndex        =   12
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FFFFC0&
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
         Height          =   1245
         Left            =   240
         Picture         =   "BaseObject.frx":1868E
         Style           =   1  '圖片外觀
         TabIndex        =   11
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00C0FFC0&
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
         Height          =   1245
         Left            =   1560
         Picture         =   "BaseObject.frx":18998
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdApply 
         BackColor       =   &H00C0FFC0&
         Caption         =   "套用"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   6840
         Picture         =   "BaseObject.frx":18CAA
         Style           =   1  '圖片外觀
         TabIndex        =   9
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0FFC0&
         Caption         =   "確定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   8160
         Picture         =   "BaseObject.frx":1A9A4
         Style           =   1  '圖片外觀
         TabIndex        =   8
         Top             =   360
         Width           =   1185
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000C&
         BackStyle       =   1  '不透明
         BorderColor     =   &H80000006&
         BorderWidth     =   2
         Height          =   1485
         Left            =   120
         Top             =   240
         Width           =   9375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "功能"
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFF80&
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
         Height          =   1125
         Left            =   4200
         Picture         =   "BaseObject.frx":1C69E
         Style           =   1  '圖片外觀
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFC0FF&
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
         Height          =   1125
         Left            =   2880
         Picture         =   "BaseObject.frx":1C9A8
         Style           =   1  '圖片外觀
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFC0C0&
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
         Height          =   1125
         Left            =   1560
         Picture         =   "BaseObject.frx":1D9EA
         Style           =   1  '圖片外觀
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdAddNew 
         BackColor       =   &H00FFFFC0&
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
         Height          =   1125
         Left            =   240
         Picture         =   "BaseObject.frx":2423C
         Style           =   1  '圖片外觀
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FF80FF&
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
         Height          =   1125
         Left            =   6840
         Picture         =   "BaseObject.frx":26366
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   360
         Width           =   1185
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FF8080&
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   5520
         Picture         =   "BaseObject.frx":4FF78
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Top             =   360
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000C&
         BackStyle       =   1  '不透明
         BorderColor     =   &H80000006&
         BorderWidth     =   2
         Height          =   1365
         Left            =   120
         Top             =   240
         Width           =   8055
      End
   End
End
Attribute VB_Name = "BaseObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
