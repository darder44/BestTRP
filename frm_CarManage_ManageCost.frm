VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_CarManage_ManageCost 
   Caption         =   "車輛管理費"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   10515
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
      Height          =   915
      Left            =   8625
      Picture         =   "frm_CarManage_ManageCost.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   0
      Width           =   1035
   End
   Begin VB.CommandButton cmd_Save 
      BackColor       =   &H00FFC0C0&
      Caption         =   "存  檔"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   7080
      Picture         =   "frm_CarManage_ManageCost.frx":0442
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   120
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   4260
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   9825
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gd_ProdSec 
         Height          =   4095
         Left            =   30
         TabIndex        =   2
         Top             =   120
         Width           =   9750
         _ExtentX        =   17198
         _ExtentY        =   7223
         _Version        =   393216
         BackColor       =   -2147483624
         Rows            =   10
         Cols            =   9
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
      End
   End
   Begin VB.CommandButton cmd_Query 
      Caption         =   "查 詢"
      Height          =   975
      Left            =   4920
      TabIndex        =   0
      Top             =   0
      Width           =   1080
   End
End
Attribute VB_Name = "frm_CarManage_ManageCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
