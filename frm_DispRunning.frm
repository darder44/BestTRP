VERSION 5.00
Begin VB.Form frm_DispRunning 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "系統狀態"
   ClientHeight    =   2160
   ClientLeft      =   3165
   ClientTop       =   3060
   ClientWidth     =   6195
   ControlBox      =   0   'False
   Icon            =   "frm_DispRunning.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '螢幕中央
   Begin VB.ListBox lst_Msg 
      BackColor       =   &H00C0C0FF&
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
      Height          =   1620
      Left            =   30
      TabIndex        =   0
      Top             =   495
      Width           =   6120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "執行中，請稍後 - - - - - - - - - - - - - "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   285
      Left            =   855
      TabIndex        =   1
      Top             =   90
      Width           =   4725
   End
End
Attribute VB_Name = "frm_DispRunning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

