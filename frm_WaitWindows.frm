VERSION 5.00
Begin VB.Form frm_WaitWindows 
   BorderStyle     =   3  '雙線固定對話方塊
   Caption         =   "系統狀態‧‧‧‧‧‧"
   ClientHeight    =   2610
   ClientLeft      =   3420
   ClientTop       =   3315
   ClientWidth     =   6180
   Icon            =   "frm_WaitWindows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_Cancel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "取  消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   60
      Picture         =   "frm_WaitWindows.frx":030A
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   1260
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5610
      Top             =   150
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   4380
      Picture         =   "frm_WaitWindows.frx":0BD4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   285
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   1575
      Picture         =   "frm_WaitWindows.frx":0EDE
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   270
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1200
      TabIndex        =   4
      Top             =   1935
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1200
      TabIndex        =   3
      Top             =   1590
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "00000"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1200
      TabIndex        =   2
      Top             =   1230
      Width           =   600
   End
End
Attribute VB_Name = "frm_WaitWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dtStart As Date
Dim updLabel3 As Boolean

Private Sub cmd_Cancel_Click()
Select Case UCase(Me.Tag)
       Case "TRANSFERTOEXCEL"
            fgTransferToExcel = False
End Select

End Sub

Private Sub Form_Activate()
dtStart = Now: updLabel3 = True
Select Case UCase(Me.Tag)
       Case "FRM_MDIFORM"
            Label1.Caption = "與資料庫建立連線中，請稍後‧‧‧‧‧‧"
            Label2.Caption = "連線起始時間：" & Format(Now, "yyyy/mm/dd ttttt")
            Label3.Caption = "連線時間： 00 分 00 秒"
            Call DB_connect
            updLabel3 = False
            Unload Me
       Case "IMPORTDATA"       '資料匯入作業
            Label1.Caption = "資料匯入作業執行中，請稍後‧‧‧‧‧‧"
            Label2.Caption = "起始時間：" & Format(Now, "yyyy/mm/dd ttttt")
            Label3.Caption = "匯入時間： 00 分 00 秒"
       Case "TRANSFERTOEXCEL"  '轉資料至 Excel
            Label1.Caption = "[轉存 Excel 檔] 作業執行中，請稍後‧‧‧‧‧‧"
            Label2.Caption = "起始時間：" & Format(Now, "yyyy/mm/dd ttttt")
            Label3.Caption = "轉檔時間： 00 分 00 秒"
            cmd_Cancel.Visible = True
       Case Else
            Label1.Caption = "系統作業執行中，請稍後‧‧‧‧‧‧"
            Label2.Caption = "起始時間：" & Format(Now, "yyyy/mm/dd ttttt")
            Label3.Caption = "查詢時間： 00 分 00 秒"
End Select
End Sub

Private Sub Form_Load()
Me.Height = 3000: Me.Width = 6300
Me.Left = ((frm_MDIForm.ScaleWidth - Me.Width) / 2) + 600
Me.Top = ((frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2) - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm_WaitWindows = Nothing
End Sub

Private Sub Timer1_Timer()
Select Case UCase(Me.Tag)
       Case "FRM_MDIFORM"
            If updLabel3 = True Then
               Label3.Caption = "連線時間：  " & Format(Now - dtStart, "nn") & " 分 " & Format(Now - dtStart, "ss") & " 秒"
            End If
       Case "TRANSFERTOEXCEL"
            If updLabel3 = True Then
               Label3.Caption = "轉檔時間：  " & Format(Now - dtStart, "nn") & " 分 " & Format(Now - dtStart, "ss") & " 秒"
            End If
       Case "IMPORTDATA"
            If updLabel3 = True Then
               Label3.Caption = "匯入時間：  " & Format(Now - dtStart, "nn") & " 分 " & Format(Now - dtStart, "ss") & " 秒"
            End If
      Case Else
            If updLabel3 = True Then
               Label3.Caption = "執行時間：  " & Format(Now - dtStart, "nn") & " 分 " & Format(Now - dtStart, "ss") & " 秒"
            End If
End Select
End Sub
