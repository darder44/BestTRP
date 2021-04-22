VERSION 5.00
Begin VB.Form frm_SwitchDB 
   Caption         =   "系統資料庫切換"
   ClientHeight    =   2880
   ClientLeft      =   3525
   ClientTop       =   2805
   ClientWidth     =   6180
   Icon            =   "frm_SwitchDB.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   6180
   Begin VB.CommandButton cmd_Switch 
      BackColor       =   &H00C0E0FF&
      Caption         =   "資料庫切換"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   3390
      Picture         =   "frm_SwitchDB.frx":030A
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   1590
      Width           =   1290
   End
   Begin VB.CommandButton cmd_Exit 
      BackColor       =   &H00FFC0FF&
      Caption         =   "離  開"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   4755
      Picture         =   "frm_SwitchDB.frx":0614
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   1605
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   15
      TabIndex        =   6
      Top             =   -30
      Width           =   6135
      Begin VB.ListBox lst_CnPropertyList 
         BackColor       =   &H00C0FFC0&
         Height          =   1140
         Left            =   390
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         Width           =   5685
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "連線狀態"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   795
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   330
         Width           =   210
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '平面
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   15
      TabIndex        =   8
      Top             =   1275
      Width           =   3270
      Begin VB.TextBox txt_NewPassword 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         IMEMode         =   3  '暫止
         Left            =   1695
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1155
         Width           =   1440
      End
      Begin VB.TextBox txt_NewUserName 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         IMEMode         =   3  '暫止
         Left            =   1695
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   825
         Width           =   1440
      End
      Begin VB.TextBox txt_NewDBName 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1695
         TabIndex        =   1
         Top             =   495
         Width           =   1440
      End
      Begin VB.TextBox txt_NewSrvName 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1695
         TabIndex        =   0
         Top             =   165
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "使用者密碼："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Index           =   3
         Left            =   435
         TabIndex        =   14
         Top             =   1215
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "使用者名稱："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Index           =   2
         Left            =   435
         TabIndex        =   12
         Top             =   885
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "資料庫名稱："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Index           =   1
         Left            =   435
         TabIndex        =   11
         Top             =   555
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  '透明
         Caption         =   "切換資料庫"
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
         Height          =   1020
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   345
         Width           =   210
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "伺服器名稱："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   195
         Index           =   0
         Left            =   435
         TabIndex        =   9
         Top             =   240
         Width           =   1260
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '不透明
      FillColor       =   &H00404000&
      Height          =   1485
      Left            =   3285
      Shape           =   4  '圓角矩形
      Top             =   1350
      Width           =   2865
   End
End
Attribute VB_Name = "frm_SwitchDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frmLoop As Integer

Private Sub cmd_Exit_Click()
Unload Me
End Sub

Private Sub cmd_Switch_Click()
If Not CheckLoginUser(4) Then
   msg_text = "先生！非常抱歉，您沒有執行此作業之權限"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Unload Me
   Exit Sub
End If

If Len(Trim(txt_NewSrvName.Text)) = 0 Or Len(Trim(txt_NewDBName.Text)) = 0 Or _
   Len(Trim(txt_NewUserName.Text)) = 0 Or Len(Trim(txt_NewPassword.Text)) = 0 Then
   msg_text = "資料庫切換所需資訊不全，請輸入所有連線所需資訊"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

'組合連線字串
cn.Close
cn_string = "Provider=SQLOLEDB;Persist Security Info=False;User ID=" & txt_NewUserName.Text & _
            ";Password=" & txt_NewPassword & ";Initial Catalog=" & txt_NewDBName.Text & ";Data Source=" & txt_NewSrvName.Text
Call DB_connect(cn_string)
Me.Caption = "資料庫切換作業     連線資料庫：" & cn.DefaultDatabase

'設定欄位顯示格式
Dim clnvalue(2) As Long, lstString As String
clnvalue(0) = 150       '屬性名稱
clnvalue(1) = 1000     '屬性設定值
Call ListBox_SetTabStops(lst_CnPropertyList.hwnd, 2, clnvalue)

lst_CnPropertyList.Clear
For frmLoop = 0 To cn.Properties.Count - 1
    lst_CnPropertyList.AddItem cn.Properties(frmLoop).Name & vbTab & cn.Properties(frmLoop).Value
Next frmLoop
'顯示目前連線資訊
txt_NewSrvName.Text = cn.Properties("Data Source Name").Value
txt_NewDBName.Text = cn.Properties("Current Catalog").Value
txt_NewUserName.Text = cn.Properties("User ID").Value

End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "資料庫切換作業"
End Sub

Private Sub Form_Load()
Me.Height = 3285: Me.Width = 6300
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200

'設定欄位顯示格式
Dim clnvalue(2) As Long, lstString As String
clnvalue(0) = 150       '屬性名稱
clnvalue(1) = 1000     '屬性設定值
Call ListBox_SetTabStops(lst_CnPropertyList.hwnd, 2, clnvalue)
'顯示 Connection 所有屬性值
lst_CnPropertyList.Clear
For frmLoop = 0 To cn.Properties.Count - 1
    lst_CnPropertyList.AddItem cn.Properties(frmLoop).Name & vbTab & cn.Properties(frmLoop).Value
Next frmLoop

'顯示目前連線資訊
txt_NewSrvName.Text = cn.Properties("Data Source Name").Value
txt_NewDBName.Text = cn.Properties("Current Catalog").Value
txt_NewUserName.Text = cn.Properties("User ID").Value

Me.Caption = "資料庫切換作業     連線資料庫：" & cn.DefaultDatabase
End Sub

Private Sub Form_Terminate()
'更新 Menu [視窗]→[已開視窗清單]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'從記憶體中移除表單，藉此引起 [Terminate] 事件
Set frm_SwitchDB = Nothing
End Sub

Private Sub txt_NewDBName_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
     Case 97 To 122   '小寫字元改為大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          txt_NewUserName.SetFocus
   End Select
End Sub

Private Sub txt_NewPassword_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
     Case vbKeyReturn
          KeyAscii = 0
          cmd_Switch.SetFocus
   End Select
End Sub

Private Sub txt_NewSrvName_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
     Case 97 To 122   '小寫字元改為大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          txt_NewDBName.SetFocus
   End Select
End Sub

Private Sub txt_NewUserName_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
     Case 97 To 122   '小寫字元改為大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
          txt_NewPassword.SetFocus
   End Select
End Sub


