VERSION 5.00
Begin VB.Form frm_UserLogin 
   BackColor       =   &H00C0C0C0&
   Caption         =   " 系 統 登 入 程 序 "
   ClientHeight    =   3660
   ClientLeft      =   3165
   ClientTop       =   2715
   ClientWidth     =   6195
   Icon            =   "frm_UserLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6195
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   30
      Width           =   5925
      Begin VB.TextBox txt_Password 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  '暫止
         Left            =   4185
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   420
         Width           =   1560
      End
      Begin VB.TextBox txt_UserID 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  '暫止
         Left            =   1485
         TabIndex        =   0
         Top             =   420
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "密   碼"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   3390
         TabIndex        =   6
         Top             =   465
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "使用者代號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   480
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmd_Login 
      BackColor       =   &H00C0FFFF&
      Caption         =   "登入系統"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   1515
      Picture         =   "frm_UserLogin.frx":0442
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   2655
      Width           =   1755
   End
   Begin VB.CommandButton cmd_Cancel 
      BackColor       =   &H00FFFFC0&
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
      Height          =   930
      Left            =   3450
      Picture         =   "frm_UserLogin.frx":074C
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   2655
      Width           =   1755
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   1485
      Left            =   120
      TabIndex        =   7
      Top             =   1065
      Width           =   5925
      Begin VB.Label lbl_CompanyName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "歸屬公司："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   165
         TabIndex        =   10
         Top             =   285
         Width           =   1275
      End
      Begin VB.Label lbl_UserName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "姓　　名："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   165
         TabIndex        =   9
         Top             =   1035
         Width           =   1275
      End
      Begin VB.Label lbl_GroupName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "歸屬群組："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   165
         TabIndex        =   8
         Top             =   660
         Width           =   1275
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404000&
      BackStyle       =   1  '不透明
      BorderColor     =   &H00400040&
      BorderWidth     =   2
      Height          =   1050
      Index           =   1
      Left            =   3375
      Top             =   2595
      Width           =   1905
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400040&
      BackStyle       =   1  '不透明
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      Height          =   1050
      Index           =   0
      Left            =   1440
      Top             =   2595
      Width           =   1905
   End
End
Attribute VB_Name = "frm_UserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim locstrPassword As String

Dim obj As Object

Private Sub cmd_Cancel_Click()
'取消
Call DB_Disconnect(cn)
End
End Sub

Private Sub cmd_Login_Click()
'登入系統
On Error GoTo err_Handle

txt_UserID.Text = Trim(txt_UserID.Text)
If Len(txt_UserID.Text) = 0 Then Exit Sub

Screen.MousePointer = vbHourglass

'檢查密碼是否正確
If locstrPassword <> Trim(txt_Password.Text) Then
   msg_text = "密碼錯誤"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   txt_Password.SelStart = 0: txt_Password.SelLength = Len(txt_Password.Text): txt_Password.SetFocus
   Exit Sub
End If

'設定 功能表
If blAdmin Then
   For Each obj In frm_MDIForm.Controls
       If TypeName(obj) = "Menu" Then
          If Right(obj.Name, 1) <> "x" Then
             obj.Enabled = True
          End If
       End If
   Next
Else
   Call Confirm_Recordset_Closed(tmp_Rs)
   str_SQL = "Select APCode,role_RUN From CodeRole(nolock) Where user_LoginID = '" & myFilter(txt_UserID.Text) & "' and role_RUN = '1'"
   tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
   If tmp_Rs.EOF Then
      tmp_Rs.Close
      msg_text = "權限設定錯誤：User 未有任何可執行的權限設定"
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
      Screen.MousePointer = vbDefault
      txt_Password.Text = ""
      txt_UserID.SelStart = 0: txt_UserID.SelLength = Len(txt_UserID.Text): txt_UserID.SetFocus
      Exit Sub
   End If
   For Each obj In frm_MDIForm.Controls
       If TypeName(obj) = "Menu" Then
          tmp_Rs.Filter = adFilterNone
          tmp_Rs.Filter = "(APCode = '" & obj.Name & "' and role_RUN = '1')"
          If Not tmp_Rs.EOF Then
             obj.Enabled = True
          End If
       End If
   Next
   tmp_Rs.Filter = adFilterNone
   tmp_Rs.Close
End If

Screen.MousePointer = vbDefault
Unload Me
Exit Sub

err_Handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-登入系統", Me.Caption, "cmd_Login_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
  msg_title = "User Login 作業"
End Sub

Private Sub Form_Load()
'設定 Form 大小、位置
Me.Height = 4170 + 100: Me.Width = 6300
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm_UserLogin = Nothing
End Sub


Private Sub txt_Password_KeyPress(KeyAscii As Integer)
'密碼
If KeyAscii = vbKeyReturn Then Call cmd_Login_Click
'   KeyAscii = 0
'   cmd_Login.SetFocus
'End If
End Sub

Private Sub txt_UserID_KeyPress(KeyAscii As Integer)
'使用者單號
'Select Case KeyAscii
'Case 65 To 90
'     KeyAscii = KeyAscii + 32
'Case vbKeyReturn
'     txt_Password.SelStart = 0
'     txt_Password.SelLength = Len(txt_Password.Text)
'     txt_Password.SetFocus
'End Select
End Sub

Private Sub txt_UserID_LostFocus()
'取得使用者相關資料
txt_UserID.Text = Trim(txt_UserID.Text)
If Len(txt_UserID.Text) = 0 Then Exit Sub

locstrPassword = ""
If Len(Trim(txt_UserID.Text)) = 0 Then Exit Sub
str_SQL = "select Rtrim(a.user_Password) as 'Password' , Rtrim(a.user_Name) as 'UserName'  , Rtrim(b.Description) as 'GroupName' " & _
         ", Company = a.user_company , Rtrim(a.user_Group) as 'Groupid' , Rtrim(a.user_Status) as 'CloseCode' , Rtrim(c.Description) as 'CompanyName' " & _
         "From CodeUSER a " & _
         "Inner Join CodeLKUP b on b.ListName = 'USERGROUP' and b.Code = a.user_Group " & _
         "Inner Join CodeLKUP c on c.ListName = 'USERCOMPANY' and c.Code = a.user_Company " & _
         "Where a.user_LoginID = '" & Trim(myFilter(txt_UserID.Text)) & "'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
   '顯示使用者資料
   lbl_CompanyName.Caption = "歸屬公司：" & tmp_Rs.Fields("CompanyName").Value
   lbl_GroupName.Caption = "歸屬群組：" & tmp_Rs.Fields("GroupName").Value
   lbl_UserName.Caption = "姓　　名：" & tmp_Rs.Fields("UserName").Value
   User_id = Trim(myFilter(UCase(txt_UserID.Text)))
   User_Name = Trim(tmp_Rs("UserName"))
   Group_id = tmp_Rs.Fields("GroupID").Value
   locstrPassword = tmp_Rs.Fields("Password").Value
   Company_id = RTrim(tmp_Rs("company"))
   If Group_id = "ADMIN" Then
      blAdmin = True
   Else
      blAdmin = False
   End If
   '停用註記
   If tmp_Rs.Fields("CloseCode").Value = "0" Then
      msg_text = "登入失敗：使用者帳號已停用！"
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
      txt_UserID.SelStart = 0: txt_UserID.SelLength = Len(txt_UserID.Text)
      txt_UserID.SetFocus
      Exit Sub
   End If
Else
   msg_text = "登入失敗：使用者帳號不存在！"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_UserID.SelStart = 0: txt_UserID.SelLength = Len(txt_UserID.Text)
   txt_UserID.SetFocus
   Exit Sub
End If
tmp_Rs.Close
End Sub
