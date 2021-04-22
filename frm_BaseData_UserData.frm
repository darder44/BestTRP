VERSION 5.00
Begin VB.Form frm_BaseData_UserData 
   Caption         =   "使用者資料維護 "
   ClientHeight    =   4230
   ClientLeft      =   2670
   ClientTop       =   2115
   ClientWidth     =   6075
   Icon            =   "frm_BaseData_UserData.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   6075
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Height          =   690
      Left            =   75
      TabIndex        =   1
      Top             =   -30
      Width           =   5925
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
         IMEMode         =   2  '關閉
         Left            =   1485
         TabIndex        =   0
         Top             =   210
         Width           =   1800
      End
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
         Height          =   360
         IMEMode         =   3  '暫止
         Left            =   4065
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   210
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "使用者帳號"
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
         TabIndex        =   4
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "密碼"
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
         TabIndex        =   3
         Top             =   255
         Width           =   510
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000000&
      Height          =   2265
      Left            =   75
      TabIndex        =   5
      Top             =   615
      Width           =   5925
      Begin VB.CheckBox chk_AdminCode 
         BackColor       =   &H8000000A&
         Caption         =   "系統管理員"
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
         Left            =   2520
         TabIndex        =   23
         Top             =   1800
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.ComboBox cmb_Company 
         Height          =   300
         Left            =   1485
         Style           =   2  '單純下拉式
         TabIndex        =   21
         Top             =   225
         Width           =   4275
      End
      Begin VB.CommandButton cmd_Clear 
         BackColor       =   &H00C0E0FF&
         Caption         =   "清除"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4785
         Style           =   1  '圖片外觀
         TabIndex        =   20
         Top             =   1770
         Width           =   915
      End
      Begin VB.TextBox txt_Notes 
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
         IMEMode         =   2  '關閉
         Left            =   1485
         TabIndex        =   17
         Top             =   1290
         Width           =   4260
      End
      Begin VB.CheckBox chk_CloseCode 
         BackColor       =   &H8000000A&
         Caption         =   "停用"
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
         Left            =   1575
         TabIndex        =   15
         Top             =   1800
         Width           =   960
      End
      Begin VB.TextBox txt_Name 
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
         IMEMode         =   2  '關閉
         Left            =   1485
         TabIndex        =   12
         Top             =   900
         Width           =   4260
      End
      Begin VB.ComboBox cmb_Group 
         Height          =   300
         Left            =   1485
         Style           =   2  '單純下拉式
         TabIndex        =   8
         Top             =   555
         Width           =   4275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "所屬公司："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   22
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "帳號狀態："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   18
         Top             =   1785
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "備註說明："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   210
         TabIndex        =   16
         Top             =   1365
         Width           =   1275
      End
      Begin VB.Label lbl_GroupName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "帳號群組："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   210
         TabIndex        =   7
         Top             =   600
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
         Height          =   240
         Left            =   210
         TabIndex        =   6
         Top             =   945
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00008000&
      Height          =   1335
      Left            =   75
      TabIndex        =   9
      Top             =   2835
      Width           =   5925
      Begin VB.CommandButton cmd_Query 
         BackColor       =   &H00C0FFFF&
         Caption         =   "查  詢"
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
         Height          =   945
         Left            =   3585
         Picture         =   "frm_BaseData_UserData.frx":030A
         Style           =   1  '圖片外觀
         TabIndex        =   19
         Top             =   255
         Width           =   1035
      End
      Begin VB.CommandButton cmd_Delete 
         BackColor       =   &H00C0FFC0&
         Caption         =   "刪  除"
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
         Height          =   945
         Left            =   2445
         Picture         =   "frm_BaseData_UserData.frx":0614
         Style           =   1  '圖片外觀
         TabIndex        =   14
         Top             =   255
         Width           =   1035
      End
      Begin VB.CommandButton cmd_Add 
         BackColor       =   &H00FFFFC0&
         Caption         =   "新  增"
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
         Height          =   945
         Left            =   1290
         Picture         =   "frm_BaseData_UserData.frx":091E
         Style           =   1  '圖片外觀
         TabIndex        =   13
         Top             =   255
         Width           =   1035
      End
      Begin VB.CommandButton cmd_Save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改存檔"
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
         Left            =   135
         Picture         =   "frm_BaseData_UserData.frx":0C28
         Style           =   1  '圖片外觀
         TabIndex        =   11
         Top             =   255
         Width           =   1035
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
         Height          =   945
         Left            =   4755
         Picture         =   "frm_BaseData_UserData.frx":0F32
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Top             =   255
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frm_BaseData_UserData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private arCompanyID() As String
Private argroupID() As String
Private arFacility() As String
Private locUserID As String

Private Sub cmd_Add_Click()
'新增
On Error GoTo err_handle

'新增作業資料查核
If CheckData() = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

'檢查 [使用者代號] 是否重複
str_SQL = "Select Rtrim(user_LoginID) as 'UserID' From CodeUser Where user_LoginID = '" & RTrim(txt_UserID.Text) & "'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "使用者代號 [" & txt_UserID.Text & "] 資料已經存在，使用者代號不允許重複"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   If GetUserData(txt_UserID.Text) = False Then
      MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
      Exit Sub
   End If
   Exit Sub
End If
tmp_Rs.Close

'新增作業
Tran_Level = cn.BeginTrans
str_SQL = "Insert into CodeUser (user_LoginID,user_Password,user_Status,user_Company,user_Group,user_Name,user_Facility,user_Notes,user_AddDate,user_AddWho) Values (" & _
          "'" & Trim(txt_UserID.Text) & "','" & RTrim(txt_Password.Text) & "','" & IIf(chk_CloseCode.Value = vbChecked, "0", "1") & "','" & arCompanyID(cmb_Company.ListIndex) & "','" & _
          argroupID(cmb_Group.ListIndex) & "','" & RTrim(txt_Name.Text) & "','" & "" & "','" & Trim(txt_Notes.Text) & "',Getdate(),'" & User_id & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
cn.CommitTrans
Tran_Level = 0
'清螢幕
Call cmd_Clear_Click
Exit Sub

err_handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-新增", Me.Caption, "cmd_Add_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Clear_Click()
'清螢幕
Call ClearForm_AllField(frm_BaseData_UserData)
locUserID = ""
If blAdmin Then
   txt_UserID.SetFocus
End If
End Sub

Private Sub cmd_Delete_Click()
'刪除
If Len(Trim(txt_UserID.Text)) = 0 Then Exit Sub
msg_text = "確認刪除使用者代號 [" & RTrim(txt_UserID.Text) & "]？"
If MsgBox(msg_text, vbOKCancel + vbInformation, msg_title) = vbCancel Then Exit Sub
Tran_Level = cn.BeginTrans

'刪除使用者之權限設定資料
str_SQL = "Delete From CodeRole Where Rtrim(user_LoginID) = '" & txt_UserID.Text & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'刪除使用者基本資料
str_SQL = "Delete From CodeUser Where Rtrim(user_LoginID) = '" & txt_UserID.Text & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
cn.CommitTrans
Tran_Level = 0
'清螢幕
Call cmd_Clear_Click
Exit Sub

err_handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-新增", Me.Caption, "cmd_Add_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Exit_Click()
'離開
Unload Me
End Sub

Private Sub cmd_Query_Click()
'查詢
txt_UserID.Text = Trim(txt_UserID.Text)
If Len(txt_UserID.Text) = 0 Then
   msg_text = "資料錯誤：查詢請輸入 [使用者代號]"
   txt_UserID.SetFocus
   Exit Sub
End If
'取出使用者明細資料
If GetUserData(Trim(txt_UserID.Text)) = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
End If
locUserID = txt_UserID.Text
txt_UserID.SelStart = 0: txt_UserID.SelLength = Len(txt_UserID.Text): txt_UserID.SetFocus
End Sub

Private Sub cmd_Save_Click()
'修改存檔
If locUserID = "" Then
   msg_text = "使用者資料修改程序：" & vbCrLf & "輸入使用者代號 → [查詢] → 修改資料 → [修改存檔]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

'修改作業資料查核
If CheckData() = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

'修改存檔
Tran_Level = cn.BeginTrans
If blAdmin Then
   '系統管理員：先刪除後新增
   str_SQL = "Delete From CodeUser Where user_LoginID = '" & locUserID & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   str_SQL = "Insert into CodeUser (user_LoginID,user_Password,user_Status,user_Company,user_Group,user_Name,user_Facility,user_Notes,user_AddDate,user_AddWho) Values (" & _
             "'" & Trim(txt_UserID.Text) & "','" & RTrim(txt_Password.Text) & "','" & IIf(chk_CloseCode.Value = vbChecked, "0", "1") & "','" & arCompanyID(cmb_Company.ListIndex) & "','" & _
             argroupID(cmb_Group.ListIndex) & "','" & RTrim(txt_Name.Text) & "','" & "" & "','" & Trim(txt_Notes.Text) & "',Getdate(),'" & User_id & "')"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   str_SQL = "Update CodeRole Set user_LoginID = '" & Trim(txt_UserID.Text) & "' Where user_LoginID = '" & locUserID & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
Else
   '非系統管理人員只允許修改：Password , Name , Notes
    str_SQL = "Update CodeUser Set user_Password='" & Trim(txt_Password.Text) & "'," & _
              "user_Name='" & Trim(txt_Name.Text) & "',user_Notes = '" & txt_Notes.Text & "',user_AddDate=Getdate(),user_AddWho='" & User_id & "' Where user_LoginID = '" & locUserID & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
End If
cn.CommitTrans
Tran_Level = 0

'清螢幕
Call cmd_Clear_Click

Exit Sub

err_handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-修改存檔", Me.Caption, "cmd_Save_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "使用者基本資料維護"
End Sub

Private Sub Form_Load()
'設定 Form 大小、位置
Me.Height = 4740: Me.Width = 6200
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200

'取出所有公司資料
Dim tmp_cnt As Integer
cmb_Company.Clear
str_SQL = "Select Rtrim(Code) as 'Code',Rtrim(Description) as 'CompanyName' From CodeLKUP Where ListName = 'USERCOMPANY' Order by Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim arCompanyID(10) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      arCompanyID(tmp_cnt) = tmp_Rs.Fields("Code").Value
      cmb_Company.AddItem tmp_Rs.Fields("Code").Value & Space(10 - Len(Trim(tmp_Rs.Fields("Code").Value))) & tmp_Rs.Fields("CompanyName").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(arCompanyID) Then
         ReDim Preserve arCompanyID(UBound(arCompanyID) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close

'取出所有群組資料
cmb_Group.Clear
str_SQL = "Select Rtrim(Code) as 'Code',Rtrim(Description) as 'GroupName' From CodeLKUP Where ListName = 'USERGROUP' Order by Code"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
ReDim argroupID(10) As String
If Not tmp_Rs.EOF Then
   tmp_cnt = 0
   Do While Not tmp_Rs.EOF
      argroupID(tmp_cnt) = tmp_Rs.Fields("Code").Value
      cmb_Group.AddItem tmp_Rs.Fields("Code").Value & Space(10 - Len(Trim(tmp_Rs.Fields("Code").Value))) & tmp_Rs.Fields("GroupName").Value
      tmp_Rs.MoveNext
      tmp_cnt = tmp_cnt + 1
      If tmp_cnt = UBound(argroupID) Then
         ReDim Preserve argroupID(UBound(argroupID) + 10) As String
      End If
   Loop
End If
tmp_Rs.Close

'取出使用者明細資料
If GetUserData(User_id) = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
End If
locUserID = User_id

'設定功能
If blAdmin Then
   '系統管理員可以新增、刪除，變更群組，權限設定，狀態
   cmd_Add.Enabled = True           '新增
   cmd_Delete.Enabled = True        '刪除
   cmd_Query.Enabled = True         '查詢
   cmd_Clear.Enabled = True         '清螢幕
   cmb_Company.Enabled = True       '使用者代號歸屬公司
   cmb_Group.Enabled = True         '使用者代號歸屬群組
   chk_CloseCode.Enabled = True     '狀態：停用
   chk_AdminCode.Enabled = True     '使用權限：系統管理員
Else
   '除 [系統管理員] 之外的所有 User 只允許修改 Password
   txt_UserID.Enabled = False
   cmd_Add.Enabled = False
   cmd_Delete.Enabled = False
   cmd_Query.Enabled = False
   cmd_Clear.Enabled = False
   cmb_Company.Enabled = False
   cmb_Group.Enabled = False
   chk_CloseCode.Enabled = False
   chk_AdminCode.Enabled = False
End If

End Sub

Private Sub Form_Terminate()
'更新 Menu [視窗]→[已開視窗清單]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'從記憶體中移除表單，藉此引起 [Terminate] 事件
Set frm_BaseData_UserData = Nothing
End Sub

Private Sub txt_UserID_KeyPress(KeyAscii As Integer)
'使用者代號
If KeyAscii = vbKeyReturn Then
   txt_Password.SelStart = 0
   txt_Password.SelLength = Len(txt_Password.Text)
   txt_Password.SetFocus
End If
End Sub

Private Function GetUserData(strUserID As String) As Boolean
'取出使用者明細資料
Dim tmpI As Integer
On Error GoTo err_handle
Call ClearForm_AllField(frm_BaseData_UserData)
GetUserData = False
str_SQL = "Select Rtrim(a.user_LoginID) as 'UserID' , Rtrim(a.user_Password) as 'Password' , Rtrim(a.user_Name) as 'UserName' , Rtrim(b.Description) as 'GroupName' , " & _
          "       Rtrim(a.user_Group) as 'GroupID' , Rtrim(c.Description) as 'CompanyName' , Rtrim(a.user_Company) as 'CompanyID' , Rtrim(a.user_Status) as 'UserStatus', " & _
          "       Rtrim(Cast(isnull(a.user_Notes,'') as varchar(300))) as 'Notes'  " & _
          "From CodeUSER a " & _
          "Inner join CodeLKUP b on b.ListName = 'USERGROUP' and Rtrim(b.Code) = Rtrim(a.user_Group) " & _
          "Inner join CodeLKUP c on c.ListName = 'USERCOMPANY' and Rtrim(c.Code) = Rtrim(a.user_Company) " & _
          "Where a.user_LoginID = '" & strUserID & "'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
If Not tmp_Rs.EOF Then
   txt_UserID.Text = strUserID
   txt_Password.Text = tmp_Rs.Fields("Password").Value
   '找出 User 所屬的 Company
   For tmpI = 0 To UBound(arCompanyID) - 1
       If arCompanyID(tmpI) = tmp_Rs.Fields("CompanyID").Value Then
          cmb_Company.ListIndex = tmpI
       End If
   Next tmpI
   '找出 User 所屬的 Group
   For tmpI = 0 To UBound(argroupID) - 1
       If argroupID(tmpI) = tmp_Rs.Fields("GroupID").Value Then
          cmb_Group.ListIndex = tmpI
       End If
   Next tmpI
   
   txt_Name.Text = tmp_Rs.Fields("UserName").Value
   If tmp_Rs.Fields("UserStatus").Value = "1" Then  '帳號是否已經停用
      chk_CloseCode.Value = vbUnchecked
   Else
      chk_CloseCode.Value = vbChecked
   End If
   txt_Notes.Text = tmp_Rs.Fields("Notes").Value
   If tmp_Rs.Fields("GroupID").Value = "ADMIN" Then  '系統管理員權限
      chk_AdminCode.Value = vbChecked
      blAdmin = True
   Else
'      chk_AdminCode.Value = vbUnchecked
'      blAdmin = False
   End If
Else
   funRtn_msg = "使用者資料錯誤：使用者代號 [" & strUserID & "] 資料不存在"
   Exit Function
End If
locUserID = strUserID
GetUserData = True
Exit Function

err_handle:
   funRtn_msg = "Function [Private GetUserData()] RunTime Error：" & vbCrLf & _
                "Err Code：" & err.Number & vbCrLf & "Err Descr：" & err.Description
End Function

Private Function CheckData() As Boolean
'資料查核
On Error GoTo err_handle
CheckData = False
If Len(RTrim(txt_UserID.Text)) = 0 Then
   funRtn_msg = "新增作業資料錯誤：使用者代號不能為空白"
   txt_UserID.SetFocus
   Exit Function
End If
If cmb_Company.ListIndex = -1 Then
   funRtn_msg = "新增作業資料錯誤：必須選取使用者之 [歸屬公司] 作為權限控管依據"
   cmb_Company.SetFocus
   Exit Function
End If
If cmb_Group.ListIndex = -1 Then
   funRtn_msg = "新增作業資料錯誤：必須選取使用者之 [歸屬群組] 作為權限控管依據"
   cmb_Group.SetFocus
   Exit Function
End If
If Len(Trim(txt_Name.Text)) = 0 Then
   funRtn_msg = "新增作業資料錯誤：使用者 [姓名] 不能為空白"
   txt_Name.SetFocus
   Exit Function
End If
CheckData = True
Exit Function

err_handle:
   funRtn_msg = "Function [Private CheckData()] RunTime Error：" & vbCrLf & _
                "Err Code：" & err.Number & vbCrLf & "Err Descr：" & err.Description
End Function
