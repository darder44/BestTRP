VERSION 5.00
Begin VB.Form frm_BaseData_UserData 
   Caption         =   "�ϥΪ̸�ƺ��@ "
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
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   2  '����
         Left            =   1485
         TabIndex        =   0
         Top             =   210
         Width           =   1800
      End
      Begin VB.TextBox txt_Password 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  '�Ȥ�
         Left            =   4065
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   210
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ϥΪ̱b��"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         BackStyle       =   0  '�z��
         Caption         =   "�K�X"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Caption         =   "�t�κ޲z��"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   21
         Top             =   225
         Width           =   4275
      End
      Begin VB.CommandButton cmd_Clear 
         BackColor       =   &H00C0E0FF&
         Caption         =   "�M��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4785
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   20
         Top             =   1770
         Width           =   915
      End
      Begin VB.TextBox txt_Notes 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   2  '����
         Left            =   1485
         TabIndex        =   17
         Top             =   1290
         Width           =   4260
      End
      Begin VB.CheckBox chk_CloseCode 
         BackColor       =   &H8000000A&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   2  '����
         Left            =   1485
         TabIndex        =   12
         Top             =   900
         Width           =   4260
      End
      Begin VB.ComboBox cmb_Group 
         Height          =   300
         Left            =   1485
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   8
         Top             =   555
         Width           =   4275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���ݤ��q�G"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         BackStyle       =   0  '�z��
         Caption         =   "�b�����A�G"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         BackStyle       =   0  '�z��
         Caption         =   "�Ƶ������G"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         BackStyle       =   0  '�z��
         Caption         =   "�b���s�աG"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         BackStyle       =   0  '�z��
         Caption         =   "�m�@�@�W�G"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Caption         =   "�d  ��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   19
         Top             =   255
         Width           =   1035
      End
      Begin VB.CommandButton cmd_Delete 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�R  ��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   14
         Top             =   255
         Width           =   1035
      End
      Begin VB.CommandButton cmd_Add 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�s  �W"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   13
         Top             =   255
         Width           =   1035
      End
      Begin VB.CommandButton cmd_Save 
         BackColor       =   &H00FFC0C0&
         Caption         =   "�ק�s��"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   11
         Top             =   255
         Width           =   1035
      End
      Begin VB.CommandButton cmd_Exit 
         BackColor       =   &H00FFC0FF&
         Caption         =   "��  �}"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         Style           =   1  '�Ϥ��~�[
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
'�s�W
On Error GoTo err_handle

'�s�W�@�~��Ƭd��
If CheckData() = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

'�ˬd [�ϥΪ̥N��] �O�_����
str_SQL = "Select Rtrim(user_LoginID) as 'UserID' From CodeUser Where user_LoginID = '" & RTrim(txt_UserID.Text) & "'"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
   tmp_Rs.Close
   msg_text = "�ϥΪ̥N�� [" & txt_UserID.Text & "] ��Ƥw�g�s�b�A�ϥΪ̥N�������\����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   If GetUserData(txt_UserID.Text) = False Then
      MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
      Exit Sub
   End If
   Exit Sub
End If
tmp_Rs.Close

'�s�W�@�~
Tran_Level = cn.BeginTrans
str_SQL = "Insert into CodeUser (user_LoginID,user_Password,user_Status,user_Company,user_Group,user_Name,user_Facility,user_Notes,user_AddDate,user_AddWho) Values (" & _
          "'" & Trim(txt_UserID.Text) & "','" & RTrim(txt_Password.Text) & "','" & IIf(chk_CloseCode.Value = vbChecked, "0", "1") & "','" & arCompanyID(cmb_Company.ListIndex) & "','" & _
          argroupID(cmb_Group.ListIndex) & "','" & RTrim(txt_Name.Text) & "','" & "" & "','" & Trim(txt_Notes.Text) & "',Getdate(),'" & User_id & "')"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
cn.CommitTrans
Tran_Level = 0
'�M�ù�
Call cmd_Clear_Click
Exit Sub

err_handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�s�W", Me.Caption, "cmd_Add_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Clear_Click()
'�M�ù�
Call ClearForm_AllField(frm_BaseData_UserData)
locUserID = ""
If blAdmin Then
   txt_UserID.SetFocus
End If
End Sub

Private Sub cmd_Delete_Click()
'�R��
If Len(Trim(txt_UserID.Text)) = 0 Then Exit Sub
msg_text = "�T�{�R���ϥΪ̥N�� [" & RTrim(txt_UserID.Text) & "]�H"
If MsgBox(msg_text, vbOKCancel + vbInformation, msg_title) = vbCancel Then Exit Sub
Tran_Level = cn.BeginTrans

'�R���ϥΪ̤��v���]�w���
str_SQL = "Delete From CodeRole Where Rtrim(user_LoginID) = '" & txt_UserID.Text & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'�R���ϥΪ̰򥻸��
str_SQL = "Delete From CodeUser Where Rtrim(user_LoginID) = '" & txt_UserID.Text & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
cn.CommitTrans
Tran_Level = 0
'�M�ù�
Call cmd_Clear_Click
Exit Sub

err_handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�s�W", Me.Caption, "cmd_Add_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Exit_Click()
'���}
Unload Me
End Sub

Private Sub cmd_Query_Click()
'�d��
txt_UserID.Text = Trim(txt_UserID.Text)
If Len(txt_UserID.Text) = 0 Then
   msg_text = "��ƿ��~�G�d�߽п�J [�ϥΪ̥N��]"
   txt_UserID.SetFocus
   Exit Sub
End If
'���X�ϥΪ̩��Ӹ��
If GetUserData(Trim(txt_UserID.Text)) = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
End If
locUserID = txt_UserID.Text
txt_UserID.SelStart = 0: txt_UserID.SelLength = Len(txt_UserID.Text): txt_UserID.SetFocus
End Sub

Private Sub cmd_Save_Click()
'�ק�s��
If locUserID = "" Then
   msg_text = "�ϥΪ̸�ƭק�{�ǡG" & vbCrLf & "��J�ϥΪ̥N�� �� [�d��] �� �ק��� �� [�ק�s��]"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

'�ק�@�~��Ƭd��
If CheckData() = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If

'�ק�s��
Tran_Level = cn.BeginTrans
If blAdmin Then
   '�t�κ޲z���G���R����s�W
   str_SQL = "Delete From CodeUser Where user_LoginID = '" & locUserID & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   str_SQL = "Insert into CodeUser (user_LoginID,user_Password,user_Status,user_Company,user_Group,user_Name,user_Facility,user_Notes,user_AddDate,user_AddWho) Values (" & _
             "'" & Trim(txt_UserID.Text) & "','" & RTrim(txt_Password.Text) & "','" & IIf(chk_CloseCode.Value = vbChecked, "0", "1") & "','" & arCompanyID(cmb_Company.ListIndex) & "','" & _
             argroupID(cmb_Group.ListIndex) & "','" & RTrim(txt_Name.Text) & "','" & "" & "','" & Trim(txt_Notes.Text) & "',Getdate(),'" & User_id & "')"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
   
   str_SQL = "Update CodeRole Set user_LoginID = '" & Trim(txt_UserID.Text) & "' Where user_LoginID = '" & locUserID & "'"
   cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
Else
   '�D�t�κ޲z�H���u���\�ק�GPassword , Name , Notes
    str_SQL = "Update CodeUser Set user_Password='" & Trim(txt_Password.Text) & "'," & _
              "user_Name='" & Trim(txt_Name.Text) & "',user_Notes = '" & txt_Notes.Text & "',user_AddDate=Getdate(),user_AddWho='" & User_id & "' Where user_LoginID = '" & locUserID & "'"
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
End If
cn.CommitTrans
Tran_Level = 0

'�M�ù�
Call cmd_Clear_Click

Exit Sub

err_handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�ק�s��", Me.Caption, "cmd_Save_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "�ϥΪ̰򥻸�ƺ��@"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
Me.Height = 4740: Me.Width = 6200
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200

'���X�Ҧ����q���
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

'���X�Ҧ��s�ո��
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

'���X�ϥΪ̩��Ӹ��
If GetUserData(User_id) = False Then
   MsgBox funRtn_msg, vbOKOnly + vbInformation, msg_title
End If
locUserID = User_id

'�]�w�\��
If blAdmin Then
   '�t�κ޲z���i�H�s�W�B�R���A�ܧ�s�աA�v���]�w�A���A
   cmd_Add.Enabled = True           '�s�W
   cmd_Delete.Enabled = True        '�R��
   cmd_Query.Enabled = True         '�d��
   cmd_Clear.Enabled = True         '�M�ù�
   cmb_Company.Enabled = True       '�ϥΪ̥N���k�ݤ��q
   cmb_Group.Enabled = True         '�ϥΪ̥N���k�ݸs��
   chk_CloseCode.Enabled = True     '���A�G����
   chk_AdminCode.Enabled = True     '�ϥ��v���G�t�κ޲z��
Else
   '�� [�t�κ޲z��] ���~���Ҧ� User �u���\�ק� Password
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
'��s Menu [����]��[�w�}�����M��]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
Set frm_BaseData_UserData = Nothing
End Sub

Private Sub txt_UserID_KeyPress(KeyAscii As Integer)
'�ϥΪ̥N��
If KeyAscii = vbKeyReturn Then
   txt_Password.SelStart = 0
   txt_Password.SelLength = Len(txt_Password.Text)
   txt_Password.SetFocus
End If
End Sub

Private Function GetUserData(strUserID As String) As Boolean
'���X�ϥΪ̩��Ӹ��
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
   '��X User ���ݪ� Company
   For tmpI = 0 To UBound(arCompanyID) - 1
       If arCompanyID(tmpI) = tmp_Rs.Fields("CompanyID").Value Then
          cmb_Company.ListIndex = tmpI
       End If
   Next tmpI
   '��X User ���ݪ� Group
   For tmpI = 0 To UBound(argroupID) - 1
       If argroupID(tmpI) = tmp_Rs.Fields("GroupID").Value Then
          cmb_Group.ListIndex = tmpI
       End If
   Next tmpI
   
   txt_Name.Text = tmp_Rs.Fields("UserName").Value
   If tmp_Rs.Fields("UserStatus").Value = "1" Then  '�b���O�_�w�g����
      chk_CloseCode.Value = vbUnchecked
   Else
      chk_CloseCode.Value = vbChecked
   End If
   txt_Notes.Text = tmp_Rs.Fields("Notes").Value
   If tmp_Rs.Fields("GroupID").Value = "ADMIN" Then  '�t�κ޲z���v��
      chk_AdminCode.Value = vbChecked
      blAdmin = True
   Else
'      chk_AdminCode.Value = vbUnchecked
'      blAdmin = False
   End If
Else
   funRtn_msg = "�ϥΪ̸�ƿ��~�G�ϥΪ̥N�� [" & strUserID & "] ��Ƥ��s�b"
   Exit Function
End If
locUserID = strUserID
GetUserData = True
Exit Function

err_handle:
   funRtn_msg = "Function [Private GetUserData()] RunTime Error�G" & vbCrLf & _
                "Err Code�G" & err.Number & vbCrLf & "Err Descr�G" & err.Description
End Function

Private Function CheckData() As Boolean
'��Ƭd��
On Error GoTo err_handle
CheckData = False
If Len(RTrim(txt_UserID.Text)) = 0 Then
   funRtn_msg = "�s�W�@�~��ƿ��~�G�ϥΪ̥N�����ର�ť�"
   txt_UserID.SetFocus
   Exit Function
End If
If cmb_Company.ListIndex = -1 Then
   funRtn_msg = "�s�W�@�~��ƿ��~�G��������ϥΪ̤� [�k�ݤ��q] �@���v�����ި̾�"
   cmb_Company.SetFocus
   Exit Function
End If
If cmb_Group.ListIndex = -1 Then
   funRtn_msg = "�s�W�@�~��ƿ��~�G��������ϥΪ̤� [�k�ݸs��] �@���v�����ި̾�"
   cmb_Group.SetFocus
   Exit Function
End If
If Len(Trim(txt_Name.Text)) = 0 Then
   funRtn_msg = "�s�W�@�~��ƿ��~�G�ϥΪ� [�m�W] ���ର�ť�"
   txt_Name.SetFocus
   Exit Function
End If
CheckData = True
Exit Function

err_handle:
   funRtn_msg = "Function [Private CheckData()] RunTime Error�G" & vbCrLf & _
                "Err Code�G" & err.Number & vbCrLf & "Err Descr�G" & err.Description
End Function
