VERSION 5.00
Begin VB.Form frm_UserLogin 
   BackColor       =   &H00C0C0C0&
   Caption         =   " �t �� �n �J �{ �� "
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
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  '�Ȥ�
         Left            =   4185
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   420
         Width           =   1560
      End
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
         IMEMode         =   3  '�Ȥ�
         Left            =   1485
         TabIndex        =   0
         Top             =   420
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�K   �X"
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
         TabIndex        =   6
         Top             =   465
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�ϥΪ̥N��"
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
         TabIndex        =   5
         Top             =   480
         Width           =   1290
      End
   End
   Begin VB.CommandButton cmd_Login 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�n�J�t��"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   2
      Top             =   2655
      Width           =   1755
   End
   Begin VB.CommandButton cmd_Cancel 
      BackColor       =   &H00FFFFC0&
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Style           =   1  '�Ϥ��~�[
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
         BackStyle       =   0  '�z��
         Caption         =   "�k�ݤ��q�G"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   165
         TabIndex        =   9
         Top             =   1035
         Width           =   1275
      End
      Begin VB.Label lbl_GroupName 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�k�ݸs�աG"
         BeginProperty Font 
            Name            =   "�s�ө���"
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
      BackStyle       =   1  '���z��
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
      BackStyle       =   1  '���z��
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
'����
Call DB_Disconnect(cn)
End
End Sub

Private Sub cmd_Login_Click()
'�n�J�t��
On Error GoTo err_Handle

txt_UserID.Text = Trim(txt_UserID.Text)
If Len(txt_UserID.Text) = 0 Then Exit Sub

Screen.MousePointer = vbHourglass

'�ˬd�K�X�O�_���T
If locstrPassword <> Trim(txt_Password.Text) Then
   msg_text = "�K�X���~"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   txt_Password.SelStart = 0: txt_Password.SelLength = Len(txt_Password.Text): txt_Password.SetFocus
   Exit Sub
End If

'�]�w �\���
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
      msg_text = "�v���]�w���~�GUser ��������i���檺�v���]�w"
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
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�n�J�t��", Me.Caption, "cmd_Login_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
  msg_title = "User Login �@�~"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
Me.Height = 4170 + 100: Me.Width = 6300
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm_UserLogin = Nothing
End Sub


Private Sub txt_Password_KeyPress(KeyAscii As Integer)
'�K�X
If KeyAscii = vbKeyReturn Then Call cmd_Login_Click
'   KeyAscii = 0
'   cmd_Login.SetFocus
'End If
End Sub

Private Sub txt_UserID_KeyPress(KeyAscii As Integer)
'�ϥΪ̳渹
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
'���o�ϥΪ̬������
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
   '��ܨϥΪ̸��
   lbl_CompanyName.Caption = "�k�ݤ��q�G" & tmp_Rs.Fields("CompanyName").Value
   lbl_GroupName.Caption = "�k�ݸs�աG" & tmp_Rs.Fields("GroupName").Value
   lbl_UserName.Caption = "�m�@�@�W�G" & tmp_Rs.Fields("UserName").Value
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
   '���ε��O
   If tmp_Rs.Fields("CloseCode").Value = "0" Then
      msg_text = "�n�J���ѡG�ϥΪ̱b���w���ΡI"
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
      txt_UserID.SelStart = 0: txt_UserID.SelLength = Len(txt_UserID.Text)
      txt_UserID.SetFocus
      Exit Sub
   End If
Else
   msg_text = "�n�J���ѡG�ϥΪ̱b�����s�b�I"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   txt_UserID.SelStart = 0: txt_UserID.SelLength = Len(txt_UserID.Text)
   txt_UserID.SetFocus
   Exit Sub
End If
tmp_Rs.Close
End Sub
