VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_BaseData_UserSecurity 
   Caption         =   " User  ��  ��  �v  ��  �]  �w"
   ClientHeight    =   6135
   ClientLeft      =   1020
   ClientTop       =   1665
   ClientWidth     =   9885
   Icon            =   "frm_BaseData_GroupSecurity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9885
   Begin VB.TextBox Text2 
      Appearance      =   0  '����
      BackColor       =   &H00FFFFC0&
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
      Left            =   1290
      TabIndex        =   14
      Top             =   750
      Width           =   1980
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '����
      BackColor       =   &H00FFFFC0&
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
      Left            =   1290
      TabIndex        =   13
      Top             =   1440
      Width           =   1980
   End
   Begin VB.TextBox txt_Name 
      Appearance      =   0  '����
      BackColor       =   &H00FFFFC0&
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
      Left            =   1290
      TabIndex        =   9
      Top             =   1095
      Width           =   1980
   End
   Begin VB.CheckBox chk_CloseCode 
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
      Left            =   4575
      TabIndex        =   8
      Top             =   720
      Width           =   960
   End
   Begin VB.CheckBox chk_AdminCode 
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
      Left            =   4575
      TabIndex        =   7
      Top             =   345
      Width           =   1725
   End
   Begin VB.CommandButton cmd_Query 
      Caption         =   "�d ��"
      Height          =   375
      Left            =   3420
      TabIndex        =   6
      Top             =   300
      Width           =   720
   End
   Begin VB.ComboBox cmb_User 
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
      Left            =   1290
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   4
      Top             =   315
      Width           =   1980
   End
   Begin VB.Frame Frame1 
      Height          =   4260
      Left            =   30
      TabIndex        =   2
      Top             =   1845
      Width           =   9825
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gd_ProdSec 
         Height          =   4095
         Left            =   30
         TabIndex        =   3
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
   Begin VB.CommandButton cmd_Save 
      BackColor       =   &H00FFC0C0&
      Caption         =   "�s  ��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   7350
      Picture         =   "frm_BaseData_GroupSecurity.frx":030A
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   1
      Top             =   420
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
      Height          =   915
      Left            =   8565
      Picture         =   "frm_BaseData_GroupSecurity.frx":0614
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   0
      Top             =   420
      Width           =   1035
   End
   Begin VB.Label lbl_UserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�m�@�@�W"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      TabIndex        =   12
      Top             =   795
      Width           =   1020
   End
   Begin VB.Label lbl_GroupName 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�k�ݸs��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      TabIndex        =   11
      Top             =   1485
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�k�ݤ��q"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      TabIndex        =   10
      Top             =   1140
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   840
   End
End
Attribute VB_Name = "frm_BaseData_UserSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private arUserID() As String




Private Sub cmb_Group_Click()
Call cmd_Query_Click
End Sub

Private Sub cmd_Exit_Click()
'���}
Unload Me
End Sub

Private Sub cmd_Query_Click()
'User �ϥ��v���]�w���
End Sub

Private Sub cmd_Save_Click()
'�s��


End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "User �ϥ��v���]�w"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
Me.Height = 6650: Me.Width = 10000
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200

'���X�Ҧ� UserID ���
Dim i As Integer
cmb_User.Clear: i = 0
ReDim arUserID(1) As String
str_SQL = "Select Rtrim(user_LoginID) as 'UserID',Rtrim(user_Name) as 'UserName' From ExceedAddin.dbo.CodeUSER Order by user_LoginID"
Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
If Not tmp_rs.EOF Then
   Do While Not tmp_rs.EOF
      cmb_User.AddItem tmp_rs.Fields("UserID").Value & vbTab & tmp_rs.Fields("UserName").Value
      i = i + 1
      If UBound(arUserID) < i Then
         ReDim Preserve arUserID(i) As String
      End If
      arUserID(i - 1) = tmp_rs.Fields("UserID").Value
      tmp_rs.MoveNext
   Loop
End If
cmb_User.ListIndex = -1
tmp_rs.Close


'�]�w Grid �榡
Call SetDBGrid
'���X�Ҧ��{�����
Dim tmpRec As Double
gd_ProdSec.Visible = False
gd_ProdSec.Rows = 2
gd_ProdSec.Row = 1
str_SQL = "Select Code as 'ProgID',Rtrim(Description1) as 'Descr' From ExceedAddin.dbo.CodeLKUP Where ListName = 'APMENU' Order by Code"
Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_rs.EOF Then
   tmpRec = 1
   Do While Not tmp_rs.EOF
      With gd_ProdSec
        .Row = tmpRec
        .Col = 0: .Text = tmpRec      '�Ǹ�
        .Col = 1: .Text = tmp_rs.Fields("ProgID").Value
        .Col = 2: .Text = tmp_rs.Fields("Descr").Value
        .Col = 3: .Text = ""
        .Col = 4: .Text = ""
        .Col = 5: .Text = ""
        .Col = 6: .Text = ""
        .Col = 7: .Text = ""
        .Col = 8: .Text = ""
        tmpRec = tmpRec + 1
        If tmpRec = .Rows Then .Rows = .Rows + 1
      End With
      tmp_rs.MoveNext
   Loop
   gd_ProdSec.Visible = True
End If
tmp_rs.Close


End Sub

Private Sub Form_Terminate()
'��s Menu [����]��[�w�}�����M��]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
Set frm_BaseData_GroupSecurity = Nothing
End Sub

Private Sub gd_ProdSec_Click()
'Program Data List
Dim SelectedCol As Integer, SelectedRow As Integer, i As Integer
With gd_ProdSec
     SelectedCol = .Col: SelectedRow = .Row
     .Col = 0    '�Ǹ�
     If Len(.Text) = 0 Then Exit Sub
     Select Case SelectedCol
            Case 2, 3, 4, 5, 6     '�s�ɡB�R���B�p��B��L�B����
                 .Col = SelectedCol
                 If Len(.Text) = 0 Then
                    .Text = "��"
                 Else
                    .Text = ""
                 End If
                 .Col = 0
            Case Else
                 Exit Sub
     End Select
End With
End Sub

Private Sub SetDBGrid()
'�W�١GSetGridFormat_OrderDetail
'���O�G�Ƶ{��
'�\��G�M���ó]�w [�ɡB���f�@�~] ��� [�w�s���Ӹ��] ��ܮ榡
'�ѼơG�ǤJ�ȡG�L
Dim sub_var1 As Integer, sub_var2 As Integer
gd_ProdSec.Visible = False
With gd_ProdSec
     .FixedRows = 1
     '�]�w���\��C���
     .AllowBigSelection = True
     '�]�w�C����r�r��
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "�s�ө���": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '�]�w�C�����e��
     .ColWidth(0) = 500
     .ColWidth(1) = 2000
     .ColWidth(2) = 3500
     .ColWidth(3) = 500
     .ColWidth(4) = 500
     .ColWidth(5) = 500
     .ColWidth(6) = 500
     .ColWidth(7) = 500
     .ColWidth(8) = 500
     
     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "�Ǹ�"
     .Col = 1: .Text = "ProgID"
     .Col = 2: .Text = "�{���W��"
     .Col = 3: .Text = "����"
     .Col = 4: .Text = "�s��"
     .Col = 5: .Text = "�R��"
     .Col = 6: .Text = "�d��"
     .Col = 7: .Text = "�ץX"
     .Col = 8: .Text = "�C�L"
     '�]�w�C����r���
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignLeftCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignCenterCenter
     .ColAlignment(4) = flexAlignCenterCenter
     .ColAlignment(5) = flexAlignCenterCenter
     .ColAlignment(6) = flexAlignCenterCenter
     .ColAlignment(7) = flexAlignCenterCenter
     .ColAlignment(8) = flexAlignCenterCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
End With
gd_ProdSec.Visible = True
End Sub
