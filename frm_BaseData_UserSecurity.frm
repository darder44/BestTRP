VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frm_BaseData_UserSecurity 
   Caption         =   " User  ��  ��  �v  ��  �]  �w"
   ClientHeight    =   6825
   ClientLeft      =   1020
   ClientTop       =   1665
   ClientWidth     =   9960
   Icon            =   "frm_BaseData_UserSecurity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9960
   Begin VB.CommandButton cmd_import 
      Caption         =   "�פJ�v���M��"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6240
      TabIndex        =   20
      Top             =   1125
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   0
      TabIndex        =   17
      Top             =   2160
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "�v���]�w"
      TabPicture(0)   =   "frm_BaseData_UserSecurity.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "�פJ�v�����t������"
      TabPicture(1)   =   "frm_BaseData_UserSecurity.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dg_import"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4260
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   9585
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gd_ProdSec 
            Height          =   4095
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   9390
            _ExtentX        =   16563
            _ExtentY        =   7223
            _Version        =   393216
            BackColor       =   -2147483624
            Rows            =   10
            Cols            =   9
            _NumberOfBands  =   1
            _Band(0).Cols   =   9
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_import 
         Height          =   4095
         Left            =   -74760
         TabIndex        =   21
         Top             =   480
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   7223
         _Version        =   393216
         BackColor       =   -2147483624
         Rows            =   10
         Cols            =   9
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
      End
   End
   Begin VB.CommandButton cmd2Excel 
      Caption         =   "�ץX�v���M��"
      Height          =   495
      Left            =   4800
      TabIndex        =   16
      Top             =   1125
      Width           =   1320
   End
   Begin VB.TextBox txt_UserName 
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
      TabIndex        =   10
      Top             =   750
      Width           =   1245
   End
   Begin VB.TextBox txt_GroupName 
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
      Top             =   1125
      Width           =   2985
   End
   Begin VB.TextBox txt_CompanyName 
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
      Left            =   2550
      TabIndex        =   6
      Top             =   750
      Width           =   4485
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
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   3240
      TabIndex        =   5
      Top             =   1695
      Width           =   960
   End
   Begin VB.CommandButton cmd_Query 
      Caption         =   "�d ��"
      Height          =   375
      Left            =   4485
      TabIndex        =   4
      Top             =   285
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
      TabIndex        =   2
      Top             =   315
      Width           =   3195
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
      Left            =   7365
      Picture         =   "frm_BaseData_UserSecurity.frx":0342
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   1
      Top             =   210
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
      Left            =   8655
      Picture         =   "frm_BaseData_UserSecurity.frx":064C
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   0
      Top             =   210
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   615
      Left            =   1290
      TabIndex        =   13
      Top             =   1440
      Width           =   3015
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
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1725
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   8400
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�b�����A"
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
      Index           =   1
      Left            =   180
      TabIndex        =   15
      Top             =   1620
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�ި�@�~"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   1
      Left            =   5685
      TabIndex        =   12
      Top             =   270
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�`�N�G�t�κ޲z�����ݳ]�w�����v�����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Index           =   1
      Left            =   4875
      TabIndex        =   11
      Top             =   1695
      Width           =   4590
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   1170
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
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   360
      Width           =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808000&
      BackStyle       =   1  '���z��
      Height          =   405
      Index           =   1
      Left            =   4785
      Top             =   1620
      Width           =   4770
   End
End
Attribute VB_Name = "frm_BaseData_UserSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arrTmp() As String
Private arUserID() As String
Private strLocUserID As String
Private locAdmin As Boolean
Private cn_Self As ADODB.Connection
Private rs_Diff As ADODB.Recordset

Private Sub cmb_Group_Click()
Call cmd_Query_Click
End Sub

Private Sub cmd_Exit_Click()
'���}
Unload Me
End Sub



Private Sub cmd_import_Click()
On Error GoTo err_Handle
Dim strFileName As String, strFieldName As String, str_TmpSQL As String
Dim i As Integer, j As Integer, k As Integer, x As Integer
Screen.MousePointer = 11: SSTab1.Tab = 1

'�פJ�v���M��
With dlgCommonDialog
    .DialogTitle = "TMS�v���פJ"
    .CancelError = True
    .InitDir = App.Path
    'ToDo: �]�w�q�ι�ܤ��������X�Ф��ݩ�
    .Filter = "*.xls|*.xls"
    .ShowOpen
    strFileName = .FileName
    
    If err.Number = cdlCancel Then strFileName = "": Exit Sub
    If Len(strFileName) = 0 Then Exit Sub
End With

'�M��
dg_import.Rows = 2
dg_import.Row = 1
For i = 0 To dg_import.Cols - 1
    dg_import.Col = i
    dg_import.Text = ""
Next
dg_import.Rows = 2
dg_import.Row = 1
For i = 0 To dg_import.Cols - 1
    dg_import.Col = i
    dg_import.Text = ""
Next

If Dir(strFileName) = "" Then MsgBox "�䤣���ɮסI", vbOKOnly + vbInformation, "TMS�v���פJ": Exit Sub '�䤣���ɮ�

'�}��EXCEL����
Set MyXlsApp = CreateObject("Excel.Application")

With MyXlsApp
    .Visible = False
    .Workbooks.Open (strFileName)
    .Sheets(1).Select '�������w�Ĥ@�Ӥu�@��
            
    '�����W��
    For i = 1 To 255
        If Len(RTrim(.Cells(1, i) & "")) = 0 Then Exit For
           strFieldName = strFieldName & myExCharFilter(RTrim(.Cells(1, i))) & Chr(9)
    Next i
    k = 2 '�ѲĤG�C�}�l�פJ
    
    arrTmp = Split(strFieldName, Chr(9))
    
    Dim rsTmp As New ADODB.Recordset
    
    If UBound(arrTmp) < 1 Then GoTo endsub
    '�إ�Recordset
    For i = 0 To UBound(arrTmp) - 1
        If Len(RTrim(arrTmp(i))) = 0 Then MsgBox "�� " & i & " ���W�� (" & arrTmp(i) & ") ���~�A�ɮ׸��J�פ�!", 64, "Excel2Recordset": GoTo endsub
        rsTmp.Fields.Append arrTmp(i), adChar, 255, adFldUpdatable
    Next i
    
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockOptimistic
    rsTmp.Open
    
    '�g�JRecordset
    Do While Len(RTrim(.Cells(k, 1))) > 0 'Or Len(RTrim(.Cells(k, 2))) > 0
    rsTmp.AddNew
        For j = 1 To UBound(arrTmp)
            rsTmp(j - 1) = RTrim(myExCharFilter(.Cells(k, j)))
        Next j
    rsTmp.Update
    k = k + 1
    Loop
    
    MyXlsApp.Quit: Set MyXlsApp = Nothing
    
endsub:
.DisplayAlerts = False: .Quit: Set MyXlsApp = Nothing
End With

Call DB_Connect_Self(cn_string) '�إ߷s�s�u
If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst

    cn_Self.Execute "if object_id ('##import') is not null drop table ##import", RowsAffect, adExecuteNoRecords
    
    str_TmpSQL = "CREATE TABLE ##import(�b�� varchar(100),���q varchar(100),�m�W varchar(100),�s�� varchar(100),���A varchar(100),ProgID varchar(100),�{�����| varchar(100),���� varchar(100))"
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly

    dg_import.Visible = False
    
    Do While Not rsTmp.EOF
        If Val(Trim(rsTmp("����").Value)) > 1 Or Len(Trim(Val(Trim(rsTmp("����").Value)))) > 1 Then MsgBox "�פJ�v����0,1�H�~��>_<" & Chr(13) & Chr(13) & "�b��:" & Trim(rsTmp("�b��").Value) & Chr(13) & "ProgID:" & Trim(rsTmp("ProgID").Value) & Chr(13) & "����:" & Trim(rsTmp("����").Value) & "", vbOKOnly, "�פJ����": GoTo final
        '�N�פJ��excel��� �s�J�Ȧs����ƪ�##import
        str_TmpSQL = "INSERT INTO ##import (�b��,���q,�m�W,�s��,���A,ProgID,�{�����|,����) " & _
                     "VALUES ('" & Trim(rsTmp("�b��").Value) & "','" & Trim(rsTmp("���q").Value) & "','" & Trim(rsTmp("�m�W").Value) & "','" & _
                     "" & Trim(rsTmp("�s��").Value) & "','" & Trim(rsTmp("���A").Value) & "','" & Trim(rsTmp("ProgID").Value) & "','" & Trim(rsTmp("�{�����|").Value) & "','" & _
                     "" & Trim(rsTmp("����").Value) & "')"

        Call Confirm_Recordset_Closed(tmp_Rs)
        tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
        rsTmp.MoveNext
    Loop

    '����t���v���t������
    str_TmpSQL = "select im.* " & _
                 "from CodeRole cr  right join ##import im on im.ProgID = cr.APCode and im.�b�� = cr.user_loginID " & _
                 "where cr.role_run <> im.���� "

    Call Confirm_Recordset_Closed(rs_Diff)
    rs_Diff.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly

If rs_Diff.BOF Then
    Call DB_Disconnect(cn_Self):
    rsTmp.Close: Set rsTmp = Nothing
    Screen.MousePointer = 0: Set MyXlsApp = Nothing
    MsgBox "�פJ���v���S���t������>_<", vbOKOnly, "�t�����"
    Exit Sub
End If

Dim intLine As Integer
intLine = 1
Do While Not rs_Diff.EOF
    dg_import.Col = 0: dg_import.Text = intLine
    dg_import.Col = 1: dg_import.Text = RTrim(rs_Diff("�b��"))
    dg_import.Col = 2: dg_import.Text = RTrim(rs_Diff("���q"))
    dg_import.Col = 3: dg_import.Text = RTrim(rs_Diff("�m�W"))
    dg_import.Col = 4: dg_import.Text = RTrim(rs_Diff("�s��"))
    dg_import.Col = 5: dg_import.Text = RTrim(rs_Diff("���A"))
    dg_import.Col = 6: dg_import.Text = RTrim(rs_Diff("ProgID"))
    dg_import.Col = 7: dg_import.Text = RTrim(rs_Diff("�{�����|"))
    dg_import.Col = 8: dg_import.Text = RTrim(rs_Diff("����"))
    rs_Diff.MoveNext
    intLine = intLine + 1
    dg_import.Rows = dg_import.Rows + 1
    dg_import.Row = dg_import.Row + 1
Loop
dg_import.Visible = True
dg_import.Rows = dg_import.Rows - 1

'�T�{�ק��v��,�i�������ݦ��X���t���ӹ���
    x = MsgBox("�нT�{�O�_��s�v��", vbQuestion + vbYesNo, "TMS�v���פJ�]�w") '�������U���O�T�w�άO����
    If x = 6 Then 'yes
        Tran_Level = cn.BeginTrans
        rs_Diff.MoveFirst
    
        Call Confirm_Recordset_Closed(tmp_Rs)
        Do While Not rs_Diff.EOF
            str_TmpSQL = "update coderole " & _
                         "Set role_run = '" & rs_Diff("����").Value & "' " & _
                         "Where user_loginid = '" & rs_Diff("�b��").Value & "' And apcode = '" & rs_Diff("ProgID").Value & "' "
    
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_TmpSQL, cn_Self, adOpenForwardOnly, adLockReadOnly
            rs_Diff.MoveNext
        Loop
        
        rs_Diff.Close: Set tmp_Rs = Nothing: Set rs_Diff = Nothing: dg_import.Clear
        cn.CommitTrans
        Tran_Level = 0
        MsgBox "��s�t���v��,����^_^", vbOKOnly, "��s���\"
    End If

final:
        rsTmp.Close: Set rsTmp = Nothing
        Screen.MousePointer = 0: cmd_Import.Enabled = True: Set MyXlsApp = Nothing
        Call setdbgrid1
        Call DB_Disconnect(cn_Self) '�����s�u
        
        '�M�ù�
        Call ResetDBGrid
        Call ClearForm_AllField(Me)

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
Dim str As String
If MyXlsApp Is Nothing = False Then MyXlsApp.Quit: Set MyXlsApp = Nothing

If err.Number = 3367 Then
    str = "���W��( " & arrTmp(i) & ")���ơI"
    
ElseIf err.Number = -2147217887 Then
    str = "�нT�{�x�s��(" & Chr(64 + j) & k & ")�A��ƬO�_���~�I"

Else
     str = "Exceed2Recordset"
End If

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmd_Query_Click()
'�d�ߨ���� User �ϥ��v���]�w���
Dim intUserID As Integer

On Error GoTo err_Handle
If cmb_User.ListIndex < 0 Then Exit Sub
strLocUserID = arUserID(cmb_User.ListIndex)
intUserID = cmb_User.ListIndex
Screen.MousePointer = 11: SSTab1.Tab = 0

Call ClearForm_AllField(Me)
Call ResetDBGrid
cmb_User.ListIndex = intUserID

'���^ User �򥻸��
str_SQL = "Select Rtrim(a.user_LoginID) as 'UserID' , Rtrim(a.user_Name) as 'UserName' , Rtrim(b.Description) as 'GroupName' , Rtrim(a.user_Group) as 'GroupID' , Rtrim(c.Description) as 'CompanyName' , Rtrim(a.user_Status) as 'UserStatus' " & _
        "From CodeUSER a " & _
        "Inner join CodeLKUP b on b.ListName = 'USERGROUP' and Rtrim(b.Code) = Rtrim(a.user_Group) " & _
        "Inner join CodeLKUP c on c.ListName = 'USERCOMPANY' and Rtrim(c.Code) = Rtrim(a.user_Company) " & _
        "Where a.user_LoginID = '" & arUserID(cmb_User.ListIndex) & "'"
Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
If Not tmp_Rs.EOF Then
   txt_UserName.Text = tmp_Rs.Fields("UserName").Value
   txt_CompanyName.Text = tmp_Rs.Fields("CompanyName").Value
   txt_GroupName.Text = tmp_Rs.Fields("GroupName").Value
   If tmp_Rs.Fields("GroupID").Value = "ADMIN" Then
      chk_AdminCode.Value = vbChecked
      locAdmin = True
   Else
      chk_AdminCode.Value = vbUnchecked
      locAdmin = False
   End If
   If tmp_Rs.Fields("UserStatus").Value = "1" Then
      chk_CloseCode.Value = vbUnchecked
   Else
      chk_CloseCode.Value = vbChecked
   End If
   strLocUserID = arUserID(cmb_User.ListIndex)
End If
tmp_Rs.Close

'���^�v���]�w���
Dim i As Integer
Dim j As Integer
Dim strAPCode As String
str_SQL = "Select * From CodeRole Where user_LoginID = '" & arUserID(cmb_User.ListIndex) & "'"
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
If tmp_Rs.EOF Then
   tmp_Rs.Close
   Screen.MousePointer = 0
   Exit Sub
End If

'����v���]�w���
   With gd_ProdSec
        For i = 1 To .Rows - 2
            .Row = i
            .Col = 1: strAPCode = .Text
            tmp_Rs.Filter = adFilterNone
            tmp_Rs.Filter = "APCode = '" & strAPCode & "'"
            If tmp_Rs.RecordCount <> 0 And (Not tmp_Rs.EOF) Then
               .Col = 3   '����
               If tmp_Rs.Fields("role_RUN").Value = "0" Then
                  .Text = ""
               Else
                  .Text = "V"
               End If
               .Col = 4   '�s��
               If tmp_Rs.Fields("role_SAVE").Value = "0" Then
                  .Text = ""
               Else
                  .Text = "V"
               End If
               .Col = 5   '�R��
               If tmp_Rs.Fields("role_DELETE").Value = "0" Then
                  .Text = ""
               Else
                  .Text = "V"
               End If
               .Col = 6   '�d��
               If tmp_Rs.Fields("role_QUERY").Value = "0" Then
                  .Text = ""
               Else
                  .Text = "V"
               End If
               .Col = 7   '�ץX
               If tmp_Rs.Fields("role_EXPORT").Value = "0" Then
                  .Text = ""
               Else
                  .Text = "V"
               End If
               .Col = 8   '�C�L
               If tmp_Rs.Fields("role_PRINT").Value = "0" Then
                  .Text = ""
               Else
                  .Text = "V"
               End If
            End If
        Next i
   End With
tmp_Rs.Filter = adFilterNone
tmp_Rs.Close
Screen.MousePointer = 0: cmd_Save.Enabled = True

Exit Sub

err_Handle:
    Screen.MousePointer = 0
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�d��", Me.Caption, "cmd_Query_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Save_Click()
'�s��
If locAdmin = True Then
   msg_text = "�`�N�G�t�κ޲z�����ݳ]�w�����v��"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Exit Sub
End If
If strLocUserID = "" Then Exit Sub

On Error GoTo err_Handle
'�R�����v���]�w���
Tran_Level = cn.BeginTrans
str_SQL = "Delete From CodeRole Where Rtrim(user_LoginID) = '" & Trim(arUserID(cmb_User.ListIndex)) & "'"
cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'�s�W��s�� [�v���]�w
Dim i As Integer
   With gd_ProdSec
        For i = 1 To .Rows - 2
            .Row = i: .Col = 1  'APCode
            str_SQL = "Insert into CodeRole " & _
                      "   (user_LoginID,APCode,role_RUN,role_SAVE,role_DELETE,role_QUERY,role_EXPORT,role_Print,AddWho) " & _
                      "Values ('" & arUserID(cmb_User.ListIndex) & "',"
            str_SQL = str_SQL & "'" & Trim(.Text) & "',"
            .Col = 3   '����
            If Len(Trim(.Text)) > 0 Then
               str_SQL = str_SQL & "'1',"
            Else
               str_SQL = str_SQL & "'0',"
            End If
            .Col = 4   '�s��
            If Len(Trim(.Text)) > 0 Then
               str_SQL = str_SQL & "'1',"
            Else
               str_SQL = str_SQL & "'0',"
            End If
            .Col = 5   '�R��
            If Len(Trim(.Text)) > 0 Then
               str_SQL = str_SQL & "'1',"
            Else
               str_SQL = str_SQL & "'0',"
            End If
            .Col = 6   '�d��
            If Len(Trim(.Text)) > 0 Then
               str_SQL = str_SQL & "'1',"
            Else
               str_SQL = str_SQL & "'0',"
            End If
            .Col = 7   '�ץX
            If Len(Trim(.Text)) > 0 Then
               str_SQL = str_SQL & "'1',"
            Else
               str_SQL = str_SQL & "'0',"
            End If
            .Col = 8   '�C�L
            If Len(Trim(.Text)) > 0 Then
               str_SQL = str_SQL & "'1',"
            Else
               str_SQL = str_SQL & "'0',"
            End If
            'userid
            str_SQL = str_SQL & "'" & arUserID(cmb_User.ListIndex) & "')"
            'Debug.Print str_SQL
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        Next i
   End With
cn.CommitTrans
Tran_Level = 0

'�M�ù�
Call ResetDBGrid
Call ClearForm_AllField(Me)

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then
      Tran_Level = 0
      cn.RollbackTrans
   End If
    Dim tmpString As String
    msg_text = "���~�T���G" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "-�s��", Me.Caption, "cmd_Save_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd2Excel_Click()

Screen.MousePointer = 11
Dim strUserID As String
If Len(RTrim(cmb_User)) > 0 Then strUserID = "and cr.user_loginid = '" & mySplit(cmb_User, " ", 0) & "' "

str_SQL = "Select �b�� = rtrim(cr.user_loginid) " & _
            ",���q = Rtrim(c.Description) " & _
            ",�m�W = Rtrim(cu.user_Name) " & _
            ",�s�� = Rtrim(cc.Description) " & _
            ",���A = case when cu.user_Status = 0 then '����' else '���`' end " & _
            ",Rtrim(cl.Code) as 'ProgID' " & _
            ",�{�����| = Rtrim(cl.Description) " & _
            ",���� = cr.role_run " & _
            "From CodeLKUP cl left join CodeRole cr on cl.Code =cr.apcode " & _
            "left join CodeUSER cu on cu.user_LoginID = cr.user_loginid " & _
            "left join CodeLKUP c on c.ListName = 'USERCOMPANY' and Rtrim(c.Code) = Rtrim(cu.user_Company) " & _
            "left join CodeLKUP cc on cc.ListName = 'USERGROUP' and Rtrim(cc.Code) = Rtrim(cu.user_Group) " & _
            "Where Rtrim(cl.Description) <> '' and cl.ListName = 'APMENU' " & strUserID & _
            "Order by cr.user_loginid, cl.Description "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic

'�t�sExcel
Call Recordset2Excel("�ϥΪ��v��", tmp_Rs)

Set MyXlsApp = Nothing

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
  '��s MDIForm �� Menu [����]��[�w��ܵ���] �O�_�ֿ�
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "User �ϥ��v���]�w"
End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
Me.Height = 7500: Me.Width = 10150
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200

If UCase(User_id) = "ADMINISTRATOR" Then cmd_Import.Enabled = True
'���X�Ҧ� UserID ���
Dim i As Integer
cmb_User.Clear: i = 0
ReDim arUserID(1) As String
str_SQL = "Select Rtrim(user_LoginID) as 'UserID',Rtrim(user_Name) as 'UserName' From CodeUSER Order by user_LoginID"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenStatic, adLockOptimistic
If Not tmp_Rs.EOF Then
   Do While Not tmp_Rs.EOF
      cmb_User.AddItem tmp_Rs.Fields("UserID").Value & Space(20 - Len(Trim(tmp_Rs.Fields("UserID").Value))) & tmp_Rs.Fields("UserName").Value
      i = i + 1
      If UBound(arUserID) < i Then
         ReDim Preserve arUserID(i) As String
      End If
      arUserID(i - 1) = tmp_Rs.Fields("UserID").Value
      tmp_Rs.MoveNext
   Loop
End If
cmb_User.ListIndex = -1
tmp_Rs.Close
strLocUserID = ""

'�]�w Grid �榡
Call SetDBGrid
'���X�Ҧ��{�����
Dim tmpRec As Double
gd_ProdSec.Visible = False
gd_ProdSec.Rows = 2
gd_ProdSec.Row = 1
str_SQL = "Select Rtrim(Code) as 'ProgID',Rtrim(Description) as 'Descr' From CodeLKUP Where ListName = 'APMENU' and Rtrim(Description) <> '' Order by Description"
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_Rs.EOF Then
   tmpRec = 1
   Do While Not tmp_Rs.EOF
      With gd_ProdSec
        .Row = tmpRec
        .Col = 0: .Text = tmpRec      '�Ǹ�
        .Col = 1: .Text = tmp_Rs.Fields("ProgID").Value
        .Col = 2: .Text = tmp_Rs.Fields("Descr").Value
        .Col = 3: .Text = ""
        .Col = 4: .Text = ""
        .Col = 5: .Text = ""
        .Col = 6: .Text = ""
        .Col = 7: .Text = ""
        .Col = 8: .Text = ""
        tmpRec = tmpRec + 1
        If tmpRec = .Rows Then .Rows = .Rows + 1
      End With
      tmp_Rs.MoveNext
   Loop
   gd_ProdSec.Visible = True
End If
tmp_Rs.Close
Call setdbgrid1

End Sub

Private Sub Form_Terminate()
'��s Menu [����]��[�w�}�����M��]
Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Form_Unload(Cancel As Integer)
'�q�O���餤�������A�Ǧ��ް_ [Terminate] �ƥ�
Set frm_BaseData_UserSecurity = Nothing
End Sub

Private Sub gd_ProdSec_Click()
'Program Data List
Dim SelectedCol As Integer, SelectedRow As Integer, i As Integer
With gd_ProdSec
     SelectedCol = .Col: SelectedRow = .Row
     .Col = 0    '�Ǹ�
     If Len(.Text) = 0 Then Exit Sub
     Select Case SelectedCol
            Case 3, 4, 5, 6, 7, 8   '����B�s�ɡB�R���B�d�ߡB�ץX�B�C�L
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
'�W�١GSetDBGrid
'���O�G�Ƶ{��
'�\��G�M���ó]�w [�l�t���v���]�w] ��� [�{���\��]�w��] ��ܮ榡
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
     .ColWidth(1) = 3200
     .ColWidth(2) = 2800
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

Private Sub setdbgrid1()    '�]�w�פJ�v����datagrd
Dim sub_var1 As Integer, sub_var2 As Integer
dg_import.Visible = False
With dg_import
     .FixedRows = 1
     '�]�w���\��C���
     .AllowBigSelection = True
     '�]�w�C����r�r��
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "�s�ө���": .CellFontSize = 10
         Next sub_var2
     Next sub_var1
     '�]�w�C�����e��
     .ColWidth(0) = 500
     .ColWidth(1) = 1000
     .ColWidth(2) = 2800
     .ColWidth(3) = 1000
     .ColWidth(4) = 1500
     .ColWidth(5) = 1000
     .ColWidth(6) = 3000
     .ColWidth(7) = 3500
     .ColWidth(8) = 500
     
     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "�Ǹ�"
     .Col = 1: .Text = "�b��"
     .Col = 2: .Text = "���q"
     .Col = 3: .Text = "�m�W"
     .Col = 4: .Text = "�s��"
     .Col = 5: .Text = "���A"
     .Col = 6: .Text = "ProgID"
     .Col = 7: .Text = "�{�����|"
     .Col = 8: .Text = "����"
     
     '�]�w�C����r���
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignLeftCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignCenterCenter
     .ColAlignment(4) = flexAlignCenterCenter
     .ColAlignment(5) = flexAlignCenterCenter
     .ColAlignment(6) = flexAlignLeftCenter
     .ColAlignment(7) = flexAlignLeftCenter
     .ColAlignment(8) = flexAlignCenterCenter

     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
End With
dg_import.Visible = True

End Sub

Private Sub ResetDBGrid()
'�W�١GResetDBGrid
'���O�G�Ƶ{��
'�\��G�M�� [�l�t���v���]�w] ��� [�{���\��]�w��] �]�w���
'�ѼơG�ǤJ�ȡG�L
Dim i As Integer
Dim j As Integer
With gd_ProdSec
     For i = 1 To .Rows - 1
         .Row = i
         For j = 3 To .Cols - 1
             .Col = j
             .Text = ""
         Next j
     Next i
End With
End Sub


Private Sub DB_Connect_Self(connection_string As String)
'ADO [Connection] Object connect
On Error GoTo err_Handle
Set cn_Self = New ADODB.Connection
cn_Self.CommandTimeout = 300
cn_Self.ConnectionTimeout = 20
cn_Self.ConnectionString = connection_string
cn_Self.Open Options:=adAsyncConnect
Do While cn_Self.State = adStateConnecting
   DoEvents: DoEvents
Loop
Exit Sub

err_Handle:
   msg_text = "�s�u���~�G�L�k�P��Ʈw�إ߳s�u�A�гq�� ��T�� "
   MsgBox msg_text, vbOKOnly + vbInformation, ""
   End
End Sub
