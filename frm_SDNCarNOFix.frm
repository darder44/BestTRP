VERSION 5.00
Begin VB.Form frm_SDNCarNOFix 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�����ܧ�"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   3015
   StartUpPosition =   2  '�ù�����
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   2775
      Begin VB.ComboBox cboType 
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
         ItemData        =   "frm_SDNCarNOFix.frx":0000
         Left            =   1200
         List            =   "frm_SDNCarNOFix.frx":0002
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   15
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txt_Driver 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txt_TRPCompany 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txt_ArriveDate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txt_DeliveryDate 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   405
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txt_C_Route_NO 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   405
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cbo_VehicleID 
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
         Left            =   1200
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   4
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "����"
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "�T�{�x�s"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "�@������"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "�r�p�m�W"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "�B�餽�q"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "��f���"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "�X�����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "�G�����s"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "�����T�{"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   975
      End
   End
End
Attribute VB_Name = "frm_SDNCarNOFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As New ADODB.Recordset
Dim strSql As String
Dim strOriginalCarNo As String


Private Sub cbo_VehicleID_Click()
If Len(RTrim(cbo_VehicleID.Text)) = 0 Then Exit Sub

Dim rsTmp As New ADODB.Recordset
rsTmp.Open "select driver = rtrim(isnull(t9.driver,'')) ,trpcompany = rtrim(isnull(t8.short_name,'')) from trp09m t9 left join trp08m t8 on t9.trp_company_code = t8.company_code where t9.vehicle_id_no = '" & cbo_VehicleID.Text & "' ", cn

If rsTmp.EOF Then
    MsgBox "�����򥻸�Ƥ��d�L�������I", vbOKOnly, Me.Caption
    cbo_VehicleID.ListIndex = -1
    txt_Driver = ""
    txt_TRPCompany = ""
    Exit Sub
End If

txt_Driver = rsTmp("driver")
txt_TRPCompany = rsTmp("trpcompany")

rsTmp.Close: Set rsTmp = Nothing
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If Len(RTrim(cbo_VehicleID.Text)) = 0 Then MsgBox "�п���G������!", 64, Me.Caption: Exit Sub
If Len(RTrim(cboType.Text)) = 0 Then MsgBox "�п���@������!", 64, Me.Caption: Exit Sub

Dim blCarNoCheck As Boolean
blCarNoCheck = False
'Terry 20190402 �������ܥi�x�s�ܧ� (�i�H���y����ܧ�дڤH)
If cbo_VehicleID.Text <> strOriginalCarNo Then
    blCarNoCheck = True
End If

If blCarNoCheck Then
    'Terry 20190320 �s�W���b �w���@�̪O�����s���i�ק郞��
    Dim rsTmp As New ADODB.Recordset
    Call ReDim_Recordset(rsTmp)
    str_SQL = "select count(*) from pallet_cds where checkno = '" & txt_C_Route_NO.Text & "'"
    rsTmp.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If rsTmp.Fields(0).Value > 0 Then
        rsTmp.Close
        MsgBox ("�����s�w���@�̪O�A�L�k�ܧ󨮸�!")
        Exit Sub
    End If
    rsTmp.Close
    
    'Terry 20190327 �s�W���b �w���p�O��Ƥ����s���i�ק郞��
    Call ReDim_Recordset(tmp_Rs)
    str_SQL = "select count(*) from sdn05t where c_route_no = '" & txt_C_Route_NO.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields(0).Value > 0 Then
        tmp_Rs.Close
        MsgBox ("�����s�w���p�O��ơA�L�k�ܧ󨮸�!")
        Exit Sub
    End If
    tmp_Rs.Close
End If

If MsgBox("�����ܧ�O�w��o�ӤG�����s���Ҧ��q��A�T�{�ק�?", vbYesNo, "�����ק�") <> vbYes Then Exit Sub

Screen.MousePointer = 11
cn.BeginTrans

'��sorders
strSql = "update sdn01t set c_vehicle_id_no = '" & cbo_VehicleID.Text & "', driver = '" & Trim(txt_Driver.Text) & _
            "' ,editdate = getdate(),edituser = '" & User_id & "' where c_route_no = '" & txt_C_Route_NO.Text & "' "
cn.Execute strSql, RowsAffect, adExecuteNoRecords

'�t�e���O
If cboType = "���e" Then
    cn.Execute "update sdn02t set vehicle_id_no = '" & cbo_VehicleID & "' where c_route_no = '" & txt_C_Route_NO & "' ", RowsAffect, adExecuteNoRecords
ElseIf cboType = "������B" Then
    cn.Execute "update sdn02t set vehicle_id_no = '000-31' where c_route_no = '" & txt_C_Route_NO & "' ", RowsAffect, adExecuteNoRecords
ElseIf cboType = "�n����B" Then
    cn.Execute "update sdn02t set vehicle_id_no = '002-34' where c_route_no = '" & txt_C_Route_NO & "' ", RowsAffect, adExecuteNoRecords

Else '����User�ۭq
    cn.Execute "update sdn02t set vehicle_id_no = '" & cboType & "' where c_route_no = '" & txt_C_Route_NO & "' ", RowsAffect, adExecuteNoRecords

End If

'��s�дڤH
cn.Execute "update SDN01T set receiver = (select isnull(receiver,driver) from trp09m where vehicle_id_no = '" & cbo_VehicleID.Text & "') where c_route_no = '" & txt_C_Route_NO.Text & "'", RowsAffect, adExecuteNoRecords

cn.CommitTrans

If intSDNCarChange = 0 Then
    Call frm_OP_SDNConfirm.cmd_OrderQuery_Click
Else
    Call frm_OP_SDNAbnormal.cmd_OrderQuery_Click
End If

'frm_OP_SDNConfirm.txt_OneOrder_VehicleID.Text = cbo_VehicleID.Text
'frm_OP_SDNConfirm.txt_OneOrder_Driver.Text = txt_Driver.Text
'frm_OP_SDNConfirm.txt_OneOrder_TRPCompany.Text = txt_TRPCompany.Text

Screen.MousePointer = 0
Call cmdCancel_Click

End Sub

Private Sub Form_Load()
Screen.MousePointer = 11

Dim rsTmp As New ADODB.Recordset
rsTmp.Open "select Carno = rtrim(vehicle_id_no) from trp09m order by vehicle_id_no", cn
rsTmp.MoveFirst

cboType.AddItem "���e"
cboType.AddItem "������B"
cboType.AddItem "�n����B"
'cboType.AddItem "�~�q"

Do While Not rsTmp.EOF

    cbo_VehicleID.AddItem rsTmp("carno")
    cboType.AddItem rsTmp("carno")
    
    rsTmp.MoveNext

Loop
rsTmp.Close: Set rsTmp = Nothing

If intSDNCarChange = 0 Then
    txt_C_Route_NO.Text = frm_OP_SDNConfirm.txt_C_Route_NO.Text
    cbo_VehicleID.Text = mySplit(frm_OP_SDNConfirm.txt_OneOrder_VehicleID.Text, "_", 0) & ""
    txt_Driver.Text = frm_OP_SDNConfirm.txt_OneOrder_Driver.Text
    txt_DeliveryDate.Text = frm_OP_SDNConfirm.txt_OneOrder_DeliveryDate.Text
    txt_ArriveDate.Text = RTrim(frm_OP_SDNConfirm.txt_OneOrder_ArriveDate.Text)
    txt_TRPCompany.Text = RTrim(frm_OP_SDNConfirm.txt_OneOrder_TRPCompany.Text)
    
    'Terry 20190402 ��郞���O�_�ܧ�
    strOriginalCarNo = cbo_VehicleID.Text
Else
    txt_C_Route_NO.Text = frm_OP_SDNAbnormal.txt_C_Route_NO.Text
    cbo_VehicleID.Text = mySplit(frm_OP_SDNAbnormal.txt_OneOrder_VehicleID.Text, "_", 0) & ""
    txt_Driver.Text = frm_OP_SDNAbnormal.txt_OneOrder_Driver.Text
    txt_DeliveryDate.Text = frm_OP_SDNAbnormal.txt_OneOrder_DeliveryDate.Text
    txt_ArriveDate.Text = RTrim(frm_OP_SDNAbnormal.txt_OneOrder_ArriveDate.Text)
    txt_TRPCompany.Text = RTrim(frm_OP_SDNAbnormal.txt_OneOrder_TRPCompany.Text)
    
    'Terry 20190402 ��郞���O�_�ܧ�
    strOriginalCarNo = cbo_VehicleID.Text
End If

Screen.MousePointer = 0
End Sub
