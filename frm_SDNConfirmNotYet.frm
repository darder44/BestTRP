VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frm_SDNConfirmNotYet 
   BorderStyle     =   1  '��u�T�w
   Caption         =   "�ֳtñ��T�{"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   13995
   StartUpPosition =   2  '�ù�����
   Begin VB.ComboBox cboCarNo 
      BeginProperty Font 
         Name            =   "�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3240
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.ComboBox cboStorerkey 
      BeginProperty Font 
         Name            =   "�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   720
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   4
      Top             =   120
      Width           =   2445
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "����"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdSdnConfirmExpress 
      Caption         =   "���`ñ��T�{"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      ToolTipText     =   "���@���`ñ�椣�p�B�O"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmd2Excel 
      Caption         =   "��Excel"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dgMain 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   20
      TabAction       =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1028
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      AutoSize        =   -1  'True
      BackStyle       =   0  '�z��
      Caption         =   "�z��"
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
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frm_SDNConfirmNotYet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As New ADODB.Recordset
Private intColumnIndex As Integer

Private Sub cboStorerkey_Click()
Call Filter
End Sub
Private Sub cboCarNo_Click()
Call Filter
End Sub

Private Sub Filter()

If cboStorerKey <> "" And cboCarno <> "" Then
    rsMain.Filter = "(�f�D = '" & mySplit(cboStorerKey, "_", 0) & "' and ���P���X = '" & mySplit(cboCarno, "_", 0) & "')"
ElseIf cboStorerKey = "" And cboCarno <> "" Then
    rsMain.Filter = "(���P���X = '" & mySplit(cboCarno, "_", 0) & "')"
ElseIf cboStorerKey <> "" And cboCarno = "" Then
    rsMain.Filter = "(�f�D = '" & mySplit(cboStorerKey, "_", 0) & "')"
Else
    rsMain.Filter = ""
    rsMain.Sort = "�s��"
End If

End Sub

Private Sub cmd2Excel_Click()

'��ƱƧ�
Recordset2Excel Me.Caption, rsMain

'..�b���s��EXCEL
If rsMain Is Nothing Then
Else
    With MyXlsApp

                
    End With
End If
Set MyXlsApp = Nothing

End Sub

Private Sub cmdSdnConfirmExpress_Click()

rsMain.Filter = "�ֳt�T�{ = 'V'"

If rsMain.RecordCount = 0 Then Call Form_Load: Exit Sub

rsMain.MoveFirst

Do While Not rsMain.EOF
    
    If RTrim(rsMain("�ֳt�T�{")) = "V" Then
    
    cn.Execute "select receipt_no from sdn02t Where len(rtrim(isnull(Confirm_Notes,''))) > 0 and receipt_no = '" & rsMain("TMS�渹") & "' ", RowsAffect, adExecuteNoRecords
    If RowsAffect <> 0 Then GoTo NextRow
    
    '��s SDN01T
    cn.Execute "Update SDN01T Set sdn_Date = getdate() Where c_route_no = '" & rsMain("�G�����s") & "'", RowsAffect, adExecuteNoRecords
    
    '��s SDN02T
    cn.Execute "Update SDN02T Set Confirm_UserID = '" & User_id & "',Confirm_Date = getdate(),Confirm_Notes = '���`�q��' ,SDNSendDate = getdate() , SDN_NOTE = '�ֳtñ��T�{' ,sdnback = '1', custsigndate = isnull(CustSignDate,isnull(SCHEDULEDATE,Arrive_Date)) Where Receipt_No = '" & rsMain("TMS�渹") & "'", RowsAffect, adExecuteNoRecords

    End If
    
NextRow:
rsMain.MoveNext
Loop

Call Form_Load

End Sub

Private Sub cmdSelectAll_Click()

rsMain.MoveFirst
Do While Not rsMain.EOF
    rsMain("�ֳt�T�{") = "V"
rsMain.MoveNext
Loop

End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

    If Len(dgMain.Columns(ColIndex).DataField) = 0 Then Exit Sub
    SaveSetting App.title, Me.Caption & "dgMain", dgMain.Columns(ColIndex).DataField, dgMain.Columns(ColIndex).Width
    
End Sub

Private Sub dgMain_DblClick()
frm_OP_SDNConfirm.txt_OrderKey.Text = rsMain("TMS�渹"): frm_OP_SDNConfirm.cmbOrderkey.Text = "TMS�渹"
Call frm_OP_SDNConfirm.cmd_OrderQuery_Click
Unload Me
End Sub

Private Sub dgMain_HeadClick(ByVal ColIndex As Integer)
If dgMain.Row = -1 Then Exit Sub
If intColumnIndex = ColIndex Then
    rsMain.Sort = dgMain.Columns(ColIndex).Caption & " DESC"
    dgMain.ClearSelCols
    intColumnIndex = 255

Else
    rsMain.Sort = dgMain.Columns(ColIndex).Caption
    dgMain.ClearSelCols
    intColumnIndex = ColIndex

End If
End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

With dgMain
    
    '�����\���ܯS�w���
    If .Col <> 6 Then Exit Sub
    If dgMain = " " Then
        dgMain = "V"
    Else
        dgMain = " "
    End If
    .Col = 5

End With
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11

'str_SQL = "Select �X����� , �G�����s ,���u�s�� " & _
'            ",�f�D " & _
'            ",�f�D�渹 " & _
'            ",�ֳt�T�{ = ' ' " & _
'            ",TMS�渹= �q��s�� " & _
'            ",�q�����O " & _
'            ",���P���X " & _
'            ",�r�p�H " & _
'            ",�f�B���q " & _
'            ",�Ƶ� = ���� " & _
'            ",�Ȥ�s�� " & _
'            ",�Ȥ�W�� " & _
'            ",�e�f�a�} " & _
'            ",�q���� " & _
'            ",��f��� " & _
'            "From SDNConfirm_OrderDate_One " & _
'            "where ñ��w�^ = 0 order by �X����� ,�G�����s ,���u�s�� ,�q��s�� "


str_SQL = "select �X����� = convert(varchar,t01t.Delivery_Date,112),�G�����s = t02t.c_Route_No,���u�s�� = t02t.Route_No " & _
",�f�D = Rtrim(t02t.StorerKey),�q�渹�X = Rtrim(t02t.Extern) ,�ֳt�T�{ = ' ',TMS�渹= rtrim(t02t.Receipt_No) " & _
",�q�����O = rtrim(t02t.priority),���P���X = Rtrim(t01t.c_Vehicle_ID_No),�r�p�H = Rtrim(t01t.driver) " & _
",�f�B���q = Isnull(Rtrim(t8m.Short_Name),''),�Ƶ� = Rtrim(Isnull(t02t.Description,'')) " & _
",�Ȥ�s�� = Rtrim(t02t.ConsigneeKey),�Ȥ�W�� = Rtrim(Isnull(t1m.Short_Name,'')) " & _
",�e�f�a�} = Rtrim(Isnull(t1m.Address,'')),�q���� = rtrim(t02t.Receipt_Date) " & _
",��f��� = rtrim(t02t.Arrive_Date) " & _
"From SDN02T t02t (nolock) join SDN01T t01t (nolock) on t02t.c_route_no = t01t.c_route_no " & _
"join trp01m t1m (nolock) on t02t.consigneekey = t1m.consigneekey and t02t.storerkey = t1m.storerkey " & _
"left join trp09m t9m (nolock) on t9m.vehicle_id_no = t01t.c_vehicle_id_no " & _
"left join trp08m t8m (nolock) on t8m.company_code = t9m.trp_company_code " & _
"where t02t.sdnback = 0 order by convert(varchar,t01t.Delivery_Date,112) ,t02t.c_Route_No ,t02t.Route_No ,t02t.Receipt_No "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

Call Replication_Recordset(tmp_Rs, rsMain)
tmp_Rs.Close

If Not rsMain.EOF Then rsMain.MoveFirst

'�f�D
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select distinct(storerkey) , short_name from trp16M order by storerkey ", cn, adOpenKeyset, adLockPessimistic

cboStorerKey.Clear
cboStorerKey.AddItem ""
Do While Not tmp_Rs.EOF
    cboStorerKey.AddItem RTrim(tmp_Rs("storerkey")) & "_" & RTrim(tmp_Rs("short_name"))
tmp_Rs.MoveNext
Loop
tmp_Rs.Close

'����
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select vehicle_id_no , driver from trp09M order by vehicle_id_no ", cn, adOpenKeyset, adLockPessimistic

cboCarno.Clear
cboCarno.AddItem ""
Do While Not tmp_Rs.EOF
    cboCarno.AddItem RTrim(tmp_Rs("vehicle_id_no")) & "_" & RTrim(tmp_Rs("driver"))
tmp_Rs.MoveNext
Loop
tmp_Rs.Close

Set dgMain.DataSource = rsMain

'�����e��
SetDataGridColWidth Me.Caption, dgMain

Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
End Sub
