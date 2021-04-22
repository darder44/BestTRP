VERSION 5.00
Begin VB.Form frm_SDNCarNOFix 
   BorderStyle     =   1  '單線固定
   Caption         =   "車號變更"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   3015
   StartUpPosition =   2  '螢幕中央
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   2775
      Begin VB.ComboBox cboType 
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
         ItemData        =   "frm_SDNCarNOFix.frx":0000
         Left            =   1200
         List            =   "frm_SDNCarNOFix.frx":0002
         Style           =   2  '單純下拉式
         TabIndex        =   15
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox txt_Driver 
         BeginProperty Font 
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Style           =   2  '單純下拉式
         TabIndex        =   4
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消"
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   3600
         Width           =   975
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "確認儲存"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "一次車號"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Caption         =   "駕駛姓名"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Caption         =   "運輸公司"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Caption         =   "到貨日期"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Caption         =   "出車日期"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Caption         =   "二次路編"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Caption         =   "車號確認"
         BeginProperty Font 
            Name            =   "新細明體"
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
    MsgBox "車輛基本資料中查無此車號！", vbOKOnly, Me.Caption
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
If Len(RTrim(cbo_VehicleID.Text)) = 0 Then MsgBox "請選取二次車號!", 64, Me.Caption: Exit Sub
If Len(RTrim(cboType.Text)) = 0 Then MsgBox "請選取一次車號!", 64, Me.Caption: Exit Sub

Dim blCarNoCheck As Boolean
blCarNoCheck = False
'Terry 20190402 車號不變可儲存變更 (可隨車籍資料變更請款人)
If cbo_VehicleID.Text <> strOriginalCarNo Then
    blCarNoCheck = True
End If

If blCarNoCheck Then
    'Terry 20190320 新增防呆 已維護棧板之路編不可修改車號
    Dim rsTmp As New ADODB.Recordset
    Call ReDim_Recordset(rsTmp)
    str_SQL = "select count(*) from pallet_cds where checkno = '" & txt_C_Route_NO.Text & "'"
    rsTmp.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If rsTmp.Fields(0).Value > 0 Then
        rsTmp.Close
        MsgBox ("此路編已維護棧板，無法變更車號!")
        Exit Sub
    End If
    rsTmp.Close
    
    'Terry 20190327 新增防呆 已有計費資料之路編不可修改車號
    Call ReDim_Recordset(tmp_Rs)
    str_SQL = "select count(*) from sdn05t where c_route_no = '" & txt_C_Route_NO.Text & "'"
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.Fields(0).Value > 0 Then
        tmp_Rs.Close
        MsgBox ("此路編已有計費資料，無法變更車號!")
        Exit Sub
    End If
    tmp_Rs.Close
End If

If MsgBox("車號變更是針對這個二次路編的所有訂單，確認修改?", vbYesNo, "車號修改") <> vbYes Then Exit Sub

Screen.MousePointer = 11
cn.BeginTrans

'更新orders
strSql = "update sdn01t set c_vehicle_id_no = '" & cbo_VehicleID.Text & "', driver = '" & Trim(txt_Driver.Text) & _
            "' ,editdate = getdate(),edituser = '" & User_id & "' where c_route_no = '" & txt_C_Route_NO.Text & "' "
cn.Execute strSql, RowsAffect, adExecuteNoRecords

'配送類別
If cboType = "直送" Then
    cn.Execute "update sdn02t set vehicle_id_no = '" & cbo_VehicleID & "' where c_route_no = '" & txt_C_Route_NO & "' ", RowsAffect, adExecuteNoRecords
ElseIf cboType = "中區轉運" Then
    cn.Execute "update sdn02t set vehicle_id_no = '000-31' where c_route_no = '" & txt_C_Route_NO & "' ", RowsAffect, adExecuteNoRecords
ElseIf cboType = "南區轉運" Then
    cn.Execute "update sdn02t set vehicle_id_no = '002-34' where c_route_no = '" & txt_C_Route_NO & "' ", RowsAffect, adExecuteNoRecords

Else '車號User自訂
    cn.Execute "update sdn02t set vehicle_id_no = '" & cboType & "' where c_route_no = '" & txt_C_Route_NO & "' ", RowsAffect, adExecuteNoRecords

End If

'更新請款人
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

cboType.AddItem "直送"
cboType.AddItem "中區轉運"
cboType.AddItem "南區轉運"
'cboType.AddItem "外島"

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
    
    'Terry 20190402 比對車號是否變更
    strOriginalCarNo = cbo_VehicleID.Text
Else
    txt_C_Route_NO.Text = frm_OP_SDNAbnormal.txt_C_Route_NO.Text
    cbo_VehicleID.Text = mySplit(frm_OP_SDNAbnormal.txt_OneOrder_VehicleID.Text, "_", 0) & ""
    txt_Driver.Text = frm_OP_SDNAbnormal.txt_OneOrder_Driver.Text
    txt_DeliveryDate.Text = frm_OP_SDNAbnormal.txt_OneOrder_DeliveryDate.Text
    txt_ArriveDate.Text = RTrim(frm_OP_SDNAbnormal.txt_OneOrder_ArriveDate.Text)
    txt_TRPCompany.Text = RTrim(frm_OP_SDNAbnormal.txt_OneOrder_TRPCompany.Text)
    
    'Terry 20190402 比對車號是否變更
    strOriginalCarNo = cbo_VehicleID.Text
End If

Screen.MousePointer = 0
End Sub
