VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frm_SDNConfirmNotYet 
   BorderStyle     =   1  '單線固定
   Caption         =   "快速簽單確認"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   13995
   StartUpPosition =   2  '螢幕中央
   Begin VB.ComboBox cboCarNo 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3240
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.ComboBox cboStorerkey 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   720
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   120
      Width           =   2445
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "全選"
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdSdnConfirmExpress 
      Caption         =   "正常簽單確認"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      ToolTipText     =   "維護正常簽單不計運費"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmd2Excel 
      Caption         =   "轉Excel"
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
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
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
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "篩選"
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
    rsMain.Filter = "(貨主 = '" & mySplit(cboStorerKey, "_", 0) & "' and 車牌號碼 = '" & mySplit(cboCarno, "_", 0) & "')"
ElseIf cboStorerKey = "" And cboCarno <> "" Then
    rsMain.Filter = "(車牌號碼 = '" & mySplit(cboCarno, "_", 0) & "')"
ElseIf cboStorerKey <> "" And cboCarno = "" Then
    rsMain.Filter = "(貨主 = '" & mySplit(cboStorerKey, "_", 0) & "')"
Else
    rsMain.Filter = ""
    rsMain.Sort = "編號"
End If

End Sub

Private Sub cmd2Excel_Click()

'資料排序
Recordset2Excel Me.Caption, rsMain

'..在此編輯EXCEL
If rsMain Is Nothing Then
Else
    With MyXlsApp

                
    End With
End If
Set MyXlsApp = Nothing

End Sub

Private Sub cmdSdnConfirmExpress_Click()

rsMain.Filter = "快速確認 = 'V'"

If rsMain.RecordCount = 0 Then Call Form_Load: Exit Sub

rsMain.MoveFirst

Do While Not rsMain.EOF
    
    If RTrim(rsMain("快速確認")) = "V" Then
    
    cn.Execute "select receipt_no from sdn02t Where len(rtrim(isnull(Confirm_Notes,''))) > 0 and receipt_no = '" & rsMain("TMS單號") & "' ", RowsAffect, adExecuteNoRecords
    If RowsAffect <> 0 Then GoTo NextRow
    
    '更新 SDN01T
    cn.Execute "Update SDN01T Set sdn_Date = getdate() Where c_route_no = '" & rsMain("二次路編") & "'", RowsAffect, adExecuteNoRecords
    
    '更新 SDN02T
    cn.Execute "Update SDN02T Set Confirm_UserID = '" & User_id & "',Confirm_Date = getdate(),Confirm_Notes = '正常訂單' ,SDNSendDate = getdate() , SDN_NOTE = '快速簽單確認' ,sdnback = '1', custsigndate = isnull(CustSignDate,isnull(SCHEDULEDATE,Arrive_Date)) Where Receipt_No = '" & rsMain("TMS單號") & "'", RowsAffect, adExecuteNoRecords

    End If
    
NextRow:
rsMain.MoveNext
Loop

Call Form_Load

End Sub

Private Sub cmdSelectAll_Click()

rsMain.MoveFirst
Do While Not rsMain.EOF
    rsMain("快速確認") = "V"
rsMain.MoveNext
Loop

End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

    If Len(dgMain.Columns(ColIndex).DataField) = 0 Then Exit Sub
    SaveSetting App.title, Me.Caption & "dgMain", dgMain.Columns(ColIndex).DataField, dgMain.Columns(ColIndex).Width
    
End Sub

Private Sub dgMain_DblClick()
frm_OP_SDNConfirm.txt_OrderKey.Text = rsMain("TMS單號"): frm_OP_SDNConfirm.cmbOrderkey.Text = "TMS單號"
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
    
    '不允許移至特定欄位
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

'str_SQL = "Select 出車日期 , 二次路編 ,路線編號 " & _
'            ",貨主 " & _
'            ",貨主單號 " & _
'            ",快速確認 = ' ' " & _
'            ",TMS單號= 訂單編號 " & _
'            ",訂單類別 " & _
'            ",車牌號碼 " & _
'            ",駕駛人 " & _
'            ",貨運公司 " & _
'            ",備註 = 說明 " & _
'            ",客戶編號 " & _
'            ",客戶名稱 " & _
'            ",送貨地址 " & _
'            ",訂單日期 " & _
'            ",到貨日期 " & _
'            "From SDNConfirm_OrderDate_One " & _
'            "where 簽單已回 = 0 order by 出車日期 ,二次路編 ,路線編號 ,訂單編號 "


str_SQL = "select 出車日期 = convert(varchar,t01t.Delivery_Date,112),二次路編 = t02t.c_Route_No,路線編號 = t02t.Route_No " & _
",貨主 = Rtrim(t02t.StorerKey),訂單號碼 = Rtrim(t02t.Extern) ,快速確認 = ' ',TMS單號= rtrim(t02t.Receipt_No) " & _
",訂單類別 = rtrim(t02t.priority),車牌號碼 = Rtrim(t01t.c_Vehicle_ID_No),駕駛人 = Rtrim(t01t.driver) " & _
",貨運公司 = Isnull(Rtrim(t8m.Short_Name),''),備註 = Rtrim(Isnull(t02t.Description,'')) " & _
",客戶編號 = Rtrim(t02t.ConsigneeKey),客戶名稱 = Rtrim(Isnull(t1m.Short_Name,'')) " & _
",送貨地址 = Rtrim(Isnull(t1m.Address,'')),訂單日期 = rtrim(t02t.Receipt_Date) " & _
",到貨日期 = rtrim(t02t.Arrive_Date) " & _
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

'貨主
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open "select distinct(storerkey) , short_name from trp16M order by storerkey ", cn, adOpenKeyset, adLockPessimistic

cboStorerKey.Clear
cboStorerKey.AddItem ""
Do While Not tmp_Rs.EOF
    cboStorerKey.AddItem RTrim(tmp_Rs("storerkey")) & "_" & RTrim(tmp_Rs("short_name"))
tmp_Rs.MoveNext
Loop
tmp_Rs.Close

'車號
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

'取欄位寬度
SetDataGridColWidth Me.Caption, dgMain

Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
End Sub
