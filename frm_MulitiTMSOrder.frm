VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_MulitiTMSOrder 
   BorderStyle     =   1  '單線固定
   Caption         =   "貨主單號查詢"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   10410
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmd2Excel 
      Caption         =   "轉Excel"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dgMain 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4895
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
End
Attribute VB_Name = "frm_MulitiTMSOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As New ADODB.Recordset
Private intColumnIndex As Integer

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

Private Sub Form_Load()
Screen.MousePointer = 11

If Trim(frm_OP_SDNConfirm.cmbOrderkey.Text) = "" Then
        str_SQL = "select 狀態 = Isnull(Rtrim(t02t.Confirm_Notes),''),出車日期 = convert(varchar,t01t.Delivery_Date,112) ,二次路編 = t02t.c_Route_No ,路線編號 = t02t.Route_No " & _
                    ",貨主 = Rtrim(t02t.StorerKey),貨主名稱 =rtrim(t16.c_name),TMS單號 = rtrim(t02t.Receipt_No),貨主單號 = Rtrim(t02t.Extern) " & _
                    ",訂單類別 = rtrim(t02t.priority) ,客戶訂單類別 = isnull(o.externordertype,'') ,車牌號碼 = Rtrim(t01t.c_Vehicle_ID_No),駕駛人 = Rtrim(t01t.driver) " & _
                    ",備註 = Rtrim(Isnull(t02t.Description,'')),客戶編號 = Rtrim(t02t.ConsigneeKey),客戶名稱 = Rtrim(Isnull(t1m.Short_Name,'')) " & _
                    ",送貨地址 = Rtrim(Isnull(t1m.Address,'')),訂單日期 = rtrim(t02t.Receipt_Date),到貨日期 = rtrim(t02t.Arrive_Date) " & _
                    "From SDN02T t02t (nolock) join SDN01T t01t (nolock) on t02t.c_route_no = t01t.c_route_no " & _
                    "join orders o (nolock) on o.orderkey = t02t.c_receipt_no " & _
                    "join trp16m t16 (nolock) on t16.STORERKEY = t02t.storerkey " & _
                    "join trp01m t1m (nolock) on t02t.consigneekey = t1m.consigneekey and t02t.storerkey = t1m.storerkey " & _
                    "where t02t.Extern like '" & Trim(frm_OP_SDNConfirm.txt_OrderKey.Text) & "%' or  t02t.Receipt_No like '" & Trim(frm_OP_SDNConfirm.txt_OrderKey.Text) & "%' or t02t.Receipt_No like '" & Format(Trim(frm_OP_SDNConfirm.txt_OrderKey.Text), "0000000000") & "%' " & _
                    "order by convert(varchar,t01t.Delivery_Date,112) ,t02t.c_Route_No ,t02t.Route_No ,t02t.Receipt_No "
ElseIf Trim(frm_OP_SDNConfirm.cmbOrderkey.Text) = "TMS單號" Then
        str_SQL = "select 狀態 = Isnull(Rtrim(t02t.Confirm_Notes),''),出車日期 = convert(varchar,t01t.Delivery_Date,112) ,二次路編 = t02t.c_Route_No ,路線編號 = t02t.Route_No " & _
                    ",貨主 = Rtrim(t02t.StorerKey),貨主名稱 =rtrim(t16.c_name),TMS單號 = rtrim(t02t.Receipt_No),貨主單號 = Rtrim(t02t.Extern) " & _
                    ",訂單類別 = rtrim(t02t.priority) ,客戶訂單類別 = isnull(o.externordertype,'') ,車牌號碼 = Rtrim(t01t.c_Vehicle_ID_No),駕駛人 = Rtrim(t01t.driver) " & _
                    ",備註 = Rtrim(Isnull(t02t.Description,'')),客戶編號 = Rtrim(t02t.ConsigneeKey),客戶名稱 = Rtrim(Isnull(t1m.Short_Name,'')) " & _
                    ",送貨地址 = Rtrim(Isnull(t1m.Address,'')),訂單日期 = rtrim(t02t.Receipt_Date),到貨日期 = rtrim(t02t.Arrive_Date) " & _
                    "From SDN02T t02t (nolock) join SDN01T t01t (nolock) on t02t.c_route_no = t01t.c_route_no " & _
                    "join orders o (nolock) on o.orderkey = t02t.c_receipt_no " & _
                    "join trp16m t16 (nolock) on t16.STORERKEY = t02t.storerkey " & _
                    "join trp01m t1m (nolock) on t02t.consigneekey = t1m.consigneekey and t02t.storerkey = t1m.storerkey " & _
                    "where t02t.Receipt_No like '" & Format(Trim(frm_OP_SDNConfirm.txt_OrderKey.Text), "0000000000") & "%' " & _
                    "order by convert(varchar,t01t.Delivery_Date,112) ,t02t.c_Route_No ,t02t.Route_No ,t02t.Receipt_No "
    'Where receipt_no = '" & strOrderkey & "' "
ElseIf Trim(frm_OP_SDNConfirm.cmbOrderkey.Text) = "貨主單號" Then
        str_SQL = "select 狀態 = Isnull(Rtrim(t02t.Confirm_Notes),''),出車日期 = convert(varchar,t01t.Delivery_Date,112) ,二次路編 = t02t.c_Route_No ,路線編號 = t02t.Route_No " & _
                    ",貨主 = Rtrim(t02t.StorerKey),貨主名稱 =rtrim(t16.c_name),TMS單號 = rtrim(t02t.Receipt_No),貨主單號 = Rtrim(t02t.Extern) " & _
                    ",訂單類別 = rtrim(t02t.priority) ,客戶訂單類別 = isnull(o.externordertype,'') ,車牌號碼 = Rtrim(t01t.c_Vehicle_ID_No),駕駛人 = Rtrim(t01t.driver) " & _
                    ",備註 = Rtrim(Isnull(t02t.Description,'')),客戶編號 = Rtrim(t02t.ConsigneeKey),客戶名稱 = Rtrim(Isnull(t1m.Short_Name,'')) " & _
                    ",送貨地址 = Rtrim(Isnull(t1m.Address,'')),訂單日期 = rtrim(t02t.Receipt_Date),到貨日期 = rtrim(t02t.Arrive_Date) " & _
                    "From SDN02T t02t (nolock) join SDN01T t01t (nolock) on t02t.c_route_no = t01t.c_route_no " & _
                    "join orders o (nolock) on o.orderkey = t02t.c_receipt_no " & _
                    "join trp16m t16 (nolock) on t16.STORERKEY = t02t.storerkey " & _
                    "join trp01m t1m (nolock) on t02t.consigneekey = t1m.consigneekey and t02t.storerkey = t1m.storerkey " & _
                    "where t02t.Extern like '" & Trim(frm_OP_SDNConfirm.txt_OrderKey.Text) & "%' " & _
                    "order by convert(varchar,t01t.Delivery_Date,112) ,t02t.c_Route_No ,t02t.Route_No ,t02t.Receipt_No "
    'Where extern
End If

'str_SQL = "select 狀態 = Isnull(Rtrim(t02t.Confirm_Notes),''),出車日期 = convert(varchar,t01t.Delivery_Date,112) ,二次路編 = t02t.c_Route_No ,路線編號 = t02t.Route_No " & _
'            ",貨主 = Rtrim(t02t.StorerKey),貨主名稱 =rtrim(t16.c_name),TMS單號 = rtrim(t02t.Receipt_No),貨主單號 = Rtrim(t02t.Extern) " & _
'            ",訂單類別 = rtrim(t02t.priority) ,車牌號碼 = Rtrim(t01t.c_Vehicle_ID_No),駕駛人 = Rtrim(t01t.driver) " & _
'            ",備註 = Rtrim(Isnull(t02t.Description,'')),客戶編號 = Rtrim(t02t.ConsigneeKey),客戶名稱 = Rtrim(Isnull(t1m.Short_Name,'')) " & _
'            ",送貨地址 = Rtrim(Isnull(t1m.Address,'')),訂單日期 = rtrim(t02t.Receipt_Date),到貨日期 = rtrim(t02t.Arrive_Date) " & _
'            "From SDN02T t02t (nolock) join SDN01T t01t (nolock) on t02t.c_route_no = t01t.c_route_no " & _
'            "join trp16m t16 (nolock) on t16.STORERKEY = t02t.storerkey " & _
'            "join trp01m t1m (nolock) on t02t.consigneekey = t1m.consigneekey and t02t.storerkey = t1m.storerkey " & _
'            "where t02t.Extern = '" & Trim(frm_OP_SDNConfirm.txt_OrderKey.Text) & "' or  t02t.Receipt_No = '" & Trim(frm_OP_SDNConfirm.txt_OrderKey.Text) & "' or t02t.Receipt_No = '" & Format(Trim(frm_OP_SDNConfirm.txt_OrderKey.Text), "0000000000") & "' " & _
'            "order by convert(varchar,t01t.Delivery_Date,112) ,t02t.c_Route_No ,t02t.Route_No ,t02t.Receipt_No "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

Call Replication_Recordset(tmp_Rs, rsMain)
tmp_Rs.Close

rsMain.MoveFirst
Set dgMain.DataSource = rsMain

'取欄位寬度
SetDataGridColWidth Me.Caption, dgMain

Screen.MousePointer = 0
End Sub
