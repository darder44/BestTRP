VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frm_Pallet_CSVehicle_id_no 
   Caption         =   "中南區車號匯入"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   10230
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "配送資料"
      TabPicture(0)   =   "frm_Pallet_CSVehicle_id_no.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dgMain"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "拆單"
      TabPicture(1)   =   "frm_Pallet_CSVehicle_id_no.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dgMainCT"
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   3975
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761087
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
      Begin MSDataGridLib.DataGrid dgMainCT 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   9
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8454016
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   20
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9
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
   Begin VB.Frame Frame20 
      Caption         =   "中南區車號匯入"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2805
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9840
      Begin VB.ComboBox Cbx_Area 
         Height          =   300
         ItemData        =   "frm_Pallet_CSVehicle_id_no.frx":0038
         Left            =   2520
         List            =   "frm_Pallet_CSVehicle_id_no.frx":003A
         Style           =   2  '單純下拉式
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Cmd_Openfiles 
         BackColor       =   &H0080FFFF&
         Caption         =   "開啟檔案"
         Height          =   375
         Left            =   3480
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cboSheet 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5400
         Style           =   2  '單純下拉式
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   4365
      End
      Begin VB.FileListBox filLocalFile 
         Height          =   1530
         Left            =   4560
         Pattern         =   "*.xls"
         TabIndex        =   4
         ToolTipText     =   "僅顯示 ""*.xls"" 檔案"
         Top             =   1200
         Width           =   5190
      End
      Begin VB.DirListBox dirLocalDir 
         Height          =   1560
         Left            =   135
         TabIndex        =   3
         ToolTipText     =   "Local Directory"
         Top             =   1200
         Width           =   4335
      End
      Begin VB.DriveListBox drvLocalDrive 
         Height          =   300
         Left            =   135
         TabIndex        =   2
         ToolTipText     =   "Local Drive List"
         Top             =   750
         Width           =   2040
      End
      Begin VB.CommandButton cmdImport 
         BackColor       =   &H0080FFFF&
         Caption         =   "開始匯入"
         Height          =   375
         Left            =   2400
         Style           =   1  '圖片外觀
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "請先選擇區域再進行匯入:"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "工作表"
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
         Index           =   17
         Left            =   4560
         TabIndex        =   6
         Top             =   900
         Visible         =   0   'False
         Width           =   720
      End
   End
End
Attribute VB_Name = "frm_Pallet_CSVehicle_id_no"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsMain As ADODB.Recordset
Private rsMainCT As ADODB.Recordset
Private rsMainItrn As ADODB.Recordset

Private Sub cboSheet_Click()

On Error GoTo err_Handle
Dim str As String, strFieldName As String, strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFile.Path, 1) = "\" Then
    strFilePath = filLocalFile.Path
Else
    strFilePath = filLocalFile.Path & "\"
End If

'建立欄位名稱陣列
strFieldName = ""

If Right(filLocalFile.Path, 1) <> "\" Then
    strFilePath = filLocalFile.Path & "\"
Else
    strFilePath = filLocalFile.Path
End If

Set rsMain = New ADODB.Recordset

Call Excel2Recordset(strFilePath & filLocalFile.FileName, cboSheet, strFieldName, rsMain)

Set dgMain.DataSource = rsMain

If rsMain Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMain
    MsgBox "此工作表共 " & rsMain.RecordCount & "筆明細，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
End If


Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub



Private Sub Cbx_Area_Click()

If RTrim(Cbx_Area.Text) <> "" Then
    Call EnableMenu   '開啟所有功能
    If Left(RTrim(Cbx_Area.Text), 2) = "CB" Then
        filLocalFile.Pattern = "*中區配送資料車號匯入.xls"
    Else
        filLocalFile.Pattern = "*南區配送資料車號匯入.xls"
    End If
Else
    Call DisableMenu   '關閉所有功能
    
End If

End Sub

Private Sub Cmd_Openfiles_Click()

On Error GoTo err_Handle
Dim str As String, strFieldName As String, strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFile.Path, 1) = "\" Then
    strFilePath = filLocalFile.Path
Else
    strFilePath = filLocalFile.Path & "\"
End If

'建立欄位名稱陣列
strFieldName = ""

If Right(filLocalFile.Path, 1) <> "\" Then
    strFilePath = filLocalFile.Path & "\"
Else
    strFilePath = filLocalFile.Path
End If

Set rsMain = New ADODB.Recordset

Call Excel2Recordset(strFilePath & filLocalFile.FileName, "配送資料", strFieldName, rsMain)

'建立欄位名稱陣列
strFieldName = ""

Call Excel2Recordset(strFilePath & filLocalFile.FileName, "拆單", strFieldName, rsMainCT)

Set dgMain.DataSource = rsMain
Set dgMainCT.DataSource = rsMainCT

If rsMain Is Nothing And rsMainCT Is Nothing Then

    MsgBox "查無資料!", 64, "Excel2Recordset"
    
Else
    SetDataGridColWidth Me.Caption, dgMain
    SetDataGridColWidth Me.Caption, dgMainCT
    rsMain.Sort = "路線編號"
    rsMainCT.Sort = "路線編號"
    MsgBox "配送資料共: " & rsMain.RecordCount & " 筆，拆單資料共: " & rsMainCT.RecordCount & " 筆，請確認筆數與內容是否與原始檔案相符!!", 64, "Excel2Recordset"
    
End If


Exit Sub
err_Handle:

Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdImport_Click()

Dim Str_RouteNo As String, Str_CarNo As String, str_MaxRouteNo As String, str_TMSOrders As String, str_CTOrders As String
Dim str_TMSALLOrders As String, str_TRPType As String, intDriveTimes As Integer
Dim Long_CTcs As Long '拆單箱數
Dim rsTmp As New ADODB.Recordset


On Error GoTo err_Handle

str_TMSALLOrders = ""
If Cbx_Area.Text = "CB中區車號匯入" Then str_TRPType = "CB"
If Cbx_Area.Text = "SB南區車號匯入" Then str_TRPType = "SB"

Tran_Level = 0
'===============================================檢查==================================================================

If (rsMain.RecordCount = 0 Or rsMain Is Nothing) And (rsMainCT.RecordCount = 0 Or rsMainCT Is Nothing) Then Exit Sub

'配送資料
If rsMain.RecordCount = 0 Or rsMain Is Nothing Then
Else
    '配送資料有資料，檢查配送資料有無異常
    Str_RouteNo = "": Str_CarNo = "": str_MaxRouteNo = "": str_TMSOrders = ""
    rsMain.MoveFirst
    Do While Not rsMain.EOF
            '一個路線編號只能有一個車號
            If Str_RouteNo <> Trim(rsMain("路線編號")) Then
                '車號不同則紀錄車號和路線編號
                Str_RouteNo = Trim(rsMain("路線編號"))
                Str_CarNo = Trim(rsMain("車號"))
            Else
                '路線編號相同，則比較車號是否相同
                If Trim(rsMain("車號")) <> Str_CarNo Then
                    Screen.MousePointer = vbDefault
                    MsgBox "路線編號:" & Str_RouteNo & "  出現兩種以上車號:" & Str_CarNo & " ; " & Trim(rsMain("車號")) & "，請確認車號。匯入結束", vbOKOnly + vbCritical, "車號檢查"
                    Exit Sub
                End If
            End If
            
            '車號檢查
            str_SQL = "select vehicle_id_no from trp09m(nolock) where vehicle_id_no = '" & RTrim(rsMain.Fields("車號")) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        
            If tmp_Rs.EOF Then  '按鈕那些要改
                MsgBox "車籍主檔中，找不到此:" & RTrim(rsMain.Fields("車號")) & " 車號，請先於商品主檔新建商品資料，訂單轉入終止!!": Screen.MousePointer = 0
                Exit Sub
            End If
        rsMain.MoveNext
    Loop
    rsMain.MoveFirst
End If

'拆單部份檢查拆單數量&車號重複
If rsMainCT.RecordCount = 0 Or rsMainCT Is Nothing Then
Else
    '拆單資料有資料，檢查拆單資料有無異常
    Str_RouteNo = "": Str_CarNo = "": str_MaxRouteNo = "": str_TMSOrders = "": Long_CTcs = 0
    rsMainCT.MoveFirst
    Do While Not rsMainCT.EOF
            '一個路線編號只能有一個車號
            If Str_RouteNo <> Trim(rsMainCT("路線編號")) Then
                '車號不同則紀錄車號和路線編號
                Str_RouteNo = Trim(rsMainCT("路線編號"))
                Str_CarNo = Trim(rsMainCT("車號"))
            Else
                '路線編號相同，則比較車號是否相同
                If Trim(rsMainCT("車號")) <> Str_CarNo Then
                    Screen.MousePointer = vbDefault
                    MsgBox "路線編號:" & Str_RouteNo & "  出現兩種以上車號:" & Str_CarNo & " ; " & Trim(rsMainCT("車號")) & "，請確認車號。匯入結束", vbOKOnly + vbCritical, "車號檢查"
                    Exit Sub
                End If
            End If
            
            '拆單箱數不可以>=原單箱數
            If Val(rsMainCT.Fields("拆單箱數")) >= Val(rsMainCT.Fields("訂單箱數")) Then
                    Screen.MousePointer = vbDefault
                    MsgBox "TMS單號:" & Val(rsMainCT.Fields("TMS單號")) & "  的拆單箱數:" & Val(rsMainCT.Fields("拆單箱數")) & "大於等於訂單箱數: " & Val(rsMainCT.Fields("訂單箱數")) & " ，請確認車號。匯入結束", vbOKOnly + vbCritical, "車號檢查"
                    Exit Sub
            End If
            
            '車號檢查
            str_SQL = "select vehicle_id_no from trp09m(nolock) where vehicle_id_no = '" & RTrim(rsMainCT.Fields("車號")) & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        
            If tmp_Rs.EOF Then  '按鈕那些要改
                MsgBox "車籍主檔中，找不到此:" & RTrim(rsMainCT.Fields("車號")) & " 車號，請先於商品主檔新建商品資料，訂單轉入終止!!": Screen.MousePointer = 0
                Exit Sub
            End If
            
        rsMainCT.MoveNext
    Loop
    rsMainCT.MoveFirst
End If

'===============================================更新==================================================================
Tran_Level = cn.BeginTrans
DoEvents: DoEvents

'開始更新配送資料總表部份
If rsMain.RecordCount = 0 Or rsMain Is Nothing Then
Else
    Str_RouteNo = "": Str_CarNo = ""
    dgMain.Enabled = False
    rsMain.MoveFirst
    '開始更新sdn02t車號

    Do While Not rsMain.EOF
        If RTrim(rsMain.Fields("車號")) = "000-31" Then '非外車不更新
        Else
            str_TMSALLOrders = str_TMSALLOrders & "'" & Format(Trim(rsMain("TMS單號")), "0000000000") & "',"
                If Str_RouteNo <> Trim(rsMain("路線編號")) Then
                    '車號不同則紀錄車號和路線編號
                    Str_RouteNo = Trim(rsMain("路線編號"))
                    Str_CarNo = Trim(rsMain("車號"))
                    '取得最大路現編號進行新增
                        str_SQL = "select MaxRouteNO = right(max(c_route_no),3)+1,最大路線編號=max(c_route_no) from sdn01t where left(c_route_no,1) = 'N' and convert(char(8),delivery_date,112) = '" & RTrim(rsMain.Fields("到貨日期")) & "'"
                        Call Confirm_Recordset_Closed(tmp_Rs)
                        Call ReDim_Recordset(tmp_Rs)
                        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                        str_MaxRouteNo = "N" & Right(RTrim(rsMain.Fields("到貨日期")), 6) & Format(tmp_Rs.Fields("MaxRouteNO"), "000")
                        tmp_Rs.Close
                        
                        '產生車次
                        str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
                                  "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & RTrim(rsMain.Fields("到貨日期")) & "' and Vehicle_ID_No = '" & RTrim(rsMain.Fields("車號")) & "'"
                        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                        intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
                        tmp_Rs.Close
                        
                        '新增一個路編資料到sdn01t
                        str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,Receiver,SDNStatus,AddUser,Drive_Times) " & _
                        "select " & _
                        "'" & RTrim(rsMain.Fields("到貨日期")) & "' " & _
                        ",'" & str_MaxRouteNo & "' " & _
                        ",'" & RTrim(rsMain.Fields("車號")) & "' " & _
                        ",駕駛 = rtrim(isnull(t9.driver,'')) " & _
                        ",請款人 = rtrim(isnull(t9.receiver,'')) " & _
                        ",sdnstatus = 0 " & _
                        ",adduser = 'Vehicle_Update' " & _
                        ",'" & intDriveTimes & "' " & _
                        "from trp09m t9 " & _
                        "where vehicle_id_no = '" & RTrim(rsMain.Fields("車號")) & "'"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                End If
                
                'sdn02t,sdn03t 更新一次路線編號及車號，
                '新增兩欄位，紀錄更新前的車號和路線編號
                '還沒補TRP_TYPE = "CB"
'                str_SQL = "update s2 " & _
'                        "Set s2.RouteNo_old = s2.Route_no " & _
'                        ",s2.CarNo_old = s2.vehicle_id_no " & _
'                        ",s2.Route_no = '" & str_MaxRouteNo & "' " & _
'                        ",s2.vehicle_id_no = '" & Trim(rsMain("車號")) & "' " & _
'                        ",s2.trp_type = '" & str_TRPType & "' " & _
'                        "from sdn02t s2 " & _
'                        "where receipt_no = '" & Format(Trim(rsMain("TMS單號")), "0000000000") & "'"
                str_SQL = "update s2 " & _
                        "Set s2.Route_no = '" & str_MaxRouteNo & "' " & _
                        ",s2.vehicle_id_no = '" & Trim(rsMain("車號")) & "' " & _
                        ",s2.trp_type = '" & str_TRPType & "' " & _
                        "from sdn02t s2 " & _
                        "where receipt_no = '" & Format(Trim(rsMain("TMS單號")), "0000000000") & "'"
                 cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                 cn.Execute "update s3 set s3.Route_no = '" & str_MaxRouteNo & "' from sdn03t s3 where receipt_no = '" & Format(Trim(rsMain("TMS單號")), "0000000000") & "'", RowsAffect, adExecuteNoRecords
        End If
        rsMain.MoveNext
    Loop
End If


'拆單部份，by訂單排序，先新增有拆單的訂單，再開始更新車號

If rsMainCT.RecordCount = 0 Or rsMainCT Is Nothing Then
Else
rsMainCT.MoveFirst
rsMainCT.Sort = "TMS單號"
Str_RouteNo = "": Str_CarNo = "": str_MaxRouteNo = "": Long_CTcs = 0: str_CTOrders = "": str_TMSOrders = ""
Do While Not rsMainCT.EOF
    If RTrim(rsMainCT.Fields("拆單箱數")) > 0 And RTrim(rsMainCT.Fields("TMS單號")) <> str_TMSOrders Then
        str_TMSOrders = RTrim(rsMainCT.Fields("TMS單號"))
        '取得最大拆單單號進行新增
            str_SQL = "select AvailNo = cast(code as integer) from codelkup where listname = 'cutordersno'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            Call ReDim_Recordset(tmp_Rs)
            tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
            str_CTOrders = "CT" & Format(tmp_Rs.Fields("AvailNo"), "00000000")
            cn.Execute "update codelkup set code = '" & Val(tmp_Rs.Fields("AvailNo")) + 1 & "' where listname = 'cutordersno'", RowsAffect, adExecuteNoRecords
            tmp_Rs.Close
            
        '新增一個sdn02t拆單的訂單資料。
            str_SQL = "insert sdn02t(C_ROUTE_NO, ROUTE_NO, STORERKEY, EXTERN, RECEIPT_DATE, ARRIVE_DATE, CUST_NAME, SHIP_CS, SHIP_CBM, SHIP_WT, CAR_NOTES, SDNStatus, SDN_NOTE, C_Route_Time, C_Route_Total, RECEIPT_NO, OnTimeDelivery, PODOnTime, RejectOrder, DESCRIPTION, CONFIRM_DATE, CONSIGNEEKEY, CONFIRM_USERID, CUSTSIGNDATE, RBCCode, RSCCode, CONFIRM_Notes, PRIORITY, SCHEDULEDATE, CustomerOrderkey1, Scan, SDNSendDate, CUST_Handle, TRP_Handle, Advance, INV_Handle, TRP_Cost, Sorting_Cost, Total_Cost, VEHICLE_ID_NO, ExpectReceiptOK, SdnFeedBack, InvBack, C_RECEIPT_NO, SDNBack, OTQty, OTConfirmUser, Facility, BConsigneekey, ReturnStatus) " & _
                    "select s2.C_ROUTE_NO, s2.ROUTE_NO, s2.STORERKEY, s2.EXTERN, s2.RECEIPT_DATE, s2.ARRIVE_DATE, s2.CUST_NAME, s2.SHIP_CS, s2.SHIP_CBM, s2.SHIP_WT, s2.CAR_NOTES, s2.SDNStatus, s2.SDN_NOTE, s2.C_Route_Time, s2.C_Route_Total, " & _
                    "'" & str_CTOrders & "', s2.OnTimeDelivery, s2.PODOnTime, s2.RejectOrder, s2.DESCRIPTION, s2.CONFIRM_DATE, s2.CONSIGNEEKEY, s2.CONFIRM_USERID, s2.CUSTSIGNDATE, s2.RBCCode, s2.RSCCode, s2.CONFIRM_Notes, s2.PRIORITY, s2.SCHEDULEDATE, " & _
                    "s2.CustomerOrderkey1, s2.Scan, s2.SDNSendDate, s2.CUST_Handle, s2.TRP_Handle, s2.Advance, s2.INV_Handle, s2.TRP_Cost, s2.Sorting_Cost, s2.Total_Cost, s2.VEHICLE_ID_NO, s2.ExpectReceiptOK, s2.SdnFeedBack, s2.InvBack, s2.C_RECEIPT_NO, s2.SDNBack, " & _
                    "s2.OTQty , s2.OTConfirmUser, s2.Facility, s2.BConsigneekey, s2.ReturnStatus " & _
                    "from sdn02t s2 " & _
                    "where s2.receipt_no = '" & Format(Trim(rsMainCT("TMS單號")), "0000000000") & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
         '抓取原本的SDN03T明細，除一筆對品項。進行拆單。
            str_SQL = "select S3.*,sp.casecnt " & _
                    "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no " & _
                    "join sdn01t s1 on s1.c_route_no = s2.c_route_no " & _
                    "join gv_skuxpack sp on sp.storerkey = s2.storerkey and s3.product_no = sp.sku " & _
                    "where s2.receipt_no = '" & Format(Trim(rsMainCT("TMS單號")), "0000000000") & "'"
            Call Confirm_Recordset_Closed(tmp_Rs)
            Call ReDim_Recordset(tmp_Rs)
            tmp_Rs.CursorLocation = adUseClient
            tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
            If tmp_Rs.RecordCount <> 0 Then
                tmp_Rs.MoveFirst
                Do While Not tmp_Rs.EOF
                    Long_CTcs = Val(Trim(rsMainCT("拆單箱數")))
                    If Long_CTcs > 0 Then
                        If Val(tmp_Rs.Fields("ship_qty")) / Val(tmp_Rs.Fields("casecnt")) >= Val(Trim(rsMainCT("拆單箱數"))) Then
                            '夠拆，則insert一筆新的明細
                            str_SQL = "insert sdn03t(C_ROUTE_NO, ROUTE_NO, STORERKEY, RECEIPT_NO, SEQ_NO, SubSeq_No, EXTERN, PRODUCT_NO, SHIP_UNIT, SHIP_QTY, SIGN_QTY, WEIGHT, VOLUMN_WEIGHT, RSC_CODE, RBC_CODE, CONFIRM_DATE, DESCRIPTION, ORDER_QTY, SHIP_TIME, Responsible) " & _
                            "values('" & RTrim(tmp_Rs.Fields("C_ROUTE_NO")) & "','" & RTrim(tmp_Rs.Fields("ROUTE_NO")) & "','" & RTrim(tmp_Rs.Fields("STORERKEY")) & "','" & str_CTOrders & "','" & RTrim(tmp_Rs.Fields("SEQ_NO")) & "','" & _
                             RTrim(tmp_Rs.Fields("SubSeq_No")) & "','" & RTrim(tmp_Rs.Fields("EXTERN")) & "','" & RTrim(tmp_Rs.Fields("PRODUCT_NO")) & "','" & RTrim(tmp_Rs.Fields("SHIP_UNIT")) & "','" & Val(tmp_Rs.Fields("casecnt")) * Val(Trim(rsMainCT("拆單箱數"))) & "','0" & _
                            "','0','0','" & RTrim(tmp_Rs.Fields("RSC_CODE")) & "','" & RTrim(tmp_Rs.Fields("RBC_CODE")) & "','" & RTrim(tmp_Rs.Fields("CONFIRM_DATE")) & "','" & RTrim(tmp_Rs.Fields("DESCRIPTION")) & "','" & Val(tmp_Rs.Fields("casecnt")) * Val(Trim(rsMainCT("拆單箱數"))) & "','" & RTrim(tmp_Rs.Fields("SHIP_TIME")) & "','" & RTrim(tmp_Rs.Fields("Responsible")) & "')"
                            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                            Long_CTcs = Long_CTcs - Val(Trim(rsMainCT("拆單箱數")))
                            '更新sdn03t的數量
                            str_SQL = "update sdn03t set ship_qty = ship_qty - '" & Val(tmp_Rs.Fields("casecnt")) * Val(Trim(rsMainCT("拆單箱數"))) & "',order_qty = order_qty - '" & Val(tmp_Rs.Fields("casecnt")) * Val(Trim(rsMainCT("拆單箱數"))) & "' from sdn03t  where receipt_no = '" & Format(Trim(rsMainCT("TMS單號")), "0000000000") & "' and seq_no = '" & RTrim(tmp_Rs.Fields("seq_no")) & "'"
                            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        Else
                            '不夠拆，直接將sdn03t的line insert 過去
                            str_SQL = "insert sdn03t(C_ROUTE_NO, ROUTE_NO, STORERKEY, RECEIPT_NO, SEQ_NO, SubSeq_No, EXTERN, PRODUCT_NO, SHIP_UNIT, SHIP_QTY, SIGN_QTY, WEIGHT, VOLUMN_WEIGHT, RSC_CODE, RBC_CODE, CONFIRM_DATE, DESCRIPTION, ORDER_QTY, SHIP_TIME, Responsible) " & _
                            "values('" & RTrim(tmp_Rs.Fields("C_ROUTE_NO")) & "','" & RTrim(tmp_Rs.Fields("ROUTE_NO")) & "','" & RTrim(tmp_Rs.Fields("STORERKEY")) & "', '" & str_CTOrders & "', '" & RTrim(tmp_Rs.Fields("SEQ_NO")) & "', '" & RTrim(tmp_Rs.Fields("SubSeq_No")) & "', '" & RTrim(tmp_Rs.Fields("EXTERN")) & _
                            "', '" & RTrim(tmp_Rs.Fields("PRODUCT_NO")) & "', '" & RTrim(tmp_Rs.Fields("SHIP_UNIT")) & "', '" & RTrim(tmp_Rs.Fields("SHIP_QTY")) & "', '" & RTrim(tmp_Rs.Fields("SIGN_QTY")) & "', '" & RTrim(tmp_Rs.Fields("Weight")) & "', '" & RTrim(tmp_Rs.Fields("VOLUMN_WEIGHT")) & "', '" & RTrim(tmp_Rs.Fields("RSC_CODE")) & "', '" & RTrim(tmp_Rs.Fields("RBC_CODE")) & _
                            "', '" & RTrim(tmp_Rs.Fields("CONFIRM_DATE")) & "', '" & RTrim(tmp_Rs.Fields("Description")) & "', '" & Val(RTrim(tmp_Rs.Fields("ORDER_QTY"))) & "', '" & RTrim(tmp_Rs.Fields("SHIP_TIME")) & "', '" & RTrim(tmp_Rs.Fields("Responsible")) & "')"
                            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                            Long_CTcs = Long_CTcs - Val(tmp_Rs.Fields("casecnt")) * Val(tmp_Rs.Fields("ship_qty"))
                        '更新sdn03t的數量
                            str_SQL = "update sdn03t set ship_qty = 0,order_qty = 0 from sdn03t where receipt_no = '" & Format(Trim(rsMainCT("TMS單號")), "0000000000") & "' and seq_no = '" & RTrim(tmp_Rs.Fields("seq_no")) & "'"
                            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                        End If
                    End If
                    tmp_Rs.MoveNext
                Loop
                tmp_Rs.Close
            End If
          '清除sdn03t中，出貨量=0的line
          str_SQL = "delete s3 " & _
                    "from sdn03t s3 " & _
                    "where s3.ship_qty = 0 and s3.receipt_no in ('" & Format(Trim(rsMainCT("TMS單號")), "0000000000") & "','" & str_CTOrders & "') "
          cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
          '最後重新更新兩個單號的sdn02t,sdn03t才積重量，箱數資料
          str_SQL = "Update s2 " & _
                    "set ship_CS = (select sum(sdn03t.ship_qty/sp.casecnt) from sdn03t sdn03t join gv_skuxpack sp on sp.storerkey = sdn03t.storerkey and sdn03t.product_no = sp.sku where sdn03t.receipt_no = s2.receipt_no), " & _
                    "ship_CBM =  (select sum(sdn03t.ship_qty*sp.stdcube) from sdn03t sdn03t join gv_skuxpack sp on sp.storerkey = sdn03t.storerkey and sdn03t.product_no = sp.sku  where sdn03t.receipt_no = s2.receipt_no), " & _
                    "ship_WT = (select sum( sdn03t.ship_qty*sp.stdgrosswgt)from sdn03t sdn03t join gv_skuxpack sp on sp.storerkey = sdn03t.storerkey and sdn03t.product_no = sp.sku  where sdn03t.receipt_no = s2.receipt_no) " & _
                    "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no " & _
                    "where s2.receipt_no in ('" & Format(Trim(rsMainCT("TMS單號")), "0000000000") & "','" & str_CTOrders & "') "
          cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
          str_SQL = "Update s3 " & _
                    "set s3.weight = (select sum(case when isnull(sp.casecnt,0) = 0 then 0 else sdn03t.ship_qty*sp.stdgrosswgt end) from sdn03t sdn03t join gv_skuxpack sp on sp.storerkey = sdn03t.storerkey and sdn03t.product_no = sp.sku where sdn03t.receipt_no = s2.receipt_no), " & _
                    "s3.volumn_weight =(select sum(case when isnull(sp.casecnt,0) = 0 then 0 else sdn03t.ship_qty*sp.stdcube end) from sdn03t sdn03t join gv_skuxpack sp on sp.storerkey = sdn03t.storerkey and sdn03t.product_no = sp.sku where sdn03t.receipt_no = s2.receipt_no) " & _
                    "from sdn02t s2 join sdn03t s3 on s2.receipt_no = s3.receipt_no " & _
                    "where s2.receipt_no in ('" & Format(Trim(rsMainCT("TMS單號")), "0000000000") & "','" & str_CTOrders & "') "
          cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            '更新recordset的訂單號碼=拆單後的訂單。
            rsMainCT.Fields("TMS單號") = str_CTOrders
    End If
    rsMainCT.MoveNext
Loop
    '開始更新車號
    rsMainCT.MoveFirst
    rsMainCT.Sort = "路線編號"
    Str_RouteNo = "": Str_CarNo = ""
    dgMainCT.Enabled = False
    rsMainCT.MoveFirst
    '開始更新sdn02t車號
    Do While Not rsMainCT.EOF
        If RTrim(rsMainCT.Fields("車號")) = "000-31" Then '非外車不更新
        Else
        str_TMSALLOrders = str_TMSALLOrders & "'" & Format(Trim(rsMainCT("TMS單號")), "0000000000") & "',"
                If Str_RouteNo <> Trim(rsMainCT("路線編號")) Then
                    '車號不同則紀錄車號和路線編號
                    Str_RouteNo = Trim(rsMainCT("路線編號"))
                    Str_CarNo = Trim(rsMainCT("車號"))
                    '取得最大路現編號進行新增
                        str_SQL = "select MaxRouteNO = right(max(c_route_no),3)+1,最大路線編號=max(c_route_no) from sdn01t where left(c_route_no,1) = 'N' and convert(char(8),delivery_date,112) = '" & RTrim(rsMainCT.Fields("到貨日期")) & "'"
                        Call Confirm_Recordset_Closed(tmp_Rs)
                        Call ReDim_Recordset(tmp_Rs)
                        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                        str_MaxRouteNo = "N" & Right(RTrim(rsMainCT.Fields("到貨日期")), 6) & Format(tmp_Rs.Fields("MaxRouteNO"), "000")
                        tmp_Rs.Close
                        
                        '產生車次
                        str_SQL = "Select Isnull(Max(Drive_Times)+1,1) as Drive_Times " & _
                                  "From TRP05T Where Convert(varchar(8),Delivery_Date,112) = '" & RTrim(rsMainCT.Fields("到貨日期")) & "' and Vehicle_ID_No = '" & RTrim(rsMainCT.Fields("車號")) & "'"
                        tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
                        intDriveTimes = tmp_Rs.Fields("Drive_Times").Value
                        tmp_Rs.Close
                        
                        '新增一個路編資料到sdn01t
                        str_SQL = "Insert into SDN01T (DELIVERY_DATE,C_Route_No,C_VEHICLE_ID_NO,Driver,Receiver,SDNStatus,AddUser,Drive_Times) " & _
                        "select " & _
                        "'" & RTrim(rsMainCT.Fields("到貨日期")) & "' " & _
                        ",'" & str_MaxRouteNo & "' " & _
                        ",'" & RTrim(rsMainCT.Fields("車號")) & "' " & _
                        ",駕駛 = rtrim(isnull(t9.driver,'')) " & _
                        ",請款人 = rtrim(isnull(t9.receiver,'')) " & _
                        ",sdnstatus = 0 " & _
                        ",adduser = 'Vehicle_Update' " & _
                        ",'" & intDriveTimes & "' " & _
                        "from trp09m t9 " & _
                        "where vehicle_id_no = '" & RTrim(rsMainCT.Fields("車號")) & "'"
                        cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                End If
                
                'sdn02t,sdn03t 更新一次路線編號及車號，
                '新增兩欄位，紀錄更新前的車號和路線編號
                '還沒補TRP_TYPE = "CB"
                str_SQL = "update s2 " & _
                        "Set s2.Route_no = '" & str_MaxRouteNo & "' " & _
                        ",s2.vehicle_id_no = '" & Trim(rsMainCT("車號")) & "' " & _
                        ",s2.trp_type = '" & str_TRPType & "'" & _
                        "from sdn02t s2 " & _
                        "where receipt_no = '" & Format(Trim(rsMainCT("TMS單號")), "0000000000") & "'"
                 cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
                 cn.Execute "update s3 set s3.Route_no = '" & str_MaxRouteNo & "' from sdn03t s3 where receipt_no = '" & Format(Trim(rsMainCT("TMS單號")), "0000000000") & "'", RowsAffect, adExecuteNoRecords
                 

        End If

        rsMainCT.MoveNext
    Loop
End If

    str_TMSALLOrders = Mid(str_TMSALLOrders, 1, Len(str_TMSALLOrders) - 1)
    cn.CommitTrans: Tran_Level = 0
    '撈出這次更新的Sdn資料
    str_SQL = "select " & _
            "二次路線編號 = RTrim(s2.c_route_no) " & _
            ",路線編號=rtrim(s2.route_no) " & _
            ",車號=rtrim(s2.vehicle_id_no) " & _
            ",TMS單號=rtrim(s2.receipt_no) " & _
            ",訂單號碼=rtrim(s2.extern) " & _
            ",客戶名稱 = rtrim(s2.cust_name) " & _
            ",品號 = rtrim(s3.product_no) " & _
            ",箱數 = rtrim(s2.ship_cs) " & _
            ",材積 = rtrim(s2.ship_CBM) " & _
            ",重量 =  rtrim(s2.ship_WT) " & _
            ",標示 = rtrim(s2.trp_type) " & _
            ",訂單量 = sum(s3.order_qty) " & _
            ",出貨量 = sum(s3.ship_qty) " & _
            "from sdn02t s2 (nolock) join sdn03t s3 (nolock) on s2.receipt_no = s3.receipt_no " & _
            "where s2.receipt_no in (" & str_TMSALLOrders & ") " & _
            "group by  RTrim(s2.c_route_no),rtrim(s2.route_no) ,rtrim(s2.vehicle_id_no) ,rtrim(s2.routeno_old) ,rtrim(s2.carno_old) ,rtrim(s2.receipt_no),rtrim(s2.extern) ,rtrim(s2.cust_name) ,rtrim(s3.product_no),rtrim(s2.ship_cs), rtrim(s2.ship_CBM), rtrim(s2.ship_WT),rtrim(s2.trp_type)  " & _
            "order by RTrim(s2.c_route_no) "

    Call Confirm_Recordset_Closed(tmp_Rs)
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        
    Recordset2Excel "中南倉車號更新", tmp_Rs
    '轉出交易EXCEL
    Set MyXlsApp = Nothing
    tmp_Rs.Close
    dgMain.Enabled = True
    msg_text = "中南區車號匯入完成"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    
    
    Set dgMain.DataSource = Nothing
    Set dgMainCT.DataSource = Nothing
    
    Cbx_Area.Text = Cbx_Area.List(0)
    Call DisableMenu
    
    Exit Sub

err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
    Set dgMain.DataSource = Nothing
    Set dgMainCT.DataSource = Nothing
End Sub





Private Sub dirLocalDir_Change()
    filLocalFile.Path = dirLocalDir.Path
End Sub

Private Sub drvLocalDrive_Change()
On Error GoTo DriveError
dirLocalDir.Path = drvLocalDrive.Drive
Exit Sub

DriveError:
MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
Resume Next
End Sub

Private Sub filLocalFile_Click()

On Error GoTo err_Handle
Set rsMain = Nothing: Set dgMain.DataSource = rsMain
Set rsMainCT = Nothing: Set dgMainCT.DataSource = rsMainCT
Dim strFilePath As String

'確認路徑是否帶"\"
If Right(filLocalFile.Path, 1) = "\" Then
    strFilePath = filLocalFile.Path
Else
    strFilePath = filLocalFile.Path & "\"
End If

If Dir(strFilePath & filLocalFile.FileName) = "" Then: filLocalFile.Refresh: Exit Sub

cboSheet.Clear

If UCase(mySplit(filLocalFile.FileName, ".", -1)) = "XLS" Then
    Set MyXlsApp = CreateObject("Excel.Application")
    MyXlsApp.Visible = False
    MyXlsApp.Workbooks.Open (strFilePath & filLocalFile.FileName)
    MyXlsApp.DisplayAlerts = False

    '列出所有工作表
    blDo = False
    cboSheet.Clear

    Dim i As Integer
    For i = 1 To MyXlsApp.Sheets.Count
'        DoEvents: DoEvents
        
        cboSheet.AddItem MyXlsApp.Sheets(i).Name
  
        '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
'        MyXlsApp.Sheets(i).Name = MyXlsApp.Sheets(i).Name
    Next
    cboSheet.ListIndex = -1

    '有些時候使用Microsoft.Jet.OLEDB.4.0來讀取XLS，Sheet必須重新命名存檔才能正確抓到
    'MyXlsApp.ActiveWorkbook.SaveAs strFilePath & filLocalFileT5.FileName

    MyXlsApp.Quit: Set MyXlsApp = Nothing
    blDo = True
Else
    cboSheet.Clear

End If

Exit Sub
err_Handle:
Set MyXlsApp = Nothing
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "中南區車號匯入")
End Sub


Private Sub Form_Load()
Me.Height = 8000: Me.Width = 10500
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200

Call DisableMenu   '關閉所有功能
Cbx_Area.AddItem "    "
Cbx_Area.AddItem "CB中區車號匯入"
Cbx_Area.AddItem "SB南區車號匯入"


End Sub

Public Function EnableMenu()
'打開所有Menu
drvLocalDrive.Enabled = True
dirLocalDir.Enabled = True
cmdImport.Enabled = True
Cmd_Openfiles.Enabled = True
SSTab1.Enabled = True
End Function

Public Function DisableMenu()
'關閉所有Menu
drvLocalDrive.Enabled = False
dirLocalDir.Enabled = False
cmdImport.Enabled = False
Cmd_Openfiles.Enabled = False
SSTab1.Enabled = False
End Function

