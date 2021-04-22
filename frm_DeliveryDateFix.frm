VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_DeliveryDateFix 
   BorderStyle     =   1  '單線固定
   Caption         =   "預定到貨時間安排"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   12735
   StartUpPosition =   2  '螢幕中央
   Begin MSComCtl2.DTPicker dtp_OneOrder_SignDate 
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy/MM/dd HH:mm:ss"
      Format          =   78381059
      UpDown          =   -1  'True
      CurrentDate     =   39438
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dgMain 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   12495
      _ExtentX        =   22040
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
   Begin VB.Label lblStatus 
      Alignment       =   2  '置中對齊
      Caption         =   "有備註需預估到貨日期時間"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   240
      TabIndex        =   3
      ToolTipText     =   "Connection Status"
      Top             =   180
      Width           =   3135
   End
End
Attribute VB_Name = "frm_DeliveryDateFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As New ADODB.Recordset

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()

On Error GoTo err_Handle
Screen.MousePointer = 11

rsMain.MoveFirst
Do While Not rsMain.EOF
    '更新TRP02T
    Tran_Level = cn.BeginTrans
    
    str_SQL = "update trp02t set scheduledate = '" & rsMain("預估到貨日期時間") & "' where receipt_no = '" & rsMain("TMS單號") & "' "
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    cn.CommitTrans: Tran_Level = 0
    rsMain.MoveNext
Loop

Screen.MousePointer = 0
Call cmdExit_Click
Exit Sub

err_Handle:
    Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

    If Len(dgMain.Columns(ColIndex).DataField) = 0 Then Exit Sub
    SaveSetting App.title, Me.Caption & "dgMain", dgMain.Columns(ColIndex).DataField, dgMain.Columns(ColIndex).Width
    
End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

'不允許移至特定欄位
With dgMain
    If .Col = 0 Or .Col > 1 Then .Col = 1
End With
If rsMain.EOF = False Then dtp_OneOrder_SignDate.Value = rsMain("預估到貨日期時間")

End Sub

Private Sub dtp_OneOrder_SignDate_Change()
rsMain("預估到貨日期時間") = dtp_OneOrder_SignDate.Value
End Sub

Private Sub Form_Load()
Screen.MousePointer = 11

str_SQL = "Select 預估到貨日期時間 = isnull(a1.scheduledate,a1.Arrive_Date) " & _
        ",車號 = a1.vehicle_id_no " & _
        ",送貨日 = a1.Arrive_Date " & _
        ",訂單備註 = Rtrim(a1.description) " & _
        ",貨主 = RTrim(a1.StorerKey) " & _
        ",客戶名稱 = isnull( Rtrim(a2.Full_Name),'x') " & _
        ",客戶簡稱 = Rtrim(a2.Short_Name) " & _
        ",運送地址 = isnull(Rtrim(a2.Address),'x') " & _
        ",車種需求 = Rtrim(Isnull(a2.Vehicle_Type,'x')) " & _
        ",特殊需求1 = Case When b2.Description = '無特殊需求' Then 'X' else Rtrim(Isnull(b2.Description,'')) End " & _
        ",特殊需求2 = Case When b3.Description = '無特殊需求' Then 'X' else Rtrim(Isnull(b3.Description,'')) End " & _
        ",TMS單號 = a1.receipt_no " & _
        "From TRP02t a1 " & _
        "left outer join TRP01M a2 on a2.ConsigneeKey = a1.ConsigneeKey and a1.storerkey = a2.storerkey " & _
        "Left outer join TRP04M b2 on b2.Extra_Demand_Code = a2.Extra_Demand_Code " & _
        "Left outer join TRP04M b3 on b3.Extra_Demand_Code = a2.Extra_Demand_Code2 " & _
        "where len(Rtrim(a1.description)) > 0 and a1.route_no = '" & strDeliveryDateFiRouteNo & "' " & _
        "order by isnull(a1.scheduledate,a1.Arrive_Date) "

Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

Call Replication_Recordset(tmp_rs, rsMain)

rsMain.MoveFirst
Set dgMain.DataSource = rsMain

'取欄位寬度
SetDataGridColWidth Me.Caption, dgMain

dtp_OneOrder_SignDate.Value = rsMain("預估到貨日期時間")

Screen.MousePointer = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
End Sub
