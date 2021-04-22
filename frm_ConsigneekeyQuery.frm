VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_ConsigneekeyQuery 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶編號查詢"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   14025
   StartUpPosition =   2  '螢幕中央
   Begin VB.TextBox txtCostomerName 
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
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   1665
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "搜尋"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmd2Excel 
      Caption         =   "轉Excel"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dgMain 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   7435
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
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '對齊表單下方
      Height          =   270
      Left            =   0
      TabIndex        =   5
      Top             =   4845
      Width           =   14025
      _ExtentX        =   24739
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Text            =   "狀態"
            TextSave        =   "狀態"
            Object.ToolTipText     =   "狀態"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   18574
            MinWidth        =   2646
            Object.ToolTipText     =   "資料筆數"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.ToolTipText     =   "使用者"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "客戶名稱"
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
      TabIndex        =   3
      Top             =   180
      Width           =   960
   End
End
Attribute VB_Name = "frm_ConsigneekeyQuery"
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

Private Sub cmdQuery_Click()

If Len(RTrim(txtCostomerName)) = 0 Then
    rsMain.Filter = ""
    rsMain.Sort = "編號"
Else
    rsMain.Filter = "客戶名稱 like '%" & RTrim(txtCostomerName) & "%'"
End If

StatusBar.Panels(2).Text = rsMain.RecordCount & " 筆資料列"

End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

    If Len(dgMain.Columns(ColIndex).DataField) = 0 Then Exit Sub
    SaveSetting App.title, Me.Caption & "dgMain", dgMain.Columns(ColIndex).DataField, dgMain.Columns(ColIndex).Width
    
End Sub

Private Sub dgMain_DblClick()

If mySplit(strDataList_Caller, " ", 0) & " " & mySplit(strDataList_Caller, " ", 1) = "frm_OP_ManualOrders txt_ConsigneeKey" Then
    frm_OP_ManualOrders.txt_ConsigneeKey = rsMain("客戶編號")
    Call frm_OP_ManualOrders.txt_ConsigneeKey_LostFocus
ElseIf mySplit(strDataList_Caller, " ", 0) & " " & mySplit(strDataList_Caller, " ", 1) = "frm_OP_ManualOrders txtShipToKey" Then
    frm_OP_ManualOrders.txtShipToKey = rsMain("客戶編號")
    Call frm_OP_ManualOrders.txtShipToKey_LostFocus
End If

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

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call cmdQuery_Click
End Sub

Private Sub Form_Load()

Screen.MousePointer = 11
'strDataList_Caller
str_SQL = "Select 貨主編號 = rtrim(a1.storerkey) " & _
            ",Rtrim(a1.ConsigneeKey) as 客戶編號 " & _
            ", Rtrim(Isnull(a1.Full_Name,'')) as 客戶名稱 " & _
            ", Rtrim(Isnull(a1.Short_Name,'')) as 客戶簡稱 " & _
            ", Rtrim(Isnull(a1.Area_Code,'')) as 運送區碼 " & _
            ", Rtrim(Isnull(a1.ZIP,'')) as 郵遞區號 " & _
            ", Rtrim(Isnull(a1.Address,'')) as 運送地址 " & _
            ", Rtrim(Isnull(a1.Contact,'')) as 聯絡人 " & _
            ", Rtrim(Isnull(a1.Phone,'')) as 電話 " & _
            ", IsNull(Rtrim(g1.Description),' ') as 特殊需求1 " & _
            ", IsNull(Rtrim(g2.Description),' ') as 特殊需求2 " & _
            "From TRP01M a1 Left outer join TRP03M b1 on b1.Area_Code = a1.Area_Code " & _
            "Left outer join TRP02M b2 on b2.ZIP = a1.ZIP " & _
            "Left outer join TRP04M g1 on g1.Extra_Demand_Code = a1.Extra_Demand_Code " & _
            "Left outer join TRP04M g2 on g2.Extra_Demand_Code = a1.Extra_Demand_Code2  " & _
            "where a1.storerkey = '" & mySplit(strDataList_Caller, " ", 2) & "' order by a1.consigneekey"

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

Call Replication_Recordset(tmp_Rs, rsMain)
tmp_Rs.Close

If Not rsMain.EOF Then rsMain.MoveFirst

Set dgMain.DataSource = rsMain

StatusBar.Panels(2).Text = rsMain.RecordCount & " 筆資料列"
StatusBar.Panels(3).Text = User_id

'取欄位寬度
SetDataGridColWidth Me.Caption, dgMain

Screen.MousePointer = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
End Sub
