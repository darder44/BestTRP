VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frm_OP_PalletExport 
   Caption         =   "交易資料匯出"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   BeginProperty Font 
      Name            =   "細明體"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   8370
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3240
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   3916
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   -2147483624
      StartOfWeek     =   103022593
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   12
      Top             =   2160
      Width           =   8295
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   2295
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4048
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8295
      Begin VB.TextBox txtOrderDateE 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1485
      End
      Begin VB.TextBox txtOrderDateS 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1485
      End
      Begin VB.TextBox txtCheckNoS 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1080
         Width           =   1485
      End
      Begin VB.TextBox txtCheckNoE 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1080
         Width           =   1485
      End
      Begin VB.TextBox txtExternOrderkeyS 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.TextBox txtExternOrderkeyE 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         MaxLength       =   12
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.ComboBox cboUserType 
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
         TabIndex        =   0
         Top             =   240
         Width           =   1515
      End
      Begin VB.CommandButton cmd2TXT 
         BackColor       =   &H00FF8080&
         Caption         =   "轉CSV"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   5880
         Picture         =   "frm_OP_PalletExport.frx":0000
         Style           =   1  '圖片外觀
         TabIndex        =   10
         Top             =   1200
         Width           =   1065
      End
      Begin VB.CommandButton cmd2Excel 
         BackColor       =   &H00FFC0C0&
         Caption         =   "轉Excel"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   5880
         Picture         =   "frm_OP_PalletExport.frx":212A
         Style           =   1  '圖片外觀
         TabIndex        =   8
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FF80FF&
         Caption         =   "離開"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   7080
         Picture         =   "frm_OP_PalletExport.frx":3424
         Style           =   1  '圖片外觀
         TabIndex        =   11
         Top             =   1200
         Width           =   1065
      End
      Begin VB.CommandButton cmdReset 
         BackColor       =   &H00C0FFC0&
         Caption         =   "重設"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   7080
         Picture         =   "frm_OP_PalletExport.frx":2D036
         Style           =   1  '圖片外觀
         TabIndex        =   9
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmdQuery 
         BackColor       =   &H00FFFFC0&
         Caption         =   "查詢"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   4680
         Picture         =   "frm_OP_PalletExport.frx":2D348
         Style           =   1  '圖片外觀
         TabIndex        =   7
         Top             =   240
         Width           =   1065
      End
      Begin MSComDlg.CommonDialog CmnDialog 
         Left            =   4920
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "Key單日期"
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
         TabIndex        =   23
         Top             =   765
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "～"
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
         Index           =   1
         Left            =   2655
         TabIndex        =   22
         Top             =   780
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "～"
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
         Index           =   23
         Left            =   2655
         TabIndex        =   21
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "單號"
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
         Index           =   22
         Left            =   120
         TabIndex        =   20
         Top             =   1125
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "～"
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
         Index           =   3
         Left            =   2655
         TabIndex        =   19
         Top             =   1500
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "訂單號碼"
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
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   1485
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "倉庫別"
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
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   300
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '對齊表單下方
      Height          =   270
      Left            =   0
      TabIndex        =   16
      Top             =   6030
      Width           =   8370
      _ExtentX        =   14764
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
            Object.Width           =   8123
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
End
Attribute VB_Name = "frm_OP_PalletExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cmd2Excel_Click()

'資料排序
If rsMain Is Nothing Then MsgBox "無資料可供轉檔！", vbOKOnly + vbInformation, "轉檔": Exit Sub
rsMain.Sort = "單據日期,單號,項次"
Recordset2Excel Me.Caption, rsMain
'..在此編輯EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmd2TXT_Click()

'轉文字檔
If rsMain Is Nothing Then MsgBox "無資料可供轉檔！", vbOKOnly + vbInformation, "轉檔": Exit Sub
If rsMain.RecordCount = 0 Then MsgBox "無資料可供轉檔！", vbOKOnly + vbInformation, "轉檔": Exit Sub

rsMain.Sort = "單據日期,單號,項次"

Call DocStoreDirectory(strDocPath)

Dim strTranFileName As String           '文字檔檔案名稱
CmnDialog.DialogTitle = "轉存文字檔"
CmnDialog.InitDir = "C:\"
CmnDialog.FileName = "WH2_Pallet" & Format(Now, "YYYYMMDD")
CmnDialog.Filter = "純文字檔(*.csv)|*.csv"
CmnDialog.FilterIndex = 1
CmnDialog.CancelError = True
On Error Resume Next
CmnDialog.Flags = cdlOFNHideReadOnly    '隱藏唯讀核取方塊
CmnDialog.ShowOpen

'If Err.Number = cdlCancel Then Exit Sub
If err.Number = cdlCancel Then          '於 [開啟舊檔] 對話方塊中，按下 [取消] 鈕
    Exit Sub
'   msg_text = "選擇 [取消] 按鈕，必須於文字檔中自行存檔"
'   MsgBox msg_text, vbQuestion + vbOKOnly, msg_title
'   strTranFileName = "C:\WH2_Pallet" & Format(Now, "YYYYMMDD") & ".csv"
Else
   strTranFileName = CmnDialog.FileName
   If Dir(strTranFileName) <> "" Then
      Kill strTranFileName
   End If
End If

On Error GoTo err_Handle
Screen.MousePointer = 11: cmd2TXT.Enabled = False: dgMain.Enabled = False: DoEvents
If SaveTo_TextFile(strTranFileName, rsMain, Me.Name & Me.Caption & "列印") = 1 Then
   MsgBox funRtn_msg, vbInformation + vbOKOnly, msg_title
   If rsMain Is Nothing Then Exit Sub
   rsMain.MoveFirst
Else
   If Len(strTranFileName) > 0 Then
      msg_text = "轉存作業完成，檔案存放位置：" & strTranFileName
      MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   End If
End If
Me.WindowState = 2
Screen.MousePointer = 0: cmd2TXT.Enabled = True: dgMain.Enabled = True
Exit Sub

err_Handle:
Screen.MousePointer = 0: cmd2TXT.Enabled = True: dgMain.Enabled = True
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_DeliveryDate As String, chc_Orderdate As String, chc_ExternOrderkey, chc_Status As String

str_SQL = "select " & _
"  單號 = CheckNo " & _
", 項次 = LineNumber " & _
", 貨主 = Storer " & _
", 車號 = CarNo " & _
", 倉庫別 = UserType " & _
", 客戶 = Customer " & _
", 客戶單號 = CustomerSheetNo " & _
", 接收日期 = ChargeDate " & _
", 借出 = QtyIn " & _
", 還入 = QtyOut " & _
", 備註 = Notes " & _
", 單據日期 = AddDate " & _
", 輸入日期 = KeyinDate " & _
", EditDate = isnull(EditDate , '') " & _
", CheckDate = isnull(CheckDate , '') " & _
", Adduser " & _
", EditUser = isnull(EditUser , '') " & _
", CheckUser = isnull(CheckUser , '') " & _
", KeyID = isnull(KeyID , '') " & _
"from pallet_cst where 1 = 1 "

'Keyin日期
chc_Orderdate = ""
If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and convert(char(8) , keyindate,112) between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
   chc_Orderdate = "and convert(char(8) , keyindate,112) = '" & txtOrderDateS.Text & "' "
ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and convert(char(8) , keyindate,112) = '" & txtOrderDateE.Text & "' "
End If

''出車日期
'chc_DeliveryDate = ""
'If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
'   chc_Orderdate = "and convert(char(8) , s01t.delivery_date,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
'ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
'   chc_Orderdate = "and convert(char(8) , s01t.delivery_date,112) = '" & txtDeliveryDateS.Text & "' "
'ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
'   chc_Orderdate = "and convert(char(8) , s01t.delivery_date,112) = '" & txtDeliveryDateE.Text & "' "
'End If
'

'單號
If (Len(RTrim(txtCheckNoS.Text)) > 0 And Len(RTrim(txtCheckNoE.Text)) = 0) Or (Len(RTrim(txtCheckNoS.Text)) = 0 And Len(RTrim(txtCheckNoE.Text)) > 0) Then str_SQL = str_SQL & "and CheckNo = '" & RTrim(txtCheckNoS.Text) & RTrim(txtCheckNoE.Text) & "' "
If (Len(RTrim(txtCheckNoS.Text)) > 0 And Len(RTrim(txtCheckNoE.Text)) > 0) Then str_SQL = str_SQL & "and CheckNo between '" & RTrim(txtCheckNoS.Text) & "'and'" & RTrim(txtCheckNoE.Text) & "' "


''訂單號碼
'chc_ExternOrderkey = ""
'If Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) > 0 Then
'   chc_ExternOrderkey = "and rtrim(s02t.extern) between '" & txtExternOrderkeyS.Text & "' and '" & txtExternOrderkeyE.Text & "' "
'ElseIf Len(txtExternOrderkeyS.Text) > 0 And Len(txtExternOrderkeyE.Text) = 0 Then
'   chc_ExternOrderkey = "and rtrim(s02t.extern) = '" & txtExternOrderkeyS.Text & "' "
'ElseIf Len(txtExternOrderkeyS.Text) = 0 And Len(txtExternOrderkeyE.Text) > 0 Then
'   chc_ExternOrderkey = "and rtrim(s02t.extern) = '" & txtExternOrderkeyE.Text & "' "
'End If

'庫別
Dim chcUsertype As String
chcUsertype = "and rtrim(usertype) = '" & cboUserType.Text & "' "
If Len(Trim(cboUserType)) = 0 Then chcUsertype = ""

'組合字串
str_SQL = str_SQL & chc_DeliveryDate & chc_Orderdate & chc_Status & chc_ExternOrderkey & chcUsertype

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
tmp_Rs.Sort = "單據日期,單號,項次"
Call Replication_Recordset(tmp_Rs, rsMain)
tmp_Rs.Close: Set tmp_Rs = Nothing

Set dgMain.DataSource = rsMain: dgMain.Visible = False
rsMain.MoveFirst

With dgMain
Set dgMain.DataSource = rsMain

    .ColumnHeaders = True        '標題行顯示
    .RowHeight = 300
    .Columns(0).Alignment = dbgCenter
    .Columns(2).Alignment = dbgCenter
    .Columns(9).Alignment = dbgRight
    .Columns(10).Alignment = dbgRight

End With

SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMain.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & dg.Name, dg.Columns(ColIndex).DataField, dg.Columns(ColIndex).Width
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '最小化

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    Frame2.Height = Me.ScaleHeight - Frame1.Height - Frame1.Top - StatusBar.Height
    dgMain.Height = Frame2.Height - 360
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    Frame2.Width = Me.ScaleWidth - 120
    dgMain.Width = Frame2.Width - 240
End If

End Sub

Private Sub cmdReset_Click()

'重設
'txtDeliveryDateS.Text = "": txtDeliveryDateE.Text = ""
txtOrderDateS.Text = "": txtOrderDateE.Text = ""
txtExternOrderkeyS.Text = "": txtExternOrderkeyE.Text = ""

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
Private Sub dgmain_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub cmdExit_Click()
Unload Me '結束此程序
'End 結束應用程式
End Sub

Private Sub Form_Load()
On Error GoTo err_Handle

StatusBar.Panels(2).Text = "0 筆資料列"
StatusBar.Panels(3).Text = User_id

'倉庫別
    '取參數
    Dim objIni As vbIniFile, arrTmp, i As Integer
    Set objIni = New vbIniFile
    objIni.FileName = striniFileName_FullPath
    
    arrTmp = Split(objIni.ReadData("OPTION", "WAREHOUSE", "0"), ";")
    
    cboUserType.AddItem ""
    For i = 0 To UBound(arrTmp)
        cboUserType.AddItem arrTmp(i)
    Next
    cboUserType.ListIndex = 0
    

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txtOrderDateS_Click()

Set objMvdateTarget = txtOrderDateS
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtOrderDateE_Click()

Set objMvdateTarget = txtOrderDateE
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txtOrderDateS_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub txtOrderDateE_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub
