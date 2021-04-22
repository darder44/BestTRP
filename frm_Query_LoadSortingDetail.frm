VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frm_Query_LoadSortingDetail 
   Caption         =   "翻板理貨簽收單"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
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
   ScaleWidth      =   8400
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3240
      TabIndex        =   15
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
      StartOfWeek     =   230227969
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
      TabIndex        =   13
      Top             =   2160
      Width           =   8295
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   2295
         Left            =   120
         TabIndex        =   16
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
      TabIndex        =   14
      Top             =   0
      Width           =   8295
      Begin VB.CheckBox chkNSL 
         Caption         =   "僅雀巢"
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
         Height          =   240
         Left            =   4320
         TabIndex        =   25
         Top             =   1440
         Width           =   1425
      End
      Begin VB.CheckBox chkPreView 
         Caption         =   "預覽列印"
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
         Height          =   240
         Left            =   4320
         TabIndex        =   10
         Top             =   1800
         Value           =   1  '核取
         Width           =   1425
      End
      Begin VB.CommandButton cmdPrintReport 
         BackColor       =   &H00C0FFC0&
         Caption         =   "報表列印"
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
         Picture         =   "frm_Query_LoadSortingDetail.frx":0000
         Style           =   1  '圖片外觀
         TabIndex        =   11
         Top             =   1200
         Width           =   1065
      End
      Begin VB.TextBox txtCheckKeyE 
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
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1485
      End
      Begin VB.TextBox txtCheckKeyS 
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
         Left            =   720
         MaxLength       =   15
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1485
      End
      Begin VB.ComboBox cboCustomer 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   3645
      End
      Begin VB.ComboBox cboCarno 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2880
         TabIndex        =   1
         Top             =   240
         Width           =   1485
      End
      Begin VB.TextBox txtDeliveryDateS 
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
         Left            =   720
         MaxLength       =   8
         TabIndex        =   3
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtDeliveryDateE 
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
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   4
         Top             =   960
         Width           =   1485
      End
      Begin VB.ComboBox cboUserType 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   1485
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
         Picture         =   "frm_Query_LoadSortingDetail.frx":030A
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
         Picture         =   "frm_Query_LoadSortingDetail.frx":1604
         Style           =   1  '圖片外觀
         TabIndex        =   12
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
         Picture         =   "frm_Query_LoadSortingDetail.frx":2B216
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
         Picture         =   "frm_Query_LoadSortingDetail.frx":2B528
         Style           =   1  '圖片外觀
         TabIndex        =   7
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "∼"
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
         Index           =   6
         Left            =   2160
         TabIndex        =   24
         Top             =   1380
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
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
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   1365
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "客戶"
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
         Left            =   120
         TabIndex        =   22
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "車號"
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
         Left            =   2280
         TabIndex        =   21
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "∼"
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
         Left            =   2160
         TabIndex        =   20
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "日期"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1005
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "類別"
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
         Top             =   300
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '對齊表單下方
      Height          =   270
      Left            =   0
      TabIndex        =   17
      Top             =   6030
      Width           =   8400
      _ExtentX        =   14817
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
            Object.Width           =   8202
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
Attribute VB_Name = "frm_Query_LoadSortingDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cmdPrintReport_Click()
Dim i As Integer, j As Integer, strTmp As String
Dim MSAccessAP As access.Application

'str_sql = "select 日期 " & _
',單號" & _
',客戶編號" & _
',客戶名稱" & _
',翻板 = sum(翻板)" & _
',理貨 = sum(理貨)" & _
',貼標 = sum(貼標)" & _
',蓋章 = sum(貼標)" & _
' From gv_loadsorting " & _
'group by 日期,單號,客戶編號,客戶名稱 "

'報表列印
If rsMain Is Nothing Then MsgBox "無資料可供列印！", vbOKOnly + vbInformation, "報表列印": Exit Sub

On Error GoTo err_Handle

'資料寫出 Access 資料庫
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From 翻板理貨簽收單"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Dim rs_Access As New ADODB.Recordset
rs_Access.Open "翻板理貨簽收單", cnAccess, adOpenStatic, adLockOptimistic
rsMain.MoveFirst
Do While Not rsMain.EOF

   rs_Access.AddNew
   For i = 0 To rsMain.Fields.Count - 1
         rs_Access.Fields(i).Value = RTrim(rsMain.Fields(i).Value)
   Next i
   rs_Access.Update
   rsMain.MoveNext
Loop
rsMain.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'寫入User_ID
With MSAccessAP
    .DoCmd.OpenReport "翻板理貨簽收單", acViewDesign
    .Reports("翻板理貨簽收單").[User_id].Caption = User_id
    .DoCmd.Close 'acReport, Me.Caption, acViewDesign
End With

'[報表列印] 命令鈕 -- 利用 Access 報表
If chkPreView = 1 Then
   '預覽列印
    MSAccessAP.Visible = True
    MSAccessAP.DoCmd.OpenReport "翻板理貨簽收單", acViewPreview
    MSAccessAP.DoCmd.Maximize
   
Else
   '直接列印至印表機
    MSAccessAP.Visible = False
    MSAccessAP.DoCmd.OpenReport "翻板理貨簽收單", acViewNormal
    MSAccessAP.CloseCurrentDatabase
    MSAccessAP.Quit: Set MSAccessAP = Nothing
End If

Exit Sub

err_Handle:
   If Tran_Level <> 0 Then cnAccess.RollbackTrans
   Tran_Level = 0
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
   End If
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-翻板理貨簽收單-列印", Me.Caption, "cmdPrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmd2Excel_Click()

'資料排序
Recordset2Excel Me.Caption, rsMain

'..在此編輯EXCEL
With MyXlsApp

  
End With
Set MyXlsApp = Nothing

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_Orderdate As String, chc_DeliveryDate As String, chc_Storerkey As String
    
str_SQL = "select * from gv_loadsorting" & " where 1 = 1 "
    
'日期
If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
   str_SQL = str_SQL & "and 日期 between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
   str_SQL = str_SQL & "and 日期 = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
   str_SQL = str_SQL & "and 日期 = '" & txtDeliveryDateE.Text & "' "
End If

'單號
If Len(txtCheckKeyS.Text) > 0 And Len(txtCheckKeyE.Text) > 0 Then
   str_SQL = str_SQL & "and 單號 between '" & txtCheckKeyS.Text & "' and '" & txtCheckKeyE.Text & "' "
ElseIf Len(txtCheckKeyS.Text) > 0 And Len(txtCheckKeyE.Text) = 0 Then
   str_SQL = str_SQL & "and 單號 = '" & txtCheckKeyS.Text & "' "
ElseIf Len(txtCheckKeyS.Text) = 0 And Len(txtCheckKeyE.Text) > 0 Then
   str_SQL = str_SQL & "and 單號 = '" & txtCheckKeyE.Text & "' "
End If

''類別
'If Len(RTrim(cboUserType.Text)) > 0 Then str_SQL = str_SQL & " and usertype ='" & cboUserType.Text & "' "
'客戶
If Len(RTrim(cboCustomer.Text)) > 0 Then str_SQL = str_SQL & " and 客戶名稱 ='" & cboCustomer.Text & "' "
'車號
If Len(RTrim(cboCarno.Text)) > 0 Then str_SQL = str_SQL & " and 車號 ='" & cboCarno.Text & "' "
'雀巢
If chkNSL = 1 Then
    
    str_SQL = str_SQL & " and 貨主 = 'LNSL01' "
Else
    str_SQL = str_SQL & " and 貨主 <> 'LNSL01' "
    
End If

Set rsMain = New ADODB.Recordset
rsMain.CursorLocation = adUseClient
rsMain.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMain.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
rsMain.Sort = "日期,單號,項次"

rsMain.MoveFirst

Set dgMain.DataSource = rsMain: dgMain.Visible = False

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
Call ClearForm_AllField(Me)

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
Dim i As Integer
StatusBar.Panels(2).Text = "0 筆資料列"
StatusBar.Panels(3).Text = User_id

''類別
'Call Confirm_Recordset_Closed(tmp_rs)
'tmp_rs.CursorLocation = adUseClient
'tmp_rs.Open "select distinct(usertype) from pallet_cst order by usertype", cn, adOpenKeyset, adLockPessimistic
'tmp_rs.MoveFirst
'For i = 0 To tmp_rs.RecordCount - 1
'    cboUserType.AddItem RTrim(tmp_rs("usertype"))
'    tmp_rs.MoveNext
'Next
'tmp_rs.Close
'cboUserType.ListIndex = -1

'客戶
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct (company) from gt_loadsorting order by company ", cn, adOpenKeyset, adLockPessimistic
tmp_Rs.MoveFirst
For i = 0 To tmp_Rs.RecordCount - 1
    cboCustomer.AddItem RTrim(tmp_Rs("company"))
    tmp_Rs.MoveNext
Next
tmp_Rs.Close
cboCustomer.ListIndex = -1

'車號
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(carno) from gt_loadsorting order by carno ", cn, adOpenKeyset, adLockPessimistic
tmp_Rs.MoveFirst
For i = 0 To tmp_Rs.RecordCount - 1
    cboCarno.AddItem RTrim(tmp_Rs("carno"))
    tmp_Rs.MoveNext
Next
tmp_Rs.Close
cboCarno.ListIndex = 0

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txtDeliveryDateS_Click()
Set objMvdateTarget = txtDeliveryDateS
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
End Sub
Private Sub txtDeliveryDateE_Click()
Set objMvdateTarget = txtDeliveryDateE
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
End Sub
Private Sub txtDeliveryDateS_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateE_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then mvDate.Visible = False

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub
