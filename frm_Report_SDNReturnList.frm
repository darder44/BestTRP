VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Report_SDNReturnList 
   Caption         =   "回單檢核表"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14235
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
   Picture         =   "frm_Report_SDNReturnList.frx":0000
   ScaleHeight     =   6300
   ScaleWidth      =   14235
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2400
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
      StartOfWeek     =   97452033
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
      TabIndex        =   6
      Top             =   2280
      Width           =   8295
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   2295
         Left            =   120
         TabIndex        =   9
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
      Height          =   2295
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14175
      Begin VB.CheckBox optSdnback 
         Caption         =   "簽單已回"
         Height          =   255
         Left            =   3480
         TabIndex        =   32
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ListBox List4 
         Columns         =   1
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         ItemData        =   "frm_Report_SDNReturnList.frx":0342
         Left            =   8880
         List            =   "frm_Report_SDNReturnList.frx":0344
         Style           =   1  '項目包含核取方塊
         TabIndex        =   31
         ToolTipText     =   "配送倉別"
         Top             =   240
         Width           =   1575
      End
      Begin VB.ListBox List3 
         Columns         =   3
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         ItemData        =   "frm_Report_SDNReturnList.frx":0346
         Left            =   8880
         List            =   "frm_Report_SDNReturnList.frx":0348
         Style           =   1  '項目包含核取方塊
         TabIndex        =   30
         ToolTipText     =   "訂單類別"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         ItemData        =   "frm_Report_SDNReturnList.frx":034A
         Left            =   6360
         List            =   "frm_Report_SDNReturnList.frx":034C
         Style           =   1  '項目包含核取方塊
         TabIndex        =   29
         ToolTipText     =   "貨運公司"
         Top             =   240
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frm_Report_SDNReturnList.frx":034E
         Left            =   1200
         List            =   "frm_Report_SDNReturnList.frx":035E
         Style           =   2  '單純下拉式
         TabIndex        =   27
         Top             =   1920
         Width           =   2325
      End
      Begin VB.CheckBox optNotYet 
         Caption         =   "未確認簽單"
         Height          =   255
         Left            =   2280
         TabIndex        =   26
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox optAbnormal 
         Caption         =   "異常簽單"
         Height          =   255
         Left            =   1200
         TabIndex        =   25
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox optNormal 
         Caption         =   "正常簽單"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdSaveToText 
         BackColor       =   &H00C0E0FF&
         Caption         =   "檢核表"
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
         Left            =   11760
         Picture         =   "frm_Report_SDNReturnList.frx":03A0
         Style           =   1  '圖片外觀
         TabIndex        =   23
         Top             =   1200
         Width           =   1065
      End
      Begin VB.ListBox List1 
         Columns         =   3
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   4560
         Style           =   1  '項目包含核取方塊
         TabIndex        =   21
         ToolTipText     =   "區碼"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdPreView 
         BackColor       =   &H00C0FFFF&
         Caption         =   "預覽列印"
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
         Left            =   4560
         Picture         =   "frm_Report_SDNReturnList.frx":06AA
         Style           =   1  '圖片外觀
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FF8080&
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
         Left            =   12120
         Picture         =   "frm_Report_SDNReturnList.frx":09B4
         Style           =   1  '圖片外觀
         TabIndex        =   19
         Top             =   1560
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox txtDeliveryDateS 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   16
         Top             =   960
         Width           =   1485
      End
      Begin VB.TextBox txtDeliveryDateE 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   15
         Top             =   960
         Width           =   1485
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   1200
         Style           =   2  '單純下拉式
         TabIndex        =   13
         Top             =   240
         Width           =   1485
      End
      Begin VB.TextBox txtOrderDateE 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   1
         Top             =   600
         Width           =   1485
      End
      Begin VB.TextBox txtOrderDateS 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   0
         Top             =   600
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
         Left            =   11760
         Picture         =   "frm_Report_SDNReturnList.frx":0CBE
         Style           =   1  '圖片外觀
         TabIndex        =   3
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
         Left            =   12960
         Picture         =   "frm_Report_SDNReturnList.frx":1FB8
         Style           =   1  '圖片外觀
         TabIndex        =   5
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
         Left            =   12960
         Picture         =   "frm_Report_SDNReturnList.frx":2BBCA
         Style           =   1  '圖片外觀
         TabIndex        =   4
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
         Left            =   10560
         Picture         =   "frm_Report_SDNReturnList.frx":2BEDC
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "排序"
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
         Left            =   360
         TabIndex        =   28
         Top             =   1980
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "需作完出車確認"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   5
         Left            =   2760
         TabIndex        =   22
         Top             =   240
         Width           =   1680
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
         Index           =   4
         Left            =   2640
         TabIndex        =   18
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
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
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1005
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "貨主"
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
         Left            =   360
         TabIndex        =   14
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "維護日期"
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
         TabIndex        =   12
         Top             =   645
         Width           =   960
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
         TabIndex        =   11
         Top             =   660
         Width           =   360
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '對齊表單下方
      Height          =   270
      Left            =   0
      TabIndex        =   10
      Top             =   6030
      Width           =   14235
      _ExtentX        =   25109
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
            Object.Width           =   18468
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
Attribute VB_Name = "frm_Report_SDNReturnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cmdPreView_Click()

Dim i As Integer, j As Integer
On Error GoTo err_Handle

If rsMain Is Nothing Then MsgBox "無資料可供列印！", vbOKOnly + vbInformation, "報表列印": Exit Sub
Screen.MousePointer = 11

'資料寫入 Access 資料庫
Call AccessDB_Connect
cnAccess.BeginTrans

cnAccess.Execute "Delete From 回單檢核表", RowsAffect, adExecuteNoRecords

Dim rs_Access As New ADODB.Recordset
rs_Access.Open "回單檢核表", cnAccess, adOpenStatic, adLockOptimistic

With rsMain
    .MoveFirst
    Do While Not .EOF
       rs_Access.AddNew
       For i = 0 To .Fields.Count - 1
           rs_Access.Fields(i).Value = .Fields(i).Value
       Next i
       rs_Access.Update
       .MoveNext
    Loop
    .MoveFirst
    cnAccess.CommitTrans
    
    Call DB_Disconnect(cnAccess)
End With

strAccessDBFileName_FullPath = GetAccessDBFileName
Dim MSAccessAP As New access.Application
With MSAccessAP
    .Visible = False
    .OpenCurrentDatabase (strAccessDBFileName_FullPath)
    .DoCmd.Maximize
    
    '寫入USER_ID
    .DoCmd.OpenReport Me.Caption, acViewDesign
    .Reports(Me.Caption).[User_id].Caption = User_id
    .DoCmd.Close

    .DoCmd.OpenReport "回單檢核表", acViewPreview
    .Visible = True

End With

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")

End Sub

Private Sub cmdPrint_Click()
Dim i As Integer, j As Integer
On Error GoTo err_Handle

If rsMain Is Nothing Then MsgBox "無資料可供列印！", vbOKOnly + vbInformation, "報表列印": Exit Sub
Screen.MousePointer = 11

'資料寫入 Access 資料庫
Call AccessDB_Connect
cnAccess.BeginTrans

cnAccess.Execute "Delete From 回單檢核表", RowsAffect, adExecuteNoRecords

Dim rs_Access As New ADODB.Recordset
rs_Access.Open "回單檢核表", cnAccess, adOpenStatic, adLockOptimistic

With rsMain
    .MoveFirst
    Do While Not .EOF
       rs_Access.AddNew
       For i = 0 To .Fields.Count - 1
           rs_Access.Fields(i).Value = .Fields(i).Value
       Next i
       rs_Access.Update
       .MoveNext
    Loop
    .MoveFirst
    cnAccess.CommitTrans
    
    Call DB_Disconnect(cnAccess)
End With

strAccessDBFileName_FullPath = GetAccessDBFileName
Dim MSAccessAP As New access.Application
With MSAccessAP
    .Visible = False
    .OpenCurrentDatabase (strAccessDBFileName_FullPath)
    
    '寫入USER_ID
    .DoCmd.OpenReport Me.Caption, acViewDesign
    .Reports(Me.Caption).[User_id].Caption = User_id
    .DoCmd.Close
    
    '直接列印至印表機
    .Visible = False
    .DoCmd.OpenReport "回單檢核表", acViewNormal
    .CloseCurrentDatabase
    .Quit: Set MSAccessAP = Nothing

End With

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
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

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_Orderdate As String, chc_DeliveryDate As String, i As Integer, strSelected As String

str_SQL = "select 通路別 = rtrim(isnull(t1m.channel,'')) " & _
            ",客戶簡稱 = rtrim(isnull(t1m.short_name,'')) " & _
            ",客戶全名 = isnull(t1m.full_name,'') " & _
            ",到貨日 = rtrim(isnull(s2.arrive_date,'')) " & _
            ",訂單類別 = rtrim(isnull(s2.priority,'')) " & _
            ",訂單號碼 = rtrim(isnull(s2.extern,'')) " & _
            ",驗收單號 = rtrim(isnull(s2.customerorderkey1,'')) " & _
            ",異常狀況 = case when s2.confirm_notes = '正常訂單' then 'N' when len(rtrim(isnull(s2.confirm_notes,''))) = 0 then '未維護' else 'Y' end " & _
            ",發票退回 = s2.invback " & _
            ",預定寄送日 = convert(varchar,getdate(),111) " & _
            ",簽單確認時間 = isnull(convert(char(19),s2.confirm_date,121),'') " & _
            ",二次車號 = rtrim(s1.c_vehicle_id_no) " & _
            ",一次車號 = rtrim(s2.vehicle_id_no) " & _
            "from sdn01t s1 join sdn02t s2 on s1.c_route_no = s2.c_route_no " & _
            "join orders o on o.orderkey = s2.c_receipt_no and o.storerkey = s2.storerkey " & _
            "join trp01m t1m on  t1m.storerkey = o.storerkey and case when rtrim(isnull(s2.priority,'')) = 'A2B' then o.b_company else o.consigneekey end = t1m.consigneekey " & _
            "left join trp09m t9m on s1.c_vehicle_id_no = t9m.vehicle_id_no " & _
            "left join trp08m t8m on t8m.company_code = t9m.trp_company_code Where 1 = 1 "
            
'區碼
For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then strSelected = strSelected & "'" & Left(List1.List(i), 2) & "',"
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & " and t1m.area_code in ( " & strSelected & "'') "

'貨運公司
strSelected = ""
For i = 0 To List2.ListCount - 1
    If List2.Selected(i) Then strSelected = strSelected & "'" & mySplit(List2.List(i), "_", 0) & "',"
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & " and t8m.company_code in ( " & strSelected & "'') "

'單別
strSelected = ""
For i = 0 To List3.ListCount - 1
    If List3.Selected(i) Then strSelected = strSelected & "'" & mySplit(List3.List(i), "_", 0) & "',"
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & " and isnull(s2.priority,'') in ( " & strSelected & "'') "

'配送倉別
strSelected = ""
For i = 0 To List4.ListCount - 1
    If List4.Selected(i) Then strSelected = strSelected & "'" & List4.List(i) & "',"
Next

If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & " and isnull(o.facility,'') in (" & Left(strSelected, Len(strSelected) - 1) & ") "

'維護日期
chc_Orderdate = ""
If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and convert(Char(8),s2.confirm_date,112) between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
   chc_Orderdate = "and convert(Char(8),s2.confirm_date,112) = '" & txtOrderDateS.Text & "' "
ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and convert(Char(8),s2.confirm_date,112) = '" & txtOrderDateE.Text & "' "
End If

'到貨日期
chc_DeliveryDate = ""
If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and convert(Char(8),s2.arrive_date,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
   chc_DeliveryDate = "and convert(Char(8),arrive_date,112) = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and convert(Char(8),arrive_date,112) = '" & txtDeliveryDateE.Text & "' "
End If

'簽單類別
If optNormal = 0 And optAbnormal = 0 And optNotYet = 0 Then GoTo NextStep
Dim strStatus As String

strStatus = "and s2.confirm_notes in ("

If optNormal = 1 Then strStatus = strStatus & "'正常訂單',"
If optAbnormal = 1 Then strStatus = strStatus & "'異常訂單','未出訂單',"
If optNotYet = 1 Then strStatus = strStatus & "'',"

str_SQL = str_SQL & Left(strStatus, Len(strStatus) - 1) & ") "

NextStep:

'簽單已回的資料
If optSdnback = 1 Then str_SQL = str_SQL & " and s2.sdnback = 1 ": MsgBox "此查詢只顯示簽單已回的資料！", vbOKOnly + vbInformation, Me.Caption:

'貨主
If Len(RTrim(Combo1.Text)) > 0 Then str_SQL = str_SQL & chc_Orderdate & chc_DeliveryDate & " and s2.storerkey ='" & Combo1.Text & "' "

If Combo2.Text = "使用者、維護時間" Then
    str_SQL = str_SQL & "order by s2.confirm_userid,isnull(convert(char(19),s2.confirm_date,121),'') "
ElseIf Combo2.Text = "通路別、客戶簡稱" Then
    str_SQL = str_SQL & "order by isnull(t1m.channel,''),isnull(t1m.short_name,'') "
ElseIf Combo2.Text = "訂單號碼" Then
    str_SQL = str_SQL & "order by s2.extern "
Else
    '外務，客戶單號花王用
    str_SQL = str_SQL & "order by isnull(t1m.外務,''),isnull(t1m.consigneekey,'')"
End If

'If Combo1.Text = "LVTL01" Or Combo1.Text = "LNSL01" Or Combo1.Text = "LTHL01" Or Combo1.Text = "LNIP01" Then
'    str_SQL = str_SQL & "order by s2.confirm_userid,isnull(convert(char(19),s2.confirm_date,121),'') "
'Else
'    str_SQL = str_SQL & "order by isnull(t1m.channel,''),isnull(t1m.short_name,'') "
'End If

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = 3
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub

Call Replication_Recordset(tmp_Rs, rsMain)

Set dgMain.DataSource = rsMain: dgMain.Visible = False
rsMain.MoveFirst

With dgMain
Set dgMain.DataSource = rsMain

End With

SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMain.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdSaveToText_Click()
'資料排序
Recordset2Excel "佰事達物流回單檢核表", rsMain

'..在此編輯EXCEL
If rsMain Is Nothing Then
Else
    With MyXlsApp
'        .Columns("L").Select
'        .Selection.ClearContents
        .Range("B3").Value = Combo1
        .Range("A1").Select
        '備份檔案
        '    If Dir("C:\LTKK01\DelievryTrack", vbDirectory) = "" Then MkDirs "C:\LTKK01\DelievryTrack"
        '    .ActiveWorkbook.SaveAs "C:\LTKK01\DelievryTrack\DelievryTrack" & Format(Now, "yyyymmddhhMMss") & ".xls"
                
    End With
End If
Set MyXlsApp = Nothing
    
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

'貨主
Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.CursorLocation = adUseClient
tmp_Rs.Open "select distinct(storerkey) from trp16M", cn, adOpenKeyset, adLockPessimistic

If Not tmp_Rs.EOF Then
    tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        Combo1.AddItem tmp_Rs("storerkey")
        tmp_Rs.MoveNext
    Next
    Combo1.ListIndex = 0
End If
tmp_Rs.Close
    
'區域
With tmp_Rs
    .Open "select area_code from trp03m order by area_code ", cn
    
    If Not .EOF Then
        .MoveFirst
        For i = 0 To .RecordCount - 1
            List1.AddItem RTrim(tmp_Rs("area_code"))
            .MoveNext
        Next
    
    End If
    .Close
    
'貨運公司
    .Open "select company_code,short_name from trp08m order by company_code ", cn
    
If Not .EOF Then
    .MoveFirst
    For i = 0 To .RecordCount - 1
        List2.AddItem RTrim(tmp_Rs("company_code")) & "_" & RTrim(tmp_Rs("short_name"))
        .MoveNext
    Next
End If
.Close

'單別
    .Open "select distinct rtrim(isnull(priority,'')) as Priority from sdn02t order by priority ", cn
    
If Not .EOF Then
    .MoveFirst
    For i = 0 To .RecordCount - 1
        List3.AddItem RTrim(tmp_Rs("Priority"))
        .MoveNext
    Next
End If
.Close

'配送倉別
    .Open "select distinct rtrim(isnull(facility,'')) as facility from Orders order by facility ", cn
    
If Not .EOF Then
    .MoveFirst
    For i = 0 To .RecordCount - 1
        List4.AddItem RTrim(tmp_Rs("facility"))
        .MoveNext
    Next
End If
.Close

End With

Combo2.ListIndex = 0
optNormal = 1
optAbnormal = 1
txtDeliveryDateS = Format(Now - 1, "YYYYMMDD")

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

mvDate.Visible = False

End Sub

Private Sub txtDeliveryDateE_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

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

mvDate.Visible = False

End Sub

Private Sub txtOrderDateE_KeyPress(KeyAscii As Integer)

mvDate.Visible = False

End Sub

Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False

End Sub
