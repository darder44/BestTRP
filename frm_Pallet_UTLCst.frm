VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Pallet_UTLCst 
   Caption         =   "經銷商棧板管理"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   10155
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   4320
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   4410
      _ExtentX        =   7779
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
      StartOfWeek     =   61472769
      TitleBackColor  =   -2147483646
      TitleForeColor  =   16777215
      TrailingForeColor=   -2147483643
      CurrentDate     =   38233
      MaxDate         =   2958455
   End
   Begin VB.Frame Frame3 
      Caption         =   "功能"
      Height          =   5175
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   1575
      Begin VB.CommandButton cmdPickCancel 
         BackColor       =   &H00FF8080&
         Caption         =   "取消"
         Enabled         =   0   'False
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
         Left            =   240
         Picture         =   "frm_Pallet_UTLCst.frx":0000
         Style           =   1  '圖片外觀
         TabIndex        =   12
         Top             =   4080
         Width           =   1065
      End
      Begin VB.CommandButton cmdPickAddNew 
         BackColor       =   &H00FFFFC0&
         Caption         =   "新增"
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
         Left            =   240
         Picture         =   "frm_Pallet_UTLCst.frx":6852
         Style           =   1  '圖片外觀
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmdPickEdit 
         BackColor       =   &H00FFC0C0&
         Caption         =   "修改"
         Enabled         =   0   'False
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
         Left            =   240
         Picture         =   "frm_Pallet_UTLCst.frx":897C
         Style           =   1  '圖片外觀
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1065
      End
      Begin VB.CommandButton cmdPickDelete 
         BackColor       =   &H00FFC0FF&
         Caption         =   "刪除"
         Enabled         =   0   'False
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
         Left            =   240
         Picture         =   "frm_Pallet_UTLCst.frx":F1CE
         Style           =   1  '圖片外觀
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1065
      End
      Begin VB.CommandButton cmdPickSave 
         BackColor       =   &H00FFFF80&
         Caption         =   "存檔"
         Enabled         =   0   'False
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
         Left            =   240
         Picture         =   "frm_Pallet_UTLCst.frx":10210
         Style           =   1  '圖片外觀
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1065
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5175
      Left            =   1800
      TabIndex        =   18
      Top             =   1320
      Width           =   8295
      Begin VB.ComboBox cboFloatCustomer 
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4920
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid dgMain 
         Height          =   2295
         Left            =   120
         TabIndex        =   7
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
      Caption         =   "查詢"
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   9975
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
         Left            =   8760
         Picture         =   "frm_Pallet_UTLCst.frx":1051A
         Style           =   1  '圖片外觀
         TabIndex        =   6
         Top             =   240
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
         Left            =   7560
         Picture         =   "frm_Pallet_UTLCst.frx":3A12C
         Style           =   1  '圖片外觀
         TabIndex        =   5
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
         Left            =   6360
         Picture         =   "frm_Pallet_UTLCst.frx":3A43E
         Style           =   1  '圖片外觀
         TabIndex        =   4
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox txt2E 
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
         MaxLength       =   8
         TabIndex        =   1
         Top             =   180
         Width           =   1485
      End
      Begin VB.TextBox txt2S 
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
         MaxLength       =   8
         TabIndex        =   0
         Top             =   180
         Width           =   1485
      End
      Begin VB.ComboBox cboCarno 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         ItemData        =   "frm_Pallet_UTLCst.frx":3A748
         Left            =   1200
         List            =   "frm_Pallet_UTLCst.frx":3A74A
         Style           =   2  '單純下拉式
         TabIndex        =   3
         Top             =   900
         Width           =   2085
      End
      Begin VB.ComboBox cboCustomer 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         ItemData        =   "frm_Pallet_UTLCst.frx":3A74C
         Left            =   1200
         List            =   "frm_Pallet_UTLCst.frx":3A74E
         TabIndex        =   2
         Top             =   540
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "簽收日期"
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
         TabIndex        =   17
         Top             =   225
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
         Index           =   23
         Left            =   2655
         TabIndex        =   16
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label3 
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
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label3 
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
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   960
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '對齊表單下方
      Height          =   270
      Left            =   0
      TabIndex        =   22
      Top             =   6495
      Width           =   10155
      _ExtentX        =   17912
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
            Object.Width           =   11298
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
Attribute VB_Name = "frm_Pallet_UTLCst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long

Private Sub cboCarno_GotFocus()
'取出車號
cboCarno.Clear
str_SQL = "select distinct Carno = rtrim(carno) From pallet_utlcst"
Dim rsTmp As New ADODB.Recordset
rsTmp.CursorLocation = 3
rsTmp.Open str_SQL, cn ', adOpenForwardOnly, adLockPessimistic
rsTmp.Sort = "Carno"
If Not rsTmp.EOF Then
        rsTmp.MoveFirst
        Do While Not rsTmp.EOF
            If IsNull(rsTmp("carno")) = False Then cboCarno.AddItem rsTmp("carno")
            rsTmp.MoveNext
        Loop
End If

rsTmp.Close: Set rsTmp = Nothing

End Sub

Private Sub cboFloatCustomer_LostFocus()
cboFloatCustomer.Visible = False
End Sub

Private Sub cmdPickAddNew_Click()
Dim i As Integer
If rsMain.EOF = False Then rsMain.MoveLast
With rsMain
    i = 1
    If .RecordCount > 0 Then .MoveLast: i = .Fields("編號") + 1
    .AddNew
    .Fields("編號") = i
    .Fields("簽收日期") = ""
    .Fields("客戶名稱") = ""
    .Fields("單號") = ""
    .Fields("車號") = ""
    .Fields("借入") = "0"
    .Fields("還回") = "0"
End With

dgMain.AllowUpdate = True
cmdPickSave.Enabled = True: cmdPickCancel.Enabled = True
cmdPickDelete.Enabled = False: cmdPickEdit.Enabled = False: cmdPickAddNew.Enabled = False
dgMain.Col = 1: dgMain.SetFocus
intPickRow = dgMain.Row
intLastCol = dgMain.Col

End Sub
Private Sub cmdPickEdit_Click()

If Len(rsMain("checkuser")) > 0 Then MsgBox "已確認資料無法修改!!", vbInformation: Exit Sub

dgMain.AllowUpdate = True
cmdPickSave.Enabled = True: cmdPickCancel.Enabled = True
cmdPickDelete.Enabled = False: cmdPickEdit.Enabled = False: cmdPickAddNew.Enabled = False
dgMain.Col = 1: dgMain.SetFocus
intPickRow = dgMain.Row
intLastCol = dgMain.Col

End Sub
Private Sub cmdPickDelete_Click()
On Error GoTo err_Handle
Dim confirm As Integer

If rsMain.BOF Then cmdPickDelete.Enabled = False: cmdPickEdit.Enabled = False: Exit Sub
If Len(rsMain("checkuser")) > 0 Then MsgBox "已確認資料無法刪除!!", vbInformation, Me.Caption: Exit Sub
confirm = MsgBox("確定刪除?", vbQuestion + vbOKCancel, Me.Caption)
If confirm <> 1 Then Exit Sub

str_SQL = "delete from pallet_utlcst where keyid = '" & rsMain("keyid") & "' "
cn.BeginTrans
    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
cn.CommitTrans

'更新dgmain資料
rsMain.Delete: If rsMain.EOF Then rsMain.MovePrevious
If rsMain.RecordCount = 0 Then cmdPickDelete.Enabled = False: cmdPickEdit.Enabled = False
cmdPickAddNew.SetFocus

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdPickSave_Click()
On Error GoTo err_Handle

'If myIsDate(Trim(rsMain("簽收日期"))) = False Then MsgBox "請檢查日期格式!!", vbOKOnly, Me.Caption: dgMain.SetFocus: Exit Sub
If Len(RTrim(rsMain("簽收日期") & "")) = 0 Then MsgBox "請輸入簽收日期!!", vbOKOnly + vbInformation, Me.Caption: dgMain.Col = 1: dgMain.SetFocus: Exit Sub
If myIsDate(rsMain("簽收日期") & "") = False Then Exit Sub
If Len(RTrim(rsMain("客戶名稱") & "")) = 0 Then MsgBox "請輸入客戶名稱!!", vbOKOnly + vbInformation, Me.Caption: dgMain.Col = 2: dgMain.SetFocus: Exit Sub
If Len(RTrim(rsMain("車號") & "")) = 0 Then MsgBox "請輸入車號!!", vbOKOnly + vbInformation, Me.Caption: dgMain.Col = 4: dgMain.SetFocus: Exit Sub
If Val(Trim(rsMain("借入"))) + Val(Trim(rsMain("還回"))) = 0 Then MsgBox "請確認數量!!", vbOKOnly + vbInformation, Me.Caption: dgMain.Col = 5: dgMain.SetFocus: Exit Sub

'檢查是否重複
Dim rsTmp1 As New ADODB.Recordset
With rsTmp1
    .CursorLocation = adUseClient
    str_SQL = "select * from pallet_utlcst where keyid = '" & rsMain("keyid") & "' "
    .Open str_SQL, cn, adOpenStatic, adLockOptimistic
        
    If .EOF Then
    
    Dim rsTmp As New ADODB.Recordset, keyid As String
    rsTmp.Open "select keyid = isnull(max(keyid),0) from pallet_utlcst", cn
    keyid = Format(Val(rsTmp("keyid")) + 1, "0000000000")
    rsTmp.Close: Set rsTmp = Nothing
    
        If rsMain("借入") <> 0 And rsMain("還回") <> 0 Then
        Dim intIn As Integer, intOut As Integer
        intIn = rsMain("借入"): intOut = rsMain("還回")
            '新增資料庫資料
            .AddNew
            .Fields("keyid") = keyid
            .Fields("Storer") = "UTL"
            .Fields("chargedate") = Trim(rsMain("簽收日期"))
            .Fields("customer") = Trim(rsMain("客戶名稱"))
            .Fields("customersheetno") = Trim(rsMain("單號"))
            .Fields("carno") = UCase(rsMain("車號"))
            .Fields("qtyin") = intIn
            .Fields("qtyout") = 0
            .Fields("notes") = Trim(rsMain.Fields("備註"))
            .Fields("Adduser") = User_id
            .Fields("Adddate") = Now()
            .Update
            
            '更新dgmain
            rsMain("簽收日期") = .Fields("chargedate")
            rsMain.Fields("還回") = .Fields("qtyout")
            rsMain.Fields("keyid") = keyid
            rsMain.Fields("Adduser") = User_id
            rsMain.Fields("Adddate") = Format(Now(), "yyyy/mm/dd hh:MM:ss")
            rsMain.Update
            
            With rsMain
            Dim i As Integer
                i = 1
                If .RecordCount > 0 Then .MoveLast: i = .Fields("編號") + 1
                .AddNew
                .Fields("編號") = i
                .Fields("簽收日期") = rsTmp1("chargedate")
                .Fields("客戶名稱") = rsTmp1("customer")
                .Fields("單號") = rsTmp1("customersheetno")
                .Fields("車號") = rsTmp1("carno")
                .Fields("借入") = 0
                .Fields("還回") = intOut
                .Fields("備註") = rsTmp1.Fields("notes")
                .Fields("keyid") = Format(Val(keyid) + 1, "0000000000")
                .Fields("Adduser") = User_id
                .Fields("Adddate") = Format(Now(), "yyyy/mm/dd hh:MM:ss")
            End With

            .AddNew
            .Fields("keyid") = rsMain("keyid")
            .Fields("Storer") = "UTL"
            .Fields("chargedate") = rsMain("簽收日期")
            .Fields("customer") = rsMain("客戶名稱")
            .Fields("customersheetno") = rsMain("單號")
            .Fields("carno") = UCase(rsMain("車號"))
            .Fields("qtyin") = rsMain("借入")
            .Fields("qtyout") = rsMain("還回")
            .Fields("notes") = rsMain.Fields("備註")
            .Fields("Adduser") = User_id
            .Fields("Adddate") = Now()
            .Update
            .MoveLast
            cmdPickEdit.Enabled = False: cmdPickDelete.Enabled = False
             
        Else
    
        '新增資料庫資料
            .AddNew
            .Fields("keyid") = keyid
            .Fields("Storer") = "UTL"
            .Fields("chargedate") = rsMain("簽收日期")
            .Fields("customer") = rsMain("客戶名稱")
            .Fields("customersheetno") = rsMain("單號")
            .Fields("carno") = UCase(rsMain("車號"))
            .Fields("qtyin") = rsMain("借入")
            .Fields("qtyout") = rsMain("還回")
            .Fields("notes") = rsMain.Fields("備註")
            .Fields("Adduser") = User_id
            .Fields("Adddate") = Now()
            .Update
            
            '更新dgmain
            rsMain.Fields("keyid") = keyid
            rsMain.Fields("Adduser") = User_id
            rsMain.Fields("Adddate") = Format(Now(), "yyyy/mm/dd hh:MM:ss")
            rsMain.Update
            dgMain.Row = rsMain.RecordCount - 1
            
        End If
    Else

        '修改資料
            .Fields("chargedate") = Trim(rsMain("簽收日期"))
            .Fields("customer") = Trim(rsMain("客戶名稱"))
            .Fields("customersheetno") = Trim(rsMain("單號"))
            .Fields("carno") = UCase(rsMain("車號"))
            .Fields("qtyin") = rsMain("借入")
            .Fields("qtyout") = rsMain("還回")
            .Fields("notes") = Trim(rsMain.Fields("備註"))
            .Fields("Edituser") = User_id
            .Fields("Editdate") = Now()
            .Update
            
            '更新dgmain
            rsMain.Fields("Edituser") = User_id
            rsMain.Fields("Editdate") = Format(Now(), "yyyy/mm/dd hh:MM:ss")
            rsMain.Update
            
    End If
    rsTmp1.Close: Set rsTmp1 = Nothing
End With

cmdPickAddNew.Enabled = True: cmdPickEdit.Enabled = True: cmdPickDelete.Enabled = True: dgMain.AllowUpdate = False: cmdPickSave.Enabled = False: cmdPickCancel.Enabled = False
'Call Update
cmdPickAddNew.SetFocus
dgMain.AllowUpdate = False

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub
Private Sub cmdPickCancel_Click()

cmdPickSave.Enabled = False: cmdPickCancel.Enabled = False
cmdPickAddNew.Enabled = True
If rsMain.RecordCount > 0 Then cmdPickEdit.Enabled = True: cmdPickDelete.Enabled = True
cmdPickAddNew.SetFocus
dgMain.AllowUpdate = False
If Len(RTrim(rsMain("Keyid"))) = 0 Then rsMain.Delete
End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
'If myIsDate(txt2S.Text) = False Or myIsDate(txt2E.Text) = False Then Exit Sub
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_Chargedate As String, chc_Carno As String, chc_Customer As String

'取出客戶棧板資料
str_SQL = "select 簽收日期 = rtrim(chargedate) " & _
          ", 客戶名稱 = rtrim(customer) " & _
          ", 單號 = rtrim(customersheetno) " & _
          ", 車號 = rtrim(carno) " & _
          ", 借入= rtrim(qtyin) " & _
          ", 還回 = rtrim(qtyout) " & _
          ", 備註 = rtrim(notes) " & _
          ", AddUser = rtrim(adduser) " & _
          ", Adddate = rtrim(convert( char(20) , adddate , 120 )) " & _
          ", CheckUser = rtrim(CheckUser) " & _
          ", Checkdate = rtrim(convert( char(20) , Checkdate , 120 )) " & _
          ", EditUser = rtrim(EditUser) " & _
          ", Editdate = rtrim(convert( char(20) , Editdate , 120 )) " & _
          ", KeyID " & _
          "from pallet_UTLcst "

'客戶名稱
chc_Customer = ""
If Len(cboCustomer.Text) > 0 Then chc_Customer = "and Customer = '" & cboCustomer.Text & "' "

'車號
chc_Carno = ""
If Len(cboCarno.Text) > 0 Then chc_Carno = "and carno = '" & cboCarno.Text & "' "

'排出日期
chc_Chargedate = ""
If Len(txt2S.Text) > 0 And Len(txt2E.Text) > 0 Then
   chc_Chargedate = "and Chargedate between '" & txt2S.Text & "' and '" & txt2E.Text & "' "
ElseIf Len(txt2S.Text) > 0 And Len(txt2E.Text) = 0 Then
   chc_Chargedate = "and Chargedate = '" & txt2S.Text & "' "
ElseIf Len(txt2S.Text) = 0 And Len(txt2E.Text) > 0 Then
   chc_Chargedate = "and Chargedate = '" & txt2E.Text & "' "
End If

'組合字串
If Len(chc_Chargedate & chc_Carno & chc_Customer) = 0 Then MsgBox "請至少指定一項查詢條件!!", vbOKOnly, Me.Caption: Screen.MousePointer = 0: Exit Sub
str_SQL = str_SQL & "where 1 = 1 " & chc_Chargedate & chc_Carno & chc_Customer

Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.CursorLocation = adUseClient
tmp_rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_rs.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
tmp_rs.Sort = "簽收日期,車號"
Call Replication_Recordset(tmp_rs, rsMain)
tmp_rs.Close: Set tmp_rs = Nothing
rsMain.MoveFirst

Set dgMain.DataSource = rsMain: dgMain.Visible = False

With dgMain

    .ColumnHeaders = True        '標題行顯示
    .RowHeight = 300
    .Columns(0).Alignment = dbgCenter
    .Columns(1).Alignment = dbgCenter
    .Columns(5).Alignment = dbgRight
    .Columns(6).Alignment = dbgRight

End With
cmdPickEdit.Enabled = True: cmdPickDelete.Enabled = True
SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMain.Visible = True
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub dgMain_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo err_Handle

cboFloatCustomer.Visible = False
mvDate.Visible = False
'If dgSku.Row = -1 Then Exit Sub

'新增狀態下無法變更資料列
If cmdPickSave.Enabled = True And LastRow <> Empty Then
    dgMain.Col = intLastCol
    dgMain.Row = intPickRow
    
    Exit Sub
End If

'是否顯示日期
If dgMain.Col = 1 And cmdPickSave.Enabled = True Then
    Set objMvdateTarget = dgMain: mvDate.Visible = True: mvDate.Value = Now()
    mvDate.Move dgMain.Columns(dgMain.Col).Left + dgMain.Columns(dgMain.Col).Width + dgMain.Left + Frame2.Left, dgMain.RowTop(dgMain.Row) + dgMain.Top + Frame2.Top
End If

'不允許移至特定欄位
If dgMain.Col = 0 Or dgMain.Col > 7 Then dgMain.Col = Abs(LastCol): Exit Sub
'If dgMain.Col = 4 Then
'    If LastCol = 3 Then dgMain.Col = 5: Exit Sub
'    If LastCol = 5 Then dgMain.Col = 2: Exit Sub
'    dgMain.Col = IIf(LastCol = -1, 5, LastCol)
'End If
'是否顯示客戶選單
If dgMain.Col = 2 And cmdPickSave.Enabled = True Then ShowList
'資料列是否變更
If LastRow <> Empty Then cmdPickEdit.Enabled = True: cmdPickDelete.Enabled = True

Screen.MousePointer = 0

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub Form_Resize()

If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub '最小化

If Me.ScaleHeight > Frame1.Top + Frame1.Height + 500 Then
    Frame2.Height = Me.ScaleHeight - Frame1.Height - Frame1.Top - StatusBar.Height
    dgMain.Height = Frame2.Height - 360
End If

If Me.ScaleWidth > Frame1.Width + Frame1.Left Then
    Frame2.Width = Me.ScaleWidth - Frame3.Width - Frame3.Left - 120
    dgMain.Width = Frame2.Width - 240
End If

End Sub

Private Sub cmdReset_Click()

'重設
txt2S.Text = "": txt2E.Text = ""
cboCustomer.ListIndex = -1
cboCarno.ListIndex = -1

End Sub

Private Sub dgMain_HeadClick(ByVal ColIndex As Integer)

If dgMain.Row = -1 Or cmdPickSave.Enabled = True Then Exit Sub
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
If KeyAscii = 13 And dgMain.Col = 7 And cmdPickSave.Enabled = True Then
    Dim ok
    ok = MsgBox("是否存檔?", vbYesNo, Me.Caption)
    If ok = 6 Then Call cmdPickSave_Click: Exit Sub
End If

If KeyAscii = 13 Then SendKeys "{tab}"
If KeyAscii = 27 And mvDate.Visible = True Then mvDate.Visible = False

End Sub
Private Sub cboFloatCustomer_Click()

dgMain.Text = cboFloatCustomer.Text

End Sub
Private Sub ShowList()

With dgMain
.RowHeight = cboFloatCustomer.Height - 10
If .Col = 2 Then
    If .Columns(.Col).Left > 0 Then
            cboFloatCustomer.Visible = True
            cboFloatCustomer.Move .Left + .Columns(.Col).Left + 15, .Top + .RowTop(.Row), .Columns(.Col).Width
            If cboFloatCustomer.Left + cboFloatCustomer.Width > .Left + .Width Then '如果欄位超出DataGrid的顯示範圍的處理
                cboFloatCustomer.Width = cboFloatCustomer.Width + .Left + .Width - cboFloatCustomer.Left - cboFloatCustomer.Width
            End If
            cboFloatCustomer.Text = RTrim(dgMain.Text)  '更新Combo的值
            cboFloatCustomer.SetFocus
    Else '如果用捲軸捲動出了DataGrid的顯示範圍，值會小於0
        cboFloatCustomer.Visible = False
    End If
Else
    cboFloatCustomer.Visible = False
End If
End With
End Sub
Private Sub dgMain_Scroll(Cancel As Integer)
ShowList
End Sub
Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
ShowList
'無資料或欄寬太小，不存寬度
If Len(dgMain.Columns(ColIndex).DataField) < 0 Or dgMain.Columns(ColIndex).Width < 50 Then Exit Sub
SaveSetting App.title, Me.Caption & "dgMain", dgMain.Columns(ColIndex).DataField, dgMain.Columns(ColIndex).Width
End Sub
Private Sub dgMain_RowResize(Cancel As Integer)
ShowList
End Sub
Private Sub cmdExit_Click()
Unload Me '結束此程序
'End 結束應用程式
End Sub

Private Sub Form_Load()
On Error GoTo err_Handle

'取出客戶名稱
Call Confirm_Recordset_Closed(tmp_rs)
str_SQL = "select code from CodeLkup where listname='Cust_CDS'"
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
If Not tmp_rs.EOF Then
   Do While Not tmp_rs.EOF
      cboCustomer.AddItem Trim(tmp_rs.Fields("code"))
      cboFloatCustomer.AddItem Trim(tmp_rs.Fields("code"))
      tmp_rs.MoveNext
   Loop
End If
tmp_rs.Close

'建立dgMain格式
str_SQL = "select 簽收日期 = rtrim(chargedate) " & _
          ", 客戶名稱 = rtrim(customer) " & _
          ", 單號 = rtrim(customersheetno) " & _
          ", 車號 = rtrim(carno) " & _
          ", 借入= rtrim(qtyin) " & _
          ", 還回 = rtrim(qtyout) " & _
          ", 備註 = rtrim(notes) " & _
          ", AddUser = rtrim(adduser) " & _
          ", Adddate = rtrim(convert( char(20) , adddate , 120 )) " & _
          ", CheckUser = rtrim(CheckUser) " & _
          ", Checkdate = rtrim(convert( char(20) , Checkdate , 120 )) " & _
          ", EditUser = rtrim(EditUser) " & _
          ", Editdate = rtrim(convert( char(20) , Editdate , 120 )) " & _
          ", KeyID " & _
          "from pallet_UTLcst where 1 = 2"
          
Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.CursorLocation = adUseClient
tmp_rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic

Call Replication_Recordset(tmp_rs, rsMain)
tmp_rs.Close: Set tmp_rs = Nothing
Set dgMain.DataSource = rsMain

With dgMain
Set dgMain.DataSource = rsMain
    .ColumnHeaders = True        '標題行顯示
    .RowHeight = 300
    .Columns(0).Width = 600:       .Columns(0).Alignment = dbgCenter
    .Columns(1).Width = 1000:       .Columns(1).Alignment = dbgCenter
    .Columns(2).Width = 1500:    .Columns(2).Alignment = dbgLeft
    .Columns(3).Width = 1000:    .Columns(3).Alignment = dbgLeft
    .Columns(4).Width = 1000:    .Columns(4).Alignment = dbgLeft
    .Columns(5).Width = 600:    .Columns(5).Alignment = dbgRight
    .Columns(6).Width = 600:    .Columns(6).Alignment = dbgRight
    .Columns(7).Width = 2000:    .Columns(7).Alignment = dbgLeft
    .Columns(8).Width = 1000:    .Columns(8).Alignment = dbgLeft
    .Columns(9).Width = 1500:    .Columns(9).Alignment = dbgLeft
    .Columns(10).Width = 1000:    .Columns(10).Alignment = dbgLeft
    .Columns(11).Width = 1500:   .Columns(11).Alignment = dbgLeft
    .Columns(12).Width = 1000:    .Columns(12).Alignment = dbgLeft
    .Columns(13).Width = 1500:    .Columns(13).Alignment = dbgLeft
    .Columns(14).Width = 1000:    .Columns(14).Alignment = dbgLeft
End With

cboCustomer.ListIndex = -1: cboFloatCustomer.ListIndex = -1
StatusBar.Panels(2).Text = "0 筆資料列"
StatusBar.Panels(3).Text = User_id

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")

End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = vbRightButton Then mvDate.Visible = False
End Sub

Private Sub txt2S_Click()
Set objMvdateTarget = txt2S
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub txt2E_Click()
Set objMvdateTarget = txt2E
mvDate.Move objMvdateTarget.Left + objMvdateTarget.Width, objMvdateTarget.Top + objMvdateTarget.Height
mvDate.Visible = True: mvDate.Value = Now
    
End Sub
Private Sub mvDate_DateClick(ByVal DateClicked As Date)

objMvdateTarget.Text = Format(mvDate.Value, "yyyymmdd")
mvDate.Visible = False
dgMain.SetFocus ': dgMain.Col = dgMain.Col + 1

End Sub

Private Sub txt2S_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then mvDate.Visible = False
End Sub
Private Sub txt2E_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then mvDate.Visible = False
End Sub
