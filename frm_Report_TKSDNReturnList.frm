VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Report_TKSDNReturnList 
   Caption         =   "送貨回單檢核表"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
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
   Picture         =   "frm_Report_TKSDNReturnList.frx":0000
   ScaleHeight     =   6300
   ScaleWidth      =   10170
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3360
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2400
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
      StartOfWeek     =   78512129
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
      Top             =   2160
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
      Height          =   2175
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10095
      Begin VB.CommandButton cmdSaveToText 
         BackColor       =   &H00C0E0FF&
         Caption         =   "轉文字檔"
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
         Left            =   7680
         Picture         =   "frm_Report_TKSDNReturnList.frx":0342
         Style           =   1  '圖片外觀
         TabIndex        =   23
         Top             =   1200
         Visible         =   0   'False
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
         Left            =   4680
         Style           =   1  '項目包含核取方塊
         TabIndex        =   21
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
         Left            =   120
         Picture         =   "frm_Report_TKSDNReturnList.frx":064C
         Style           =   1  '圖片外觀
         TabIndex        =   20
         Top             =   1200
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
         Left            =   1320
         Picture         =   "frm_Report_TKSDNReturnList.frx":0956
         Style           =   1  '圖片外觀
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   1065
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
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   16
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
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   15
         Top             =   960
         Width           =   1485
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   1200
         Style           =   2  '單純下拉式
         TabIndex        =   13
         Top             =   240
         Width           =   1485
      End
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
         MaxLength       =   8
         TabIndex        =   1
         Top             =   600
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
         Left            =   7680
         Picture         =   "frm_Report_TKSDNReturnList.frx":0C60
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
         Left            =   8880
         Picture         =   "frm_Report_TKSDNReturnList.frx":1F5A
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
         Left            =   8880
         Picture         =   "frm_Report_TKSDNReturnList.frx":2BB6C
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
         Left            =   6480
         Picture         =   "frm_Report_TKSDNReturnList.frx":2BE7E
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   240
         Width           =   1065
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
         Left            =   1500
         TabIndex        =   22
         Top             =   1320
         Width           =   1680
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
         Caption         =   "簽單日期"
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
      Width           =   10170
      _ExtentX        =   17939
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
            Object.Width           =   11324
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
Attribute VB_Name = "frm_Report_TKSDNReturnList"
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

cnAccess.Execute "Delete From 送貨回單檢核表", RowsAffect, adExecuteNoRecords

Dim rs_Access As New ADODB.Recordset
rs_Access.Open "送貨回單檢核表", cnAccess, adOpenStatic, adLockOptimistic

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

    .DoCmd.OpenReport "送貨回單檢核表", acViewPreview
    .Visible = True

End With

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")

End Sub

Private Sub cmdPrint_Click()
Dim i As Integer, j As Integer
On Error GoTo err_Handle

If rsMain Is Nothing Then MsgBox "無資料可供列印！", vbOKOnly + vbInformation, "報表列印": Exit Sub
Screen.MousePointer = 11

'資料寫入 Access 資料庫
Call AccessDB_Connect
cnAccess.BeginTrans

cnAccess.Execute "Delete From 送貨回單檢核表", RowsAffect, adExecuteNoRecords

Dim rs_Access As New ADODB.Recordset
rs_Access.Open "送貨回單檢核表", cnAccess, adOpenStatic, adLockOptimistic

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
    .DoCmd.OpenReport "送貨回單檢核表", acViewNormal
    .CloseCurrentDatabase
    .Quit: Set MSAccessAP = Nothing

End With

Screen.MousePointer = 0
Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub cmd2Excel_Click()
'資料排序
Recordset2Excel Me.Caption, rsMain

'..在此編輯EXCEL
If rsMain Is Nothing Then
Else
    With MyXlsApp
        .Columns("J:K").Select
        .Selection.ClearContents
        .Range("A1").Select
        '備份檔案
        '    If Dir("C:\LTKK01\DelievryTrack", vbDirectory) = "" Then MkDirs "C:\LTKK01\DelievryTrack"
        '    .ActiveWorkbook.SaveAs "C:\LTKK01\DelievryTrack\DelievryTrack" & Format(Now, "yyyymmddhhMMss") & ".xls"
                
    End With
End If
Set MyXlsApp = Nothing

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_Orderdate As String, chc_DeliveryDate As String, i As Integer, strSelected As String

For i = 0 To List1.ListCount - 1
    If List1.Selected(i) Then strSelected = strSelected & "'" & Left(List1.List(i), 2) & "',"
Next

str_SQL = "select 通路別 = rtrim(isnull(t1m.channel,'')) " & _
            ",客戶名稱 = rtrim(isnull(t1m.short_name,'')) " & _
            ",地址別中文 = isnull(t1m.full_name,'') " & _
            ",到貨日 = rtrim(isnull(s2.arrive_date,'')) " & _
            ",訂單類別 = rtrim(isnull(s2.priority,'')) " & _
            ",訂單號碼 = rtrim(isnull(s2.extern,'')) " & _
            ",驗收單號 = rtrim(isnull(s2.customerorderkey1,'')) " & _
            ",異常狀況 = case when s2.confirm_notes = '正常訂單' then 'N' else 'Y' end " & _
            ",預定寄送日 = convert(varchar,getdate(),111) " & _
            ",簽單確認時間 = isnull(s2.confirm_date,'') " & _
            "from trp01m t1m right join sdn02t s2 on s2.consigneekey = t1m.consigneekey " & _
            "Where s2.priority <> 'R' "
If Len(RTrim(strSelected)) > 0 Then str_SQL = str_SQL & " and t1m.area_code in ( " & strSelected & "'')"

'簽單日期
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

'組合字串
str_SQL = str_SQL & chc_Orderdate & chc_DeliveryDate & " and s2.storerkey ='" & Combo1.Text & "'order by t1m.channel ,t1m.short_name, t1m.full_name ,s2.arrive_date"

Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_rs.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
'rsMain.Sort = "貨主,訂單號碼,項次"

Call Replication_Recordset(tmp_rs, rsMain)

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
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
End Sub

Private Sub cmdSaveToText_Click()
'資料排序
Recordset2Excel Me.Caption, rsMain

'..在此編輯EXCEL
If rsMain Is Nothing Then
Else
    With MyXlsApp
        .Columns("J:K").Select
        .Selection.ClearContents
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
    Call Confirm_Recordset_Closed(tmp_rs)
    tmp_rs.CursorLocation = adUseClient
    tmp_rs.Open "select distinct(storerkey) from trp16M", cn, adOpenKeyset, adLockPessimistic
    tmp_rs.MoveFirst
    For i = 0 To tmp_rs.RecordCount - 1
        Combo1.AddItem tmp_rs("storerkey")
        tmp_rs.MoveNext
    Next
    tmp_rs.Close: Set tmp_rs = Nothing
'    Combo1.ListIndex = 0
    Combo1.Text = "LTKK01"
    
'區域
Dim rsTmp As New ADODB.Recordset
With rsTmp
    .CursorLocation = 3
    .Open "select area_code from trp03m order by area_code ", cn
    .MoveFirst
    For i = 0 To .RecordCount - 1
        List1.AddItem rsTmp("area_code")
        .MoveNext
    Next
    .Close: Set rsTmp = Nothing
End With

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
End Sub

Private Sub mvDate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
