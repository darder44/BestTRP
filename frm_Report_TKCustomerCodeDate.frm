VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_Report_TKCustomerCodeDate 
   Caption         =   "客戶進貨有效期限明細表"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
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
   Picture         =   "frm_Report_TKCustomerCodeDate.frx":0000
   ScaleHeight     =   6300
   ScaleWidth      =   10335
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   2760
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3240
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
      StartOfWeek     =   135593985
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
      Width           =   8295
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
         Left            =   5880
         Picture         =   "frm_Report_TKCustomerCodeDate.frx":0342
         Style           =   1  '圖片外觀
         TabIndex        =   20
         Top             =   240
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
         Left            =   5880
         Picture         =   "frm_Report_TKCustomerCodeDate.frx":064C
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
         Left            =   4680
         Picture         =   "frm_Report_TKCustomerCodeDate.frx":0956
         Style           =   1  '圖片外觀
         TabIndex        =   3
         Top             =   1200
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
         Picture         =   "frm_Report_TKCustomerCodeDate.frx":1C50
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
         Left            =   7080
         Picture         =   "frm_Report_TKCustomerCodeDate.frx":2B862
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
         Left            =   4680
         Picture         =   "frm_Report_TKCustomerCodeDate.frx":2BB74
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "需作完簽單維護"
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
         Left            =   2040
         TabIndex        =   21
         Top             =   1320
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
         Caption         =   "訂單日期"
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
      Width           =   10335
      _ExtentX        =   18230
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
            Object.Width           =   11589
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
Attribute VB_Name = "frm_Report_TKCustomerCodeDate"
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

cnAccess.Execute "Delete From 店家進貨有效期限明細表", RowsAffect, adExecuteNoRecords

Dim rs_Access As New ADODB.Recordset
rs_Access.Open "店家進貨有效期限明細表", cnAccess, adOpenStatic, adLockOptimistic

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

    .DoCmd.OpenReport "店家進貨有效期限明細表", acViewPreview
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

cnAccess.Execute "Delete From 店家進貨有效期限明細表", RowsAffect, adExecuteNoRecords

Dim rs_Access As New ADODB.Recordset
rs_Access.Open "店家進貨有效期限明細表", cnAccess, adOpenStatic, adLockOptimistic

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
    .DoCmd.OpenReport "店家進貨有效期限明細表", acViewNormal
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
Recordset2Excel "客戶進貨", rsMain

'..在此編輯EXCEL
With MyXlsApp

    .Range("Z:AD").NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    '備份檔案
    If Dir("C:\LTKK01\客戶進貨", vbDirectory) = "" Then MkDirs "C:\LTKK01\客戶進貨"
    .ActiveWorkbook.SaveAs "C:\LTKK01\客戶進貨\客戶進貨_" & Format(Now, "yyyymmddhhMMss") & ".xls"
    
End With

Set MyXlsApp = Nothing

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_Orderdate As String, chc_DeliveryDate As String

str_SQL = "select 貨主 ,店家代碼,客戶名稱,客戶單號,料號,儲位,訂單箱數,訂單大單位,訂單組數,訂單中單位,訂單個數,訂單小單位,出貨箱數,出貨大單位,出貨組數,出貨中單位,出貨個數,出貨小單位 " & _
",驗收箱數,驗收大單位,驗收組數,驗收中單位,驗收個數,驗收小單位,製造日,到期日,訂單日期,到貨日期,寄驗收單日 " & _
",店家驗收單號,店家驗收SCAN,有無異常 from gv_DeliveryCodeDate where 1 = 1 "

'訂單日期
chc_Orderdate = ""
If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "replace(訂單日期,'/','') between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
   chc_Orderdate = "and replace(訂單日期,'/','') = '" & txtOrderDateS.Text & "' "
ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and replace(訂單日期,'/','') = '" & txtOrderDateE.Text & "' "
End If

'到貨日期
chc_DeliveryDate = ""
If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and replace(到貨日期,'/','') between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
   chc_DeliveryDate = "and replace(到貨日期,'/','') = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and replace(到貨日期,'/','') = '" & txtDeliveryDateE.Text & "' "
End If

'組合字串
str_SQL = str_SQL & chc_Orderdate & chc_DeliveryDate & " and 貨主 ='" & Combo1.Text & "' order by 店家代碼 ,到貨日期 ,料號 "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
'rsMain.Sort = "貨主,訂單號碼,項次"

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

If rsMain Is Nothing Then Exit Sub
On Error GoTo err_Handle
Screen.MousePointer = 11

Dim i As Integer, strFileName As String, strFileName1 As String, strCheck As String

'轉文字檔
If Dir("C:\LTKK01\Ship2TKK", vbDirectory) = "" Then MkDirs "C:\LTKK01\Ship2TKK"
If Dir("C:\LTKK01\Ship2TKK\Backup", vbDirectory) = "" Then MkDirs "C:\LTKK01\Ship2TKK\Backup"
strFileName = "出貨回檔-系統" & Format(Now, "yyyymmddhhMMss") & ".csv"
strFileName1 = "出貨回檔-手開" & Format(Now, "yyyymmddhhMMss") & ".csv"

Open "C:\LTKK01\Ship2TKK\" & strFileName For Output As #1
Open "C:\LTKK01\Ship2TKK\" & strFileName1 For Output As #2

'交易開始
Tran_Level = cn.BeginTrans

'手開單寫入第一筆資料
Print #2, "交貨單號"; ","; "到貨日"; ","; "到貨日"; ","; "訂單號碼"; ","; "項次"; ","; "數量"; ","; "製造日"; ","; "B"; ","; "地址別"; ","; "客戶名稱"; ","; "料號"
Dim strA As String, strB As String, strC As String, strD As String, strE As String, intF As Integer, strG As String, strH As String, strI As String, strJ As String, strK As String, strL As String, strM As String

rsMain.MoveFirst
strA = RTrim(rsMain("交貨單號"))
strB = RTrim(rsMain("到貨日"))
strC = RTrim(rsMain("到貨日"))
strD = RTrim(rsMain("訂單號碼"))
strE = RTrim(rsMain("項次"))
strH = RTrim(rsMain("B"))
strI = RTrim(rsMain("地址別"))
strJ = RTrim(rsMain("客戶名稱"))
strK = RTrim(rsMain("料號"))
strL = RTrim(rsMain("訂單來源"))
strM = RTrim(rsMain("WMS單號"))
strCheck = RTrim(rsMain("訂單號碼")) & RTrim(rsMain("項次"))

Do While Not rsMain.EOF

    If strCheck = RTrim(rsMain("訂單號碼")) & RTrim(rsMain("項次")) Then
        '同單號品項數量相加
        intF = intF + RTrim(rsMain("數量")): strG = strG & RTrim(rsMain("製造日")) & ";"
    Else
        '不同單號品項
        '檢查是否系統單
        If Len(strL) > 0 Then
            '系統單
            Print #1, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH
            
        Else
            '手開單
            Print #1, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH; ","; strI; ","; strJ; ","; strK
        
        End If
        
    '更新為已回傳
    cn.Execute "update " & strWMSDB & "..orders set yfystatus = '2' ,TrafficCop = null where orderkey = '" & strM & "' ", RowsAffect, adExecuteNoRecords
    
    '歸零
    strA = RTrim(rsMain("交貨單號"))
    strB = RTrim(rsMain("到貨日"))
    strC = RTrim(rsMain("到貨日"))
    strD = RTrim(rsMain("訂單號碼"))
    strE = RTrim(rsMain("項次"))
    intF = RTrim(rsMain("數量"))
    strG = RTrim(rsMain("製造日")) & ";"
    strH = RTrim(rsMain("B"))
    strI = RTrim(rsMain("地址別"))
    strJ = RTrim(rsMain("客戶名稱"))
    strK = RTrim(rsMain("料號"))
    strL = RTrim(rsMain("訂單來源"))
    strM = RTrim(rsMain("WMS單號"))
    strCheck = RTrim(rsMain("訂單號碼")) & RTrim(rsMain("項次"))
    End If
    rsMain.MoveNext
Loop

'寫入最後資料
'檢查是否系統單
If Len(strL) > 0 Then
    '系統單
    Print #1, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH
    
Else
    '手開單
    Print #1, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH; ","; strI; ","; strJ; ","; strK

End If

'更新為已回傳
cn.Execute "update " & strWMSDB & "..orders set yfystatus = '2' ,TrafficCop = null where orderkey = '" & strM & "' ", RowsAffect, adExecuteNoRecords

cn.CommitTrans: Tran_Level = 0

'關閉檔案
Close

'備份檔案
FileCopy "C:\LTKK01\Ship2TKK\" & strFileName, "C:\LTKK01\Ship2TKK\Backup\" & strFileName
FileCopy "C:\LTKK01\Ship2TKK\" & strFileName1, "C:\LTKK01\Ship2TKK\Backup\" & strFileName1

Set rsMain = Nothing: Set dgMain.DataSource = Nothing
Screen.MousePointer = 0
MsgBox "出貨資料轉出完成!!" & vbCrLf & "C:\LTKK01\Ship2TKK\Backup\" & strFileName & vbCrLf & "C:\LTKK01\Ship2TKK\Backup\" & strFileName1, vbOKOnly, Me.Caption
Exit Sub

err_Handle:
    Close
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

'貨主
    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.CursorLocation = adUseClient
    tmp_Rs.Open "select distinct(storerkey) from trp16M", cn, adOpenKeyset, adLockPessimistic
    tmp_Rs.MoveFirst
    For i = 0 To tmp_Rs.RecordCount - 1
        Combo1.AddItem tmp_Rs("storerkey")
        tmp_Rs.MoveNext
    Next
    tmp_Rs.Close: Set tmp_Rs = Nothing
'    Combo1.ListIndex = 0
    Combo1.Text = "LTKK01"
    
    txtDeliveryDateS = Format(Now, "YYYYMM") & "01"
    txtDeliveryDateE = Format(Now, "YYYYMMDD")

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
