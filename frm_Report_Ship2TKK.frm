VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_Report_Ship2TKK 
   Caption         =   "出貨資料回傳"
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
   ScaleHeight     =   6300
   ScaleWidth      =   10335
   Begin MSComCtl2.MonthView mvDate 
      Height          =   2220
      Left            =   3240
      TabIndex        =   8
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
         Left            =   5880
         Picture         =   "frm_Report_Ship2TKK.frx":0000
         Style           =   1  '圖片外觀
         TabIndex        =   15
         Top             =   1200
         Width           =   1065
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
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
         Enabled         =   0   'False
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
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.TextBox txtOrderDateS 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
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
         Picture         =   "frm_Report_Ship2TKK.frx":030A
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
         Left            =   7080
         Picture         =   "frm_Report_Ship2TKK.frx":1604
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
         Picture         =   "frm_Report_Ship2TKK.frx":2B216
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
         Picture         =   "frm_Report_Ship2TKK.frx":2B528
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "倉庫扣帳後每日1300前回傳"
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
         Left            =   1440
         TabIndex        =   16
         Top             =   960
         Width           =   2880
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
         Caption         =   "出車日期"
         Enabled         =   0   'False
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
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "～"
         Enabled         =   0   'False
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
         Visible         =   0   'False
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
Attribute VB_Name = "frm_Report_Ship2TKK"
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
Recordset2Excel Me.Caption, rsMain
'..在此編輯EXCEL
Set MyXlsApp = Nothing

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_Orderdate As String

str_SQL = "exec gs_ship2tkk "

Set rsMain = New ADODB.Recordset
rsMain.CursorLocation = adUseClient
rsMain.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMain.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
'rsMain.Sort = "貨主編號,訂單號碼,項次"

Set dgMain.DataSource = rsMain: dgMain.Visible = False
rsMain.MoveFirst

SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMain.Visible = True

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, err.Number, err.Description, "")
End Sub

Private Sub cmdSaveToText_Click()

If rsMain Is Nothing Then Exit Sub: If rsMain.EOF Then Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11: cmdSaveToText.Enabled = False: dgMain.Enabled = False

Dim i As Integer, strFileName1 As String, strFileName2 As String, strFileName3 As String, strCheck As String

'轉文字檔
If Dir("C:\LTKK01\出貨資料回傳", vbDirectory) = "" Then MkDirs "C:\LTKK01\出貨資料回傳"
'If Dir("C:\LTKK01\DELI\Backup", vbDirectory) = "" Then MkDirs "C:\LTKK01\DELI\Backup"
strFileName1 = "DELI-系統_" & Format(Now, "yyyymmddhhMMss") & ".csv"
strFileName2 = "DELI-手開_" & Format(Now, "yyyymmddhhMMss") & ".csv"
strFileName3 = "轉倉_" & Format(Now, "yyyymmddhhMMss") & ".csv"

Open "C:\LTKK01\出貨資料回傳\" & strFileName1 For Output As #1
Open "C:\LTKK01\出貨資料回傳\" & strFileName2 For Output As #2
Open "C:\LTKK01\出貨資料回傳\" & strFileName3 For Output As #3

'交易開始
Tran_Level = cn.BeginTrans

rsMain.MoveFirst
'手開單寫入抬頭
'Print #2, "交貨單號"; ","; "到貨日"; ","; "到貨日"; ","; "訂單號碼"; ","; "項次"; ","; "數量"; ","; "製造日"; ","; "B"; ","; "地址別"; ","; "客戶名稱"; ","; "料號"
Dim strA As String, strB As String, strC As String, strD As String, strE As String, intF As Long, strG As String, strH As String, strI As String, strJ As String, strK As String, strL As String, strM As String

rsMain.MoveFirst
strA = RTrim(rsMain("交貨單號"))
strB = RTrim(rsMain("到貨日"))
strC = RTrim(rsMain("到貨日"))
strD = RTrim(rsMain("訂單號碼")) & ""
strE = RTrim(rsMain("項次")) & ""
strH = RTrim(rsMain("B"))
strI = RTrim(rsMain("地址別")) & ""
strJ = RTrim(rsMain("客戶名稱")) & ""
strK = RTrim(rsMain("料號"))
strL = RTrim(rsMain("類別")) & ""
strM = RTrim(rsMain("WMS單號"))
strCheck = RTrim(rsMain("訂單號碼")) & RTrim(rsMain("項次")) & RTrim(rsMain("客戶名稱")) & RTrim(rsMain("交貨單號")) & RTrim(rsMain("料號"))

Do While Not rsMain.EOF

    If strCheck = RTrim(rsMain("訂單號碼")) & RTrim(rsMain("項次")) & RTrim(rsMain("客戶名稱")) & RTrim(rsMain("交貨單號")) & RTrim(rsMain("料號")) Then
        '同單號品項數量相加
        intF = intF + RTrim(rsMain("數量")): strG = strG & RTrim(rsMain("製造日")) & ";"
    Else
        '不同單號品項
        '檢查是否系統單
        If Len(strA) > 0 Then
        
            '系統單
            Print #1, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH
          
        Else
            '手開單
            If strL = "A" Then ' 轉倉單
            
                Print #3, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH; ","; strI; ","; strJ; ","; strK
            Else
                Print #2, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH; ","; strI; ","; strJ; ","; strK
                    
            End If
        End If
        
        '歸零
        strA = RTrim(rsMain("交貨單號"))
        strB = RTrim(rsMain("到貨日"))
        strC = RTrim(rsMain("到貨日"))
        strD = RTrim(rsMain("訂單號碼")) & ""
        strE = RTrim(rsMain("項次")) & ""
        intF = RTrim(rsMain("數量"))
        strG = RTrim(rsMain("製造日")) & ";"
        strH = RTrim(rsMain("B"))
        strI = RTrim(rsMain("地址別")) & ""
        strJ = RTrim(rsMain("客戶名稱")) & ""
        strK = RTrim(rsMain("料號"))
        strL = RTrim(rsMain("類別")) & ""
        strM = RTrim(rsMain("WMS單號"))
        strCheck = RTrim(rsMain("訂單號碼")) & RTrim(rsMain("項次")) & RTrim(rsMain("客戶名稱")) & RTrim(rsMain("交貨單號")) & RTrim(rsMain("料號"))
    End If
    
    '更新為已回傳
    str_SQL = "update " & strWMSDB & "..orders " & _
                "set yfystatus = '2' ,TrafficCop = null where orderkey = '" & RTrim(rsMain("WMS單號")) & "' and storerkey = 'LTKK01'"
'                "where orderkey in (select od.orderkey from " & strWMSDB & "..orderdetail od where od.externorderkey = '" & strD & "' and od.externlineno = '" & RTrim(strE) & RTrim(strA) & "')"

    cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
    
    rsMain.MoveNext
Loop

'寫入最後資料
'檢查是否系統單
If Len(strA) > 0 Then
    '系統單
    Print #1, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH
    
Else
    If strL = "A" Then ' 轉倉單
    
        Print #3, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH; ","; strI; ","; strJ; ","; strK
    Else
        Print #2, strA; ","; strB; ","; strC; ","; strD; ","; strE; ","; intF; ","; strG; ","; strH; ","; strI; ","; strJ; ","; strK
            
    End If
End If

'更新為已回傳
str_SQL = "update " & strWMSDB & "..orders " & _
            "set yfystatus = '2' ,TrafficCop = null where orderkey = '" & strM & "' and storerkey = 'LTKK01' "
'            "where orderkey in (select od.orderkey from " & strWMSDB & "..orderdetail od where od.externorderkey = '" & strD & "' and od.externlineno = '" & RTrim(strE) & RTrim(strA) & "')"

cn.Execute str_SQL, RowsAffect, adExecuteNoRecords

'關閉檔案
Close

'備份檔案
'If Dir("C:\LTKK01\出貨資料回傳\", vbDirectory) = "" Then MkDirs "C:\LTKK01\出貨資料回傳"
'FileCopy "C:\LTKK01\Ship2TKK\" & strFileName, "C:\LTKK01\Ship2TKK\Backup\" & strFileName
'FileCopy "C:\LTKK01\Ship2TKK\" & strFileName1, "C:\LTKK01\Ship2TKK\Backup\" & strFileName1

cn.CommitTrans: Tran_Level = 0

Set rsMain = Nothing: Set dgMain.DataSource = Nothing
Screen.MousePointer = 0: cmdSaveToText.Enabled = True: dgMain.Enabled = True
MsgBox "出貨資料轉出完成!!" & vbCrLf & "C:\LTKK01\出貨資料回傳\" & strFileName1 & vbCrLf & "C:\LTKK01\出貨資料回傳\" & strFileName2 & vbCrLf & "C:\LTKK01\出貨資料回傳\" & strFileName3 & vbCrLf, vbOKOnly, Me.Caption
Exit Sub

err_Handle:
    Screen.MousePointer = 0: cmdSaveToText.Enabled = True: dgMain.Enabled = True
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
txtOrderDateS.Text = "": txtOrderDateE.Text = ""

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
    Combo1.ListIndex = 0
    Combo1.Text = "LTKK01"

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
