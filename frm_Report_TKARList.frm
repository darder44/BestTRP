VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_Report_TKARList 
   Caption         =   "應收帳款明細表"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
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
   ScaleHeight     =   10950
   ScaleWidth      =   20250
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   960
         Width           =   1485
      End
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
         Picture         =   "frm_Report_TKARList.frx":0000
         Style           =   1  '圖片外觀
         TabIndex        =   15
         Top             =   1200
         Visible         =   0   'False
         Width           =   1065
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
         Visible         =   0   'False
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
         Picture         =   "frm_Report_TKARList.frx":030A
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
         Picture         =   "frm_Report_TKARList.frx":1604
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
         Picture         =   "frm_Report_TKARList.frx":2B216
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
         Picture         =   "frm_Report_TKARList.frx":2B528
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   240
         Width           =   1065
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         Caption         =   "接單日期"
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
      Top             =   10680
      Width           =   20250
      _ExtentX        =   35719
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
            Object.Width           =   29078
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
Attribute VB_Name = "frm_Report_TKARList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As ADODB.Recordset
Private intColumnIndex As Integer
Private objMvdateTarget As Object
Private intPickRow As Long, intLastCol As Long, intOrderRow As Long, intSkuRow As Long, intPickqty As Long, strDeliveryDateS As String, strDeliveryDateE As String

Private Sub cmd2Excel_Click()

Call WriteOut_RunLog("1/16.轉出計費明細資料")
Recordset2Excel "LTKK01應收帳款明細表", rsMain
If rsMain Is Nothing Then Call Unload_RunLogForm: Exit Sub

'..在此編輯EXCEL
Screen.MousePointer = 11
Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String

With MyXlsApp: .Visible = False

If RTrim(Combo1) = "LTKK01" Then
    cn.Execute "if object_id ('tempdb..##LTKK01ARList') is not null drop table ##LTKK01ARList exec gs_LTKK01ARList '" & strDeliveryDateS & "' , '" & strDeliveryDateE & "' ", RowsAffect, adExecuteNoRecords
    
    Dim rsTmp As New ADODB.Recordset

'日報表

    '尋找工作表
    strSheet = "日報表"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增DATA工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
        
'    .Sheets.Add: .ActiveSheet.Name = "會計請付款資料"
    str_SQL = "select * from gv_" & Combo1.Text & "Charge where 1 = 1 " & "and 載貨日期 between '" & txtDeliveryDateS & "' and '" & txtDeliveryDateE & "' order by 請款類別,序號,載貨日期,車號 "

    Call Confirm_Recordset_Closed(tmp_Rs)
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("2/16.轉出日報表資料")
    Call OffLineRecordset(tmp_Rs, rsTmp)

    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i

    .Range("A2").CopyFromRecordset rsTmp

    rsTmp.Close
    
'專車及理貨
Screen.MousePointer = 11
'尋找工作表
strSheet = "專車及理貨"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
Next

'找不到新增工作表
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "SELECT 下單日期 = cast(RECEIPT_DATE as datetime),載貨日期 = cast(ARRIVE_DATE as datetime) " & _
            ",簽單日期 = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
            ",車號 = C_VEHICLE_ID_NO ,訖點 = areaend ,客戶單號 = orderkey,店家代碼 = SHIPTO " & _
            ",客戶名稱 = FULL_NAME ,品項 = reason ,產品別 = SUSR1 ,品牌別 = SUSR3 " & _
            ",用途別 = NOTES1  " & _
            ",區別 = NOTES2 " & _
            ",出貨箱數 = ship_cs " & _
            ",數量 = chargeqty " & _
            ",單位 = uom " & _
            ",不足公斤數 = FULL_KG " & _
            ",配送費單價 = receivable " & _
            ",配送費總價 = sumreceivable " & _
            ",理貨費單價 = SortingAR " & _
            ",理貨費總價 = SUMSortingAR " & _
            ",路線編號 = route_no " & _
            ",通路別 = channel " & _
            ",地址別中文 = short_name ,備註 = note " & _
            "from ##LTKK01ARList " & _
            "where priority <> 'R' " & _
            "and costkind <> '原車退回' and note like ('專車%') " & _
            "order by ARRIVE_DATE,orderkey,SUSR1 "
        
Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open str_SQL, cn
Call WriteOut_RunLog("3/16.轉出專車運費資料")
Call Replication_Recordset(tmp_Rs, rsTmp)

'寫入標題列
k = 65: j = 1: strCol = ""
For i = 0 To rsTmp.Fields.Count - 1
    l = i Mod 26
    .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
    '欄位超過26
    If Chr(65 + l) = "Z" Then
        If strCol = "" Then
            strCol = "A"
        Else
            strCol = Chr(Asc(strCol) + 1)
        End If
    End If
Next i

.Range("A2").CopyFromRecordset rsTmp

rsTmp.Close

'專車運費分析
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "專車運費分析"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT 請款類別 = '專車運費',載貨日期 = cast(ARRIVE_DATE as datetime) " & _
                ",簽單日期 = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",品項 = reason ,產品別 = SUSR1 ,品牌別 = SUSR3 " & _
                ",用途別 = NOTES1  " & _
                ",區別 = NOTES2 " & _
                ",配送費總價 = sumreceivable " & _
                ",客戶名稱 = FULL_NAME ,店家代碼 = SHIPTO " & _
                ",通路別 = channel " & _
                ",地址別中文 = short_name " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '原車退回' and note like ('專車%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("4/16.轉出專車運費分析資料")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'專車理貨分析
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "專車理貨分析"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT 請款類別 = '專車理貨費',載貨日期 = cast(ARRIVE_DATE as datetime) " & _
                ",簽單日期 = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",品項 = reason ,產品別 = SUSR1 ,品牌別 = SUSR3 " & _
                ",用途別 = NOTES1  " & _
                ",區別 = NOTES2 " & _
                ",理貨費總價 = SUMSortingAR " & _
                ",客戶名稱 = FULL_NAME ,店家代碼 = SHIPTO " & _
                ",通路別 = channel " & _
                ",地址別中文 = short_name " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '原車退回' and note like ('專車%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("5/16.轉出專車理貨分析資料")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close

'外島運費
Screen.MousePointer = 11
'尋找工作表
strSheet = "外島運費"
For i = 1 To .Sheets.Count
    If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
Next

'找不到新增工作表
If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "SELECT 下單日期 = cast(RECEIPT_DATE as datetime),載貨日期 = cast(ARRIVE_DATE as datetime) " & _
            ",簽單日期 = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
            ",車號 = C_VEHICLE_ID_NO ,訖點 = areaend ,客戶單號 = orderkey,店家代碼 = SHIPTO " & _
            ",客戶名稱 = FULL_NAME ,品項 = reason ,產品別 = SUSR1 ,品牌別 = SUSR3 " & _
            ",用途別 = NOTES1  " & _
            ",區別 = NOTES2 " & _
            ",出貨箱數 = ship_cs " & _
            ",數量 = chargeqty " & _
            ",單位 = uom " & _
            ",不足公斤數 = FULL_KG " & _
            ",配送費總價 = sumreceivable " & _
            ",路線編號 = route_no " & _
            ",通路別 = channel " & _
            ",地址別中文 = short_name " & _
            "from ##LTKK01ARList " & _
            "where priority <> 'R' " & _
            "and costkind <> '原車退回' and rtrim(costcode) in ('000-67','002-09','002-43') " & _
            "order by ARRIVE_DATE,orderkey,SUSR1 "
        
Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open str_SQL, cn
Call WriteOut_RunLog("6/16.轉出外島運費資料")
Call Replication_Recordset(tmp_Rs, rsTmp)

'寫入標題列
k = 65: j = 1: strCol = ""
For i = 0 To rsTmp.Fields.Count - 1
    l = i Mod 26
    .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
    '欄位超過26
    If Chr(65 + l) = "Z" Then
        If strCol = "" Then
            strCol = "A"
        Else
            strCol = Chr(Asc(strCol) + 1)
        End If
    End If
Next i

.Range("A2").CopyFromRecordset rsTmp

rsTmp.Close

'外島運費分析
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "外島運費分析"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT 請款類別 = '外島運費',載貨日期 = cast(ARRIVE_DATE as datetime) " & _
                ",簽單日期 = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",品項 = reason ,產品別 = SUSR1 ,品牌別 = SUSR3 " & _
                ",用途別 = NOTES1  " & _
                ",區別 = NOTES2 " & _
                ",配送費總價 = sumreceivable " & _
                ",客戶名稱 = FULL_NAME ,店家代碼 = SHIPTO " & _
                ",通路別 = channel " & _
                ",地址別中文 = short_name " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '原車退回' and rtrim(costcode) in ('000-67','002-09','002-43') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("7/16.轉出外島運費分析資料")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'退貨及理貨
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "退貨及理貨"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet
'    .Sheets.Add: .ActiveSheet.Name = "退貨及理貨"

    str_SQL = "SELECT 下單日期 = cast(RECEIPT_DATE as datetime),載貨日期 = cast(ARRIVE_DATE as datetime) " & _
            ",簽單日期 = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
            ",車號 = C_VEHICLE_ID_NO,起點 = areastart,客戶單號 = orderkey,店家代碼 = SHIPTO,客戶名稱 = FULL_NAME,品項 = reason " & _
            ",原因 = case when priority = 'R' then '通知收退回' else rtrim(costkind) end,產品別 = SUSR1,品牌別 = SUSR3,用途別 = NOTES1 " & _
            ",區別 = NOTES2,出貨箱數 = ship_cs,數量 = chargeqty,單位 = uom,配送費單價 = receivable,配送費總價 = sumreceivable,理貨費單價 = SortingAR " & _
            ",理貨費總價 = SUMSortingAR,路線編號 = route_no,通路別 = channel,地址別中文 = short_name ,備註 = note " & _
            "from ##LTKK01ARList " & _
            "where (priority = 'R' or costkind = '原車退回') " & _
            "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("8/16.轉出退貨及理貨資料")
    Call Replication_Recordset(tmp_Rs, rsTmp)
    
    '寫入標題列
    k = 65: j = 1
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'退貨運費分析
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "退貨運費分析"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT 請款類別 = '退貨運費',載貨日期 = cast(ARRIVE_DATE as datetime) " & _
            ",簽單日期 = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
            ",品項 = reason ,產品別 = SUSR1,品牌別 = SUSR3,用途別 = NOTES1 ,區別 = NOTES2,配送費總價 = sumreceivable " & _
            ",客戶名稱 = FULL_NAME ,店家代碼 = SHIPTO " & _
            ",通路別 = channel,地址別中文 = short_name " & _
            "from ##LTKK01ARList " & _
            "where (priority = 'R' or costkind = '原車退回') " & _
            "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("9/16.轉出退貨運費分析資料")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'退貨理貨分析
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "退貨理貨分析"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT 請款類別 = '退貨理貨',載貨日期 = cast(ARRIVE_DATE as datetime) " & _
            ",簽單日期 = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
            ",品項 = reason ,產品別 = SUSR1,品牌別 = SUSR3,用途別 = NOTES1 ,區別 = NOTES2,理貨費總價 = SUMSortingAR " & _
            ",客戶名稱 = FULL_NAME ,店家代碼 = SHIPTO " & _
            ",通路別 = channel,地址別中文 = short_name " & _
            "from ##LTKK01ARList " & _
            "where (priority = 'R' or costkind = '原車退回') " & _
            "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("10/16.轉出退貨理貨分析資料")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'配送及理貨
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "配送及理貨"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT 下單日期 = cast(RECEIPT_DATE as datetime),載貨日期 = cast(ARRIVE_DATE as datetime) " & _
                ",簽單日期 = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",車號 = C_VEHICLE_ID_NO ,訖點 = areaend ,客戶單號 = orderkey,店家代碼 = SHIPTO " & _
                ",客戶名稱 = FULL_NAME ,品項 = reason ,產品別 = SUSR1 ,品牌別 = SUSR3 " & _
                ",用途別 = NOTES1  " & _
                ",區別 = NOTES2 " & _
                ",出貨箱數 = ship_cs " & _
                ",數量 = chargeqty " & _
                ",單位 = uom " & _
                ",不足公斤數 = FULL_KG " & _
                ",配送費單價 = receivable " & _
                ",配送費總價 = sumreceivable " & _
                ",理貨費單價 = SortingAR " & _
                ",理貨費總價 = SUMSortingAR " & _
                ",路線編號 = route_no " & _
                ",通路別 = channel " & _
                ",地址別中文 = short_name ,備註 = note " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '原車退回' and rtrim(costcode) not in ('000-67','002-09','002-43','Bonded') and note not like ('專車%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("11/16.轉出配送及理貨資料")
    Call Replication_Recordset(tmp_Rs, rsTmp)
    
    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    
'配送分析
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "配送分析"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT 請款類別 = '配送費',載貨日期 = cast(ARRIVE_DATE as datetime) " & _
                ",簽單日期 = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",品項 = reason ,產品別 = SUSR1 ,品牌別 = SUSR3 " & _
                ",用途別 = NOTES1  " & _
                ",區別 = NOTES2 " & _
                ",配送費總價 = sumreceivable " & _
                ",客戶名稱 = FULL_NAME ,店家代碼 = SHIPTO " & _
                ",通路別 = channel " & _
                ",地址別中文 = short_name " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '原車退回' and rtrim(costcode) not in ('000-67','002-09','002-43','Bonded') and note not like ('專車%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("12/16.轉出配送分析資料")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close

    '保稅及理貨
    Screen.MousePointer = 11
    '尋找工作表
    strSheet = "保稅及理貨"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next

    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT 下單日期 = cast(RECEIPT_DATE as datetime),載貨日期 = cast(ARRIVE_DATE as datetime) " & _
                ",簽單日期 = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",車號 = C_VEHICLE_ID_NO ,訖點 = areaend ,客戶單號 = orderkey,店家代碼 = SHIPTO " & _
                ",客戶名稱 = FULL_NAME ,品項 = reason ,產品別 = SUSR1 ,品牌別 = SUSR3 " & _
                ",用途別 = NOTES1  " & _
                ",區別 = NOTES2 " & _
                ",出貨箱數 = ship_cs " & _
                ",數量 = chargeqty " & _
                ",單位 = uom " & _
                ",不足公斤數 = FULL_KG " & _
                ",配送費單價 = receivable " & _
                ",配送費總價 = sumreceivable " & _
                ",理貨費單價 = SortingAR " & _
                ",理貨費總價 = SUMSortingAR " & _
                ",路線編號 = route_no " & _
                ",通路別 = channel " & _
                ",地址別中文 = short_name ,備註 = note " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '原車退回' and rtrim(costcode) = 'Bonded' and note not like ('專車%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "

    Call Confirm_Recordset_Closed(tmp_Rs)

    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("13/16.轉出保稅及理貨資料")
    Call Replication_Recordset(tmp_Rs, rsTmp)

    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i

    .Range("A2").CopyFromRecordset rsTmp

    rsTmp.Close

'配送分析
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "保稅分析"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next

    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT 請款類別 = '配送費',載貨日期 = cast(ARRIVE_DATE as datetime) " & _
                ",簽單日期 = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",品項 = reason ,產品別 = SUSR1 ,品牌別 = SUSR3 " & _
                ",用途別 = NOTES1  " & _
                ",區別 = NOTES2 " & _
                ",配送費總價 = sumreceivable " & _
                ",客戶名稱 = FULL_NAME ,店家代碼 = SHIPTO " & _
                ",通路別 = channel " & _
                ",地址別中文 = short_name " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '原車退回' and costcode = 'Bonded' and note not like ('專車%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "

    Call Confirm_Recordset_Closed(tmp_Rs)

    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("14/16.轉出保稅分析資料")
    Call OffLineRecordset(tmp_Rs, rsTmp)

    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i

    .Range("A2").CopyFromRecordset rsTmp

    rsTmp.Close
    
'理貨分析
Screen.MousePointer = 11
    '尋找工作表
    strSheet = "理貨分析"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "SELECT 請款類別 = '理貨費',載貨日期 = cast(ARRIVE_DATE as datetime) " & _
                ",簽單日期 = cast(case when Right(RTrim(ARRIVE_DATE), 2) > 25 then convert(char(6),dateadd(m,1,cast(ARRIVE_DATE as datetime)),112) + '01' Else ARRIVE_DATE End as datetime) " & _
                ",品項 = reason ,產品別 = SUSR1 ,品牌別 = SUSR3 " & _
                ",用途別 = NOTES1  " & _
                ",區別 = NOTES2 " & _
                ",理貨費總價 = SUMSortingAR " & _
                ",客戶名稱 = FULL_NAME ,店家代碼 = SHIPTO " & _
                ",通路別 = channel " & _
                ",地址別中文 = short_name " & _
                "from ##LTKK01ARList " & _
                "where priority <> 'R' " & _
                "and costkind <> '原車退回' and rtrim(costcode) not in ('000-67','002-09','002-43','Bonded') and note not like ('專車%') " & _
                "order by ARRIVE_DATE,orderkey,SUSR1 "
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("15/16.轉出理貨分析資料")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close


'應收付
    Screen.MousePointer = 11
    '尋找工作表
    strSheet = "應收付"
    For i = 1 To .Sheets.Count
        If UCase(RTrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
    Next
    
    '找不到新增工作表
    If UCase(RTrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

    str_SQL = "exec es_LTKK01ARP '" & txtDeliveryDateS & "','" & txtDeliveryDateE & "'"
            
    Call Confirm_Recordset_Closed(tmp_Rs)
    
    tmp_Rs.Open str_SQL, cn
    Call WriteOut_RunLog("16/16.轉出應收付資料")
    Call OffLineRecordset(tmp_Rs, rsTmp)
    
    '寫入標題列
    k = 65: j = 1: strCol = ""
    For i = 0 To rsTmp.Fields.Count - 1
        l = i Mod 26
        .Range(strCol & Chr(k + l) & j).Value = rsTmp.Fields(i).Name
        '欄位超過26
        If Chr(65 + l) = "Z" Then
            If strCol = "" Then
                strCol = "A"
            Else
                strCol = Chr(Asc(strCol) + 1)
            End If
        End If
    Next i
    
    .Range("A2").CopyFromRecordset rsTmp
    
    rsTmp.Close
    cn.Execute "if object_id ('tempdb..##LTKK01ARList') is not null drop table ##LTKK01ARList ", RowsAffect, adExecuteNoRecords
End If
.Visible = True: End With

Call Unload_RunLogForm
Set MyXlsApp = Nothing
Screen.MousePointer = 0

End Sub

Private Sub cmdQuery_Click()
On Error GoTo err_Handle
If Len(txtDeliveryDateS.Text) = 0 Or Len(txtDeliveryDateE.Text) = 0 Then MsgBox "請輸入起訖日期區間！", vbOKOnly, Me.Caption: Exit Sub
strDeliveryDateS = txtDeliveryDateS.Text: strDeliveryDateE = txtDeliveryDateE.Text
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_Orderdate As String, chc_DeliveryDate As String
    
'訂單日期
chc_Orderdate = ""
If Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and YMD between '" & txtOrderDateS.Text & "' and '" & txtOrderDateE.Text & "' "
ElseIf Len(txtOrderDateS.Text) > 0 And Len(txtOrderDateE.Text) = 0 Then
   chc_Orderdate = "and YMD = '" & txtOrderDateS.Text & "' "
ElseIf Len(txtOrderDateS.Text) = 0 And Len(txtOrderDateE.Text) > 0 Then
   chc_Orderdate = "and YMD = '" & txtOrderDateE.Text & "' "
End If

'到貨日期
chc_DeliveryDate = "and 到貨日 between '" & strDeliveryDateS & "' and '" & strDeliveryDateE & "' "

str_SQL = "select * from gv_sdn05tdetail where 1 = 1 " & chc_Orderdate & chc_DeliveryDate

'貨主
If Len(RTrim(Combo1.Text)) > 0 Then str_SQL = str_SQL & "and 貨主 = '" & RTrim(Combo1.Text) & "' "

Set rsMain = New ADODB.Recordset
rsMain.CursorLocation = adUseClient
rsMain.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If rsMain.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
rsMain.Sort = "到貨日,路線編號,貨主單號"

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
Combo1.ListIndex = 0

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
