VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Report_TKKPI 
   Caption         =   "單量明細"
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
      StartOfWeek     =   60948481
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
      Begin VB.CheckBox chkIncData 
         Caption         =   "含明細"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5880
         TabIndex        =   23
         Top             =   1320
         Width           =   945
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frm_Report_TKKPI.frx":0000
         Left            =   120
         List            =   "frm_Report_TKKPI.frx":000A
         Style           =   2  '單純下拉式
         TabIndex        =   19
         Top             =   600
         Width           =   1125
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
         Visible         =   0   'False
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
         Visible         =   0   'False
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
         Picture         =   "frm_Report_TKKPI.frx":001E
         Style           =   1  '圖片外觀
         TabIndex        =   14
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
         Left            =   4320
         Style           =   2  '單純下拉式
         TabIndex        =   12
         Top             =   1680
         Visible         =   0   'False
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
         Left            =   5880
         Picture         =   "frm_Report_TKKPI.frx":0328
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
         Picture         =   "frm_Report_TKKPI.frx":1622
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
         Picture         =   "frm_Report_TKKPI.frx":2B234
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
         Picture         =   "frm_Report_TKKPI.frx":2B546
         Style           =   1  '圖片外觀
         TabIndex        =   2
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "結束日期"
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
         Left            =   3240
         TabIndex        =   22
         Top             =   360
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "起始日期"
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
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   960
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
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   360
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
         Left            =   2640
         TabIndex        =   18
         Top             =   1020
         Visible         =   0   'False
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
         Visible         =   0   'False
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
         Left            =   3480
         TabIndex        =   13
         Top             =   1740
         Visible         =   0   'False
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
            Object.Width           =   11615
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
Attribute VB_Name = "frm_Report_TKKPI"
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

If Combo2 = "接單日" Then
    Call cmd2Excel_AddDate
Else
    Call cmd2Excel_DeliveryDate
End If

End Sub

Private Sub cmd2Excel_AddDate()

Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String

Recordset2Excel Combo2 & Me.Caption, rsMain

'在此編輯EXCEL
Screen.MousePointer = 11
With MyXlsApp: .Visible = False

Dim rsTmp As New ADODB.Recordset

'其他
Screen.MousePointer = 11
'尋找工作表
strSheet = "其他"
For i = 1 To .Sheets.Count
    If UCase(Rtrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
Next

'找不到新增工作表
If UCase(Rtrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_TKKPIxAddDatexChannel '其他','" & txtOrderDateS & "','" & txtOrderDateE & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open str_SQL, cn
Call WriteOut_RunLog("1/4.轉出其他通路KPI")
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

'SD
Screen.MousePointer = 11
'尋找工作表
strSheet = "SD"
For i = 1 To .Sheets.Count
    If UCase(Rtrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
Next

'找不到新增工作表
If UCase(Rtrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_TKKPIxAddDatexChannel 'SD','" & txtOrderDateS & "','" & txtOrderDateE & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open str_SQL, cn
Call WriteOut_RunLog("2/4.轉出SD通路KPI")
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
    
'KA
Screen.MousePointer = 11
'尋找工作表
strSheet = "KA"
For i = 1 To .Sheets.Count
    If UCase(Rtrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
Next

'找不到新增工作表
If UCase(Rtrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_TKKPIxAddDatexChannel 'KA','" & txtOrderDateS & "','" & txtOrderDateE & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open str_SQL, cn
Call WriteOut_RunLog("3/4.轉出KA通路KPI")
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

'ALL
Screen.MousePointer = 11
'尋找工作表
strSheet = "ALL"
For i = 1 To .Sheets.Count
    If UCase(Rtrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
Next

'找不到新增工作表
If UCase(Rtrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_TKKPIxAddDatexChannel 'ALL','" & txtOrderDateS & "','" & txtOrderDateE & "' "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn
Call WriteOut_RunLog("4/4.轉出所有通路KPI")
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

'刪除DATA工作表
If chkIncData = 0 Then
    '尋找工作表
    For i = 1 To .Sheets.Count
        If UCase(Rtrim(.Sheets(i).Name)) = "DATA" Then .Sheets("DATA").Delete: Exit For
        If UCase(Rtrim(.Sheets(i).Name)) = Combo2 & Me.Caption Then .Sheets(Combo2 & Me.Caption).Delete: Exit For
    Next
End If

.Visible = True: End With

Call Unload_RunLogForm
Set MyXlsApp = Nothing
Screen.MousePointer = 0

End Sub

Private Sub cmd2Excel_DeliveryDate()

Dim i As Integer, j As Integer, k As Integer, l As Integer, strCol As String, strSheet As String

Recordset2Excel Combo2 & Me.Caption, rsMain

'在此編輯EXCEL
Screen.MousePointer = 11
With MyXlsApp: .Visible = False

Dim rsTmp As New ADODB.Recordset

'其他
Screen.MousePointer = 11
'尋找工作表
strSheet = "其他"
For i = 1 To .Sheets.Count
    If UCase(Rtrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
Next

'找不到新增工作表
If UCase(Rtrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_TKKPIxdeliveryDatexChannel '其他','" & txtOrderDateS & "','" & txtOrderDateE & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open str_SQL, cn
Call WriteOut_RunLog("1/4.轉出其他通路KPI")
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

'SD
Screen.MousePointer = 11
'尋找工作表
strSheet = "SD"
For i = 1 To .Sheets.Count
    If UCase(Rtrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
Next

'找不到新增工作表
If UCase(Rtrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_TKKPIxdeliveryDatexChannel 'SD','" & txtOrderDateS & "','" & txtOrderDateE & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open str_SQL, cn
Call WriteOut_RunLog("2/4.轉出SD通路KPI")
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

'KA
Screen.MousePointer = 11
'尋找工作表
strSheet = "KA"
For i = 1 To .Sheets.Count
    If UCase(Rtrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
Next

'找不到新增工作表
If UCase(Rtrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_TKKPIxdeliveryDatexChannel 'KA','" & txtOrderDateS & "','" & txtOrderDateE & "' "
        
Call Confirm_Recordset_Closed(tmp_Rs)

tmp_Rs.Open str_SQL, cn
Call WriteOut_RunLog("3/4.轉出KA通路KPI")
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

'ALL
Screen.MousePointer = 11
'尋找工作表
strSheet = "ALL"
For i = 1 To .Sheets.Count
    If UCase(Rtrim(.Sheets(i).Name)) = strSheet Then .Sheets(strSheet).Select: Exit For '選定工作表
Next

'找不到新增工作表
If UCase(Rtrim(.ActiveSheet.Name)) <> strSheet Then .Sheets.Add: .ActiveSheet.Name = strSheet

str_SQL = "exec gs_TKKPIxdeliveryDatexChannel 'ALL','" & txtOrderDateS & "','" & txtOrderDateE & "' "

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn
Call WriteOut_RunLog("4/4.轉出所有通路KPI")
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

'刪除DATA工作表
If chkIncData = 0 Then
    '尋找工作表
    For i = 1 To .Sheets.Count
        If UCase(Rtrim(.Sheets(i).Name)) = "DATA" Then .Sheets("DATA").Delete: Exit For
        If UCase(Rtrim(.Sheets(i).Name)) = Combo2 & Me.Caption Then .Sheets(Combo2 & Me.Caption).Delete: Exit For
    Next
End If

.Visible = True: End With

Call Unload_RunLogForm
Set MyXlsApp = Nothing
Screen.MousePointer = 0

End Sub

Private Sub cmdQuery_Click()

If Len(Rtrim(txtOrderDateS)) = 0 Or Len(Rtrim(txtOrderDateE)) = 0 Then MsgBox "請輸入日期區間", 64, "查詢": Exit Sub

On Error GoTo err_Handle
Screen.MousePointer = 11
Set dgMain.DataSource = Nothing: StatusBar.Panels(2).Text = "0 筆資料列"
Dim chc_Orderdate As String, chc_DeliveryDate As String

'str_SQL = "exec gs_TKkpi '" & Format(Now, "YYYYMMDD") & "' " & _
'            "select 廠別 = 'BL01' " & _
'            ",客戶代碼 = storerkey " & _
'            ",單別 = Type " & _
'            ",拉單日 = ymd " & _
'            ",張數 = orders " & _
'            ",店數 = consignee " & _
'            ",筆數 = orderline " & _
'            ",品項數 = sku " & _
'            ",訂貨箱數 = ordercs " & _
'            ",配貨箱數 = shipcs " & _
'            ",配貨重量 = round(shipwg,3) " & _
'            ",配貨材積 = round(shipcube,3) " & _
'            ",配貨板數 = round(shippl,3) " & _
'            ",箱 = shipcs " & _
'            ",零散 = shipea " & _
'            " From gt_kpi " & _
'            " where type in ('TK出貨單','TK退貨單') "
    
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
chc_DeliveryDate = ""
If Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and convert(Char(8),到貨日期,112) between '" & txtDeliveryDateS.Text & "' and '" & txtDeliveryDateE.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) > 0 And Len(txtDeliveryDateE.Text) = 0 Then
   chc_DeliveryDate = "and convert(Char(8),到貨日期,112) = '" & txtDeliveryDateS.Text & "' "
ElseIf Len(txtDeliveryDateS.Text) = 0 And Len(txtDeliveryDateE.Text) > 0 Then
   chc_DeliveryDate = "and convert(Char(8),到貨日期,112) = '" & txtDeliveryDateE.Text & "' "
End If

'組合字串
'str_SQL = str_SQL & chc_Orderdate & chc_DeliveryDate & " and storerkey ='" & Combo1.Text & "'order by YMD ,Type "

If Combo2 = "接單日" Then
    str_SQL = "exec gs_TKKPIxAddDate '" & txtOrderDateS & "','" & txtOrderDateE & "' "
Else
    str_SQL = "exec gs_TKKPIxDeliveryDate '" & txtOrderDateS & "','" & txtOrderDateE & "' "
End If

Call Confirm_Recordset_Closed(tmp_Rs)
tmp_Rs.Open str_SQL, cn, adOpenKeyset, adLockPessimistic
If tmp_Rs.EOF = True Then Screen.MousePointer = 0: MsgBox "查無資料！", vbOKOnly + vbInformation, Me.Caption: Exit Sub
'rsMain.Sort = "貨主,訂單號碼,項次"

Set rsMain = New ADODB.Recordset
Call OffLineRecordset(tmp_Rs, rsMain)
tmp_Rs.Close

Set dgMain.DataSource = rsMain: dgMain.Visible = False
rsMain.MoveFirst

With dgMain
Set dgMain.DataSource = rsMain

End With

SetDataGridColWidth Me.Caption, dgMain
StatusBar.Panels(2).Text = rsMain.RecordCount & " 筆資料列"
Screen.MousePointer = 0: dgMain.Visible = True

'cn.Execute "if object_id ('tempdb..##TKKPIxDeliveryDate') is not null drop table ##TKKPIxDeliveryDate", RowsAffect, adExecuteNoRecords
'cn.Execute "if object_id ('tempdb..##TKKPIxAddDate') is not null drop table ##TKKPIxAddDate", RowsAffect, adExecuteNoRecords

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
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
strA = Rtrim(rsMain("交貨單號"))
strB = Rtrim(rsMain("到貨日"))
strC = Rtrim(rsMain("到貨日"))
strD = Rtrim(rsMain("訂單號碼"))
strE = Rtrim(rsMain("項次"))
strH = Rtrim(rsMain("B"))
strI = Rtrim(rsMain("地址別"))
strJ = Rtrim(rsMain("客戶名稱"))
strK = Rtrim(rsMain("料號"))
strL = Rtrim(rsMain("訂單來源"))
strM = Rtrim(rsMain("WMS單號"))
strCheck = Rtrim(rsMain("訂單號碼")) & Rtrim(rsMain("項次"))

Do While Not rsMain.EOF

    If strCheck = Rtrim(rsMain("訂單號碼")) & Rtrim(rsMain("項次")) Then
        '同單號品項數量相加
        intF = intF + Rtrim(rsMain("數量")): strG = strG & Rtrim(rsMain("製造日")) & ";"
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
    strA = Rtrim(rsMain("交貨單號"))
    strB = Rtrim(rsMain("到貨日"))
    strC = Rtrim(rsMain("到貨日"))
    strD = Rtrim(rsMain("訂單號碼"))
    strE = Rtrim(rsMain("項次"))
    intF = Rtrim(rsMain("數量"))
    strG = Rtrim(rsMain("製造日")) & ";"
    strH = Rtrim(rsMain("B"))
    strI = Rtrim(rsMain("地址別"))
    strJ = Rtrim(rsMain("客戶名稱"))
    strK = Rtrim(rsMain("料號"))
    strL = Rtrim(rsMain("訂單來源"))
    strM = Rtrim(rsMain("WMS單號"))
    strCheck = Rtrim(rsMain("訂單號碼")) & Rtrim(rsMain("項次"))
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
    Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")
    
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim dg As Object: Set dg = dgMain
'無資料或欄寬太小，不存寬度
If Len(dg.Columns(ColIndex).DataField) < 0 Or dg.Columns(ColIndex).Width < 200 Then Exit Sub
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
Private Sub dgMain_KeyPress(KeyAscii As Integer)

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
    Combo2.ListIndex = 0
    Combo1.Text = "LTKK01"

'    txtOrderDateS.Text = Format(DateAdd("M", -1, Now), "YYYYMM") & "01"
'    txtOrderDateE.Text = Format(Now, "YYYYMMDD")

Exit Sub
err_Handle:
Call ErrorMsgbox(Me.Caption, Err.Number, Err.Description, "")

End Sub

Private Sub Form_Unload(Cancel As Integer)
cn.Execute "if object_id ('tempdb..##TKKPIxAddDate') is not null drop table ##TKKPIxAddDate ", RowsAffect, adExecuteNoRecords
cn.Execute "if object_id ('tempdb..##TKKPIxDeliveryDate') is not null drop table ##TKKPIxDeliveryDate ", RowsAffect, adExecuteNoRecords
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
