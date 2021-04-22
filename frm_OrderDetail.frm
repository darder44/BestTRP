VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_OrderDetail 
   BorderStyle     =   1  '單線固定
   Caption         =   "訂單明細"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   10410
   StartUpPosition =   2  '螢幕中央
   Begin VB.CommandButton cmdExit 
      Caption         =   "離開"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmd2Excel 
      Caption         =   "轉Excel"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid dgMain 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4895
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
Attribute VB_Name = "frm_OrderDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMain As New ADODB.Recordset
Private intColumnIndex As Integer

Private Sub cmd2Excel_Click()

'資料排序
Recordset2Excel "訂單明細", rsMain

'..在此編輯EXCEL
If rsMain Is Nothing Then
Else
    With MyXlsApp

                
    End With
End If
Set MyXlsApp = Nothing

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub dgMain_ColResize(ByVal ColIndex As Integer, Cancel As Integer)

    If Len(dgMain.Columns(ColIndex).DataField) = 0 Then Exit Sub
    SaveSetting App.title, Me.Caption & "dgMain", dgMain.Columns(ColIndex).DataField, dgMain.Columns(ColIndex).Width
    
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

Private Sub Form_Load()
Screen.MousePointer = 11

str_SQL = "select 小單位名稱 = isnull(sp.busr1,'') " & _
            ",箱數 = round(sum(case when sp.casecnt = 0 then 0 else od.originalqty/sp.casecnt end),3) " & _
            ",板數 = round(sum(case when sp.pallet = 0 then 0 else od.originalqty/sp.pallet end),3) " & _
            ",總材積 = round(sum(isnull(sp.stdcube,0)*od.originalqty),3) " & _
            ",總重量 = round(sum(isnull(sp.stdgrosswgt,0)*od.originalqty),3) " & _
            ",總個數 = sum(od.originalqty) " & _
            "from orderdetail od join gv_skuxpack sp on sp.storerkey = od.storerkey and sp.sku = od.sku " & _
            "where orderkey = '" & RTrim(frm_OP_TRPPlan.txtReceipt_no) & "' " & _
            "group by isnull(sp.busr1,'') "

Call Confirm_Recordset_Closed(tmp_rs)
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

Call Replication_Recordset(tmp_rs, rsMain)

If Not rsMain.EOF Then rsMain.MoveFirst
Set dgMain.DataSource = rsMain

'取欄位寬度
SetDataGridColWidth Me.Caption, dgMain

Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsMain = Nothing
End Sub
