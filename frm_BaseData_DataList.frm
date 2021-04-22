VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_BaseData_DataList 
   Caption         =   "資料值選取....."
   ClientHeight    =   4440
   ClientLeft      =   2040
   ClientTop       =   2445
   ClientWidth     =   9255
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fam_DataLoading 
      Caption         =   "資料載入進度"
      Height          =   1380
      Left            =   1260
      TabIndex        =   11
      Top             =   1170
      Visible         =   0   'False
      Width           =   6540
      Begin VB.TextBox txt_DataLoading 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   13
         Top             =   840
         Width           =   6300
      End
      Begin MSComctlLib.ProgressBar pb_DataLoading 
         Height          =   420
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   741
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton cmd_LoadData 
      BackColor       =   &H00FF8080&
      Caption         =   "資料載入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   3975
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmd_Query 
      BackColor       =   &H0080FF80&
      Height          =   345
      Left            =   8415
      Picture         =   "frm_BaseData_DataList.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   9
      Top             =   4080
      Width           =   345
   End
   Begin VB.TextBox txt_Query 
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
      Left            =   7155
      TabIndex        =   8
      Top             =   4080
      Width           =   1260
   End
   Begin VB.ComboBox cmb_Query 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5760
      Style           =   2  '單純下拉式
      TabIndex        =   6
      Top             =   4080
      Width           =   1410
   End
   Begin VB.CommandButton cmd_OrderBy 
      Height          =   345
      Left            =   4410
      Picture         =   "frm_BaseData_DataList.frx":058A
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   4050
      Width           =   360
   End
   Begin VB.ComboBox cmb_OrderBy 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   2
      Left            =   3165
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   4080
      Width           =   1275
   End
   Begin VB.ComboBox cmb_OrderBy 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   1
      Left            =   1905
      Style           =   2  '單純下拉式
      TabIndex        =   3
      Top             =   4080
      Width           =   1275
   End
   Begin VB.ComboBox cmb_OrderBy 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   0
      Left            =   645
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   4080
      Width           =   1275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_DataList 
      Height          =   3915
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   6906
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "尋找"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   1
      Left            =   5160
      TabIndex        =   7
      Top             =   4095
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "排序"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   4110
      Width           =   510
   End
End
Attribute VB_Name = "frm_BaseData_DataList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private arFieldName() As String      '欄位名稱
Private dLoopVar1 As Double          '迴圈變數
Private dLoopVar2 As Double          '迴圈變數

Private blEventFlag As Boolean       '事件是否有效

Private Sub cmb_Query_Click()
'資料尋找欄位指定
If blEventFlag Then
   txt_Query.SelStart = 0: txt_Query.SelLength = Len(txt_Query.Text)
   txt_Query.SetFocus
End If
End Sub

Private Sub cmd_LoadData_Click()
'載入資料
Select Case UCase(strDataList_Caller)
    Case "FRM_OP_MANUALORDERS_CMD_CONSIGNEELIST"
         '客戶資料量過大，由使用者自行決定是否載入
         blEventFlag = False
         Screen.MousePointer = vbHourglass
         Call frm_OP_ManualOrders_cmd_ConsigneeList
         blEventFlag = True
         Screen.MousePointer = vbDefault
    Case "FRM_OP_MANUALORDERS_CMDSHIPTOLIST"
         '客戶資料量過大，由使用者自行決定是否載入
         blEventFlag = False
         Screen.MousePointer = vbHourglass
         Call frm_OP_ManualOrders_cmdShipToList
         blEventFlag = True
         Screen.MousePointer = vbDefault
End Select

End Sub

Private Sub cmd_OrderBy_Click()
'排序
dg_DataList.Visible = False
dg_DataList.Sort = 9   '自訂
dg_DataList.Visible = True
End Sub

Private Sub cmd_Query_Click()
'尋找
If cmb_Query.Text = "" Then Exit Sub
If Len(Trim(txt_Query.Text)) = 0 Then Exit Sub

'依所查詢欄位排序
cmb_OrderBy(0).ListIndex = cmb_Query.ListIndex
cmb_OrderBy(1).ListIndex = -1
cmb_OrderBy(2).ListIndex = -1
Call cmd_OrderBy_Click

txt_Query.Text = Trim(txt_Query.Text)
With dg_DataList
     .Visible = False
     .Col = cmb_Query.ListIndex
     For dLoopVar1 = 1 To .Rows - 2
         .Row = dLoopVar1
         '字元型態資料尋找
         If Fun_ChkNumber(Trim(.Text)) = 1 Then
            If InStr(.Text, txt_Query.Text) > 0 Then
               .Visible = True
               .TopRow = dLoopVar1
               .LeftCol = cmb_Query.ListIndex
               Call dg_DataList_Click
               Exit Sub
            End If
          Else
          '數字型態資料尋找
            If .Text = txt_Query.Text Then
               .Visible = True
               .TopRow = dLoopVar1
               .LeftCol = cmb_Query.ListIndex
               Call dg_DataList_Click
               Exit Sub
            End If
          End If
     Next dLoopVar1
     .Visible = True
     msg_text = "狀態回報：找不到符合條件之資料"
     MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End With

End Sub

Private Sub dg_DataList_Click()
'點一次：選取，點第二次：取消選取
Dim i As Double
With dg_DataList
     .Col = 0   '編號
     If Len(Trim(.Text)) = 0 Then Exit Sub
     .Col = 0
     For i = 0 To .Cols - 1
         .ColSel = i
     Next i
End With
End Sub

Private Sub dg_DataList_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
'自訂排序
Dim strValue1 As String, strValue2 As String
strValue1 = "": strValue2 = ""
With dg_DataList
     .Row = Row1
     For dLoopVar1 = 0 To cmb_OrderBy.Count - 1
         If Trim(cmb_OrderBy(dLoopVar1).Text) <> "" Then
            .Col = cmb_OrderBy(dLoopVar1).ListIndex
            If Fun_ChkNumber(Trim(.Text)) = 1 Then
               strValue1 = strValue1 & StrPadRight(Trim(.Text), 60, " ")
            Else
               strValue1 = strValue1 & StrPadLeft(.Text, 60, "0")
            End If
         End If
     Next dLoopVar1
     
     .Row = Row2
     For dLoopVar2 = 0 To cmb_OrderBy.Count - 1
         If Trim(cmb_OrderBy(dLoopVar2).Text) <> "" Then
            .Col = cmb_OrderBy(dLoopVar2).ListIndex
            If Fun_ChkNumber(Trim(.Text)) = 1 Then
               strValue2 = strValue2 & StrPadRight(Trim(.Text), 60, " ")
            Else
               strValue2 = strValue2 & StrPadLeft(.Text, 60, "0")
            End If
         End If
     Next dLoopVar2
     
     strValue1 = Trim(strValue1)
     strValue2 = Trim(strValue2)
     If strValue1 > strValue2 Then
        Cmp = -1
     ElseIf strValue1 < strValue2 Then
        Cmp = 1
     Else
        Cmp = 0
     End If
End With


End Sub

Private Sub dg_DataList_DblClick()
'DoubleClick >> 選取資料並將資料傳回至呼叫者
With dg_DataList
     .Col = 0   '編號
     If Len(Trim(.Text)) = 0 Then Exit Sub
     Call ReturnToCaller
End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'攔截整個表單鍵盤按鍵事件
'用途：使用者按下 Esc 則不傳回任何資料，且關閉視窗
If KeyCode = vbKeyEscape Then
   Select Case UCase(strDataList_Caller)
          Case "FRM_OP_TRPPLAN_CMD_TAB0_SELECTCAR1", "FRM_OP_TRPPLAN_CMD_TAB0_SELECTCAR2"
               frm_OP_TRPPlan.WindowState = 2
          Case "FRM_OP_DCROUTEMERGE_CMD_TAB0_SELECTCAR1", "FRM_OP_DCROUTEMERGE_CMD_TAB0_SELECTCAR2"
               frm_OP_DCRouteMerge.WindowState = 2
          Case "FRM_OP_MANUALORDERS_CMD_CONSIGNEELIST"
               frm_OP_ManualOrders.WindowState = 2
          Case "FRM_OP_ROUTEDATA_CMD_SELECTCAR"
               frm_OP_RouteData.WindowState = 2
   End Select
   Unload Me
End If
End Sub

Private Sub Form_Load()
'設定 Form 大小、位置
Me.Height = 5000: Me.Width = 8900
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

blEventFlag = False
Select Case UCase(strDataList_Caller)
    Case "FRM_OP_TRPPLAN_CMD_TAB0_SELECTCAR1", "FRM_OP_TRPPLAN_CMD_TAB0_SELECTCAR2"
         '排車處理作業 >> 排車作業 >> 司機資料選取
         'Form_Name：frm_OP_TRPPlan
         msg_title = "運送之車輛選取..."
         Me.Caption = "請選取執行運送之車輛......."
         Call frm_OP_TRPPlan_cmd_Tab0_SelectCar
    Case "FRM_OP_DCROUTEMERGE_CMD_TAB0_SELECTCAR1", "FRM_OP_DCROUTEMERGE_CMD_TAB0_SELECTCAR2"
         '排車處理作業 >> DC 併車作業 >> 司機資料選取
         'Form_Name：frm_OP_DCRouteMerge
         msg_title = "運送之車輛選取..."
         Me.Caption = "請選取執行運送之車輛......."
         Call frm_OP_DCRouteMerge_cmd_Tab0_SelectCar
    Case "FRM_OP_MANUALORDERS_CMDSHIPTOLIST"
         '訂單維護作業 >> 轉運到貨客戶資料選取
         'Form_Name：frm_OP_ManualOrders
         msg_title = "訂單之客戶選取..."
         Me.Caption = "請選取訂單之客戶......."
         '客戶資料量過大，載入費時，由使用者決定是否載入
         cmb_OrderBy(2).Visible = False
         cmd_OrderBy.Left = cmd_OrderBy.Left - cmb_OrderBy(2).Width
         cmd_LoadData.Visible = True
    Case "FRM_OP_MANUALORDERS_CMD_CONSIGNEELIST"
         '訂單維護作業 >> 客戶司機資料選取
         'Form_Name：frm_OP_ManualOrders
         msg_title = "訂單之客戶選取..."
         Me.Caption = "請選取訂單之客戶......."
         '客戶資料量過大，載入費時，由使用者決定是否載入
         cmb_OrderBy(2).Visible = False
         cmd_OrderBy.Left = cmd_OrderBy.Left - cmb_OrderBy(2).Width
         cmd_LoadData.Visible = True
    Case "FRM_OP_ROUTEDATA_CMD_SELECTCAR"
         '排車處理作業 >> 路線編號維護作業 >> 司機資料選取
         'Form_Name：frm_OP_TRPPlan
         msg_title = "運送之車輛選取..."
         Me.Caption = "請選取執行運送之車輛......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB0_SELECTCAR02" '一定要大寫
         '排車處理作業 >> 出車確認 >> 司機資料選取
         'Form_Name：frm_OP_RouteConfirm
         msg_title = "運送之車輛選取..."
         Me.Caption = "請選取執行運送之車輛......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB0_SELECTCAR12" '一定要大寫
         '排車處理作業 >> 出車確認 >> 司機資料選取
         'Form_Name：frm_OP_RouteConfirm
         msg_title = "運送之車輛選取..."
         Me.Caption = "請選取執行運送之車輛......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB1_SELECTCAR2" '一定要大寫
         '排車處理作業 >> 出車確認 >> 司機資料選取
         'Form_Name：frm_OP_RouteConfirm
         msg_title = "運送之車輛選取..."
         Me.Caption = "請選取執行運送之車輛......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB2_SELECTCAR2" '一定要大寫
         '排車處理作業 >> 出車確認 >> 司機資料選取
         'Form_Name：frm_OP_RouteConfirm
         msg_title = "運送之車輛選取..."
         Me.Caption = "請選取執行運送之車輛......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
    Case "FRM_OP_SDNCONFIRM_CMD_TAB2_SELECTCAR2" '一定要大寫
         '排車處理作業 >> 出車確認 >> 司機資料選取
         'Form_Name：frm_OP_RouteConfirm
         msg_title = "運送之車輛選取..."
         Me.Caption = "請選取執行運送之車輛......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
         'frm_OP_TRPPlan
    Case "FRM_OTHER_OPTPLAN_CMD_TAB0_SELECTCAR2" '一定要大寫
         '排車處理作業 >> 出車確認 >> 司機資料選取
         'Form_Name：frm_OP_RouteConfirm
         msg_title = "運送之車輛選取..."
         Me.Caption = "請選取執行運送之車輛......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
    Case Else
         msg_text = "傳入資料錯誤：未告知呼叫者"
         MsgBox msg_text, vbOKOnly + vbInformation, msg_title
         Unload Me
End Select
blEventFlag = True

End Sub

Private Sub txt_Query_KeyPress(KeyAscii As Integer)
'尋找條件
If KeyAscii = vbKeyReturn Then
   cmd_Query.SetFocus
End If
End Sub


Private Sub frm_OP_TRPPlan_cmd_Tab0_SelectCar()
'排車處理作業 >> 排車作業 >> 司機資料選取
'Form_Name：frm_OP_TRPPlan

'設定 DataGrid 格式
Dim sub_var1 As Integer, sub_var2 As Integer
dg_DataList.Visible = False
With dg_DataList
     .FixedRows = 1: .Cols = 11
     '設定允許整列選取
     .AllowBigSelection = True
     '設定列表之文字字型
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "新細明體": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '設定列表之欄位寬度
     .ColWidth(0) = 350
     .ColWidth(1) = 450
     .ColWidth(2) = 500
     .ColWidth(3) = 1000
     .ColWidth(4) = 750
     .ColWidth(5) = 850
     .ColWidth(6) = 1100
     .ColWidth(7) = 1500
     .ColWidth(8) = 500
     .ColWidth(9) = 2000
     .ColWidth(10) = 2000
     '設定列表之標題
     .Row = 0
     .Col = 0: .Text = "No."
     .Col = 1: .Text = "車種"
     .Col = 2: .Text = "公司"
     .Col = 3: .Text = "車牌號碼"
     .Col = 4: .Text = "可載重"
     .Col = 5: .Text = "駕駛人"
     .Col = 6: .Text = "電話"
     .Col = 7: .Text = "說明"
     .Col = 8: .Text = "車次"
     .Col = 9: .Text = "車種說明"
     .Col = 10: .Text = "車種說明"
     '設定列表之文字對齊
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignCenterCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignRightCenter
     .ColAlignment(5) = flexAlignLeftCenter
     .ColAlignment(6) = flexAlignLeftCenter
     .ColAlignment(7) = flexAlignLeftCenter
     .ColAlignment(8) = flexAlignCenterCenter
     .ColAlignment(9) = flexAlignLeftCenter
     .ColAlignment(10) = flexAlignLeftCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With

'取得運送車輛基本資料
str_SQL = "Select 車種,公司別,車牌號碼,可載重,駕駛人,電話,說明,車種代碼,公司代碼 From BaseData_TRPCarList Order by 車種"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
cn.CommandTimeout = 0   '無限期等待
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "查詢結果：無符合設定條件之運輸車輛資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End If

Do While Not tmp_rs.EOF
   With dg_DataList
       .Rows = .Rows + 1
       .Row = .Rows - 2
       .Col = 0    '序號
       .Text = .Rows - 2
       .Col = 1    '車種代碼
       .Text = tmp_rs.Fields("車種代碼").Value
       .Col = 2    '運輸公司
       .Text = tmp_rs.Fields("公司代碼").Value
       .Col = 3    '車牌號碼
       .Text = tmp_rs.Fields("車牌號碼").Value
       .Col = 4    '可載重
       .Text = tmp_rs.Fields("可載重").Value
       .Col = 5    '駕駛人
       .Text = tmp_rs.Fields("駕駛人").Value
       .Col = 6    '電話
       .Text = tmp_rs.Fields("電話").Value
       .Col = 7    '說明
       .Text = tmp_rs.Fields("說明").Value
       .Col = 8    '車次
       .Text = "？"
       .Col = 9    '車種
       .Text = tmp_rs.Fields("車種").Value
       .Col = 10    '車種
       .Text = tmp_rs.Fields("公司別").Value
  End With
  tmp_rs.MoveNext
Loop
tmp_rs.Close

If UCase(strDataList_Caller) = "FRM_OP_TRPPLAN_CMD_TAB0_SELECTCAR1" Then
   '查詢各車輛當日運送已排定之車次編號
   With dg_DataList
     For dLoopVar1 = 1 To .Rows - 2
        .Row = dLoopVar1
        .Col = 3   '車牌號碼
        str_SQL = "Select Isnull(Max(Cast(Drive_TimeS as varchar)),'') as Drive_Times From TRP05T Where Vehicle_ID_NO = '" & .Text & "' and " & _
                  "  Convert(varchar(8),Delivery_Date,112) = '" & frm_OP_TRPPlan.txt_Tab0_TRPDate.Text & "' and Route_No <> 'D'"
        tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        .Col = 8
        .Text = tmp_rs.Fields("Drive_Times").Value
        tmp_rs.Close
     Next dLoopVar1
   End With
End If

dg_DataList.Visible = True


'欄位名稱提取供使用者進行自訂排序、尋找
ReDim arFieldName(1) As String
dLoopVar2 = 0
dg_DataList.Row = 0
For dLoopVar1 = 0 To dg_DataList.Cols - 1
   dLoopVar2 = dLoopVar2 + 1           '陣列以 0 開始，因此最後ㄧ個會是空白，當成 [未設定排序]
   If UBound(arFieldName) < dLoopVar2 Then
      ReDim Preserve arFieldName(dLoopVar2) As String
   End If
   dg_DataList.Col = dLoopVar1
   arFieldName(dLoopVar1) = Trim(dg_DataList.Text)
Next dLoopVar1
For dLoopVar1 = LBound(arFieldName) To UBound(arFieldName)
    For dLoopVar2 = 0 To cmb_OrderBy.Count - 1
        cmb_OrderBy(dLoopVar2).AddItem arFieldName(dLoopVar1)
    Next dLoopVar2
    cmb_Query.AddItem arFieldName(dLoopVar1)
Next dLoopVar1
'自訂排序、尋找：預設選取最後ㄧ個：空白
For dLoopVar1 = 0 To cmb_OrderBy.Count - 1
    cmb_OrderBy(dLoopVar1).ListIndex = cmb_OrderBy(dLoopVar1).ListCount - 1
Next dLoopVar1
cmb_Query.ListIndex = cmb_Query.ListCount - 1

End Sub

Private Sub ReturnToCaller()
'排車處理作業 >> 排車作業 >> 司機資料選取
'Form_Name：frm_OP_TRPPlan
Select Case UCase(strDataList_Caller)
    Case "FRM_OP_TRPPLAN_CMD_TAB0_SELECTCAR1", "FRM_OP_TRPPLAN_CMD_TAB0_SELECTCAR2"
         '排車處理作業 >> 排車作業 >> 司機資料選取
         'Form_Name：frm_OP_TRPPlan
         With dg_DataList
              .Col = 3     '車牌號碼
              frm_OP_TRPPlan.txt_Tab0_DeliveryCarNo.Text = Trim(.Text)
              .Col = 2     '運輸公司
              frm_OP_TRPPlan.txt_Tab0_DeliveryCompany.Text = Trim(.Text)
              .Col = 5     '駕駛人
              frm_OP_TRPPlan.txt_Tab0_DeliveryDriver.Text = Trim(.Text)
              .Col = 6     '電話
              frm_OP_TRPPlan.txt_Tab0_DeliveryPhone.Text = Trim(.Text)
              .Col = 1     '車種
              frm_OP_TRPPlan.txt_Tab0_DeliveryCarType.Text = Trim(.Text)
         End With
         frm_OP_TRPPlan.WindowState = 2   '最大化
    Case "FRM_OP_DCROUTEMERGE_CMD_TAB0_SELECTCAR1", "FRM_OP_DCROUTEMERGE_CMD_TAB0_SELECTCAR2"
         '排車處理作業 >> DC併車作業 >> 司機資料選取
         'Form_Name：frm_OP_DCRouteMerge
         With dg_DataList
              .Col = 3     '車牌號碼
              frm_OP_DCRouteMerge.txt_Tab0_DeliveryCarNo.Text = Trim(.Text)
              .Col = 2     '運輸公司
              frm_OP_DCRouteMerge.txt_Tab0_DeliveryCompany.Text = Trim(.Text)
              .Col = 5     '駕駛人
              frm_OP_DCRouteMerge.txt_Tab0_DeliveryDriver.Text = Trim(.Text)
              .Col = 6     '電話
              frm_OP_DCRouteMerge.txt_Tab0_DeliveryPhone.Text = Trim(.Text)
              .Col = 1     '車種
              frm_OP_DCRouteMerge.txt_Tab0_DeliveryCarType.Text = Trim(.Text)
              .Col = 1     '車種代碼
              frm_OP_DCRouteMerge.txt_Tab0_DeliveryCarTypeCode.Text = Trim(.Text)
         End With
         frm_OP_DCRouteMerge.WindowState = 2
    Case "FRM_OTHER_OPTPLAN_CMD_TAB0_SELECTCAR2"
         '退貨排車 >> 退貨排車 >> 司機資料選取
         'Form_Name：frm_OP_TRPPlan
         With dg_DataList
              .Col = 3     '車牌號碼
              frm_Other_OPTPlan.txt_Tab0_DeliveryCarNo.Text = Trim(.Text)
              .Col = 2     '運輸公司
              frm_Other_OPTPlan.txt_Tab0_DeliveryCompany.Text = Trim(.Text)
              .Col = 5     '駕駛人
              frm_Other_OPTPlan.txt_Tab0_DeliveryDriver.Text = Trim(.Text)
              .Col = 6     '電話
              frm_Other_OPTPlan.txt_Tab0_DeliveryPhone.Text = Trim(.Text)
              .Col = 1     '車種
              frm_Other_OPTPlan.txt_Tab0_DeliveryCarType.Text = Trim(.Text)
         End With
         frm_Other_OPTPlan.WindowState = 2   '最大化
    Case "FRM_OP_MANUALORDERS_CMD_CONSIGNEELIST"
         '訂單維護作業 >> 客戶資料選取
         'Form_Name：frm_OP_ManualOrders
          With dg_DataList
               .Col = 1    '客戶編號
               frm_OP_ManualOrders.txt_ConsigneeKey.Text = .Text
               Call frm_OP_ManualOrders.txt_ConsigneeKey_LostFocus
          End With
          frm_OP_ManualOrders.WindowState = 2
    Case "FRM_OP_MANUALORDERS_CMDSHIPTOLIST"
         '訂單維護作業 >> 轉運到貨客戶資料選取
         'Form_Name：frm_OP_ManualOrders
          With dg_DataList
               .Col = 1    '客戶編號
               frm_OP_ManualOrders.txtShipToKey.Text = .Text
               Call frm_OP_ManualOrders.txtShipToKey_LostFocus
          End With
          frm_OP_ManualOrders.WindowState = 2
    Case "FRM_OP_ROUTEDATA_CMD_SELECTCAR"
         '排車處理作業 >> 路線編號維護作業 >> 司機資料選取
         'Form_Name：frm_OP_RouteData
         With dg_DataList
              .Col = 3     '車牌號碼
              frm_OP_RouteData.txt_VehicleNo.Text = Trim(.Text)
              .Col = 2     '運輸公司
              frm_OP_RouteData.txt_TRPCompany.Text = Trim(.Text)
              .Col = 5     '駕駛人
              frm_OP_RouteData.txt_Driver.Text = Trim(.Text)
              .Col = 6     '電話
              frm_OP_RouteData.txt_Phone.Text = Trim(.Text)
              .Col = 1     '車種
              frm_OP_RouteData.txt_VehicleType.Text = Trim(.Text)
         End With
         frm_OP_RouteData.WindowState = 2   '最大化
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB0_SELECTCAR02" '一定要大寫"
         '排車處理作業 >> 出車確認 >> 司機資料選取
         'Form_Name：frm_OP_RouteConfirm
         With dg_DataList
              .Col = 3     '車牌號碼
              frm_OP_RouteConfirm.txt_VehicleNo0.Text = Trim(.Text)
              .Col = 5     '駕駛人
              frm_OP_RouteConfirm.txt_Driver0.Text = Trim(.Text)
         End With
         frm_OP_RouteConfirm.WindowState = 2   '最大化
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB0_SELECTCAR12" '一定要大寫"
         '排車處理作業 >> 出車確認 >> 司機資料選取
         'Form_Name：frm_OP_RouteConfirm
         With dg_DataList
              .Col = 3     '車牌號碼
              frm_OP_RouteConfirm.txt_VehicleNo1.Text = Trim(.Text)
              .Col = 5     '駕駛人
              frm_OP_RouteConfirm.txt_Driver1.Text = Trim(.Text)
         End With
         frm_OP_RouteConfirm.WindowState = 2   '最大化
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB1_SELECTCAR2" '一定要大寫"
         '排車處理作業 >> 出車確認 >> 司機資料選取
         'Form_Name：frm_OP_RouteConfirm
         With dg_DataList
              .Col = 3     '車牌號碼
              frm_OP_RouteConfirm.txt_Tab1_VehicleNo.Text = Trim(.Text)
              .Col = 5     '駕駛人
              frm_OP_RouteConfirm.txt_Tab1_Driver0.Text = Trim(.Text)
         End With
         frm_OP_RouteConfirm.WindowState = 2   '最大化
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB2_SELECTCAR2" '一定要大寫"
         '排車處理作業 >> 出車確認 >> 司機資料選取
         'Form_Name：frm_OP_RouteConfirm
         With dg_DataList
              .Col = 3     '車牌號碼
              frm_OP_RouteConfirm.txt_Tab2_VehicleNo.Text = Trim(.Text)
              .Col = 5     '駕駛人
              frm_OP_RouteConfirm.txt_Tab2_Driver.Text = Trim(.Text)
         End With
         frm_OP_RouteConfirm.WindowState = 2   '最大化
    Case "FRM_OP_SDNCONFIRM_CMD_TAB2_SELECTCAR2" '一定要大寫"
         '排車處理作業 >> 出車確認 >> 司機資料選取
         'Form_Name：frm_OP_RouteConfirm
         With dg_DataList
              .Col = 3     '車牌號碼
              frm_OP_SDNConfirm.txt_Tab02_C_VEHICLE_ID_NO.Text = Trim(.Text)
              .Col = 5     '駕駛人
              frm_OP_SDNConfirm.txt_Tab02_Driver.Text = Trim(.Text)
              frm_OP_SDNConfirm.txt_Tab02_Receiver.Text = Trim(.Text)
              frm_OP_SDNConfirm.NextPositionTab2Detail 1, 2
         End With
         frm_OP_SDNConfirm.WindowState = 2   '最大化
    Case Else
         msg_text = "未指明呼叫者，資料不知要傳回給誰"
         MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Select
Unload Me
End Sub

Private Sub frm_OP_DCRouteMerge_cmd_Tab0_SelectCar()
'排車處理作業 >> DC 併車作業 >> 司機資料選取
'Form_Name：frm_OP_DCRouteMerge

'設定 DataGrid 格式
Dim sub_var1 As Integer, sub_var2 As Integer
dg_DataList.Visible = False
With dg_DataList
     .FixedRows = 1: .Cols = 10
     '設定允許整列選取
     .AllowBigSelection = True
     '設定列表之文字字型
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "新細明體": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '設定列表之欄位寬度
     .ColWidth(0) = 350
     .ColWidth(1) = 500
     .ColWidth(2) = 700
     .ColWidth(3) = 1000
     .ColWidth(4) = 750
     .ColWidth(5) = 850
     .ColWidth(6) = 1100
     .ColWidth(7) = 1500
     .ColWidth(8) = 500
     .ColWidth(9) = 2000
     '設定列表之標題
     .Row = 0
     .Col = 0: .Text = "No."
     .Col = 1: .Text = "車種"
     .Col = 2: .Text = "公司"
     .Col = 3: .Text = "車牌號碼"
     .Col = 4: .Text = "可載重"
     .Col = 5: .Text = "駕駛人"
     .Col = 6: .Text = "電話"
     .Col = 7: .Text = "說明"
     .Col = 8: .Text = "車次"
     .Col = 9: .Text = "車種說明"
     '設定列表之文字對齊
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignLeftCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignRightCenter
     .ColAlignment(5) = flexAlignLeftCenter
     .ColAlignment(6) = flexAlignLeftCenter
     .ColAlignment(7) = flexAlignLeftCenter
     .ColAlignment(8) = flexAlignCenterCenter
     .ColAlignment(9) = flexAlignCenterCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With

'取得運送車輛基本資料
str_SQL = "Select 車種,公司別,車牌號碼,可載重,駕駛人,電話,說明,車種代碼 From BaseData_TRPCarList Order by 車種"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
cn.CommandTimeout = 0   '無限期等待
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "查詢結果：無符合設定條件之運輸車輛資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End If

Do While Not tmp_rs.EOF
   With dg_DataList
       .Rows = .Rows + 1
       .Row = .Rows - 2
       .Col = 0    '序號
       .Text = .Rows - 2
       .Col = 1    '車種代碼
       .Text = tmp_rs.Fields("車種代碼").Value
       .Col = 2    '運輸公司
       .Text = tmp_rs.Fields("公司別").Value
       .Col = 3    '車牌號碼
       .Text = tmp_rs.Fields("車牌號碼").Value
       .Col = 4    '可載重
       .Text = tmp_rs.Fields("可載重").Value
       .Col = 5    '駕駛人
       .Text = tmp_rs.Fields("駕駛人").Value
       .Col = 6    '電話
       .Text = tmp_rs.Fields("電話").Value
       .Col = 7    '說明
       .Text = tmp_rs.Fields("說明").Value
       .Col = 8    '車次
       .Text = "？"
       .Col = 9    '車種代碼
       .Text = tmp_rs.Fields("車種").Value
  End With
  tmp_rs.MoveNext
Loop
tmp_rs.Close

If UCase(strDataList_Caller) = "FRM_OP_DCROUTEMERGE_CMD_TAB0_SELECTCAR1" Then
   '查詢各車輛當日運送已排定之車次編號
   With dg_DataList
     For dLoopVar1 = 1 To .Rows - 2
        .Row = dLoopVar1
        .Col = 3   '車牌號碼
        str_SQL = "Select Isnull(Max(Cast(Drive_TimeS as varchar)),'') as Drive_Times From TRP05T Where Vehicle_ID_NO = '" & .Text & "' and " & _
                  "  Convert(varchar(8),Delivery_Date,112) = '" & frm_OP_DCRouteMerge.txt_Tab0_TRPDate.Text & "' and Route_No <> 'D'"
        tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
        .Col = 8
        .Text = tmp_rs.Fields("Drive_Times").Value
        tmp_rs.Close
     Next dLoopVar1
   End With
End If
dg_DataList.Visible = True


'欄位名稱提取供使用者進行自訂排序、尋找
ReDim arFieldName(1) As String
dLoopVar2 = 0
dg_DataList.Row = 0
For dLoopVar1 = 0 To dg_DataList.Cols - 1
   dLoopVar2 = dLoopVar2 + 1           '陣列以 0 開始，因此最後ㄧ個會是空白，當成 [未設定排序]
   If UBound(arFieldName) < dLoopVar2 Then
      ReDim Preserve arFieldName(dLoopVar2) As String
   End If
   dg_DataList.Col = dLoopVar1
   arFieldName(dLoopVar1) = Trim(dg_DataList.Text)
Next dLoopVar1
For dLoopVar1 = LBound(arFieldName) To UBound(arFieldName)
    For dLoopVar2 = 0 To cmb_OrderBy.Count - 1
        cmb_OrderBy(dLoopVar2).AddItem arFieldName(dLoopVar1)
    Next dLoopVar2
    cmb_Query.AddItem arFieldName(dLoopVar1)
Next dLoopVar1
'自訂排序、尋找：預設選取最後ㄧ個：空白
For dLoopVar1 = 0 To cmb_OrderBy.Count - 1
    cmb_OrderBy(dLoopVar1).ListIndex = cmb_OrderBy(dLoopVar1).ListCount - 1
Next dLoopVar1
cmb_Query.ListIndex = cmb_Query.ListCount - 1

End Sub

Private Sub frm_OP_ManualOrders_cmd_ConsigneeList()
'訂單維護作業 >> 客戶資料選取
'Form_Name：frm_OP_ManualOrders

'設定 DataGrid 格式
Dim sub_var1 As Integer, sub_var2 As Integer
dg_DataList.Visible = False
With dg_DataList
     .FixedRows = 1: .Cols = 15
     '設定允許整列選取
     .AllowBigSelection = True
     '設定列表之文字字型
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "新細明體": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '設定列表之欄位寬度
     .ColWidth(0) = 400
     .ColWidth(1) = 1500
     .ColWidth(2) = 1500
     .ColWidth(3) = 1500
     .ColWidth(4) = 800
     .ColWidth(5) = 4000
     .ColWidth(6) = 3000
     .ColWidth(7) = 1500
     .ColWidth(9) = 1500
     .ColWidth(10) = 1500
     .ColWidth(11) = 300
     .ColWidth(12) = 300
     .ColWidth(13) = 300
     .ColWidth(14) = 300

     '設定列表之標題
     .Row = 0
     .Col = 0: .Text = "No."
     .Col = 1: .Text = "客戶編號"
     .Col = 2: .Text = "客戶名稱"
     .Col = 3: .Text = "客戶簡稱"
     .Col = 4: .Text = "郵遞區號"
     .Col = 5: .Text = "運送地址"
     .Col = 6: .Text = "運送區域"
     .Col = 7: .Text = "特殊需求-1"
     .Col = 8: .Text = "特殊需求-2"
     .Col = 9: .Text = "聯絡人"
     .Col = 10: .Text = "電話"
     .Col = 11: .Text = "運送區域碼"
     .Col = 12: .Text = "郵遞區號碼"
     .Col = 13: .Text = "特殊需求1"
     .Col = 14: .Text = "特殊需求2"
     '設定列表之文字對齊
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignLeftCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignCenterCenter
     .ColAlignment(5) = flexAlignLeftCenter
     .ColAlignment(6) = flexAlignLeftCenter
     .ColAlignment(7) = flexAlignLeftCenter
     .ColAlignment(8) = flexAlignLeftCenter
     .ColAlignment(9) = flexAlignLeftCenter
     .ColAlignment(10) = flexAlignLeftCenter
     .ColAlignment(11) = flexAlignCenterCenter
     .ColAlignment(12) = flexAlignCenterCenter
     .ColAlignment(13) = flexAlignCenterCenter
     .ColAlignment(14) = flexAlignCenterCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With

Dim dbTotal As Double, dbNow As Double, strStorerkey As String
fam_DataLoading.Visible = True

strStorerkey = mySplit(frm_OP_ManualOrders.cmbStorerkey, " ", 0)

'取貨主
If Len(RTrim(strStorerkey)) = 0 Then
    str_SQL = "Select count(*) as RecCount From TRP01M"
Else
    str_SQL = "Select count(*) as RecCount From TRP01M where storerkey = '" & strStorerkey & "' "
End If

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
cn.CommandTimeout = 0   '無限期等待
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "查詢結果：無符合設定條件之客戶資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
Else
   dbTotal = tmp_rs.Fields("RecCount").Value
   pb_DataLoading.Max = dbTotal
End If
tmp_rs.Close

'貨主條件
If Len(RTrim(strStorerkey)) = 0 Then
    strStorerkey = ""
Else
    strStorerkey = "where 貨主編號 = '" & strStorerkey & "' "
End If
   
'取得客戶基本資料
str_SQL = "Select 客戶編號,客戶名稱,客戶簡稱,郵遞區號,運送地址,運送區域,特殊需求1,特殊需求2,聯絡人,電話," & _
          "  運送區域碼,郵遞區號碼,特殊需求碼1,特殊需求碼2 " & _
          "From BaseData_ConsigneeList " & strStorerkey & " Order by 客戶編號"
cn.CommandTimeout = 0   '無限期等待
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "查詢結果：無符合設定條件之客戶資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

dbNow = 0
Do While Not tmp_rs.EOF
   dbNow = dbNow + 1
   pb_DataLoading.Value = dbNow
   txt_DataLoading.Text = "客戶資料共 " & dbTotal & " 已載入 " & dbNow & " 筆"
   DoEvents
   With dg_DataList
       .Rows = .Rows + 1
       .Row = .Rows - 2
       .Col = 0    '序號
       .Text = .Rows - 2
       .Col = 1    '客戶編號
       .Text = tmp_rs.Fields("客戶編號").Value
       .Col = 2    '客戶名稱
       .Text = tmp_rs.Fields("客戶名稱").Value
       .Col = 3    '客戶簡稱
       .Text = tmp_rs.Fields("客戶簡稱").Value
       .Col = 4    '郵遞區號
       .Text = tmp_rs.Fields("郵遞區號").Value
       .Col = 5    '運送區域
       .Text = tmp_rs.Fields("運送地址").Value
       .Col = 6    '運送地址
       .Text = tmp_rs.Fields("運送區域").Value
       .Col = 7    '特殊需求 1
       .Text = tmp_rs.Fields("特殊需求1").Value
       .Col = 8    '特殊需求 2
       .Text = tmp_rs.Fields("特殊需求2").Value
       .Col = 9    '聯絡人
       .Text = tmp_rs.Fields("聯絡人").Value
       .Col = 10   '電話
       .Text = tmp_rs.Fields("電話").Value
       .Col = 11   '運送區域碼
       .Text = tmp_rs.Fields("運送區域碼").Value
       .Col = 12   '郵遞區號代碼
       .Text = tmp_rs.Fields("郵遞區號碼").Value
       .Col = 13   '特殊需求碼1
       .Text = tmp_rs.Fields("特殊需求碼1").Value
       .Col = 14   '特殊需求碼2
       .Text = tmp_rs.Fields("特殊需求碼2").Value
  End With
  tmp_rs.MoveNext
Loop
tmp_rs.Close
fam_DataLoading.Visible = False
dg_DataList.Visible = True

'欄位名稱提取供使用者進行自訂排序、尋找
ReDim arFieldName(1) As String
dLoopVar2 = 0
dg_DataList.Row = 0
For dLoopVar1 = 0 To dg_DataList.Cols - 1
   dLoopVar2 = dLoopVar2 + 1           '陣列以 0 開始，因此最後ㄧ個會是空白，當成 [未設定排序]
   If UBound(arFieldName) < dLoopVar2 Then
      ReDim Preserve arFieldName(dLoopVar2) As String
   End If
   dg_DataList.Col = dLoopVar1
   arFieldName(dLoopVar1) = Trim(dg_DataList.Text)
Next dLoopVar1
For dLoopVar1 = LBound(arFieldName) To UBound(arFieldName)
    For dLoopVar2 = 0 To cmb_OrderBy.Count - 1
        cmb_OrderBy(dLoopVar2).AddItem arFieldName(dLoopVar1)
    Next dLoopVar2
    cmb_Query.AddItem arFieldName(dLoopVar1)
Next dLoopVar1
'自訂排序、尋找：預設選取最後ㄧ個：空白
For dLoopVar1 = 0 To cmb_OrderBy.Count - 1
    cmb_OrderBy(dLoopVar1).ListIndex = cmb_OrderBy(dLoopVar1).ListCount - 1
Next dLoopVar1
cmb_Query.ListIndex = cmb_Query.ListCount - 1

End Sub
Private Sub frm_OP_ManualOrders_cmdShipToList()
'訂單維護作業 >> 轉運到貨客戶資料選取
'Form_Name：frm_OP_ManualOrders

'設定 DataGrid 格式
Dim sub_var1 As Integer, sub_var2 As Integer
dg_DataList.Visible = False
With dg_DataList
     .FixedRows = 1: .Cols = 15
     '設定允許整列選取
     .AllowBigSelection = True
     '設定列表之文字字型
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "新細明體": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '設定列表之欄位寬度
     .ColWidth(0) = 400
     .ColWidth(1) = 1500
     .ColWidth(2) = 1500
     .ColWidth(3) = 1500
     .ColWidth(4) = 800
     .ColWidth(5) = 4000
     .ColWidth(6) = 3000
     .ColWidth(7) = 1500
     .ColWidth(9) = 1500
     .ColWidth(10) = 1500
     .ColWidth(11) = 300
     .ColWidth(12) = 300
     .ColWidth(13) = 300
     .ColWidth(14) = 300

     '設定列表之標題
     .Row = 0
     .Col = 0: .Text = "No."
     .Col = 1: .Text = "客戶編號"
     .Col = 2: .Text = "客戶名稱"
     .Col = 3: .Text = "客戶簡稱"
     .Col = 4: .Text = "郵遞區號"
     .Col = 5: .Text = "運送地址"
     .Col = 6: .Text = "運送區域"
     .Col = 7: .Text = "特殊需求-1"
     .Col = 8: .Text = "特殊需求-2"
     .Col = 9: .Text = "聯絡人"
     .Col = 10: .Text = "電話"
     .Col = 11: .Text = "運送區域碼"
     .Col = 12: .Text = "郵遞區號碼"
     .Col = 13: .Text = "特殊需求1"
     .Col = 14: .Text = "特殊需求2"
     '設定列表之文字對齊
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignLeftCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignCenterCenter
     .ColAlignment(5) = flexAlignLeftCenter
     .ColAlignment(6) = flexAlignLeftCenter
     .ColAlignment(7) = flexAlignLeftCenter
     .ColAlignment(8) = flexAlignLeftCenter
     .ColAlignment(9) = flexAlignLeftCenter
     .ColAlignment(10) = flexAlignLeftCenter
     .ColAlignment(11) = flexAlignCenterCenter
     .ColAlignment(12) = flexAlignCenterCenter
     .ColAlignment(13) = flexAlignCenterCenter
     .ColAlignment(14) = flexAlignCenterCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With

Dim dbTotal As Double, dbNow As Double, strStorerkey As String
fam_DataLoading.Visible = True

strStorerkey = mySplit(frm_OP_ManualOrders.cmbStorerkey, " ", 0)

'取貨主
If Len(RTrim(strStorerkey)) = 0 Then
    str_SQL = "Select count(*) as RecCount From TRP01M"
Else
    str_SQL = "Select count(*) as RecCount From TRP01M where storerkey = '" & strStorerkey & "' "
End If

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
cn.CommandTimeout = 0   '無限期等待
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "查詢結果：無符合設定條件之客戶資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
Else
   dbTotal = tmp_rs.Fields("RecCount").Value
   pb_DataLoading.Max = dbTotal
End If
tmp_rs.Close
   
'貨主條件
If Len(RTrim(strStorerkey)) = 0 Then
    strStorerkey = ""
Else
    strStorerkey = "where 貨主編號 = '" & strStorerkey & "' "
End If
   
'取得客戶基本資料
str_SQL = "Select 客戶編號,客戶名稱,客戶簡稱,郵遞區號,運送地址,運送區域,特殊需求1,特殊需求2,聯絡人,電話," & _
          "  運送區域碼,郵遞區號碼,特殊需求碼1,特殊需求碼2 " & _
          "From BaseData_ConsigneeList " & strStorerkey & " Order by 客戶編號"
cn.CommandTimeout = 0   '無限期等待
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120

If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "查詢結果：無符合設定條件之客戶資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

dbNow = 0
Do While Not tmp_rs.EOF
   dbNow = dbNow + 1
   pb_DataLoading.Value = dbNow
   txt_DataLoading.Text = "客戶資料共 " & dbTotal & " 已載入 " & dbNow & " 筆"
   DoEvents
   With dg_DataList
       .Rows = .Rows + 1
       .Row = .Rows - 2
       .Col = 0    '序號
       .Text = .Rows - 2
       .Col = 1    '客戶編號
       .Text = tmp_rs.Fields("客戶編號").Value
       .Col = 2    '客戶名稱
       .Text = tmp_rs.Fields("客戶名稱").Value
       .Col = 3    '客戶簡稱
       .Text = tmp_rs.Fields("客戶簡稱").Value
       .Col = 4    '郵遞區號
       .Text = tmp_rs.Fields("郵遞區號").Value
       .Col = 5    '運送區域
       .Text = tmp_rs.Fields("運送地址").Value
       .Col = 6    '運送地址
       .Text = tmp_rs.Fields("運送區域").Value
       .Col = 7    '特殊需求 1
       .Text = tmp_rs.Fields("特殊需求1").Value
       .Col = 8    '特殊需求 2
       .Text = tmp_rs.Fields("特殊需求2").Value
       .Col = 9    '聯絡人
       .Text = tmp_rs.Fields("聯絡人").Value
       .Col = 10   '電話
       .Text = tmp_rs.Fields("電話").Value
       .Col = 11   '運送區域碼
       .Text = tmp_rs.Fields("運送區域碼").Value
       .Col = 12   '郵遞區號代碼
       .Text = tmp_rs.Fields("郵遞區號碼").Value
       .Col = 13   '特殊需求碼1
       .Text = tmp_rs.Fields("特殊需求碼1").Value
       .Col = 14   '特殊需求碼2
       .Text = tmp_rs.Fields("特殊需求碼2").Value
  End With
  tmp_rs.MoveNext
Loop
tmp_rs.Close
fam_DataLoading.Visible = False
dg_DataList.Visible = True

'欄位名稱提取供使用者進行自訂排序、尋找
ReDim arFieldName(1) As String
dLoopVar2 = 0
dg_DataList.Row = 0
For dLoopVar1 = 0 To dg_DataList.Cols - 1
   dLoopVar2 = dLoopVar2 + 1           '陣列以 0 開始，因此最後ㄧ個會是空白，當成 [未設定排序]
   If UBound(arFieldName) < dLoopVar2 Then
      ReDim Preserve arFieldName(dLoopVar2) As String
   End If
   dg_DataList.Col = dLoopVar1
   arFieldName(dLoopVar1) = Trim(dg_DataList.Text)
Next dLoopVar1
For dLoopVar1 = LBound(arFieldName) To UBound(arFieldName)
    For dLoopVar2 = 0 To cmb_OrderBy.Count - 1
        cmb_OrderBy(dLoopVar2).AddItem arFieldName(dLoopVar1)
    Next dLoopVar2
    cmb_Query.AddItem arFieldName(dLoopVar1)
Next dLoopVar1
'自訂排序、尋找：預設選取最後ㄧ個：空白
For dLoopVar1 = 0 To cmb_OrderBy.Count - 1
    cmb_OrderBy(dLoopVar1).ListIndex = cmb_OrderBy(dLoopVar1).ListCount - 1
Next dLoopVar1
cmb_Query.ListIndex = cmb_Query.ListCount - 1

End Sub
Private Sub frm_OP_ROUTEDATA_cmd_SelectCar()
'排車處理作業 >> 路線編號維護作業 >> 司機資料選取
'Form_Name：frm_OP_RouteData

'設定 DataGrid 格式
Dim sub_var1 As Integer, sub_var2 As Integer
dg_DataList.Visible = False
With dg_DataList
     .FixedRows = 1: .Cols = 8
     '設定允許整列選取
     .AllowBigSelection = True
     '設定列表之文字字型
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "新細明體": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '設定列表之欄位寬度
     .ColWidth(0) = 350
     .ColWidth(1) = 1600
     .ColWidth(2) = 700
     .ColWidth(3) = 1000
     .ColWidth(4) = 750
     .ColWidth(5) = 850
     .ColWidth(6) = 1100
     .ColWidth(7) = 1500
     '設定列表之標題
     .Row = 0
     .Col = 0: .Text = "No."
     .Col = 1: .Text = "車種"
     .Col = 2: .Text = "公司"
     .Col = 3: .Text = "車牌號碼"
     .Col = 4: .Text = "可載重"
     .Col = 5: .Text = "駕駛人"
     .Col = 6: .Text = "電話"
     .Col = 7: .Text = "說明"
     '設定列表之文字對齊
     .ColAlignment(0) = flexAlignCenterCenter
     .ColAlignment(1) = flexAlignLeftCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignLeftCenter
     .ColAlignment(4) = flexAlignRightCenter
     .ColAlignment(5) = flexAlignLeftCenter
     .ColAlignment(6) = flexAlignLeftCenter
     .ColAlignment(7) = flexAlignLeftCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With

'取得運送車輛基本資料
str_SQL = "Select 車種,公司別,車牌號碼,可載重,駕駛人,電話,說明 From BaseData_TRPCarList Order by 車種"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
cn.CommandTimeout = 0   '無限期等待
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "查詢結果：無符合設定條件之運輸車輛資料"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End If

Do While Not tmp_rs.EOF
   With dg_DataList
       .Rows = .Rows + 1
       .Row = .Rows - 2
       .Col = 0    '序號
       .Text = .Rows - 2
       .Col = 1    '車種代碼
       .Text = tmp_rs.Fields("車種").Value
       .Col = 2    '運輸公司
       .Text = tmp_rs.Fields("公司別").Value
       .Col = 3    '車牌號碼
       .Text = tmp_rs.Fields("車牌號碼").Value
       .Col = 4    '可載重
       .Text = tmp_rs.Fields("可載重").Value
       .Col = 5    '駕駛人
       .Text = tmp_rs.Fields("駕駛人").Value
       .Col = 6    '電話
       .Text = tmp_rs.Fields("電話").Value
       .Col = 7    '說明
       .Text = tmp_rs.Fields("說明").Value
  End With
  tmp_rs.MoveNext
Loop
tmp_rs.Close
dg_DataList.Visible = True


'欄位名稱提取供使用者進行自訂排序、尋找
ReDim arFieldName(1) As String
dLoopVar2 = 0
dg_DataList.Row = 0
For dLoopVar1 = 0 To dg_DataList.Cols - 1
   dLoopVar2 = dLoopVar2 + 1           '陣列以 0 開始，因此最後ㄧ個會是空白，當成 [未設定排序]
   If UBound(arFieldName) < dLoopVar2 Then
      ReDim Preserve arFieldName(dLoopVar2) As String
   End If
   dg_DataList.Col = dLoopVar1
   arFieldName(dLoopVar1) = Trim(dg_DataList.Text)
Next dLoopVar1
For dLoopVar1 = LBound(arFieldName) To UBound(arFieldName)
    For dLoopVar2 = 0 To cmb_OrderBy.Count - 1
        cmb_OrderBy(dLoopVar2).AddItem arFieldName(dLoopVar1)
    Next dLoopVar2
    cmb_Query.AddItem arFieldName(dLoopVar1)
Next dLoopVar1
'自訂排序、尋找：預設選取最後ㄧ個：空白
For dLoopVar1 = 0 To cmb_OrderBy.Count - 1
    cmb_OrderBy(dLoopVar1).ListIndex = cmb_OrderBy(dLoopVar1).ListCount - 1
Next dLoopVar1
cmb_Query.ListIndex = cmb_Query.ListCount - 1

End Sub
