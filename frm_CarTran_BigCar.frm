VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_CarTran_BigCar 
   Caption         =   "並車"
   ClientHeight    =   6990
   ClientLeft      =   360
   ClientTop       =   1245
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   11055
   WindowState     =   1  '最小化
   Begin MSDataGridLib.DataGrid dg_Car 
      Height          =   1575
      Left            =   4440
      TabIndex        =   22
      Top             =   5175
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2778
      _Version        =   393216
      BackColor       =   16744576
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   2055
      Left            =   5880
      TabIndex        =   7
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txt_Type 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txt_Phone 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   18
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txt_Driver 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txt_AbleCBM 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txt_AbleWT 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txt_CarNo 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txt_CarTime 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "車種"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "電話"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "駕駛人"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "可載容積"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "可載重"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "車次"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "車輛號碼"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmd_Save 
      BackColor       =   &H00FFFFC0&
      Caption         =   "存  檔"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   120
      Width           =   1395
   End
   Begin VB.CommandButton cmd_Route 
      BackColor       =   &H00C0E0FF&
      Caption         =   "取得排車資料"
      Height          =   375
      Left            =   2400
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txt_Date 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_Route 
      Height          =   1920
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   3387
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      FocusRect       =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_RouteDetail 
      Height          =   3120
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   5503
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      FocusRect       =   2
      HighLight       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin VB.Label Label10 
      Caption         =   "CS"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "送貨日"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "WT"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "CBM"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "frm_CarTran_BigCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs_route As ADODB.Recordset
Private rs_routedetail As ADODB.Recordset
Private rs_car As ADODB.Recordset
Private disp_rsde As ADODB.Recordset
Private i, j As Integer
Private db As Connection


Private Sub cmd_route_Click()   '匯入排車資料
    dg_Route.Rows = 2
    '未派車資料
    On Error GoTo err_handle
    If Len(Trim(Me.txt_Date.Text)) = 0 Then Exit Sub
    'select wavekey,orderwt as 重量,ordercbm as 材積,areacode as 地區,custname as 客戶,custnum as 運送點 from BestWave where status <> 9 order by areacode
    str_SQL = "select T1.ROUTE_NO,T1.CASE_CNT,T1.WEIGHT,T1.VOLUMN_WEIGHT,T5.C_ROUTE_NO " & _
            "from LOGICTOWN.dbo.TRP01T T1 " & _
            "inner join LOGICTOWN.dbo.TRP05T T5  on T1.ROUTE_NO=T5.ROUTE_NO " & _
            "where Convert(Varchar,T1.DELIVERY_DATE,112)='" & Me.txt_Date.Text & "' " & _
            "and left(T1.ROUTE_NO,1) <> 'S' and len(rtrim(isnull(T5.C_ROUTE_NO,'')))=0 "
    Set rs_route = New ADODB.Recordset
    rs_route.Open str_SQL, db, adOpenForwardOnly, adLockReadOnly
    If rs_route.EOF Then
        Screen.MousePointer = vbDefault
        msg_text = "查詢結果：無庫存資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    rs_route.MoveFirst
    j = 1
    Do While Not rs_route.EOF
        i = 0
        dg_Route.Row = j
        For i = 0 To 3
            dg_Route.Col = i + 1
            dg_Route.Text = Trim(rs_route.Fields(i))
        Next
        j = j + 1
        If j > 1 Then
            With dg_Route
                .Rows = .Rows + 1
            End With
        End If
    rs_route.MoveNext
    Loop

    rs_route.Close
    Screen.MousePointer = vbDefault
Exit Sub

err_handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
   tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
   CreateErrorLog Me.Name & "並車", Me.Caption, "cmd_Route_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub



Private Sub cmd_Save_Click()
    If Len(Trim(Me.txt_CarNo.Text)) = 0 Then Exit Sub
    If Len(Trim(Me.txt_CarTime.Text)) = 0 Then Exit Sub
    If Len(Trim(Me.txt_Date.Text)) = 0 Then Exit Sub
    On Error GoTo err_handle
    SumCBM = 0
    sumwt = 0
    Sumcs = 0
    j = 0
    For i = 1 To dg_Route.Rows - 1
        dg_Route.Row = i
        dg_Route.Col = 0
        If dg_Route.Text = "Ｖ" Then
            j = j + 1
            dg_Route.Col = 4
            SumCBM = SumCBM + dg_Route.Text
            dg_Route.Col = 3
            sumwt = sumwt + dg_Route.Text
            dg_Route.Col = 2
            Sumcs = Sumcs + dg_Route.Text
        End If
    Next i
    If j = 0 Then Exit Sub  '沒有選任何資料
    '取得路編,"S"+"YYMMDD"+"XXX"
    str_Date = Format(Now, "YYMMDD")
    str_SQL = "select isnull(max(right(ROUTE_NO,3)),0) from LOGICTOWN.dbo.TRP01T where left(ROUTE_NO,1)='S' " & _
            "and SUBSTRING(ROUTE_NO,2,6)= '" & str_Date & "'"
    tmp_rs.Open str_SQL, db, adOpenForwardOnly, adLockReadOnly
    str_route = "S" & str_Date & StrPadLeft(Val(tmp_rs.Fields(0)) + 1, 3, 0)
    tmp_rs.Close
    '新增路編於TRP01T
    db.BeginTrans
    str_SQL = "insert into LOGICTOWN.dbo.TRP01T (ROUTE_NO,DELIVERY_DATE,CASE_CNT,WEIGHT,VOLUMN_WEIGHT,DESCRIPTION,EXE_CONFIRM) values " & _
            "('" & str_route & "','" & Me.txt_Date.Text & "','" & Sumcs & "','" & sumwt & "','" & sumwt & "','並車','0')"
    db.Execute str_SQL
    db.CommitTrans
nextadd:
    For i = 1 To dg_Route.Rows - 1
        dg_Route.Row = i
        dg_Route.Col = 0
        If dg_Route.Text = "Ｖ" Then
            With dg_Route
                dg_Route.Col = 1
                '更新TRP05T資料,寫入路編;車號;車次
                db.BeginTrans
                str_SQL = "update LOGICTOWN.dbo.TRP05T set C_ROUTE_NO='" & str_route & "', " & _
                        "C_VEHICLE_ID_NO='" & Me.txt_CarNo.Text & "',C_DRIVE_TIMES='" & Me.txt_CarTime.Text & "' " & _
                        "Where ROUTE_NO= '" & Me.dg_Route.Text & "'"
                db.Execute str_SQL
                db.CommitTrans
                '下面往上補
                For k = i To .Rows - 2
                    .Row = k
                    For j = 0 To .Cols - 1
                    .Col = j
                    .Text = .TextArray((.Row + 1) * .Cols + .Col)
                    Next j
                Next k
                .Rows = .Rows - 1   '會有多一行空白列
            End With
            GoTo nextadd    '因為資料已由下往上補,所以需重新跑一遍
        End If
    Next i
    Call clear_RouteDetail
    msg_text = "完成"
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
err_handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
   tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
   CreateErrorLog Me.Name & "並車", Me.Caption, "cmd_Save_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub dg_Car_DblClick()
    Me.txt_CarNo.Text = rs_car.Fields(0).Value
    Me.txt_Type.Text = rs_car.Fields(1).Value
    Me.txt_AbleWT.Text = rs_car.Fields(2).Value
    Me.txt_AbleCBM.Text = rs_car.Fields(3).Value
    Me.txt_Driver.Text = rs_car.Fields(4).Value
    Me.txt_Phone.Text = rs_car.Fields(5).Value
    dg_Car.Visible = False
End Sub

Private Sub dg_route_Click()
    dg_Route.Col = 2
    If Trim(dg_Route.Text) = "" Then Exit Sub
    With dg_Route
         .Col = 0    '※
         If Len(.Text) = 0 Then
            .Text = "Ｖ"
         Else
            .Text = ""
         End If
         .Col = 0
         Dim i As Integer
         For i = 0 To .Cols - 1
             .ColSel = i
         Next i
    End With
    dg_Route.SelectionMode = flexSelectionByRow
    Call sum_route
    dg_Route.Col = 1
    Call display_RouteDetail(Trim(dg_Route.Text))
End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "並車"
End Sub

Private Sub Form_Load()
    Me.Height = 7600: Me.Width = 11500
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200
    'Me.MonthView1.Visible = False
    With Me.dg_Route
         .FixedRows = 1
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
         .ColWidth(0) = 500
         .ColWidth(1) = 900
         .ColWidth(2) = 600
         .ColWidth(3) = 700
         .ColWidth(4) = 700
         '設定列表之標題
         'Convert(Varchar,AddDate,112),Route,CarNo,CarWT,Driver,AreaStart,AreaEnd,CustNum,CustCost,DriverCost,Note from bestroute"
         .Row = 0
         .Col = 0: .Text = "選取"
         .Col = 1: .Text = "路線編號"
         .Col = 2: .Text = "箱數"
         .Col = 3: .Text = "重量"
         .Col = 4: .Text = "才積"
         '設定列表之文字對齊
         .ColAlignment(0) = flexAlignCenterCenter
         .ColAlignment(1) = flexAlignRight
         .ColAlignment(2) = flexAlignRight
         .ColAlignment(5) = flexAlignLeft
         .ColAlignment(6) = flexAlignLeft
         For sub_var1 = 0 To .Cols - 1
             .Col = sub_var1
             .CellAlignment = flexAlignLeft
         Next sub_var1
    End With
    With Me.dg_RouteDetail
         .FixedRows = 1
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
         '設定列表之標題
    
         .Row = 0
         .Col = 0: .Text = "路編"
         .Col = 1: .Text = "貨主"
         .Col = 2: .Text = "單號"
         .Col = 3: .Text = "箱數"
         .Col = 4: .Text = "重量"
         .Col = 5: .Text = "才積"
         .Col = 6: .Text = "車號"
         .Col = 7: .Text = "車次"
         .Col = 8: .Text = "公司"
         '設定列表之文字對齊
         For sub_var1 = 0 To .Cols - 1
             .Col = sub_var1
             .CellAlignment = flexAlignLeft
         Next sub_var1
    End With
    '連線
    Call ConDB
End Sub

Private Sub sum_route()
    SumCBM = 0
    sumwt = 0
    Sumcs = 0
    j = dg_Route.Row
    For i = 1 To dg_Route.Rows - 1
        dg_Route.Row = i
        dg_Route.Col = 0
        If dg_Route.Text = "Ｖ" Then
            dg_Route.Col = 4
            SumCBM = SumCBM + dg_Route.Text
            dg_Route.Col = 3
            sumwt = sumwt + dg_Route.Text
            dg_Route.Col = 2
            Sumcs = Sumcs + dg_Route.Text
        End If
    Next i
    Label12.Caption = "重量:" & Round(sumwt, 2)
    Label1.Caption = "材績:" & Round(SumCBM, 2)
    Label10.Caption = "箱數:" & Round(Sumcs, 2)
    dg_Route.Row = j
End Sub

Public Sub display_RouteDetail(str_routeno As String)
    dg_RouteDetail.Rows = 2
    '未派車資料
    On Error GoTo err_handle
    If Len(Trim(Me.txt_Date.Text)) = 0 Then Exit Sub
    'select wavekey,orderwt as 重量,ordercbm as 材積,areacode as 地區,custname as 客戶,custnum as 運送點 from BestWave where status <> 9 order by areacode
    str_SQL = "select T2.ROUTE_NO,T2.STORERKEY,T2.EXTERN,T2.CASE_CNT,T2.WEIGHT,T2.VOLUMN_WEIGHT,T2.VEHICLE_ID_NO,T2.DRIVE_TIMES,o.C_Company " & _
            "from LOGICTOWN.dbo.TRP02T T2 " & _
            "inner join LOGICTOWN.dbo.orders o on T2.EXTERN=o.ExternOrderKey where T2.ROUTE_NO='" & str_routeno & "'"
    Set rs_routedetail = New ADODB.Recordset
    rs_routedetail.Open str_SQL, db, adOpenForwardOnly, adLockReadOnly
    If rs_routedetail.EOF Then
       Screen.MousePointer = vbDefault
       msg_text = "查詢結果：無明細資料"
       MsgBox msg_text, vbOKOnly + vbInformation, msg_title
       Exit Sub
    End If
    rs_routedetail.MoveFirst
    j = 1
    Do While Not rs_routedetail.EOF
        i = 0
        dg_RouteDetail.Row = j
        For i = 0 To 8
            dg_RouteDetail.Col = i
            dg_RouteDetail.Text = Trim(rs_routedetail.Fields(i))
        Next
        j = j + 1
        If j > 1 Then
            With dg_RouteDetail
                .Rows = .Rows + 1
            End With
        End If
    rs_routedetail.MoveNext
    Loop

    rs_routedetail.Close
    Screen.MousePointer = vbDefault
Exit Sub

err_handle:
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
   tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
   CreateErrorLog Me.Name & "並車", Me.Caption, "display_RouteDetail", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
  
End Sub

Public Sub clear_RouteDetail()     '
    On Error GoTo NextError
    dg_RouteDetail.Rows = 2
    dg_RouteDetail.Row = 1
    For i = 0 To dg_RouteDetail.Cols - 1
        dg_RouteDetail.Col = i
        dg_RouteDetail.Text = ""
    Next i
    Exit Sub
NextError:
    MsgBox Err.Description
    tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
    CreateErrorLog Me.Name & "並車", Me.Caption, "clear_RouteDetail", tmpString
End Sub


Private Sub txt_CarNo_GotFocus()
    On Error GoTo NextError

    str_SQL = "select VEHICLE_ID_NO as 車號,isnull(VEHICLE_TYPE,'') as 車型,isnull(LOADING_SIZE,'0') as 可載重,isnull(MAX_CUBIC_CAPACITY,'0') as 可載容積, " & _
            "isnull(DRIVER,'') as 司機,isnull(DRIVER_PHONE,'') as 電話,isnull(DESCRIPTION,'') as 備註 from LOGICTOWN.dbo.TRP09M"
    Set rs_car = New Recordset
    rs_car.Open str_SQL, db, adOpenDynamic, adLockPessimistic
    If rs_car.EOF Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    Set dg_Car.DataSource = rs_car
    Me.dg_Car.Visible = True
    Exit Sub
NextError:
    MsgBox Err.Description
    tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
    CreateErrorLog Me.Name & "並車", Me.Caption, "txt_CarNo_GotFocus", tmpString
End Sub

Private Sub txt_CarNo_LostFocus()
    On Error GoTo NextError
    str_SQL = "select isnull(max(C_DRIVE_TIMES),'0') from LOGICTOWN.dbo.TRP01T T1 " & _
            "inner join LOGICTOWN.dbo.TRP05T T5 on T1.ROUTE_NO=T5.ROUTE_NO " & _
            "where Convert(Varchar,DELIVERY_DATE,112)='" & Me.txt_Date.Text & "' and T5.C_VEHICLE_ID_NO='" & Me.txt_CarNo.Text & "'"
    tmp_rs.Open str_SQL, db, adOpenForwardOnly, adLockReadOnly
    Me.txt_CarTime = Trim(tmp_rs.Fields(0).Value) + 1
    tmp_rs.Close
    Exit Sub
NextError:
    MsgBox Err.Description
    tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
    CreateErrorLog Me.Name & "並車", Me.Caption, "txt_CarNo_LostFocus", tmpString
End Sub


Private Function ConDB() As Boolean
'取得資料庫連線資訊
Dim objIni As vbIniFile
Dim srvName As String, dbName As String, urName As String, urPassword As String

On Error GoTo err_handle
Set objIni = New vbIniFile
Retrive_ConDBInfo = True

If Dir(striniFileName_FullPath, vbHidden + vbReadOnly) = "" Then
   Retrive_ConDBInfo = False
   funRtn_msg = "指定設定檔 [" & striniFileName_FullPath & " 不存在" & vbCrLf & _
                "請通知 系統部"
   Exit Function
End If

'指定 INI 檔案存放位置與檔案名稱
objIni.FileName = striniFileName_FullPath
'取得 Server Name , DataBase Name , Login User , Login Password
srvName = objIni.ReadData("DBCONNECT", "SERVER_NAME", "0")
dbName = objIni.ReadData("DBCONNECT", "DATABASE_NAME", "0")
urName = objIni.ReadData("DBCONNECT", "LOGIN_USER", "0")
urPassword = objIni.ReadData("DBCONNECT", "USER_PASSWORD", "0")
strDefaultStorer = objIni.ReadData("DBCONNECT", "DEFAULTSTORER", "UTL")
If objIni.ReadData("DBCONNECT", "CHECKSTORER", "0") = "0" Then
   blCheckStorer = False
Else
   blCheckStorer = True
End If
'組合連線字串
    Set db = New Connection
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=MSDASQL;driver={SQL Server};server=" & srvName & ";uid=" & urName & ";pwd=" & urPassword & ";database=" & dbName & ";"
Set objIni = Nothing

Exit Function

err_handle:
   Retrive_ConDBInfo = False
   funRtn_msg = "取得資料庫連線資訊錯誤：" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Desc:" & Err.Description
End Function


