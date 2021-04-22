VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_OP_ShipQty 
   Caption         =   "揀貨數量確認"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   11475
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   12515
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14215660
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "揀貨確認"
      TabPicture(0)   =   "frm_OP_ShipQty.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "dg_Tab0_Ship"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "揀貨單"
      TabPicture(1)   =   "frm_OP_ShipQty.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "dg_Tab1_RouteOrders"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   885
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   10605
         Begin VB.CheckBox chk_Tab1_PreView 
            Caption         =   "預覽列印"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400040&
            Height          =   240
            Left            =   4080
            TabIndex        =   21
            Top             =   360
            Width           =   1425
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H00FFC0FF&
            Caption         =   "離  開"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   0
            Left            =   8280
            Style           =   1  '圖片外觀
            TabIndex        =   17
            Top             =   240
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Tab1_Print 
            Appearance      =   0  '平面
            BackColor       =   &H00C0FFC0&
            Caption         =   "列  印"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   7080
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   16
            Top             =   240
            Width           =   1200
         End
         Begin VB.TextBox txt_Tab1_RouteNo_End 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2565
            TabIndex        =   15
            Top             =   360
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab1_RouteNo_Start 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1050
            TabIndex        =   14
            Top             =   360
            Width           =   1245
         End
         Begin VB.CommandButton cmd_Tab1_Query 
            Appearance      =   0  '平面
            BackColor       =   &H00C0C000&
            Caption         =   "查  詢"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   5880
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   13
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "路線編號"
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
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   19
            Top             =   405
            Width           =   840
         End
         Begin VB.Label Label3 
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
            Index           =   0
            Left            =   2325
            TabIndex        =   18
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   -73800
         TabIndex        =   11
         Top             =   3360
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Frame Frame11 
         Height          =   885
         Left            =   -74640
         TabIndex        =   1
         Top             =   600
         Width           =   10605
         Begin VB.CheckBox ck_Tab0_ship 
            Caption         =   "未回傳揀貨數量"
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
            Height          =   195
            Left            =   3960
            TabIndex        =   9
            Top             =   405
            Width           =   1815
         End
         Begin VB.CommandButton cmd_Tab0_Query 
            Appearance      =   0  '平面
            BackColor       =   &H00C0C000&
            Caption         =   "查  詢"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   5880
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   6
            Top             =   240
            Width           =   1200
         End
         Begin VB.TextBox txt_Tab0_RouteNo_Start 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1050
            TabIndex        =   5
            Top             =   360
            Width           =   1245
         End
         Begin VB.TextBox txt_Tab0_RouteNo_End 
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.75
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2565
            TabIndex        =   4
            Top             =   360
            Width           =   1245
         End
         Begin VB.CommandButton cmd_Tab0_Save 
            Appearance      =   0  '平面
            BackColor       =   &H008080FF&
            Caption         =   "確認存檔"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   7080
            MaskColor       =   &H00FFC0C0&
            Style           =   1  '圖片外觀
            TabIndex        =   3
            Top             =   240
            Width           =   1200
         End
         Begin VB.CommandButton cmd_Exit 
            BackColor       =   &H00FFC0FF&
            Caption         =   "離  開"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   3
            Left            =   8280
            Style           =   1  '圖片外觀
            TabIndex        =   2
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label3 
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
            Index           =   21
            Left            =   2325
            TabIndex        =   8
            Top             =   360
            Width           =   240
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "路線編號"
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
            Height          =   195
            Index           =   22
            Left            =   150
            TabIndex        =   7
            Top             =   405
            Width           =   840
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dg_Tab0_Ship 
         Height          =   4560
         Left            =   -74640
         TabIndex        =   10
         Top             =   1800
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   8043
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         Cols            =   14
         FixedCols       =   0
         BackColorSel    =   10354595
         ForeColorSel    =   8454016
         BackColorBkg    =   -2147483644
         AllowBigSelection=   0   'False
         FocusRect       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   14
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   1
      End
      Begin MSDataGridLib.DataGrid dg_Tab1_RouteOrders 
         Height          =   5040
         Left            =   360
         TabIndex        =   20
         Top             =   1680
         Width           =   10620
         _ExtentX        =   18733
         _ExtentY        =   8890
         _Version        =   393216
         AllowUpdate     =   0   'False
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
   End
End
Attribute VB_Name = "frm_OP_ShipQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dbsrcFormHeight As Double                'Form 設計時期的高
Private dbsrcFormWidth As Double                 'Form 設計時期的寬
Private rs_Tab1_RouteOrders As ADODB.Recordset   '揀貨表Recordset
Private rs_Access As ADODB.Recordset             'Access Recordset

Private Sub cmd_Exit_Click(Index As Integer)
    '離開
    Unload Me
End Sub

Private Sub cmd_Tab0_Query_Click()
    '訂單查詢
    Dim ExternOrderKey As String
    Dim strSubwhere As String
    Dim str_Where As String
    On Error GoTo err_handle
    dg_Tab0_Ship.Rows = 2
    For i = 0 To 7
        dg_Tab0_Ship.Col = i
        dg_Tab0_Ship.Text = ""
    Next
    str_SQL = "select ROUTE_NO,EXTERN,PRODUCT_NO,SEQ_NO,isnull(SubSeq_No,'0') as SubSeq,ORDER_QTY,isnull(SHIP_TIME,'未回傳'),isnull(SHIP_QTY,0),isnull(SHIP_QTY,ORDER_QTY) " & _
            "From trp03t where ROUTE_NO<>'D'"

    '一次排車路線編號
    txt_Tab0_RouteNo_Start.Text = Trim(txt_Tab0_RouteNo_Start.Text)
    txt_Tab0_RouteNo_End.Text = Trim(txt_Tab0_RouteNo_End.Text)
    strSubwhere = ""
    If Len(txt_Tab0_RouteNo_Start.Text) > 0 And Len(txt_Tab0_RouteNo_End.Text) > 0 Then
        strSubwhere = " Route_No Between '" & txt_Tab0_RouteNo_Start.Text & "' and '" & txt_Tab0_RouteNo_End.Text & "' "
    ElseIf Len(txt_Tab0_RouteNo_Start.Text) > 0 And Len(txt_Tab0_RouteNo_End.Text) = 0 Then
        strSubwhere = " Route_No = '" & txt_Tab0_RouteNo_Start.Text & "' "
    ElseIf Len(txt_Tab0_RouteNo_Start.Text) = 0 And Len(txt_Tab0_RouteNo_End.Text) > 0 Then
        strSubwhere = " Route_No = '" & txt_Tab0_RouteNo_End.Text & "' "
    End If
    If Len(strSubwhere) > 0 Then
        str_Where = str_Where & " and " & strSubwhere
    End If
    
    '未確認
  
    If ck_Tab0_ship.Value = vbChecked Then
        str_Where = str_Where & " and ship_time is null "
    End If
    
    
    If Len(str_Where) = 0 Then
        str_SQL = str_SQL & " order by ROUTE_NO,EXTERN,SEQ_NO,SubSeq_No"
    Else
        str_SQL = str_SQL & str_Where & " order by ROUTE_NO,EXTERN,SEQ_NO,SubSeq_No"
    End If
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "查詢結果：不存在於排車系統"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    'Dim i, j As Integer
    j = 1
    Do While Not tmp_Rs.EOF
        i = 0
        dg_Tab0_Ship.Row = j
        For i = 0 To 8
            dg_Tab0_Ship.Col = i
            dg_Tab0_Ship.Text = Trim(tmp_Rs.Fields(i))
        Next
        '新增一行
        j = j + 1
        If j > 1 Then
            With dg_Tab0_Ship
                .Rows = .Rows + 1
            End With
        End If
        tmp_Rs.MoveNext
    Loop
    dg_Tab0_Ship.Rows = dg_Tab0_Ship.Rows - 1
    tmp_Rs.Close
    Exit Sub
    
err_handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "揀貨數量確認-查詢", Me.Caption, "cmd_Tab0_Query_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab0_Save_Click()
    On Error GoTo err_handle
    cn.BeginTrans
        '存表身:裝載明細,SDN02T
        For i = 1 To dg_Tab0_Ship.Rows - 1
        
            dg_Tab0_Ship.Row = i
            ''ROUTE_NO,EXTERN,PRODUCT_NO,SEQ_NO,isnull(SubSeq_No,'0'),ORDER_QTY,SHIP_TIME,SHIP_QTY
            '路線編號,客戶單號,貨號,項次,拆單,排車數,回傳時間,目前數量,確認數量
            dg_Tab0_Ship.Col = 0: str_ROUTE_NO = Trim(dg_Tab0_Ship.Text)
            dg_Tab0_Ship.Col = 1: str_EXTERN = Trim(dg_Tab0_Ship.Text)
            dg_Tab0_Ship.Col = 2: str_PRODUCT_NO = Trim(dg_Tab0_Ship.Text)
            dg_Tab0_Ship.Col = 3: str_SEQ_NO = Trim(dg_Tab0_Ship.Text)
            dg_Tab0_Ship.Col = 4: str_SubSeq_No = Trim(dg_Tab0_Ship.Text)
            dg_Tab0_Ship.Col = 8: str_SHIP_QTY = Trim(dg_Tab0_Ship.Text)
            dg_Tab0_Ship.Col = 1
            str_SDNStatus = 0
            str_SQL = "update trp03t set SHIP_QTY='" & str_SHIP_QTY & "' where " & _
                    "ROUTE_NO='" & str_ROUTE_NO & "' and PRODUCT_NO ='" & str_PRODUCT_NO & "' and EXTERN ='" & str_EXTERN & "' " & _
                    "and SEQ_NO='" & str_SEQ_NO & "' and isnull(SubSeq_No,'0') ='" & str_SubSeq_No & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
            str_SQL = "update sdn03t set SHIP_QTY='" & str_SHIP_QTY & "' where " & _
            "ROUTE_NO='" & str_ROUTE_NO & "' and PRODUCT_NO ='" & str_PRODUCT_NO & "' and EXTERN ='" & str_EXTERN & "' " & _
            "and SEQ_NO='" & str_SEQ_NO & "' and isnull(SubSeq_No,'0') ='" & str_SubSeq_No & "'"
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            
        Next
        
    cn.CommitTrans
    '清除dg_Tab0_Ship
    dg_Tab0_Ship.Rows = 2: dg_Tab0_Ship.Row = 1
    For i = 0 To dg_Tab0_Ship.Cols - 1
        dg_Tab0_Ship.Col = i
        dg_Tab0_Ship.Text = ""
    Next

    Exit Sub
    
err_handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "揀貨數量確認-存檔", Me.Caption, "cmd_Tab0_Save_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmd_Tab1_Print_Click()
'排車一覽表 >> 報表列印
If rs_Tab1_RouteOrders Is Nothing Then Exit Sub
If rs_Tab1_RouteOrders.RecordCount = 0 Then Exit Sub

On Error GoTo err_handle

'1. 資料寫出 Access 資料庫 >> 車輛裝載匯總表
Dim iLoop As Double
Call AccessDB_Connect
Tran_Level = 0
Tran_Level = cnAccess.BeginTrans
str_SQL = "Delete From 揀貨單"
cnAccess.Execute str_SQL, RowsAffect, adExecuteNoRecords
Call ReDim_Recordset(rs_Access)
rs_Access.Open "揀貨單", cnAccess, adOpenStatic, adLockOptimistic
rs_Tab1_RouteOrders.MoveFirst
Do While Not rs_Tab1_RouteOrders.EOF
   rs_Access.AddNew
   For iLoop = 0 To rs_Tab1_RouteOrders.Fields.Count - 1
       rs_Access.Fields(iLoop).Value = rs_Tab1_RouteOrders.Fields(iLoop).Value
   Next iLoop
   rs_Access.Update
   rs_Tab1_RouteOrders.MoveNext
Loop
rs_Tab1_RouteOrders.MoveFirst
cnAccess.CommitTrans
Tran_Level = 0
Call DB_Disconnect(cnAccess)

'2. call Access 列印報表
strAccessDBFileName_FullPath = GetAccessDBFileName
Set MSAccessAP = New access.Application
MSAccessAP.Visible = False
MSAccessAP.OpenCurrentDatabase (strAccessDBFileName_FullPath)

'[報表列印] 命令鈕 -- 利用 Access 報表
If chk_Tab1_PreView.Value = vbChecked Then
   '預覽列印
   MSAccessAP.Visible = True
   MSAccessAP.DoCmd.OpenReport "揀貨單", acViewPreview
Else
   '直接列印至印表機
   MSAccessAP.Visible = False
   MSAccessAP.DoCmd.OpenReport "揀貨單", acViewNormal
   MSAccessAP.CloseCurrentDatabase
   MSAccessAP.Quit
   Set MSAccessAP = Nothing
End If
Exit Sub

err_handle:
   If Tran_Level <> 0 Then cnAccess.RollbackTrans
   Tran_Level = 0
   If Not (MSAccessAP Is Nothing) Then
      If Len(MSAccessAP.CurrentObjectName) <> 0 Then
         MSAccessAP.CloseCurrentDatabase
      End If
      MSAccessAP.Quit
      Set MSAccessAP = Nothing
   End If
   Call Unload_RunLogForm
   Dim tmpString As String
   msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
   tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
   CreateErrorLog Me.Name & "-揀貨單-列印", Me.Caption, "cmd_Tab1_PrintReport_Click", tmpString
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Sub

Private Sub cmd_Tab1_Query_Click()
    '訂單查詢
    Dim ExternOrderKey As String
    Dim strSubwhere As String
    Dim str_Where As String
    On Error GoTo err_handle
    
    str_SQL = "select rtrim(convert(char(8),t2.receipt_date,112)) as 訂單日期,t3.ROUTE_NO as 路線編號, " & _
            "(select top 1 t1m.Short_Name from trp02t m2t inner join TRP01M t1m on t1m.ConsigneeKey = m2t.ConsigneeKey  where m2t.route_no=t3.ROUTE_NO) as  客戶簡稱, " & _
            "rtrim(t3.PRODUCT_NO) as 貨號,rtrim(s.DESCR) as 中文品名,sum(t3.ORDER_QTY) as 數量,rtrim(od.LotTable02) as 指定日期 " & _
            "from trp03t t3 " & _
            "inner join trp02t t2 on t3.route_no=t2.route_no  and t3.RECEIPT_NO=t2.RECEIPT_NO " & _
            "inner join sku s on s.sku=t3.PRODUCT_NO and s.storerkey = t3.storerkey " & _
            "inner join orderdetail od on od.sku=t3.PRODUCT_NO and od.externorderkey=t3.EXTERN and od.orderlinenumber=t3.SEQ_NO "

    '一次排車路線編號
    txt_Tab1_RouteNo_Start.Text = Trim(txt_Tab1_RouteNo_Start.Text)
    txt_Tab1_RouteNo_End.Text = Trim(txt_Tab1_RouteNo_End.Text)
    strSubwhere = ""
    If Len(txt_Tab1_RouteNo_Start.Text) > 0 And Len(txt_Tab1_RouteNo_End.Text) > 0 Then
        strSubwhere = " t3.Route_No Between '" & txt_Tab1_RouteNo_Start.Text & "' and '" & txt_Tab1_RouteNo_End.Text & "' "
    ElseIf Len(txt_Tab1_RouteNo_Start.Text) > 0 And Len(txt_Tab1_RouteNo_End.Text) = 0 Then
        strSubwhere = " t3.Route_No = '" & txt_Tab1_RouteNo_Start.Text & "' "
    ElseIf Len(txt_Tab1_RouteNo_Start.Text) = 0 And Len(txt_Tab1_RouteNo_End.Text) > 0 Then
        strSubwhere = " t3.Route_No = '" & txt_Tab1_RouteNo_End.Text & "' "
    End If
    
    If Len(strSubwhere) > 0 Then
        str_Where = str_Where & " where " & strSubwhere
    End If
    
    If Len(str_Where) = 0 Then
        msg_text = "請輸入查詢條件"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    Else
        str_SQL = str_SQL & str_Where & " group by t2.RECEIPT_DATE,t3.ROUTE_NO,t3.PRODUCT_NO,s.DESCR,od.Lottable02 "
    End If
    Call DB_CheckConnectStatus
    Call ReDim_Recordset(tmp_Rs)
    tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
    If tmp_Rs.EOF Then
        tmp_Rs.Close
        msg_text = "查詢結果：不存在於排車系統"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    
    Call Replication_Recordset(tmp_Rs, rs_Tab1_RouteOrders)
    tmp_Rs.Close
    
    rs_Tab1_RouteOrders.MoveFirst
    Set dg_Tab1_RouteOrders.DataSource = rs_Tab1_RouteOrders
    With dg_Tab1_RouteOrders
        .ColumnHeaders = True         '標題行顯示
        .RowHeight = 250
        .Columns(0).Width = 500       '序號
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Width = 1000      '訂單日期
        .Columns(1).Alignment = dbgLeft
        .Columns(2).Width = 1000      '路線編號
        .Columns(2).Alignment = dbgLeft
        .Columns(3).Width = 1500       '客戶名稱
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Width = 1000       '貨號
        .Columns(4).Alignment = dbgLeft
        .Columns(5).Width = 2000      '中文品名
        .Columns(5).Alignment = dbgLeft
        .Columns(6).Width = 500       '數量
        .Columns(6).Alignment = dbgCenter
        .Columns(7).Width = 1000      '指定日期
        .Columns(7).Alignment = dbgLeft
    End With
    
    DoEvents: DoEvents
    Screen.MousePointer = vbDefault
    Exit Sub
    
err_handle:
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & err.Number & vbCrLf & "Error Descr:" & err.Description
    tmpString = "Error Code:" & err.Number & vbTab & "Error Descr:" & err.Description
    CreateErrorLog Me.Name & "揀貨數量確認-查詢", Me.Caption, "cmd_Tab0_Query_Click", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub

Private Sub dg_Tab0_Ship_Click()
    If dg_Tab0_Ship.Col = 8 Then
        NextPositionCost dg_Tab0_Ship.Row, dg_Tab0_Ship.Col
    End If
End Sub

Public Sub NextPositionCost(ByVal r As Integer, ByVal C As Integer)     '移動文字方塊
    On Error GoTo NextError
    Text1.Width = dg_Tab0_Ship.CellWidth                     '寬度
    Text1.Height = dg_Tab0_Ship.CellHeight                   '高度
    Text1.Left = dg_Tab0_Ship.Left + dg_Tab0_Ship.ColPos(C) + 30 '左側
    Text1.Top = dg_Tab0_Ship.Top + dg_Tab0_Ship.RowPos(r)     '上方
    Text1.Text = dg_Tab0_Ship.Text       '將MSFlexGrid目前作用儲存格內容放置於文字方塊
    Text1.Visible = True                '將文字方塊顯示於畫面上
    Text1.SetFocus                      '將游標移至文字方塊
    Exit Sub
NextError:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    '設定 Form 大小、位置
    dbsrcFormHeight = 7140
    dbsrcFormWidth = 11475
    Me.Height = 7650: Me.Width = 11600
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Left = 200
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300
    SSTab1.Tab = 0
    ck_Tab0_ship.Value = vbChecked
    Call SetGridFormat_Tab0_Ship
End Sub

Private Sub Form_Resize()
    If Me.ScaleHeight = 0 Or Me.ScaleWidth = 0 Then Exit Sub
    If Me.ScaleHeight < dbsrcFormHeight Then
        '變小
        SSTab1.Top = (SSTab1.Top - ((dbsrcFormHeight - Me.ScaleHeight) / 2))
        SSTab1.Left = (SSTab1.Left - ((dbsrcFormWidth - Me.ScaleWidth) / 2))
          
        dbsrcFormHeight = Me.ScaleHeight
        dbsrcFormWidth = Me.ScaleWidth
    Else
        SSTab1.Top = (SSTab1.Top + ((Me.ScaleHeight - dbsrcFormHeight) / 2))
        SSTab1.Left = (SSTab1.Left + ((Me.ScaleWidth - dbsrcFormWidth) / 2))
        
        dbsrcFormHeight = Me.ScaleHeight
        dbsrcFormWidth = Me.ScaleWidth
    End If
End Sub


Private Sub SetGridFormat_Tab0_Ship()
'回傳欲設定之路線編號資料
Dim sub_var1 As Integer, sub_var2 As Integer
dg_Tab0_Ship.Visible = False
With dg_Tab0_Ship
     .Rows = 2: .Cols = 9
     .FixedRows = 1
     '設定允許整列選取
     .AllowBigSelection = True
     '設定列表之文字字型
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             '.CellFontName = "新細明體": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '設定列表之欄位寬度
     .ColWidth(0) = 1500
     .ColWidth(1) = 1500
     .ColWidth(2) = 1000
     .ColWidth(3) = 700
     .ColWidth(4) = 700
     .ColWidth(5) = 1000
     .ColWidth(6) = 1200
     .ColWidth(7) = 1000
     .ColWidth(8) = 1000
     
     'ROUTE_NO,EXTERN,PRODUCT_NO,SEQ_NO,isnull(SubSeq_No,'0'),ORDER_QTY,SHIP_TIME,SHIP_QTY
     '設定列表之標題:路線編號,客戶單號,貨號,項次,拆單,排車數,回傳時間,目前數量,確認數量
     .Row = 0
     .Col = 0: .Text = "路線編號"
     .Col = 1: .Text = "客戶單號"
     .Col = 2: .Text = "貨號"
     .Col = 3: .Text = "項次"
     .Col = 4: .Text = "拆單"
     .Col = 5: .Text = "排車數量"
     .Col = 6: .Text = "回傳時間"
     .Col = 7: .Text = "目前數量"
     .Col = 8: .Text = "確認數量"
     '設定列表之文字對齊
     .ColAlignment(0) = flexAlignLeftCenter
     .ColAlignment(1) = flexAlignLeftCenter
     .ColAlignment(2) = flexAlignLeftCenter
     .ColAlignment(3) = flexAlignCenterCenter
     .ColAlignment(4) = flexAlignCenterCenter
     .ColAlignment(5) = flexAlignRightCenter
     .ColAlignment(6) = flexAlignLeftCenter
     .ColAlignment(7) = flexAlignRightCenter
     .ColAlignment(8) = flexAlignRightCenter
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1
         .CellAlignment = flexAlignCenterCenter
     Next sub_var1
     .Rows = 2: .Row = 1
     For sub_var1 = 0 To .Cols - 1
         .Col = sub_var1: .Text = ""
     Next sub_var1
End With
dg_Tab0_Ship.Visible = True
End Sub

Private Sub Text1_LostFocus()
    On Error GoTo TextError
        Text1.Visible = False
    Exit Sub
TextError:
    MsgBox err.Description
End Sub

Private Sub Text1_Change()  '將文字方塊內容寫至對應儲存格
    On Error GoTo TextError
    dg_Tab0_Ship.Text = Text1.Text   '將文字方塊內容寫至對應儲存格
    Exit Sub
TextError:
    MsgBox err.Description
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    On Error GoTo TextError
    If KeyAscii = vbKeyReturn Then  '在按下Enter時，決定下個grid的位置
        dg_Tab0_Ship.Row = dg_Tab0_Ship.Row + 1
        NextPositionCost dg_Tab0_Ship.Row, dg_Tab0_Ship.Col
    End If
    Exit Sub
TextError:
    MsgBox err.Description
End Sub

