VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_Pallet_Import 
   Caption         =   "棧板資料匯入"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   10680
   Begin VB.Frame fraLocalFiles 
      Caption         =   "檔案總管"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10080
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
         Index           =   2
         Left            =   3840
         Style           =   1  '圖片外觀
         TabIndex        =   7
         Top             =   2400
         Width           =   1200
      End
      Begin VB.CommandButton cmd_Tab0_Import 
         BackColor       =   &H00FFC0C0&
         Caption         =   "匯  入"
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
         Left            =   2640
         Style           =   1  '圖片外觀
         TabIndex        =   6
         Top             =   2400
         Width           =   1200
      End
      Begin VB.ComboBox cmb_Tab0_Storer 
         BackColor       =   &H00C0FFC0&
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
         ItemData        =   "frm_Pallet_Import.frx":0000
         Left            =   960
         List            =   "frm_Pallet_Import.frx":000A
         Style           =   2  '單純下拉式
         TabIndex        =   5
         Top             =   2520
         Width           =   1605
      End
      Begin VB.DriveListBox drvLocalDrive 
         Height          =   300
         Left            =   135
         TabIndex        =   3
         ToolTipText     =   "Local Drive List"
         Top             =   270
         Width           =   2040
      End
      Begin VB.DirListBox dirLocalDir 
         Height          =   1560
         Left            =   135
         TabIndex        =   2
         ToolTipText     =   "Local Directory"
         Top             =   720
         Width           =   4335
      End
      Begin VB.FileListBox filLocalFile 
         Height          =   2070
         Left            =   4560
         Pattern         =   "*.xls"
         TabIndex        =   1
         ToolTipText     =   "Local Files"
         Top             =   240
         Width           =   5190
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "類別："
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
         Index           =   11
         Left            =   240
         TabIndex        =   8
         Top             =   2580
         Width           =   630
      End
   End
   Begin MSDataGridLib.DataGrid dg_Tab0_Pallet 
      Height          =   3255
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5741
      _Version        =   393216
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
Attribute VB_Name = "frm_Pallet_Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rs_Excel As ADODB.Recordset         'excel
Private arStorer() As String                '貨主
Private dbsrcFormHeight As Double           'Form 設計時期的高
Private dbsrcFormWidth As Double            'Form 設計時期的寬
Private str_SDN_Date, str_PalletNo, str_CarNo, str_Type, str_AreaStart, str_AreaEnd, str_uom, str_Cost, str_QTy, str_in, str_out, str_customer As String

Private Sub cmd_Exit_Click(Index As Integer)
    Set rs_Excel = Nothing
    '離開
    Unload Me
End Sub

Private Sub cmd_Tab0_Import_Click()
    If Len(Trim(cmb_Tab0_Storer.Text)) > 0 Then
        Select Case Trim(cmb_Tab0_Storer.Text)
            Case "B&Q南區"
                 Call ImportB
            Case "B&Q中區"
                 Call ImportC
            Case Else
                 msg_text = "無此貨主之匯入程式"
                 MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                 Exit Sub
        End Select
    Else
        msg_text = "請先點選貨主再匯入"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
End Sub

Private Sub dirLocalDir_Change()
    filLocalFile.Path = dirLocalDir.Path
End Sub

Private Sub drvLocalDrive_Change()
On Error GoTo DriveError
    dirLocalDir.Path = drvLocalDrive.Drive
    Exit Sub
DriveError:
    MsgBox "Error accessing selected drive.", vbCritical + vbOKOnly, "Error"
    Resume Next
End Sub

Private Sub Form_Activate()
  '更新 MDIForm 之 Menu [視窗]→[已顯示視窗] 是否核選
  Call UpdateMDIForm_Menu_WindowName_Checked(Me.Tag)
  msg_title = "上下傳"
End Sub

Private Sub Form_Load()
    Me.Height = 7600: Me.Width = 11000
    Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
    Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200
    
    'dirLocalDir.Path = "C:"
    RecievingSize = False
End Sub

Private Sub ImportB()

    '開始匯入檔案
    strExcelFileName = filLocalFile.Path & "\" & filLocalFile.FileName
    If Len(Trim(filLocalFile.FileName)) = 0 Then
        Exit Sub
    End If
    
    If strExcelFileName = "" Then
        '無選取檔案
        Exit Sub
    End If
    If FileLen(strExcelFileName) = 0 Then
        msg_text = "檔案大小=0,檔名:" & str_file
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    On Error GoTo err_Handle
    str_file = Trim(filLocalFile.FileName)
    '檢查是否重複轉檔
'    Call Confirm_Recordset_Closed(tmp_rs)
'    str_SQL = "select * from bestroute where import_file='" & str_file & "'"
'    tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If Not tmp_rs.EOF Then
'        msg_text = "這個檔案資料已匯入"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        Exit Sub
'    End If
    
    '建立 Excel 報表資料庫連接
    strExcel = "Provider=MSDASQL.1;Persist Security Info=False;Driver={Microsoft Excel Driver (*.xls)};DBQ= " & strExcelFileName
    Set cnExcel = New ADODB.Connection
    cnExcel.ConnectionString = strExcel
    cnExcel.Open
    Call ReDim_Recordset(rs_Excel)
    
    rs_Excel.CursorLocation = 3
    str_SQL = "select * from [載具管理$]"
    'rs_Excel.Open str_SQL, cnExcel, adOpenForwardOnly, adLockReadOnly      '無法執行 Set dg_Tab0_Import.DataSource = rs_Excel
    rs_Excel.Open str_SQL, cnExcel, adOpenStatic, adLockOptimistic
    
    If rs_Excel.EOF Then
        rs_Excel.Close
        msg_text = "查詢結果：excel無資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
         cnExcel.Close
         Exit Sub
    Else
        rs_Excel.Sort = "單號 desc"
'        Call OffLineRecordset(tmp_rs, rs_Excel)
        Set dg_Tab0_Pallet.DataSource = rs_Excel
        rs_Excel.MoveFirst
        
        Do While Not rs_Excel.EOF
            str_AreaStart = Trim(rs_Excel.Fields(4).Value)
            str_AreaEnd = Trim(rs_Excel.Fields(5).Value)
            If str_AreaStart = str_AreaEnd Then
                msg_text = "錯誤訊息：起點 " & str_AreaStart & "與迄點:" & str_AreaEnd & "相同"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                'rs_Excel.Close
                Exit Sub
            End If
            
            If str_AreaStart <> "南區轉運站" And str_AreaEnd <> "南區轉運站" Then
                msg_text = "錯誤訊息：起點迄點要有南區轉運站"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                'rs_Excel.Close
                Exit Sub
            End If
            
            If Len(Trim(rs_Excel("日期"))) < 8 Or IsNull(rs_Excel("日期")) Then MsgBox "日期欄位有誤!", 64, "棧板資料匯入": Exit Sub
            If Len(Trim(rs_Excel("單號"))) = 0 Or IsNull(rs_Excel("單號")) Then MsgBox "單號不能空白!", 64, "棧板資料匯入": Exit Sub
            If Len(Trim(rs_Excel("類別"))) = 0 Or IsNull(rs_Excel("類別")) Then MsgBox "類別不能空白!", 64, "棧板資料匯入": Exit Sub
            If Len(Trim(rs_Excel("起點"))) = 0 Or IsNull(rs_Excel("起點")) Then MsgBox "起點不能空白!", 64, "棧板資料匯入": Exit Sub
            If Len(Trim(rs_Excel("迄點"))) = 0 Or IsNull(rs_Excel("迄點")) Then MsgBox "迄點不能空白!", 64, "棧板資料匯入": Exit Sub
            If IsNull(rs_Excel("數量")) Then MsgBox "數量不能為零!", 64, "棧板資料匯入": Exit Sub
            If Val(rs_Excel("數量")) = 0 Then MsgBox "數量不能為零!", 64, "棧板資料匯入": Exit Sub
            
            Call ReDim_Recordset(tmp_rs)
            tmp_rs.Open "select * from pallet_cst where rtrim(checkno) = '" & RTrim(rs_Excel("單號")) & "' ", cn
            If Not tmp_rs.EOF Then MsgBox "單號重複，轉入終止!", 64, "匯入": Exit Sub
            
            rs_Excel.MoveNext
        Loop
        
        rs_Excel.MoveFirst
        int_order = 0: intLine = 0
        Tran_Level = 0
        Tran_Level = cn.BeginTrans
        
        Do While Not rs_Excel.EOF
            DoEvents: DoEvents
'            If IsNull(rs_Excel.Fields(0).Value) Then GoTo exitloop
            If str_PalletNo <> Trim(rs_Excel.Fields(1).Value) Then '資料檢驗--判斷訂單編號已訂是否要在 [明細檔] 中新增一筆
                str_PalletNo = Trim(rs_Excel.Fields(1).Value)
            
            '寫入表頭資料
            str_SQL = "insert into pallet_cds(checkno,storer,carno,usertype,adddate,adduser,edituser,keyindate,editdate) " & _
                "values('" & Trim(rs_Excel("單號")) & "','BEST','" & UCase(Trim(rs_Excel("車號"))) & "','','" & Trim(rs_Excel("日期")) & "','南區轉運站','" & User_id & "','" & Trim(rs_Excel("日期")) & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
            intLine = 1
            
            End If
            
            'excel:日期,單號,車號,類別,起點,迄點,單位,單價,數量
            '新增明細檔
            str_SDN_Date = Format(rs_Excel.Fields(0).Value, "YYYYMMDD")
            str_PalletNo = Trim(rs_Excel.Fields(1).Value)
            str_CarNo = Trim(rs_Excel.Fields(2).Value)
            str_Type = Trim(rs_Excel.Fields(3).Value)
            str_AreaStart = Trim(rs_Excel.Fields(4).Value)
            str_AreaEnd = Trim(rs_Excel.Fields(5).Value)
            str_uom = Trim(rs_Excel.Fields(6).Value)
            str_Cost = Trim(rs_Excel.Fields(7).Value)
            str_QTy = Trim(rs_Excel.Fields(8).Value)
            
            If str_AreaStart = "南區轉運站" Then
                str_in = str_QTy
                str_out = 0
                str_customer = str_AreaEnd
            Else
                str_in = 0
                str_out = str_QTy
                str_customer = str_AreaStart
            End If
             
            'checkno,linenumber,storer,carno,usertype,customer,customernoSheetno,chargedate,qtyin,qtyout,sortingqty,notes,adddate,keyinDate,Editdate,checkDate,AddUser,EditUser,CheckUser,KeyID
            str_SQL = "INSERT Pallet_Cst (checkno,linenumber,storer,carno,usertype,customer,chargedate,adddate,qtyin,qtyout,sortingqty,AddUser,keyindate)" & _
                     "VALUES ('" & str_PalletNo & "','" & intLine & "','Best','" & str_CarNo & "','" & str_Type & "', " & _
                     "'" & str_customer & "','" & str_SDN_Date & "','" & str_SDN_Date & "','" & str_in & "','" & str_out & "','0', " & _
                     "'南區轉運站',getdate())"
                      
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_order = int_order + 1
            intLine = intLine + 1
            rs_Excel.MoveNext
        Loop
exitloop:
        cn.CommitTrans
        Tran_Level = 0
        msg_text = "匯入筆數:" & int_order
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    End If
    Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then
        Tran_Level = 0
        cn.RollbackTrans
    End If
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
    tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
    CreateErrorLog Me.Name & "匯入日報表-匯入", Me.Caption, "Import_other", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault

End Sub

Private Sub ImportC()

    '開始匯入檔案
    strExcelFileName = filLocalFile.Path & "\" & filLocalFile.FileName
    If Len(Trim(filLocalFile.FileName)) = 0 Then
        Exit Sub
    End If
    
    If strExcelFileName = "" Then
        '無選取檔案
        Exit Sub
    End If
    If FileLen(strExcelFileName) = 0 Then
        msg_text = "檔案大小=0,檔名:" & str_file
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
        Exit Sub
    End If
    On Error GoTo err_Handle
    str_file = Trim(filLocalFile.FileName)
    '檢查是否重複轉檔
'    Call Confirm_Recordset_Closed(tmp_rs)
'    str_SQL = "select * from bestroute where import_file='" & str_file & "'"
'    tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
'    If Not tmp_rs.EOF Then
'        msg_text = "這個檔案資料已匯入"
'        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
'        Exit Sub
'    End If
    
    '建立 Excel 報表資料庫連接
    strExcel = "Provider=MSDASQL.1;Persist Security Info=False;Driver={Microsoft Excel Driver (*.xls)};DBQ= " & strExcelFileName
    Set cnExcel = New ADODB.Connection
    cnExcel.ConnectionString = strExcel
    cnExcel.Open
    Call ReDim_Recordset(rs_Excel)
    
    rs_Excel.CursorLocation = 3
    str_SQL = "select * from [載具管理$]"
    'rs_Excel.Open str_SQL, cnExcel, adOpenForwardOnly, adLockReadOnly      '無法執行 Set dg_Tab0_Import.DataSource = rs_Excel
    rs_Excel.Open str_SQL, cnExcel, adOpenStatic, adLockOptimistic
    
    If rs_Excel.EOF Then
        rs_Excel.Close
        msg_text = "查詢結果：excel無資料"
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
         cnExcel.Close
         Exit Sub
    Else
        rs_Excel.Sort = "單號 desc"
'        Call OffLineRecordset(tmp_rs, rs_Excel)
        Set dg_Tab0_Pallet.DataSource = rs_Excel
        rs_Excel.MoveFirst
        
        Do While Not rs_Excel.EOF
            str_AreaStart = Trim(rs_Excel.Fields(4).Value)
            str_AreaEnd = Trim(rs_Excel.Fields(5).Value)
            If str_AreaStart = str_AreaEnd Then
                msg_text = "錯誤訊息：起點 " & str_AreaStart & "與迄點:" & str_AreaEnd & "相同"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                'rs_Excel.Close
                Exit Sub
            End If
            
            If str_AreaStart <> "中區轉運站" And str_AreaEnd <> "中區轉運站" Then
                msg_text = "錯誤訊息：起點迄點要有中區轉運站"
                MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                'rs_Excel.Close
                Exit Sub
            End If
            
            If Len(Trim(rs_Excel("日期"))) < 8 Or IsNull(rs_Excel("日期")) Then MsgBox "日期欄位有誤!", 64, "棧板資料匯入": Exit Sub
            If Len(Trim(rs_Excel("單號"))) = 0 Or IsNull(rs_Excel("單號")) Then MsgBox "單號不能空白!", 64, "棧板資料匯入": Exit Sub
            If Len(Trim(rs_Excel("類別"))) = 0 Or IsNull(rs_Excel("類別")) Then MsgBox "類別不能空白!", 64, "棧板資料匯入": Exit Sub
            If Len(Trim(rs_Excel("起點"))) = 0 Or IsNull(rs_Excel("起點")) Then MsgBox "起點不能空白!", 64, "棧板資料匯入": Exit Sub
            If Len(Trim(rs_Excel("迄點"))) = 0 Or IsNull(rs_Excel("迄點")) Then MsgBox "迄點不能空白!", 64, "棧板資料匯入": Exit Sub
            If IsNull(rs_Excel("數量")) Then MsgBox "數量不能為零!", 64, "棧板資料匯入": Exit Sub
            If Val(rs_Excel("數量")) = 0 Then MsgBox "數量不能為零!", 64, "棧板資料匯入": Exit Sub
            
            Call ReDim_Recordset(tmp_rs)
            tmp_rs.Open "select * from pallet_cst where rtrim(checkno) = '" & RTrim(rs_Excel("單號")) & "' ", cn
            If Not tmp_rs.EOF Then MsgBox "單號重複，轉入終止!", 64, "匯入": Exit Sub
            
            rs_Excel.MoveNext
        Loop
        
        rs_Excel.MoveFirst
        int_order = 0: intLine = 0
        Tran_Level = 0
        Tran_Level = cn.BeginTrans
        
        Do While Not rs_Excel.EOF
            DoEvents: DoEvents
'            If IsNull(rs_Excel.Fields(0).Value) Then GoTo exitloop
            If str_PalletNo <> Trim(rs_Excel.Fields(1).Value) Then '資料檢驗--判斷訂單編號已訂是否要在 [明細檔] 中新增一筆
                str_PalletNo = Trim(rs_Excel.Fields(1).Value)
            
            '寫入表頭資料
            str_SQL = "insert into pallet_cds(checkno,storer,carno,usertype,adddate,adduser,edituser,keyindate,editdate) " & _
                "values('" & Trim(rs_Excel("單號")) & "','BEST','" & UCase(Trim(rs_Excel("車號"))) & "','','" & Trim(rs_Excel("日期")) & "','中區轉運站','" & User_id & "','" & Trim(rs_Excel("日期")) & "','" & Format(Now, "yyyy-MM-dd hh:mm:ss") & "') "
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
                
            intLine = 1
            
            End If
            
            'excel:日期,單號,車號,類別,起點,迄點,單位,單價,數量
            '新增明細檔
            str_SDN_Date = Format(rs_Excel.Fields(0).Value, "YYYYMMDD")
            str_PalletNo = Trim(rs_Excel.Fields(1).Value)
            str_CarNo = Trim(rs_Excel.Fields(2).Value)
            str_Type = Trim(rs_Excel.Fields(3).Value)
            str_AreaStart = Trim(rs_Excel.Fields(4).Value)
            str_AreaEnd = Trim(rs_Excel.Fields(5).Value)
            str_uom = Trim(rs_Excel.Fields(6).Value)
            str_Cost = Trim(rs_Excel.Fields(7).Value)
            str_QTy = Trim(rs_Excel.Fields(8).Value)
            
            If str_AreaStart = "中區轉運站" Then
                str_in = str_QTy
                str_out = 0
                str_customer = str_AreaEnd
            Else
                str_in = 0
                str_out = str_QTy
                str_customer = str_AreaStart
            End If
             
            'checkno,linenumber,storer,carno,usertype,customer,customernoSheetno,chargedate,qtyin,qtyout,sortingqty,notes,adddate,keyinDate,Editdate,checkDate,AddUser,EditUser,CheckUser,KeyID
            str_SQL = "INSERT Pallet_Cst (checkno,linenumber,storer,carno,usertype,customer,chargedate,adddate,qtyin,qtyout,sortingqty,AddUser,keyindate)" & _
                     "VALUES ('" & str_PalletNo & "','" & intLine & "','Best','" & str_CarNo & "','" & str_Type & "', " & _
                     "'" & str_customer & "','" & str_SDN_Date & "','" & str_SDN_Date & "','" & str_in & "','" & str_out & "','0', " & _
                     "'中區轉運站',getdate())"
                      
            cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
            int_order = int_order + 1
            intLine = intLine + 1
            rs_Excel.MoveNext
        Loop
exitloop:
        cn.CommitTrans
        Tran_Level = 0
        msg_text = "匯入筆數:" & int_order
        MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    End If
    Exit Sub
    
err_Handle:
    If Tran_Level <> 0 Then
        Tran_Level = 0
        cn.RollbackTrans
    End If
    Dim tmpString As String
    msg_text = "錯誤訊息：" & vbCrLf & "Error Code:" & Err.Number & vbCrLf & "Error Descr:" & Err.Description
    tmpString = "Error Code:" & Err.Number & vbTab & "Error Descr:" & Err.Description
    CreateErrorLog Me.Name & "匯入日報表-匯入", Me.Caption, "Import_other", tmpString
    MsgBox msg_text, vbOKOnly + vbInformation, msg_title
    Screen.MousePointer = vbDefault
End Sub



