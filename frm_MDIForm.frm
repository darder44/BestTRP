VERSION 5.00
Begin VB.MDIForm frm_MDIForm 
   BackColor       =   &H8000000C&
   Caption         =   "車輛派遣系統"
   ClientHeight    =   5925
   ClientLeft      =   915
   ClientTop       =   2010
   ClientWidth     =   11370
   Icon            =   "frm_MDIForm.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frm_MDIForm.frx":0442
   WindowState     =   2  '最大化
   Begin VB.Menu Menu_Orders 
      Caption         =   "訂單處理作業"
      Begin VB.Menu Menu_Upload_FTP 
         Caption         =   "訂單接收"
      End
      Begin VB.Menu Menu_TRPPlan_ManualOrders 
         Caption         =   "訂單維護"
      End
      Begin VB.Menu Menu_TRPPlan_OrderImport 
         Caption         =   "訂單轉入及客戶異動維護"
      End
      Begin VB.Menu Menu_TRPPlan_Query 
         Caption         =   "訂單查詢作業"
      End
   End
   Begin VB.Menu Menu_TRPPlan 
      Caption         =   "一般排車作業"
      Begin VB.Menu Menu_TRPPlan_CutOrders 
         Caption         =   "ㄧ單多車訂單切割"
      End
      Begin VB.Menu Menu_TRPPlan_TRPPlan 
         Caption         =   "一般排車作業"
      End
      Begin VB.Menu Menu_TRPPlan_DCRouteMerge 
         Caption         =   "二次排車作業"
      End
      Begin VB.Menu Menu_TRPPlan_BacktoEXE 
         Caption         =   "排車資料回傳設定"
      End
      Begin VB.Menu Menu_TRPPlan_Route 
         Caption         =   "路線編號維護作業"
      End
      Begin VB.Menu Menu_TRPPlan_ReDelivery 
         Caption         =   "未收訂單再配送作業"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_TRPPlan_Report 
         Caption         =   "排車作業報表"
      End
      Begin VB.Menu Menu_TRPPlan_RouteConfirm 
         Caption         =   "出車確認"
      End
      Begin VB.Menu Menu_TRPPlan_SDNAbnormal 
         Caption         =   "配送異常維護"
      End
      Begin VB.Menu Menu_TRPPlan_SDNConfirm 
         Caption         =   "簽單確認"
      End
      Begin VB.Menu Menu_TRPPlan_ShipQty 
         Caption         =   "揀貨數量確認"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_TRPPlan_Cost 
         Caption         =   "運費分析"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Menu_Other 
      Caption         =   "其它排車作業"
      Begin VB.Menu Menu_OP_Other_OrderImport 
         Caption         =   "訂單轉入及客戶異動維護"
      End
      Begin VB.Menu Menu_Other_ORTPlan 
         Caption         =   "排車作業"
      End
      Begin VB.Menu Menu_Other_Report 
         Caption         =   "排車作業報表"
      End
      Begin VB.Menu Menu_OP_RSDNConfirm 
         Caption         =   "退貨簽單維護"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Menu_Pallet 
      Caption         =   "其它管理作業"
      Begin VB.Menu Menu_BQControlSheet 
         Caption         =   "BQ管制表"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_PalletxSorting 
         Caption         =   "棧板管理"
      End
      Begin VB.Menu Menu_LoadSorting 
         Caption         =   "翻板理貨管理"
      End
      Begin VB.Menu Menu_Pallet_UTLCst 
         Caption         =   "經銷商棧板管理"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_Pallet_Match 
         Caption         =   "棧板資料確認"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_Pallet_CSVehicle_id_no 
         Caption         =   "中南區車號匯入"
      End
      Begin VB.Menu Menu_OP_CaseConfirm 
         Caption         =   "出貨件數確認"
      End
      Begin VB.Menu Menu_Line3x 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Query_Pallet 
         Caption         =   "棧板對帳表"
      End
      Begin VB.Menu Menu_Query_PalletCST 
         Caption         =   "棧板統計結餘"
      End
      Begin VB.Menu Menu_Query_PalletDetail 
         Caption         =   "棧板明細查詢"
      End
      Begin VB.Menu Menu_Query_loadsortingDetail 
         Caption         =   "翻板理貨簽收單"
      End
      Begin VB.Menu Menu_Query_PalletRent 
         Caption         =   "租金計算"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_OP_PalletExport 
         Caption         =   "交易資料匯出"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_OP_PalletImport 
         Caption         =   "棧板資料匯入"
      End
   End
   Begin VB.Menu Menu_Report 
      Caption         =   "報表"
      Begin VB.Menu Menu_Report_DivideSku 
         Caption         =   "花王分貨表"
      End
      Begin VB.Menu Menu_Report_MboReport 
         Caption         =   "毛寶需求報表"
         Begin VB.Menu Menu_Report_MBO_Cod 
            Caption         =   "毛寶下貨收現查詢"
         End
         Begin VB.Menu Menu_Report_MboReport_PodRetrun 
            Caption         =   "POD回傳"
         End
         Begin VB.Menu Menu_Report_MboReport_SDNReturnList 
            Caption         =   "回單檢核表"
         End
      End
      Begin VB.Menu Menu_TKReport 
         Caption         =   "TK需求報表"
         Begin VB.Menu Menu_Report_Ship2TKK 
            Caption         =   "出貨資料回傳"
            Visible         =   0   'False
         End
         Begin VB.Menu Menu_Report_DelOrder 
            Caption         =   "訂單刪除明細"
         End
         Begin VB.Menu Menu_Report_DeliveryTrack 
            Caption         =   "客戶到貨追蹤表"
         End
         Begin VB.Menu Menu_Report_TKExpect 
            Caption         =   "訂單退回明細"
         End
         Begin VB.Menu Menu_Report_TKExpect1 
            Caption         =   "訂單配送異常表"
         End
         Begin VB.Menu Menu_Report_TKCustomerCodeDate 
            Caption         =   "客戶進貨有效期限明細表"
         End
         Begin VB.Menu Menu_Report_TKKSDNReturnList 
            Caption         =   "送貨回單檢核表"
         End
         Begin VB.Menu Menu_Report_TKKRSDNReturnList 
            Caption         =   "退貨回單檢核表"
         End
         Begin VB.Menu Menu_Report_TKKPI 
            Caption         =   "單量明細"
         End
         Begin VB.Menu Menu_Report_TKARList 
            Caption         =   "應收帳款明細表"
         End
      End
      Begin VB.Menu Menu_VTLReport 
         Caption         =   "VTL需求報表"
      End
      Begin VB.Menu Menu_THLReport 
         Caption         =   "THL需求報表"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_NSLReport 
         Caption         =   "NSL需求報表"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_ABTReport 
         Caption         =   "ABT需求報表"
      End
      Begin VB.Menu Menu_Report_SDNReturnList 
         Caption         =   "回單檢核表"
      End
      Begin VB.Menu Menu_Report_TRPTrack 
         Caption         =   "到貨追蹤表"
      End
      Begin VB.Menu Menu_Report_TMSAbnormal 
         Caption         =   "配送異常表"
      End
      Begin VB.Menu Menu_Report_APPSdnDetail 
         Caption         =   "簽單明細表"
      End
   End
   Begin VB.Menu Menu_Query 
      Caption         =   "查詢"
      Begin VB.Menu Menu_Query_InterfaceLog 
         Caption         =   "InterFaceLog"
      End
      Begin VB.Menu Menu_Query_KPI 
         Caption         =   "管理KPI"
         Begin VB.Menu Menu_Query_KPI_KPI 
            Caption         =   "每日KPI"
         End
         Begin VB.Menu Menu_Query_KPI_CarCount 
            Caption         =   "每日區域車型車次KPI"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Menu_Query_Charge 
         Caption         =   "請付款日報表"
      End
      Begin VB.Menu Menu_Query_Account_LoadSorting 
         Caption         =   "會計翻板與理貨資料"
      End
      Begin VB.Menu Menu_BackOrderDetail 
         Caption         =   "退換貨與拒短收明細"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Menu_BaseData 
      Caption         =   "基本資料"
      Begin VB.Menu Menu_BaseData_Car 
         Caption         =   "車輛/貨運公司"
      End
      Begin VB.Menu Menu_BaseData_ConsigCar 
         Caption         =   "客戶/車輛/貨運公司"
      End
      Begin VB.Menu Menu_DY_BaseData_ConsigCar 
         Caption         =   "允收期匯入"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_BaseData_SKU 
         Caption         =   "亞培貨號資料維護"
      End
      Begin VB.Menu Menu_BaseData_OP 
         Caption         =   "作業代碼資料維護"
      End
      Begin VB.Menu Menu_BaseData_OP_1 
         Caption         =   "進階代碼資料維護"
      End
      Begin VB.Menu Menu_BaseData_UserData 
         Caption         =   "使用者資料維護"
      End
      Begin VB.Menu Menu_BaseData_Code 
         Caption         =   "系統代碼維護"
      End
      Begin VB.Menu Menu_BaseData_UserSecutiry 
         Caption         =   "系統權限設定"
      End
   End
   Begin VB.Menu Menu_System 
      Caption         =   "系統設定"
      Begin VB.Menu Menu_SwitchDB 
         Caption         =   "資料庫切換"
      End
      Begin VB.Menu Menu_SystemUpdate 
         Caption         =   "系統更新"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_Line2x 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Options 
         Caption         =   "選項"
      End
   End
   Begin VB.Menu Menu_Windowsx 
      Caption         =   "視窗排列"
      Begin VB.Menu Menu_WindowMinx 
         Caption         =   "最小化"
      End
      Begin VB.Menu mnuWindowCascadex 
         Caption         =   "重疊顯示"
      End
      Begin VB.Menu mnuWindowTileHorizontalx 
         Caption         =   "水平並排"
      End
      Begin VB.Menu mnuWindowTileVerticalx 
         Caption         =   "垂直並排"
      End
      Begin VB.Menu mnuWindowArrangeIconsx 
         Caption         =   "排列圖示"
      End
      Begin VB.Menu Menu_WindowSourceSizex 
         Caption         =   "原始視窗"
      End
      Begin VB.Menu Menu_Line1x 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1未指定"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1未指定"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1未指定"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1未指定"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1未指定"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1未指定"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1未指定"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1未指定"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1未指定"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1未指定"
         Index           =   9
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Menu_Exitx 
      Caption         =   "離開"
   End
End
Attribute VB_Name = "frm_MDIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strScreenRes As String '螢幕解析度

Private Sub MDIForm_Load()

'螢幕解析度
 strScreenRes = Screen.Width \ Screen.TwipsPerPixelX & "x" & Screen.Height \ Screen.TwipsPerPixelY
 '底圖
 If Dir(App.Path & "\" & strScreenRes & ".pic") <> "" Then Me.Picture = LoadPicture(App.Path & "\" & strScreenRes & ".pic")

Me.Caption = Me.Caption & "(" & App.Major & "." & App.Minor & "." & App.Revision & ")"

Load frm_WaitWindows
frm_WaitWindows.Tag = "frm_MDIForm"
frm_WaitWindows.ZOrder
  
Do While TypeName(cn) = "Nothing"
   DoEvents
Loop
Do While cn.State = adStateConnecting
   DoEvents
Loop

'確認codelist是否有此項目
tmp_Rs.Open "select listname from codelist where listname = 'Options'", cn
If tmp_Rs.EOF Then cn.Execute "insert into codelist(listname,description,adddate,addwho,editdate,editwho) values ('Options','選項設定值',getdate(),'dbo',getdate(),'dbo')", RowsAffect, adExecuteNoRecords
tmp_Rs.Close

  '由資料庫參數設定，取得 Security Control 設定值
  blSecurityControl = True
  cn.Execute "select listname from codelkup where listname = 'Options' and code = 'LoginControl' and Description = 0 ", RowsAffect, adExecuteNoRecords
  If RowsAffect <> 0 Then blSecurityControl = False: User_id = strComputerName: blAdmin = True 'RouteModify= 0 時表示資料庫無相符資料
  
  '排車資料修改是否限制原使用者
  cn.Execute "select listname from codelkup where listname = 'Options' and code = 'RouteModify' and Description = 0 ", RowsAffect, adExecuteNoRecords
  If RowsAffect = 0 Then blRouteModifyControl = True '沒找到時要控管
 
'資料維護期限
str_SQL = "select DueDate = convert(char(8),getdate()- cast(isnull(description,0) as int),112) from codelkup where listname = 'Options' and code = 'DueDate' "
tmp_Rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly

If tmp_Rs.EOF Then
    cn.Execute "insert into codelkup(listname,code,description,short,long,notes,adddate,addwho,editdate,editwho) values ('Options','DueDate',60,'','','',getdate(),'dbo',getdate(),'dbo')", RowsAffect, adExecuteNoRecords
    lngDueDate = Format(Now - 60, "yyyymmdd")
Else
    lngDueDate = tmp_Rs("Duedate")
End If

tmp_Rs.Close
  
If blSecurityControl Then
   Call Disable_Menu
   '把所有功能表都設為 Disable，再依使用者權限設定值逐一開啟：enable
   '電腦名稱是BESTPREPARES or gemini則跳開，避開開發時版本不同的問題
   If UCase(strComputerName) <> "BESTPREPARES" And UCase(strComputerName) <> "BEST_ALICENB" And UCase(strComputerName) <> "BEST-TERRY" And UCase(strComputerName) <> "BEST-TEST" And UCase(strComputerName) <> "GEMINI_NB" And UCase(strComputerName) <> "GEMINI_VPC" Then
         '檢查系統版本是否為最新版本，不是則需手動更新，不然無法使用
        If RTrim(App.EXEName) = "BestTRP" Then
          '正常版的TMS
          '檢查版本編號是否為最新版本
          tmp_Rs.Open "select top 1 version from VersionCheck where project = 'BestTms' order by adddate desc", cn, adOpenForwardOnly, adLockReadOnly
          If RTrim(tmp_Rs.Fields("version")) = RTrim(App.Major & "." & App.Minor & "." & App.Revision) Then
              tmp_Rs.Close
          Else
              MsgBox "TMS有新版本發佈:" & RTrim(tmp_Rs.Fields("version")) & "，請關閉您的TMS並更新您的TMS，確保系統的正確性!" & Chr(13) & "否則為了資料的正確性，您將無法繼續使用!", vbOKOnly + vbCritical, "TMS版本檢查"
              tmp_Rs.Close
              Exit Sub
          End If
        Else
          '舊版備用的TMS
          MsgBox "你使用的版本是TMS old版本，請確認捷徑的路徑是否正確!" & Chr(13) & "您仍可繼續使用!但建議您使用最新版本，確保資料正確性!" & Chr(9) & Chr(9) & Chr(9) & "系統將紀錄您的帳號做為日後稽核。", vbOKOnly + vbExclamation, "TMS版本檢查"
          str_SQL = "Insert into gt_Logs(APName,APVer,APCaption,Code,Description,Notes,ComputerName,AddWho) Values ('" & _
                        App.EXEName & "','" & App.Major & "." & App.Minor & "." & App.Revision & "','" & Me.Caption & "','" & "" & "','" & "使用舊版TMS系統" & "','" & "使用舊版TMS系統" & "','" & strComputerName & "','" & User_id & "')"
          cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        End If
    End If
    
     Load frm_UserLogin
     frm_UserLogin.Visible = False: frm_UserLogin.WindowState = vbNormal
     frm_UserLogin.Visible = True
     frm_UserLogin.ZOrder
     frm_UserLogin.Tag = "系統登入"
End If

  Call HideMenu

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Call DB_Disconnect(cn)
  End
End Sub

Private Sub Menu_ABTReport_Click()
    '報表 → ABT需求報表
    If CheckOpenForm("ABT需求報表") = 1 Then Exit Sub
    Dim obj As Object
    Set obj = frm_Report_ABT
    Load obj
    obj.Visible = False
    obj.Visible = True
    obj.ZOrder
    obj.Tag = "ABT需求報表"
    obj.WindowState = 2
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_BackOrderDetail_Click()

    '查詢 → 退換貨與拒短收明細
    If CheckOpenForm("退換貨與拒短收明細") = 1 Then Exit Sub
    Load frm_Query_BackOrderDetail
    frm_Query_BackOrderDetail.Visible = False: frm_Query_BackOrderDetail.WindowState = 2
    frm_Query_BackOrderDetail.Visible = True
    frm_Query_BackOrderDetail.ZOrder
    frm_Query_BackOrderDetail.Tag = "退換貨與拒短收明細"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BaseData_Code_Click()
    'Menu 基本資料 → 子系統代碼資料維護
    If CheckOpenForm("子系統代碼資料維護") = 1 Then Exit Sub
    Load frm_BaseData_Code
    frm_BaseData_Code.Visible = False: frm_BaseData_Code.WindowState = vbNormal
    frm_BaseData_Code.Visible = True
    frm_BaseData_Code.ZOrder
    frm_BaseData_Code.Tag = "子系統代碼資料維護"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BaseData_Car_Click()
    'Menu 基本資料 → 車輛基本資料維護
    If CheckOpenForm("車輛基本資料維護") = 1 Then Exit Sub
    Load frm_BaseData_Car
    frm_BaseData_Car.Visible = False: frm_BaseData_Car.WindowState = 2
    frm_BaseData_Car.Visible = True
    frm_BaseData_Car.ZOrder
    frm_BaseData_Car.Tag = "車輛基本資料維護"
    Call UpdateMDIForm_Menu_WindowName
End Sub
Private Sub Menu_BaseData_ConsigCar_Click()
    'Menu 基本資料 → 客戶/車輛基本資料維護
    If CheckOpenForm("客戶/車輛基本資料維護") = 1 Then Exit Sub
    Load frm_BaseData_ConsigCar
    frm_BaseData_ConsigCar.Visible = False: frm_BaseData_ConsigCar.WindowState = 2
    frm_BaseData_ConsigCar.Visible = True
    frm_BaseData_ConsigCar.ZOrder
    frm_BaseData_ConsigCar.Tag = "客戶/車輛基本資料維護"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_BaseData_OP_1_Click()
    'Menu 基本資料 → 進階代碼資料維護
    If CheckOpenForm("進階代碼資料維護") = 1 Then Exit Sub
    Load frm_BaseData_OPCode_1
    frm_BaseData_OPCode_1.Visible = False: frm_BaseData_OPCode_1.WindowState = 2
    frm_BaseData_OPCode_1.Visible = True
    frm_BaseData_OPCode_1.ZOrder
    frm_BaseData_OPCode_1.Tag = "進階代碼資料維護"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BaseData_OP_Click()
    'Menu 基本資料 → 作業代碼資料維護
    If CheckOpenForm("作業代碼資料維護") = 1 Then Exit Sub
    Load frm_BaseData_OPCode
    frm_BaseData_OPCode.Visible = False: frm_BaseData_OPCode.WindowState = 2
    frm_BaseData_OPCode.Visible = True
    frm_BaseData_OPCode.ZOrder
    frm_BaseData_OPCode.Tag = "作業代碼資料維護"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BaseData_SKU_Click()
    'Menu 基本資料 → 亞培商品資料維護
    If CheckOpenForm("亞培商品資料維護") = 1 Then Exit Sub
    Load frm_BaseData_Sku
    frm_BaseData_Sku.Visible = False: frm_BaseData_Sku.WindowState = vbNormal
    frm_BaseData_Sku.Visible = True
    frm_BaseData_Sku.ZOrder
    frm_BaseData_Sku.Tag = "亞培商品資料維護"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BaseData_UserData_Click()
    'Menu 基本資料 → 使用者資料維護
    If CheckOpenForm("使用者資料維護") = 1 Then Exit Sub
    Load frm_BaseData_UserData
    frm_BaseData_UserData.Visible = False: frm_BaseData_UserData.WindowState = vbNormal
    frm_BaseData_UserData.Visible = True
    frm_BaseData_UserData.ZOrder
    frm_BaseData_UserData.Tag = "使用者資料維護"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BaseData_UserSecutiry_Click()
    'Menu 基本資料 → 系統權限設定
    If CheckOpenForm("系統權限設定") = 1 Then Exit Sub
    Load frm_BaseData_UserSecurity
    frm_BaseData_UserSecurity.Visible = False: frm_BaseData_UserSecurity.WindowState = vbNormal
    frm_BaseData_UserSecurity.Visible = True
    frm_BaseData_UserSecurity.ZOrder
    frm_BaseData_UserSecurity.Tag = "系統權限設定"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_CaeManage_CarControl_Click()
    '排車處理作業 →  車輛進出管制作業
    If CheckOpenForm("車輛進出管制作業") = 1 Then Exit Sub
    Load frm_OP_CarControl
    frm_OP_CarControl.Visible = False: frm_OP_CarControl.WindowState = 2
    frm_OP_CarControl.Visible = True
    frm_OP_CarControl.ZOrder
    frm_OP_CarControl.Tag = "車輛進出管制作業"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BQControlSheet_Click()
    '其他管理作業→BQ管制表
    If CheckOpenForm("BQ管制表") = 1 Then Exit Sub
    
    Load frm_OP_BQControlSheet
    frm_OP_BQControlSheet.Visible = False: frm_OP_BQControlSheet.WindowState = 2
    frm_OP_BQControlSheet.Visible = True
    frm_OP_BQControlSheet.ZOrder
    frm_OP_BQControlSheet.Tag = "BQ管制表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Exitx_Click()
  Call DB_Disconnect(cn)
  End
End Sub

Private Sub Menu_FormNamex_Click(Index As Integer)
'Menu 之 [視窗]→[已顯示視窗]
'將被選取之表單調到最前端
Dim i As Integer, SelectedForm As Integer
For i = 0 To Forms.Count - 1
    If Not (Forms(i) Is frm_MDIForm) Then
       If Forms(i).Tag = frm_MDIForm.Menu_FormNamex(Index).Caption Then
          SelectedForm = i
       Else
          Forms(i).WindowState = vbMinimized
       End If
    End If
Next i
Forms(SelectedForm).WindowState = 2
Forms(SelectedForm).ZOrder
End Sub

Private Sub Menu_LoadSorting_Click()
    '其他管理作業→翻板理貨管理
    If CheckOpenForm("翻板理貨管理") = 1 Then Exit Sub
    
    Load frm_OP_LoadSorting
    frm_OP_LoadSorting.Visible = False: frm_OP_LoadSorting.WindowState = 2
    frm_OP_LoadSorting.Visible = True
    frm_OP_LoadSorting.ZOrder
    frm_OP_LoadSorting.Tag = "翻板理貨管理"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_NSLReport_Click()
    '報表 → NSL需求報表
    If CheckOpenForm("NSL需求報表") = 1 Then Exit Sub
    Dim obj As Object
    Set obj = frm_Report_NSL
    Load obj
    obj.Visible = False
    obj.Visible = True
    obj.ZOrder
    obj.Tag = "NSL需求報表"
    obj.WindowState = 2
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_OP_CaseConfirm_Click()
    '其他管理作業
    If CheckOpenForm("出貨件數確認") = 1 Then Exit Sub
    
    Load frm_OP_CaseConfirm
'    frm_OP_Other_OrderImport.Visible = False: frm_OP_Other_OrderImport.WindowState = 2
'    frm_OP_Other_OrderImport.Visible = True
'    frm_OP_Other_OrderImport.ZOrder
    frm_OP_CaseConfirm.Tag = frm_OP_CaseConfirm.Caption
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_OP_Other_OrderImport_Click()
    '退貨訂單轉入及客戶異動維護
    If CheckOpenForm("退貨訂單轉入及客戶異動維護") = 1 Then Exit Sub
    
    Load frm_OP_Other_OrderImport
    frm_OP_Other_OrderImport.Visible = False: frm_OP_Other_OrderImport.WindowState = 2
    frm_OP_Other_OrderImport.Visible = True
    frm_OP_Other_OrderImport.ZOrder
    frm_OP_Other_OrderImport.Tag = "退貨訂單轉入及客戶異動維護"
    Call UpdateMDIForm_Menu_WindowName
End Sub



Private Sub Menu_OP_RSDNConfirm_Click()
    '退貨排車
    If CheckOpenForm("退貨簽單維護") = 1 Then Exit Sub
    
    Load frm_OP_RSDNConfirm
    frm_OP_RSDNConfirm.Visible = False: frm_OP_RSDNConfirm.WindowState = 2
    frm_OP_RSDNConfirm.Visible = True
    frm_OP_RSDNConfirm.ZOrder
    frm_OP_RSDNConfirm.Tag = "退貨簽單維護"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Options_Click()
    '選項
    If CheckOpenForm("選項") = 1 Then Exit Sub
    
    Load frm_Options
    frm_Options.Visible = False ': frm_Options.WindowState = 2
    frm_Options.Visible = True
    frm_Options.ZOrder
    frm_Options.Tag = "選項"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Other_ORTPlan_Click()
    '退貨排車
    If CheckOpenForm("退貨排車") = 1 Then Exit Sub
    
    Load frm_Other_OPTPlan
    frm_Other_OPTPlan.Visible = False: frm_Other_OPTPlan.WindowState = 2
    frm_Other_OPTPlan.Visible = True
    frm_Other_OPTPlan.ZOrder
    frm_Other_OPTPlan.Tag = "退貨排車"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Other_Report_Click()
    '退貨報表
    If CheckOpenForm("退貨報表") = 1 Then Exit Sub
    Load frm_Report_Other
    frm_Report_Other.Visible = False: frm_Report_Other.WindowState = 2
    frm_Report_Other.Visible = True
    frm_Report_Other.ZOrder
    frm_Report_Other.Tag = "退貨報表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Pallet_CSVehicle_id_no_Click()
    '其他管理作業
    If CheckOpenForm("中南區車號匯入") = 1 Then Exit Sub
    
    Load frm_Pallet_CSVehicle_id_no
    frm_Pallet_CSVehicle_id_no.Tag = frm_Pallet_CSVehicle_id_no.Caption
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Pallet_Match_Click()
    '棧板資料確認
    If CheckOpenForm("棧板資料確認") = 1 Then Exit Sub
    Load frm_Pallet_Match
    frm_Pallet_Match.Visible = False: frm_Pallet_Match.WindowState = 2
    frm_Pallet_Match.Visible = True
    frm_Pallet_Match.ZOrder
    frm_Pallet_Match.Tag = "棧板資料確認"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Pallet_UTL_Click()
    '棧板管理
    If CheckOpenForm("棧板管理") = 1 Then Exit Sub
    Load frm_Pallet_UTL
    frm_Pallet_UTL.Visible = False: frm_Pallet_UTL.WindowState = 2
    frm_Pallet_UTL.Visible = True
    frm_Pallet_UTL.ZOrder
    frm_Pallet_UTL.Tag = "棧板管理"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Pallet_UTLCst_Click()
    '經銷商棧板管理
    If CheckOpenForm("經銷商棧板管理") = 1 Then Exit Sub
    Load frm_Pallet_UTLCst
    frm_Pallet_UTLCst.Visible = False: frm_Pallet_UTLCst.WindowState = 2
    frm_Pallet_UTLCst.Visible = True
    frm_Pallet_UTLCst.ZOrder
    frm_Pallet_UTLCst.Tag = "經銷商棧板管理"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_PalletxSorting_Click()
    '其他管理作業->棧板與理貨管理
    If CheckOpenForm("棧板與理貨管理") = 1 Then Exit Sub
    Load frm_OP_PalletxSorting
    frm_OP_PalletxSorting.Visible = False: frm_OP_PalletxSorting.WindowState = 2
    frm_OP_PalletxSorting.Visible = True
    frm_OP_PalletxSorting.ZOrder
    frm_OP_PalletxSorting.Tag = "棧板與理貨管理"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_Account_LoadSorting_Click()
    '查詢->翻板與理貨資料
    If CheckOpenForm("翻板與理貨資料") = 1 Then Exit Sub
    Load frm_Query_Account_LoadSorting
    frm_Query_Account_LoadSorting.Visible = False: frm_Query_Account_LoadSorting.WindowState = 2
    frm_Query_Account_LoadSorting.Visible = True
    frm_Query_Account_LoadSorting.ZOrder
    frm_Query_Account_LoadSorting.Tag = "翻板與理貨資料"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_Charge_Click()

    '查詢->客戶請款資料
    If CheckOpenForm("客戶請款資料") = 1 Then Exit Sub
    Load frm_Query_Charge
    frm_Query_Charge.Visible = False: frm_Query_Charge.WindowState = 2
    frm_Query_Charge.Visible = True
    frm_Query_Charge.ZOrder
    frm_Query_Charge.Tag = "客戶請款資料"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_Query_InterfaceLog_Click()
Dim obj
Set obj = frm_Query_InterfaceLog
If CheckOpenForm(obj.Caption) = 1 Then Exit Sub

Load obj
obj.Visible = False: obj.WindowState = 2
obj.Visible = True
obj.ZOrder
obj.Tag = obj.Caption
Call UpdateMDIForm_Menu_WindowName
End Sub

'Private Sub Menu_Query_Inventory_Click()
'    '查詢 → 即時庫存查詢
'    If CheckOpenForm("即時庫存查詢") = 1 Then Exit Sub
'    Load frm_Query_Inventory
'    frm_Query_Inventory.Visible = False
'    frm_Query_Inventory.WindowState = 2 '  = vbNormal
'    frm_Query_Inventory.Visible = True
'    frm_Query_Inventory.ZOrder
'    frm_Query_Inventory.Tag = "即時庫存查詢"
'    Call UpdateMDIForm_Menu_WindowName
'End Sub

Private Sub Menu_Query_KPI_CarCount_Click()
    '管理KPI → 每日區域車型車次KPI
    If CheckOpenForm("每日區域車型車次KPI") = 1 Then Exit Sub
    Load frm_Query_KPI_CarCount
    frm_Query_KPI_CarCount.Visible = False
    frm_Query_KPI_CarCount.WindowState = 2 '  = vbNormal
    frm_Query_KPI_CarCount.Visible = True
    frm_Query_KPI_CarCount.ZOrder
    frm_Query_KPI_CarCount.Tag = "每日區域車型車次KPI"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_loadsortingDetail_Click()
    '翻板理貨明細查詢
    If CheckOpenForm("翻板理貨明細查詢") = 1 Then Exit Sub
    Load frm_Query_LoadSortingDetail
    frm_Query_LoadSortingDetail.Visible = False: frm_Query_LoadSortingDetail.WindowState = 2
    frm_Query_LoadSortingDetail.Visible = True
    frm_Query_LoadSortingDetail.ZOrder
    frm_Query_LoadSortingDetail.Tag = "翻板理貨明細查詢"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_Pallet_Click()
    '對帳表
    If CheckOpenForm("對帳表") = 1 Then Exit Sub
    Load frm_Query_Pallet
    frm_Query_Pallet.Visible = False: frm_Query_Pallet.WindowState = 2
    frm_Query_Pallet.Visible = True
    frm_Query_Pallet.ZOrder
    frm_Query_Pallet.Tag = "對帳表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_PalletCST_Click()
    '統計結餘
    If CheckOpenForm("統計結餘") = 1 Then Exit Sub
    Load frm_Query_PalletCst
    frm_Query_PalletCst.Visible = False: frm_Query_PalletCst.WindowState = 2
    frm_Query_PalletCst.Visible = True
    frm_Query_PalletCst.ZOrder
    frm_Query_PalletCst.Tag = "統計結餘"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_PalletDetail_Click()
    '棧板明細查詢
    If CheckOpenForm("棧板明細查詢") = 1 Then Exit Sub
    Load frm_Query_PalletDetail
    frm_Query_PalletDetail.Visible = False: frm_Query_PalletDetail.WindowState = 2
    frm_Query_PalletDetail.Visible = True
    frm_Query_PalletDetail.ZOrder
    frm_Query_PalletDetail.Tag = "棧板明細查詢"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_PalletRent_Click()
    '租金計算
    If CheckOpenForm("租金計算") = 1 Then Exit Sub
    Load frm_Query_PalletRent
    frm_Query_PalletRent.Visible = False: frm_Query_PalletRent.WindowState = 2
    frm_Query_PalletRent.Visible = True
    frm_Query_PalletRent.ZOrder
    frm_Query_PalletRent.Tag = "租金計算"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_OP_PalletExport_Click()
    '棧板資料匯出
    If CheckOpenForm("棧板資料匯出") = 1 Then Exit Sub
    Load frm_OP_PalletExport
    frm_OP_PalletExport.Visible = False: frm_OP_PalletExport.WindowState = 2
    frm_OP_PalletExport.Visible = True
    frm_OP_PalletExport.ZOrder
    frm_OP_PalletExport.Tag = "棧板資料匯出"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_OP_PalletImport_Click()
'    '棧板資料匯入
'    If CheckOpenForm("棧板資料匯入") = 1 Then Exit Sub
'    Load frm_OP_PalletImport
'    frm_OP_PalletImport.Visible = False: frm_OP_PalletImport.WindowState = 2
'    frm_OP_PalletImport.Visible = True
'    frm_OP_PalletImport.ZOrder
'    frm_OP_PalletImport.Tag = "棧板資料匯入"
'    Call UpdateMDIForm_Menu_WindowName
    
    '棧板資料匯入
    If CheckOpenForm("棧板資料匯入") = 1 Then Exit Sub
    Load frm_Pallet_Import
    frm_Pallet_Import.Visible = False: frm_Pallet_Import.WindowState = 2
    frm_Pallet_Import.Visible = True
    frm_Pallet_Import.ZOrder
    frm_Pallet_Import.Tag = "棧板資料匯入"
    Call UpdateMDIForm_Menu_WindowName
    
End Sub

'Private Sub Menu_Query_ReceiptDetail_Click()
'    '查詢 → 入庫明細資料查詢
'    If CheckOpenForm("入庫明細資料查詢") = 1 Then Exit Sub
'    Load frm_Query_ReceiptDetail
'    frm_Query_ReceiptDetail.Visible = False: frm_Query_ReceiptDetail.WindowState = 2
'    frm_Query_ReceiptDetail.Visible = True
'    frm_Query_ReceiptDetail.ZOrder
'    frm_Query_ReceiptDetail.Tag = "入庫明細資料查詢"
'    Call UpdateMDIForm_Menu_WindowName
'End Sub

'Private Sub Menu_Query_ShipDetail_Click()
'    '查詢 → 出貨明細資料查詢
'    If CheckOpenForm("出貨明細資料查詢") = 1 Then Exit Sub
'    Load frm_Query_ShipDetail
'    frm_Query_ShipDetail.Visible = False: frm_Query_ShipDetail.WindowState = 2
'    frm_Query_ShipDetail.Visible = True
'    frm_Query_ShipDetail.ZOrder
'    frm_Query_ShipDetail.Tag = "出貨明細資料查詢"
'    Call UpdateMDIForm_Menu_WindowName
'End Sub

Private Sub Menu_Report_DeliveryTrack_Click()
    'FTP上下傳 → 客戶到貨追蹤表
    If CheckOpenForm("客戶到貨追蹤表") = 1 Then Exit Sub
    Load frm_Report_DeliveryTrack
    frm_Report_DeliveryTrack.Visible = False: frm_Report_DeliveryTrack.WindowState = 2
    frm_Report_DeliveryTrack.Visible = True
    frm_Report_DeliveryTrack.ZOrder
    frm_Report_DeliveryTrack.Tag = "客戶到貨追蹤表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_DelOrder_Click()
    'TK需求報表 → 訂單刪除明細
    If CheckOpenForm("訂單刪除明細") = 1 Then Exit Sub
    Load frm_Report_DelOrder
    frm_Report_DelOrder.Visible = False: frm_Report_DelOrder.WindowState = 2
    frm_Report_DelOrder.Visible = True
    frm_Report_DelOrder.ZOrder
    frm_Report_DelOrder.Tag = "訂單刪除明細"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_DivideSku_Click()
    '報表 → 花王分貨表
    If CheckOpenForm("花王分貨表") = 1 Then Exit Sub
    Load frm_Report_DivideSku
    frm_Report_DivideSku.Visible = False: frm_Report_DivideSku.WindowState = 2
    frm_Report_DivideSku.Visible = True
    frm_Report_DivideSku.ZOrder
    frm_Report_DivideSku.Tag = "花王分貨表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_MBO_Cod_Click()
    '報表 → POD回傳
    If CheckOpenForm("毛寶下貨收現查詢") = 1 Then Exit Sub
    Load frm_Report_MBO_Cod
    frm_Report_MBO_Cod.Visible = False: frm_Report_MBO_Cod.WindowState = 2
    frm_Report_MBO_Cod.Visible = True
    frm_Report_MBO_Cod.ZOrder
    frm_Report_MBO_Cod.Tag = "毛寶下貨收現查詢"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_MboReport_PodRetrun_Click()
    '報表 → POD回傳
    If CheckOpenForm("POD回傳") = 1 Then Exit Sub
    Load frm_Report_PodRetrun
    frm_Report_PodRetrun.Visible = False: frm_Report_PodRetrun.WindowState = 2
    frm_Report_PodRetrun.Visible = True
    frm_Report_PodRetrun.ZOrder
    frm_Report_PodRetrun.Tag = "POD回傳"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_MboReport_SDNReturnList_Click()
    '報表 → POD回傳
    If CheckOpenForm("毛寶回單檢核表") = 1 Then Exit Sub
    Load frm_Report_MBO_SDNReturnList
    frm_Report_MBO_SDNReturnList.Visible = False: frm_Report_MBO_SDNReturnList.WindowState = 2
    frm_Report_MBO_SDNReturnList.Visible = True
    frm_Report_MBO_SDNReturnList.ZOrder
    frm_Report_MBO_SDNReturnList.Tag = "毛寶回單檢核表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_SDNReturnList_Click()
    '報表 → 回單檢核表
    If CheckOpenForm("回單檢核表") = 1 Then Exit Sub
    Load frm_Report_SDNReturnList
    frm_Report_SDNReturnList.Visible = False: frm_Report_SDNReturnList.WindowState = 2
    frm_Report_SDNReturnList.Visible = True
    frm_Report_SDNReturnList.ZOrder
    frm_Report_SDNReturnList.Tag = "送貨回單檢核表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_Ship2TKK_Click()
    'TK需求報表 → 出貨資料回傳
    If CheckOpenForm("出貨資料回傳") = 1 Then Exit Sub
    Load frm_Report_Ship2TKK
    frm_Report_Ship2TKK.Visible = False: frm_Report_Ship2TKK.WindowState = 2
    frm_Report_Ship2TKK.Visible = True
    frm_Report_Ship2TKK.ZOrder
    frm_Report_Ship2TKK.Tag = "出貨資料回傳"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_Report_TKARList_Click()
'TK需求報表 → 客戶應收明細表
    If CheckOpenForm("客戶應收明細表") = 1 Then Exit Sub
    Load frm_Report_TKARList
    frm_Report_TKARList.Visible = False: frm_Report_TKARList.WindowState = 2
    frm_Report_TKARList.Visible = True
    frm_Report_TKARList.ZOrder
    frm_Report_TKARList.Tag = "客戶應收明細表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TKCustomerCodeDate_Click()
    'TK需求報表 → 客戶進貨有效期限明細表
    If CheckOpenForm("客戶進貨有效期限明細表") = 1 Then Exit Sub
    Load frm_Report_TKCustomerCodeDate
    frm_Report_TKCustomerCodeDate.Visible = False: frm_Report_TKCustomerCodeDate.WindowState = 2
    frm_Report_TKCustomerCodeDate.Visible = True
    frm_Report_TKCustomerCodeDate.ZOrder
    frm_Report_TKCustomerCodeDate.Tag = "客戶進貨有效期限明細表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TKExpect1_Click()
    'TK需求報表 → 訂單配送異常表
    If CheckOpenForm("訂單配送異常表") = 1 Then Exit Sub
    Load frm_Report_TKExpect1
    frm_Report_TKExpect1.Visible = False: frm_Report_TKExpect1.WindowState = 2
    frm_Report_TKExpect1.Visible = True
    frm_Report_TKExpect1.ZOrder
    frm_Report_TKExpect1.Tag = "訂單配送異常表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TKKPI_Click()
    'TK需求報表 → 單量明細
    If CheckOpenForm("單量明細") = 1 Then Exit Sub
    Load frm_Report_TKKPI
    frm_Report_TKKPI.Visible = False: frm_Report_TKKPI.WindowState = 2
    frm_Report_TKKPI.Visible = True
    frm_Report_TKKPI.ZOrder
    frm_Report_TKKPI.Tag = "單量明細"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TKKRSDNReturnList_Click()
    'TK需求報表 → 退貨回單檢核表
    If CheckOpenForm("退貨回單檢核表") = 1 Then Exit Sub
    Load frm_Report_TKRSDNReturnList
    frm_Report_TKRSDNReturnList.Visible = False: frm_Report_TKRSDNReturnList.WindowState = 2
    frm_Report_TKRSDNReturnList.Visible = True
    frm_Report_TKRSDNReturnList.ZOrder
    frm_Report_TKRSDNReturnList.Tag = "退貨回單檢核表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TKKSDNReturnList_Click()
    'TK需求報表 → 送貨回單檢核表
    If CheckOpenForm("送貨回單檢核表") = 1 Then Exit Sub
    Load frm_Report_TKSDNReturnList
    frm_Report_TKSDNReturnList.Visible = False: frm_Report_TKSDNReturnList.WindowState = 2
    frm_Report_TKSDNReturnList.Visible = True
    frm_Report_TKSDNReturnList.ZOrder
    frm_Report_TKSDNReturnList.Tag = "送貨回單檢核表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TMSAbnormal_Click()
    
    '報表 → 訂單配送異常表表
    If CheckOpenForm("配送異常表") = 1 Then Exit Sub
    Load frm_Report_TMSAbnormal
    frm_Report_TMSAbnormal.Visible = False: frm_Report_TMSAbnormal.WindowState = 2
    frm_Report_TMSAbnormal.Visible = True
    frm_Report_TMSAbnormal.ZOrder
    frm_Report_TMSAbnormal.Tag = "配送異常表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_APPSdnDetail_Click()
    
    '報表 → 訂單配送異常表表
    If CheckOpenForm("簽單明細表") = 1 Then Exit Sub
    Load frm_Report_APPSdnDetail
    frm_Report_APPSdnDetail.Visible = False: frm_Report_APPSdnDetail.WindowState = 2
    frm_Report_APPSdnDetail.Visible = True
    frm_Report_APPSdnDetail.ZOrder
    frm_Report_APPSdnDetail.Tag = "簽單明細表"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TRPTrack_Click()
    '報表→到貨追蹤表
    If CheckOpenForm("到貨追蹤表") = 1 Then Exit Sub
    
    Load frm_Report_TRPTrack
    frm_Report_TRPTrack.Visible = False: frm_Report_TRPTrack.WindowState = 2
    frm_Report_TRPTrack.Visible = True
    frm_Report_TRPTrack.ZOrder
    frm_Report_TRPTrack.Tag = "到貨追蹤表"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_SwitchDB_Click()
    '資料庫切換
    If CheckOpenForm("資料庫切換") = 1 Then Exit Sub
    Load frm_SwitchDB
    frm_SwitchDB.Visible = False: frm_SwitchDB.WindowState = vbNormal
    frm_SwitchDB.Visible = True
    frm_SwitchDB.ZOrder
    frm_SwitchDB.Tag = "資料庫切換"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_SystemUpdate_Click()
'系統更新
If CheckOpenForm("系統更新") = 1 Then Exit Sub
   Load frm_SystemUpdate
   frm_SystemUpdate.Visible = False: frm_SystemUpdate.WindowState = vbNormal
   frm_SystemUpdate.Visible = True
   frm_SystemUpdate.ZOrder
   frm_SystemUpdate.Tag = "系統更新"
   Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_Report_TKExpect_Click()
    'TK需求報表 → 訂單退回明細
    If CheckOpenForm("訂單退回明細") = 1 Then Exit Sub
    Load frm_Report_TKExpect
    frm_Report_TKExpect.Visible = False: frm_Report_TKExpect.WindowState = 2
    frm_Report_TKExpect.Visible = True
    frm_Report_TKExpect.ZOrder
    frm_Report_TKExpect.Tag = "訂單退回明細"
    Call UpdateMDIForm_Menu_WindowName
End Sub


Private Sub Menu_THLReport_Click()
    '報表 → THL需求報表
    If CheckOpenForm("THL需求報表") = 1 Then Exit Sub
    Dim obj As Object
    Set obj = frm_Report_THL
    Load obj
    obj.Visible = False
    obj.Visible = True
    obj.ZOrder
    obj.Tag = "VLT需求報表"
    obj.WindowState = 2
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_TRPPlan_BacktoEXE_Click()
    '排車處理作業 → 訂單排車資料回傳設定
    If CheckOpenForm("訂單排車資料回傳設定") = 1 Then Exit Sub
    Load frm_OP_BacktoEXE
    frm_OP_BacktoEXE.Visible = False: frm_OP_BacktoEXE.WindowState = 2
    frm_OP_BacktoEXE.Visible = True
    frm_OP_BacktoEXE.ZOrder
    frm_OP_BacktoEXE.Tag = "訂單排車資料回傳設定"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_Cost_Click()
    '排車處理作業 → 運費分析
    If CheckOpenForm("運費分析") = 1 Then Exit Sub
    Load frm_OP_TRPCost
    frm_OP_TRPCost.Visible = False: frm_OP_TRPCost.WindowState = 2
    frm_OP_TRPCost.Visible = True
    frm_OP_TRPCost.ZOrder
    frm_OP_TRPCost.Tag = "運費分析"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_TRPPlan_CutOrders_Click()
    '排車處理作業 → ㄧ單多車訂單切割
    If CheckOpenForm("ㄧ單多車訂單切割") = 1 Then Exit Sub
    Load frm_OP_CutOrders
    frm_OP_CutOrders.Visible = False: frm_OP_CutOrders.WindowState = 2
    frm_OP_CutOrders.Visible = True
    frm_OP_CutOrders.ZOrder
    frm_OP_CutOrders.Tag = "ㄧ單多車訂單切割"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_DCRouteMerge_Click()
    '排車處理作業 →  二次排車作業
    If CheckOpenForm("二次排車作業") = 1 Then Exit Sub
    Load frm_OP_DCRouteMerge
    frm_OP_DCRouteMerge.Visible = False: frm_OP_DCRouteMerge.WindowState = 2
    frm_OP_DCRouteMerge.Visible = True
    frm_OP_DCRouteMerge.ZOrder
    frm_OP_DCRouteMerge.Tag = "二次排車作業"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_ManualOrders_Click()
    '排車處理作業 → 訂單維護作業
    If CheckOpenForm("訂單維護作業") = 1 Then Exit Sub
    Load frm_OP_ManualOrders
    frm_OP_ManualOrders.Visible = False: frm_OP_ManualOrders.WindowState = 2 '  = vbNormal
    frm_OP_ManualOrders.Visible = True
    frm_OP_ManualOrders.ZOrder
    frm_OP_ManualOrders.Tag = "訂單維護作業"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_OrderImport_Click()
    '排車處理作業 → 訂單轉入及客戶異動維護
    If CheckOpenForm("訂單轉入及客戶異動維護") = 1 Then Exit Sub
    Load frm_OP_OrderImport
    frm_OP_OrderImport.Visible = False: frm_OP_OrderImport.WindowState = 2 '  = vbNormal
    frm_OP_OrderImport.Visible = True
    frm_OP_OrderImport.ZOrder
    frm_OP_OrderImport.Tag = "訂單轉入及客戶異動維護"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_Query_Click()
    '排車處理作業 → 訂單查詢作業
    If CheckOpenForm("訂單查詢作業") = 1 Then Exit Sub
    Load frm_Query_Orders
    frm_Query_Orders.Visible = False: frm_Query_Orders.WindowState = 2 '  = vbNormal
    frm_Query_Orders.Visible = True
    frm_Query_Orders.ZOrder
    frm_Query_Orders.Tag = "訂單查詢作業"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_Report_Click()
    '排車處理作業 → 排車作業報表
    If CheckOpenForm("排車作業報表") = 1 Then Exit Sub
    Load frm_Report_TRPPlan
    frm_Report_TRPPlan.Visible = False: frm_Report_TRPPlan.WindowState = 2
    frm_Report_TRPPlan.Visible = True
    frm_Report_TRPPlan.ZOrder
    frm_Report_TRPPlan.Tag = "排車作業報表"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_ReDelivery_Click()
    '排車處理作業 → 未收訂單在配送作業
    If CheckOpenForm("未收訂單在配送作業") = 1 Then Exit Sub
    Load frm_OP_ReDelivery
    frm_OP_ReDelivery.Visible = False: frm_OP_ReDelivery.WindowState = 2
    frm_OP_ReDelivery.Visible = True
    frm_OP_ReDelivery.ZOrder
    frm_OP_ReDelivery.Tag = "未收訂單在配送作業"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_Route_Click()
    '排車處理作業 → 路線編號維護作業
    If CheckOpenForm("路線編號維護作業") = 1 Then Exit Sub
    Load frm_OP_RouteData
    frm_OP_RouteData.Visible = False: frm_OP_RouteData.WindowState = 2 '  = vbNormal
    frm_OP_RouteData.Visible = True
    frm_OP_RouteData.ZOrder
    frm_OP_RouteData.Tag = "路線編號維護作業"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_RouteConfirm_Click()
    '排車處理作業 →  出車確認
    If CheckOpenForm("出車確認") = 1 Then Exit Sub
    Load frm_OP_RouteConfirm
    frm_OP_RouteConfirm.Visible = False: frm_OP_RouteConfirm.WindowState = 2
    frm_OP_RouteConfirm.Visible = True
    frm_OP_RouteConfirm.ZOrder
    frm_OP_RouteConfirm.Tag = "出車確認"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_SDNAbnormal_Click()
    '排車處理作業 →  配送異常維護
    If CheckOpenForm("配送異常維護") = 1 Then Exit Sub
    Load frm_OP_SDNAbnormal
    frm_OP_SDNAbnormal.Visible = False: frm_OP_SDNAbnormal.WindowState = 2
    frm_OP_SDNAbnormal.Visible = True
    frm_OP_SDNAbnormal.ZOrder
    frm_OP_SDNAbnormal.Tag = "配送異常維護"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_TRPPlan_SDNConfirm_Click()
    '排車處理作業 →  簽單確認
    If CheckOpenForm("簽單確認") = 1 Then Exit Sub
    Load frm_OP_SDNConfirm
    frm_OP_SDNConfirm.Visible = False: frm_OP_SDNConfirm.WindowState = 2
    frm_OP_SDNConfirm.Visible = True
    frm_OP_SDNConfirm.ZOrder
    frm_OP_SDNConfirm.Tag = "簽單確認"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_ShipQty_Click()
    '排車處理作業 →  揀貨數量確認
    If CheckOpenForm("揀貨數量確認") = 1 Then Exit Sub
    Load frm_OP_ShipQty
    frm_OP_ShipQty.Visible = False: frm_OP_ShipQty.WindowState = 2
    frm_OP_ShipQty.Visible = True
    frm_OP_ShipQty.ZOrder
    frm_OP_ShipQty.Tag = "揀貨數量確認"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_TRPPlan_TRPPlan_Click()
    '排車處理作業 →  排車作業
    If CheckOpenForm("排車作業") = 1 Then Exit Sub
    Load frm_OP_TRPPlan
    frm_OP_TRPPlan.Visible = False: frm_OP_TRPPlan.WindowState = 2
    frm_OP_TRPPlan.Visible = True
    frm_OP_TRPPlan.ZOrder
    frm_OP_TRPPlan.Tag = "排車作業"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_Upload_FTP_Click()
    '上下傳作業 → FTP 上下傳
    If CheckOpenForm("FTP上下傳") = 1 Then Exit Sub
    Load frm_FTP
    frm_FTP.Visible = False
    frm_FTP.WindowState = 2 '  = vbNormal
    frm_FTP.Visible = True
    frm_FTP.ZOrder
    frm_FTP.Tag = "FTP上下傳"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_Query_KPI_KPI_Click()
    '管理KPI → 每日KPI
    If CheckOpenForm("每日KPI") = 1 Then Exit Sub
    Load frm_Query_KPI
    frm_Query_KPI.Visible = False
    frm_Query_KPI.WindowState = 2 '  = vbNormal
    frm_Query_KPI.Visible = True
    frm_Query_KPI.ZOrder
    frm_Query_KPI.Tag = "每日KPI"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_VTLReport_Click()

    '報表 → VLT需求報表
    If CheckOpenForm("VLT需求報表") = 1 Then Exit Sub
    Dim obj As Object
    Set obj = frm_Report_VTL
    Load obj
    obj.Visible = False
    obj.Visible = True
    obj.ZOrder
    obj.Tag = "VLT需求報表"
    obj.WindowState = 2
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_WindowMinx_Click()
'Menu 之 [視窗]→[最小化]
Dim i As Integer
For i = 0 To Forms.Count - 1
    If Not (Forms(i) Is frm_MDIForm) Then
       Forms(i).WindowState = vbMinimized
    End If
Next i
End Sub

Private Sub Menu_WindowSourceSizex_Click()
'Menu 之 [視窗]→[原始視窗]
Dim i As Integer
Dim frmHeight As Long, frmWidth As Long, frmTopNum As Long
For i = 0 To Forms.Count - 1
    If Not (Forms(i) Is frm_MDIForm) Then
       Forms(i).WindowState = 2
    End If
Next i
End Sub

Private Sub Disable_Menu()
'預設動作：Disable 所有功能表
Dim obj As Object
   For Each obj In frm_MDIForm.Controls
       If TypeName(obj) = "Menu" Then
          If Right(Trim(obj.Name), 1) <> "x" Then obj.Enabled = False
       End If
   Next
   Menu_Exitx.Enabled = True
End Sub

Private Sub HideMenu()
'*****************************
'由ini檔定義隱藏選單
'Create by Gemini @20070416
'
'
'
'*****************************
'取參數
Dim objIni As vbIniFile, arrTmp
Set objIni = New vbIniFile

With objIni

    .FileName = striniFileName_FullPath
    Dim i As Integer, obj As Object
    
    arrTmp = Split(.ReadData("OPTION", "HIDEMENU", "0"), ";")
    
    For Each obj In frm_MDIForm.Controls
      For i = 0 To UBound(arrTmp)
        If TypeName(obj) = "Menu" Then If Trim(obj.Caption) = Trim(arrTmp(i)) Then obj.Visible = False
      Next i
    Next

End With

Set objIni = Nothing
   
End Sub
Private Sub mnuWindowCascadex_Click()
    Me.Arrange vbCascade
End Sub
Private Sub mnuWindowTileHorizontalx_Click()
    Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuWindowTileVerticalx_Click()
    Me.Arrange vbTileVertical
End Sub
Private Sub mnuWindowArrangeIconsx_Click()
    Me.Arrange vbArrangeIcons
End Sub
