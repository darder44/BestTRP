VERSION 5.00
Begin VB.MDIForm frm_MDIForm 
   BackColor       =   &H8000000C&
   Caption         =   "���������t��"
   ClientHeight    =   5925
   ClientLeft      =   915
   ClientTop       =   2010
   ClientWidth     =   11370
   Icon            =   "frm_MDIForm.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frm_MDIForm.frx":0442
   WindowState     =   2  '�̤j��
   Begin VB.Menu Menu_Orders 
      Caption         =   "�q��B�z�@�~"
      Begin VB.Menu Menu_Upload_FTP 
         Caption         =   "�q�汵��"
      End
      Begin VB.Menu Menu_TRPPlan_ManualOrders 
         Caption         =   "�q����@"
      End
      Begin VB.Menu Menu_TRPPlan_OrderImport 
         Caption         =   "�q����J�ΫȤᲧ�ʺ��@"
      End
      Begin VB.Menu Menu_TRPPlan_Query 
         Caption         =   "�q��d�ߧ@�~"
      End
   End
   Begin VB.Menu Menu_TRPPlan 
      Caption         =   "�@��ƨ��@�~"
      Begin VB.Menu Menu_TRPPlan_CutOrders 
         Caption         =   "����h���q�����"
      End
      Begin VB.Menu Menu_TRPPlan_TRPPlan 
         Caption         =   "�@��ƨ��@�~"
      End
      Begin VB.Menu Menu_TRPPlan_DCRouteMerge 
         Caption         =   "�G���ƨ��@�~"
      End
      Begin VB.Menu Menu_TRPPlan_BacktoEXE 
         Caption         =   "�ƨ���Ʀ^�ǳ]�w"
      End
      Begin VB.Menu Menu_TRPPlan_Route 
         Caption         =   "���u�s�����@�@�~"
      End
      Begin VB.Menu Menu_TRPPlan_ReDelivery 
         Caption         =   "�����q��A�t�e�@�~"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_TRPPlan_Report 
         Caption         =   "�ƨ��@�~����"
      End
      Begin VB.Menu Menu_TRPPlan_RouteConfirm 
         Caption         =   "�X���T�{"
      End
      Begin VB.Menu Menu_TRPPlan_SDNAbnormal 
         Caption         =   "�t�e���`���@"
      End
      Begin VB.Menu Menu_TRPPlan_SDNConfirm 
         Caption         =   "ñ��T�{"
      End
      Begin VB.Menu Menu_TRPPlan_ShipQty 
         Caption         =   "�z�f�ƶq�T�{"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_TRPPlan_Cost 
         Caption         =   "�B�O���R"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Menu_Other 
      Caption         =   "�䥦�ƨ��@�~"
      Begin VB.Menu Menu_OP_Other_OrderImport 
         Caption         =   "�q����J�ΫȤᲧ�ʺ��@"
      End
      Begin VB.Menu Menu_Other_ORTPlan 
         Caption         =   "�ƨ��@�~"
      End
      Begin VB.Menu Menu_Other_Report 
         Caption         =   "�ƨ��@�~����"
      End
      Begin VB.Menu Menu_OP_RSDNConfirm 
         Caption         =   "�h�fñ����@"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Menu_Pallet 
      Caption         =   "�䥦�޲z�@�~"
      Begin VB.Menu Menu_BQControlSheet 
         Caption         =   "BQ�ި��"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_PalletxSorting 
         Caption         =   "�̪O�޲z"
      End
      Begin VB.Menu Menu_LoadSorting 
         Caption         =   "½�O�z�f�޲z"
      End
      Begin VB.Menu Menu_Pallet_UTLCst 
         Caption         =   "�g�P�Ӵ̪O�޲z"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_Pallet_Match 
         Caption         =   "�̪O��ƽT�{"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_Pallet_CSVehicle_id_no 
         Caption         =   "���n�Ϩ����פJ"
      End
      Begin VB.Menu Menu_OP_CaseConfirm 
         Caption         =   "�X�f��ƽT�{"
      End
      Begin VB.Menu Menu_Line3x 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Query_Pallet 
         Caption         =   "�̪O��b��"
      End
      Begin VB.Menu Menu_Query_PalletCST 
         Caption         =   "�̪O�έp���l"
      End
      Begin VB.Menu Menu_Query_PalletDetail 
         Caption         =   "�̪O���Ӭd��"
      End
      Begin VB.Menu Menu_Query_loadsortingDetail 
         Caption         =   "½�O�z�fñ����"
      End
      Begin VB.Menu Menu_Query_PalletRent 
         Caption         =   "�����p��"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_OP_PalletExport 
         Caption         =   "�����ƶץX"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_OP_PalletImport 
         Caption         =   "�̪O��ƶפJ"
      End
   End
   Begin VB.Menu Menu_Report 
      Caption         =   "����"
      Begin VB.Menu Menu_Report_DivideSku 
         Caption         =   "������f��"
      End
      Begin VB.Menu Menu_Report_MboReport 
         Caption         =   "���_�ݨD����"
         Begin VB.Menu Menu_Report_MBO_Cod 
            Caption         =   "���_�U�f���{�d��"
         End
         Begin VB.Menu Menu_Report_MboReport_PodRetrun 
            Caption         =   "POD�^��"
         End
         Begin VB.Menu Menu_Report_MboReport_SDNReturnList 
            Caption         =   "�^���ˮ֪�"
         End
      End
      Begin VB.Menu Menu_TKReport 
         Caption         =   "TK�ݨD����"
         Begin VB.Menu Menu_Report_Ship2TKK 
            Caption         =   "�X�f��Ʀ^��"
            Visible         =   0   'False
         End
         Begin VB.Menu Menu_Report_DelOrder 
            Caption         =   "�q��R������"
         End
         Begin VB.Menu Menu_Report_DeliveryTrack 
            Caption         =   "�Ȥ��f�l�ܪ�"
         End
         Begin VB.Menu Menu_Report_TKExpect 
            Caption         =   "�q��h�^����"
         End
         Begin VB.Menu Menu_Report_TKExpect1 
            Caption         =   "�q��t�e���`��"
         End
         Begin VB.Menu Menu_Report_TKCustomerCodeDate 
            Caption         =   "�Ȥ�i�f���Ĵ������Ӫ�"
         End
         Begin VB.Menu Menu_Report_TKKSDNReturnList 
            Caption         =   "�e�f�^���ˮ֪�"
         End
         Begin VB.Menu Menu_Report_TKKRSDNReturnList 
            Caption         =   "�h�f�^���ˮ֪�"
         End
         Begin VB.Menu Menu_Report_TKKPI 
            Caption         =   "��q����"
         End
         Begin VB.Menu Menu_Report_TKARList 
            Caption         =   "�����b�ک��Ӫ�"
         End
      End
      Begin VB.Menu Menu_VTLReport 
         Caption         =   "VTL�ݨD����"
      End
      Begin VB.Menu Menu_THLReport 
         Caption         =   "THL�ݨD����"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_NSLReport 
         Caption         =   "NSL�ݨD����"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_ABTReport 
         Caption         =   "ABT�ݨD����"
      End
      Begin VB.Menu Menu_Report_SDNReturnList 
         Caption         =   "�^���ˮ֪�"
      End
      Begin VB.Menu Menu_Report_TRPTrack 
         Caption         =   "��f�l�ܪ�"
      End
      Begin VB.Menu Menu_Report_TMSAbnormal 
         Caption         =   "�t�e���`��"
      End
      Begin VB.Menu Menu_Report_APPSdnDetail 
         Caption         =   "ñ����Ӫ�"
      End
   End
   Begin VB.Menu Menu_Query 
      Caption         =   "�d��"
      Begin VB.Menu Menu_Query_InterfaceLog 
         Caption         =   "InterFaceLog"
      End
      Begin VB.Menu Menu_Query_KPI 
         Caption         =   "�޲zKPI"
         Begin VB.Menu Menu_Query_KPI_KPI 
            Caption         =   "�C��KPI"
         End
         Begin VB.Menu Menu_Query_KPI_CarCount 
            Caption         =   "�C��ϰ쨮������KPI"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu Menu_Query_Charge 
         Caption         =   "�ХI�ڤ����"
      End
      Begin VB.Menu Menu_Query_Account_LoadSorting 
         Caption         =   "�|�p½�O�P�z�f���"
      End
      Begin VB.Menu Menu_BackOrderDetail 
         Caption         =   "�h���f�P�ڵu������"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Menu_BaseData 
      Caption         =   "�򥻸��"
      Begin VB.Menu Menu_BaseData_Car 
         Caption         =   "����/�f�B���q"
      End
      Begin VB.Menu Menu_BaseData_ConsigCar 
         Caption         =   "�Ȥ�/����/�f�B���q"
      End
      Begin VB.Menu Menu_DY_BaseData_ConsigCar 
         Caption         =   "�������פJ"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_BaseData_SKU 
         Caption         =   "�Ȱ��f����ƺ��@"
      End
      Begin VB.Menu Menu_BaseData_OP 
         Caption         =   "�@�~�N�X��ƺ��@"
      End
      Begin VB.Menu Menu_BaseData_OP_1 
         Caption         =   "�i���N�X��ƺ��@"
      End
      Begin VB.Menu Menu_BaseData_UserData 
         Caption         =   "�ϥΪ̸�ƺ��@"
      End
      Begin VB.Menu Menu_BaseData_Code 
         Caption         =   "�t�ΥN�X���@"
      End
      Begin VB.Menu Menu_BaseData_UserSecutiry 
         Caption         =   "�t���v���]�w"
      End
   End
   Begin VB.Menu Menu_System 
      Caption         =   "�t�γ]�w"
      Begin VB.Menu Menu_SwitchDB 
         Caption         =   "��Ʈw����"
      End
      Begin VB.Menu Menu_SystemUpdate 
         Caption         =   "�t�Χ�s"
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_Line2x 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Options 
         Caption         =   "�ﶵ"
      End
   End
   Begin VB.Menu Menu_Windowsx 
      Caption         =   "�����ƦC"
      Begin VB.Menu Menu_WindowMinx 
         Caption         =   "�̤p��"
      End
      Begin VB.Menu mnuWindowCascadex 
         Caption         =   "���|���"
      End
      Begin VB.Menu mnuWindowTileHorizontalx 
         Caption         =   "�����ñ�"
      End
      Begin VB.Menu mnuWindowTileVerticalx 
         Caption         =   "�����ñ�"
      End
      Begin VB.Menu mnuWindowArrangeIconsx 
         Caption         =   "�ƦC�ϥ�"
      End
      Begin VB.Menu Menu_WindowSourceSizex 
         Caption         =   "��l����"
      End
      Begin VB.Menu Menu_Line1x 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1�����w"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1�����w"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1�����w"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1�����w"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1�����w"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1�����w"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1�����w"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1�����w"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1�����w"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu Menu_FormNamex 
         Caption         =   "&1�����w"
         Index           =   9
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Menu_Exitx 
      Caption         =   "���}"
   End
End
Attribute VB_Name = "frm_MDIForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strScreenRes As String '�ù��ѪR��

Private Sub MDIForm_Load()

'�ù��ѪR��
 strScreenRes = Screen.Width \ Screen.TwipsPerPixelX & "x" & Screen.Height \ Screen.TwipsPerPixelY
 '����
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

'�T�{codelist�O�_��������
tmp_Rs.Open "select listname from codelist where listname = 'Options'", cn
If tmp_Rs.EOF Then cn.Execute "insert into codelist(listname,description,adddate,addwho,editdate,editwho) values ('Options','�ﶵ�]�w��',getdate(),'dbo',getdate(),'dbo')", RowsAffect, adExecuteNoRecords
tmp_Rs.Close

  '�Ѹ�Ʈw�ѼƳ]�w�A���o Security Control �]�w��
  blSecurityControl = True
  cn.Execute "select listname from codelkup where listname = 'Options' and code = 'LoginControl' and Description = 0 ", RowsAffect, adExecuteNoRecords
  If RowsAffect <> 0 Then blSecurityControl = False: User_id = strComputerName: blAdmin = True 'RouteModify= 0 �ɪ�ܸ�Ʈw�L�۲Ÿ��
  
  '�ƨ���ƭק�O�_�����ϥΪ�
  cn.Execute "select listname from codelkup where listname = 'Options' and code = 'RouteModify' and Description = 0 ", RowsAffect, adExecuteNoRecords
  If RowsAffect = 0 Then blRouteModifyControl = True '�S���ɭn����
 
'��ƺ��@����
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
   '��Ҧ��\����]�� Disable�A�A�̨ϥΪ��v���]�w�ȳv�@�}�ҡGenable
   '�q���W�٬OBESTPREPARES or gemini�h���}�A�׶}�}�o�ɪ������P�����D
   If UCase(strComputerName) <> "BESTPREPARES" And UCase(strComputerName) <> "BEST_ALICENB" And UCase(strComputerName) <> "BEST-TERRY" And UCase(strComputerName) <> "BEST-TEST" And UCase(strComputerName) <> "GEMINI_NB" And UCase(strComputerName) <> "GEMINI_VPC" Then
         '�ˬd�t�Ϊ����O�_���̷s�����A���O�h�ݤ�ʧ�s�A���M�L�k�ϥ�
        If RTrim(App.EXEName) = "BestTRP" Then
          '���`����TMS
          '�ˬd�����s���O�_���̷s����
          tmp_Rs.Open "select top 1 version from VersionCheck where project = 'BestTms' order by adddate desc", cn, adOpenForwardOnly, adLockReadOnly
          If RTrim(tmp_Rs.Fields("version")) = RTrim(App.Major & "." & App.Minor & "." & App.Revision) Then
              tmp_Rs.Close
          Else
              MsgBox "TMS���s�����o�G:" & RTrim(tmp_Rs.Fields("version")) & "�A�������z��TMS�ç�s�z��TMS�A�T�O�t�Ϊ����T��!" & Chr(13) & "�_�h���F��ƪ����T�ʡA�z�N�L�k�~��ϥ�!", vbOKOnly + vbCritical, "TMS�����ˬd"
              tmp_Rs.Close
              Exit Sub
          End If
        Else
          '�ª��ƥΪ�TMS
          MsgBox "�A�ϥΪ������OTMS old�����A�нT�{���|�����|�O�_���T!" & Chr(13) & "�z���i�~��ϥ�!����ĳ�z�ϥγ̷s�����A�T�O��ƥ��T��!" & Chr(9) & Chr(9) & Chr(9) & "�t�αN�����z���b���������]�֡C", vbOKOnly + vbExclamation, "TMS�����ˬd"
          str_SQL = "Insert into gt_Logs(APName,APVer,APCaption,Code,Description,Notes,ComputerName,AddWho) Values ('" & _
                        App.EXEName & "','" & App.Major & "." & App.Minor & "." & App.Revision & "','" & Me.Caption & "','" & "" & "','" & "�ϥ��ª�TMS�t��" & "','" & "�ϥ��ª�TMS�t��" & "','" & strComputerName & "','" & User_id & "')"
          cn.Execute str_SQL, RowsAffect, adExecuteNoRecords
        End If
    End If
    
     Load frm_UserLogin
     frm_UserLogin.Visible = False: frm_UserLogin.WindowState = vbNormal
     frm_UserLogin.Visible = True
     frm_UserLogin.ZOrder
     frm_UserLogin.Tag = "�t�εn�J"
End If

  Call HideMenu

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Call DB_Disconnect(cn)
  End
End Sub

Private Sub Menu_ABTReport_Click()
    '���� �� ABT�ݨD����
    If CheckOpenForm("ABT�ݨD����") = 1 Then Exit Sub
    Dim obj As Object
    Set obj = frm_Report_ABT
    Load obj
    obj.Visible = False
    obj.Visible = True
    obj.ZOrder
    obj.Tag = "ABT�ݨD����"
    obj.WindowState = 2
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_BackOrderDetail_Click()

    '�d�� �� �h���f�P�ڵu������
    If CheckOpenForm("�h���f�P�ڵu������") = 1 Then Exit Sub
    Load frm_Query_BackOrderDetail
    frm_Query_BackOrderDetail.Visible = False: frm_Query_BackOrderDetail.WindowState = 2
    frm_Query_BackOrderDetail.Visible = True
    frm_Query_BackOrderDetail.ZOrder
    frm_Query_BackOrderDetail.Tag = "�h���f�P�ڵu������"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BaseData_Code_Click()
    'Menu �򥻸�� �� �l�t�ΥN�X��ƺ��@
    If CheckOpenForm("�l�t�ΥN�X��ƺ��@") = 1 Then Exit Sub
    Load frm_BaseData_Code
    frm_BaseData_Code.Visible = False: frm_BaseData_Code.WindowState = vbNormal
    frm_BaseData_Code.Visible = True
    frm_BaseData_Code.ZOrder
    frm_BaseData_Code.Tag = "�l�t�ΥN�X��ƺ��@"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BaseData_Car_Click()
    'Menu �򥻸�� �� �����򥻸�ƺ��@
    If CheckOpenForm("�����򥻸�ƺ��@") = 1 Then Exit Sub
    Load frm_BaseData_Car
    frm_BaseData_Car.Visible = False: frm_BaseData_Car.WindowState = 2
    frm_BaseData_Car.Visible = True
    frm_BaseData_Car.ZOrder
    frm_BaseData_Car.Tag = "�����򥻸�ƺ��@"
    Call UpdateMDIForm_Menu_WindowName
End Sub
Private Sub Menu_BaseData_ConsigCar_Click()
    'Menu �򥻸�� �� �Ȥ�/�����򥻸�ƺ��@
    If CheckOpenForm("�Ȥ�/�����򥻸�ƺ��@") = 1 Then Exit Sub
    Load frm_BaseData_ConsigCar
    frm_BaseData_ConsigCar.Visible = False: frm_BaseData_ConsigCar.WindowState = 2
    frm_BaseData_ConsigCar.Visible = True
    frm_BaseData_ConsigCar.ZOrder
    frm_BaseData_ConsigCar.Tag = "�Ȥ�/�����򥻸�ƺ��@"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_BaseData_OP_1_Click()
    'Menu �򥻸�� �� �i���N�X��ƺ��@
    If CheckOpenForm("�i���N�X��ƺ��@") = 1 Then Exit Sub
    Load frm_BaseData_OPCode_1
    frm_BaseData_OPCode_1.Visible = False: frm_BaseData_OPCode_1.WindowState = 2
    frm_BaseData_OPCode_1.Visible = True
    frm_BaseData_OPCode_1.ZOrder
    frm_BaseData_OPCode_1.Tag = "�i���N�X��ƺ��@"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BaseData_OP_Click()
    'Menu �򥻸�� �� �@�~�N�X��ƺ��@
    If CheckOpenForm("�@�~�N�X��ƺ��@") = 1 Then Exit Sub
    Load frm_BaseData_OPCode
    frm_BaseData_OPCode.Visible = False: frm_BaseData_OPCode.WindowState = 2
    frm_BaseData_OPCode.Visible = True
    frm_BaseData_OPCode.ZOrder
    frm_BaseData_OPCode.Tag = "�@�~�N�X��ƺ��@"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BaseData_SKU_Click()
    'Menu �򥻸�� �� �Ȱ��ӫ~��ƺ��@
    If CheckOpenForm("�Ȱ��ӫ~��ƺ��@") = 1 Then Exit Sub
    Load frm_BaseData_Sku
    frm_BaseData_Sku.Visible = False: frm_BaseData_Sku.WindowState = vbNormal
    frm_BaseData_Sku.Visible = True
    frm_BaseData_Sku.ZOrder
    frm_BaseData_Sku.Tag = "�Ȱ��ӫ~��ƺ��@"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BaseData_UserData_Click()
    'Menu �򥻸�� �� �ϥΪ̸�ƺ��@
    If CheckOpenForm("�ϥΪ̸�ƺ��@") = 1 Then Exit Sub
    Load frm_BaseData_UserData
    frm_BaseData_UserData.Visible = False: frm_BaseData_UserData.WindowState = vbNormal
    frm_BaseData_UserData.Visible = True
    frm_BaseData_UserData.ZOrder
    frm_BaseData_UserData.Tag = "�ϥΪ̸�ƺ��@"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BaseData_UserSecutiry_Click()
    'Menu �򥻸�� �� �t���v���]�w
    If CheckOpenForm("�t���v���]�w") = 1 Then Exit Sub
    Load frm_BaseData_UserSecurity
    frm_BaseData_UserSecurity.Visible = False: frm_BaseData_UserSecurity.WindowState = vbNormal
    frm_BaseData_UserSecurity.Visible = True
    frm_BaseData_UserSecurity.ZOrder
    frm_BaseData_UserSecurity.Tag = "�t���v���]�w"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_CaeManage_CarControl_Click()
    '�ƨ��B�z�@�~ ��  �����i�X�ި�@�~
    If CheckOpenForm("�����i�X�ި�@�~") = 1 Then Exit Sub
    Load frm_OP_CarControl
    frm_OP_CarControl.Visible = False: frm_OP_CarControl.WindowState = 2
    frm_OP_CarControl.Visible = True
    frm_OP_CarControl.ZOrder
    frm_OP_CarControl.Tag = "�����i�X�ި�@�~"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_BQControlSheet_Click()
    '��L�޲z�@�~��BQ�ި��
    If CheckOpenForm("BQ�ި��") = 1 Then Exit Sub
    
    Load frm_OP_BQControlSheet
    frm_OP_BQControlSheet.Visible = False: frm_OP_BQControlSheet.WindowState = 2
    frm_OP_BQControlSheet.Visible = True
    frm_OP_BQControlSheet.ZOrder
    frm_OP_BQControlSheet.Tag = "BQ�ި��"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Exitx_Click()
  Call DB_Disconnect(cn)
  End
End Sub

Private Sub Menu_FormNamex_Click(Index As Integer)
'Menu �� [����]��[�w��ܵ���]
'�N�Q��������ը�̫e��
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
    '��L�޲z�@�~��½�O�z�f�޲z
    If CheckOpenForm("½�O�z�f�޲z") = 1 Then Exit Sub
    
    Load frm_OP_LoadSorting
    frm_OP_LoadSorting.Visible = False: frm_OP_LoadSorting.WindowState = 2
    frm_OP_LoadSorting.Visible = True
    frm_OP_LoadSorting.ZOrder
    frm_OP_LoadSorting.Tag = "½�O�z�f�޲z"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_NSLReport_Click()
    '���� �� NSL�ݨD����
    If CheckOpenForm("NSL�ݨD����") = 1 Then Exit Sub
    Dim obj As Object
    Set obj = frm_Report_NSL
    Load obj
    obj.Visible = False
    obj.Visible = True
    obj.ZOrder
    obj.Tag = "NSL�ݨD����"
    obj.WindowState = 2
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_OP_CaseConfirm_Click()
    '��L�޲z�@�~
    If CheckOpenForm("�X�f��ƽT�{") = 1 Then Exit Sub
    
    Load frm_OP_CaseConfirm
'    frm_OP_Other_OrderImport.Visible = False: frm_OP_Other_OrderImport.WindowState = 2
'    frm_OP_Other_OrderImport.Visible = True
'    frm_OP_Other_OrderImport.ZOrder
    frm_OP_CaseConfirm.Tag = frm_OP_CaseConfirm.Caption
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_OP_Other_OrderImport_Click()
    '�h�f�q����J�ΫȤᲧ�ʺ��@
    If CheckOpenForm("�h�f�q����J�ΫȤᲧ�ʺ��@") = 1 Then Exit Sub
    
    Load frm_OP_Other_OrderImport
    frm_OP_Other_OrderImport.Visible = False: frm_OP_Other_OrderImport.WindowState = 2
    frm_OP_Other_OrderImport.Visible = True
    frm_OP_Other_OrderImport.ZOrder
    frm_OP_Other_OrderImport.Tag = "�h�f�q����J�ΫȤᲧ�ʺ��@"
    Call UpdateMDIForm_Menu_WindowName
End Sub



Private Sub Menu_OP_RSDNConfirm_Click()
    '�h�f�ƨ�
    If CheckOpenForm("�h�fñ����@") = 1 Then Exit Sub
    
    Load frm_OP_RSDNConfirm
    frm_OP_RSDNConfirm.Visible = False: frm_OP_RSDNConfirm.WindowState = 2
    frm_OP_RSDNConfirm.Visible = True
    frm_OP_RSDNConfirm.ZOrder
    frm_OP_RSDNConfirm.Tag = "�h�fñ����@"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Options_Click()
    '�ﶵ
    If CheckOpenForm("�ﶵ") = 1 Then Exit Sub
    
    Load frm_Options
    frm_Options.Visible = False ': frm_Options.WindowState = 2
    frm_Options.Visible = True
    frm_Options.ZOrder
    frm_Options.Tag = "�ﶵ"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Other_ORTPlan_Click()
    '�h�f�ƨ�
    If CheckOpenForm("�h�f�ƨ�") = 1 Then Exit Sub
    
    Load frm_Other_OPTPlan
    frm_Other_OPTPlan.Visible = False: frm_Other_OPTPlan.WindowState = 2
    frm_Other_OPTPlan.Visible = True
    frm_Other_OPTPlan.ZOrder
    frm_Other_OPTPlan.Tag = "�h�f�ƨ�"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Other_Report_Click()
    '�h�f����
    If CheckOpenForm("�h�f����") = 1 Then Exit Sub
    Load frm_Report_Other
    frm_Report_Other.Visible = False: frm_Report_Other.WindowState = 2
    frm_Report_Other.Visible = True
    frm_Report_Other.ZOrder
    frm_Report_Other.Tag = "�h�f����"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Pallet_CSVehicle_id_no_Click()
    '��L�޲z�@�~
    If CheckOpenForm("���n�Ϩ����פJ") = 1 Then Exit Sub
    
    Load frm_Pallet_CSVehicle_id_no
    frm_Pallet_CSVehicle_id_no.Tag = frm_Pallet_CSVehicle_id_no.Caption
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Pallet_Match_Click()
    '�̪O��ƽT�{
    If CheckOpenForm("�̪O��ƽT�{") = 1 Then Exit Sub
    Load frm_Pallet_Match
    frm_Pallet_Match.Visible = False: frm_Pallet_Match.WindowState = 2
    frm_Pallet_Match.Visible = True
    frm_Pallet_Match.ZOrder
    frm_Pallet_Match.Tag = "�̪O��ƽT�{"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Pallet_UTL_Click()
    '�̪O�޲z
    If CheckOpenForm("�̪O�޲z") = 1 Then Exit Sub
    Load frm_Pallet_UTL
    frm_Pallet_UTL.Visible = False: frm_Pallet_UTL.WindowState = 2
    frm_Pallet_UTL.Visible = True
    frm_Pallet_UTL.ZOrder
    frm_Pallet_UTL.Tag = "�̪O�޲z"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Pallet_UTLCst_Click()
    '�g�P�Ӵ̪O�޲z
    If CheckOpenForm("�g�P�Ӵ̪O�޲z") = 1 Then Exit Sub
    Load frm_Pallet_UTLCst
    frm_Pallet_UTLCst.Visible = False: frm_Pallet_UTLCst.WindowState = 2
    frm_Pallet_UTLCst.Visible = True
    frm_Pallet_UTLCst.ZOrder
    frm_Pallet_UTLCst.Tag = "�g�P�Ӵ̪O�޲z"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_PalletxSorting_Click()
    '��L�޲z�@�~->�̪O�P�z�f�޲z
    If CheckOpenForm("�̪O�P�z�f�޲z") = 1 Then Exit Sub
    Load frm_OP_PalletxSorting
    frm_OP_PalletxSorting.Visible = False: frm_OP_PalletxSorting.WindowState = 2
    frm_OP_PalletxSorting.Visible = True
    frm_OP_PalletxSorting.ZOrder
    frm_OP_PalletxSorting.Tag = "�̪O�P�z�f�޲z"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_Account_LoadSorting_Click()
    '�d��->½�O�P�z�f���
    If CheckOpenForm("½�O�P�z�f���") = 1 Then Exit Sub
    Load frm_Query_Account_LoadSorting
    frm_Query_Account_LoadSorting.Visible = False: frm_Query_Account_LoadSorting.WindowState = 2
    frm_Query_Account_LoadSorting.Visible = True
    frm_Query_Account_LoadSorting.ZOrder
    frm_Query_Account_LoadSorting.Tag = "½�O�P�z�f���"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_Charge_Click()

    '�d��->�Ȥ�дڸ��
    If CheckOpenForm("�Ȥ�дڸ��") = 1 Then Exit Sub
    Load frm_Query_Charge
    frm_Query_Charge.Visible = False: frm_Query_Charge.WindowState = 2
    frm_Query_Charge.Visible = True
    frm_Query_Charge.ZOrder
    frm_Query_Charge.Tag = "�Ȥ�дڸ��"
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
'    '�d�� �� �Y�ɮw�s�d��
'    If CheckOpenForm("�Y�ɮw�s�d��") = 1 Then Exit Sub
'    Load frm_Query_Inventory
'    frm_Query_Inventory.Visible = False
'    frm_Query_Inventory.WindowState = 2 '  = vbNormal
'    frm_Query_Inventory.Visible = True
'    frm_Query_Inventory.ZOrder
'    frm_Query_Inventory.Tag = "�Y�ɮw�s�d��"
'    Call UpdateMDIForm_Menu_WindowName
'End Sub

Private Sub Menu_Query_KPI_CarCount_Click()
    '�޲zKPI �� �C��ϰ쨮������KPI
    If CheckOpenForm("�C��ϰ쨮������KPI") = 1 Then Exit Sub
    Load frm_Query_KPI_CarCount
    frm_Query_KPI_CarCount.Visible = False
    frm_Query_KPI_CarCount.WindowState = 2 '  = vbNormal
    frm_Query_KPI_CarCount.Visible = True
    frm_Query_KPI_CarCount.ZOrder
    frm_Query_KPI_CarCount.Tag = "�C��ϰ쨮������KPI"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_loadsortingDetail_Click()
    '½�O�z�f���Ӭd��
    If CheckOpenForm("½�O�z�f���Ӭd��") = 1 Then Exit Sub
    Load frm_Query_LoadSortingDetail
    frm_Query_LoadSortingDetail.Visible = False: frm_Query_LoadSortingDetail.WindowState = 2
    frm_Query_LoadSortingDetail.Visible = True
    frm_Query_LoadSortingDetail.ZOrder
    frm_Query_LoadSortingDetail.Tag = "½�O�z�f���Ӭd��"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_Pallet_Click()
    '��b��
    If CheckOpenForm("��b��") = 1 Then Exit Sub
    Load frm_Query_Pallet
    frm_Query_Pallet.Visible = False: frm_Query_Pallet.WindowState = 2
    frm_Query_Pallet.Visible = True
    frm_Query_Pallet.ZOrder
    frm_Query_Pallet.Tag = "��b��"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_PalletCST_Click()
    '�έp���l
    If CheckOpenForm("�έp���l") = 1 Then Exit Sub
    Load frm_Query_PalletCst
    frm_Query_PalletCst.Visible = False: frm_Query_PalletCst.WindowState = 2
    frm_Query_PalletCst.Visible = True
    frm_Query_PalletCst.ZOrder
    frm_Query_PalletCst.Tag = "�έp���l"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_PalletDetail_Click()
    '�̪O���Ӭd��
    If CheckOpenForm("�̪O���Ӭd��") = 1 Then Exit Sub
    Load frm_Query_PalletDetail
    frm_Query_PalletDetail.Visible = False: frm_Query_PalletDetail.WindowState = 2
    frm_Query_PalletDetail.Visible = True
    frm_Query_PalletDetail.ZOrder
    frm_Query_PalletDetail.Tag = "�̪O���Ӭd��"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Query_PalletRent_Click()
    '�����p��
    If CheckOpenForm("�����p��") = 1 Then Exit Sub
    Load frm_Query_PalletRent
    frm_Query_PalletRent.Visible = False: frm_Query_PalletRent.WindowState = 2
    frm_Query_PalletRent.Visible = True
    frm_Query_PalletRent.ZOrder
    frm_Query_PalletRent.Tag = "�����p��"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_OP_PalletExport_Click()
    '�̪O��ƶץX
    If CheckOpenForm("�̪O��ƶץX") = 1 Then Exit Sub
    Load frm_OP_PalletExport
    frm_OP_PalletExport.Visible = False: frm_OP_PalletExport.WindowState = 2
    frm_OP_PalletExport.Visible = True
    frm_OP_PalletExport.ZOrder
    frm_OP_PalletExport.Tag = "�̪O��ƶץX"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_OP_PalletImport_Click()
'    '�̪O��ƶפJ
'    If CheckOpenForm("�̪O��ƶפJ") = 1 Then Exit Sub
'    Load frm_OP_PalletImport
'    frm_OP_PalletImport.Visible = False: frm_OP_PalletImport.WindowState = 2
'    frm_OP_PalletImport.Visible = True
'    frm_OP_PalletImport.ZOrder
'    frm_OP_PalletImport.Tag = "�̪O��ƶפJ"
'    Call UpdateMDIForm_Menu_WindowName
    
    '�̪O��ƶפJ
    If CheckOpenForm("�̪O��ƶפJ") = 1 Then Exit Sub
    Load frm_Pallet_Import
    frm_Pallet_Import.Visible = False: frm_Pallet_Import.WindowState = 2
    frm_Pallet_Import.Visible = True
    frm_Pallet_Import.ZOrder
    frm_Pallet_Import.Tag = "�̪O��ƶפJ"
    Call UpdateMDIForm_Menu_WindowName
    
End Sub

'Private Sub Menu_Query_ReceiptDetail_Click()
'    '�d�� �� �J�w���Ӹ�Ƭd��
'    If CheckOpenForm("�J�w���Ӹ�Ƭd��") = 1 Then Exit Sub
'    Load frm_Query_ReceiptDetail
'    frm_Query_ReceiptDetail.Visible = False: frm_Query_ReceiptDetail.WindowState = 2
'    frm_Query_ReceiptDetail.Visible = True
'    frm_Query_ReceiptDetail.ZOrder
'    frm_Query_ReceiptDetail.Tag = "�J�w���Ӹ�Ƭd��"
'    Call UpdateMDIForm_Menu_WindowName
'End Sub

'Private Sub Menu_Query_ShipDetail_Click()
'    '�d�� �� �X�f���Ӹ�Ƭd��
'    If CheckOpenForm("�X�f���Ӹ�Ƭd��") = 1 Then Exit Sub
'    Load frm_Query_ShipDetail
'    frm_Query_ShipDetail.Visible = False: frm_Query_ShipDetail.WindowState = 2
'    frm_Query_ShipDetail.Visible = True
'    frm_Query_ShipDetail.ZOrder
'    frm_Query_ShipDetail.Tag = "�X�f���Ӹ�Ƭd��"
'    Call UpdateMDIForm_Menu_WindowName
'End Sub

Private Sub Menu_Report_DeliveryTrack_Click()
    'FTP�W�U�� �� �Ȥ��f�l�ܪ�
    If CheckOpenForm("�Ȥ��f�l�ܪ�") = 1 Then Exit Sub
    Load frm_Report_DeliveryTrack
    frm_Report_DeliveryTrack.Visible = False: frm_Report_DeliveryTrack.WindowState = 2
    frm_Report_DeliveryTrack.Visible = True
    frm_Report_DeliveryTrack.ZOrder
    frm_Report_DeliveryTrack.Tag = "�Ȥ��f�l�ܪ�"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_DelOrder_Click()
    'TK�ݨD���� �� �q��R������
    If CheckOpenForm("�q��R������") = 1 Then Exit Sub
    Load frm_Report_DelOrder
    frm_Report_DelOrder.Visible = False: frm_Report_DelOrder.WindowState = 2
    frm_Report_DelOrder.Visible = True
    frm_Report_DelOrder.ZOrder
    frm_Report_DelOrder.Tag = "�q��R������"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_DivideSku_Click()
    '���� �� ������f��
    If CheckOpenForm("������f��") = 1 Then Exit Sub
    Load frm_Report_DivideSku
    frm_Report_DivideSku.Visible = False: frm_Report_DivideSku.WindowState = 2
    frm_Report_DivideSku.Visible = True
    frm_Report_DivideSku.ZOrder
    frm_Report_DivideSku.Tag = "������f��"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_MBO_Cod_Click()
    '���� �� POD�^��
    If CheckOpenForm("���_�U�f���{�d��") = 1 Then Exit Sub
    Load frm_Report_MBO_Cod
    frm_Report_MBO_Cod.Visible = False: frm_Report_MBO_Cod.WindowState = 2
    frm_Report_MBO_Cod.Visible = True
    frm_Report_MBO_Cod.ZOrder
    frm_Report_MBO_Cod.Tag = "���_�U�f���{�d��"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_MboReport_PodRetrun_Click()
    '���� �� POD�^��
    If CheckOpenForm("POD�^��") = 1 Then Exit Sub
    Load frm_Report_PodRetrun
    frm_Report_PodRetrun.Visible = False: frm_Report_PodRetrun.WindowState = 2
    frm_Report_PodRetrun.Visible = True
    frm_Report_PodRetrun.ZOrder
    frm_Report_PodRetrun.Tag = "POD�^��"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_MboReport_SDNReturnList_Click()
    '���� �� POD�^��
    If CheckOpenForm("���_�^���ˮ֪�") = 1 Then Exit Sub
    Load frm_Report_MBO_SDNReturnList
    frm_Report_MBO_SDNReturnList.Visible = False: frm_Report_MBO_SDNReturnList.WindowState = 2
    frm_Report_MBO_SDNReturnList.Visible = True
    frm_Report_MBO_SDNReturnList.ZOrder
    frm_Report_MBO_SDNReturnList.Tag = "���_�^���ˮ֪�"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_SDNReturnList_Click()
    '���� �� �^���ˮ֪�
    If CheckOpenForm("�^���ˮ֪�") = 1 Then Exit Sub
    Load frm_Report_SDNReturnList
    frm_Report_SDNReturnList.Visible = False: frm_Report_SDNReturnList.WindowState = 2
    frm_Report_SDNReturnList.Visible = True
    frm_Report_SDNReturnList.ZOrder
    frm_Report_SDNReturnList.Tag = "�e�f�^���ˮ֪�"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_Ship2TKK_Click()
    'TK�ݨD���� �� �X�f��Ʀ^��
    If CheckOpenForm("�X�f��Ʀ^��") = 1 Then Exit Sub
    Load frm_Report_Ship2TKK
    frm_Report_Ship2TKK.Visible = False: frm_Report_Ship2TKK.WindowState = 2
    frm_Report_Ship2TKK.Visible = True
    frm_Report_Ship2TKK.ZOrder
    frm_Report_Ship2TKK.Tag = "�X�f��Ʀ^��"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_Report_TKARList_Click()
'TK�ݨD���� �� �Ȥ��������Ӫ�
    If CheckOpenForm("�Ȥ��������Ӫ�") = 1 Then Exit Sub
    Load frm_Report_TKARList
    frm_Report_TKARList.Visible = False: frm_Report_TKARList.WindowState = 2
    frm_Report_TKARList.Visible = True
    frm_Report_TKARList.ZOrder
    frm_Report_TKARList.Tag = "�Ȥ��������Ӫ�"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TKCustomerCodeDate_Click()
    'TK�ݨD���� �� �Ȥ�i�f���Ĵ������Ӫ�
    If CheckOpenForm("�Ȥ�i�f���Ĵ������Ӫ�") = 1 Then Exit Sub
    Load frm_Report_TKCustomerCodeDate
    frm_Report_TKCustomerCodeDate.Visible = False: frm_Report_TKCustomerCodeDate.WindowState = 2
    frm_Report_TKCustomerCodeDate.Visible = True
    frm_Report_TKCustomerCodeDate.ZOrder
    frm_Report_TKCustomerCodeDate.Tag = "�Ȥ�i�f���Ĵ������Ӫ�"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TKExpect1_Click()
    'TK�ݨD���� �� �q��t�e���`��
    If CheckOpenForm("�q��t�e���`��") = 1 Then Exit Sub
    Load frm_Report_TKExpect1
    frm_Report_TKExpect1.Visible = False: frm_Report_TKExpect1.WindowState = 2
    frm_Report_TKExpect1.Visible = True
    frm_Report_TKExpect1.ZOrder
    frm_Report_TKExpect1.Tag = "�q��t�e���`��"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TKKPI_Click()
    'TK�ݨD���� �� ��q����
    If CheckOpenForm("��q����") = 1 Then Exit Sub
    Load frm_Report_TKKPI
    frm_Report_TKKPI.Visible = False: frm_Report_TKKPI.WindowState = 2
    frm_Report_TKKPI.Visible = True
    frm_Report_TKKPI.ZOrder
    frm_Report_TKKPI.Tag = "��q����"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TKKRSDNReturnList_Click()
    'TK�ݨD���� �� �h�f�^���ˮ֪�
    If CheckOpenForm("�h�f�^���ˮ֪�") = 1 Then Exit Sub
    Load frm_Report_TKRSDNReturnList
    frm_Report_TKRSDNReturnList.Visible = False: frm_Report_TKRSDNReturnList.WindowState = 2
    frm_Report_TKRSDNReturnList.Visible = True
    frm_Report_TKRSDNReturnList.ZOrder
    frm_Report_TKRSDNReturnList.Tag = "�h�f�^���ˮ֪�"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TKKSDNReturnList_Click()
    'TK�ݨD���� �� �e�f�^���ˮ֪�
    If CheckOpenForm("�e�f�^���ˮ֪�") = 1 Then Exit Sub
    Load frm_Report_TKSDNReturnList
    frm_Report_TKSDNReturnList.Visible = False: frm_Report_TKSDNReturnList.WindowState = 2
    frm_Report_TKSDNReturnList.Visible = True
    frm_Report_TKSDNReturnList.ZOrder
    frm_Report_TKSDNReturnList.Tag = "�e�f�^���ˮ֪�"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TMSAbnormal_Click()
    
    '���� �� �q��t�e���`���
    If CheckOpenForm("�t�e���`��") = 1 Then Exit Sub
    Load frm_Report_TMSAbnormal
    frm_Report_TMSAbnormal.Visible = False: frm_Report_TMSAbnormal.WindowState = 2
    frm_Report_TMSAbnormal.Visible = True
    frm_Report_TMSAbnormal.ZOrder
    frm_Report_TMSAbnormal.Tag = "�t�e���`��"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_APPSdnDetail_Click()
    
    '���� �� �q��t�e���`���
    If CheckOpenForm("ñ����Ӫ�") = 1 Then Exit Sub
    Load frm_Report_APPSdnDetail
    frm_Report_APPSdnDetail.Visible = False: frm_Report_APPSdnDetail.WindowState = 2
    frm_Report_APPSdnDetail.Visible = True
    frm_Report_APPSdnDetail.ZOrder
    frm_Report_APPSdnDetail.Tag = "ñ����Ӫ�"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_Report_TRPTrack_Click()
    '�������f�l�ܪ�
    If CheckOpenForm("��f�l�ܪ�") = 1 Then Exit Sub
    
    Load frm_Report_TRPTrack
    frm_Report_TRPTrack.Visible = False: frm_Report_TRPTrack.WindowState = 2
    frm_Report_TRPTrack.Visible = True
    frm_Report_TRPTrack.ZOrder
    frm_Report_TRPTrack.Tag = "��f�l�ܪ�"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_SwitchDB_Click()
    '��Ʈw����
    If CheckOpenForm("��Ʈw����") = 1 Then Exit Sub
    Load frm_SwitchDB
    frm_SwitchDB.Visible = False: frm_SwitchDB.WindowState = vbNormal
    frm_SwitchDB.Visible = True
    frm_SwitchDB.ZOrder
    frm_SwitchDB.Tag = "��Ʈw����"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_SystemUpdate_Click()
'�t�Χ�s
If CheckOpenForm("�t�Χ�s") = 1 Then Exit Sub
   Load frm_SystemUpdate
   frm_SystemUpdate.Visible = False: frm_SystemUpdate.WindowState = vbNormal
   frm_SystemUpdate.Visible = True
   frm_SystemUpdate.ZOrder
   frm_SystemUpdate.Tag = "�t�Χ�s"
   Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_Report_TKExpect_Click()
    'TK�ݨD���� �� �q��h�^����
    If CheckOpenForm("�q��h�^����") = 1 Then Exit Sub
    Load frm_Report_TKExpect
    frm_Report_TKExpect.Visible = False: frm_Report_TKExpect.WindowState = 2
    frm_Report_TKExpect.Visible = True
    frm_Report_TKExpect.ZOrder
    frm_Report_TKExpect.Tag = "�q��h�^����"
    Call UpdateMDIForm_Menu_WindowName
End Sub


Private Sub Menu_THLReport_Click()
    '���� �� THL�ݨD����
    If CheckOpenForm("THL�ݨD����") = 1 Then Exit Sub
    Dim obj As Object
    Set obj = frm_Report_THL
    Load obj
    obj.Visible = False
    obj.Visible = True
    obj.ZOrder
    obj.Tag = "VLT�ݨD����"
    obj.WindowState = 2
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_TRPPlan_BacktoEXE_Click()
    '�ƨ��B�z�@�~ �� �q��ƨ���Ʀ^�ǳ]�w
    If CheckOpenForm("�q��ƨ���Ʀ^�ǳ]�w") = 1 Then Exit Sub
    Load frm_OP_BacktoEXE
    frm_OP_BacktoEXE.Visible = False: frm_OP_BacktoEXE.WindowState = 2
    frm_OP_BacktoEXE.Visible = True
    frm_OP_BacktoEXE.ZOrder
    frm_OP_BacktoEXE.Tag = "�q��ƨ���Ʀ^�ǳ]�w"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_Cost_Click()
    '�ƨ��B�z�@�~ �� �B�O���R
    If CheckOpenForm("�B�O���R") = 1 Then Exit Sub
    Load frm_OP_TRPCost
    frm_OP_TRPCost.Visible = False: frm_OP_TRPCost.WindowState = 2
    frm_OP_TRPCost.Visible = True
    frm_OP_TRPCost.ZOrder
    frm_OP_TRPCost.Tag = "�B�O���R"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_TRPPlan_CutOrders_Click()
    '�ƨ��B�z�@�~ �� ����h���q�����
    If CheckOpenForm("����h���q�����") = 1 Then Exit Sub
    Load frm_OP_CutOrders
    frm_OP_CutOrders.Visible = False: frm_OP_CutOrders.WindowState = 2
    frm_OP_CutOrders.Visible = True
    frm_OP_CutOrders.ZOrder
    frm_OP_CutOrders.Tag = "����h���q�����"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_DCRouteMerge_Click()
    '�ƨ��B�z�@�~ ��  �G���ƨ��@�~
    If CheckOpenForm("�G���ƨ��@�~") = 1 Then Exit Sub
    Load frm_OP_DCRouteMerge
    frm_OP_DCRouteMerge.Visible = False: frm_OP_DCRouteMerge.WindowState = 2
    frm_OP_DCRouteMerge.Visible = True
    frm_OP_DCRouteMerge.ZOrder
    frm_OP_DCRouteMerge.Tag = "�G���ƨ��@�~"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_ManualOrders_Click()
    '�ƨ��B�z�@�~ �� �q����@�@�~
    If CheckOpenForm("�q����@�@�~") = 1 Then Exit Sub
    Load frm_OP_ManualOrders
    frm_OP_ManualOrders.Visible = False: frm_OP_ManualOrders.WindowState = 2 '  = vbNormal
    frm_OP_ManualOrders.Visible = True
    frm_OP_ManualOrders.ZOrder
    frm_OP_ManualOrders.Tag = "�q����@�@�~"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_OrderImport_Click()
    '�ƨ��B�z�@�~ �� �q����J�ΫȤᲧ�ʺ��@
    If CheckOpenForm("�q����J�ΫȤᲧ�ʺ��@") = 1 Then Exit Sub
    Load frm_OP_OrderImport
    frm_OP_OrderImport.Visible = False: frm_OP_OrderImport.WindowState = 2 '  = vbNormal
    frm_OP_OrderImport.Visible = True
    frm_OP_OrderImport.ZOrder
    frm_OP_OrderImport.Tag = "�q����J�ΫȤᲧ�ʺ��@"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_Query_Click()
    '�ƨ��B�z�@�~ �� �q��d�ߧ@�~
    If CheckOpenForm("�q��d�ߧ@�~") = 1 Then Exit Sub
    Load frm_Query_Orders
    frm_Query_Orders.Visible = False: frm_Query_Orders.WindowState = 2 '  = vbNormal
    frm_Query_Orders.Visible = True
    frm_Query_Orders.ZOrder
    frm_Query_Orders.Tag = "�q��d�ߧ@�~"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_Report_Click()
    '�ƨ��B�z�@�~ �� �ƨ��@�~����
    If CheckOpenForm("�ƨ��@�~����") = 1 Then Exit Sub
    Load frm_Report_TRPPlan
    frm_Report_TRPPlan.Visible = False: frm_Report_TRPPlan.WindowState = 2
    frm_Report_TRPPlan.Visible = True
    frm_Report_TRPPlan.ZOrder
    frm_Report_TRPPlan.Tag = "�ƨ��@�~����"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_ReDelivery_Click()
    '�ƨ��B�z�@�~ �� �����q��b�t�e�@�~
    If CheckOpenForm("�����q��b�t�e�@�~") = 1 Then Exit Sub
    Load frm_OP_ReDelivery
    frm_OP_ReDelivery.Visible = False: frm_OP_ReDelivery.WindowState = 2
    frm_OP_ReDelivery.Visible = True
    frm_OP_ReDelivery.ZOrder
    frm_OP_ReDelivery.Tag = "�����q��b�t�e�@�~"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_Route_Click()
    '�ƨ��B�z�@�~ �� ���u�s�����@�@�~
    If CheckOpenForm("���u�s�����@�@�~") = 1 Then Exit Sub
    Load frm_OP_RouteData
    frm_OP_RouteData.Visible = False: frm_OP_RouteData.WindowState = 2 '  = vbNormal
    frm_OP_RouteData.Visible = True
    frm_OP_RouteData.ZOrder
    frm_OP_RouteData.Tag = "���u�s�����@�@�~"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_RouteConfirm_Click()
    '�ƨ��B�z�@�~ ��  �X���T�{
    If CheckOpenForm("�X���T�{") = 1 Then Exit Sub
    Load frm_OP_RouteConfirm
    frm_OP_RouteConfirm.Visible = False: frm_OP_RouteConfirm.WindowState = 2
    frm_OP_RouteConfirm.Visible = True
    frm_OP_RouteConfirm.ZOrder
    frm_OP_RouteConfirm.Tag = "�X���T�{"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_SDNAbnormal_Click()
    '�ƨ��B�z�@�~ ��  �t�e���`���@
    If CheckOpenForm("�t�e���`���@") = 1 Then Exit Sub
    Load frm_OP_SDNAbnormal
    frm_OP_SDNAbnormal.Visible = False: frm_OP_SDNAbnormal.WindowState = 2
    frm_OP_SDNAbnormal.Visible = True
    frm_OP_SDNAbnormal.ZOrder
    frm_OP_SDNAbnormal.Tag = "�t�e���`���@"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_TRPPlan_SDNConfirm_Click()
    '�ƨ��B�z�@�~ ��  ñ��T�{
    If CheckOpenForm("ñ��T�{") = 1 Then Exit Sub
    Load frm_OP_SDNConfirm
    frm_OP_SDNConfirm.Visible = False: frm_OP_SDNConfirm.WindowState = 2
    frm_OP_SDNConfirm.Visible = True
    frm_OP_SDNConfirm.ZOrder
    frm_OP_SDNConfirm.Tag = "ñ��T�{"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_TRPPlan_ShipQty_Click()
    '�ƨ��B�z�@�~ ��  �z�f�ƶq�T�{
    If CheckOpenForm("�z�f�ƶq�T�{") = 1 Then Exit Sub
    Load frm_OP_ShipQty
    frm_OP_ShipQty.Visible = False: frm_OP_ShipQty.WindowState = 2
    frm_OP_ShipQty.Visible = True
    frm_OP_ShipQty.ZOrder
    frm_OP_ShipQty.Tag = "�z�f�ƶq�T�{"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_TRPPlan_TRPPlan_Click()
    '�ƨ��B�z�@�~ ��  �ƨ��@�~
    If CheckOpenForm("�ƨ��@�~") = 1 Then Exit Sub
    Load frm_OP_TRPPlan
    frm_OP_TRPPlan.Visible = False: frm_OP_TRPPlan.WindowState = 2
    frm_OP_TRPPlan.Visible = True
    frm_OP_TRPPlan.ZOrder
    frm_OP_TRPPlan.Tag = "�ƨ��@�~"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_Upload_FTP_Click()
    '�W�U�ǧ@�~ �� FTP �W�U��
    If CheckOpenForm("FTP�W�U��") = 1 Then Exit Sub
    Load frm_FTP
    frm_FTP.Visible = False
    frm_FTP.WindowState = 2 '  = vbNormal
    frm_FTP.Visible = True
    frm_FTP.ZOrder
    frm_FTP.Tag = "FTP�W�U��"
    Call UpdateMDIForm_Menu_WindowName

End Sub

Private Sub Menu_Query_KPI_KPI_Click()
    '�޲zKPI �� �C��KPI
    If CheckOpenForm("�C��KPI") = 1 Then Exit Sub
    Load frm_Query_KPI
    frm_Query_KPI.Visible = False
    frm_Query_KPI.WindowState = 2 '  = vbNormal
    frm_Query_KPI.Visible = True
    frm_Query_KPI.ZOrder
    frm_Query_KPI.Tag = "�C��KPI"
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_VTLReport_Click()

    '���� �� VLT�ݨD����
    If CheckOpenForm("VLT�ݨD����") = 1 Then Exit Sub
    Dim obj As Object
    Set obj = frm_Report_VTL
    Load obj
    obj.Visible = False
    obj.Visible = True
    obj.ZOrder
    obj.Tag = "VLT�ݨD����"
    obj.WindowState = 2
    Call UpdateMDIForm_Menu_WindowName
End Sub

Private Sub Menu_WindowMinx_Click()
'Menu �� [����]��[�̤p��]
Dim i As Integer
For i = 0 To Forms.Count - 1
    If Not (Forms(i) Is frm_MDIForm) Then
       Forms(i).WindowState = vbMinimized
    End If
Next i
End Sub

Private Sub Menu_WindowSourceSizex_Click()
'Menu �� [����]��[��l����]
Dim i As Integer
Dim frmHeight As Long, frmWidth As Long, frmTopNum As Long
For i = 0 To Forms.Count - 1
    If Not (Forms(i) Is frm_MDIForm) Then
       Forms(i).WindowState = 2
    End If
Next i
End Sub

Private Sub Disable_Menu()
'�w�]�ʧ@�GDisable �Ҧ��\���
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
'��ini�ɩw�q���ÿ��
'Create by Gemini @20070416
'
'
'
'*****************************
'���Ѽ�
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
