VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm_BaseData_DataList 
   Caption         =   "��ƭȿ��....."
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
      Caption         =   "��Ƹ��J�i��"
      Height          =   1380
      Left            =   1260
      TabIndex        =   11
      Top             =   1170
      Visible         =   0   'False
      Width           =   6540
      Begin VB.TextBox txt_DataLoading 
         Appearance      =   0  '����
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '�S���ؽu
         BeginProperty Font 
            Name            =   "�s�ө���"
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
      Caption         =   "��Ƹ��J"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      Style           =   1  '�Ϥ��~�[
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
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   9
      Top             =   4080
      Width           =   345
   End
   Begin VB.TextBox txt_Query 
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   6
      Top             =   4080
      Width           =   1410
   End
   Begin VB.CommandButton cmd_OrderBy 
      Height          =   345
      Left            =   4410
      Picture         =   "frm_BaseData_DataList.frx":058A
      Style           =   1  '�Ϥ��~�[
      TabIndex        =   5
      Top             =   4050
      Width           =   360
   End
   Begin VB.ComboBox cmb_OrderBy 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   4
      Top             =   4080
      Width           =   1275
   End
   Begin VB.ComboBox cmb_OrderBy 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   3
      Top             =   4080
      Width           =   1275
   End
   Begin VB.ComboBox cmb_OrderBy 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      Style           =   2  '��¤U�Ԧ�
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
      BackStyle       =   0  '�z��
      Caption         =   "�M��"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
      BackStyle       =   0  '�z��
      Caption         =   "�Ƨ�"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
Private arFieldName() As String      '���W��
Private dLoopVar1 As Double          '�j���ܼ�
Private dLoopVar2 As Double          '�j���ܼ�

Private blEventFlag As Boolean       '�ƥ�O�_����

Private Sub cmb_Query_Click()
'��ƴM�������w
If blEventFlag Then
   txt_Query.SelStart = 0: txt_Query.SelLength = Len(txt_Query.Text)
   txt_Query.SetFocus
End If
End Sub

Private Sub cmd_LoadData_Click()
'���J���
Select Case UCase(strDataList_Caller)
    Case "FRM_OP_MANUALORDERS_CMD_CONSIGNEELIST"
         '�Ȥ��ƶq�L�j�A�ѨϥΪ̦ۦ�M�w�O�_���J
         blEventFlag = False
         Screen.MousePointer = vbHourglass
         Call frm_OP_ManualOrders_cmd_ConsigneeList
         blEventFlag = True
         Screen.MousePointer = vbDefault
    Case "FRM_OP_MANUALORDERS_CMDSHIPTOLIST"
         '�Ȥ��ƶq�L�j�A�ѨϥΪ̦ۦ�M�w�O�_���J
         blEventFlag = False
         Screen.MousePointer = vbHourglass
         Call frm_OP_ManualOrders_cmdShipToList
         blEventFlag = True
         Screen.MousePointer = vbDefault
End Select

End Sub

Private Sub cmd_OrderBy_Click()
'�Ƨ�
dg_DataList.Visible = False
dg_DataList.Sort = 9   '�ۭq
dg_DataList.Visible = True
End Sub

Private Sub cmd_Query_Click()
'�M��
If cmb_Query.Text = "" Then Exit Sub
If Len(Trim(txt_Query.Text)) = 0 Then Exit Sub

'�̩Ҭd�����Ƨ�
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
         '�r�����A��ƴM��
         If Fun_ChkNumber(Trim(.Text)) = 1 Then
            If InStr(.Text, txt_Query.Text) > 0 Then
               .Visible = True
               .TopRow = dLoopVar1
               .LeftCol = cmb_Query.ListIndex
               Call dg_DataList_Click
               Exit Sub
            End If
          Else
          '�Ʀr���A��ƴM��
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
     msg_text = "���A�^���G�䤣��ŦX���󤧸��"
     MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End With

End Sub

Private Sub dg_DataList_Click()
'�I�@���G����A�I�ĤG���G�������
Dim i As Double
With dg_DataList
     .Col = 0   '�s��
     If Len(Trim(.Text)) = 0 Then Exit Sub
     .Col = 0
     For i = 0 To .Cols - 1
         .ColSel = i
     Next i
End With
End Sub

Private Sub dg_DataList_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
'�ۭq�Ƨ�
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
'DoubleClick >> �����ƨñN��ƶǦ^�ܩI�s��
With dg_DataList
     .Col = 0   '�s��
     If Len(Trim(.Text)) = 0 Then Exit Sub
     Call ReturnToCaller
End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'�d�I��Ӫ����L����ƥ�
'�γ~�G�ϥΪ̫��U Esc �h���Ǧ^�����ơA�B��������
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
'�]�w Form �j�p�B��m
Me.Height = 5000: Me.Width = 8900
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 300

blEventFlag = False
Select Case UCase(strDataList_Caller)
    Case "FRM_OP_TRPPLAN_CMD_TAB0_SELECTCAR1", "FRM_OP_TRPPLAN_CMD_TAB0_SELECTCAR2"
         '�ƨ��B�z�@�~ >> �ƨ��@�~ >> �q����ƿ��
         'Form_Name�Gfrm_OP_TRPPlan
         msg_title = "�B�e���������..."
         Me.Caption = "�п������B�e������......."
         Call frm_OP_TRPPlan_cmd_Tab0_SelectCar
    Case "FRM_OP_DCROUTEMERGE_CMD_TAB0_SELECTCAR1", "FRM_OP_DCROUTEMERGE_CMD_TAB0_SELECTCAR2"
         '�ƨ��B�z�@�~ >> DC �֨��@�~ >> �q����ƿ��
         'Form_Name�Gfrm_OP_DCRouteMerge
         msg_title = "�B�e���������..."
         Me.Caption = "�п������B�e������......."
         Call frm_OP_DCRouteMerge_cmd_Tab0_SelectCar
    Case "FRM_OP_MANUALORDERS_CMDSHIPTOLIST"
         '�q����@�@�~ >> ��B��f�Ȥ��ƿ��
         'Form_Name�Gfrm_OP_ManualOrders
         msg_title = "�q�椧�Ȥ���..."
         Me.Caption = "�п���q�椧�Ȥ�......."
         '�Ȥ��ƶq�L�j�A���J�O�ɡA�ѨϥΪ̨M�w�O�_���J
         cmb_OrderBy(2).Visible = False
         cmd_OrderBy.Left = cmd_OrderBy.Left - cmb_OrderBy(2).Width
         cmd_LoadData.Visible = True
    Case "FRM_OP_MANUALORDERS_CMD_CONSIGNEELIST"
         '�q����@�@�~ >> �Ȥ�q����ƿ��
         'Form_Name�Gfrm_OP_ManualOrders
         msg_title = "�q�椧�Ȥ���..."
         Me.Caption = "�п���q�椧�Ȥ�......."
         '�Ȥ��ƶq�L�j�A���J�O�ɡA�ѨϥΪ̨M�w�O�_���J
         cmb_OrderBy(2).Visible = False
         cmd_OrderBy.Left = cmd_OrderBy.Left - cmb_OrderBy(2).Width
         cmd_LoadData.Visible = True
    Case "FRM_OP_ROUTEDATA_CMD_SELECTCAR"
         '�ƨ��B�z�@�~ >> ���u�s�����@�@�~ >> �q����ƿ��
         'Form_Name�Gfrm_OP_TRPPlan
         msg_title = "�B�e���������..."
         Me.Caption = "�п������B�e������......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB0_SELECTCAR02" '�@�w�n�j�g
         '�ƨ��B�z�@�~ >> �X���T�{ >> �q����ƿ��
         'Form_Name�Gfrm_OP_RouteConfirm
         msg_title = "�B�e���������..."
         Me.Caption = "�п������B�e������......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB0_SELECTCAR12" '�@�w�n�j�g
         '�ƨ��B�z�@�~ >> �X���T�{ >> �q����ƿ��
         'Form_Name�Gfrm_OP_RouteConfirm
         msg_title = "�B�e���������..."
         Me.Caption = "�п������B�e������......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB1_SELECTCAR2" '�@�w�n�j�g
         '�ƨ��B�z�@�~ >> �X���T�{ >> �q����ƿ��
         'Form_Name�Gfrm_OP_RouteConfirm
         msg_title = "�B�e���������..."
         Me.Caption = "�п������B�e������......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB2_SELECTCAR2" '�@�w�n�j�g
         '�ƨ��B�z�@�~ >> �X���T�{ >> �q����ƿ��
         'Form_Name�Gfrm_OP_RouteConfirm
         msg_title = "�B�e���������..."
         Me.Caption = "�п������B�e������......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
    Case "FRM_OP_SDNCONFIRM_CMD_TAB2_SELECTCAR2" '�@�w�n�j�g
         '�ƨ��B�z�@�~ >> �X���T�{ >> �q����ƿ��
         'Form_Name�Gfrm_OP_RouteConfirm
         msg_title = "�B�e���������..."
         Me.Caption = "�п������B�e������......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
         'frm_OP_TRPPlan
    Case "FRM_OTHER_OPTPLAN_CMD_TAB0_SELECTCAR2" '�@�w�n�j�g
         '�ƨ��B�z�@�~ >> �X���T�{ >> �q����ƿ��
         'Form_Name�Gfrm_OP_RouteConfirm
         msg_title = "�B�e���������..."
         Me.Caption = "�п������B�e������......."
         Call frm_OP_ROUTEDATA_cmd_SelectCar
    Case Else
         msg_text = "�ǤJ��ƿ��~�G���i���I�s��"
         MsgBox msg_text, vbOKOnly + vbInformation, msg_title
         Unload Me
End Select
blEventFlag = True

End Sub

Private Sub txt_Query_KeyPress(KeyAscii As Integer)
'�M�����
If KeyAscii = vbKeyReturn Then
   cmd_Query.SetFocus
End If
End Sub


Private Sub frm_OP_TRPPlan_cmd_Tab0_SelectCar()
'�ƨ��B�z�@�~ >> �ƨ��@�~ >> �q����ƿ��
'Form_Name�Gfrm_OP_TRPPlan

'�]�w DataGrid �榡
Dim sub_var1 As Integer, sub_var2 As Integer
dg_DataList.Visible = False
With dg_DataList
     .FixedRows = 1: .Cols = 11
     '�]�w���\��C���
     .AllowBigSelection = True
     '�]�w�C����r�r��
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "�s�ө���": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '�]�w�C�����e��
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
     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "No."
     .Col = 1: .Text = "����"
     .Col = 2: .Text = "���q"
     .Col = 3: .Text = "���P���X"
     .Col = 4: .Text = "�i����"
     .Col = 5: .Text = "�r�p�H"
     .Col = 6: .Text = "�q��"
     .Col = 7: .Text = "����"
     .Col = 8: .Text = "����"
     .Col = 9: .Text = "���ػ���"
     .Col = 10: .Text = "���ػ���"
     '�]�w�C����r���
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

'���o�B�e�����򥻸��
str_SQL = "Select ����,���q�O,���P���X,�i����,�r�p�H,�q��,����,���إN�X,���q�N�X From BaseData_TRPCarList Order by ����"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
cn.CommandTimeout = 0   '�L��������
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧹B�騮�����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End If

Do While Not tmp_rs.EOF
   With dg_DataList
       .Rows = .Rows + 1
       .Row = .Rows - 2
       .Col = 0    '�Ǹ�
       .Text = .Rows - 2
       .Col = 1    '���إN�X
       .Text = tmp_rs.Fields("���إN�X").Value
       .Col = 2    '�B�餽�q
       .Text = tmp_rs.Fields("���q�N�X").Value
       .Col = 3    '���P���X
       .Text = tmp_rs.Fields("���P���X").Value
       .Col = 4    '�i����
       .Text = tmp_rs.Fields("�i����").Value
       .Col = 5    '�r�p�H
       .Text = tmp_rs.Fields("�r�p�H").Value
       .Col = 6    '�q��
       .Text = tmp_rs.Fields("�q��").Value
       .Col = 7    '����
       .Text = tmp_rs.Fields("����").Value
       .Col = 8    '����
       .Text = "�H"
       .Col = 9    '����
       .Text = tmp_rs.Fields("����").Value
       .Col = 10    '����
       .Text = tmp_rs.Fields("���q�O").Value
  End With
  tmp_rs.MoveNext
Loop
tmp_rs.Close

If UCase(strDataList_Caller) = "FRM_OP_TRPPLAN_CMD_TAB0_SELECTCAR1" Then
   '�d�ߦU�������B�e�w�Ʃw�������s��
   With dg_DataList
     For dLoopVar1 = 1 To .Rows - 2
        .Row = dLoopVar1
        .Col = 3   '���P���X
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


'���W�ٴ����ѨϥΪ̶i��ۭq�ƧǡB�M��
ReDim arFieldName(1) As String
dLoopVar2 = 0
dg_DataList.Row = 0
For dLoopVar1 = 0 To dg_DataList.Cols - 1
   dLoopVar2 = dLoopVar2 + 1           '�}�C�H 0 �}�l�A�]���̫᣸�ӷ|�O�ťաA�� [���]�w�Ƨ�]
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
'�ۭq�ƧǡB�M��G�w�]����̫᣸�ӡG�ť�
For dLoopVar1 = 0 To cmb_OrderBy.Count - 1
    cmb_OrderBy(dLoopVar1).ListIndex = cmb_OrderBy(dLoopVar1).ListCount - 1
Next dLoopVar1
cmb_Query.ListIndex = cmb_Query.ListCount - 1

End Sub

Private Sub ReturnToCaller()
'�ƨ��B�z�@�~ >> �ƨ��@�~ >> �q����ƿ��
'Form_Name�Gfrm_OP_TRPPlan
Select Case UCase(strDataList_Caller)
    Case "FRM_OP_TRPPLAN_CMD_TAB0_SELECTCAR1", "FRM_OP_TRPPLAN_CMD_TAB0_SELECTCAR2"
         '�ƨ��B�z�@�~ >> �ƨ��@�~ >> �q����ƿ��
         'Form_Name�Gfrm_OP_TRPPlan
         With dg_DataList
              .Col = 3     '���P���X
              frm_OP_TRPPlan.txt_Tab0_DeliveryCarNo.Text = Trim(.Text)
              .Col = 2     '�B�餽�q
              frm_OP_TRPPlan.txt_Tab0_DeliveryCompany.Text = Trim(.Text)
              .Col = 5     '�r�p�H
              frm_OP_TRPPlan.txt_Tab0_DeliveryDriver.Text = Trim(.Text)
              .Col = 6     '�q��
              frm_OP_TRPPlan.txt_Tab0_DeliveryPhone.Text = Trim(.Text)
              .Col = 1     '����
              frm_OP_TRPPlan.txt_Tab0_DeliveryCarType.Text = Trim(.Text)
         End With
         frm_OP_TRPPlan.WindowState = 2   '�̤j��
    Case "FRM_OP_DCROUTEMERGE_CMD_TAB0_SELECTCAR1", "FRM_OP_DCROUTEMERGE_CMD_TAB0_SELECTCAR2"
         '�ƨ��B�z�@�~ >> DC�֨��@�~ >> �q����ƿ��
         'Form_Name�Gfrm_OP_DCRouteMerge
         With dg_DataList
              .Col = 3     '���P���X
              frm_OP_DCRouteMerge.txt_Tab0_DeliveryCarNo.Text = Trim(.Text)
              .Col = 2     '�B�餽�q
              frm_OP_DCRouteMerge.txt_Tab0_DeliveryCompany.Text = Trim(.Text)
              .Col = 5     '�r�p�H
              frm_OP_DCRouteMerge.txt_Tab0_DeliveryDriver.Text = Trim(.Text)
              .Col = 6     '�q��
              frm_OP_DCRouteMerge.txt_Tab0_DeliveryPhone.Text = Trim(.Text)
              .Col = 1     '����
              frm_OP_DCRouteMerge.txt_Tab0_DeliveryCarType.Text = Trim(.Text)
              .Col = 1     '���إN�X
              frm_OP_DCRouteMerge.txt_Tab0_DeliveryCarTypeCode.Text = Trim(.Text)
         End With
         frm_OP_DCRouteMerge.WindowState = 2
    Case "FRM_OTHER_OPTPLAN_CMD_TAB0_SELECTCAR2"
         '�h�f�ƨ� >> �h�f�ƨ� >> �q����ƿ��
         'Form_Name�Gfrm_OP_TRPPlan
         With dg_DataList
              .Col = 3     '���P���X
              frm_Other_OPTPlan.txt_Tab0_DeliveryCarNo.Text = Trim(.Text)
              .Col = 2     '�B�餽�q
              frm_Other_OPTPlan.txt_Tab0_DeliveryCompany.Text = Trim(.Text)
              .Col = 5     '�r�p�H
              frm_Other_OPTPlan.txt_Tab0_DeliveryDriver.Text = Trim(.Text)
              .Col = 6     '�q��
              frm_Other_OPTPlan.txt_Tab0_DeliveryPhone.Text = Trim(.Text)
              .Col = 1     '����
              frm_Other_OPTPlan.txt_Tab0_DeliveryCarType.Text = Trim(.Text)
         End With
         frm_Other_OPTPlan.WindowState = 2   '�̤j��
    Case "FRM_OP_MANUALORDERS_CMD_CONSIGNEELIST"
         '�q����@�@�~ >> �Ȥ��ƿ��
         'Form_Name�Gfrm_OP_ManualOrders
          With dg_DataList
               .Col = 1    '�Ȥ�s��
               frm_OP_ManualOrders.txt_ConsigneeKey.Text = .Text
               Call frm_OP_ManualOrders.txt_ConsigneeKey_LostFocus
          End With
          frm_OP_ManualOrders.WindowState = 2
    Case "FRM_OP_MANUALORDERS_CMDSHIPTOLIST"
         '�q����@�@�~ >> ��B��f�Ȥ��ƿ��
         'Form_Name�Gfrm_OP_ManualOrders
          With dg_DataList
               .Col = 1    '�Ȥ�s��
               frm_OP_ManualOrders.txtShipToKey.Text = .Text
               Call frm_OP_ManualOrders.txtShipToKey_LostFocus
          End With
          frm_OP_ManualOrders.WindowState = 2
    Case "FRM_OP_ROUTEDATA_CMD_SELECTCAR"
         '�ƨ��B�z�@�~ >> ���u�s�����@�@�~ >> �q����ƿ��
         'Form_Name�Gfrm_OP_RouteData
         With dg_DataList
              .Col = 3     '���P���X
              frm_OP_RouteData.txt_VehicleNo.Text = Trim(.Text)
              .Col = 2     '�B�餽�q
              frm_OP_RouteData.txt_TRPCompany.Text = Trim(.Text)
              .Col = 5     '�r�p�H
              frm_OP_RouteData.txt_Driver.Text = Trim(.Text)
              .Col = 6     '�q��
              frm_OP_RouteData.txt_Phone.Text = Trim(.Text)
              .Col = 1     '����
              frm_OP_RouteData.txt_VehicleType.Text = Trim(.Text)
         End With
         frm_OP_RouteData.WindowState = 2   '�̤j��
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB0_SELECTCAR02" '�@�w�n�j�g"
         '�ƨ��B�z�@�~ >> �X���T�{ >> �q����ƿ��
         'Form_Name�Gfrm_OP_RouteConfirm
         With dg_DataList
              .Col = 3     '���P���X
              frm_OP_RouteConfirm.txt_VehicleNo0.Text = Trim(.Text)
              .Col = 5     '�r�p�H
              frm_OP_RouteConfirm.txt_Driver0.Text = Trim(.Text)
         End With
         frm_OP_RouteConfirm.WindowState = 2   '�̤j��
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB0_SELECTCAR12" '�@�w�n�j�g"
         '�ƨ��B�z�@�~ >> �X���T�{ >> �q����ƿ��
         'Form_Name�Gfrm_OP_RouteConfirm
         With dg_DataList
              .Col = 3     '���P���X
              frm_OP_RouteConfirm.txt_VehicleNo1.Text = Trim(.Text)
              .Col = 5     '�r�p�H
              frm_OP_RouteConfirm.txt_Driver1.Text = Trim(.Text)
         End With
         frm_OP_RouteConfirm.WindowState = 2   '�̤j��
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB1_SELECTCAR2" '�@�w�n�j�g"
         '�ƨ��B�z�@�~ >> �X���T�{ >> �q����ƿ��
         'Form_Name�Gfrm_OP_RouteConfirm
         With dg_DataList
              .Col = 3     '���P���X
              frm_OP_RouteConfirm.txt_Tab1_VehicleNo.Text = Trim(.Text)
              .Col = 5     '�r�p�H
              frm_OP_RouteConfirm.txt_Tab1_Driver0.Text = Trim(.Text)
         End With
         frm_OP_RouteConfirm.WindowState = 2   '�̤j��
    Case "FRM_OP_ROUTECONFIRM_CMD_TAB2_SELECTCAR2" '�@�w�n�j�g"
         '�ƨ��B�z�@�~ >> �X���T�{ >> �q����ƿ��
         'Form_Name�Gfrm_OP_RouteConfirm
         With dg_DataList
              .Col = 3     '���P���X
              frm_OP_RouteConfirm.txt_Tab2_VehicleNo.Text = Trim(.Text)
              .Col = 5     '�r�p�H
              frm_OP_RouteConfirm.txt_Tab2_Driver.Text = Trim(.Text)
         End With
         frm_OP_RouteConfirm.WindowState = 2   '�̤j��
    Case "FRM_OP_SDNCONFIRM_CMD_TAB2_SELECTCAR2" '�@�w�n�j�g"
         '�ƨ��B�z�@�~ >> �X���T�{ >> �q����ƿ��
         'Form_Name�Gfrm_OP_RouteConfirm
         With dg_DataList
              .Col = 3     '���P���X
              frm_OP_SDNConfirm.txt_Tab02_C_VEHICLE_ID_NO.Text = Trim(.Text)
              .Col = 5     '�r�p�H
              frm_OP_SDNConfirm.txt_Tab02_Driver.Text = Trim(.Text)
              frm_OP_SDNConfirm.txt_Tab02_Receiver.Text = Trim(.Text)
              frm_OP_SDNConfirm.NextPositionTab2Detail 1, 2
         End With
         frm_OP_SDNConfirm.WindowState = 2   '�̤j��
    Case Else
         msg_text = "�������I�s�̡A��Ƥ����n�Ǧ^����"
         MsgBox msg_text, vbOKOnly + vbInformation, msg_title
End Select
Unload Me
End Sub

Private Sub frm_OP_DCRouteMerge_cmd_Tab0_SelectCar()
'�ƨ��B�z�@�~ >> DC �֨��@�~ >> �q����ƿ��
'Form_Name�Gfrm_OP_DCRouteMerge

'�]�w DataGrid �榡
Dim sub_var1 As Integer, sub_var2 As Integer
dg_DataList.Visible = False
With dg_DataList
     .FixedRows = 1: .Cols = 10
     '�]�w���\��C���
     .AllowBigSelection = True
     '�]�w�C����r�r��
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "�s�ө���": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '�]�w�C�����e��
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
     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "No."
     .Col = 1: .Text = "����"
     .Col = 2: .Text = "���q"
     .Col = 3: .Text = "���P���X"
     .Col = 4: .Text = "�i����"
     .Col = 5: .Text = "�r�p�H"
     .Col = 6: .Text = "�q��"
     .Col = 7: .Text = "����"
     .Col = 8: .Text = "����"
     .Col = 9: .Text = "���ػ���"
     '�]�w�C����r���
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

'���o�B�e�����򥻸��
str_SQL = "Select ����,���q�O,���P���X,�i����,�r�p�H,�q��,����,���إN�X From BaseData_TRPCarList Order by ����"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
cn.CommandTimeout = 0   '�L��������
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧹B�騮�����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End If

Do While Not tmp_rs.EOF
   With dg_DataList
       .Rows = .Rows + 1
       .Row = .Rows - 2
       .Col = 0    '�Ǹ�
       .Text = .Rows - 2
       .Col = 1    '���إN�X
       .Text = tmp_rs.Fields("���إN�X").Value
       .Col = 2    '�B�餽�q
       .Text = tmp_rs.Fields("���q�O").Value
       .Col = 3    '���P���X
       .Text = tmp_rs.Fields("���P���X").Value
       .Col = 4    '�i����
       .Text = tmp_rs.Fields("�i����").Value
       .Col = 5    '�r�p�H
       .Text = tmp_rs.Fields("�r�p�H").Value
       .Col = 6    '�q��
       .Text = tmp_rs.Fields("�q��").Value
       .Col = 7    '����
       .Text = tmp_rs.Fields("����").Value
       .Col = 8    '����
       .Text = "�H"
       .Col = 9    '���إN�X
       .Text = tmp_rs.Fields("����").Value
  End With
  tmp_rs.MoveNext
Loop
tmp_rs.Close

If UCase(strDataList_Caller) = "FRM_OP_DCROUTEMERGE_CMD_TAB0_SELECTCAR1" Then
   '�d�ߦU�������B�e�w�Ʃw�������s��
   With dg_DataList
     For dLoopVar1 = 1 To .Rows - 2
        .Row = dLoopVar1
        .Col = 3   '���P���X
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


'���W�ٴ����ѨϥΪ̶i��ۭq�ƧǡB�M��
ReDim arFieldName(1) As String
dLoopVar2 = 0
dg_DataList.Row = 0
For dLoopVar1 = 0 To dg_DataList.Cols - 1
   dLoopVar2 = dLoopVar2 + 1           '�}�C�H 0 �}�l�A�]���̫᣸�ӷ|�O�ťաA�� [���]�w�Ƨ�]
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
'�ۭq�ƧǡB�M��G�w�]����̫᣸�ӡG�ť�
For dLoopVar1 = 0 To cmb_OrderBy.Count - 1
    cmb_OrderBy(dLoopVar1).ListIndex = cmb_OrderBy(dLoopVar1).ListCount - 1
Next dLoopVar1
cmb_Query.ListIndex = cmb_Query.ListCount - 1

End Sub

Private Sub frm_OP_ManualOrders_cmd_ConsigneeList()
'�q����@�@�~ >> �Ȥ��ƿ��
'Form_Name�Gfrm_OP_ManualOrders

'�]�w DataGrid �榡
Dim sub_var1 As Integer, sub_var2 As Integer
dg_DataList.Visible = False
With dg_DataList
     .FixedRows = 1: .Cols = 15
     '�]�w���\��C���
     .AllowBigSelection = True
     '�]�w�C����r�r��
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "�s�ө���": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '�]�w�C�����e��
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

     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "No."
     .Col = 1: .Text = "�Ȥ�s��"
     .Col = 2: .Text = "�Ȥ�W��"
     .Col = 3: .Text = "�Ȥ�²��"
     .Col = 4: .Text = "�l���ϸ�"
     .Col = 5: .Text = "�B�e�a�}"
     .Col = 6: .Text = "�B�e�ϰ�"
     .Col = 7: .Text = "�S��ݨD-1"
     .Col = 8: .Text = "�S��ݨD-2"
     .Col = 9: .Text = "�p���H"
     .Col = 10: .Text = "�q��"
     .Col = 11: .Text = "�B�e�ϰ�X"
     .Col = 12: .Text = "�l���ϸ��X"
     .Col = 13: .Text = "�S��ݨD1"
     .Col = 14: .Text = "�S��ݨD2"
     '�]�w�C����r���
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

'���f�D
If Len(RTrim(strStorerkey)) = 0 Then
    str_SQL = "Select count(*) as RecCount From TRP01M"
Else
    str_SQL = "Select count(*) as RecCount From TRP01M where storerkey = '" & strStorerkey & "' "
End If

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
cn.CommandTimeout = 0   '�L��������
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫Ȥ���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
Else
   dbTotal = tmp_rs.Fields("RecCount").Value
   pb_DataLoading.Max = dbTotal
End If
tmp_rs.Close

'�f�D����
If Len(RTrim(strStorerkey)) = 0 Then
    strStorerkey = ""
Else
    strStorerkey = "where �f�D�s�� = '" & strStorerkey & "' "
End If
   
'���o�Ȥ�򥻸��
str_SQL = "Select �Ȥ�s��,�Ȥ�W��,�Ȥ�²��,�l���ϸ�,�B�e�a�},�B�e�ϰ�,�S��ݨD1,�S��ݨD2,�p���H,�q��," & _
          "  �B�e�ϰ�X,�l���ϸ��X,�S��ݨD�X1,�S��ݨD�X2 " & _
          "From BaseData_ConsigneeList " & strStorerkey & " Order by �Ȥ�s��"
cn.CommandTimeout = 0   '�L��������
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫Ȥ���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

dbNow = 0
Do While Not tmp_rs.EOF
   dbNow = dbNow + 1
   pb_DataLoading.Value = dbNow
   txt_DataLoading.Text = "�Ȥ��Ʀ@ " & dbTotal & " �w���J " & dbNow & " ��"
   DoEvents
   With dg_DataList
       .Rows = .Rows + 1
       .Row = .Rows - 2
       .Col = 0    '�Ǹ�
       .Text = .Rows - 2
       .Col = 1    '�Ȥ�s��
       .Text = tmp_rs.Fields("�Ȥ�s��").Value
       .Col = 2    '�Ȥ�W��
       .Text = tmp_rs.Fields("�Ȥ�W��").Value
       .Col = 3    '�Ȥ�²��
       .Text = tmp_rs.Fields("�Ȥ�²��").Value
       .Col = 4    '�l���ϸ�
       .Text = tmp_rs.Fields("�l���ϸ�").Value
       .Col = 5    '�B�e�ϰ�
       .Text = tmp_rs.Fields("�B�e�a�}").Value
       .Col = 6    '�B�e�a�}
       .Text = tmp_rs.Fields("�B�e�ϰ�").Value
       .Col = 7    '�S��ݨD 1
       .Text = tmp_rs.Fields("�S��ݨD1").Value
       .Col = 8    '�S��ݨD 2
       .Text = tmp_rs.Fields("�S��ݨD2").Value
       .Col = 9    '�p���H
       .Text = tmp_rs.Fields("�p���H").Value
       .Col = 10   '�q��
       .Text = tmp_rs.Fields("�q��").Value
       .Col = 11   '�B�e�ϰ�X
       .Text = tmp_rs.Fields("�B�e�ϰ�X").Value
       .Col = 12   '�l���ϸ��N�X
       .Text = tmp_rs.Fields("�l���ϸ��X").Value
       .Col = 13   '�S��ݨD�X1
       .Text = tmp_rs.Fields("�S��ݨD�X1").Value
       .Col = 14   '�S��ݨD�X2
       .Text = tmp_rs.Fields("�S��ݨD�X2").Value
  End With
  tmp_rs.MoveNext
Loop
tmp_rs.Close
fam_DataLoading.Visible = False
dg_DataList.Visible = True

'���W�ٴ����ѨϥΪ̶i��ۭq�ƧǡB�M��
ReDim arFieldName(1) As String
dLoopVar2 = 0
dg_DataList.Row = 0
For dLoopVar1 = 0 To dg_DataList.Cols - 1
   dLoopVar2 = dLoopVar2 + 1           '�}�C�H 0 �}�l�A�]���̫᣸�ӷ|�O�ťաA�� [���]�w�Ƨ�]
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
'�ۭq�ƧǡB�M��G�w�]����̫᣸�ӡG�ť�
For dLoopVar1 = 0 To cmb_OrderBy.Count - 1
    cmb_OrderBy(dLoopVar1).ListIndex = cmb_OrderBy(dLoopVar1).ListCount - 1
Next dLoopVar1
cmb_Query.ListIndex = cmb_Query.ListCount - 1

End Sub
Private Sub frm_OP_ManualOrders_cmdShipToList()
'�q����@�@�~ >> ��B��f�Ȥ��ƿ��
'Form_Name�Gfrm_OP_ManualOrders

'�]�w DataGrid �榡
Dim sub_var1 As Integer, sub_var2 As Integer
dg_DataList.Visible = False
With dg_DataList
     .FixedRows = 1: .Cols = 15
     '�]�w���\��C���
     .AllowBigSelection = True
     '�]�w�C����r�r��
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "�s�ө���": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '�]�w�C�����e��
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

     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "No."
     .Col = 1: .Text = "�Ȥ�s��"
     .Col = 2: .Text = "�Ȥ�W��"
     .Col = 3: .Text = "�Ȥ�²��"
     .Col = 4: .Text = "�l���ϸ�"
     .Col = 5: .Text = "�B�e�a�}"
     .Col = 6: .Text = "�B�e�ϰ�"
     .Col = 7: .Text = "�S��ݨD-1"
     .Col = 8: .Text = "�S��ݨD-2"
     .Col = 9: .Text = "�p���H"
     .Col = 10: .Text = "�q��"
     .Col = 11: .Text = "�B�e�ϰ�X"
     .Col = 12: .Text = "�l���ϸ��X"
     .Col = 13: .Text = "�S��ݨD1"
     .Col = 14: .Text = "�S��ݨD2"
     '�]�w�C����r���
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

'���f�D
If Len(RTrim(strStorerkey)) = 0 Then
    str_SQL = "Select count(*) as RecCount From TRP01M"
Else
    str_SQL = "Select count(*) as RecCount From TRP01M where storerkey = '" & strStorerkey & "' "
End If

Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
cn.CommandTimeout = 0   '�L��������
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫Ȥ���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
Else
   dbTotal = tmp_rs.Fields("RecCount").Value
   pb_DataLoading.Max = dbTotal
End If
tmp_rs.Close
   
'�f�D����
If Len(RTrim(strStorerkey)) = 0 Then
    strStorerkey = ""
Else
    strStorerkey = "where �f�D�s�� = '" & strStorerkey & "' "
End If
   
'���o�Ȥ�򥻸��
str_SQL = "Select �Ȥ�s��,�Ȥ�W��,�Ȥ�²��,�l���ϸ�,�B�e�a�},�B�e�ϰ�,�S��ݨD1,�S��ݨD2,�p���H,�q��," & _
          "  �B�e�ϰ�X,�l���ϸ��X,�S��ݨD�X1,�S��ݨD�X2 " & _
          "From BaseData_ConsigneeList " & strStorerkey & " Order by �Ȥ�s��"
cn.CommandTimeout = 0   '�L��������
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120

If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧫Ȥ���"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
   Exit Sub
End If

dbNow = 0
Do While Not tmp_rs.EOF
   dbNow = dbNow + 1
   pb_DataLoading.Value = dbNow
   txt_DataLoading.Text = "�Ȥ��Ʀ@ " & dbTotal & " �w���J " & dbNow & " ��"
   DoEvents
   With dg_DataList
       .Rows = .Rows + 1
       .Row = .Rows - 2
       .Col = 0    '�Ǹ�
       .Text = .Rows - 2
       .Col = 1    '�Ȥ�s��
       .Text = tmp_rs.Fields("�Ȥ�s��").Value
       .Col = 2    '�Ȥ�W��
       .Text = tmp_rs.Fields("�Ȥ�W��").Value
       .Col = 3    '�Ȥ�²��
       .Text = tmp_rs.Fields("�Ȥ�²��").Value
       .Col = 4    '�l���ϸ�
       .Text = tmp_rs.Fields("�l���ϸ�").Value
       .Col = 5    '�B�e�ϰ�
       .Text = tmp_rs.Fields("�B�e�a�}").Value
       .Col = 6    '�B�e�a�}
       .Text = tmp_rs.Fields("�B�e�ϰ�").Value
       .Col = 7    '�S��ݨD 1
       .Text = tmp_rs.Fields("�S��ݨD1").Value
       .Col = 8    '�S��ݨD 2
       .Text = tmp_rs.Fields("�S��ݨD2").Value
       .Col = 9    '�p���H
       .Text = tmp_rs.Fields("�p���H").Value
       .Col = 10   '�q��
       .Text = tmp_rs.Fields("�q��").Value
       .Col = 11   '�B�e�ϰ�X
       .Text = tmp_rs.Fields("�B�e�ϰ�X").Value
       .Col = 12   '�l���ϸ��N�X
       .Text = tmp_rs.Fields("�l���ϸ��X").Value
       .Col = 13   '�S��ݨD�X1
       .Text = tmp_rs.Fields("�S��ݨD�X1").Value
       .Col = 14   '�S��ݨD�X2
       .Text = tmp_rs.Fields("�S��ݨD�X2").Value
  End With
  tmp_rs.MoveNext
Loop
tmp_rs.Close
fam_DataLoading.Visible = False
dg_DataList.Visible = True

'���W�ٴ����ѨϥΪ̶i��ۭq�ƧǡB�M��
ReDim arFieldName(1) As String
dLoopVar2 = 0
dg_DataList.Row = 0
For dLoopVar1 = 0 To dg_DataList.Cols - 1
   dLoopVar2 = dLoopVar2 + 1           '�}�C�H 0 �}�l�A�]���̫᣸�ӷ|�O�ťաA�� [���]�w�Ƨ�]
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
'�ۭq�ƧǡB�M��G�w�]����̫᣸�ӡG�ť�
For dLoopVar1 = 0 To cmb_OrderBy.Count - 1
    cmb_OrderBy(dLoopVar1).ListIndex = cmb_OrderBy(dLoopVar1).ListCount - 1
Next dLoopVar1
cmb_Query.ListIndex = cmb_Query.ListCount - 1

End Sub
Private Sub frm_OP_ROUTEDATA_cmd_SelectCar()
'�ƨ��B�z�@�~ >> ���u�s�����@�@�~ >> �q����ƿ��
'Form_Name�Gfrm_OP_RouteData

'�]�w DataGrid �榡
Dim sub_var1 As Integer, sub_var2 As Integer
dg_DataList.Visible = False
With dg_DataList
     .FixedRows = 1: .Cols = 8
     '�]�w���\��C���
     .AllowBigSelection = True
     '�]�w�C����r�r��
     For sub_var1 = 0 To .Rows - 1
         .Row = sub_var1: .RowHeight(sub_var1) = 250
         For sub_var2 = 0 To .Cols - 1
             .Col = sub_var2
             .CellFontName = "�s�ө���": .CellFontSize = 9
         Next sub_var2
     Next sub_var1
     '�]�w�C�����e��
     .ColWidth(0) = 350
     .ColWidth(1) = 1600
     .ColWidth(2) = 700
     .ColWidth(3) = 1000
     .ColWidth(4) = 750
     .ColWidth(5) = 850
     .ColWidth(6) = 1100
     .ColWidth(7) = 1500
     '�]�w�C�����D
     .Row = 0
     .Col = 0: .Text = "No."
     .Col = 1: .Text = "����"
     .Col = 2: .Text = "���q"
     .Col = 3: .Text = "���P���X"
     .Col = 4: .Text = "�i����"
     .Col = 5: .Text = "�r�p�H"
     .Col = 6: .Text = "�q��"
     .Col = 7: .Text = "����"
     '�]�w�C����r���
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

'���o�B�e�����򥻸��
str_SQL = "Select ����,���q�O,���P���X,�i����,�r�p�H,�q��,���� From BaseData_TRPCarList Order by ����"
Call DB_CheckConnectStatus
Call ReDim_Recordset(tmp_rs)
cn.CommandTimeout = 0   '�L��������
tmp_rs.Open str_SQL, cn, adOpenForwardOnly, adLockReadOnly
cn.CommandTimeout = 120
If tmp_rs.EOF Then
   tmp_rs.Close
   msg_text = "�d�ߵ��G�G�L�ŦX�]�w���󤧹B�騮�����"
   MsgBox msg_text, vbOKOnly + vbInformation, msg_title
   Screen.MousePointer = vbDefault
End If

Do While Not tmp_rs.EOF
   With dg_DataList
       .Rows = .Rows + 1
       .Row = .Rows - 2
       .Col = 0    '�Ǹ�
       .Text = .Rows - 2
       .Col = 1    '���إN�X
       .Text = tmp_rs.Fields("����").Value
       .Col = 2    '�B�餽�q
       .Text = tmp_rs.Fields("���q�O").Value
       .Col = 3    '���P���X
       .Text = tmp_rs.Fields("���P���X").Value
       .Col = 4    '�i����
       .Text = tmp_rs.Fields("�i����").Value
       .Col = 5    '�r�p�H
       .Text = tmp_rs.Fields("�r�p�H").Value
       .Col = 6    '�q��
       .Text = tmp_rs.Fields("�q��").Value
       .Col = 7    '����
       .Text = tmp_rs.Fields("����").Value
  End With
  tmp_rs.MoveNext
Loop
tmp_rs.Close
dg_DataList.Visible = True


'���W�ٴ����ѨϥΪ̶i��ۭq�ƧǡB�M��
ReDim arFieldName(1) As String
dLoopVar2 = 0
dg_DataList.Row = 0
For dLoopVar1 = 0 To dg_DataList.Cols - 1
   dLoopVar2 = dLoopVar2 + 1           '�}�C�H 0 �}�l�A�]���̫᣸�ӷ|�O�ťաA�� [���]�w�Ƨ�]
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
'�ۭq�ƧǡB�M��G�w�]����̫᣸�ӡG�ť�
For dLoopVar1 = 0 To cmb_OrderBy.Count - 1
    cmb_OrderBy(dLoopVar1).ListIndex = cmb_OrderBy(dLoopVar1).ListCount - 1
Next dLoopVar1
cmb_Query.ListIndex = cmb_Query.ListCount - 1

End Sub
