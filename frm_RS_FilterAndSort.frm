VERSION 5.00
Begin VB.Form frm_RS_FilterAndSort 
   Caption         =   "�� �q �� �t �z ��"
   ClientHeight    =   5070
   ClientLeft      =   2715
   ClientTop       =   2040
   ClientWidth     =   7095
   Icon            =   "frm_RS_FilterAndSort.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  '����
      BackColor       =   &H8000000A&
      Caption         =   "��ܲŦX���󪺸��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2430
      Left            =   150
      TabIndex        =   0
      Top             =   75
      Width           =   6810
      Begin VB.OptionButton opt_OR 
         BackColor       =   &H8000000A&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1470
         TabIndex        =   27
         Top             =   840
         Width           =   615
      End
      Begin VB.OptionButton opt_AND 
         BackColor       =   &H8000000A&
         Caption         =   "�B"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   540
         TabIndex        =   26
         Top             =   840
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.ComboBox cmb_FieldList2 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   300
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   25
         Top             =   1140
         Width           =   2055
      End
      Begin VB.ComboBox cmb_Operator2 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2490
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   24
         Top             =   1140
         Width           =   1575
      End
      Begin VB.TextBox txt_Value2 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         TabIndex        =   23
         Top             =   1140
         Width           =   2310
      End
      Begin VB.TextBox txt_Value1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         TabIndex        =   3
         Top             =   405
         Width           =   2310
      End
      Begin VB.ComboBox cmb_Operator1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2490
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   2
         Top             =   405
         Width           =   1575
      End
      Begin VB.ComboBox cmb_FieldList1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   300
         Style           =   2  '��¤U�Ԧ�
         TabIndex        =   1
         Top             =   405
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "���\�ϥ� ���P �H�A�����O�r�ꤺ�̫᪺�r��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   615
         TabIndex        =   5
         Top             =   2025
         Width           =   4110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '�z��
         Caption         =   "�`�N�ƶ��G"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   315
         TabIndex        =   4
         Top             =   1680
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '����
      BackColor       =   &H8000000A&
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   2295
      Left            =   135
      TabIndex        =   6
      Top             =   2640
      Width           =   6825
      Begin VB.CommandButton cmd_DoCommand 
         BackColor       =   &H00FF8080&
         Caption         =   "�T  �w"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2445
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   28
         Top             =   1635
         Width           =   2145
      End
      Begin VB.CommandButton cmd_Cancel 
         BackColor       =   &H00FFC0FF&
         Caption         =   "��  ��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   4620
         Style           =   1  '�Ϥ��~�[
         TabIndex        =   22
         Top             =   1635
         Width           =   2145
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '����
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   1320
         Index           =   2
         Left            =   4575
         TabIndex        =   17
         Top             =   270
         Width           =   2175
         Begin VB.OptionButton opt_Order3_DESC 
            BackColor       =   &H8000000A&
            Caption         =   "�� ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1155
            TabIndex        =   20
            Top             =   990
            Width           =   825
         End
         Begin VB.OptionButton opt_Order3_ASC 
            BackColor       =   &H8000000A&
            Caption         =   "�� �W"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   990
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.ComboBox cmb_Order3_FieldList 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   18
            Top             =   510
            Width           =   2055
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��  �T  ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   2
            Left            =   60
            TabIndex        =   21
            Top             =   225
            Width           =   1020
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '����
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   1320
         Index           =   1
         Left            =   2370
         TabIndex        =   12
         Top             =   270
         Width           =   2175
         Begin VB.ComboBox cmb_Order2_FieldList 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   15
            Top             =   510
            Width           =   2055
         End
         Begin VB.OptionButton opt_Order2_ASC 
            BackColor       =   &H8000000A&
            Caption         =   "�� �W"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   195
            TabIndex        =   14
            Top             =   990
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.OptionButton opt_Order2_DESC 
            BackColor       =   &H8000000A&
            Caption         =   "�� ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1125
            TabIndex        =   13
            Top             =   990
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "��  �n  ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   1
            Left            =   60
            TabIndex        =   16
            Top             =   225
            Width           =   1020
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '����
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   1320
         Index           =   0
         Left            =   165
         TabIndex        =   7
         Top             =   270
         Width           =   2175
         Begin VB.OptionButton opt_Order1_DESC 
            BackColor       =   &H8000000A&
            Caption         =   "�� ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1140
            TabIndex        =   11
            Top             =   990
            Width           =   825
         End
         Begin VB.OptionButton opt_Order1_ASC 
            BackColor       =   &H8000000A&
            Caption         =   "�� �W"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   9.75
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   210
            TabIndex        =   10
            Top             =   990
            Value           =   -1  'True
            Width           =   825
         End
         Begin VB.ComboBox cmb_Order1_FieldList 
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            Style           =   2  '��¤U�Ԧ�
            TabIndex        =   9
            Top             =   510
            Width           =   2055
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '�z��
            Caption         =   "�D  �n  ��"
            BeginProperty Font 
               Name            =   "�s�ө���"
               Size            =   11.25
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Top             =   225
            Width           =   1020
         End
      End
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '���z��
      BorderColor     =   &H00000080&
      FillColor       =   &H00808000&
      Height          =   2220
      Index           =   1
      Left            =   90
      Top             =   2760
      Width           =   6915
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '���z��
      BorderColor     =   &H00000080&
      FillColor       =   &H00808000&
      Height          =   2385
      Index           =   0
      Left            =   105
      Top             =   165
      Width           =   6900
   End
End
Attribute VB_Name = "frm_RS_FilterAndSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�ۭq�z��B�⤸
Private arOperator(8) As String


Private Sub cmd_Cancel_Click(Index As Integer)
'����
Call ShowForm_byName(strFormName_FilterAndSort)
Unload Me

End Sub

Private Sub cmd_CreateString_Click()
'�j�M
Dim strFilter As String
txt_Value1.Text = RTrim(txt_Value1.Text)
txt_Value2.Text = RTrim(txt_Value2.Text)
If cmb_FieldList1.ListIndex > 0 And cmb_Operator1.ListIndex > 0 And Len(txt_Value1.Text) > 0 Then
   strFilter = ""
   Select Case UCase(Right(cmb_FieldList1.List(cmb_FieldList1.ListIndex), 3))
          Case "(N)"  '����ƫ��A�G�ƭ�
               If InStr(txt_Value1.Text, "*") <> 0 Or InStr(txt_Value1.Text, "%") <> 0 Then
                  msg_text = "��ƿz�������~�G�ƭȸ�Ƥ��i�H�� �� or �H "
                  MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                  txt_Value1.SelStart = 0: txt_Value1.SelLength = Len(txt_Value1.Text): txt_Value1.SetFocus
                  Exit Sub
               End If
               strFilter = Mid(cmb_FieldList1.List(cmb_FieldList1.ListIndex), 1, Len(cmb_FieldList1.List(cmb_FieldList1.ListIndex)) - 3) & _
                           arOperator(cmb_Operator1.ListIndex) & txt_Value1.Text
          Case "(D)"  '����ƫ��A�G���  �e��[ #
               strFilter = Mid(cmb_FieldList1.List(cmb_FieldList1.ListIndex), 1, Len(cmb_FieldList1.List(cmb_FieldList1.ListIndex)) - 3) & _
                           arOperator(cmb_Operator1.ListIndex) & "#" & txt_Value1.Text & "#"
          Case "(S)"
               If arOperator(cmb_Operator1.ListIndex) = " Like " Then '����ƫ��A�G�r��H�ΨϥΥ]�t���� �e��[ '* *' add by gemini @ 20081223 4 �]�t����d�L���
                strFilter = Mid(cmb_FieldList1.List(cmb_FieldList1.ListIndex), 1, Len(cmb_FieldList1.List(cmb_FieldList1.ListIndex)) - 3) & arOperator(cmb_Operator1.ListIndex) & "'*" & txt_Value1.Text & "*'"
               Else '����ƫ��A�G�r��  �e��[ '
                strFilter = Mid(cmb_FieldList1.List(cmb_FieldList1.ListIndex), 1, Len(cmb_FieldList1.List(cmb_FieldList1.ListIndex)) - 3) & arOperator(cmb_Operator1.ListIndex) & "'" & txt_Value1.Text & "'"
               End If
'          Case "(S)"  '����ƫ��A�G�r��  �e��[ '
'               strFilter = Mid(cmb_FieldList1.List(cmb_FieldList1.ListIndex), 1, Len(cmb_FieldList1.List(cmb_FieldList1.ListIndex)) - 3) & _
'                           arOperator(cmb_Operator1.ListIndex) & "'" & txt_Value1.Text & "'"
          Case Else   '�L���A�ѧO����쳣�⥦���O [�r��]
               strFilter = Mid(cmb_FieldList1.List(cmb_FieldList1.ListIndex), 1, Len(cmb_FieldList1.List(cmb_FieldList1.ListIndex)) - 3) & _
                           arOperator(cmb_Operator1.ListIndex) & "'" & txt_Value1.Text & "'"
   End Select
End If
If cmb_FieldList2.ListIndex > 0 And cmb_Operator2.ListIndex > 0 And Len(txt_Value2.Text) > 0 Then
   If Len(strFilter) > 0 Then
      If opt_AND.Value Then
         strFilter = strFilter & " And "
      ElseIf opt_OR.Value Then
         strFilter = strFilter & " or "
      Else
         Unload Me
      End If
   End If
   Select Case UCase(Right(cmb_FieldList2.List(cmb_FieldList2.ListIndex), 3))
          Case "(N)"  '����ƫ��A�G�ƭ�
               If InStr(txt_Value2.Text, "*") <> 0 Or InStr(txt_Value2.Text, "%") <> 0 Then
                  msg_text = "��ƿz�������~�G�ƭȸ�Ƥ��i�H�� �� or �H "
                  MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                  txt_Value2.SelStart = 0: txt_Value2.SelLength = Len(txt_Value2.Text): txt_Value2.SetFocus
                  Exit Sub
               End If
               strFilter = strFilter & Mid(cmb_FieldList2.List(cmb_FieldList2.ListIndex), 1, Len(cmb_FieldList2.List(cmb_FieldList2.ListIndex)) - 3) & _
                           arOperator(cmb_Operator2.ListIndex) & txt_Value2.Text
          Case "(D)"  '����ƫ��A�G���  �e��[ #
               strFilter = strFilter & Mid(cmb_FieldList2.List(cmb_FieldList2.ListIndex), 1, Len(cmb_FieldList2.List(cmb_FieldList2.ListIndex)) - 3) & _
                           arOperator(cmb_Operator2.ListIndex) & "#" & txt_Value2.Text & "#"
          Case "(S)"  '����ƫ��A�G�r��  �e��[ *'
               If arOperator(cmb_Operator2.ListIndex) = " Like " Then '����ƫ��A�G�r��H�ΨϥΥ]�t���� �e��[ '* *' add by gemini @ 20081223 4 �]�t����d�L���
                strFilter = strFilter & Mid(cmb_FieldList2.List(cmb_FieldList2.ListIndex), 1, Len(cmb_FieldList2.List(cmb_FieldList2.ListIndex)) - 3) & arOperator(cmb_Operator2.ListIndex) & "'*" & txt_Value2.Text & "*'"
               Else '����ƫ��A�G�r��  �e��[ '
                strFilter = strFilter & Mid(cmb_FieldList2.List(cmb_FieldList2.ListIndex), 1, Len(cmb_FieldList2.List(cmb_FieldList2.ListIndex)) - 3) & arOperator(cmb_Operator2.ListIndex) & "'" & txt_Value2.Text & "'"
               End If
'               strFilter = strFilter & Mid(cmb_FieldList1.List(cmb_FieldList2.ListIndex), 1, Len(cmb_FieldList2.List(cmb_FieldList2.ListIndex)) - 3) & _
'                           arOperator(cmb_Operator2.ListIndex) & "'*" & txt_Value2.Text & "*'"
          Case Else   '�L���A�ѧO����쳣�⥦���O [�r��]
               strFilter = strFilter & Mid(cmb_FieldList1.List(cmb_FieldList2.ListIndex), 1, Len(cmb_FieldList2.List(cmb_FieldList2.ListIndex)) - 3) & _
                           arOperator(cmb_Operator2.ListIndex) & "'" & txt_Value2.Text & "'"
   End Select
End If
If Len(strFilter) > 0 Then strFilter = "(" & strFilter & ")"

'�^��I�s�� Form �� Public Sub
Select Case UCase(strFormName_FilterAndSort)
       Case "FRM_OP_CUTORDERS"       '�ƨ��B�z�@�~ >> ����h���q�����
            frm_OP_CutOrders.frm_OP_CutOrders_rsFilterAndSort "FILTER", strFilter
            frm_OP_CutOrders.WindowState = 2
            'Unload Me
       Case "FRM_OP_TRPPLAN"         '�ƨ��B�z�@�~ >> �ƨ��@�~
            frm_OP_TRPPlan.frm_OP_TRPPlan_rsFilterAndSort "FILTER", strFilter
            frm_OP_TRPPlan.WindowState = 2
            'Unload Me
       Case "FRM_OP_DCROUTEMERGE"    '�ƨ��B�z�@�~ >> DC �֨��@�~
            frm_OP_DCRouteMerge.frm_OP_DCRouteMerge_rsFilterAndSort "FILTER", strFilter
            frm_OP_DCRouteMerge.WindowState = 2
            'Unload Me
       Case "FRM_BASEDATA_CONSIGCAR"    '�򥻸�ƺ��@ >> �Ȥ�/����/�f�B���q �򥻸��
            frm_BaseData_ConsigCar.frm_BaseData_ConsigCar_rsFilterAndSort "FILTER", strFilter
            frm_BaseData_ConsigCar.WindowState = 2
            'Unload Me
       Case "FRM_OTHER_OPTPLAN"    '�h�f�ƨ� >> �ƨ��@�~
            frm_Other_OPTPlan.frm_OP_TRPPlan_rsFilterAndSort "FILTER", strFilter
            frm_Other_OPTPlan.WindowState = 2
            'Unload Me
       Case Else
End Select

End Sub

Private Sub cmd_DoCommand_Click()
Call cmd_CreateString_Click
Call cmd_OrderBy_Click
Unload Me
End Sub

Private Sub cmd_OrderBy_Click()
'�Ƨǳ]�w
Dim strOrder As String
strOrder = ""
If cmb_Order1_FieldList.ListIndex <> -1 And Len(cmb_Order1_FieldList.List(cmb_Order1_FieldList.ListIndex)) > 0 Then
   If Len(strOrder) > 0 Then
      If opt_Order1_ASC.Value Then
         strOrder = strOrder & "," & Mid(cmb_Order1_FieldList.List(cmb_Order1_FieldList.ListIndex), 1, Len(cmb_Order1_FieldList.List(cmb_Order1_FieldList.ListIndex)) - 3) & " ASC "
      Else
         strOrder = strOrder & "," & Mid(cmb_Order1_FieldList.List(cmb_Order1_FieldList.ListIndex), 1, Len(cmb_Order1_FieldList.List(cmb_Order1_FieldList.ListIndex)) - 3) & " DESC "
      End If
   Else
      If opt_Order1_ASC.Value Then
         strOrder = Mid(cmb_Order1_FieldList.List(cmb_Order1_FieldList.ListIndex), 1, Len(cmb_Order1_FieldList.List(cmb_Order1_FieldList.ListIndex)) - 3) & " ASC "
      Else
         strOrder = Mid(cmb_Order1_FieldList.List(cmb_Order1_FieldList.ListIndex), 1, Len(cmb_Order1_FieldList.List(cmb_Order1_FieldList.ListIndex)) - 3) & " DESC "
      End If
   End If
End If
If cmb_Order2_FieldList.ListIndex <> -1 And Len(cmb_Order1_FieldList.List(cmb_Order1_FieldList.ListIndex)) > 0 Then
   If Len(strOrder) > 0 Then
      If opt_Order2_ASC.Value Then
         strOrder = strOrder & "," & Mid(cmb_Order2_FieldList.List(cmb_Order2_FieldList.ListIndex), 1, Len(cmb_Order2_FieldList.List(cmb_Order2_FieldList.ListIndex)) - 3) & " ASC "
      Else
         strOrder = strOrder & "," & Mid(cmb_Order2_FieldList.List(cmb_Order2_FieldList.ListIndex), 1, Len(cmb_Order2_FieldList.List(cmb_Order2_FieldList.ListIndex)) - 3) & " DESC "
      End If
   Else
      If opt_Order2_ASC.Value Then
         strOrder = Mid(cmb_Order2_FieldList.List(cmb_Order2_FieldList.ListIndex), 1, Len(cmb_Order2_FieldList.List(cmb_Order2_FieldList.ListIndex)) - 3) & " ASC "
      Else
         strOrder = Mid(cmb_Order2_FieldList.List(cmb_Order2_FieldList.ListIndex), 1, Len(cmb_Order2_FieldList.List(cmb_Order2_FieldList.ListIndex)) - 3) & " DESC "
      End If
   End If
End If
If cmb_Order3_FieldList.ListIndex <> -1 And Len(cmb_Order1_FieldList.List(cmb_Order1_FieldList.ListIndex)) > 0 Then
   If Len(strOrder) > 0 Then
      If opt_Order3_ASC.Value Then
         strOrder = strOrder & "," & Mid(cmb_Order3_FieldList.List(cmb_Order3_FieldList.ListIndex), 1, Len(cmb_Order3_FieldList.List(cmb_Order3_FieldList.ListIndex)) - 3) & " ASC "
      Else
         strOrder = strOrder & "," & Mid(cmb_Order3_FieldList.List(cmb_Order3_FieldList.ListIndex), 1, Len(cmb_Order3_FieldList.List(cmb_Order3_FieldList.ListIndex)) - 3) & " DESC "
      End If
   Else
      If opt_Order3_ASC.Value Then
         strOrder = Mid(cmb_Order3_FieldList.List(cmb_Order3_FieldList.ListIndex), 1, Len(cmb_Order3_FieldList.List(cmb_Order3_FieldList.ListIndex)) - 3) & " ASC "
      Else
         strOrder = Mid(cmb_Order3_FieldList.List(cmb_Order3_FieldList.ListIndex), 1, Len(cmb_Order3_FieldList.List(cmb_Order3_FieldList.ListIndex)) - 3) & " DESC "
      End If
   End If
End If

If Len(strOrder) > 0 Then
   '�^��I�s�� Form �� Public Sub
   Select Case UCase(strFormName_FilterAndSort)
          Case "FRM_OP_CUTORDERS"        '�ƨ��B�z�@�~ >> ����h���q�����
               frm_OP_CutOrders.frm_OP_CutOrders_rsFilterAndSort "SORT", strOrder
               frm_OP_CutOrders.WindowState = 2
               'Unload Me
          Case "FRM_OP_TRPPLAN"          '�ƨ��B�z�@�~ >> �ƨ��@�~,
               frm_OP_TRPPlan.frm_OP_TRPPlan_rsFilterAndSort "SORT", strOrder
               frm_OP_TRPPlan.WindowState = 2
               'Unload Me
          Case "FRM_OP_DCROUTEMERGE"     '�ƨ��B�z�@�~ >> DC �֨��@�~
               frm_OP_DCRouteMerge.frm_OP_DCRouteMerge_rsFilterAndSort "SORT", strOrder
               frm_OP_DCRouteMerge.WindowState = 2
               'Unload Me
          Case "FRM_BASEDATA_CONSIGCAR"    '�򥻸�ƺ��@ >> �Ȥ�/����/�f�B���q �򥻸��
               frm_BaseData_ConsigCar.frm_BaseData_ConsigCar_rsFilterAndSort "SORT", strOrder
               frm_BaseData_ConsigCar.WindowState = 2
               'Unload Me
          Case "FRM_OTHER_OPTPLAN"    '�h�f�ƨ� >> �ƨ��@�~
               frm_Other_OPTPlan.frm_OP_TRPPlan_rsFilterAndSort "SORT", strOrder
               frm_Other_OPTPlan.WindowState = 2
               'Unload Me
          Case Else
   End Select
End If
'Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'�d�I��Ӫ����L����ƥ�
'�γ~�G�ϥΪ̫��U Esc �h���Ǧ^�����ơA�B��������
If KeyCode = vbKeyEscape Then
   Call ShowForm_byName(strFormName_FilterAndSort)
   Unload Me
End If

End Sub

Private Sub Form_Load()
'�]�w Form �j�p�B��m
Me.Height = 5475: Me.Width = 7215
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200

'�ۭq�z���Ƥ��B�⤸
arOperator(0) = "": cmb_Operator1.AddItem "": cmb_Operator2.AddItem ""
arOperator(1) = " = ": cmb_Operator1.AddItem "����": cmb_Operator2.AddItem "����"
arOperator(2) = " > ": cmb_Operator1.AddItem "�j��": cmb_Operator2.AddItem "�j��"
arOperator(3) = " >= ": cmb_Operator1.AddItem "�j��ε���": cmb_Operator2.AddItem "�j��ε���"
arOperator(4) = " < ": cmb_Operator1.AddItem "�p��": cmb_Operator2.AddItem "�p��"
arOperator(5) = " <= ": cmb_Operator1.AddItem "�p��ε���": cmb_Operator2.AddItem "�p��ε���"
arOperator(6) = " <> ": cmb_Operator1.AddItem "������": cmb_Operator2.AddItem "������"
arOperator(7) = " Like ": cmb_Operator1.AddItem "�]�t": cmb_Operator2.AddItem "�]�t"

End Sub

Private Sub txt_Value1_KeyPress(KeyAscii As Integer)
'�ۭq�z�� >> ���� 1
Select Case KeyAscii
     Case 97 To 122     '��w���j�g�r��
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
End Select
End Sub

Private Sub txt_Value2_KeyPress(KeyAscii As Integer)
'�ۭq�z�� >> ���� 2
Select Case KeyAscii
     Case 97 To 122     '��w���j�g�r��
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
End Select
End Sub
