VERSION 5.00
Begin VB.Form frm_RS_FilterAndSort 
   Caption         =   "自 訂 快 速 篩 選"
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
      Appearance      =   0  '平面
      BackColor       =   &H8000000A&
      Caption         =   "顯示符合條件的資料"
      BeginProperty Font 
         Name            =   "新細明體"
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
         Caption         =   "或"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Caption         =   "且"
         BeginProperty Font 
            Name            =   "新細明體"
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
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   300
         Style           =   2  '單純下拉式
         TabIndex        =   25
         Top             =   1140
         Width           =   2055
      End
      Begin VB.ComboBox cmb_Operator2 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2490
         Style           =   2  '單純下拉式
         TabIndex        =   24
         Top             =   1140
         Width           =   1575
      End
      Begin VB.TextBox txt_Value2 
         BeginProperty Font 
            Name            =   "新細明體"
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
            Name            =   "新細明體"
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
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2490
         Style           =   2  '單純下拉式
         TabIndex        =   2
         Top             =   405
         Width           =   1575
      End
      Begin VB.ComboBox cmb_FieldList1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   300
         Style           =   2  '單純下拉式
         TabIndex        =   1
         Top             =   405
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "允許使用 ＊與 ％，但須是字串內最後的字元"
         BeginProperty Font 
            Name            =   "新細明體"
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
         BackStyle       =   0  '透明
         Caption         =   "注意事項："
         BeginProperty Font 
            Name            =   "新細明體"
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
      Appearance      =   0  '平面
      BackColor       =   &H8000000A&
      Caption         =   "排 序"
      BeginProperty Font 
         Name            =   "新細明體"
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
         Caption         =   "確  定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2445
         Style           =   1  '圖片外觀
         TabIndex        =   28
         Top             =   1635
         Width           =   2145
      End
      Begin VB.CommandButton cmd_Cancel 
         BackColor       =   &H00FFC0FF&
         Caption         =   "取  消"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Style           =   1  '圖片外觀
         TabIndex        =   22
         Top             =   1635
         Width           =   2145
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  '平面
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
            Caption         =   "遞 減"
            BeginProperty Font 
               Name            =   "新細明體"
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
            Caption         =   "遞 增"
            BeginProperty Font 
               Name            =   "新細明體"
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
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            Style           =   2  '單純下拉式
            TabIndex        =   18
            Top             =   510
            Width           =   2055
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "第  三  鍵"
            BeginProperty Font 
               Name            =   "新細明體"
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
         Appearance      =   0  '平面
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
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            Style           =   2  '單純下拉式
            TabIndex        =   15
            Top             =   510
            Width           =   2055
         End
         Begin VB.OptionButton opt_Order2_ASC 
            BackColor       =   &H8000000A&
            Caption         =   "遞 增"
            BeginProperty Font 
               Name            =   "新細明體"
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
            Caption         =   "遞 減"
            BeginProperty Font 
               Name            =   "新細明體"
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
            BackStyle       =   0  '透明
            Caption         =   "次  要  鍵"
            BeginProperty Font 
               Name            =   "新細明體"
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
         Appearance      =   0  '平面
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
            Caption         =   "遞 減"
            BeginProperty Font 
               Name            =   "新細明體"
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
            Caption         =   "遞 增"
            BeginProperty Font 
               Name            =   "新細明體"
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
               Name            =   "新細明體"
               Size            =   12
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            Style           =   2  '單純下拉式
            TabIndex        =   9
            Top             =   510
            Width           =   2055
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  '透明
            Caption         =   "主  要  鍵"
            BeginProperty Font 
               Name            =   "新細明體"
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
      BackStyle       =   1  '不透明
      BorderColor     =   &H00000080&
      FillColor       =   &H00808000&
      Height          =   2220
      Index           =   1
      Left            =   90
      Top             =   2760
      Width           =   6915
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  '不透明
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

'自訂篩選運算元
Private arOperator(8) As String


Private Sub cmd_Cancel_Click(Index As Integer)
'取消
Call ShowForm_byName(strFormName_FilterAndSort)
Unload Me

End Sub

Private Sub cmd_CreateString_Click()
'搜尋
Dim strFilter As String
txt_Value1.Text = RTrim(txt_Value1.Text)
txt_Value2.Text = RTrim(txt_Value2.Text)
If cmb_FieldList1.ListIndex > 0 And cmb_Operator1.ListIndex > 0 And Len(txt_Value1.Text) > 0 Then
   strFilter = ""
   Select Case UCase(Right(cmb_FieldList1.List(cmb_FieldList1.ListIndex), 3))
          Case "(N)"  '欄位資料型態：數值
               If InStr(txt_Value1.Text, "*") <> 0 Or InStr(txt_Value1.Text, "%") <> 0 Then
                  msg_text = "資料篩選條件錯誤：數值資料不可以用 ＊ or ％ "
                  MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                  txt_Value1.SelStart = 0: txt_Value1.SelLength = Len(txt_Value1.Text): txt_Value1.SetFocus
                  Exit Sub
               End If
               strFilter = Mid(cmb_FieldList1.List(cmb_FieldList1.ListIndex), 1, Len(cmb_FieldList1.List(cmb_FieldList1.ListIndex)) - 3) & _
                           arOperator(cmb_Operator1.ListIndex) & txt_Value1.Text
          Case "(D)"  '欄位資料型態：日期  前後加 #
               strFilter = Mid(cmb_FieldList1.List(cmb_FieldList1.ListIndex), 1, Len(cmb_FieldList1.List(cmb_FieldList1.ListIndex)) - 3) & _
                           arOperator(cmb_Operator1.ListIndex) & "#" & txt_Value1.Text & "#"
          Case "(S)"
               If arOperator(cmb_Operator1.ListIndex) = " Like " Then '欄位資料型態：字串以及使用包含條件 前後加 '* *' add by gemini @ 20081223 4 包含條件查無資料
                strFilter = Mid(cmb_FieldList1.List(cmb_FieldList1.ListIndex), 1, Len(cmb_FieldList1.List(cmb_FieldList1.ListIndex)) - 3) & arOperator(cmb_Operator1.ListIndex) & "'*" & txt_Value1.Text & "*'"
               Else '欄位資料型態：字串  前後加 '
                strFilter = Mid(cmb_FieldList1.List(cmb_FieldList1.ListIndex), 1, Len(cmb_FieldList1.List(cmb_FieldList1.ListIndex)) - 3) & arOperator(cmb_Operator1.ListIndex) & "'" & txt_Value1.Text & "'"
               End If
'          Case "(S)"  '欄位資料型態：字串  前後加 '
'               strFilter = Mid(cmb_FieldList1.List(cmb_FieldList1.ListIndex), 1, Len(cmb_FieldList1.List(cmb_FieldList1.ListIndex)) - 3) & _
'                           arOperator(cmb_Operator1.ListIndex) & "'" & txt_Value1.Text & "'"
          Case Else   '無型態識別的欄位都把它當成是 [字串]
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
          Case "(N)"  '欄位資料型態：數值
               If InStr(txt_Value2.Text, "*") <> 0 Or InStr(txt_Value2.Text, "%") <> 0 Then
                  msg_text = "資料篩選條件錯誤：數值資料不可以用 ＊ or ％ "
                  MsgBox msg_text, vbOKOnly + vbInformation, msg_title
                  txt_Value2.SelStart = 0: txt_Value2.SelLength = Len(txt_Value2.Text): txt_Value2.SetFocus
                  Exit Sub
               End If
               strFilter = strFilter & Mid(cmb_FieldList2.List(cmb_FieldList2.ListIndex), 1, Len(cmb_FieldList2.List(cmb_FieldList2.ListIndex)) - 3) & _
                           arOperator(cmb_Operator2.ListIndex) & txt_Value2.Text
          Case "(D)"  '欄位資料型態：日期  前後加 #
               strFilter = strFilter & Mid(cmb_FieldList2.List(cmb_FieldList2.ListIndex), 1, Len(cmb_FieldList2.List(cmb_FieldList2.ListIndex)) - 3) & _
                           arOperator(cmb_Operator2.ListIndex) & "#" & txt_Value2.Text & "#"
          Case "(S)"  '欄位資料型態：字串  前後加 *'
               If arOperator(cmb_Operator2.ListIndex) = " Like " Then '欄位資料型態：字串以及使用包含條件 前後加 '* *' add by gemini @ 20081223 4 包含條件查無資料
                strFilter = strFilter & Mid(cmb_FieldList2.List(cmb_FieldList2.ListIndex), 1, Len(cmb_FieldList2.List(cmb_FieldList2.ListIndex)) - 3) & arOperator(cmb_Operator2.ListIndex) & "'*" & txt_Value2.Text & "*'"
               Else '欄位資料型態：字串  前後加 '
                strFilter = strFilter & Mid(cmb_FieldList2.List(cmb_FieldList2.ListIndex), 1, Len(cmb_FieldList2.List(cmb_FieldList2.ListIndex)) - 3) & arOperator(cmb_Operator2.ListIndex) & "'" & txt_Value2.Text & "'"
               End If
'               strFilter = strFilter & Mid(cmb_FieldList1.List(cmb_FieldList2.ListIndex), 1, Len(cmb_FieldList2.List(cmb_FieldList2.ListIndex)) - 3) & _
'                           arOperator(cmb_Operator2.ListIndex) & "'*" & txt_Value2.Text & "*'"
          Case Else   '無型態識別的欄位都把它當成是 [字串]
               strFilter = strFilter & Mid(cmb_FieldList1.List(cmb_FieldList2.ListIndex), 1, Len(cmb_FieldList2.List(cmb_FieldList2.ListIndex)) - 3) & _
                           arOperator(cmb_Operator2.ListIndex) & "'" & txt_Value2.Text & "'"
   End Select
End If
If Len(strFilter) > 0 Then strFilter = "(" & strFilter & ")"

'回到呼叫的 Form 的 Public Sub
Select Case UCase(strFormName_FilterAndSort)
       Case "FRM_OP_CUTORDERS"       '排車處理作業 >> ㄧ單多車訂單切割
            frm_OP_CutOrders.frm_OP_CutOrders_rsFilterAndSort "FILTER", strFilter
            frm_OP_CutOrders.WindowState = 2
            'Unload Me
       Case "FRM_OP_TRPPLAN"         '排車處理作業 >> 排車作業
            frm_OP_TRPPlan.frm_OP_TRPPlan_rsFilterAndSort "FILTER", strFilter
            frm_OP_TRPPlan.WindowState = 2
            'Unload Me
       Case "FRM_OP_DCROUTEMERGE"    '排車處理作業 >> DC 併車作業
            frm_OP_DCRouteMerge.frm_OP_DCRouteMerge_rsFilterAndSort "FILTER", strFilter
            frm_OP_DCRouteMerge.WindowState = 2
            'Unload Me
       Case "FRM_BASEDATA_CONSIGCAR"    '基本資料維護 >> 客戶/車輛/貨運公司 基本資料
            frm_BaseData_ConsigCar.frm_BaseData_ConsigCar_rsFilterAndSort "FILTER", strFilter
            frm_BaseData_ConsigCar.WindowState = 2
            'Unload Me
       Case "FRM_OTHER_OPTPLAN"    '退貨排車 >> 排車作業
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
'排序設定
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
   '回到呼叫的 Form 的 Public Sub
   Select Case UCase(strFormName_FilterAndSort)
          Case "FRM_OP_CUTORDERS"        '排車處理作業 >> ㄧ單多車訂單切割
               frm_OP_CutOrders.frm_OP_CutOrders_rsFilterAndSort "SORT", strOrder
               frm_OP_CutOrders.WindowState = 2
               'Unload Me
          Case "FRM_OP_TRPPLAN"          '排車處理作業 >> 排車作業,
               frm_OP_TRPPlan.frm_OP_TRPPlan_rsFilterAndSort "SORT", strOrder
               frm_OP_TRPPlan.WindowState = 2
               'Unload Me
          Case "FRM_OP_DCROUTEMERGE"     '排車處理作業 >> DC 併車作業
               frm_OP_DCRouteMerge.frm_OP_DCRouteMerge_rsFilterAndSort "SORT", strOrder
               frm_OP_DCRouteMerge.WindowState = 2
               'Unload Me
          Case "FRM_BASEDATA_CONSIGCAR"    '基本資料維護 >> 客戶/車輛/貨運公司 基本資料
               frm_BaseData_ConsigCar.frm_BaseData_ConsigCar_rsFilterAndSort "SORT", strOrder
               frm_BaseData_ConsigCar.WindowState = 2
               'Unload Me
          Case "FRM_OTHER_OPTPLAN"    '退貨排車 >> 排車作業
               frm_Other_OPTPlan.frm_OP_TRPPlan_rsFilterAndSort "SORT", strOrder
               frm_Other_OPTPlan.WindowState = 2
               'Unload Me
          Case Else
   End Select
End If
'Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'攔截整個表單鍵盤按鍵事件
'用途：使用者按下 Esc 則不傳回任何資料，且關閉視窗
If KeyCode = vbKeyEscape Then
   Call ShowForm_byName(strFormName_FilterAndSort)
   Unload Me
End If

End Sub

Private Sub Form_Load()
'設定 Form 大小、位置
Me.Height = 5475: Me.Width = 7215
Me.Left = (frm_MDIForm.ScaleWidth - Me.Width) / 2
Me.Top = (frm_MDIForm.ScaleHeight - Me.ScaleHeight) / 2 - 200

'自訂篩選資料比對運算元
arOperator(0) = "": cmb_Operator1.AddItem "": cmb_Operator2.AddItem ""
arOperator(1) = " = ": cmb_Operator1.AddItem "等於": cmb_Operator2.AddItem "等於"
arOperator(2) = " > ": cmb_Operator1.AddItem "大於": cmb_Operator2.AddItem "大於"
arOperator(3) = " >= ": cmb_Operator1.AddItem "大於或等於": cmb_Operator2.AddItem "大於或等於"
arOperator(4) = " < ": cmb_Operator1.AddItem "小於": cmb_Operator2.AddItem "小於"
arOperator(5) = " <= ": cmb_Operator1.AddItem "小於或等於": cmb_Operator2.AddItem "小於或等於"
arOperator(6) = " <> ": cmb_Operator1.AddItem "不等於": cmb_Operator2.AddItem "不等於"
arOperator(7) = " Like ": cmb_Operator1.AddItem "包含": cmb_Operator2.AddItem "包含"

End Sub

Private Sub txt_Value1_KeyPress(KeyAscii As Integer)
'自訂篩選 >> 條件 1
Select Case KeyAscii
     Case 97 To 122     '轉緩為大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
End Select
End Sub

Private Sub txt_Value2_KeyPress(KeyAscii As Integer)
'自訂篩選 >> 條件 2
Select Case KeyAscii
     Case 97 To 122     '轉緩為大寫字元
          KeyAscii = KeyAscii - 32
     Case vbKeyReturn
          KeyAscii = 0
End Select
End Sub
